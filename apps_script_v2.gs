/**
 * Waqti cafe — Sheet bridge v2 (orders + owner admin: list/cancel orders, menu management).
 * Paste into Extensions -> Apps Script on the CLIENT'S OWN spreadsheet
 * (one deployment per client — this is a template, not shared code).
 *
 * This is apps_script.gs (Elbob's live script) + admin + menu-management actions.
 * Elbob's live deployment is untouched — this is for new clients (Squeeze first).
 *
 * doPost                          -> append an order row (no auth, matches v1). Returns
 *                                     {error:...} instead of throwing on bad input, so the
 *                                     page can tell a real failure apart from a real success.
 *                                     Items are stored as a JSON array (name/qty/price per
 *                                     line), not a flattened display string -- lets the
 *                                     printer bridge show a price per line on the cashier
 *                                     copy and a no-price kitchen copy from the same data.
 *                                     Also writes an optional name+phone (if either was
 *                                     given) to a "Customers" tab, for a gift/offer later --
 *                                     silently skipped if that tab doesn't exist yet.
 * doGet                           -> today's order count + expected cash (unchanged, n8n cron)
 * doGet?action=list&key=K         -> today's orders as JSON (admin dashboard)
 * doGet?action=cancel&id=I&key=K  -> mark one order 'cancelled' by its order_id (admin dashboard)
 * doGet?action=history&from=&to=&key=K -> per-day order count + total in a date range
 *                                     (yyyy-MM-dd, inclusive), for the admin dashboard's
 *                                     history view. Reads full Orders history on purpose
 *                                     (owner-initiated, not polled) -- see readOrdersTail_
 *                                     below for why today/list/cancel don't.
 * doGet?action=menu               -> full Menu tab as JSON rows (no auth — same exposure the
 *                                     old published-CSV link already had; it's just a menu)
 * doGet?action=menu_set&...&key=K -> add or edit ONE Menu row by item_id+variant (owner
 *                                     menu-management screen: stock toggle, price/name edit,
 *                                     new item). Upsert: matching row updates in place, no
 *                                     match appends a new row.
 * doGet?action=order_edit&id=&table=&items=<JSON>&total=&key=K -> fix a mistake (wrong item/
 *                                     qty/price/table) on an existing order without cancelling
 *                                     it. Overwrites only table/items/total at that row.
 * doGet?action=expense_add&amount=&note=&key=K     -> log a cash expense (تقفيل)
 * doGet?action=expense_void&id=&key=K              -> void a mistaken expense entry
 * doGet?action=expense_list&key=K                  -> today's expenses + running total.
 *                                     Optional &date=yyyy-MM-dd views any past day instead
 *                                     (admin dashboard's ◀ السابق / التالي ▶ day-nav).
 * doGet?action=report_request&from=&to=&key=K      -> queue a تقفيل print job (owner's phone
 *                                     tap) for print_bridge.py to pick up and print
 * doGet?action=report_pending&key=K                -> queued report print jobs (print_bridge.py)
 * doGet?action=report_done&id=&key=K               -> delete a report job once printed
 * doGet?action=report_data&from=&to=&key=K         -> sales+expenses+net aggregate, PLUS a
 *                                     per-item qty+total breakdown (items_sold, biggest
 *                                     earners first) for a تقفيل ticket's content
 *                                     (print_bridge.py). Legacy pre-JSON orders (items stored
 *                                     as a display string) still count toward sales_total/
 *                                     count but can't be itemized -- silently skipped there.
 *
 * SET THIS before deploying — same key you put in the client's admin_key field:
 */
var ADMIN_KEY = 'CHANGE_ME';

var TZ = 'Africa/Cairo';
var DAY_START_HOUR = 6;   // business-day boundary (6 AM) — see apps_script.gs for why

// doPost has no auth by design (customers order with no login) -- these bounds exist so an
// arbitrary client-supplied total/price can't land as real cash-reconciliation data, and so
// a scripted flood can't blow up a single Sheet cell or row. Generous on purpose: picked to
// never reject a real order, only an obviously-wrong or adversarial one.
var MAX_ORDER_TOTAL = 5000;    // EGP -- biggest plausible real order has headroom to spare
var MAX_ITEMS_PER_ORDER = 20;
var MAX_FIELD_LEN = 300;       // per item name / table label
var MAX_MENU_PRICE = 2000;     // EGP -- menu_set is admin_key-gated, but bound it too in case that key ever leaks
var MAX_EXPENSE_AMOUNT = 50000; // EGP -- expense_add is admin_key-gated, bounded same spirit as menu price

function bizToken(d) {
  var hour = Number(Utilities.formatDate(d, TZ, 'H'));
  var base = (hour < DAY_START_HOUR) ? new Date(d.getTime() - 86400000) : d;
  return 'b' + Utilities.formatDate(base, TZ, 'yyyyMMdd');
}
function bizDateHuman(d) {
  var hour = Number(Utilities.formatDate(d, TZ, 'H'));
  var base = (hour < DAY_START_HOUR) ? new Date(d.getTime() - 86400000) : d;
  return Utilities.formatDate(base, TZ, 'yyyy-MM-dd');
}
function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Writes a row, forcing plain-text format ('@') on column 1 (always a timestamp) PLUS
// whichever other columns are passed in textCols BEFORE the values land -- order matters,
// Sheets decides whether to auto-convert a string into its own Date type (using the
// SPREADSHEET'S OWN timezone setting, not necessarily Cairo) at the moment of write, not
// retroactively. Setting the format after the fact would be too late. Column 1 alone was
// the actual fix for the Oct-2026 wrong-receipt-time bug (root cause: the spreadsheet's
// timezone wasn't Cairo, so Sheets silently reinterpreted the already-Cairo-formatted
// timestamp string through a different offset). textCols covers the SAME class of bug for
// any other plain "YYYY-MM-DD"-shaped column -- e.g. reportRequestAdd_'s from/to, which
// Sheets was silently turning into real Date cells, breaking reportData_'s strict regex
// check on read-back (confirmed live on the Squeeze cafe PC, 2026-10-12: تقفيل printing
// failed with "bad_range" every time because of exactly this).
function appendTextSafeRow_(sheet, values, textCols) {
  var row = sheet.getLastRow() + 1;
  (textCols || [1]).forEach(function (col) { sheet.getRange(row, col).setNumberFormat('@'); });
  sheet.getRange(row, 1, 1, values.length).setValues([values]);
}

// Shared by doPost (new order) and orderEdit_ (fix a mistake) -- same slicing/bounds logic,
// not duplicated between the two write paths.
function sanitizeItems_(rawItems) {
  return rawItems.map(function (i) {
    return {
      name: String(i.name || '').slice(0, MAX_FIELD_LEN),
      qty: Number(i.qty) || 0,
      price: Number(i.price) || 0,
    };
  });
}

function doPost(e) {
  var data;
  try { data = JSON.parse(e.postData.contents); } catch (err) { return json({ error: 'bad_json' }); }
  if (!data || !Array.isArray(data.items) || !data.items.length) return json({ error: 'empty_order' });
  if (data.items.length > MAX_ITEMS_PER_ORDER) return json({ error: 'too_many_items' });

  var total = Number(data.total);
  if (isNaN(total) || total <= 0 || total > MAX_ORDER_TOTAL) return json({ error: 'bad_total' });

  var table = String(data.table || '').slice(0, MAX_FIELD_LEN);
  // Stored as JSON (not the old flattened "qty× name، qty× name" display string) so the
  // printer bridge can show a price per line on the cashier copy and a no-price kitchen
  // copy from the same data -- the old flattened string threw the price away entirely and
  // only the grand total ever printed. Also incidentally fixes the pre-existing bug where an
  // Arabic comma inside a customer's own note got mis-split into extra receipt lines, since
  // there's no more string-splitting involved at all.
  var items = sanitizeItems_(data.items);

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Orders');
  var now = new Date();
  var ts = Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm:ss');
  var orderId = Utilities.getUuid().slice(0, 8);
  appendTextSafeRow_(sh, [ts, table, JSON.stringify(items), total, 'new', bizToken(now), orderId]);

  // Optional customer capture (name + phone, for a gift/offer later) -- never blocks the
  // order itself. Written only if the owner has created a "Customers" tab; if she hasn't
  // yet, this just silently does nothing rather than failing the order.
  var custName = String(data.customer_name || '').trim().slice(0, MAX_FIELD_LEN);
  var custPhone = String(data.customer_phone || '').trim().slice(0, MAX_FIELD_LEN);
  if (custName || custPhone) {
    var custSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Customers');
    if (custSheet) {
      appendTextSafeRow_(custSheet, [ts, custName, custPhone, table, orderId]);
    }
  }

  return json({ ok: true, order_id: orderId });
}

// Orders only ever grows by doPost appending a row at the bottom -- today's orders are
// always the most recent rows, so list/today/cancel only need the TAIL of the sheet instead
// of its entire history via getDataRange(). Without this, every 15s admin.html poll and 2s
// print_bridge.py poll re-reads every order ever placed, getting slower forever as Orders
// grows. ORDERS_TAIL_ROWS is generous headroom over any single real business day's volume.
var ORDERS_TAIL_ROWS = 500;
var ORDERS_COLS = 7; // ts, table, items(json), total, status, bizToken, order_id -- doPost's row shape

// Generic tail-read -- any sheet that only grows by appending at the bottom should read its
// tail, not its full history, in a hot/polled path. Returns { rows, startRow } where startRow
// is the 1-indexed absolute sheet row rows[0] came from. Callers that write back to a row
// found by its index in `rows` (cancel, expenseVoid_, orderEdit_) MUST use startRow + i, not a
// hardcoded "+2" -- the tail window no longer always starts at row 2 once a sheet has grown
// past maxRows.
function readTail_(sheet, numCols, maxRows) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { rows: [], startRow: 2 }; // header only, no data rows yet
  var startRow = Math.max(2, lastRow - maxRows + 1);
  return { rows: sheet.getRange(startRow, 1, lastRow - startRow + 1, numCols).getValues(), startRow: startRow };
}
function readOrdersTail_(sheet) { return readTail_(sheet, ORDERS_COLS, ORDERS_TAIL_ROWS); }

function doGet(e) {
  var action = (e.parameter.action || 'today');

  // Menu actions don't touch the Orders sheet at all — handle first.
  if (action === 'menu') return menuList_();
  if (action === 'menu_set') return menuSet_(e);

  // Expenses + ReportQueue actions touch their own sheets, not the Orders tail read below
  // (report_data reads Orders separately, full-history, on purpose -- see reportData_).
  if (action === 'expense_add') return expenseAdd_(e);
  if (action === 'expense_void') return expenseVoid_(e);
  if (action === 'expense_list') return expenseList_(e);
  if (action === 'report_request') return reportRequestAdd_(e);
  if (action === 'report_pending') return reportPendingList_(e);
  if (action === 'report_done') return reportDone_(e);
  if (action === 'report_data') return reportData_(e);

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Orders');
  var now = new Date();
  var token = bizToken(now);

  // history spans an owner-chosen date range that can reach further back than the tail
  // window -- it's not in the 15s/2s hot path (owner-initiated, not polled), so it reads
  // full history on purpose. Handle it before the tail read the other actions share.
  if (action === 'history') return ordersHistory_(sh, e);

  var tail = readOrdersTail_(sh);
  var rows = tail.rows;

  if (action === 'today') {
    var count = 0, total = 0;
    rows.forEach(function (r) {
      if (String(r[5]).trim() === token && r[4] !== 'cancelled') { count++; total += Number(r[3]) || 0; }
    });
    return json({ date: bizDateHuman(now), orders: count, expected_cash: total, _v: 'v2' });
  }

  if (action === 'list') {
    if (e.parameter.key !== ADMIN_KEY) return json({ error: 'unauthorized' });
    var out = [];
    rows.forEach(function (r) {
      if (String(r[5]).trim() === token) {
        out.push({ id: r[6], time: r[0], table: r[1], items: r[2], total: Number(r[3]) || 0, status: r[4] });
      }
    });
    var total = out.filter(function (o) { return o.status !== 'cancelled'; })
      .reduce(function (s, o) { return s + o.total; }, 0);
    return json({ date: bizDateHuman(now), orders: out, expected_cash: total });
  }

  if (action === 'cancel') {
    if (e.parameter.key !== ADMIN_KEY) return json({ error: 'unauthorized' });
    var id = e.parameter.id;
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][6]).trim() === String(id).trim()) {
        sh.getRange(tail.startRow + i, 5).setValue('cancelled'); // column E = status
        return json({ ok: true });
      }
    }
    return json({ error: 'not_found' });
  }

  if (action === 'order_edit') return orderEdit_(sh, tail, e);

  return json({ error: 'unknown_action' });
}

// doGet?action=order_edit&id=&table=&items=<JSON>&total=&key=K -- fix a mistake (wrong
// item/qty/price/table) on an existing order without cancelling it and losing the record of
// what actually happened. Same tail-read + tail.startRow+i write-back pattern as `cancel`
// above, so it's correct even once Orders has grown past the tail window. Overwrites only
// columns B/C/D (table/items/total) at that absolute row -- ts/status/bizToken/order_id
// untouched. total is re-validated server-side independently rather than trusted blindly,
// even though the client always computes it as the sum of the edited line items.
function orderEdit_(sh, tail, e) {
  if (e.parameter.key !== ADMIN_KEY) return json({ error: 'unauthorized' });

  var table = String(e.parameter.table || '').slice(0, MAX_FIELD_LEN);
  if (!table) return json({ error: 'bad_table' });

  var rawItems;
  try { rawItems = JSON.parse(e.parameter.items); } catch (err) { return json({ error: 'bad_items' }); }
  if (!Array.isArray(rawItems) || !rawItems.length || rawItems.length > MAX_ITEMS_PER_ORDER) return json({ error: 'bad_items' });
  var items = sanitizeItems_(rawItems);

  var total = Number(e.parameter.total);
  if (isNaN(total) || total <= 0 || total > MAX_ORDER_TOTAL) return json({ error: 'bad_total' });

  var id = e.parameter.id;
  for (var i = 0; i < tail.rows.length; i++) {
    if (String(tail.rows[i][6]).trim() === String(id).trim()) {
      sh.getRange(tail.startRow + i, 2, 1, 3).setValues([[table, JSON.stringify(items), total]]); // B/C/D
      return json({ ok: true });
    }
  }
  return json({ error: 'not_found' });
}

// doGet?action=history&from=yyyy-MM-dd&to=yyyy-MM-dd&key=K -- order count + total per
// business day in [from, to] (inclusive), for the admin dashboard's history view. bizToken
// strings ('byyyyMMdd') compare correctly with plain < / > since they're always the same
// length and zero-padded.
function ordersHistory_(sh, e) {
  if (e.parameter.key !== ADMIN_KEY) return json({ error: 'unauthorized' });
  var from = String(e.parameter.from || '').trim();
  var to = String(e.parameter.to || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(from) || !/^\d{4}-\d{2}-\d{2}$/.test(to)) return json({ error: 'bad_range' });
  var fromToken = 'b' + from.replace(/-/g, '');
  var toToken = 'b' + to.replace(/-/g, '');

  var rows = sh.getDataRange().getValues().slice(1); // full history on purpose -- see doGet
  var byDay = {};
  rows.forEach(function (r) {
    var tok = String(r[5]).trim();
    if (tok < fromToken || tok > toToken) return;
    if (r[4] === 'cancelled') return;
    if (!byDay[tok]) byDay[tok] = { orders: 0, total: 0 };
    byDay[tok].orders++;
    byDay[tok].total += Number(r[3]) || 0;
  });
  var days = Object.keys(byDay).sort().map(function (tok) {
    return {
      date: tok.slice(1, 5) + '-' + tok.slice(5, 7) + '-' + tok.slice(7, 9),
      orders: byDay[tok].orders, total: byDay[tok].total,
    };
  });
  return json({ from: from, to: to, days: days });
}

// ---------- Menu management (owner self-serve: stock toggle, price/name edit, add item,
// per-item offers) ----------
// Menu tab columns: category, item_id, name_ar, variant, price_egp, active, sale_price_egp,
// on_offer. sale_price_egp/on_offer added 2026-10-10 for per-item discounts: on_offer is the
// owner's manual on/off toggle (same pattern as `active` for stock); sale_price_egp is the
// discount price she set once and keeps even while toggled off, so turning an offer back on
// later doesn't require re-entering the price.

function menuSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Menu');
}

function menuList_() {
  var rows = menuSheet_().getDataRange().getValues();
  rows.shift(); // drop header
  var out = rows
    .filter(function (r) { return String(r[1]).trim() !== ''; })
    .map(function (r) {
      return {
        category: r[0], item_id: r[1], name_ar: r[2], variant: r[3], price_egp: r[4], active: r[5],
        sale_price_egp: r[6], on_offer: r[7],
      };
    });
  return json({ rows: out });
}

function menuSet_(e) {
  if (e.parameter.key !== ADMIN_KEY) return json({ error: 'unauthorized' });

  var itemId = (e.parameter.item_id || '').trim().slice(0, MAX_FIELD_LEN);
  var category = (e.parameter.category || '').trim().slice(0, MAX_FIELD_LEN);
  var nameAr = (e.parameter.name_ar || '').trim().slice(0, MAX_FIELD_LEN);
  var variant = (e.parameter.variant || '').trim().slice(0, MAX_FIELD_LEN);
  var price = Number(e.parameter.price_egp);
  var active = e.parameter.active !== 'false'; // default true unless explicitly "false"

  if (!itemId || !category || !nameAr || isNaN(price)) return json({ error: 'missing_fields' });
  if (price < 0 || price > MAX_MENU_PRICE) return json({ error: 'bad_price' });

  var salePriceRaw = (e.parameter.sale_price_egp || '').trim();
  var salePrice = salePriceRaw === '' ? null : Number(salePriceRaw);
  if (salePrice != null && (isNaN(salePrice) || salePrice < 0 || salePrice > MAX_MENU_PRICE)) return json({ error: 'bad_price' });
  var onOffer = e.parameter.on_offer === 'true';
  if (onOffer && salePrice == null) onOffer = false; // can't be on-offer with no discount price configured

  var sh = menuSheet_();
  var rows = sh.getDataRange().getValues(); // includes header at index 0
  var newRow = [category, itemId, nameAr, variant, price, active ? 'TRUE' : 'FALSE',
    salePrice == null ? '' : salePrice, onOffer ? 'TRUE' : 'FALSE'];
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).trim() === itemId && String(rows[i][3]).trim() === variant) {
      sh.getRange(i + 1, 1, 1, 8).setValues([newRow]);
      return json({ ok: true, item_id: itemId });
    }
  }
  sh.appendRow(newRow);
  return json({ ok: true, item_id: itemId });
}

// ---------- Expenses (تقفيل: so the owner can see what she spent, not just what she made)
// ----------
// Expenses tab columns: ts, amount_egp, note, status, bizToken, expense_id. Same
// appendTextSafeRow_/tail-read/write-back-by-absolute-row patterns already proven for Orders.

var EXPENSES_COLS = 6;
var EXPENSES_TAIL_ROWS = 500;

function expensesSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Expenses');
}

function expenseAdd_(e) {
  if (e.parameter.key !== ADMIN_KEY) return json({ error: 'unauthorized' });
  var amount = Number(e.parameter.amount);
  if (isNaN(amount) || amount <= 0 || amount > MAX_EXPENSE_AMOUNT) return json({ error: 'bad_amount' });
  var note = String(e.parameter.note || '').trim().slice(0, MAX_FIELD_LEN);

  var sh = expensesSheet_();
  var now = new Date();
  var ts = Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm:ss');
  var id = Utilities.getUuid().slice(0, 8);
  appendTextSafeRow_(sh, [ts, amount, note, 'active', bizToken(now), id]);
  return json({ ok: true, expense_id: id });
}

function expenseVoid_(e) {
  if (e.parameter.key !== ADMIN_KEY) return json({ error: 'unauthorized' });
  var sh = expensesSheet_();
  var tail = readTail_(sh, EXPENSES_COLS, EXPENSES_TAIL_ROWS);
  var id = e.parameter.id;
  for (var i = 0; i < tail.rows.length; i++) {
    if (String(tail.rows[i][5]).trim() === String(id).trim()) {
      sh.getRange(tail.startRow + i, 4).setValue('voided'); // column D = status
      return json({ ok: true });
    }
  }
  return json({ error: 'not_found' });
}

// Optional &date=yyyy-MM-dd lets the admin dashboard's day-nav (◀ السابق / التالي ▶) browse
// any past day's expenses within the tail window, not just today -- defaults to today when
// omitted (unchanged behavior for any existing caller). expense_add still always logs to
// today regardless of which day is being viewed (see addExpense in admin.html).
function expenseList_(e) {
  if (e.parameter.key !== ADMIN_KEY) return json({ error: 'unauthorized' });
  var sh = expensesSheet_();
  var dateParam = String(e.parameter.date || '').trim();
  var token, dateHuman;
  if (/^\d{4}-\d{2}-\d{2}$/.test(dateParam)) {
    token = 'b' + dateParam.replace(/-/g, '');
    dateHuman = dateParam;
  } else {
    var now = new Date();
    token = bizToken(now);
    dateHuman = bizDateHuman(now);
  }
  var tail = readTail_(sh, EXPENSES_COLS, EXPENSES_TAIL_ROWS);
  var out = [];
  tail.rows.forEach(function (r) {
    if (String(r[4]).trim() === token) {
      out.push({ id: r[5], ts: r[0], amount: Number(r[1]) || 0, note: r[2], status: r[3] });
    }
  });
  var total = out.filter(function (x) { return x.status !== 'voided'; })
    .reduce(function (s, x) { return s + x.amount; }, 0);
  return json({ date: dateHuman, expenses: out, total: total });
}

// ---------- ReportQueue (تقفيل print bridge: owner taps a button, print_bridge.py prints a
// closing ticket on the real counter printer) ----------
// ReportQueue tab columns: ts, from_date, to_date, request_id. No bizToken -- ephemeral, not
// day-partitioned. Kept perpetually tiny by hard-deleting rows once printed (no audit value
// once done) rather than soft-marking status, so it never needs tail-read logic itself.

function reportQueueSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('ReportQueue');
}

function reportRequestAdd_(e) {
  if (e.parameter.key !== ADMIN_KEY) return json({ error: 'unauthorized' });
  var from = String(e.parameter.from || '').trim();
  var to = String(e.parameter.to || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(from) || !/^\d{4}-\d{2}-\d{2}$/.test(to)) return json({ error: 'bad_range' });

  var sh = reportQueueSheet_();
  var now = new Date();
  var ts = Utilities.formatDate(now, TZ, 'yyyy-MM-dd HH:mm:ss');
  var id = Utilities.getUuid().slice(0, 8);
  appendTextSafeRow_(sh, [ts, from, to, id], [1, 2, 3]); // from/to (cols 2-3) are plain YYYY-MM-DD -- see appendTextSafeRow_
  return json({ ok: true, request_id: id });
}

function reportPendingList_(e) {
  if (e.parameter.key !== ADMIN_KEY) return json({ error: 'unauthorized' });
  var rows = reportQueueSheet_().getDataRange().getValues();
  rows.shift(); // drop header
  var out = rows
    .filter(function (r) { return String(r[3]).trim() !== ''; })
    .map(function (r) { return { ts: r[0], from: r[1], to: r[2], request_id: r[3] }; });
  return json({ requests: out });
}

function reportDone_(e) {
  if (e.parameter.key !== ADMIN_KEY) return json({ error: 'unauthorized' });
  var sh = reportQueueSheet_();
  var rows = sh.getDataRange().getValues();
  var id = e.parameter.id;
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][3]).trim() === String(id).trim()) {
      sh.deleteRow(i + 1);
      return json({ ok: true });
    }
  }
  return json({ error: 'not_found' });
}

// doGet?action=report_data&from=&to=&key=K -- sales+expenses+net aggregate for a تقفيل
// ticket's content. Full-history scan of both Orders (sales, excluding cancelled) and
// Expenses (excluding voided) within [fromToken, toToken] -- same string-comparison trick
// already used by ordersHistory_. Full read is justified the same way: button/queue-
// triggered, not polled.
function reportData_(e) {
  if (e.parameter.key !== ADMIN_KEY) return json({ error: 'unauthorized' });
  var from = String(e.parameter.from || '').trim();
  var to = String(e.parameter.to || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(from) || !/^\d{4}-\d{2}-\d{2}$/.test(to)) return json({ error: 'bad_range' });
  var fromToken = 'b' + from.replace(/-/g, '');
  var toToken = 'b' + to.replace(/-/g, '');

  var orderRows = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Orders').getDataRange().getValues().slice(1);
  var salesTotal = 0, salesCount = 0;
  var itemAgg = {}; // name -> { qty, total } -- "what did we actually sell" breakdown for the ticket
  orderRows.forEach(function (r) {
    var tok = String(r[5]).trim();
    if (tok < fromToken || tok > toToken) return;
    if (r[4] === 'cancelled') return;
    salesCount++;
    salesTotal += Number(r[3]) || 0;
    var items;
    try { items = JSON.parse(r[2]); } catch (e2) { items = null; } // legacy pilot-era orders stored items as a display string, not JSON -- can't itemize those, but they still count toward sales_total/count above
    if (Array.isArray(items)) {
      items.forEach(function (it) {
        var name = String((it && it.name) || '').trim();
        if (!name) return;
        var qty = Number(it.qty) || 0;
        var lineTotal = qty * (Number(it.price) || 0);
        if (!itemAgg[name]) itemAgg[name] = { qty: 0, total: 0 };
        itemAgg[name].qty += qty;
        itemAgg[name].total += lineTotal;
      });
    }
  });
  var itemsSold = Object.keys(itemAgg).map(function (name) {
    return { name: name, qty: itemAgg[name].qty, total: itemAgg[name].total };
  }).sort(function (a, b) { return b.total - a.total; }); // biggest earners first

  var expenseRows = expensesSheet_().getDataRange().getValues().slice(1);
  var expensesTotal = 0, expensesCount = 0;
  expenseRows.forEach(function (r) {
    var tok = String(r[4]).trim();
    if (tok < fromToken || tok > toToken) return;
    if (r[3] === 'voided') return;
    expensesCount++;
    expensesTotal += Number(r[1]) || 0;
  });

  return json({
    from: from, to: to,
    sales_total: salesTotal, sales_count: salesCount,
    items_sold: itemsSold,
    expenses_total: expensesTotal, expenses_count: expensesCount,
    net: salesTotal - expensesTotal,
  });
}
