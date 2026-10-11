"""
Waqti cafe — printer bridge (Squeeze).

Confirmed hardware (checked on-site 2026-10-04/05):
  Printer: Xprinter XP-Q808K, ESC/POS, USB+LAN (currently wired via USB)
  Windows print queue name: "cashier" (NOT "Printer POS-80" — that was just the
    USB device's hardware description in Device Manager, not the queue name)
    -> we print through the "cashier" queue (raw ESC/POS bytes via the Windows
       spooler), not by opening the USB device directly. No driver swap needed.

ARABIC: ESC/POS printers have no built-in Arabic character set/shaping engine —
text-mode printing renders Arabic as garbled disconnected glyphs no matter what
encoding you declare. Fix: render the whole receipt as a bitmap image instead
(printers can always print pixels). Within that image, Arabic letters and
digits/symbols are drawn with SEPARATE fonts and composited run-by-run — tested
locally and confirmed no single common font has correct glyphs for both shaped
Arabic presentation forms AND plain digits at once.

Runs on their PC. Polls the same Apps Script endpoint the admin dashboard uses
(?action=list) for new orders and prints each one on the Xprinter the moment
it shows up. This only prints — it never touches their existing POS software.

Each order prints TWO copies (2026-10-10): a cashier copy (price per line + grand
total, for cash reconciliation) and a kitchen copy right after it (same items, NO
prices anywhere -- what the barista actually works from).
"""

import json
import os
import sys
import time
import urllib.request
from datetime import datetime, timedelta

import win32print
from escpos.printer import Dummy
from PIL import Image, ImageDraw, ImageFont
import arabic_reshaper
from bidi.algorithm import get_display

# ---------------- Config ----------------
PRINTER_NAME = "cashier"   # exact Windows printer queue name (confirmed via win32print.EnumPrinters, 2026-10-05)

ADMIN_WEBHOOK = "https://script.google.com/macros/s/AKfycbxh3-0tNyirbMBFXuIfpcx_NJkvh0bLBavgcJoGHvEQ2TvA6R_t-bMa-vYn-v0lhoQTgQ/exec"
ADMIN_KEY = "69e947ed5e8b6b8d"
POLL_SECONDS = 2

ARABIC_FONT_PATH = r"C:\Windows\Fonts\tahoma.ttf"   # Arabic letters (shaped presentation forms)
LATIN_FONT_PATH = r"C:\Windows\Fonts\arial.ttf"     # digits, ×, :, ج, timestamps — ships on every Windows PC

HERE = os.path.dirname(os.path.abspath(__file__))
STATE_PATH = os.path.join(HERE, ".printed_order_ids.json")
CAFE_NAME = "Squeeze"
CAIRO_OFFSET = timedelta(hours=2)   # Africa/Cairo, no DST currently observed
RECEIPT_WIDTH = 576                 # px, standard 80mm thermal printer raster width @ 203dpi


def load_printed_ids():
    if os.path.exists(STATE_PATH):
        with open(STATE_PATH, encoding="utf-8") as f:
            return set(json.load(f))
    return set()


def save_printed_ids(ids):
    with open(STATE_PATH, "w", encoding="utf-8") as f:
        json.dump(sorted(ids), f)


def fetch_orders():
    url = f"{ADMIN_WEBHOOK}?action=list&key={ADMIN_KEY}"
    with urllib.request.urlopen(url, timeout=10) as r:
        data = json.loads(r.read().decode("utf-8"))
    if "error" in data:
        raise RuntimeError(data["error"])
    return data.get("orders", [])


def fetch_pending_reports():
    url = f"{ADMIN_WEBHOOK}?action=report_pending&key={ADMIN_KEY}"
    with urllib.request.urlopen(url, timeout=10) as r:
        data = json.loads(r.read().decode("utf-8"))
    if "error" in data:
        raise RuntimeError(data["error"])
    return data.get("requests", [])


def fetch_report_data(from_date, to_date):
    url = f"{ADMIN_WEBHOOK}?action=report_data&from={from_date}&to={to_date}&key={ADMIN_KEY}"
    with urllib.request.urlopen(url, timeout=10) as r:
        data = json.loads(r.read().decode("utf-8"))
    if "error" in data:
        raise RuntimeError(data["error"])
    return data


def mark_report_done(request_id):
    url = f"{ADMIN_WEBHOOK}?action=report_done&id={request_id}&key={ADMIN_KEY}"
    with urllib.request.urlopen(url, timeout=10) as r:
        data = json.loads(r.read().decode("utf-8"))
    if "error" in data:
        raise RuntimeError(data["error"])


def _local_datetime(raw):
    """Sheet cells sometimes come back as ISO-UTC ('...T..Z', Sheets auto-typed
    the cell as a Date) and sometimes as the plain local string we wrote
    ('YYYY-MM-DD HH:MM:SS') — handle both, always return Cairo local time."""
    if raw.endswith("Z"):
        return datetime.strptime(raw, "%Y-%m-%dT%H:%M:%S.%fZ") + CAIRO_OFFSET
    return datetime.strptime(raw[:19], "%Y-%m-%d %H:%M:%S")


def format_time_local(raw):
    try:
        return _local_datetime(raw).strftime("%H:%M")
    except Exception:
        return str(raw)


def format_date_local(raw):
    try:
        return _local_datetime(raw).strftime("%d/%m/%Y")
    except Exception:
        return ""


def ar(text):
    """Reshape + bidi-reorder Arabic into final visual draw order."""
    return get_display(arabic_reshaper.reshape(text))


def is_arabic_char(ch):
    o = ord(ch)
    return (0x0600 <= o <= 0x06FF) or (0xFB50 <= o <= 0xFDFF) or (0xFE70 <= o <= 0xFEFF)


def segment_runs(text):
    """Split already-reordered text into consecutive (is_arabic, substring) runs
    so each run can be drawn with the font that actually has its glyphs."""
    runs, cur_type, cur = [], None, ""
    for ch in text:
        t = is_arabic_char(ch)
        if cur_type is None:
            cur_type, cur = t, ch
        elif t == cur_type:
            cur += ch
        else:
            runs.append((cur_type, cur))
            cur_type, cur = t, ch
    if cur:
        runs.append((cur_type, cur))
    return runs


def line_width(draw, text, ar_font, latin_font):
    return sum(
        draw.textbbox((0, 0), s, font=(ar_font if is_a else latin_font))[2]
        for is_a, s in segment_runs(text)
    )


def fit_arabic_line(draw, raw_text, max_width, ar_font, latin_font):
    """Reshape+bidi raw_text, trimming from its logical end (append '…') until it fits
    max_width. Trims the RAW string (not the already-reordered one) so the cut lands on
    the sentence's actual end -- trimming a bidi-reordered RTL string from its storage-end
    would chip away at the sentence's logical START instead. Without this, a long customer
    note (the note placeholder itself is ~35 chars) can draw past the 80mm receipt's edge
    and get clipped/overlapped on the printed ticket the kitchen works from."""
    shaped = ar(raw_text)
    if line_width(draw, shaped, ar_font, latin_font) <= max_width:
        return shaped
    text = raw_text
    while text:
        text = text[:-1]
        shaped = ar(text + "…")
        if line_width(draw, shaped, ar_font, latin_font) <= max_width:
            return shaped
    return "…"


def draw_right(draw, right_x, y, text, ar_font, latin_font, fill=0):
    x = right_x - line_width(draw, text, ar_font, latin_font)
    for is_a, s in segment_runs(text):
        f = ar_font if is_a else latin_font
        draw.text((x, y), s, font=f, fill=fill)
        x += draw.textbbox((0, 0), s, font=f)[2]


def draw_center(draw, center_x, y, text, font, fill=0):
    w = draw.textbbox((0, 0), text, font=font)[2]
    draw.text((center_x - w / 2, y), text, font=font, fill=fill)


def parse_items(raw):
    """Orders since the 2026-10-10 receipts overhaul store `items` as a JSON array
    (name/qty/price per line) instead of a flattened display string -- lets the cashier
    copy show a price per line and the kitchen copy show no prices, from the same data.
    Older orders (pre-dating this change, possibly still sitting in a real Sheet from the
    pilot period) are the old flattened string -- fall back to one no-price line per
    '، '-joined fragment rather than crashing on json.loads."""
    try:
        parsed = json.loads(raw)
        if isinstance(parsed, list):
            return parsed
    except (ValueError, TypeError):
        pass
    return [{"name": x, "qty": None, "price": None} for x in str(raw).split("، ") if x]


def _item_line_text(item, with_price):
    name = str(item.get("name", ""))
    qty = item.get("qty")
    price = item.get("price")
    prefix = f"{qty}× " if qty is not None else ""
    if with_price and qty is not None and price is not None:
        return f"{prefix}{name} — {qty * price} ج"
    return f"{prefix}{name}"


def build_receipt_image(order, order_number):
    """Cashier copy: price per line + the grand total -- what she reconciles against the
    drawer at end of day. order_number is this order's 1-based position in TODAY's order
    list (resets every business day) -- a short, speakable reference ("order 7"), not the
    internal random order id."""
    pad = 20
    ar_header = ImageFont.truetype(ARABIC_FONT_PATH, 46)
    latin_header = ImageFont.truetype(LATIN_FONT_PATH, 46)
    ar_bold = ImageFont.truetype(ARABIC_FONT_PATH, 30)
    latin_bold = ImageFont.truetype(LATIN_FONT_PATH, 30)
    ar_reg = ImageFont.truetype(ARABIC_FONT_PATH, 28)
    latin_reg = ImageFont.truetype(LATIN_FONT_PATH, 28)

    # Measure/fit against a throwaway 1x1 canvas -- text measurement doesn't need the real
    # image, and the real image's height depends only on line COUNT (fixed below), never
    # on measured width, so this has no effect on the layout math that follows.
    measure = ImageDraw.Draw(Image.new("L", (1, 1)))
    max_text_width = RECEIPT_WIDTH - 2 * pad

    table_line = fit_arabic_line(measure, f"ترابيزة {order['table']}", max_text_width, ar_bold, latin_bold)
    datetime_str = f"{format_date_local(order['time'])} {format_time_local(order['time'])}"
    order_no_line = ar(f"طلب رقم {order_number}")
    items = parse_items(order["items"])
    item_lines = [
        fit_arabic_line(measure, _item_line_text(it, True), max_text_width, ar_reg, latin_reg)
        for it in items
    ]
    total_line = fit_arabic_line(measure, f"الإجمالي: {order['total']} ج", max_text_width, ar_header, latin_header)

    line_h_reg, line_h_bold, line_h_header = 40, 44, 60
    divider_h = 26
    height = (
        pad + line_h_header + divider_h
        + line_h_bold + line_h_reg + line_h_reg + divider_h
        + len(item_lines) * line_h_reg + divider_h
        + line_h_header + pad
    )

    img = Image.new("L", (RECEIPT_WIDTH, height), 255)
    d = ImageDraw.Draw(img)
    y = pad

    draw_center(d, RECEIPT_WIDTH / 2, y, CAFE_NAME, latin_header)
    y += line_h_header
    d.line([(pad, y + divider_h / 2), (RECEIPT_WIDTH - pad, y + divider_h / 2)], fill=0, width=2)
    y += divider_h

    draw_right(d, RECEIPT_WIDTH - pad, y, table_line, ar_bold, latin_bold)
    y += line_h_bold
    draw_right(d, RECEIPT_WIDTH - pad, y, datetime_str, ar_reg, latin_reg)
    y += line_h_reg
    draw_right(d, RECEIPT_WIDTH - pad, y, order_no_line, ar_reg, latin_reg)
    y += line_h_reg
    d.line([(pad, y + divider_h / 2), (RECEIPT_WIDTH - pad, y + divider_h / 2)], fill=0, width=2)
    y += divider_h

    for line in item_lines:
        draw_right(d, RECEIPT_WIDTH - pad, y, line, ar_reg, latin_reg)
        y += line_h_reg
    d.line([(pad, y + divider_h / 2), (RECEIPT_WIDTH - pad, y + divider_h / 2)], fill=0, width=2)
    y += divider_h

    draw_right(d, RECEIPT_WIDTH - pad, y, total_line, ar_header, latin_header)

    return img


def build_kitchen_receipt_image(order, order_number):
    """Kitchen copy: same items, NO prices anywhere -- what the barista actually works
    from. Printed right after the cashier copy, same printer, every order."""
    pad = 20
    ar_header = ImageFont.truetype(ARABIC_FONT_PATH, 46)
    latin_header = ImageFont.truetype(LATIN_FONT_PATH, 46)
    ar_bold = ImageFont.truetype(ARABIC_FONT_PATH, 30)
    latin_bold = ImageFont.truetype(LATIN_FONT_PATH, 30)
    ar_reg = ImageFont.truetype(ARABIC_FONT_PATH, 28)
    latin_reg = ImageFont.truetype(LATIN_FONT_PATH, 28)

    measure = ImageDraw.Draw(Image.new("L", (1, 1)))
    max_text_width = RECEIPT_WIDTH - 2 * pad

    kitchen_label = fit_arabic_line(measure, "نسخة المطبخ", max_text_width, ar_bold, latin_bold)
    table_line = fit_arabic_line(measure, f"ترابيزة {order['table']}", max_text_width, ar_bold, latin_bold)
    datetime_str = f"{format_date_local(order['time'])} {format_time_local(order['time'])}"
    order_no_line = ar(f"طلب رقم {order_number}")
    items = parse_items(order["items"])
    item_lines = [
        fit_arabic_line(measure, _item_line_text(it, False), max_text_width, ar_reg, latin_reg)
        for it in items
    ]

    line_h_reg, line_h_bold, line_h_header = 40, 44, 60
    divider_h = 26
    height = (
        pad + line_h_header + divider_h
        + line_h_bold + divider_h
        + line_h_bold + line_h_reg + line_h_reg + divider_h
        + len(item_lines) * line_h_reg + pad
    )

    img = Image.new("L", (RECEIPT_WIDTH, height), 255)
    d = ImageDraw.Draw(img)
    y = pad

    draw_center(d, RECEIPT_WIDTH / 2, y, CAFE_NAME, latin_header)
    y += line_h_header
    d.line([(pad, y + divider_h / 2), (RECEIPT_WIDTH - pad, y + divider_h / 2)], fill=0, width=2)
    y += divider_h

    draw_right(d, RECEIPT_WIDTH - pad, y, kitchen_label, ar_bold, latin_bold)
    y += line_h_bold
    d.line([(pad, y + divider_h / 2), (RECEIPT_WIDTH - pad, y + divider_h / 2)], fill=0, width=2)
    y += divider_h

    draw_right(d, RECEIPT_WIDTH - pad, y, table_line, ar_bold, latin_bold)
    y += line_h_bold
    draw_right(d, RECEIPT_WIDTH - pad, y, datetime_str, ar_reg, latin_reg)
    y += line_h_reg
    draw_right(d, RECEIPT_WIDTH - pad, y, order_no_line, ar_reg, latin_reg)
    y += line_h_reg
    d.line([(pad, y + divider_h / 2), (RECEIPT_WIDTH - pad, y + divider_h / 2)], fill=0, width=2)
    y += divider_h

    for line in item_lines:
        draw_right(d, RECEIPT_WIDTH - pad, y, line, ar_reg, latin_reg)
        y += line_h_reg

    return img


def _send_to_printer(img):
    p = Dummy()
    p.image(img)
    p._raw(b"\n\n")
    p.cut()
    data = p.output

    hprinter = win32print.OpenPrinter(PRINTER_NAME)
    try:
        win32print.StartDocPrinter(hprinter, 1, ("Waqti order", None, "RAW"))
        win32print.StartPagePrinter(hprinter)
        win32print.WritePrinter(hprinter, data)
        win32print.EndPagePrinter(hprinter)
        win32print.EndDocPrinter(hprinter)
    finally:
        win32print.ClosePrinter(hprinter)


def print_order(order, order_number):
    _send_to_printer(build_receipt_image(order, order_number))
    _send_to_printer(build_kitchen_receipt_image(order, order_number))


def format_date_range_ar(from_date, to_date):
    """Compose a right-to-left-readable 'from X to Y' line for two raw ISO dates, WITHOUT
    routing the combined string through ar()/get_display() as one string -- python-bidi
    mishandles two separate hyphenated-date runs inside one RTL-majority line, silently
    reversing the SECOND date's digit groups (e.g. "2026-10-11" -> "11-10-2026") while
    leaving the first alone. Confirmed via direct repro, not hypothetical. Only "إلى" (a lone
    Arabic word, no ambiguity) goes through ar(); the two raw dates are placed around it in
    final left-to-right PIXEL order directly -- to_date first (leftmost = read last), then
    إلى, then from_date (rightmost = read first), so an RTL reader sees "from_date إلى
    to_date"."""
    return f"{to_date} {ar('إلى')} {from_date}"


def build_report_image(report_data):
    """تقفيل (closing) ticket: what sold (qty + line total per item, biggest earner first),
    then sales+expenses+net. Same fonts/RECEIPT_WIDTH/pad/line_h_*/divider_h constants and
    fit_arabic_line/draw_right/draw_center helpers as build_receipt_image, so it visually
    belongs to the same system. The itemized section is OPTIONAL (items_sold missing/empty --
    e.g. a day with zero orders, or an older caller) -- falls back to exactly the original
    fixed height with no section/divider added, so this never changes anything for existing
    callers."""
    pad = 20
    ar_header = ImageFont.truetype(ARABIC_FONT_PATH, 46)
    latin_header = ImageFont.truetype(LATIN_FONT_PATH, 46)
    ar_bold = ImageFont.truetype(ARABIC_FONT_PATH, 30)
    latin_bold = ImageFont.truetype(LATIN_FONT_PATH, 30)
    ar_reg = ImageFont.truetype(ARABIC_FONT_PATH, 28)
    latin_reg = ImageFont.truetype(LATIN_FONT_PATH, 28)

    measure = ImageDraw.Draw(Image.new("L", (1, 1)))
    max_text_width = RECEIPT_WIDTH - 2 * pad

    title_line = fit_arabic_line(measure, "تقفيل", max_text_width, ar_bold, latin_bold)
    # Not run through fit_arabic_line (see format_date_range_ar for why) -- always short/
    # fixed-format enough to fit the receipt width without truncation.
    range_line = format_date_range_ar(report_data['from'], report_data['to'])
    item_lines = [
        fit_arabic_line(measure, f"{it['name']} × {it['qty']} = {it['total']} ج", max_text_width, ar_reg, latin_reg)
        for it in (report_data.get('items_sold') or [])
    ]
    sales_line = fit_arabic_line(
        measure, f"المبيعات: {report_data['sales_count']} طلب — {report_data['sales_total']} ج", max_text_width, ar_reg, latin_reg
    )
    expenses_line = fit_arabic_line(
        measure, f"المصروفات: {report_data['expenses_count']} — {report_data['expenses_total']} ج", max_text_width, ar_reg, latin_reg
    )
    net_line = fit_arabic_line(measure, f"الصافي: {report_data['net']} ج", max_text_width, ar_header, latin_header)

    line_h_reg, line_h_bold, line_h_header = 40, 44, 60
    divider_h = 26
    items_block_h = (len(item_lines) * line_h_reg + divider_h) if item_lines else 0
    height = (
        pad + line_h_header + divider_h
        + line_h_bold + line_h_reg + divider_h
        + items_block_h
        + line_h_reg + line_h_reg + divider_h
        + line_h_header + pad
    )

    img = Image.new("L", (RECEIPT_WIDTH, height), 255)
    d = ImageDraw.Draw(img)
    y = pad

    draw_center(d, RECEIPT_WIDTH / 2, y, CAFE_NAME, latin_header)
    y += line_h_header
    d.line([(pad, y + divider_h / 2), (RECEIPT_WIDTH - pad, y + divider_h / 2)], fill=0, width=2)
    y += divider_h

    draw_right(d, RECEIPT_WIDTH - pad, y, title_line, ar_bold, latin_bold)
    y += line_h_bold
    draw_right(d, RECEIPT_WIDTH - pad, y, range_line, ar_reg, latin_reg)
    y += line_h_reg
    d.line([(pad, y + divider_h / 2), (RECEIPT_WIDTH - pad, y + divider_h / 2)], fill=0, width=2)
    y += divider_h

    if item_lines:
        for line in item_lines:
            draw_right(d, RECEIPT_WIDTH - pad, y, line, ar_reg, latin_reg)
            y += line_h_reg
        d.line([(pad, y + divider_h / 2), (RECEIPT_WIDTH - pad, y + divider_h / 2)], fill=0, width=2)
        y += divider_h

    draw_right(d, RECEIPT_WIDTH - pad, y, sales_line, ar_reg, latin_reg)
    y += line_h_reg
    draw_right(d, RECEIPT_WIDTH - pad, y, expenses_line, ar_reg, latin_reg)
    y += line_h_reg
    d.line([(pad, y + divider_h / 2), (RECEIPT_WIDTH - pad, y + divider_h / 2)], fill=0, width=2)
    y += divider_h

    draw_right(d, RECEIPT_WIDTH - pad, y, net_line, ar_header, latin_header)

    return img


def print_report(report_data):
    _send_to_printer(build_report_image(report_data))


def check_report_queue():
    """Called every poll cycle, same 2s cadence as orders -- ReportQueue stays near-empty by
    construction (hard-delete on done), so the cost is a trivial read on a tiny sheet,
    negligible next to the 500-row Orders tail-read that already happens every cycle. A
    failed request is left pending (not marked done) so it retries next cycle instead of
    silently vanishing from the queue."""
    for req in fetch_pending_reports():
        try:
            print_report(fetch_report_data(req["from"], req["to"]))
            mark_report_done(req["request_id"])
        except Exception as e:
            print(f"[warn] report {req['request_id']} failed: {e} -- retries next cycle")


def main():
    if not ADMIN_WEBHOOK:
        print("ADMIN_WEBHOOK is empty — paste the apps_script_v2.gs /exec URL first.")
        sys.exit(1)

    printed = load_printed_ids()
    print(f"Printer bridge running -> '{PRINTER_NAME}'. Watching for new orders every {POLL_SECONDS}s...")

    while True:
        try:
            orders = fetch_orders()
            # action=list only ever returns TODAY's orders, so any tracked id no longer in
            # it belongs to a business day that's gone for good -- drop it. Self-cleans at
            # every day boundary instead of growing this file forever.
            before = len(printed)
            printed &= {o["id"] for o in orders}
            pruned = len(printed) != before

            # order_number = this order's 1-based position in TODAY's full list (action=list
            # is always insertion-order for the business day) -- a stable, speakable "order
            # N" that resets every business day, independent of which orders get printed in
            # which batch or cancelled along the way.
            new_orders = [
                (i, o) for i, o in enumerate(orders, start=1)
                if o["id"] not in printed and o["status"] != "cancelled"
            ]
            for order_number, order in new_orders:
                print(f"Printing order {order['id']} — table {order['table']} (#{order_number})")
                print_order(order, order_number)
                printed.add(order["id"])
            if new_orders or pruned:
                save_printed_ids(printed)

            check_report_queue()
        except Exception as e:
            print(f"[warn] {e} — retrying next cycle")
        time.sleep(POLL_SECONDS)


if __name__ == "__main__":
    main()
