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


def format_time_local(raw):
    """Sheet cells sometimes come back as ISO-UTC ('...T..Z', Sheets auto-typed
    the cell as a Date) and sometimes as the plain local string we wrote
    ('YYYY-MM-DD HH:MM:SS') — handle both, always show Cairo local HH:MM."""
    try:
        if raw.endswith("Z"):
            dt = datetime.strptime(raw, "%Y-%m-%dT%H:%M:%S.%fZ") + CAIRO_OFFSET
        else:
            dt = datetime.strptime(raw[:19], "%Y-%m-%d %H:%M:%S")
        return dt.strftime("%H:%M")
    except Exception:
        return str(raw)


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


def draw_right(draw, right_x, y, text, ar_font, latin_font, fill=0):
    x = right_x - line_width(draw, text, ar_font, latin_font)
    for is_a, s in segment_runs(text):
        f = ar_font if is_a else latin_font
        draw.text((x, y), s, font=f, fill=fill)
        x += draw.textbbox((0, 0), s, font=f)[2]


def draw_center(draw, center_x, y, text, font, fill=0):
    w = draw.textbbox((0, 0), text, font=font)[2]
    draw.text((center_x - w / 2, y), text, font=font, fill=fill)


def build_receipt_image(order):
    pad = 20
    ar_header = ImageFont.truetype(ARABIC_FONT_PATH, 46)
    latin_header = ImageFont.truetype(LATIN_FONT_PATH, 46)
    ar_bold = ImageFont.truetype(ARABIC_FONT_PATH, 30)
    latin_bold = ImageFont.truetype(LATIN_FONT_PATH, 30)
    ar_reg = ImageFont.truetype(ARABIC_FONT_PATH, 28)
    latin_reg = ImageFont.truetype(LATIN_FONT_PATH, 28)

    table_line = ar(f"ترابيزة {order['table']}")
    time_str = format_time_local(order["time"])
    item_lines = [ar(x) for x in str(order["items"]).split("، ") if x]
    total_line = ar(f"الإجمالي: {order['total']} ج")

    line_h_reg, line_h_bold, line_h_header = 40, 44, 60
    divider_h = 26
    height = (
        pad + line_h_header + divider_h
        + line_h_bold + line_h_reg + divider_h
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
    draw_right(d, RECEIPT_WIDTH - pad, y, time_str, ar_reg, latin_reg)
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


def print_order(order):
    img = build_receipt_image(order)
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


def main():
    if not ADMIN_WEBHOOK:
        print("ADMIN_WEBHOOK is empty — paste the apps_script_v2.gs /exec URL first.")
        sys.exit(1)

    printed = load_printed_ids()
    print(f"Printer bridge running -> '{PRINTER_NAME}'. Watching for new orders every {POLL_SECONDS}s...")

    while True:
        try:
            orders = fetch_orders()
            new_orders = [o for o in orders if o["id"] not in printed and o["status"] != "cancelled"]
            for order in new_orders:
                print(f"Printing order {order['id']} — table {order['table']}")
                print_order(order)
                printed.add(order["id"])
            if new_orders:
                save_printed_ids(printed)
        except Exception as e:
            print(f"[warn] {e} — retrying next cycle")
        time.sleep(POLL_SECONDS)


if __name__ == "__main__":
    main()
