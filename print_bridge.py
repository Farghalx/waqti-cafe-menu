"""
Waqti cafe — printer bridge (Squeeze).

Confirmed hardware (checked on-site 2026-10-04):
  Printer: Xprinter XP-Q808K, ESC/POS, USB+LAN (currently wired via USB)
  USB ID : VID_1FC9&PID_2016
  Windows already has it installed as a working printer queue named "Printer POS-80"
    -> we print through THAT queue (raw ESC/POS bytes via the Windows spooler),
       not by opening the USB device directly. No driver swap needed (no Zadig/
       libusb), and it can't conflict with their existing POS software also
       printing to the same queue.

Runs on their PC. Polls the same Apps Script endpoint the admin dashboard uses
(?action=list) for new orders and prints each one on the Xprinter the moment
it shows up.

SETUP (one-time, on their PC):
  1. pip install python-escpos pywin32
  2. Confirm the exact printer name: Settings -> Bluetooth & devices ->
     Printers & scanners. If it's not literally "Printer POS-80", update
     PRINTER_NAME below to match exactly (case-sensitive).
  3. Set ADMIN_WEBHOOK (the apps_script_v2.gs /exec URL once deployed) + ADMIN_KEY
     (must match the ADMIN_KEY constant set inside that script).
  4. Run: python print_bridge.py
     -> it should print nothing yet (no orders), just sit there polling.
  5. Make it start automatically: Task Scheduler -> Create Task -> trigger
     "At log on" -> action "python print_bridge.py" -> keeps it running
     whenever the PC is on, same way the kitchen phone just stays open on Telegram.

This only prints — it never touches their existing cafe-management software.
"""

import json
import os
import sys
import time
import urllib.request

import win32print
from escpos.printer import Dummy
import arabic_reshaper
from bidi.algorithm import get_display

# ---------------- Config ----------------
PRINTER_NAME = "Printer POS-80"   # exact Windows printer queue name — verify in step 2 above

ADMIN_WEBHOOK = "https://script.google.com/macros/s/AKfycbxh3-0tNyirbMBFXuIfpcx_NJkvh0bLBavgcJoGHvEQ2TvA6R_t-bMa-vYn-v0lhoQTgQ/exec"
ADMIN_KEY = "69e947ed5e8b6b8d"
POLL_SECONDS = 6

HERE = os.path.dirname(os.path.abspath(__file__))
STATE_PATH = os.path.join(HERE, ".printed_order_ids.json")
CAFE_NAME = "Squeeze"


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


def ar(text):
    """Thermal printers have no text-shaping engine — unlike WeasyPrint (used for
    the PDF proposals/reports), raw ESC/POS just prints isolated glyph bytes, so
    Arabic prints as disconnected letters unless reshaped first. Same fix already
    used in leads-finder/core/execution/pdf_renderer.py."""
    return get_display(arabic_reshaper.reshape(text))


def build_receipt_bytes(order):
    """Use python-escpos's Dummy printer just to build the ESC/POS byte sequence
    (text formatting, cut command, etc.) without opening any USB/network connection
    itself — we hand those bytes to Windows' own spooler instead."""
    p = Dummy()
    p.set(align="center", bold=True, width=2, height=2)
    p.text(f"{CAFE_NAME}\n")
    p.set(align="center", bold=False, width=1, height=1)
    p.text("-" * 32 + "\n")
    p.set(align="right", bold=True)
    p.text(ar(f"ترابيزة {order['table']}") + "\n")
    p.set(bold=False)
    p.text(f"{order['time']}\n")
    p.text("-" * 32 + "\n")
    p.set(align="right")
    for line in str(order["items"]).split("، "):
        p.text(ar(line) + "\n")
    p.text("-" * 32 + "\n")
    p.set(bold=True, width=2, height=2)
    p.text(ar(f"الإجمالي: {order['total']} ج") + "\n")
    p.set(bold=False, width=1, height=1)
    p.text("\n")
    p.cut()
    return p.output


def print_order(order):
    data = build_receipt_bytes(order)
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
