#!/usr/bin/env python3
"""
generate_bundle_tickets.py
------------------------------------------------------------------
Quick batch ticket generator for the Roaring Raffle Draw.

WHAT IT DOES
  1. Reads a list of names from a .md file (one name per line).
  2. For each name, generates N tickets by calling your existing
     Google Apps Script backend (same WEBAPP_URL / API_SECRET that
     generate-ticket.html uses) - so tickets land in the real
     GenerateTicket sheet, same as filling the form by hand:
        Status:            Verified (auto-verified, same as the form)
        Payment Method:    Cash
        Transaction Number: 1234
  3. Writes ONE plain-text (.txt) file per person with their ticket IDs +
     QR links, laid out one ticket at a time so it's easy to read and
     copy - e.g.  tickets/Wolverine_20_Tickets.txt

HOW MANY TICKETS PER PERSON
  Type the number when you run it:
      python3 generate_bundle_tickets.py quick_gen.md 20
  -> every name in quick_gen.md gets 20 tickets.

  Want a specific person to get a different amount? Just add it after
  their name in the .md file, and it overrides the number above:
      wolverine: 5
      cyclops
      jane doe: 30

BEFORE YOU RUN IT FOR REAL
  Edit the CONFIG block below and fill in TICKET_SITE_BASE_URL - the
  address where your view-ticket.html actually lives. Without it the
  ticket links saved to the .md files will be wrong.

  Test safely first with --dry-run (fakes ticket IDs, does NOT touch
  your real Google Sheet):
      python3 generate_bundle_tickets.py quick_gen.md 20 --dry-run

REQUIREMENTS
  Python 3 only - no installs needed (uses the standard library).
------------------------------------------------------------------
"""

import sys
import os
import re
import json
import time
import urllib.request
import urllib.error
from datetime import datetime

# ===================== CONFIG - edit these =====================

# Copied from your generate-ticket.html - already correct.
WEBAPP_URL = "https://script.google.com/macros/s/AKfycbxrJUdm_fgywvzr8v1aWjbiQMGR2K94WyesWjsCchRTEY_ANcBFJXu-ig13hWygjDOxrA/exec"
API_SECRET = "L10nsClvbQCBELC"

# vvv FILL THIS IN vvv  - where view-ticket.html is hosted, no trailing slash.
# Example: "https://qcbelc.org/raffle" or "https://yourname.github.io/raffle"
TICKET_SITE_BASE_URL = "https://qcbahaghariexecutivelionsclub.github.io/Digital-Raffle-Draw"

MEMBER_NAME = "Quezon City Bahaghari Executive Lions Club"     # shown as "Sold by" on every ticket
PAYMENT_METHOD = "Cash"
TRANSACTION_NUMBER = "1234"        # 4-digit transaction number / PIN

OUTPUT_DIR = "tickets"             # folder where per-person .txt files go
MAX_PER_BACKEND_CALL = 20          # backend caps each call at 20, don't change
DELAY_BETWEEN_CALLS_SEC = 0.6      # be gentle with Apps Script's rate limits

# =================================================================


def slugify(name: str) -> str:
    slug = name.strip().lower()
    slug = re.sub(r"[^a-z0-9]+", "_", slug)
    return slug.strip("_") or "unnamed"


def friendly_filename(name: str, qty: int) -> str:
    """Turns a name into a clean 'First_Last_20_Tickets.txt' style filename -
    easier to recognize at a glance than the old qr_..._tickets.md style."""
    cleaned = re.sub(r"[^A-Za-z0-9]+", "_", name.strip()).strip("_")
    parts = [p.capitalize() for p in cleaned.split("_") if p]
    nice_name = "_".join(parts) or "Unnamed"
    return f"{nice_name}_{qty}_Tickets.txt"


def parse_names_file(path: str, default_qty: int):
    """Returns a list of (name, qty) tuples. Skips blank lines and lines
    starting with '#'. A trailing ': N' / '- N' / ', N' on a line overrides
    the default quantity for that one person."""
    entries = []
    override_re = re.compile(r"^(.*?)[\s:,\-]+(\d+)\s*$")

    with open(path, "r", encoding="utf-8") as f:
        for raw_line in f:
            line = raw_line.strip()
            if not line or line.startswith("#"):
                continue

            m = override_re.match(line)
            if m and m.group(1).strip():
                name = m.group(1).strip()
                qty = int(m.group(2))
            else:
                name = line
                qty = default_qty

            if qty < 1:
                print(f"  ! Skipping '{name}': quantity must be at least 1.")
                continue

            entries.append((name, qty))

    return entries


def call_generate_ticket(donor_name: str, qty: int):
    """Calls the real backend. Returns the parsed JSON response."""
    payload = {
        "action": "generateTicket",
        "secret": API_SECRET,
        "memberName": MEMBER_NAME,
        "donorName": donor_name,
        "donorPhone": "",
        "ticketQty": qty,
        "paymentMethod": PAYMENT_METHOD,
        "transactionNumberId": TRANSACTION_NUMBER,
        "receiptImage": None,
        "receiptFileName": None,
    }
    body = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        WEBAPP_URL,
        data=body,
        headers={"Content-Type": "text/plain;charset=utf-8"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read().decode("utf-8"))


def fake_tickets_for_dry_run(qty: int, seed: int):
    """Mimics the demo-mode fallback in generate-ticket.html, purely for
    testing the script locally without touching the real sheet."""
    ts = str(int(time.time()))[-6:]
    return [
        {"ticketId": f"TKT-{ts}-{seed + i}", "qrCodeId": f"QR-DEMO{ts}{seed + i}"}
        for i in range(qty)
    ]


def generate_for_name(name: str, total_qty: int, dry_run: bool):
    """Handles chunking into batches of <=20 (the backend's per-call cap)
    and returns the full list of {ticketId, qrCodeId} for this person."""
    all_tickets = []
    remaining = total_qty
    seed = 0

    while remaining > 0:
        chunk = min(MAX_PER_BACKEND_CALL, remaining)

        if dry_run:
            tickets = fake_tickets_for_dry_run(chunk, seed)
        else:
            result = call_generate_ticket(name, chunk)
            if not result.get("success"):
                raise RuntimeError(result.get("message", "Unknown error from backend"))
            tickets = result["tickets"]
            time.sleep(DELAY_BETWEEN_CALLS_SEC)

        all_tickets.extend(tickets)
        remaining -= chunk
        seed += chunk

    return all_tickets


def build_view_url(qr_code_id: str, all_qr_ids: list) -> str:
    siblings = [q for q in all_qr_ids if q != qr_code_id]
    url = f"{TICKET_SITE_BASE_URL}/view-ticket.html?qr={qr_code_id}"
    if siblings:
        url += "&siblings=" + ",".join(siblings)
    return url


def write_person_md(name: str, tickets: list, dry_run: bool):
    """Writes a plain .txt file. If the file already exists, it APPENDS 
    new tickets to it rather than overwriting.
    
    NOTE: We now use a consistent filename (Name_Tickets.txt) so that 
    running the script multiple times adds to the same file."""
    
    qty = len(tickets)
    # 1. Use a consistent filename (drop the quantity from the name)
    cleaned = re.sub(r"[^A-Za-z0-9]+", "_", name.strip()).strip("_")
    nice_name = cleaned.capitalize() or "Unnamed"
    filename = f"{nice_name}_Tickets.txt"  # e.g., Wolverine_Tickets.txt
    filepath = os.path.join(OUTPUT_DIR, filename)

    all_qr_ids = [t["qrCodeId"] for t in tickets]
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # 2. Check if file exists to decide how to open it
    is_new_file = not os.path.exists(filepath)
    mode = 'w' if is_new_file else 'a'  # 'w'=Write (new), 'a'=Append (existing)

    divider = "=" * 50
    thin_divider = "-" * 50
    
    # Build header text (only used for new files)
    header_lines = [
        divider,
        f"  TICKETS FOR: {name}",
        divider,
        "",
        f"You have {qty} ticket(s) below.",  # Note: this is just for the current batch
        "Each ticket has ONE link. To view a ticket:",
        "  1. Copy the link under that ticket (select it, then copy).",
        "  2. Paste it into your phone or computer's web browser.",
        "",
        f"Sold by: {MEMBER_NAME}",
        f"Payment method: {PAYMENT_METHOD}",
        f"Transaction number: {TRANSACTION_NUMBER} (PIN: {TRANSACTION_NUMBER[-4:]})",
        "Status: Verified",
        f"Generated: {now}{'  (DRY RUN - not saved to the real sheet)' if dry_run else ''}",
        "",
    ]

    # 3. Write to file (either new or append)
    with open(filepath, mode, encoding='utf-8') as f:
        if is_new_file:
            # Only write the header for the very first batch
            f.write("\n".join(header_lines) + "\n")
        else:
            # If appending, just add a separator so we know where the new batch starts
            f.write("\n" + divider + "  NEW BATCH ADDED  " + divider + "\n\n")

        # Write the ticket details (loop over tickets)
        for i, t in enumerate(tickets, start=1):
            url = build_view_url(t["qrCodeId"], all_qr_ids)
            f.write(f"{thin_divider}\n")
            f.write(f"TICKET {i} of {qty}   (Ticket ID: {t['ticketId']})\n\n")
            f.write("COPY THIS LINK TO VIEW YOUR TICKET:\n")
            f.write(url + "\n\n")

    return filepath


def main():
    args = [a for a in sys.argv[1:] if a != "--dry-run"]
    dry_run = "--dry-run" in sys.argv

    if len(args) < 1:
        print("Usage: python3 generate_bundle_tickets.py <names_file.md> [default_qty] [--dry-run]")
        sys.exit(1)

    names_file = args[0]
    default_qty = int(args[1]) if len(args) > 1 else None

    if not os.path.exists(names_file):
        print(f"Could not find '{names_file}'.")
        sys.exit(1)

    entries = parse_names_file(names_file, default_qty or 0)

    # Any entry that ended up with 0 means no CLI qty AND no inline override.
    missing = [name for name, qty in entries if qty == 0]
    if missing:
        print("These names have no ticket count (add it to the file, or pass one on the command line):")
        for name in missing:
            print(f"  - {name}")
        sys.exit(1)

    if TICKET_SITE_BASE_URL == "https://YOUR-SITE-URL-HERE":
        print("⚠️  Heads up: TICKET_SITE_BASE_URL is still the placeholder at the top of this script.")
        print("   Links saved to the .txt files won't work until you fill it in.\n")

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    print(f"{'[DRY RUN] ' if dry_run else ''}Generating tickets for {len(entries)} name(s)...\n")

    ok, failed = 0, 0
    for name, qty in entries:
        print(f"  {name}: {qty} ticket(s)...", end=" ", flush=True)
        try:
            tickets = generate_for_name(name, qty, dry_run)
            filepath = write_person_md(name, tickets, dry_run)
            print(f"done -> {filepath}")
            ok += 1
        except (RuntimeError, urllib.error.URLError) as e:
            print(f"FAILED ({e})")
            failed += 1

    print(f"\nDone. {ok} succeeded, {failed} failed.")


if __name__ == "__main__":
    main()