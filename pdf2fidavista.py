#!/usr/bin/env python3
"""
pdf2fidavista.py - convert a Swedbank (Latvia) account statement PDF ("Konta Izraksts")
into a FiDAViSta XML statement in the same shape as Swedbank's own FiDAViSta export
(version 1.01, WINDOWS-1257 encoded).

    pip install pdfplumber

    # Inspect the PDF text (one line per printed row):
    python pdf2fidavista.py statement.pdf --dump-text

    # Convert (bank/client/IBAN/period/balances are read from the PDF):
    python pdf2fidavista.py statement.pdf -o statement.xml

Swedbank prints every transaction as a block of rows:

    01.09.2026  123-1  JANIS BERZINS               LV80BANK0000435195001
    2026090100000001  MP                           BANKLV2X          500.00 |
                       PAR augustu
    <date>      <doc nr> <counterparty name>       <counterparty acc> Debets | Kredīts
    <archive nr> <code>  <reg. nr / details>       <BIC>
                         <payment details ...>

Debit vs credit is only visible from which column the amount sits in, so the parser works
on word coordinates (pdfplumber.extract_words) rather than on plain text lines.
Anything that can't be found in the PDF can be passed on the command line.
"""
import argparse
import re
import sys
from datetime import datetime, date
from decimal import Decimal
import xml.etree.ElementTree as ET

import pdfplumber

# ============================== CONFIG ==============================
NAMESPACES = {
    "1.01": "http://www.bankasoc.lv/fidavista/fidavista0101.xsd",
    "1.2": "http://bankasoc.lv/fidavista/fidavista_1-2.xsd",
}
XSI_NS = "http://www.w3.org/2001/XMLSchema-instance"
OUTPUT_ENCODING = "WINDOWS-1257"

# Fixed values Swedbank writes into its own FiDAViSta export.
SWEDBANK_FROM = "Swedbank, business.swedbank.lv, Customer support Tel. +371 67444444"
SWEDBANK_ADDRESS = "bankAddress"

DATE_RE = re.compile(r"^\d{2}\.\d{2}\.\d{4}$")
ARCHIVE_RE = re.compile(r"^\d{16}$")
AMOUNT_RE = re.compile(r"^\d+\.\d{2}$")
THOUSANDS_RE = re.compile(r"^\d{1,3}$")
THOUSANDS_GAP = 4.0  # max gap (pt) between "1" and "266.13" in "1 266.13"
BIC_RE = re.compile(r"^[A-Z]{4}[A-Z]{2}[A-Z0-9]{2}(?:[A-Z0-9]{3})?$")
ACC_RE = re.compile(r"^[A-Z]{2}\d{2}[A-Z0-9]{11,30}$")
CCY_RE = re.compile(r"^[A-Z]{3}$")

# Column x-boundaries (PDF points). Derived from the table header on each page when
# possible ("Dok.", "Pretējās", "Konta", "Debets", "Kredīts"); these are fallbacks.
DEFAULT_COLS = {"doc": 150, "name": 205, "acc": 360, "amount": 460, "credit": 518}
ROW_TOLERANCE = 2.0  # words whose 'top' differs by less than this are on the same row

# Swedbank operation code -> FiDAViSta TypeCode (TypeName is the Swedbank code itself).
# Anything not listed here (MP salary, KOM fee, CTX card ...) is exported by Swedbank as OTHR.
SWED_TYPE_CODES = {"INB": "INP", "IZP": "OUTP", "PRV": "OUTP"}

# Detail rows that only hold the payer / beneficiary registration numbers
# ("40000000000 / 40000000001"); Swedbank leaves them out of PmtInfo.
REG_NO_LINE_RE = re.compile(r"^\d{11}(?: / \d{11})*$")
CARD_DATE_RE = re.compile(r"\b(\d{2}\.\d{2}\.(?:\d{4}|\d{2}))\b")
# ====================================================================


def parse_date(s):
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d.%m.%y"):
        try:
            return datetime.strptime(s.strip(), fmt).date()
        except ValueError:
            pass
    raise ValueError(f"Unrecognised date: {s!r}")


def parse_amount(s):
    """'1 266.13' / '1266,13' / '-12.50' -> Decimal."""
    s = s.strip().replace(" ", "").replace(" ", "")
    sign = -1 if s.startswith("-") else 1
    s = s.lstrip("+-")
    s = s[:-3].replace(",", "").replace(".", "") + "." + s[-2:]
    return sign * Decimal(s)


def group_rows(words):
    """Group pdfplumber words into rows (sorted top-to-bottom, left-to-right)."""
    rows = []
    for w in sorted(words, key=lambda w: (round(w["top"]), w["x0"])):
        if rows and abs(rows[-1][0]["top"] - w["top"]) < ROW_TOLERANCE:
            rows[-1].append(w)
        else:
            rows.append([w])
    return [sorted(r, key=lambda w: w["x0"]) for r in rows]


def text_of(ws):
    return " ".join(w["text"] for w in ws)


def extract_pages(pdf_path):
    """Return a list of pages, each a list of rows (list of word dicts)."""
    with pdfplumber.open(pdf_path) as pdf:
        return [group_rows(p.extract_words(keep_blank_chars=False)) for p in pdf.pages]


def page_columns(rows):
    """Find column boundaries from the table header row, if present on this page."""
    cols = dict(DEFAULT_COLS)
    for r in rows:
        names = {w["text"]: w for w in r}
        if "Debets" in names and "Kredīts" in names:
            cols["doc"] = names.get("Dok.", {}).get("x0", cols["doc"] + 3) - 5
            cols["name"] = names.get("Pretējās", {}).get("x0", cols["name"] + 3) - 5
            cols["acc"] = names.get("Konta", {}).get("x0", cols["acc"] + 3) - 5
            cols["amount"] = names["Debets"]["x0"] - 10
            # Amounts are right-aligned: debit ends ~20pt right of "Debets", credit at the page edge.
            cols["credit"] = (names["Debets"]["x1"] + names["Kredīts"]["x1"]) / 2 + 10
            break
    return cols


def split_amount(row, cols):
    """Split a row into (other words, amount Decimal or None, 'C'/'D' or None)."""
    if not row or row[-1]["x0"] < cols["amount"] or not AMOUNT_RE.match(row[-1]["text"]):
        return row, None, None
    # Walk left over thousands groups that sit right next to each other ("1 266.13").
    i = len(row) - 1
    while i > 0 and THOUSANDS_RE.match(row[i - 1]["text"]) \
            and row[i]["x0"] - row[i - 1]["x1"] < THOUSANDS_GAP:
        i -= 1
    amt_words, rest = row[i:], row[:i]
    amount = parse_amount("".join(w["text"] for w in amt_words))
    cd = "C" if amt_words[-1]["x1"] > cols["credit"] else "D"
    return rest, amount, cd


def zone(ws, lo, hi):
    return [w for w in ws if lo <= w["x0"] < hi]


def parse_header(pages):
    """Bank, client, IBAN, period and balances from the statement header / footer."""
    info = {}
    first = pages[0]
    all_rows = [r for p in pages for r in p]
    full = "\n".join(text_of(r) for r in all_rows)

    m = re.search(r"(\d{2}\.\d{2}\.\d{4})\s*-\s*(\d{2}\.\d{2}\.\d{4})", full)
    if m:
        info["start"], info["end"] = parse_date(m.group(1)), parse_date(m.group(2))
    m = re.search(r"IBAN:\s*([A-Z]{2}\d{2}(?:\s?[A-Z0-9]{1,4}\b)+)", full)
    if m:
        info["iban"] = m.group(1).replace(" ", "")
    m = re.search(r"^(\d{2}\.\d{2}\.\d{4} \d{2}:\d{2}:\d{2})$", full, re.M)  # statement printed on
    if m:
        info["prepared"] = datetime.strptime(m.group(1), "%d.%m.%Y %H:%M:%S")
    m = re.search(r"Izraksta Nr\.\s*(\d+)", full)
    if m:
        info["stmt_no"] = m.group(1)

    # Page 1 header: bank on the left, client on the right, above the table header.
    for r in first:
        t = text_of(r)
        if "Op.datums" in t:
            break
        left = [w for w in r if w["x0"] < 230]
        right = [w for w in r if w["x0"] >= 460]
        lt, rt = text_of(left), text_of(right)
        if lt.startswith("Reģ.Nr:"):
            info["bank_id"] = lt.split(":", 1)[1].strip()
            if rt.startswith("Reģ.Nr:"):
                info["client_id"] = rt.split(":", 1)[1].strip()
        elif lt.startswith("BIC:"):
            info["bank_bic"] = lt.split(":", 1)[1].strip()
        elif left and right and "bank_name" not in info and not DATE_RE.match(r[0]["text"]):
            info["bank_name"], info["client_name"] = lt, rt

    # Balances / turnovers (labels in the middle, amount on the right).
    labels = {
        "Sākuma atlikums": "open_bal", "Beigu atlikums": "close_bal",
        "Kredīta apgrozījums": "credit_turnover", "Debeta apgrozījums": "debit_turnover",
    }
    for r in all_rows:
        t = text_of(r)
        for label, key in labels.items():
            if label in t and key not in info:
                m = re.search(rf"{label}\s+(-?[\d ]+\.\d{{2}})\s*$", t)
                if m:
                    info[key] = parse_amount(m.group(1))
        if CCY_RE.match(r[0]["text"]) and r[0]["x0"] < 75 and "ccy" not in info:
            info["ccy"] = r[0]["text"]
    return info


def parse_transactions(pages):
    trx = []
    cur = None
    for rows in pages:
        cols = page_columns(rows)
        in_table = False
        for r in rows:
            t = text_of(r)
            first = r[0]["text"]
            if "Informācija saņēmējam" in t:      # last row of the table header
                in_table = True
                continue
            if not in_table:
                continue
            # Summary / balance rows ("EUR 01.09.2026 Sākuma atlikums ...") end a block.
            if (CCY_RE.match(first) and r[0]["x0"] < cols["doc"]) or "apgrozījums" in t or "atlikums" in t \
                    or "Komisiju kopsumma" in t or "Elektroniskais paraksts" in t:
                cur = None
                continue

            if DATE_RE.match(first) and r[0]["x0"] < cols["doc"]:
                # Row 1: <date> <doc nr> <name> <account>
                cur = {
                    "book": parse_date(first),
                    "doc_no": text_of(zone(r[1:], cols["doc"], cols["name"])),
                    "cp_name": text_of(zone(r[1:], cols["name"], cols["acc"])),
                    "cp_acc": text_of(zone(r[1:], cols["acc"], cols["amount"])).replace(" ", ""),
                    "info": [],
                    "amount": None,
                }
                trx.append(cur)
            elif cur and ARCHIVE_RE.match(first) and cur["amount"] is None:
                # Row 2: <archive nr> <code> <reg nr | details> <BIC> <amount>
                rest, amount, cd = split_amount(r[1:], cols)
                if amount is None:
                    sys.exit(f"Can't find amount in row: {t}")
                cur["bank_ref"] = first
                cur["code"] = text_of(zone(rest, cols["doc"], cols["name"]))
                cur["amount"], cur["cd"] = amount, cd
                middle = zone(rest, cols["name"], float("inf"))
                if cur["cp_acc"]:
                    if middle and middle[-1]["x0"] >= cols["acc"] - 5 and BIC_RE.match(middle[-1]["text"]):
                        cur["cp_bic"] = middle.pop()["text"]
                    cur["cp_legal_id"] = text_of(middle)
                else:
                    cur["info"].append(text_of(middle))   # KOM / CTX: details on this row
            elif cur and cur["amount"] is not None and r[0]["x0"] >= cols["name"] - 5:
                cur["info"].append(t)                    # Row 3+: payment details
            else:
                cur = None
    for t in trx:
        if t["amount"] is None:
            sys.exit(f"Transaction on {t['book']} ({t['cp_name']}) has no amount row.")
    return trx


def classify(t):
    code = t.get("code", "")
    return SWED_TYPE_CODES.get(code, "OTHR"), code


def value_date(t):
    """Card transactions are valued on the card date printed in the details; others on book date."""
    if t.get("code") == "CTX":
        m = CARD_DATE_RE.search(" ".join(t["info"]))
        if m:
            return parse_date(m.group(1))
    return t["book"]


def pmt_info(t):
    lines = [l for i, l in enumerate(t["info"]) if l and not (i and REG_NO_LINE_RE.match(l))]
    return " ".join(lines) or t["type_name"]


def sub(parent, tag, text=None):
    el = ET.SubElement(parent, tag)
    if text is not None:
        el.text = str(text)
    return el


def build_xml(info, trx, version):
    root = ET.Element("FIDAVISTA", {"xmlns:xsi": XSI_NS, "xmlns": NAMESPACES[version]})
    prepared = info.get("prepared") or datetime.now()
    header = sub(root, "Header")
    sub(header, "Timestamp", prepared.strftime("%Y%m%d%H%M%S%f")[:17])
    sub(header, "From", (info.get("from") or SWEDBANK_FROM)[:70])

    st = sub(root, "Statement")
    period = sub(st, "Period")
    sub(period, "StartDate", info["start"].isoformat())
    sub(period, "EndDate", info["end"].isoformat())
    sub(period, "PrepDate", prepared.date().isoformat())

    if info.get("bank_name") or info.get("bank_id"):
        bank = sub(st, "BankSet")
        sub(bank, "Name", (info.get("bank_name") or "")[:140])
        sub(bank, "LegalId", (info.get("bank_id") or "")[:20])
        sub(bank, "Address", (info.get("bank_address") or SWEDBANK_ADDRESS)[:70])
    if info.get("client_name") or info.get("client_id"):
        client = sub(st, "ClientSet")
        sub(client, "Name", (info.get("client_name") or "")[:140])
        sub(client, "LegalId", (info.get("client_id") or "")[:20])

    acc = sub(st, "AccountSet")
    sub(acc, "AccNo", info["iban"])
    ccy = sub(acc, "CcyStmt")
    sub(ccy, "Ccy", info["ccy"])
    sub(ccy, "OpenBal", f"{info['open_bal']:.2f}")
    sub(ccy, "CloseBal", f"{info['close_bal']:.2f}")

    for t in trx:
        ts = sub(ccy, "TrxSet")
        sub(ts, "TypeCode", t["type_code"])
        sub(ts, "TypeName", t["type_name"][:70])
        # Only payment orders have a registration date; bank fees / card transactions leave it empty.
        sub(ts, "RegDate", t["book"].isoformat() if t["cp_acc"] else "")
        sub(ts, "BookDate", t["book"].isoformat())
        sub(ts, "ValueDate", value_date(t).isoformat())
        sub(ts, "BankRef", t["bank_ref"][:25])
        if t["doc_no"]:
            sub(ts, "DocNo", t["doc_no"][:25])
        sub(ts, "CorD", t["cd"])
        sub(ts, "AccAmt", f"{t['amount']:.2f}")
        sub(ts, "PmtInfo", pmt_info(t)[:200])
        cp = sub(ts, "CPartySet")
        sub(cp, "AccNo", t["cp_acc"][:34])
        if t["cp_acc"]:
            holder = sub(cp, "AccHolder")
            sub(holder, "Name", t["cp_name"][:140])
            sub(holder, "LegalId", (t.get("cp_legal_id") or "")[:35])
        sub(cp, "BankCode", (t.get("cp_bic") or "")[:20])
        sub(cp, "Ccy", info["ccy"])
        sub(cp, "Amt", f"{t['amount']:.2f}")

    body = ET.tostring(root, encoding="unicode", short_empty_elements=False)
    decl = f'<?xml version="1.0" encoding="{OUTPUT_ENCODING}"?>'
    return (decl + body).encode(OUTPUT_ENCODING, errors="xmlcharrefreplace")


def main():
    ap = argparse.ArgumentParser(description="Convert a Swedbank statement PDF to FiDAViSta XML")
    ap.add_argument("pdf")
    ap.add_argument("-o", "--output", help="output .xml (default: <pdf name>.xml)")
    ap.add_argument("--dump-text", action="store_true", help="print extracted PDF rows and exit")
    ap.add_argument("--version", choices=NAMESPACES, default="1.01", help="FiDAViSta version")
    ap.add_argument("--iban")
    ap.add_argument("--ccy")
    ap.add_argument("--open-bal")
    ap.add_argument("--close-bal")
    ap.add_argument("--start", help="period start, e.g. 2026-09-01")
    ap.add_argument("--end", help="period end, e.g. 2026-09-30")
    ap.add_argument("--bank-name")
    ap.add_argument("--bank-id")
    ap.add_argument("--client-name")
    ap.add_argument("--client-id")
    ap.add_argument("--from", dest="from_", help="Header/From value")
    a = ap.parse_args()

    pages = extract_pages(a.pdf)
    if not any(pages):
        sys.exit("No text found - the PDF is probably a scan. Run OCR first (e.g. ocrmypdf).")
    if a.dump_text:
        n = 0
        for p in pages:
            for r in p:
                n += 1
                print(f"{n:4}: {text_of(r)}")
        return

    info = parse_header(pages)
    overrides = {
        "iban": a.iban, "ccy": a.ccy, "bank_name": a.bank_name, "bank_id": a.bank_id,
        "client_name": a.client_name, "client_id": a.client_id, "from": a.from_,
        "open_bal": a.open_bal and parse_amount(a.open_bal),
        "close_bal": a.close_bal and parse_amount(a.close_bal),
        "start": a.start and parse_date(a.start), "end": a.end and parse_date(a.end),
    }
    info.update({k: v for k, v in overrides.items() if v is not None})
    info.setdefault("ccy", "EUR")

    trx = parse_transactions(pages)
    if not trx:
        sys.exit("No transactions found. Run with --dump-text and check the layout.")
    for t in trx:
        t["type_code"], t["type_name"] = classify(t)
    info.setdefault("start", min(t["book"] for t in trx))
    info.setdefault("end", max(t["book"] for t in trx))

    missing = [n for n in ("iban", "open_bal") if info.get(n) is None]
    if missing:
        sys.exit(f"Couldn't find {', '.join('--' + m.replace('_', '-') for m in missing)} "
                 "in the PDF; pass it on the command line.")

    # Sanity checks: turnovers and opening + credits - debits = closing
    credits = sum(t["amount"] for t in trx if t["cd"] == "C")
    debits = sum(t["amount"] for t in trx if t["cd"] == "D")
    computed = info["open_bal"] + credits - debits
    print(f"Account {info['iban']} {info['ccy']}  {info['start']} .. {info['end']}")
    print(f"Transactions: {len(trx)}  (credits {sum(t['cd'] == 'C' for t in trx)}, "
          f"debits {sum(t['cd'] == 'D' for t in trx)})")
    ok = True
    for label, got, key in (("Credit turnover", credits, "credit_turnover"),
                            ("Debit turnover", debits, "debit_turnover"),
                            ("Closing balance", computed, "close_bal")):
        exp = info.get(key)
        status = "" if exp is None else ("OK" if got == exp else f"MISMATCH (PDF says {exp:.2f})")
        ok &= exp is None or got == exp
        print(f"  {label:16} {got:>12.2f}  {status}")
    info.setdefault("close_bal", computed)

    out = a.output or re.sub(r"\.pdf$", "", a.pdf, flags=re.I) + ".xml"
    with open(out, "wb") as f:
        f.write(build_xml(info, trx, a.version))
    print(f"Written {out}")
    if not ok:
        sys.exit("Totals don't match the PDF - check the parsed transactions!")


if __name__ == "__main__":
    main()
