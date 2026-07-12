#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_medicines.py  —  ΑΥΤΟΜΑΤΗ ανανέωση medicines.json

Βρίσκει μόνο του το πιο πρόσφατο ΠΛΗΡΕΣ δελτίο τιμών φαρμάκων («Δελτίο
αναθεωρημένων τιμών φαρμάκων ανθρώπινης χρήσης») στη σελίδα του Υπουργείου
Υγείας, κατεβάζει το xlsx, το διαβάζει, και γράφει το medicines.json στη μορφή
που περιμένει η εφαρμογή.

Τρέχει από GitHub Action (βλ. .github/workflows/update-medicines.yml). Τοπικά:
    pip install requests pandas openpyxl
    python build_medicines.py medicines.json
"""

import sys
import io
import re
import json
import unicodedata
import datetime
from urllib.parse import urljoin
import requests
import pandas as pd

LISTING_URL = "https://www.moh.gov.gr/articles/times-farmakwn/deltia-timwn"
# Χαρακτηριστικό κομμάτι του slug ΜΟΝΟ του πλήρους δελτίου (ΟΧΙ Συμπληρωματικό /
# Νέων Γενοσήμων / Τριμήνου). Στη διεύθυνση εμφανίζεται ως "<αριθμός>-deltio-
# anathewrhmenwn-timwn-farmakwn-...".
FULL_BULLETIN_SLUG = "deltio-anathewrhmenwn-timwn-farmakwn"
HEADERS = {"User-Agent": "Mozilla/5.0 (medicines-updater)"}

COLUMN_KEYWORDS = {
    "id":     ["κωδικ"],                          # "Κωδικός" (ΟΧΙ BARCODE)
    "name":   ["περιγραφ", "ονομασ", "εμπορικ"],  # "Περιγραφή Προϊόντος"
    "atc":    ["atc"],
    "active": ["δραστ"],                          # "Δραστική/ές"
    "price":  ["λιανικ"],                         # "Λιανική Τιμή"
}


def norm(s):
    s = "" if s is None else str(s)
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.lower().strip()


def get_html(url):
    r = requests.get(url, headers=HEADERS, timeout=30)
    r.raise_for_status()
    return r.text


def find_latest_full_bulletin_url():
    """Σαρώνει τις πρώτες σελίδες της λίστας (νεότερο-πρώτο) και επιστρέφει το
    URL του άρθρου του νεότερου ΠΛΗΡΟΥΣ δελτίου."""
    # ΔΕΝ απαιτεί "/" πριν το slug: στη διεύθυνση προηγείται ο αριθμός άρθρου.
    pattern = re.compile(r'href="([^"]*' + re.escape(FULL_BULLETIN_SLUG) + r'[^"]*)"')
    for page in range(1, 4):
        url = LISTING_URL if page == 1 else f"{LISTING_URL}?page={page}"
        html = get_html(url)
        m = pattern.search(html)
        if m:
            return urljoin(url, m.group(1).replace("&amp;", "&"))
    sys.exit("Δεν βρέθηκε πλήρες δελτίο («αναθεωρημένων τιμών») στη σελίδα του "
             "Υπουργείου. Ίσως άλλαξε η δομή — στείλε τη σελίδα για προσαρμογή.")


def find_xlsx_download_url(article_url):
    """Στη σελίδα του άρθρου, βρίσκει τον σύνδεσμο λήψης που είναι .xlsx. Ψάχνει
    <a> tags που περιέχουν ΚΑΙ "?fdl=" ΚΑΙ ".xlsx" (σε href ή title), ανεξάρτητα
    από τη σειρά των attributes."""
    html = get_html(article_url)
    for tag in re.findall(r'<a\b[^>]*>', html):
        low = tag.lower()
        if "?fdl=" in low and ".xlsx" in low:
            href_m = re.search(r'href="([^"]+)"', tag)
            if href_m:
                return urljoin(article_url, href_m.group(1).replace("&amp;", "&"))
    sys.exit("Δεν βρέθηκε αρχείο .xlsx στη σελίδα του δελτίου.")


def find_header(raw):
    best_row, best_score = None, -1
    for i in range(min(25, len(raw))):
        cells = [norm(x) for x in raw.iloc[i].tolist()]
        score = sum(
            1 for kws in COLUMN_KEYWORDS.values()
            if any(any(kw in c for kw in kws) for c in cells)
        )
        if score > best_score:
            best_score, best_row = score, i
    if best_row is None or best_score < 3:
        sys.exit("ΛΕΙΠΟΥΝ αναμενόμενες στήλες — το Υπουργείο ίσως άλλαξε τη δομή.")
    header = [norm(x) for x in raw.iloc[best_row].tolist()]
    original = [("" if pd.isna(x) else str(x)) for x in raw.iloc[best_row].tolist()]
    colmap, colnames = {}, {}
    for key, kws in COLUMN_KEYWORDS.items():
        for idx, cell_ in enumerate(header):
            if any(kw in cell_ for kw in kws):
                colmap[key] = idx
                colnames[key] = original[idx].strip()
                break
    missing = [k for k in COLUMN_KEYWORDS if k not in colmap]
    if missing:
        sys.exit(f"ΛΕΙΠΟΥΝ αναμενόμενες στήλες: {missing}. Κεφαλίδα: {[c for c in original if c]}")
    return raw.iloc[best_row + 1:].reset_index(drop=True), colmap, colnames


def parse_price(v):
    if v is None:
        return None
    s = str(v).strip().replace("€", "").replace(" ", "")
    if s == "" or s.lower() == "nan":
        return None
    if "," in s:
        s = s.replace(".", "").replace(",", ".")
    try:
        return round(float(s), 2)
    except ValueError:
        return None


def cell(row, idx):
    v = row.iloc[idx]
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip()
    return "" if s.lower() == "nan" else s


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else "medicines.json"

    article = find_latest_full_bulletin_url()
    print("Δελτίο:", article)
    xlsx_url = find_xlsx_download_url(article)
    print("xlsx:", xlsx_url)

    resp = requests.get(xlsx_url, headers=HEADERS, timeout=60)
    resp.raise_for_status()
    raw = pd.read_excel(io.BytesIO(resp.content), header=None, dtype=str)
    data, col, colnames = find_header(raw)
    print("Στήλες:", {k: colnames[k] for k in ("id", "name", "atc", "active", "price")})

    meds = []
    for _, row in data.iterrows():
        name = cell(row, col["name"])
        price = parse_price(row.iloc[col["price"]])
        if not name or price is None:
            continue
        active = cell(row, col["active"]).upper()
        search_key = name.lower()
        meds.append({
            "id": cell(row, col["id"]),
            "name": name,
            "search_key": search_key,
            "atc": cell(row, col["atc"]),
            "active_ingredient": active,
            "price": price,
            "full_text": (search_key + " " + active.lower()).strip(),
        })

    if len(meds) < 2000:
        sys.exit(f"ΣΦΑΛΜΑ: μόνο {len(meds)} φάρμακα — μάλλον μερικό/χαλασμένο δελτίο. "
                 "Δεν γράφτηκε τίποτα.")

    payload = {"updated_at": datetime.date.today().isoformat(), "medicines": meds}
    with open(out, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, separators=(",", ":"))

    print(f"Γράφτηκαν {len(meds)} φάρμακα στο {out} (updated_at {payload['updated_at']})")


if __name__ == "__main__":
    main()
