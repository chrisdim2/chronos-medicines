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

------------------------------------------------------------------------------
ΔΥΟ ΕΙΔΗ ΑΠΟΤΥΧΙΑΣ, ΚΑΙ ΓΙΑΤΙ ΔΕΝ ΕΙΝΑΙ ΤΟ ΙΔΙΟ ΠΡΑΓΜΑ

1. «Η ΠΗΓΗ ΕΙΝΑΙ ΠΡΟΣΩΡΙΝΑ ΚΑΤΩ» (503, 502, timeout, μπλοκάρισμα IP).
   ΔΕΝ φταίει τίποτα δικό μας και δεν υπάρχει τίποτα να διορθώσουμε. Το script
   τελειώνει με exit 0, δεν γράφει τίποτα, η ροή βγαίνει ΠΡΑΣΙΝΗ και δεν
   στέλνεται email. Την επόμενη Δευτέρα ξαναδοκιμάζει.

   Πριν, αυτό έριχνε τη ροή με exit 1 και έστελνε email αποτυχίας για κάτι που
   ο παραλήπτης δεν μπορούσε να διορθώσει.

2. «ΑΛΛΑΞΕ Η ΔΟΜΗ ΤΗΣ ΠΗΓΗΣ» (δεν βρέθηκε δελτίο, δεν βρέθηκε xlsx, λείπουν
   στήλες, ελάχιστα φάρμακα). ΑΥΤΟ θέλει άνθρωπο. Τελειώνει με exit 1 και
   στέλνεται email, όπως πρέπει.

Έτσι, ένα κόκκινο Χ σημαίνει πάντα «χρειάζομαι εσένα», ποτέ «το moh.gov.gr
είχε κακή μέρα».
------------------------------------------------------------------------------
"""

import sys
import io
import re
import json
import time
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

# ΓΙΑΤΙ ΚΑΝΟΝΙΚΟΣ USER-AGENT ΦΥΛΛΟΜΕΤΡΗΤΗ: ο προηγούμενος έγραφε ρητά
# «medicines-updater», δηλαδή αυτοσυστηνόταν ως script. Τα τείχη προστασίας
# κρατικών ιστότοπων κόβουν τέτοια κεφαλίδα, και μαζί κόβουν συχνά και ολόκληρα
# εύρη IP των GitHub Actions. Τα Accept/Accept-Language υπάρχουν για τον ίδιο
# λόγο: ένα αίτημα χωρίς αυτά ξεχωρίζει αμέσως από αίτημα φυλλομετρητή.
HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/126.0.0.0 Safari/537.36"),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "el-GR,el;q=0.9,en-US;q=0.8,en;q=0.7",
    "Connection": "keep-alive",
}

# Κωδικοί που σημαίνουν «ξαναδοκίμασε», όχι «κάτι έσπασε».
TRANSIENT_STATUS = {408, 425, 429, 500, 502, 503, 504}
RETRIES = 4
RETRY_SLEEP = 6          # δευτερόλεπτα, πολλαπλασιάζεται σε κάθε προσπάθεια

COLUMN_KEYWORDS = {
    "id":     ["κωδικ"],                                   # "Κωδικός" (ΟΧΙ BARCODE)
    "name":   ["περιγραφ", "ονομασ", "εμπορικ", "προιον"],  # "Περιγραφή Προϊόντος" / "Προϊόν"
    "atc":    ["atc"],
    "active": ["δραστ"],                          # "Δραστική/ές"
    "price":  ["λιανικ"],                         # "Λιανική Τιμή"
}


def norm(s):
    s = "" if s is None else str(s)
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return s.lower().strip()


def source_unavailable(reason):
    """Έξοδος ΧΩΡΙΣ σφάλμα: η πηγή δεν απαντά, δεν φταίει τίποτα δικό μας.

    Το exit 0 είναι σκόπιμο. Κρατά τη ροή πράσινη ώστε τα emails αποτυχίας να
    σημαίνουν πάντα κάτι που χρειάζεται άνθρωπο.
    """
    print(f"[ΠΑΡΑΛΕΙΨΗ] {reason}")
    print("Το medicines.json ΔΕΝ πειράχτηκε. Θα ξαναδοκιμάσει στον επόμενο κύκλο.")
    sys.exit(0)


def structure_changed(reason):
    """Έξοδος ΜΕ σφάλμα: η πηγή απάντησε, αλλά δεν είναι πια αυτό που περιμέναμε.
    Αυτό θέλει χειροκίνητη προσαρμογή, οπότε πρέπει να φανεί."""
    print(f"[ΣΦΑΛΜΑ] {reason}", file=sys.stderr)
    sys.exit(1)


def fetch(url, timeout=30):
    """GET με επαναλήψεις στα προσωρινά σφάλματα. Επιστρέφει το response.

    Δεν κάνει raise_for_status εδώ: ο καλών αποφασίζει αν ένα 404 σημαίνει
    «άλλαξε η δομή» ή κάτι άλλο.
    """
    last = None
    for attempt in range(RETRIES):
        try:
            r = requests.get(url, headers=HEADERS, timeout=timeout)
        except requests.RequestException as e:
            last = f"σφάλμα δικτύου ({e.__class__.__name__})"
        else:
            if r.status_code not in TRANSIENT_STATUS:
                return r
            last = f"HTTP {r.status_code}"

        if attempt < RETRIES - 1:
            wait = RETRY_SLEEP * (attempt + 1)
            print(f"  {last}, νέα προσπάθεια σε {wait}s "
                  f"({attempt + 2} από {RETRIES})…")
            time.sleep(wait)

    source_unavailable(f"Η πηγή δεν απάντησε μετά από {RETRIES} προσπάθειες "
                       f"({last}), {url}")


def get_html(url):
    r = fetch(url)
    if r.status_code >= 400:
        structure_changed(f"HTTP {r.status_code} στο {url}. "
                          f"Η διεύθυνση μάλλον άλλαξε.")
    return r.text


def find_latest_full_bulletin_url():
    """Σαρώνει τις πρώτες σελίδες της λίστας (νεότερο-πρώτο) και επιστρέφει το
    URL του άρθρου του νεότερου ΠΛΗΡΟΥΣ δελτίου."""
    # Απαιτεί το slug ΑΜΕΣΩΣ μετά τον αριθμό άρθρου: "/<αριθμός>-deltio-
    # anathewrhmenwn-...". Έτσι αποκλείονται άρθρα «Τροποποίηση/Ορθή Επανάληψη»
    # που απλώς ΑΝΑΦΕΡΟΥΝ το δελτίο στον τίτλο τους (slug "...-laquo-tropopoihsh-
    # ...-deltio-anathewrhmenwn-...") και δεν είναι το ίδιο το πλήρες δελτίο.
    pattern = re.compile(r'href="([^"]*/\d+-' + re.escape(FULL_BULLETIN_SLUG) + r'[^"]*)"')
    for page in range(1, 4):
        url = LISTING_URL if page == 1 else f"{LISTING_URL}?page={page}"
        html = get_html(url)
        m = pattern.search(html)
        if m:
            return urljoin(url, m.group(1).replace("&amp;", "&"))
    structure_changed("Δεν βρέθηκε πλήρες δελτίο («αναθεωρημένων τιμών») στη "
                      "σελίδα του Υπουργείου. Ίσως άλλαξε η δομή ή το slug.")


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
    structure_changed("Δεν βρέθηκε αρχείο .xlsx στη σελίδα του δελτίου.")


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
        structure_changed("ΛΕΙΠΟΥΝ αναμενόμενες στήλες — το Υπουργείο ίσως "
                          "άλλαξε τη δομή του Excel.")
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
        structure_changed(f"ΛΕΙΠΟΥΝ αναμενόμενες στήλες: {missing}. "
                          f"Κεφαλίδα: {[c for c in original if c]}")
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

    # Το xlsx είναι πολλών megabyte, οπότε πιο γενναιόδωρο timeout. Οι ίδιες
    # επαναλήψεις ισχύουν και εδώ: μια διακοπή στη μέση του κατεβάσματος δεν
    # είναι λόγος να πέσει η ροή.
    resp = fetch(xlsx_url, timeout=120)
    if resp.status_code >= 400:
        structure_changed(f"HTTP {resp.status_code} στο κατέβασμα του xlsx.")

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
        structure_changed(f"Μόνο {len(meds)} φάρμακα — μάλλον μερικό ή χαλασμένο "
                          f"δελτίο. Δεν γράφτηκε τίποτα.")

    payload = {"updated_at": datetime.date.today().isoformat(), "medicines": meds}
    with open(out, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, separators=(",", ":"))

    print(f"Γράφτηκαν {len(meds)} φάρμακα στο {out} (updated_at {payload['updated_at']})")


if __name__ == "__main__":
    main()
