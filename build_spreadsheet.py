#!/usr/bin/env python3
"""
build_spreadsheet.py  -  PSN Library Exporter, spreadsheet builder (v8)

Contains NO credentials, tokens or account data. Emails and card digits are removed from anything it reads.
Your CSVs / workbook DO contain your game list, purchase history and your PSN online ID (a public handle) - keep them private.

EASIEST USE - put this script, your exports and your previous workbook in one folder, then:
    python build_spreadsheet.py
It auto-picks psn-library_*.csv, psn-transactions_*.csv and the newest PlayStation-Game-Collection*.xlsx.

OTHER WAYS
    python build_spreadsheet.py a.csv b.csv transactions.csv previous.xlsx      (any order)
    python build_spreadsheet.py --base USD        (currency all money is converted to; asked once, then remembered in the workbook)
    python build_spreadsheet.py --pick | --fresh | --partial
    python build_spreadsheet.py --psnp-html "My PSNProfiles page.html"          (OPTIONAL trophies; repeat the flag per file)
Optional files in this folder (see /sample):
    collections.csv     Collection;Component 1;Component 2...   one purchase that unlocks several library games
    aliases.csv         Region;Name;DisplayName                 split / rename games that share a name (e.g. two Resident Evil 4s)
    trophy_aliases.csv  TrophyName;LibraryGame                  PSNProfiles lists that live inside one library game (trilogies)
    type_overrides.csv  Name;Type

OPTIONAL ratings from IGDB (needs the free Twitch developer credentials in the environment variables TWITCH_CLIENT_ID / TWITCH_CLIENT_SECRET):
    python build_spreadsheet.py --igdb          (results are cached in igdb_cache.json; later builds reuse it automatically, even without --igdb)
    igdb_overrides.csv                          Name;IgdbID   (or Region;Name;IgdbID)   forces a specific IGDB game

Needs:  pip install pandas openpyxl        (+ beautifulsoup4 only if you use --psnp-html)
"""
import argparse, glob, os, re, sys, unicodedata, datetime as dt
from collections import defaultdict
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

TODAY = dt.date.today().isoformat()
PURCHASED_TAB = "recently-purchased"
GAME_TYPES = {"Game"}
ALL_TYPES = ["Game", "Demo", "App", "Soundtrack", "Theme", "Avatar", "Add-on", "Extra"]
STATUSES = ["Backlog", "Playing", "Finished", "Dropped", "Won't play"]
PRIORITIES = ["High", "Medium", "Low"]
ISO = {"\u00a3": "GBP", "GBP": "GBP", "$": "USD", "US$": "USD", "USD": "USD", "TL": "TRY", "\u20ba": "TRY", "TRY": "TRY", "\u20ac": "EUR", "EUR": "EUR"}
CURRENCY_REGION = {"GBP": "UK", "TRY": "TR", "USD": "US", "EUR": "EU"}
REGION_CURRENCY = {v: k for k, v in CURRENCY_REGION.items()}
STRONG = {"GBP", "TRY"}
REGION_LOCALE = {"UK": "en-gb", "TR": "en-tr", "US": "en-us"}
SYMBOL = {"GBP": "\u00a3", "USD": "$", "TRY": "TL ", "EUR": "\u20ac", "CAD": "CA$", "AUD": "A$", "JPY": "\u00a5"}
COLLECTIONS, ALIASES, TROPHY_ALIASES = {}, {}, {}
# collections whose components were checked against real receipts (Triple Pack / Raccoon City Edition / Yakuza)
COLLECTIONS_RAW = {
    "yakuzaremasteredcollection": ["yakuza3", "yakuza4", "yakuza5"],
    "residenteviltriplepack": ["residentevil4", "residentevil5", "residentevil6"],
    "raccooncityedition": ["residentevil2", "residentevil3", "residentevilresistance"],
}

# ----------------------------------------------------------------------------- helpers
def s(v):
    if v is None: return ""
    if isinstance(v, float) and pd.isna(v): return ""
    if isinstance(v, (dt.datetime, dt.date)): return v.strftime("%Y-%m-%d")
    t = str(v).strip()
    return "" if t.lower() in ("nan", "nat", "none") else t
def truthy(v): return s(v).lower() in ("true", "1", "1.0", "yes", "y")
def flt(v):
    try: return float(s(v))
    except ValueError: return None
def order_id(v):
    t = s(v)
    if re.fullmatch(r"[\d.]+e\+\d+", t, re.I): t = str(int(float(t)))
    return t
def acct_label(region, account): return f"{region} \u00b7 {account}" if account else region
def is_online_id(a): return bool(re.fullmatch(r"[A-Za-z][A-Za-z0-9_-]{1,15}", s(a)))     # PSN online IDs never contain spaces

EMAIL_RE = re.compile(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+")
CARD_RE = re.compile(r"\(\s*\*{2,}\s*\d{2,4}\s*\)")
def redact(text):
    """Remove e-mail addresses and card digits from expanded order text (older exports contained them)."""
    t = s(text)
    t = re.sub(r"(ACCOUNT\s*\|\s*)[^|]*", r"\1[removed] ", t, flags=re.I)
    return CARD_RE.sub("(****)", EMAIL_RE.sub("[email]", t))

PLATFORM_WORDS = r"\b(ps4|ps5)\b(\s*(&|and|/|\+)\s*\b(ps4|ps5)\b)?"
EDITION_WORDS = {"digital", "deluxe", "standard", "ultimate", "premium", "collector", "collectors", "gold", "complete",
    "definitive", "goty", "game", "of", "the", "year", "edition", "legendary", "launch", "anniversary", "director",
    "directors", "cut", "enhanced", "special", "bundle", "version", "and", "ps4", "ps5", "upgrade", "cross", "gen", "crossgen"}
SOFT_EDITION = {"remastered", "remaster", "hd", "redux", "reloaded", "collection", "trilogy"}      # NOTE: "remake" is deliberately NOT here
TAIL_KEYWORDS = {"edition", "bundle", "version", "cut", "collection"}      # NOT "pack": "... Content Pack" is DLC, not the base game
PUBLISHER_PREFIXES = (("ea", "sports"), ("ea",))
NUMBER_WORDS = {"zero": "0", "one": "1", "two": "2", "three": "3", "four": "4", "five": "5", "six": "6", "seven": "7", "eight": "8", "nine": "9", "ten": "10",
                "ii": "2", "iii": "3", "iv": "4", "v": "5", "vi": "6", "vii": "7", "viii": "8", "ix": "9"}

def clean_name(n):
    n = re.sub(r"[\u2122\u00ae\u00a9]", "", s(n))
    n = unicodedata.normalize("NFKD", n)
    n = "".join(ch for ch in n if not unicodedata.combining(ch)).lower().replace("\u00a0", " ")
    n = re.sub(r"\[[^\]]*\]", " ", n)
    n = re.sub(r"[\u2018\u2019']", "", n)
    n = n.replace("&", " and ")
    return re.sub(PLATFORM_WORDS, " ", n)
def words_of(name): return re.findall(r"[a-z0-9]+", clean_name(name))
def norm(name): return re.sub(r"[^a-z0-9]+", "", clean_name(name))
def base_key(name):
    w = words_of(name)
    while len(w) > 1 and w[-1] in EDITION_WORDS: w.pop()
    return "".join(w)
def loose_key(name):
    txt = re.sub(r"\s*[-:]\s*(episode|chapter)\s*\d+.*$", "", clean_name(name))
    txt = re.sub(r"\b(f2p|free to play|free access|free experience|free edition)\b", " ", txt)
    w = re.findall(r"[a-z0-9]+", txt)
    w = [NUMBER_WORDS.get(x, x) if len(w) > 1 else x for x in w]
    while len(w) > 1 and (w[-1] in EDITION_WORDS or w[-1] in SOFT_EDITION): w.pop()
    if len(w) > 1 and w[0] == "the": w = w[1:]
    return "".join(w)
def weak_keys(name):
    """Generic suffix/prefix rules: 'EA SPORTS FIFA 23 Standard Edition' -> fifa23, 'Jotun: Valhalla Edition' -> jotun.
    Only ever used when the result identifies exactly ONE library game."""
    w = words_of(name); cands = [w]
    for p in PUBLISHER_PREFIXES:
        if tuple(w[:len(p)]) == p and len(w) > len(p): cands.append(w[len(p):])
    out = set()
    for c in cands:
        out.add(base_key(" ".join(c)))
        if c and c[-1] in TAIL_KEYWORDS:
            for k in range(2, 6):
                if len(c) - k >= 1: out.add("".join(c[:len(c) - k]))
    return {k for k in out if len(k) >= 5}
def keys_of(name): return {base_key(name), loose_key(name)} - {""}
def coll_key(name): return re.sub(r"^the", "", norm(name))
COLLECTIONS.update({coll_key(k): v for k, v in COLLECTIONS_RAW.items()})

REGION_FIX = {"GB": "UK"}
def region_from_url(url, default=""):
    m = re.search(r"/([a-z]{2})-([a-z]{2})/", s(url))
    return REGION_FIX.get(m.group(2).upper(), m.group(2).upper()) if m else default
def region_from_locale(loc):
    m = re.search(r"([a-z]{2})[-_]([a-z]{2})", s(loc), re.I)
    return REGION_FIX.get(m.group(2).upper(), m.group(2).upper()) if m else ""
def link_for_region(url, region):
    loc = REGION_LOCALE.get(region)
    return re.sub(r"/[a-z]{2}-[a-z]{2}/", f"/{loc}/", url, count=1) if loc and url else url
def platform_from_id(tid):
    tid = s(tid); return "PS5" if tid.startswith("PPSA") else "PS4" if tid.startswith("CUSA") else ""
def title_number(tid):
    m = re.match(r"(?:CUSA|PPSA)(\d+)", s(tid)); return int(m.group(1)) if m else None

def infer_dayfirst(values, locale=""):
    d_first = m_first = False
    for v in values:
        m = re.match(r"\s*(\d{1,2})[/.\-](\d{1,2})[/.\-]\d{4}", s(v))
        if m:
            a, b = int(m.group(1)), int(m.group(2)); d_first |= a > 12; m_first |= b > 12
    if d_first and not m_first: return True
    if m_first and not d_first: return False
    return not s(locale).lower().startswith("en-us")
def parse_date(raw, dayfirst=True):
    raw = s(raw)
    if not raw: return ""
    if re.match(r"\d{4}-\d{2}-\d{2}", raw): return raw[:10]
    d = pd.to_datetime(raw, dayfirst=dayfirst, errors="coerce")
    return d.date().isoformat() if pd.notna(d) else ""
def parse_amount(raw):
    raw = s(raw).replace("\u00a0", " ")
    sym = re.sub(r"[\d\s.,\-+()\u2212]", "", raw)
    cur = ISO.get(sym, ISO.get(sym.upper(), sym))
    num = re.sub(r"[^\d.,]", "", raw)
    if not num: return cur, None
    if "," in num and "." in num:
        dec = "," if num.rfind(",") > num.rfind(".") else "."
        num = num.replace("." if dec == "," else ",", "").replace(dec, ".")
    elif "," in num:
        head, _, tail = num.rpartition(",")
        num = (head.replace(",", "") + "." + tail) if len(tail) in (1, 2) else num.replace(",", "")
    try: val = float(num)
    except ValueError: return cur, None
    if raw.strip().startswith(("(", "-", "\u2212")): val = -val
    return cur, val
def tx_region(cur, csv_region):
    if cur in STRONG: return CURRENCY_REGION[cur]
    return csv_region or CURRENCY_REGION.get(cur, "")

# ----------------------------------------------------------------------------- classification
APP_NAMES = {"youtube", "bbciplayer", "mediaplayer", "netflix", "plex", "primevideo", "amazonprimevideo", "spotify", "twitch", "ign", "itvx", "itv",
    "disney", "disneyplus", "crunchyroll", "appletv", "nowtv", "now", "channel4", "all4", "funimationnow", "vudu", "hulu", "peacock", "btsport", "tntsports",
    "skysports", "skygo", "skyticket", "eurosport", "dazn", "espn", "beinsports", "hayu", "rakutentv", "paramountplus", "my5", "uktvplay", "stvplayer",
    "redbulltv", "musicunlimited", "videounlimited", "internetbrowser", "twitter", "facebook", "tunein", "deezer", "vevo", "wwenetwork", "nflgamepass", "nhl", "mlbtv",
    "playstationvue", "playstationstore", "playstationtv", "livefromplaystation", "amazonmusic", "applemusic", "youtubetv", "kodi", "emby", "vimeo", "dailymotion"}
BROADCASTERS = r"(bt|sky|tnt|espn|eurosport|dazn|bein|fox|nbc|cbs|itv|bbc|stv|cnn|abc|nfl|nba|wwe|ufc|mlb|nhl|sport|red bull|amazon|apple|disney)"
def classify_library(name):
    n = s(name).lower(); flat = re.sub(r"[^a-z0-9]+", "", n)
    if flat in APP_NAMES: return "App"
    words = re.findall(r"[a-z0-9+]+", n)
    if len(words) <= 3 and re.search(r"\b(tv|iplayer|radio|streaming|player|browser)\b", n) and not re.search(r"\b(game|simulator|manager|tycoon|quest)\b", n): return "App"
    if len(words) <= 3 and re.search(r"\b(sport|sports|news|network|play|go|now|plus)\b", n) and re.match(r"\s*" + BROADCASTERS + r"\b", n): return "App"
    if re.search(r"\b(original |digital )?(game )?soundtrack\b|\bost\b", n): return "Soundtrack"
    if re.search(r"\bdynamic theme\b|\btheme (pack|bundle)\b|\btheme$", n): return "Theme"
    if re.search(r"\bavatars?\b.*\b(pack|bundle|set|collection)\b|\b(pack|bundle|set)\b.*\bavatars?\b", n): return "Avatar"
    if re.search(r"\b(demo|trial|teaser)\b|\b(open |closed )?beta( test)?\b|\bplaytest\b", n): return "Demo"
    if re.search(r"companion app|\bapp$", n): return "App"
    if re.search(r"season pass|expansion pass|\bdlc\b|add-on|\bupgrade\b", n): return "Add-on"
    if re.search(r"wrap-?up|art ?book|digital art|the art of|comic books?|bonus content", n): return "Extra"
    return "Game"
def classify_tx_name(name):
    n = s(name).lower()
    if re.search(r"playstation ?plus|ps ?plus (essential|extra|premium|deluxe)|ps ?now|\bsubscription\b|membership", n) and not re.search(r"\bpack\b|\bavatar", n): return "Subscription"
    if re.search(r"wallet|top-?up|funds", n): return "Wallet top-up"
    t = classify_library(name)
    if t != "Game": return t
    if re.search(r"currency|credits|coins|gems|\bpoints\b|\bpack\b|bundle|booster|starter|ticket|outfit|skin", n): return "In-game item / pack"
    if re.search(r"season pass|expansion|episode|chapter|\bdlc\b", n): return "Add-on"
    return ""

# ----------------------------------------------------------------------------- readers
def read_scan_csv(path):
    df = pd.read_csv(path, encoding="utf-8-sig", dtype=str).fillna("")
    low = {c.lower().replace(" ", "").replace("_", ""): c for c in df.columns}
    g = lambda *k: next((df[low[x]] for x in k if x in low), pd.Series([""] * len(df), index=df.index))
    out = pd.DataFrame({
        "Name": g("name", "title", "titlename").str.replace("\u00a0", " ").str.strip(),
        "Platform": g("platform"), "TitleID": g("titleid"), "ProductID": g("productid"), "ConceptID": g("conceptid"),
        "Subscription": g("subscription"), "StoreURL": g("storeurl", "url"), "Image": g("image", "imageurl"),
        "Region": g("region"), "Account": g("account"), "Locale": g("locale"), "Tab": g("tab"), "ReportedTotal": g("reportedtotal"), "ScannedAt": g("scannedat")})
    out = out[out["Name"] != ""].copy()
    for i, r in out.iterrows():
        if not r["Platform"]: out.at[i, "Platform"] = platform_from_id(r["TitleID"])
        if not r["Region"]: out.at[i, "Region"] = region_from_url(r["StoreURL"], "UNKNOWN")
        if not r["Tab"]: out.at[i, "Tab"] = PURCHASED_TAB
        if not r["Subscription"]: out.at[i, "Subscription"] = "NONE"
    return out.reset_index(drop=True)
def read_rb4_csv(path):
    df = pd.read_csv(path, encoding="utf-8-sig", dtype=str).fillna("")
    return pd.DataFrame({"Name": df["Name"], "Platform": "", "TitleID": "", "ProductID": df["productID"], "ConceptID": "", "Subscription": "NONE",
        "StoreURL": "", "Image": "", "Region": "UNKNOWN", "Account": "", "Locale": "", "Tab": PURCHASED_TAB, "ReportedTotal": "", "ScannedAt": TODAY})
def read_tx_csv(path):
    df = pd.read_csv(path, encoding="utf-8-sig", dtype=str).fillna("")
    low = {c.lower().replace(" ", ""): c for c in df.columns}
    g = lambda k: df[low[k]] if k in low else pd.Series([""] * len(df), index=df.index)
    return pd.DataFrame({"DateRaw": g("date"), "Name": g("name"), "AmountRaw": g("amount"), "OrderID": g("orderid"),
                         "Locale": g("locale"), "Region": g("region"), "Account": g("account"), "Details": g("details").map(redact)})
def read_trophy_csv(path):
    df = pd.read_csv(path, encoding="utf-8-sig", dtype=str).fillna("")
    cols = {re.sub(r"[^a-z]", "", c.lower()): c for c in df.columns}
    name = next((cols[k] for k in ("name", "game", "title", "gamename", "gametitle") if k in cols), None)
    prog = next((cols[k] for k in ("progress", "completion", "percent", "percentage", "complete") if k in cols), None)
    plat = next((cols[k] for k in ("platinum", "platinumearned", "hasplatinum", "plat") if k in cols), None)
    if not name: raise SystemExit(f"{path}: no game-name column found. Columns: {list(df.columns)}")
    out = []
    for _, r in df.iterrows():
        m = re.search(r"[\d.]+", r[prog]) if prog else None
        out.append({"label": r[name], "progress": float(m.group()) if m else None, "platinum": truthy(r[plat]) if plat else False})
    return out
def read_psnp_html(path):
    try: from bs4 import BeautifulSoup
    except ImportError: raise SystemExit("--psnp-html needs:  pip install beautifulsoup4")
    so = BeautifulSoup(open(path, encoding="utf-8", errors="ignore").read(), "html.parser")
    out = []
    for r in so.select("#gamesTable tr"):
        t = r.select_one("a.title")
        if not t or re.match(r"show \d+ more", t.get_text(strip=True), re.I): continue
        pg = r.select_one(".progress-bar span")
        m = re.search(r"[\d.]+", pg.get_text()) if pg else None
        out.append({"label": t.get_text(strip=True), "progress": float(m.group()) if m else None,
                    "platinum": r.select_one("img.platinum-18.earned, .icon-sprite.platinum-18.earned") is not None})
    if not out: print(f"  WARNING: [{os.path.basename(path)}] no games found - is this a saved PSNProfiles profile page with the games list loaded?")
    else: print(f"  [{os.path.basename(path)}] {len(out)} games read from the saved PSNProfiles page")
    return out

# ---- order breakdown ("Details") --------------------------------------------------------------------------
PRICE = re.compile(r"^[+\-\u2212]?\s*[^\d\s]{0,3}\s?\d[\d.,]*\s?(?:TL|\u20ba|GBP|USD|EUR)?$")
STOP = re.compile(r"^(discount|total\b|subtotal\b|tax\b|wallet balance\b|payment method\b|charge\b|request refund\b|refund\b)", re.I)
LABELS = re.compile(r"^(date|amount|description|type|account|order( number| date)?|transaction( id| amount| date)?|total|subtotal|discount|payment( method)?|items?|price|quantity|qty|details?|refund|request refund)\b", re.I)
def parse_details(raw, order_amount):
    """Item prices from the DETAILS section of an expanded order. Payment lines ('Charge wallet', 'Charge MC'), 'Discount' and 'Total' are ignored.
    Item prices are what you paid; a trailing 'Discount' line is only the saving versus list price."""
    toks = [x.strip().replace("\u00a0", " ") for x in redact(raw).replace("|", "\n").splitlines() if x.strip()]
    start = next((i + 1 for i, x in enumerate(toks) if x.upper() == "DETAILS"), 0)
    sub, items, i, disc = toks[start:], [], 0, 0.0
    while i < len(sub):
        tok = sub[i]
        if STOP.match(tok):
            if re.match(r"discount", tok, re.I) and i + 1 < len(sub) and PRICE.match(sub[i + 1]): disc = abs(parse_amount(sub[i + 1])[1] or 0)
            break
        if i + 1 < len(sub) and PRICE.match(sub[i + 1]) and not PRICE.match(tok) and not LABELS.match(tok):
            val = parse_amount(sub[i + 1])[1]
            if val is not None and val >= 0: items.append((tok, val)); i += 2; continue
        i += 1
    if not items or order_amount is None: return []
    total = sum(v for _, v in items); amt = abs(order_amount)
    if abs(total - amt) < 0.05: return items
    if disc and abs(total - disc - amt) < 0.05 and total > 0: return [(n, v * amt / total) for n, v in items]     # prices were list prices
    return []

def sheet_by_headers(wb, must, forbid=()):
    for ws in wb.worksheets:
        hdr = [s(c.value) for c in next(ws.iter_rows(min_row=1, max_row=1))] if ws.max_row else []
        if all(m in hdr for m in must) and not any(f in hdr for f in forbid): return ws, hdr
    return None, []
def sheet_rows(ws, hdr):
    for row in ws.iter_rows(min_row=2, values_only=True):
        d = {h: row[i] if i < len(row) else None for i, h in enumerate(hdr) if h}
        if any(s(v) for v in d.values()): yield d

# ----------------------------------------------------------------------------- state
EDITION_COLS = ["Region", "Account", "Platform", "Name", "Type", "Title ID", "Product ID", "Concept ID", "Access (tile)",
                "Store link", "Image URL", "First seen", "Last seen", "Status", "Auto type"]

class State:
    def __init__(self):
        self.editions, self.user, self.tx, self.scans, self.changes, self.legacy_tx = [], {}, [], [], [], []
        self.warnings, self.type_overrides, self.base = [], {}, ""

    def load_workbook_file(self, path):
        wb = load_workbook(path, data_only=False)
        st_ws, st_h = sheet_by_headers(wb, ["Setting", "Value"])
        if st_ws is not None:
            for d in sheet_rows(st_ws, st_h):
                if s(d.get("Setting")).lower().startswith("base currency"): self.base = s(d.get("Value")).upper()
        ed_ws, ed_h = sheet_by_headers(wb, ["Title ID", "Product ID", "First seen"])
        lib_ws, lib_h = sheet_by_headers(wb, ["Name"], forbid=("Title ID",))
        if lib_ws is not None and not any(h in lib_h for h in ("Played?", "Status", "Played")): lib_ws = None
        if lib_ws is not None:
            for d in sheet_rows(lib_ws, lib_h):
                nm = s(d.get("Name"))
                if not nm: continue
                u = {}
                if s(d.get("Status")): u["Status"] = s(d["Status"])
                if s(d.get("Priority")): u["Priority"] = s(d["Priority"])
                if s(d.get("Notes")): u["Notes"] = s(d["Notes"])
                if s(d.get("Platinum")).lower() == "yes": u["Platinum"] = True
                if flt(d.get("Trophy %")) is not None: u["Trophy"] = flt(d["Trophy %"])
                rk = next((h for h in lib_h if h.startswith("Rating")), None)
                if rk and flt(d.get(rk)) is not None: u["Rating"] = flt(d[rk])
                if truthy(d.get("Platinum'd?")): u["Platinum"] = True
                if truthy(d.get("Completed?")) or truthy(d.get("Platinum'd?")): u.setdefault("Status", "Finished")
                elif truthy(d.get("Played?")): u.setdefault("Status", "Playing")
                if u: self.user[norm(nm)] = u
                if ed_ws is None:
                    sub = s(d.get("Subscription"))
                    self.editions.append({**{c: "" for c in EDITION_COLS}, "Region": s(d.get("Region")) or "UNKNOWN", "Platform": s(d.get("Platform")), "Name": nm,
                        "Type": classify_library(nm), "Access (tile)": "PS_PLUS" if sub.upper() == "PS_PLUS" else "NONE", "Image URL": s(d.get("Image")) or s(d.get("Image URL")),
                        "First seen": "legacy", "Last seen": "legacy", "Status": "Present", "Auto type": classify_library(nm)})
        if ed_ws is not None:
            for d in sheet_rows(ed_ws, ed_h):
                base = {c: s(d.get(c)) for c in EDITION_COLS}
                if not base["Type"]: base["Type"] = classify_library(base["Name"])
                if not base["Auto type"]: base["Auto type"] = base["Type"]
                if "Regions" in ed_h:
                    present = [r.strip() for r in s(d.get("Regions")).split(",") if r.strip()]
                    missing = [r.strip() for r in s(d.get("Missing from")).split(",") if r.strip()]
                    accts = {}
                    for part in [p.strip() for p in s(d.get("Accounts")).split(";") if p.strip()]:
                        rg, _, an = part.partition(":"); accts.setdefault(rg.strip(), []).append(an.strip())
                    for reg in present + missing:
                        for an in (accts.get(reg) or [""]):
                            e = dict(base); e["Region"] = reg; e["Account"] = an; e["Status"] = "Present" if reg in present else "Missing"
                            e["Store link"] = link_for_region(base["Store link"], reg); self.editions.append(e)
                else: self.editions.append(base)
        tx_ws, tx_h = sheet_by_headers(wb, ["Date", "Name", "Order ID", "Category"])
        if tx_ws is not None:
            amt_col = "Amount (original)" if "Amount (original)" in tx_h else "Amount"
            for d in sheet_rows(tx_ws, tx_h):
                cur = ISO.get(s(d.get("Currency")), s(d.get("Currency")))
                self.tx.append({"Date": parse_date(d.get("Date")), "Name": s(d.get("Name")), "Amount": flt(d.get(amt_col)), "Currency": cur, "Order ID": order_id(d.get("Order ID")),
                                "Region": tx_region(cur, s(d.get("Region"))), "Account": s(d.get("Account")), "Details": redact(d.get("Details (raw)"))})
        old_tx, old_h = sheet_by_headers(wb, ["Name", "Amount", "Currency", "Subscription"])
        if old_tx is not None:
            for d in sheet_rows(old_tx, old_h):
                if s(d.get("Name")) in ("", "Name"): continue
                self.legacy_tx.append({"Date": "", "Name": s(d["Name"]), "Amount": flt(d["Amount"]), "Currency": ISO.get(s(d["Currency"]), s(d["Currency"])), "Order ID": "", "Region": "", "Account": "", "Details": ""})
        sc_ws, sc_h = sheet_by_headers(wb, ["Scan date", "Region", "Complete?"])
        if sc_ws is not None:
            for d in sheet_rows(sc_ws, sc_h): self.scans.append({k: s(d.get(k)) for k in ("Scan date", "Region", "Account", "Tab", "Rows captured", "Reported total", "Complete?", "Source file")})
        ch_ws, ch_h = sheet_by_headers(wb, ["Run date", "Change", "Name"])
        if ch_ws is not None:
            for d in sheet_rows(ch_ws, ch_h): self.changes.append({k: s(d.get(k)) for k in ("Run date", "Change", "Name", "Platform", "Region", "Account", "Access", "Title ID")})

    def find_edition(self, r):
        same = lambda e: e["Region"] == r["Region"] and e["Account"] in ("", r["Account"])
        for e in self.editions:
            if same(e) and r["ProductID"] and e["Product ID"] == r["ProductID"] and e["Title ID"] == r["TitleID"]: return e
        if r["TitleID"]:
            for e in self.editions:
                if same(e) and e["Title ID"] == r["TitleID"]: return e
        for e in self.editions:
            if same(e) and not e["Title ID"] and norm(e["Name"]) == norm(r["Name"]) and (not e["Platform"] or not r["Platform"] or e["Platform"] == r["Platform"]): return e
        for e in self.editions:
            if e["Region"] in ("UNKNOWN", "") and not e["Title ID"] and norm(e["Name"]) == norm(r["Name"]) and e["Platform"] == r["Platform"]:
                e["Region"] = r["Region"]; return e
        return None

    def merge_scan(self, df, source, partial):
        region, account, tab = df["Region"].iloc[0], df["Account"].iloc[0], df["Tab"].iloc[0]
        total = next((int(x) for x in df["ReportedTotal"] if s(x).isdigit()), None)
        date = (s(df["ScannedAt"].iloc[0])[:10]) or TODAY
        complete = (not partial) and (total is None or len(df) >= total)
        if tab != PURCHASED_TAB:
            print(f"  [{source}] tab '{tab}' is not used any more - ignored ({len(df)} rows)"); return
        self.scans = [x for x in self.scans if (x["Scan date"], x["Region"], x["Account"], x["Tab"], x["Source file"]) != (date, region, account, tab, source)]
        self.scans.append({"Scan date": date, "Region": region, "Account": account, "Tab": tab, "Rows captured": len(df),
            "Reported total": total if total is not None else "", "Complete?": "Yes" if complete else "No", "Source file": source})
        had_any = any(e["Region"] == region and e["Account"] in ("", account) for e in self.editions)
        seen, added_rows = set(), []
        for _, r in df.iterrows():
            e = self.find_edition(r)
            if e is None:
                e = {c: "" for c in EDITION_COLS}; e["First seen"] = date; self.editions.append(e)
                e["Type"] = e["Auto type"] = classify_library(r["Name"]); added_rows.append(r)
            elif e["Status"] == "Missing":
                self.changes.append({"Run date": date, "Change": "Returned", "Name": r["Name"], "Platform": r["Platform"], "Region": region, "Account": account, "Access": e["Access (tile)"], "Title ID": r["TitleID"]})
            if not e["Type"]: e["Type"] = e["Auto type"] = classify_library(r["Name"])
            e.update({"Region": region, "Account": account or e["Account"], "Platform": r["Platform"] or e["Platform"], "Name": r["Name"], "Title ID": r["TitleID"],
                      "Product ID": r["ProductID"], "Concept ID": r["ConceptID"] or e["Concept ID"], "Access (tile)": r["Subscription"],
                      "Store link": r["StoreURL"] or e["Store link"], "Image URL": r["Image"] or e["Image URL"], "Last seen": date, "Status": "Present"})
            seen.add(id(e))
        if had_any or len(added_rows) < 25:
            for r in added_rows:
                self.changes.append({"Run date": date, "Change": "Added", "Name": r["Name"], "Platform": r["Platform"], "Region": region, "Account": account, "Access": "", "Title ID": r["TitleID"]})
        elif added_rows:
            self.changes.append({"Run date": date, "Change": "Initial import", "Name": f"{len(added_rows)} titles", "Platform": "", "Region": region, "Account": account, "Access": "", "Title ID": ""})
        missing = 0
        if complete:
            for e in self.editions:
                if e["Region"] == region and e["Account"] in ("", account) and id(e) not in seen and e["Status"] != "Missing":
                    e["Status"] = "Missing"; missing += 1
                    self.changes.append({"Run date": date, "Change": "Missing", "Name": e["Name"], "Platform": e["Platform"], "Region": region, "Account": account, "Access": e["Access (tile)"], "Title ID": e["Title ID"]})
        print(f"  [{source}] {acct_label(region, account)}: {len(df)} rows, {len(added_rows)} new, {missing} now missing, complete={complete}")

    def merge_tx(self, df, source):
        dayfirst = infer_dayfirst(df["DateRaw"], next((l for l in df["Locale"] if l), ""))
        key = lambda t: (t["Date"], t["Account"], norm(t["Name"]), abs(t["Amount"] or 0), t["Order ID"])
        book = {key(t): t for t in self.tx}; new = 0; regs = {}
        for _, r in df.iterrows():
            cur, amt = parse_amount(r["AmountRaw"])
            reg = s(r["Region"]) or region_from_locale(r["Locale"])
            if not cur and amt is not None: cur = REGION_CURRENCY.get(reg, "")
            row = {"Date": parse_date(r["DateRaw"], dayfirst), "Name": s(r["Name"]), "Amount": amt, "Currency": cur, "Order ID": order_id(r["OrderID"]),
                   "Region": tx_region(cur, reg), "Account": s(r["Account"]), "Details": s(r["Details"])}
            lab = acct_label(row["Region"], row["Account"]); regs[lab] = regs.get(lab, 0) + 1
            if key(row) not in book: new += 1
            elif not row["Details"]: row["Details"] = book[key(row)].get("Details", "")
            book[key(row)] = row
        self.tx = list(book.values())
        print(f"  [{source}] transactions: {len(df)} rows, {new} new, accounts {regs}, dates read {'day-first' if dayfirst else 'month-first'}")

# ----------------------------------------------------------------------------- one-off: unify old free-text labels with PSN online IDs
def plan_account_map(st, frames):
    seen = defaultdict(lambda: {"online": set(), "legacy": set()})
    def note(region, acc):
        if not region or region == "UNKNOWN": return
        seen[region]["online" if is_online_id(acc) else "legacy"].add(s(acc))
    for e in st.editions: note(e["Region"], e["Account"])
    for t in st.tx: note(t["Region"], t["Account"])
    for x in st.scans: note(x["Region"], x["Account"])
    for df in frames:
        if "Region" in df.columns and "Account" in df.columns:
            for reg, acc in set(zip(df["Region"], df["Account"])):
                if reg: note(s(reg), s(acc))
    mapping = {}
    for region, d in seen.items():
        onl, leg = d["online"], d["legacy"]
        if len(onl) == 1 and leg:
            for a in leg: mapping[(region, a)] = next(iter(onl))
        elif not onl and len([a for a in leg if a]) >= 1:
            named = [a for a in leg if a]
            main = [a for a in named if not a.lower().endswith("transactions")]
            if len(main) == 1:
                for a in leg:
                    if a != main[0]: mapping[(region, a)] = main[0]
    return mapping
def apply_account_map(st, frames, mapping):
    if not mapping: return
    for e in st.editions: e["Account"] = mapping.get((e["Region"], e["Account"]), e["Account"])
    for t in st.tx: t["Account"] = mapping.get((t["Region"], t["Account"]), t["Account"])
    for x in st.scans: x["Account"] = mapping.get((x["Region"], x["Account"]), x["Account"])
    for c in st.changes: c["Account"] = mapping.get((c["Region"], c["Account"]), c["Account"])
    for df in frames:
        if "Account" in df.columns and "Region" in df.columns:
            df["Account"] = [mapping.get((s(r), s(a)), s(a)) for r, a in zip(df["Region"], df["Account"])]
    for (reg, old), new in sorted(mapping.items()):
        if old != new: print(f"  account labels unified: {reg}: '{old or '(blank)'}' -> '{new}'")

# ----------------------------------------------------------------------------- extra user files
def read_rows(path):
    rows = []
    for line in open(path, encoding="utf-8-sig").read().splitlines()[1:]:
        parts = re.split(r"[;\t]", line) if (";" in line or "\t" in line) else line.split(",")
        parts = [x.strip() for x in parts if x.strip()]
        if parts: rows.append(parts)
    return rows
def load_extra_files(here, st):
    p = os.path.join(here, "collections.csv")
    if os.path.exists(p):
        for parts in read_rows(p):
            if len(parts) >= 2: COLLECTIONS[coll_key(parts[0])] = [base_key(c) for c in parts[1:]]
        print(f"  collections.csv: {len(COLLECTIONS)} collections known")
    p = os.path.join(here, "aliases.csv")
    if os.path.exists(p):
        for parts in read_rows(p):
            if len(parts) >= 3: ALIASES[(parts[0].upper(), norm(parts[1]))] = parts[2]
            elif len(parts) == 2: ALIASES[("*", norm(parts[0]))] = parts[1]
        print(f"  aliases.csv: {len(ALIASES)} renames")
    p = os.path.join(here, "trophy_aliases.csv")
    if os.path.exists(p):
        for parts in read_rows(p):
            if len(parts) >= 2: TROPHY_ALIASES[norm(parts[0])] = parts[1]
        print(f"  trophy_aliases.csv: {len(TROPHY_ALIASES)} entries")
    p = os.path.join(here, "type_overrides.csv")
    if os.path.exists(p):
        for parts in read_rows(p):
            if len(parts) >= 2: st.type_overrides[norm(parts[0])] = parts[1]
def display_name(e): return ALIASES.get((e["Region"], norm(e["Name"])), ALIASES.get(("*", norm(e["Name"])), e["Name"]))

# ----------------------------------------------------------------------------- matching with an ambiguity guard
def split_item(text):
    t = re.sub(r"^\s*bundle\s*:\s*", "", s(text), flags=re.I)
    parts = [p.strip() for p in re.split(r"\s\+\s", t) if p.strip()]
    if len(parts) > 1 and ":" in parts[0]:
        prefix = parts[0].split(":")[0].strip(); parts = [parts[0]] + [p if ":" in p else f"{prefix}: {p}" for p in parts[1:]]
    return parts or [t]
def split_basket(name, known):
    parts = [p.strip() for p in s(name).split(", ")]; items, i = [], 0
    while i < len(parts):
        for j in range(len(parts), i, -1):
            cand = ", ".join(parts[i:j])
            if j - i == 1 or keys_of(cand) & known: items.extend(split_item(cand)); i = j; break
    return items

class Index:
    """key -> set of game ids. A tx / trophy name only attaches to a game when exactly ONE game fits (after filtering by region)."""
    def __init__(self, groups, disp):
        self.owner, self.regions, self.names = defaultdict(set), {}, {}
        for gid, eds in groups.items():
            self.regions[gid] = {e["Region"] for e in eds if e["Status"] != "Missing"} or {e["Region"] for e in eds}
            for e in eds:
                for n in (e["Name"], disp(e)):
                    for k in keys_of(n): self.owner[k].add(gid)
            self.names[gid] = disp(eds[0])
        self.known = set(self.owner)
    def resolve(self, name, region=None):
        """-> (gid or None, how). how: exact / loose / weak / ambiguous / none"""
        for kind, keys in (("exact", {base_key(name)}), ("loose", {loose_key(name)}), ("weak", weak_keys(name))):
            gids = set()
            for k in keys - {""}: gids |= self.owner.get(k, set())
            if not gids: continue
            use = [g for g in gids if not region or region in self.regions[g]] or list(gids)
            return (use[0], kind) if len(use) == 1 else (None, "ambiguous")
        return None, "none"

def analyse_transactions(st, groups, disp):
    idx = Index(groups, disp)
    type_by_gid = {gid: eds[0]["Type"] for gid, eds in groups.items()}
    evidence, ambiguous = defaultdict(list), []
    known = idx.known
    for t in st.tx:
        t["Refund"] = bool(re.match(r"refund for\s", s(t["Name"]), re.I))
        eff = re.sub(r"^refund for\s+", "", s(t["Name"]), flags=re.I)
        t["Breakdown"] = parse_details(t.get("Details", ""), t["Amount"]) if t.get("Details") else []
        raw_items = [(n, v) for n, v in t["Breakdown"]] if t["Breakdown"] else [(n, None) for n in (split_basket(eff, known) if eff else [])]
        t["Items"] = [n for n, _ in raw_items]; t["Basket"] = len(raw_items) > 1
        cat = classify_tx_name(eff) if len(raw_items) <= 1 else ""
        if t["Refund"]: cat = "Refund"
        elif len(raw_items) > 1: cat = "Bundle / basket"
        t["Category"] = cat; t["Matched"] = []
        n_items = max(1, len(raw_items)); total = t["Amount"]; first_gid = None
        for n, price in raw_items:
            if price is not None and total: frac, basis = price / total, "Order breakdown"
            else: frac, basis = 1.0 / n_items, ("Paid" if n_items == 1 else "Bundle / basket split (est.)")
            gid, how = idx.resolve(n, t["Region"])
            if gid is None and how == "none" and coll_key(n) in COLLECTIONS:
                comps = COLLECTIONS[coll_key(n)]; hit = 0
                for c in comps:
                    g2 = {g for g in idx.owner.get(c, set()) if not t["Region"] or t["Region"] in idx.regions[g]} or set(idx.owner.get(c, set()))
                    if len(g2) == 1: evidence[next(iter(g2))].append((t, frac / len(comps), "Collection split")); hit += 1
                    elif len(g2) > 1: ambiguous.append((n, t, "collection component '" + c + "' fits several games"))
                t["Matched"].append(f"{n} -> {hit} of {len(comps)} games in library"); continue
            if gid is None:
                if how == "ambiguous": ambiguous.append((n, t, "name fits more than one game"))
                continue
            t["Matched"].append(n if how == "exact" else f"{n} ({how} match)"); first_gid = first_gid or gid
            evidence[gid].append((t, frac, basis))
        if not cat:
            cat = "DLC / other (not a library game)"
            if first_gid is not None: cat = "Game (in library)" if type_by_gid.get(first_gid) == "Game" else type_by_gid.get(first_gid, cat)
            t["Category"] = cat
    return evidence, ambiguous, idx

def access_for(tile_plus, evid):
    net = sum((t["Amount"] or 0) * f for t, f, b in evid if t["Category"] != "Subscription")
    if net > 0.005: return "Purchased"
    if tile_plus: return "PS Plus"
    if evid: return "Free / claimed"
    return "Free (assumed)"

def find_collisions(st):
    """Same name on the same platform with very different title numbers (e.g. RE4 2005 vs RE4 2023) -> probably two games."""
    by = defaultdict(set)
    for e in st.editions:
        n = title_number(e["Title ID"])
        if n is not None and e["Platform"]: by[(norm(e["Name"]), e["Platform"])].add((n, e["Title ID"], e["Region"], e["Name"]))
    out = []
    for (k, plat), v in by.items():
        nums = sorted({x[0] for x in v})
        if len(nums) >= 2 and nums[-1] - nums[0] > 15000 and not any((x[2], k) in ALIASES or ("*", k) in ALIASES for x in v):
            out.append((sorted(v)[0][3], plat, ", ".join(sorted(f"{x[2]}:{x[1]}" for x in v))))
    return out

# ----------------------------------------------------------------------------- library
def refresh_types(st):
    for e in st.editions:
        auto = classify_library(e["Name"]); ov = st.type_overrides.get(norm(e["Name"]))
        if ov: e["Type"] = ov
        elif not e["Type"] or e["Type"] == e["Auto type"]: e["Type"] = auto
        e["Auto type"] = auto

def build_library(st, trophies):
    refresh_types(st)
    groups = defaultdict(list)
    for e in st.editions:
        if e["Name"]: groups[base_key(display_name(e)) or norm(display_name(e))].append(e)
    user_of = lambda eds: next((st.user[k] for e in eds for k in (norm(e["Name"]), norm(display_name(e))) if k in st.user), None)
    lib_gids = [g for g, eds in groups.items() if any(e["Type"] in GAME_TYPES for e in eds) or user_of(eds)]
    evidence, ambiguous, idx = analyse_transactions(st, groups, display_name)
    # trophies -> games (ambiguity guard; trilogy children via trophy_aliases.csv are averaged, plain duplicates take the max)
    plain, roll, unmatched = defaultdict(list), defaultdict(list), []
    disp_index = {norm(display_name(e)): gid for gid, eds in groups.items() for e in eds}       # exact display names win (needed for split games)
    for r in trophies:
        tgt = TROPHY_ALIASES.get(norm(r["label"]))
        if tgt:
            gid, how = (disp_index.get(norm(tgt)), "alias") if norm(tgt) in disp_index else idx.resolve(tgt)
            (roll[gid].append(r) if gid else unmatched.append((r["label"], f"alias target '{tgt}' not found")))
            continue
        gid, how = idx.resolve(r["label"])
        if gid: plain[gid].append(r)
        else: unmatched.append((r["label"], "fits more than one game - add it to trophy_aliases.csv" if how == "ambiguous" else "no matching game in the library"))
    rows, price_map = [], {}
    for gid in lib_gids:
        eds = sorted(groups[gid], key=lambda e: (e["Platform"] != "PS5", e["Region"] != "UK"))
        name = min((display_name(e) for e in eds), key=lambda x: (len(x), x))
        plats = sorted({p.strip() for e in eds for p in e["Platform"].split(",") if p.strip()})
        evid = evidence.get(gid, [])
        access = access_for(any(e["Access (tile)"] == "PS_PLUS" for e in eds), evid)
        dated = sorted([(t["Date"], t, f, b) for t, f, b in evid if t["Date"] and t["Category"] not in ("Subscription", "Refund")], key=lambda x: x[0])
        paid = [(d, t, f, b) for d, t, f, b in dated if (t["Amount"] or 0) > 0]
        u = user_of(eds) or {}
        progs = [x["progress"] for x in plain[gid] if x["progress"] is not None]
        rp = [x["progress"] for x in roll[gid] if x["progress"] is not None]
        if rp: progs.append(sum(rp) / len(rp))
        if "Trophy" in u: progs.append(u["Trophy"])
        plat = bool(u.get("Platinum")) or any(x["platinum"] for x in plain[gid] + roll[gid])
        present = sorted({e["Region"] for e in eds if e["Region"] and e["Status"] != "Missing"} or {e["Region"] for e in eds if e["Region"]})
        if paid and access == "Purchased": price_map[gid] = (paid[0][1], paid[0][2], paid[0][3])
        rows.append({"Name": name, "Platforms": " + ".join(plats), "Access": access, "Status": u.get("Status", ""), "Priority": u.get("Priority", ""),
            "Trophy %": max(progs) if progs else "", "Platinum": "Yes" if plat else "", "Rating": u.get("Rating", ""), "Acquired": dated[0][0] if dated else "",
            "PriceBasis": price_map[gid][2] if gid in price_map else "", "Notes": u.get("Notes", ""), "Regions": ", ".join(present),
            "Store link": next((link_for_region(e["Store link"], "UK") if "UK" in present and e["Store link"] else e["Store link"] for e in eds if e["Store link"]), ""),
            "Image URL": next((e["Image URL"] for e in eds if e["Image URL"]), ""), "_key": gid})
    lib = pd.DataFrame(rows).sort_values("Name", key=lambda x: x.str.lower()).reset_index(drop=True)
    return lib, price_map, ambiguous, unmatched, idx

def review_rows(st, lib, ambiguous, idx):
    out = []
    for n, t, why in ambiguous: out.append(["(ambiguous)", "", n, t["Date"], t["Amount"], t["Currency"], acct_label(t["Region"], t["Account"]), "", why])
    for name, plat, ids in find_collisions(st):
        out.append([name, "", "same name, very different title numbers", "", "", "", ids, "", "Possibly TWO games (e.g. original vs remake). Add a line to aliases.csv: Region;Name;DisplayName"])
    skip = {"In-game item / pack", "Add-on", "Theme", "Avatar", "Demo", "Extra", "Soundtrack", "Subscription", "Refund", "Wallet top-up"}
    items = []
    for t in st.tx:
        if t["Category"] in skip: continue
        for n in t.get("Items", []) or ([t["Name"]] if t["Name"] else []):
            if not any(norm(n) == norm(m.split(" (")[0]) for m in t.get("Matched", [])) and not re.search(r"\b(pack|episode|dlc|season pass|ticket|currency)\b", n, re.I): items.append((n, t))
    import difflib
    for _, r in lib[lib["Access"] == "Free (assumed)"].iterrows():
        lk = loose_key(r["Name"]); best = []
        for n, t in items:
            ik = loose_key(n)
            if not ik or not lk or t["Category"] == "Refund": continue
            ratio = difflib.SequenceMatcher(None, lk, ik).ratio()
            contained = (len(lk) >= 6 and lk in ik) or (len(ik) >= 6 and ik in lk)
            collection = bool(re.search(r"collection|trilogy|anthology|saga|bundle", n, re.I)) and len(lk) >= 5 and re.sub(r"\d+$", "", lk)[:6] in ik
            if ratio >= 0.82 or contained or collection: best.append((ratio, n, t, "Collection / bundle?" if collection else "Similar name"))
        best.sort(key=lambda x: -x[0])
        for ratio, n, t, why in best[:2]: out.append([r["Name"], r["Access"], n, t["Date"], t["Amount"], t["Currency"], acct_label(t["Region"], t["Account"]), round(ratio, 2), why])
    return out

# ----------------------------------------------------------------------------- workbook
HEAD_FILL = PatternFill("solid", fgColor="0B1E3A")
def style_header(ws, widths, freeze="A2"):
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF"); c.fill = HEAD_FILL; c.alignment = Alignment(wrap_text=True, vertical="center")
    for i, w in enumerate(widths, 1): ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = freeze
def as_date(v): return dt.date.fromisoformat(v) if isinstance(v, str) and re.match(r"\d{4}-\d{2}-\d{2}$", v) else (v or None)

def merged_editions(st):
    groups = defaultdict(list)
    for e in st.editions:
        k = (e["Title ID"], e["Product ID"]) if e["Title ID"] and e["Product ID"] else (e["Region"], e["Account"], norm(e["Name"]), e["Platform"])
        groups[k].append(e)
    rows = []
    for k, eds in groups.items():
        present = sorted({e["Region"] for e in eds if e["Status"] != "Missing"}); missing = sorted({e["Region"] for e in eds if e["Status"] == "Missing"} - set(present))
        f = sorted(eds, key=lambda e: (e["Region"] != "UK", e["Status"] == "Missing"))[0]
        link = link_for_region(f["Store link"], "UK") if f["Store link"] and (set(present) | set(missing)) & {"UK"} else f["Store link"]
        accts = sorted({f"{e['Region']}:{e['Account']}" for e in eds if e["Account"]})
        rows.append({"Platform": f["Platform"], "Name": f["Name"], "Type": f["Type"], "Title ID": f["Title ID"], "Product ID": f["Product ID"],
            "Concept ID": next((e["Concept ID"] for e in eds if e["Concept ID"]), ""), "Regions": ", ".join(present), "Missing from": ", ".join(missing), "Accounts": "; ".join(accts),
            "Access (tile)": "PS_PLUS" if any(e["Access (tile)"] == "PS_PLUS" for e in eds) else "NONE", "Store link": link,
            "Image URL": next((e["Image URL"] for e in eds if e["Image URL"]), ""),
            "First seen": min([e["First seen"] for e in eds if e["First seen"]] or [""]), "Last seen": max([e["Last seen"] for e in eds if e["Last seen"]] or [""]), "Auto type": f["Auto type"]})
    return sorted(rows, key=lambda r: (r["Name"].lower(), r["Platform"]))

def wallet_rows(st):
    """Wallet ledger = 'Wallet funding' top-ups minus the 'Charge wallet' part of each order's payment breakdown (same-day top-ups are counted first)."""
    rows = []
    for t in st.tx:
        if not t.get("Date"): continue
        name, det = s(t["Name"]), s(t.get("Details", "")); cur = t.get("Currency") or REGION_CURRENCY.get(t["Region"], "")
        if re.match(r"wallet funding", name, re.I) and (t["Amount"] or 0) > 0:
            rows.append((acct_label(t["Region"], t["Account"]), t["Date"], 0, "Top-up", name, abs(t["Amount"]), cur, t.get("Order ID", "")))
        m = re.search(r"Charge wallet \| ([^|]+)", det)
        if m:
            c2, v = parse_amount(m.group(1))
            if v: rows.append((acct_label(t["Region"], t["Account"]), t["Date"], 1, "Spend", name[:80], -abs(v), c2 or cur, t.get("Order ID", "")))
    return sorted(rows, key=lambda r: (r[0], r[1], r[2], r[4]))

def write_wallet(wb, st):
    rows = wallet_rows(st)
    if not rows: return
    ws = wb.create_sheet("Wallet"); ws.append(["Wallet", "Date", "Type", "Description", "Amount", "Running balance", "Currency", "Order ID"])
    for i, r in enumerate(rows):
        x = i + 2; ws.append([r[0], as_date(r[1]), r[3], r[4], r[5], f"=IF(A{x}=A{x-1},F{x-1},0)+E{x}", r[6], r[7] or None])
        ws.cell(x, 2).number_format = "yyyy-mm-dd"; ws.cell(x, 5).number_format = ws.cell(x, 6).number_format = "#,##0.00"
    n = len(rows) + 1; A_, C_, E_, F_ = (f"${c}$2:${c}${n}" for c in "ACEF")
    ws.cell(1, 10, "Wallet").font = ws.cell(1, 11, "Currency").font = Font(bold=True)
    for j, h in enumerate(["Wallet", "Currency", "Top-ups", "Spent from wallet", "Calculated balance", "Balance shown by PSN (enter)", "Difference", "Lowest running balance"]):
        c = ws.cell(1, 10 + j, h); c.font = Font(bold=True, color="FFFFFF"); c.fill = HEAD_FILL; c.alignment = Alignment(wrap_text=True, vertical="center")
    wallets = sorted({r[0] for r in rows})
    for k, w in enumerate(wallets):
        y = k + 2; cur = next(r[6] for r in rows if r[0] == w)
        ws.cell(y, 10, w); ws.cell(y, 11, cur)
        ws.cell(y, 12, f'=SUMIFS({E_},{A_},J{y},{C_},"Top-up")'); ws.cell(y, 13, f'=-SUMIFS({E_},{A_},J{y},{C_},"Spend")'); ws.cell(y, 14, f"=L{y}-M{y}")
        ws.cell(y, 16, f'=IF(O{y}="","",ROUND(N{y}-O{y},2))'); ws.cell(y, 17, f'=MINIFS({F_},{A_},J{y})')
        for col in (12, 13, 14, 15, 16, 17): ws.cell(y, col).number_format = "#,##0.00"
        ws.cell(y, 15).fill = PatternFill("solid", fgColor="FFF2CC")
    r0 = len(wallets) + 3
    for t_ in ["Enter the balance PlayStation shows for each wallet in the yellow cell: Difference should be 0.00.",
               "A negative 'Lowest running balance' means same-day order is uncertain (top-ups are counted before spends) or an older top-up is missing.",
               "Refunds are NOT added back: nothing in the export shows whether they went to the wallet or the card."]:
        ws.cell(r0, 10, t_).font = Font(italic=True, color="666666"); r0 += 1
    for col, w in zip("ABCDEFGHIJKLMNOPQ", [20, 12, 9, 46, 11, 14, 9, 22, 2, 20, 9, 11, 14, 14, 16, 11, 14]): ws.column_dimensions[col].width = w
    ws.row_dimensions[1].height = 32; ws.freeze_panes = "A2"; ws.auto_filter.ref = f"A1:H{n}"
    for c in ws[1][:8]: c.font = Font(bold=True, color="FFFFFF"); c.fill = HEAD_FILL; c.alignment = Alignment(wrap_text=True, vertical="center")

def write_workbook(st, lib, price_map, review, base, out):
    sym = SYMBOL.get(base, base + " "); FMT = f'"{sym}"#,##0.00'
    wb = Workbook(); sm = wb.active; sm.title = "Summary"
    stg = wb.create_sheet("Settings"); stg.append(["Setting", "Value", "Notes"])
    stg.append(["Base currency", base, "All money is converted to this. To change it: edit this cell OR run build_spreadsheet.py --base XXX, then rebuild."])
    stg.append(["Built on", TODAY, ""]); style_header(stg, [22, 12, 90])
    TC = ["Date", "Year", "Name", f"Amount ({base})", "Category", "Region", "Account", "Items in order", "Item prices (from breakdown)", "Matched library games",
          "Order ID", "Amount (original)", "Currency", "Details (raw, redacted)"]
    L_ = {n: get_column_letter(i + 1) for i, n in enumerate(TC)}; CB = L_[f"Amount ({base})"]
    tx = wb.create_sheet("Transactions"); tx.append(TC)
    txs = sorted(st.tx, key=lambda t: t["Date"], reverse=True); fx_needed = {}
    for i, t in enumerate(txs):
        r_ = i + 2; t["_row"] = r_; d = as_date(t["Date"]); cur = t["Currency"] or REGION_CURRENCY.get(t["Region"], base)
        if cur != base and t["Amount"] and d: fx_needed[(cur, t["Date"])] = None
        K, Lc = L_["Amount (original)"], L_["Currency"]
        f = f'=IF({K}{r_}="","",IF({K}{r_}=0,0,IF({Lc}{r_}="{base}",{K}{r_},IFERROR({K}{r_}*VLOOKUP({Lc}{r_}&"|"&TEXT(A{r_},"yyyy-mm-dd"),FX!$C$2:$D$5000,2,FALSE),"FX?"))))'
        tx.append([d, d.year if d else None, t["Name"], f, t["Category"], t["Region"] or None, t["Account"] or None,
                   " | ".join(t["Items"]) if t["Basket"] else None, " | ".join(f"{n}: {v:.2f}" for n, v in t["Breakdown"]) if t["Breakdown"] else None,
                   " | ".join(t["Matched"]) if t["Matched"] else None, t["Order ID"] or None, t["Amount"], cur, redact(t.get("Details")) or None])
        tx.cell(r_, 1).number_format = "yyyy-mm-dd"; tx.cell(r_, 11).number_format = "@"; tx.cell(r_, 4).number_format = FMT
    style_header(tx, [11, 6, 52, 13, 20, 7, 16, 40, 40, 40, 17, 12, 8, 20]); tx.auto_filter.ref = f"A1:{get_column_letter(len(TC))}{len(txs)+1}"
    for n in ("Amount (original)", "Currency", "Details (raw, redacted)"): tx.column_dimensions[L_[n]].hidden = True
    txn = len(txs) + 1
    fx = wb.create_sheet("FX"); fx.append(["Currency", "Date", "Key", f"Rate to {base} (GOOGLEFINANCE, falls back up to 3 days)"])
    for i, (cur, date) in enumerate(sorted(fx_needed, key=lambda x: (x[0], x[1])), 2):
        pair = f'"CURRENCY:{cur}{base}"'
        f = "=" + "".join(f'IFERROR(INDEX(GOOGLEFINANCE({pair},"price",B{i}-{k}),2,2),' for k in range(3)) + f'INDEX(GOOGLEFINANCE({pair},"price",B{i}-3),2,2)' + ")" * 3
        fx.append([cur, as_date(date), f"{cur}|{date}", f]); fx.cell(i, 2).number_format = "yyyy-mm-dd"
    style_header(fx, [10, 12, 18, 40])

    ws = wb.create_sheet("Library", 1)
    LC = ["Artwork", "Name", "Platforms", "Access", "Status", "Priority", "Played", "Completed", "Trophy %", "Platinum", "Rating (0-10)", "Acquired",
          f"Price paid ({base})", "Price basis", "Notes", "Regions", "Store", "Image URL"]
    IG = [c for c in ("Critic (IGDB)", "Users (IGDB)", "IGDB type", "Year", "IGDB match", "IGDB") if c in lib.columns]; LC += IG
    ws.append(LC); n = len(lib); last = n + 1000
    for i, r in lib.iterrows():
        x = i + 2; price = None
        if r["_key"] in price_map:
            ptx, frac, _b = price_map[r["_key"]]; price = f'=IFERROR(Transactions!${CB}${ptx["_row"]}*{frac:.6f},"")'
        ws.append([f'=IF(R{x}="","",IMAGE(R{x},1))', r["Name"], r["Platforms"], r["Access"], r["Status"] or None, r["Priority"] or None,
            f'=IF(OR(N(I{x})>0,E{x}="Playing",E{x}="Finished",E{x}="Dropped"),"Yes","")', f'=IF(OR(E{x}="Finished",N(I{x})>=100),"Yes","")',
            r["Trophy %"] if r["Trophy %"] != "" else None, r["Platinum"] or None, r["Rating"] if r["Rating"] != "" else None,
            as_date(r["Acquired"]), price, r["PriceBasis"] or None, r["Notes"] or None, r["Regions"],
            f'=HYPERLINK("{r["Store link"]}","Store")' if r["Store link"] else None, r["Image URL"] or None] + [(f'=HYPERLINK("{r[c]}","IGDB")' if c == "IGDB" and r[c] else (None if pd.isna(r[c]) or r[c] == "" else r[c])) for c in IG])
        ws.row_dimensions[x].height = 46; ws.cell(x, 12).number_format = "yyyy-mm-dd"; ws.cell(x, 13).number_format = FMT
    style_header(ws, [10, 40, 11, 17, 11, 10, 8, 10, 9, 9, 10, 12, 12, 20, 30, 11, 8, 20] + [10, 10, 13, 7, 12, 7][:len(IG)], "C2")
    ws.column_dimensions["R"].hidden = True; ws.auto_filter.ref = f"A1:{get_column_letter(len(LC))}{n+1}"
    if "IGDB match" in IG:
        gl = get_column_letter(LC.index("IGDB match") + 1)
        ws.conditional_formatting.add(f"{gl}2:{gl}{last}", FormulaRule(formula=[f'AND({gl}2<>"",{gl}2<>"OK",{gl}2<>"Override")'], fill=PatternFill("solid", bgColor="FCE4D6")))
    def dv(vals, rng):
        d = DataValidation(type="list", formula1='"' + ",".join(vals) + '"', allow_blank=True); ws.add_data_validation(d); d.add(rng)
    dv(STATUSES, f"E2:E{last}"); dv(PRIORITIES, f"F2:F{last}"); dv(["Yes"], f"J2:J{last}")
    d2 = DataValidation(type="decimal", operator="between", formula1="0", formula2="10", allow_blank=True); ws.add_data_validation(d2); d2.add(f"K2:K{last}")
    for rng, formula, color in [(f"E2:E{last}", 'E2="Finished"', "C6EFCE"), (f"E2:E{last}", 'E2="Playing"', "BDD7EE"), (f"E2:E{last}", 'E2="Dropped"', "F8CBAD"), (f"J2:J{last}", 'J2="Yes"', "FFE699")]:
        ws.conditional_formatting.add(rng, FormulaRule(formula=[formula], fill=PatternFill("solid", bgColor=color)))
    ed = wb.create_sheet("Editions"); EC = ["Platform", "Name", "Type", "Title ID", "Product ID", "Concept ID", "Regions", "Missing from", "Accounts", "Access (tile)", "Store link", "Image URL", "First seen", "Last seen", "Auto type"]
    ed.append(EC); mrows = merged_editions(st)
    for r in mrows: ed.append([as_date(r[c]) if c in ("First seen", "Last seen") else (r[c] or None) for c in EC])
    style_header(ed, [9, 40, 11, 15, 38, 11, 12, 12, 24, 13, 40, 25, 12, 12, 10], "C2")
    ed.column_dimensions["L"].hidden = True; ed.column_dimensions["O"].hidden = True; ed.auto_filter.ref = f"A1:O{len(mrows)+1}"
    d3 = DataValidation(type="list", formula1='"' + ",".join(ALL_TYPES) + '"', allow_blank=False); ed.add_data_validation(d3); d3.add(f"C2:C{len(mrows)+500}")
    ed.conditional_formatting.add(f"A2:O{len(mrows)+1}", FormulaRule(formula=['$H2<>""'], fill=PatternFill("solid", bgColor="FCE4D6")))
    rv = wb.create_sheet("Match review"); rv.append(["Library game", "Access now", "Transaction item / detail", "Date", "Amount (original)", "Currency", "Account / IDs", "Similarity", "Why"])
    for row in review: rv.append([row[0], row[1], row[2], as_date(row[3]) if row[3] else None, row[4] if row[4] != "" else None, row[5] or None, row[6], row[7] if row[7] != "" else None, row[8]])
    style_header(rv, [36, 16, 52, 12, 12, 9, 28, 9, 60])
    try:
        import igdb_enrich
        if igdb_enrich.REVIEW:
            ir = wb.create_sheet("IGDB review"); ir.append(["Library game", "Best IGDB guess", "IGDB type", "IGDB id", "Result", "Similarity", "Why", "Other candidates"])
            for row in igdb_enrich.REVIEW: ir.append(row)
            style_header(ir, [36, 36, 14, 10, 12, 10, 50, 70])
    except ImportError: pass
    write_wallet(wb, st)
    sc = wb.create_sheet("Scans"); SC = ["Scan date", "Region", "Account", "Tab", "Rows captured", "Reported total", "Complete?", "Source file"]; sc.append(SC)
    for x in sorted(st.scans, key=lambda x: x["Scan date"], reverse=True): sc.append([x.get(c, "") for c in SC])
    style_header(sc, [12, 8, 18, 20, 14, 14, 11, 40])
    ch = wb.create_sheet("Changes"); CH = ["Run date", "Change", "Name", "Platform", "Region", "Account", "Access", "Title ID"]; ch.append(CH)
    for x in sorted(st.changes, key=lambda x: x["Run date"], reverse=True): ch.append([x.get(c, "") for c in CH])
    style_header(ch, [12, 14, 45, 10, 8, 18, 18, 15])

    LL = lambda col: f"Library!${col}$2:${col}${last}"
    T_ = lambda name: f"Transactions!${L_[name]}$2:${L_[name]}${txn+2000}"
    AM, CAT, RG, AC, YR = T_(f"Amount ({base})"), T_("Category"), T_("Region"), T_("Account"), T_("Year")
    accounts = sorted({(t["Region"], t["Account"]) for t in st.tx if t["Region"]} | {(e["Region"], e["Account"]) for e in st.editions if e["Region"]})
    lines = [("PlayStation library tracker", None), ("Built on", TODAY), ("", None), ("COLLECTION", None),
        ("Games in library", f"=COUNTA({LL('B')})"),
        ("Non-game items kept on Editions (demos, apps, soundtracks...)", '=COUNTIF(Editions!$C$2:$C$5000,"<>Game")-COUNTBLANK(Editions!$C$2:$C$5000)'),
        ("Games on PS5 (incl. both)", f'=COUNTIF({LL("C")},"*PS5*")'), ("Games on PS4 (incl. both)", f'=COUNTIF({LL("C")},"*PS4*")'),
        ("Access: Purchased", f'=COUNTIF({LL("D")},"Purchased")'), ("Access: PS Plus (tile label)", f'=COUNTIF({LL("D")},"PS Plus")'),
        ("Access: Free / claimed (transaction at 0.00)", f'=COUNTIF({LL("D")},"Free / claimed")'), ("Access: Free (assumed, no record)", f'=COUNTIF({LL("D")},"Free (assumed)")'),
        ("", None), ("PROGRESS", None),
        ("Played (trophies, Playing, Finished or Dropped)", f'=COUNTIF({LL("G")},"Yes")'), ("Finished / 100% trophies", f'=COUNTIF({LL("H")},"Yes")'), ("Platinum", f'=COUNTIF({LL("J")},"Yes")'),
        ("Playing now", f'=COUNTIF({LL("E")},"Playing")'),
        ("Backlog (not played, not dropped / won't play)", f'=COUNTIFS({LL("G")},"",{LL("E")},"")+COUNTIFS({LL("G")},"",{LL("E")},"Backlog")'),
        ("High-priority backlog", f'=COUNTIFS({LL("G")},"",{LL("F")},"High")'),
        ("Dropped / won't play", f'=COUNTIF({LL("E")},"Dropped")+COUNTIF({LL("E")},"Won\'t play")'), ("Average rating", f'=IFERROR(AVERAGE({LL("K")}),"")')]
    r = 1
    for k, v in lines:
        sm.cell(r, 1, k)
        if v is not None: sm.cell(r, 2, v)
        if k.isupper() or r == 1: sm.cell(r, 1).font = Font(bold=True, size=14 if r == 1 else 11)
        r += 1
    r += 1; sm.cell(r, 1, f"SPEND IN {base} (excludes wallet top-ups)").font = Font(bold=True); r += 1
    sm.cell(r, 1, "Total spent, all accounts, all time"); sm.cell(r, 2, f'=SUMIFS({AM},{CAT},"<>Wallet top-up")').number_format = FMT; r += 1
    for reg, an in accounts:
        crit = f'{RG},"{reg}"' + (f',{AC},"{an}"' if an else "")
        sm.cell(r, 1, f"  {acct_label(reg, an)}"); sm.cell(r, 2, f'=SUMIFS({AM},{crit},{CAT},"<>Wallet top-up")').number_format = FMT; r += 1
    sm.cell(r, 1, "  of which PS Plus subscriptions"); sm.cell(r, 2, f'=SUMIFS({AM},{CAT},"Subscription")').number_format = FMT; r += 1
    sm.cell(r, 1, "  wallet top-ups (for reference)"); sm.cell(r, 2, f'=SUMIFS({AM},{CAT},"Wallet top-up")').number_format = FMT; r += 1
    sm.cell(r, 1, "  rows still missing an FX rate (see FX sheet)"); sm.cell(r, 2, f'=COUNTIF({AM},"FX?")'); r += 1
    years = sorted({int(t["Date"][:4]) for t in st.tx if t["Date"]}, reverse=True); regions = sorted({reg for reg, _ in accounts})
    if years:
        r += 1; sm.cell(r, 1, f"Spend per year ({base})").font = Font(bold=True); sm.cell(r, 2, "All").font = Font(bold=True)
        for j, reg in enumerate(regions): sm.cell(r, 3 + j, reg).font = Font(bold=True)
        first = r + 1; r += 1
        for y in years:
            sm.cell(r, 1, y); sm.cell(r, 2, f'=SUMIFS({AM},{YR},{y},{CAT},"<>Wallet top-up")').number_format = FMT
            for j, reg in enumerate(regions): sm.cell(r, 3 + j, f'=SUMIFS({AM},{YR},{y},{RG},"{reg}",{CAT},"<>Wallet top-up")').number_format = FMT
            r += 1
        sm.cell(r, 1, "Total").font = Font(bold=True)
        for j in range(2, 3 + len(regions)): sm.cell(r, j, f"=SUM({get_column_letter(j)}{first}:{get_column_letter(j)}{r-1})").number_format = FMT
        r += 1
    r += 1; sm.cell(r, 1, "ACCOUNTS CAPTURED").font = Font(bold=True); sm.cell(r, 2, "Library items").font = Font(bold=True); sm.cell(r, 3, "Transactions").font = Font(bold=True); r += 1
    for reg, an in accounts:
        n_ed = sum(1 for e in st.editions if e["Region"] == reg and e["Account"] == an)
        crit = f'{RG},"{reg}"' + (f',{AC},"{an}"' if an else "")
        sm.cell(r, 1, acct_label(reg, an)); sm.cell(r, 2, n_ed); sm.cell(r, 3, f"=COUNTIFS({crit})"); r += 1
    r += 1; sm.cell(r, 1, "LATEST SCANS").font = Font(bold=True); r += 1
    for j, h in enumerate(["Account", "Scan date", "Rows", "Reported", "Complete?"]): sm.cell(r, 1 + j, h).font = Font(bold=True)
    latest = {}
    for x in st.scans:
        k = (x["Region"], x["Account"])
        if k not in latest or x["Scan date"] >= latest[k]["Scan date"]: latest[k] = x
    for (reg, an), x in sorted(latest.items()):
        r += 1
        for j, v in enumerate([acct_label(reg, an), x["Scan date"], x["Rows captured"], x["Reported total"], x["Complete?"]]): sm.cell(r, 1 + j, v)
    sm.column_dimensions["A"].width = 56
    for col in "BCDEF": sm.column_dimensions[col].width = 14
    wb.save(out)

# ----------------------------------------------------------------------------- main
def classify_files(paths):
    scans, txs, tro, books, html = [], [], [], [], []
    skip_names = ("collections.csv", "type_overrides.csv", "aliases.csv", "trophy_aliases.csv")
    for p in paths:
        pl = os.path.basename(p).lower()
        if pl.endswith((".xlsx", ".xlsm")): books.append(p)
        elif pl.endswith((".html", ".htm")): html.append(p)
        elif pl.endswith(".csv") and pl not in skip_names: (txs if "transaction" in pl else tro if "troph" in pl else scans).append(p)
    return scans, txs, tro, books, html

def choose_base(arg, prev, txs):
    if arg: return arg.upper()
    if prev: return prev
    counts = {}
    for t in txs:
        if t["Currency"]: counts[t["Currency"]] = counts.get(t["Currency"], 0) + 1
    guess = max(counts, key=counts.get) if counts else "GBP"
    if sys.stdin.isatty():
        ans = input(f"Which currency should all money be shown in? [{guess}] (e.g. GBP, USD, EUR, TRY): ").strip().upper()
        return ans or guess
    return guess

def main():
    ap = argparse.ArgumentParser(description="Build the PlayStation library tracker workbook.")
    ap.add_argument("files", nargs="*")
    for opt in ("--export", "--transactions", "--trophies", "--psnp-html"): ap.add_argument(opt, action="append", default=[])
    ap.add_argument("--existing"); ap.add_argument("--out"); ap.add_argument("--base", help="base currency code for all totals, e.g. GBP, USD, EUR")
    ap.add_argument("--fresh", action="store_true"); ap.add_argument("--partial", action="store_true"); ap.add_argument("--pick", action="store_true"); ap.add_argument("--igdb", action="store_true", help="look games up on IGDB (needs TWITCH_CLIENT_ID / TWITCH_CLIENT_SECRET)")
    a = ap.parse_args()
    here = os.path.dirname(os.path.abspath(__file__)); files = list(a.files)
    if a.pick:
        import tkinter as tk; from tkinter import filedialog
        tk.Tk().withdraw(); files += list(filedialog.askopenfilenames(title="Select exports (CSV) and your previous workbook (xlsx)"))
    scans, txs, tro, books, html = classify_files(files)
    scans += a.export; txs += a.transactions; tro += a.trophies; html += a.psnp_html
    if a.existing: books.append(a.existing)
    if not scans and not txs:
        g = lambda pat: sorted(glob.glob(os.path.join(here, pat)), key=os.path.getmtime)
        scans, txs = g("psn-library_*.csv"), g("psn-transactions_*.csv"); tro = tro or g("psn-trophies*.csv")
        if not books: books = [b for b in g("PlayStation-Game-Collection*.xlsx") if not os.path.basename(b).startswith("~$")][-1:]
        print("Auto-detected:", [os.path.basename(x) for x in scans + txs + tro + books + html])
    if not scans and not txs and not books: sys.exit("No exports found. Run the userscript first (see README).")
    if a.fresh: books = []
    st = State(); load_extra_files(here, st)
    if books: print("Previous workbook:", os.path.basename(books[-1])); st.load_workbook_file(books[-1])
    def when(p):
        try: return s(pd.read_csv(p, nrows=1, dtype=str)["ScannedAt"].iloc[0])
        except Exception: return ""
    scan_frames, tx_frames = [], []
    for p in sorted(scans, key=lambda p: (when(p), p)):
        cols = pd.read_csv(p, nrows=0).columns
        scan_frames.append((os.path.basename(p), read_scan_csv(p) if ("ProductID" in cols or "TitleID" in cols) else read_rb4_csv(p)))
    for p in txs: tx_frames.append((os.path.basename(p), read_tx_csv(p)))
    frames = [f for _, f in scan_frames] + [f for _, f in tx_frames]
    apply_account_map(st, frames, plan_account_map(st, frames))          # old free-text labels -> PSN online ID
    for name, df in scan_frames:
        for (reg, acc, tab), part in df.groupby(["Region", "Account", "Tab"]): st.merge_scan(part.reset_index(drop=True), name, a.partial)
    for name, df in tx_frames: st.merge_tx(df, name)
    have = {(norm(t["Name"]), abs(t["Amount"] or 0)) for t in st.tx}
    st.tx += [t for t in st.legacy_tx if (norm(t["Name"]), abs(t["Amount"] or 0)) not in have]
    base = choose_base(a.base, st.base, st.tx); print(f"  base currency: {base}")
    trophies = []
    for p in tro + html: trophies += (read_trophy_csv(p) if p.lower().endswith(".csv") else read_psnp_html(p))
    lib, price_map, ambiguous, unmatched, idx = build_library(st, trophies)
    try:
        import igdb_enrich
        lib = igdb_enrich.enrich_library(lib, here, offline=not a.igdb)
    except ImportError:
        if a.igdb: print("  IGDB: igdb_enrich.py must sit next to this script.")
    if trophies:
        print(f"  trophy data: {len(trophies)-len(unmatched)}/{len(trophies)} lists attached to a library game")
        if unmatched:
            open(os.path.join(here, "unmatched_trophies.txt"), "w", encoding="utf-8").write("\n".join(f"{n}\t{why}" for n, why in unmatched))
            print(f"  {len(unmatched)} not attached - see unmatched_trophies.txt (disc copies, other accounts, trilogy children: add to trophy_aliases.csv)")
    review = review_rows(st, lib, ambiguous, idx)
    if find_collisions(st): print("  NOTE: possible same-name / different-game pairs found - see the Match review sheet (aliases.csv)")
    out = a.out or os.path.join(here, f"PlayStation-Game-Collection_{TODAY}.xlsx"); b0, i = out, 2
    while os.path.exists(out) and not a.out: out = b0.replace(".xlsx", f"-{i}.xlsx"); i += 1
    write_workbook(st, lib, price_map, review, base, out)
    for w in st.warnings: print("  NOTE:", w)
    print(f"Saved {out}: {len(lib)} games, {len(st.editions)} edition records, {len(st.tx)} transactions.")

if __name__ == "__main__":
    main()
