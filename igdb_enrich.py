#!/usr/bin/env python3
"""
igdb_enrich.py  -  OPTIONAL IGDB ratings / type / release-year enrichment for build_spreadsheet.py

Credentials are read ONLY from environment variables (never stored, never written to the workbook or the cache):
    TWITCH_CLIENT_ID      TWITCH_CLIENT_SECRET
Windows (PowerShell, this session only):   $env:TWITCH_CLIENT_ID="..."; $env:TWITCH_CLIENT_SECRET="..."
macOS / Linux:                             export TWITCH_CLIENT_ID=...; export TWITCH_CLIENT_SECRET=...

Every lookup is cached in igdb_cache.json (next to this script), so later builds are instant and work offline.
Optional igdb_overrides.csv   (Name;IgdbID  or  Region;Name;IgdbID)   forces a specific IGDB game, e.g.  TR;Resident Evil 4 (2005);20065
IGDB allows ~4 requests/second; the first run for ~650 games takes about 3-4 minutes.
"""
import csv, difflib, json, os, re, time, unicodedata, urllib.error, urllib.parse, urllib.request

CACHE_FILE, OVERRIDE_FILE, CACHE_DAYS = "igdb_cache.json", "igdb_overrides.csv", 45
PLATFORM_IDS = {"PS3": 9, "PS4": 48, "PS5": 167, "PS VITA": 46, "VITA": 46, "PSP": 38, "VR": 165, "PSVR": 165, "PSVR2": 390}
KEEP_TYPES = (0, 8, 9, 10, 11)                                   # Main Game, Remake, Remaster, Expanded Game, Port
TYPE_NAMES = {0: "Main Game", 1: "DLC", 2: "Expansion", 3: "Bundle", 4: "Standalone Expansion", 5: "Mod", 6: "Episode", 7: "Season",
              8: "Remake", 9: "Remaster", 10: "Expanded Game", 11: "Port", 12: "Fork", 13: "Pack / Addon", 14: "Update"}
FIELDS = ("name,slug,game_type,aggregated_rating,aggregated_rating_count,rating,rating_count,total_rating,total_rating_count,"
          "first_release_date,platforms")
COLUMNS = ["Critic (IGDB)", "Users (IGDB)", "IGDB type", "Year", "IGDB match", "IGDB"]
REVIEW = []                                                      # filled by enrich_library(), written as the 'IGDB review' sheet

_EDITION = re.compile(r"\b(standard|deluxe|digital deluxe|ultimate|gold|complete|definitive|game of the year|goty|premium|special|collector'?s|"
                      r"anniversary|launch|day one|enhanced|director'?s cut|legacy)\s+edition\b|\b(ps4|ps5|ps4 & ps5|ps4 and ps5|playstation ?[45])\b", re.I)
def clean(name):
    t = unicodedata.normalize("NFKD", str(name or "")).replace("\u00ae", "").replace("\u2122", "").replace("\u00a9", "")
    t = re.sub(r"\(\s*(19|20)\d\d\s*\)", " ", t)                 # "(2005)" display suffixes
    t = _EDITION.sub(" ", t)
    return re.sub(r"\s+", " ", t).strip(" -:\u2013")
def norm(name):
    t = unicodedata.normalize("NFKD", clean(name)).encode("ascii", "ignore").decode().lower().replace("&", " and ")
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9 ]+", " ", t)).strip()
def sim(a, b):
    a, b = norm(a), norm(b)
    if not a or not b: return 0.0
    return 1.0 if a == b else difflib.SequenceMatcher(None, a, b).ratio()
def plat_ids(platforms):
    return {PLATFORM_IDS[p.strip().upper()] for p in re.split(r"[+/,]", platforms or "") if p.strip().upper() in PLATFORM_IDS}

class Igdb:
    def __init__(self, folder, log=print):
        self.folder, self.log, self.token, self.exp, self.last = folder, log, None, 0, 0.0
        self.cid, self.secret = os.environ.get("TWITCH_CLIENT_ID", ""), os.environ.get("TWITCH_CLIENT_SECRET", "")
        self.path = os.path.join(folder, CACHE_FILE)
        try: self.cache = json.load(open(self.path, encoding="utf-8"))
        except Exception: self.cache = {}
        self.calls = 0
    @property
    def online(self): return bool(self.cid and self.secret)
    def save(self):
        json.dump(self.cache, open(self.path, "w", encoding="utf-8"), ensure_ascii=False)
    def _http(self, url, data=None, headers=None):
        req = urllib.request.Request(url, data=data, headers=headers or {}, method="POST")
        with urllib.request.urlopen(req, timeout=30) as r: return json.loads(r.read().decode("utf-8"))
    def _token(self):
        if self.token and time.time() < self.exp - 60: return self.token
        q = urllib.parse.urlencode({"client_id": self.cid, "client_secret": self.secret, "grant_type": "client_credentials"})
        j = self._http("https://id.twitch.tv/oauth2/token?" + q)
        self.token, self.exp = j["access_token"], time.time() + int(j.get("expires_in", 3600)); return self.token
    def _post(self, endpoint, body):
        for attempt in range(4):
            wait = 0.28 - (time.time() - self.last)
            if wait > 0: time.sleep(wait)
            self.last = time.time()
            try:
                return self._http("https://api.igdb.com/v4/" + endpoint, body.encode("utf-8"), {"Client-ID": self.cid, "Authorization": "Bearer " + self._token(), "Accept": "application/json"})
            except urllib.error.HTTPError as e:
                if e.code == 429 and attempt < 3: time.sleep(1.5 * (attempt + 1)); continue
                if e.code in (401, 403): self.token = None
                if attempt == 3 or e.code not in (401, 403, 500, 502, 503): raise
            except urllib.error.URLError:
                if attempt == 3: raise
                time.sleep(2)
    def _cached(self, key):
        c = self.cache.get(key)
        return c["r"] if c and time.time() - c["t"] < CACHE_DAYS * 86400 else None
    def search(self, term):
        key = "q|" + norm(term); hit = self._cached(key)
        if hit is not None or not self.online: return hit
        safe = clean(term).replace("\\", " ").replace('"', " ")
        r = self._post("games", f'search "{safe}"; fields {FIELDS}; where game_type = ({",".join(map(str, KEEP_TYPES))}); limit 15;')
        self.cache[key] = {"t": time.time(), "r": r}; self.calls += 1
        if self.calls % 25 == 0: self.save(); self.log(f"    IGDB: {self.calls} lookups done...")
        return r
    def by_id(self, gid):
        key = f"id|{gid}"; hit = self._cached(key)
        if hit is not None or not self.online: return (hit or [None])[0] if hit else None
        r = self._post("games", f"fields {FIELDS}; where id = {int(gid)};")
        self.cache[key] = {"t": time.time(), "r": r}; self.calls += 1
        return r[0] if r else None

def load_overrides(folder):
    out, p = {}, os.path.join(folder, OVERRIDE_FILE)
    if os.path.exists(p):
        for row in csv.reader(open(p, encoding="utf-8-sig"), delimiter=";"):
            row = [x.strip() for x in row]
            if len(row) == 2 and row[1].isdigit(): out[("*", norm(row[0]))] = int(row[1])
            elif len(row) >= 3 and row[2].isdigit(): out[(row[0].upper(), norm(row[1]))] = int(row[2])
    return out

def year_hint(name):
    m = re.search(r"\(\s*((?:19|20)\d\d)\s*\)", str(name or "")); return int(m.group(1)) if m else None

def choose(cands, name, platforms):
    """-> (best candidate or None, confidence label, similarity, reason, other candidates)"""
    want = plat_ids(platforms); scored = []; yh = year_hint(name)
    for c in cands or []:
        sm = sim(name, c.get("name", ""))
        sm_bonus = 0.05 * len(want & set(c["platforms"])) / len(want) if want and c.get("platforms") else 0
        if yh and c.get("first_release_date") and time.gmtime(c["first_release_date"]).tm_year == yh: sm_bonus += 0.05   # "(2005)" in the display name
        scored.append((sm + sm_bonus, sm, c))
    scored.sort(key=lambda x: -x[0])
    if not scored or scored[0][1] < 0.75: return None, "No match", scored[0][1] if scored else 0, "no candidate with a similar name", scored[:3]
    best = scored[0]; ties = [x for x in scored[1:] if x[1] >= 0.9 and x[2]["id"] != best[2]["id"] and best[0] - x[0] < 0.02]
    if best[1] < 0.9: return best[2], "Check", best[1], "name only loosely similar", scored[1:3]
    if ties:
        kinds = " / ".join(sorted({TYPE_NAMES.get(x[2].get("game_type"), "?") for x in [best] + ties}))
        return best[2], "Check", best[1], f"same name, several IGDB games ({kinds}) - add igdb_overrides.csv if wrong", ties[:3]
    return best[2], "OK", best[1], "", scored[1:3]

def enrich_library(lib, folder, offline=False, log=print):
    """Adds the IGDB columns to the library DataFrame. offline=True uses the cache only (no credentials needed)."""
    global REVIEW
    REVIEW = []; cli = Igdb(folder, log); ov = load_overrides(folder)
    if offline and not os.path.exists(cli.path): return lib
    if not offline and not cli.online:
        log("  IGDB: TWITCH_CLIENT_ID / TWITCH_CLIENT_SECRET are not set - using cached data only."); offline = True
    cols = {c: [] for c in COLUMNS}
    for _, r in lib.iterrows():
        name, plats = r["Name"], r.get("Platforms", ""); region = (str(r.get("Regions", "")).split(",")[0].strip() or "*").upper()
        g, label, sc, why, others = None, "", 0.0, "", []
        try:
            forced = ov.get((region, norm(name))) or ov.get(("*", norm(name)))
            if forced: g = cli.by_id(forced); label, sc, why = ("Override", 1.0, "igdb_overrides.csv") if g else ("No match", 0, f"override id {forced} not found")
            else:
                cands = cli.search(name)
                if cands is None: label, why = "Not looked up", "no cached result and no credentials"
                else: g, label, sc, why, others = choose(cands, name, plats)
        except Exception as e:
            label, why = "Error", f"{type(e).__name__}: {e}"[:120]
        if g:
            ts = g.get("first_release_date")
            cols["Critic (IGDB)"].append(round(g["aggregated_rating"], 1) if g.get("aggregated_rating") is not None else None)
            cols["Users (IGDB)"].append(round(g["rating"], 1) if g.get("rating") is not None else None)
            cols["IGDB type"].append(TYPE_NAMES.get(g.get("game_type"), ""))
            cols["Year"].append(time.gmtime(ts).tm_year if ts else None)
            cols["IGDB match"].append(label if label != "OK" else "OK")
            cols["IGDB"].append(f"https://www.igdb.com/games/{g['slug']}" if g.get("slug") else None)
        else:
            for c in COLUMNS: cols[c].append(None)
            cols["IGDB match"][-1] = label or None
        if label not in ("OK", "Override"):
            alt = "; ".join(f"{x[2].get('name')} [{TYPE_NAMES.get(x[2].get('game_type'), '?')}, id {x[2]['id']}]" for x in (others or [])[:3])
            REVIEW.append([name, g.get("name") if g else "", TYPE_NAMES.get(g.get("game_type"), "") if g else "", g["id"] if g else "", label, round(sc, 2), why, alt])
    cli.save(); out = lib.copy()
    for c in COLUMNS: out[c] = cols[c]
    n_ok = sum(1 for x in cols["IGDB match"] if x in ("OK", "Override"))
    log(f"  IGDB: {n_ok}/{len(lib)} matched confidently, {len(REVIEW)} to review (see the 'IGDB review' sheet); {cli.calls} new lookups.")
    return out
