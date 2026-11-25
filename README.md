# PSN Library Exporter

A local-first, two-stage ETL pipeline that exports your PlayStation Store library and transaction history straight from the browser—without credentials, NPSSO tokens, or unofficial Sony API calls—and reconciles everything across multiple PSN accounts and regions into a formatted Excel workbook.

> **Status:** Active work in progress (Userscript `v8.6.0`, Builder `v8`). Usable end-to-end today.

---

## Architecture

```text
+--------------------------------------------+    +----------------------------------------------+
|  Browser (Tampermonkey: PSNLX-8.6.0)       |    |  Local Python ETL Pipeline                   |
|                                            |    |                                              |
|  library.playstation.com                   |    |  build_spreadsheet.py                        |
|   └─ Checkpoint-resumable page scraper     |    |   ├─ Schema normalisation & PII redaction    |
|                                            |    |   ├─ Ambiguity-guarded entity resolution     |
|  checkout.playstation.com (Order History)  |    |   │   (exact -> loose -> weak + collisions)  |
|   └─ Infinite-scroll + paid-order expander |    |   ├─ Basket/collection splitting & wallet    |
|                                            |    |   │   ledger reconciliation                  |
|  Cross-origin sync (GM_* / postMessage)    |    |   ├─ Optional PSNProfiles HTML trophy parser |
|  Regex PII redaction (email + card digits) |    |   └─ Optional IGDB enrichment (igdb_enrich)  |
+----------------------+---------------------+    +----------------------+-----------------------+
                       |                                                 |
                       v                                                 v
         psn-library_*.csv / psn-transactions_*.csv      PlayStation-Game-Collection_YYYY-MM-DD.xlsx
                                                         (Summary, Library, Transactions, FX,
                                                          Wallet, Editions, Match/IGDB Review, Scans)
```

---

## Privacy by Design

- **No credentials or tokens:** Never reads cookies, passwords, or NPSSO tokens.
- **Minimal identity capture:** Reads only your public PSN `onlineId` and `country` from the page's own session object; emails, real names, and internal account IDs are ignored.
- **Double-pass PII redaction:** Email addresses and masked card digits are stripped from expanded order text in the browser (`PSNLX-8.6.0.user.js`) and stripped again by `build_spreadsheet.py` before writing to Excel.
- **Keep exports private:** Generated `.csv`, `.xlsx`, `.html`, and `igdb_cache.json` files are ignored via `.gitignore`.

---

## Usage

### 1. Export Data From PlayStation (Per Account)
Install `PSNLX-8.6.0.user.js` in Tampermonkey, sign in to PlayStation in your browser, and:
1. Open **PlayStation Library → Purchased** (`library.playstation.com/recently-purchased`), press **▶ Start**, then click **⬇ Games**. Stopped scans resume from their last checkpoint automatically.
2. Open **Profile Menu → Order History**, press **▶ Start** in the Transactions Exporter panel (loads full history and expands paid orders for itemised prices), then click **⬇ Transactions** (or download from the Library panel).
3. Repeat for any additional regional PSN accounts.

### 2. Build the Workbook
```bash
pip install pandas openpyxl
# Optional (only if parsing saved PSNProfiles HTML pages):
pip install beautifulsoup4
```

Place your exported CSVs (and optional previous `.xlsx` workbook to preserve personal status/ratings/notes) in the project folder and run:

```bash
python build_spreadsheet.py
```

#### Useful Flags
| Flag | Purpose |
| :--- | :--- |
| `--base GBP` | Set base currency for totals (`GBP`, `USD`, `EUR`, `TRY`, etc.; remembered in the workbook). |
| `--psnp-html file.html` | Attach trophy progress and platinum status from a saved PSNProfiles HTML page (repeatable). |
| `--igdb` | Enrich games with critic/user ratings, game type, and release year via IGDB API (cached 45 days in `igdb_cache.json`). |
| `--fresh` | Ignore any previous workbook and build from raw CSVs only. |
| `--partial` | Merge a partial library scan without marking unseen titles as `Missing`. |
| `--pick` | Select input CSVs / previous `.xlsx` via a file picker dialog. |

#### Optional IGDB Enrichment
Requires free Twitch developer credentials supplied strictly via environment variables:
```powershell
$env:TWITCH_CLIENT_ID="your-client-id"
$env:TWITCH_CLIENT_SECRET="your-client-secret"
python build_spreadsheet.py --igdb
```
Subsequent builds reuse `igdb_cache.json` automatically offline without needing `--igdb` or credentials.

#### Optional Override Files (Semicolon/Tab/Comma-Delimited)
- `aliases.csv` (`Region;Name;DisplayName`) — Split or rename titles sharing a name (e.g. original vs. remake).
- `collections.csv` (`Collection;Component 1;Component 2;...`) — Map multi-game bundles to individual library entries.
- `trophy_aliases.csv` (`TrophyName;LibraryGame`) — Map standalone trilogy trophy lists to a single library title.
- `type_overrides.csv` (`Name;Type`) — Override auto-detected item types (`Game`, `Demo`, `App`, `Add-on`, etc.).
- `igdb_overrides.csv` (`Name;IgdbID` or `Region;Name;IgdbID`) — Pin a specific IGDB ID for ambiguous titles.

---

## Trade-Offs & Design Decisions

- **DOM scraping over Sony's internal GraphQL/NPSSO APIs:** Using a browser userscript against pages the user already opened avoids handling authentication tokens, reduces account-ban risk, and stays resilient to auth-flow changes—at the cost of slower paginated capture and reliance on front-end DOM selectors.
- **Ambiguity-guarded matching over greedy fuzzy joins:** Transaction lines and trophy lists only attach to a library title when a key (`exact` → `loose` → `weak`) resolves to **one** candidate (after regional filtering). Collisions (e.g. same-name remakes with widely separated `CUSA`/`PPSA` title numbers) and uncertain matches are routed to a dedicated **`Match review`** sheet rather than silently corrupting spend or completion data.
- **Live Excel/Sheets formulas (`GOOGLEFINANCE`, `SUMIFS`, `MINIFS`) over static values:** Writing dynamic formulas keeps historical multi-currency conversion, spend breakdowns, and wallet running balances transparent and auditable inside the workbook, trading offline instant FX values for zero third-party FX API dependencies.
- **Stateful incremental rebuilds:** Passing the prior workbook preserves manual columns (`Status`, `Priority`, `Rating`, `Notes`) and scan/change history across re-exports without requiring a separate database.

