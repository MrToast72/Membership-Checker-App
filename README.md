# Membership Card Verifier

Desktop app to verify memberships by scanning a barcode card and matching it against an Excel workbook.

## What it does

- Uses a USB barcode scanner as keyboard input (standard HID scanner behavior).
- Loads your Excel workbook (`.xlsx` / `.xlsm`) as the membership database.
- Verifies a scanned value against membership number, email, or member name.
- Shows match details in modern result cards.
- Writes every scan result to a CSV history log.
- Automatically clears previous scan context when the next card starts scanning.
- Tracks and updates `Membership Amount Used` on each verified scan.
- Tracks scans in a local pending-usage cache and periodically syncs counts to Excel.
- Supports undo of the last scan to correct accidental scans.
- Includes in-app editing for member details, including `Includes Cart` and `Includes Range`.
- Loads all membership sheets in the workbook (every non-total tab), not just a single tab.
- Uses a modern rounded UI powered by CustomTkinter.
- Responsive layout down to `445px` app width (compact/stacked mode).
- Full window viewport is vertically scrollable, so match results and editor controls remain reachable on short app heights.
- Supports app icon configuration and packaged default icon assets.
- Uses code-based icon config from `Icon.png` beside `app.py` (no in-app icon picker).

## Data integrity and safety

- Every workbook write uses an atomic save (`.tmp` + replace) to reduce corruption risk.
- A timestamped backup copy is created before each save in a `backups/` folder beside the workbook.
- Save operations fail fast if the workbook changed on disk after loading, preventing silent overwrite.
- Basic input hardening is applied to block formula-injection-like values in editable text fields.
- Public scan logs are hash-tracked and verified on startup to detect tampering.
- A hidden internal audit trail (hash-chained) records app activity and log operations.

## Requirements

- Python 3.10+
- Windows, macOS, or Linux
- Dependencies from `requirements.txt` (`openpyxl`, `customtkinter`)
- Dev dependencies for tests: `requirements-dev.txt` (`pytest`)

## Tests

Run tests locally:

```bash
python -m pip install -r requirements-dev.txt
python -m pytest -q
```

Current tests cover:

- Responsive layout mode thresholds.
- Platform app-data path resolution.
- Formula-injection protections for edited values/log CSV.
- Hidden audit trail hash-chain integrity and tamper detection.
- Icon asset preprocessing behavior (format handling + background inference).

## App icon setup (code-based)

- Place your source icon file beside `app.py` as `Icon.png` (or update `APP_ICON_SOURCE` in `app.py`).
- Supported source formats include `.png`, `.jpg/.jpeg`, and `.webp`.
- During startup/build preparation, the app generates:
  - `assets/app_icon.png`
  - `assets/app_icon.ico`
- Processing rule:
  - The image height is preserved and scaled to icon height.
  - Width is centered with letterbox fill.
  - Fill color is inferred as white or black based on the source background.

## Setup

```bash
python -m venv .venv
```

## Web app (Docker)

Run the membership checker as a web app with Docker Compose:

```bash
docker compose up --build
```

Then open `http://localhost:8000` (or route the container through your reverse proxy so it is served via
`https://member.cyberconnectit.com`).

Uploads sync the workbook into SQLite at `./data/membership.sqlite3`. You can override the database path by setting
`MEMBERSHIP_DB_PATH`.

### Windows

```bash
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

### macOS / Linux

```bash
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

## Scanner setup

Most USB barcode scanners act like keyboards:

1. Click into the scanner input field.
2. Scan the card.
3. If your scanner sends Enter after scan, verification runs automatically.

If your scanner does not send Enter, press `Verify Membership`.

## Matching behavior

The app attempts lookup in this order:

1. Membership Number
2. Email
3. Name (`First Last` and `Last, First`)
4. Fallback partial-text match

## Scan history log

- Log file path:
  - Windows: `%LOCALAPPDATA%\MembershipVerifier\scan_history.csv`
  - macOS: `~/Library/Application Support/MembershipVerifier/scan_history.csv`
  - Linux: `~/.local/share/MembershipVerifier/scan_history.csv`
- A row is added for every scan with timestamp, scan value, result (`verified`, `multiple_matches`, `not_found`), and matched member details.
- Additional audit entries are logged for `undo` and `edit` actions.

## Usage count syncing

- Scans are first recorded in a local temp file: `pending_usage.json` in the app data folder.
- The app flushes pending usage deltas to Excel periodically, at threshold, and on clean app close.
- This reduces frequent workbook writes and lowers file-lock contention on Windows.
- If the app cannot flush on close, it keeps the window open and shows an error so no usage data is lost.

## Internal audit trail

- In addition to the visible `scan_history.csv`, the app writes a hidden internal audit journal.
- It stores hash-chained records for app lifecycle events, scan/log operations, and critical state updates.
- The app verifies chain integrity on startup and flags integrity failures.

## Notes about your current workbook

- Your workbook has a `Membership Number` column, but many rows are currently blank.
- Best reliability comes from barcode values matching `Membership Number` in Excel.
- If barcodes contain encoded names/emails instead, the app can still match those.

## Build a Windows .exe (embedded Python/runtime)

```bash
pip install pyinstaller
pyinstaller --noconsole --windowed --onefile --name MembershipVerifier app.py
```

Output executable will be in `dist/MembershipVerifier.exe`.

Important: a Windows `.exe` must be built on Windows (PyInstaller is platform-specific).

You can also run:

```bat
build_windows.bat
```

## Build Windows app in GitHub Actions

This repository includes a workflow at `.github/workflows/build-windows.yml`.

How to use it:

1. Push your latest code to GitHub.
2. Open GitHub -> `Actions` -> `Build Windows App`.
3. Click `Run workflow`.
4. Download `MembershipVerifier-windows` artifact (contains `MembershipVerifier.exe`).

## Build a macOS app and DMG (optional)

```bash
pip install pyinstaller
pyinstaller --noconsole --windowed --name MembershipVerifier app.py
hdiutil create -volname "MembershipVerifier" -srcfolder "dist/MembershipVerifier.app" -ov -format UDZO "dist/MembershipVerifier.dmg"
```

Output app bundle will be `dist/MembershipVerifier.app` and DMG will be `dist/MembershipVerifier.dmg`.

You can also run:

```bash
chmod +x build_mac.sh
./build_mac.sh
```

## Embedded runtime note

- PyInstaller bundles Python + dependencies into the app/exe, so target machines do not need Python or pip installed.
- You still need to distribute your Excel file with the app, or place it in the same folder as the executable/app.
