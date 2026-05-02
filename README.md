# Membership WebApp

Web app to verify memberships by scanning a barcode card and matching it against a SQLite master database.

## What it does

- Runs as a web app in Docker Compose behind Traefik.
- Uses SQLite as the master copy of all membership data.
- Imports Excel rows into SQLite and only adds or merges missing DB data.
- Verifies a scanned value against membership number, email, or member name.
- Edits member details in the browser.
- Updates membership usage in SQLite.
- Supports basic auth via Traefik middleware.
- Uses a simple single-container Python web server.

## Data integrity and safety

- SQLite is stored on a Docker volume.
- Excel imports never replace populated DB fields with blank workbook values.
- Basic input hardening is applied to block formula-injection-like values in editable text fields.
- A hidden internal audit trail (hash-chained) records app activity.

## Requirements

- Docker and Docker Compose
- Traefik network named `traefik`
- Dependencies from `requirements.txt` (`openpyxl`, `Pillow`)

## Run locally

```bash
docker compose up --build
```

The app listens on port `8000` inside the container and is intended to be reached through Traefik.

For Portainer, deploy the stack using a prebuilt image tag instead of `build:`.

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

## Docker Compose

`docker-compose.yml` includes the Traefik labels you requested:

```yaml
networks:
  - traefik

labels:
  - "traefik.enable=true"
  - "traefik.docker.network=traefik"
  - "traefik.http.routers.member-web.rule=Host(`member.cyberconnectit.com`)"
  - "traefik.http.routers.member-web.entrypoints=websecure"
  - "traefik.http.routers.member-web.tls.certresolver=cloudflare"
  - "traefik.http.routers.member-web.middlewares=member-auth"
  - "traefik.http.middlewares.member-auth.basicauth.users=..."
  - "traefik.http.services.member-web.loadbalancer.server.port=8000"
  - "traefik.http.services.member-web.loadbalancer.server.scheme=http"

networks:
  traefik:
    external: true
```

## Excel import behavior

- Uploading a workbook imports rows into SQLite.
- Existing DB rows are matched by sheet and row source key.
- Blank workbook fields do not overwrite populated DB fields.
- Usage changes happen in SQLite, not Excel.

## Setup

```bash
python -m venv .venv
```

### Local file upload

```bash
Open the app through your Traefik hostname and upload the workbook from the dashboard.
```

## Scanner setup

Most USB barcode scanners act like keyboards in the browser too:

1. Focus the search field.
2. Scan the card.
3. Search runs from the barcode value.

## Matching behavior

The app attempts lookup in this order:

1. Membership Number
2. Email
3. Name (`First Last` and `Last, First`)
4. Fallback partial-text match

## Scan history log

- Audit events are written to the app data directory.

## Usage count syncing

- Usage is updated in SQLite immediately.

## Internal audit trail

- In addition to the visible `scan_history.csv`, the app writes a hidden internal audit journal.
- It stores hash-chained records for app lifecycle events, scan/log operations, and critical state updates.
- The app verifies chain integrity on startup and flags integrity failures.

## Notes

- If you want Traefik-only auth, leave `MEMBER_BASIC_AUTH_USER` and `MEMBER_BASIC_AUTH_PASS` empty.
- If you want app-layer auth too, set both env vars in Portainer.
