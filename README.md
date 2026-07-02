# Work log

A small, self-hosted work-logging web app. Log daily work descriptions and time blocks,
switch between companies, get monthly overviews, and export to CSV / Markdown / XLSX.

- **Stack:** a single-file Python / Flask app with inline templates, SQLite storage,
  served by gunicorn. No external database or services required.
- **Auth:** email + password, invitation-only account creation (no LDAP, no third-party auth).
- **Multi-tenant:** work is organised into *companies*; users switch between the companies
  they belong to, and each company keeps its own records.
- **License:** GNU GPL v3.0.

---

## Table of contents

- [Quick start (Docker)](#quick-start-docker)
- [Manual installation](#manual-installation-without-docker)
- [Configuration](#configuration)
- [First run](#first-run)
- [Roles & usage](#roles--usage)
- [Running behind a reverse proxy](#running-behind-a-reverse-proxy)
- [Build the image yourself](#build-the-image-yourself)
- [Data & backup](#data--backup)
- [License](#license)

---

## Quick start (Docker)

Requires Docker + the Docker Compose plugin.

```bash
# 1. Get the compose file and env template
mkdir worklog && cd worklog
curl -O https://raw.githubusercontent.com/Hannibalus/worklog/main/docker-compose.yml
curl -o .env https://raw.githubusercontent.com/Hannibalus/worklog/main/.env.example

# 2. Edit .env - set a secret and the bootstrap admin
#    (generate a secret: python3 -c "import secrets;print(secrets.token_hex(32))")
vim .env

# 3. Start it
docker compose up -d
```

The app is now on **http://localhost:5000**. Sign in with the bootstrap admin email/password
you set in `.env`.

### Or a single `docker run`

```bash
docker run -d --name worklog -p 5000:5000 \
  -e WORKLOG_SECRET_KEY="$(python3 -c 'import secrets;print(secrets.token_hex(32))')" \
  -e WORKLOG_ADMIN_EMAIL="you@example.com" \
  -e WORKLOG_ADMIN_PASSWORD="change-me" \
  -v "$(pwd)/data:/data" \
  ghcr.io/hannibalus/worklog:latest
```

Images are published to **`ghcr.io/hannibalus/worklog`** (tags: `latest`, and versioned tags).

---

## Manual installation (without Docker)

Requires **Python 3.11+**.

```bash
# 1. Clone
git clone https://github.com/Hannibalus/worklog.git
cd worklog

# 2. Virtual environment + dependencies
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# 3. Configuration (environment variables)
export WORKLOG_SECRET_KEY="$(python3 -c 'import secrets;print(secrets.token_hex(32))')"
export WORKLOG_ADMIN_EMAIL="you@example.com"
export WORKLOG_ADMIN_PASSWORD="change-me-strong"

# 4a. Run for development (Flask dev server on :5000)
python3 worklog_web.py

# 4b. Or run for production with gunicorn
gunicorn -w 4 -b 0.0.0.0:5000 worklog_web:APP
```

The SQLite database is created at **`/data/worklog.sqlite3`**. Make sure that path is
writable, or change `DB_FILE` near the top of `worklog_web.py` to a path you control
(e.g. `./data/worklog.sqlite3`).

### Run as a systemd service (optional)

Create `/etc/systemd/system/worklog.service`:

```ini
[Unit]
Description=Work log
After=network.target

[Service]
WorkingDirectory=/opt/worklog
Environment=WORKLOG_SECRET_KEY=REPLACE_ME
Environment=WORKLOG_ADMIN_EMAIL=you@example.com
Environment=WORKLOG_ADMIN_PASSWORD=change-me
ExecStart=/opt/worklog/.venv/bin/gunicorn -w 4 -b 0.0.0.0:5000 worklog_web:APP
Restart=always

[Install]
WantedBy=multi-user.target
```

```bash
sudo systemctl daemon-reload
sudo systemctl enable --now worklog
```

---

## Configuration

All configuration is via environment variables.

| Variable | Default | Purpose |
|---|---|---|
| `WORKLOG_SECRET_KEY` | dev fallback | Flask session secret. **Set a random value in production.** |
| `WORKLOG_ADMIN_EMAIL` | - | Bootstrap superadmin email, created on first start if absent. |
| `WORKLOG_ADMIN_PASSWORD` | - | Bootstrap superadmin password. |
| `WORKLOG_ADMIN_NAME` | `Admin` | Bootstrap superadmin display name. |
| `WORKLOG_BRAND_NAME` | `Work log` | Name shown in the UI. |
| `WORKLOG_BRAND_MARK` | `W` | Single-letter logo mark. |
| `WORKLOG_INVITE_TTL_HOURS` | `24` | Invitation-link lifetime, in hours. |

The database lives at `/data/worklog.sqlite3` - mount/persist that directory.

---

## First run

1. Set `WORKLOG_SECRET_KEY`, `WORKLOG_ADMIN_EMAIL`, `WORKLOG_ADMIN_PASSWORD`.
2. Start the app - the first superadmin account is created automatically.
3. Sign in, open **Admin -> Companies**, and create a company.
4. Open the company's **Manage members**, invite users by email, and send them the
   registration links the app displays (the app does not send email itself).

---

## Roles & usage

- **Superadmin** - creates/deletes companies, invites users, assigns roles, sees everything
  (has the **Admin** button). The first superadmin comes from the bootstrap env vars.
- **Manager** (per-company role) - read-only overseer of one company: can view all members'
  records in that company via the member selector on the **Month** page. Cannot manage.
- **Member** - logs their own work and sees only their own records.

Anyone who logs work in a company automatically becomes a member of it, so their records
are visible to that company's managers.

**Invitations:** invitation links are single-use and expire after `WORKLOG_INVITE_TTL_HOURS`.
Every user can change their own password via the **Account** button.

---

## Running behind a reverse proxy

The app listens on port `5000` and speaks plain HTTP - terminate TLS at your proxy.

**Traefik** (labels; drop the `ports:` mapping and attach to your proxy network):

```yaml
    labels:
      - "traefik.enable=true"
      - "traefik.http.routers.worklog.rule=Host(`worklog.example.com`)"
      - "traefik.http.routers.worklog.entrypoints=https"
      - "traefik.http.routers.worklog.tls=true"
      - "traefik.http.services.worklog.loadbalancer.server.port=5000"
```

**nginx** (reverse proxy):

```nginx
location / {
    proxy_pass http://127.0.0.1:5000;
    proxy_set_header Host $host;
    proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
    proxy_set_header X-Forwarded-Proto $scheme;
}
```

---

## Build the image yourself

```bash
docker build -t worklog:local .
docker run -d -p 5000:5000 -v "$(pwd)/data:/data" \
  -e WORKLOG_ADMIN_EMAIL=you@example.com -e WORKLOG_ADMIN_PASSWORD=change-me worklog:local
```

---

## Data & backup

Everything lives in a single SQLite file at `/data/worklog.sqlite3` (persisted via the
`./data` volume). To back up, just copy that file:

```bash
cp data/worklog.sqlite3 data/worklog.sqlite3.bak.$(date +%F)
```

---

## License

This project is licensed under the **GNU General Public License v3.0** - see the
[LICENSE](LICENSE) file. You may use, study, share, and modify it; if you distribute a
modified version, you must also release your source under the GPL.