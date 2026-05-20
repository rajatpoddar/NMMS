# NMMS Tracking Report

> Automated attendance tracking and reporting tool for MNREGA NMMS. Extracts muster roll data, generates professional Excel reports with embedded photos.

## Architecture

```
┌──────────────────────┐      HTTP/API       ┌──────────────────────┐
│   Desktop App (GUI)  │ ◄─────────────────► │   Server (Flask)     │
│   - customtkinter    │     Port 6667        │   - License Mgmt     │
│   - Selenium Scraper │                      │   - Admin Panel      │
│   - Excel Export     │                      │   - PostgreSQL DB    │
│   Runs on your PC    │                      │   Runs on your NAS   │
└──────────────────────┘                      └──────────────────────┘
```

## Quick Start

### 1. Deploy the Server on your NAS

```bash
# Clone the repo on your NAS
git clone https://github.com/rajatpoddar/NMMS.git nmms-tracker
cd nmms-tracker

# Copy and edit env config (optional - defaults work out of the box)
cp .env.example .env

# Start the server with PostgreSQL
docker compose up -d

# Check it's running
curl http://localhost:6667/health
```

**Expected response:**
```json
{"status": "healthy", "database": "postgresql", "port": 6667}
```

### 2. Configure the Desktop App

On the machine where you run the GUI app:

**Option A — Environment variable:**
```bash
# Windows (Command Prompt)
set NMMS_SERVER_URL=https://nmms.palojori.in
python app.py

# Mac / Linux
export NMMS_SERVER_URL=https://nmms.palojori.in
python app.py
```

**Option B — Config file:**
Create `~/.nmms_config.json` (or `config.json` next to `app.py`):
```json
{"server_url": "https://nmms.palojori.in"}
```

Then just run `python app.py`.

### 3. Access the Admin Panel

Open in your browser:

```
https://nmms.palojori.in/admin
```

From here you can:
- View all registered users with expiry dates
- Toggle user active/disabled status
- Extend licenses by 30 days

## Docker Deployment (Production)

### Prerequisites

- Docker & Docker Compose on your NAS
- Port 6667 available (change via `HOST_PORT` in `.env`)
- Git installed to clone and pull updates

### Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `HOST_PORT` | `6667` | External port for the server |
| `DB_PASSWORD` | `nmms_password` | PostgreSQL password |
| `GUNICORN_WORKERS` | `2` | Number of Gunicorn workers |
| `GUNICORN_TIMEOUT` | `120` | Worker timeout (seconds) |
| `LOG_LEVEL` | `info` | Log verbosity |
| `TRIAL_DAYS` | `30` | Trial license duration (days) |

### Deploy Updates with deploy.sh

To deploy updates, run the included deployment script on your NAS:

```bash
./deploy.sh
```

This script will:
1. Pull the latest code from GitHub
2. Automatically backup your PostgreSQL database to `./backups/`
3. Stop running containers
4. Rebuild and restart containers
5. Wait for the health check to confirm everything is working

### Manual Backup & Restore

```bash
# Backup the database
docker exec nmms-tracker-db pg_dump -U nmms nmms_tracker > backup_$(date +%Y%m%d).sql

# Restore
docker exec -i nmms-tracker-db psql -U nmms nmms_tracker < backup.sql
```

### Persistent Data

The PostgreSQL database is stored in a Docker volume called `nmms_postgres_data`. It persists across container restarts and updates.

## Desktop App Usage

The desktop app (`app.py`) is a GUI application that:

1. **Connects to your server** for license verification
2. **Scrapes NREGA NMMS website** to extract muster roll data
3. **Generates professional Excel reports** with worker details, photos, and geo-coordinates

### Requirements

- Python 3.8+
- Google Chrome installed
- Dependencies: `pip install -r requirements.txt`

### Running

```bash
# Set your server URL first, then:
python app.py
```

### First-time Setup

1. The app will ask you to register via a web portal
2. Register with your details and MAC address
3. Refresh the app — you'll get a 30-day trial
4. Admin can extend via the admin panel

## Building a Windows Installer

A GitHub Actions workflow is included. To build a standalone `.exe`:

1. Push to `main` or `master` branch — the workflow runs automatically
2. Or trigger manually from GitHub Actions UI:
   - Go to **Actions** → **Build Windows Installer** → **Run workflow**
3. Download the `.exe` from the workflow artifacts

The installer includes Python, all dependencies, and the app in a single executable.

## Admin Dashboard

Accessible at `http://<your-nas-ip>:6667/admin`:

![Admin Dashboard](https://via.placeholder.com/800x400?text=NMMS+Admin+Dashboard)

Features:
- **Stats cards**: Total/Active/Expired/Disabled users
- **User table**: See name, location, registration & expiry dates
- **Days left**: Visual progress bar for remaining trial days
- **Actions**: Toggle status or extend license with one click

## API Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/register?mac=xxx` | GET | Registration page |
| `/register_user` | POST | Register a device |
| `/api/check_status?mac=xxx` | GET | Check license status |
| `/admin` | GET | Admin dashboard |
| `/admin/toggle/<mac>` | GET | Toggle user active status |
| `/admin/extend/<mac>` | GET | Extend license by 30 days |
| `/health` | GET | Health check endpoint |

## Cloudflare Deployment

This application is designed to run behind [Cloudflare](https://www.cloudflare.com/) for SSL/TLS and domain management.

### Setup Steps

1. **Point your domain** to your NAS IP in Cloudflare DNS (A record):
   ```
   Type:  A
   Name:  nmms
   IPv4:  <YOUR-NAS-PUBLIC-IP>
   Proxy:  Proxied (orange cloud) — enables SSL, DDoS protection
   ```

2. **Configure SSL/TLS** in Cloudflare Dashboard → SSL/TLS:
   - Set to **Full (strict)** for maximum security
   - Cloudflare handles HTTPS termination — your server stays on HTTP behind it

3. **Create an Origin Certificate** (Cloudflare → SSL/TLS → Origin Server):
   - Generate a certificate for `nmms.palojori.in`
   - This is optional but recommended for encryption between Cloudflare and your NAS

4. **Open port 6667** on your NAS firewall/router and forward to the Docker host

5. **Desktop clients** connect using:
   ```bash
   export NMMS_SERVER_URL=https://nmms.palojori.in
   ```

### How it Works

```
┌─────────────┐     HTTPS     ┌────────────┐     HTTP      ┌─────────────┐
│ Desktop App │ ──────────►  │  Cloudflare │ ──────────►  │ Flask Server │
│ (Client)    │               │  (Proxy)    │  Port 6667   │ (NAS Docker) │
└─────────────┘               └────────────┘               └─────────────┘
```

- Cloudflare terminates HTTPS at the edge
- Forwards traffic as HTTP to your NAS on port 6667
- The server uses `ProxyFix` middleware to correctly detect HTTPS for redirects
- All admin panel links and API responses work seamlessly behind Cloudflare

### Security

- Cloudflare DDoS protection shields your NAS
- SSL/TLS encryption between clients and Cloudflare
- Optionally encrypt traffic between Cloudflare and your server with Origin Certificates
- The server runs as a non-root user inside Docker
- PostgreSQL requires authentication and is not exposed externally

## Development

### Project Structure

```
├── app.py                    # Desktop GUI client
├── server.py                 # Flask server (license management)
├── requirements.txt           # All dependencies
├── requirements-server.txt   # Server-only dependencies
├── Dockerfile                # Server Docker image
├── docker-compose.yml        # Production deployment
├── entrypoint.sh             # Docker entrypoint
├── .env.example              # Environment config template
├── .gitignore                # Git ignore rules
├── .github/
│   └── workflows/
│       └── build-windows.yml # Windows installer builder
└── README.md                 # This file
```

## License Management

The server manages device licenses with:
- **30-day trials** (configurable via `TRIAL_DAYS`)
- **MAC-based device registration**
- **Admin-controlled extensions** (clicking "+30 Days" adds 30 days to current expiry)
- **Toggle activation** to enable/disable devices
