import os
from flask import Flask, request, jsonify, render_template_string, redirect, url_for
from flask_cors import CORS
from werkzeug.middleware.proxy_fix import ProxyFix
import psycopg2
import psycopg2.extras
import psycopg2.errors
from datetime import datetime, timedelta
from contextlib import contextmanager

app = Flask(__name__)

# Trust Cloudflare proxy headers (X-Forwarded-Proto, X-Forwarded-For)
# This ensures redirects and URL generation use HTTPS when behind Cloudflare
app.wsgi_app = ProxyFix(
    app.wsgi_app,
    x_for=1,      # Trust one X-Forwarded-For hop
    x_proto=1,    # Trust one X-Forwarded-Proto hop
    x_host=1,     # Trust one X-Forwarded-Host hop
    x_prefix=1    # Trust one X-Forwarded-Prefix hop
)

CORS(app)

# PostgreSQL Configuration from environment variables
SERVER_HOST = os.environ.get('SERVER_HOST', '0.0.0.0')
SERVER_PORT = int(os.environ.get('SERVER_PORT', '6667'))
DB_HOST = os.environ.get('DB_HOST', 'localhost')
DB_PORT = int(os.environ.get('DB_PORT', '5432'))
DB_NAME = os.environ.get('DB_NAME', 'nmms_tracker')
DB_USER = os.environ.get('DB_USER', 'nmms')
DB_PASSWORD = os.environ.get('DB_PASSWORD', 'nmms_password')
DEBUG_MODE = os.environ.get('DEBUG', 'false').lower() == 'true'

TRIAL_DAYS = int(os.environ.get('TRIAL_DAYS', '30'))


@contextmanager
def get_db():
    """Context manager for PostgreSQL database connections."""
    conn = psycopg2.connect(
        host=DB_HOST,
        port=DB_PORT,
        dbname=DB_NAME,
        user=DB_USER,
        password=DB_PASSWORD
    )
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db():
    """Initialize the PostgreSQL database schema."""
    with get_db() as conn:
        c = conn.cursor()
        c.execute('''
            CREATE TABLE IF NOT EXISTS users (
                mac_id VARCHAR(255) PRIMARY KEY,
                name VARCHAR(255) NOT NULL,
                phone VARCHAR(50),
                state VARCHAR(100),
                district VARCHAR(100),
                block VARCHAR(100),
                registration_date DATE NOT NULL,
                is_active BOOLEAN DEFAULT TRUE
            )
        ''')


# ==========================================
# Helper: Calculate days left
# ==========================================
def get_days_left(reg_date):
    """Calculate remaining trial days from registration date."""
    if reg_date is None:
        return 0
    expiry = reg_date + timedelta(days=TRIAL_DAYS)
    remaining = (expiry - datetime.now().date()).days
    return max(0, remaining)


def get_expiry_date(reg_date):
    """Get the expiry date from registration date."""
    if reg_date is None:
        return None
    return reg_date + timedelta(days=TRIAL_DAYS)


# ==========================================
# LANDING & REGISTRATION PAGE
# ==========================================
LANDING_PAGE_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>NMMS Tracking Report - Pro Automation</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f7f6; }
        .hero { background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%); color: white; padding: 60px 0; border-radius: 0 0 30px 30px; margin-bottom: 40px; box-shadow: 0 10px 20px rgba(0,0,0,0.1); }
        .card { border: none; border-radius: 15px; box-shadow: 0 8px 30px rgba(0,0,0,0.08); }
        .btn-custom { background-color: #ff9800; color: white; font-weight: bold; border-radius: 8px; padding: 12px; transition: 0.3s; }
        .btn-custom:hover { background-color: #e68a00; color: white; transform: translateY(-2px); }
        input, select { text-transform: uppercase; }
    </style>
</head>
<body>
    <div class="hero text-center">
        <div class="container">
            <h1 class="display-4 fw-bold">NMMS Tracking Report</h1>
            <p class="lead mt-3">The ultimate professional tool for MNREGA reporting. Automate data extraction, download high-quality Excel sheets with embedded photos, and save hours of manual work.</p>
            <a href="#" class="btn btn-light btn-lg mt-3 fw-bold text-primary px-5 rounded-pill">Download Application</a>
        </div>
    </div>

    <div class="container mb-5">
        <div class="row justify-content-center">
            <div class="col-md-6">
                <div class="card p-4">
                    <h3 class="text-center mb-1 fw-bold text-dark">Start 30-Day Free Trial</h3>
                    <p class="text-center text-muted mb-4">Register your device to activate the software.</p>
                    <form action="/register_user" method="POST">
                        <input type="hidden" name="mac_id" value="{{ mac_id }}">
                        <div class="mb-3">
                            <label class="form-label fw-bold">Full Name</label>
                            <input type="text" name="name" class="form-control form-control-lg" oninput="this.value = this.value.toUpperCase()" required placeholder="ENTER YOUR NAME">
                        </div>
                        <div class="mb-3">
                            <label class="form-label fw-bold">Phone Number</label>
                            <input type="number" name="phone" class="form-control form-control-lg" required placeholder="10-DIGIT MOBILE NO">
                        </div>
                        <div class="mb-3">
                            <label class="form-label fw-bold">State</label>
                            <select name="state" class="form-select form-control-lg" required>
                                <option value="">-- SELECT STATE --</option>
                                <option value="JHARKHAND">JHARKHAND</option>
                                <option value="BIHAR">BIHAR</option>
                                <option value="WEST BENGAL">WEST BENGAL</option>
                                <option value="UTTAR PRADESH">UTTAR PRADESH</option>
                            </select>
                        </div>
                        <div class="row">
                            <div class="col-md-6 mb-3">
                                <label class="form-label fw-bold">District</label>
                                <input type="text" name="district" class="form-control form-control-lg" oninput="this.value = this.value.toUpperCase()" required placeholder="E.G. DEOGHAR">
                            </div>
                            <div class="col-md-6 mb-3">
                                <label class="form-label fw-bold">Block</label>
                                <input type="text" name="block" class="form-control form-control-lg" oninput="this.value = this.value.toUpperCase()" required placeholder="E.G. PALOJORI">
                            </div>
                        </div>
                        <button type="submit" class="btn btn-custom w-100 mt-2 fs-5">Register & Activate Software</button>
                    </form>
                </div>
            </div>
            <div class="col-md-5 mt-4 mt-md-0 d-flex flex-column justify-content-center px-4">
                <h4 class="fw-bold mb-3">Why choose our tool?</h4>
                <ul class="list-group list-group-flush fs-5">
                    <li class="list-group-item bg-transparent border-0">1-Click Automation: Scrape 50+ Muster rolls instantly.</li>
                    <li class="list-group-item bg-transparent border-0">Auto Photo Embed: Group Photos download & fit directly into Excel cells.</li>
                    <li class="list-group-item bg-transparent border-0">BDO/DM Ready Reports: Professional formatting applied automatically.</li>
                    <li class="list-group-item bg-transparent border-0">Secure & Fast: Works quietly in the background.</li>
                </ul>
            </div>
        </div>
    </div>
</body>
</html>
"""

# ==========================================
# PROFESSIONAL ADMIN DASHBOARD
# ==========================================
ADMIN_PAGE_HTML = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Admin Panel - NMMS Tracker</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.0/font/bootstrap-icons.css">
    <style>
        :root {
            --sidebar-width: 250px;
            --primary-dark: #1a1a2e;
            --secondary-dark: #16213e;
            --accent: #0f3460;
            --gold: #e94560;
        }
        * { box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
            background: #f0f2f5;
            margin: 0;
            min-height: 100vh;
            display: flex;
        }
        /* Sidebar */
        .sidebar {
            width: var(--sidebar-width);
            background: linear-gradient(180deg, var(--primary-dark) 0%, var(--secondary-dark) 100%);
            color: white;
            padding: 0;
            position: fixed;
            top: 0;
            left: 0;
            bottom: 0;
            z-index: 100;
            box-shadow: 2px 0 15px rgba(0,0,0,0.1);
        }
        .sidebar-header {
            padding: 24px 20px;
            border-bottom: 1px solid rgba(255,255,255,0.1);
            text-align: center;
        }
        .sidebar-header h4 {
            margin: 0;
            font-weight: 700;
            font-size: 1.1rem;
            letter-spacing: 0.5px;
        }
        .sidebar-header small {
            opacity: 0.7;
            font-size: 0.75rem;
        }
        .sidebar .nav-item {
            padding: 14px 24px;
            display: flex;
            align-items: center;
            gap: 12px;
            color: rgba(255,255,255,0.7);
            transition: all 0.2s;
            cursor: default;
            border-left: 3px solid transparent;
        }
        .sidebar .nav-item.active {
            background: rgba(255,255,255,0.08);
            color: white;
            border-left-color: var(--gold);
        }
        .sidebar .nav-item i { font-size: 1.2rem; }
        /* Main content */
        .main-content {
            margin-left: var(--sidebar-width);
            flex: 1;
            padding: 30px;
            min-height: 100vh;
        }
        .page-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 28px;
            flex-wrap: wrap;
            gap: 12px;
        }
        .page-header h2 {
            margin: 0;
            font-weight: 700;
            color: var(--primary-dark);
            font-size: 1.6rem;
        }
        .page-header .subtitle {
            color: #666;
            font-size: 0.9rem;
        }
        /* Stats cards */
        .stats-row {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 18px;
            margin-bottom: 28px;
        }
        .stat-card {
            background: white;
            border-radius: 12px;
            padding: 20px 24px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.05);
            transition: transform 0.2s;
        }
        .stat-card:hover { transform: translateY(-2px); }
        .stat-card .stat-label {
            font-size: 0.8rem;
            text-transform: uppercase;
            letter-spacing: 1px;
            color: #888;
            font-weight: 600;
        }
        .stat-card .stat-value {
            font-size: 2rem;
            font-weight: 700;
            margin: 4px 0 0;
        }
        .stat-card .stat-value.green { color: #00b894; }
        .stat-card .stat-value.red { color: var(--gold); }
        .stat-card .stat-value.blue { color: #0984e3; }
        .stat-card .stat-value.orange { color: #fdcb6e; }
        /* Table */
        .table-container {
            background: white;
            border-radius: 12px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.05);
            overflow: hidden;
        }
        .table-container .table-header {
            padding: 18px 24px;
            border-bottom: 1px solid #eee;
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 12px;
        }
        .table-container .table-header h5 {
            margin: 0;
            font-weight: 700;
        }
        .table-responsive { margin: 0; }
        .table {
            margin: 0;
            font-size: 0.9rem;
        }
        .table thead th {
            background: #f8f9fa;
            border-bottom: 2px solid #dee2e6;
            color: #495057;
            font-weight: 600;
            text-transform: uppercase;
            font-size: 0.75rem;
            letter-spacing: 0.5px;
            padding: 14px 16px;
            white-space: nowrap;
        }
        .table tbody td {
            padding: 14px 16px;
            vertical-align: middle;
        }
        .table tbody tr:hover { background: #f8f9ff; }
        .badge-status {
            padding: 5px 14px;
            border-radius: 50px;
            font-weight: 600;
            font-size: 0.75rem;
        }
        .badge-active {
            background: #d4edda;
            color: #155724;
        }
        .badge-expired {
            background: #f8d7da;
            color: #721c24;
        }
        .badge-disabled {
            background: #e2e3e5;
            color: #383d41;
        }
        .days-bar {
            display: inline-block;
            width: 100px;
            height: 6px;
            border-radius: 3px;
            background: #eee;
            margin-right: 10px;
            vertical-align: middle;
        }
        .days-bar-fill {
            height: 100%;
            border-radius: 3px;
            transition: width 0.3s;
        }
        .btn-action {
            padding: 5px 14px;
            font-size: 0.8rem;
            border-radius: 6px;
            font-weight: 600;
            border: none;
            transition: all 0.2s;
            text-decoration: none;
            display: inline-block;
        }
        .btn-action:hover { transform: translateY(-1px); }
        .btn-extend {
            background: #007bff;
            color: white;
        }
        .btn-extend:hover { background: #0056b3; color: white; }
        .btn-toggle {
            background: #6c757d;
            color: white;
        }
        .btn-toggle:hover { background: #5a6268; color: white; }
        .empty-state {
            text-align: center;
            padding: 60px 20px;
            color: #aaa;
        }
        .empty-state i { font-size: 3rem; }
        @media (max-width: 768px) {
            .sidebar { width: 60px; }
            .sidebar .nav-item span, .sidebar-header small, .sidebar-header h4 { display: none; }
            .main-content { margin-left: 60px; padding: 16px; }
        }
    </style>
</head>
<body>
    <!-- Sidebar -->
    <div class="sidebar">
        <div class="sidebar-header">
            <h4>NMMS Tracker</h4>
            <small>Admin Panel v2.0</small>
        </div>
        <div class="nav-item active">
            <i class="bi bi-people"></i>
            <span>Users</span>
        </div>
        <div class="nav-item">
            <i class="bi bi-gear"></i>
            <span>Settings</span>
        </div>
    </div>

    <!-- Main Content -->
    <div class="main-content">
        <div class="page-header">
            <div>
                <h2>User Management</h2>
                <span class="subtitle">Manage registered devices and license subscriptions</span>
            </div>
            <div>
                <span class="badge bg-light text-dark px-3 py-2">
                    <i class="bi bi-clock"></i> {{ now }}
                </span>
            </div>
        </div>

        <!-- Stats Cards -->
        <div class="stats-row">
            <div class="stat-card">
                <div class="stat-label">Total Users</div>
                <div class="stat-value blue">{{ stats.total }}</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">Active</div>
                <div class="stat-value green">{{ stats.active }}</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">Expired</div>
                <div class="stat-value red">{{ stats.expired }}</div>
            </div>
            <div class="stat-card">
                <div class="stat-label">Disabled</div>
                <div class="stat-value orange">{{ stats.disabled }}</div>
            </div>
        </div>

        <!-- Users Table -->
        <div class="table-container">
            <div class="table-header">
                <h5><i class="bi bi-list-ul"></i> Registered Devices</h5>
                <span class="text-muted small">{{ users|length }} user(s)</span>
            </div>
            <div class="table-responsive">
                <table class="table table-hover align-middle mb-0">
                    <thead>
                        <tr>
                            <th>#</th>
                            <th>Name</th>
                            <th>Phone / MAC</th>
                            <th>Location</th>
                            <th>Registered</th>
                            <th>Expires</th>
                            <th>Days Left</th>
                            <th>Status</th>
                            <th>Actions</th>
                        </tr>
                    </thead>
                    <tbody>
                        {% for u in users %}
                        {% set expiry = u.expiry_date %}
                        {% set days_left = u.days_left %}
                        {% set status_class = 'badge-active' if u.is_active and days_left > 0 else ('badge-expired' if days_left == 0 else 'badge-disabled') %}
                        {% set status_text = 'Active' if u.is_active and days_left > 0 else ('Expired' if days_left == 0 else 'Disabled') %}
                        {% set bar_pct = (days_left / 30 * 100)|round|int if days_left <= 30 else 100 %}
                        {% set bar_color = '#00b894' if days_left > 15 else ('#fdcb6e' if days_left > 5 else '#e94560') %}
                        <tr>
                            <td class="text-muted">{{ loop.index }}</td>
                            <td class="fw-bold">{{ u.name }}</td>
                            <td>
                                {{ u.phone }}<br>
                                <small class="text-muted" style="font-size:0.7rem;"><i class="bi bi-hash"></i> {{ u.mac_id[:20] }}...</small>
                            </td>
                            <td>{{ u.state }} > {{ u.district }} > {{ u.block }}</td>
                            <td>{{ u.registration_date }}</td>
                            <td>
                                {% if expiry %}
                                    <span class="{% if days_left <= 7 %}text-danger fw-bold{% endif %}">{{ expiry }}</span>
                                {% else %}-{% endif %}
                            </td>
                            <td>
                                {% if days_left > 0 %}
                                    <div style="display:flex;align-items:center;gap:8px;">
                                        <div class="days-bar">
                                            <div class="days-bar-fill" style="width:{{ bar_pct }}%;background:{{ bar_color }};"></div>
                                        </div>
                                        <span class="fw-bold" style="color:{{ bar_color }};">{{ days_left }}d</span>
                                    </div>
                                {% else %}
                                    <span class="text-muted">0 days</span>
                                {% endif %}
                            </td>
                            <td><span class="badge-status {{ status_class }}">{{ status_text }}</span></td>
                            <td>
                                <div style="display:flex;gap:6px;flex-wrap:wrap;">
                                    <a href="/admin/toggle/{{ u.mac_id }}" class="btn-action btn-toggle">
                                        <i class="bi bi-toggle-off"></i> Toggle
                                    </a>
                                    <a href="/admin/extend/{{ u.mac_id }}" class="btn-action btn-extend">
                                        <i class="bi bi-plus-circle"></i> +30 Days
                                    </a>
                                </div>
                            </td>
                        </tr>
                        {% endfor %}
                        {% if users|length == 0 %}
                        <tr>
                            <td colspan="9">
                                <div class="empty-state">
                                    <i class="bi bi-inbox"></i>
                                    <p class="mt-2">No registered users yet.</p>
                                </div>
                            </td>
                        </tr>
                        {% endif %}
                    </tbody>
                </table>
            </div>
        </div>
        <div class="text-center mt-3">
            <small class="text-muted">NMMS Tracking Report - License Management Server</small>
        </div>
    </div>
</body>
</html>
"""

# ==========================================
# ROUTES
# ==========================================

@app.route('/register')
def register_page():
    mac_id = request.args.get('mac')
    return render_template_string(LANDING_PAGE_HTML, mac_id=mac_id)


@app.route('/register_user', methods=['POST'])
def register_user():
    data = request.form
    reg_date = datetime.now().date()
    try:
        with get_db() as conn:
            c = conn.cursor()
            c.execute(
                "INSERT INTO users (mac_id, name, phone, state, district, block, registration_date, is_active) "
                "VALUES (%s, %s, %s, %s, %s, %s, %s, %s)",
                (data['mac_id'], data['name'].upper(), data['phone'],
                 data['state'].upper(), data['district'].upper(), data['block'].upper(),
                 reg_date, True)
            )
        return "<h2 style='text-align:center; color:green; margin-top:50px; font-family:sans-serif;'>Registration Successful!</h2>"
    except psycopg2.errors.UniqueViolation:
        return "<h2 style='text-align:center; color:red; margin-top:50px; font-family:sans-serif;'>Device Already Registered!</h2>"
    except Exception as e:
        return f"<h2 style='text-align:center; color:red; margin-top:50px; font-family:sans-serif;'>Error: {str(e)}</h2>"


@app.route('/api/check_status', methods=['GET'])
def check_status():
    mac_id = request.args.get('mac')
    try:
        with get_db() as conn:
            c = conn.cursor(cursor_factory=psycopg2.extras.DictCursor)
            c.execute(
                "SELECT name, state, district, block, registration_date, is_active FROM users WHERE mac_id=%s",
                (mac_id,)
            )
            row = c.fetchone()

        if not row:
            return jsonify({"registered": False, "active": False})

        reg_date = row['registration_date']
        is_active = row['is_active']
        days_left = get_days_left(reg_date)

        if days_left == 0:
            is_active = False

        return jsonify({
            "registered": True,
            "active": bool(is_active),
            "days_left": days_left,
            "user_data": {
                "name": row['name'],
                "state": row['state'],
                "district": row['district'],
                "block": row['block']
            }
        })
    except Exception as e:
        return jsonify({"error": f"Server error: {str(e)}"}), 500


@app.route('/admin')
def admin_dashboard():
    try:
        with get_db() as conn:
            c = conn.cursor(cursor_factory=psycopg2.extras.DictCursor)
            c.execute("SELECT * FROM users ORDER BY registration_date DESC")
            rows = c.fetchall()

        users = []
        now = datetime.now().strftime('%d %b %Y, %I:%M %p')
        stats = {"total": 0, "active": 0, "expired": 0, "disabled": 0}

        for row in rows:
            days = get_days_left(row['registration_date'])
            is_active = row['is_active'] and days > 0
            expiry = get_expiry_date(row['registration_date'])

            users.append({
                'mac_id': row['mac_id'],
                'name': row['name'],
                'phone': row['phone'],
                'state': row['state'],
                'district': row['district'],
                'block': row['block'],
                'registration_date': row['registration_date'].strftime('%d %b %Y') if row['registration_date'] else '-',
                'expiry_date': expiry.strftime('%d %b %Y') if expiry else '-',
                'is_active': row['is_active'],
                'days_left': days
            })

            stats['total'] += 1
            if is_active:
                stats['active'] += 1
            elif not row['is_active']:
                stats['disabled'] += 1
            else:
                stats['expired'] += 1

        return render_template_string(ADMIN_PAGE_HTML, users=users, stats=stats, now=now)
    except Exception as e:
        return f"<h2>Database Error: {str(e)}</h2>", 500


@app.route('/admin/toggle/<mac>')
def admin_toggle(mac):
    with get_db() as conn:
        c = conn.cursor()
        c.execute("UPDATE users SET is_active = NOT is_active WHERE mac_id=%s", (mac,))
    return redirect(url_for('admin_dashboard'))


@app.route('/admin/extend/<mac>')
def admin_extend(mac):
    """
    Extend user license by TRIAL_DAYS from the current expiry date.
    This means if a user has 5 days left, they'll get 35 days total.
    """
    with get_db() as conn:
        c = conn.cursor()
        # Add TRIAL_DAYS to the existing registration_date, effectively extending the expiry
        c.execute(
            "UPDATE users SET registration_date = registration_date + INTERVAL %s, is_active = TRUE WHERE mac_id=%s",
            (f'{TRIAL_DAYS} days', mac)
        )
    return redirect(url_for('admin_dashboard'))


@app.route('/health')
def health_check():
    """Health check endpoint for Docker container monitoring."""
    try:
        with get_db() as conn:
            c = conn.cursor()
            c.execute("SELECT 1")
        return jsonify({
            "status": "healthy",
            "timestamp": datetime.now().isoformat(),
            "database": "postgresql",
            "port": SERVER_PORT
        })
    except Exception as e:
        return jsonify({"status": "unhealthy", "error": str(e)}), 500


if __name__ == '__main__':
    init_db()
    app.run(host=SERVER_HOST, port=SERVER_PORT, debug=DEBUG_MODE)
