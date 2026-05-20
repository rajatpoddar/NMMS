#!/bin/bash
# ============================================
# NMMS Server - Docker Entrypoint
# ============================================
set -e

echo "========================================"
echo " NMMS Tracking Report - Server Starting"
echo "========================================"

# Wait for PostgreSQL to be ready
echo "[init] Waiting for PostgreSQL connection..."
export PGPASSWORD=${DB_PASSWORD:-nmms_password}
for i in $(seq 1 30); do
    if pg_isready -h ${DB_HOST:-nmms-db} -p ${DB_PORT:-5432} -U ${DB_USER:-nmms} > /dev/null 2>&1; then
        echo "[init] PostgreSQL is ready."
        break
    fi
    echo "[init] Waiting for PostgreSQL... attempt $i/30"
    sleep 2
done

# Initialize the database schema
echo "[init] Initializing database schema..."
python -c "
import server
server.init_db()
print('[init] Database schema initialized successfully.')
"

# Number of workers (default: 2, can be overridden)
WORKERS=${GUNICORN_WORKERS:-2}
TIMEOUT=${GUNICORN_TIMEOUT:-120}
SERVER_PORT=${SERVER_PORT:-6667}

echo "[server] Starting Gunicorn with $WORKERS workers..."
echo "[server] Listening on 0.0.0.0:${SERVER_PORT}"

exec gunicorn \
    --bind 0.0.0.0:${SERVER_PORT} \
    --workers $WORKERS \
    --timeout $TIMEOUT \
    --access-logfile - \
    --error-logfile - \
    --log-level ${LOG_LEVEL:-info} \
    --worker-tmp-dir /dev/shm \
    server:app
