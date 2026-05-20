#!/bin/bash
# ============================================
# NMMS Tracker - Deployment Script
# Pulls latest code, backs up DB, stops
# containers, rebuilds, and restarts.
# Run this on your NAS to deploy updates.
# ============================================
set -e

# Colors
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m'

echo -e "${GREEN}"
echo "============================================"
echo "  NMMS Tracker - Deployment Script"
echo "  https://nmms.palojori.in"
echo "============================================"
echo -e "${NC}"

# 1. Navigate to script directory
cd "$(dirname "$0")"

# 2. Pull latest code
echo ""
echo -e "${YELLOW}[1/5] Pulling latest code from GitHub...${NC}"
git pull origin main
echo -e "${GREEN}  ✔ Code updated${NC}"

# 3. Backup database
echo ""
echo -e "${YELLOW}[2/5] Backing up PostgreSQL database...${NC}"
BACKUP_DIR="./backups"
mkdir -p "$BACKUP_DIR"
BACKUP_FILE="$BACKUP_DIR/nmms_backup_$(date +%Y%m%d_%H%M%S).sql"

if docker ps --format '{{.Names}}' 2>/dev/null | grep -q 'nmms-tracker-db'; then
    docker exec nmms-tracker-db pg_dump -U nmms nmms_tracker > "$BACKUP_FILE"
    echo -e "${GREEN}  ✔ Backup saved: $BACKUP_FILE${NC}"
else
    echo -e "${YELLOW}  ⚠ DB container not running — skipping backup${NC}"
fi

# 4. Stop running containers
echo ""
echo -e "${YELLOW}[3/5] Stopping running containers...${NC}"
docker compose down
echo -e "${GREEN}  ✔ Containers stopped${NC}"

# 5. Rebuild and start
echo ""
echo -e "${YELLOW}[4/5] Rebuilding and starting containers...${NC}"
docker compose up -d --build
echo -e "${GREEN}  ✔ Containers started${NC}"

# 6. Health check
echo ""
echo -e "${YELLOW}[5/5] Waiting for server health check...${NC}"
echo ""

for i in $(seq 1 15); do
    sleep 4
    if curl -sf http://localhost:6667/health > /dev/null 2>&1; then
        echo -e "${GREEN}"
        echo "============================================"
        echo "  ✔ Deployment Successful!"
        echo ""
        echo "  Server:     http://localhost:6667"
        echo "  Admin:      http://localhost:6667/admin"
        echo "  Domain:     https://nmms.palojori.in"
        echo ""
        echo "  To check logs:  docker compose logs -f"
        echo "============================================"
        echo -e "${NC}"
        exit 0
    fi
    echo "  Waiting for server... attempt $i/15"
done

echo -e "${RED}"
echo "============================================"
echo "  ✘ Deployment may have issues."
echo ""
echo "  Check logs:  docker compose logs"
echo "  Check DB:    docker compose logs nmms-db"
echo "============================================"
echo -e "${NC}"
exit 1
