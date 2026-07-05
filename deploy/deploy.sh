#!/bin/bash
# Run this script ON the VPS (187.124.15.14) as root:
#   curl -sSL https://raw.githubusercontent.com/Refaat1942/Lotus-Offer-Generator/main/deploy/deploy.sh | bash

set -e

APP_DIR="/opt/lotus-offer-generator"
PORT=17800
REPO="https://github.com/Refaat1942/Lotus-Offer-Generator.git"

echo "=== Lotus Offer Generator Deployment ==="

apt-get update -qq
apt-get install -y -qq python3 python3-pip python3-venv git

if [ -d "$APP_DIR/.git" ]; then
    cd "$APP_DIR"
    git pull origin main
else
    rm -rf "$APP_DIR"
    git clone "$REPO" "$APP_DIR"
    cd "$APP_DIR"
fi

python3 -m venv venv
source venv/bin/activate
pip install --upgrade pip -q
pip install -r requirements.txt -q

mkdir -p data static/uploads

cp deploy/lotus-offer.service /etc/systemd/system/
systemctl daemon-reload
systemctl enable lotus-offer
systemctl restart lotus-offer

# Open firewall port if ufw is active
if command -v ufw &>/dev/null && ufw status | grep -q "Status: active"; then
    ufw allow $PORT/tcp
fi

echo ""
echo "=== Deployment complete ==="
echo "App URL: http://187.124.15.14:$PORT"
echo "Login: admin / admin"
systemctl status lotus-offer --no-pager
