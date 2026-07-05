#!/bin/bash
set -e

APP_DIR="/opt/lotus-offer-generator"
PORT=17800

echo "=== Lotus Offer Generator Deployment ==="

# Install system dependencies
apt-get update -qq
apt-get install -y -qq python3 python3-pip python3-venv git

# Clone or update repo
if [ -d "$APP_DIR/.git" ]; then
    cd "$APP_DIR"
    git pull origin main
else
    rm -rf "$APP_DIR"
    git clone https://github.com/Refaat1942/Lotus-Offer-Generator.git "$APP_DIR"
    cd "$APP_DIR"
fi

# Setup virtual environment
python3 -m venv venv
source venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt

# Create directories
mkdir -p data static/uploads

# Install systemd service
cp deploy/lotus-offer.service /etc/systemd/system/
systemctl daemon-reload
systemctl enable lotus-offer
systemctl restart lotus-offer

echo "=== Deployment complete ==="
echo "App running on port $PORT"
systemctl status lotus-offer --no-pager
