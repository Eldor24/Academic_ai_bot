#!/usr/bin/env bash
# Talaba_Bot -> Oracle Linux 9 (opc foydalanuvchi).
set -euo pipefail
APP=/home/opc/talaba_bot
SVC=talaba-bot
ARX=/home/opc/talaba_bot.tar.gz

echo "== 1/5 Python 3.12 =="
rpm -q python3.12 >/dev/null 2>&1 || sudo dnf install -y python3.12 python3.12-pip tzdata

echo "== 2/5 Zaxira =="
mkdir -p "$APP" "$APP/zaxira"
if [ -f "$APP/users.json" ]; then
  cp "$APP/users.json" "$APP/zaxira/users_$(date +%F_%H%M).json"
  ls -1t "$APP"/zaxira/users_*.json | tail -n +21 | xargs -r rm -f     # oxirgi 20 tasi saqlanadi
fi
if tar -tzf "$ARX" | grep -qx './users.json'; then
  echo "!! Arxivda users.json bor — serverdagisi ALMASHTIRILADI (eski nusxa zaxira/ da)."
fi

echo "== 3/5 Fayllar =="
REQ_OLD=$(md5sum "$APP/requirements.txt" 2>/dev/null | cut -d' ' -f1 || true)
tar -xzf "$ARX" -C "$APP"
rm -f "$ARX"
[ -f "$APP/.env" ] && chmod 600 "$APP/.env" || echo "!! .env topilmadi — birinchi o'rnatishda -WithEnv bilan yuklang"

echo "== 4/5 Virtual muhit =="
REQ_NEW=$(md5sum "$APP/requirements.txt" | cut -d' ' -f1)
if [ ! -d "$APP/venv" ] || [ "$REQ_OLD" != "$REQ_NEW" ]; then
  [ -d "$APP/venv" ] || python3.12 -m venv "$APP/venv"
  "$APP/venv/bin/pip" install --no-cache-dir --upgrade pip
  "$APP/venv/bin/pip" install --no-cache-dir -r "$APP/requirements.txt"
else
  echo "requirements.txt o'zgarmagan — kutubxonalar qayta o'rnatilmaydi"
fi
# SELinux: systemd uy papkasidagi venv python'ni ishga tushira olishi uchun
sudo chcon -R -t bin_t "$APP/venv/bin" || true

echo "== 5/5 systemd xizmati =="
sudo tee /etc/systemd/system/$SVC.service >/dev/null <<UNIT
[Unit]
Description=Talaba (Academic) Telegram Bot
After=network-online.target
Wants=network-online.target

[Service]
Type=simple
User=opc
WorkingDirectory=$APP
Environment=PYTHONUNBUFFERED=1
Environment=PYTHONUTF8=1
ExecStart=$APP/venv/bin/python bot.py
Restart=always
RestartSec=10
# 1 GB RAM'li server: bot xotirani hammasini egallab, serverni qotirib qo'ymasligi uchun
MemoryHigh=650M
MemoryMax=850M
# to'xtatilganda bot users.json ni saqlab ulgurishi uchun
TimeoutStopSec=30

[Install]
WantedBy=multi-user.target
UNIT
sudo systemctl daemon-reload
sudo systemctl enable $SVC
sudo systemctl restart $SVC
sleep 8
sudo systemctl is-active $SVC || true
journalctl -u $SVC -n 25 --no-pager
free -m
