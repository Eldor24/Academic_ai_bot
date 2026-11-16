# Serverga yuklash (xavfsiz usul)

> ⚠️ Server parolini hech qachon faylga yozmang va GitHub'ga yuklamang. Parol o'rniga SSH kalitdan foydalaning.

## 1. Bir martalik sozlash: parol o'rniga SSH kalit

Windows'da `cmd` yoki PowerShell oching:

```bat
ssh-keygen -t ed25519
type %USERPROFILE%\.ssh\id_ed25519.pub | ssh root@SERVER_IP "mkdir -p ~/.ssh && cat >> ~/.ssh/authorized_keys"
```

Serverda (`ssh root@SERVER_IP`):

```bash
passwd                                  # kuchli parol o'rnating
sed -i 's/^#\?PasswordAuthentication.*/PasswordAuthentication no/' /etc/ssh/sshd_config
systemctl restart ssh
apt install -y ufw fail2ban
ufw allow OpenSSH && ufw --force enable  # bot faqat chiquvchi ulanish ishlatadi, port ochish shart emas
```

## 2. Fayllarni yuklash (har yangilanishda)

Endi bot bir nechta fayldan iborat. Barchasini yuklang:

```bat
cd "C:\...\Academic_ai_bot"   & rem bot papkangiz
scp bot.py security.py pptx_themes.py xizmatlar_yangi.py requirements.txt root@SERVER_IP:/root/academic_bot/
scp -r assets root@SERVER_IP:/root/academic_bot/
```

`.env` faqat o'zgarganda yuklanadi:

```bat
scp .env root@SERVER_IP:/root/academic_bot/
```

## 3. Serverda kutubxonalarni yangilash va qayta ishga tushirish

```bash
cd /root/academic_bot
pip install -r requirements.txt        # (venv bo'lsa: ./venv/bin/pip install -r requirements.txt)
chmod 600 .env
sudo systemctl restart academicbot
sudo systemctl status academicbot
journalctl -u academicbot -n 50 --no-pager   # xatolarni ko'rish
```

> Muhim: `requirements.txt` da `httpx<0.28` bor. Usiz `openai==1.14.0` yangi serverda
> `unexpected keyword argument 'proxies'` xatosi bilan ishga tushmaydi.

## 4. (Tavsiya) systemd xizmatini qattiqlashtirish

`/etc/systemd/system/academicbot.service` ichida `[Service]` bo'limiga qo'shing:

```ini
Restart=always
RestartSec=5
NoNewPrivileges=true
PrivateTmp=true
ProtectSystem=full
ProtectHome=read-only
ReadWritePaths=/root/academic_bot
```

Keyin: `systemctl daemon-reload && systemctl restart academicbot`.

## 5. Yangilashdan oldin zaxira

```bash
cp /root/academic_bot/users.json /root/academic_bot/users_$(date +%F).json
```
