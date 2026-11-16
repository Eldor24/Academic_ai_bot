"""
cloud_db.py — users.json ni Telegram'ning o'zida saqlash
═══════════════════════════════════════════════════════════════════
Render kabi bepul hostinglarda disk vaqtinchalik: har qayta ishga tushish yoki
deploy'da users.json o'chib ketadi. Bu modul bazani Telegram chatiga fayl qilib
yuboradi va QADAB (pin) qo'yadi:

  • Bot ishga tushganda: lokal baza yo'q yoki bo'sh bo'lsa — qadalgan fayldan tiklaydi
  • Ishlash davomida: har 5 daqiqada baza o'zgargan bo'lsa — qadalgan faylni yangilaydi
    (yangi xabar yubormaydi, o'sha xabarni tahrirlaydi — chat to'lib ketmaydi)
  • Bot to'xtaganda: oxirgi holatni yuboradi

Qayerda saqlanadi: DB_CHAT_ID (.env) yoki u bo'lmasa — admin bilan shaxsiy chat.
Qachon yoqiladi: Render'da avtomatik (RENDER o'zgaruvchisi bor) yoki CLOUD_DB=1 bo'lsa.
"""
from __future__ import annotations

import hashlib
import io
import json
import logging
import os
import threading
import time

log = logging.getLogger(__name__)

CAPTION = "🗄 Academic Bot bazasi (avtomatik). Bu xabarni o'chirmang va qadalganini olib tashlamang!"
FILE_NAME = "users.json"


def enabled() -> bool:
    return bool(os.getenv("RENDER")) or os.getenv("CLOUD_DB", "").strip() in ("1", "true", "yes")


def chat_id(admin_id: int) -> int:
    raw = os.getenv("DB_CHAT_ID", "").strip()
    try:
        return int(raw) if raw else admin_id
    except ValueError:
        return admin_id


def _valid(data) -> bool:
    return isinstance(data, dict) and isinstance(data.get("all_users"), list) and isinstance(data.get("coins", {}), dict)


def _local_user_count(path: str) -> int:
    try:
        with open(path, "r", encoding="utf-8") as f:
            return len(json.load(f).get("all_users", []))
    except Exception:
        return 0


class CloudDB:
    def __init__(self, bot, chat: int, users_file: str):
        self.bot = bot
        self.chat = chat
        self.users_file = users_file
        self.msg_id: int | None = None
        self.last_hash: str | None = None
        self.lock = threading.Lock()

    # ── Tiklash ────────────────────────────────────────────────────
    def restore_if_needed(self) -> bool:
        """Lokal baza bo'sh bo'lsa, qadalgan fayldan tiklaydi. Tiklangan bo'lsa True."""
        try:
            pinned = self.bot.get_chat(self.chat).pinned_message
        except Exception as e:
            log.warning(f"☁️ Bulut bazani tekshirib bo'lmadi: {e}")
            return False
        if not pinned or not getattr(pinned, "document", None):
            log.info("☁️ Bulutda saqlangan baza hali yo'q.")
            return False
        self.msg_id = pinned.message_id
        local_n = _local_user_count(self.users_file)
        try:
            fi = self.bot.get_file(pinned.document.file_id)
            raw = self.bot.download_file(fi.file_path)
            data = json.loads(raw.decode("utf-8"))
        except Exception as e:
            log.error(f"☁️ Bulut bazani yuklab bo'lmadi: {e}")
            return False
        if not _valid(data):
            log.error("☁️ Bulutdagi fayl baza formatida emas — e'tiborsiz qoldirildi.")
            return False
        cloud_n = len(data.get("all_users", []))
        self.last_hash = hashlib.sha256(raw).hexdigest()
        if local_n >= cloud_n and local_n > 0:
            log.info(f"☁️ Lokal baza ({local_n}) bulutdagidan ({cloud_n}) kam emas — tiklash shart emas.")
            return False
        with open(self.users_file, "w", encoding="utf-8") as f:
            f.write(raw.decode("utf-8"))
        log.info(f"☁️ ✅ Baza bulutdan tiklandi: {cloud_n} foydalanuvchi.")
        return True

    # ── Saqlash ────────────────────────────────────────────────────
    def push(self, snapshot: bytes, force: bool = False) -> bool:
        """Baza o'zgargan bo'lsa Telegram'dagi qadalgan faylni yangilaydi."""
        from telebot.types import InputMediaDocument
        h = hashlib.sha256(snapshot).hexdigest()
        with self.lock:
            if not force and h == self.last_hash:
                return False
            f = io.BytesIO(snapshot)
            f.name = FILE_NAME
            try:
                if self.msg_id:
                    try:
                        self.bot.edit_message_media(
                            media=InputMediaDocument(f, caption=CAPTION),
                            chat_id=self.chat, message_id=self.msg_id)
                        self.last_hash = h
                        return True
                    except Exception as e:
                        if "not modified" in str(e).lower():
                            self.last_hash = h
                            return False
                        log.warning(f"☁️ Qadalgan faylni tahrirlab bo'lmadi, yangisi yuboriladi: {e}")
                        f.seek(0)
                m = self.bot.send_document(self.chat, f, caption=CAPTION, visible_file_name=FILE_NAME,
                                           disable_notification=True)
                self.msg_id = m.message_id
                try:
                    self.bot.pin_chat_message(self.chat, m.message_id, disable_notification=True)
                except Exception as e:
                    log.error(f"☁️ Faylni qadab bo'lmadi (kanal bo'lsa botga 'Pin' huquqini bering): {e}")
                self.last_hash = h
                log.info("☁️ Baza Telegram'ga saqlandi.")
                return True
            except Exception as e:
                log.error(f"☁️ Bazani Telegram'ga saqlab bo'lmadi: {e}")
                return False

    def start(self, get_snapshot, interval: int = 300):
        def _loop():
            while True:
                time.sleep(interval)
                try:
                    self.push(get_snapshot())
                except Exception as e:
                    log.error(f"☁️ Sinxronlash xatosi: {e}")
        threading.Thread(target=_loop, daemon=True).start()
        log.info(f"☁️ Bulut baza sinxronlash yoqildi: chat={self.chat}, har {interval // 60} daqiqada.")
