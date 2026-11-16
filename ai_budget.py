"""
ai_budget.py — AI tokenlari sarfini nazorat qilish
═══════════════════════════════════════════════════════════════════
Botdagi BARCHA AI chaqiruvlari (DeepSeek va Gemini) shu yerdan o'tadi:

  1. Har bir so'rovning token sarfi hisoblanadi (API qaytargan haqiqiy son)
  2. Kunlik umumiy limit (AI_DAILY_TOKENS). 80% ga yetganda adminga ogohlantirish,
     100% da yangi AI ishlar to'xtatiladi (boshlangan ishlar oxiriga yetkaziladi)
  3. Har bir foydalanuvchi uchun kunlik AI ishlar limiti (AI_USER_DAILY_JOBS)
  4. Bitta so'rovda javob uzunligi (max_tokens) cheklanadi
  5. Juda katta so'rov (prompt) umuman yuborilmaydi
  6. /aistat — admin uchun bugungi va oxirgi 7 kunlik statistika

Hisob users_db["ai_usage"] ichida saqlanadi (bulut bazasi bilan birga sinxronlanadi).
Sozlash (.env, ixtiyoriy):
    AI_DAILY_TOKENS=3000000      kunlik umumiy token limiti
    AI_USER_DAILY_JOBS=25        bitta foydalanuvchi kuniga nechta AI ish boshlay oladi
    AI_MAX_TOKENS_PER_CALL=8000  bitta javobning maksimal uzunligi
    AI_MAX_PROMPT_CHARS=100000   bitta so'rovning maksimal hajmi (belgi)
"""
from __future__ import annotations

import datetime
import logging
import os
import threading

log = logging.getLogger(__name__)


def _env_int(name: str, default: int) -> int:
    try:
        return int(os.getenv(name, "").strip() or default)
    except ValueError:
        return default


DAILY_TOKENS = _env_int("AI_DAILY_TOKENS", 3_000_000)
USER_DAILY_JOBS = _env_int("AI_USER_DAILY_JOBS", 25)
MAX_TOKENS_PER_CALL = _env_int("AI_MAX_TOKENS_PER_CALL", 8000)
MAX_PROMPT_CHARS = _env_int("AI_MAX_PROMPT_CHARS", 100_000)
HARD_CEILING = 1.25          # limitdan 25% oshsa — hatto boshlangan ishlar ham to'xtatiladi

LIMIT_TXT = {
    "user": {
        "uz": "⏳ Bugungi AI so'rovlar limitingiz ({n} ta) tugadi. Ertaga yana davom etamiz! 🙏",
        "ru": "⏳ Ваш дневной лимит ИИ-запросов ({n}) исчерпан. Продолжим завтра! 🙏",
        "en": "⏳ You've reached today's AI request limit ({n}). See you tomorrow! 🙏",
    },
    "global": {
        "uz": "⏳ Bot bugun juda band — kunlik AI quvvati tugadi. Iltimos, ertaga urinib ko'ring. 🙏",
        "ru": "⏳ Бот сегодня перегружен — дневной ресурс ИИ исчерпан. Попробуйте завтра. 🙏",
        "en": "⏳ The bot is very busy today — the daily AI capacity is used up. Please try tomorrow. 🙏",
    },
}


class BudgetExceeded(RuntimeError):
    pass


class AIBudget:
    def __init__(self, get_db, mark_dirty, db_lock, admin_id: int, notify=None, tz=None):
        self.get_db = get_db
        self.mark_dirty = mark_dirty
        self.db_lock = db_lock
        self.admin_id = admin_id
        self.notify = notify
        self.tz = tz
        self.lock = threading.Lock()
        self.denied: dict[int, str] = {}       # uid -> sabab (oxirgi rad etish)

    # ── Kunlik hisob ──────────────────────────────────────────────
    def _today(self) -> str:
        now = datetime.datetime.now(self.tz) if self.tz else datetime.datetime.now()
        return now.strftime("%Y-%m-%d")

    def _usage(self) -> dict:
        """users_db["ai_usage"] — sana o'zgarganda yangi kun boshlanadi."""
        db = self.get_db()
        u = db.setdefault("ai_usage", {})
        today = self._today()
        if u.get("sana") != today:
            if u.get("sana"):
                tarix = u.setdefault("tarix", {})
                tarix[u["sana"]] = {"tokens": u.get("tokens", 0), "sorovlar": u.get("sorovlar", 0),
                                    "ishlar": sum(u.get("jobs", {}).values())}
                for k in sorted(tarix)[:-30]:
                    tarix.pop(k, None)
            u.update({"sana": today, "tokens": 0, "kirish": 0, "chiqish": 0, "sorovlar": 0,
                      "jobs": {}, "ogoh80": False, "ogoh100": False})
        return u

    # ── Ishni boshlashga ruxsat ───────────────────────────────────
    def allow_job(self, uid: int) -> bool:
        with self.db_lock:
            u = self._usage()
            if u["tokens"] >= DAILY_TOKENS and uid != self.admin_id:
                self.denied[uid] = "global"
                return False
            jobs = u.setdefault("jobs", {})
            key = str(uid)
            if uid != self.admin_id and jobs.get(key, 0) >= USER_DAILY_JOBS:
                self.denied[uid] = "user"
                return False
            jobs[key] = jobs.get(key, 0) + 1
        self.denied.pop(uid, None)
        self.mark_dirty()
        return True

    def deny_text(self, uid: int, til: str) -> str | None:
        """Oxirgi rad etish sababi bo'yicha foydalanuvchiga xabar (bo'lmasa None)."""
        why = self.denied.pop(uid, None)
        if not why:
            return None
        t = LIMIT_TXT[why].get(til, LIMIT_TXT[why]["uz"])
        return t.format(n=USER_DAILY_JOBS)

    # ── Token hisobi ──────────────────────────────────────────────
    def _check_hard(self):
        with self.db_lock:
            if self._usage()["tokens"] >= DAILY_TOKENS * HARD_CEILING:
                raise BudgetExceeded("Kunlik AI token limiti keskin oshib ketdi — so'rov to'xtatildi")

    def record(self, tokens_in: int, tokens_out: int, model: str = ""):
        alert = None
        with self.db_lock:
            u = self._usage()
            u["kirish"] = u.get("kirish", 0) + int(tokens_in or 0)
            u["chiqish"] = u.get("chiqish", 0) + int(tokens_out or 0)
            u["tokens"] = u["kirish"] + u["chiqish"]
            u["sorovlar"] = u.get("sorovlar", 0) + 1
            m = u.setdefault("modellar", {})
            m[model or "?"] = m.get(model or "?", 0) + int(tokens_in or 0) + int(tokens_out or 0)
            if u["tokens"] >= DAILY_TOKENS and not u.get("ogoh100"):
                u["ogoh100"] = True
                alert = f"🛑 <b>AI kunlik limiti tugadi</b>: {u['tokens']:,} / {DAILY_TOKENS:,} token. Yangi AI ishlar ertagacha to'xtatildi."
            elif u["tokens"] >= DAILY_TOKENS * 0.8 and not u.get("ogoh80"):
                u["ogoh80"] = True
                alert = f"⚠️ <b>AI sarfi 80% ga yetdi</b>: {u['tokens']:,} / {DAILY_TOKENS:,} token."
        self.mark_dirty()
        if alert and self.notify:
            try:
                self.notify(self.admin_id, alert.replace(",", " "))
            except Exception:
                pass

    # ── Klientlarni o'rash ────────────────────────────────────────
    def wrap_openai(self, client):
        comp = client.chat.completions
        orig = comp.create

        def create(*args, **kwargs):
            msgs = kwargs.get("messages") or []
            size = sum(len(str(m.get("content", ""))) for m in msgs if isinstance(m, dict))
            if size > MAX_PROMPT_CHARS:
                raise ValueError(f"AI so'rovi juda katta ({size} belgi > {MAX_PROMPT_CHARS})")
            if kwargs.get("max_tokens"):
                kwargs["max_tokens"] = min(int(kwargs["max_tokens"]), MAX_TOKENS_PER_CALL)
            self._check_hard()
            resp = orig(*args, **kwargs)
            try:
                us = resp.usage
                self.record(us.prompt_tokens, us.completion_tokens, kwargs.get("model", "deepseek"))
            except Exception:
                self.record(size // 3, 0, kwargs.get("model", "deepseek"))
            return resp

        comp.create = create
        log.info("💰 AI token nazorati: DeepSeek ulandi")

    def wrap_genai(self, client):
        if client is None:
            return
        models = client.models
        orig = models.generate_content

        def generate_content(*args, **kwargs):
            self._check_hard()
            resp = orig(*args, **kwargs)
            try:
                md = resp.usage_metadata
                self.record(md.prompt_token_count or 0, (md.candidates_token_count or 0) +
                            (getattr(md, "thoughts_token_count", 0) or 0), kwargs.get("model", "gemini"))
            except Exception:
                self.record(1000, 500, "gemini")
            return resp

        models.generate_content = generate_content
        log.info("💰 AI token nazorati: Gemini ulandi")

    # ── Statistika ────────────────────────────────────────────────
    def stats_text(self) -> str:
        with self.db_lock:
            u = dict(self._usage())
            tarix = dict(u.get("tarix", {}))
            jobs = dict(u.get("jobs", {}))
            modellar = dict(u.get("modellar", {}))
        pct = u["tokens"] / DAILY_TOKENS * 100 if DAILY_TOKENS else 0
        top = sorted(jobs.items(), key=lambda kv: -kv[1])[:5]
        lines = [
            "💰 <b>AI sarfi — bugun</b> (" + u["sana"] + ")",
            f"🔢 Tokenlar: <b>{u['tokens']:,}</b> / {DAILY_TOKENS:,} ({pct:.1f}%)",
            f"   ↳ kirish {u.get('kirish', 0):,} • chiqish {u.get('chiqish', 0):,}",
            f"📨 AI so'rovlar: <b>{u.get('sorovlar', 0)}</b> • ishlar: <b>{sum(jobs.values())}</b>",
        ]
        if modellar:
            lines.append("🤖 " + " • ".join(f"{k}: {v:,}" for k, v in modellar.items()))
        if top:
            lines.append("\n👤 <b>Eng faol (ishlar soni):</b>")
            lines += [f"  <code>{k}</code> — {v} ta" for k, v in top]
        if tarix:
            lines.append("\n📅 <b>Oxirgi kunlar:</b>")
            for d in sorted(tarix)[-7:]:
                t = tarix[d]
                lines.append(f"  {d}: {t.get('tokens', 0):,} token • {t.get('ishlar', 0)} ish")
        lines.append(f"\n⚙️ Limitlar: kuniga {DAILY_TOKENS:,} token • foydalanuvchiga {USER_DAILY_JOBS} ish • "
                     f"javob ≤ {MAX_TOKENS_PER_CALL} token")
        return "\n".join(lines).replace(",", " ")
