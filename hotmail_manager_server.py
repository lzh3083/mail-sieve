#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Hotmail / Outlook 批量邮箱获取与深度取证审查管理面板后端服务
功能：
  1. 从指定邮箱（IMAP/POP3/Outlook OAuth2/Cloudflare 临时邮箱）自动拉取发货邮件并提取购买的 Hotmail/Outlook 卡密
  2. 对获取的邮箱进行多线程并发深度审查（OAuth2 Token 测活 + IMAP/Graph 收件箱历史邮件扫描）
  3. 自动识别：
     - 是否为号商“改别名翻新（Alias-Recycled）”二手机箱（检测前身别名、Thunderbird 连接掩码不匹配、添加登录名记录）
     - 是否有 OpenAI / ChatGPT 封号黑历史（Access Deactivated）或风控锁号通知（Account Locked）
     - 历史注册过哪些服务（OpenAI/ChatGPT, xAI/Grok, Claude, Cursor, X/Twitter, Apple ID, 1040.com, Uber, Google, GitHub 等）
     - 评估能否用于注册目标业务（如 OpenAI / ChatGPT / Grok 等）
  4. 一键生成《号商翻新与黑历史取证报告》（Markdown）以及一键同步可用邮箱至 gptcli-2api 注册机邮箱池
"""
from __future__ import annotations

import argparse
import base64
import calendar
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timezone
import email as email_lib
from email.header import decode_header
from email.utils import parseaddr, parsedate_to_datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import imaplib
import json
import os
from pathlib import Path
import poplib
import re
import sqlite3
import sys
import threading
import time
import traceback
from typing import Any
import urllib.error
import urllib.parse
import urllib.request

WORKSPACE_DIR = Path(__file__).resolve().parent


def _resolve_gptcli_root() -> Path:
    env_dir = os.environ.get("GPTCLI_ROOT_DIR", "").strip()
    if env_dir and Path(env_dir).is_dir():
        return Path(env_dir)
    for cand in (
        WORKSPACE_DIR / "gptcli-2api",
        WORKSPACE_DIR.parent / "gptcli-2api",
        Path("/www/wwwroot/gptcli-2api"),
    ):
        if cand.is_dir():
            return cand
    return WORKSPACE_DIR / "gptcli-2api"


GPTCLI_ROOT = _resolve_gptcli_root()
DB_PATH = Path(os.environ.get("MAIL_SIEVE_DB_PATH") or (WORKSPACE_DIR / "hotmail_inventory.sqlite3"))
DASHBOARD_HTML_PATH = WORKSPACE_DIR / "hotmail-dashboard.html"
GPTCLI_OUTLOOK_TXT = GPTCLI_ROOT / "grok-build-auth" / "openai_client" / "用于注册的邮箱.txt"
GPTCLI_POOL_TXT = GPTCLI_ROOT / "output" / "mailboxes" / "outlook-accounts.txt"
GPTCLI_SQLITE_DB = GPTCLI_ROOT / "grok-build-auth" / "openai_client" / "turb.sqlite3"

_DB_LOCK = threading.RLock()
_AUDIT_STATE_LOCK = threading.Lock()
_AUDIT_PROGRESS: dict[str, Any] = {
    "running": False,
    "total": 0,
    "completed": 0,
    "current_email": "",
    "started_at": 0.0,
    "finished_at": 0.0,
    "target_service": "openai",
    "summary": {},
}

# 常见收货邮箱 IMAP 服务器自动映射
COMMON_IMAP_SERVERS: dict[str, tuple[str, int]] = {
    "qq.com": ("imap.qq.com", 993),
    "foxmail.com": ("imap.qq.com", 993),
    "163.com": ("imap.163.com", 993),
    "126.com": ("imap.126.com", 993),
    "yeah.net": ("imap.yeah.net", 993),
    "gmail.com": ("imap.gmail.com", 993),
    "outlook.com": ("outlook.office365.com", 993),
    "hotmail.com": ("outlook.office365.com", 993),
    "live.com": ("outlook.office365.com", 993),
    "msn.com": ("outlook.office365.com", 993),
    "aliyun.com": ("imap.aliyun.com", 993),
    "sina.com": ("imap.sina.com", 993),
    "sohu.com": ("imap.sohu.com", 993),
}

# 服务识别规则库
SERVICE_RULES: list[tuple[str, str, list[str], list[str]]] = [
    (
        "openai",
        "OpenAI / ChatGPT",
        ["openai.com", "chatgpt.com"],
        [
            "chatgpt",
            "openai",
            "your temporary chatgpt",
            "access deactivated",
            "secure your chatgpt account",
        ],
    ),
    (
        "grok",
        "xAI / Grok",
        ["x.ai", "grok.com"],
        ["grok", "x.ai", "xai verification"],
    ),
    (
        "claude",
        "Anthropic / Claude",
        ["anthropic.com", "claude.ai"],
        ["claude", "anthropic"],
    ),
    (
        "cursor",
        "Cursor",
        ["cursor.com", "cursor.sh"],
        ["sign up for cursor", "connect your repos to cursor", "cursor limits"],
    ),
    (
        "x_twitter",
        "X (Twitter)",
        ["x.com", "twitter.com", "e.x.com"],
        [
            "x verification code",
            "codice di convalida per x",
            "x two-factor authentication",
            "benvenuto su x",
            "welcome to x",
        ],
    ),
    (
        "apple",
        "Apple ID",
        ["apple.com", "id.apple.com", "insideapple.apple.com"],
        ["apple account", "apple id", "apple‑kontobegäran", "verify your apple"],
    ),
    (
        "tax_1040",
        "1040.com (美国报税)",
        ["1040.com"],
        ["1040.com", "irs rejected your return", "e-filing with 1040.com"],
    ),
    (
        "uber",
        "Uber / Uber Eats",
        ["uber.com"],
        ["uber eats", "uber code", "your uber"],
    ),
    (
        "google",
        "Google",
        ["google.com", "accounts.google.com"],
        ["google verification", "google account"],
    ),
    (
        "github",
        "GitHub",
        ["github.com"],
        ["[github]", "github launch code", "github verification"],
    ),
    (
        "discord",
        "Discord",
        ["discord.com"],
        ["discord", "verify email address for discord"],
    ),
    (
        "telegram",
        "Telegram",
        ["telegram.org"],
        ["telegram"],
    ),
    (
        "hk_ticketing",
        "HK Ticketing (快达票)",
        ["hkticketing.com"],
        ["hk ticketing"],
    ),
    (
        "amazon",
        "Amazon (亚马逊)",
        ["amazon.com", "amazon.co.jp", "amazon.co.uk"],
        ["amazonアカウント", "amazon account", "your amazon verification"],
    ),
    (
        "dragonnest",
        "Dragon Nest (龙之谷)",
        ["dragonnest.com"],
        ["[dnsea]", "dragon nest"],
    ),
]


def _now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def get_db_conn() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH), timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    return conn


def init_db() -> None:
    with _DB_LOCK:
        conn = get_db_conn()
        try:
            conn.executescript(
                """
                CREATE TABLE IF NOT EXISTS mailboxes (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    email TEXT NOT NULL UNIQUE COLLATE NOCASE,
                    password TEXT NOT NULL DEFAULT '',
                    client_id TEXT NOT NULL DEFAULT '',
                    refresh_token TEXT NOT NULL DEFAULT '',
                    recovery_email TEXT NOT NULL DEFAULT '',
                    recovery_code TEXT NOT NULL DEFAULT '',
                    batch_no TEXT NOT NULL DEFAULT '',
                    source_info TEXT NOT NULL DEFAULT '',
                    purpose TEXT NOT NULL DEFAULT '',
                    notes TEXT NOT NULL DEFAULT '',
                    audit_status TEXT NOT NULL DEFAULT 'pending',
                    can_register_target INTEGER NOT NULL DEFAULT 0,
                    auth_channel TEXT NOT NULL DEFAULT '',
                    auth_error TEXT NOT NULL DEFAULT '',
                    msg_count INTEGER NOT NULL DEFAULT 0,
                    earliest_msg_date TEXT NOT NULL DEFAULT '',
                    latest_msg_date TEXT NOT NULL DEFAULT '',
                    registered_services TEXT NOT NULL DEFAULT '[]',
                    risk_flags TEXT NOT NULL DEFAULT '[]',
                    former_aliases TEXT NOT NULL DEFAULT '[]',
                    messages_sample TEXT NOT NULL DEFAULT '[]',
                    last_audited_at TEXT NOT NULL DEFAULT '',
                    created_at TEXT NOT NULL DEFAULT '',
                    updated_at TEXT NOT NULL DEFAULT ''
                );
                CREATE INDEX IF NOT EXISTS idx_mailboxes_status ON mailboxes(audit_status);
                CREATE INDEX IF NOT EXISTS idx_mailboxes_batch ON mailboxes(batch_no);

                CREATE TABLE IF NOT EXISTS source_mailbox_config (
                    id INTEGER PRIMARY KEY CHECK (id = 1),
                    protocol TEXT NOT NULL DEFAULT 'imap',
                    email TEXT NOT NULL DEFAULT '',
                    password TEXT NOT NULL DEFAULT '',
                    imap_server TEXT NOT NULL DEFAULT '',
                    imap_port INTEGER NOT NULL DEFAULT 993,
                    use_ssl INTEGER NOT NULL DEFAULT 1,
                    client_id TEXT NOT NULL DEFAULT '',
                    refresh_token TEXT NOT NULL DEFAULT '',
                    cf_api_base TEXT NOT NULL DEFAULT 'https://apimail.jgxjs.com',
                    cf_api_key TEXT NOT NULL DEFAULT '',
                    cf_address TEXT NOT NULL DEFAULT '',
                    search_keyword TEXT NOT NULL DEFAULT '',
                    max_mails INTEGER NOT NULL DEFAULT 30,
                    updated_at TEXT NOT NULL DEFAULT ''
                );
                """
            )
            row = conn.execute("SELECT 1 FROM source_mailbox_config WHERE id = 1").fetchone()
            if not row:
                conn.execute(
                    """
                    INSERT INTO source_mailbox_config (
                        id, protocol, email, password, imap_server, imap_port, use_ssl,
                        client_id, refresh_token, cf_api_base, cf_api_key, cf_address,
                        search_keyword, max_mails, updated_at
                    ) VALUES (1, 'imap', '', '', '', 993, 1, '', '', 'https://apimail.jgxjs.com', '', '', '', 30, ?)
                    """,
                    (_now_iso(),),
                )
            conn.commit()
        finally:
            conn.close()


def _decode_mime_header(value: str | None) -> str:
    if not value:
        return ""
    parts: list[str] = []
    for part, charset in decode_header(str(value)):
        if isinstance(part, bytes):
            try:
                parts.append(part.decode(charset or "utf-8", errors="replace"))
            except Exception:
                parts.append(part.decode("utf-8", errors="replace"))
        else:
            parts.append(str(part))
    return "".join(parts).strip()


def _extract_email_body_and_attachments(msg: email_lib.message.Message) -> tuple[str, list[tuple[str, str]]]:
    """提取邮件纯文本/HTML正文以及所有文本类附件内容。"""
    text_parts: list[str] = []
    html_parts: list[str] = []
    attachments: list[tuple[str, str]] = []

    for part in (msg.walk() if msg.is_multipart() else [msg]):
        if part.is_multipart():
            continue
        ctype = (part.get_content_type() or "").lower()
        filename = _decode_mime_header(part.get_filename())
        payload = part.get_payload(decode=True)
        if payload is None:
            raw = part.get_payload()
            decoded = raw if isinstance(raw, str) else ""
        else:
            charset = part.get_content_charset() or "utf-8"
            try:
                decoded = payload.decode(charset, errors="replace")
            except Exception:
                decoded = payload.decode("utf-8", errors="replace")

        if filename:
            if any(filename.lower().endswith(ext) for ext in (".txt", ".csv", ".log", ".json", ".md")):
                attachments.append((filename, decoded))
            continue

        if ctype == "text/plain":
            text_parts.append(decoded)
        elif ctype == "text/html":
            html_parts.append(decoded)

    body = "\n".join(text_parts)
    if html_parts:
        html_raw = "\n".join(html_parts)
        html_clean = re.sub(r"<style[^>]*>.*?</style>", " ", html_raw, flags=re.S | re.I)
        html_clean = re.sub(r"<script[^>]*>.*?</script>", " ", html_clean, flags=re.S | re.I)
        html_clean = re.sub(r"<br\s*/?>|</p>|</div>|</li>|</tr>", "\n", html_clean, flags=re.I)
        html_clean = re.sub(r"<[^>]+>", " ", html_clean)
        body = (body + "\n" + html_clean).strip()
    return body, attachments


def _clean_snippet(raw_text: str, max_len: int = 240) -> str:
    clean = re.sub(r"<style[^>]*>.*?</style>", " ", str(raw_text or ""), flags=re.S | re.I)
    clean = re.sub(r"<script[^>]*>.*?</script>", " ", clean, flags=re.S | re.I)
    clean = re.sub(r"<[^>]+>", " ", clean)
    clean = re.sub(r"\s+", " ", clean).strip()
    return clean[:max_len]


def _format_date_short(date_str: str) -> str:
    if not date_str:
        return ""
    try:
        dt = parsedate_to_datetime(date_str)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.astimezone(timezone.utc).strftime("%Y-%m-%d %H:%M")
    except Exception:
        pass
    try:
        dt = datetime.fromisoformat(date_str.replace("Z", "+00:00"))
        return dt.astimezone(timezone.utc).strftime("%Y-%m-%d %H:%M")
    except Exception:
        return str(date_str)[:19]


# ============================================================
# 卡密/邮箱解析引擎：支持从邮件正文、附件或粘贴文本中提取购买的邮箱
# ============================================================

_EMAIL_RE = re.compile(r"^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$")
_BATCH_ORDER_RE = re.compile(r"\b(LD[0-9A-Z]{8,20}|ORD[-_]?[0-9A-Z]{6,20}|\d{14,22})\b")


def parse_purchased_accounts_text(
    raw_text: str,
    default_batch: str = "",
    default_source: str = "手动导入",
    allow_email_only: bool = True,
) -> list[dict[str, str]]:
    """
    从文本（发货邮件正文、txt/csv附件、手动粘贴）中提取购买的 Hotmail/Outlook 邮箱信息。
    支持格式：
      1. email----password----client_id----refresh_token (4/5/6段，支持 ----, ====, |, \t, 逗号)
      2. email----password 或 email:password
      3. 纯 email 列表（当 allow_email_only=True 时）
    """
    results: dict[str, dict[str, str]] = {}
    text = str(raw_text or "").replace("\r\n", "\n").replace("\r", "\n")

    # 尝试从文本中提取订单号/批次号
    detected_batch = default_batch
    if not detected_batch:
        m_order = _BATCH_ORDER_RE.search(text)
        if m_order:
            detected_batch = m_order.group(1)

    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue

        # 去掉常见列表前缀，如 "1. ", "- ", "* "
        line = re.sub(r"^(?:\d+[\.\)、]\s*|[-*•]\s*)", "", line).strip()

        parts: list[str] = []
        if "----" in line:
            # 兼容 token 自身末尾带 - 的 4 连横线分隔
            last = 0
            for match in re.finditer(r"-{4,}", line):
                parts.append((line[last : match.start()] + "-" * (len(match.group(0)) - 4)).strip())
                last = match.end()
            parts.append(line[last:].strip())
        elif "====" in line:
            parts = [p.strip() for p in line.split("====")]
        elif "|" in line:
            parts = [p.strip() for p in line.split("|")]
        elif "\t" in line:
            parts = [p.strip() for p in line.split("\t") if p.strip()]
        elif "," in line and line.count(",") >= 1:
            csv_parts = [p.strip().strip('"').strip("'") for p in line.split(",")]
            if csv_parts and _EMAIL_RE.match(csv_parts[0]):
                parts = csv_parts
        elif ":" in line and not line.lower().startswith(("http:", "https:", "mailto:")):
            colon_parts = [p.strip() for p in line.split(":")]
            if len(colon_parts) >= 2 and _EMAIL_RE.match(colon_parts[0]):
                parts = colon_parts

        if parts and _EMAIL_RE.match(parts[0]):
            email_addr = parts[0].lower()
            password = parts[1] if len(parts) >= 2 else ""
            client_id = parts[2] if len(parts) >= 3 else ""
            refresh_token = parts[3] if len(parts) >= 4 else ""
            recovery_email = ""
            recovery_code = ""
            if len(parts) >= 6:
                recovery_email = parts[4]
                recovery_code = parts[5]
            elif len(parts) == 5 and "@" in parts[4]:
                recovery_email = parts[4]

            prev = results.get(email_addr)
            # 如果之前只解析到了纯邮箱，而这次解析到了带 token/password 的完整行，则升级覆盖
            if not prev or (refresh_token and not prev.get("refresh_token")) or (password and not prev.get("password")):
                results[email_addr] = {
                    "email": parts[0].strip(),
                    "password": password,
                    "client_id": client_id,
                    "refresh_token": refresh_token,
                    "recovery_email": recovery_email,
                    "recovery_code": recovery_code,
                    "batch_no": detected_batch or "默认批次",
                    "source_info": default_source,
                }
            continue

        # 如果是单行纯邮箱或行内包含 hotmail/outlook/live/msn 邮箱
        if allow_email_only:
            if _EMAIL_RE.match(line):
                email_addr = line.lower()
                if email_addr not in results:
                    results[email_addr] = {
                        "email": line,
                        "password": "",
                        "client_id": "",
                        "refresh_token": "",
                        "recovery_email": "",
                        "recovery_code": "",
                        "batch_no": detected_batch or "默认批次",
                        "source_info": default_source,
                    }
            else:
                # 在混合文本行里查找 `email----password----client_id----refresh_token`
                for m in re.finditer(
                    r"([A-Za-z0-9._%+\-]+@(?:hotmail|outlook|live|msn)\.[A-Za-z.]+)\s*-{4,}\s*([^\s-]+)\s*-{4,}\s*([0-9a-fA-F\-]{32,36})\s*-{4,}\s*([A-Za-z0-9._!\-*$]+)",
                    line,
                ):
                    em_raw, pw, cid, rtok = m.groups()
                    results[em_raw.lower()] = {
                        "email": em_raw,
                        "password": pw,
                        "client_id": cid,
                        "refresh_token": rtok,
                        "recovery_email": "",
                        "recovery_code": "",
                        "batch_no": detected_batch or "默认批次",
                        "source_info": default_source,
                    }

    return list(results.values())


def upsert_mailboxes(records: list[dict[str, Any]]) -> dict[str, int]:
    """将提取到的邮箱写入或更新到 SQLite 台账数据库。"""
    inserted = 0
    updated = 0
    now = _now_iso()
    with _DB_LOCK:
        conn = get_db_conn()
        try:
            for item in records:
                email_raw = str(item.get("email") or "").strip()
                if not email_raw or "@" not in email_raw:
                    continue
                existing = conn.execute(
                    "SELECT * FROM mailboxes WHERE email = ? COLLATE NOCASE",
                    (email_raw,),
                ).fetchone()
                if not existing:
                    conn.execute(
                        """
                        INSERT INTO mailboxes (
                            email, password, client_id, refresh_token,
                            recovery_email, recovery_code, batch_no, source_info,
                            purpose, notes, audit_status, created_at, updated_at
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending', ?, ?)
                        """,
                        (
                            email_raw,
                            str(item.get("password") or "").strip(),
                            str(item.get("client_id") or "").strip(),
                            str(item.get("refresh_token") or "").strip(),
                            str(item.get("recovery_email") or "").strip(),
                            str(item.get("recovery_code") or "").strip(),
                            str(item.get("batch_no") or "默认批次").strip(),
                            str(item.get("source_info") or "导入").strip(),
                            str(item.get("purpose") or "").strip(),
                            str(item.get("notes") or "").strip(),
                            now,
                            now,
                        ),
                    )
                    inserted += 1
                else:
                    # 仅补充非空字段，不覆盖已有非空凭证
                    pw = str(item.get("password") or "").strip() or existing["password"]
                    cid = str(item.get("client_id") or "").strip() or existing["client_id"]
                    rtok = str(item.get("refresh_token") or "").strip() or existing["refresh_token"]
                    rec_em = str(item.get("recovery_email") or "").strip() or existing["recovery_email"]
                    rec_cd = str(item.get("recovery_code") or "").strip() or existing["recovery_code"]
                    batch_no = str(item.get("batch_no") or "").strip() or existing["batch_no"]
                    source_info = str(item.get("source_info") or "").strip() or existing["source_info"]
                    conn.execute(
                        """
                        UPDATE mailboxes
                        SET password = ?, client_id = ?, refresh_token = ?,
                            recovery_email = ?, recovery_code = ?,
                            batch_no = ?, source_info = ?, updated_at = ?
                        WHERE id = ?
                        """,
                        (pw, cid, rtok, rec_em, rec_cd, batch_no, source_info, now, existing["id"]),
                    )
                    updated += 1
            conn.commit()
        finally:
            conn.close()
    return {"inserted": inserted, "updated": updated}


# ============================================================
# 第一步：从指定收货邮箱获取购买的邮箱信息
# ============================================================

def _oauth_refresh_for_mailbox(email_addr: str, client_id: str, refresh_token: str, timeout: int = 12) -> tuple[str | None, str | None, str]:
    """
    尝试用 client_id + refresh_token 换取 access_token。
    返回 (channel_name, access_token, error_detail)
    channel_name:
      - 'live_imap': login.live.com IMAP XOAUTH2 token
      - 'ms_imap': login.microsoftonline.com IMAP XOAUTH2 token
      - 'ms_graph': login.microsoftonline.com Microsoft Graph token
    """
    if not client_id or not refresh_token:
        return None, None, "缺少 client_id 或 refresh_token"

    endpoints = [
        (
            "live_imap",
            "https://login.live.com/oauth20_token.srf",
            {
                "client_id": client_id,
                "grant_type": "refresh_token",
                "refresh_token": refresh_token,
            },
        ),
        (
            "ms_graph",
            "https://login.microsoftonline.com/common/oauth2/v2.0/token",
            {
                "client_id": client_id,
                "grant_type": "refresh_token",
                "refresh_token": refresh_token,
                "scope": "https://graph.microsoft.com/Mail.Read offline_access",
            },
        ),
        (
            "ms_imap",
            "https://login.microsoftonline.com/consumers/oauth2/v2.0/token",
            {
                "client_id": client_id,
                "grant_type": "refresh_token",
                "refresh_token": refresh_token,
                "scope": "https://outlook.office.com/IMAP.AccessAsUser.All offline_access",
            },
        ),
    ]

    errors: list[str] = []
    for channel, url, payload in endpoints:
        try:
            data_bytes = urllib.parse.urlencode(payload).encode("utf-8")
            req = urllib.request.Request(
                url,
                data=data_bytes,
                headers={
                    "Content-Type": "application/x-www-form-urlencoded",
                    "Accept": "application/json",
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
                },
            )
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                body = resp.read().decode("utf-8", errors="replace")
                parsed = json.loads(body)
                tok = str(parsed.get("access_token") or "").strip()
                if tok:
                    return channel, tok, ""
        except urllib.error.HTTPError as he:
            err_body = ""
            try:
                err_body = he.read().decode("utf-8", errors="replace")
                err_json = json.loads(err_body)
                err_code = err_json.get("error") or f"HTTP {he.code}"
                err_desc = str(err_json.get("error_description") or "")[:120]
                errors.append(f"{channel}:{err_code}({err_desc})")
            except Exception:
                errors.append(f"{channel}:HTTP {he.code}")
        except Exception as exc:
            errors.append(f"{channel}:{type(exc).__name__}({str(exc)[:80]})")

    return None, None, " | ".join(errors[:2])


def fetch_purchased_from_source_mailbox(cfg: dict[str, Any]) -> dict[str, Any]:
    """
    连接用户指定的收货邮箱（支持 IMAP 密码/授权码、Outlook OAuth2、POP3、Cloudflare 临时邮箱），
    扫描其中的发货邮件及附件，自动提取购买的 Hotmail/Outlook 邮箱信息并入库。
    """
    protocol = str(cfg.get("protocol") or "imap").strip().lower()
    email_addr = str(cfg.get("email") or "").strip()
    password = str(cfg.get("password") or "").strip()
    imap_server = str(cfg.get("imap_server") or "").strip()
    imap_port = int(cfg.get("imap_port") or 993)
    use_ssl = bool(cfg.get("use_ssl", True))
    client_id = str(cfg.get("client_id") or "").strip()
    refresh_token = str(cfg.get("refresh_token") or "").strip()
    search_keyword = str(cfg.get("search_keyword") or "").strip().lower()
    max_mails = max(1, min(200, int(cfg.get("max_mails") or 30)))
    batch_override = str(cfg.get("batch_no") or "").strip()

    # 如果用户在 email 字段直接粘贴了 4 段式卡密 email----pwd----cid----rtok，自动拆解
    if "----" in email_addr:
        parts = [p.strip() for p in email_addr.split("----")]
        if len(parts) >= 4:
            email_addr, password, client_id, refresh_token = parts[:4]
            protocol = "outlook_oauth"

    # 自动推断常见邮箱的 IMAP 服务器
    if not imap_server and "@" in email_addr:
        domain = email_addr.rsplit("@", 1)[-1].lower()
        if domain in COMMON_IMAP_SERVERS:
            imap_server, imap_port = COMMON_IMAP_SERVERS[domain]

    # 保存配置（脱敏或原样落本地库方便下次使用）
    with _DB_LOCK:
        conn = get_db_conn()
        try:
            conn.execute(
                """
                UPDATE source_mailbox_config
                SET protocol=?, email=?, password=?, imap_server=?, imap_port=?, use_ssl=?,
                    client_id=?, refresh_token=?, cf_api_base=?, cf_api_key=?, cf_address=?,
                    search_keyword=?, max_mails=?, updated_at=?
                WHERE id = 1
                """,
                (
                    protocol,
                    email_addr,
                    password,
                    imap_server,
                    imap_port,
                    1 if use_ssl else 0,
                    client_id,
                    refresh_token,
                    str(cfg.get("cf_api_base") or "https://apimail.jgxjs.com").strip(),
                    str(cfg.get("cf_api_key") or "").strip(),
                    str(cfg.get("cf_address") or "").strip(),
                    search_keyword,
                    max_mails,
                    _now_iso(),
                ),
            )
            conn.commit()
        finally:
            conn.close()

    scanned_mails: list[dict[str, Any]] = []
    extracted_accounts: list[dict[str, str]] = []

    if protocol in ("imap", "outlook_oauth"):
        mail: imaplib.IMAP4 | None = None
        try:
            if protocol == "outlook_oauth" or (client_id and refresh_token):
                ch, tok, err = _oauth_refresh_for_mailbox(email_addr, client_id, refresh_token)
                if not tok:
                    raise RuntimeError(f"指定邮箱 OAuth2 令牌刷新失败: {err}")
                if ch == "ms_graph":
                    # 用 Graph API 读取指定邮箱
                    graph_msgs = _fetch_messages_via_graph_api(tok, top=max_mails)
                    for gm in graph_msgs:
                        subj = str(gm.get("subject") or "")
                        sender = str(gm.get("from") or "")
                        dt_str = str(gm.get("date") or "")
                        body = str(gm.get("body") or "")
                        combined = f"{subj}\n{body}"
                        if search_keyword and search_keyword not in combined.lower() and search_keyword not in sender.lower():
                            continue
                        found = parse_purchased_accounts_text(
                            combined,
                            default_batch=batch_override,
                            default_source=f"指定邮箱({email_addr})",
                            allow_email_only=False,
                        )
                        scanned_mails.append({
                            "subject": subj,
                            "from": sender,
                            "date": dt_str,
                            "extracted_count": len(found),
                        })
                        extracted_accounts.extend(found)
                    upsert_res = upsert_mailboxes(extracted_accounts)
                    return {
                        "ok": True,
                        "source": f"OAuth2 Graph ({email_addr})",
                        "scanned_mail_count": len(scanned_mails),
                        "extracted_count": len(extracted_accounts),
                        "inserted": upsert_res["inserted"],
                        "updated": upsert_res["updated"],
                        "scanned_mails": scanned_mails[:20],
                    }
                mail = imaplib.IMAP4_SSL("outlook.office365.com", 993, timeout=18)
                auth_str = f"user={email_addr}\x01auth=Bearer {tok}\x01\x01".encode("utf-8")
                mail.authenticate("XOAUTH2", lambda _: auth_str)
            else:
                if not email_addr or not password or not imap_server:
                    raise ValueError("请填写完整的指定邮箱地址、密码/授权码以及 IMAP 服务器地址")
                cls = imaplib.IMAP4_SSL if use_ssl else imaplib.IMAP4
                mail = cls(imap_server, imap_port, timeout=18)
                mail.login(email_addr, password)

            status, cnt_data = mail.select("INBOX", readonly=True)
            if status != "OK":
                raise RuntimeError(f"无法打开指定邮箱 INBOX: {status}")
            total_msgs = int(cnt_data[0].decode()) if cnt_data and cnt_data[0] else 0
            if total_msgs > 0:
                start_seq = max(1, total_msgs - max_mails + 1)
                st, fetch_data = mail.fetch(f"{start_seq}:{total_msgs}", "(RFC822)")
                if st == "OK" and fetch_data:
                    for part in reversed(fetch_data):
                        if not (isinstance(part, tuple) and len(part) >= 2 and isinstance(part[1], bytes)):
                            continue
                        msg = email_lib.message_from_bytes(part[1])
                        subj = _decode_mime_header(msg.get("Subject"))
                        sender = _decode_mime_header(msg.get("From"))
                        dt_str = _format_date_short(msg.get("Date") or "")
                        body, attachments = _extract_email_body_and_attachments(msg)
                        full_text = subj + "\n" + body + "\n" + "\n".join(att[1] for att in attachments)
                        if search_keyword and search_keyword not in full_text.lower() and search_keyword not in sender.lower():
                            continue
                        # 优先提取带卡密分隔符的记录；如果邮件主题明确包含 hotmail/outlook/发货/卡密，也允许提取纯邮箱行
                        allow_plain = any(k in (subj + " " + search_keyword).lower() for k in ("hotmail", "outlook", "发货", "卡密", "邮箱", "订单"))
                        found = parse_purchased_accounts_text(
                            full_text,
                            default_batch=batch_override,
                            default_source=f"指定邮箱({email_addr})",
                            allow_email_only=allow_plain,
                        )
                        scanned_mails.append({
                            "subject": subj or "(无主题)",
                            "from": sender,
                            "date": dt_str,
                            "attachments": [a[0] for a in attachments],
                            "extracted_count": len(found),
                        })
                        extracted_accounts.extend(found)
        finally:
            if mail is not None:
                try:
                    mail.logout()
                except Exception:
                    pass

    elif protocol == "cf_worker":
        cf_base = str(cfg.get("cf_api_base") or "https://apimail.jgxjs.com").strip().rstrip("/")
        cf_key = str(cfg.get("cf_api_key") or "").strip()
        cf_addr = str(cfg.get("cf_address") or "").strip()
        params = {"limit": str(max_mails), "offset": "0"}
        if cf_addr:
            params["address"] = cf_addr
        url = f"{cf_base}/admin/mails?{urllib.parse.urlencode(params)}"
        headers = {"Accept": "application/json", "User-Agent": "Mozilla/5.0"}
        if cf_key:
            headers["x-admin-auth"] = cf_key
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=15) as resp:
            data = json.loads(resp.read().decode("utf-8", errors="replace"))
        rows = data.get("results") or data.get("mails") or data.get("data") or []
        for item in rows:
            raw_mime = str(item.get("raw") or item.get("text") or item.get("message") or "")
            subj = str(item.get("subject") or "")
            sender = str(item.get("source") or item.get("from") or "")
            dt_str = str(item.get("created_at") or "")
            if raw_mime and "Subject:" in raw_mime:
                msg = email_lib.message_from_bytes(raw_mime.encode("utf-8", errors="replace"))
                subj = subj or _decode_mime_header(msg.get("Subject"))
                sender = sender or _decode_mime_header(msg.get("From"))
                body, attachments = _extract_email_body_and_attachments(msg)
                full_text = subj + "\n" + body + "\n" + "\n".join(a[1] for a in attachments)
            else:
                full_text = f"{subj}\n{raw_mime}"
            if search_keyword and search_keyword not in full_text.lower():
                continue
            found = parse_purchased_accounts_text(
                full_text,
                default_batch=batch_override,
                default_source=f"CF邮箱({cf_addr or cf_base})",
                allow_email_only=False,
            )
            scanned_mails.append({
                "subject": subj or "(CF Mail)",
                "from": sender,
                "date": dt_str,
                "extracted_count": len(found),
            })
            extracted_accounts.extend(found)
    else:
        raise ValueError(f"不支持的收货邮箱协议: {protocol}")

    upsert_res = upsert_mailboxes(extracted_accounts)
    return {
        "ok": True,
        "source": email_addr or protocol,
        "scanned_mail_count": len(scanned_mails),
        "extracted_count": len(extracted_accounts),
        "inserted": upsert_res["inserted"],
        "updated": upsert_res["updated"],
        "scanned_mails": scanned_mails[:25],
    }


def sync_from_local_workspace_files(batch_no: str = "本地工作区邮箱池") -> dict[str, Any]:
    """一键从当前工作区已有的 用于注册的邮箱.txt / outlook-accounts.txt / turb.sqlite3 导入购买的邮箱。"""
    collected: list[dict[str, str]] = []
    sources_loaded: list[str] = []

    for path_obj, label in [
        (GPTCLI_OUTLOOK_TXT, "用于注册的邮箱.txt"),
        (GPTCLI_POOL_TXT, "outlook-accounts.txt"),
    ]:
        if path_obj.is_file():
            txt = path_obj.read_text(encoding="utf-8", errors="replace")
            recs = parse_purchased_accounts_text(txt, default_batch=batch_no, default_source=label, allow_email_only=False)
            if recs:
                collected.extend(recs)
                sources_loaded.append(f"{label}({len(recs)}条)")

    if GPTCLI_SQLITE_DB.is_file():
        try:
            conn = sqlite3.connect(str(GPTCLI_SQLITE_DB), timeout=10)
            conn.row_factory = sqlite3.Row
            rows = conn.execute("SELECT payload FROM email_pool WHERE source = 'outlook'").fetchall()
            db_count = 0
            for r in rows:
                try:
                    p = json.loads(r["payload"])
                    em = str(p.get("email") or "").strip()
                    if em:
                        collected.append({
                            "email": em,
                            "password": str(p.get("password") or ""),
                            "client_id": str(p.get("client_id") or ""),
                            "refresh_token": str(p.get("refresh_token") or ""),
                            "recovery_email": str(p.get("recovery_email") or ""),
                            "recovery_code": str(p.get("recovery_code") or ""),
                            "batch_no": batch_no,
                            "source_info": "turb.sqlite3(邮箱池)",
                        })
                        db_count += 1
                except Exception:
                    pass
            conn.close()
            if db_count:
                sources_loaded.append(f"turb.sqlite3({db_count}条)")
        except Exception:
            pass

    upsert_res = upsert_mailboxes(collected)
    return {
        "ok": True,
        "sources": sources_loaded,
        "extracted_count": len(collected),
        "inserted": upsert_res["inserted"],
        "updated": upsert_res["updated"],
    }


# ============================================================
# 第二步：对获取的邮箱进行深度取证审查与业务可用性分类
# ============================================================

def _fetch_messages_via_graph_api(access_token: str, top: int = 25) -> list[dict[str, str]]:
    """通过 Microsoft Graph API 拉取收件箱邮件列表。"""
    params = urllib.parse.urlencode({
        "$top": str(top),
        "$orderby": "receivedDateTime desc",
        "$select": "id,subject,from,toRecipients,receivedDateTime,bodyPreview,body",
    })
    url = f"https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?{params}"
    req = urllib.request.Request(
        url,
        headers={
            "Authorization": f"Bearer {access_token}",
            "Accept": "application/json",
            "Prefer": 'outlook.body-content-type="text"',
            "User-Agent": "Mozilla/5.0",
        },
    )
    with urllib.request.urlopen(req, timeout=15) as resp:
        data = json.loads(resp.read().decode("utf-8", errors="replace"))
    rows = data.get("value") or []
    messages: list[dict[str, str]] = []
    for m in rows:
        if not isinstance(m, dict):
            continue
        from_obj = ((m.get("from") or {}).get("emailAddress") or {}) if isinstance(m.get("from"), dict) else {}
        sender = f"{from_obj.get('name', '')} <{from_obj.get('address', '')}>".strip()
        to_list = m.get("toRecipients") or []
        to_addrs = []
        for t in to_list:
            if isinstance(t, dict):
                ea = t.get("emailAddress") or {}
                if ea.get("address"):
                    to_addrs.append(str(ea["address"]))
        body_obj = m.get("body") if isinstance(m.get("body"), dict) else {}
        body_text = str(body_obj.get("content") or m.get("bodyPreview") or "")
        messages.append({
            "date": _format_date_short(str(m.get("receivedDateTime") or "")),
            "from": sender,
            "to": ", ".join(to_addrs),
            "subject": str(m.get("subject") or "").strip(),
            "body": _clean_snippet(body_text, 400),
        })
    return messages


def _fetch_messages_via_imap(email_addr: str, auth_mode: str, secret: str, max_fetch: int = 30) -> tuple[int, list[dict[str, str]]]:
    """
    通过 IMAP (XOAUTH2 或密码登录) 读取 Outlook/Hotmail 收件箱历史邮件。
    一次性批量抓取邮件头与文本摘要，兼顾高速度与取证深度，并对微软瞬态断连自动重试。
    """
    last_exc: Exception | None = None
    for attempt in range(3):
        mail = imaplib.IMAP4_SSL("outlook.office365.com", 993, timeout=15)
        try:
            if auth_mode == "xoauth2":
                auth_str = f"user={email_addr}\x01auth=Bearer {secret}\x01\x01".encode("utf-8")
                mail.authenticate("XOAUTH2", lambda _: auth_str)
            else:
                mail.login(email_addr, secret)

            status, cnt_data = mail.select("INBOX", readonly=True)
            if status != "OK":
                raise RuntimeError(f"IMAP select INBOX 失败: {status}")
            total = int(cnt_data[0].decode()) if cnt_data and cnt_data[0] else 0
            if total <= 0:
                return 0, []

            start_seq = max(1, total - max_fetch + 1)
            # 一次性批量拉取邮件（包含头部和正文，最多30封，避免逐封往返耗时）
            st, fetch_data = mail.fetch(f"{start_seq}:{total}", "(RFC822)")
            if st != "OK" or not fetch_data:
                return total, []

            messages: list[dict[str, str]] = []
            for part in fetch_data:
                if not (isinstance(part, tuple) and len(part) >= 2 and isinstance(part[1], bytes)):
                    continue
                msg = email_lib.message_from_bytes(part[1])
                subj = _decode_mime_header(msg.get("Subject"))
                sender = _decode_mime_header(msg.get("From"))
                to_hdr = _decode_mime_header(msg.get("To"))
                delivered_to = _decode_mime_header(msg.get("Delivered-To"))
                orig_to = _decode_mime_header(msg.get("X-Original-To"))
                dt_str = _format_date_short(msg.get("Date") or "")
                body_text, _ = _extract_email_body_and_attachments(msg)
                combined_to = ", ".join(x for x in (to_hdr, delivered_to, orig_to) if x)
                messages.append({
                    "date": dt_str,
                    "from": sender,
                    "to": combined_to,
                    "subject": subj,
                    "body": _clean_snippet(body_text, 420),
                })
            return total, messages
        except Exception as exc:
            last_exc = exc
            if attempt < 2:
                time.sleep(1.0 * (attempt + 1))
        finally:
            try:
                mail.logout()
            except Exception:
                pass
    raise last_exc or RuntimeError("IMAP 连接失败")


def _check_masked_alias_mismatch(current_email: str, masked_account: str) -> bool:
    """
    检查 Microsoft 安全邮件中的掩码账号（如 'Da**y@hotmail.com' 或 'ox**1@outlook.com'）
    是否与当前邮箱（如 'DarylEffertzxdhby@hotmail.com'）不一致。
    如果不一致，说明该邮箱是改别名翻新（Alias-Recycled）的二手机箱！
    """
    cur = current_email.strip().lower()
    mask = masked_account.strip().lower().rstrip("。.,;)")
    if "@" not in cur or "@" not in mask or "*" not in mask:
        return False
    cur_local, cur_dom = cur.split("@", 1)
    mask_local, mask_dom = mask.split("@", 1)
    if cur_dom != mask_dom:
        return True
    prefix = mask_local.split("*", 1)[0]
    suffix = mask_local.rsplit("*", 1)[-1]
    if prefix and not cur_local.startswith(prefix):
        return True
    if suffix and not cur_local.endswith(suffix):
        return True
    return False


def analyze_mailbox_messages(
    email_addr: str,
    total_msg_count: int,
    messages: list[dict[str, str]],
    target_service: str = "openai",
) -> dict[str, Any]:
    """
    对单个邮箱的历史邮件进行深度取证分析，还原：
      - 注册过的服务列表
      - 风险标记（OpenAI 封号、ChatGPT 锁号、别名翻新、第三方应用连接等）
      - 前身别名（Alias-Recycled）
      - 能否用于注册目标业务
    """
    cur_email_low = email_addr.strip().lower()
    services_found: list[str] = []
    service_keys_found: set[str] = set()
    risk_flags: list[str] = []
    former_aliases: list[str] = []
    analyzed_samples: list[dict[str, Any]] = []

    has_openai_deactivated = False
    has_chatgpt_locked = False
    has_alias_recycled = False
    has_thunderbird_conn = False

    dates = [m["date"] for m in messages if m.get("date")]
    earliest_date = min(dates) if dates else ""
    latest_date = max(dates) if dates else ""

    for m in messages:
        subj = str(m.get("subject") or "").strip()
        subj_low = subj.lower()
        sender = str(m.get("from") or "").strip()
        sender_low = sender.lower()
        to_str = str(m.get("to") or "").strip()
        body = str(m.get("body") or "").strip()
        body_low = body.lower()
        combined_low = f"{subj_low}\n{sender_low}\n{body_low}"
        msg_tags: list[str] = []

        # 1. 检测收件人 To Header 是否包含与当前邮箱不同的 Microsoft 邮箱别名
        for addr_match in re.findall(r"[A-Za-z0-9._%+\-]+@(?:outlook|hotmail|live|msn)\.[A-Za-z.]+", to_str, flags=re.I):
            addr_low = addr_match.lower().rstrip(".")
            if addr_low != cur_email_low and not addr_low.startswith("account-security") and not addr_low.startswith("noreply"):
                if addr_low not in former_aliases:
                    former_aliases.append(addr_low)
                has_alias_recycled = True
                msg_tags.append(f"前身别名:{addr_low}")

        # 2. 检测 Microsoft 新应用连接通知（如 Thunderbird）与掩码别名翻新
        if (
            "连接到 microsoft 帐户的新应用" in subj_low
            or "new app(s) connected to your microsoft account" in subj_low
            or "connected to the microsoft account" in body_low
            or "已连接到 microsoft 帐户" in body_low
        ):
            has_thunderbird_conn = True
            msg_tags.append("Microsoft应用授权")
            if "第三方客户端授权 (Thunderbird)" not in risk_flags:
                risk_flags.append("第三方客户端授权 (Thunderbird)")
            # 提取掩码邮箱，如 Da**y@hotmail.com 或 ox**1@outlook.com
            for mask_m in re.findall(r"[A-Za-z0-9._%+\-]+\*+[A-Za-z0-9._%+\-]*@(?:outlook|hotmail|live|msn)\.[A-Za-z]+", body):
                if _check_masked_alias_mismatch(email_addr, mask_m):
                    has_alias_recycled = True
                    alias_label = f"{mask_m}(掩码不匹配)"
                    if alias_label not in former_aliases:
                        former_aliases.append(alias_label)
                    msg_tags.append(f"别名翻新:{mask_m}")

        if "添加邮箱登录名" in subj or "alias" in subj_low:
            has_alias_recycled = True
            msg_tags.append("修改登录别名")

        # 3. 检测 OpenAI 封号与 ChatGPT 锁号黑历史
        if "access deactivated" in subj_low or "account has been deactivated" in body_low:
            has_openai_deactivated = True
            msg_tags.append("🚨OpenAI封号通知")
            if "OpenAI 封号黑历史 (Access Deactivated)" not in risk_flags:
                risk_flags.append("OpenAI 封号黑历史 (Access Deactivated)")
            # 提取封号邮件正文中可能出现的前身别名邮箱
            for body_addr in re.findall(r"[A-Za-z0-9._%+\-]+@(?:outlook|hotmail|live|msn)\.[A-Za-z]+", body):
                b_low = body_addr.lower()
                if b_low != cur_email_low and b_low not in former_aliases:
                    former_aliases.append(b_low)
                    has_alias_recycled = True
                    msg_tags.append(f"封号前身别名:{b_low}")

        if (
            "action required to secure your chatgpt account" in subj_low
            or "your chatgpt account has been locked" in body_low
        ):
            has_chatgpt_locked = True
            msg_tags.append("🚨ChatGPT风控锁号")
            if "ChatGPT 风控锁号 (Account Locked)" not in risk_flags:
                risk_flags.append("ChatGPT 风控锁号 (Account Locked)")

        if "your x account has been suspended" in subj_low or "automatically suspended for violating the x rules" in body_low:
            msg_tags.append("⚠️X账号封禁")
            if "曾绑定的 X(Twitter) 账号被封禁" not in risk_flags:
                risk_flags.append("曾绑定的 X(Twitter) 账号被封禁")

        # 4. 匹配各类互联网/AI服务注册痕迹
        for svc_key, svc_label, domain_list, kw_list in SERVICE_RULES:
            matched = False
            if any(dom in sender_low for dom in domain_list):
                matched = True
            elif any(kw in combined_low for kw in kw_list):
                matched = True
            if matched:
                service_keys_found.add(svc_key)
                if svc_label not in services_found:
                    services_found.append(svc_label)
                if svc_label not in msg_tags:
                    msg_tags.append(svc_label)

        analyzed_samples.append({
            "date": m.get("date") or "",
            "from": sender,
            "to": to_str,
            "subject": subj,
            "snippet": body[:220],
            "tags": msg_tags,
        })

    if has_alias_recycled and "改别名翻新 (Alias-Recycled)" not in risk_flags:
        risk_flags.insert(0, "改别名翻新 (Alias-Recycled)")

    if "openai" in service_keys_found and not has_openai_deactivated and not has_chatgpt_locked:
        if "已存在 OpenAI/ChatGPT 账号记录" not in risk_flags:
            risk_flags.append("已存在 OpenAI/ChatGPT 账号记录")

    # 检查本地 output_cpa 目录是否已存在该邮箱的注册产物
    for cpa_dir in (WORKSPACE_DIR / "output_cpa", WORKSPACE_DIR.parent / "output_cpa", GPTCLI_ROOT / "data"):
        if cpa_dir.is_dir():
            found_local = False
            for prefix in ("chatgpt-", "codex-"):
                if (cpa_dir / f"{prefix}{email_addr}.json").is_file() or (cpa_dir / f"{prefix}{cur_email_low}.json").is_file():
                    service_keys_found.add("openai")
                    if "OpenAI / ChatGPT" not in services_found:
                        services_found.append("OpenAI / ChatGPT")
                    if "本地已存在注册凭证 (output_cpa)" not in risk_flags:
                        risk_flags.append("本地已存在注册凭证 (output_cpa)")
                    found_local = True
                    break
            if found_local:
                break

    # 综合分类判定
    target_key = (target_service or "openai").strip().lower()
    if has_openai_deactivated or has_chatgpt_locked or has_alias_recycled:
        audit_status = "blacklisted_recycled"
        can_register = 0
    elif target_key in service_keys_found:
        audit_status = "target_registered"
        can_register = 0
    elif services_found:
        # 仅注册过其他非目标服务（如 1040.com, Cursor, X, HK Ticketing），未注册过目标业务且无封号/改名黑历史
        audit_status = "other_services_only"
        can_register = 1 if target_key != "strict_clean" else 0
    else:
        audit_status = "clean_ready"
        can_register = 1

    return {
        "audit_status": audit_status,
        "can_register_target": can_register,
        "msg_count": total_msg_count,
        "earliest_msg_date": earliest_date,
        "latest_msg_date": latest_date,
        "registered_services": services_found,
        "risk_flags": risk_flags,
        "former_aliases": former_aliases,
        "messages_sample": analyzed_samples[::-1][:25],  # 最新邮件排前面，最多存25封完整取证样本
    }


def audit_single_mailbox(row: dict[str, Any], target_service: str = "openai") -> dict[str, Any]:
    """对单条邮箱记录执行实时连接与收件箱历史审查，并将结果写回数据库。"""
    email_addr = str(row.get("email") or "").strip()
    password = str(row.get("password") or "").strip()
    client_id = str(row.get("client_id") or "").strip()
    refresh_token = str(row.get("refresh_token") or "").strip()

    auth_channel = ""
    auth_error = ""
    total_msgs = 0
    messages: list[dict[str, str]] = []
    connected = False

    # 1. 优先尝试 OAuth2 (Live IMAP / MS Graph / MS IMAP)
    if client_id and refresh_token:
        ch, tok, err = _oauth_refresh_for_mailbox(email_addr, client_id, refresh_token)
        if tok:
            if ch in ("live_imap", "ms_imap"):
                try:
                    total_msgs, messages = _fetch_messages_via_imap(email_addr, "xoauth2", tok)
                    auth_channel = ch
                    connected = True
                except Exception as imap_exc:
                    auth_error = f"IMAP连接异常: {imap_exc}"
                    # 如果 IMAP 报 User is authenticated but not connected，尝试回退 Graph API
                    try:
                        g_ch, g_tok, _ = _oauth_refresh_for_mailbox(
                            email_addr,
                            client_id,
                            refresh_token,
                        )
                        if g_ch == "ms_graph" and g_tok:
                            messages = _fetch_messages_via_graph_api(g_tok)
                            total_msgs = len(messages)
                            auth_channel = "ms_graph"
                            auth_error = ""
                            connected = True
                    except Exception:
                        pass
            elif ch == "ms_graph":
                try:
                    messages = _fetch_messages_via_graph_api(tok)
                    total_msgs = len(messages)
                    auth_channel = "ms_graph"
                    connected = True
                except Exception as g_exc:
                    auth_error = f"Graph读取异常: {g_exc}"
        else:
            auth_error = err or "refresh_token 失效"

    # 2. 若 OAuth2 未连接成功且有密码，尝试 IMAP 密码登录兜底
    if not connected and password and not (client_id and refresh_token):
        try:
            total_msgs, messages = _fetch_messages_via_imap(email_addr, "password", password)
            auth_channel = "imap_password"
            auth_error = ""
            connected = True
        except Exception as pw_exc:
            auth_error = f"密码登录失败: {pw_exc}"

    now = _now_iso()
    if not connected:
        result = {
            "audit_status": "auth_failed",
            "can_register_target": 0,
            "auth_channel": "failed",
            "auth_error": auth_error or "缺少有效的 OAuth2 refresh_token 或邮箱未开通 IMAP",
            "msg_count": 0,
            "earliest_msg_date": "",
            "latest_msg_date": "",
            "registered_services": [],
            "risk_flags": [f"连接失败: {(auth_error or '凭证无效')[:90]}"],
            "former_aliases": [],
            "messages_sample": [],
            "last_audited_at": now,
        }
    else:
        analysis = analyze_mailbox_messages(email_addr, total_msgs, messages, target_service=target_service)
        result = {
            **analysis,
            "auth_channel": auth_channel,
            "auth_error": "",
            "last_audited_at": now,
        }

    with _DB_LOCK:
        conn = get_db_conn()
        try:
            conn.execute(
                """
                UPDATE mailboxes
                SET audit_status = ?,
                    can_register_target = ?,
                    auth_channel = ?,
                    auth_error = ?,
                    msg_count = ?,
                    earliest_msg_date = ?,
                    latest_msg_date = ?,
                    registered_services = ?,
                    risk_flags = ?,
                    former_aliases = ?,
                    messages_sample = ?,
                    last_audited_at = ?,
                    updated_at = ?
                WHERE email = ? COLLATE NOCASE
                """,
                (
                    result["audit_status"],
                    result["can_register_target"],
                    result["auth_channel"],
                    result["auth_error"],
                    result["msg_count"],
                    result["earliest_msg_date"],
                    result["latest_msg_date"],
                    json.dumps(result["registered_services"], ensure_ascii=False),
                    json.dumps(result["risk_flags"], ensure_ascii=False),
                    json.dumps(result["former_aliases"], ensure_ascii=False),
                    json.dumps(result["messages_sample"], ensure_ascii=False),
                    result["last_audited_at"],
                    now,
                    email_addr,
                ),
            )
            conn.commit()
        finally:
            conn.close()

    return {"email": email_addr, **result}


def run_batch_audit(
    scope: str = "all",
    emails: list[str] | None = None,
    batch_no: str = "",
    target_service: str = "openai",
    max_workers: int = 8,
    blocking: bool = True,
) -> dict[str, Any]:
    """批量并发审查邮箱。"""
    with _DB_LOCK:
        conn = get_db_conn()
        try:
            if emails:
                placeholders = ",".join("?" for _ in emails)
                rows = [
                    dict(r)
                    for r in conn.execute(
                        f"SELECT * FROM mailboxes WHERE email IN ({placeholders})",
                        emails,
                    ).fetchall()
                ]
            elif scope == "pending":
                rows = [
                    dict(r)
                    for r in conn.execute("SELECT * FROM mailboxes WHERE audit_status = 'pending'").fetchall()
                ]
            elif batch_no:
                rows = [
                    dict(r)
                    for r in conn.execute("SELECT * FROM mailboxes WHERE batch_no = ?", (batch_no,)).fetchall()
                ]
            else:
                rows = [dict(r) for r in conn.execute("SELECT * FROM mailboxes ORDER BY id ASC").fetchall()]
        finally:
            conn.close()

    total = len(rows)
    if total == 0:
        return {"ok": True, "total": 0, "completed": 0, "summary": get_dashboard_summary()}

    def _worker() -> dict[str, Any]:
        with _AUDIT_STATE_LOCK:
            _AUDIT_PROGRESS.update({
                "running": True,
                "total": total,
                "completed": 0,
                "current_email": "",
                "started_at": time.time(),
                "finished_at": 0.0,
                "target_service": target_service,
            })
        counts: dict[str, int] = {}
        try:
            with ThreadPoolExecutor(max_workers=max(1, min(16, max_workers))) as ex:
                future_map = {ex.submit(audit_single_mailbox, r, target_service): r["email"] for r in rows}
                for fut in as_completed(future_map):
                    em = future_map[fut]
                    try:
                        res = fut.result()
                        st = res.get("audit_status", "auth_failed")
                        counts[st] = counts.get(st, 0) + 1
                    except Exception:
                        counts["auth_failed"] = counts.get("auth_failed", 0) + 1
                    with _AUDIT_STATE_LOCK:
                        _AUDIT_PROGRESS["completed"] += 1
                        _AUDIT_PROGRESS["current_email"] = em
        finally:
            summary = get_dashboard_summary()
            with _AUDIT_STATE_LOCK:
                _AUDIT_PROGRESS["running"] = False
                _AUDIT_PROGRESS["finished_at"] = time.time()
                _AUDIT_PROGRESS["summary"] = summary
        return {"ok": True, "total": total, "completed": total, "counts": counts, "summary": summary}

    if blocking:
        return _worker()
    t = threading.Thread(target=_worker, daemon=True)
    t.start()
    return {"ok": True, "started": True, "total": total}


# ============================================================
# 统计、取证报告生成与注册机邮箱池同步
# ============================================================

def _deserialize_row(row: sqlite3.Row | dict[str, Any]) -> dict[str, Any]:
    d = dict(row)
    for k in ("registered_services", "risk_flags", "former_aliases", "messages_sample"):
        raw = d.get(k)
        if isinstance(raw, str):
            try:
                d[k] = json.loads(raw)
            except Exception:
                d[k] = []
        elif not isinstance(raw, list):
            d[k] = []
    d["can_register_target"] = bool(d.get("can_register_target"))
    d["has_oauth"] = bool(d.get("client_id") and d.get("refresh_token"))

    # 计算三大核心业务可用性分类：一、可以注册Codex；二、可以注册Grok；三、可以注册Claude
    cur_email_low = str(d.get("email") or "").strip().lower()
    st = str(d.get("audit_status") or "pending")
    auth_ch = str(d.get("auth_channel") or "")
    is_alive = st not in ("auth_failed", "pending") and auth_ch not in ("failed", "")
    risk_flags: list[str] = d.get("risk_flags") or []
    former_aliases: list[str] = d.get("former_aliases") or []
    reg_svcs: list[str] = d.get("registered_services") or []
    samples: list[dict[str, Any]] = d.get("messages_sample") or []

    is_shared_alias = bool(former_aliases) or ("改别名翻新 (Alias-Recycled)" in risk_flags)

    own_services: list[str] = []
    alias_services: list[str] = []
    own_openai_banned = False

    if not is_shared_alias:
        own_services = list(reg_svcs)
        if any("Access Deactivated" in rf or "Account Locked" in rf for rf in risk_flags):
            own_openai_banned = True
    else:
        for m in samples:
            to_str = str(m.get("to") or "")
            rcpts = [
                a.lower().rstrip(".")
                for a in re.findall(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}", to_str)
            ]
            is_own_msg = (not rcpts) or (cur_email_low in rcpts)
            m_tags = m.get("tags") or []
            subj_low = str(m.get("subject") or "").lower()
            snip_low = str(m.get("snippet") or "").lower()
            if is_own_msg and (
                "access deactivated" in subj_low
                or "account has been deactivated" in snip_low
                or "action required to secure your chatgpt account" in subj_low
            ):
                own_openai_banned = True
            for svc_key, svc_label, domain_list, kw_list in SERVICE_RULES:
                comb = f"{subj_low}\n{str(m.get('from') or '').lower()}\n{snip_low}"
                if (
                    svc_label in m_tags
                    or any(dom in str(m.get("from") or "").lower() for dom in domain_list)
                    or any(kw in comb for kw in kw_list)
                ):
                    if is_own_msg:
                        if svc_label not in own_services:
                            own_services.append(svc_label)
                    else:
                        if svc_label not in alias_services:
                            alias_services.append(svc_label)

    if "本地已存在注册凭证 (output_cpa)" in risk_flags and "OpenAI / ChatGPT" not in own_services:
        own_services.append("OpenAI / ChatGPT")

    def _calc_biz(svc_label: str, extra_ban: bool = False) -> tuple[bool, str]:
        if not is_alive:
            return False, "dead"
        if svc_label in own_services or extra_ban:
            return False, "registered"
        if is_shared_alias:
            return True, "alias_ok"
        return True, "clean_ok"

    can_codex, codex_status = _calc_biz("OpenAI / ChatGPT", extra_ban=own_openai_banned)
    can_grok, grok_status = _calc_biz("Grok / xAI")
    can_claude, claude_status = _calc_biz("Claude / Anthropic")

    d["is_shared_alias"] = is_shared_alias
    d["own_services"] = own_services
    d["alias_services"] = alias_services
    d["can_codex"] = can_codex
    d["codex_status"] = codex_status  # clean_ok | alias_ok | registered | dead
    d["can_grok"] = can_grok
    d["grok_status"] = grok_status    # clean_ok | alias_ok | registered | dead
    d["can_claude"] = can_claude
    d["claude_status"] = claude_status  # clean_ok | alias_ok | registered | dead
    return d


def list_mailboxes() -> list[dict[str, Any]]:
    with _DB_LOCK:
        conn = get_db_conn()
        try:
            rows = conn.execute("SELECT * FROM mailboxes ORDER BY id DESC").fetchall()
            return [_deserialize_row(r) for r in rows]
        finally:
            conn.close()


def get_dashboard_summary() -> dict[str, Any]:
    items = list_mailboxes()
    summary: dict[str, Any] = {
        "total": len(items),
        "clean_ready": 0,
        "other_services_only": 0,
        "target_registered": 0,
        "blacklisted_recycled": 0,
        "auth_failed": 0,
        "pending": 0,
        "can_register_count": 0,
        "codex_total": 0,
        "codex_clean": 0,
        "codex_alias": 0,
        "grok_total": 0,
        "grok_clean": 0,
        "grok_alias": 0,
        "claude_total": 0,
        "claude_clean": 0,
        "claude_alias": 0,
        "batches": [],
        "service_breakdown": {},
    }
    batches_set: set[str] = set()
    svc_counts: dict[str, int] = {}
    for item in items:
        st = item.get("audit_status") or "pending"
        if st in summary:
            summary[st] += 1
        else:
            summary["pending"] += 1
        if item.get("can_register_target"):
            summary["can_register_count"] += 1

        if item.get("can_codex"):
            summary["codex_total"] += 1
            if item.get("codex_status") == "clean_ok":
                summary["codex_clean"] += 1
            else:
                summary["codex_alias"] += 1

        if item.get("can_grok"):
            summary["grok_total"] += 1
            if item.get("grok_status") == "clean_ok":
                summary["grok_clean"] += 1
            else:
                summary["grok_alias"] += 1

        if item.get("can_claude"):
            summary["claude_total"] += 1
            if item.get("claude_status") == "clean_ok":
                summary["claude_clean"] += 1
            else:
                summary["claude_alias"] += 1

        if item.get("batch_no"):
            batches_set.add(item["batch_no"])
        for svc in item.get("registered_services") or []:
            svc_counts[svc] = svc_counts.get(svc, 0) + 1
    summary["batches"] = sorted(batches_set)
    summary["service_breakdown"] = svc_counts
    return summary


def generate_forensic_report(batch_no: str = "") -> str:
    """
    自动生成《号商邮箱审查与黑历史取证报告》（Markdown 格式），
    还原早期注册时间线、客户端连接记录、第三方服务验证码、OpenAI 封号/锁号通知及前身别名。
    """
    items = list_mailboxes()
    if batch_no:
        items = [x for x in items if x.get("batch_no") == batch_no]

    total = len(items)
    clean_cnt = sum(1 for x in items if x["audit_status"] == "clean_ready")
    other_cnt = sum(1 for x in items if x["audit_status"] == "other_services_only")
    reg_cnt = sum(1 for x in items if x["audit_status"] == "target_registered")
    black_cnt = sum(1 for x in items if x["audit_status"] == "blacklisted_recycled")
    dead_cnt = sum(1 for x in items if x["audit_status"] == "auth_failed")
    can_use_cnt = sum(1 for x in items if x["can_register_target"])

    batch_title = batch_no or "全部批次"
    lines = [
        f"# 邮箱资产深度审查与取证报告（{batch_title}）",
        "",
        f"- **生成时间**：`{_now_iso()}`",
        f"- **审查总数**：**{total}** 个邮箱",
        f"- **可用于注册目标业务（OpenAI/ChatGPT）**：**{can_use_cnt}** 个（其中全新纯净 **{clean_cnt}** 个，仅注册过其他非冲突服务 **{other_cnt}** 个）",
        f"- **不可用于新注册**：**{total - can_use_cnt}** 个",
        f"  - 🔴 **封号黑历史 / 别名翻新二手机箱**：**{black_cnt}** 个",
        f"  - 🟠 **已注册过 OpenAI / ChatGPT**：**{reg_cnt}** 个",
        f"  - ⚫ **OAuth2 Token 失效 / 邮箱无法连接**：**{dead_cnt}** 个",
        "",
        "---",
        "",
        "## 一、问题邮箱取证汇总（可用于售后退换维权）",
        "",
    ]

    # 1. 封号黑历史 / 别名翻新 / 已注册过服务的二手机箱取证
    dirty_items = [
        x for x in items if x["audit_status"] in ("blacklisted_recycled", "target_registered", "other_services_only")
    ]
    if dirty_items:
        lines.append(f"### 1. 历史邮件与二手翻新轨迹还原（共 {len(dirty_items)} 个存量历史邮件邮箱）")
        lines.append("")
        for idx, item in enumerate(dirty_items, 1):
            st_label = {
                "blacklisted_recycled": "🔴 封号黑历史 / 别名翻新",
                "target_registered": "🟠 已注册过 OpenAI/ChatGPT",
                "other_services_only": "🟡 曾被用于注册其他第三方服务",
            }.get(item["audit_status"], item["audit_status"])
            svcs = "、".join(item.get("registered_services") or []) or "无明显服务标签"
            flags = "、".join(item.get("risk_flags") or []) or "无"
            aliases = "、".join(item.get("former_aliases") or [])
            lines.append(
                f"#### ({idx}) `{item['email']}` — **{st_label}**（收件箱共 **{item['msg_count']}** 封历史邮件）"
            )
            lines.append(f"- **时间跨度**：`{item.get('earliest_msg_date') or '未知'}` ~ `{item.get('latest_msg_date') or '未知'}`")
            lines.append(f"- **检出已注册服务**：{svcs}")
            lines.append(f"- **风险标签**：{flags}")
            if aliases:
                lines.append(f"- **检出前身别名 (Alias-Recycled)**：`{aliases}`")
            samples = item.get("messages_sample") or []
            if samples:
                lines.append("- **IMAP 原始历史邮件证据摘录**：")
                for s in samples[:8]:
                    lines.append(
                        f"  - `[{s.get('date') or '未知时间'}]` **主题**: `{s.get('subject')}` | **发件人**: `{s.get('from')}`"
                        + (f" | **摘要**: _{s.get('snippet')[:110]}_" if s.get("snippet") else "")
                    )
            lines.append("")
    else:
        lines.append("### 1. 历史邮件与二手翻新轨迹还原\n- 暂未检出含历史邮件的二手邮箱。\n")

    # 2. 凭证失效死号列表
    dead_items = [x for x in items if x["audit_status"] == "auth_failed"]
    if dead_items:
        lines.append(f"### 2. OAuth2 令牌失效 / 邮箱未开通死号清单（共 {len(dead_items)} 个）")
        lines.append("")
        for item in dead_items:
            lines.append(f"- `{item['email']}` — 原因：`{item.get('auth_error') or 'refresh_token 无效'}`")
        lines.append("")

    # 3. 可用于注册目标业务的邮箱清单
    usable_items = [x for x in items if x["can_register_target"]]
    lines.append(f"## 二、审查通过 · 可用于注册目标业务清单（共 {len(usable_items)} 个）")
    lines.append("")
    if usable_items:
        for item in usable_items:
            svcs = "、".join(item.get("registered_services") or [])
            note = f"（全新纯净，0 历史冲突）" if item["audit_status"] == "clean_ready" else f"（未注册过 OpenAI，仅有历史邮件：{svcs}）"
            lines.append(f"- `{item['email']}` {note}")
    else:
        lines.append("- 当前无审查通过的可用邮箱。")

    return "\n".join(lines)


def sync_usable_to_gptcli_pool(include_other_services: bool = True) -> dict[str, Any]:
    """
    将面板中审查通过（可注册 OpenAI）且含完整 OAuth 凭证的邮箱同步写入：
      1. gptcli-2api/grok-build-auth/openai_client/用于注册的邮箱.txt
      2. gptcli-2api/output/mailboxes/outlook-accounts.txt
      3. gptcli-2api/grok-build-auth/openai_client/turb.sqlite3 (email_pool)
    同时将已确认失效或已注册过 OpenAI 的邮箱在 turb.sqlite3 中标记为 failed，避免注册机踩坑！
    """
    items = list_mailboxes()
    allowed_statuses = {"clean_ready", "other_services_only"} if include_other_services else {"clean_ready"}

    usable_lines: list[str] = []
    usable_items: list[dict[str, Any]] = []
    bad_items: list[dict[str, Any]] = []

    for item in items:
        if not (item.get("client_id") and item.get("refresh_token")):
            continue
        if item.get("audit_status") in allowed_statuses and item.get("can_register_target"):
            usable_items.append(item)
            usable_lines.append(
                f"{item['email']}----{item['password']}----{item['client_id']}----{item['refresh_token']}"
            )
        elif item.get("audit_status") in ("auth_failed", "blacklisted_recycled", "target_registered"):
            bad_items.append(item)

    # 写入 txt 文件
    for txt_path in (GPTCLI_OUTLOOK_TXT, GPTCLI_POOL_TXT):
        try:
            txt_path.parent.mkdir(parents=True, exist_ok=True)
            txt_path.write_text("\n".join(usable_lines) + ("\n" if usable_lines else ""), encoding="utf-8")
        except Exception:
            pass

    # 同步更新 turb.sqlite3 的 email_pool 表
    db_synced_usable = 0
    db_marked_failed = 0
    if GPTCLI_SQLITE_DB.is_file():
        try:
            sys.path.insert(0, str(GPTCLI_ROOT / "grok-build-auth"))
            sys.path.insert(0, str(GPTCLI_ROOT / "grok-build-auth" / "openai_client"))
            from core.db import import_outlook_accounts, release_outlook  # type: ignore

            if usable_items:
                ins, _skip = import_outlook_accounts([
                    {
                        "email": x["email"],
                        "password": x["password"],
                        "client_id": x["client_id"],
                        "refresh_token": x["refresh_token"],
                    }
                    for x in usable_items
                ])
                db_synced_usable = len(usable_items)
                for x in usable_items:
                    release_outlook(x["email"], status="available", note=f"audited_ok:{x['audit_status']}")

            for x in bad_items:
                reason = f"audit_{x['audit_status']}: " + (", ".join(x.get("risk_flags") or []) or x.get("auth_error") or "")
                release_outlook(x["email"], status="failed", note=reason[:180])
                db_marked_failed += 1
        except Exception:
            pass

    return {
        "ok": True,
        "usable_synced_count": len(usable_items),
        "bad_marked_failed_count": db_marked_failed,
        "files_updated": [str(GPTCLI_OUTLOOK_TXT), str(GPTCLI_POOL_TXT)],
    }


# ============================================================
# HTTP API 处理器
# ============================================================

class DashboardRequestHandler(BaseHTTPRequestHandler):
    def log_message(self, fmt: str, *args: Any) -> None:
        return

    def _send_json(self, data: Any, status: int = 200) -> None:
        raw = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()
        self.wfile.write(raw)

    def _read_json_body(self) -> dict[str, Any]:
        length = int(self.headers.get("Content-Length") or 0)
        if length <= 0:
            return {}
        raw = self.rfile.read(length).decode("utf-8", errors="replace")
        try:
            return json.loads(raw)
        except Exception:
            return {}

    def do_OPTIONS(self) -> None:
        self._send_json({"ok": True})

    def do_GET(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        if path.startswith("/mail-sieve"):
            path = path[len("/mail-sieve"):] or "/"

        if path in ("/", "/index.html", "/hotmail-dashboard.html"):
            if DASHBOARD_HTML_PATH.is_file():
                content = DASHBOARD_HTML_PATH.read_bytes()
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(content)))
                self.end_headers()
                self.wfile.write(content)
                return
            self._send_json({"error": "hotmail-dashboard.html not found"}, 404)
            return

        if path == "/health":
            self._send_json({"ok": True, "service": "mail-sieve"})
            return

        if path == "/api/state":
            with _DB_LOCK:
                conn = get_db_conn()
                try:
                    cfg_row = conn.execute("SELECT * FROM source_mailbox_config WHERE id = 1").fetchone()
                    source_cfg = dict(cfg_row) if cfg_row else {}
                finally:
                    conn.close()
            with _AUDIT_STATE_LOCK:
                audit_prog = dict(_AUDIT_PROGRESS)
            self._send_json({
                "ok": True,
                "summary": get_dashboard_summary(),
                "mailboxes": list_mailboxes(),
                "source_config": source_cfg,
                "audit_progress": audit_prog,
            })
            return

        if path == "/api/audit-status":
            with _AUDIT_STATE_LOCK:
                prog = dict(_AUDIT_PROGRESS)
            self._send_json({"ok": True, "audit_progress": prog, "summary": get_dashboard_summary()})
            return

        if path == "/api/forensic-report":
            qs = urllib.parse.parse_qs(parsed.query)
            batch_no = (qs.get("batch_no") or [""])[0]
            report_md = generate_forensic_report(batch_no=batch_no)
            self._send_json({"ok": True, "report": report_md})
            return

        self._send_json({"error": f"Not found: {path}"}, 404)

    def do_POST(self) -> None:
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        if path.startswith("/mail-sieve"):
            path = path[len("/mail-sieve"):] or "/"
        body = self._read_json_body()

        try:
            if path == "/api/fetch-from-source":
                res = fetch_purchased_from_source_mailbox(body)
                if body.get("auto_audit") and res.get("extracted_count", 0) > 0:
                    run_batch_audit(scope="pending", target_service=str(body.get("target_service") or "openai"), blocking=False)
                self._send_json(res)
                return

            if path == "/api/sync-workspace":
                batch_no = str(body.get("batch_no") or "工作区邮箱池").strip()
                res = sync_from_local_workspace_files(batch_no=batch_no)
                if body.get("auto_audit"):
                    run_batch_audit(scope="pending", target_service=str(body.get("target_service") or "openai"), blocking=False)
                self._send_json(res)
                return

            if path == "/api/import-text":
                raw_text = str(body.get("text") or "")
                batch_no = str(body.get("batch_no") or "").strip()
                source_info = str(body.get("source_info") or "批量粘贴导入").strip()
                recs = parse_purchased_accounts_text(
                    raw_text,
                    default_batch=batch_no,
                    default_source=source_info,
                    allow_email_only=True,
                )
                up = upsert_mailboxes(recs)
                if body.get("auto_audit") and recs:
                    run_batch_audit(
                        emails=[r["email"] for r in recs],
                        target_service=str(body.get("target_service") or "openai"),
                        blocking=False,
                    )
                self._send_json({
                    "ok": True,
                    "extracted_count": len(recs),
                    "inserted": up["inserted"],
                    "updated": up["updated"],
                })
                return

            if path == "/api/audit":
                scope = str(body.get("scope") or "all").strip()
                emails = body.get("emails") if isinstance(body.get("emails"), list) else None
                batch_no = str(body.get("batch_no") or "").strip()
                target_service = str(body.get("target_service") or "openai").strip()
                blocking = bool(body.get("blocking", False))
                res = run_batch_audit(
                    scope=scope,
                    emails=emails,
                    batch_no=batch_no,
                    target_service=target_service,
                    blocking=blocking,
                )
                self._send_json(res)
                return

            if path == "/api/update-mailbox":
                email_addr = str(body.get("email") or "").strip()
                if not email_addr:
                    self._send_json({"ok": False, "error": "缺少邮箱地址"}, 400)
                    return
                with _DB_LOCK:
                    conn = get_db_conn()
                    try:
                        conn.execute(
                            """
                            UPDATE mailboxes
                            SET purpose = ?, batch_no = ?, notes = ?, audit_status = ?,
                                can_register_target = ?, updated_at = ?
                            WHERE email = ? COLLATE NOCASE
                            """,
                            (
                                str(body.get("purpose") or "").strip(),
                                str(body.get("batch_no") or "").strip(),
                                str(body.get("notes") or "").strip(),
                                str(body.get("audit_status") or "pending").strip(),
                                1 if body.get("can_register_target") else 0,
                                _now_iso(),
                                email_addr,
                            ),
                        )
                        conn.commit()
                    finally:
                        conn.close()
                self._send_json({"ok": True})
                return

            if path == "/api/delete-mailboxes":
                emails = body.get("emails") or []
                if isinstance(emails, list) and emails:
                    with _DB_LOCK:
                        conn = get_db_conn()
                        try:
                            placeholders = ",".join("?" for _ in emails)
                            conn.execute(f"DELETE FROM mailboxes WHERE email IN ({placeholders})", emails)
                            conn.commit()
                        finally:
                            conn.close()
                self._send_json({"ok": True, "deleted": len(emails)})
                return

            if path == "/api/sync-to-gptcli":
                inc_other = bool(body.get("include_other_services", True))
                res = sync_usable_to_gptcli_pool(include_other_services=inc_other)
                self._send_json(res)
                return

            self._send_json({"ok": False, "error": f"Unknown endpoint: {path}"}, 404)
        except Exception as exc:
            self._send_json({"ok": False, "error": str(exc), "trace": traceback.format_exc()[-400:]}, 500)


def main() -> int:
    parser = argparse.ArgumentParser(description="Hotmail 批量邮箱获取与审查管理台服务")
    parser.add_argument("--host", default="0.0.0.0")
    parser.add_argument("--port", type=int, default=8899)
    parser.add_argument("--auto-init-audit", action="store_true", help="启动时自动导入工作区邮箱并执行首次全量审查")
    args = parser.parse_args()

    init_db()
    if args.auto_init_audit:
        sync_from_local_workspace_files()
        run_batch_audit(scope="pending", blocking=True)

    server = ThreadingHTTPServer((args.host, args.port), DashboardRequestHandler)
    print(f"[*] Hotmail 邮箱获取与审查管理面板已启动: http://127.0.0.1:{args.port}", flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
