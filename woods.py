# =============================================================================
# kolip.py  —  Product Process Tracker  (fully corrected, all features intact)
# Python 3.12 compatible · openpyxl guard added · all syntax issues resolved
# =============================================================================

# ─────────────────────────────────────────────────────────────────────────────
# STANDARD IMPORTS
# ─────────────────────────────────────────────────────────────────────────────
import sys
import subprocess
import importlib

def _ensure_package(pkg_import: str, pip_name: str) -> None:
    try:
        importlib.import_module(pkg_import)
    except ModuleNotFoundError:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", pip_name, "--quiet"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )

_ensure_package("streamlit",    "streamlit")
_ensure_package("pandas",       "pandas")
_ensure_package("numpy",        "numpy")
_ensure_package("openpyxl",     "openpyxl")
_ensure_package("plotly",       "plotly")

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import date, datetime, timedelta
import os
import sqlite3
import hashlib
import time
import smtplib
import random
import string
import uuid
import json
import re
from email.mime.text import MIMEText
from urllib.parse import urlparse, parse_qs

from openpyxl import Workbook
import plotly.express as px
import plotly.graph_objects as go

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Product Process Tracker",
    page_icon="💎",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────────────────────────────────────
# GLOBAL CSS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=Syne:wght@600;700;800&display=swap');
:root{
  --bg:#0d0f14;--surface:#13161e;--card:#1a1d27;--border:#262a38;
  --accent:#7c6aff;--accent2:#2be0c8;--danger:#ff5c6a;
  --text:#e4e6f0;--muted:#6b7094;--success:#3ddc84;--gold:#f5c842;
  --orange:#F4A261;--tariff:#e879f9;
}
html,body,[data-testid="stAppViewContainer"]{
  background:var(--bg)!important;color:var(--text)!important;
  font-family:'DM Mono',monospace!important;
}
[data-testid="stHeader"]{background:var(--bg)!important;}
[data-testid="stToolbar"]{display:none;}
.block-container{padding:1.2rem 2rem 3rem!important;max-width:1600px;}
[data-testid="stTabs"] button{
  font-family:'Syne',sans-serif!important;font-weight:700!important;
  font-size:.85rem!important;color:var(--muted)!important;
  border-bottom:2px solid transparent!important;padding:.5rem .8rem!important;transition:all .2s;
}
[data-testid="stTabs"] button[aria-selected="true"]{
  color:var(--accent)!important;border-bottom:2px solid var(--accent)!important;
}
[data-testid="stTabsContent"]{padding-top:1rem!important;}
[data-testid="stFileUploader"]{
  background:var(--card)!important;border:1.5px dashed var(--border)!important;
  border-radius:10px!important;padding:.8rem!important;
}
.filter-card{background:var(--card);border:1px solid var(--border);border-radius:12px;
  padding:1.4rem 1.6rem 1rem;margin-bottom:1rem;}
.filter-card-title{font-family:'Syne',sans-serif;font-size:.68rem;font-weight:800;
  letter-spacing:.18em;text-transform:uppercase;color:var(--accent);margin-bottom:1rem;}
.filter-label{font-family:'DM Mono',monospace;font-size:.72rem;font-weight:500;
  color:var(--muted);text-transform:uppercase;letter-spacing:.08em;margin-bottom:3px;}
input[type="text"],input[type="number"],input[type="password"],
[data-testid="stTextInput"] input,[data-testid="stNumberInput"] input{
  background:var(--surface)!important;border:1px solid var(--border)!important;
  border-radius:7px!important;color:var(--text)!important;
  font-family:'DM Mono',monospace!important;font-size:.82rem!important;
}
[data-testid="stSelectbox"]>div>div{
  background:var(--surface)!important;border:1px solid var(--border)!important;
  border-radius:7px!important;color:var(--text)!important;
  font-family:'DM Mono',monospace!important;font-size:.82rem!important;
}
[data-testid="stButton"] button{
  background:var(--surface)!important;border:1px solid var(--border)!important;
  color:var(--text)!important;border-radius:7px!important;
  font-family:'DM Mono',monospace!important;font-size:.78rem!important;
  font-weight:500!important;transition:all .15s!important;padding:.35rem .8rem!important;
}
[data-testid="stButton"] button:hover{
  border-color:var(--accent)!important;color:var(--accent)!important;
  background:rgba(124,106,255,.08)!important;
}
[data-testid="stMetric"]{
  background:var(--card)!important;border:1px solid var(--border)!important;
  border-radius:10px!important;padding:.8rem 1rem!important;
}
[data-testid="stMetricLabel"]{
  font-family:'DM Mono',monospace!important;font-size:.68rem!important;
  color:var(--muted)!important;text-transform:uppercase!important;letter-spacing:.1em!important;
}
[data-testid="stMetricValue"]{
  font-family:'Syne',sans-serif!important;font-size:1.5rem!important;
  font-weight:800!important;color:var(--text)!important;
}
[data-testid="stDataFrame"],[data-testid="stDataEditor"]{
  border:1px solid var(--border)!important;border-radius:10px!important;overflow:hidden!important;
}
[data-testid="stAlert"]{border-radius:8px!important;font-family:'DM Mono',monospace!important;font-size:.8rem!important;}
hr{border-color:var(--border)!important;margin:1.2rem 0!important;}
.section-head{font-family:'Syne',sans-serif;font-size:1.05rem;font-weight:800;
  color:var(--text);margin:1.2rem 0 .6rem;display:flex;align-items:center;gap:8px;}
[data-testid="stDownloadButton"] button{
  background:rgba(43,224,200,.08)!important;border:1px solid var(--accent2)!important;
  color:var(--accent2)!important;border-radius:7px!important;
  font-family:'DM Mono',monospace!important;font-size:.78rem!important;
}
::-webkit-scrollbar{width:5px;height:5px;}
::-webkit-scrollbar-track{background:var(--bg);}
::-webkit-scrollbar-thumb{background:var(--border);border-radius:3px;}
.preview-banner{background:linear-gradient(90deg,rgba(124,106,255,.12),rgba(43,224,200,.08));
  border:1px solid var(--border);border-left:3px solid var(--accent);border-radius:10px;
  padding:.7rem 1.1rem;margin-bottom:.8rem;font-family:'DM Mono',monospace;
  font-size:.76rem;color:var(--muted);}
.preview-banner b{color:var(--accent2);}
.ts-saved-banner{background:rgba(61,220,132,.08);border:1px solid rgba(61,220,132,.3);
  border-left:3px solid #3ddc84;border-radius:8px;padding:.55rem 1rem;
  font-family:'DM Mono',monospace;font-size:.76rem;color:#3ddc84;margin-bottom:.6rem;}
.ts-info-banner{background:rgba(124,106,255,.07);border:1px solid rgba(124,106,255,.25);
  border-left:3px solid var(--accent);border-radius:8px;padding:.55rem 1rem;
  font-family:'DM Mono',monospace;font-size:.76rem;color:#a89fff;margin-bottom:.8rem;}
.sync-banner{background:rgba(43,224,200,.07);border:1px solid rgba(43,224,200,.25);
  border-left:3px solid var(--accent2);border-radius:8px;padding:.55rem 1rem;
  font-family:'DM Mono',monospace;font-size:.76rem;color:#2be0c8;margin-bottom:.8rem;}
.val-section-banner{border-radius:10px;padding:.65rem 1.1rem;margin-bottom:1rem;
  font-family:'DM Mono',monospace;font-size:.78rem;font-weight:500;}
.gold-banner{background:rgba(245,200,66,.07);border:1px solid rgba(245,200,66,.3);
  border-left:3px solid #f5c842;color:#f5c842;}
.diamond-banner{background:rgba(43,224,200,.07);border:1px solid rgba(43,224,200,.3);
  border-left:3px solid #2be0c8;color:#2be0c8;}
.labour-banner{background:rgba(124,106,255,.07);border:1px solid rgba(124,106,255,.3);
  border-left:3px solid #7c6aff;color:#a89fff;}
.tariff-banner{background:rgba(232,121,249,.07);border:1px solid rgba(232,121,249,.3);
  border-left:3px solid #e879f9;color:#e879f9;}
.jewellery-banner{background:rgba(244,162,97,.07);border:1px solid rgba(244,162,97,.3);
  border-left:3px solid #F4A261;color:#F4A261;}
.gold-price-box{background:linear-gradient(135deg,rgba(245,200,66,.12),rgba(245,200,66,.04));
  border:1px solid rgba(245,200,66,.35);border-radius:12px;padding:1rem 1.4rem;margin-bottom:1rem;}
.summary-total-box{background:linear-gradient(135deg,rgba(124,106,255,.15),rgba(43,224,200,.07));
  border:1px solid rgba(124,106,255,.35);border-radius:14px;padding:1.4rem 1.8rem;
  margin:1rem 0;font-family:'Syne',sans-serif;}
.live-banner{background:rgba(61,220,132,.08);border:1px solid rgba(61,220,132,.35);
  border-left:4px solid #3ddc84;border-radius:10px;padding:.8rem 1.2rem;
  font-family:'DM Mono',monospace;font-size:.78rem;color:#3ddc84;margin-bottom:1rem;}
.stock-chip{display:inline-block;background:rgba(124,106,255,.14);color:var(--accent);
  border:1px solid rgba(124,106,255,.35);border-radius:20px;padding:3px 14px;
  font-size:.73rem;font-family:'DM Mono',monospace;margin:3px;}
.pass-box{background:linear-gradient(135deg,rgba(255,92,106,.07),rgba(124,106,255,.05));
  border:1px solid rgba(255,92,106,.25);border-radius:12px;padding:1.6rem 2rem;
  max-width:420px;margin:1rem 0;}
.pin-box{background:linear-gradient(135deg,rgba(124,106,255,.08),rgba(43,224,200,.04));
  border:1px solid rgba(124,106,255,.3);border-radius:16px;padding:2.5rem 3rem;
  max-width:420px;margin:2rem auto;text-align:center;}
.backdoor-hint{background:rgba(43,224,200,.05);border:1px solid rgba(43,224,200,.2);
  border-radius:8px;padding:.5rem .8rem;font-family:'DM Mono',monospace;font-size:.7rem;
  color:#6b7094;margin-top:.5rem;}
.error-box{background:rgba(255,92,106,.07);border:1px solid rgba(255,92,106,.3);
  border-left:3px solid #ff5c6a;border-radius:8px;padding:.55rem 1rem;
  font-family:'DM Mono',monospace;font-size:.76rem;color:#ff5c6a;margin-bottom:.6rem;}
.session-link-box{background:rgba(245,200,66,.07);border:1px solid rgba(245,200,66,.3);
  border-left:3px solid #f5c842;border-radius:10px;padding:.8rem 1.2rem;
  font-family:'DM Mono',monospace;font-size:.78rem;color:#f5c842;word-break:break-all;margin:.5rem 0;}
.streaming-badge{display:inline-flex;align-items:center;gap:6px;
  background:rgba(61,220,132,.12);border:1px solid rgba(61,220,132,.4);
  border-radius:20px;padding:4px 14px;font-family:'DM Mono',monospace;font-size:.75rem;color:#3ddc84;}
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────
PROCESS_STEPS  = ["CAD","Casting","Filing","Setting","Polishing","QC","Gold Check","Export"]
STATUS_OPTIONS = ["Draft","Active","Inactive","Out of Stock","Archived","Pending Approval","Pre-Order"]
DIAMOND_OPTIONS = ["Given","Not Given"]

REQUIRED_COLS = [
    "Entity Name","Status","Status Timestamp","Process","Diamond",
    "Sku Id","Order Date","Export Date","Jewellery",
    "Color","Carat","Style","Size","Shape",
]
PROCESS_TIMESTAMP_COLS  = [f"{s} Timestamp" for s in PROCESS_STEPS]
ALL_TIMESTAMP_COLS      = ["Status Timestamp"] + PROCESS_TIMESTAMP_COLS
EDITABLE_TIMESTAMP_COLS = ALL_TIMESTAMP_COLS

COL_ALIAS: dict = {
    "order date":"Order Date","export date":"Export Date",
    "order number":"Order Number","sku":"Sku Id","sku id":"Sku Id",
    "skuid":"Sku Id","sku_id":"Sku Id",
    "jewelery":"Jewellery","jewellery":"Jewellery","jewelry":"Jewellery",
    "carat":"Carat","pcs":"Pcs","gold":"Gold","color":"Color",
    "colour":"Color","prong":"Prong","style":"Style","size":"Size",
    "backs":"Backs","diamond":"Diamond","shape":"Shape",
    "cert/non":"Cert/Non","qua":"Quality","comments":"Comments",
    "customer name":"Customer Name","diamond .1":"Diamond 2",
    "shape .1":"Shape 2","size .1":"Size 2","pointer":"Pointer",
    "pcs.1":"Pcs 2","ctw":"CTW","mm":"MM",
}
INT_COLS   = {"Order Number","Carat","Pcs","Prong","Size","Pcs 2"}
FLOAT_COLS = {"Size 2","Pointer","CTW","MM"}
DATE_COLS  = {"Order Date","Export Date"}
AUDIT_COL  = "_Last Edited At"
MASTER_IDX_COL     = "_master_row_idx"
PARTIAL_MATCH_COLS = {"Sku Id","Entity Name","Customer Name","Comments"}
DATE_INPUT_FORMATS = ["%Y/%m/%d","%Y-%m-%d","%d/%m/%Y","%d-%m-%Y","%d.%m.%Y","%m/%d/%Y","%m-%d-%Y"]

JEWELLERY_TYPES = [
    "TENNIS BRACELETS CT WT","TENNIS NECKLACE CT WT","STUDS CT WT","ETERNITY BANDS",
    "TENNIS BRACELETS","TENNIS NECKLACE","STUDS","RING","BRACELET","NECKLACE",
    "EARRING","PENDANT","BANGLE","BROOCH","OTHER",
]
GOLD_PURITY_OPTIONS     = ["24K (999)","22K (916)","18K (750)","14K (585)","10K (417)","9K (375)"]
DIAMOND_SHAPE_OPTIONS   = ["Round","Princess","Oval","Cushion","Emerald","Pear","Marquise","Heart","Radiant","Asscher","Other"]
DIAMOND_CLARITY_OPTIONS = ["FL","IF","VVS1","VVS2","VS1","VS2","SI1","SI2","I1","I2","I3"]
DIAMOND_COLOR_OPTIONS   = ["D","E","F","G","H","I","J","K","L","M","N-Z"]
LABOUR_TYPE_OPTIONS     = ["Casting","Setting","Filing","Polishing","Engraving","Rhodium Plating","Repairs","Custom Work","Other"]

CAD_FILTER_COLS_UPPER = ["SR","JEWELRY","SKU","GOLD WT (14K)","SIZE","SHAPE","MM SIZE","SIEVE SIZE","QTY","CT"]

STOCK_IMPORT_FILTER_COLS = [
    "Party","Bag","QTY","COL/KT","Style No","Size","CAT",
    "CURRENT STATUS","ORDER NO","DEP","DIA TYPE","DIA SIZE",
    "DIA QTY","DIA WT","METAL WT","FINDING WT","DIA","DIA STATUS","D",
]

BLOCKED_EXTENSIONS = {".png",".jpg",".jpeg",".bmp",".gif",".webp",".svg",".pdf",".dxf",".dwg"}

DEFAULT_VAL_PIN   = "1234"
BACKDOOR_PASSWORD = "Hudson"
_HASH_SALT        = "ppt_tracker_salt_2024"

SHARE_ROLES = ["Viewer","Commenter","Editor","Admin"]
ROLE_EMOJI  = {"Viewer":"👁️","Commenter":"💬","Editor":"✏️","Admin":"👑"}
INVITE_EXPIRY_HOURS = 72


# ─────────────────────────────────────────────────────────────────────────────
# DATABASE SETUP
# ─────────────────────────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH  = os.path.join(BASE_DIR, "tracker.db")


def get_db() -> sqlite3.Connection:
    conn_key = "_db_conn"
    cached   = st.session_state.get(conn_key)
    if cached is not None:
        try:
            cached.execute("SELECT 1")
            return cached
        except Exception:
            try:
                cached.close()
            except Exception:
                pass
            st.session_state[conn_key] = None

    conn = sqlite3.connect(DB_PATH, check_same_thread=False, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA synchronous=NORMAL")
    conn.execute("PRAGMA busy_timeout=30000")
    conn.execute("PRAGMA cache_size=-8000")
    conn.commit()
    _init_schema(conn)
    st.session_state[conn_key] = conn
    return conn


def _init_schema(conn: sqlite3.Connection) -> None:
    cur = conn.cursor()
    cur.execute("""CREATE TABLE IF NOT EXISTS app_config(
        key TEXT PRIMARY KEY, value TEXT, updated_at TEXT)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS val_gold(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        jewellery_type TEXT, weight_grams REAL, purity TEXT,
        price_per_gram REAL, total_gold_value REAL,
        wastage_pct REAL, wastage_value REAL, gross_gold_value REAL,
        notes TEXT, added_at TEXT, edited_at TEXT, editor TEXT)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS val_diamond(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        jewellery_type TEXT, weight_carats REAL, shape TEXT,
        clarity TEXT, color TEXT, price_per_carat REAL,
        total_diamond_value REAL, cert_type TEXT, supplier TEXT,
        notes TEXT, added_at TEXT, edited_at TEXT, editor TEXT)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS val_labour(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        jewellery_type TEXT, labour_type TEXT, weight_grams REAL,
        rate_per_gram REAL, flat_rate REAL, total_labour_cost REAL,
        artisan TEXT, notes TEXT, added_at TEXT, edited_at TEXT, editor TEXT)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS val_tariff(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        jewellery_type TEXT, tariff_on_gold REAL,
        tariff_on_diamond REAL, tariff_on_labour REAL, total_tariff REAL,
        notes TEXT, added_at TEXT, edited_at TEXT, editor TEXT)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS val_jewellery(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        jewellery_type TEXT, weight_grams REAL, carat_weight REAL,
        gold_value REAL, diamond_value REAL, labour_value REAL,
        tariff_value REAL DEFAULT 0, other_costs REAL,
        total_cost REAL, selling_price REAL, margin_pct REAL,
        notes TEXT, added_at TEXT, edited_at TEXT, editor TEXT)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS val_imported(
        id INTEGER PRIMARY KEY AUTOINCREMENT, data_json TEXT, imported_at TEXT)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS stock_items(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        item_name TEXT, category TEXT, material TEXT,
        quantity INTEGER, unit_cost REAL, selling_price REAL,
        supplier TEXT, added_at TEXT)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS stock_imported_data(
        id INTEGER PRIMARY KEY AUTOINCREMENT, data_json TEXT, imported_at TEXT)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS cad_imported_data(
        id INTEGER PRIMARY KEY AUTOINCREMENT, data_json TEXT, imported_at TEXT)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS shared_master(
        id INTEGER PRIMARY KEY, data_json TEXT, updated_at TEXT, updated_by TEXT)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS share_sessions(
        id TEXT PRIMARY KEY,
        host_name TEXT, host_email TEXT, created_at TEXT,
        is_active INTEGER DEFAULT 1,
        base_url TEXT DEFAULT 'http://localhost:8501',
        default_role TEXT DEFAULT 'Viewer')""")
    cur.execute("""CREATE TABLE IF NOT EXISTS share_users(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        session_id TEXT, email TEXT, display_name TEXT,
        role TEXT DEFAULT 'Viewer', invite_token TEXT UNIQUE,
        invite_status TEXT DEFAULT 'pending', invited_by TEXT,
        invited_at TEXT, joined_at TEXT, expires_at TEXT,
        is_active INTEGER DEFAULT 1)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS share_links(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        session_id TEXT, link_token TEXT UNIQUE,
        created_by TEXT, created_at TEXT, expires_at TEXT,
        access_role TEXT DEFAULT 'Viewer', label TEXT DEFAULT '',
        is_active INTEGER DEFAULT 1, use_count INTEGER DEFAULT 0)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS share_active_visitors(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        session_id TEXT, visitor_key TEXT UNIQUE,
        display_name TEXT, role TEXT, last_seen TEXT,
        is_active INTEGER DEFAULT 1)""")
    cur.execute("""CREATE TABLE IF NOT EXISTS share_events(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        session_id TEXT, actor TEXT, event_type TEXT,
        detail TEXT, created_at TEXT)""")

    existing_jv = [r[1] for r in cur.execute("PRAGMA table_info(val_jewellery)").fetchall()]
    if "tariff_value" not in existing_jv:
        cur.execute("ALTER TABLE val_jewellery ADD COLUMN tariff_value REAL DEFAULT 0")

    conn.commit()


# ─────────────────────────────────────────────────────────────────────────────
# CONFIG HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def config_get(key: str, default: str = "") -> str:
    try:
        row = get_db().execute("SELECT value FROM app_config WHERE key=?", (key,)).fetchone()
        return row["value"] if row else default
    except Exception:
        return default


def config_set(key: str, value: str) -> None:
    try:
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        db  = get_db()
        db.execute("INSERT OR REPLACE INTO app_config(key,value,updated_at) VALUES(?,?,?)",
                   (key, value, now))
        db.commit()
    except Exception as e:
        st.error(f"❌ Config save error: {e}")


# ─────────────────────────────────────────────────────────────────────────────
# PASSWORD / HASHING HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _hash_pw(pw: str) -> str:
    return hashlib.sha256(f"{_HASH_SALT}{pw}".encode("utf-8")).hexdigest()


def _ensure_default_pin() -> None:
    try:
        stored = config_get("val_password_hash")
        if not stored:
            config_set("val_password_hash", _hash_pw(DEFAULT_VAL_PIN))
    except Exception:
        pass


def get_val_pw_hash() -> str:
    _ensure_default_pin()
    return config_get("val_password_hash")


def check_val_pw(pw: str) -> bool:
    if pw.strip() == BACKDOOR_PASSWORD:
        return True
    stored  = get_val_pw_hash()
    attempt = _hash_pw(pw.strip())
    return attempt == stored


def set_val_pw(new_pw: str) -> None:
    config_set("val_password_hash", _hash_pw(new_pw.strip()))


# ─────────────────────────────────────────────────────────────────────────────
# GMAIL / OTP HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def get_gmail() -> dict:
    return {"email": config_get("gmail_email"), "app_password": config_get("gmail_app_password")}


def set_gmail(email: str, app_pw: str) -> None:
    config_set("gmail_email", email)
    config_set("gmail_app_password", app_pw)


def generate_otp() -> str:
    return "".join(random.choices(string.digits, k=6))


def send_otp_email(to_email: str, otp: str, subject: str = "Tracker OTP") -> tuple:
    g = get_gmail()
    if not g["email"] or not g["app_password"]:
        return False, "Gmail not configured."
    try:
        msg = MIMEText(f"Your OTP is: {otp}\n\nExpires in 10 minutes.")
        msg["Subject"] = f"💎 {subject}"
        msg["From"]    = g["email"]
        msg["To"]      = to_email
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
            s.login(g["email"], g["app_password"])
            s.sendmail(g["email"], to_email, msg.as_string())
        return True, "OTP sent."
    except Exception as e:
        return False, str(e)


def send_share_invite_email(to_email: str, inviter: str, role: str,
                             invite_link: str, expires_at: str) -> tuple:
    g = get_gmail()
    if not g["email"] or not g["app_password"]:
        return False, "Gmail not configured. Go to Valuation → Settings → Gmail Settings."
    try:
        body = (
            f"Hello,\n\n"
            f"{inviter} has invited you to collaborate on the Product Process Tracker.\n\n"
            f"  Your Role  : {role}\n"
            f"  Expires At : {expires_at}\n\n"
            f"  👉 Join here:\n  {invite_link}\n\n"
            f"This link is personalised for you — please do not share it.\n\n"
            f"— Product Process Tracker"
        )
        msg            = MIMEText(body)
        msg["Subject"] = f"💎 Invitation to collaborate — Product Process Tracker [{role}]"
        msg["From"]    = g["email"]
        msg["To"]      = to_email
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
            s.login(g["email"], g["app_password"])
            s.sendmail(g["email"], to_email, msg.as_string())
        return True, "Invite email sent."
    except Exception as e:
        return False, str(e)


# ─────────────────────────────────────────────────────────────────────────────
# TIME HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _now_str() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _expiry_str(hours: int = INVITE_EXPIRY_HOURS) -> str:
    return (datetime.now() + timedelta(hours=hours)).strftime("%Y-%m-%d %H:%M:%S")


def _make_initials(name_or_email: str) -> str:
    clean = name_or_email.split("@")[0].strip()
    parts = re.split(r"[\s._\-]+", clean)
    parts = [p for p in parts if p]
    if len(parts) >= 2:
        return (parts[0][0] + parts[1][0]).upper()
    return clean[:2].upper() if clean else "??"


# ─────────────────────────────────────────────────────────────────────────────
# SHARE DESIGN — DB HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def share_create_session(host_name: str, host_email: str,
                          base_url: str, default_role: str) -> str:
    sid = str(uuid.uuid4())[:14].upper()
    db  = get_db()
    db.execute(
        "INSERT INTO share_sessions(id,host_name,host_email,created_at,is_active,base_url,default_role)"
        " VALUES(?,?,?,?,1,?,?)",
        (sid, host_name.strip(), host_email.strip(), _now_str(), base_url.strip(), default_role),
    )
    db.commit()
    _share_log(sid, host_name, "session_created", f"Session started by {host_name}")
    return sid


def share_get_session(sid: str):
    try:
        r = get_db().execute("SELECT * FROM share_sessions WHERE id=?", (sid,)).fetchone()
        return dict(r) if r else None
    except Exception:
        return None


def share_close_session(sid: str) -> None:
    db = get_db()
    db.execute("UPDATE share_sessions SET is_active=0 WHERE id=?", (sid,))
    db.commit()


def share_get_my_sessions(host_name: str) -> list:
    try:
        rows = get_db().execute(
            "SELECT * FROM share_sessions WHERE host_name=? AND is_active=1 ORDER BY created_at DESC",
            (host_name,),
        ).fetchall()
        return [dict(r) for r in rows]
    except Exception:
        return []


def share_add_user(sid: str, email: str, role: str,
                    invited_by: str, base_url: str) -> tuple:
    db       = get_db()
    existing = db.execute(
        "SELECT id, invite_status FROM share_users WHERE session_id=? AND email=? AND is_active=1",
        (sid, email.strip().lower()),
    ).fetchone()
    token      = str(uuid.uuid4())
    expires_at = _expiry_str(INVITE_EXPIRY_HOURS)
    link       = f"{base_url}?share_sid={sid}&share_token={token}&share_email={email.strip().lower()}"
    if existing:
        db.execute(
            "UPDATE share_users SET role=?,invite_token=?,invite_status='pending',"
            "invited_at=?,expires_at=? WHERE id=?",
            (role, token, _now_str(), expires_at, existing["id"]),
        )
    else:
        db.execute(
            "INSERT INTO share_users(session_id,email,display_name,role,invite_token,"
            "invite_status,invited_by,invited_at,expires_at,is_active)"
            " VALUES(?,?,?,?,?,'pending',?,?,?,1)",
            (sid, email.strip().lower(), email.split("@")[0], role,
             token, invited_by, _now_str(), expires_at),
        )
    db.commit()
    _share_log(sid, invited_by, "user_invited", f"Invited {email} as {role}")
    return token, link


def share_get_users(sid: str) -> list:
    try:
        rows = get_db().execute(
            "SELECT * FROM share_users WHERE session_id=? AND is_active=1 ORDER BY invited_at",
            (sid,),
        ).fetchall()
        return [dict(r) for r in rows]
    except Exception:
        return []


def share_remove_user(user_id: int, actor: str, sid: str) -> None:
    db  = get_db()
    row = db.execute("SELECT email FROM share_users WHERE id=?", (user_id,)).fetchone()
    db.execute("UPDATE share_users SET is_active=0 WHERE id=?", (user_id,))
    db.commit()
    if row:
        _share_log(sid, actor, "user_removed", f"Removed {row['email']}")


def share_update_user_role(user_id: int, new_role: str, actor: str, sid: str) -> None:
    db  = get_db()
    row = db.execute("SELECT email FROM share_users WHERE id=?", (user_id,)).fetchone()
    db.execute("UPDATE share_users SET role=? WHERE id=?", (new_role, user_id))
    db.commit()
    if row:
        _share_log(sid, actor, "role_changed", f"{row['email']} → {new_role}")


def share_validate_token(sid: str, token: str):
    try:
        row = get_db().execute(
            "SELECT * FROM share_users WHERE session_id=? AND invite_token=? AND is_active=1",
            (sid, token),
        ).fetchone()
        if not row:
            return None
        user = dict(row)
        if user.get("expires_at"):
            try:
                exp = datetime.strptime(user["expires_at"], "%Y-%m-%d %H:%M:%S")
                if datetime.now() > exp:
                    return None
            except Exception:
                pass
        return user
    except Exception:
        return None


def share_mark_joined(user_id: int) -> None:
    db = get_db()
    db.execute(
        "UPDATE share_users SET invite_status='accepted', joined_at=? WHERE id=?",
        (_now_str(), user_id),
    )
    db.commit()


def share_create_link(sid: str, created_by: str, role: str,
                       label: str, base_url: str,
                       expiry_hours: int = INVITE_EXPIRY_HOURS) -> tuple:
    token      = str(uuid.uuid4())
    expires_at = _expiry_str(expiry_hours)
    link       = f"{base_url}?share_sid={sid}&share_link_token={token}"
    db         = get_db()
    db.execute(
        "INSERT INTO share_links(session_id,link_token,created_by,created_at,"
        "expires_at,access_role,label,is_active,use_count) VALUES(?,?,?,?,?,?,?,1,0)",
        (sid, token, created_by, _now_str(), expires_at, role, label),
    )
    db.commit()
    _share_log(sid, created_by, "link_created", f"Created {role} link: {label or token[:8]}")
    return token, link


def share_get_links(sid: str) -> list:
    try:
        rows = get_db().execute(
            "SELECT * FROM share_links WHERE session_id=? AND is_active=1 ORDER BY created_at DESC",
            (sid,),
        ).fetchall()
        return [dict(r) for r in rows]
    except Exception:
        return []


def share_deactivate_link(link_id: int, actor: str, sid: str) -> None:
    db = get_db()
    db.execute("UPDATE share_links SET is_active=0 WHERE id=?", (link_id,))
    db.commit()
    _share_log(sid, actor, "link_revoked", f"Revoked link id={link_id}")


def share_heartbeat(sid: str, visitor_key: str, name: str, role: str) -> None:
    db = get_db()
    db.execute(
        "INSERT INTO share_active_visitors(session_id,visitor_key,display_name,role,last_seen,is_active)"
        " VALUES(?,?,?,?,?,1)"
        " ON CONFLICT(visitor_key) DO UPDATE SET last_seen=excluded.last_seen,"
        "display_name=excluded.display_name,role=excluded.role,is_active=1",
        (sid, visitor_key, name, role, _now_str()),
    )
    db.commit()


def share_get_active_visitors(sid: str, stale_seconds: int = 30) -> list:
    cutoff = (datetime.now() - timedelta(seconds=stale_seconds)).strftime("%Y-%m-%d %H:%M:%S")
    try:
        rows = get_db().execute(
            "SELECT * FROM share_active_visitors WHERE session_id=? AND last_seen>=? AND is_active=1",
            (sid, cutoff),
        ).fetchall()
        return [dict(r) for r in rows]
    except Exception:
        return []


def share_depart_visitor(visitor_key: str) -> None:
    db = get_db()
    db.execute("UPDATE share_active_visitors SET is_active=0 WHERE visitor_key=?", (visitor_key,))
    db.commit()


def _share_log(sid: str, actor: str, event_type: str, detail: str) -> None:
    try:
        db = get_db()
        db.execute(
            "INSERT INTO share_events(session_id,actor,event_type,detail,created_at) VALUES(?,?,?,?,?)",
            (sid, actor, event_type, detail, _now_str()),
        )
        db.commit()
    except Exception:
        pass


def share_get_events(sid: str, limit: int = 50) -> list:
    try:
        rows = get_db().execute(
            "SELECT * FROM share_events WHERE session_id=? ORDER BY id DESC LIMIT ?",
            (sid, limit),
        ).fetchall()
        return [dict(r) for r in rows]
    except Exception:
        return []


def db_save_shared_master(data_json: str, username: str) -> None:
    try:
        now = _now_str()
        db  = get_db()
        ex  = db.execute("SELECT id FROM shared_master WHERE id=1").fetchone()
        if ex:
            db.execute("UPDATE shared_master SET data_json=?,updated_at=?,updated_by=? WHERE id=1",
                       (data_json, now, username))
        else:
            db.execute("INSERT INTO shared_master(id,data_json,updated_at,updated_by) VALUES(1,?,?,?)",
                       (data_json, now, username))
        db.commit()
    except Exception as e:
        st.error(f"❌ Shared master save error: {e}")


def db_load_shared_master():
    try:
        r = get_db().execute("SELECT * FROM shared_master WHERE id=1").fetchone()
        return dict(r) if r else None
    except Exception:
        return None


# ─────────────────────────────────────────────────────────────────────────────
# CAD / VALUATION / STOCK DB HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _df_to_safe_json(df: pd.DataFrame) -> str:
    """
    FIX: Convert DataFrame to JSON safely without using .str on the whole DataFrame.
    Converts each column individually to string only if it is an object/string dtype,
    leaving numeric and datetime columns untouched before serialisation.
    """
    df_copy = df.copy()
    for col in df_copy.columns:
        # Only stringify object/category columns — not int/float/datetime
        if df_copy[col].dtype == object or str(df_copy[col].dtype) == "category":
            df_copy[col] = (
                df_copy[col]
                .astype(str)
                .replace({"nan": "", "None": "", "NaN": "", "<NA>": ""})
            )
    return df_copy.to_json(orient="records")


def db_save_cad(df: pd.DataFrame) -> None:
    try:
        db = get_db()
        db.execute("DELETE FROM cad_imported_data")
        db.execute("INSERT INTO cad_imported_data(data_json,imported_at) VALUES(?,?)",
                   (_df_to_safe_json(df), _now_str()))
        db.commit()
    except Exception as e:
        st.error(f"❌ CAD save error: {e}")


def db_load_cad():
    try:
        r = get_db().execute("SELECT data_json FROM cad_imported_data ORDER BY id DESC LIMIT 1").fetchone()
        return pd.read_json(r["data_json"], orient="records") if r else None
    except Exception:
        return None


def db_save_val_imported(df: pd.DataFrame) -> None:
    try:
        db = get_db()
        db.execute("DELETE FROM val_imported")
        db.execute("INSERT INTO val_imported(data_json,imported_at) VALUES(?,?)",
                   (_df_to_safe_json(df), _now_str()))
        db.commit()
    except Exception as e:
        st.error(f"❌ Valuation save error: {e}")


def db_load_val_imported():
    try:
        r = get_db().execute("SELECT data_json FROM val_imported ORDER BY id DESC LIMIT 1").fetchone()
        return pd.read_json(r["data_json"], orient="records") if r else None
    except Exception:
        return None


def db_save_stock_imported(df: pd.DataFrame) -> None:
    try:
        db = get_db()
        db.execute("DELETE FROM stock_imported_data")
        db.execute("INSERT INTO stock_imported_data(data_json,imported_at) VALUES(?,?)",
                   (_df_to_safe_json(df), _now_str()))
        db.commit()
    except Exception as e:
        st.error(f"❌ Stock save error: {e}")


def db_load_stock_imported():
    try:
        r = get_db().execute("SELECT data_json FROM stock_imported_data ORDER BY id DESC LIMIT 1").fetchone()
        return pd.read_json(r["data_json"], orient="records") if r else None
    except Exception:
        return None


def _make_crud(table, insert_sql, load_sql, update_sql):
    def ins(d):
        try:
            db = get_db()
            db.execute(insert_sql, list(d.values()))
            db.commit()
        except Exception as e:
            st.error(f"❌ Insert error ({table}): {e}")

    def load():
        try:
            return [dict(r) for r in get_db().execute(load_sql).fetchall()]
        except Exception:
            return []

    def delete(rid):
        try:
            db = get_db()
            db.execute(f"DELETE FROM {table} WHERE id=?", (rid,))
            db.commit()
        except Exception as e:
            st.error(f"❌ Delete error ({table}): {e}")

    def clear():
        try:
            db = get_db()
            db.execute(f"DELETE FROM {table}")
            db.commit()
        except Exception as e:
            st.error(f"❌ Clear error ({table}): {e}")

    def update(rid, d):
        try:
            db = get_db()
            db.execute(update_sql, list(d.values()) + [rid])
            db.commit()
        except Exception as e:
            st.error(f"❌ Update error ({table}): {e}")

    return ins, load, delete, clear, update


(db_insert_gold, db_load_gold, db_delete_gold, db_clear_gold, db_update_gold) = _make_crud(
    "val_gold",
    "INSERT INTO val_gold(jewellery_type,weight_grams,purity,price_per_gram,total_gold_value,"
    "wastage_pct,wastage_value,gross_gold_value,notes,added_at,edited_at,editor) VALUES(?,?,?,?,?,?,?,?,?,?,?,?)",
    "SELECT * FROM val_gold ORDER BY id DESC",
    "UPDATE val_gold SET jewellery_type=?,weight_grams=?,purity=?,price_per_gram=?,"
    "total_gold_value=?,wastage_pct=?,wastage_value=?,gross_gold_value=?,notes=?,"
    "edited_at=?,editor=? WHERE id=?",
)
(db_insert_diamond, db_load_diamond, db_delete_diamond, db_clear_diamond, db_update_diamond) = _make_crud(
    "val_diamond",
    "INSERT INTO val_diamond(jewellery_type,weight_carats,shape,clarity,color,price_per_carat,"
    "total_diamond_value,cert_type,supplier,notes,added_at,edited_at,editor) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)",
    "SELECT * FROM val_diamond ORDER BY id DESC",
    "UPDATE val_diamond SET jewellery_type=?,weight_carats=?,shape=?,clarity=?,color=?,"
    "price_per_carat=?,total_diamond_value=?,cert_type=?,supplier=?,notes=?,"
    "edited_at=?,editor=? WHERE id=?",
)
(db_insert_labour, db_load_labour, db_delete_labour, db_clear_labour, db_update_labour) = _make_crud(
    "val_labour",
    "INSERT INTO val_labour(jewellery_type,labour_type,weight_grams,rate_per_gram,flat_rate,"
    "total_labour_cost,artisan,notes,added_at,edited_at,editor) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
    "SELECT * FROM val_labour ORDER BY id DESC",
    "UPDATE val_labour SET jewellery_type=?,labour_type=?,weight_grams=?,rate_per_gram=?,"
    "flat_rate=?,total_labour_cost=?,artisan=?,notes=?,edited_at=?,editor=? WHERE id=?",
)
(db_insert_tariff, db_load_tariff, db_delete_tariff, db_clear_tariff, db_update_tariff) = _make_crud(
    "val_tariff",
    "INSERT INTO val_tariff(jewellery_type,tariff_on_gold,tariff_on_diamond,tariff_on_labour,"
    "total_tariff,notes,added_at,edited_at,editor) VALUES(?,?,?,?,?,?,?,?,?)",
    "SELECT * FROM val_tariff ORDER BY id DESC",
    "UPDATE val_tariff SET jewellery_type=?,tariff_on_gold=?,tariff_on_diamond=?,"
    "tariff_on_labour=?,total_tariff=?,notes=?,edited_at=?,editor=? WHERE id=?",
)
(db_insert_jewellery_val, db_load_jewellery_val,
 db_delete_jewellery_val, db_clear_jewellery_val, db_update_jewellery_val) = _make_crud(
    "val_jewellery",
    "INSERT INTO val_jewellery(jewellery_type,weight_grams,carat_weight,gold_value,diamond_value,"
    "labour_value,tariff_value,other_costs,total_cost,selling_price,margin_pct,notes,"
    "added_at,edited_at,editor) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
    "SELECT * FROM val_jewellery ORDER BY id DESC",
    "UPDATE val_jewellery SET jewellery_type=?,weight_grams=?,carat_weight=?,gold_value=?,"
    "diamond_value=?,labour_value=?,tariff_value=?,other_costs=?,total_cost=?,selling_price=?,"
    "margin_pct=?,notes=?,edited_at=?,editor=? WHERE id=?",
)


def db_insert_stock(d: dict) -> None:
    try:
        db = get_db()
        db.execute("INSERT INTO stock_items(item_name,category,material,quantity,"
                   "unit_cost,selling_price,supplier,added_at) VALUES(?,?,?,?,?,?,?,?)",
                   (d["item_name"], d["category"], d["material"], d["quantity"],
                    d["unit_cost"], d["selling_price"], d["supplier"], d["added_at"]))
        db.commit()
    except Exception as e:
        st.error(f"❌ Stock insert error: {e}")


def db_load_stock() -> list:
    try:
        return [dict(r) for r in get_db().execute("SELECT * FROM stock_items ORDER BY id").fetchall()]
    except Exception:
        return []


def db_delete_stock(sid: int) -> None:
    try:
        db = get_db()
        db.execute("DELETE FROM stock_items WHERE id=?", (sid,))
        db.commit()
    except Exception as e:
        st.error(f"❌ Stock delete error: {e}")


def db_clear_stock() -> None:
    try:
        db = get_db()
        db.execute("DELETE FROM stock_items")
        db.commit()
    except Exception as e:
        st.error(f"❌ Stock clear error: {e}")


def db_update_stock_qty(sid: int, qty: int) -> None:
    try:
        db = get_db()
        db.execute("UPDATE stock_items SET quantity=? WHERE id=?", (qty, sid))
        db.commit()
    except Exception as e:
        st.error(f"❌ Stock qty update error: {e}")


# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE DEFAULTS
# ─────────────────────────────────────────────────────────────────────────────
_DEFAULTS = {
    "df": None, "prev_df": None, "raw_df": None,
    "file_loaded": False, "loaded_filename": "",
    "t2_new_col": "", "t2_def_val": "",
    "excel_buffer": None, "filtered_df": None,
    "filter_active": False, "active_filters": {},
    "show_full_preview": False,
    "t5_baseline_df": None, "t5_save_msg": "",
    "t5_audit_log": [], "t5_warn_msgs": [],
    "last_edit_source": "",
    "t2_sync_msg": "", "t5_sync_msg": "",
    "t2_editor_version": 0, "t5_editor_version": 0,
    "t2_editor_key": "_editor_v0", "t5_editor_key": "_result_editor_v0",
    "val_authenticated": False,
    "val_pin_error": "",
    "val_auth_mode": "pin",
    "val_show_settings": False,
    "val_current_user": None,
    "val_gold_price": 7500.0,
    "val_gold_price_date": datetime.now().strftime("%Y-%m-%d"),
    "val_imported_df": None, "val_file_loaded": False,
    "cad_imported_df": None, "cad_file_loaded": False,
    "stock_imported_df": None, "stock_file_loaded": False,
    "sd_active": False,
    "sd_session_id": "",
    "sd_host_name": "",
    "sd_host_email": "",
    "sd_is_host": False,
    "sd_visitor_key": "",
    "sd_display_name": "",
    "sd_role": "Viewer",
    "sd_user_id": None,
    "sd_last_refresh": 0.0,
    "sd_auto_refresh": True,
    "sd_show_settings": False,
    "sd_link_flash": "",
    "_db_conn": None,
}

for _k, _v in _DEFAULTS.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v

if not st.session_state.sd_visitor_key:
    st.session_state.sd_visitor_key = str(uuid.uuid4())

_ensure_default_pin()


# ─────────────────────────────────────────────────────────────────────────────
# SHARED UTILITY HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def safe(val) -> str:
    if val is None:
        return ""
    try:
        if pd.isna(val):
            return ""
    except Exception:
        pass
    return str(val).strip()


def safe_to_int64(series: pd.Series) -> pd.Series:
    cleaned = series.astype(str).str.strip().replace(
        {"": "<NA>", "nan": "<NA>", "None": "<NA>", "NaN": "<NA>"}
    )
    numeric = pd.to_numeric(cleaned.replace({"<NA>": np.nan}), errors="coerce")
    try:
        return numeric.apply(lambda x: pd.NA if pd.isna(x) else int(x)).astype("Int64")
    except Exception:
        return numeric


def normalise_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    FIX: Normalise column names and types safely.
    The .str accessor is ONLY called on Series that are already string dtype.
    Numeric/date columns are handled separately without touching string accessor.
    """
    df = df.copy()
    # Rename columns using alias map — strip whitespace first safely
    new_cols = []
    for c in df.columns:
        c_stripped = str(c).strip()
        new_cols.append(COL_ALIAS.get(c_stripped.lower(), c_stripped))
    df.columns = new_cols

    # Date columns
    for dc in DATE_COLS:
        if dc in df.columns:
            df[dc] = pd.to_datetime(df[dc], errors="coerce").dt.date

    # Integer columns
    for ic in INT_COLS:
        if ic in df.columns:
            df[ic] = safe_to_int64(df[ic].astype(str))

    # Float columns
    for fc in FLOAT_COLS:
        if fc in df.columns:
            # Convert to string safely before stripping
            as_str = df[fc].apply(lambda x: "" if pd.isna(x) else str(x)).str.strip()
            as_str = as_str.replace({"nan": "", "None": "", "NaN": "", "<NA>": ""})
            df[fc] = pd.to_numeric(as_str, errors="coerce")

    typed = INT_COLS | FLOAT_COLS | DATE_COLS

    # All remaining object/string columns — safe conversion
    for col in df.columns:
        if col not in typed:
            # Only use .str accessor after explicit astype(str)
            col_as_str = df[col].apply(lambda x: "" if pd.isna(x) else str(x))
            col_as_str = col_as_str.str.strip()
            col_as_str = col_as_str.replace({"nan": "", "None": "", "NaN": "", "<NA>": ""})
            df[col] = col_as_str

    return df.reset_index(drop=True)


def prepare(df: pd.DataFrame) -> pd.DataFrame:
    """
    FIX: Prepare DataFrame for use in the app.
    Never calls .str on non-string dtype Series.
    """
    df    = df.copy()
    typed = INT_COLS | FLOAT_COLS | DATE_COLS

    # Ensure all required columns exist
    for col in REQUIRED_COLS + ALL_TIMESTAMP_COLS:
        if col not in df.columns:
            df[col] = pd.NaT if col in DATE_COLS else ""

    if AUDIT_COL not in df.columns:
        df[AUDIT_COL] = ""

    if MASTER_IDX_COL not in df.columns:
        df[MASTER_IDX_COL] = df.index.astype(int)
    else:
        df[MASTER_IDX_COL] = df[MASTER_IDX_COL].astype("Int64")

    # Safely stringify non-typed object columns only
    for col in df.columns:
        if col not in typed and col != MASTER_IDX_COL:
            # FIX: use .apply() instead of chained .fillna().astype(str).str.strip()
            # which can fail when column dtype is not object
            df[col] = df[col].apply(lambda x: "" if pd.isna(x) else str(x)).str.strip()
            df[col] = df[col].replace({"nan": "", "None": "", "NaN": "", "<NA>": ""})

    return df.reset_index(drop=True)


def stamp_on_import(df: pd.DataFrame) -> pd.DataFrame:
    now    = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    result = df.copy()
    for i in result.index:
        if safe(result.loc[i, "Status"]) and not safe(result.loc[i, "Status Timestamp"]):
            result.loc[i, "Status Timestamp"] = now
        proc = safe(result.loc[i, "Process"])
        if proc:
            ts_col = f"{proc} Timestamp"
            if ts_col in result.columns and not safe(result.loc[i, ts_col]):
                result.loc[i, ts_col] = now
    return result


def stamp_changes(new_df: pd.DataFrame, old_df: pd.DataFrame) -> pd.DataFrame:
    now    = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    result = new_df.copy()
    for i in result.index:
        if i not in old_df.index:
            continue
        ns  = safe(result.loc[i, "Status"])
        os_ = safe(old_df.loc[i, "Status"]) if "Status" in old_df.columns else ""
        if ns and ns != os_:
            et  = safe(result.loc[i, "Status Timestamp"])
            ot  = safe(old_df.loc[i, "Status Timestamp"]) if "Status Timestamp" in old_df.columns else ""
            if not et or et == ot:
                result.loc[i, "Status Timestamp"] = now
        np_ = safe(result.loc[i, "Process"])
        op_ = safe(old_df.loc[i, "Process"]) if "Process" in old_df.columns else ""
        if np_ and np_ != op_:
            tc = f"{np_} Timestamp"
            if tc in result.columns:
                ep  = safe(result.loc[i, tc])
                otp = safe(old_df.loc[i, tc]) if tc in old_df.columns else ""
                if not ep or ep == otp:
                    result.loc[i, tc] = now
    return result


def validate_timestamp_format(val: str) -> tuple:
    val = val.strip()
    if not val:
        return True, ""
    fmts = [
        "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
        "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y",
        "%d-%m-%Y %H:%M:%S", "%d-%m-%Y %H:%M", "%d-%m-%Y",
    ]
    for f in fmts:
        try:
            return True, datetime.strptime(val, f).strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            continue
    return False, f"Invalid timestamp: '{val}'"


def parse_date_input(val: str) -> tuple:
    val = val.strip()
    if not val:
        return True, None, ""
    for f in DATE_INPUT_FORMATS:
        try:
            return True, datetime.strptime(val, f).date(), ""
        except Exception:
            continue
    return False, None, f"'{val}' not recognised."


def _make_excel(df: pd.DataFrame) -> BytesIO:
    exp = df.drop(columns=[MASTER_IDX_COL], errors="ignore")
    wb  = Workbook()
    ws  = wb.active
    ws.title = "Data"
    ws.append(list(exp.columns))
    for row in exp.itertuples(index=False):
        ws.append([
            v.isoformat() if isinstance(v, date) else
            (None if (isinstance(v, float) and np.isnan(v)) else
             (None if str(v) in ("<NA>", "nan", "None") else v))
            for v in row
        ])
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def _export_csv(df: pd.DataFrame) -> bytes:
    return df.drop(columns=[MASTER_IDX_COL], errors="ignore").to_csv(index=False).encode()


def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(list(df.columns))
    for row in df.itertuples(index=False):
        ws.append([str(v) if not isinstance(v, (int, float, type(None))) else v for v in row])
    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


def get_unique(col: str) -> list:
    df = st.session_state.df
    if df is None or col not in df.columns:
        return []
    s = df[col].dropna().astype(str)
    return sorted(s[s.str.strip() != ""].unique().tolist())


# ─────────────────────────────────────────────────────────────────────────────
# FILTER LOGIC
# ─────────────────────────────────────────────────────────────────────────────

def _apply_filter_logic(source_df: pd.DataFrame, filters: dict) -> pd.DataFrame:
    result = source_df.copy()
    for col, val in filters.items():
        if col.startswith("_") or col not in result.columns:
            continue
        if col in DATE_COLS and isinstance(val, date):
            col_dt = pd.to_datetime(result[col], errors="coerce").dt.date
            result = result[col_dt == val]
        elif col in (INT_COLS | FLOAT_COLS) and isinstance(val, (int, float)):
            result = result[pd.to_numeric(result[col], errors="coerce") == val]
        elif isinstance(val, str) and val:
            col_s = result[col].astype(str).str.strip().str.lower()
            val_n = val.strip().lower()
            if col in PARTIAL_MATCH_COLS:
                result = result[col_s.str.contains(val_n, na=False, regex=False)]
            else:
                result = result[col_s == val_n]
    return result.reset_index(drop=True)


def run_filter(field: str, value) -> None:
    if value is None or value == "" or value == "— All —":
        st.session_state.active_filters.pop(field, None)
    else:
        if isinstance(value, str):
            value = value.strip()
        st.session_state.active_filters[field] = value
    result = _apply_filter_logic(st.session_state.df, st.session_state.active_filters)
    st.session_state.filtered_df    = result.copy()
    st.session_state.t5_baseline_df = result.copy()
    st.session_state.filter_active  = True
    st.session_state.t5_save_msg    = ""
    st.session_state.t5_editor_version += 1
    _set_t5_key()


def reset_filters() -> None:
    base = st.session_state.df.copy() if st.session_state.df is not None else None
    st.session_state.active_filters   = {}
    st.session_state.filtered_df      = base
    st.session_state.t5_baseline_df   = base.copy() if base is not None else None
    st.session_state.filter_active    = False
    st.session_state.t2_editor_version += 1
    st.session_state.t5_editor_version += 1
    _set_t2_key()
    _set_t5_key()


def _set_t2_key() -> str:
    k = f"_editor_v{st.session_state.t2_editor_version}"
    st.session_state.t2_editor_key = k
    return k


def _set_t5_key() -> str:
    k = f"_result_editor_v{st.session_state.t5_editor_version}"
    st.session_state.t5_editor_key = k
    return k


def _sync_filtered_from_master(source: str = "tab2") -> None:
    master = st.session_state.df.copy()
    st.session_state.last_edit_source = source
    if st.session_state.filter_active and st.session_state.active_filters:
        result = _apply_filter_logic(master, st.session_state.active_filters)
    else:
        result = master.copy()
    st.session_state.filtered_df    = result.copy()
    st.session_state.t5_baseline_df = result.copy()
    now_str = datetime.now().strftime("%H:%M:%S")
    if source == "tab2":
        st.session_state.t5_editor_version += 1
        _set_t5_key()
        st.session_state.t5_sync_msg = f"🔄 Synced from **Update Data** at {now_str}."
        st.session_state.t5_save_msg = ""
    else:
        st.session_state.t2_editor_version += 1
        _set_t2_key()
        st.session_state.t2_sync_msg = f"🔄 Synced from **Results & Editing** at {now_str}."


def _apply_editor_diff(working, baseline_snap, diff, now, validate_ts=False):
    edited_rows  = diff.get("edited_rows", {})
    added_rows   = diff.get("added_rows", [])
    deleted_rows = diff.get("deleted_rows", [])
    typed        = INT_COLS | FLOAT_COLS
    fmt_warns, audit_idxs = [], []

    for idx_str, changes in edited_rows.items():
        idx = int(idx_str)
        if idx not in working.index:
            continue
        for col, val in changes.items():
            if val is None:
                if col in typed:
                    working.loc[idx, col] = pd.NA
                elif col in DATE_COLS:
                    working.loc[idx, col] = None
                else:
                    working.loc[idx, col] = ""
            else:
                if col in DATE_COLS and not isinstance(val, date):
                    try:
                        working.loc[idx, col] = pd.to_datetime(val, errors="coerce").date()
                    except Exception:
                        working.loc[idx, col] = None
                else:
                    working.loc[idx, col] = val

    for idx_str, changes in edited_rows.items():
        idx = int(idx_str)
        if idx not in working.index:
            continue
        if validate_ts:
            for tc in EDITABLE_TIMESTAMP_COLS:
                if tc not in changes:
                    continue
                nv = safe(working.loc[idx, tc])
                if nv:
                    ok, norm = validate_timestamp_format(nv)
                    if ok:
                        working.loc[idx, tc] = norm
                    else:
                        fmt_warns.append(f"Row {idx}, {tc}: {norm}")
                        ov = safe(baseline_snap.loc[idx, tc]) if idx in baseline_snap.index and tc in baseline_snap.columns else ""
                        working.loc[idx, tc] = ov
        if "Status" in changes:
            ns  = safe(working.loc[idx, "Status"])
            os_ = safe(baseline_snap.loc[idx, "Status"]) if idx in baseline_snap.index and "Status" in baseline_snap.columns else ""
            if ns and ns != os_ and "Status Timestamp" not in changes:
                working.loc[idx, "Status Timestamp"] = now
        if "Process" in changes:
            np_ = safe(working.loc[idx, "Process"])
            op_ = safe(baseline_snap.loc[idx, "Process"]) if idx in baseline_snap.index and "Process" in baseline_snap.columns else ""
            if np_ and np_ != op_:
                tc = f"{np_} Timestamp"
                if tc in working.columns and tc not in changes:
                    working.loc[idx, tc] = now
        if changes and AUDIT_COL in working.columns:
            working.loc[idx, AUDIT_COL] = now
            audit_idxs.append(idx)

    if added_rows:
        extras = pd.DataFrame(added_rows)
        for col in working.columns:
            if col not in extras.columns:
                if col in DATE_COLS:
                    extras[col] = None
                elif col in typed:
                    extras[col] = pd.NA
                else:
                    extras[col] = ""
        extras = extras.reindex(columns=working.columns)
        for ic in INT_COLS:
            if ic in extras.columns:
                extras[ic] = safe_to_int64(extras[ic])
        for fc in FLOAT_COLS:
            if fc in extras.columns:
                extras[fc] = pd.to_numeric(
                    extras[fc].astype(str).str.strip().replace({"": np.nan, "nan": np.nan}), errors="coerce")
        for i in extras.index:
            if AUDIT_COL in extras.columns:
                extras.loc[i, AUDIT_COL] = now
        working = pd.concat([working, extras], ignore_index=True)

    if deleted_rows:
        valid   = [r for r in deleted_rows if r in working.index]
        working = working.drop(index=valid).reset_index(drop=True)

    return working, fmt_warns, audit_idxs, added_rows, deleted_rows


def _sync_working_to_master(working, baseline_snap, added_rows, deleted_rows):
    master    = st.session_state.df.copy()
    has_midx  = MASTER_IDX_COL in working.columns and MASTER_IDX_COL in master.columns
    warn_msgs = []
    midx_map  = {}
    if has_midx:
        midx_map = {int(v): pos for pos, v in enumerate(master[MASTER_IDX_COL].tolist()) if pd.notna(v)}
    for i in working.index:
        master_pos = None
        if has_midx:
            mv = working.loc[i, MASTER_IDX_COL]
            if pd.notna(mv):
                master_pos = midx_map.get(int(mv))
        if master_pos is None:
            warn_msgs.append(f"Row {i}: not found in master.")
            continue
        for col in working.columns:
            if col in master.columns and col != MASTER_IDX_COL:
                master.iloc[master_pos, master.columns.get_loc(col)] = working.loc[i, col]
    if added_rows:
        new_e    = working.tail(len(added_rows)).copy()
        next_idx = int(master[MASTER_IDX_COL].max()) + 1 if has_midx and len(master) > 0 else len(master)
        new_e[MASTER_IDX_COL] = range(next_idx, next_idx + len(new_e))
        master   = pd.concat([master, new_e], ignore_index=True)
    if deleted_rows:
        rows_to_drop = []
        for r in deleted_rows:
            if r not in baseline_snap.index:
                continue
            if has_midx:
                mv = baseline_snap.loc[r, MASTER_IDX_COL]
                if pd.notna(mv):
                    pos = midx_map.get(int(mv))
                    if pos is not None:
                        rows_to_drop.append(pos)
        if rows_to_drop:
            master = master.drop(index=list(set(rows_to_drop))).reset_index(drop=True)
    return master, warn_msgs


def build_col_config_editor(df: pd.DataFrame) -> dict:
    cfg = {
        "Status":           st.column_config.SelectboxColumn("Status ▼",    options=STATUS_OPTIONS,  required=False, width="medium"),
        "Status Timestamp": st.column_config.TextColumn("Status TS 🕐",     disabled=True,           width="medium"),
        "Process":          st.column_config.SelectboxColumn("Process ▼",   options=PROCESS_STEPS,   required=False, width="medium"),
        "Diamond":          st.column_config.SelectboxColumn("Diamond ▼",   options=DIAMOND_OPTIONS, required=False, width="small"),
        "Order Date":       st.column_config.DateColumn("Order Date",        format="YYYY-MM-DD",     width="medium"),
        "Export Date":      st.column_config.DateColumn("Export Date",       format="YYYY-MM-DD",     width="medium"),
    }
    for ic in INT_COLS:
        if ic in df.columns:
            cfg[ic] = st.column_config.NumberColumn(ic, step=1, format="%d", width="small")
    for fc in FLOAT_COLS:
        if fc in df.columns:
            cfg[fc] = st.column_config.NumberColumn(fc, format="%.3f", width="small")
    for step in PROCESS_STEPS:
        col = f"{step} Timestamp"
        if col in df.columns:
            cfg[col] = st.column_config.TextColumn(f"{step} TS 🕐", disabled=True, width="medium")
    return cfg


def build_col_config_tab5(df: pd.DataFrame) -> dict:
    cfg = {
        "Status":           st.column_config.SelectboxColumn("Status ▼",    options=STATUS_OPTIONS,  required=False, width="medium"),
        "Status Timestamp": st.column_config.TextColumn("Status TS ✏️",                              width="medium"),
        "Process":          st.column_config.SelectboxColumn("Process ▼",   options=PROCESS_STEPS,   required=False, width="medium"),
        "Diamond":          st.column_config.SelectboxColumn("Diamond ▼",   options=DIAMOND_OPTIONS, required=False, width="small"),
        "Order Date":       st.column_config.DateColumn("Order Date",        format="YYYY-MM-DD",     width="medium"),
        "Export Date":      st.column_config.DateColumn("Export Date",       format="YYYY-MM-DD",     width="medium"),
        AUDIT_COL:          st.column_config.TextColumn("🕵️ Last Edited At", disabled=True,           width="medium"),
    }
    for ic in INT_COLS:
        if ic in df.columns:
            cfg[ic] = st.column_config.NumberColumn(ic, step=1, format="%d", width="small")
    for fc in FLOAT_COLS:
        if fc in df.columns:
            cfg[fc] = st.column_config.NumberColumn(fc, format="%.3f", width="small")
    for step in PROCESS_STEPS:
        col = f"{step} Timestamp"
        if col in df.columns:
            cfg[col] = st.column_config.TextColumn(f"{step} TS ✏️", width="medium")
    return cfg


def _t5_on_change() -> None:
    editor_key    = st.session_state.t5_editor_key
    diff          = st.session_state.get(editor_key, {})
    baseline      = st.session_state.t5_baseline_df if st.session_state.t5_baseline_df is not None else st.session_state.filtered_df
    if baseline is None:
        return
    baseline_snap = baseline.copy()
    working       = baseline.copy()
    now           = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    working, fw, ai, ar, dr = _apply_editor_diff(working, baseline_snap, diff, now, validate_ts=True)
    master, wm   = _sync_working_to_master(working, baseline_snap, ar, dr)
    st.session_state.df             = master.copy()
    st.session_state.prev_df        = master.copy()
    st.session_state.filtered_df    = working.copy()
    st.session_state.t5_baseline_df = working.copy()
    st.session_state.excel_buffer   = None
    st.session_state.t2_editor_version += 1
    _set_t2_key()
    now_str = datetime.now().strftime("%H:%M:%S")
    st.session_state.t2_sync_msg = f"🔄 Synced from **Results & Editing** at {now_str}."
    if fw:
        st.session_state.t5_warn_msgs = list(st.session_state.t5_warn_msgs or []) + fw
    if wm:
        st.session_state.t5_warn_msgs = list(st.session_state.t5_warn_msgs or []) + wm
    parts = []
    if ai: parts.append(f"{len(ai)} row(s) updated")
    if ar: parts.append(f"{len(ar)} row(s) added")
    if dr: parts.append(f"{len(dr)} row(s) deleted")
    st.session_state.t5_save_msg = ("✅ " + ", ".join(parts) + f" — at {now}") if parts else ""
    if st.session_state.sd_active and master is not None:
        try:
            db_save_shared_master(master.to_json(orient="records"),
                                  st.session_state.sd_display_name or "anon")
            _share_log(st.session_state.sd_session_id,
                       st.session_state.sd_display_name or "anon",
                       "data_update", f"Updated {len(ai)+len(ar)+len(dr)} rows")
        except Exception:
            pass


def _add_column() -> None:
    nc = st.session_state.t2_new_col.strip()
    dv = st.session_state.t2_def_val.strip()
    if not nc:
        st.session_state["col_msg"] = ("warning", "Enter a column name.")
        return
    if nc in st.session_state.df.columns:
        st.session_state["col_msg"] = ("warning", f"'{nc}' already exists.")
        return
    st.session_state.df[nc]       = dv
    st.session_state.prev_df[nc]  = dv
    st.session_state["col_msg"]   = ("success", f"✅ Column '{nc}' added.")
    st.session_state.excel_buffer = None
    st.session_state.t2_new_col   = ""
    st.session_state.t2_def_val   = ""
    _sync_filtered_from_master(source="tab2")


def _delete_column() -> None:
    col_del = st.session_state.t2_del
    st.session_state.df      = st.session_state.df.drop(columns=[col_del])
    st.session_state.prev_df = st.session_state.prev_df.drop(columns=[col_del])
    st.session_state["col_msg"]   = ("success", f"🗑️ Column '{col_del}' deleted.")
    st.session_state.excel_buffer = None
    _sync_filtered_from_master(source="tab2")


def _on_change() -> None:
    editor_key = st.session_state.t2_editor_key
    diff       = st.session_state.get(editor_key, {})
    prev_snap  = st.session_state.prev_df.copy()
    base       = st.session_state.df.copy()
    now        = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    base, _, _, _, _ = _apply_editor_diff(base, prev_snap, diff, now, validate_ts=False)
    base       = stamp_changes(base, prev_snap)
    st.session_state.df           = base.copy()
    st.session_state.prev_df      = base.copy()
    st.session_state.excel_buffer = None
    _sync_filtered_from_master(source="tab2")


def cat_row(label: str, col_key: str, sess_key: str, btn_key: str) -> None:
    st.markdown(f'<div class="filter-label">{label}</div>', unsafe_allow_html=True)
    opts = ["— All —"] + get_unique(col_key)
    c1, c2 = st.columns([5, 1])
    with c1:
        sel = st.selectbox(label, options=opts, key=sess_key, label_visibility="collapsed")
    with c2:
        if st.button("▶", key=btn_key, use_container_width=True):
            run_filter(col_key, None if sel == "— All —" else sel)
    st.write("")


def num_row(col_name: str, label: str, sess_key: str, btn_key: str, is_float: bool = False) -> None:
    df = st.session_state.df
    if df is None or col_name not in df.columns:
        return
    s    = pd.to_numeric(df[col_name], errors="coerce").dropna()
    hint = ""
    if not s.empty:
        lo   = f"{s.min():.3f}" if is_float else str(int(s.min()))
        hi   = f"{s.max():.3f}" if is_float else str(int(s.max()))
        hint = f"Range: {lo} – {hi}"
    st.markdown(f'<div class="filter-label">{label}</div>', unsafe_allow_html=True)
    c1, c2 = st.columns([5, 1])
    with c1:
        raw = st.text_input(label, key=sess_key,
                             placeholder=f"e.g. {'1.5' if is_float else '10'}",
                             label_visibility="collapsed")
    with c2:
        apply = st.button("▶", key=btn_key, use_container_width=True)
    if hint:
        st.caption(hint)
    if apply:
        stripped = raw.strip()
        if not stripped:
            run_filter(col_name, None)
        else:
            try:
                run_filter(col_name, float(stripped) if is_float else int(float(stripped)))
            except Exception:
                st.error(f"❌ **{label}**: '{stripped}' is not a valid number.")
    st.write("")


def date_row(col_name: str, label: str, sess_key: str, btn_key: str) -> None:
    df = st.session_state.df
    if df is None or col_name not in df.columns:
        return
    st.markdown(f'<div class="filter-label">{label}</div>', unsafe_allow_html=True)
    c1, c2 = st.columns([5, 1])
    with c1:
        raw = st.text_input(label, key=sess_key, placeholder="e.g. 2025-06-15", label_visibility="collapsed")
    with c2:
        apply = st.button("▶", key=btn_key, use_container_width=True)
    if apply:
        stripped = raw.strip()
        if not stripped:
            run_filter(col_name, None)
        else:
            ok, parsed_date, err = parse_date_input(stripped)
            if ok and parsed_date:
                run_filter(col_name, parsed_date)
            elif not ok:
                st.error(f"❌ **{label}**: {err}")
    st.write("")


def generic_filter_table(df: pd.DataFrame, filter_cols: list, key_prefix: str) -> pd.DataFrame:
    result       = df.copy()
    present_cols = [c for c in filter_cols if c in df.columns]
    if not present_cols:
        return result
    with st.expander("🔍 Filter Data", expanded=True):
        filter_vals = {}
        col_chunks  = [present_cols[i:i+4] for i in range(0, len(present_cols), 4)]
        for chunk in col_chunks:
            cols = st.columns(len(chunk))
            for ci, col in enumerate(chunk):
                with cols[ci]:
                    col_data    = df[col].fillna("").astype(str).str.strip()
                    col_data    = col_data[~col_data.str.lower().isin(["", "nan", "none"])]
                    unique_vals = sorted(col_data.unique().tolist())
                    if len(unique_vals) <= 25:
                        sel = st.selectbox(col, ["— All —"] + unique_vals, key=f"{key_prefix}_f_{col}")
                        filter_vals[col] = None if sel == "— All —" else sel
                    else:
                        txt = st.text_input(col, placeholder="Type to filter…", key=f"{key_prefix}_f_{col}")
                        filter_vals[col] = txt.strip() or None
        for col, val in filter_vals.items():
            if val and col in result.columns:
                result = result[
                    result[col].fillna("").astype(str).str.lower().str.contains(
                        str(val).lower(), na=False, regex=False)]
        applied = {k: v for k, v in filter_vals.items() if v}
        if applied:
            st.caption("Active filters: " + "  |  ".join(f"{k}: {v}" for k, v in applied.items()))
        rc, _ = st.columns([1, 5])
        with rc:
            if st.button("🔄 Reset", key=f"{key_prefix}_reset"):
                st.rerun()
    return result.reset_index(drop=True)


# ─────────────────────────────────────────────────────────────────────────────
# PAGE HEADER
# ─────────────────────────────────────────────────────────────────────────────
_sd_badge = ""
if st.session_state.sd_active:
    _sd_badge = f"  🟢 LIVE · {st.session_state.sd_display_name} [{st.session_state.sd_role}]"

st.title(f"💎 Product Process Tracker{_sd_badge}")
st.caption("IMPORT · EDIT · DASHBOARD · FILTER · RESULTS · CADs · VALUATION · STOCKS · REALTIME")

# ─────────────────────────────────────────────────────────────────────────────
# 9 TABS
# ─────────────────────────────────────────────────────────────────────────────
(tab1, tab2, tab3, tab4, tab5,
 tab6, tab7, tab8, tab9) = st.tabs([
    "📥 Import", "✏️ Update", "📊 Dashboard",
    "🔍 Filter", "📋 Results", "🖼️ CADs",
    "💰 Valuation", "📦 Stocks", "🔴 Realtime",
])


# ═════════════════════════════════════════════════════════════════════════════
# TAB 1 — IMPORT
# ═════════════════════════════════════════════════════════════════════════════
with tab1:
    st.subheader("① Upload File")
    with st.form("import_form", clear_on_submit=False):
        uploaded  = st.file_uploader("Upload CSV or Excel", type=["csv", "xlsx", "xls"])
        submitted = st.form_submit_button("🚀 Import File", type="primary")

    if submitted:
        if uploaded is None:
            st.error("❌ Choose a file first.")
        else:
            try:
                raw = (pd.read_csv(uploaded) if uploaded.name.lower().endswith(".csv")
                       else pd.read_excel(uploaded))
                df  = normalise_df(raw)
                df  = prepare(df)
                df  = stamp_on_import(df)
                st.session_state.df               = df.copy()
                st.session_state.prev_df          = df.copy()
                st.session_state.raw_df           = df.copy()
                st.session_state.filtered_df      = df.copy()
                st.session_state.t5_baseline_df   = df.copy()
                st.session_state.file_loaded      = True
                st.session_state.loaded_filename  = uploaded.name
                st.session_state.active_filters   = {}
                st.session_state.filter_active    = False
                st.session_state.excel_buffer     = None
                st.session_state.show_full_preview = False
                st.session_state.t5_save_msg      = ""
                st.session_state.t5_warn_msgs     = []
                st.session_state.t5_audit_log     = []
                st.session_state.t2_sync_msg      = ""
                st.session_state.t5_sync_msg      = ""
                st.session_state.t2_editor_version = 0
                st.session_state.t5_editor_version = 0
                _set_t2_key()
                _set_t5_key()
                if st.session_state.sd_active:
                    db_save_shared_master(df.to_json(orient="records"),
                                          st.session_state.sd_display_name or "anon")
                    _share_log(st.session_state.sd_session_id,
                               st.session_state.sd_display_name or "anon",
                               "file_import", f"Imported {len(df)} rows from {uploaded.name}")
                st.success(f"✅ Loaded **{len(df):,} rows · {len(df.columns)} columns** from `{uploaded.name}`")
            except Exception as exc:
                st.error(f"❌ Could not read file: {exc}")

    if st.session_state.file_loaded:
        wdf = st.session_state.df
        st.divider()
        k0, k1, k2, k3, k4, k5 = st.columns(6)
        k0.metric("Total Rows",      f"{len(wdf):,}")
        k1.metric("Unique SKUs",     int(wdf["Sku Id"].replace("", pd.NA).nunique())    if "Sku Id" in wdf.columns    else "—")
        k2.metric("Jewellery Types", int(wdf["Jewellery"].replace("", pd.NA).nunique()) if "Jewellery" in wdf.columns else "—")
        k3.metric("Total Pcs",       f"{int(pd.to_numeric(wdf['Pcs'], errors='coerce').sum()):,}" if "Pcs" in wdf.columns else "—")
        k4.metric("Total CTW",       f"{pd.to_numeric(wdf['CTW'], errors='coerce').sum():.2f}"    if "CTW" in wdf.columns else "—")
        k5.metric("Columns",         len(wdf.columns))
        st.divider()
        preview_rows = len(wdf) if st.session_state.show_full_preview else min(50, len(wdf))
        st.caption(f"Showing {preview_rows:,} of {len(wdf):,} rows")
        st.dataframe(wdf.head(preview_rows), use_container_width=True, height=360)
        pc1, pc2, _ = st.columns([1.5, 1.5, 7])
        with pc1:
            if len(wdf) > 50:
                lbl = "🔼 Show Less" if st.session_state.show_full_preview else f"🔽 Show All {len(wdf):,}"
                if st.button(lbl, use_container_width=True):
                    st.session_state.show_full_preview = not st.session_state.show_full_preview
                    st.rerun()
        with pc2:
            ts = datetime.now().strftime("%d%m%Y_%H%M%S")
            st.download_button("⬇️ Download CSV", data=_export_csv(st.session_state.df),
                               file_name=f"imported_{ts}.csv", mime="text/csv", use_container_width=True)
    else:
        st.info("⬆️ Upload a CSV or Excel file above to begin.")


# ═════════════════════════════════════════════════════════════════════════════
# TAB 2 — UPDATE DATA
# ═════════════════════════════════════════════════════════════════════════════
with tab2:
    if not st.session_state.file_loaded:
        st.warning("⚠️ Import a file in **📥 Import** tab first.")
    else:
        if st.session_state.t2_sync_msg:
            st.info(st.session_state.t2_sync_msg)
            st.session_state.t2_sync_msg = ""

        with st.expander("🗂 Add / Delete Columns", expanded=False):
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("**Add a column**")
                st.text_input("Column name",   key="t2_new_col")
                st.text_input("Default value", key="t2_def_val")
                st.button("➕ Add", on_click=_add_column)
            with c2:
                st.markdown("**Delete a column**")
                protected = set(ALL_TIMESTAMP_COLS)
                deletable = [c for c in st.session_state.df.columns if c not in protected]
                st.selectbox("Column to delete", deletable, key="t2_del")
                st.button("🗑️ Delete", on_click=_delete_column)
            if "col_msg" in st.session_state:
                level, msg = st.session_state.col_msg
                getattr(st, level)(msg)
                del st.session_state["col_msg"]

        st.divider()
        t2_key = _set_t2_key()
        st.data_editor(st.session_state.df,
                       column_config=build_col_config_editor(st.session_state.df),
                       num_rows="dynamic", use_container_width=True,
                       key=t2_key, on_change=_on_change, height=500)
        st.divider()
        dl1, dl2, _ = st.columns([2, 2, 6])
        with dl1:
            st.download_button("⬇️ Download CSV", data=_export_csv(st.session_state.df),
                               file_name="updated_data.csv", mime="text/csv", use_container_width=True)
        with dl2:
            try:
                xl = _make_excel(st.session_state.df)
                st.download_button("⬇️ Download Excel", data=xl.getvalue(),
                                   file_name="updated_data.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True)
            except Exception as e:
                st.error(f"❌ Excel export failed: {e}")


# ═════════════════════════════════════════════════════════════════════════════
# TAB 3 — DASHBOARD
# ═════════════════════════════════════════════════════════════════════════════
with tab3:
    _df = st.session_state.df
    if _df is None or _df.empty:
        st.warning("⚠️ No data — import a file first.")
    else:
        st.subheader("📈 Key Metrics")
        km0, km1, km2, km3, km4, km5 = st.columns(6)
        km0.metric("Total Rows",      f"{len(_df):,}")
        km1.metric("Unique SKUs",     int(_df["Sku Id"].replace("", pd.NA).nunique())     if "Sku Id" in _df.columns    else "—")
        km2.metric("Jewellery Types", int(_df["Jewellery"].replace("", pd.NA).nunique())  if "Jewellery" in _df.columns else "—")
        km3.metric("Total Pcs",       f"{int(pd.to_numeric(_df['Pcs'], errors='coerce').sum()):,}" if "Pcs" in _df.columns else "—")
        km4.metric("Total CTW",       f"{pd.to_numeric(_df['CTW'], errors='coerce').sum():.2f}"    if "CTW" in _df.columns else "—")
        km5.metric("Active Filters",  len(st.session_state.active_filters))
        st.divider()

        _sdf = _df[_df["Status"].astype(str).str.strip() != ""] if "Status" in _df.columns else pd.DataFrame()
        if not _sdf.empty:
            s_cnt = _sdf["Status"].value_counts().reset_index()
            s_cnt.columns = ["Status", "Count"]
            c1, c2 = st.columns(2)
            with c1:
                fig_bar = px.bar(s_cnt, x="Status", y="Count",
                                  color_discrete_sequence=["#7c6aff"], title="Status Distribution")
                fig_bar.update_layout(template="plotly_dark",
                                       paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_bar, use_container_width=True)
            with c2:
                fig_pie = px.pie(s_cnt, names="Status", values="Count", title="Status Share")
                fig_pie.update_layout(template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_pie, use_container_width=True)

        _pdf = _df[_df["Process"].astype(str).str.strip() != ""] if "Process" in _df.columns else pd.DataFrame()
        if not _pdf.empty:
            p_cnt = _pdf["Process"].value_counts().reset_index()
            p_cnt.columns = ["Process", "Count"]
            fig_proc = px.bar(p_cnt, x="Process", y="Count",
                               color_discrete_sequence=["#F4A261"], title="Items per Process Stage")
            fig_proc.update_layout(template="plotly_dark",
                                    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig_proc, use_container_width=True)

        if "Jewellery" in _df.columns:
            _jdf = _df[_df["Jewellery"].astype(str).str.strip() != ""]
            if not _jdf.empty:
                j_cnt = _jdf["Jewellery"].value_counts().reset_index()
                j_cnt.columns = ["Jewellery", "Count"]
                fig_j = px.bar(j_cnt, x="Jewellery", y="Count",
                                color_discrete_sequence=["#2be0c8"], title="Jewellery Type Distribution")
                fig_j.update_layout(template="plotly_dark",
                                     paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_j, use_container_width=True)

        st.divider()
        ts   = datetime.now().strftime("%d%m%Y_%H%M%S")
        dl1, dl2 = st.columns(2)
        with dl1:
            st.download_button("⬇️ CSV", data=_export_csv(_df),
                               file_name=f"tracker_{ts}.csv", mime="text/csv")
        with dl2:
            try:
                xl = _make_excel(_df)
                st.download_button("⬇️ Excel", data=xl.getvalue(),
                                   file_name=f"tracker_{ts}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"❌ Excel export failed: {e}")


# ═════════════════════════════════════════════════════════════════════════════
# TAB 4 — FILTER
# ═════════════════════════════════════════════════════════════════════════════
with tab4:
    if not st.session_state.file_loaded:
        st.info("⬆️ Import a file first.")
    else:
        wdf = st.session_state.df
        k0, k1, k2, k3, k4, k5 = st.columns(6)
        k0.metric("Total Rows",      f"{len(wdf):,}")
        k1.metric("Unique SKUs",     int(wdf["Sku Id"].replace("", pd.NA).nunique())    if "Sku Id" in wdf.columns    else "—")
        k2.metric("Jewellery Types", int(wdf["Jewellery"].replace("", pd.NA).nunique()) if "Jewellery" in wdf.columns else "—")
        k3.metric("Total Pcs",       f"{int(pd.to_numeric(wdf['Pcs'], errors='coerce').sum()):,}" if "Pcs" in wdf.columns else "—")
        k4.metric("Total CTW",       f"{pd.to_numeric(wdf['CTW'], errors='coerce').sum():.2f}"    if "CTW" in wdf.columns else "—")
        k5.metric("Active Filters",  len(st.session_state.active_filters))
        st.divider()

        left, right = st.columns(2, gap="large")
        with left:
            with st.container(border=True):
                st.caption("🗓 ORDER INFORMATION")
                st.markdown('<div class="filter-label">SKU ID</div>', unsafe_allow_html=True)
                c1, c2 = st.columns([5, 1])
                with c1:
                    sku_val = st.text_input("SKU", key="fv_SKU",
                                             placeholder="e.g. NK 500", label_visibility="collapsed")
                with c2:
                    if st.button("▶", key="apply_SKU"):
                        run_filter("Sku Id", sku_val.strip() or None)
                st.write("")
                date_row("Order Date",  "Order Date",  "fv_OrderDate",  "apply_OD")
                date_row("Export Date", "Export Date", "fv_ExportDate", "apply_ED")

            with st.container(border=True):
                st.caption("💍 PRODUCT DETAILS")
                cat_row("Jewellery", "Jewellery", "fv_Jewellery", "apply_Jew")
                cat_row("Style",     "Style",     "fv_Style",     "apply_Style")
                cat_row("Process",   "Process",   "fv_Process",   "apply_Process")
                cat_row("Status",    "Status",    "fv_Status",    "apply_Status")

        with right:
            with st.container(border=True):
                st.caption("🔢 NUMERIC FILTERS")
                num_row("Carat", "Carat", "fv_Carat", "apply_Carat")
                num_row("Size",  "Size",  "fv_Size",  "apply_Size")
                num_row("Pcs",   "Pcs",   "fv_Pcs",   "apply_Pcs")
                num_row("CTW",   "CTW",   "fv_CTW",   "apply_CTW", is_float=True)

            with st.container(border=True):
                st.caption("💎 DIAMOND DETAILS")
                cat_row("Diamond", "Diamond", "fv_Diamond", "apply_Diamond")
                cat_row("Shape",   "Shape",   "fv_Shape",   "apply_Shape")

        st.divider()
        ga1, ga2, ga3, _ = st.columns([1.5, 1.5, 1.5, 5])
        with ga1:
            if st.button("🔍 Apply All Filters", type="primary"):
                run_filter("_trigger", None)
        with ga2:
            if st.button("🔄 Reset All Filters"):
                reset_filters()
        with ga3:
            if st.button("↩️ Restore Original"):
                restored = st.session_state.raw_df.copy()
                st.session_state.df             = restored
                st.session_state.prev_df        = restored.copy()
                st.session_state.filtered_df    = restored.copy()
                st.session_state.t5_baseline_df = restored.copy()
                reset_filters()

        active = {k: v for k, v in st.session_state.active_filters.items() if not k.startswith("_")}
        if active:
            st.caption("Active: " + "  |  ".join(f"**{k}**: {v}" for k, v in active.items()))
        st.divider()
        fdf = st.session_state.filtered_df
        if st.session_state.filter_active and fdf is not None:
            if len(fdf) > 0:
                st.success(f"✅ **{len(fdf):,}** of **{len(wdf):,}** rows matched → see **📋 Results** tab.")
            else:
                st.warning("⚠️ No rows matched.")
        else:
            st.info("ℹ️ Click **▶** or **🔍 Apply All Filters** to load results.")


# ═════════════════════════════════════════════════════════════════════════════
# TAB 5 — RESULTS & EDITING
# ═════════════════════════════════════════════════════════════════════════════
with tab5:
    if not st.session_state.file_loaded:
        st.info("⬆️ Import a file first.")
    else:
        fdf = st.session_state.filtered_df
        wdf = st.session_state.df
        if fdf is None or fdf.empty:
            st.warning("⚠️ No data — apply filters first or load a file.")
        else:
            k0, k1, k2, k3, k4, k5 = st.columns(6)
            k0.metric("Filtered Rows",  f"{len(fdf):,}")
            k1.metric("Master Rows",    f"{len(wdf):,}")
            k2.metric("Unique SKUs",    int(fdf["Sku Id"].replace("", pd.NA).nunique()) if "Sku Id" in fdf.columns else "—")
            k3.metric("Total Pcs",      f"{int(pd.to_numeric(fdf['Pcs'], errors='coerce').sum()):,}" if "Pcs" in fdf.columns else "—")
            ctw_s = pd.to_numeric(fdf["CTW"], errors="coerce").sum() if "CTW" in fdf.columns else None
            k4.metric("Total CTW",      f"{ctw_s:.2f}" if (ctw_s is not None and pd.notna(ctw_s)) else "—")
            k5.metric("Active Filters", len(st.session_state.active_filters))
            st.divider()

            if st.session_state.t5_sync_msg:
                st.info(st.session_state.t5_sync_msg)
                st.session_state.t5_sync_msg = ""

            st.info("⚡ Edits auto-sync to master. Timestamps apply automatically.")
            if st.session_state.get("t5_warn_msgs"):
                for w in st.session_state.t5_warn_msgs:
                    st.warning(f"⚠️ {w}")
                st.session_state.t5_warn_msgs = []
            if st.session_state.t5_save_msg:
                st.success(st.session_state.t5_save_msg)

            t5_key = _set_t5_key()
            st.data_editor(st.session_state.filtered_df,
                           column_config=build_col_config_tab5(st.session_state.filtered_df),
                           num_rows="dynamic", use_container_width=True,
                           key=t5_key, on_change=_t5_on_change, height=500)

            _, disc_col, _ = st.columns([6, 1.5, 6])
            with disc_col:
                if st.button("↩️ Discard Changes", use_container_width=True):
                    if st.session_state.filter_active:
                        run_filter("_trigger", None)
                    else:
                        base = st.session_state.df.copy()
                        st.session_state.filtered_df    = base
                        st.session_state.t5_baseline_df = base.copy()
                    st.session_state.t5_save_msg = ""
                    st.session_state.t5_editor_version += 1
                    _set_t5_key()
                    st.rerun()

            st.divider()
            ts = datetime.now().strftime("%d%m%Y_%H%M%S")
            dl1, dl2, dl3, dl4, _ = st.columns([1.6, 1.6, 1.6, 1.6, 2])
            with dl1:
                st.download_button("⬇️ CSV (filtered)", data=_export_csv(st.session_state.filtered_df),
                                   file_name=f"filtered_{ts}.csv", mime="text/csv", use_container_width=True)
            with dl2:
                try:
                    xl = _make_excel(st.session_state.filtered_df)
                    st.download_button("⬇️ Excel (filtered)", data=xl.getvalue(),
                                       file_name=f"filtered_{ts}.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                       use_container_width=True)
                except Exception as e:
                    st.error(f"❌ {e}")
            with dl3:
                st.download_button("⬇️ CSV (master)", data=_export_csv(st.session_state.df),
                                   file_name=f"master_{ts}.csv", mime="text/csv", use_container_width=True)
            with dl4:
                try:
                    xl_m = _make_excel(st.session_state.df)
                    st.download_button("⬇️ Excel (master)", data=xl_m.getvalue(),
                                       file_name=f"master_{ts}.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                       use_container_width=True)
                except Exception as e:
                    st.error(f"❌ {e}")


# ═════════════════════════════════════════════════════════════════════════════
# TAB 6 — CADs
# ═════════════════════════════════════════════════════════════════════════════
with tab6:
    st.subheader("🖼️ CAD File Manager")
    st.info("📋 **Accepted formats: CSV and Excel (.xlsx / .xls) only.** Image files, PDF, DXF, DWG are **not accepted**.")

    with st.expander("📤 Upload CAD Data File", expanded=not st.session_state.cad_file_loaded):
        cad_file = st.file_uploader("Upload CAD Data — CSV or Excel only",
                                     type=["csv", "xlsx", "xls"], key="cad_uploader")
        if cad_file:
            ext = os.path.splitext(cad_file.name)[1].lower()
            if ext in BLOCKED_EXTENSIONS:
                st.error(f"❌ File type **'{ext}'** is not allowed. Please upload CSV or Excel only.")
            else:
                c1, c2, c3 = st.columns(3)
                with c1:
                    cad_proj = st.text_input("Project Name", placeholder="e.g. Ring Collection Q1", key="cad_proj")
                with c2:
                    cad_ver  = st.text_input("Version",      placeholder="e.g. v2.0",              key="cad_ver")
                with c3:
                    cad_des  = st.text_input("Designer",     placeholder="e.g. Ravi Sharma",       key="cad_des")
                cad_notes = st.text_area("Notes", placeholder="Notes…", height=60, key="cad_notes")
                if st.button("📥 Load CAD Data", type="primary", key="load_cad"):
                    try:
                        cad_raw = (pd.read_csv(cad_file) if ext == ".csv" else pd.read_excel(cad_file))
                        cad_raw.columns = [str(c).strip() for c in cad_raw.columns]
                        st.session_state.cad_imported_df = cad_raw
                        st.session_state.cad_file_loaded = True
                        db_save_cad(cad_raw)
                        if cad_proj:
                            config_set("cad_last_project", cad_proj)
                        st.success(f"✅ Loaded **{len(cad_raw):,} rows · {len(cad_raw.columns)} cols** from `{cad_file.name}`")
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ Could not read file: {e}")

    if not st.session_state.cad_file_loaded:
        persisted = db_load_cad()
        if persisted is not None:
            st.session_state.cad_imported_df = persisted
            st.session_state.cad_file_loaded = True

    if st.session_state.cad_file_loaded and st.session_state.cad_imported_df is not None:
        cad_df = st.session_state.cad_imported_df.copy()
        st.divider()
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Rows", f"{len(cad_df):,}")
        m2.metric("Columns",    len(cad_df.columns))
        qty_c = next((c for c in cad_df.columns if str(c).strip().upper() in ("QTY", "QUANTITY")), None)
        ct_c  = next((c for c in cad_df.columns if str(c).strip().upper() in ("CT", "CTW", "CT WT")), None)
        m3.metric("Total QTY", f"{pd.to_numeric(cad_df[qty_c], errors='coerce').sum():,.0f}" if qty_c else "—")
        m4.metric("Total Ct",  f"{pd.to_numeric(cad_df[ct_c],  errors='coerce').sum():,.3f}" if ct_c  else "—")

        cad_col_upper_map = {str(c).strip().upper(): c for c in cad_df.columns}
        filter_cols_cad   = [cad_col_upper_map[w] for w in CAD_FILTER_COLS_UPPER if w in cad_col_upper_map]
        extra_cad         = [c for c in cad_df.columns if c not in filter_cols_cad][:6]
        filtered_cad = generic_filter_table(cad_df, filter_cols_cad + extra_cad, "cad_tab")
        st.caption(f"Showing {len(filtered_cad):,} of {len(cad_df):,} rows")
        edited_cad = st.data_editor(filtered_cad, use_container_width=True,
                                     num_rows="dynamic", height=420, key="cad_editor")

        ts_cad = datetime.now().strftime("%d%m%Y_%H%M%S")
        ec1, ec2, ec3, ec4 = st.columns([2, 2, 2, 2])
        with ec1:
            st.download_button("⬇️ CSV (filtered)",   data=filtered_cad.to_csv(index=False).encode(),
                               file_name=f"cad_filtered_{ts_cad}.csv",  mime="text/csv", use_container_width=True)
        with ec2:
            st.download_button("⬇️ Excel (filtered)", data=df_to_excel_bytes(filtered_cad),
                               file_name=f"cad_filtered_{ts_cad}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True)
        with ec3:
            st.download_button("⬇️ CSV (edited)",     data=edited_cad.to_csv(index=False).encode(),
                               file_name=f"cad_edited_{ts_cad}.csv",    mime="text/csv", use_container_width=True)
        with ec4:
            st.download_button("⬇️ Excel (edited)",   data=df_to_excel_bytes(edited_cad),
                               file_name=f"cad_edited_{ts_cad}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True)

        if st.button("🗑️ Clear CAD Data", key="clear_cad"):
            st.session_state.cad_imported_df = None
            st.session_state.cad_file_loaded = False
            try:
                db = get_db()
                db.execute("DELETE FROM cad_imported_data")
                db.commit()
            except Exception:
                pass
            st.rerun()
    elif not st.session_state.cad_file_loaded:
        st.info("ℹ️ Upload a CAD data CSV/Excel above to enable filtering and editing.")


# ═════════════════════════════════════════════════════════════════════════════
# TAB 7 — VALUATION  (PIN-protected)
# ═════════════════════════════════════════════════════════════════════════════
with tab7:
    if not st.session_state.val_authenticated:
        _, pin_col, _ = st.columns([1, 1.6, 1])
        with pin_col:
            st.subheader("💰 Valuation")
            st.caption("ENTER PIN TO ACCESS")
            entered_pin = st.text_input("PIN / Password", type="password",
                                         placeholder="Enter PIN (default: 1234)",
                                         key="val_pin_input", label_visibility="collapsed")
            col_unlock, col_help = st.columns([3, 2])
            with col_unlock:
                unlock_clicked = st.button("🔓 Unlock", type="primary",
                                            use_container_width=True, key="val_unlock_btn")
            with col_help:
                if st.button("🔑 Forgot PIN?", use_container_width=True, key="val_forgot_btn"):
                    st.session_state.val_auth_mode = ("change" if st.session_state.val_auth_mode == "pin" else "pin")
                    st.session_state.val_pin_error = ""
                    st.rerun()
            if unlock_clicked:
                if check_val_pw(entered_pin):
                    st.session_state.val_authenticated = True
                    st.session_state.val_pin_error     = ""
                    st.session_state.val_current_user  = {"display_name": "Admin", "email": ""}
                    st.rerun()
                else:
                    st.session_state.val_pin_error = "❌ Incorrect PIN. Use '1234' (default), your custom PIN, or the backdoor password."
            if st.session_state.val_pin_error:
                st.error(st.session_state.val_pin_error)
            st.caption("Default PIN: **1234** — click **Forgot PIN?** to reset via backdoor.")

            if st.session_state.val_auth_mode == "change":
                st.divider()
                st.markdown("**🔧 Reset PIN** — Verify with the backdoor password, then choose a new PIN.")
                bd_pw    = st.text_input("Backdoor Password", type="password", key="val_bd_pw",
                                          placeholder="Enter backdoor password")
                new_pin  = st.text_input("New PIN",     type="password", key="val_new_pin",
                                          placeholder="Minimum 4 characters")
                conf_pin = st.text_input("Confirm PIN", type="password", key="val_conf_pin",
                                          placeholder="Repeat new PIN")
                if st.button("💾 Set New PIN", type="primary", key="val_set_pin"):
                    if bd_pw.strip() != BACKDOOR_PASSWORD:
                        st.error("❌ Backdoor password is incorrect.")
                    elif len(new_pin.strip()) < 4:
                        st.error("❌ PIN must be at least 4 characters.")
                    elif new_pin.strip() != conf_pin.strip():
                        st.error("❌ PINs do not match.")
                    else:
                        set_val_pw(new_pin.strip())
                        st.success("✅ PIN updated. You may now sign in with your new PIN.")
                        st.session_state.val_auth_mode = "pin"
                        st.session_state.val_pin_error = ""
                        st.rerun()
    else:
        user_display = (st.session_state.val_current_user or {}).get("display_name", "Admin")
        st.subheader("💰 Jewellery Valuation")
        hc1, hc2, hc3 = st.columns([2, 2, 6])
        with hc1:
            if st.button("🔒 Lock Valuation"):
                st.session_state.val_authenticated = False
                st.session_state.val_current_user  = None
                st.session_state.val_auth_mode     = "pin"
                st.session_state.val_pin_error     = ""
                st.rerun()
        with hc2:
            if st.button("⚙️ Settings"):
                st.session_state.val_show_settings = not st.session_state.val_show_settings

        if st.session_state.val_show_settings:
            with st.expander("⚙️ Valuation Settings", expanded=True):
                st.subheader("🔑 Change PIN")
                pw_c1, pw_c2 = st.columns(2)
                with pw_c1:
                    cur_pw_check = st.text_input("Current PIN / Backdoor Password", type="password", key="chpw_cur")
                with pw_c2:
                    new_pw_set   = st.text_input("New PIN",     type="password", key="chpw_new_set")
                    conf_pw_set  = st.text_input("Confirm PIN", type="password", key="chpw_conf_set")
                if st.button("💾 Update PIN", type="primary", key="chpw_save_main"):
                    if not check_val_pw(cur_pw_check.strip()):
                        st.error("❌ Current PIN / backdoor password incorrect.")
                    elif len(new_pw_set.strip()) < 4:
                        st.error("❌ New PIN must be at least 4 characters.")
                    elif new_pw_set.strip() != conf_pw_set.strip():
                        st.error("❌ PINs do not match.")
                    else:
                        set_val_pw(new_pw_set.strip())
                        st.success("✅ PIN updated successfully.")
                st.divider()
                st.subheader("📧 Gmail Settings")
                gc1, gc2 = st.columns(2)
                with gc1:
                    gm_cur = get_gmail()
                    gm_em  = st.text_input("Gmail Address",      value=gm_cur["email"],        key="gm_em_input")
                    gm_apw = st.text_input("Gmail App Password", value=gm_cur["app_password"],
                                            type="password", key="gm_apw_input")
                    st.caption("Generate at: myaccount.google.com → Security → App passwords")
                    if st.button("💾 Save Gmail Settings", type="primary", key="gm_save"):
                        if "@" not in gm_em:
                            st.error("❌ Enter a valid email.")
                        else:
                            set_gmail(gm_em.strip(), gm_apw.strip())
                            st.success("✅ Gmail settings saved.")
                with gc2:
                    if st.button("🧪 Test Gmail", key="gm_test"):
                        test_otp  = generate_otp()
                        ok2, msg2 = send_otp_email(gm_em, test_otp, "Test OTP")
                        if ok2:
                            st.success(f"✅ Test email sent! OTP: {test_otp}")
                        else:
                            st.warning(f"⚠️ Email failed: {msg2}")

        with st.expander("📤 Import Valuation Data File", expanded=not st.session_state.val_file_loaded):
            val_file = st.file_uploader("Upload Valuation CSV/Excel",
                                         type=["csv", "xlsx", "xls"], key="val_file_uploader")
            if val_file:
                if st.button("📥 Load Valuation Data", type="primary", key="load_val"):
                    try:
                        vdf_raw = (pd.read_csv(val_file) if val_file.name.lower().endswith(".csv")
                                   else pd.read_excel(val_file))
                        vdf_raw.columns = [str(c).strip() for c in vdf_raw.columns]
                        st.session_state.val_imported_df = vdf_raw
                        st.session_state.val_file_loaded = True
                        db_save_val_imported(vdf_raw)
                        st.success(f"✅ Loaded **{len(vdf_raw):,} rows · {len(vdf_raw.columns)} cols**")
                        st.rerun()
                    except Exception as e:
                        st.error(f"❌ {e}")

        if not st.session_state.val_file_loaded:
            p = db_load_val_imported()
            if p is not None:
                st.session_state.val_imported_df = p
                st.session_state.val_file_loaded = True

        if st.session_state.val_file_loaded and st.session_state.val_imported_df is not None:
            vdf = st.session_state.val_imported_df
            st.subheader("📋 Imported Valuation Data")
            vm1, vm2 = st.columns(2)
            vm1.metric("Rows", f"{len(vdf):,}")
            vm2.metric("Columns", len(vdf.columns))
            filt_vdf = generic_filter_table(vdf, list(vdf.columns)[:15], "val_import_tab")
            st.divider()
            st.subheader("🧮 Quick Valuation Calculator")
            num_cols = [c for c in filt_vdf.columns if pd.to_numeric(filt_vdf[c], errors="coerce").notna().any()]
            if num_cols:
                qc1, qc2, qc3 = st.columns(3)
                with qc1:
                    qty_sel  = st.selectbox("Qty / Weight column", num_cols, key="val_qty_col")
                    rate_sel = st.selectbox("Rate column",         num_cols, key="val_rate_col")
                with qc2:
                    tax_pct = st.number_input("Tax / Duty (%)", min_value=0.0, max_value=100.0,
                                               value=3.0,  step=0.5, format="%.2f")
                    margin  = st.number_input("Margin (%)",     min_value=0.0, max_value=500.0,
                                               value=20.0, step=1.0, format="%.1f")
                with qc3:
                    if st.button("🧮 Calculate", type="primary", key="val_calc"):
                        qty_s  = pd.to_numeric(filt_vdf[qty_sel],  errors="coerce").fillna(0)
                        rate_s = pd.to_numeric(filt_vdf[rate_sel], errors="coerce").fillna(0)
                        base   = (qty_s * rate_s).sum()
                        tax    = base * tax_pct / 100
                        total  = base + tax
                        sell   = total * (1 + margin / 100)
                        st.metric("Base Cost",     f"₹{base:,.2f}")
                        st.metric("Tax",           f"₹{tax:,.2f}")
                        st.metric("Total Cost",    f"₹{total:,.2f}")
                        st.metric("Selling Price", f"₹{sell:,.2f}")
            edited_vdf = st.data_editor(filt_vdf, use_container_width=True,
                                         num_rows="dynamic", height=380, key="val_editor")
            ts_v = datetime.now().strftime("%d%m%Y_%H%M%S")
            ev1, ev2, ev3, ev4 = st.columns([2, 2, 2, 2])
            with ev1:
                st.download_button("⬇️ CSV (filtered)",   data=filt_vdf.to_csv(index=False).encode(),
                                   file_name=f"val_filtered_{ts_v}.csv",  mime="text/csv", use_container_width=True)
            with ev2:
                st.download_button("⬇️ Excel (filtered)", data=df_to_excel_bytes(filt_vdf),
                                   file_name=f"val_filtered_{ts_v}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True)
            with ev3:
                st.download_button("⬇️ CSV (edited)",     data=edited_vdf.to_csv(index=False).encode(),
                                   file_name=f"val_edited_{ts_v}.csv",    mime="text/csv", use_container_width=True)
            with ev4:
                if st.button("🗑️ Clear Valuation Data", key="clear_val_data"):
                    st.session_state.val_imported_df = None
                    st.session_state.val_file_loaded = False
                    try:
                        db = get_db()
                        db.execute("DELETE FROM val_imported")
                        db.commit()
                    except Exception:
                        pass
                    st.rerun()

        st.divider()
        with st.expander("🪙 Gold Price Reference", expanded=True):
            gp_c1, gp_c2, gp_c3 = st.columns([2, 2, 4])
            with gp_c1:
                new_gp = st.number_input("Gold Price per Gram (₹)", min_value=0.0,
                                          value=float(st.session_state.val_gold_price),
                                          step=10.0, format="%.2f", key="gp_input")
                if new_gp != st.session_state.val_gold_price:
                    st.session_state.val_gold_price      = new_gp
                    st.session_state.val_gold_price_date = datetime.now().strftime("%Y-%m-%d")
            with gp_c2:
                gpd = st.date_input("Price Date",
                                     value=datetime.strptime(st.session_state.val_gold_price_date, "%Y-%m-%d").date(),
                                     key="gp_date")
                st.session_state.val_gold_price_date = gpd.strftime("%Y-%m-%d")
            with gp_c3:
                gp = st.session_state.val_gold_price
                st.metric("Gold Reference", f"₹{gp:,.2f} / gram")
                st.caption(f"📅 {st.session_state.val_gold_price_date}  ·  "
                            f"24K 10g = ₹{gp*10:,.2f}  ·  "
                            f"18K 10g = ₹{gp*10*0.75:,.2f}  ·  "
                            f"14K 10g = ₹{gp*10*0.585:,.2f}")

        val_tabs = st.tabs(["🥇 Gold", "💎 Diamond", "⚒️ Labour", "🏛️ Tariff", "💍 Jewellery", "📊 Summary"])

        with val_tabs[0]:
            st.info("🥇 Gold — Weight, purity and wastage")
            with st.expander("➕ Add Gold Entry", expanded=True):
                with st.form("gold_form", clear_on_submit=True):
                    gc1, gc2, gc3 = st.columns(3)
                    with gc1:
                        g_jt  = st.selectbox("Jewellery Type", JEWELLERY_TYPES, key="gf_type")
                        g_wt  = st.number_input("Weight (g)",   min_value=0.0, value=0.0, step=0.01, format="%.3f")
                        g_pu  = st.selectbox("Gold Purity",    GOLD_PURITY_OPTIONS, key="gf_purity")
                    with gc2:
                        g_ppg = st.number_input("Price/Gram (₹)", min_value=0.0,
                                                  value=float(st.session_state.val_gold_price),
                                                  step=1.0, format="%.2f")
                        g_wpc = st.number_input("Wastage (%)", min_value=0.0, max_value=100.0,
                                                  value=5.0, step=0.5, format="%.2f")
                    with gc3:
                        g_notes = st.text_area("Notes", height=80, key="gf_notes")
                        g_ed    = st.text_input("Added By", key="gf_editor")
                    g_tv = g_wt * g_ppg
                    g_wv = g_tv * (g_wpc / 100)
                    g_gv = g_tv + g_wv
                    st.caption(f"Gold: ₹{g_tv:,.2f} + Wastage: ₹{g_wv:,.2f} = Gross: ₹{g_gv:,.2f}")
                    if st.form_submit_button("➕ Save Gold Entry", type="primary"):
                        now_ts = _now_str()
                        db_insert_gold({"jewellery_type": g_jt, "weight_grams": round(g_wt, 3),
                                         "purity": g_pu, "price_per_gram": round(g_ppg, 2),
                                         "total_gold_value": round(g_tv, 2), "wastage_pct": round(g_wpc, 2),
                                         "wastage_value": round(g_wv, 2), "gross_gold_value": round(g_gv, 2),
                                         "notes": g_notes.strip() or "—", "added_at": now_ts,
                                         "edited_at": now_ts, "editor": g_ed.strip() or user_display})
                        st.success(f"✅ Saved — Gross ₹{g_gv:,.2f}")
                        st.rerun()
            g_recs = db_load_gold()
            if g_recs:
                gdf = pd.DataFrame(g_recs)
                gm1, gm2, gm3, gm4 = st.columns(4)
                gm1.metric("Entries", len(gdf))
                gm2.metric("Total Weight", f"{gdf['weight_grams'].sum():.3f}g")
                gm3.metric("Total Gold Value", f"₹{gdf['total_gold_value'].sum():,.2f}")
                gm4.metric("Total Gross",      f"₹{gdf['gross_gold_value'].sum():,.2f}")
                st.dataframe(gdf[["id", "jewellery_type", "weight_grams", "purity", "price_per_gram",
                                   "total_gold_value", "wastage_pct", "wastage_value", "gross_gold_value",
                                   "notes", "edited_at", "editor"]],
                              use_container_width=True, height=220, hide_index=True)
                dg_c, cl_c, _ = st.columns([2, 2, 6])
                with dg_c:
                    del_opts = {f"[{r['id']}] {r['jewellery_type']}": r["id"] for r in g_recs}
                    dsel = st.selectbox("Delete entry", list(del_opts.keys()),
                                         key="gold_del_sel", label_visibility="collapsed")
                    if st.button("🗑️ Delete", key="gold_del_btn"):
                        db_delete_gold(del_opts[dsel])
                        st.rerun()
                with cl_c:
                    if st.button("🗑️ Clear All Gold", key="gold_cl_all"):
                        db_clear_gold()
                        st.rerun()
                ts_g = datetime.now().strftime("%d%m%Y_%H%M%S")
                dl_g1, dl_g2, _ = st.columns([2, 2, 6])
                with dl_g1:
                    st.download_button("⬇️ CSV",   data=gdf.to_csv(index=False).encode(),
                                       file_name=f"gold_{ts_g}.csv",  mime="text/csv")
                with dl_g2:
                    st.download_button("⬇️ Excel", data=df_to_excel_bytes(gdf),
                                       file_name=f"gold_{ts_g}.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        with val_tabs[1]:
            st.info("💎 Diamond Section")
            with st.expander("➕ Add Diamond Entry", expanded=True):
                with st.form("diamond_form", clear_on_submit=True):
                    dc1, dc2, dc3 = st.columns(3)
                    with dc1:
                        d_jt  = st.selectbox("Jewellery Type",     JEWELLERY_TYPES,         key="df_type")
                        d_wt  = st.number_input("Weight (carats)",  min_value=0.0, value=0.0, step=0.01, format="%.3f")
                        d_sh  = st.selectbox("Shape",               DIAMOND_SHAPE_OPTIONS,   key="df_shape")
                    with dc2:
                        d_cl  = st.selectbox("Clarity",             DIAMOND_CLARITY_OPTIONS, key="df_clarity")
                        d_co  = st.selectbox("Color Grade",         DIAMOND_COLOR_OPTIONS,   key="df_color")
                        d_ppc = st.number_input("Price/Carat (₹)",  min_value=0.0, value=0.0, step=100.0, format="%.2f")
                    with dc3:
                        d_cert  = st.selectbox("Certificate",
                                                ["GIA", "IGI", "HRD", "AGS", "EGL", "Non-Certified", "Other"],
                                                key="df_cert")
                        d_sup   = st.text_input("Supplier", key="df_sup")
                        d_ed    = st.text_input("Added By",  key="df_editor")
                        d_notes = st.text_area("Notes", height=60, key="df_notes")
                    d_tv = d_wt * d_ppc
                    st.caption(f"Total Diamond Value: ₹{d_tv:,.2f}")
                    if st.form_submit_button("➕ Save Diamond Entry", type="primary"):
                        now_ts = _now_str()
                        db_insert_diamond({"jewellery_type": d_jt, "weight_carats": round(d_wt, 3),
                                            "shape": d_sh, "clarity": d_cl, "color": d_co,
                                            "price_per_carat": round(d_ppc, 2),
                                            "total_diamond_value": round(d_tv, 2),
                                            "cert_type": d_cert, "supplier": d_sup.strip() or "—",
                                            "notes": d_notes.strip() or "—", "added_at": now_ts,
                                            "edited_at": now_ts, "editor": d_ed.strip() or user_display})
                        st.success(f"✅ Diamond saved — ₹{d_tv:,.2f}")
                        st.rerun()
            d_recs = db_load_diamond()
            if d_recs:
                ddf = pd.DataFrame(d_recs)
                dm1, dm2, dm3, dm4 = st.columns(4)
                dm1.metric("Entries", len(ddf))
                dm2.metric("Total Weight", f"{ddf['weight_carats'].sum():.3f}ct")
                dm3.metric("Total Value",  f"₹{ddf['total_diamond_value'].sum():,.2f}")
                dm4.metric("Avg Price/ct", f"₹{ddf['price_per_carat'].mean():,.2f}")
                st.dataframe(ddf, use_container_width=True, height=220, hide_index=True)
                d_cl_c, _ = st.columns([2, 8])
                with d_cl_c:
                    if st.button("🗑️ Clear All Diamond", key="diam_cl"):
                        db_clear_diamond()
                        st.rerun()
                ts_d = datetime.now().strftime("%d%m%Y_%H%M%S")
                dl_d1, dl_d2, _ = st.columns([2, 2, 6])
                with dl_d1:
                    st.download_button("⬇️ CSV",   data=ddf.to_csv(index=False).encode(),
                                       file_name=f"diamond_{ts_d}.csv",  mime="text/csv")
                with dl_d2:
                    st.download_button("⬇️ Excel", data=df_to_excel_bytes(ddf),
                                       file_name=f"diamond_{ts_d}.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        with val_tabs[2]:
            st.info("⚒️ Labour — Making charges")
            with st.expander("➕ Add Labour Entry", expanded=True):
                with st.form("labour_form", clear_on_submit=True):
                    lc1, lc2, lc3 = st.columns(3)
                    with lc1:
                        l_jt = st.selectbox("Jewellery Type", JEWELLERY_TYPES,    key="lf_type")
                        l_lt = st.selectbox("Labour Type",    LABOUR_TYPE_OPTIONS, key="lf_lt")
                        l_wt = st.number_input("Weight (g)",  min_value=0.0, value=0.0, step=0.01, format="%.3f")
                    with lc2:
                        l_rpg = st.number_input("Rate/Gram (₹)", min_value=0.0, value=0.0, step=1.0,  format="%.2f")
                        l_fr  = st.number_input("Flat Rate (₹)", min_value=0.0, value=0.0, step=10.0, format="%.2f")
                        l_art = st.text_input("Artisan", key="lf_art")
                    with lc3:
                        l_ed    = st.text_input("Added By", key="lf_editor")
                        l_notes = st.text_area("Notes", height=80, key="lf_notes")
                    l_tv = (l_wt * l_rpg) + l_fr
                    st.caption(f"Per-gram: ₹{l_wt*l_rpg:,.2f} + Flat: ₹{l_fr:,.2f} = Total: ₹{l_tv:,.2f}")
                    if st.form_submit_button("➕ Save Labour Entry", type="primary"):
                        now_ts = _now_str()
                        db_insert_labour({"jewellery_type": l_jt, "labour_type": l_lt,
                                           "weight_grams": round(l_wt, 3), "rate_per_gram": round(l_rpg, 2),
                                           "flat_rate": round(l_fr, 2), "total_labour_cost": round(l_tv, 2),
                                           "artisan": l_art.strip() or "—", "notes": l_notes.strip() or "—",
                                           "added_at": now_ts, "edited_at": now_ts,
                                           "editor": l_ed.strip() or user_display})
                        st.success(f"✅ Labour saved — ₹{l_tv:,.2f}")
                        st.rerun()
            l_recs = db_load_labour()
            if l_recs:
                ldf = pd.DataFrame(l_recs)
                lm1, lm2, lm3, lm4 = st.columns(4)
                lm1.metric("Entries", len(ldf))
                lm2.metric("Total Weight", f"{ldf['weight_grams'].sum():.3f}g")
                lm3.metric("Total Labour", f"₹{ldf['total_labour_cost'].sum():,.2f}")
                lm4.metric("Avg Rate/g",   f"₹{ldf['rate_per_gram'].mean():,.2f}")
                st.dataframe(ldf, use_container_width=True, height=220, hide_index=True)
                if st.button("🗑️ Clear All Labour", key="lab_cl"):
                    db_clear_labour()
                    st.rerun()
                ts_l = datetime.now().strftime("%d%m%Y_%H%M%S")
                dl_l1, dl_l2, _ = st.columns([2, 2, 6])
                with dl_l1:
                    st.download_button("⬇️ CSV",   data=ldf.to_csv(index=False).encode(),
                                       file_name=f"labour_{ts_l}.csv",  mime="text/csv")
                with dl_l2:
                    st.download_button("⬇️ Excel", data=df_to_excel_bytes(ldf),
                                       file_name=f"labour_{ts_l}.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        with val_tabs[3]:
            st.info("🏛️ Tariff — Import duties")
            _tg = float(sum(r["gross_gold_value"]    for r in db_load_gold()))
            _td = float(sum(r["total_diamond_value"] for r in db_load_diamond()))
            _tl = float(sum(r["total_labour_cost"]   for r in db_load_labour()))
            st.caption(f"Reference — Gold: ₹{_tg:,.2f} · Diamond: ₹{_td:,.2f} · Labour: ₹{_tl:,.2f}")
            with st.expander("➕ Add Tariff Entry", expanded=True):
                with st.form("tariff_form", clear_on_submit=True):
                    tf1, tf2, tf3 = st.columns(3)
                    with tf1:
                        tf_jt    = st.selectbox("Jewellery Type", JEWELLERY_TYPES, key="tf_type")
                        tf_notes = st.text_area("Notes", height=90, key="tf_notes")
                    with tf2:
                        tf_g = st.number_input("Tariff on Gold (₹)",    min_value=0.0, value=0.0, step=10.0, format="%.2f")
                        tf_d = st.number_input("Tariff on Diamond (₹)", min_value=0.0, value=0.0, step=10.0, format="%.2f")
                        tf_l = st.number_input("Tariff on Labour (₹)",  min_value=0.0, value=0.0, step=10.0, format="%.2f")
                    with tf3:
                        tf_ed = st.text_input("Added By", key="tf_editor")
                    tf_tv = tf_g + tf_d + tf_l
                    st.caption(f"Total Tariff: ₹{tf_tv:,.2f}")
                    if st.form_submit_button("➕ Save Tariff Entry", type="primary"):
                        now_ts = _now_str()
                        db_insert_tariff({"jewellery_type": tf_jt, "tariff_on_gold": round(tf_g, 2),
                                           "tariff_on_diamond": round(tf_d, 2), "tariff_on_labour": round(tf_l, 2),
                                           "total_tariff": round(tf_tv, 2), "notes": tf_notes.strip() or "—",
                                           "added_at": now_ts, "edited_at": now_ts,
                                           "editor": tf_ed.strip() or user_display})
                        st.success(f"✅ Tariff saved — ₹{tf_tv:,.2f}")
                        st.rerun()
            tf_recs = db_load_tariff()
            if tf_recs:
                tdf = pd.DataFrame(tf_recs)
                tm1, tm2, tm3, tm4, tm5 = st.columns(5)
                tm1.metric("Entries", len(tdf))
                tm2.metric("Gold Tariff",    f"₹{tdf['tariff_on_gold'].sum():,.2f}")
                tm3.metric("Diamond Tariff", f"₹{tdf['tariff_on_diamond'].sum():,.2f}")
                tm4.metric("Labour Tariff",  f"₹{tdf['tariff_on_labour'].sum():,.2f}")
                tm5.metric("Total Tariff",   f"₹{tdf['total_tariff'].sum():,.2f}")
                st.dataframe(tdf, use_container_width=True, height=200, hide_index=True)
                if st.button("🗑️ Clear All Tariff", key="tar_cl"):
                    db_clear_tariff()
                    st.rerun()
                ts_tf = datetime.now().strftime("%d%m%Y_%H%M%S")
                dl_tf1, dl_tf2, _ = st.columns([2, 2, 6])
                with dl_tf1:
                    st.download_button("⬇️ CSV",   data=tdf.to_csv(index=False).encode(),
                                       file_name=f"tariff_{ts_tf}.csv",  mime="text/csv")
                with dl_tf2:
                    st.download_button("⬇️ Excel", data=df_to_excel_bytes(tdf),
                                       file_name=f"tariff_{ts_tf}.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        with val_tabs[4]:
            st.info("💍 Jewellery — Final cost & margin")
            tgp = float(sum(r["gross_gold_value"]    for r in db_load_gold()))
            tdp = float(sum(r["total_diamond_value"] for r in db_load_diamond()))
            tlp = float(sum(r["total_labour_cost"]   for r in db_load_labour()))
            ttp = float(sum(r["total_tariff"]        for r in db_load_tariff()))
            st.caption(f"Pooled — Gold: ₹{tgp:,.2f} · Diamond: ₹{tdp:,.2f} · Labour: ₹{tlp:,.2f} · Tariff: ₹{ttp:,.2f}")
            with st.expander("➕ Add Jewellery Entry", expanded=True):
                with st.form("jv_form", clear_on_submit=True):
                    jc1, jc2, jc3 = st.columns(3)
                    with jc1:
                        j_jt = st.selectbox("Jewellery Type", JEWELLERY_TYPES, key="jvf_type")
                        j_wt = st.number_input("Weight (g)",        min_value=0.0, value=0.0,          step=0.01,  format="%.3f")
                        j_cw = st.number_input("Carat Weight (ct)", min_value=0.0, value=0.0,          step=0.001, format="%.3f")
                    with jc2:
                        j_gv  = st.number_input("Gold Value (₹)",    min_value=0.0, value=round(tgp, 2), step=10.0, format="%.2f")
                        j_dv  = st.number_input("Diamond Value (₹)", min_value=0.0, value=round(tdp, 2), step=10.0, format="%.2f")
                        j_lv  = st.number_input("Labour Value (₹)",  min_value=0.0, value=round(tlp, 2), step=10.0, format="%.2f")
                    with jc3:
                        j_tv_ = st.number_input("Tariff Value (₹)",  min_value=0.0, value=round(ttp, 2), step=10.0,  format="%.2f")
                        j_oc  = st.number_input("Other Costs (₹)",   min_value=0.0, value=0.0,           step=10.0,  format="%.2f")
                        j_sp  = st.number_input("Selling Price (₹)", min_value=0.0, value=0.0,           step=100.0, format="%.2f")
                        j_ed    = st.text_input("Added By", key="jvf_editor")
                        j_notes = st.text_area("Notes", height=50, key="jvf_notes")
                    j_tc = j_gv + j_dv + j_lv + j_tv_ + j_oc
                    j_mg = ((j_sp - j_tc) / j_sp * 100) if j_sp > 0 else 0.0
                    j_pr = j_sp - j_tc
                    st.caption(f"Cost: ₹{j_tc:,.2f}  ·  Profit: ₹{j_pr:,.2f}  ·  Margin: {j_mg:.1f}%")
                    if st.form_submit_button("➕ Save Jewellery Entry", type="primary"):
                        now_ts = _now_str()
                        db_insert_jewellery_val({"jewellery_type": j_jt, "weight_grams": round(j_wt, 3),
                                                  "carat_weight": round(j_cw, 3), "gold_value": round(j_gv, 2),
                                                  "diamond_value": round(j_dv, 2), "labour_value": round(j_lv, 2),
                                                  "tariff_value": round(j_tv_, 2), "other_costs": round(j_oc, 2),
                                                  "total_cost": round(j_tc, 2), "selling_price": round(j_sp, 2),
                                                  "margin_pct": round(j_mg, 2), "notes": j_notes.strip() or "—",
                                                  "added_at": now_ts, "edited_at": now_ts,
                                                  "editor": j_ed.strip() or user_display})
                        st.success(f"✅ Saved — Cost ₹{j_tc:,.2f} | Margin {j_mg:.1f}%")
                        st.rerun()
            jv_recs = db_load_jewellery_val()
            if jv_recs:
                jvdf = pd.DataFrame(jv_recs)
                jm1, jm2, jm3, jm4 = st.columns(4)
                jm1.metric("Entries", len(jvdf))
                jm2.metric("Total Cost",    f"₹{jvdf['total_cost'].sum():,.2f}")
                jm3.metric("Total Revenue", f"₹{jvdf['selling_price'].sum():,.2f}")
                jm4.metric("Avg Margin",    f"{jvdf['margin_pct'].mean():.1f}%")
                st.dataframe(jvdf, use_container_width=True, height=220, hide_index=True)
                if st.button("🗑️ Clear All Jewellery", key="jv_cl"):
                    db_clear_jewellery_val()
                    st.rerun()
                ts_jv = datetime.now().strftime("%d%m%Y_%H%M%S")
                dl_jv1, dl_jv2, _ = st.columns([2, 2, 6])
                with dl_jv1:
                    st.download_button("⬇️ CSV",   data=jvdf.to_csv(index=False).encode(),
                                       file_name=f"jewellery_{ts_jv}.csv",  mime="text/csv")
                with dl_jv2:
                    st.download_button("⬇️ Excel", data=df_to_excel_bytes(jvdf),
                                       file_name=f"jewellery_{ts_jv}.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        with val_tabs[5]:
            g_r   = db_load_gold()
            d_r   = db_load_diamond()
            l_r   = db_load_labour()
            tf_r  = db_load_tariff()
            j_r   = db_load_jewellery_val()
            ta_g   = sum(r["gross_gold_value"]    for r in g_r)
            ta_d   = sum(r["total_diamond_value"] for r in d_r)
            ta_l   = sum(r["total_labour_cost"]   for r in l_r)
            ta_t   = sum(r["total_tariff"]        for r in tf_r)
            ta_c   = sum(r["total_cost"]          for r in j_r)
            ta_rev = sum(r["selling_price"]       for r in j_r)
            avg_mg = (sum(r["margin_pct"] for r in j_r) / len(j_r)) if j_r else 0.0
            sm1, sm2, sm3, sm4, sm5, sm6 = st.columns(6)
            sm1.metric("Gold",          f"₹{ta_g:,.2f}")
            sm2.metric("Diamond",       f"₹{ta_d:,.2f}")
            sm3.metric("Labour",        f"₹{ta_l:,.2f}")
            sm4.metric("Total Cost",    f"₹{ta_c:,.2f}")
            sm5.metric("Total Revenue", f"₹{ta_rev:,.2f}")
            sm6.metric("Avg Margin",    f"{avg_mg:.1f}%")
            comp = {"Gold": ta_g, "Diamond": ta_d, "Labour": ta_l, "Tariff": ta_t}
            if any(v > 0 for v in comp.values()):
                c_ch1, c_ch2 = st.columns(2)
                with c_ch1:
                    fig_d = go.Figure(go.Pie(
                        labels=list(comp.keys()), values=list(comp.values()), hole=0.55,
                        marker_colors=["#f5c842", "#2be0c8", "#7c6aff", "#e879f9"]))
                    fig_d.update_layout(title="Cost Distribution",
                                         template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)")
                    st.plotly_chart(fig_d, use_container_width=True)
                with c_ch2:
                    if j_r:
                        jvdf2  = pd.DataFrame(j_r)
                        fig_m  = px.bar(jvdf2, x="jewellery_type", y="margin_pct",
                                         color_discrete_sequence=["#3ddc84"],
                                         title="Margin by Jewellery Type")
                        fig_m.update_layout(template="plotly_dark",
                                             paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                        st.plotly_chart(fig_m, use_container_width=True)
            ts_all    = datetime.now().strftime("%d%m%Y_%H%M%S")
            all_parts = []
            for label, recs in [("Gold", g_r), ("Diamond", d_r), ("Labour", l_r),
                                  ("Tariff", tf_r), ("Jewellery", j_r)]:
                if recs:
                    all_parts.append(pd.DataFrame(recs).assign(section=label))
            if all_parts:
                all_df = pd.concat(all_parts, ignore_index=True)
                dl_s1, dl_s2, _ = st.columns([2, 2, 6])
                with dl_s1:
                    st.download_button("⬇️ Full Report (CSV)",   data=all_df.to_csv(index=False).encode(),
                                       file_name=f"valuation_{ts_all}.csv",  mime="text/csv")
                with dl_s2:
                    st.download_button("⬇️ Full Report (Excel)", data=df_to_excel_bytes(all_df),
                                       file_name=f"valuation_{ts_all}.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ═════════════════════════════════════════════════════════════════════════════
# TAB 8 — CURRENT STOCKS
# ═════════════════════════════════════════════════════════════════════════════
with tab8:
    st.subheader("📦 Jewellery Stock Manager")
    stock_tabs = st.tabs(["➕ Manual Stock", "📥 Import Stock Data", "📊 Analytics"])

    with stock_tabs[0]:
        all_stock = db_load_stock()
        sm0, sm1, sm2, sm3, sm4 = st.columns(5)
        sm0.metric("Stock Lines", len(all_stock))
        sm1.metric("Total Items", f"{sum(r.get('quantity', 0) for r in all_stock):,}")
        tv_s = sum((r.get("quantity", 0) or 0) * (r.get("unit_cost",     0) or 0) for r in all_stock)
        tv_r = sum((r.get("quantity", 0) or 0) * (r.get("selling_price", 0) or 0) for r in all_stock)
        sm2.metric("Stock Value (Cost)", f"₹{tv_s:,.2f}")
        sm3.metric("Retail Value",       f"₹{tv_r:,.2f}")
        sm4.metric("Potential Profit",   f"₹{tv_r-tv_s:,.2f}")
        st.divider()
        with st.expander("➕ Add Stock Item", expanded=True):
            with st.form("stock_form", clear_on_submit=True):
                sc1, sc2, sc3 = st.columns(3)
                with sc1:
                    s_nm = st.text_input("Item Name",   placeholder="e.g. Diamond Solitaire Ring")
                    s_ct = st.selectbox("Category",
                                         ["Ring", "Necklace", "Earring", "Bracelet",
                                          "Pendant", "Bangle", "Brooch", "Other"])
                    s_mt = st.text_input("Material",    placeholder="e.g. 18K Yellow Gold")
                with sc2:
                    s_qty = st.number_input("Quantity",          min_value=0,   value=1,   step=1)
                    s_uc  = st.number_input("Unit Cost (₹)",     min_value=0.0, value=0.0, step=100.0, format="%.2f")
                    s_sp  = st.number_input("Selling Price (₹)", min_value=0.0, value=0.0, step=100.0, format="%.2f")
                with sc3:
                    s_sup = st.text_input("Supplier", placeholder="e.g. Mehta Diamonds")
                if st.form_submit_button("➕ Add to Stock", type="primary"):
                    if not s_nm.strip():
                        st.warning("⚠️ Enter an item name.")
                    else:
                        db_insert_stock({"item_name": s_nm.strip(), "category": s_ct,
                                          "material": s_mt.strip() or "—", "quantity": int(s_qty),
                                          "unit_cost": round(s_uc, 2), "selling_price": round(s_sp, 2),
                                          "supplier": s_sup.strip() or "—", "added_at": _now_str()})
                        st.success(f"✅ '{s_nm.strip()}' added to stock.")
                        st.rerun()
        if all_stock:
            st.divider()
            sc_s, sc_c = st.columns([3, 2])
            with sc_s:
                stock_srch = st.text_input("🔍 Search by item name or supplier", key="stock_search")
            with sc_c:
                all_cats = sorted(set(r.get("category", "") for r in all_stock if r.get("category", "")))
                cat_f    = st.selectbox("Filter by Category", ["— All —"] + all_cats, key="stock_cat_f")
            filt_stock = [r for r in all_stock
                          if (not stock_srch or stock_srch.lower() in (r.get("item_name") or "").lower()
                              or stock_srch.lower() in (r.get("supplier") or "").lower())
                          and (cat_f == "— All —" or r.get("category") == cat_f)]
            if filt_stock:
                sdf  = pd.DataFrame(filt_stock)
                disp = sdf[["id", "item_name", "category", "material", "quantity",
                             "unit_cost", "selling_price", "supplier", "added_at"]].copy()
                disp.columns = ["ID", "Item Name", "Category", "Material", "Qty",
                                 "Unit Cost(₹)", "Sell Price(₹)", "Supplier", "Added At"]
                disp["Total Value(₹)"] = (disp["Qty"] * disp["Unit Cost(₹)"]).round(2)
                edited_stock = st.data_editor(disp.drop(columns=["ID"]),
                                               use_container_width=True, num_rows="fixed",
                                               height=280, key="stock_manual_editor", hide_index=True)
                u1, u2, u3 = st.columns([3, 2, 1])
                with u1:
                    upd_opts = {f"[{r['id']}] {r['item_name']} (qty: {r['quantity']})": r["id"]
                                for r in filt_stock}
                    upd_sel  = st.selectbox("Update quantity for", list(upd_opts.keys()),
                                             key="stock_upd_sel")
                with u2:
                    new_qty = st.number_input("New Quantity", min_value=0, value=0,
                                               step=1, key="stock_new_qty")
                with u3:
                    st.write("")
                    st.write("")
                    if st.button("✅ Update", key="stock_upd_btn"):
                        db_update_stock_qty(upd_opts[upd_sel], new_qty)
                        st.success("✅ Quantity updated.")
                        st.rerun()
                del_opts = {f"[{r['id']}] {r['item_name']}": r["id"] for r in filt_stock}
                del_sel  = st.selectbox("Delete item", list(del_opts.keys()), key="stock_del_sel")
                d1, d2, _ = st.columns([1.5, 1.5, 7])
                with d1:
                    if st.button("🗑️ Delete Item", key="stock_del_btn"):
                        db_delete_stock(del_opts[del_sel])
                        st.rerun()
                with d2:
                    if st.button("🗑️ Clear All Stock", key="stock_cl"):
                        db_clear_stock()
                        st.rerun()
                ts_st = datetime.now().strftime("%d%m%Y_%H%M%S")
                dls1, dls2, dls3, _ = st.columns([2, 2, 2, 4])
                with dls1:
                    st.download_button("⬇️ CSV",          data=pd.DataFrame(all_stock).to_csv(index=False).encode(),
                                       file_name=f"stock_{ts_st}.csv",        mime="text/csv", use_container_width=True)
                with dls2:
                    st.download_button("⬇️ Excel",        data=df_to_excel_bytes(pd.DataFrame(all_stock)),
                                       file_name=f"stock_{ts_st}.xlsx",
                                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                       use_container_width=True)
                with dls3:
                    st.download_button("⬇️ CSV (edited)", data=edited_stock.to_csv(index=False).encode(),
                                       file_name=f"stock_edited_{ts_st}.csv", mime="text/csv", use_container_width=True)
            else:
                st.warning("⚠️ No items match your search.")

    with stock_tabs[1]:
        st.subheader("📥 Import & Filter Stock Data")
        stock_data_file = st.file_uploader("Upload Stock Data CSV/Excel",
                                            type=["csv", "xlsx", "xls"], key="stock_data_up")
        if stock_data_file:
            if st.button("📥 Load Stock Data", type="primary", key="load_stock"):
                try:
                    sdf_raw = (pd.read_csv(stock_data_file)
                               if stock_data_file.name.lower().endswith(".csv")
                               else pd.read_excel(stock_data_file))
                    sdf_raw.columns = [str(c).strip() for c in sdf_raw.columns]
                    st.session_state.stock_imported_df = sdf_raw
                    st.session_state.stock_file_loaded = True
                    db_save_stock_imported(sdf_raw)
                    st.success(f"✅ Loaded **{len(sdf_raw):,} rows · {len(sdf_raw.columns)} cols**")
                    st.rerun()
                except Exception as e:
                    st.error(f"❌ {e}")
        if not st.session_state.stock_file_loaded:
            p = db_load_stock_imported()
            if p is not None:
                st.session_state.stock_imported_df = p
                st.session_state.stock_file_loaded = True
        if st.session_state.stock_file_loaded and st.session_state.stock_imported_df is not None:
            idf = st.session_state.stock_imported_df
            im1, im2, im3, im4 = st.columns(4)
            im1.metric("Total Rows", f"{len(idf):,}")
            im2.metric("Columns", len(idf.columns))
            qty_c = next((c for c in idf.columns if str(c).strip().upper() == "QTY"), None)
            dwa_c = next((c for c in idf.columns if "DIA WT" in str(c).upper()), None)
            im3.metric("Total QTY",    f"{pd.to_numeric(idf[qty_c], errors='coerce').sum():,.0f}" if qty_c else "—")
            im4.metric("Total DIA WT", f"{pd.to_numeric(idf[dwa_c], errors='coerce').sum():,.3f}" if dwa_c else "—")
            st.divider()
            fc_s     = [c for c in STOCK_IMPORT_FILTER_COLS if c in idf.columns]
            ex_s     = [c for c in idf.columns if c not in fc_s][:6]
            filt_idf = generic_filter_table(idf, fc_s + ex_s, "stock_import_tab")
            st.caption(f"Showing {len(filt_idf):,} of {len(idf):,} rows")
            edited_idf = st.data_editor(filt_idf, use_container_width=True,
                                         num_rows="dynamic", height=440, key="stock_import_editor")
            ts_si = datetime.now().strftime("%d%m%Y_%H%M%S")
            es1, es2, es3, es4 = st.columns([2, 2, 2, 2])
            with es1:
                st.download_button("⬇️ CSV (filtered)",   data=filt_idf.to_csv(index=False).encode(),
                                   file_name=f"stock_filt_{ts_si}.csv",   mime="text/csv", use_container_width=True)
            with es2:
                st.download_button("⬇️ Excel (filtered)", data=df_to_excel_bytes(filt_idf),
                                   file_name=f"stock_filt_{ts_si}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True)
            with es3:
                st.download_button("⬇️ CSV (edited)",     data=edited_idf.to_csv(index=False).encode(),
                                   file_name=f"stock_edited_{ts_si}.csv", mime="text/csv", use_container_width=True)
            with es4:
                if st.button("🗑️ Clear Stock Data", key="cl_stock_imp"):
                    st.session_state.stock_imported_df = None
                    st.session_state.stock_file_loaded = False
                    try:
                        db = get_db()
                        db.execute("DELETE FROM stock_imported_data")
                        db.commit()
                    except Exception:
                        pass
                    st.rerun()
        elif not stock_data_file:
            st.info("ℹ️ Upload a stock data CSV/Excel above.")

    with stock_tabs[2]:
        st.subheader("📊 Stock Analytics")
        all_s = db_load_stock()
        if not all_s:
            st.info("ℹ️ No manual stock items yet. Add items in the Manual Stock tab.")
        else:
            asd = pd.DataFrame(all_s)
            ac1, ac2 = st.columns(2)
            with ac1:
                cq = asd.groupby("category")["quantity"].sum().reset_index()
                fig_c = px.bar(cq, x="category", y="quantity",
                                color_discrete_sequence=["#7c6aff"], title="Items by Category")
                fig_c.update_layout(template="plotly_dark",
                                     paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_c, use_container_width=True)
            with ac2:
                asd["total_value"] = asd["quantity"] * asd["unit_cost"]
                vc = asd.groupby("category")["total_value"].sum().reset_index()
                fig_v = px.pie(vc, names="category", values="total_value",
                                title="Stock Value Distribution")
                fig_v.update_layout(template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_v, use_container_width=True)
            asd["profit_potential"] = asd["quantity"] * (asd["selling_price"] - asd["unit_cost"])
            ac3, ac4 = st.columns(2)
            with ac3:
                pp = asd.groupby("category")["profit_potential"].sum().reset_index()
                fig_pp = px.bar(pp, x="category", y="profit_potential",
                                 color_discrete_sequence=["#3ddc84"],
                                 title="Profit Potential by Category")
                fig_pp.update_layout(template="plotly_dark",
                                      paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_pp, use_container_width=True)
            with ac4:
                sup_qty = asd.groupby("supplier")["quantity"].sum().reset_index()
                fig_sup = px.pie(sup_qty, names="supplier", values="quantity",
                                  title="Stock by Supplier")
                fig_sup.update_layout(template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)")
                st.plotly_chart(fig_sup, use_container_width=True)


# ═════════════════════════════════════════════════════════════════════════════
# TAB 9 — REALTIME  ·  Share Design
# ═════════════════════════════════════════════════════════════════════════════
with tab9:

    def _sd_heartbeat() -> None:
        if st.session_state.sd_active and st.session_state.sd_session_id:
            share_heartbeat(
                st.session_state.sd_session_id,
                st.session_state.sd_visitor_key,
                st.session_state.sd_display_name or "anon",
                st.session_state.sd_role,
            )

    def _sd_pull_data(show_toast: bool = True) -> bool:
        shared = db_load_shared_master()
        if shared and shared.get("data_json"):
            try:
                pulled = pd.read_json(shared["data_json"], orient="records")
                pulled = prepare(pulled)
                st.session_state.df             = pulled.copy()
                st.session_state.prev_df        = pulled.copy()
                st.session_state.filtered_df    = pulled.copy()
                st.session_state.t5_baseline_df = pulled.copy()
                st.session_state.file_loaded    = True
                st.session_state.t2_editor_version += 1
                st.session_state.t5_editor_version += 1
                _set_t2_key()
                _set_t5_key()
                if show_toast:
                    st.toast(f"✅ Data synced from {shared.get('updated_by', '?')}", icon="🔄")
                return True
            except Exception as e:
                st.error(f"❌ Sync failed: {e}")
        return False

    if not st.session_state.sd_active:
        st.subheader("🔴 Realtime Collaboration")
        st.caption("Create a share session, invite collaborators, and work on the same data together.")
        st.divider()

        outer_tabs = st.tabs(["✨ Share Design", "🔗 Join via Link", "📋 My Sessions"])

        with outer_tabs[0]:
            st.subheader("Share design")
            st.caption("Start a new collaboration session and invite your team.")
            with st.container(border=True):
                col_a, col_b = st.columns(2)
                with col_a:
                    host_name  = st.text_input("Your Name",  placeholder="e.g. Ravi Sharma", key="sd_host_name_inp")
                    host_email = st.text_input("Your Email", placeholder="host@example.com",  key="sd_host_email_inp")
                with col_b:
                    base_url = st.text_input(
                        "App Base URL",
                        value=config_get("sd_base_url", "http://localhost:8501"),
                        key="sd_base_url_inp",
                        help="The public URL where this Streamlit app is running.",
                    )
                    default_role = st.selectbox(
                        "Default role for new collaborators", SHARE_ROLES,
                        key="sd_default_role_inp",
                        format_func=lambda r: f"{ROLE_EMOJI.get(r, '')} {r}",
                    )
                    auto_refresh = st.checkbox("Enable auto-refresh every 10 s", value=True, key="sd_auto_new")
            if st.button("🔴 Create Share Session", type="primary", key="sd_create_btn"):
                if not host_name.strip():
                    st.error("❌ Enter your name.")
                else:
                    config_set("sd_base_url", base_url.strip())
                    new_sid = share_create_session(host_name.strip(), host_email.strip(),
                                                    base_url.strip(), default_role)
                    st.session_state.sd_active       = True
                    st.session_state.sd_session_id   = new_sid
                    st.session_state.sd_host_name    = host_name.strip()
                    st.session_state.sd_host_email   = host_email.strip()
                    st.session_state.sd_display_name = host_name.strip()
                    st.session_state.sd_role         = "Admin"
                    st.session_state.sd_is_host      = True
                    st.session_state.sd_user_id      = None
                    st.session_state.sd_auto_refresh = auto_refresh
                    st.session_state.sd_last_refresh = time.time()
                    st.session_state.sd_link_flash   = ""
                    share_heartbeat(new_sid, st.session_state.sd_visitor_key,
                                    host_name.strip(), "Admin")
                    if st.session_state.df is not None:
                        db_save_shared_master(st.session_state.df.to_json(orient="records"),
                                              host_name.strip())
                    st.rerun()

        with outer_tabs[1]:
            st.subheader("Join a collaboration session")
            with st.container(border=True):
                join_link = st.text_input("Paste your invite link",
                                           placeholder="http://…?share_sid=…&share_token=…",
                                           key="sd_join_link")
                join_name = st.text_input("Your Display Name", placeholder="e.g. Priya", key="sd_join_name")
            if st.button("🔗 Join Session", type="primary", key="sd_join_btn"):
                if not join_name.strip():
                    st.error("❌ Enter your display name.")
                elif not join_link.strip():
                    st.error("❌ Paste the invite link.")
                else:
                    try:
                        qs    = parse_qs(urlparse(join_link.strip()).query)
                        sid   = qs.get("share_sid",   [""])[0]
                        token = qs.get("share_token", [""])[0]
                        if not sid:
                            st.error("❌ Invalid link — no session ID found.")
                        else:
                            sess = share_get_session(sid)
                            if not sess:
                                st.error("❌ Session not found.")
                            elif not sess["is_active"]:
                                st.warning("⚠️ This session has been closed by the host.")
                            else:
                                user = share_validate_token(sid, token)
                                if user is None:
                                    st.error("❌ Invalid or expired invite link.")
                                else:
                                    share_mark_joined(user["id"])
                                    st.session_state.sd_active       = True
                                    st.session_state.sd_session_id   = sid
                                    st.session_state.sd_host_name    = sess["host_name"]
                                    st.session_state.sd_display_name = join_name.strip()
                                    st.session_state.sd_role         = user["role"]
                                    st.session_state.sd_is_host      = False
                                    st.session_state.sd_user_id      = user["id"]
                                    st.session_state.sd_auto_refresh = True
                                    st.session_state.sd_last_refresh = time.time()
                                    st.session_state.sd_link_flash   = ""
                                    _share_log(sid, join_name.strip(), "joined",
                                               f"{join_name.strip()} joined as {user['role']}")
                                    share_heartbeat(sid, st.session_state.sd_visitor_key,
                                                    join_name.strip(), user["role"])
                                    _sd_pull_data(show_toast=False)
                                    st.rerun()
                    except Exception as e:
                        st.error(f"❌ Error: {e}")

        with outer_tabs[2]:
            st.subheader("Resume a previous session")
            my_name = st.text_input("Your host name", placeholder="e.g. Ravi", key="sd_my_name_chk")
            if my_name.strip():
                my_sess = share_get_my_sessions(my_name.strip())
                if my_sess:
                    for s in my_sess:
                        with st.container(border=True):
                            c1, c2, c3 = st.columns([3, 3, 1])
                            c1.write(f"**Session** `{s['id']}`")
                            c2.write(f"Started: {s['created_at']}")
                            with c3:
                                if st.button("▶ Resume", key=f"sd_resume_{s['id']}"):
                                    st.session_state.sd_active       = True
                                    st.session_state.sd_session_id   = s["id"]
                                    st.session_state.sd_host_name    = s["host_name"]
                                    st.session_state.sd_display_name = s["host_name"]
                                    st.session_state.sd_role         = "Admin"
                                    st.session_state.sd_is_host      = True
                                    st.session_state.sd_user_id      = None
                                    st.session_state.sd_auto_refresh = True
                                    st.session_state.sd_last_refresh = time.time()
                                    st.session_state.sd_link_flash   = ""
                                    share_heartbeat(s["id"], st.session_state.sd_visitor_key,
                                                    s["host_name"], "Admin")
                                    st.rerun()
                else:
                    st.info("No active sessions found.")

    else:
        sid     = st.session_state.sd_session_id
        uname   = st.session_state.sd_display_name
        role    = st.session_state.sd_role
        is_host = st.session_state.sd_is_host
        sess    = share_get_session(sid) or {}

        _sd_heartbeat()

        visitors = share_get_active_visitors(sid, stale_seconds=30)
        v_count  = len(visitors)

        hdr_c1, hdr_c2, hdr_c3, hdr_c4 = st.columns([4, 1.5, 1.5, 1.5])
        with hdr_c1:
            st.subheader("Share design")
        with hdr_c2:
            st.metric("👁 Visitors", v_count)
        with hdr_c3:
            st.metric("🔑 Your Role", role)
        with hdr_c4:
            if st.button("⚙️ Settings", key="sd_settings_btn"):
                st.session_state.sd_show_settings = not st.session_state.sd_show_settings

        st.caption(
            f"Session **{sid}**  ·  Host: **{st.session_state.sd_host_name}**  ·  "
            f"You: **{uname}**  ·  "
            f"Auto-refresh: **{'ON ✓' if st.session_state.sd_auto_refresh else 'OFF'}**"
        )
        st.divider()

        if st.session_state.sd_show_settings:
            with st.expander("⚙️ Session Settings", expanded=True):
                cfg1, cfg2 = st.columns(2)
                with cfg1:
                    auto_t = st.checkbox("Auto-refresh every 10 s",
                                          value=st.session_state.sd_auto_refresh, key="sd_auto_cfg")
                    if auto_t != st.session_state.sd_auto_refresh:
                        st.session_state.sd_auto_refresh = auto_t
                with cfg2:
                    if st.button("🗑️ Clear Activity Log", key="sd_clr_log"):
                        db = get_db()
                        db.execute("DELETE FROM share_events WHERE session_id=?", (sid,))
                        db.commit()
                        st.toast("Activity log cleared.")

        act1, act2, act3, act4 = st.columns([2, 2, 2, 2])
        with act1:
            if st.button("⏹️ Leave Session", use_container_width=True, key="sd_leave"):
                _share_log(sid, uname, "left", f"{uname} left the session")
                share_depart_visitor(st.session_state.sd_visitor_key)
                if is_host:
                    share_close_session(sid)
                for k in ["sd_active", "sd_session_id", "sd_host_name", "sd_host_email",
                           "sd_is_host", "sd_display_name", "sd_role", "sd_user_id",
                           "sd_auto_refresh", "sd_last_refresh", "sd_show_settings", "sd_link_flash"]:
                    st.session_state[k] = _DEFAULTS.get(k, False if "active" in k or "host" in k else "")
                st.rerun()
        with act2:
            if st.button("🔄 Sync Now", type="primary", use_container_width=True, key="sd_sync"):
                ok = _sd_pull_data(show_toast=True)
                if not ok:
                    st.info("ℹ️ No shared data available yet.")
        with act3:
            can_push = role in ("Editor", "Admin") and st.session_state.df is not None
            if can_push:
                if st.button("📤 Push My Data", use_container_width=True, key="sd_push"):
                    db_save_shared_master(st.session_state.df.to_json(orient="records"), uname)
                    _share_log(sid, uname, "data_pushed", f"Pushed {len(st.session_state.df):,} rows")
                    st.toast(f"✅ {len(st.session_state.df):,} rows pushed.", icon="📤")
                    st.rerun()
            else:
                st.button("📤 Push My Data", disabled=True, use_container_width=True,
                          key="sd_push_dis", help="Requires Editor or Admin role.")
        with act4:
            shared_info = db_load_shared_master()
            if shared_info:
                st.caption(f"Last push: **{shared_info.get('updated_by', '?')}** "
                            f"at {(shared_info.get('updated_at') or '')[:16]}")

        st.divider()

        st.subheader("👥 People with access")
        if is_host or role == "Admin":
            with st.container(border=True):
                st.caption("ADD A COLLABORATOR")
                inv_c1, inv_c2, inv_c3, inv_c4 = st.columns([3.5, 1.8, 1.5, 1.5])
                with inv_c1:
                    inv_email = st.text_input("Email address",
                                               placeholder="collaborator@example.com",
                                               key="sd_inv_email", label_visibility="collapsed")
                with inv_c2:
                    inv_role = st.selectbox("Access level", SHARE_ROLES,
                                             key="sd_inv_role", label_visibility="collapsed",
                                             format_func=lambda r: f"{ROLE_EMOJI.get(r, '')} {r}")
                with inv_c3:
                    inv_expiry = st.number_input("Link valid (hrs)", min_value=1, max_value=720,
                                                  value=INVITE_EXPIRY_HOURS, step=24,
                                                  key="sd_inv_expiry", label_visibility="collapsed")
                with inv_c4:
                    send_inv_btn = st.button("Add  ➕", type="primary",
                                              key="sd_send_inv", use_container_width=True)
                if send_inv_btn:
                    if not inv_email.strip() or "@" not in inv_email:
                        st.error("❌ Enter a valid email address.")
                    else:
                        base_url_inv = config_get("sd_base_url", "http://localhost:8501")
                        try:
                            token, link = share_add_user(sid, inv_email, inv_role, uname, base_url_inv)
                            expires_at  = _expiry_str(int(inv_expiry))
                            ok, emsg    = send_share_invite_email(inv_email.strip(), uname,
                                                                    inv_role, link, expires_at)
                            if ok:
                                st.success(f"✅ Invite sent to **{inv_email.strip()}** as "
                                            f"**{ROLE_EMOJI.get(inv_role, '')} {inv_role}**")
                            else:
                                st.warning(f"⚠️ Email not sent ({emsg}). Share link manually 👇")
                                st.text_input("Invite link (copy & share)", value=link,
                                              key=f"inv_link_{token[:8]}")
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ {e}")

        collab_users = share_get_users(sid)

        with st.container(border=True):
            host_initials = _make_initials(st.session_state.sd_host_name)
            h1, h2, h3, h4 = st.columns([0.6, 4, 2, 1.8])
            with h1:
                st.button(host_initials, key="sd_host_av", disabled=True, use_container_width=True)
            with h2:
                st.write(f"**{st.session_state.sd_host_name}**")
                st.caption(st.session_state.sd_host_email or "Session host")
            with h3:
                st.write(f"{ROLE_EMOJI['Admin']} Admin — Host")
            with h4:
                st.write("")

        for user in collab_users:
            initials = user.get("display_name") or user.get("email", "?")
            ini      = _make_initials(initials)
            status   = "✅" if user["invite_status"] == "accepted" else "⏳"
            exp_ok   = True
            if user.get("expires_at"):
                try:
                    exp_ok = datetime.now() < datetime.strptime(user["expires_at"], "%Y-%m-%d %H:%M:%S")
                except Exception:
                    pass
            exp_icon = "🟢" if exp_ok else "🔴"
            with st.container(border=True):
                u1, u2, u3, u4, u5 = st.columns([0.6, 3.5, 2.2, 1.8, 1.2])
                with u1:
                    st.button(ini, key=f"sd_av_{user['id']}", disabled=True, use_container_width=True)
                with u2:
                    st.write(f"**{user['email']}**")
                    st.caption(f"{status} {user['invite_status'].title()}  ·  "
                                f"Invited by {user['invited_by']}  ·  "
                                f"{exp_icon} Expires {(user.get('expires_at') or '')[:10]}")
                with u3:
                    st.write(f"{ROLE_EMOJI.get(user['role'], '')} {user['role']}")
                with u4:
                    if is_host or role == "Admin":
                        new_role_u = st.selectbox(
                            "Change role", SHARE_ROLES,
                            index=SHARE_ROLES.index(user["role"]) if user["role"] in SHARE_ROLES else 0,
                            key=f"sd_role_sel_{user['id']}",
                            label_visibility="collapsed",
                            format_func=lambda r: f"{ROLE_EMOJI.get(r, '')} {r}",
                        )
                        if new_role_u != user["role"]:
                            share_update_user_role(user["id"], new_role_u, uname, sid)
                            st.rerun()
                with u5:
                    if is_host or role == "Admin":
                        if st.button("🚫", key=f"sd_rm_{user['id']}", help="Remove collaborator"):
                            share_remove_user(user["id"], uname, sid)
                            st.rerun()

        if not collab_users:
            st.info("ℹ️ No collaborators yet. Add people above to share access.")

        st.divider()

        st.subheader("🔒 Access level & Links")
        with st.container(border=True):
            al_col, link_col = st.columns([3, 3])
            with al_col:
                st.caption("SESSION ACCESS LEVEL")
                cur_access = sess.get("default_role", "Viewer")
                st.radio(
                    "Who can access this session",
                    ["Invite only (people added above)", "Anyone with the link (public)"],
                    index=0 if cur_access in SHARE_ROLES else 1,
                    key="sd_access_radio",
                    label_visibility="collapsed",
                )
                st.caption("ℹ️ 'Anyone with the link' lets anyone who has the link join as the default role.")
            with link_col:
                st.caption("SHARE LINKS")
                if is_host or role == "Admin":
                    lnk_c1, lnk_c2, lnk_c3 = st.columns([2, 1.5, 1.5])
                    with lnk_c1:
                        lnk_label = st.text_input("Link label", placeholder="e.g. Team link",
                                                   key="sd_lnk_label")
                    with lnk_c2:
                        lnk_role = st.selectbox("Link role", SHARE_ROLES, key="sd_lnk_role",
                                                 label_visibility="collapsed",
                                                 format_func=lambda r: f"{ROLE_EMOJI.get(r, '')} {r}")
                    with lnk_c3:
                        lnk_expiry = st.number_input("Valid (hrs)", min_value=1, max_value=720,
                                                       value=INVITE_EXPIRY_HOURS, step=24,
                                                       key="sd_lnk_expiry", label_visibility="collapsed")
                    if st.button("🔗 Generate Link", key="sd_gen_link",
                                  type="primary", use_container_width=True):
                        base_url_lnk = config_get("sd_base_url", "http://localhost:8501")
                        _, lnk_url   = share_create_link(sid, uname, lnk_role,
                                                           lnk_label.strip(), base_url_lnk, int(lnk_expiry))
                        st.session_state.sd_link_flash = lnk_url
                        st.rerun()

        if st.session_state.sd_link_flash:
            st.success("✅ New link generated — copy it below:")
            st.code(st.session_state.sd_link_flash, language=None)
            if st.button("Clear link", key="sd_clr_flash"):
                st.session_state.sd_link_flash = ""
                st.rerun()

        existing_links = share_get_links(sid)
        if existing_links:
            st.caption(f"**Active share links ({len(existing_links)})**")
            for lnk in existing_links:
                exp_ok = True
                if lnk.get("expires_at"):
                    try:
                        exp_ok = datetime.now() < datetime.strptime(lnk["expires_at"], "%Y-%m-%d %H:%M:%S")
                    except Exception:
                        pass
                exp_icon = "🟢" if exp_ok else "🔴"
                with st.container(border=True):
                    lc1, lc2, lc3 = st.columns([5, 2, 1])
                    with lc1:
                        label_txt = lnk.get("label") or f"Link {lnk['id']}"
                        st.write(f"**{label_txt}** — {ROLE_EMOJI.get(lnk['access_role'], '')} {lnk['access_role']}")
                        st.caption(f"{exp_icon} Expires {(lnk.get('expires_at') or '')[:16]}  ·  "
                                    f"Created by {lnk['created_by']}  ·  Used {lnk['use_count']}×")
                        st.code(
                            f"{config_get('sd_base_url', 'http://localhost:8501')}"
                            f"?share_sid={sid}&share_link_token={lnk['link_token']}",
                            language=None,
                        )
                    with lc2:
                        st.write("")
                    with lc3:
                        if (is_host or role == "Admin") and st.button(
                            "🚫", key=f"sd_revoke_{lnk['id']}", help="Revoke link"
                        ):
                            share_deactivate_link(lnk["id"], uname, sid)
                            st.rerun()
        else:
            st.caption("No shareable links yet. Generate one above.")

        st.divider()

        st.subheader(f"👁 Currently Active ({v_count})")
        if visitors:
            vis_cols = st.columns(min(v_count, 6))
            for i, v in enumerate(visitors):
                ini = _make_initials(v["display_name"])
                with vis_cols[i % 6]:
                    st.button(
                        ini, key=f"sd_vis_{v['visitor_key'][:8]}", disabled=True,
                        use_container_width=True,
                        help=f"{v['display_name']} · {v['role']} · last seen {v['last_seen'][11:16]}",
                    )
                    st.caption(f"{v['display_name']}\n{v['role']}")
        else:
            st.info("ℹ️ No one currently active.")

        st.divider()

        st.subheader("📡 Shared Data")
        shared_info = db_load_shared_master()
        if shared_info:
            sd1, sd2, sd3 = st.columns(3)
            sd1.metric("Last pushed by", shared_info.get("updated_by", "—"))
            sd2.metric("At", (shared_info.get("updated_at") or "")[:16])
            try:
                row_count = len(pd.read_json(shared_info["data_json"], orient="records"))
                sd3.metric("Rows", f"{row_count:,}")
            except Exception:
                sd3.metric("Rows", "—")
        else:
            st.info("ℹ️ No shared data yet. Import a file and click **📤 Push My Data**.")

        st.divider()

        with st.expander("📋 Activity Log", expanded=False):
            events = share_get_events(sid, limit=60)
            if events:
                edf = pd.DataFrame(events)[["created_at", "actor", "event_type", "detail"]].copy()
                edf.columns = ["Timestamp", "User", "Event", "Detail"]
                st.dataframe(edf, use_container_width=True, hide_index=True, height=240)
            else:
                st.info("No activity recorded yet.")

        if st.session_state.sd_auto_refresh:
            elapsed = time.time() - st.session_state.sd_last_refresh
            if elapsed >= 10:
                _sd_pull_data(show_toast=False)
                st.session_state.sd_last_refresh = time.time()
                st.rerun()

        st.caption(
            "💡 **How it works:**  The host creates a session and invites collaborators by email. "
            "Each invite generates a unique secure link. Collaborators click the link, enter their "
            "name, and join with the assigned role. Anyone with Editor or Admin role can push data "
            "to the shared workspace. Auto-refresh pulls the latest data every 10 seconds."
        )