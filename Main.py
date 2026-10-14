import streamlit as st
import pandas as pd
import json
import io
import math
import re
import hmac
import threading
import time
import gc
import google.generativeai as genai
import pdfplumber
from google.api_core.exceptions import ResourceExhausted
import gspread
from google.oauth2.service_account import Credentials
import datetime

# --- CONFIG ---
BRANCH_TAB = "ตั้งค่า_สาขา"
MENU_TAB = "ตั้งค่า_เมนู"
BRANCH_HEADER = ["รหัสบนสลิป", "สาขา", "ป้ายบนสลิป (ถ้ามี)"]
MENU_HEADER = ["สาขา", "รหัสสินค้า", "ชื่อเมนู", "ราคา"]

# Only used to fill the two config tabs the first time they are created.
SEED_BRANCHES = {
    "208413": "อโศก",
    "205711": "พระราม 3",
    "206025": "แฟชั่น 3",
    "207033": "แฟชั่น B",
    "990221": "เอสพลานาด"
}
SEED_MENU_FILE = "item_master.json"

# Names this code used earlier for a default store; a sheet row with exactly one of these is renamed once.
LEGACY_STORE_NAMES = {"000020": ["หมูพริ้ว เยาวราช"]}

# Stores added to the sheet automatically if their slip code is not in the branch tab yet.
# Once a store is in the sheet, the sheet is the source of truth and is never overwritten.
DEFAULT_STORES = [
    {
        "slip_id": "000020",
        "name": "JLK หมูพริ้ว",
        "label": "ร้านค้า",
        "menu": [
            ("JL2001", "ข้าวขาหมูมหาชน (อิ่มพอดี)", 50),
            ("JL2002", "ข้าวขาหมู", 70),
            ("JL2003", "ข้าวกะเพราขาหมู", 70),
            ("JL2004", "ข้าวขาหมูลีน", 80),
            ("JL2005", "ข้าวขาหมูรวมพล (เนื้อ คากิ ไส้ ไข่)", 90),
            ("JL2006", "ขาหมูรวมพล (เนื้อ คากิ ไส้ ไข่)", 160),
            ("JL2007", "ข้าวหมูทอด", 65),
            ("JL2008", "ข้าวหมู2สหาย", 80),
            ("JL2009", "ข้าวหมูทอดพริกเกลือ", 65),
            ("JL2010", "เพิ่ม ข้าวเปล่า", 10),
            ("JL2011", "เพิ่ม ไส้พะโล้", 15),
            ("JL2012", "เพิ่ม ไข่พะโล้", 10),
            ("JL2013", "เพิ่ม พิเศษ", 10),
            ("JL2014", "เพิ่มหมูทอด", 15),
            ("JL2015", "เพิ่มกานาฉ่าย", 10),
            ("JL2016", "น้ำเปล่า", 10),
            ("JL2017", "ใส่กล่อง", 5),
        ],
    },
]

CSV_CONFIG = {
    "date_col": "วันที่เปิดบิล",
    "item_col": "ชื่อเมนู",
    "qty_col": "จำนวน",
    "amount_col": "ยอดขายสุทธิ"
}

# --- Google Sheets ---
@st.cache_resource
def get_gspread_client():
    creds_info = json.loads(st.secrets["GCP_SERVICE_ACCOUNT"])
    creds = Credentials.from_service_account_info(creds_info, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    return gspread.authorize(creds)

def get_spreadsheet():
    return get_gspread_client().open_by_key(st.secrets["SHEET_ID"])

def get_google_sheet():
    return get_spreadsheet().worksheet("Data")

@st.cache_resource
def get_seed_lock():
    return threading.Lock()

def get_or_create_tab(spreadsheet, title, header, seed_rows):
    with get_seed_lock():
        try:
            return spreadsheet.worksheet(title)
        except gspread.exceptions.WorksheetNotFound:
            ws = spreadsheet.add_worksheet(title=title, rows=max(len(seed_rows) + 50, 100), cols=len(header))
            ws.update([header] + seed_rows, value_input_option="RAW")
            return ws

def seed_menu_rows():
    with open(SEED_MENU_FILE, "r", encoding="utf-8") as f:
        items = json.load(f)
    return [[branch, item["code"], item["name"], item["price"]]
            for branch, branch_items in items.items()
            for item in branch_items.values()]

def normalize_id(raw):
    value = re.sub(r"\s+", "", str(raw)).upper()
    # Sheets drops leading zeros from numbers typed into a cell, so 000020 and 20 must match.
    return (value.lstrip("0") or "0") if value.isdigit() else value

def parse_price(raw):
    return float(str(raw).replace(",", "").replace("฿", "").strip())

def rename_legacy_store(store, branch_ws, menu_ws, branch_rows):
    """Renames a default store the sheet still holds under an old default name (menu rows first, then the branch)."""
    legacy = LEGACY_STORE_NAMES.get(store["slip_id"], [])
    branch_row = next((i for i, r in enumerate(branch_rows, start=1) if i > 1 and r and len(r) > 1
                       and normalize_id(r[0]) == normalize_id(store["slip_id"]) and r[1].strip() in legacy), None)
    if branch_row is None:
        return
    old_name = branch_rows[branch_row - 1][1].strip()
    menu_cells = [gspread.Cell(i, 1, store["name"])
                  for i, r in enumerate(menu_ws.get_all_values(), start=1)
                  if i > 1 and r and r[0].strip() == old_name]
    if menu_cells:
        menu_ws.update_cells(menu_cells, value_input_option="RAW")
    branch_ws.update_cell(branch_row, 2, store["name"])

def add_default_stores(branch_ws, menu_ws):
    """Appends any DEFAULT_STORES missing from the branch tab. Returns the names added."""
    added = []
    with get_seed_lock():
        branch_rows = branch_ws.get_all_values()
        existing = {normalize_id(r[0]) for r in branch_rows[1:] if r and r[0].strip()}
        for store in DEFAULT_STORES:
            if normalize_id(store["slip_id"]) in existing:
                rename_legacy_store(store, branch_ws, menu_ws, branch_rows)
                continue
            menu_ws.append_rows([[store["name"], code, name, price] for code, name, price in store["menu"]],
                                value_input_option="RAW")
            branch_ws.append_rows([[store["slip_id"], store["name"], store["label"]]], value_input_option="RAW")
            added.append(store["name"])
    return added

@st.cache_data(ttl=300, show_spinner="กำลังโหลดสาขาและเมนูจาก Google Sheet ...")
def load_config():
    """Returns (branch_config {slip_id: branch}, branch_labels {slip_id: label}, master_data {branch: {code: item}}, warnings)."""
    spreadsheet = get_spreadsheet()
    branch_ws = get_or_create_tab(spreadsheet, BRANCH_TAB, BRANCH_HEADER,
                                  [[vid, name, "VID"] for vid, name in SEED_BRANCHES.items()])
    menu_ws = get_or_create_tab(spreadsheet, MENU_TAB, MENU_HEADER, seed_menu_rows())

    warnings = []
    try:
        add_default_stores(branch_ws, menu_ws)
    except Exception as e:
        warnings.append(f"เพิ่มร้านเริ่มต้นลงแท็บ {BRANCH_TAB}/{MENU_TAB} ไม่ได้: {e}")
    branch_config = {}
    branch_labels = {}
    seen_ids = set()
    for i, row in enumerate(branch_ws.get_all_values()[1:], start=2):
        slip_id, name, label = [(row[j].strip() if j < len(row) else "") for j in range(3)]
        if not slip_id and not name:
            continue
        if not slip_id or not name:
            warnings.append(f"แท็บ {BRANCH_TAB} แถว {i}: ต้องมีทั้งรหัสบนสลิปและชื่อสาขา — ข้ามแถวนี้")
        elif normalize_id(slip_id) in seen_ids:
            warnings.append(f"แท็บ {BRANCH_TAB} แถว {i}: รหัส {slip_id} ซ้ำ — ใช้แถวแรกที่เจอ")
        else:
            seen_ids.add(normalize_id(slip_id))
            branch_config[slip_id] = name
            branch_labels[slip_id] = label

    known_branches = set(branch_config.values())
    master_data = {name: {} for name in known_branches}
    for i, row in enumerate(menu_ws.get_all_values()[1:], start=2):
        branch, code, name, price = [(row[j].strip() if j < len(row) else "") for j in range(4)]
        if not any([branch, code, name, price]):
            continue
        if branch not in known_branches:
            warnings.append(f"แท็บ {MENU_TAB} แถว {i}: ไม่มีสาขา \"{branch}\" ในแท็บ {BRANCH_TAB} — ข้ามแถวนี้")
            continue
        if not code:
            warnings.append(f"แท็บ {MENU_TAB} แถว {i}: ไม่มีรหัสสินค้า — ข้ามแถวนี้")
            continue
        try:
            price_value = parse_price(price)
        except ValueError:
            warnings.append(f"แท็บ {MENU_TAB} แถว {i}: ราคา \"{price}\" ไม่ใช่ตัวเลข — ข้ามแถวนี้")
            continue
        if code in master_data[branch]:
            warnings.append(f"แท็บ {MENU_TAB} แถว {i}: รหัส {code} ซ้ำในสาขา {branch} — ใช้แถวแรกที่เจอ")
            continue
        master_data[branch][code] = {"code": code, "name": name, "price": price_value}

    return branch_config, branch_labels, master_data, warnings

@st.cache_resource
def get_vision_semaphore():
    """Shared semaphore (across all sessions) — limits concurrent Gemini Vision API calls to 2."""
    return threading.Semaphore(2)

class TokenBucket:
    """Shared rate limiter — refills continuously so idle periods let bursts through immediately."""
    def __init__(self, rate_per_minute):
        self.rate = rate_per_minute / 60.0
        self.capacity = rate_per_minute
        self.tokens = float(rate_per_minute)
        self.last_refill = time.time()
        self.lock = threading.Lock()

    def wait_for_token(self):
        while True:
            with self.lock:
                now = time.time()
                self.tokens = min(self.capacity, self.tokens + (now - self.last_refill) * self.rate)
                self.last_refill = now
                if self.tokens >= 1:
                    self.tokens -= 1
                    return
                wait_time = (1 - self.tokens) / self.rate
            time.sleep(wait_time)

@st.cache_resource
def get_rate_limiters():
    """Shared per-model rate limiters (across all sessions) — kept just under each model's real RPM quota."""
    return {
        'gemini-2.5-flash': TokenBucket(rate_per_minute=4),        # real limit: 5/min
        'gemini-2.5-flash-lite': TokenBucket(rate_per_minute=8),   # real limit: 10/min
    }

def clean_for_sheets(value):
    if value is None:
        return ""
    if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
        return ""
    return value

# --- AI ---
def analyze_receipts(images, model_version, branch_ids):
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
    preferred_model = 'gemini-2.5-flash' if model_version == "Flash (เน้นแม่นยำ)" else 'gemini-2.5-flash-lite'
    fallback_model = 'gemini-2.5-flash-lite' if preferred_model == 'gemini-2.5-flash' else 'gemini-2.5-flash'
    prompt = f"""
    This document is a sales receipt or sales report (photo or PDF). Identify which branch each
    item belongs to. The document prints one of these branch identifiers (the label shown is the
    text printed next to it, if known):
{branch_ids}
    Put that identifier, copied exactly as listed above, in the "vid" field of every item.
    A report may contain several store sections, each with its own identifier — use the
    identifier of the section the item is listed under.

    Extract every item line whose total amount is > 0. Skip subtotal and grand-total rows
    (e.g. "ยอดรวมตามร้านค้า", "ยอดรวมทั้งสิ้น").

    The document uses one of these layouts:

    Layout A — 2 rows per item:
    Row 1: [Line Number]. [Item Name]
    Row 2: [Item Code]   [Unit Price]   [Quantity]   [Total Amount]
    Example:
    2. ข้าวมันไก่ทอด
    FMFC033-002  60  33 1,980.00
    → code: "FMFC033-002", unit_price: 60.0, qty: 33, total_amount: 1980.0

    Layout B — 1 row per item, columns:
    [Item Code] [Item Name] [Quantity] [Total incl. VAT "ยอดขายรวมทั้งสิ้น"] [VAT] [Before VAT "ก่อน VAT"]
    There is no unit price column: total_amount is the "ยอดขายรวมทั้งสิ้น" column (incl. VAT),
    never the VAT or before-VAT column, and unit_price = total_amount / qty.
    Example:
    JL2002  ข้าวขาหมู  17  1,190.00  77.85  1,112.15
    → code: "JL2002", unit_price: 70.0, qty: 17, total_amount: 1190.0

    Return ONLY JSON list: [{{"vid": "str", "code": "str", "unit_price": float, "qty": int, "total_amount": float}}]
    """
    sem = get_vision_semaphore()
    limiters = get_rate_limiters()
    last_error = None

    for api_model_name in (preferred_model, fallback_model):
        model = genai.GenerativeModel(api_model_name)
        for attempt in range(2):
            try:
                limiters[api_model_name].wait_for_token()
                with sem:
                    response = model.generate_content([prompt] + images)
                text = response.text.strip()
                if not text:
                    raise ValueError("AI returned empty response")
                return json.loads(text.replace("```json", "").replace("```", "").strip())
            except ResourceExhausted as e:
                last_error = e
                break  # quota full on this model — skip straight to the fallback model
            except (ValueError, json.JSONDecodeError) as e:
                last_error = e
                if attempt == 0:
                    time.sleep(2)

    raise ValueError(f"AI ไม่สามารถอ่านสลิปได้ ({last_error}) — ลองใหม่อีกครั้งในอีกสักครู่")


# --- PDF sales report (text-based, no AI) ---
ITEM_CODE_RE = re.compile(r"^[A-Z][A-Z0-9-]*\d[A-Z0-9-]*$")
AMOUNT_RE = re.compile(r"^-?[\d,]+(\.\d+)?$")
DATE_RE = re.compile(r"^(\d{2})/(\d{2})/(\d{4})$")
TIME_RE = re.compile(r"^\d{1,2}:\d{2}(:\d{2})?$")

def pdf_rows(pdf):
    """Yields each visual line of the PDF as a list of words, left to right."""
    for page in pdf.pages:
        row, row_top = [], None
        for w in sorted(page.extract_words(), key=lambda w: (w["top"], w["x0"])):
            if row and w["top"] - row_top > 3:
                yield [x["text"] for x in sorted(row, key=lambda x: x["x0"])]
                row = []
            if not row:
                row_top = w["top"]
            row.append(w)
        if row:
            yield [x["text"] for x in sorted(row, key=lambda x: x["x0"])]

def to_amount(token):
    return float(token.replace(",", ""))

def parse_sales_report_pdf(data):
    """Parses the "รายงานสรุปยอดขาย Item สินค้า แยกตามร้านค้า" PDF.

    Returns (items, report_date) or None when the file is not that report or any store
    section's items don't add up to its printed subtotal — the caller then falls back to AI.
    """
    try:
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            rows = list(pdf_rows(pdf))
    except Exception:
        return None

    items, section, store, report_date = [], [], None, None
    for row in rows:
        dates = [t for t in row if DATE_RE.match(t)]
        if report_date is None and len(dates) == 2 and not any(TIME_RE.match(t) for t in row):
            d1, d2 = (datetime.date(int(y), int(m), int(d)) for d, m, y in (DATE_RE.match(t).groups() for t in dates))
            report_date = d1 if d1 == d2 else (d1, d2)

        if row[0] == "ร้านค้า" and len(row) > 1:
            if section:
                return None  # previous store had no subtotal row to check against
            store = row[1]
            continue

        tail = row[-4:]
        numeric_tail = len(row) >= 5 and all(AMOUNT_RE.match(t) for t in tail)

        if store and ITEM_CODE_RE.match(row[0]) and numeric_tail:
            qty, total, vat, before_vat = (to_amount(t) for t in tail)
            if qty <= 0 or qty != int(qty) or abs(total - vat - before_vat) > 0.02:
                return None
            section.append({"vid": store, "code": row[0], "qty": int(qty),
                            "total_amount": total, "unit_price": total / qty})
        elif row[0].startswith("ยอดรวมตามร้านค้า") and len(row) >= 5 and numeric_tail:
            qty_sum, total_sum = to_amount(tail[0]), to_amount(tail[1])
            if (not section or qty_sum != sum(i["qty"] for i in section)
                    or abs(total_sum - sum(i["total_amount"] for i in section)) > 0.01):
                return None
            items.extend(section)
            section, store = [], None

    if section or not items:
        return None
    return items, report_date

# ============================================================
# PAGE CONFIG & STYLES
# ============================================================
st.set_page_config(page_title="Power One One-Stop", page_icon="⚡", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Sarabun:wght@300;400;500;600;700;800&display=swap');

/* ── Force light mode on all elements (fixes iOS dark mode) ── */
:root { color-scheme: light; }

html, body, [class*="css"] {
    font-family: 'Sarabun', sans-serif !important;
    color-scheme: light;
}

/* ── App background ── */
.stApp {
    background-color: #F1F5F9 !important;
}
.block-container {
    padding: 1.5rem 1.5rem 3rem !important;
    max-width: 900px;
    background-color: #F1F5F9 !important;
}

/* ── Force ALL text dark so it shows on light background ── */
p, span, label, div, li, td, th {
    color: #1E293B !important;
}

/* ── Labels above inputs ── */
[data-testid="stDateInput"] label,
[data-testid="stRadio"] label,
[data-testid="stSelectbox"] label,
.stRadio > label,
.stDateInput > label,
.stSelectbox > label,
.stFileUploader > label {
    color: #1E293B !important;
    font-size: 14px !important;
    font-weight: 600 !important;
}

/* ── Radio option text ── */
[data-testid="stRadio"] label span,
[data-testid="stRadio"] p {
    color: #1E293B !important;
    font-size: 14px !important;
}

/* ── Force inputs to light mode ── */
input, select, textarea {
    color-scheme: light !important;
    background-color: #FFFFFF !important;
    color: #1E293B !important;
    border: 1px solid #CBD5E1 !important;
    border-radius: 8px !important;
}

/* ── Date input wrapper ── */
[data-testid="stDateInput"] > div > div {
    background-color: #FFFFFF !important;
    border: 1px solid #CBD5E1 !important;
    border-radius: 8px !important;
}
[data-testid="stDateInput"] input {
    color: #1E293B !important;
    background-color: #FFFFFF !important;
}

/* ── Selectbox ── */
[data-testid="stSelectbox"] > div > div {
    background-color: #FFFFFF !important;
    border: 1px solid #CBD5E1 !important;
    border-radius: 8px !important;
    color: #1E293B !important;
}

/* ── Hero header ── */
.hero {
    background: linear-gradient(135deg, #1E3A5F 0%, #2563EB 100%);
    border-radius: 16px;
    padding: 24px 28px 20px;
    margin-bottom: 24px;
    box-shadow: 0 4px 24px rgba(37,99,235,0.25);
}
.hero-title {
    font-size: 22px;
    font-weight: 800;
    color: #FFFFFF !important;
    margin: 0 0 4px 0;
}
.hero-sub {
    font-size: 13px;
    color: rgba(255,255,255,0.85) !important;
    margin: 0 0 10px 0;
}
.hero-badge {
    display: inline-block;
    background: #FCD34D;
    color: #1E3A5F !important;
    font-size: 10px;
    font-weight: 800;
    padding: 2px 10px;
    border-radius: 20px;
    letter-spacing: 1.5px;
    text-transform: uppercase;
}

/* ── Section cards ── */
.card {
    background: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 14px;
    padding: 18px 22px 14px;
    margin-bottom: 16px;
    box-shadow: 0 1px 6px rgba(0,0,0,0.06);
}
.card-header {
    display: flex;
    align-items: center;
    gap: 10px;
    margin-bottom: 6px;
}
.step-num {
    background: linear-gradient(135deg, #2563EB, #1D4ED8);
    color: #fff !important;
    font-size: 12px;
    font-weight: 700;
    width: 28px;
    height: 28px;
    border-radius: 50%;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    flex-shrink: 0;
}
.card-title {
    color: #1E293B !important;
    font-size: 15px;
    font-weight: 700;
    margin: 0;
}

/* ── Date info banner ── */
.info-banner {
    background: #EFF6FF;
    border: 1px solid #BFDBFE;
    border-left: 4px solid #3B82F6;
    border-radius: 8px;
    padding: 10px 14px;
    color: #1E40AF !important;
    font-size: 14px;
    margin-top: 12px;
}
.info-banner strong { color: #1E40AF !important; }

/* ── Primary button (scan) ── */
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #2563EB, #1D4ED8) !important;
    color: #fff !important;
    border: none !important;
    border-radius: 10px !important;
    font-size: 15px !important;
    font-weight: 600 !important;
    padding: 0.65rem 1.2rem !important;
    box-shadow: 0 4px 14px rgba(37,99,235,0.35) !important;
    font-family: 'Sarabun', sans-serif !important;
}
.stButton > button[kind="primary"]:hover {
    background: linear-gradient(135deg, #1D4ED8, #1E40AF) !important;
    box-shadow: 0 6px 20px rgba(37,99,235,0.45) !important;
    transform: translateY(-1px) !important;
}

/* ── Login form button ── */
[data-testid="stFormSubmitButton"] button {
    background: linear-gradient(135deg, #2563EB, #1D4ED8) !important;
    color: #fff !important;
    border: none !important;
    border-radius: 10px !important;
    font-size: 15px !important;
    font-weight: 600 !important;
}
[data-testid="stFormSubmitButton"] button p { color: #fff !important; }

/* ── Save button (green) ── */
.save-wrap .stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #059669, #047857) !important;
    box-shadow: 0 4px 14px rgba(5,150,105,0.35) !important;
}
.save-wrap .stButton > button[kind="primary"]:hover {
    background: linear-gradient(135deg, #047857, #065F46) !important;
    box-shadow: 0 6px 20px rgba(5,150,105,0.45) !important;
}

/* ── Secondary button (cancel) ── */
.stButton > button:not([kind="primary"]) {
    color: #374151 !important;
    background: #F9FAFB !important;
    border: 1px solid #D1D5DB !important;
    border-radius: 10px !important;
    font-size: 14px !important;
    font-weight: 500 !important;
}

/* ── Summary card ── */
.summary-box {
    background: #F0FDF4;
    border: 1px solid #BBF7D0;
    border-radius: 12px;
    padding: 16px 20px 12px;
    margin: 16px 0 12px;
}
.summary-box-title {
    color: #166534 !important;
    font-size: 14px;
    font-weight: 700;
    margin-bottom: 12px;
}

/* ── Metric box ── */
[data-testid="metric-container"] {
    background: linear-gradient(135deg, #EFF6FF, #DBEAFE) !important;
    border: 1px solid #BFDBFE !important;
    border-radius: 12px !important;
    padding: 14px 20px !important;
}
[data-testid="stMetricValue"] {
    color: #1E40AF !important;
    font-size: 28px !important;
    font-weight: 800 !important;
}
[data-testid="stMetricLabel"] {
    color: #3B82F6 !important;
    font-size: 13px !important;
    font-weight: 600 !important;
}

/* ── Expander ── */
[data-testid="stExpander"] summary {
    color: #374151 !important;
    font-size: 14px !important;
    font-weight: 600 !important;
}

/* ── Divider ── */
hr { border-color: #E2E8F0 !important; margin: 16px 0 !important; }

/* ── Alerts ── */
[data-testid="stAlert"] { border-radius: 10px !important; font-size: 14px !important; }
[data-testid="stAlert"] p { color: inherit !important; }

/* ── File uploader — force light background ── */
[data-testid="stFileUploader"],
[data-testid="stFileUploader"] > div,
[data-testid="stFileUploader"] section,
[data-testid="stFileUploader"] section > div,
[data-testid="stFileUploaderDropzone"] {
    background-color: #FFFFFF !important;
    color-scheme: light !important;
}
[data-testid="stFileUploader"] section {
    border: 2px dashed #CBD5E1 !important;
    border-radius: 10px !important;
}
[data-testid="stFileUploader"] small,
[data-testid="stFileUploader"] span,
[data-testid="stFileUploader"] p {
    color: #64748B !important;
}
/* Upload button inside file uploader */
[data-testid="stFileUploader"] button,
[data-testid="baseButton-secondary"] {
    background-color: #F1F5F9 !important;
    color: #1E293B !important;
    border: 1px solid #CBD5E1 !important;
    border-radius: 8px !important;
}

/* ── Dataframe ── */
[data-testid="stDataFrame"] { border-radius: 10px; overflow: hidden; }

/* ── Spinner text ── */
[data-testid="stSpinner"] p { color: #374151 !important; }

/* ── Hide Streamlit branding ── */
#MainMenu { visibility: hidden; }
footer { visibility: hidden; }
header { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ── Hero ─────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
    <div class="hero-title">⚡ Power One One-Stop</div>
    <div class="hero-sub">ระบบบันทึกยอดขาย · Food Market Dashboard</div>
    <span class="hero-badge">AI Powered</span>
</div>
""", unsafe_allow_html=True)

# ── Password gate ────────────────────────────────────────────
if not st.session_state.get("authenticated"):
    app_password = st.secrets.get("APP_PASSWORD", "")
    if not app_password:
        st.error("ยังไม่ได้ตั้งรหัสผ่านของแอป — ผู้ดูแลต้องเพิ่ม APP_PASSWORD ใน Streamlit secrets ก่อน")
        st.stop()

    with st.form("login_form"):
        entered = st.text_input("รหัสผ่าน", type="password")
        submitted = st.form_submit_button("เข้าสู่ระบบ", type="primary", use_container_width=True)

    if submitted:
        if hmac.compare_digest(entered.encode(), app_password.encode()):
            st.session_state["authenticated"] = True
            st.rerun()
        time.sleep(1)  # slows down password guessing
        st.error("รหัสผ่านไม่ถูกต้อง")
    st.stop()

try:
    BRANCH_CONFIG, branch_labels, master_data, config_warnings = load_config()
except Exception as e:
    st.error(f"โหลดสาขาและเมนูจาก Google Sheet ไม่ได้: {e}")
    st.stop()

branch_by_id = {normalize_id(slip_id): name for slip_id, name in BRANCH_CONFIG.items()}
branch_id_prompt = "\n".join(
    f'    - "{slip_id}"' + (f' (labeled "{branch_labels[slip_id]}")' if branch_labels[slip_id] else "")
    for slip_id in BRANCH_CONFIG
)

# ── Step 1: Settings ─────────────────────────────────────────
st.markdown("""
<div class="card">
    <div class="card-header">
        <span class="step-num">1</span>
        <span class="card-title">ตั้งค่าการทำงาน</span>
    </div>
</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    selected_date = st.date_input("วันที่ยอดขาย", datetime.date.today())
    formatted_date_for_sheet = selected_date.strftime("%Y-%m-%d")
with col2:
    ai_choice = st.radio("ขุมพลัง AI", ["Flash (เน้นแม่นยำ)", "Flash Lite (เน้นความเร็ว)"])

st.markdown(f"""
<div class="info-banner">
    📅 ล็อคข้อมูลทั้งหมดเป็นวันที่ <strong>{selected_date.strftime('%d/%m/%Y')}</strong>
    &nbsp;·&nbsp; วันอื่นใน CSV จะถูกตัดออก
</div>
""", unsafe_allow_html=True)

st.markdown("<div style='height:18px'></div>", unsafe_allow_html=True)

# ── Step 2: Upload ───────────────────────────────────────────
st.markdown("""
<div class="card">
    <div class="card-header">
        <span class="step-num">2</span>
        <span class="card-title">นำเข้าสลิปและไฟล์ CSV</span>
    </div>
</div>
""", unsafe_allow_html=True)

files = st.file_uploader("ถ่ายรูปสลิป / ไฟล์ PDF รายงานยอดขาย", type=['jpg', 'jpeg', 'png', 'pdf'], accept_multiple_files=True)
csv_file = st.file_uploader("อัปโหลดไฟล์ CSV (เอสพลานาด)", type=['csv'])

st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

if st.button("🔍  สแกนและตรวจสอบข้อมูล", type="primary", use_container_width=True):
    temp_data = []

    read_notes = []
    if files:
        with st.spinner(f"กำลังสแกนด้วย {ai_choice} ..."):
            try:
                ai_results = []
                for f in files:
                    img_bytes = f.read()
                    is_pdf = f.type == "application/pdf" or f.name.lower().endswith(".pdf")
                    parsed = parse_sales_report_pdf(img_bytes) if is_pdf else None
                    if parsed:
                        results, report_date = parsed
                        read_notes.append(("info", f"📄 {f.name}: อ่านจากข้อความใน PDF โดยตรง ไม่ใช้ AI — {len(results)} รายการ ยอดตรงกับยอดรวมในรายงาน"))
                        if isinstance(report_date, tuple):
                            read_notes.append(("warning", f"{f.name}: รายงานนี้ครอบคลุมหลายวัน ({report_date[0].strftime('%d/%m/%Y')}–{report_date[1].strftime('%d/%m/%Y')}) แต่จะถูกบันทึกเป็นวันที่ {selected_date.strftime('%d/%m/%Y')} ทั้งหมด"))
                        elif report_date and report_date != selected_date:
                            read_notes.append(("warning", f"{f.name}: รายงานเป็นยอดของวันที่ {report_date.strftime('%d/%m/%Y')} แต่เลือกวันที่ไว้ {selected_date.strftime('%d/%m/%Y')} — ตรวจวันที่ก่อนบันทึก"))
                    else:
                        if is_pdf:
                            read_notes.append(("info", f"🤖 {f.name}: อ่านข้อความใน PDF ไม่ได้หรือไม่ใช่รูปแบบรายงานที่รู้จัก — ใช้ AI อ่านแทน"))
                        img_part = {"mime_type": f.type or "image/jpeg", "data": img_bytes}
                        results = analyze_receipts([img_part], ai_choice, branch_id_prompt)
                        del img_part
                    ai_results.extend(results)
                    del img_bytes
                    gc.collect()

                for d in ai_results:
                    branch = branch_by_id.get(normalize_id(d.get('vid', '')), "ไม่ทราบสาขา")
                    extracted_code = str(d.get('code', '')).strip()
                    ai_unit_price = float(d.get('unit_price', 0))
                    ai_qty = int(d.get('qty', 0))
                    ai_total = float(d.get('total_amount', 0))

                    if ai_total > 0:
                        branch_items = master_data.get(branch, {})
                        match = next((item for item in branch_items.values() if item.get('code') == extracted_code), None)

                        if match:
                            db_price = float(match['price'])
                            if (ai_unit_price == db_price) and (ai_qty * ai_unit_price == ai_total):
                                final_qty = ai_qty
                                final_total = ai_total
                                status = "✅ ผ่าน"
                            else:
                                final_qty = round(ai_total / db_price) if db_price > 0 else 0
                                final_total = ai_total
                                status = "⚠️ ปรับยอดจาก Total"

                            if final_qty > 0:
                                temp_data.append({
                                    "วันที่": formatted_date_for_sheet,
                                    "สาขา (จาก CSV)": branch,
                                    "รหัสสินค้า": match['code'],
                                    "ชื่อเมนู": match['name'],
                                    "ราคา": db_price,
                                    "จำนวน": int(final_qty),
                                    "ยอด (฿)": final_total,
                                    "ตรวจสอบ": status
                                })
                        else:
                            temp_data.append({
                                "วันที่": formatted_date_for_sheet,
                                "สาขา (จาก CSV)": branch,
                                "รหัสสินค้า": extracted_code,
                                "ชื่อเมนู": "⚠️ รหัสไม่ตรง Database",
                                "ราคา": ai_unit_price,
                                "จำนวน": ai_qty,
                                "ยอด (฿)": ai_total,
                                "ตรวจสอบ": "❌ ขัดข้อง"
                            })
            except Exception as e:
                st.error(f"เกิดข้อผิดพลาดในการสแกนภาพ: {e}")

    if csv_file:
        try:
            try:
                df_csv = pd.read_csv(csv_file, encoding='utf-8')
            except:
                csv_file.seek(0)
                try:
                    df_csv = pd.read_csv(csv_file, encoding='utf-8-sig')
                except:
                    csv_file.seek(0)
                    df_csv = pd.read_csv(csv_file, encoding='tis-620')

            date_col = CSV_CONFIG["date_col"]
            if date_col in df_csv.columns:
                def parse_date(x):
                    try:
                        if str(x).replace('.', '', 1).isdigit():
                            return pd.to_datetime('1899-12-30') + pd.to_timedelta(float(x), 'D')
                        return pd.to_datetime(x, dayfirst=True)
                    except:
                        return pd.NaT

                df_csv['parsed_date'] = df_csv[date_col].apply(parse_date).dt.date
                df_csv = df_csv[df_csv['parsed_date'] == selected_date]
            else:
                st.warning(f"ไม่พบคอลัมน์ '{date_col}'")

            for _, row in df_csv.iterrows():
                qty = float(row.get(CSV_CONFIG["qty_col"], 0))
                if pd.notna(qty) and qty > 0:
                    raw_item_name = str(row.get(CSV_CONFIG["item_col"], 'ไม่ระบุชื่อ'))
                    item_name = raw_item_name.split('(')[0].strip()
                    item_code = str(row.get('รหัสเมนู', 'ไม่ระบุ'))
                    unit_price = float(row.get('ราคาต่อหน่วย', 0))
                    amount = float(row.get(CSV_CONFIG["amount_col"], row.get('ยอดขาย', 0)))

                    raw_branch = str(row.get('สาขา', 'เอสพลานาด'))
                    if "สยาม" in raw_branch or "ข้าวมันไก่สยาม" in raw_branch:
                        mapped_branch = "เอสพลานาด"
                    else:
                        mapped_branch = raw_branch.strip()

                    temp_data.append({
                        "วันที่": formatted_date_for_sheet,
                        "สาขา (จาก CSV)": mapped_branch,
                        "รหัสสินค้า": item_code,
                        "ชื่อเมนู": item_name,
                        "ราคา": unit_price,
                        "จำนวน": int(qty),
                        "ยอด (฿)": amount,
                        "ตรวจสอบ": "✅ ผ่าน (CSV)"
                    })
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการอ่าน CSV: {e}")

    if temp_data:
        st.session_state['preview_data'] = temp_data
        st.success(f"สแกนสำเร็จ — พบ {len(temp_data)} รายการ ของวันที่ {selected_date.strftime('%d/%m/%Y')}")
    elif csv_file or files:
        st.warning(f"ไม่พบข้อมูลของวันที่ {selected_date.strftime('%d/%m/%Y')}")
    for level, note in read_notes:
        (st.warning if level == "warning" else st.info)(note)

# ── Step 3: Review & Save ────────────────────────────────────
if 'preview_data' in st.session_state and st.session_state['preview_data']:

    st.markdown("""
    <div class="card">
        <div class="card-header">
            <span class="step-num">3</span>
            <span class="card-title">ตรวจสอบและยืนยันข้อมูล</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    df_preview = pd.DataFrame(st.session_state['preview_data'])
    df_edited = st.data_editor(df_preview, use_container_width=True)

    # Summary
    df_summary = (
        df_edited
        .groupby("สาขา (จาก CSV)", as_index=False)
        .agg(จำนวนรายการ=("ชื่อเมนู", "count"), ยอดรวม_บาท=("ยอด (฿)", "sum"))
        .rename(columns={"สาขา (จาก CSV)": "สาขา", "ยอดรวม_บาท": "ยอดรวม (฿)"})
        .sort_values("สาขา")
    )
    df_summary["ยอดรวม (฿)"] = df_summary["ยอดรวม (฿)"].map("{:,.2f}".format)
    total_all = df_edited["ยอด (฿)"].sum()

    st.markdown('<div class="summary-box"><div class="summary-box-title">📊 สรุปยอดขายตามสาขา</div>', unsafe_allow_html=True)
    st.dataframe(df_summary, use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

    st.metric("ยอดรวมทั้งหมด", f"฿{total_all:,.2f}")
    st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)

    st.markdown('<div class="save-wrap">', unsafe_allow_html=True)
    if st.button("✅  ยืนยันและบันทึกลง Google Sheets", type="primary", use_container_width=True):
        try:
            sheet = get_google_sheet()
            raw_rows = df_edited.drop(columns=['ตรวจสอบ']).values.tolist()
            data_to_save = []
            for row in raw_rows:
                cleaned = [clean_for_sheets(v) for v in row]
                cleaned.extend(["", ""])
                data_to_save.append(cleaned)

            sheet.append_rows(data_to_save, value_input_option="USER_ENTERED")
            st.success("บันทึกข้อมูลสำเร็จ — ยอดขายวิ่งเข้าชีตเรียบร้อย 🎉")
            st.balloons()
            st.session_state['preview_data'] = []
        except Exception as e:
            st.error(f"ไม่สามารถบันทึกลงชีตได้: {e}")
    st.markdown('</div>', unsafe_allow_html=True)

# ── Section 4: Delete wrong upload ──────────────────────────
st.markdown("<div style='height:24px'></div>", unsafe_allow_html=True)

with st.expander("🗑️  แก้ไขข้อมูลที่อัปโหลดผิด — ลบตามวันที่และสาขา"):
    st.markdown("""
    <div style='color:#6B7280;font-size:13px;margin-bottom:16px;'>
    เลือกวันที่และสาขาที่ต้องการลบออกจาก Google Sheets — ระบบแสดง preview ก่อนลบจริง
    </div>
    """, unsafe_allow_html=True)

    col_d, col_b = st.columns(2)
    with col_d:
        del_date = st.date_input("วันที่ที่ต้องการลบ", datetime.date.today(), key="del_date")
        del_date_str = del_date.strftime("%Y-%m-%d")
    with col_b:
        branch_options = ["ทุกสาขา"] + list(BRANCH_CONFIG.values())
        del_branch = st.selectbox("สาขา", branch_options, key="del_branch")

    if st.button("🔎  ดูข้อมูลที่จะถูกลบ", key="preview_delete"):
        try:
            sheet = get_google_sheet()
            all_rows = sheet.get_all_values()

            if len(all_rows) < 2:
                st.info("ไม่มีข้อมูลในชีท")
            else:
                header = all_rows[0]
                data_rows = all_rows[1:]
                # คอลัมน์ A=วันที่ (index 0), คอลัมน์ B=สาขา (index 1)
                matched, matched_indices = [], []
                for i, row in enumerate(data_rows):
                    if not row or row[0] != del_date_str:
                        continue
                    row_branch = row[1] if len(row) > 1 else ""
                    if del_branch != "ทุกสาขา" and row_branch != del_branch:
                        continue
                    padded = row + [''] * max(0, len(header) - len(row))
                    matched.append(padded[:len(header)])
                    matched_indices.append(i + 2)

                if matched:
                    st.session_state['delete_rows_indices'] = matched_indices
                    st.session_state['delete_date_str'] = del_date_str
                    st.session_state['delete_branch'] = del_branch
                    df_del = pd.DataFrame(matched, columns=header)
                    label = del_branch if del_branch != "ทุกสาขา" else "ทุกสาขา"
                    st.warning(f"พบ **{len(matched)} แถว** — วันที่ {del_date.strftime('%d/%m/%Y')} · {label}")
                    st.dataframe(df_del, use_container_width=True, hide_index=True)
                else:
                    label = del_branch if del_branch != "ทุกสาขา" else "ทุกสาขา"
                    st.info(f"ไม่พบข้อมูลของวันที่ {del_date.strftime('%d/%m/%Y')} · {label}")
                    st.session_state.pop('delete_rows_indices', None)
        except Exception as e:
            st.error(f"ดึงข้อมูลไม่ได้: {e}")

    # ปุ่มยืนยัน — แสดงเมื่อ filter ตรงกับ preview ที่ค้างอยู่
    indices = st.session_state.get('delete_rows_indices', [])
    same_filter = (
        st.session_state.get('delete_date_str') == del_date_str and
        st.session_state.get('delete_branch') == del_branch
    )
    if indices and same_filter:
        n = len(indices)
        st.markdown(f"<p style='color:#DC2626;font-weight:600;font-size:14px;margin-top:12px;'>⚠️ การลบ {n} แถวนี้ไม่สามารถย้อนกลับได้</p>", unsafe_allow_html=True)
        col_ok, col_no = st.columns(2)
        with col_ok:
            if st.button(f"🗑️  ยืนยันลบ {n} แถว", type="primary", use_container_width=True, key="confirm_delete"):
                try:
                    sheet = get_google_sheet()
                    for row_idx in sorted(indices, reverse=True):
                        sheet.delete_rows(row_idx)
                    st.success(f"ลบ {n} แถวสำเร็จ")
                    st.session_state.pop('delete_rows_indices', None)
                    st.session_state.pop('delete_date_str', None)
                    st.session_state.pop('delete_branch', None)
                except Exception as e:
                    st.error(f"ลบไม่สำเร็จ: {e}")
        with col_no:
            if st.button("ยกเลิก", use_container_width=True, key="cancel_delete"):
                st.session_state.pop('delete_rows_indices', None)
                st.session_state.pop('delete_date_str', None)
                st.session_state.pop('delete_branch', None)
                st.rerun()

# ── Section 5: Branch & menu settings ───────────────────────
if config_warnings:
    st.warning(f"พบปัญหา {len(config_warnings)} จุดในแท็บตั้งค่าสาขา/เมนู — ดูรายละเอียดในหัวข้อ \"สาขาและเมนู\" ด้านล่าง")

with st.expander("⚙️  สาขาและเมนู — แก้ไขได้ใน Google Sheet"):
    st.markdown(f"""
    <div style='color:#6B7280;font-size:13px;margin-bottom:12px;'>
    เพิ่ม/แก้สาขาที่แท็บ <strong>{BRANCH_TAB}</strong> และเมนูที่แท็บ <strong>{MENU_TAB}</strong>
    ชื่อสาขาในสองแท็บต้องสะกดตรงกันทุกตัวอักษร · แอปดึงข้อมูลใหม่ทุก 5 นาที หรือกดปุ่มด้านล่างเพื่อโหลดทันที
    </div>
    """, unsafe_allow_html=True)

    df_branches = pd.DataFrame(
        [{"รหัสบนสลิป": slip_id, "ป้ายบนสลิป": branch_labels[slip_id], "สาขา": name,
          "จำนวนเมนู": len(master_data.get(name, {}))}
         for slip_id, name in BRANCH_CONFIG.items()]
    )
    st.dataframe(df_branches, use_container_width=True, hide_index=True)

    for w in config_warnings:
        st.warning(w)

    if st.button("🔄  โหลดสาขาและเมนูใหม่จากชีท", key="reload_config"):
        load_config.clear()
        st.rerun()
