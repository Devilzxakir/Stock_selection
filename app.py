"""
Stock Analyzer Web App — Flask Backend
Run: python app.py
Visit: http://localhost:5000
"""

import logging
from flask import Flask, request, jsonify, render_template, send_file
import requests as req
import pandas as pd
import math, re, os, io
import time
from datetime import datetime, timedelta
from io import BytesIO

# PDF
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
)
from reportlab.lib.enums import TA_CENTER, TA_RIGHT

# Logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
log = logging.getLogger("stocksense")

app = Flask(__name__)

# ── HARD LIMITS (production) ──────────────────────────────────────────────────
MAX_UPLOAD_BYTES  = 10 * 1024 * 1024          # 10 MB cap on Excel upload
MAX_REMOTE_BYTES  = 25 * 1024 * 1024          # 25 MB cap on Screener.in response
REQUEST_TIMEOUT   = (5, 12)                   # (connect, read) for cheap calls
SLUG_RE           = re.compile(r"^[a-z0-9][a-z0-9\-]{0,79}$")  # validate slugs
TICKER_RE         = re.compile(r"^[A-Z][A-Z0-9\.\-]{0,9}$")    # validate tickers

app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_BYTES

# ── CACHE (in-memory, per-instance) ───────────────────────────────────────────
_ANALYZE_CACHE = {}              # slug -> (ts, result_dict)
_ANALYZE_TTL   = timedelta(minutes=15)
_US_ANALYZE_CACHE = {}           # ticker -> (ts, result_dict)
_US_ANALYZE_TTL = timedelta(minutes=15)
_LIVE_CACHE    = {}              # slug -> (ts, dict)
_LIVE_TTL      = timedelta(minutes=2)
_RATE_BUCKET   = {}              # ip   -> [timestamps]
_RATE_WINDOW   = 60              # seconds
_RATE_MAX      = 30              # requests per window per IP

def _rate_limit():
    """Per-IP fixed window limiter. Returns None if OK, or a 429 response."""
    ip = request.headers.get("X-Forwarded-For", request.remote_addr or "anon").split(",")[0].strip()
    now = time.time()
    bucket = _RATE_BUCKET.setdefault(ip, [])
    cutoff = now - _RATE_WINDOW
    _RATE_BUCKET[ip] = [t for t in bucket if t > cutoff]
    if len(_RATE_BUCKET[ip]) >= _RATE_MAX:
        resp = jsonify({"error": "Too many requests. Please slow down."})
        resp.status_code = 429
        resp.headers["Retry-After"] = str(_RATE_WINDOW)
        return resp
    _RATE_BUCKET[ip].append(now)
    return None

# ── SECURITY HEADERS ──────────────────────────────────────────────────────────
@app.after_request
def _security_headers(resp):
    resp.headers.setdefault("X-Content-Type-Options", "nosniff")
    resp.headers.setdefault("X-Frame-Options", "DENY")
    resp.headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")
    resp.headers.setdefault("Permissions-Policy", "geolocation=(), microphone=(), camera=()")
    if request.endpoint in (None, "index"):
        resp.headers.setdefault(
            "Content-Security-Policy",
            "default-src 'self'; "
            "img-src 'self' data:; "
            "style-src 'self' https://fonts.googleapis.com 'unsafe-inline'; "
            "font-src https://fonts.gstatic.com; "
            "connect-src 'self'; "
            "script-src 'self' 'unsafe-inline'; "
            "frame-ancestors 'none'"
        )
    return resp

# ── ERROR HANDLERS (no trace leaks, no leaked stack to client) ───────────────
@app.errorhandler(404)
def _not_found(_):
    return jsonify({"error": "Not found"}), 404

@app.errorhandler(413)
def _too_large(_):
    return jsonify({"error": f"File too large. Max {MAX_UPLOAD_BYTES // (1024*1024)} MB."}), 413

@app.errorhandler(500)
def _server_error(e):
    log.exception("Unhandled server error")
    return jsonify({"error": "Internal server error"}), 500

# ── CONFIG ────────────────────────────────────────────────────────────────────
# Screener.in session cookie. Required in production; falls back to a dev value
# only when explicitly running locally.
SESSION_ID = os.environ.get("SESSION_ID")
if not SESSION_ID:
    if os.environ.get("VERCEL"):
        # Fail closed on Vercel if not configured
        log.warning("SESSION_ID env var not set on Vercel — Screener scraping will fail.")
        SESSION_ID = ""
    else:
        SESSION_ID = "wd45cahfg2g5q6tqyabmkfw7zjih1bbb"

BASE    = "https://www.screener.in"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    "Accept": "application/json, text/html, */*",
    "Referer": "https://www.screener.in/",
}
COOKIES = {"sessionid": SESSION_ID} if SESSION_ID else {}

# ── Industry classification for 20-point framework ──
GROWING_INDUSTRIES = [
    "ai", "artificial intelligence", "semiconductor", "renewable", "solar", "wind",
    "banking", "pharma", "pharmaceutical", "defense", "defence", "it services",
    "software", "cloud", "data center", "electric vehicle", "ev", "battery",
    "healthcare", "hospital", "fintech", "digital", "e-commerce", "telecom",
    "5g", "electronics", "manufacturing", "infrastructure", "chemical", "green energy"
]
STABLE_INDUSTRIES = [
    "fmcg", "consumer", "food", "beverage", "insurance", "power", "utility",
    "gas", "oil", "metal", "mining", "cement", "construction", "auto", "automobile"
]

# ══════════════════════════════════════════════════════════════════════════════
# SCREENER HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def safe_float(val):
    try:
        v = float(str(val).replace(",", "").strip())
        return None if math.isnan(v) else v
    except (ValueError, TypeError, OverflowError):
        return None

def parse_sheet(xls, sheet_name):
    try:
        df = xls.parse(sheet_name, header=None)
    except (ValueError, KeyError):
        return None, []
    years = [str(c).strip() for c in df.iloc[0, 1:].tolist()]
    data  = {}
    for _, row in df.iloc[1:].iterrows():
        metric = str(row.iloc[0]).strip()
        if not metric or metric == "nan":
            continue
        data[metric] = {}
        for yr, val in zip(years, row.iloc[1:].tolist()):
            data[metric][yr] = safe_float(val)
    return data, years

def parse_data_sheet(df):
    """Parse the Data Sheet into sections. Returns dict of {section: {data: {metric: {year: val}}, years: []}}."""
    sections = {}
    cur_sec = None
    years = []
    col_start = 2

    SECTIONS = {"PROFIT & LOSS": "Profit & Loss", "BALANCE SHEET": "Balance Sheet",
                "CASH FLOW:": "Cash Flow", "QUARTERS": "Quarters"}

    for i in range(len(df)):
        col0 = str(df.iloc[i, 0]).strip()
        col0u = col0.upper()

        if col0u in SECTIONS:
            cur_sec = SECTIONS[col0u]
            years = []
            col_start = 2
            sections[cur_sec] = {"data": {}, "years": [], "col_start": 2}
            continue

        if not cur_sec:
            continue

        col0l = col0.lower()

        if "report date" in col0l:
            years = []
            col_start = 2
            for c in range(2, len(df.columns)):
                y = str(df.iloc[i, c]).strip()
                if y not in ("nan", ""):
                    if not years:
                        col_start = c
                    years.append(y)
            sections[cur_sec]["years"] = years
            sections[cur_sec]["col_start"] = col_start
            continue

        if years and col0 and col0 != "nan":
            # Check if next row is a section header — but still process this row first
            is_last_in_section = (i + 1 < len(df) and
                                  str(df.iloc[i+1, 0]).strip().upper() in SECTIONS)
            has_val = any(str(df.iloc[i, c]).strip() not in ("nan", "") for c in range(1, len(df.columns)))
            if not has_val:
                continue
            result = {}
            for j, c in enumerate(range(col_start, min(len(df.columns), col_start + len(years)))):
                if j < len(years):
                    result[years[j]] = safe_float(df.iloc[i, c])
            sections[cur_sec]["data"][col0] = result
            if is_last_in_section:
                cur_sec = None
                years = []

    # Parse PRICE row
    for i in range(len(df)):
        col0 = str(df.iloc[i, 0]).strip().upper()
        if col0 == "PRICE:":
            bs = sections.get("Balance Sheet", {})
            bs_years = bs.get("years", [])
            bs_start = bs.get("col_start", 2)
            if bs_years:
                price_data = {"PRICE": {}}
                for j, c in enumerate(range(bs_start, min(len(df.columns), bs_start + len(bs_years)))):
                    if j < len(bs_years):
                        val = safe_float(df.iloc[i, c])
                        if val is not None:
                            price_data["PRICE"][bs_years[j]] = val
                if price_data["PRICE"]:
                    sections["Price"] = {"data": price_data, "years": bs_years}

    # Parse META section for current price, shares, etc.
    meta_data = {}
    for i in range(len(df)):
        col0 = str(df.iloc[i, 0]).strip()
        if col0.upper() == "META":
            for j in range(i + 1, min(i + 10, len(df))):
                label = str(df.iloc[j, 0]).strip()
                val = str(df.iloc[j, 1]).strip() if j < len(df) else ""
                if label and label != "nan" and val and val != "nan":
                    meta_data[label] = safe_float(val)
            break
    if meta_data:
        sections["Meta"] = {"data": {"Meta": meta_data}, "years": []}

    return sections

def metric_match(section_data, keywords):
    """Find a metric in section data using substring matching (more robust than exact key lookup)."""
    for name, data in section_data.items():
        name_lower = name.lower()
        if any(kw.lower() in name_lower for kw in keywords):
            return data
    return {}

def compute_ratios(sections):
    """Compute derived ratios and inject them into existing sections."""
    pl = sections.get("Profit & Loss", {}).get("data", {})
    bs = sections.get("Balance Sheet", {}).get("data", {})
    cf = sections.get("Cash Flow", {}).get("data", {})
    pr = sections.get("Price", {}).get("data", {}).get("PRICE", {})
    mt = sections.get("Meta", {}).get("data", {}).get("Meta", {})

    pl_years = sections.get("Profit & Loss", {}).get("years", [])
    bs_years = sections.get("Balance Sheet", {}).get("years", [])
    cf_data = metric_match(cf, ["cash from operating", "operating activity", "cfo"])

    def norm_year(y): return y.strip().replace(",", "").replace(" ", "").lower()
    pl_years_norm = {norm_year(y): y for y in pl_years}
    bs_years_norm = {norm_year(y): y for y in bs_years}
    common_norm = sorted(set(pl_years_norm.keys()) & set(bs_years_norm.keys()))
    if not common_norm:
        # Fallback: just use P&L years if Balance Sheet years are unavailable
        if pl_years:
            common_norm = [norm_year(y) for y in pl_years]
        else:
            return
    # Map back to original P&L year strings (canonical)
    all_years = [pl_years_norm[y] for y in common_norm if y in pl_years_norm]

    sales_d = metric_match(pl, ["sales", "revenue", "net sales"])
    np_d = metric_match(pl, ["net profit", "profit after", "net income"])
    pbt_d = metric_match(pl, ["profit before tax", "pbt", "profit before"])
    dep_d = metric_match(pl, ["depreciation", "dep"])
    int_d = metric_match(pl, ["interest"])
    oi_d = metric_match(pl, ["other income"])
    res_d = metric_match(bs, ["reserves"])
    eq_d = metric_match(bs, ["equity share capital", "equity capital", "share capital"])
    debt_d = metric_match(bs, ["borrowings", "debt", "total debt", "loans"])
    sh_d = metric_match(bs, ["no. of equity shares", "number of equity shares", "equity shares", "no of shares"])
    tot_d = metric_match(bs, ["total"])
    ol_d = metric_match(bs, ["other liabilities", "other liability"])
    cash_d = metric_match(bs, ["cash", "cash & bank", "cash and bank", "cash equivalent"])
    inv_d = metric_match(bs, ["inventory", "inventories", "stock"])
    recv_d = metric_match(bs, ["receivables", "debtors", "trade receivables", "account receivables"])
    meta_price = mt.get("Current Price", None)

    pl_data = sections.get("Profit & Loss", {}).get("data", {})
    key_data = {}
    key_years = list(all_years)

    for yr in all_years:
        s = sales_d.get(yr)
        np_ = np_d.get(yr)
        pbt = pbt_d.get(yr)
        dp = dep_d.get(yr)
        int_ = int_d.get(yr)
        oi_ = oi_d.get(yr)
        res = res_d.get(yr)
        eq = eq_d.get(yr)
        debt = debt_d.get(yr)
        shares = sh_d.get(yr)
        tot = tot_d.get(yr)
        ol = ol_d.get(yr)
        cash = cash_d.get(yr)
        inv = inv_d.get(yr)
        recv = recv_d.get(yr)
        cfo_val = cf_data.get(yr)
        price = meta_price if meta_price is not None else pr.get(yr)

        eq_val = eq if eq is not None else 0
        res_val = res if res is not None else 0
        equity = eq_val + res_val

        # Operating Profit = PBT + Interest + Depreciation - Other Income
        op = None
        if pbt is not None and int_ is not None and dp is not None:
            op = pbt + int_ + dp
            if oi_ is not None:
                op -= oi_
        if op is not None:
            pl_data.setdefault("Operating Profit", {})[yr] = round(op, 2)

        # EBITDA = Operating Profit + Depreciation
        ebitda_ = round(op + dp, 2) if op is not None and dp is not None else None
        if ebitda_ is not None:
            pl_data.setdefault("EBITDA", {})[yr] = ebitda_
            key_data.setdefault("EBITDA", {})[yr] = ebitda_

        # Net Profit Margin
        if np_ is not None and s and s > 0:
            key_data.setdefault("Net Margin", {})[yr] = round((np_ / s) * 100, 1)

        # OPM
        if op is not None and s and s > 0:
            opm_val = round((op / s) * 100, 1)
            pl_data.setdefault("OPM", {})[yr] = opm_val
            key_data.setdefault("OPM", {})[yr] = opm_val

        # EBITDA Margin
        if ebitda_ is not None and s and s > 0:
            key_data.setdefault("EBITDA Margin", {})[yr] = round((ebitda_ / s) * 100, 1)

        # Normalize shares: Screener.in provides actual count (e.g., 100,000,000 for 10Cr shares)
        # If value is small (< 1M), it might be in Crores already — multiply back
        effective_shares = None
        if shares is not None and shares > 0:
            if shares < 1e6:
                effective_shares = shares * 1e7
            else:
                effective_shares = shares

        # EPS
        if np_ is not None and effective_shares and effective_shares > 0:
            eps_val = round(np_ / (effective_shares / 1e7), 2)
            key_data.setdefault("EPS", {})[yr] = eps_val

        # P/E
        eps_val = key_data.get("EPS", {}).get(yr)
        if price is not None and eps_val and eps_val > 0:
            key_data.setdefault("P/E", {})[yr] = round(price / eps_val, 1)

        # ROE
        if np_ is not None and equity > 0:
            key_data.setdefault("ROE", {})[yr] = round((np_ / equity) * 100, 1)

        # ROCE = Operating Profit / Capital Employed (Equity + Total Borrowings)
        if op is not None and (debt is not None or (equity) > 0):
            capital_employed = (equity) + (debt or 0)
            if capital_employed > 0:
                key_data.setdefault("ROCE", {})[yr] = round((op / capital_employed) * 100, 1)

        # D/E
        if debt is not None and equity > 0:
            key_data.setdefault("Debt/Equity", {})[yr] = round(debt / equity, 2)

        # Interest Coverage Ratio
        if op is not None and int_ is not None and int_ > 0:
            key_data.setdefault("Interest Coverage", {})[yr] = round(op / int_, 1)

        # BVPS
        if effective_shares and effective_shares > 0:
            key_data.setdefault("BVPS", {})[yr] = round(equity / (effective_shares / 1e7), 2)

        # Current Ratio (Current Assets / Current Liabilities — approximated as (Cash + Inventory + Receivables) / Borrowings)
        ca = (cash or 0) + (inv or 0) + (recv or 0)
        cl = debt or 0
        if ca > 0 and cl > 0:
            key_data.setdefault("Current Ratio", {})[yr] = round(ca / cl, 1)

        # Free Cash Flow (CFO - Capex approximated)
        cap = cf.get("Cash from Investing Activity", {}).get(yr)
        if cfo_val is not None and cap is not None:
            key_data.setdefault("Free Cash Flow", {})[yr] = round(cfo_val + cap, 2)  # cap is negative

        # Price
        if price is not None:
            key_data.setdefault("Price", {})[yr] = price

    if key_data:
        sections["Key"] = {"data": key_data, "years": key_years}

def _parse_html_table(table):
    """Parse an HTML <table> into {metric: {year: value}}, years list."""
    from bs4 import BeautifulSoup
    rows = table.find_all("tr")
    if not rows:
        return {}, []
    header = [c.get_text(strip=True) for c in rows[0].find_all(["th", "td"])]
    years = [h for h in header[1:] if h]
    data = {}
    for row in rows[1:]:
        cells = [c.get_text(strip=True) for c in row.find_all(["th", "td"])]
        if not cells or not cells[0]:
            continue
        metric = cells[0].rstrip("+")
        series = {}
        for yr, val in zip(years, cells[1:]):
            fv = safe_float(val.replace("%", ""))
            if fv is not None:
                series[yr] = fv
        if series:
            data[metric] = series
    return data, years


def _merge_meta_into_biz(sheets, biz_info):
    """Copy industry from sheets Meta into biz_info if biz_info is missing it."""
    meta = sheets.get("Meta", {}).get("data", {}).get("Meta", {})
    if not biz_info.get("industry") and meta.get("Industry"):
        biz_info["industry"] = meta["Industry"]
    if not biz_info.get("market_cap") and meta.get("Market Cap"):
        biz_info["market_cap"] = meta["Market Cap"]
    return biz_info


def _enrich_with_yahoo(slug, biz_info):
    """Enrich biz_info and return (biz_info, yahoo_data) with Yahoo Finance data."""
    yf_data = scrape_yahoo(slug)
    if not yf_data:
        return biz_info, {}
    if not biz_info.get("industry") and yf_data.get("industry"):
        biz_info["industry"] = yf_data["industry"]
    if not biz_info.get("sector") and yf_data.get("sector"):
        biz_info["sector"] = yf_data["sector"]
    if not biz_info.get("market_cap") and yf_data.get("market_cap_yf"):
        biz_info["market_cap"] = yf_data["market_cap_yf"]
    if not biz_info.get("promoter_holding") and yf_data.get("held_by_insiders"):
        biz_info["promoter_holding"] = yf_data["held_by_insiders"]
    if not biz_info.get("institutional_holding") and yf_data.get("held_by_institutions"):
        biz_info["institutional_holding"] = yf_data["held_by_institutions"]
    return biz_info, yf_data


def scrape_html(slug):
    """Scrape financial data directly from the Screener.in HTML page (no login needed).
    Returns the same sheets dict that parse_excel produces."""
    from bs4 import BeautifulSoup

    url = f"{BASE}/company/{slug}/consolidated/"
    r = req.get(url, headers=HEADERS, timeout=REQUEST_TIMEOUT)
    r.raise_for_status()
    soup = BeautifulSoup(r.text, "html.parser")

    SECTION_MAP = {
        "profit-loss": "Profit & Loss",
        "quarters":    "Quarters",
        "balance-sheet": "Balance Sheet",
        "cash-flow":   "Cash Flow",
        "ratios":      "Ratios",
    }
    sections = {}
    for sec_id, sheet_name in SECTION_MAP.items():
        sec = soup.find("section", id=sec_id)
        if not sec:
            continue
        table = sec.find("table")
        if not table:
            continue
        data, years = _parse_html_table(table)
        if data:
            sections[sheet_name] = {"data": data, "years": years}

    # Parse top-level key ratios using .name/.number selectors (no colon in text)
    meta_data = {}
    for li in soup.select(".company-ratios li"):
        name_el = li.select_one(".name")
        val_el = li.select_one(".value")
        if not name_el or not val_el:
            continue
        key = name_el.get_text(strip=True)
        num_el = val_el.select_one(".number")
        raw_val = num_el.get_text(strip=True) if num_el else val_el.get_text(strip=True)
        raw_val = raw_val.replace("₹", "").replace(",", "").replace("%", "").replace("Cr.", "").strip()
        fv = safe_float(raw_val)
        if fv is not None:
            meta_data[key] = fv

    # Also inject industry and description from page
    industry_el = soup.find("a", href=lambda h: h and "/industry/" in h)
    if industry_el:
        meta_data["Industry"] = industry_el.get_text(strip=True)
    else:
        peer_sec = soup.find("section", id="peers")
        if peer_sec:
            peer_text = peer_sec.get_text(" ", strip=True)
            parts = peer_text.split("Peer comparison")
            if len(parts) > 1:
                chain = parts[1].split("Part of")[0].strip()
                industry_words = chain.split()
                if len(industry_words) >= 4:
                    meta_data["Industry"] = " ".join(industry_words[-4:])
                elif industry_words:
                    meta_data["Industry"] = chain
    desc_el = soup.find("meta", attrs={"name": "description"})
    if desc_el and desc_el.get("content"):
        meta_data["Description"] = desc_el["content"]

    if meta_data:
        sections["Meta"] = {"data": {"Meta": meta_data}, "years": []}

    # Inject a Price section so compute_ratios can find the current price
    current_price = meta_data.get("Current Price")
    if current_price:
        pl_years = sections.get("Profit & Loss", {}).get("years", [])
        if pl_years:
            sections["Price"] = {"data": {"PRICE": {pl_years[-1]: current_price}}, "years": pl_years}
        else:
            sections["Price"] = {"data": {"PRICE": {"current": current_price}}, "years": ["current"]}

    # Inject BVPS into Balance Sheet from Book Value top ratio
    book_value = meta_data.get("Book Value")
    if book_value:
        bs_data = sections.get("Balance Sheet", {}).get("data", {})
        bs_years = sections.get("Balance Sheet", {}).get("years", [])
        if bs_years:
            bv_series = {yr: book_value for yr in bs_years}
            bs_data["Book Value per Share"] = bv_series

    # Estimate share count from Market Cap and Current Price for EPS computation
    mcap = meta_data.get("Market Cap")
    if mcap and current_price and current_price > 0:
        est_shares_cr = mcap / current_price  # in Crores
        bs_data = sections.get("Balance Sheet", {}).get("data", {})
        bs_years = sections.get("Balance Sheet", {}).get("years", [])
        if bs_years and "Book Value per Share" in bs_data:
            # Compute shares = (Equity + Reserves) / BVPS
            eq_cap = bs_data.get("Equity Capital", {})
            reserves = bs_data.get("Reserves", {})
            shares_series = {}
            for yr in bs_years:
                eq = eq_cap.get(yr, 0) or 0
                res = reserves.get(yr, 0) or 0
                bv = bs_data.get("Book Value per Share", {}).get(yr)
                if bv and bv > 0:
                    shares_series[yr] = round((eq + res) / bv, 2)
            if shares_series:
                bs_data["No. of Equity Shares"] = shares_series

    return sections


def parse_excel(excel_bytes):
    xls    = pd.ExcelFile(excel_bytes)
    result = {}
    if "Data Sheet" in xls.sheet_names:
        df = xls.parse("Data Sheet", header=None)
        ds_sections = parse_data_sheet(df)
        compute_ratios(ds_sections)
        result.update(ds_sections)
    else:
        for sheet in xls.sheet_names:
            data, years = parse_sheet(xls, sheet)
            if data:
                result[sheet] = {"data": data, "years": years}
    return result

def get_metric(sheets, sheet_kws, metric_kws):
    for sh_name, sh_val in sheets.items():
        if not any(k.lower() in sh_name.lower() for k in sheet_kws):
            continue
        for m_name, series in sh_val["data"].items():
            if any(k.lower() in m_name.lower() for k in metric_kws):
                clean = {yr: v for yr, v in series.items() if v is not None}
                if clean:
                    return m_name, clean
    return None, {}

def fv(s): return list(s.values())[0]  if s else None
def lv(s): return list(s.values())[-1] if s else None

def normalize_shares(shares):
    """Screener.in provides share count in actual number (e.g., 100M for 10Cr shares).
    If value appears to be in Crores (< 1M), multiply back to actual count."""
    if shares is None or shares <= 0:
        return None
    if shares < 1e6:
        return shares * 1e7
    return shares
def yrs(s):
    k = list(s.keys())
    return len(k) - 1 if len(k) > 1 else 1

def cagr(start, end, n):
    if not start or not end or start <= 0 or end <= 0 or n <= 0:
        return None
    return round(((end / start) ** (1 / n) - 1) * 100, 1)

def score_cagr(c):
    if c is None: return 5
    if c >= 25: return 10
    if c >= 20: return 9
    if c >= 15: return 8
    if c >= 10: return 7
    if c >= 5:  return 6
    return 4

def score_val(v, thresholds):
    """thresholds: [(min_val, score), ...] sorted descending."""
    if v is None: return 5
    for min_v, s in thresholds:
        if v >= min_v: return s
    return 4

# ══════════════════════════════════════════════════════════════════════════════
# ADVANCED ANALYSIS FUNCTIONS
# ══════════════════════════════════════════════════════════════════════════════

def piotroski_f_score(pl, bs, cf, key):
    """Compute Piotroski F-Score (0-9) from available financial data."""
    score = 0
    details = {}
    # 1. Net Income positive
    np_vals = list(pl.get("Net profit", {}).values())
    if np_vals and np_vals[-1] and np_vals[-1] > 0:
        score += 1; details["Net Income > 0"] = True
    else:
        details["Net Income > 0"] = False

    # 2. CFO positive
    cfo_vals = list(cf.get("Cash from Operating Activity", {}).values())
    if cfo_vals and cfo_vals[-1] and cfo_vals[-1] > 0:
        score += 1; details["CFO > 0"] = True
    else:
        details["CFO > 0"] = False

    # 3. ROE increasing (YoY)
    roe_vals = list(key.get("ROE", {}).values())
    if len(roe_vals) >= 2 and roe_vals[-1] and roe_vals[-2] and roe_vals[-1] > roe_vals[-2]:
        score += 1; details["ROE Increasing"] = True
    else:
        details["ROE Increasing"] = False

    # 4. CFO > Net Profit (earnings quality)
    if cfo_vals and np_vals and cfo_vals[-1] and np_vals[-1] and cfo_vals[-1] > np_vals[-1]:
        score += 1; details["CFO > Net Profit"] = True
    else:
        details["CFO > Net Profit"] = False

    # 5. Debt/Equity decreased
    de_vals = list(key.get("Debt/Equity", {}).values())
    if len(de_vals) >= 2 and de_vals[-1] is not None and de_vals[-2] is not None and de_vals[-1] < de_vals[-2]:
        score += 1; details["D/E Decreased"] = True
    else:
        details["D/E Decreased"] = False

    # 6. Current Ratio increased
    cr_vals = list(key.get("Current Ratio", {}).values())
    if len(cr_vals) >= 2 and cr_vals[-1] and cr_vals[-2] and cr_vals[-1] > cr_vals[-2]:
        score += 1; details["Current Ratio Increased"] = True
    else:
        details["Current Ratio Increased"] = False

    # 7. No share dilution (shares count didn't increase)
    shares = pl.get("No. of Equity Shares", {}) or bs.get("No. of Equity Shares", {})
    sh_vals = list(shares.values())
    if len(sh_vals) >= 2 and sh_vals[-1] and sh_vals[-2] and sh_vals[-1] <= sh_vals[-2]:
        score += 1; details["No Share Dilution"] = True
    else:
        details["No Share Dilution"] = False

    # 8. Net Margin increased (approximation for Gross Margin)
    nm_vals = list(key.get("Net Margin", {}).values())
    if len(nm_vals) >= 2 and nm_vals[-1] and nm_vals[-2] and nm_vals[-1] > nm_vals[-2]:
        score += 1; details["Net Margin Increased"] = True
    else:
        details["Net Margin Increased"] = False

    # 9. Asset Turnover increased (Sales/Total Assets)
    ta = bs.get("Total", {}) or bs.get("Total Assets", {})
    sales = pl.get("Sales", {})
    at_years = sorted(set(ta.keys()) & set(sales.keys()))
    at_vals = []
    for yr in at_years:
        s = sales.get(yr)
        t = ta.get(yr)
        if s and t and t > 0:
            at_vals.append(s / t)
    if len(at_vals) >= 2 and at_vals[-1] > at_vals[-2]:
        score += 1; details["Asset Turnover Increased"] = True
    else:
        details["Asset Turnover Increased"] = False

    return score, details


def earnings_quality_analysis(cfo, profit):
    """Compare CFO vs Net Profit to assess earnings quality."""
    if not cfo or not profit:
        return {"quality": "Insufficient data", "ratio": None}
    common = sorted(set(cfo.keys()) & set(profit.keys()))
    ratios = {}
    for yr in common:
        c = cfo.get(yr)
        p = profit.get(yr)
        if c and p and p > 0:
            ratios[yr] = round(c / p, 2)
    if not ratios:
        return {"quality": "Insufficient data", "ratio": None}
    latest = list(ratios.values())[-1]
    if latest > 1.0:
        quality = "Excellent"
    elif latest > 0.7:
        quality = "Good"
    elif latest > 0.3:
        quality = "Fair"
    else:
        quality = "Poor"
    return {"quality": quality, "ratio": ratios, "latest": latest}


def revenue_acceleration(sales):
    """Check if revenue growth is accelerating by comparing last 3 periods."""
    if not sales or len(sales) < 4:
        return {"status": "Insufficient data"}
    vals = list(sales.values())
    # Calc 3 YoY growth rates
    rates = []
    for i in range(1, len(vals)):
        if vals[i-1] and vals[i-1] > 0:
            rates.append((vals[i] - vals[i-1]) / vals[i-1] * 100)
    if len(rates) < 3:
        return {"status": "Insufficient data"}
    r1, r2, r3 = rates[-3], rates[-2], rates[-1]
    accelerating = r3 > r2 > r1
    decelerating = r3 < r2 < r1
    return {
        "status": "Accelerating" if accelerating else "Decelerating" if decelerating else "Mixed/Stable",
        "recent_rates": [round(x, 1) for x in rates[-4:]],
        "latest_growth": round(r3, 1) if r3 else None,
    }


def altman_z_score(pl, bs, key, price, shares_outstanding):
    """Calculate Altman Z-Score for bankruptcy risk assessment."""
    try:
        # Get latest year available
        sales = metric_match(pl, ["sales", "revenue", "net sales"])
        np_ = metric_match(pl, ["net profit", "profit after", "net income"])
        op = metric_match(pl, ["operating profit", "ebitda", "ebit"]) or metric_match(key, ["ebitda"])
        ta = metric_match(bs, ["total", "total assets", "total liabilities"])
        cl = metric_match(bs, ["borrowings", "debt", "total debt", "loans"])
        ca = metric_match(bs, ["cash", "cash & bank", "cash and bank", "cash equivalent"])
        inv = metric_match(bs, ["inventory", "inventories", "stock"])
        recv = metric_match(bs, ["receivables", "debtors", "trade receivables"])
        reserves = metric_match(bs, ["reserves"])
        eq_sc = metric_match(bs, ["equity share capital", "equity capital", "share capital"])
        borrowings = metric_match(bs, ["borrowings", "debt", "total debt"])

        common = sorted(set(sales.keys()) & set(ta.keys()) & set(borrowings.keys()))
        if not common:
            return {"z_score": None, "zone": "Insufficient data"}

        yr = common[-1]
        s = sales.get(yr, 0) or 0
        t = ta.get(yr, 1) or 1
        op_val = op.get(yr, 0) or 0
        np_val = np_.get(yr, 0) or 0
        ca_val = (ca.get(yr, 0) or 0) + (inv.get(yr, 0) or 0) + (recv.get(yr, 0) or 0)
        cl_val = cl.get(yr, 1) or 1
        re_val = reserves.get(yr, 0) or 0
        de_val = cl.get(yr, 0) or 0
        eq_val = (eq_sc.get(yr, 0) or 0) + re_val
        shares_norm = normalize_shares(shares_outstanding)
        mv = (price or 0) * (shares_norm or 1) / 1e7  # Cr

        wc = ca_val - cl_val
        total_liab = t - eq_val

        A = wc / t if t != 0 else 0
        B = re_val / t if t != 0 else 0
        C = op_val / t if t != 0 else 0
        D = mv / total_liab if total_liab and total_liab != 0 else 0
        E = s / t if t != 0 else 0

        z = 1.2 * A + 1.4 * B + 3.3 * C + 0.6 * D + 1.0 * E
        if z > 2.99:
            zone = "Safe"
        elif z > 1.81:
            zone = "Grey Zone"
        else:
            zone = "Distress Zone"
        return {
            "z_score": round(z, 2),
            "zone": zone,
            "components": {
                "A (WC/TA)": round(A, 3),
                "B (RE/TA)": round(B, 3),
                "C (EBIT/TA)": round(C, 3),
                "D (MVE/TL)": round(D, 3),
                "E (S/TA)": round(E, 3),
            }
        }
    except Exception:
        return {"z_score": None, "zone": "Error computing"}


def entry_zones(latest_price, val_info):
    """Define bear/base/bull price targets and entry recommendation."""
    if not latest_price or not val_info or not val_info.get("weighted_iv"):
        return {"entry": "Insufficient data for entry zones"}
    iv = val_info["weighted_iv"]
    upside = val_info.get("upside_pct", 0)
    zones = {
        "strong_buy_below": round(iv * 0.8, 2),
        "buy_below": round(iv * 0.95, 2),
        "fair_value": round(iv, 2),
        "overvalued_above": round(iv * 1.2, 2),
        "current_price": round(latest_price, 2),
        "current_upside": upside,
    }
    # Recommendation
    if upside >= 30:
        zones["recommendation"] = "Strong Buy — deep value"
    elif upside >= 15:
        zones["recommendation"] = "Buy — good margin of safety"
    elif upside >= 0:
        zones["recommendation"] = "Hold — near fair value"
    elif upside >= -15:
        zones["recommendation"] = "Cautious — slightly overvalued"
    else:
        zones["recommendation"] = "Avoid — significantly overvalued, wait for correction"
    return zones


def magic_formula_rank(key, pl, bs, price, shares_outstanding):
    """Compute Magic Formula (ROCE rank + Earnings Yield rank) for value investing."""
    try:
        roce_vals = list(key.get("ROE", {}).values())
        roce_latest = roce_vals[-1] if roce_vals else None
        # Earnings Yield = EBIT / Enterprise Value
        ebit = metric_match(key, ["ebitda"])
        ebit_vals = list(ebit.values())
        ebit_latest = ebit_vals[-1] if ebit_vals else None
        debt = list(metric_match(bs, ["borrowings", "debt", "total debt"]).values())
        debt_latest = debt[-1] if debt else 0
        cash = list(metric_match(bs, ["cash", "cash & bank", "cash and bank"]).values())
        cash_latest = cash[-1] if cash else 0

        shares_norm = normalize_shares(shares_outstanding)
        ev = None
        if price and shares_norm:
            mcap = price * shares_norm / 1e7  # in Cr
            ev = mcap + (debt_latest or 0) - (cash_latest or 0)

        earnings_yield = None
        if ebit_latest and ev and ev > 0:
            earnings_yield = round(ebit_latest / ev * 100, 1)

        formulas = {}
        if roce_latest:
            formulas["roce"] = f"ROCE = Operating Profit / (Equity + Borrowings)\n= EBIT / (Shareholders' Equity + Total Debt)"
        if earnings_yield:
            formulas["earnings_yield"] = f"Earnings Yield = EBIT / Enterprise Value\n= EBITDA / (Market Cap + Debt - Cash)"

        return {
            "roce": round(roce_latest, 1) if roce_latest else None,
            "earnings_yield": earnings_yield,
            "magic_formula_rank": "Value + Quality" if (roce_latest and roce_latest > 15 and earnings_yield and earnings_yield > 8) else "Quality" if (roce_latest and roce_latest > 15) else "Value" if (earnings_yield and earnings_yield > 8) else "Below Thresholds",
            "ev": round(ev, 2) if ev else None,
            "formulas": formulas,
        }
    except (TypeError, ValueError, IndexError, ZeroDivisionError) as e:
        log.warning("magic_formula_rank failed: %s", e)
        return {"roce": None, "earnings_yield": None}


# ══════════════════════════════════════════════════════════════════════════════
# INVESTOR FRAMEWORK SCORING ENGINES (0-100 each)
# ══════════════════════════════════════════════════════════════════════════════

def _clip(v, lo=0, hi=100):
    return max(lo, min(hi, v))

def _framework_base(name, investor_tagline):
    return {"framework": name, "tagline": investor_tagline, "score": 0, "max": 100, "components": {}, "verdict": "Insufficient data"}

def _fw_verdict(score):
    if score >= 80: return "STRONG BUY"
    if score >= 60: return "BUY"
    if score >= 40: return "HOLD"
    if score >= 20: return "CAUTION"
    return "AVOID"

def buffett_score(pl, bs, cf, key, latest_roe_v, latest_roce_v, latest_opm_v, latest_de_v, latest_int_cov, cfo_positive, rev_cagr_v, pro_cagr_v, consistent_growth, debt_reduced, industry):
    """Warren Buffett — Buy wonderful businesses at fair prices.
    Durable moat, consistent ROE, low debt, owner earnings."""
    fw = _framework_base("Buffett (Buffettology)",
        '"It\'s far better to buy a wonderful company at a fair price than a fair company at a wonderful price."')
    c = {}
    s = 0

    # 1. ROE quality (30 pts)
    roe_score = 0
    if latest_roe_v:
        if latest_roe_v >= 25: roe_score = 30
        elif latest_roe_v >= 20: roe_score = 26
        elif latest_roe_v >= 15: roe_score = 20
        elif latest_roe_v >= 10: roe_score = 12
        else: roe_score = 5
    c["ROE Quality"] = {"score": roe_score, "max": 30, "detail": f"ROE: {latest_roe_v}% (threshold: >15% sustained)" if latest_roe_v else "N/A"}
    s += roe_score

    # 2. Durable competitive advantage (25 pts)
    moat_score = 0
    if latest_roce_v and latest_roce_v > 20: moat_score += 8
    elif latest_roce_v and latest_roce_v > 15: moat_score += 5
    if latest_opm_v and latest_opm_v > 20: moat_score += 8
    elif latest_opm_v and latest_opm_v > 15: moat_score += 5
    if latest_roe_v and latest_roe_v > 20: moat_score += 9
    elif latest_roe_v and latest_roe_v > 15: moat_score += 5
    moat_score = min(moat_score, 25)
    c["Moat (Durable Advantage)"] = {"score": moat_score, "max": 25, "detail": f"ROCE: {latest_roce_v}%, OPM: {latest_opm_v}%, ROE: {latest_roe_v}%" if latest_roce_v else "N/A"}
    s += moat_score

    # 3. Low debt (15 pts)
    debt_score = 0
    if latest_de_v is not None:
        if latest_de_v < 0.25: debt_score = 15
        elif latest_de_v < 0.5: debt_score = 13
        elif latest_de_v < 1.0: debt_score = 9
        elif latest_de_v < 2.0: debt_score = 5
        else: debt_score = 2
    if debt_reduced: debt_score = min(debt_score + 2, 15)
    if latest_int_cov and latest_int_cov < 2: debt_score = max(debt_score - 4, 0)
    c["Debt Discipline"] = {"score": debt_score, "max": 15, "detail": f"D/E: {latest_de_v}" if latest_de_v is not None else "N/A"}
    s += debt_score

    # 4. Owner earnings quality (15 pts)
    eq_score = 0
    if cfo_positive: eq_score += 8
    np_vals = list(pl.get("Net profit", {}).values()) if pl else []
    cfo_vals = list(cf.get("Cash from Operating Activity", {}).values()) if cf else []
    if len(cfo_vals) >= 1 and len(np_vals) >= 1 and cfo_vals[-1] and np_vals[-1] and np_vals[-1] > 0:
        ratio = cfo_vals[-1] / np_vals[-1]
        if ratio > 1.0: eq_score += 7
        elif ratio > 0.7: eq_score += 5
        elif ratio > 0.3: eq_score += 3
    c["Owner Earnings Quality"] = {"score": eq_score, "max": 15, "detail": f"CFO positive: {cfo_positive}"}
    s += eq_score

    # 5. Consistent earnings growth (15 pts)
    gr_score = 0
    if consistent_growth: gr_score += 5
    if rev_cagr_v and rev_cagr_v > 10: gr_score += 5
    elif rev_cagr_v and rev_cagr_v > 5: gr_score += 3
    if pro_cagr_v and pro_cagr_v > 12: gr_score += 5
    elif pro_cagr_v and pro_cagr_v > 5: gr_score += 3
    gr_score = min(gr_score, 15)
    c["Earnings Consistency"] = {"score": gr_score, "max": 15, "detail": f"Rev CAGR: {rev_cagr_v}%, Profit CAGR: {pro_cagr_v}%"}
    s += gr_score

    s = _clip(s)
    fw["score"] = round(s, 1)
    fw["components"] = c
    fw["verdict"] = _fw_verdict(s)
    return fw


def lynch_score(pl, bs, key, latest_pe_v, latest_eps_v, latest_de_v, latest_roe_v, latest_opm_v, rev_cagr_v, pro_cagr_v, consistent_growth, cfo_positive, promoter_info, inst_holding):
    """Peter Lynch — Growth at a reasonable price (GARP).
    PEG < 1, understand the business, debt-free, earnings story."""
    fw = _framework_base("Lynch (GARP)",
        '"Know what you own, and know why you own it."')
    c = {}
    s = 0

    # 1. PEG Ratio (25 pts)
    peg_score = 0
    if latest_pe_v and latest_eps_v and pro_cagr_v and pro_cagr_v > 0 and latest_pe_v > 0:
        peg = latest_pe_v / pro_cagr_v
        if peg < 0.5: peg_score = 25
        elif peg < 0.8: peg_score = 22
        elif peg < 1.0: peg_score = 18
        elif peg < 1.5: peg_score = 12
        elif peg < 2.0: peg_score = 6
        else: peg_score = 2
        c["PEG Ratio"] = {"score": peg_score, "max": 25, "detail": f"PEG: {peg:.2f} (P/E {latest_pe_v}x / Growth {pro_cagr_v}%)"}
    else:
        c["PEG Ratio"] = {"score": 0, "max": 25, "detail": "Insufficient data"}
    s += peg_score

    # 2. Debt-free or low debt (20 pts)
    debt_score = 0
    if latest_de_v is not None:
        if latest_de_v < 0.1: debt_score = 20
        elif latest_de_v < 0.3: debt_score = 17
        elif latest_de_v < 0.5: debt_score = 14
        elif latest_de_v < 1.0: debt_score = 8
        elif latest_de_v < 2.0: debt_score = 4
        else: debt_score = 1
    c["Balance Sheet Strength"] = {"score": debt_score, "max": 20, "detail": f"D/E: {latest_de_v}" if latest_de_v is not None else "N/A"}
    s += debt_score

    # 3. Earnings story (20 pts)
    earn_score = 0
    if consistent_growth: earn_score += 5
    if pro_cagr_v and pro_cagr_v > 15: earn_score += 8
    elif pro_cagr_v and pro_cagr_v > 10: earn_score += 5
    elif pro_cagr_v and pro_cagr_v > 5: earn_score += 3
    if rev_cagr_v and rev_cagr_v > 10: earn_score += 5
    elif rev_cagr_v and rev_cagr_v > 5: earn_score += 3
    if latest_roe_v and latest_roe_v > 15: earn_score += 2
    earn_score = min(earn_score, 20)
    c["Earnings Story"] = {"score": earn_score, "max": 20, "detail": f"Profit CAGR: {pro_cagr_v}%, Consistent: {consistent_growth}"}
    s += earn_score

    # 4. Cash flow health (15 pts)
    cf_score = 0
    if cfo_positive: cf_score += 8
    np_vals = list(pl.get("Net profit", {}).values()) if pl else []
    if np_vals and len(np_vals) >= 3:
        losses = sum(1 for v in np_vals[-3:] if v and v < 0)
        if losses == 0: cf_score += 7
        elif losses <= 1: cf_score += 4
    c["Cash Flow Health"] = {"score": cf_score, "max": 15, "detail": f"CFO positive: {cfo_positive}"}
    s += cf_score

    # 5. Company type & institutional sweet spot (20 pts)
    type_score = 0
    if latest_opm_v and latest_opm_v > 20: type_score += 5
    if latest_roe_v and latest_roe_v > 20: type_score += 5
    if inst_holding is not None:
        if 10 <= inst_holding <= 60:
            type_score += 5
        elif inst_holding > 60:
            type_score += 2
    if promoter_info and promoter_info > 50: type_score += 5
    elif promoter_info and promoter_info > 30: type_score += 3
    type_score = min(type_score, 20)
    c["Company Type & Smart Money"] = {"score": type_score, "max": 20, "detail": f"Institutional: {inst_holding}%, Promoter: {promoter_info}%" if inst_holding else "N/A"}
    s += type_score

    s = _clip(s)
    fw["score"] = round(s, 1)
    fw["components"] = c
    fw["verdict"] = _fw_verdict(s)
    return fw


def munger_score(latest_roe_v, latest_roce_v, latest_opm_v, latest_de_v, latest_int_cov, rev_cagr_v, pro_cagr_v, consistent_growth, debt_reduced, cfo_positive, piotroski, altman_zone, industry, red_flags_list, promoter_info):
    """Charlie Munger — Invert, always invert. Mental models, moat, rationality.
    Lollapalooza effect: multiple factors compounding together."""
    fw = _framework_base("Munger (Mental Models)",
        '"The big money is not in the buying and selling, but in the waiting."')
    c = {}
    s = 0

    # 1. Lollapalooza — multiple reinforcing factors (30 pts)
    lollapalooza = 0
    if latest_roe_v and latest_roe_v > 20: lollapalooza += 6
    if latest_roce_v and latest_roce_v > 20: lollapalooza += 6
    if latest_opm_v and latest_opm_v > 20: lollapalooza += 6
    if consistent_growth: lollapalooza += 6
    if latest_de_v is not None and latest_de_v < 0.3: lollapalooza += 6
    elif latest_de_v is not None and latest_de_v < 0.5: lollapalooza += 4
    lollapalooza = min(lollapalooza, 30)
    c["Lollapalooza Effect"] = {"score": lollapalooza, "max": 30, "detail": "Reinforcing factors: ROE, ROCE, OPM, growth, low debt"}
    s += lollapalooza

    # 2. Circle of competence / predictability (25 pts)
    pred_score = 0
    if consistent_growth: pred_score += 6
    if rev_cagr_v and rev_cagr_v > 10: pred_score += 5
    elif rev_cagr_v and rev_cagr_v > 5: pred_score += 3
    if pro_cagr_v and pro_cagr_v > 10: pred_score += 5
    elif pro_cagr_v and pro_cagr_v > 5: pred_score += 3
    if latest_int_cov and latest_int_cov > 3: pred_score += 5
    elif latest_int_cov and latest_int_cov > 2: pred_score += 3
    if cfo_positive: pred_score += 4
    pred_score = min(pred_score, 25)
    c["Predictability (Circle of Competence)"] = {"score": pred_score, "max": 25, "detail": f"Consistent growth: {consistent_growth}, Interest coverage: {latest_int_cov}"}
    s += pred_score

    # 3. Inversion — absence of red flags (20 pts)
    inv_score = 20
    critical_count = sum(1 for f in red_flags_list if f.get("severity") == "critical") if red_flags_list else 0
    high_count = sum(1 for f in red_flags_list if f.get("severity") == "high") if red_flags_list else 0
    inv_score -= critical_count * 6
    inv_score -= high_count * 3
    inv_score = max(inv_score, 0)
    c["Inversion (No Red Flags)"] = {"score": inv_score, "max": 20, "detail": f"Critical: {critical_count}, High: {high_count}"}
    s += inv_score

    # 4. Patience & moat durability (15 pts)
    patience = 0
    if latest_roe_v and latest_roe_v > 15: patience += 4
    if latest_roce_v and latest_roce_v > 15: patience += 4
    if latest_de_v is not None and latest_de_v < 0.5: patience += 4
    if promoter_info and promoter_info > 50: patience += 3
    patience = min(patience, 15)
    c["Moat Durability (Patience)"] = {"score": patience, "max": 15, "detail": "Wide moat + low debt = hold forever quality"}
    s += patience

    # 5. Rationality check — debt & dilution (10 pts)
    rational = 10
    if latest_de_v is not None and latest_de_v > 1.0: rational -= 4
    if latest_de_v is not None and latest_de_v > 2.0: rational -= 3
    if not debt_reduced and latest_de_v and latest_de_v > 0.5: rational -= 2
    rational = max(rational, 0)
    c["Capital Allocation Rationality"] = {"score": rational, "max": 10, "detail": f"Debt trend: {'Reducing' if debt_reduced else 'Increasing'}"}
    s += rational

    s = _clip(s)
    fw["score"] = round(s, 1)
    fw["components"] = c
    fw["verdict"] = _fw_verdict(s)
    return fw


def fisher_score(pl, bs, key, latest_opm_v, latest_roe_v, latest_roce_v, latest_de_v, rev_cagr_v, pro_cagr_v, consistent_growth, cfo_positive, debt_reduced, promoter_info, inst_holding, industry, mcap):
    """Phil Fisher — Scuttlebutt method. Invest in companies with high innovation,
    strong R&D, excellent management, and long-term growth horizons."""
    fw = _framework_base("Fisher (Scuttlebutt)",
        '"The stock market is filled with individuals who know the price of everything, but the value of nothing."')
    c = {}
    s = 0

    # 1. Innovation & R&D potential (20 pts)
    innovation = 0
    if industry:
        ind_lower = industry.lower()
        innovative_kw = ["software", "pharma", "ai", "semiconductor", "renewable", "defense", "digital", "fintech", "electronics", "healthcare", "biotech", "technology"]
        for kw in innovative_kw:
            if kw in ind_lower:
                innovation += 15
                break
    if rev_cagr_v and rev_cagr_v > 15: innovation += 5
    elif rev_cagr_v and rev_cagr_v > 10: innovation += 3
    innovation = min(innovation, 20)
    c["Innovation & R&D Potential"] = {"score": innovation, "max": 20, "detail": f"Industry: {industry}" if industry else "N/A"}
    s += innovation

    # 2. Margin quality & cost efficiency (25 pts)
    margin_score = 0
    if latest_opm_v:
        if latest_opm_v > 25: margin_score = 25
        elif latest_opm_v > 20: margin_score = 21
        elif latest_opm_v > 15: margin_score = 16
        elif latest_opm_v > 10: margin_score = 10
        else: margin_score = 5
    c["Profit Margins & Cost Efficiency"] = {"score": margin_score, "max": 25, "detail": f"OPM: {latest_opm_v}%"}
    s += margin_score

    # 3. Management quality (25 pts)
    mgmt_score = 0
    if promoter_info:
        if promoter_info > 70: mgmt_score += 10
        elif promoter_info > 50: mgmt_score += 8
        elif promoter_info > 30: mgmt_score += 5
        else: mgmt_score += 2
    if debt_reduced: mgmt_score += 5
    if latest_de_v is not None and latest_de_v < 0.5: mgmt_score += 5
    if latest_roe_v and latest_roe_v > 18: mgmt_score += 5
    elif latest_roe_v and latest_roe_v > 12: mgmt_score += 3
    mgmt_score = min(mgmt_score, 25)
    c["Management & Leadership"] = {"score": mgmt_score, "max": 25, "detail": f"Promoter: {promoter_info}%, Debt reduced: {debt_reduced}"}
    s += mgmt_score

    # 4. Sales organization & growth (15 pts)
    sales_score = 0
    if consistent_growth: sales_score += 5
    if rev_cagr_v and rev_cagr_v > 15: sales_score += 5
    elif rev_cagr_v and rev_cagr_v > 10: sales_score += 3
    if pro_cagr_v and pro_cagr_v > 15: sales_score += 5
    elif pro_cagr_v and pro_cagr_v > 10: sales_score += 3
    sales_score = min(sales_score, 15)
    c["Sales & Growth Organization"] = {"score": sales_score, "max": 15, "detail": f"Rev CAGR: {rev_cagr_v}%, Profit CAGR: {pro_cagr_v}%"}
    s += sales_score

    # 5. Long-term horizon (15 pts)
    lt_score = 0
    if mcap and mcap > 5000: lt_score += 4
    if latest_roe_v and latest_roe_v > 15: lt_score += 4
    if latest_de_v is not None and latest_de_v < 0.5: lt_score += 4
    if cfo_positive: lt_score += 3
    lt_score = min(lt_score, 15)
    c["Long-Term Horizon"] = {"score": lt_score, "max": 15, "detail": f"Market cap: {mcap}Cr" if mcap else "N/A"}
    s += lt_score

    s = _clip(s)
    fw["score"] = round(s, 1)
    fw["components"] = c
    fw["verdict"] = _fw_verdict(s)
    return fw


def canslim_score(pl, bs, key, latest_eps_v, latest_pe_v, latest_roe_v, latest_opm_v, rev_cagr_v, pro_cagr_v, eps_ser, sales, profit, latest_price_v, shares_units, promoter_info, inst_holding, mcap, industry, consistent_growth, revenue_accel_status):
    """William O'Neil — CAN SLIM. Momentum + fundamentals.
    C: Current earnings, A: Annual earnings, N: New, S: Supply, L: Leader, I: Institutional, M: Market."""
    fw = _framework_base("CAN SLIM (O'Neil)",
        '"To make money in stocks you must have a clear plan and follow it carefully."')
    c = {}
    s = 0

    # C — Current quarterly earnings (20 pts)
    c_score = 0
    if eps_ser and len(eps_ser) >= 2:
        eps_vals = list(eps_ser.values())
        if len(eps_vals) >= 2 and eps_vals[-1] and eps_vals[-2]:
            eps_growth = (eps_vals[-1] - eps_vals[-2]) / abs(eps_vals[-2]) * 100
            if eps_growth > 50: c_score = 20
            elif eps_growth > 25: c_score = 17
            elif eps_growth > 10: c_score = 12
            elif eps_growth > 0: c_score = 7
            else: c_score = 2
    if pro_cagr_v:
        if pro_cagr_v > 25: c_score = max(c_score, 16)
        elif pro_cagr_v > 15: c_score = max(c_score, 12)
    c["C: Current Earnings"] = {"score": c_score, "max": 20, "detail": f"Profit CAGR: {pro_cagr_v}%" if pro_cagr_v else "N/A"}
    s += c_score

    # A — Annual earnings growth (15 pts)
    a_score = 0
    if rev_cagr_v:
        if rev_cagr_v > 25: a_score = 15
        elif rev_cagr_v > 15: a_score = 12
        elif rev_cagr_v > 10: a_score = 8
        elif rev_cagr_v > 5: a_score = 5
        else: a_score = 2
    if consistent_growth: a_score = min(a_score + 3, 15)
    c["A: Annual Earnings"] = {"score": a_score, "max": 15, "detail": f"Rev CAGR: {rev_cagr_v}%" if rev_cagr_v else "N/A"}
    s += a_score

    # N — New product/management/high (15 pts)
    n_score = 0
    if revenue_accel_status == "Accelerating": n_score += 6
    if industry:
        ind_lower = industry.lower()
        new_kw = ["software", "ai", "fintech", "semiconductor", "renewable", "ev", "defense", "pharma", "digital", "cloud"]
        for kw in new_kw:
            if kw in ind_lower:
                n_score += 5
                break
    if latest_roe_v and latest_roe_v > 20: n_score += 4
    n_score = min(n_score, 15)
    c["N: New (Products/Highs)"] = {"score": n_score, "max": 15, "detail": f"Revenue: {revenue_accel_status}"}
    s += n_score

    # S — Supply and demand (15 pts)
    s_score = 0
    if shares_units:
        if shares_units < 1e7: s_score += 6
        elif shares_units < 5e7: s_score += 4
        else: s_score += 2
    if promoter_info and promoter_info > 60: s_score += 5
    if mcap and mcap > 5000: s_score += 4
    s_score = min(s_score, 15)
    c["S: Supply & Demand"] = {"score": s_score, "max": 15, "detail": f"Shares: {shares_units}" if shares_units else "N/A"}
    s += s_score

    # L — Leader or laggard (15 pts)
    l_score = 0
    if latest_opm_v:
        if latest_opm_v > 25: l_score += 5
        elif latest_opm_v > 15: l_score += 3
    if latest_roe_v:
        if latest_roe_v > 25: l_score += 5
        elif latest_roe_v > 15: l_score += 3
    if mcap and mcap > 20000: l_score += 5
    elif mcap and mcap > 5000: l_score += 3
    l_score = min(l_score, 15)
    c["L: Leader or Laggard"] = {"score": l_score, "max": 15, "detail": f"Margins: {latest_opm_v}%, Mcap: {mcap}Cr" if mcap else "N/A"}
    s += l_score

    # I — Institutional sponsorship (10 pts)
    i_score = 0
    if inst_holding is not None:
        if inst_holding > 30: i_score = 10
        elif inst_holding > 15: i_score = 7
        elif inst_holding > 5: i_score = 4
        else: i_score = 2
    c["I: Institutional Sponsorship"] = {"score": i_score, "max": 10, "detail": f"Institutional: {inst_holding}%" if inst_holding is not None else "N/A"}
    s += i_score

    # M — Market direction (10 pts) — proxy with consistency
    m_score = 0
    if consistent_growth: m_score += 5
    if rev_cagr_v and rev_cagr_v > 10: m_score += 5
    elif rev_cagr_v and rev_cagr_v > 5: m_score += 3
    m_score = min(m_score, 10)
    c["M: Market Direction (Proxy)"] = {"score": m_score, "max": 10, "detail": "Based on growth consistency as market proxy"}
    s += m_score

    s = _clip(s)
    fw["score"] = round(s, 1)
    fw["components"] = c
    fw["verdict"] = _fw_verdict(s)
    return fw


def overall_investor_score(framework_results):
    """Aggregate all 5 frameworks into one Overall Investment Score (0-100)."""
    names = [f["framework"] for f in framework_results]
    scores = [f["score"] for f in framework_results]
    avg = sum(scores) / len(scores) if scores else 0
    # Weighted: Buffett & Lynch higher weight for value investors
    weighted = 0
    weights = {"Buffett (Buffettology)": 0.25, "Lynch (GARP)": 0.20, "Munger (Mental Models)": 0.20, "Fisher (Scuttlebutt)": 0.15, "CAN SLIM (O'Neil)": 0.20}
    total_w = 0
    for fw in framework_results:
        w = weights.get(fw["framework"], 0.2)
        weighted += fw["score"] * w
        total_w += w
    weighted_score = weighted / total_w if total_w > 0 else avg

    if weighted_score >= 80: overall_verdict = "STRONG BUY — Exceptional across all frameworks"
    elif weighted_score >= 60: overall_verdict = "BUY — Good alignment with multiple strategies"
    elif weighted_score >= 40: overall_verdict = "HOLD — Mixed signals across frameworks"
    elif weighted_score >= 20: overall_verdict = "CAUTION — Most frameworks disagree"
    else: overall_verdict = "AVOID — Consistently poor across all frameworks"

    return {
        "overall_score": round(weighted_score, 1),
        "simple_avg": round(avg, 1),
        "verdict": overall_verdict,
        "frameworks_included": len(framework_results),
    }


def overall_signal(scores, overall_score, val_info, risks, piotroski, promoter_holding):
    """Aggregate all signals into a single BUY/SELL/WAIT recommendation."""
    signal = {"strength": 0, "label": "NEUTRAL", "details": []}

    # Score signal
    if overall_score >= 8.0:
        signal["strength"] += 2
        signal["details"].append("Strong fundamentals")
    elif overall_score >= 6.5:
        signal["strength"] += 1
        signal["details"].append("Decent fundamentals")
    elif overall_score < 5.0:
        signal["strength"] -= 1
        signal["details"].append("Weak fundamentals")

    # Valuation signal
    if val_info:
        ups = val_info.get("upside_pct", 0)
        if ups >= 20:
            signal["strength"] += 2
            signal["details"].append("Undervalued")
        elif ups >= 5:
            signal["strength"] += 1
            signal["details"].append("Slightly undervalued")
        elif ups <= -20:
            signal["strength"] -= 1
            signal["details"].append("Overvalued")

    # Piotroski signal
    if piotroski is not None:
        if piotroski >= 7:
            signal["strength"] += 1
            signal["details"].append(f"Strong financial health (F-Score {piotroski}/9)")
        elif piotroski <= 3:
            signal["strength"] -= 1
            signal["details"].append(f"Weak financial health (F-Score {piotroski}/9)")

    # Risk signal
    if len(risks) <= 1:
        signal["strength"] += 1
        signal["details"].append("Low risk profile")
    elif len(risks) >= 4:
        signal["strength"] -= 1
        signal["details"].append("Elevated risk profile")

    # Promoter signal
    if promoter_holding and promoter_holding > 60:
        signal["strength"] += 1
        signal["details"].append("Strong promoter holding")
    elif promoter_holding and promoter_holding < 25:
        signal["strength"] -= 1
        signal["details"].append("Low promoter holding")

    if signal["strength"] >= 4:
        signal["label"] = "STRONG BUY"
    elif signal["strength"] >= 2:
        signal["label"] = "BUY"
    elif signal["strength"] >= 0:
        signal["label"] = "WAIT & WATCH"
    elif signal["strength"] >= -2:
        signal["label"] = "CAUTION"
    else:
        signal["label"] = "SELL / AVOID"

    return signal


# ══════════════════════════════════════════════════════════════════════════════
# COMPREHENSIVE 20-POINT FRAMEWORK & BUY CONFIRMATION
# ══════════════════════════════════════════════════════════════════════════════

def assess_industry_future(industry):
    """Assess if industry has strong future potential."""
    if not industry:
        return {"score": 5, "label": "Unknown Industry", "detail": "Industry not identified"}
    ind_lower = industry.lower()
    for kw in GROWING_INDUSTRIES:
        if kw in ind_lower:
            return {"score": 9, "label": "Growing Industry", "detail": f"{industry} is in a high-growth sector"}
    for kw in STABLE_INDUSTRIES:
        if kw in ind_lower:
            return {"score": 7, "label": "Stable Industry", "detail": f"{industry} is a stable, established sector"}
    return {"score": 5, "label": "Neutral Industry", "detail": f"{industry} has moderate growth prospects"}


def assess_market_position(mcap, industry):
    """Assess market position based on market cap and industry."""
    if not mcap:
        return {"score": 5, "label": "Unknown Position", "detail": "Insufficient data"}
    if mcap > 50000:
        return {"score": 9, "label": "Market Leader", "detail": f"Large cap with ₹{mcap:,.0f} Cr market cap — likely market leader"}
    elif mcap > 20000:
        return {"score": 8, "label": "Strong Position", "detail": f"Large cap with ₹{mcap:,.0f} Cr market cap"}
    elif mcap > 5000:
        return {"score": 7, "label": "Mid Cap Leader", "detail": f"Mid cap with ₹{mcap:,.0f} Cr market cap"}
    elif mcap > 1000:
        return {"score": 6, "label": "Small Cap", "detail": f"Small cap with ₹{mcap:,.0f} Cr market cap — higher risk, higher reward"}
    else:
        return {"score": 4, "label": "Micro Cap", "detail": f"Very small market cap — high risk, low liquidity"}


def assess_management_quality(promoter_holding, debt_trend, roce, roe):
    """Assess management quality from available indicators."""
    score = 5
    reasons = []
    flags = []

    # Promoter holding
    if promoter_holding and promoter_holding > 60:
        score += 2
        reasons.append(f"High promoter holding ({promoter_holding}%) shows confidence")
    elif promoter_holding and promoter_holding > 40:
        score += 1
        reasons.append(f"Moderate promoter holding ({promoter_holding}%)")
    elif promoter_holding and promoter_holding < 25:
        score -= 1
        flags.append(f"Low promoter holding ({promoter_holding}%) — lack of skin in the game")

    # Debt management
    if debt_trend == "Reducing":
        score += 1
        reasons.append("Management is reducing debt — good capital allocation")
    else:
        flags.append("Debt is increasing — monitor debt levels")

    # Efficiency
    if roce and roce > 20:
        score += 1
        reasons.append(f"Strong ROCE ({roce}%) — efficient capital use")
    elif roce and roce < 10:
        flags.append(f"Low ROCE ({roce}%) — inefficient capital use")
    if roe and roe > 18:
        score += 1
        reasons.append(f"Strong ROE ({roe}%) — shareholder value creation")
    elif roe and roe < 10:
        flags.append(f"Low ROE ({roe}%) — weak shareholder returns")

    score = max(1, min(10, score))
    return {
        "score": score,
        "label": "Good" if score >= 7 else "Average" if score >= 5 else "Poor",
        "reasons": reasons,
        "flags": flags,
    }


def assess_dividend_history(latest_eps, series_profit_list):
    """Assess dividend payment potential from profit series list."""
    if not latest_eps or not series_profit_list:
        return {"score": 5, "label": "Unknown", "detail": "Insufficient data"}
    np_vals = [d["value"] for d in series_profit_list if d.get("value") is not None]
    if not np_vals:
        return {"score": 5, "label": "Unknown", "detail": "Insufficient data"}
    latest_np = np_vals[-1]
    if latest_eps > 0 and latest_np > 0:
        consecutive_profits = sum(1 for v in np_vals[-5:] if v > 0) if len(np_vals) >= 5 else len(np_vals)
        if consecutive_profits >= 5:
            return {"score": 9, "label": "Consistent Profits", "detail": f"Profitable for {consecutive_profits} years — strong dividend potential"}
        elif consecutive_profits >= 3:
            return {"score": 7, "label": "Mostly Profitable", "detail": f"Profitable for {consecutive_profits}/5 years"}
        else:
            return {"score": 5, "label": "Inconsistent Profits", "detail": "Profit history is inconsistent"}
    elif latest_np < 0:
        return {"score": 2, "label": "Loss Making", "detail": "Company is making losses — no dividend likely"}
    return {"score": 5, "label": "Unknown", "detail": "Insufficient data"}


def assess_economic_sensitivity(industry, debt, interest_coverage):
    """Assess sensitivity to economic conditions."""
    score = 7
    warnings = []
    if not industry:
        return {"score": 5, "label": "Unknown", "detail": "Insufficient data", "warnings": warnings}
    ind_lower = industry.lower()
    # Cyclical industries
    cyclical_kw = ["metal", "mining", "commodity", "oil", "gas", "real estate", "construction", "auto"]
    for kw in cyclical_kw:
        if kw in ind_lower:
            score -= 2
            warnings.append(f"{industry} is cyclical — sensitive to economic downturns")
            break
    # Defensive industries
    defensive_kw = ["fmcg", "pharma", "healthcare", "food", "beverage", "utility", "insurance"]
    for kw in defensive_kw:
        if kw in ind_lower:
            score += 1
            break
    # High debt increases economic sensitivity
    if debt and debt > 1:
        score -= 1
        warnings.append("High debt makes company vulnerable to rising interest rates")
    if interest_coverage and interest_coverage < 2:
        score -= 1
        warnings.append("Low interest coverage — at risk during economic slowdown")
    score = max(1, min(10, score))
    return {
        "score": score,
        "label": "Low Sensitivity" if score >= 7 else "Moderate Sensitivity" if score >= 5 else "High Sensitivity",
        "detail": f"Economic sensitivity score: {score}/10",
        "warnings": warnings,
    }


def assess_institutional_buying(institutional_holding, promoter_holding):
    """Assess institutional interest."""
    score = 5
    detail = "Insufficient data"
    if institutional_holding is not None:
        if institutional_holding > 30:
            score = 9
            detail = f"High institutional holding ({institutional_holding}%) — smart money is interested"
        elif institutional_holding > 15:
            score = 7
            detail = f"Moderate institutional holding ({institutional_holding}%) — some institutional interest"
        elif institutional_holding > 5:
            score = 5
            detail = f"Low institutional holding ({institutional_holding}%) — limited institutional interest"
        else:
            score = 3
            detail = f"Very low institutional holding ({institutional_holding}%) — institutions are avoiding"
    return {"score": score, "label": detail, "detail": detail}


def detect_red_flags(pl, bs, cf, key, sales, profit, borrowings, cfo, promoter_holding,
                     rev_cagr, pro_cagr, latest_de, latest_pe, latest_roe, latest_roce,
                     latest_int_cov, debt_reduced, cfo_positive, eps_ser, industry):
    """Comprehensive red flag detection — flags that should make you AVOID a stock."""
    flags = []
    severity = "low"

    # 1. Fake accounting signs
    cfo_vals = list(cf.get("Cash from Operating Activity", {}).values()) if cf else []
    np_vals = list(pl.get("Net profit", {}).values()) if pl else []
    if len(cfo_vals) >= 3 and len(np_vals) >= 3:
        cfo_sum = sum(v for v in cfo_vals[-3:] if v)
        np_sum = sum(v for v in np_vals[-3:] if v)
        if cfo_sum and np_sum and cfo_sum < np_sum * 0.5:
            flags.append({
                "flag": "🚨 Earnings Quality Warning",
                "detail": "Operating cash flow is significantly lower than net profit over last 3 years — possible fake profits",
                "severity": "critical"
            })

    # 2. Continuous losses
    if profit and len(profit) >= 3:
        loss_years = sum(1 for v in list(profit.values())[-3:] if v and v < 0)
        if loss_years >= 3:
            flags.append({
                "flag": "🚨 Continuous Losses",
                "detail": "Company has been loss-making for 3 consecutive years",
                "severity": "critical"
            })
        elif loss_years >= 2:
            flags.append({
                "flag": "⚠️ Recurring Losses",
                "detail": "Losses in 2 of last 3 years — profitability issues",
                "severity": "high"
            })

    # 3. Huge debt
    if latest_de is not None:
        if latest_de > 3:
            flags.append({
                "flag": "🚨 Critically High Debt",
                "detail": f"Debt-to-Equity ratio is {latest_de:.2f}x — dangerously high leverage",
                "severity": "critical"
            })
        elif latest_de > 2:
            flags.append({
                "flag": "⚠️ Very High Debt",
                "detail": f"Debt-to-Equity ratio is {latest_de:.2f}x — high leverage risk",
                "severity": "high"
            })

    # 4. Negative cash flow
    if cfo_positive is False:
        flags.append({
            "flag": "🚨 Negative Operating Cash Flow",
            "detail": "Company is not generating cash from core operations",
            "severity": "critical"
        })

    # 5. Falling revenue
    if sales and len(sales) >= 3:
        sv = list(sales.values())
        if len(sv) >= 3 and sv[-1] and sv[-2] and sv[-1] < sv[-2]:
            if len(sv) >= 3 and sv[-2] and sv[-3] and sv[-2] < sv[-3]:
                flags.append({
                    "flag": "🚨 Consistent Revenue Decline",
                    "detail": "Revenue declining for 2 consecutive years",
                    "severity": "critical"
                })
            else:
                flags.append({
                    "flag": "⚠️ Revenue Decline",
                    "detail": "Revenue declined in the latest year",
                    "severity": "high"
                })

    # 6. Revenue up but profit down
    if sales and profit and len(sales) >= 2 and len(profit) >= 2:
        sv = list(sales.values())
        pv = list(profit.values())
        if sv[-1] and sv[-2] and pv[-1] and pv[-2]:
            if sv[-1] > sv[-2] and pv[-1] < pv[-2]:
                flags.append({
                    "flag": "⚠️ Revenue Up, Profit Down",
                    "detail": "Revenue grew but profits fell — margin compression or cost issues",
                    "severity": "high"
                })

    # 7. Promoter selling
    if promoter_holding is not None and promoter_holding < 25:
            flags.append({
                "flag": "⚠️ Low Promoter Holding",
                "detail": f"Promoters hold only {promoter_holding}% — lack of confidence in own business",
                "severity": "high"
            })

    # 8. Overvaluation
    if latest_pe is not None and latest_pe > 50:
        flags.append({
            "flag": "⚠️ Highly Overvalued",
            "detail": f"P/E ratio is {latest_pe:.1f}x — extremely expensive valuation",
            "severity": "high"
        })
    elif latest_pe is not None and latest_pe > 30:
        flags.append({
            "flag": "⚠️ Expensive Valuation",
            "detail": f"P/E ratio is {latest_pe:.1f}x — above average valuation",
            "severity": "medium"
        })

    # 9. Low interest coverage
    if latest_int_cov is not None and latest_int_cov < 1.5:
        flags.append({
            "flag": "🚨 Debt Servicing Risk",
            "detail": f"Interest coverage ratio is {latest_int_cov:.1f}x — profits barely cover interest payments",
            "severity": "critical"
        })

    # 10. Low profitability
    if latest_roe is not None and latest_roe < 5:
        flags.append({
            "flag": "⚠️ Low Return on Equity",
            "detail": f"ROE is only {latest_roe:.1f}% — shareholder value being destroyed",
            "severity": "high"
        })

    # 11. Negative growth
    if rev_cagr is not None and rev_cagr < 0:
        flags.append({
            "flag": "⚠️ Negative Revenue Growth",
            "detail": f"Revenue CAGR is {rev_cagr:.1f}% — business is shrinking",
            "severity": "high"
        })

    # Determine overall severity
    critical_count = sum(1 for f in flags if f["severity"] == "critical")
    high_count = sum(1 for f in flags if f["severity"] == "high")
    if critical_count >= 2:
        severity = "critical"
    elif critical_count >= 1 or high_count >= 3:
        severity = "high"
    elif high_count >= 1:
        severity = "medium"

    return {
        "red_flags": flags,
        "total_flags": len(flags),
        "critical_count": critical_count,
        "severity": severity,
        "verdict": "🚨 AVOID — Multiple Critical Red Flags" if severity == "critical" else \
                   "⚠️ CAUTION — Significant Risk Factors Present" if severity == "high" else \
                   "👁️ Monitor — Some Risk Factors" if severity == "medium" else \
                   "✅ Clean — No Major Red Flags"
    }


def twenty_point_checklist(analyze_result):
    """Evaluate all 20 points from the investment framework and return pass/fail for each."""
    m = analyze_result["metrics"]
    b = analyze_result["business"]
    gr = analyze_result["growth"]
    bs_data = analyze_result["balance_sheet"]
    cf_data = analyze_result["cash_flow"]
    val = analyze_result.get("val")
    lt = analyze_result.get("long_term", {})
    pf = analyze_result.get("piotroski_fscore", {})
    az = analyze_result.get("altman_z", {})

    checklist = []

    # 1. Business Understanding
    bm_pass = bool(b.get("industry") and b.get("description") and len(b.get("description", "")) > 20)
    checklist.append({
        "point": 1, "category": "Business Model",
        "question": "Do you understand what the company does?",
        "pass": bm_pass,
        "detail": b.get("description", "N/A")[:100] + "..." if b.get("description") else "No description available",
        "weight": "high",
        "reason": f"Industry: {b.get('industry', 'N/A')}. Description available: {'Yes ✅' if b.get('description') else 'No ❌'}. {'You can understand the business model clearly.' if bm_pass else 'Cannot evaluate — insufficient business description.'}"
    })

    # 2. Revenue Growth
    rev_cagr = m.get("rev_cagr")
    rev_pass = rev_cagr is not None and rev_cagr > 5
    checklist.append({
        "point": 2, "category": "Revenue Growth",
        "question": "Is revenue growing consistently?",
        "pass": rev_pass,
        "detail": f"Revenue CAGR: {rev_cagr:.1f}%" if rev_cagr else "Insufficient data",
        "weight": "high",
        "reason": f"Revenue CAGR: {rev_cagr:.1f}% (threshold: >5%). {'✅ Growing above threshold — consistent revenue expansion.' if rev_pass else '❌ Below 5% threshold — weak or negative revenue growth.'}" if rev_cagr else "Insufficient revenue data to evaluate growth."
    })

    # 3. Profit Growth
    pro_cagr = m.get("pro_cagr")
    rev_cagr_val = rev_cagr or 0
    pro_cagr_val = pro_cagr or 0
    profit_faster = pro_cagr_val > rev_cagr_val if rev_cagr is not None and pro_cagr is not None else None
    pro_pass = pro_cagr is not None and pro_cagr > 5
    checklist.append({
        "point": 3, "category": "Profit Growth",
        "question": "Are profits growing? Is profit growing faster than revenue?",
        "pass": pro_pass,
        "detail": f"Profit CAGR: {pro_cagr:.1f}%{' — Growing faster than revenue ✅' if profit_faster else ''}" if pro_cagr else "Insufficient data",
        "weight": "high",
        "reason": f"Profit CAGR: {pro_cagr:.1f}% (threshold: >5%). {'✅ Strong profit growth.' if pro_pass else '❌ Below 5% threshold.'} {'Revenue CAGR: ' + str(rev_cagr) + '%.' if rev_cagr else ''} {'Profits growing faster than revenue — improving margins ✅.' if profit_faster else 'Revenue growing faster than profits — margin pressure ⚠️.' if profit_faster is False else ''}" if pro_cagr else "Insufficient profit data to evaluate."
    })

    # 4. Debt Analysis
    latest_de = m.get("latest_de")
    de_pass = latest_de is not None and latest_de < 1
    checklist.append({
        "point": 4, "category": "Debt Analysis",
        "question": "Is debt level safe? (D/E below 1 is ideal)",
        "pass": de_pass,
        "detail": f"Debt/Equity: {latest_de:.2f}x{' — Below 1 ✅' if latest_de and latest_de < 1 else ' — Above 1 ⚠️' if latest_de else ''}" if latest_de else "Insufficient data",
        "weight": "high",
        "reason": f"D/E Ratio: {latest_de:.2f}x (threshold: <1.0x). {'✅ Low debt — financially stable.' if de_pass else '❌ D/E above 1 — elevated debt risk.'} {'Healthy balance sheet with manageable leverage.' if latest_de and latest_de < 0.5 else 'Moderate debt levels — monitor closely.' if latest_de and latest_de < 1 else 'High debt — interest burden may impact profits.' if latest_de and latest_de >= 1 else ''}" if latest_de else "Insufficient debt data to evaluate."
    })

    # 5. Cash Flow Analysis
    cfo_pos = m.get("cfo_positive")
    fcf_series = analyze_result.get("series", {}).get("fcf", [])
    latest_fcf = fcf_series[-1]["value"] if fcf_series else None
    cf_pass = cfo_pos is True
    checklist.append({
        "point": 5, "category": "Cash Flow",
        "question": "Is company generating real cash? Is FCF positive?",
        "pass": cf_pass,
        "detail": "Operating cash flow is positive ✅" if cfo_pos else "Operating cash flow is negative 🚨" if cfo_pos is False else "Insufficient data",
        "weight": "high",
        "reason": f"Operating Cash Flow: {'Positive ✅ — company generates real cash from operations.' if cf_pass else 'Negative 🚨 — company burns cash, may need external funding.' if cfo_pos is False else 'Data unavailable.'} {'Cash from operations exceeds net profit — earnings quality is good.' if cf_pass else 'Negative CFO is a major red flag for long-term survival.' if cfo_pos is False else ''}"
    })

    # 6. ROE
    latest_roe = m.get("latest_roe")
    roe_pass = latest_roe is not None and latest_roe > 15
    checklist.append({
        "point": 6, "category": "ROE",
        "question": "Is ROE above 15%?",
        "pass": roe_pass,
        "detail": f"ROE: {latest_roe:.1f}%{' — Above 15% ✅' if latest_roe and latest_roe > 15 else ''}" if latest_roe else "Insufficient data",
        "weight": "medium",
        "reason": f"ROE: {latest_roe:.1f}% (threshold: >15%). {'✅ Strong return on equity — management generates good profits from shareholder capital.' if roe_pass else '❌ Below 15% threshold — capital efficiency needs improvement.'} {'ROE > 20% indicates a strong competitive advantage.' if latest_roe and latest_roe > 20 else ''}" if latest_roe else "Insufficient data to calculate ROE."
    })

    # 7. ROCE
    latest_roce = m.get("latest_roce")
    roce_pass = latest_roce is not None and latest_roce > 15
    checklist.append({
        "point": 7, "category": "ROCE",
        "question": "Is ROCE above 15%?",
        "pass": roce_pass,
        "detail": f"ROCE: {latest_roce:.1f}%{' — Above 15% ✅' if latest_roce and latest_roce > 15 else ''}" if latest_roce else "Insufficient data",
        "weight": "medium",
        "reason": f"ROCE: {latest_roce:.1f}% (threshold: >15%). {'✅ Efficient use of capital — company earns good returns on total capital employed.' if roce_pass else '❌ Below 15% threshold — capital efficiency below ideal.'} {'ROCE > 20% indicates a strong moat.' if latest_roce and latest_roce > 20 else ''}" if latest_roce else "Insufficient data to calculate ROCE."
    })

    # 8. Valuation
    latest_pe = m.get("latest_pe")
    upside = val.get("upside_pct") if val else None
    val_pass = (upside is not None and upside > 0) or (latest_pe is not None and latest_pe < 25)
    checklist.append({
        "point": 8, "category": "Valuation",
        "question": "Is valuation reasonable? (P/E reasonable, upside positive)",
        "pass": val_pass,
        "detail": f"P/E: {latest_pe:.1f}x, Upside: {upside:+.1f}%" if latest_pe and upside is not None else f"P/E: {latest_pe:.1f}x" if latest_pe else "Insufficient data",
        "weight": "high",
        "reason": f"P/E: {latest_pe:.1f}x | Intrinsic upside: {upside:+.1f}% (threshold: >0% or P/E <25x). {'✅ Valuation is reasonable — upside potential or reasonable P/E.' if val_pass else '❌ Stock appears overvalued — limited margin of safety.'}" if latest_pe and upside is not None else "Insufficient pricing data for full valuation assessment."
    })

    # 9. Competitive Advantage (Moat)
    opm = m.get("latest_opm")
    moat_score = 0
    if latest_roe and latest_roe > 20: moat_score += 1
    if latest_roce and latest_roce > 20: moat_score += 1
    if opm and opm > 20: moat_score += 1
    if pf.get("score") and pf["score"] >= 7: moat_score += 1
    moat_pass = moat_score >= 2
    checklist.append({
        "point": 9, "category": "Competitive Moat",
        "question": "Does company have a durable competitive advantage?",
        "pass": moat_pass,
        "detail": f"High ROE/ROCE/OPM ({moat_score}/4 indicators positive)" if moat_score >= 2 else "Limited moat indicators",
        "weight": "medium",
        "reason": f"Moat Score: {moat_score}/4 indicators positive. ROE>20%: {'✅' if latest_roe and latest_roe > 20 else '❌'} ({f'{latest_roe:.1f}%' if latest_roe else 'N/A'}), ROCE>20%: {'✅' if latest_roce and latest_roce > 20 else '❌'} ({f'{latest_roce:.1f}%' if latest_roce else 'N/A'}), OPM>20%: {'✅' if opm and opm > 20 else '❌'} ({f'{opm:.1f}%' if opm else 'N/A'}), Piotroski≥7: {'✅' if pf.get('score') and pf['score'] >= 7 else '❌'} ({pf.get('score', 'N/A')}/9). {'✅ Company shows signs of a competitive moat.' if moat_pass else '❌ Limited evidence of durable competitive advantage.'}"
    })

    # 10. Management Quality
    mgmt = assess_management_quality(b.get("promoter_holding"), bs_data.get("debt_trend"), latest_roce, latest_roe)
    mgmt_pass = mgmt["score"] >= 7
    checklist.append({
        "point": 10, "category": "Management Quality",
        "question": "Is management trustworthy with good capital allocation?",
        "pass": mgmt_pass,
        "detail": f"Score: {mgmt['score']}/10 — {mgmt['label']}" if mgmt["reasons"] else "Based on available indicators",
        "weight": "high",
        "reason": f"Management Score: {mgmt['score']}/10. {'✅ Management appears capable with good capital allocation.' if mgmt_pass else '❌ Management quality concerns.'} {'Promoter holding: ' + str(b.get('promoter_holding', 'N/A')) + '%.' if b.get('promoter_holding') else ''} {'Debt trend: ' + bs_data.get('debt_trend', 'N/A') + '.' if bs_data.get('debt_trend') else ''} {'ROCE: ' + str(latest_roce) + '%, ROE: ' + str(latest_roe) + '%.' if latest_roce and latest_roe else ''}"
    })

    # 11. Promoter Holding
    promoter = b.get("promoter_holding")
    prom_pass = promoter is not None and promoter > 50
    checklist.append({
        "point": 11, "category": "Promoter Holding",
        "question": "Is promoter holding high and stable?",
        "pass": prom_pass,
        "detail": f"Promoter holding: {promoter:.1f}%" if promoter else "Insufficient data",
        "weight": "medium",
        "reason": f"Promoter Holding: {promoter:.1f}% (threshold: >50%). {'✅ High promoter stake — management is aligned with shareholders.' if prom_pass else '❌ Below 50% — low promoter confidence.'} {'Holding >60% indicates strong promoter conviction.' if promoter and promoter > 60 else ''}" if promoter else "Insufficient promoter data."
    })

    # 12. Industry Future
    industry_future = assess_industry_future(b.get("industry"))
    ind_pass = industry_future["score"] >= 7
    checklist.append({
        "point": 12, "category": "Industry Future",
        "question": "Is the industry growing? (AI, Semi, Pharma, Defense, etc.)",
        "pass": ind_pass,
        "detail": industry_future["detail"],
        "weight": "medium",
        "reason": f"Industry: {b.get('industry', 'N/A')}. Assessment: {industry_future['label']} (Score: {industry_future['score']}/10). {'✅ Company operates in a growing/favorable industry.' if ind_pass else '❌ Industry outlook is neutral or challenging.'} {'Growing industries provide tailwinds for revenue expansion.' if industry_future['score'] >= 7 else 'Mature/declining industries face headwinds for growth.'}"
    })

    # 13. Market Position
    market_pos = assess_market_position(b.get("market_cap"), b.get("industry"))
    mkt_pass = market_pos["score"] >= 7
    checklist.append({
        "point": 13, "category": "Market Position",
        "question": "Is company a market leader or gaining share?",
        "pass": mkt_pass,
        "detail": market_pos["detail"],
        "weight": "medium",
        "reason": f"Market Cap: {'₹' + str(b.get('market_cap', '')) + 'Cr' if b.get('market_cap') else 'N/A'}. Assessment: {market_pos['label']} (Score: {market_pos['score']}/10). {'✅ Strong market position — likely has pricing power.' if mkt_pass else '❌ Limited market presence — may lack pricing power.'}"
    })

    # 14. Technical Analysis (entry timing)
    ent_pass = upside is not None and upside > 0
    checklist.append({
        "point": 14, "category": "Entry Timing",
        "question": "Is the entry price favorable based on valuation zones?",
        "pass": ent_pass,
        "detail": f"Entry recommendation: {val.get('val_verdict', 'N/A')}" if val else "Insufficient data",
        "weight": "low",
        "reason": f"Upside: {upside:+.1f}% | Intrinsic Value: ₹{val.get('weighted_iv', 'N/A')} | Current Price: ₹{val.get('current_price', 'N/A')}. {'✅ Entry price is favorable — margin of safety exists.' if ent_pass else '❌ Stock is overvalued — wait for better entry.'} Valuation verdict: {val.get('val_verdict', 'N/A')}." if val else "Insufficient data for entry timing analysis."
    })

    # 15. Risk Analysis
    risk_count = len(analyze_result.get("risks", []))
    risk_pass = risk_count <= 2
    checklist.append({
        "point": 15, "category": "Risk Assessment",
        "question": "What can go wrong? Are risks manageable?",
        "pass": risk_pass,
        "detail": f"{risk_count} risk factor(s) identified" if risk_count > 1 else "No major risks identified" if risk_count == 1 else "Clean risk profile",
        "weight": "high",
        "reason": f"Risk factors identified: {risk_count} (threshold: ≤2). {'✅ Risk profile is clean/acceptable.' if risk_pass else '❌ Too many risk factors — proceed with caution.'} Risks: {', '.join(analyze_result.get('risks', [])) if analyze_result.get('risks') else 'None identified.'}"
    })

    # 16. Dividend History
    div = assess_dividend_history(m.get("latest_eps"), analyze_result.get("series", {}).get("profit", []))
    div_pass = div["score"] >= 7
    checklist.append({
        "point": 16, "category": "Dividend History",
        "question": "Does company have consistent dividend payments?",
        "pass": div_pass,
        "detail": div["detail"],
        "weight": "low",
        "reason": f"Dividend Assessment: {div['label'] if 'label' in div else 'N/A'} (Score: {div['score']}/10). {'✅ Consistent dividend history — income investors may find this attractive.' if div_pass else '❌ Inconsistent or no dividend payments.'} {'Dividends indicate management confidence in cash flows.' if div_pass else 'Growth companies often reinvest profits rather than pay dividends.'}"
    })

    # 17. Economic Conditions sensitivity
    econ = assess_economic_sensitivity(b.get("industry"), latest_de, m.get("interest_coverage"))
    econ_pass = econ["score"] >= 6
    checklist.append({
        "point": 17, "category": "Economic Resilience",
        "question": "Can business withstand inflation, high interest rates, recession?",
        "pass": econ_pass,
        "detail": f"{econ['label']} (Score: {econ['score']}/10)",
        "weight": "medium",
        "reason": f"Economic Sensitivity: {econ['label']} (Score: {econ['score']}/10, threshold: ≥6). {'✅ Business is resilient to economic cycles.' if econ_pass else '❌ Highly sensitive to economic downturns.'} {'Low debt and essential-demand products provide stability.' if econ_pass else 'High cyclical exposure and/or debt increases vulnerability.'} Warnings: {', '.join(econ.get('warnings', [])) if econ.get('warnings') else 'None.'}"
    })

    # 18. Long-Term Growth Potential
    multibagger = lt.get("multibagger_possibility", "Low")
    lt_pass = multibagger in ("High", "Moderate")
    checklist.append({
        "point": 18, "category": "Long-Term Growth",
        "question": "Can company become 2x-10x bigger in 5-10 years?",
        "pass": lt_pass,
        "detail": f"Multibagger potential: {multibagger}",
        "weight": "medium",
        "reason": f"Long-Term Potential: {multibagger}. {'✅ Strong long-term wealth creation potential.' if multibagger == 'High' else '⚠️ Moderate potential — decent but not exceptional.' if multibagger == 'Moderate' else '❌ Limited long-term growth prospects.'} {'High growth + low debt = multibagger formula.' if multibagger == 'High' else 'Growth prospects are limited by market size or competition.' if multibagger == 'Low' else ''}"
    })

    # 19. Institutional Buying
    inst = b.get("institutional_holding")
    inst_pass = inst is not None and inst > 10
    checklist.append({
        "point": 19, "category": "Institutional Interest",
        "question": "Are mutual funds, FIIs, DIIs increasing holdings?",
        "pass": inst_pass,
        "detail": f"Institutional holding: {inst:.1f}%" if inst else "Insufficient data",
        "weight": "low",
        "reason": f"Institutional Holding: {inst:.1f}% (threshold: >10%). {'✅ Meaningful institutional presence — professional investors see value.' if inst_pass else '❌ Low institutional interest — may indicate governance or growth concerns.'} {'High institutional holding often correlates with better governance.' if inst and inst > 20 else ''}" if inst else "Insufficient institutional data."
    })

    # 20. Red Flags
    red_flags = analyze_result.get("red_flags", {}).get("red_flags", [])
    rf_pass = len(red_flags) == 0
    rf_detail_list = [f"{f['flag']}: {f['detail']}" for f in red_flags] if red_flags else []
    checklist.append({
        "point": 20, "category": "Red Flags",
        "question": "Are there any major red flags? (Fraud, losses, debt, negative cash flow)",
        "pass": rf_pass,
        "detail": f"{len(red_flags)} red flag(s) detected" if red_flags else "No red flags — Clean 🟢",
        "weight": "high",
        "reason": f"Red Flags: {len(red_flags)} detected (threshold: 0). {'✅ No red flags — company appears clean.' if rf_pass else '❌ Red flags present — investigate before investing.'} {'Details: ' + ' | '.join(rf_detail_list) if rf_detail_list else ''}"
    })

    # Calculate overall pass rate
    total_points = len(checklist)
    passed = sum(1 for c in checklist if c["pass"])
    pass_rate = (passed / total_points * 100) if total_points > 0 else 0

    # Weighted pass rate (high weight items count more)
    weight_map = {"high": 3, "medium": 2, "low": 1}
    total_weight = sum(weight_map.get(c["weight"], 1) for c in checklist)
    passed_weight = sum(weight_map.get(c["weight"], 1) for c in checklist if c["pass"])
    weighted_pass_rate = (passed_weight / total_weight * 100) if total_weight > 0 else 0

    # Generate verdict
    critical_fails = sum(1 for c in checklist if not c["pass"] and c["weight"] == "high")
    if critical_fails >= 3:
        decision = "🚨 AVOID — Multiple critical criteria failed"
        action = "DO NOT BUY. Critical issues need resolution first."
    elif pass_rate >= 80 and weighted_pass_rate >= 75:
        decision = "✅ STRONG BUY — Most criteria passed"
        action = "Good entry point for long-term investment."
    elif pass_rate >= 60 and weighted_pass_rate >= 55:
        decision = "⚠️ CAUTIOUS BUY — Some concerns exist"
        action = "Consider buying but monitor risk factors closely."
    elif pass_rate >= 40:
        decision = "👁️ WAIT & WATCH — Several concerns"
        action = "Wait for improvements in key areas before investing."
    else:
        decision = "🚨 AVOID — Too many criteria failed"
        action = "Not a suitable investment at this time."

    return {
        "checklist": checklist,
        "summary": {
            "total": total_points,
            "passed": passed,
            "failed": total_points - passed,
            "pass_rate": round(pass_rate, 1),
            "weighted_pass_rate": round(weighted_pass_rate, 1),
            "critical_fails": critical_fails,
        },
        "decision": decision,
        "action": action,
    }


def buy_confirmation_gate(twenty_point, red_flags_data, val_info, overall_score):
    """Final gate that must be passed before recommending a buy."""
    gate = {
        "buy_signal": False,
        "reasons_to_buy": [],
        "reasons_to_avoid": [],
        "final_verdict": "",
        "required_checks": [],
    }

    # Check 1: Pass rate >= 60%
    pass_rate = twenty_point["summary"]["pass_rate"]
    if pass_rate >= 60:
        gate["reasons_to_buy"].append(f"✅ {pass_rate:.0f}% of 20-point criteria passed")
    else:
        gate["reasons_to_avoid"].append(f"❌ Only {pass_rate:.0f}% of criteria passed (need ≥60%)")

    # Check 2: No critical red flags
    red_flags = red_flags_data.get("red_flags", [])
    critical_red_flags = [f for f in red_flags if f.get("severity") == "critical"]
    if len(critical_red_flags) == 0:
        gate["reasons_to_buy"].append("✅ No critical red flags detected")
    else:
        gate["reasons_to_avoid"].append(f"❌ {len(critical_red_flags)} critical red flag(s) — safety concern")

    # Check 3: Undervalued or fairly valued
    upside = val_info.get("upside_pct") if val_info else None
    if upside is not None and upside >= 0:
        gate["reasons_to_buy"].append(f"✅ Upside potential of {upside:+.1f}%")
    elif upside is not None and upside < 0:
        gate["reasons_to_avoid"].append(f"❌ Stock is overvalued (upside: {upside:+.1f}%)")
    else:
        gate["reasons_to_avoid"].append("❌ Valuation data insufficient")

    # Check 4: Overall score >= 6
    if overall_score >= 6:
        gate["reasons_to_buy"].append(f"✅ Fundamentals score {overall_score}/10")
    else:
        gate["reasons_to_avoid"].append(f"❌ Weak fundamentals ({overall_score}/10)")

    # Check 5: Revenue growing
    # Check 6: Profit growing (implied in overall score)

    # Final verdict
    buy_count = len(gate["reasons_to_buy"])
    avoid_count = len(gate["reasons_to_avoid"])

    gate["required_checks"] = {
        i+1: r.replace('✅ ', '').replace('❌ ', '')
        for i, r in enumerate(gate["reasons_to_buy"] + gate["reasons_to_avoid"])
    }

    if buy_count >= 3 and avoid_count == 0:
        gate["buy_signal"] = True
        gate["final_verdict"] = "✅ GREEN LIGHT — All major checks passed. Good to buy."
    elif buy_count >= 2 and avoid_count <= 1:
        gate["buy_signal"] = True
        gate["final_verdict"] = "🟡 CONDITIONAL BUY — Most checks passed. Buy with position sizing."
    elif avoid_count >= 3:
        gate["buy_signal"] = False
        gate["final_verdict"] = "🔴 RED LIGHT — Multiple reasons to avoid. Do not buy."
    else:
        gate["buy_signal"] = False
        gate["final_verdict"] = "🟠 WAIT — Mixed signals. Wait for better clarity."

    return gate


# ══════════════════════════════════════════════════════════════════════════════
# LIVE DATA
# ══════════════════════════════════════════════════════════════════════════════

def get_live_price(slug, company_name=""):
    """Fetch live price from Yahoo Finance using company slug."""
    try:
        import yfinance as yf

        tickers_to_try = [f"{slug.upper()}.NS"]
        # Name based variations
        if company_name:
            name_clean = company_name.replace(" ", "").replace(".", "").replace("Ltd", "").replace("Limited", "").upper()
            tickers_to_try.append(f"{name_clean}.NS")
            name_words = company_name.strip().split()
            if len(name_words) >= 2:
                abbrev = "".join(w[0] for w in name_words if w[0].isalpha()).upper()
                tickers_to_try.append(f"{abbrev}.NS")
        # Also try with no suffix for some tickers
        tickers_to_try.append(slug.upper())

        for t in tickers_to_try:
            try:
                tk = yf.Ticker(t)
                info = tk.info or {}
                price = (info.get("currentPrice") or info.get("regularMarketPrice")
                         or info.get("previousClose") or info.get("ask") or info.get("bid"))
                if price:
                    return {
                        "price": price,
                        "change": info.get("regularMarketChange"),
                        "change_pct": info.get("regularMarketChangePercent"),
                        "day_high": (info.get("dayHigh") or info.get("regularMarketDayHigh")),
                        "day_low": (info.get("dayLow") or info.get("regularMarketDayLow")),
                        "volume": (info.get("volume") or info.get("regularMarketVolume")),
                        "market_cap": info.get("marketCap"),
                        "source": f"Yahoo Finance ({t})",
                    }
            except (KeyError, TypeError):
                continue
    except ImportError:
        log.warning("yfinance not installed — live price unavailable")
    except Exception as e:
        log.warning("Live price fetch failed for %s: %s", slug, e)

    # Fallback: scrape from Screener.in company page
    try:
        url = f"{BASE}/company/{slug}/"
        r = req.get(url, headers=HEADERS, cookies=COOKIES, timeout=10)
        m = re.search(r'id="price"\s*[^>]*>\s*₹?\s*([\d.,]+)', r.text, re.I)
        if not m:
            m = re.search(r'<span[^>]*class="[^"]*price[^"]*"[^>]*>\s*₹?\s*([\d.,]+)', r.text, re.I)
        if not m:
            m = re.search(r'₹\s*([\d,]+\.\d{2})', r.text)
        if m:
            price = safe_float(m.group(1))
            if price:
                return {"price": price, "source": "Screener.in (live)"}
    except (req.RequestException, re.error) as e:
        log.warning("Screener.in price scrape failed for %s: %s", slug, e)
    return None


# ══════════════════════════════════════════════════════════════════════════════
# BUSINESS INFO SCRAPER
# ══════════════════════════════════════════════════════════════════════════════

def scrape_business_info(slug):
    """Scrape company page for business description, industry, and shareholding."""
    from bs4 import BeautifulSoup
    info = {"industry": None, "description": None, "market_cap": None,
            "promoter_holding": None, "institutional_holding": None}
    try:
        url = f"{BASE}/company/{slug}/"
        r = req.get(url, headers=HEADERS, cookies=COOKIES, timeout=15)
        if r.status_code != 200:
            return info

        soup = BeautifulSoup(r.text, "html.parser")

        # Industry — found in breadcrumb
        m = re.search(r'<a[^>]*href="/industry/[^"]*"[^>]*>([^<]+)</a>', r.text)
        if m:
            info["industry"] = m.group(1).strip()

        # Business description — meta description
        m = re.search(r'<meta[^>]*name="description"[^>]*content="([^"]+)"', r.text, re.I)
        if m:
            info["description"] = m.group(1).strip()

        # Market cap — use CSS selectors (reliable, no regex needed)
        for li in soup.select(".company-ratios li"):
            name_el = li.select_one(".name")
            num_el = li.select_one(".number")
            if not name_el or not num_el:
                continue
            key = name_el.get_text(strip=True)
            raw_val = num_el.get_text(strip=True).replace(",", "")
            if key == "Market Cap":
                info["market_cap"] = safe_float(raw_val)
            elif key == "ROE":
                info["institutional_holding"] = info.get("institutional_holding")  # keep existing

        # Promoter holding from shareholding table
        sh_sec = soup.find("section", id="shareholding")
        if sh_sec:
            table = sh_sec.find("table")
            if table:
                rows = table.find_all("tr")
                for row in rows:
                    cells = [c.get_text(strip=True) for c in row.find_all(["th", "td"])]
                    if cells and "promoter" in cells[0].lower():
                        if len(cells) >= 2:
                            info["promoter_holding"] = safe_float(cells[-1].replace("%", ""))
                        break

    except (req.RequestException, re.error, TypeError) as e:
        log.warning("scrape_business_info failed for %s: %s", slug, e)
    return info


def scrape_yahoo(slug):
    """Fetch supplementary data from Yahoo Finance for Indian stocks (.NS suffix).
    Returns dict with enriched metrics: industry, sector, shareholding, 52wk, beta, analyst target, etc."""
    result = {}
    try:
        import yfinance as yf
        ticker_obj = yf.Ticker(f"{slug.upper()}.NS")
        info = ticker_obj.info or {}
        if not info or info.get("trailingPegRatio") is None and info.get("shortName") is None:
            return result

        # Industry & sector (more reliable than Screener)
        result["industry"] = info.get("industry")
        result["sector"] = info.get("sector")
        result["country"] = info.get("country", "India")

        # Market data
        result["current_price_yf"] = info.get("currentPrice") or info.get("regularMarketPrice")
        result["market_cap_yf"] = round(info["marketCap"] / 1e7, 2) if info.get("marketCap") else None
        result["fifty_two_week_high"] = info.get("fiftyTwoWeekHigh")
        result["fifty_two_week_low"] = info.get("fiftyTwoWeekLow")
        result["fifty_day_avg"] = info.get("fiftyDayAverage")
        result["two_hundred_day_avg"] = info.get("twoHundredDayAverage")
        result["beta"] = info.get("beta")

        # Ratios
        result["trailing_pe"] = info.get("trailingPE")
        result["forward_pe"] = info.get("forwardPE")
        result["price_to_book"] = info.get("priceToBook")
        result["price_to_sales"] = info.get("priceToSalesTrailing12Months")
        result["trailing_eps"] = info.get("trailingEps")
        result["forward_eps"] = info.get("forwardEps")
        result["book_value_yf"] = info.get("bookValue")
        result["dividend_yield"] = info.get("dividendYield")

        # Profitability
        result["roe_yf"] = round(info["returnOnEquity"] * 100, 1) if info.get("returnOnEquity") else None
        result["operating_margin_yf"] = round(info["operatingMargins"] * 100, 1) if info.get("operatingMargins") else None
        result["profit_margin_yf"] = round(info["profitMargins"] * 100, 1) if info.get("profitMargins") else None

        # Growth
        result["revenue_growth"] = round(info["revenueGrowth"] * 100, 1) if info.get("revenueGrowth") else None
        result["earnings_growth"] = round(info["earningsGrowth"] * 100, 1) if info.get("earningsGrowth") else None

        # Financial health
        result["debt_to_equity_yf"] = info.get("debtToEquity")
        result["total_cash"] = round(info["totalCash"] / 1e7, 2) if info.get("totalCash") else None
        result["total_debt_yf"] = round(info["totalDebt"] / 1e7, 2) if info.get("totalDebt") else None
        result["free_cashflow_yf"] = round(info["freeCashflow"] / 1e7, 2) if info.get("freeCashflow") else None

        # Shareholding
        result["held_by_insiders"] = round(info["heldPercentInsiders"] * 100, 1) if info.get("heldPercentInsiders") else None
        result["held_by_institutions"] = round(info["heldPercentInstitutions"] * 100, 1) if info.get("heldPercentInstitutions") else None
        result["shares_outstanding"] = info.get("sharesOutstanding")

        # Analyst
        result["analyst_target"] = info.get("targetMeanPrice")
        result["analyst_recommendation"] = info.get("recommendationKey")

        # Enterprise value
        result["enterprise_value"] = round(info["enterpriseValue"] / 1e7, 2) if info.get("enterpriseValue") else None

    except ImportError:
        log.warning("yfinance not installed — skipping Yahoo enrichment")
    except Exception as e:
        log.warning("scrape_yahoo failed for %s: %s", slug, e)
    return result

# ══════════════════════════════════════════════════════════════════════════════
# VALUATION ENGINE
# ══════════════════════════════════════════════════════════════════════════════

def valuation_engine(eps, bvps, current_price, rev_cagr, pro_cagr,
                     ebitda_val, total_debt, cash, shares_outstanding, market="India"):
    results  = {}
    g_rate   = (pro_cagr or rev_cagr or 15) / 100
    AAA      = 7.2
    CUR      = "$" if market == "US" else "₹"
    CR_CONV  = 1     if market == "US" else 1e7   # Indian values in crores need ×1e7

    # Graham Number
    try:
        if eps and eps > 0 and bvps and bvps > 0:
            gn = math.sqrt(22.5 * eps * bvps)
            results["Graham Number"] = {
                "iv": round(gn, 2),
                "formula": f"√(22.5 × EPS {CUR}{eps:.1f} × BVPS {CUR}{bvps:.1f})",
                "weight": 0.20, "desc": "Benjamin Graham conservative formula"
            }
    except (TypeError, ValueError):
        pass

    # Graham Growth
    try:
        if eps and eps > 0 and pro_cagr is not None and pro_cagr > 0:
            g_pct = min(pro_cagr, 25)
            gg    = eps * (8.5 + 2 * g_pct) * 4.4 / AAA
            results["Graham Growth"] = {
                "iv": round(gg, 2),
                "formula": f"EPS {CUR}{eps:.1f} × (8.5 + 2×{g_pct:.0f}%) × 4.4 / {AAA}%",
                "weight": 0.25, "desc": "Growth-adjusted Graham formula"
            }
    except (TypeError, ValueError):
        pass

    # DCF
    try:
        if eps and eps > 0:
            r = 0.10 if market == "US" else 0.12
            g1 = max(min(g_rate, 0.25), 0)          # floor at 0%
            g2 = max(min(g_rate * 0.6, 0.15), 0)    # floor at 0%
            gt = 0.03 if market == "US" else 0.04
            fcf = eps
            pv  = 0
            for yr in range(1, 11):
                fcf  = fcf * (1 + (g1 if yr <= 5 else g2))
                pv  += fcf / (1 + r) ** yr
            tv   = fcf * (1 + gt) / (r - gt)
            dcf  = pv + tv / (1 + r) ** 10
            results["DCF (10yr)"] = {
                "iv": round(dcf, 2),
                "formula": f"r={r*100:.0f}%, g1={g1*100:.0f}%, g2={g2*100:.0f}%, gT={gt*100:.0f}%",
                "weight": 0.30, "desc": "10-year discounted cash flow"
            }
    except (TypeError, ValueError, ZeroDivisionError):
        pass

    # PEG
    try:
        if eps and eps > 0 and pro_cagr and pro_cagr > 0:
            fair_pe    = min(pro_cagr, 40)
            peg_iv     = fair_pe * eps
            peg_actual = (current_price / eps / pro_cagr) if current_price else None
            results["PEG Model"] = {
                "iv": round(peg_iv, 2),
                "formula": f"Fair P/E={fair_pe:.0f} (=growth%) × EPS {CUR}{eps:.1f}",
                "weight": 0.15, "desc": "Peter Lynch PEG=1 fair value",
                "peg_actual": round(peg_actual, 2) if peg_actual else None
            }
    except (TypeError, ValueError, ZeroDivisionError):
        pass

    # EV/EBITDA
    try:
        shares_norm = normalize_shares(shares_outstanding)
        if ebitda_val and ebitda_val > 0 and shares_norm and shares_norm > 0:
            EVX  = 14
            debt = total_debt or 0
            c    = cash or 0
            fair_price = ((EVX * ebitda_val - debt + c) * CR_CONV) / shares_norm
            ebitda_fmt = _fmt_usd(ebitda_val) if market == "US" else f"{CUR}{ebitda_val:.0f}Cr"
            results["EV/EBITDA"] = {
                "iv": round(fair_price, 2),
                "formula": f"14x × EBITDA {ebitda_fmt} − Debt + Cash / shares",
                "weight": 0.10, "desc": "Enterprise value to EBITDA multiple"
            }
    except (TypeError, ValueError, ZeroDivisionError):
        pass

    if not results or not current_price:
        return None

    total_w = sum(m["weight"] for m in results.values())
    wtd_iv  = round(sum(m["iv"] * m["weight"] for m in results.values()) / total_w, 2)
    upside  = round(((wtd_iv - current_price) / current_price) * 100, 1)
    margin  = round(((wtd_iv - current_price) / wtd_iv) * 100, 1)

    if upside >= 30:  v = ("UNDERVALUED",      "strong",  "🟢 Strong margin of safety — Good entry")
    elif upside >= 10: v = ("SLIGHTLY UNDER",   "mild",    "🟡 Reasonable entry point")
    elif upside >= -10: v= ("FAIRLY VALUED",    "fair",    "🟠 Fairly priced — Limited upside")
    elif upside >= -30: v= ("OVERVALUED",       "over",    "🔴 Wait for correction")
    else:               v= ("HIGHLY OVERVALUED","danger",  "🔴 Significant downside risk")

    peg_actual = results.get("PEG Model", {}).get("peg_actual")
    val_formulas = {}
    for name, md in results.items():
        key = f"val_{name.lower().replace(' ', '_').replace('(', '').replace(')', '').replace('/', '_')}"
        val_formulas[key] = md.get("formula", "")
    val_formulas["weighted_iv"] = f"Weighted Average:\n" + " + ".join([f"{m['weight']*100:.0f}% × {CUR}{m['iv']}" for m in results.values()])
    return {
        "models":        results,
        "weighted_iv":   wtd_iv,
        "current_price": current_price,
        "upside_pct":    upside,
        "margin_safety": margin,
        "val_label":     v[0],
        "val_class":     v[1],
        "val_verdict":   v[2],
        "peg_actual":    peg_actual,
        "val_formulas":  val_formulas,
    }


# ══════════════════════════════════════════════════════════════════════════════
# COMPREHENSIVE ANALYSIS ENGINE
# ══════════════════════════════════════════════════════════════════════════════

def _fmt_usd(val, dp=1):
    """Format a raw USD number as $XB / $XM / $XK for display."""
    if val is None:
        return None
    v = float(val)
    av = abs(v)
    sign = "-" if v < 0 else ""
    if av >= 1e12:
        return f"{sign}${av/1e12:.{dp}f}T"
    if av >= 1e9:
        return f"{sign}${av/1e9:.{dp}f}B"
    if av >= 1e6:
        return f"{sign}${av/1e6:.{dp}f}M"
    if av >= 1e3:
        return f"{sign}${av/1e3:.{dp}f}K"
    return f"{sign}${av:.{dp}f}"

def series_list(s, n=12):
    if not s: return []
    ks = list(s.keys())[-n:]
    return [{"year": k, "value": s[k]} for k in ks]

def series_list_us(s, n=12):
    if not s: return []
    ks = list(s.keys())[-n:]
    return [{"year": k, "value": s[k], "formatted": _fmt_usd(s[k])} for k in ks]

def analyze_full(company_name, slug, sheets, biz_info, yahoo_data=None):
    """Comprehensive analysis returning all sections including 10 advanced features."""
    # ── Extract all raw metrics ──
    _, sales      = get_metric(sheets, ["profit","loss","p&l","income"], ["sales","revenue","net sales"])
    _, profit     = get_metric(sheets, ["profit","loss","p&l","income"], ["net profit","profit after"])
    _, ebitda     = get_metric(sheets, ["profit","loss","p&l","income"], ["ebitda","operating profit"])
    _, opm_ser    = get_metric(sheets, ["profit","loss","p&l","income"], ["opm","operating margin"])
    _, borrowings = get_metric(sheets, ["balance","sheet"], ["borrowings","debt","total debt"])
    _, reserves   = get_metric(sheets, ["balance","sheet"], ["reserves"])
    _, cash_ser   = get_metric(sheets, ["balance","sheet"], ["cash","cash equivalent"])
    _, cfo        = get_metric(sheets, ["cash"], ["operating","from operations"])
    _, roe_ser    = get_metric(sheets, ["ratio","key"], ["roe","return on equity"])
    _, roce_ser   = get_metric(sheets, ["ratio","key"], ["roce","return on capital"])
    _, pe_ser     = get_metric(sheets, ["ratio","key"], ["p/e","price to earn","pe ratio"])
    _, eps_ser    = get_metric(sheets, ["ratio","key","profit"], ["eps","earning per share"])
    _, deb_eq     = get_metric(sheets, ["ratio","key","balance"], ["debt/equity","d/e ratio"])
    _, bvps_ser   = get_metric(sheets, ["ratio","key","balance"], ["book value","bvps"])
    _, price_ser  = get_metric(sheets, ["ratio","key"], ["price","cmp","share price","market price"])
    _, shares_ser = get_metric(sheets, ["ratio","key","balance"], ["shares","no. of shares","equity shares"])
    _, int_cov    = get_metric(sheets, ["ratio","key"], ["interest coverage"])
    _, net_margin = get_metric(sheets, ["ratio","key","profit"], ["net margin"])
    _, ebitda_mrg = get_metric(sheets, ["ratio","key","profit"], ["ebitda margin"])

    # ── Quarterly data for Historical Trends ──
    qdata = sheets.get("Quarters", {}).get("data", {})
    q_sales      = metric_match(qdata, ["sales","revenue","net sales"])
    q_profit     = metric_match(qdata, ["net profit","profit after"])
    q_ebitda     = metric_match(qdata, ["ebitda","operating profit"])
    q_borrowings = metric_match(qdata, ["borrowings","debt","total debt"])
    q_reserves   = metric_match(qdata, ["reserves"])
    q_cfo        = metric_match(qdata, ["operating","from operations"])

    # ── Latest values ──
    rev_cagr_v     = cagr(fv(sales), lv(sales), yrs(sales)) if sales else None
    pro_cagr_v     = cagr(fv(profit), lv(profit), yrs(profit)) if profit else None
    # Fallback: if first profit was negative, find first positive year for CAGR
    if pro_cagr_v is None and profit:
        pos_vals = {k: v for k, v in profit.items() if v and v > 0}
        if len(pos_vals) >= 2:
            pro_cagr_v = cagr(list(pos_vals.values())[0], list(pos_vals.values())[-1], len(pos_vals) - 1)
    latest_opm_v   = lv(opm_ser) or (lv(ebitda)/lv(sales)*100 if ebitda and sales and lv(sales) else None)
    latest_roe_v   = lv(roe_ser)
    latest_roce_v  = lv(roce_ser)
    latest_pe_v    = lv(pe_ser)
    latest_de_v    = lv(deb_eq)
    latest_eps_v   = lv(eps_ser)
    latest_bvps_v  = lv(bvps_ser)
    latest_price_v = lv(price_ser)
    latest_int_cov = lv(int_cov)
    latest_nm_v    = lv(net_margin)
    latest_em_v    = lv(ebitda_mrg)

    # ── Fallback: compute EPS, P/E, ROE, ROCE directly from raw data if computed ratios are missing ──
    if latest_eps_v is None and profit and shares_ser:
        np_latest = lv(profit)
        sh_latest = lv(shares_ser)
        if np_latest and sh_latest:
            sh_norm = normalize_shares(sh_latest)
            if sh_norm:
                latest_eps_v = round(np_latest / (sh_norm / 1e7), 2)
    if latest_bvps_v is None and shares_ser and reserves:
        eq_sc = get_metric(sheets, ["balance","sheet"], ["equity share capital","equity capital","share capital"])[1]
        eq_val = (lv(eq_sc) or 0) + (lv(reserves) or 0)
        sh_latest = lv(shares_ser)
        if eq_val > 0 and sh_latest:
            sh_norm = normalize_shares(sh_latest)
            if sh_norm:
                latest_bvps_v = round(eq_val / (sh_norm / 1e7), 2)
    if latest_pe_v is None and latest_eps_v and latest_price_v and latest_eps_v > 0:
        latest_pe_v = round(latest_price_v / latest_eps_v, 1)
    if latest_roe_v is None and profit and reserves:
        eq_sc = get_metric(sheets, ["balance","sheet"], ["equity share capital","equity capital","share capital"])[1]
        np_latest = lv(profit)
        eq_latest = (lv(eq_sc) or 0) + (lv(reserves) or 0)
        if np_latest and eq_latest > 0:
            latest_roe_v = round((np_latest / eq_latest) * 100, 1)
    if latest_roce_v is None and ebitda and borrowings and reserves:
        eq_sc = get_metric(sheets, ["balance","sheet"], ["equity share capital","equity capital","share capital"])[1]
        op_latest = lv(ebitda)
        eq_latest = (lv(eq_sc) or 0) + (lv(reserves) or 0)
        debt_latest = lv(borrowings) or 0
        capital_employed = eq_latest + debt_latest
        if op_latest and capital_employed > 0:
            latest_roce_v = round((op_latest / capital_employed) * 100, 1)
    if latest_de_v is None and borrowings and reserves:
        eq_sc = get_metric(sheets, ["balance","sheet"], ["equity share capital","equity capital","share capital"])[1]
        eq_latest = (lv(eq_sc) or 0) + (lv(reserves) or 0)
        debt_latest = lv(borrowings)
        if debt_latest and eq_latest > 0:
            latest_de_v = round(debt_latest / eq_latest, 2)

    # ── Prefer Meta Current Price over historical PRICE ──
    meta = sheets.get("Meta", {}).get("data", {}).get("Meta", {})
    meta_current_price = meta.get("Current Price")
    if meta_current_price is not None:
        latest_price_v = meta_current_price

    # ── Yahoo Finance fallback: fill any remaining None values ──
    if yahoo_data:
        if latest_price_v is None and yahoo_data.get("current_price_yf"):
            latest_price_v = yahoo_data["current_price_yf"]
        if latest_pe_v is None and yahoo_data.get("trailing_pe"):
            latest_pe_v = yahoo_data["trailing_pe"]
        if latest_eps_v is None and yahoo_data.get("trailing_eps"):
            latest_eps_v = yahoo_data["trailing_eps"]
        if latest_bvps_v is None and yahoo_data.get("book_value_yf"):
            latest_bvps_v = yahoo_data["book_value_yf"]
        if latest_roe_v is None and yahoo_data.get("roe_yf"):
            latest_roe_v = yahoo_data["roe_yf"]
        if latest_de_v is None and yahoo_data.get("debt_to_equity_yf"):
            latest_de_v = round(yahoo_data["debt_to_equity_yf"] / 100, 2)
        if latest_nm_v is None and yahoo_data.get("profit_margin_yf"):
            latest_nm_v = yahoo_data["profit_margin_yf"]
        if latest_em_v is None and yahoo_data.get("operating_margin_yf"):
            latest_em_v = yahoo_data["operating_margin_yf"]
        if latest_int_cov is None and opm_ser and int_d:
            op_l = lv(opm_ser)
            int_l = lv(int_d)
            if op_l and int_l and int_l > 0:
                latest_int_cov = round(op_l / int_l, 1)

    # ── Trend analysis ──
    debt_reduced = False
    if borrowings and len(borrowings) >= 2:
        vals = list(borrowings.values())
        debt_reduced = vals[-1] < vals[-2]
    cfo_positive = (lv(cfo) or 0) > 0 if cfo else None

    # ── Growth consistency ──
    sales_growth_years = []
    if sales and len(sales) >= 3:
        svals = list(sales.values())
        for i in range(1, len(svals)):
            if svals[i-1] and svals[i-1] > 0:
                sales_growth_years.append((svals[i] - svals[i-1]) / svals[i-1] * 100)
    profit_growth_years = []
    if profit and len(profit) >= 3:
        pvals = list(profit.values())
        for i in range(1, len(pvals)):
            if pvals[i-1] and pvals[i-1] > 0:
                profit_growth_years.append((pvals[i] - pvals[i-1]) / pvals[i-1] * 100)
    consistent_growth = (len(sales_growth_years) >= 3 and
                         sum(1 for g in sales_growth_years if g > 0) >= 2 and
                         sum(1 for g in profit_growth_years if g > 0) >= 2)

    # ── Peers from biz_info ──
    mcap = biz_info.get("market_cap")
    mcap_category = "Large Cap"
    if mcap:
        if mcap < 5000: mcap_category = "Small Cap"
        elif mcap < 20000: mcap_category = "Mid Cap"

    # ── Risk factors ──
    risks = []
    if latest_de_v and latest_de_v > 1.0:
        risks.append("High Debt-to-Equity ratio")
    if latest_de_v and latest_de_v > 0.5 and latest_de_v <= 1.0:
        risks.append("Moderate debt levels — monitor")
    if not cfo_positive:
        risks.append("Negative operating cash flow")
    if sales_growth_years and len(sales_growth_years) >= 2 and sales_growth_years[-1] < 0:
        risks.append("Declining revenue in latest year")
    if rev_cagr_v is None or (rev_cagr_v < 5):
        risks.append("Slow or negative revenue growth")
    if latest_int_cov and latest_int_cov < 2:
        risks.append("Low interest coverage — debt servicing risk")
    if not risks:
        risks.append("No major risks identified from available data")

    # ── Scores ──
    scores = {}
    scores["Revenue Growth"]   = score_cagr(rev_cagr_v)
    scores["Profit Growth"]    = score_cagr(pro_cagr_v)
    scores["Oper. Margin"]     = score_val(latest_opm_v, [(25,10),(20,9),(15,8),(10,6)])
    scores["ROE"]              = score_val(latest_roe_v, [(25,10),(15,8),(10,6)])
    scores["ROCE"]             = score_val(latest_roce_v, [(20,10),(15,8),(10,6)])
    scores["Debt Mgmt"]        = (9 if debt_reduced else 8 if (latest_de_v or 99) < 0.5
                                  else 6 if (latest_de_v or 99) < 1.0 else 4)
    scores["Cash Flow"]        = 8 if cfo_positive else 5
    scores["EPS Growth"]       = score_cagr(cagr(fv(eps_ser), lv(eps_ser), yrs(eps_ser))) if eps_ser else 5
    overall = round(sum(scores.values()) / len(scores), 1)

    if overall >= 8.5:   verdict, v_cls = "STRONG BUY",  "strong-buy"
    elif overall >= 7.5: verdict, v_cls = "BUY",          "buy"
    elif overall >= 6.5: verdict, v_cls = "ACCUMULATE",   "watch"
    elif overall >= 5.5: verdict, v_cls = "HOLD",         "hold"
    else:                verdict, v_cls = "AVOID",        "avoid"

    shares_units = lv(shares_ser) if lv(shares_ser) else None

    # ── Valuation ──
    val = valuation_engine(
        eps=latest_eps_v, bvps=latest_bvps_v,
        current_price=latest_price_v,
        rev_cagr=rev_cagr_v, pro_cagr=pro_cagr_v,
        ebitda_val=lv(ebitda),
        total_debt=lv(borrowings), cash=lv(cash_ser),
        shares_outstanding=shares_units,
    )

    # ── Growth potential assessment ──
    growth_potential = "Low"
    if rev_cagr_v and rev_cagr_v > 15: growth_potential = "High"
    elif rev_cagr_v and rev_cagr_v > 8: growth_potential = "Moderate"

    div_potential = "Low"
    if latest_eps_v and latest_price_v and (latest_eps_v / latest_price_v) > 0.02:
        div_potential = "Moderate"

    # ── Professional checklist ──
    checklist = {
        "Growth": scores["Revenue Growth"],
        "Profitability": round((scores["Oper. Margin"] + scores["ROE"] + scores["ROCE"]) / 3, 1),
        "Debt": scores["Debt Mgmt"],
        "Cash Flow": scores["Cash Flow"],
        "Valuation": score_val(val["upside_pct"] if val else None, [(30,9),(10,7),(-10,5),(-30,3)]) if val else 5,
        "Consistency": 8 if consistent_growth else 5,
    }

    # ═════════════════════════════════════════════════════════════════════
    # NEW ADVANCED ANALYSIS FEATURES
    # ═════════════════════════════════════════════════════════════════════

    # 1. Piotroski F-Score
    pl_section = sheets.get("Profit & Loss", {}).get("data", {})
    bs_section = sheets.get("Balance Sheet", {}).get("data", {})
    cf_section = sheets.get("Cash Flow", {}).get("data", {})
    key_section = sheets.get("Key", {}).get("data", {})

    piotroski, piotroski_details = piotroski_f_score(pl_section, bs_section, cf_section, key_section)

    # 2. Earnings Quality
    eq = earnings_quality_analysis(cfo, profit)

    # 3. Revenue Acceleration
    rev_acc = revenue_acceleration(sales)

    # 4. Altman Z-Score
    altman = altman_z_score(pl_section, bs_section, key_section, latest_price_v, shares_units)

    # 5. Entry Zones
    zones = entry_zones(latest_price_v, val)

    # 6. Magic Formula
    mf = magic_formula_rank(key_section, pl_section, bs_section, latest_price_v, shares_units)

    # 7. Promoter Trend
    promoter_info = biz_info.get("promoter_holding")
    promoter_trend = {
        "current_holding": promoter_info,
        "level": "High" if promoter_info and promoter_info > 60 else "Moderate" if promoter_info and promoter_info > 30 else "Low" if promoter_info else "N/A",
        "institutional_holding": biz_info.get("institutional_holding"),
    }

    # 8. Overall Signal
    osig = overall_signal(scores, overall, val, risks, piotroski, promoter_info)

    # 9. Interest Coverage (already computed, upgrade section)
    int_cov_years = list(key_section.get("Interest Coverage", {}).values())
    interest_coverage_analysis = {
        "latest": latest_int_cov,
        "trend": int_cov_years,
        "status": "Safe" if latest_int_cov and latest_int_cov > 3 else "Moderate" if latest_int_cov and latest_int_cov > 1.5 else "Concerning",
    }

    # 10. Bear/Base/Bull Targets
    if val and val.get("weighted_iv"):
        iv = val["weighted_iv"]
        cp = latest_price_v
        bbb = {
            "bear": round(min(cp * 0.85, iv * 0.7), 2) if cp and iv else None,
            "base": round(iv, 2),
            "bull": round(max(iv * 1.3, cp * 1.5), 2) if cp and iv else None,
            "description": {
                "bear": "Pessimistic — 30% below fair value or 15% below CMP",
                "base": "Fair value as per weighted intrinsic value",
                "bull": "Optimistic — 30% above fair value or 50% above CMP",
            }
        }
    else:
        bbb = None

    # ── Formula strings for double-click display ──
    formulas = {}
    pl_raw = sheets.get("Profit & Loss", {}).get("data", {})
    bs_raw = sheets.get("Balance Sheet", {}).get("data", {})

    # Get operating profit for formulas
    op_ser = metric_match(pl_raw, ["operating profit"])
    op_val = lv(op_ser)
    if op_val is None:
        pbt_s = metric_match(pl_raw, ["profit before tax", "pbt"])
        int_s_ = metric_match(pl_raw, ["interest"])
        dep_s = metric_match(pl_raw, ["depreciation", "dep"])
        oi_s_ = metric_match(pl_raw, ["other income"])
        p_v, i_v, d_v = lv(pbt_s), lv(int_s_), lv(dep_s)
        if p_v is not None and i_v is not None and d_v is not None:
            op_val = p_v + i_v + d_v
            oi_v_ = lv(oi_s_)
            if oi_v_ is not None: op_val -= oi_v_

    # Interest for coverage formula
    int_cov_raw = lv(metric_match(pl_raw, ["interest"]))

    # Equity for formulas
    eq_sc_f = get_metric(sheets, ["balance","sheet"], ["equity share capital","equity capital","share capital"])[1]
    eq_sc_v = lv(eq_sc_f) or 0
    res_v = lv(reserves) or 0
    eq_v = eq_sc_v + res_v

    # Normalized shares for formulas
    sh_norm_f = normalize_shares(shares_units) if shares_units else None

    def _cr(x):
        return f"₹{x:.1f}Cr" if x is not None else "N/A"

    # Revenue CAGR
    if sales and lv(sales) and fv(sales) and yrs(sales) > 0:
        f_s, l_s, ny = fv(sales), lv(sales), yrs(sales)
        formulas["rev_cagr"] = f"CAGR = ((Latest / First)^(1/Years) - 1) × 100\n= (({l_s:.1f} / {f_s:.1f})^(1/{ny}) - 1) × 100 = {rev_cagr_v}%"

    # Profit CAGR
    if profit and lv(profit) and fv(profit) and yrs(profit) > 0:
        f_p, l_p, ny = fv(profit), lv(profit), yrs(profit)
        formulas["pro_cagr"] = f"CAGR = ((Latest / First)^(1/Years) - 1) × 100\n= (({l_p:.1f} / {f_p:.1f})^(1/{ny}) - 1) × 100 = {pro_cagr_v}%"

    # EPS CAGR
    if eps_ser and lv(eps_ser) and fv(eps_ser) and yrs(eps_ser) > 0:
        eps_cagr_v = cagr(fv(eps_ser), lv(eps_ser), yrs(eps_ser))
        f_e, l_e, ny = fv(eps_ser), lv(eps_ser), yrs(eps_ser)
        formulas["eps_cagr"] = f"CAGR = ((Latest EPS / First EPS)^(1/Years) - 1) × 100\n= ((₹{l_e:.2f} / ₹{f_e:.2f})^(1/{ny}) - 1) × 100 = {eps_cagr_v}%"

    # OPM
    if latest_opm_v is not None and op_val is not None and lv(sales):
        s_v = lv(sales)
        formulas["latest_opm"] = f"OPM = (Operating Profit / Revenue) × 100\n= ({_cr(op_val)} / {_cr(s_v)}) × 100 = {latest_opm_v}%"

    # Net Margin
    if latest_nm_v is not None and lv(profit) and lv(sales):
        formulas["net_margin"] = f"Net Margin = (Net Profit / Revenue) × 100\n= ({_cr(lv(profit))} / {_cr(lv(sales))}) × 100 = {latest_nm_v}%"

    # EBITDA Margin
    if latest_em_v is not None and lv(ebitda) and lv(sales):
        formulas["ebitda_margin"] = f"EBITDA Margin = (EBITDA / Revenue) × 100\n= ({_cr(lv(ebitda))} / {_cr(lv(sales))}) × 100 = {latest_em_v}%"

    # ROE
    if latest_roe_v is not None and lv(profit) and eq_v > 0:
        formulas["latest_roe"] = f"ROE = (Net Profit / Shareholders' Equity) × 100\n= ({_cr(lv(profit))} / {_cr(eq_v)}) × 100 = {latest_roe_v}%"

    # ROCE
    if latest_roce_v is not None and op_val is not None:
        debt_v = lv(borrowings) or 0
        ce_val = eq_v + debt_v
        formulas["latest_roce"] = f"ROCE = (Operating Profit / (Equity + Borrowings)) × 100\n= ({_cr(op_val)} / ({_cr(eq_v)} + {_cr(debt_v)})) × 100 = {latest_roce_v}%"

    # P/E
    if latest_pe_v is not None and latest_price_v and latest_eps_v and latest_eps_v > 0:
        formulas["latest_pe"] = f"P/E = Price / EPS\n= (₹{latest_price_v:.2f} / ₹{latest_eps_v:.2f}) = {latest_pe_v}x"

    # D/E
    if latest_de_v is not None and lv(borrowings) is not None and eq_v > 0:
        formulas["latest_de"] = f"D/E = Total Borrowings / Shareholders' Equity\n= ({_cr(lv(borrowings))} / {_cr(eq_v)}) = {latest_de_v}"

    # Interest Coverage
    if latest_int_cov is not None and op_val is not None and int_cov_raw is not None and int_cov_raw > 0:
        formulas["interest_coverage"] = f"Interest Coverage = Operating Profit / Interest Expense\n= ({_cr(op_val)} / {_cr(int_cov_raw)}) = {latest_int_cov}x"

    # EPS
    if latest_eps_v is not None and lv(profit) and sh_norm_f:
        formulas["latest_eps"] = f"EPS = Net Profit / (Shares / 10⁷)\n= ({_cr(lv(profit))} / ({sh_norm_f/1e7:.1f}Cr / 10⁷)) = ₹{latest_eps_v}"

    # BVPS
    if latest_bvps_v is not None and eq_v > 0 and sh_norm_f:
        formulas["latest_bvps"] = f"BVPS = Shareholders' Equity / (Shares / 10⁷)\n= ({_cr(eq_v)} / ({sh_norm_f/1e7:.1f}Cr / 10⁷)) = ₹{latest_bvps_v}"

    # ── Build all sections (will add advanced sections after building the base) ──
    result = {
        "company_name": company_name,
        "generated_at": datetime.now().strftime("%d %b %Y %H:%M"),
        "slug": slug,
        "market": "India",

        "business": {
            "industry": biz_info.get("industry", "N/A"),
            "description": biz_info.get("description", ""),
            "market_cap": mcap,
            "market_cap_category": mcap_category if mcap else "N/A",
            "promoter_holding": biz_info.get("promoter_holding"),
            "institutional_holding": biz_info.get("institutional_holding"),
        },

        "metrics": {
            "rev_cagr":     round(rev_cagr_v, 1) if rev_cagr_v else None,
            "pro_cagr":     round(pro_cagr_v, 1) if pro_cagr_v else None,
            "latest_opm":   round(latest_opm_v, 1) if latest_opm_v else None,
            "latest_roe":   round(latest_roe_v, 1) if latest_roe_v else None,
            "latest_roce":  round(latest_roce_v, 1) if latest_roce_v else None,
            "latest_pe":    round(latest_pe_v, 1) if latest_pe_v else None,
            "latest_de":    round(latest_de_v, 2) if latest_de_v else None,
            "latest_eps":   round(latest_eps_v, 2) if latest_eps_v else None,
            "latest_bvps":  round(latest_bvps_v, 2) if latest_bvps_v else None,
            "latest_price": round(latest_price_v, 2) if latest_price_v else None,
            "debt_reduced": debt_reduced,
            "cfo_positive": cfo_positive,
            "interest_coverage": round(latest_int_cov, 1) if latest_int_cov else None,
            "net_margin": round(latest_nm_v, 1) if latest_nm_v else None,
            "ebitda_margin": round(latest_em_v, 1) if latest_em_v else None,
            # Yahoo Finance enriched fields
            "forward_pe": yahoo_data.get("forward_pe") if yahoo_data else None,
            "forward_eps": yahoo_data.get("forward_eps") if yahoo_data else None,
            "beta": yahoo_data.get("beta") if yahoo_data else None,
            "fifty_two_week_high": yahoo_data.get("fifty_two_week_high") if yahoo_data else None,
            "fifty_two_week_low": yahoo_data.get("fifty_two_week_low") if yahoo_data else None,
            "fifty_day_avg": yahoo_data.get("fifty_day_avg") if yahoo_data else None,
            "two_hundred_day_avg": yahoo_data.get("two_hundred_day_avg") if yahoo_data else None,
            "analyst_target": yahoo_data.get("analyst_target") if yahoo_data else None,
            "revenue_growth_yf": yahoo_data.get("revenue_growth") if yahoo_data else None,
        },

        "series": {
            "sales":      series_list(sales, 12),
            "profit":     series_list(profit, 12),
            "ebitda":     series_list(ebitda, 12),
            "borrowings": series_list(borrowings, 12),
            "reserves":   series_list(reserves, 12),
            "cfo":        series_list(cfo, 12),
        },

        "ratios": {
            "valuation": {
                "P/E": fmt(latest_pe_v, "x", 1),
                "PEG": fmt(val["peg_actual"], "", 2) if val and val.get("peg_actual") else "N/A",
                "P/BV": fmt(latest_price_v / latest_bvps_v, "x", 1) if latest_price_v and latest_bvps_v and latest_bvps_v > 0 else "N/A",
            },
            "profitability": {
                "ROE": fmt(latest_roe_v, "%", 1),
                "ROCE": fmt(latest_roce_v, "%", 1),
                "OPM": fmt(latest_opm_v, "%", 1),
                "Net Margin": fmt(latest_nm_v, "%", 1),
            },
            "debt": {
                "D/E": fmt(latest_de_v, "x", 2),
                "Interest Coverage": fmt(latest_int_cov, "x", 1),
            },
        },

        "balance_sheet": {
            "debt_trend": "Reducing" if debt_reduced else "Increasing",
            "debt_level": "High" if (latest_de_v or 0) > 1.0 else "Moderate" if (latest_de_v or 0) > 0.5 else "Low",
            "cash_reserves": lv(cash_ser),
            "borrowings_latest": lv(borrowings),
            "reserves_latest": lv(reserves),
        },

        "cash_flow": {
            "cfo_positive": cfo_positive,
            "cfo_latest": lv(cfo),
            "profit_cash_consistent": "Good" if cfo_positive and lv(cfo) and lv(profit) and abs(lv(cfo)/lv(profit)) > 0.5 else "Weak",
        },

        "growth": {
            "rev_cagr": rev_cagr_v,
            "pro_cagr": pro_cagr_v,
            "consistent_growth": "Yes" if consistent_growth else "No",
            "growth_potential": growth_potential,
            "eps_cagr": cagr(fv(eps_ser), lv(eps_ser), yrs(eps_ser)) if eps_ser else None,
        },

        "risks": risks,
        "val": val,
        "formulas": {**formulas, **(val.get("val_formulas") if val else {}), **(mf.get("formulas") if mf else {})},
        "scores": scores,
        "overall": overall,
        "verdict": verdict,
        "v_cls": v_cls,

        "long_term": {
            "growth_potential": growth_potential,
            "dividend_potential": div_potential,
            "multibagger_possibility": "High" if (rev_cagr_v and rev_cagr_v > 20 and latest_de_v and latest_de_v < 0.5) else "Moderate" if (rev_cagr_v and rev_cagr_v > 10) else "Low",
        },

        "checklist": checklist,
        "checklist_overall": round(sum(checklist.values()) / len(checklist), 1),

        # ── NEW ADVANCED SECTIONS ──
        "piotroski_fscore": {
            "score": piotroski,
            "max_score": 9,
            "details": piotroski_details,
            "rating": "Strong" if piotroski >= 7 else "Moderate" if piotroski >= 5 else "Weak",
        },

        "earnings_quality": eq,

        "revenue_acceleration": rev_acc,

        "altman_z": altman,

        "entry_zones": zones,

        "magic_formula": mf,

        "promoter_trend": promoter_trend,

        "overall_signal": osig,

        "interest_coverage_analysis": interest_coverage_analysis,

        "bear_base_bull": bbb,
    }

    # ═════════════════════════════════════════════════════════════════════
    # NEW: 20-POINT FRAMEWORK, RED FLAGS & BUY CONFIRMATION
    # ═════════════════════════════════════════════════════════════════════

    # Run comprehensive red flag detection
    result["red_flags"] = detect_red_flags(
        pl=pl_section, bs=bs_section, cf=cf_section, key=key_section,
        sales=sales, profit=profit, borrowings=borrowings, cfo=cfo,
        promoter_holding=biz_info.get("promoter_holding"),
        rev_cagr=rev_cagr_v, pro_cagr=pro_cagr_v,
        latest_de=latest_de_v, latest_pe=latest_pe_v,
        latest_roe=latest_roe_v, latest_roce=latest_roce_v,
        latest_int_cov=latest_int_cov, debt_reduced=debt_reduced,
        cfo_positive=cfo_positive, eps_ser=eps_ser,
        industry=biz_info.get("industry"),
    )

    # 20-Point Framework (pass the already-built result as reference)
    twenty_point = twenty_point_checklist(result)

    # Buy Confirmation Gate
    buy_gate = buy_confirmation_gate(twenty_point, result["red_flags"], val, overall)

    result["twenty_point_checklist"] = twenty_point
    result["buy_confirmation"] = buy_gate

    # Also add industry assessment
    result["industry_future"] = assess_industry_future(biz_info.get("industry"))
    result["management_quality"] = assess_management_quality(
        biz_info.get("promoter_holding"), 
        "Reducing" if debt_reduced else "Increasing",
        latest_roce_v, latest_roe_v
    )
    result["economic_sensitivity"] = assess_economic_sensitivity(
        biz_info.get("industry"), latest_de_v, latest_int_cov
    )
    result["institutional_assessment"] = assess_institutional_buying(
        biz_info.get("institutional_holding"), biz_info.get("promoter_holding")
    )

    # ═════════════════════════════════════════════════════════════════════
    # INVESTOR FRAMEWORK SCORING (Buffett, Lynch, Munger, Fisher, CAN SLIM)
    # ═════════════════════════════════════════════════════════════════════
    fw_buffett  = buffett_score(pl_section, bs_section, cf_section, key_section,
                    latest_roe_v, latest_roce_v, latest_opm_v, latest_de_v,
                    latest_int_cov, cfo_positive, rev_cagr_v, pro_cagr_v,
                    consistent_growth, debt_reduced, biz_info.get("industry"))
    fw_lynch    = lynch_score(pl_section, bs_section, key_section,
                    latest_pe_v, latest_eps_v, latest_de_v, latest_roe_v,
                    latest_opm_v, rev_cagr_v, pro_cagr_v, consistent_growth,
                    cfo_positive, promoter_info, biz_info.get("institutional_holding"))
    fw_munger   = munger_score(latest_roe_v, latest_roce_v, latest_opm_v,
                    latest_de_v, latest_int_cov, rev_cagr_v, pro_cagr_v,
                    consistent_growth, debt_reduced, cfo_positive,
                    piotroski, altman.get("zone", ""), biz_info.get("industry"),
                    result.get("red_flags", {}).get("red_flags", []),
                    promoter_info)
    fw_fisher   = fisher_score(pl_section, bs_section, key_section,
                    latest_opm_v, latest_roe_v, latest_roce_v, latest_de_v,
                    rev_cagr_v, pro_cagr_v, consistent_growth, cfo_positive,
                    debt_reduced, promoter_info, biz_info.get("institutional_holding"),
                    biz_info.get("industry"), mcap)
    fw_canslim  = canslim_score(pl_section, bs_section, key_section,
                    latest_eps_v, latest_pe_v, latest_roe_v, latest_opm_v,
                    rev_cagr_v, pro_cagr_v, eps_ser, sales, profit,
                    latest_price_v, shares_units, promoter_info,
                    biz_info.get("institutional_holding"), mcap,
                    biz_info.get("industry"), consistent_growth,
                    rev_acc.get("status", "Insufficient data"))

    framework_results = [fw_buffett, fw_lynch, fw_munger, fw_fisher, fw_canslim]
    fw_overall = overall_investor_score(framework_results)

    result["investor_frameworks"] = {
        "frameworks": framework_results,
        "overall": fw_overall,
    }

    return result


# ══════════════════════════════════════════════════════════════════════════════
# US STOCK MARKET ANALYSIS (via Yahoo Finance)
# ══════════════════════════════════════════════════════════════════════════════

US_SECTORS = {
    "Technology": ["AAPL", "MSFT", "NVDA", "AVGO", "ORCL", "CRM", "ADBE", "INTC", "AMD", "CSCO", "IBM", "QCOM"],
    "Healthcare": ["JNJ", "UNH", "PFE", "ABBV", "MRK", "TMO", "ABT", "MDT", "BMY", "LLY", "ISRG", "SYK", "BSX"],
    "Financials": ["JPM", "BAC", "WFC", "C", "GS", "MS", "V", "MA", "AXP", "BLK", "SCHW", "USB", "PNC", "BK", "COF"],
    "Consumer Cyclical": ["TSLA", "AMZN", "HD", "MCD", "NKE", "SBUX", "LOW", "TGT", "GM", "F", "LULU", "TJX", "ROST", "MAR", "BKNG"],
    "Energy": ["XOM", "CVX", "COP", "EOG", "SLB", "OXY", "PSX", "VLO", "MPC", "HAL", "BKR", "HES", "DVN", "FANG", "CTRA"],
    "Communication Services": ["GOOGL", "META", "NFLX", "DIS", "CMCSA", "T", "VZ", "EA", "TTWO", "WBD", "PARA", "LYV", "MTCH", "ROKU", "SNAP"],
    "Industrials": ["CAT", "DE", "GE", "HON", "UPS", "BA", "LMT", "RTX", "MMM", "UNP", "CSX", "EMR", "ETN", "ITW", "GD"],
    "Utilities": ["NEE", "DUK", "SO", "D", "AEP", "EXC", "SRE", "PEG", "ED", "XEL", "WEC", "ES", "AWK", "EIX", "DTE"],
    "Real Estate": ["PLD", "AMT", "CCI", "EQIX", "PSA", "O", "DLR", "SPG", "WY", "AVB", "EQR", "WELL", "ARE", "BXP", "VICI"],
    "Basic Materials": ["LIN", "SHW", "APD", "ECL", "FCX", "NEM", "DOW", "DD", "PPG", "NUE", "STLD", "ALB", "MOS", "CF", "EMN"],
    "Biotechnology": ["REGN", "VRTX", "BIIB", "ILMN", "MRNA", "ALXN", "INCY", "UTHR", "EXEL", "NBIX", "ALNY", "SRPT", "BGNE", "AMGN", "GILD"],
    "Consumer Defensive": ["PG", "KO", "PEP", "WMT", "COST", "CL", "KMB", "MDLZ", "KHC", "GIS", "CPB", "SJM", "CAG", "HRL", "K"],
}

# ── US stock autocomplete cache ──
_us_search_cache = {}
_us_search_cache_time = 0

def _us_label(ticker, info=None):
    """Build display label for a US stock."""
    if info is None:
        try:
            import yfinance as yf
            info = yf.Ticker(ticker).info or {}
        except (ImportError, KeyError, TypeError):
            info = {}
    name = info.get("longName") or info.get("shortName") or ticker
    sector = info.get("sector", "")
    industry = info.get("industry", "")
    price = info.get("currentPrice") or info.get("regularMarketPrice") or info.get("previousClose")
    price_str = f"${price:.2f}" if price else ""
    return {"ticker": ticker, "name": name, "sector": sector, "industry": industry, "price": price_str}


def us_search(q):
    """Search US stocks by ticker or company name using Yahoo Finance search API."""
    import yfinance as yf
    results = []
    q_upper = q.strip().upper()
    q_lower = q.strip().lower()
    NON_US = {"SAO", "NEO", "MX", "FRA", "EPA", "LON", "TSE", "HKG", "SHA", "SYD",
              "NSE", "BSE", "JKT", "SET", "KLSE", "PSE", "VIH", "SGO", "BUE",
              "MEX", "CSE", "STU", "MUN", "DUS", "BER", "HAM", "GER", "AMS",
              "BRU", "MIL", "TOR", "VAN", "CNQ", "ASX", "NZE", "OSL", "STO",
              "CPH", "WSE", "PRA", "BUD", "IST", "WSE", "RIS", "LIT", "LAT",
              "TAL", "UKX", "FSX", "SWX", "CDI", "IOB", "NOK", "OMC", "ICX",
              "KSC", "TWO", "TAI", "KUW", "ADX", "DFM", "QAT", "BAH", "EGX",
              "JSE", "Cairo", "DFM", "ADX"}

    # 1. Yahoo Finance search API (fast, single HTTP call)
    try:
        import requests as _rq
        sr = _rq.get(
            "https://query2.finance.yahoo.com/v1/finance/search",
            params={"q": q, "quotesCount": 10, "newsCount": 0,
                    "listsCount": 0, "types": "equity"},
            headers={"User-Agent": "Mozilla/5.0"},
            timeout=5,
        )
        if sr.ok:
            for item in sr.json().get("quotes", []):
                sym = item.get("symbol", "")
                exch = item.get("exchange", "")
                if item.get("quoteType") == "EQUITY" and sym:
                    if exch in NON_US:
                        continue
                    name = item.get("longname") or item.get("shortname") or sym
                    sector = item.get("sector", "")
                    industry = item.get("industry", "")
                    results.append({
                        "ticker": sym,
                        "name": name,
                        "sector": sector,
                        "industry": industry,
                        "price": "",
                        "exchange": exch,
                    })
                    if len(results) >= 6:
                        break
    except Exception:
        pass

    # 2. Fallback: direct ticker lookup if search returned nothing
    if not results:
        try:
            tk = yf.Ticker(q_upper)
            info = tk.info or {}
            if info.get("longName") or info.get("shortName"):
                results.append(_us_label(q_upper, info))
        except (KeyError, TypeError):
            pass

    return results[:6]




def fetch_us_financials(ticker):
    """Fetch US stock data from Yahoo Finance, return in sections format compatible with scoring engines."""
    import yfinance as yf
    tk = yf.Ticker(ticker)
    info = tk.info or {}

    def _to_num(v):
        try:
            if isinstance(v, complex): return round(v.real, 2)
            if isinstance(v, (int, float)): return round(float(v), 2)
            return round(float(v), 2)
        except (TypeError, ValueError, OverflowError):
            return None

    def collect(fin_df):
        """Convert yfinance DataFrame to {metric: {year: val}} — sorted oldest-first."""
        d = {}
        if fin_df is None or fin_df.empty:
            return d
        for idx in fin_df.index:
            series = fin_df.loc[idx]
            yr_vals = {}
            for col, val in zip(series.index, series.values):
                if val is not None and pd.notna(val):
                    yr = str(col.year) if hasattr(col, 'year') else str(col)
                    num = _to_num(val)
                    if num is not None:
                        yr_vals[yr] = num
            if yr_vals:
                d[str(idx)] = dict(sorted(yr_vals.items()))
        return d

    pl_raw = collect(tk.financials)
    bs_raw = collect(tk.balance_sheet)
    cf_raw = collect(tk.cashflow)

    # Build sections dict
    pl_section = {"data": {}, "years": []}
    bs_section = {"data": {}, "years": []}
    cf_section = {"data": {}, "years": []}
    key_section = {"data": {}, "years": []}

    # Map yfinance financials to our metric names
    rev_map = {"Total Revenue": "Sales", "Revenue": "Sales", "Operating Revenue": "Sales"}
    for yf_key, our_key in rev_map.items():
        if yf_key in pl_raw:
            pl_section["data"][our_key] = pl_raw[yf_key]
            break
    if "Sales" not in pl_section["data"]:
        # Try keyword match
        for k, v in pl_raw.items():
            kl = k.lower()
            if any(w in kl for w in ["total revenue", "revenue", "sales", "operating revenue"]):
                pl_section["data"]["Sales"] = v
                break

    np_map = {"Net Income": "Net profit", "Net Income From Continuing Operation Net Minority Interest": "Net profit"}
    for yf_key, our_key in np_map.items():
        if yf_key in pl_raw:
            pl_section["data"][our_key] = pl_raw[yf_key]
            break
    if "Net profit" not in pl_section["data"]:
        for k, v in pl_raw.items():
            kl = k.lower()
            if any(w in kl for w in ["net income", "net profit", "net earnings"]):
                if "discontinued" not in kl and "comprehensive" not in kl:
                    pl_section["data"]["Net profit"] = v
                    break

    for yf_key, our_key in [("EBITDA", "EBITDA"), ("EBIT", "Operating Profit")]:
        if yf_key in pl_raw:
            pl_section["data"][our_key] = pl_raw[yf_key]
    if "Operating Profit" not in pl_section["data"]:
        for k, v in pl_raw.items():
            kl = k.lower()
            if any(w in kl for w in ["operating income", "ebit"]):
                pl_section["data"]["Operating Profit"] = v
                break

    if "Interest Expense" in pl_raw:
        pl_section["data"]["Interest"] = {k: abs(v) if v else 0 for k, v in pl_raw["Interest Expense"].items()}
    if "Depreciation And Amortization" in pl_raw:
        pl_section["data"]["Depreciation"] = {k: abs(v) if v else 0 for k, v in pl_raw["Depreciation And Amortization"].items()}

    # Balance sheet
    for yf_key, our_key in [("Total Debt", "Borrowings"), ("Stockholders Equity", "Reserves"),
                              ("Total Equity Gross Minority Interest", "Equity"),
                              ("Cash And Cash Equivalents", "Cash"),
                              ("Cash And Cash Equivalents, Reported", "Cash"),
                              ("Current Assets", "Current Assets"),
                              ("Current Liabilities", "Current Liabilities"),
                              ("Inventory", "Inventory"),
                              ("Accounts Receivable", "Receivables"),
                              ("Total Assets", "Total Assets")]:
        if yf_key in bs_raw:
            bs_section["data"][our_key] = bs_raw[yf_key]
    # Try keyword fallback for key BS items
    if "Borrowings" not in bs_section["data"]:
        for k, v in bs_raw.items():
            kl = k.lower()
            if any(w in kl for w in ["total debt", "borrowings"]):
                bs_section["data"]["Borrowings"] = v
                break
    if "Cash" not in bs_section["data"]:
        for k, v in bs_raw.items():
            kl = k.lower()
            if any(w in kl for w in ["cash and cash equivalents", "cash", "cash equivalents"]):
                if "restricted" not in kl:
                    bs_section["data"]["Cash"] = v
                    break
    if "Reserves" not in bs_section["data"]:
        for k, v in bs_raw.items():
            kl = k.lower()
            if any(w in kl for w in ["stockholders equity", "total equity", "shareholders equity"]):
                bs_section["data"]["Reserves"] = v
                break
    if "No. of Equity Shares" not in bs_section["data"]:
        shares = info.get("sharesOutstanding")
        if shares:
            bs_section["data"]["No. of Equity Shares"] = {"Latest": float(shares)}

    # Cash flow items
    for yf_key, our_key in [("Free Cash Flow", "Free Cash Flow"),
                              ("Capital Expenditure", "Capex"),
                              ("Operating Cash Flow", "Cash from Operating Activity")]:
        if yf_key in cf_raw:
            cf_section["data"][our_key] = {k: abs(v) if v else 0 for k, v in cf_raw[yf_key].items()}
    if "Cash from Operating Activity" not in cf_section["data"]:
        cfo = info.get("operatingCashflow")
        if cfo:
            cf_section["data"]["Cash from Operating Activity"] = {"Latest": float(cfo)}
    if "Capex" not in cf_section["data"]:
        for k, v in cf_raw.items():
            kl = k.lower()
            if any(w in kl for w in ["capital expenditure", "capex"]):
                cf_section["data"]["Capex"] = {k2: abs(v2) if v2 else 0 for k2, v2 in v.items()}
                break

    # Key ratios from info dict
    cur_price = info.get("currentPrice") or info.get("regularMarketPrice") or info.get("previousClose")
    if cur_price:
        key_section["data"]["Price"] = {"Latest": float(cur_price)}
    if info.get("trailingEps"):
        key_section["data"]["EPS"] = {"Latest": float(info["trailingEps"])}
    if info.get("bookValue") and info["bookValue"] > 0:
        key_section["data"]["BVPS"] = {"Latest": float(info["bookValue"])}
    if info.get("trailingPE"):
        key_section["data"]["P/E"] = {"Latest": float(info["trailingPE"])}
    if info.get("returnOnEquity") is not None:
        key_section["data"]["ROE"] = {"Latest": round(float(info["returnOnEquity"]) * 100, 1)}
    if info.get("priceToBook") and info["priceToBook"] > 0:
        key_section["data"]["P/B"] = {"Latest": float(info["priceToBook"])}
    if info.get("debtToEquity") is not None:
        key_section["data"]["Debt/Equity"] = {"Latest": round(float(info["debtToEquity"]) / 100, 2)}
    if info.get("currentRatio"):
        key_section["data"]["Current Ratio"] = {"Latest": float(info["currentRatio"])}
    if info.get("profitMargins") is not None:
        key_section["data"]["Net Margin"] = {"Latest": round(float(info["profitMargins"]) * 100, 1)}
    if info.get("operatingMargins") is not None:
        key_section["data"]["OPM"] = {"Latest": round(float(info["operatingMargins"]) * 100, 1)}
    if info.get("freeCashflow"):
        key_section["data"]["Free Cash Flow"] = {"Latest": float(info["freeCashflow"])}
    if info.get("revenueGrowth") is not None:
        key_section["data"]["Revenue Growth"] = {"Latest": round(float(info["revenueGrowth"]) * 100, 1)}

    # Build all_years from sales data
    sales_data = pl_section["data"].get("Sales", {})
    pl_years = sorted(sales_data.keys(), key=lambda y: (len(y), y)) if sales_data else []
    bs_years = sorted(bs_section["data"].get("Borrowings", {}).keys(), key=lambda y: (len(y), y)) if bs_section["data"].get("Borrowings") else pl_years

    pl_section["years"] = pl_years
    bs_section["years"] = bs_years
    key_section["years"] = ["Latest"]

    return {
        "Profit & Loss": pl_section,
        "Balance Sheet": bs_section,
        "Cash Flow": cf_section,
        "Key": key_section,
        "_info": info,
        "_ticker": ticker,
    }


def analyze_us_stock(ticker):
    """Full analysis of a US stock using the same investor frameworks as Indian stocks."""
    sections = fetch_us_financials(ticker)
    info = sections.get("_info", {})

    pl = sections.get("Profit & Loss", {}).get("data", {})
    bs = sections.get("Balance Sheet", {}).get("data", {})
    cf = sections.get("Cash Flow", {}).get("data", {})
    key = sections.get("Key", {}).get("data", {})

    def lv(d): return list(d.values())[-1] if d else None
    def fv(d): return list(d.values())[0] if d else None
    def yrs(d): return max(len(d) - 1, 1) if d else 1

    sales_d = pl.get("Sales", {})
    profit_d = pl.get("Net profit", {})
    ebitda_d = pl.get("EBITDA", {})
    op_d = pl.get("Operating Profit", {})
    int_d = pl.get("Interest", {})
    dep_d = pl.get("Depreciation", {})
    borrowings_d = bs.get("Borrowings", {})
    reserves_d = bs.get("Reserves", {})
    cash_d = bs.get("Cash", {})
    shares_d = bs.get("No. of Equity Shares", {})
    cfo_d = cf.get("Cash from Operating Activity", {})
    fcf_d = cf.get("Free Cash Flow", {})

    eps_ser = key.get("EPS", {})

    rev_cagr_v = cagr(fv(sales_d), lv(sales_d), yrs(sales_d)) if sales_d else None
    pro_cagr_v = cagr(fv(profit_d), lv(profit_d), yrs(profit_d)) if profit_d else None
    eps_cagr_v = cagr(fv(eps_ser), lv(eps_ser), yrs(eps_ser)) if eps_ser else None

    latest_price_v = lv(key.get("Price", {}))
    latest_eps_v = lv(key.get("EPS", {}))
    latest_bvps_v = lv(key.get("BVPS", {}))
    latest_pe_v = lv(key.get("P/E", {}))
    latest_roe_v = lv(key.get("ROE", {}))
    latest_de_v = lv(key.get("Debt/Equity", {}))
    latest_opm_v = lv(key.get("OPM", {}))
    latest_nm_v = lv(key.get("Net Margin", {}))

    latest_em_v = None
    if ebitda_d and sales_d:
        ebitda_l = lv(ebitda_d)
        sales_l = lv(sales_d)
        if ebitda_l and sales_l and sales_l > 0:
            latest_em_v = round((ebitda_l / sales_l) * 100, 1)
    if latest_em_v is None and info.get("operatingMargins") is not None:
        latest_em_v = round(float(info["operatingMargins"]) * 100, 1)

    latest_int_cov = None
    if op_d and int_d:
        op_l = lv(op_d)
        int_l = lv(int_d)
        if op_l and int_l and int_l > 0:
            latest_int_cov = round(op_l / int_l, 1)
    latest_roce_v = None
    if op_d and borrowings_d and reserves_d:
        op_l = lv(op_d)
        debt_l = lv(borrowings_d) or 0
        eq_l = lv(reserves_d) or 0
        ce = eq_l + debt_l
        if op_l and ce > 0:
            latest_roce_v = round((op_l / ce) * 100, 1)

    debt_reduced = False
    if borrowings_d and len(borrowings_d) >= 2:
        vals = list(borrowings_d.values())
        debt_reduced = vals[-1] < vals[-2]
    cfo_val = lv(cfo_d) or info.get("operatingCashflow")
    cfo_positive = (cfo_val or 0) > 0 if cfo_val is not None else None

    sales_vals = list(sales_d.values())
    profit_vals = list(profit_d.values())
    sales_growth_years = []
    for i in range(1, len(sales_vals)):
        if sales_vals[i-1] and sales_vals[i-1] > 0:
            sales_growth_years.append((sales_vals[i] - sales_vals[i-1]) / sales_vals[i-1] * 100)
    profit_growth_years = []
    for i in range(1, len(profit_vals)):
        if profit_vals[i-1] and profit_vals[i-1] > 0:
            profit_growth_years.append((profit_vals[i] - profit_vals[i-1]) / profit_vals[i-1] * 100)
    consistent_growth = (len(sales_growth_years) >= 3 and
                         sum(1 for g in sales_growth_years if g > 0) >= 2 and
                         sum(1 for g in profit_growth_years if g > 0) >= 2)

    mcap = info.get("marketCap")
    mcap_b = round(mcap / 1e9, 2) if mcap else None
    mcap_category = "N/A"
    if mcap_b:
        if mcap_b >= 200: mcap_category = "Mega Cap"
        elif mcap_b >= 10: mcap_category = "Large Cap"
        elif mcap_b >= 2: mcap_category = "Mid Cap"
        else: mcap_category = "Small Cap"

    promoter_info = None
    inst_holding = info.get("heldPercentInstitutions")
    if inst_holding is not None:
        inst_holding = round(inst_holding * 100, 1)

    scores = {}
    scores["Revenue Growth"] = score_cagr(rev_cagr_v)
    scores["Profit Growth"] = score_cagr(pro_cagr_v)
    scores["Oper. Margin"] = score_val(latest_opm_v, [(25,10),(20,9),(15,8),(10,6)])
    scores["ROE"] = score_val(latest_roe_v, [(25,10),(15,8),(10,6)])
    scores["ROCE"] = score_val(latest_roce_v, [(20,10),(15,8),(10,6)])
    scores["Debt Mgmt"] = (9 if debt_reduced else 8 if (latest_de_v or 99) < 0.5
                           else 6 if (latest_de_v or 99) < 1.0 else 4)
    scores["Cash Flow"] = 8 if cfo_positive else 5
    scores["EPS Growth"] = score_cagr(eps_cagr_v) if eps_cagr_v else 5
    overall = round(sum(scores.values()) / len(scores), 1)

    if overall >= 8.5: verdict, v_cls = "STRONG BUY", "strong-buy"
    elif overall >= 7.5: verdict, v_cls = "BUY", "buy"
    elif overall >= 6.5: verdict, v_cls = "ACCUMULATE", "watch"
    elif overall >= 5.5: verdict, v_cls = "HOLD", "hold"
    else: verdict, v_cls = "AVOID", "avoid"

    company_name = info.get("longName") or info.get("shortName") or ticker
    sector = info.get("sector", "N/A")
    industry = info.get("industry", "N/A")

    val = None
    if latest_price_v and latest_eps_v and latest_eps_v > 0:
        val = valuation_engine(
            eps=latest_eps_v, bvps=latest_bvps_v,
            current_price=latest_price_v,
            rev_cagr=rev_cagr_v, pro_cagr=pro_cagr_v,
            ebitda_val=lv(ebitda_d),
            total_debt=lv(borrowings_d), cash=lv(cash_d),
            shares_outstanding=info.get("sharesOutstanding"),
            market="US",
        )

    biz_info = {
        "industry": industry,
        "sector": sector,
        "description": info.get("longBusinessSummary", ""),
        "market_cap": mcap_b,
        "market_cap_category": mcap_category,
        "promoter_holding": None,
        "institutional_holding": inst_holding,
    }

    risks = []
    if latest_de_v and latest_de_v > 1.0: risks.append("High Debt-to-Equity ratio")
    if cfo_positive is False: risks.append("Negative operating cash flow")
    if rev_cagr_v is None or (rev_cagr_v < 5): risks.append("Slow or negative revenue growth")
    if latest_int_cov and latest_int_cov < 2: risks.append("Low interest coverage")

    piotroski, piotroski_details = piotroski_f_score(pl, bs, cf, key)
    altman = altman_z_score(pl, bs, key, latest_price_v, info.get("sharesOutstanding"))
    eq = earnings_quality_analysis(cfo_d, profit_d)
    rev_acc = revenue_acceleration(sales_d)
    zones = entry_zones(latest_price_v, val)
    mf = magic_formula_rank(key, pl, bs, latest_price_v, info.get("sharesOutstanding"))

    checklist = {
        "Growth": scores["Revenue Growth"],
        "Profitability": round((scores["Oper. Margin"] + scores["ROE"] + scores["ROCE"]) / 3, 1),
        "Debt": scores["Debt Mgmt"],
        "Cash Flow": scores["Cash Flow"],
        "Valuation": score_val(val["upside_pct"] if val else None, [(30,9),(10,7),(-10,5),(-30,3)]) if val else 5,
        "Consistency": 8 if consistent_growth else 5,
    }

    growth_potential = "Low"
    if rev_cagr_v and rev_cagr_v > 15: growth_potential = "High"
    elif rev_cagr_v and rev_cagr_v > 8: growth_potential = "Moderate"

    div_potential = "Low"
    if latest_eps_v and latest_price_v and (latest_eps_v / latest_price_v) > 0.02:
        div_potential = "Moderate"

    debt_level = "High" if (latest_de_v or 0) > 1.0 else "Moderate" if (latest_de_v or 0) > 0.5 else "Low"

    osig = overall_signal(scores, overall, val, risks, piotroski, promoter_info)

    interest_coverage_analysis = {
        "latest": latest_int_cov,
        "trend": [],
        "status": "Safe" if latest_int_cov and latest_int_cov > 3 else "Moderate" if latest_int_cov and latest_int_cov > 1.5 else "Concerning",
    }

    if val and val.get("weighted_iv"):
        iv = val["weighted_iv"]
        cp = latest_price_v
        bbb = {
            "bear": round(min(cp * 0.85, iv * 0.7), 2) if cp and iv else None,
            "base": round(iv, 2),
            "bull": round(max(iv * 1.3, cp * 1.5), 2) if cp and iv else None,
            "description": {
                "bear": "Pessimistic — 30% below fair value or 15% below CMP",
                "base": "Fair value as per weighted intrinsic value",
                "bull": "Optimistic — 30% above fair value or 50% above CMP",
            }
        }
    else:
        bbb = None

    fw_buffett = buffett_score(pl, bs, cf, key,
                    latest_roe_v, latest_roce_v, latest_opm_v, latest_de_v,
                    latest_int_cov, cfo_positive, rev_cagr_v, pro_cagr_v,
                    consistent_growth, debt_reduced, industry)
    fw_lynch = lynch_score(pl, bs, key,
                    latest_pe_v, latest_eps_v, latest_de_v, latest_roe_v,
                    latest_opm_v, rev_cagr_v, pro_cagr_v, consistent_growth,
                    cfo_positive, promoter_info, inst_holding)
    fw_munger = munger_score(latest_roe_v, latest_roce_v, latest_opm_v,
                    latest_de_v, latest_int_cov, rev_cagr_v, pro_cagr_v,
                    consistent_growth, debt_reduced, cfo_positive,
                    piotroski, altman.get("zone", ""), industry,
                    [], promoter_info)
    fw_fisher = fisher_score(pl, bs, key,
                    latest_opm_v, latest_roe_v, latest_roce_v, latest_de_v,
                    rev_cagr_v, pro_cagr_v, consistent_growth, cfo_positive,
                    debt_reduced, promoter_info, inst_holding, industry, mcap_b)
    fw_canslim = canslim_score(pl, bs, key,
                    latest_eps_v, latest_pe_v, latest_roe_v, latest_opm_v,
                    rev_cagr_v, pro_cagr_v, eps_ser, sales_d, profit_d,
                    latest_price_v, info.get("sharesOutstanding"), promoter_info,
                    inst_holding, mcap_b, industry, consistent_growth,
                    rev_acc.get("status", "Insufficient data"))

    framework_results = [fw_buffett, fw_lynch, fw_munger, fw_fisher, fw_canslim]
    fw_overall = overall_investor_score(framework_results)

    result = {
        "company_name": company_name,
        "ticker": ticker,
        "market": "US",
        "generated_at": datetime.now().strftime("%d %b %Y %H:%M"),
        "business": biz_info,
        "metrics": {
            "rev_cagr": round(rev_cagr_v, 1) if rev_cagr_v else None,
            "pro_cagr": round(pro_cagr_v, 1) if pro_cagr_v else None,
            "latest_opm": round(latest_opm_v, 1) if latest_opm_v else None,
            "latest_roe": round(latest_roe_v, 1) if latest_roe_v else None,
            "latest_roce": round(latest_roce_v, 1) if latest_roce_v else None,
            "latest_pe": round(latest_pe_v, 1) if latest_pe_v else None,
            "latest_de": round(latest_de_v, 2) if latest_de_v else None,
            "latest_eps": round(latest_eps_v, 2) if latest_eps_v else None,
            "latest_bvps": round(latest_bvps_v, 2) if latest_bvps_v else None,
            "latest_price": round(latest_price_v, 2) if latest_price_v else None,
            "debt_reduced": debt_reduced,
            "cfo_positive": cfo_positive,
            "interest_coverage": round(latest_int_cov, 1) if latest_int_cov else None,
            "net_margin": round(latest_nm_v, 1) if latest_nm_v else None,
            "ebitda_margin": round(latest_em_v, 1) if latest_em_v else None,
            "forward_pe": info.get("forwardPE"),
            "forward_eps": info.get("forwardEps"),
            "beta": info.get("beta"),
            "fifty_two_week_high": info.get("fiftyTwoWeekHigh"),
            "fifty_two_week_low": info.get("fiftyTwoWeekLow"),
            "fifty_day_avg": info.get("fiftyDayAverage"),
            "two_hundred_day_avg": info.get("twoHundredDayAverage"),
            "analyst_target": info.get("targetMeanPrice"),
            "revenue_growth_yf": round(info["revenueGrowth"] * 100, 1) if info.get("revenueGrowth") else None,
        },
        "series": {
            "sales": series_list_us(sales_d, 12),
            "profit": series_list_us(profit_d, 12),
            "ebitda": series_list_us(ebitda_d, 12),
            "borrowings": series_list_us(borrowings_d, 12),
            "reserves": series_list_us(reserves_d, 12),
            "cfo": series_list_us(cfo_d, 12),
        },
        "ratios": {
            "valuation": {
                "P/E": fmt(latest_pe_v, "x", 1),
                "PEG": fmt(val.get("peg_actual") if val else None, "", 2) if val and val.get("peg_actual") else "N/A",
                "P/BV": fmt(latest_price_v / latest_bvps_v, "x", 1) if latest_price_v and latest_bvps_v and latest_bvps_v > 0 else "N/A",
            },
            "profitability": {
                "ROE": fmt(latest_roe_v, "%", 1),
                "ROCE": fmt(latest_roce_v, "%", 1),
                "OPM": fmt(latest_opm_v, "%", 1),
                "Net Margin": fmt(latest_nm_v, "%", 1),
            },
            "debt": {
                "D/E": fmt(latest_de_v, "x", 2),
                "Interest Coverage": fmt(latest_int_cov, "x", 1),
            },
        },
        "balance_sheet": {
            "debt_trend": "Reducing" if debt_reduced else "Increasing",
            "debt_level": debt_level,
            "cash_reserves": lv(cash_d),
            "borrowings_latest": lv(borrowings_d),
            "reserves_latest": lv(reserves_d),
            "cash_reserves_fmt": _fmt_usd(lv(cash_d)),
            "borrowings_latest_fmt": _fmt_usd(lv(borrowings_d)),
            "reserves_latest_fmt": _fmt_usd(lv(reserves_d)),
        },
        "cash_flow": {
            "cfo_positive": cfo_positive,
            "cfo_latest": lv(cfo_d),
            "cfo_latest_fmt": _fmt_usd(lv(cfo_d)),
            "profit_cash_consistent": "Good" if cfo_positive and lv(cfo_d) and lv(profit_d) and abs(lv(cfo_d)/lv(profit_d)) > 0.5 else "Weak",
        },
        "growth": {
            "rev_cagr": rev_cagr_v,
            "pro_cagr": pro_cagr_v,
            "consistent_growth": "Yes" if consistent_growth else "No",
            "growth_potential": growth_potential,
            "eps_cagr": eps_cagr_v,
        },
        "long_term": {
            "growth_potential": growth_potential,
            "dividend_potential": div_potential,
            "multibagger_possibility": "High" if (rev_cagr_v and rev_cagr_v > 20 and latest_de_v and latest_de_v < 0.5) else "Moderate" if (rev_cagr_v and rev_cagr_v > 10) else "Low",
        },
        "checklist": checklist,
        "checklist_overall": round(sum(checklist.values()) / len(checklist), 1),
        "val": val,
        "risks": risks,
        "scores": scores,
        "overall": overall,
        "verdict": verdict,
        "v_cls": v_cls,
        "piotroski_fscore": {"score": piotroski, "max_score": 9, "details": piotroski_details,
                             "rating": "Strong" if piotroski >= 7 else "Moderate" if piotroski >= 5 else "Weak"},
        "earnings_quality": eq,
        "revenue_acceleration": rev_acc,
        "altman_z": altman,
        "entry_zones": zones,
        "magic_formula": mf,
        "overall_signal": osig,
        "industry_future": assess_industry_future(industry),
        "economic_sensitivity": assess_economic_sensitivity(industry, latest_de_v, latest_int_cov),
        "institutional_assessment": assess_institutional_buying(inst_holding, promoter_info),
        "interest_coverage_analysis": interest_coverage_analysis,
        "bear_base_bull": bbb,
        "investor_frameworks": {"frameworks": framework_results, "overall": fw_overall},
    }

    result["red_flags"] = detect_red_flags(
        pl=pl, bs=bs, cf=cf, key=key,
        sales=sales_d, profit=profit_d, borrowings=borrowings_d, cfo=cfo_d,
        promoter_holding=None,
        rev_cagr=rev_cagr_v, pro_cagr=pro_cagr_v,
        latest_de=latest_de_v, latest_pe=latest_pe_v,
        latest_roe=latest_roe_v, latest_roce=latest_roce_v,
        latest_int_cov=latest_int_cov, debt_reduced=debt_reduced,
        cfo_positive=cfo_positive, eps_ser=eps_ser,
        industry=industry,
    )

    twenty_point = twenty_point_checklist(result)
    buy_gate = buy_confirmation_gate(twenty_point, result["red_flags"], val, overall)
    result["twenty_point_checklist"] = twenty_point
    result["buy_confirmation"] = buy_gate
    result["management_quality"] = assess_management_quality(
        None, "Reducing" if debt_reduced else "Increasing",
        latest_roce_v, latest_roe_v
    )

    return result


def us_sector_analysis(sector_name):
    """Analyze all tickers in a sector and return ranked results."""
    sector_key = None
    for key in US_SECTORS:
        if key.lower() == sector_name.lower():
            sector_key = key
            break
    if not sector_key:
        return {"error": f"Sector '{sector_name}' not found. Available: {', '.join(US_SECTORS.keys())}"}

    tickers = US_SECTORS[sector_key]
    results = []
    errors = []
    for ticker in tickers:
        try:
            data = analyze_us_stock(ticker)
            results.append({
                "ticker": ticker,
                "name": data["company_name"],
                "overall_score": data.get("investor_frameworks", {}).get("overall", {}).get("overall_score", data["overall"]),
                "verdict": data["verdict"],
                "overall_fundamentals": data["overall"],
                "frameworks": data.get("investor_frameworks", {}).get("frameworks", []),
                "metrics": {
                    "pe": data["metrics"]["latest_pe"],
                    "roe": data["metrics"]["latest_roe"],
                    "rev_cagr": data["metrics"]["rev_cagr"],
                    "pro_cagr": data["metrics"]["pro_cagr"],
                    "de": data["metrics"]["latest_de"],
                    "price": data["metrics"]["latest_price"],
                }
            })
        except Exception as e:
            errors.append({"ticker": ticker, "error": str(e)})

    results.sort(key=lambda r: r["overall_score"], reverse=True)
    return {
        "sector": sector_key,
        "total_stocks": len(tickers),
        "analyzed": len(results),
        "errors": errors,
        "ranked_stocks": results,
    }


# ── SECTOR ANALYSIS CACHE ────────────────────────────────────────────────────
SECTOR_CACHE = {}
SECTOR_CACHE_TTL = timedelta(minutes=30)

def get_cached_sector_analysis(sector_name):
    now = datetime.now()
    key = sector_name.lower()
    if key in SECTOR_CACHE:
        entry = SECTOR_CACHE[key]
        if now - entry["ts"] < SECTOR_CACHE_TTL:
            return entry["data"]
    data = us_sector_analysis(sector_name)
    SECTOR_CACHE[key] = {"ts": now, "data": data}
    return data

def find_sector_for_ticker(ticker):
    """Find which predefined US sector a ticker belongs to."""
    t = ticker.upper()
    for sector, tickers in US_SECTORS.items():
        if t in tickers:
            return sector
    return None

def compute_sector_context(ticker, framework_results, overall_fundamentals_score):
    """Compute peer ranking & comparison data within the stock's sector."""
    sector = find_sector_for_ticker(ticker)
    if not sector:
        return None

    sector_data = get_cached_sector_analysis(sector)
    if "error" in sector_data:
        return None

    ranked = sector_data.get("ranked_stocks", [])
    total = sector_data.get("total_stocks", 0)
    analyzed = sector_data.get("analyzed", 0)

    my_rank = next((i + 1 for i, s in enumerate(ranked) if s["ticker"].upper() == ticker.upper()), None)

    # Sector averages for framework scores
    sector_avg_scores = {}
    if ranked:
        fw_keys = ["Buffett (Buffettology)", "Lynch (GARP)", "Munger (Mental Models)",
                    "Fisher (Scuttlebutt)", "CAN SLIM (O'Neil)"]
        for fk in fw_keys:
            scores = []
            for s in ranked:
                for fw in s.get("frameworks", []):
                    if fw["framework"] == fk and fw["score"] is not None:
                        scores.append(fw["score"])
            sector_avg_scores[fk] = round(sum(scores) / len(scores), 1) if scores else None

    # Stock's own framework scores for comparison
    my_fw_scores = {}
    for fw in framework_results or []:
        my_fw_scores[fw["framework"]] = fw["score"]

    # Compute percentiles
    percentiles = {}
    for fk in my_fw_scores:
        scores = []
        for s in ranked:
            for fw in s.get("frameworks", []):
                if fw["framework"] == fk and fw["score"] is not None:
                    scores.append(fw["score"])
        my_score = my_fw_scores[fk]
        if scores and my_score is not None:
            below = sum(1 for x in scores if x < my_score)
            pct = round(below / len(scores) * 100) if scores else 50
            percentiles[fk] = pct

    return {
        "sector": sector,
        "total_stocks": total,
        "analyzed_stocks": analyzed,
        "rank": my_rank,
        "percentile": round((1 - (my_rank or 1) / max(total, 1)) * 100) if my_rank else None,
        "sector_averages": sector_avg_scores,
        "my_scores": my_fw_scores,
        "percentiles": percentiles,
        "overall_sector_leader": ranked[0]["ticker"] if ranked else None,
        "overall_sector_leader_score": ranked[0]["overall_score"] if ranked else None,
    }


# ══════════════════════════════════════════════════════════════════════════════
# ROUTES
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/search")
def search():
    rl = _rate_limit()
    if rl:
        return rl
    q = request.args.get("q", "").strip()
    if not q:
        return jsonify({"error": "Empty query"}), 400
    if len(q) > 100:
        return jsonify({"error": "Query too long"}), 400
    try:
        url = f"{BASE}/api/company/search/?q={req.utils.quote(q)}&v=3"
        r   = req.get(url, headers=HEADERS, cookies=COOKIES, timeout=REQUEST_TIMEOUT)
        r.raise_for_status()
        return jsonify(r.json()[:6])
    except req.RequestException as e:
        log.warning("/api/search failed: %s", e)
        return jsonify({"error": "Search unavailable"}), 502


@app.route("/api/analyze")
def analyze_route():
    rl = _rate_limit()
    if rl:
        return rl
    slug = request.args.get("slug", "").strip().lower()
    name = request.args.get("name", slug)
    if not slug:
        return jsonify({"error": "No slug"}), 400
    if not SLUG_RE.match(slug):
        return jsonify({"error": "Invalid slug format"}), 400
    if len(name) > 200:
        name = name[:200]

    # Check cache
    now = datetime.now()
    cache_entry = _ANALYZE_CACHE.get(slug)
    if cache_entry and (now - cache_entry[0]) < _ANALYZE_TTL:
        log.info("Cache hit for /api/analyze slug=%s", slug)
        resp = jsonify(cache_entry[1])
        resp.headers["Cache-Control"] = "public, max-age=900"
        return resp

    try:
        # get export url
        page_url = f"{BASE}/company/{slug}/consolidated/"
        r        = req.get(page_url, headers=HEADERS, cookies=COOKIES, timeout=REQUEST_TIMEOUT)
        match    = re.search(r'(?:href|formaction)="(/user/company/export/\d+/)"', r.text)
        if not match:
            page_url = f"{BASE}/company/{slug}/"
            r        = req.get(page_url, headers=HEADERS, cookies=COOKIES, timeout=REQUEST_TIMEOUT)
            match    = re.search(r'(?:href|formaction)="(/user/company/export/\d+/)"', r.text)
        if not match:
            log.info("No export URL for slug=%s, falling back to HTML scrape", slug)
            sheets = scrape_html(slug)
            if not sheets or "Profit & Loss" not in sheets:
                resp = jsonify({"error": "Could not fetch data from Screener.in. Try uploading an Excel file instead."})
                resp.headers["Cache-Control"] = "no-store"
                return resp, 400
            compute_ratios(sheets)
            biz_info, yahoo_data = _enrich_with_yahoo(slug, _merge_meta_into_biz(sheets, scrape_business_info(slug)))
            result = analyze_full(name, slug, sheets, biz_info, yahoo_data)
            _ANALYZE_CACHE[slug] = (now, result)
            resp = jsonify(result)
            resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
            resp.headers["Pragma"] = "no-cache"
            return resp

        export_url = BASE + match.group(1)
        # GET the export page to obtain CSRF token
        r2 = req.get(export_url, headers=HEADERS, cookies=COOKIES, timeout=(5, 15))
        r2.raise_for_status()
        csrf_match = re.search(r'csrfmiddlewaretoken[^>]+value="([^"]+)"', r2.text)
        if not csrf_match:
            resp = jsonify({"error": "Couldn't fetch data. The Screener.in session may have expired. Try uploading an Excel file instead."})
            resp.headers["Cache-Control"] = "no-store"
            return resp, 401
        csrf_token = csrf_match.group(1)
        # Merge COOKIES with any new cookies from the response (csrftoken)
        all_cookies = {**COOKIES, **{k: v for k, v in r2.cookies.items()}}
        # POST to export URL with CSRF token to download the actual Excel file
        r3 = req.post(export_url, headers=HEADERS, cookies=all_cookies,
                      data={"csrfmiddlewaretoken": csrf_token, "next": page_url},
                      timeout=(5, 15))
        r3.raise_for_status()

        # Validate response is actually an Excel file, not a login/register redirect
        content = r3.content
        is_excel = (content[:2] == b'PK'                    # .xlsx (ZIP)
                    or content[:4] == b'\xd0\xcf\x11\xe0'  # .xls (OLE2)
                    or (len(content) > 500
                        and (b'worksheet' in content[:2000].lower()
                             or b'<?xml' in content[:200].lower())))
        if not is_excel:
            log.info("Excel export failed for slug=%s, falling back to HTML scrape", slug)
            sheets = scrape_html(slug)
            if not sheets or "Profit & Loss" not in sheets:
                resp = jsonify({"error": "Could not fetch data from Screener.in. Try uploading an Excel file instead."})
                resp.headers["Cache-Control"] = "no-store"
                return resp, 401
            compute_ratios(sheets)
            biz_info, yahoo_data = _enrich_with_yahoo(slug, _merge_meta_into_biz(sheets, scrape_business_info(slug)))
            result = analyze_full(name, slug, sheets, biz_info, yahoo_data)
        else:
            sheets = parse_excel(BytesIO(r3.content))
            biz_info, yahoo_data = _enrich_with_yahoo(slug, scrape_business_info(slug))
            result = analyze_full(name, slug, sheets, biz_info, yahoo_data)

        # Store in cache
        _ANALYZE_CACHE[slug] = (now, result)
        # Evict old entries
        for k, (ts, _) in list(_ANALYZE_CACHE.items()):
            if (now - ts) > _ANALYZE_TTL:
                del _ANALYZE_CACHE[k]

        resp = jsonify(result)
        resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        resp.headers["Pragma"] = "no-cache"
        return resp
    except req.RequestException as e:
        log.warning("/api/analyze failed for slug=%s: %s", slug, e)
        resp = jsonify({"error": "Failed to fetch data from Screener.in. Please try again."})
        resp.headers["Cache-Control"] = "no-store"
        return resp, 502
    except Exception as e:
        log.exception("/api/analyze unexpected error for slug=%s", slug)
        resp = jsonify({"error": "Analysis failed. Please try again or upload an Excel file."})
        resp.headers["Cache-Control"] = "no-store"
        return resp, 500


@app.route("/api/analyze/excel", methods=["POST"])
def analyze_excel():
    """Analyze from uploaded Screener.in Excel file."""
    rl = _rate_limit()
    if rl:
        return rl
    try:
        if 'file' not in request.files:
            return jsonify({"error": "No file selected. Please choose an Excel file from your computer."}), 400

        file = request.files['file']
        if not file.filename:
            return jsonify({"error": "The selected file appears to be empty. Please try a different file."}), 400

        if not file.filename.lower().endswith(('.xlsx', '.xls')):
            return jsonify({"error": "Please upload an Excel file (.xlsx or .xls) exported from Screener.in."}), 400

        # Read Excel file
        excel_bytes = file.read()
        if len(excel_bytes) < 1000:
            return jsonify({"error": "The file seems too small or empty. Please check the file and try again."}), 400

        # Cap remote size
        if len(excel_bytes) > MAX_UPLOAD_BYTES:
            return jsonify({"error": f"File too large. Max {MAX_UPLOAD_BYTES // (1024*1024)} MB."}), 413

        # Parse Excel
        sheets = parse_excel(BytesIO(excel_bytes))

        if not sheets:
            return jsonify({"error": "Couldn't read this Excel file. Make sure you're uploading a Screener.in data sheet export (not a custom spreadsheet)."}), 400

        # Try to extract company info
        meta = sheets.get("Meta", {}).get("data", {}).get("Meta", {})
        company_name = meta.get("Company Name") or "Unknown"
        industry = meta.get("Industry") or "Unknown"

        # Get business info
        biz_info = {
            "industry": industry,
            "description": meta.get("Description", ""),
            "market_cap": meta.get("Market Cap"),
            "promoter_holding": meta.get("Promoter Holding"),
            "institutional_holding": meta.get("Institutional Holding"),
        }

        # If no company name from Meta, try to get from filename
        if company_name == "Unknown":
            company_name = re.sub(r'[^a-zA-Z0-9\s\-.]', ' ',
                                  file.filename.rsplit('.', 1)[0].replace('_', ' ')).strip()
            if company_name.lower().endswith('consolidated'):
                company_name = company_name[:-12].strip()

        # Generate slug from company name
        slug = re.sub(r'[^a-z0-9\-]', '-', company_name.lower().strip())[:80]

        # Run analysis
        result = analyze_full(company_name, slug, sheets, biz_info)

        resp = jsonify(result)
        resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        resp.headers["Pragma"] = "no-cache"
        return resp

    except Exception as e:
        log.exception("/api/analyze/excel unexpected error")
        resp = jsonify({"error": "Analysis failed. Make sure the file is a valid Screener.in export."})
        resp.headers["Cache-Control"] = "no-store"
        return resp, 500


@app.route("/api/pdf")
def generate_pdf():
    """Quick PDF download — pass same params as /analyze."""
    rl = _rate_limit()
    if rl:
        return rl
    slug = request.args.get("slug", "").strip().lower()
    name = request.args.get("name", slug)
    if not slug:
        return jsonify({"error": "No slug"}), 400
    if not SLUG_RE.match(slug):
        return jsonify({"error": "Invalid slug format"}), 400

    # Try cache first, otherwise analyze
    now = datetime.now()
    cache_entry = _ANALYZE_CACHE.get(slug)
    if cache_entry and (now - cache_entry[0]) < _ANALYZE_TTL:
        data = cache_entry[1]
    else:
        try:
            # Reuse the analysis logic directly instead of calling test_client
            page_url = f"{BASE}/company/{slug}/consolidated/"
            r = req.get(page_url, headers=HEADERS, cookies=COOKIES, timeout=REQUEST_TIMEOUT)
            match = re.search(r'(?:href|formaction)="(/user/company/export/\d+/)"', r.text)
            if not match:
                page_url = f"{BASE}/company/{slug}/"
                r = req.get(page_url, headers=HEADERS, cookies=COOKIES, timeout=REQUEST_TIMEOUT)
                match = re.search(r'(?:href|formaction)="(/user/company/export/\d+/)"', r.text)
            if not match:
                log.info("No export URL for PDF slug=%s, falling back to HTML scrape", slug)
                sheets = scrape_html(slug)
                if not sheets or "Profit & Loss" not in sheets:
                    return jsonify({"error": "Could not fetch data from Screener.in."}), 400
                compute_ratios(sheets)
                biz_info, yahoo_data = _enrich_with_yahoo(slug, _merge_meta_into_biz(sheets, scrape_business_info(slug)))
                data = analyze_full(name, slug, sheets, biz_info, yahoo_data)
            else:
                export_url = BASE + match.group(1)
                r2 = req.get(export_url, headers=HEADERS, cookies=COOKIES, timeout=(5, 15))
                r2.raise_for_status()
                csrf_match = re.search(r'csrfmiddlewaretoken[^>]+value="([^"]+)"', r2.text)
                if not csrf_match:
                    return jsonify({"error": "Screener.in session expired."}), 401
                all_cookies = {**COOKIES, **{k: v for k, v in r2.cookies.items()}}
                r3 = req.post(export_url, headers=HEADERS, cookies=all_cookies,
                              data={"csrfmiddlewaretoken": csrf_match.group(1), "next": page_url},
                              timeout=(5, 15))
                r3.raise_for_status()
                content = r3.content
                is_excel = (content[:2] == b'PK'
                            or content[:4] == b'\xd0\xcf\x11\xe0'
                            or (len(content) > 500
                                and (b'worksheet' in content[:2000].lower()
                                     or b'<?xml' in content[:200].lower())))
                if not is_excel:
                    log.info("Excel export failed for PDF slug=%s, falling back to HTML scrape", slug)
                    sheets = scrape_html(slug)
                    if not sheets or "Profit & Loss" not in sheets:
                        return jsonify({"error": "Could not fetch data from Screener.in."}), 401
                    compute_ratios(sheets)
                    biz_info, yahoo_data = _enrich_with_yahoo(slug, _merge_meta_into_biz(sheets, scrape_business_info(slug)))
                    data = analyze_full(name, slug, sheets, biz_info, yahoo_data)
                else:
                    sheets = parse_excel(BytesIO(r3.content))
                    biz_info, yahoo_data = _enrich_with_yahoo(slug, _merge_meta_into_biz(sheets, scrape_business_info(slug)))
                    data = analyze_full(name, slug, sheets, biz_info, yahoo_data)
        except req.RequestException as e:
            log.warning("/api/pdf fetch failed for slug=%s: %s", slug, e)
            return jsonify({"error": "Failed to fetch data from Screener.in."}), 502
        except Exception as e:
            log.exception("/api/pdf analysis failed for slug=%s", slug)
            return jsonify({"error": "Analysis failed."}), 500

    try:
        buf = build_pdf(data)
        buf.seek(0)
        safe = re.sub(r"[^\w]", "_", name)[:100]
        return send_file(buf, as_attachment=True,
                         download_name=f"{safe}_analysis.pdf",
                         mimetype="application/pdf")
    except Exception as e:
        log.exception("/api/pdf build failed for slug=%s", slug)
        return jsonify({"error": "PDF generation failed."}), 500


@app.route("/api/live")
def live_price():
    rl = _rate_limit()
    if rl:
        return rl
    slug = request.args.get("slug", "").strip().lower()
    name = request.args.get("name", slug)
    if not slug:
        return jsonify({"error": "No slug"}), 400
    if not SLUG_RE.match(slug):
        return jsonify({"error": "Invalid slug format"}), 400

    # Check cache
    now = datetime.now()
    cache_entry = _LIVE_CACHE.get(slug)
    if cache_entry and (now - cache_entry[0]) < _LIVE_TTL:
        resp = jsonify(cache_entry[1])
        resp.headers["Cache-Control"] = "public, max-age=120"
        return resp

    live = get_live_price(slug, name)
    if live:
        _LIVE_CACHE[slug] = (now, live)
        resp = jsonify(live)
    else:
        resp = jsonify({"error": "Live price not available", "price": None})
    resp.headers["Cache-Control"] = "no-store"
    return resp


# ══════════════════════════════════════════════════════════════════════════════
# US STOCK ROUTES
# ══════════════════════════════════════════════════════════════════════════════

@app.route("/api/us/search")
def us_search_route():
    rl = _rate_limit()
    if rl:
        return rl
    q = request.args.get("q", "").strip()
    if not q or len(q) < 1:
        return jsonify([])
    if len(q) > 50:
        return jsonify({"error": "Query too long"}), 400
    try:
        results = us_search(q)
        return jsonify(results)
    except Exception as e:
        log.warning("/api/us/search failed: %s", e)
        return jsonify({"error": "Search unavailable"}), 502


@app.route("/api/us/analyze")
def us_analyze_route():
    rl = _rate_limit()
    if rl:
        return rl
    ticker = request.args.get("ticker", "").strip().upper()
    if not ticker:
        return jsonify({"error": "No ticker provided"}), 400
    if not TICKER_RE.match(ticker):
        return jsonify({"error": "Invalid ticker format"}), 400
    try:
        now = datetime.now()
        cache_entry = _US_ANALYZE_CACHE.get(ticker)
        if cache_entry and (now - cache_entry[0]) < _US_ANALYZE_TTL:
            log.info("Cache hit for /api/us/analyze ticker=%s", ticker)
            resp = jsonify(cache_entry[1])
            resp.headers["Cache-Control"] = "public, max-age=900"
            return resp

        result = analyze_us_stock(ticker)
        fw_results = result.get("investor_frameworks", {}).get("frameworks", [])
        result["sector_context"] = compute_sector_context(
            ticker, fw_results, result.get("overall")
        )
        _US_ANALYZE_CACHE[ticker] = (now, result)
        for k, (ts, _) in list(_US_ANALYZE_CACHE.items()):
            if (now - ts) > _US_ANALYZE_TTL:
                del _US_ANALYZE_CACHE[k]
        resp = jsonify(result)
        resp.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        return resp
    except Exception as e:
        log.exception("/api/us/analyze failed for ticker=%s", ticker)
        return jsonify({"error": f"Analysis failed for {ticker}."}), 500


@app.route("/api/us/sector")
def us_sector_route():
    rl = _rate_limit()
    if rl:
        return rl
    sector = request.args.get("sector", "").strip()
    if not sector:
        return jsonify({"error": "No sector provided", "available": list(US_SECTORS.keys())}), 400
    try:
        result = us_sector_analysis(sector)
        resp = jsonify(result)
        resp.headers["Cache-Control"] = "no-store"
        return resp
    except Exception as e:
        log.exception("/api/us/sector failed for sector=%s", sector)
        return jsonify({"error": "Sector analysis failed."}), 500


@app.route("/api/us/sectors")
def us_sectors_list():
    return jsonify(list(US_SECTORS.keys()))


# ══════════════════════════════════════════════════════════════════════════════
# PDF BUILDER
# ══════════════════════════════════════════════════════════════════════════════

C_DARK  = colors.HexColor("#0D1117")
C_ACC   = colors.HexColor("#1F6FEB")
C_GREEN = colors.HexColor("#2EA043")
C_RED   = colors.HexColor("#DA3633")
C_AMBER = colors.HexColor("#D29922")
C_LIGHT = colors.HexColor("#F6F8FA")
C_MID   = colors.HexColor("#30363D")
C_TEXT  = colors.HexColor("#24292F")
C_WHITE = colors.white
C_TEAL  = colors.HexColor("#0891B2")

def ST(name, **kw): return ParagraphStyle(name, **kw)

def kv_tbl(pairs, cw=(5.5*cm, 2.8*cm)):
    data = [[Paragraph(f"<b>{k}</b>", ST("k",fontName="Helvetica-Bold",fontSize=9,textColor=C_TEXT)),
             Paragraph(str(v), ST("v",fontName="Helvetica",fontSize=9,textColor=C_TEXT,alignment=TA_RIGHT))]
            for k, v in pairs]
    t = Table(data, colWidths=cw)
    t.setStyle(TableStyle([
        ("ROWBACKGROUNDS",(0,0),(-1,-1),[C_WHITE,C_LIGHT]),
        ("GRID",(0,0),(-1,-1),0.3,colors.HexColor("#D0D7DE")),
        ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),
        ("LEFTPADDING",(0,0),(-1,-1),7),("RIGHTPADDING",(0,0),(-1,-1),7),
    ]))
    return t

def series_tbl_pdf(series_list_data, title, n=7):
    if not series_list_data: return []
    items = series_list_data[-n:]
    def c(txt, bold=False, color=C_TEXT):
        return Paragraph(str(txt), ST("c",fontName="Helvetica-Bold" if bold else "Helvetica",
                                      fontSize=8.5,textColor=color,alignment=TA_CENTER))
    header = [c("Metric",bold=True)] + [c(i["year"],bold=True) for i in items]
    row    = [c(title,bold=True,color=C_ACC)] + [
        c(f"{i['value']:,.0f}Cr" if i["value"] is not None else "N/A") for i in items]
    cw = [3.5*cm] + [2.0*cm]*len(items)
    t  = Table([header,row], colWidths=cw)
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),C_MID),("TEXTCOLOR",(0,0),(-1,0),C_WHITE),
        ("BACKGROUND",(0,1),(-1,-1),C_LIGHT),
        ("GRID",(0,0),(-1,-1),0.3,colors.HexColor("#D0D7DE")),
        ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),
    ]))
    return [t, Spacer(1,5)]

def fmt(v, s="", dp=1):
    if v is None: return "N/A"
    return f"{v:.{dp}f}{s}"

def build_pdf(A):
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
          leftMargin=1.8*cm,rightMargin=1.8*cm,topMargin=1.5*cm,bottomMargin=1.5*cm)
    m   = A["metrics"]
    val = A.get("val")
    story = []

    # Header
    banner = Table([[Paragraph(
        f"<b>{A['company_name'].upper()}</b><br/>"
        f"<font size='9' color='#8B949E'>Stock Analysis  •  {A['generated_at']}</font>",
        ST("bh",fontName="Helvetica-Bold",fontSize=18,textColor=C_WHITE,alignment=TA_CENTER,leading=24))]],
        colWidths=[17*cm])
    banner.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),C_DARK),
        ("TOPPADDING",(0,0),(-1,-1),16),("BOTTOMPADDING",(0,0),(-1,-1),16)]))
    story.append(banner); story.append(Spacer(1,10))

    # Verdict
    vc_map = {"strong-buy":C_GREEN,"buy":C_GREEN,"watch":C_AMBER,"hold":C_AMBER,"avoid":C_RED}
    vc = vc_map.get(A["v_cls"], C_MID)
    vb = Table([[Paragraph(f"FUNDAMENTALS: {A['verdict']} — {A['overall']}/10",
        ST("vb",fontName="Helvetica-Bold",fontSize=12,textColor=C_WHITE,alignment=TA_CENTER))]],
        colWidths=[17*cm])
    vb.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),vc),
        ("TOPPADDING",(0,0),(-1,-1),8),("BOTTOMPADDING",(0,0),(-1,-1),8)]))
    story.append(vb); story.append(Spacer(1,12))

    # Business Overview
    biz = A.get("business", {})
    if biz.get("industry") or biz.get("market_cap"):
        story.append(Paragraph("BUSINESS OVERVIEW",ST("biz_sec",fontName="Helvetica-Bold",fontSize=12,textColor=C_ACC,spaceAfter=5)))
        biz_pairs = []
        if biz.get("industry"): biz_pairs.append(("Industry", biz["industry"]))
        if biz.get("market_cap"): biz_pairs.append(("Market Cap", f"₹{biz['market_cap']:,.0f} Cr"))
        if biz.get("market_cap_category"): biz_pairs.append(("Category", biz["market_cap_category"]))
        if biz.get("promoter_holding") is not None: biz_pairs.append(("Promoter Holding", f"{biz['promoter_holding']}%"))
        if biz_pairs:
            story.append(kv_tbl(biz_pairs,(5.5*cm,2.8*cm))); story.append(Spacer(1,8))

    # Key metrics
    story.append(Paragraph("KEY METRICS",ST("sec",fontName="Helvetica-Bold",fontSize=12,textColor=C_ACC,spaceAfter=5)))
    qp = [
        ("Revenue CAGR",  fmt(m["rev_cagr"],"%")),
        ("Profit CAGR",   fmt(m["pro_cagr"],"%")),
        ("Oper. Margin",  fmt(m["latest_opm"],"%")),
        ("ROE",           fmt(m["latest_roe"],"%")),
        ("ROCE",          fmt(m["latest_roce"],"%")),
        ("P/E",           fmt(m["latest_pe"],"x")),
        ("D/E",           fmt(m["latest_de"],"x")),
        ("EPS",           f"₹{fmt(m['latest_eps'])}"),
        ("BVPS",          f"₹{fmt(m['latest_bvps'])}"),
        ("Debt",          "Reducing ✓" if m["debt_reduced"] else "Increasing ↑"),
    ]
    qrow = [[kv_tbl(qp[:5],(5.0*cm,3.0*cm)),Spacer(0.3*cm,0),kv_tbl(qp[5:],(5.0*cm,3.0*cm))]]
    qt   = Table(qrow,colWidths=[8.4*cm,0.2*cm,8.4*cm])
    qt.setStyle(TableStyle([("VALIGN",(0,0),(-1,-1),"TOP"),("LEFTPADDING",(0,0),(-1,-1),2),("RIGHTPADDING",(0,0),(-1,-1),2)]))
    story.append(qt); story.append(Spacer(1,10))

    # Trends
    story.append(Paragraph("GROWTH TRENDS",ST("s2",fontName="Helvetica-Bold",fontSize=12,textColor=C_ACC,spaceAfter=5)))
    story += series_tbl_pdf(A["series"]["sales"],  "Sales (Cr)")
    story += series_tbl_pdf(A["series"]["profit"], "Net Profit (Cr)")
    story += series_tbl_pdf(A["series"]["ebitda"], "EBITDA (Cr)")
    story.append(Paragraph("BALANCE SHEET",ST("s3",fontName="Helvetica-Bold",fontSize=12,textColor=C_ACC,spaceAfter=5)))
    story += series_tbl_pdf(A["series"]["borrowings"],"Borrowings (Cr)")
    story += series_tbl_pdf(A["series"]["reserves"],  "Reserves (Cr)")
    story.append(Paragraph("CASH FLOW",ST("s4",fontName="Helvetica-Bold",fontSize=12,textColor=C_ACC,spaceAfter=5)))
    story += series_tbl_pdf(A["series"]["cfo"],"CFO (Cr)")

    # Valuation
    story.append(Paragraph("VALUATION — UNDERVALUED OR OVERVALUED?",
        ST("s5",fontName="Helvetica-Bold",fontSize=12,textColor=C_ACC,spaceAfter=5)))
    if val:
        vsum = [
            ("Current Price",   f"₹{val['current_price']:,.2f}"),
            ("Intrinsic Value", f"₹{val['weighted_iv']:,.2f}"),
            ("Upside/Downside", f"{val['upside_pct']:+.1f}%"),
            ("Margin of Safety",f"{val['margin_safety']:.1f}%"),
        ]
        if val.get("peg_actual"):
            vsum.append(("PEG Ratio",f"{val['peg_actual']:.2f}"))
        story.append(kv_tbl(vsum,(6*cm,3.5*cm))); story.append(Spacer(1,6))

        mhdr = [Paragraph(h,ST("mh",fontName="Helvetica-Bold",fontSize=9,textColor=C_WHITE,alignment=TA_CENTER))
                for h in ["Model","Intrinsic Value","Weight","Formula"]]
        mrows = [mhdr]
        for nm, md in val["models"].items():
            diff = md["iv"] - val["current_price"]
            ic   = C_GREEN if diff > 0 else C_RED
            mrows.append([
                Paragraph(nm, ST("mn",fontName="Helvetica-Bold",fontSize=9,textColor=C_TEXT)),
                Paragraph(f"₹{md['iv']:,.2f}",ST("iv",fontName="Helvetica-Bold",fontSize=9,textColor=ic,alignment=TA_CENTER)),
                Paragraph(f"{int(md['weight']*100)}%",ST("w",fontName="Helvetica",fontSize=9,textColor=C_MID,alignment=TA_CENTER)),
                Paragraph(md["formula"][:45],ST("fm",fontName="Helvetica",fontSize=8,textColor=C_MID)),
            ])
        mrows.append([
            Paragraph("<b>WEIGHTED IV</b>",ST("wi",fontName="Helvetica-Bold",fontSize=9,textColor=C_WHITE)),
            Paragraph(f"<b>₹{val['weighted_iv']:,.2f}</b>",ST("wv",fontName="Helvetica-Bold",fontSize=9,textColor=C_WHITE,alignment=TA_CENTER)),
            Paragraph("",ST("we")),
            Paragraph("Weighted avg of all models",ST("wf",fontName="Helvetica",fontSize=8,textColor=C_WHITE)),
        ])
        n = len(mrows)
        mt = Table(mrows,colWidths=[3.5*cm,3.5*cm,1.2*cm,9.8*cm])
        mt.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),C_MID),
            ("ROWBACKGROUNDS",(0,1),(-1,n-2),[C_WHITE,C_LIGHT]),
            ("BACKGROUND",(0,n-1),(-1,-1),C_TEAL),
            ("GRID",(0,0),(-1,-1),0.3,colors.HexColor("#D0D7DE")),
            ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5),
            ("LEFTPADDING",(0,0),(-1,-1),6),
        ]))
        story.append(mt); story.append(Spacer(1,8))

        vc2_map = {"strong":C_GREEN,"mild":C_GREEN,"fair":C_AMBER,"over":C_RED,"danger":C_RED}
        vc2 = vc2_map.get(val["val_class"], C_MID)
        vvb = Table([[Paragraph(val["val_verdict"],
            ST("vvb",fontName="Helvetica-Bold",fontSize=12,textColor=C_WHITE,alignment=TA_CENTER))]],
            colWidths=[17*cm])
        vvb.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),vc2),
            ("TOPPADDING",(0,0),(-1,-1),9),("BOTTOMPADDING",(0,0),(-1,-1),9)]))
        story.append(vvb); story.append(Spacer(1,12))
    else:
        story.append(Paragraph("Insufficient data for valuation (need CMP / EPS / BVPS)",
            ST("na",fontName="Helvetica",fontSize=10,textColor=C_RED))); story.append(Spacer(1,10))

    # Professional Checklist
    chk = A.get("checklist", {})
    if chk:
        story.append(Paragraph("PROFESSIONAL INVESTOR CHECKLIST",ST("chk_sec",fontName="Helvetica-Bold",fontSize=12,textColor=C_ACC,spaceAfter=5)))
        chk_pairs = [(k, f"{v}/10") for k, v in chk.items()]
        chk_pairs.append(("Final Score", f"{A.get('checklist_overall', 'N/A')}/10"))
        story.append(kv_tbl(chk_pairs,(6*cm,2.5*cm))); story.append(Spacer(1,8))

    # Risk section
    risks = A.get("risks", [])
    if risks:
        story.append(Paragraph("RISK ASSESSMENT",ST("risk_sec",fontName="Helvetica-Bold",fontSize=12,textColor=C_ACC,spaceAfter=5)))
        risk_text = "<br/>".join(f"• {r}" for r in risks)
        story.append(Paragraph(risk_text, ST("rt",fontName="Helvetica",fontSize=9,textColor=C_RED,leading=14)))
        story.append(Spacer(1,8))

    # ── Advanced Analysis: Piotroski F-Score ──
    pf = A.get("piotroski_fscore", {})
    if pf and pf.get("score") is not None:
        story.append(Paragraph("PIOTROSKI F-SCORE",ST("pf_sec",fontName="Helvetica-Bold",fontSize=12,textColor=C_ACC,spaceAfter=5)))
        pf_color = C_GREEN if pf["score"] >= 7 else C_AMBER if pf["score"] >= 5 else C_RED
        story.append(kv_tbl([
            ("Score", f"{pf['score']}/{pf['max_score']}"),
            ("Rating", pf.get("rating","N/A")),
        ], (6*cm, 3*cm)))
        details = pf.get("details", {})
        if details:
            detail_lines = [f"{'✓' if v else '✗'} {k}" for k, v in details.items()]
            story.append(Paragraph("<br/>".join(detail_lines),
                ST("pf_dtl",fontName="Helvetica",fontSize=8.5,textColor=C_TEXT,leading=14)))
            story.append(Spacer(1,6))

    # ── Advanced Analysis: Altman Z-Score ──
    az = A.get("altman_z", {})
    if az and az.get("z_score") is not None:
        story.append(Paragraph("ALTMAN Z-SCORE (BANKRUPTCY RISK)",ST("az_sec",fontName="Helvetica-Bold",fontSize=12,textColor=C_ACC,spaceAfter=5)))
        az_color = C_GREEN if az["zone"] == "Safe" else C_AMBER if az["zone"] == "Grey Zone" else C_RED
        story.append(kv_tbl([
            ("Z-Score", f"{az['z_score']:.2f}"),
            ("Zone", az["zone"]),
        ], (6*cm, 3*cm)))
        comps = az.get("components", {})
        if comps:
            comp_text = " | ".join(f"{k}: {v}" for k, v in comps.items())
            story.append(Paragraph(comp_text, ST("az_cmp",fontName="Helvetica",fontSize=8,textColor=C_MID)))
            story.append(Spacer(1,6))

    # ── Advanced Analysis: Earnings Quality + Revenue Acceleration ──
    eq = A.get("earnings_quality", {})
    ra = A.get("revenue_acceleration", {})
    if eq or ra:
        story.append(Paragraph("EARNINGS QUALITY & REVENUE ACCELERATION",ST("eq_sec",fontName="Helvetica-Bold",fontSize=12,textColor=C_ACC,spaceAfter=5)))
        eq_pairs = []
        if eq.get("quality"):
            eq_pairs.append(("Earnings Quality", f"{eq['quality']} ({eq.get('latest',0):.2f}x)"))
        if ra.get("status"):
            eq_pairs.append(("Revenue Trend", ra["status"]))
        if ra.get("latest_growth"):
            eq_pairs.append(("Latest Growth", f"+{ra['latest_growth']}%"))
        if eq_pairs:
            story.append(kv_tbl(eq_pairs, (6*cm, 3*cm)))
            story.append(Spacer(1,6))

    # ── Advanced Analysis: Entry Zones & Bear/Base/Bull ──
    ez = A.get("entry_zones", {})
    bbb = A.get("bear_base_bull", {})
    if ez.get("strong_buy_below") or bbb.get("base"):
        story.append(Paragraph("ENTRY ZONES & PRICE TARGETS",ST("ez_sec",fontName="Helvetica-Bold",fontSize=12,textColor=C_ACC,spaceAfter=5)))
        ez_pairs = []
        if ez.get("recommendation"): ez_pairs.append(("Recommendation", ez["recommendation"]))
        if ez.get("strong_buy_below"): ez_pairs.append(("Strong Buy Below", f"₹{ez['strong_buy_below']:,.0f}"))
        if ez.get("buy_below"): ez_pairs.append(("Buy Below", f"₹{ez['buy_below']:,.0f}"))
        if ez.get("fair_value"): ez_pairs.append(("Fair Value", f"₹{ez['fair_value']:,.0f}"))
        if ez.get("overvalued_above"): ez_pairs.append(("Overvalued Above", f"₹{ez['overvalued_above']:,.0f}"))
        if bbb.get("bear"): ez_pairs.append(("Bear Target", f"₹{bbb['bear']:,.0f}"))
        if bbb.get("base"): ez_pairs.append(("Base Target", f"₹{bbb['base']:,.0f}"))
        if bbb.get("bull"): ez_pairs.append(("Bull Target", f"₹{bbb['bull']:,.0f}"))
        if ez_pairs:
            story.append(kv_tbl(ez_pairs, (6*cm, 3*cm)))
            story.append(Spacer(1,6))

    # ── Advanced Analysis: Magic Formula ──
    mf = A.get("magic_formula", {})
    if mf and mf.get("roce") is not None:
        story.append(Paragraph("MAGIC FORMULA (GREENBLATT)",ST("mf_sec",fontName="Helvetica-Bold",fontSize=12,textColor=C_ACC,spaceAfter=5)))
        story.append(kv_tbl([
            ("ROCE", f"{mf['roce']:.1f}%"),
            ("Earnings Yield", f"{mf.get('earnings_yield',0):.1f}%" if mf.get('earnings_yield') else "N/A"),
            ("Rank", mf.get("magic_formula_rank","N/A")),
        ], (6*cm, 3*cm)))
        story.append(Spacer(1,6))

    # ── Advanced Analysis: Overall Signal ──
    osig = A.get("overall_signal", {})
    if osig and osig.get("label"):
        sig_color = C_GREEN if osig["label"] in ("STRONG BUY","BUY") else C_AMBER if osig["label"] == "WAIT & WATCH" else C_RED
        sig_banner = Table([[Paragraph(
            f"COMBINED SIGNAL: {osig['label']}",
            ST("sig_t",fontName="Helvetica-Bold",fontSize=14,textColor=C_WHITE,alignment=TA_CENTER))]],
            colWidths=[17*cm])
        sig_banner.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),sig_color),
            ("TOPPADDING",(0,0),(-1,-1),9),("BOTTOMPADDING",(0,0),(-1,-1),9)]))
        story.append(sig_banner); story.append(Spacer(1,8))
        if osig.get("details"):
            det_text = "  •  ".join(osig["details"])
            story.append(Paragraph(det_text,
                ST("sig_d",fontName="Helvetica",fontSize=8.5,textColor=C_MID,alignment=TA_CENTER,leading=12)))
            story.append(Spacer(1,10))

    # ── Investor Framework Scores ──
    fws = A.get("investor_frameworks", {})
    if fws and fws.get("frameworks"):
        story.append(Paragraph("HOW LEGENDARY INVESTORS SCORE THIS STOCK",
            ST("fw_sec",fontName="Helvetica-Bold",fontSize=12,textColor=C_ACC,spaceAfter=5)))
        overall = fws.get("overall", {})
        if overall:
            ov_color = C_GREEN if overall.get("overall_score", 0) >= 60 else C_AMBER if overall.get("overall_score", 0) >= 40 else C_RED
            ov_banner = Table([[Paragraph(
                f"OVERALL INVESTMENT SCORE: {overall.get('overall_score', 'N/A')}/100  —  {overall.get('verdict', '')}",
                ST("ov_t",fontName="Helvetica-Bold",fontSize=11,textColor=C_WHITE,alignment=TA_CENTER))]],
                colWidths=[17*cm])
            ov_banner.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),ov_color),
                ("TOPPADDING",(0,0),(-1,-1),8),("BOTTOMPADDING",(0,0),(-1,-1),8)]))
            story.append(ov_banner); story.append(Spacer(1,8))
        for fw in fws["frameworks"]:
            sc = fw.get("score", 0)
            sc_color = C_GREEN if sc >= 60 else C_AMBER if sc >= 40 else C_RED
            fw_pairs = [("Score", f"{sc}/100")]
            comps = fw.get("components", {})
            for k, v in comps.items():
                fw_pairs.append((f"  {k}", f"{v.get('score',0)}/{v.get('max',0)}"))
            story.append(Paragraph(f"<b>{fw['framework']}</b> — {fw.get('verdict','')}",
                ST("fw_n",fontName="Helvetica-Bold",fontSize=10,textColor=sc_color,spaceAfter=2)))
            story.append(Paragraph(f"<font size='8' color='#8B949E'>{fw.get('tagline','')}</font>",
                ST("fw_t",fontName="Helvetica",fontSize=8,textColor=C_MID,spaceAfter=4)))
            story.append(kv_tbl(fw_pairs, (6*cm, 3.5*cm)))
            story.append(Spacer(1,6))

    # Scores
    story.append(Paragraph("SCORE CARD",ST("s6",fontName="Helvetica-Bold",fontSize=12,textColor=C_ACC,spaceAfter=5)))
    def sc_col(s): return C_GREEN if s>=8 else C_AMBER if s>=6 else C_RED
    shdr = [Paragraph(h,ST("sh",fontName="Helvetica-Bold",fontSize=9,textColor=C_WHITE))
            for h in ["Category","Score","Visual"]]
    srows = [shdr]
    for cat, sc in A["scores"].items():
        bar = "■"*int(sc)+"□"*(10-int(sc))
        srows.append([
            Paragraph(cat,ST("sc",fontName="Helvetica",fontSize=9,textColor=C_TEXT)),
            Paragraph(f"{sc}/10",ST("ss",fontName="Helvetica-Bold",fontSize=9,textColor=sc_col(sc),alignment=TA_CENTER)),
            Paragraph(bar,ST("sb",fontName="Helvetica",fontSize=8,textColor=sc_col(sc),alignment=TA_CENTER)),
        ])
    srows.append([
        Paragraph("<b>OVERALL</b>",ST("so",fontName="Helvetica-Bold",fontSize=10,textColor=C_WHITE)),
        Paragraph(f"<b>{A['overall']}/10</b>",ST("so2",fontName="Helvetica-Bold",fontSize=10,textColor=C_WHITE,alignment=TA_CENTER)),
        Paragraph("",ST("so3")),
    ])
    ns = len(srows)
    st = Table(srows,colWidths=[7*cm,2.5*cm,4*cm])
    st.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),C_MID),
        ("ROWBACKGROUNDS",(0,1),(-1,ns-2),[C_WHITE,C_LIGHT]),
        ("BACKGROUND",(0,ns-1),(-1,-1),C_ACC),
        ("GRID",(0,0),(-1,-1),0.3,colors.HexColor("#D0D7DE")),
        ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5),
        ("LEFTPADDING",(0,0),(-1,-1),7),
    ]))
    story.append(st); story.append(Spacer(1,16))

    story.append(HRFlowable(width="100%",thickness=0.5,color=colors.HexColor("#D0D7DE")))
    story.append(Spacer(1,5))
    story.append(Paragraph("Data: Screener.in  •  Educational only  •  Not financial advice",
        ST("ft",fontName="Helvetica",fontSize=7.5,textColor=colors.HexColor("#8B949E"),alignment=TA_CENTER)))
    doc.build(story)
    return buf


if __name__ == "__main__":
    print("\n*** Stock Analyzer running at http://localhost:5000 ***\n")
    app.run(debug=True, port=5000)
