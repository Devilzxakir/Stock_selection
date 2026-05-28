<div align="center">

<img src="https://img.shields.io/badge/Live_Demo-stock--what--you--want.vercel.app-4f8ef7?style=for-the-badge&logo=vercel&logoColor=white" />

# 📈 StockSense — Smart Stock Analyzer

**Institutional-grade fundamental analysis for Indian stocks. Free. No login needed.**

![Python](https://img.shields.io/badge/Python-3.8+-3776AB?style=flat-square&logo=python&logoColor=white)
![Flask](https://img.shields.io/badge/Flask-3.0-000000?style=flat-square&logo=flask&logoColor=white)
![Vercel](https://img.shields.io/badge/Deployed-Vercel-000000?style=flat-square&logo=vercel&logoColor=white)
![Screener](https://img.shields.io/badge/Data-Screener.in-FF6B35?style=flat-square)
![Yahoo Finance](https://img.shields.io/badge/Live_Price-Yahoo_Finance-6001D2?style=flat-square)
![License](https://img.shields.io/badge/License-MIT-22c55e?style=flat-square)
![Stars](https://img.shields.io/github/stars/Devilzxakir/Stock_selection?style=flat-square&color=ffd700)

### 🌐 [stock-what-you-want.vercel.app](https://stock-what-you-want.vercel.app/)

Search any NSE/BSE stock → get a full analysis in seconds.  
**BUY / HOLD / AVOID** verdict backed by 10+ models + downloadable PDF report.

[🚀 Try Live Demo](https://stock-what-you-want.vercel.app/) · [📖 Features](#-features) · [📊 Models](#-valuation-models) · [🔧 Self Host](#-self-host) · [🌐 API](#-api-endpoints)

</div>

---

## 🖥️ Live Demo

> **👉 [stock-what-you-want.vercel.app](https://stock-what-you-want.vercel.app/)**

No installation needed. Just open and search any company.

**Two ways to analyze:**
1. **🔍 Search by name** — type company name → live autocomplete → instant analysis
2. **📊 Upload Excel** — download Excel from Screener.in → upload here → analyze offline

---

## 🖼️ How It Works

```
Option A — Live Search:
  Type "Tata Motors" → select from dropdown
  → fetches Excel from Screener.in automatically
  → computes 40+ metrics + 10 models
  → shows full dashboard + BUY/AVOID verdict

Option B — Excel Upload:
  Download Excel from screener.in/company/YOURSTOCK
  → upload .xlsx file to the website
  → same full analysis, no session needed
```

---

## ✨ Features

### 📊 Core Analysis Engine

| Feature | What It Does |
|---|---|
| **Screener.in Integration** | Full P&L, Balance Sheet, Cash Flow, Ratios via authenticated Excel |
| **Excel Upload Mode** | Upload `.xlsx` directly — works without session ID |
| **5-Model Valuation** | Weighted intrinsic value from 5 independent models |
| **Piotroski F-Score** | 9 binary financial health tests → score 0–9 |
| **Altman Z-Score** | Bankruptcy risk → Safe / Grey Zone / Distress |
| **Magic Formula Rank** | Greenblatt: ROCE quality + Earnings Yield value |
| **Earnings Quality** | CFO ÷ Net Profit — detects paper profits vs real cash |
| **Revenue Acceleration** | Is growth speeding up or slowing down? |
| **Entry Zones** | 4 price levels: Strong Buy / Buy / Fair Value / Overvalued |
| **Live Price** | Yahoo Finance auto-refresh every 30 seconds |

---

### 🚩 Red Flag Detection (11 Checks)

| Severity | Flags Checked |
|---|---|
| 🔴 **Critical** | CFO < 50% of Net Profit (fake earnings), recurring losses, critically high debt |
| 🟠 **High** | Negative CFO, promoter selling stake, revenue declining YoY |
| 🟡 **Medium** | Low interest coverage (<2x), overvaluation, decelerating growth |

---

### ✅ 20-Point Pre-Buy Checklist

Every point is weighted and scored before a BUY signal is shown:

```
 1. Business Understanding       11. Promoter Holding Trend
 2. Revenue Growth               12. Industry Outlook
 3. Profit Growth                13. Market Position
 4. Debt Analysis                14. Entry Timing
 5. Cash Flow Health             15. Risk Assessment
 6. ROE Quality                  16. Dividend History
 7. ROCE Efficiency              17. Economic Resilience
 8. Valuation (P/E vs Growth)    18. Long-Term Potential
 9. Economic Moat                19. Institutional Interest
10. Management Quality           20. Red Flag Safety
```

---

### 🔴 Buy Confirmation Gate

All 4 gates must pass before BUY is recommended:

```
✅  Checklist pass rate  ≥ 60%
✅  Zero critical red flags
✅  Positive upside vs intrinsic value
✅  Fundamentals score   ≥ 6 / 10
```

---

### 🏭 Qualitative Assessments

| Assessment | How It's Measured |
|---|---|
| **Industry Future** | Keyword scoring → Growing / Stable / Neutral |
| **Management Quality** | Promoter holding + debt trend + ROCE + ROE |
| **Economic Sensitivity** | Cyclical vs Defensive exposure |
| **Institutional Activity** | Smart money signal |

---

### 📄 PDF Report Contents

```
📋 Banner Header            →  Company + date
📋 Business Overview        →  Sector, key facts
📋 Key Metrics (40+)        →  All ratios in one view
📋 Growth Trends            →  Sales, Profit, EBITDA (8-year table)
📋 Balance Sheet            →  Borrowings, Reserves trend
📋 Cash Flow                →  Operating CFO trend
📋 Valuation Breakdown      →  All 5 models with formulas
📋 20-Point Checklist       →  Full scoring table
📋 Red Flag Report          →  All flags with severity
📋 Piotroski F-Score        →  9-point breakdown
📋 Altman Z-Score           →  Bankruptcy risk gauge
📋 Entry Zones              →  Price levels with % discount to IV
📋 Magic Formula Rank       →  Quality + Value score
📋 Score Card               →  7-parameter bar visual
📋 Final Verdict            →  BUY / HOLD / AVOID with reasoning
```

---

## 📊 Valuation Models

| # | Model | Formula | Weight |
|---|---|---|:---:|
| 1 | **Graham Number** | `√(22.5 × EPS × BVPS)` | 20% |
| 2 | **Graham Growth** | `EPS × (8.5 + 2g) × 4.4 ÷ AAA_yield` | 25% |
| 3 | **DCF — 10 Year** | Phase1 (high growth) + Phase2 (slowdown) + Terminal | 30% |
| 4 | **PEG Model** | `Fair P/E = Growth%` → `Fair Price = Fair P/E × EPS` | 15% |
| 5 | **EV / EBITDA** | `14× EBITDA − Debt + Cash ÷ Shares` | 10% |

**Output verdict:**

| Upside vs Intrinsic Value | Signal |
|---|---|
| ≥ +30% | 🟢 **UNDERVALUED** — Strong margin of safety |
| +10% to +30% | 🟡 **SLIGHTLY UNDERVALUED** — Reasonable entry |
| −10% to +10% | 🟠 **FAIRLY VALUED** — Limited upside |
| −10% to −30% | 🔴 **OVERVALUED** — Wait for correction |
| < −30% | 🔴 **HIGHLY OVERVALUED** — Significant downside risk |

---

## 🔧 Self Host

```bash
# 1. Clone
git clone https://github.com/Devilzxakir/Stock_selection.git
cd Stock_selection

# 2. Install
pip install -r requirements_web.txt

# 3. Optional: live price
pip install yfinance

# 4. Set session ID in app.py line ~27
SESSION_ID = "your_screener_session_id"

# 5. Run
python app.py
# → http://localhost:5000
```

> **💡 Tip:** Excel Upload mode works without a session ID — great for local testing.

---

### 🔑 Get Screener Session ID

> Only needed for live search mode. Excel upload works without it.

1. Open [screener.in](https://www.screener.in) → **Login**
2. Press **F12** → **Application** tab
3. Left panel → **Cookies** → `https://www.screener.in`
4. Find `sessionid` row → copy the **Value**
5. Paste in `app.py`:
```python
SESSION_ID = "paste_your_sessionid_here"
```

> ⚠️ Session expires every few days. Repeat above when you see `Session expired` error.

### 📁 Project Structure

```
Stock_selection/
├── app.py                   ← Flask backend + full analysis engine
├── requirements_web.txt     ← Python dependencies
└── templates/
    └── index.html           ← Frontend (dark theme, vanilla JS)
```

---

## 🌐 API Endpoints

| Endpoint | Method | Description |
|---|---|---|
| `/` | GET | Web dashboard |
| `/api/search?q=<query>` | GET | Live company autocomplete |
| `/api/analyze?slug=<slug>&name=<name>` | GET | Full analysis (JSON) |
| `/api/pdf?slug=<slug>&name=<name>` | GET | Download PDF report |
| `/api/live?slug=<slug>&name=<name>` | GET | Live price via Yahoo Finance |

### Sample Response — `/api/analyze`

```json
{
  "company_name": "Emmvee Photovoltaic Power Ltd",
  "metrics": {
    "rev_cagr": 27.4,
    "pro_cagr": 31.2,
    "latest_opm": 32.8,
    "latest_roe": 24.1,
    "latest_pe": 18.3,
    "debt_reduced": true
  },
  "val": {
    "weighted_iv": 842.50,
    "current_price": 610.00,
    "upside_pct": 38.1,
    "val_verdict": "🟢 UNDERVALUED — Strong margin of safety"
  },
  "overall": 8.5,
  "verdict": "STRONG BUY"
}
```

---

## ⚙️ Tech Stack

| Layer | Technology |
|---|---|
| **Backend** | Python 3.8+, Flask 3.0, Pandas, NumPy, Requests |
| **Frontend** | Vanilla HTML/CSS/JS — zero frameworks, dark theme |
| **Fonts** | Google Fonts: Syne + DM Mono + DM Sans |
| **Data** | Screener.in Excel export + Yahoo Finance (`yfinance`) |
| **PDF** | ReportLab — fully custom A4 layout |
| **Hosting** | Vercel |

---

## ⚠️ Known Limitations

| Issue | Details |
|---|---|
| **Session Expiry** | Screener.in session ID needs manual refresh every few days |
| **NSE Suffix** | Live price uses `.NS` — BSE-only stocks may not resolve |
| **No Technical Analysis** | Fundamentals only — no price charts or patterns |
| **No Cache/DB** | Each request re-fetches and computes fresh |
| **Valuation Gaps** | Models needing CMP/EPS/BVPS are skipped if data missing |

---

## 🗺️ Roadmap

- [ ] Piotroski F-Score full 9-point UI breakdown
- [ ] Promoter holding trend chart (quarterly)
- [ ] Bear / Base / Bull 3-scenario price targets
- [ ] Peer comparison (vs sector average)
- [ ] Portfolio tracker (watch multiple stocks)
- [ ] Email alert when stock enters Buy Zone
- [ ] Screener.in session auto-refresh

---

## ⚠️ Disclaimer

> This tool is for **educational and research purposes only**.  
> It is **not financial advice**. Always do your own due diligence.  
> The author is not responsible for any investment decisions or losses.

---

<div align="center">

Made with ❤️ by [Devilzxakir](https://github.com/Devilzxakir)

**⭐ Star this repo if it helped you!**

[🌐 Live App](https://stock-what-you-want.vercel.app/) · [🐛 Report Bug](https://github.com/Devilzxakir/Stock_selection/issues) · [💡 Request Feature](https://github.com/Devilzxakir/Stock_selection/issues)

</div>
