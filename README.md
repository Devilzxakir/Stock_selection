<div align="center">

# 📈 StockSense

**Deep Fundamental Analysis Engine for Indian Stocks**

![Python](https://img.shields.io/badge/Python-3.8+-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Flask](https://img.shields.io/badge/Flask-3.0-000000?style=for-the-badge&logo=flask&logoColor=white)
![Screener](https://img.shields.io/badge/Data-Screener.in-FF6B35?style=for-the-badge)
![Yahoo Finance](https://img.shields.io/badge/Live_Price-Yahoo_Finance-6001D2?style=for-the-badge)
![License](https://img.shields.io/badge/License-MIT-22c55e?style=for-the-badge)

Search any Indian stock → get a full institutional-grade analysis in seconds.  
**BUY / HOLD / AVOID** verdict backed by 10+ analytical models + printable PDF report.

[🚀 Quick Start](#-installation) · [✨ Features](#-features) · [📊 Models](#-valuation-models) · [🔧 API](#-api-endpoints) · [⚠️ Limitations](#-limitations)

</div>

---

## 🖼️ What It Does

```
You type:   "Tata Motors"
It fetches: Full Excel data from Screener.in (P&L, Balance Sheet, Cash Flow, Ratios)
It computes: 40+ metrics, 5-model valuation, 11 red flag checks, 20-point checklist
You get:    BUY / AVOID verdict + PDF report in < 10 seconds
```

---

## ✨ Features

### 📊 Data & Analysis

| Feature | Description |
|---|---|
| **Screener.in Integration** | Full P&L, Balance Sheet, Cash Flow, Ratios via authenticated Excel export |
| **5-Model Valuation** | Graham Number, Growth Graham, DCF 10yr, PEG, EV/EBITDA → single weighted IV |
| **Piotroski F-Score** | 9 binary financial health tests → score 0–9 |
| **Altman Z-Score** | Bankruptcy risk → Safe / Grey Zone / Distress |
| **Magic Formula** | Greenblatt: ROCE quality + Earnings Yield value rank |
| **Earnings Quality** | CFO vs Net Profit divergence — detects paper profits |
| **Revenue Acceleration** | Is growth speeding up or slowing down? |
| **Entry Zones** | 4 price zones: Strong Buy / Buy / Fair Value / Overvalued |
| **Live Price** | Yahoo Finance auto-refresh every 30 seconds |

---

### 🚩 Red Flag Detection (11 Checks)

Automatically flags dangerous signals with severity tags:

| Severity | Examples |
|---|---|
| 🔴 **Critical** | CFO < 50% of profit (fake earnings), recurring losses, critically high debt |
| 🟠 **High** | Negative CFO, promoter selling stake, revenue declining |
| 🟡 **Medium** | Low interest coverage, overvaluation, negative growth trend |

---

### ✅ 20-Point Pre-Buy Checklist

A weighted investment framework — every point scored before a BUY is recommended:

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

All 4 conditions must pass before showing a BUY signal:

```
✅  Checklist pass rate ≥ 60%
✅  Zero critical red flags
✅  Positive upside vs intrinsic value
✅  Fundamentals score ≥ 6 / 10
```

---

### 🏭 Category Assessments

| Assessment | Method |
|---|---|
| **Industry Future** | Keyword-based scoring → Growing / Stable / Neutral |
| **Management Quality** | Promoter holding + debt trend + ROCE + ROE |
| **Economic Sensitivity** | Cyclical vs Defensive exposure |
| **Institutional Activity** | Smart money indicator |

---

### 📄 PDF Report Sections

Professional A4 dark-theme report includes:

```
Banner Header          →  Company name + generated date
Business Overview      →  Sector, industry, key facts
Key Metrics            →  40+ financial ratios
Growth Trends          →  Sales, Profit, EBITDA (historical table)
Balance Sheet          →  Borrowings, Reserves trend
Cash Flow              →  CFO trend
Valuation Breakdown    →  All 5 models with formulas
20-Point Checklist     →  Full scoring table
Red Flag Report        →  All flags with severity
Piotroski F-Score      →  9-point breakdown
Altman Z-Score         →  Bankruptcy risk gauge
Entry Zones            →  Price levels with % discount
Magic Formula          →  Quality + Value rank
Score Card             →  7-parameter bar chart
Final Verdict          →  BUY / HOLD / AVOID with reasoning
```

---

## 📊 Valuation Models

| Model | Formula | Weight |
|---|---|:---:|
| **Graham Number** | `√(22.5 × EPS × BVPS)` | 20% |
| **Graham Growth** | `EPS × (8.5 + 2g) × 4.4 / AAA_yield` | 25% |
| **DCF — 10 Year** | Phase 1 (5yr high growth) + Phase 2 (5yr slowdown) + Terminal Value | 30% |
| **PEG Model** | `Fair P/E = Growth% → Fair Price = Fair P/E × EPS` (Lynch PEG=1 rule) | 15% |
| **EV / EBITDA** | `14× EBITDA − Debt + Cash ÷ Shares Outstanding` | 10% |

**Verdict scale:**

| Upside vs IV | Signal |
|---|---|
| ≥ +30% | 🟢 **UNDERVALUED** — Strong margin of safety |
| +10% to +30% | 🟡 **SLIGHTLY UNDERVALUED** — Reasonable entry |
| −10% to +10% | 🟠 **FAIRLY VALUED** — Limited upside |
| −10% to −30% | 🔴 **OVERVALUED** — Wait for correction |
| < −30% | 🔴 **HIGHLY OVERVALUED** — Significant downside risk |

---

## ⚙️ Tech Stack

| Layer | Technology |
|---|---|
| **Backend** | Python 3.8+, Flask 3.0, Requests, Pandas, NumPy |
| **Frontend** | Vanilla HTML / CSS / JS — dark theme, zero frameworks |
| **Fonts** | Google Fonts (Syne + DM Mono + DM Sans) |
| **Data Source** | Screener.in Excel export (authenticated session) |
| **Live Price** | Yahoo Finance (`yfinance`) with Screener fallback |
| **PDF Engine** | ReportLab — fully custom A4 layout |

---

## 🔧 Installation

```bash
# 1. Clone
git clone https://github.com/Devilzxakir/Stock_selection.git
cd Stock_selection

# 2. Install dependencies
pip install -r requirements_web.txt

# 3. Optional: live price support
pip install yfinance
```

---

## 🔑 Get Your Screener Session ID

> Required one-time setup — takes 1 minute.

1. Open [screener.in](https://www.screener.in) → **Login**
2. Press **F12** → go to **Application** tab
3. Left panel → **Cookies** → `https://www.screener.in`
4. Find row: **Name = `sessionid`** → copy the **Value**
5. Paste into `app.py` line ~27:

```python
SESSION_ID = "paste_your_value_here"
```

> ⚠️ Session expires periodically. Repeat steps above if you see `Session expired` error.

---

## ▶️ Usage

```bash
python app.py
```

Then open → **http://localhost:5000**

1. Type company name (e.g. `Emmvee`, `Reliance`, `TCS`)
2. Select from live autocomplete dropdown
3. Analysis loads in ~5–10 seconds
4. Click **⬇️ Download PDF** for full printable report

---

## 🌐 API Endpoints

| Endpoint | Method | Description |
|---|---|---|
| `/` | GET | Web UI dashboard |
| `/api/search?q=<query>` | GET | Company autocomplete from Screener.in |
| `/api/analyze?slug=<slug>&name=<name>` | GET | Full fundamental analysis (JSON) |
| `/api/pdf?slug=<slug>&name=<name>` | GET | Download PDF report |
| `/api/live?slug=<slug>&name=<name>` | GET | Live price via Yahoo Finance |

### Example Response — `/api/analyze`

```json
{
  "company_name": "Emmvee Photovoltaic Power Ltd",
  "generated_at": "28 May 2026 14:32",
  "metrics": {
    "rev_cagr": 27.4,
    "pro_cagr": 31.2,
    "latest_opm": 32.8,
    "latest_roe": 24.1,
    "latest_roce": 21.5,
    "latest_pe": 18.3,
    "latest_de": 0.12,
    "debt_reduced": true
  },
  "val": {
    "weighted_iv": 842.50,
    "current_price": 610.00,
    "upside_pct": 38.1,
    "val_verdict": "🟢 UNDERVALUED — Strong margin of safety"
  },
  "scores": { "Revenue Growth": 9, "Profit Growth": 9, "...": "..." },
  "overall": 8.5,
  "verdict": "STRONG BUY"
}
```

---

## 📁 Project Structure

```
Stock_selection/
├── app.py                   ← Flask backend + full analysis engine
│   ├── /api/search          ← Screener autocomplete
│   ├── /api/analyze         ← Core analysis (40+ metrics, 10+ models)
│   ├── /api/pdf             ← ReportLab PDF builder
│   └── /api/live            ← Yahoo Finance live price
├── templates/
│   └── index.html           ← Frontend UI (dark theme, vanilla JS)
├── requirements_web.txt     ← Python dependencies
└── README.md
```

---

## ⚠️ Limitations

| Limitation | Detail |
|---|---|
| **Session Expiry** | Screener session ID must be refreshed manually when expired |
| **NSE Only** | Live price uses `.NS` suffix — BSE-only stocks may not work |
| **No Technical Analysis** | Fundamental analysis only — no charts or price patterns |
| **No Database** | All analysis computed fresh per request — no caching |
| **Valuation Gaps** | Some models need CMP + EPS + BVPS — missing data = skipped model |

---

## 🗺️ Roadmap

- [ ] Piotroski F-Score UI breakdown
- [ ] Promoter holding chart (quarterly trend)
- [ ] Bear / Base / Bull price target scenarios
- [ ] Sector comparison (vs industry peers)
- [ ] Portfolio tracker (track multiple stocks)
- [ ] Email alert when stock enters Buy Zone

---

## ⚠️ Disclaimer

> This tool is for **educational and research purposes only**.  
> It is **not financial advice**. Past performance does not guarantee future results.  
> Always do your own due diligence before making any investment decision.  
> The author is not liable for any financial loss arising from use of this tool.

---

<div align="center">

Made with ❤️ by [Devilzxakir](https://github.com/Devilzxakir)

If this helped you — **⭐ star the repo!**

</div>
