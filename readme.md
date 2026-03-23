# 📈 India Stock Portfolio Tracker — Google Sheets + Apps Script

> A fully automated NSE stock portfolio tracker built inside Google Sheets.
> **Zero manual work after setup.** Live prices, daily snapshots, email alerts, sector tagging, health score, and 6 charts — all automatic.

---

## ✨ What Makes This Different

Most portfolio trackers make you update prices manually, need a paid API, or break every few months. This one:

- Uses **Google Finance** — free, built into Google Sheets, always works
- Runs **entirely inside Google Sheets** — no app to install, no account to create
- **Auto-triggers at 4 PM IST** every trading day — you don't press anything
- Sends **email alerts** when your stocks hit target or stop loss
- Builds **weekly and monthly summaries** automatically from daily history
- Scores your portfolio **health out of 100** every day

---

## 🚀 Quick Setup (5 minutes, one time only)

### Step 1 — Create a new Google Sheet

Go to [sheets.google.com](https://sheets.google.com) and create a blank spreadsheet.

---

### Step 2 — Open Apps Script

Inside your Google Sheet:

```
Extensions → Apps Script
```

---

### Step 3 — Paste the code

1. Delete everything in the editor (select all → delete)
2. Paste the entire contents of `portfolio-tracker.gs`
3. Click **Save** (floppy disk icon or `Ctrl+S`)

---

### Step 4 — Set the timezone

1. Click the **⚙️ gear icon** on the left sidebar → **Project Settings**
2. Scroll to **Time zone** → set it to `Asia/Kolkata`
3. Click **Save**

> This is important. Without this, the 4 PM IST trigger will fire at the wrong time.

---

### Step 5 — Run Setup

1. In the function dropdown at the top, select **`setup`**
2. Click **▶ Run**
3. A permissions popup will appear — click **Review permissions → Allow**
4. Wait 10–15 seconds while all sheets are created automatically
5. A popup confirms: **"Setup complete!"**

---

### Step 6 — Add your stocks

Go to the **Portfolio** sheet and fill only these 5 columns:

| Column | What to enter | Example |
|--------|--------------|---------|
| **A** — Stock Name | Any name you want | `HDFC Bank` |
| **B** — NSE Symbol | Exact NSE ticker | `HDFCBANK` |
| **C** — Shares | How many shares you own | `50` |
| **D** — Avg Buy Price ₹ | Price you paid per share | `1650.00` |
| **E** — Buy Date | Date of purchase (DD/MM/YYYY) | `15/03/2024` |

> **Columns F onwards fill automatically. Do not type anything there.**

After adding your stocks, run **`fillNewStockFormulas`** from Apps Script to apply formulas to your new rows.

---

### ✅ Done. Everything else runs automatically.

---

## 📋 All 8 Sheets — Explained

### 1. 📊 Portfolio

The main sheet. You fill only columns A–E. Everything else is calculated automatically.

| Column | Name | Fills automatically? | What it shows |
|--------|------|---------------------|--------------|
| A | Stock Name | No — you fill this | Name you entered |
| B | NSE Symbol | No — you fill this | Symbol you entered |
| C | Shares | No — you fill this | Shares you entered |
| D | Avg Buy Price ₹ | No — you fill this | Buy price you entered |
| E | Buy Date | No — you fill this | Buy date you entered |
| F | Current Price ₹ | ✅ Auto | Live NSE price via Google Finance |
| G | Current Value ₹ | ✅ Auto | Shares × Current Price |
| H | Invested ₹ | ✅ Auto | Shares × Buy Price |
| I | P&L ₹ | ✅ Auto | Current Value − Invested |
| J | P&L % | ✅ Auto | Overall return % |
| K | CAGR % | ✅ Auto | Annualised return based on buy date |
| L | Day Change % | ✅ Auto | Today's price change % |
| M | Day Change ₹ | ✅ Auto | Today's change in ₹ for your holding |
| N | 52W High ₹ | ✅ Auto | 52-week highest price |
| O | 52W Low ₹ | ✅ Auto | 52-week lowest price |
| P | % from 52W High | ✅ Auto | How far below the 52-week high |
| Q | Sector | ✅ Auto | Auto-tagged from 150+ stock database |
| R | Target Price ₹ | ✅ Auto (editable) | Default: buy price +20% |
| S | Stop Loss ₹ | ✅ Auto (editable) | Default: buy price −10% |
| T | To Target % | ✅ Auto | How far away from your target |
| U | Status | ✅ Auto | TARGET HIT / SL HIT / NEAR SL / NEAR TARGET / HOLD |

> **Tip:** You can override Target (col R) and Stop Loss (col S) by typing your own value. The script preserves custom values on rebuild.

---

### 2. 🎛️ Dashboard

A visual summary of your entire portfolio. Auto-refreshes every time you open the sheet.

**Section 1 — Market Overview**

Live prices and day change % for the 4 major Indian indices:

| Index | Ticker used |
|-------|------------|
| Nifty 50 | INDEXNSE:NIFTY_50 |
| Sensex | INDEXBOM:SENSEX |
| Nifty Bank | INDEXNSE:NIFTY_BANK |
| Nifty IT | INDEXNSE:NIFTY_IT |

**Section 2 — Portfolio Overview (4 cards)**

| Card | What it shows |
|------|--------------|
| NET WORTH ₹ | Total current value of all your holdings |
| TOTAL INVESTED ₹ | Total money you have put in |
| TOTAL P&L ₹ | Total profit or loss in ₹ |
| P&L % | Your overall return percentage |

**Section 3 — Today (4 cards)**

| Card | What it shows |
|------|--------------|
| TODAY'S CHANGE ₹ | How much your portfolio moved today |
| TODAY'S CHANGE % | Same expressed as % of invested amount |
| BEST TODAY | Which stock gained the most today |
| WORST TODAY | Which stock fell the most today |

**Section 4 — Health Score (top right)**

Your portfolio score out of 100, updated at 4 PM every trading day.
See [Health Score section](#-portfolio-health-score) below for how it's calculated.

**Section 5 — Holdings Breakdown (bottom table)**

A clean, colour-coded view of all 30 stocks with key metrics: buy price, current price, P&L ₹, P&L %, CAGR, day change, 52W high, sector, target, stop loss, and status.

---

### 3. 📅 Daily History

Automatically records a snapshot of your portfolio every trading day at 4 PM IST.

| Column | What it records |
|--------|----------------|
| Date | The trading date |
| Net Worth ₹ | Total portfolio value that day |
| Invested ₹ | Total amount invested |
| Total P&L ₹ | Total profit or loss |
| P&L % | Overall return % |
| Day Change ₹ | Change vs the previous day in ₹ |
| Day Change % | Change vs the previous day in % |

> This sheet powers all the charts and the weekly/monthly summaries. The longer you use the tracker, the more useful this becomes.

---

### 4. 📆 Weekly Summary

Auto-built from Daily History every trading day. Each row represents one week.

| Column | What it shows |
|--------|--------------|
| Week Of (Mon) | Start date of that week |
| Opening ₹ | Portfolio value at start of week |
| Closing ₹ | Portfolio value at end of week |
| Change ₹ | Weekly gain or loss in ₹ |
| Change % | Weekly gain or loss % |
| Weekly High ₹ | Highest value during the week |
| Weekly Low ₹ | Lowest value during the week |

Rows are colour-coded: **green background** = profitable week, **red background** = losing week.

---

### 5. 🗓️ Monthly Summary

Same structure as Weekly Summary but grouped by calendar month.

| Column | What it shows |
|--------|--------------|
| Month | e.g. `Mar 2025` |
| Opening ₹ | Portfolio value at start of month |
| Closing ₹ | Portfolio value at end of month |
| Change ₹ | Monthly gain or loss in ₹ |
| Change % | Monthly gain or loss % |
| Monthly High ₹ | Highest value that month |
| Monthly Low ₹ | Lowest value that month |

---

### 6. 👁️ Watchlist

Track stocks you are interested in but haven't bought yet. Fill only:
- **Col A** — Stock Name
- **Col B** — NSE Symbol

Everything else fills automatically: current price, day change, 52-week high/low, % from high, and sector.

To add a stock with a dialog box: run `addToWatchlist` from Apps Script.

---

### 7. 📉 Charts

Six charts are auto-generated and refreshed every trading day at 4 PM:

| # | Chart Title | Chart Type | Data Source |
|---|-------------|-----------|------------|
| 1 | Net Worth Over Time | Line | Daily History |
| 2 | Portfolio Allocation | Donut / Pie | Portfolio (current values) |
| 3 | Monthly P&L ₹ | Column | Monthly Summary |
| 4 | Daily Change ₹ | Area | Daily History |
| 5 | Overall P&L % per Stock | Horizontal Bar | Portfolio |
| 6 | Weekly P&L ₹ | Column | Weekly Summary |

---

### 8. 🗂️ _SectorData (hidden)

A hidden helper sheet containing 150+ NSE stock → sector mappings. This powers the auto-tagging in Portfolio col Q. You never need to open or edit this sheet.

**Sectors included:** IT, Banking, NBFC, Insurance, Energy, Defence, Healthcare, Pharma, Auto, Auto Ancillary, Metals, Mining, Cement, FMCG, Consumer, Retail, Chemicals, Agro Chemicals, Infrastructure, Capital Goods, Real Estate, Telecom, Fintech, Consumer Tech, Tourism, Consumer Durables, Conglomerate, Logistics.

---

## 💡 How CAGR is Calculated

CAGR (Compound Annual Growth Rate) shows your annualised return, so you can fairly compare stocks held for different periods.

**Formula:**
```
CAGR % = ((Current Price / Buy Price) ^ (365 / Days Held) − 1) × 100
```

**Why it matters:**
- A stock up **30% in 6 months** → CAGR of ~69% (high — it's moving fast)
- A stock up **30% in 2 years** → CAGR of ~14% (moderate)
- Simple P&L % treats both the same. CAGR shows the difference.

> This is why Buy Date (col E) is required.

---

## 🏥 Portfolio Health Score

Scored out of **100**, updated every trading day at 4 PM. Displayed in the top-right card of the Dashboard.

### How points are calculated:

| Component | Max Points | Calculation |
|-----------|-----------|-------------|
| Profitability | 40 pts | (Stocks in profit / Total stocks) × 40 |
| Sector Diversity | 30 pts | Number of sectors × 6 pts, capped at 5 sectors |
| Concentration | 30 pts | (1 − biggest holding % of portfolio) × 30 |

**Total = Profitability + Diversity + Concentration**

### Score labels:

| Score Range | Label | Meaning |
|-------------|-------|---------|
| 80–100 | Excellent | Well diversified, most stocks profitable |
| 60–79 | Good | Solid portfolio, minor risks |
| 40–59 | Fair | Some issues — review concentration or loss-making stocks |
| 0–39 | Needs Review | High concentration or too many losing positions |

**Example:** 8 of 10 stocks in profit (32 pts) + 4 sectors (24 pts) + biggest stock is 18% of portfolio (24.6 pts) = **80/100 — Excellent**

---

## 🔔 Email Alerts

Checked automatically every trading day at **3:35 PM IST** (25 minutes before market close).

| Alert Type | Trigger Condition | Email colour |
|-----------|------------------|-------------|
| 🟢 TARGET HIT | Current price ≥ Target Price (col R) | Green |
| 🔴 STOP LOSS HIT | Current price ≤ Stop Loss (col S) | Red |
| 🟡 NEAR STOP LOSS | Current price is within 3% above Stop Loss | Orange |

**What the email contains:**
- Stock name and symbol
- Alert type
- Current price vs the trigger level
- Overall P&L %
- Today's change %

**Duplicate prevention:** Each alert fires only **once per stock per day**. You won't get the same alert repeatedly.

**Email goes to:** The Google account you are logged into. No configuration needed.

---

## ⏰ Automatic Triggers

These run silently in the background every weekday. Weekends are skipped automatically.

| Time (IST) | Function | What it does |
|-----------|---------|--------------|
| 3:35 PM | `checkAlerts` | Checks all stocks for target/stop loss conditions, sends email if triggered |
| 4:00 PM | `runDailySnapshot` | Reads portfolio values, saves to Daily History, rebuilds summaries, updates health score, refreshes charts |

> Triggers are set up automatically when you run `setup`. You never need to configure them manually.

---

## 🛠️ Utility Functions

Run these manually from Apps Script when needed (select function → click ▶ Run):

| Function | When to use it |
|----------|---------------|
| `setup` | **First time only.** Creates all 8 sheets and sets up the two daily triggers. |
| `fixFormats` | If any ₹ or % cells show `#VALUE!` — fixes number formatting without touching your data. |
| `forceSnapshotNow` | Takes a snapshot immediately. Use after adding stocks or if history looks stale. |
| `fillNewStockFormulas` | After manually adding new rows to the Portfolio sheet — applies all formulas to those rows. |
| `addToWatchlist` | Opens a dialog to enter a stock name and symbol, adds it to the Watchlist sheet. |

---

## ❓ Frequently Asked Questions

**Q: My stock price shows 0 or doesn't load.**

The NSE symbol is probably wrong or has a special format. Check the exact ticker on [nseindia.com](https://www.nseindia.com). Examples of tricky symbols: `BAJAJ-AUTO` (has a dash), `M&M` (has an ampersand), `NIFTYBEES` (ETF).

---

**Q: ₹ columns show `#VALUE!` instead of numbers.**

Run `fixFormats` from Apps Script. This re-applies the correct number format to all sheets without changing any data.

---

**Q: I added new stocks but columns F–U are empty.**

Run `fillNewStockFormulas`. It scans all rows and adds formulas to any that are missing them.

---

**Q: Daily History is empty / summaries show nothing.**

Run `forceSnapshotNow` to take an immediate snapshot. The history builds day by day as the 4 PM trigger fires — it won't be populated from the past.

---

**Q: I want a different Target or Stop Loss for a specific stock.**

Type your value directly into col R (Target) or col S (Stop Loss) for that row. The script detects non-zero custom values and preserves them on rebuild.

---

**Q: Can I have more than one portfolio?**

Yes. Create a second Google Sheet, paste the script, and run `setup`. Each sheet is completely independent.

---

**Q: Does this work on mobile?**

Yes. Google Sheets works on Android and iOS. The Dashboard sheet is the best view on mobile.

---

**Q: Will my data be lost if I re-run `setup`?**

No. The script reads and saves all existing rows (cols A–S) before clearing the sheet, then writes them back with the updated structure.

---

**Q: I don't get email alerts. What's wrong?**

Make sure `checkAlerts` trigger is set. Go to Apps Script → **Triggers** (clock icon on left) and confirm there is a trigger for `checkAlerts` set to run daily between 3–4 PM. If missing, re-run `setup`.

---

## 📦 Code Structure

```
portfolio-tracker.gs
│
├── CONSTANTS
│   ├── SHEETS        Sheet name strings
│   ├── COL           Column index numbers (A=1, B=2 ...)
│   ├── COLORS        Hex colors used across all sheets
│   ├── FORMATS       Number format strings (₹, %, dates)
│   └── ALERT_STYLES  Alert type → background/font color map
│
├── MAIN FUNCTIONS  (you run these)
│   ├── setup()                  One-time setup
│   ├── runDailySnapshot()       4 PM auto-trigger
│   ├── checkAlerts()            3:35 PM auto-trigger
│   ├── fixFormats()             Repair ₹ display issues
│   ├── forceSnapshotNow()       Manual snapshot
│   ├── fillNewStockFormulas()   Repair missing formulas
│   └── addToWatchlist()         Add stock to Watchlist via dialog
│
├── SHEET BUILDERS  (called once during setup)
│   ├── _buildPortfolioSheet()   21-column main data sheet
│   ├── _buildDashboardSheet()   Cards + holdings table
│   ├── _buildHistorySheet()     Daily log sheet
│   ├── _buildWeeklySheet()      Weekly summary sheet
│   ├── _buildMonthlySheet()     Monthly summary sheet
│   ├── _buildWatchlistSheet()   Watchlist sheet
│   ├── _buildChartsSheet()      Chart container sheet
│   └── _buildSectorMapSheet()   Hidden sector lookup sheet
│
├── DATA & LOGIC  (called during daily snapshot)
│   ├── _saveHistoryRow()           Read portfolio → append to history
│   ├── _rebuildPeriodSummary()     Unified weekly + monthly rebuild
│   ├── _updateHealthScore()        Calculate and write health score
│   ├── _refreshCharts()            Remove old charts, insert 6 new ones
│   └── _applyDashboardColors()     Apply green/red font colors
│
└── HELPERS
    ├── _buildPortfolioFormulaRow() 14-formula array for one portfolio row
    ├── _buildWatchlistFormulaRow() 7-formula array for one watchlist row
    ├── _applyHeaderStyle()         Shared blue header row styling
    ├── _applyPeriodFormats()       ₹/% formats for summary sheets
    ├── _parseBuyDate()             Convert DD/MM/YYYY string to Date
    ├── _weekKey()                  Get Monday date string for a date
    ├── _monthKey()                 Get "MMM YYYY" string for a date
    └── _clearTriggers()            Delete old triggers before re-setup
```

---

## 🔒 Privacy & Security

- All data lives in **your own Google Sheet** — nothing is sent to any external server
- The script only accesses the spreadsheet it is attached to
- Email alerts go only to **your own Google account email**
- No API keys, no third-party services, no external accounts needed

---

## 📌 Known Limitations

- **Google Finance data** is generally reliable but can be delayed 15–20 minutes during market hours
- **ETFs and index funds** may use different ticker formats — check NSE India if prices don't load
- **Sector auto-tagging** covers 150+ popular NSE stocks; unknown symbols are tagged as `Other`
- **Dashboard Holdings table** shows up to 30 stocks; Portfolio sheet supports up to 100
- **Historical per-stock charts** are not included — Google Sheets charts don't support individual stock candlesticks

---

## 📄 License

MIT — free to use, modify, and distribute. Attribution appreciated but not required.

---

*Built with Google Apps Script + Google Finance. No external APIs. No subscriptions. No installations. Just a spreadsheet that works.*
