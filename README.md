# Titan Capital Markets — Trading Journal

Professional forex journal for the Currency & Futures Trading Group / Training & Proprietary Capital Division. Simple daily entry, deep analysis. Two surfaces, one model.

| Surface | File | Use it for |
|---|---|---|
| **Interactive desk** | `index.html` | Live ticket, filters, charts, JSON/CSV backup |
| **Excel workbook** | `Titan_Capital_Markets_Trading_Journal.xlsx` | Portable source of truth, dropdowns, formulas, dashboard |

No sample statistics are stored. The journal opens empty. Settings → *Load labeled example trades* is opt-in preview data tagged **EXAMPLE** in notes.

## Start here (web desk)

1. Open the live preview, or serve this folder (`python3 -m http.server`).
2. **Settings** → set **trader name**, **Trader ID**, and **initial capital**.
3. **New trade** → fill the left ticket. The right ticket calculates day, pips, planned/realized RR, duration, PnL, win/loss.
4. Leave the exit blank to keep a trade **open**. Open trades never enter win rate, profit factor, or the equity curve.
5. **Dashboard** and **Analytics** share the same filters: pair, weekday, session, setup, date range.

Keyboard: `Ctrl/Cmd + N` opens a new ticket.

## Start here (Excel)

1. Open `Titan_Capital_Markets_Trading_Journal.xlsx`.
2. `01_Settings` → yellow cells **B8 trader name**, **B9 Trader ID**, **B11 initial capital**.
3. `02_Trade_Log` → type only in **yellow** cells from row 2. Grey cells are formulas (pre-filled through row 401).
4. `03_Dashboard` → dropdown filters. `04_Breakdowns` → pair / day / session / setup tables.

Yellow = you type. Grey = formula. Do not overwrite grey cells.

## Pairs, sessions, setups

**Pairs:** AUDCHF, AUDJPY, AUDCAD, AUDUSD, AUDNZD, EURGBP, EURCHF, EURJPY, EURCAD, EURUSD, EURNZD, EURAUD, GBPCHF, GBPJPY, GBPCAD, GBPUSD, GBPNZD, GBPAUD, NZDCHF, NZDJPY, NZDCAD, NZDUSD, CADJPY, CADCHF, USDCHF, USDJPY, USDCAD, XAUUSD, CHFJPY

**Setups:** Safety trade · H1 5050 Bounce · H4 50505 Bounce · Other

**Sessions:** Asian · Tokyo · London · New York

**Position:** Buy / Long · Sell / Short

## How numbers are calculated

- **Pip size:** JPY pairs `0.01` · XAUUSD `0.10` · other FX `0.0001`
- **Contract:** FX 100,000 · gold 100 oz
- **Direction:** Buy = +1 · Sell = −1
- **Result pips** = (exit − entry) × direction ÷ pip size
- **Planned RR** = TP pips ÷ SL pips
- **Realized RR** = result pips ÷ SL pips (negative on a loss)
- **Duration** = exit datetime − entry datetime
- **Day of week** from the entry calendar date
- **PnL (USD-style):**
  - XXXUSD & XAUUSD → price move × contract × lots
  - USDXXX → that amount ÷ exit
  - Crosses → estimated at ~$10 per pip per standard lot — **paste broker PnL into the override**
- **Net PnL** = (override if present, else estimate) − commission
- **Result:** Open if no exit · BE if |Net PnL| ≤ threshold (default 0.50) · else sign of Net PnL
- **Equity** = initial capital + sum of closed net PnL (exit order in the desk, row order in Excel)
- **Win rate** = wins ÷ closed
- **Profit factor** = gross profit ÷ |gross loss| · `∞` if no losses and some profit · `—` if both zero
- **Expectancy** = total net PnL ÷ closed count
- **Max drawdown** = peak-to-trough on the equity curve (desk)

Open trades are excluded from every performance figure.

## Fields added beyond the original brief

Initial capital, account currency, commission/swap, broker PnL override, open vs closed, pips, planned vs realized RR, exit reason, plan adherence, market condition, emotion, timeframe, ticket, chart URL, MAE/MFE, risk %, expectancy, max drawdown, profit-factor zero-loss guard.

The in-app **Blueprint** page is the full specification (structure, formulas, charts, validation, QC).

## Data

The desk stores trades in **this browser only** (`localStorage`). Use **Export CSV** / **Backup JSON** before clearing cache. Excel is the file you can email, version, and archive.
