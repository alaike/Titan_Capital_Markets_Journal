#!/usr/bin/env python3
"""Build Aurum Ledger — professional Excel trading journal."""
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.chart import BarChart, DoughnutChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.line import LineProperties
from openpyxl.chart.layout import Layout, ManualLayout
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import (
    Alignment, Border, Font, PatternFill, Protection, Side,
    NamedStyle,
)
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows  # noqa: kept unused-safe
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.chart.marker import Marker
from openpyxl.chart.series import SeriesLabel
from copy import copy

OUT = "/home/user/aurum-journal/Titan_Capital_Markets_Trading_Journal.xlsx"
LOGO_HEADER = "/home/user/aurum-journal/assets/tcm-logo-header.png"
LOGO_MARK = "/home/user/aurum-journal/assets/tcm-mark-sm.png"
N_ROWS = 400  # formula rows 2..401

# Palette
INK = "0B1220"
INK2 = "10151F"
NAVY = "161D2A"
GOLD = "C9A84C"
GOLD2 = "E7D392"
CREAM = "F7F4EC"
PAPER = "FFFdf8"
GREEN = "1F8A5A"
GREEN_BG = "E3F6EC"
RED = "C0392B"
RED_BG = "FDE8EA"
MUTED = "6D7A8C"
LINE = "D9D1C0"
WHITE = "FFFFFF"
TEAL = "2A6F7F"

thin = Border(
    left=Side(style="thin", color=LINE),
    right=Side(style="thin", color=LINE),
    top=Side(style="thin", color=LINE),
    bottom=Side(style="thin", color=LINE),
)
none_b = Border()

fill_ink = PatternFill("solid", fgColor=INK)
fill_navy = PatternFill("solid", fgColor=NAVY)
fill_gold = PatternFill("solid", fgColor=GOLD)
fill_cream = PatternFill("solid", fgColor=CREAM)
fill_paper = PatternFill("solid", fgColor=PAPER)
fill_green = PatternFill("solid", fgColor=GREEN_BG)
fill_red = PatternFill("solid", fgColor=RED_BG)
fill_kpi = PatternFill("solid", fgColor="F3EBD2")
fill_alt = PatternFill("solid", fgColor="FBF8F1")
fill_input = PatternFill("solid", fgColor="FFF9E6")
fill_calc = PatternFill("solid", fgColor="EEF3F7")
fill_open = PatternFill("solid", fgColor="F7E7B4")

font_white = Font(name="Calibri", color=WHITE, size=11)
font_gold = Font(name="Calibri", color=GOLD, size=11, bold=True)
font_title = Font(name="Calibri", color=GOLD2, size=22, bold=True)
font_h = Font(name="Calibri", color=WHITE, size=11, bold=True)
font_label = Font(name="Calibri", color=MUTED, size=9, bold=True)
font_input = Font(name="Calibri", color=INK, size=11)
font_calc = Font(name="Calibri", color="2A3A4A", size=11, italic=True)
font_kpi = Font(name="Calibri", color=INK, size=18, bold=True)
font_small = Font(name="Calibri", size=10, color="3A4654")

center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left = Alignment(horizontal="left", vertical="center", wrap_text=True)
right = Alignment(horizontal="right", vertical="center")

PAIRS = [
    "AUDCHF","AUDJPY","AUDCAD","AUDUSD","AUDNZD",
    "EURGBP","EURCHF","EURJPY","EURCAD","EURUSD","EURNZD","EURAUD",
    "GBPCHF","GBPJPY","GBPCAD","GBPUSD","GBPNZD","GBPAUD",
    "NZDCHF","NZDJPY","NZDCAD","NZDUSD",
    "CADJPY","CADCHF","USDCHF","USDJPY","USDCAD",
    "XAUUSD","CHFJPY",
]
SETUPS = ["Safety trade", "H1 5050 Bounce", "H4 50505 Bounce", "Other"]
SESSIONS = ["Asian", "Tokyo", "London", "New York"]
POSITIONS = ["Buy", "Sell"]
PLAN = ["Yes", "No", "Partial"]
EXIT_REASONS = [
    "Take profit hit", "Stop loss hit", "Breakeven / scratch",
    "Manual close", "Partial close then full", "News / invalidation",
]
MARKETS = ["Trending", "Ranging", "Volatile / news", "Low liquidity"]
WEEKDAYS = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]

# Trade log columns (1-indexed)
# A SN  B Day  C EntryDate  D EntryTime  E Pair  F Position  G Session
# H Lots  I RiskAmt  J RiskPct  K Entry  L SL  M TP  N Setup
# O Exit  P ExitDate  Q ExitTime  R Commission  S PnLOverride  T Notes
# U FollowedPlan  V ExitReason  W Market  X Ticket
# --- calculated ---
# Y PipSize  Z Direction  AA Contract  AB SLPips  AC TPPips  AD ResultPips
# AE PlannedRR  AF RealizedRR  AG DurationHrs  AH DurationTxt
# AI EstPnL  AJ NetPnL  AK Result  AL Balance  AM ClosedFlag

HEADERS = [
    "S/N", "Day of week", "Entry date", "Entry time", "Currency pair", "Position",
    "Session", "Lot size", "Risk per trade", "Risk %", "Entry price", "Stop loss",
    "Take profit", "Trading setup", "Exit price", "Exit date", "Exit time",
    "Commission / swap", "PnL override (broker)", "Notes",
    "Followed plan", "Exit reason", "Market condition", "Ticket / ID",
    "Pip size", "Direction", "Contract size", "SL pips", "TP pips", "Result pips",
    "Planned RR", "Realized RR", "Duration (hours)", "Trade duration",
    "Estimated PnL", "Net PnL", "Result", "Balance after trade", "Closed?",
]

INPUT_COLS = set(range(3, 25))  # C..X
CALC_COLS = set(range(1, 40)) - INPUT_COLS  # A,B and Y..AM

LAST = 1 + N_ROWS  # 401


def apply_col_widths(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def header_bar(ws, title, subtitle, last_col="L"):
    ws.merge_cells(f"A1:{last_col}1")
    ws.merge_cells(f"A2:{last_col}2")
    a1 = ws["A1"]
    a1.value = "TITAN CAPITAL MARKETS  ·  " + title
    a1.font = Font(name="Calibri", color=GOLD2, size=24, bold=True)
    a1.fill = fill_ink
    a1.alignment = Alignment(horizontal="left", vertical="center")
    a2 = ws["A2"]
    a2.value = subtitle
    a2.font = Font(name="Calibri", color="9AA6B8", size=11)
    a2.fill = fill_ink
    a2.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 32
    ws.row_dimensions[2].height = 20
    max_col = ws[last_col + "1"].column
    for col in range(1, max_col + 1):
        for r in (1, 2):
            cell = ws.cell(r, col)
            cell.fill = fill_ink


def section_label(ws, cell, text):
    c = ws[cell]
    c.value = text
    c.font = Font(name="Calibri", color=GOLD, size=10, bold=True)
    c.alignment = Alignment(horizontal="left", vertical="center")


def kpi_box(ws, r, c, label, formula, num_fmt=None, span=2):
    ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + span - 1)
    lab = ws.cell(r, c, label)
    lab.font = Font(name="Calibri", size=8, bold=True, color=MUTED)
    lab.fill = fill_kpi
    lab.alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=r + 1, start_column=c, end_row=r + 1, end_column=c + span - 1)
    val = ws.cell(r + 1, c, formula)
    val.font = font_kpi
    val.fill = fill_kpi
    val.alignment = Alignment(horizontal="center", vertical="center")
    if num_fmt:
        val.number_format = num_fmt
    for col in range(c, c + span):
        ws.cell(r, col).fill = fill_kpi
        ws.cell(r + 1, col).fill = fill_kpi
        ws.cell(r, col).border = thin
        ws.cell(r + 1, col).border = thin
    ws.row_dimensions[r].height = 16
    ws.row_dimensions[r + 1].height = 28


def build_lists(wb):
    ws = wb.create_sheet("99_Lists")
    ws.sheet_state = "hidden"
    headers = ["Pairs", "Setups", "Sessions", "Positions", "Plan", "ExitReason", "Market", "Weekdays", "FilterPair", "FilterDay", "FilterSession", "FilterSetup"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(1, i, h)
        cell.font = font_h
        cell.fill = fill_ink
    for i, p in enumerate(PAIRS, 2):
        ws.cell(i, 1, p)
    for i, p in enumerate(SETUPS, 2):
        ws.cell(i, 2, p)
    for i, p in enumerate(SESSIONS, 2):
        ws.cell(i, 3, p)
    for i, p in enumerate(POSITIONS, 2):
        ws.cell(i, 4, p)
    for i, p in enumerate(PLAN, 2):
        ws.cell(i, 5, p)
    for i, p in enumerate(EXIT_REASONS, 2):
        ws.cell(i, 6, p)
    for i, p in enumerate(MARKETS, 2):
        ws.cell(i, 7, p)
    for i, p in enumerate(WEEKDAYS, 2):
        ws.cell(i, 8, p)
    ws.cell(2, 9, "All")
    for i, p in enumerate(PAIRS, 3):
        ws.cell(i, 9, p)
    ws.cell(2, 10, "All")
    for i, p in enumerate(WEEKDAYS, 3):
        ws.cell(i, 10, p)
    ws.cell(2, 11, "All")
    for i, p in enumerate(SESSIONS, 3):
        ws.cell(i, 11, p)
    ws.cell(2, 12, "All")
    for i, p in enumerate(SETUPS, 3):
        ws.cell(i, 12, p)

    wb.defined_names.add(DefinedName("ListPairs", attr_text="99_Lists!$A$2:$A$30"))
    wb.defined_names.add(DefinedName("ListSetups", attr_text="99_Lists!$B$2:$B$5"))
    wb.defined_names.add(DefinedName("ListSessions", attr_text="99_Lists!$C$2:$C$5"))
    wb.defined_names.add(DefinedName("ListPositions", attr_text="99_Lists!$D$2:$D$3"))
    wb.defined_names.add(DefinedName("ListPlan", attr_text="99_Lists!$E$2:$E$4"))
    wb.defined_names.add(DefinedName("ListExitReason", attr_text="99_Lists!$F$2:$F$7"))
    wb.defined_names.add(DefinedName("ListMarket", attr_text="99_Lists!$G$2:$G$5"))
    wb.defined_names.add(DefinedName("FilterPairs", attr_text="99_Lists!$I$2:$I$31"))
    wb.defined_names.add(DefinedName("FilterDays", attr_text="99_Lists!$J$2:$J$9"))
    wb.defined_names.add(DefinedName("FilterSessions", attr_text="99_Lists!$K$2:$K$6"))
    wb.defined_names.add(DefinedName("FilterSetups", attr_text="99_Lists!$L$2:$L$6"))
    return ws


def build_settings(wb):
    ws = wb.create_sheet("01_Settings", 0)
    header_bar(ws, "TRADING JOURNAL", "Account settings  ·  trader identity and initial capital live here  ·  every formula starts from this sheet", "G")
    ws.row_dimensions[3].height = 8

    # Logo on a white strip
    ws.merge_cells("A4:C6")
    for r in range(4, 7):
        for c in range(1, 4):
            ws.cell(r, c).fill = PatternFill("solid", fgColor="FFFFFF")
            ws.cell(r, c).border = thin
    ws.row_dimensions[4].height = 22
    ws.row_dimensions[5].height = 22
    ws.row_dimensions[6].height = 22
    logo = XLImage(LOGO_HEADER)
    logo.width = 240
    logo.height = 113
    ws.add_image(logo, "A4")

    ws["D4"] = "FIRM"
    ws["D4"].font = Font(name="Calibri", size=8, bold=True, color=MUTED)
    ws.merge_cells("E4:G4")
    ws["E4"] = "Titan Capital Markets"
    ws["E4"].font = Font(name="Calibri", bold=True, size=14, color="1B2A4A")
    ws["D5"] = "DIVISION"
    ws["D5"].font = Font(name="Calibri", size=8, bold=True, color=MUTED)
    ws.merge_cells("E5:G5")
    ws["E5"] = "Currency & Futures Trading Group  ·  Training & Proprietary Capital Division"
    ws["E5"].font = font_small

    labels = [
        (8, "Trader name", "", "Full name as it should appear on this journal."),
        (9, "Trader ID", "", "Desk / prop ID, e.g. TCM-0042."),
        (10, "Account name", "", "Optional label (challenge / live / broker)."),
        (11, "Initial capital", 10000, "Starting cash. Equity curve begins at this number."),
        (12, "Account currency", "USD", "Used for number formats and the dashboard."),
        (13, "Default risk %", 1, "Informational default for the Risk % column."),
        (14, "Breakeven threshold", 0.5, "|Net PnL| at or below this is BE, not a win or a loss."),
    ]
    ws["A7"] = "INPUTS  (yellow cells)"
    ws["A7"].font = Font(name="Calibri", bold=True, color=GOLD, size=11)

    for r, label, val, hint in labels:
        ws.cell(r, 1, label).font = Font(name="Calibri", bold=True, size=11)
        cell = ws.cell(r, 2, val)
        cell.fill = fill_input
        cell.border = thin
        cell.font = Font(name="Calibri", bold=True, size=12)
        ws.cell(r, 3, hint).font = font_small
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=6)

    ws["B11"].number_format = '#,##0.00'
    ws["B13"].number_format = '0.00'
    ws["B14"].number_format = '0.00'

    dv_ccy = DataValidation(type="list", formula1='"USD,EUR,GBP,NGN,ZAR,CAD,AUD,CHF,JPY"', allow_blank=False)
    ws.add_data_validation(dv_ccy)
    dv_ccy.add("B12")

    ws["A16"] = "CALCULATED"
    ws["A16"].font = Font(name="Calibri", bold=True, color=GOLD, size=11)
    computed = [
        (17, "Closed trades", '=COUNTIF(\'02_Trade_Log\'!AK:AK,"Win")+COUNTIF(\'02_Trade_Log\'!AK:AK,"Loss")+COUNTIF(\'02_Trade_Log\'!AK:AK,"BE")'),
        (18, "Open trades", '=COUNTIF(\'02_Trade_Log\'!AK:AK,"Open")'),
        (19, "Total net PnL", '=SUMIF(\'02_Trade_Log\'!AK:AK,"<>Open",\'02_Trade_Log\'!AJ:AJ)'),
        (20, "Current equity", "=B11+B19"),
        (21, "Wins", '=COUNTIF(\'02_Trade_Log\'!AK:AK,"Win")'),
        (22, "Losses", '=COUNTIF(\'02_Trade_Log\'!AK:AK,"Loss")'),
        (23, "Breakevens", '=COUNTIF(\'02_Trade_Log\'!AK:AK,"BE")'),
        (24, "Win rate", '=IF(B17=0,"—",B21/B17)'),
        (25, "Gross profit", '=SUMIFS(\'02_Trade_Log\'!AJ:AJ,\'02_Trade_Log\'!AJ:AJ,">0",\'02_Trade_Log\'!AK:AK,"<>Open")'),
        (26, "Gross loss (abs)", '=ABS(SUMIFS(\'02_Trade_Log\'!AJ:AJ,\'02_Trade_Log\'!AJ:AJ,"<0",\'02_Trade_Log\'!AK:AK,"<>Open"))'),
        (27, "Profit factor", '=IF(B26=0,IF(B25>0,"∞","—"),B25/B26)'),
    ]
    for r, label, formula in computed:
        ws.cell(r, 1, label).font = Font(name="Calibri", size=11)
        cell = ws.cell(r, 2, formula)
        cell.fill = fill_calc
        cell.border = thin
        cell.font = font_calc
    ws["B19"].number_format = '#,##0.00'
    ws["B20"].number_format = '#,##0.00'
    ws["B24"].number_format = '0.0%'
    ws["B25"].number_format = '#,##0.00'
    ws["B26"].number_format = '#,##0.00'
    ws["B27"].number_format = '0.00'

    ws["A29"] = "RULES"
    ws["A24"].font = Font(name="Calibri", bold=True, color=GOLD, size=11)
    rules = [
        "Yellow cells are the only settings you type. Grey cells are formulas — do not overwrite them.",
        "Enter trades on 02_Trade_Log, starting at row 2. Do not insert columns; adding rows inside the table is safe.",
        "Leave Exit price and Exit date blank while a trade is open. Open trades never change equity, win rate, or profit factor.",
        "If the pair is a non-USD cross, paste the broker PnL into “PnL override”. Estimated PnL for crosses uses ~10 account-units per pip per standard lot.",
        "Pip size: JPY pairs 0.01 · XAUUSD 0.10 · other FX 0.0001. Contract: 100,000 FX · 100 oz gold.",
        "Buy: direction +1. Sell: direction −1. Result pips = (exit − entry) × direction ÷ pip size.",
        "Planned RR = TP pips ÷ SL pips. Realized RR = result pips ÷ SL pips (negative on a loss).",
        "Net PnL = (override if present, else estimate) − commission. BE when |Net PnL| ≤ breakeven threshold.",
        "Profit factor = gross profit ÷ gross loss. If gross loss is 0 and gross profit > 0 the cell shows ∞.",
        "Balance after trade is a running sum in row order. Sort the log by exit date before reading the equity path.",
        "S/N and Day of week are automatic. Dropdowns lock Pair, Position, Session, Setup, Plan, Exit reason, Market.",
        "Example / demo rows are not included. Do not invent results; log only tickets you actually took.",
    ]
    for i, t in enumerate(rules):
        ws.merge_cells(start_row=30 + i, start_column=1, end_row=30 + i, end_column=7)
        ws.cell(30 + i, 1, f"{i+1}.  {t}").font = font_small
        ws.row_dimensions[30 + i].height = 18

    apply_col_widths(ws, {"A": 28, "B": 22, "C": 22, "D": 16, "E": 16, "F": 16, "G": 16})
    ws.freeze_panes = "A8"
    ws.sheet_properties.tabColor = GOLD
    wb.defined_names.add(DefinedName("TraderName", attr_text="'01_Settings'!$B$8"))
    wb.defined_names.add(DefinedName("TraderId", attr_text="'01_Settings'!$B$9"))
    wb.defined_names.add(DefinedName("InitialCapital", attr_text="'01_Settings'!$B$11"))
    wb.defined_names.add(DefinedName("BEThreshold", attr_text="'01_Settings'!$B$14"))
    wb.defined_names.add(DefinedName("AccountCurrency", attr_text="'01_Settings'!$B$12"))
    return ws


def f_if_row(row, formula):
    """Wrap a formula so blank entry-date rows stay empty."""
    return f'IF($C{row}="","",{formula})'


def build_log(wb):
    ws = wb.create_sheet("02_Trade_Log", 1)
    # headers
    for i, h in enumerate(HEADERS, 1):
        cell = ws.cell(1, i, h)
        cell.font = Font(name="Calibri", color=WHITE, size=9, bold=True)
        cell.fill = fill_ink
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 36
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:AM{LAST}"

    # column widths
    widths = {
        "A": 6, "B": 14, "C": 13, "D": 12, "E": 14, "F": 11, "G": 12,
        "H": 11, "I": 14, "J": 10, "K": 13, "L": 12, "M": 13, "N": 18,
        "O": 12, "P": 13, "Q": 12, "R": 16, "S": 18, "T": 28,
        "U": 14, "V": 20, "W": 16, "X": 14,
        "Y": 11, "Z": 11, "AA": 13, "AB": 11, "AC": 11, "AD": 12,
        "AE": 12, "AF": 12, "AG": 14, "AH": 14, "AI": 14, "AJ": 12,
        "AK": 12, "AL": 18, "AM": 10,
    }
    apply_col_widths(ws, widths)

    # helper lambdas for formulas using current row
    def formulas(r):
        pip = f'IF(E{r}="XAUUSD",0.1,IF(ISNUMBER(SEARCH("JPY",E{r})),0.01,0.0001))'
        direc = f'IF(F{r}="Sell",-1,1)'
        cs = f'IF(E{r}="XAUUSD",100,100000)'
        slp = f'IF(OR(K{r}="",L{r}=""),"",ABS(K{r}-L{r})/Y{r})'
        tpp = f'IF(OR(K{r}="",M{r}=""),"",ABS(M{r}-K{r})/Y{r})'
        rp = f'IF(OR(K{r}="",O{r}=""),"",(O{r}-K{r})*Z{r}/Y{r})'
        prr = f'IF(OR(AB{r}="",AB{r}=0,AC{r}=""),"",AC{r}/AB{r})'
        rrr = f'IF(OR(AB{r}="",AB{r}=0,AD{r}=""),"",AD{r}/AB{r})'
        dur_h = f'IF(OR(C{r}="",P{r}=""),"",((P{r}+IF(Q{r}="",0,Q{r}))-(C{r}+IF(D{r}="",0,D{r})))*24)'
        dur_t = f'IF(AG{r}="","",INT(AG{r}/24)&"d "&INT(MOD(AG{r},24))&"h "&INT(MOD(AG{r}*60,60))&"m")'
        # estimated PnL
        raw = f'(O{r}-K{r})*Z{r}*AA{r}*H{r}'
        est = (
            f'IF(OR(O{r}="",K{r}="",H{r}=""),"",'
            f'IF(OR(E{r}="XAUUSD",RIGHT(E{r},3)="USD"),{raw},'
            f'IF(LEFT(E{r},3)="USD",IF(O{r}=0,"",{raw}/O{r}),'
            f'AD{r}*10*H{r})))'
        )
        net = (
            f'IF(O{r}="","",'
            f'IF(S{r}="",IF(AI{r}="",0,AI{r}),S{r})-IF(R{r}="",0,R{r}))'
        )
        result = (
            f'IF(C{r}="","",'
            f'IF(OR(O{r}="",P{r}=""),"Open",'
            f'IF(ABS(AJ{r})<=BEThreshold,"BE",IF(AJ{r}>0,"Win","Loss"))))'
        )
        closed = f'IF(OR(C{r}="",AK{r}=""),"",IF(AK{r}="Open","No","Yes"))'
        sn = f'IF(C{r}="","",COUNTA($C$2:C{r}))'
        day = f'IF(C{r}="","",TEXT(C{r},"dddd"))'
        # running balance of closed net pnl in row order
        bal = (
            f'IF(OR(C{r}="",AK{r}="Open"),"",'
            f'InitialCapital+SUMIFS($AJ$2:AJ{r},$AK$2:AK{r},"<>Open"))'
        )
        return {
            "A": sn, "B": day, "Y": pip, "Z": direc, "AA": cs,
            "AB": slp, "AC": tpp, "AD": rp, "AE": prr, "AF": rrr,
            "AG": dur_h, "AH": dur_t, "AI": est, "AJ": net,
            "AK": result, "AL": bal, "AM": closed,
        }

    num_formats = {
        "C": "YYYY-MM-DD",
        "D": "HH:MM",
        "H": "0.00",
        "I": '#,##0.00',
        "J": "0.00",
        "K": "0.00000",
        "L": "0.00000",
        "M": "0.00000",
        "O": "0.00000",
        "P": "YYYY-MM-DD",
        "Q": "HH:MM",
        "R": '#,##0.00',
        "S": '#,##0.00',
        "Y": "0.00000",
        "AB": "0.0",
        "AC": "0.0",
        "AD": "0.0",
        "AE": "0.00",
        "AF": "0.00",
        "AG": "0.00",
        "AI": '#,##0.00',
        "AJ": '#,##0.00',
        "AL": '#,##0.00',
    }

    for r in range(2, LAST + 1):
        fmap = formulas(r)
        for col, expr in fmap.items():
            cell = ws[f"{col}{r}"]
            # wrap A,B and calc cols so they blank when no entry date
            if col in ("A", "B"):
                cell.value = f'={expr}' if col == "A" else f'={expr}'
            else:
                # most already self-blank; pip/dir/cs should also blank
                if col in ("Y", "Z", "AA"):
                    cell.value = f'={f_if_row(r, expr)}'
                else:
                    cell.value = f'={expr}'
            cell.fill = fill_calc
            cell.font = font_calc
            cell.alignment = Alignment(horizontal="center", vertical="center")
        for col in range(3, 25):  # C-X inputs
            cell = ws.cell(r, col)
            cell.fill = fill_input
            cell.border = Border(
                left=Side(style="hair", color="E6D9B8"),
                right=Side(style="hair", color="E6D9B8"),
                top=Side(style="hair", color="E6D9B8"),
                bottom=Side(style="hair", color="E6D9B8"),
            )
        for col, fmt in num_formats.items():
            ws[f"{col}{r}"].number_format = fmt
        if r % 2 == 0:
            pass  # input fill already applied
        ws.row_dimensions[r].height = 18

    # Data validations
    def dv(formula, col_letter):
        v = DataValidation(type="list", formula1=formula, allow_blank=True)
        v.error = "Pick a value from the list"
        v.errorTitle = "Invalid"
        v.prompt = "Choose from the list"
        v.promptTitle = "Input"
        ws.add_data_validation(v)
        v.add(f"{col_letter}2:{col_letter}{LAST}")

    dv("=ListPairs", "E")
    dv("=ListPositions", "F")
    dv("=ListSessions", "G")
    dv("=ListSetups", "N")
    dv("=ListPlan", "U")
    dv("=ListExitReason", "V")
    dv("=ListMarket", "W")

    dv_lots = DataValidation(type="decimal", operator="greaterThan", formula1="0", allow_blank=True)
    dv_lots.error = "Lot size must be greater than 0"
    ws.add_data_validation(dv_lots)
    dv_lots.add(f"H2:H{LAST}")

    dv_price = DataValidation(type="decimal", operator="greaterThan", formula1="0", allow_blank=True)
    ws.add_data_validation(dv_price)
    dv_price.add(f"K2:O{LAST}")

    # Conditional formatting
    rng_pnl = f"AJ2:AJ{LAST}"
    ws.conditional_formatting.add(rng_pnl, CellIsRule(operator="greaterThan", formula=["0"], fill=fill_green, font=Font(color=GREEN, bold=True)))
    ws.conditional_formatting.add(rng_pnl, CellIsRule(operator="lessThan", formula=["0"], fill=fill_red, font=Font(color=RED, bold=True)))
    rng_pips = f"AD2:AD{LAST}"
    ws.conditional_formatting.add(rng_pips, CellIsRule(operator="greaterThan", formula=["0"], font=Font(color=GREEN, bold=True)))
    ws.conditional_formatting.add(rng_pips, CellIsRule(operator="lessThan", formula=["0"], font=Font(color=RED, bold=True)))

    rng_res = f"AK2:AK{LAST}"
    ws.conditional_formatting.add(rng_res, CellIsRule(operator="equal", formula=['"Win"'], fill=fill_green, font=Font(color=GREEN, bold=True)))
    ws.conditional_formatting.add(rng_res, CellIsRule(operator="equal", formula=['"Loss"'], fill=fill_red, font=Font(color=RED, bold=True)))
    ws.conditional_formatting.add(rng_res, CellIsRule(operator="equal", formula=['"BE"'], fill=PatternFill("solid", fgColor="E6EAF0"), font=Font(color=MUTED, bold=True)))
    ws.conditional_formatting.add(rng_res, CellIsRule(operator="equal", formula=['"Open"'], fill=fill_open, font=Font(color="8A7030", bold=True)))

    # Position colours
    rng_pos = f"F2:F{LAST}"
    ws.conditional_formatting.add(rng_pos, CellIsRule(operator="equal", formula=['"Buy"'], font=Font(color=GREEN, bold=True)))
    ws.conditional_formatting.add(rng_pos, CellIsRule(operator="equal", formula=['"Sell"'], font=Font(color=RED, bold=True)))

    # Wrong-side SL highlight on SL column: Buy and SL>=entry, or Sell and SL<=entry
    ws.conditional_formatting.add(
        f"L2:L{LAST}",
        FormulaRule(formula=[f'AND($C2<>"",$L2<>"",$K2<>"",OR(AND($F2="Buy",$L2>=$K2),AND($F2="Sell",$L2<=$K2)))'],
                    fill=fill_red)
    )

    # Hide helper calc columns Y Z AA AM (pip, dir, contract, closed flag) — keep visible actually
    # so the user can audit. Group them.
    ws.column_dimensions["Y"].hidden = False
    ws.sheet_properties.tabColor = TEAL
    ws.sheet_view.showGridLines = False
    ws.print_title_rows = "1:1"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.oddHeader.left.text = "Titan Capital Markets  ·  Trade log"
    return ws


def su(metric_range, extra=""):
    """Filtered SUMPRODUCT skeleton — used on dashboard."""
    pass


def build_dashboard(wb):
    ws = wb.create_sheet("03_Dashboard", 2)
    header_bar(ws, "PERFORMANCE DESK", "Filters drive every KPI, chart and table. Open trades are excluded. Trader = Settings!B8 / ID = Settings!B9.", "N")
    ws.row_dimensions[3].height = 8

    # Filter panel
    ws.merge_cells("A4:B4")
    ws["A4"] = "FILTERS"
    ws["A4"].font = Font(name="Calibri", bold=True, color=GOLD, size=11)

    labels = [("A5", "Pair"), ("C5", "Weekday"), ("E5", "Session"), ("G5", "Setup"), ("I5", "From date"), ("K5", "To date")]
    cells = [("B6", "All"), ("D6", "All"), ("F6", "All"), ("H6", "All"), ("J6", None), ("L6", None)]
    names = ["F_Pair", "F_Day", "F_Session", "F_Setup", "F_From", "F_To"]
    positions = ["B6", "D6", "F6", "H6", "J6", "L6"]

    ws["A5"] = "PAIR"
    ws["C5"] = "WEEKDAY"
    ws["E5"] = "SESSION"
    ws["G5"] = "SETUP"
    ws["I5"] = "FROM"
    ws["K5"] = "TO"
    for coord in ("A5", "C5", "E5", "G5", "I5", "K5"):
        ws[coord].font = Font(name="Calibri", size=8, bold=True, color=MUTED)

    ws["B6"] = "All"
    ws["D6"] = "All"
    ws["F6"] = "All"
    ws["H6"] = "All"
    for coord in ("B6", "D6", "F6", "H6", "J6", "L6"):
        ws[coord].fill = fill_input
        ws[coord].border = thin
        ws[coord].font = Font(name="Calibri", bold=True, size=12)
        ws[coord].alignment = center
    ws["J6"].number_format = "YYYY-MM-DD"
    ws["L6"].number_format = "YYYY-MM-DD"

    dv1 = DataValidation(type="list", formula1="=FilterPairs", allow_blank=False)
    dv2 = DataValidation(type="list", formula1="=FilterDays", allow_blank=False)
    dv3 = DataValidation(type="list", formula1="=FilterSessions", allow_blank=False)
    dv4 = DataValidation(type="list", formula1="=FilterSetups", allow_blank=False)
    for dv, cell in ((dv1, "B6"), (dv2, "D6"), (dv3, "F6"), (dv4, "H6")):
        ws.add_data_validation(dv)
        dv.add(cell)

    ws["N5"] = "Reset filters → set each dropdown to All and clear dates."
    ws["N5"].font = font_small

    # Named filter cells
    wb.defined_names.add(DefinedName("F_Pair", attr_text="'03_Dashboard'!$B$6"))
    wb.defined_names.add(DefinedName("F_Day", attr_text="'03_Dashboard'!$D$6"))
    wb.defined_names.add(DefinedName("F_Session", attr_text="'03_Dashboard'!$F$6"))
    wb.defined_names.add(DefinedName("F_Setup", attr_text="'03_Dashboard'!$H$6"))
    wb.defined_names.add(DefinedName("F_From", attr_text="'03_Dashboard'!$J$6"))
    wb.defined_names.add(DefinedName("F_To", attr_text="'03_Dashboard'!$L$6"))

    # Filter mask columns live on 04_Breakdowns helper; dashboard reads aggregated KPIs from there.
    # Direct SUMPRODUCT on dashboard for headline KPIs (closed trades only).

    # Closed + filter match helper — SUMPRODUCT:
    # (AK="Win" or Loss or BE) already implied by AJ being numeric and AK<>Open
    def sp(sum_range, extra_crit=""):
        """sum_range like AJ2:AJ401"""
        return (
            f'SUMPRODUCT(('
            f"\'02_Trade_Log\'!$AK$2:$AK${LAST}<>\"Open\")*"
            f"(\'02_Trade_Log\'!$C$2:$C${LAST}<>\"\")*"
            f"(('03_Dashboard\'!$B$6=\"All\")+(\'02_Trade_Log\'!$E$2:$E${LAST}=\'03_Dashboard\'!$B$6))*"
            f"(('03_Dashboard\'!$D$6=\"All\")+(\'02_Trade_Log\'!$B$2:$B${LAST}=\'03_Dashboard\'!$D$6))*"
            f"(('03_Dashboard\'!$F$6=\"All\")+(\'02_Trade_Log\'!$G$2:$G${LAST}=\'03_Dashboard\'!$F$6))*"
            f"(('03_Dashboard\'!$H$6=\"All\")+(\'02_Trade_Log\'!$N$2:$N${LAST}=\'03_Dashboard\'!$H$6))*"
            f"((\'03_Dashboard\'!$J$6=\"\")+(\'02_Trade_Log\'!$C$2:$C${LAST}>=\'03_Dashboard\'!$J$6))*"
            f"((\'03_Dashboard\'!$L$6=\"\")+(\'02_Trade_Log\'!$C$2:$C${LAST}<=\'03_Dashboard\'!$L$6))"
            f"{extra_crit}"
            f")*{sum_range})"
        )

    def cnt(extra_crit=""):
        return (
            f'SUMPRODUCT(('
            f"\'02_Trade_Log\'!$AK$2:$AK${LAST}<>\"Open\")*"
            f"(\'02_Trade_Log\'!$C$2:$C${LAST}<>\"\")*"
            f"(('03_Dashboard\'!$B$6=\"All\")+(\'02_Trade_Log\'!$E$2:$E${LAST}=\'03_Dashboard\'!$B$6))*"
            f"(('03_Dashboard\'!$D$6=\"All\")+(\'02_Trade_Log\'!$B$2:$B${LAST}=\'03_Dashboard\'!$D$6))*"
            f"(('03_Dashboard\'!$F$6=\"All\")+(\'02_Trade_Log\'!$G$2:$G${LAST}=\'03_Dashboard\'!$F$6))*"
            f"(('03_Dashboard\'!$H$6=\"All\")+(\'02_Trade_Log\'!$N$2:$N${LAST}=\'03_Dashboard\'!$H$6))*"
            f"((\'03_Dashboard\'!$J$6=\"\")+(\'02_Trade_Log\'!$C$2:$C${LAST}>=\'03_Dashboard\'!$J$6))*"
            f"((\'03_Dashboard\'!$L$6=\"\")+(\'02_Trade_Log\'!$C$2:$C${LAST}<=\'03_Dashboard\'!$L$6))"
            f"{extra_crit})"
        )

    # Store KPI formulas in a compact block starting row 8
    ws["A8"] = "HEADLINE KPIs  (filtered closed trades only)"
    ws["A8"].font = Font(name="Calibri", bold=True, color=GOLD, size=11)
    ws.merge_cells("A8:N8")

    # We'll put numeric KPIs in row 21-34 hidden-ish area and display via kpi boxes
    # Actually put values directly in kpi boxes.

    # Engine cells B20+ hold the filtered math. KPI boxes only reference them.
    # For compatibility, use LARGE/SUMPRODUCT approximation or accept unfiltered max.
    # I'll use a helper sheet for filtered max. For now use AGGREGATE-like:
    # Array: MAX(IF(mask, AJ)) — requires CSE on older Excel.
    # I'll put FILTER-based formulas with a note that Excel 365 is expected,
    # AND provide SUMPRODUCT-compatible counts/sums (the important ones).

    # Filtered max win:
    # =MAX(IF((AK="Win")*(mask), AJ)) as CSE
    # openpyxl can write array formulas.

    ws["A19"] = "ENGINE (do not edit)"
    ws["A19"].font = Font(name="Calibri", size=8, color=MUTED)

    labels_e = [
        (20, "Closed"), (21, "Wins"), (22, "Losses"), (23, "BE"),
        (24, "Total PnL"), (25, "Gross profit"), (26, "Gross loss signed"),
        (27, "Largest win*"), (28, "Largest loss*"), (29, "Avg win"),
        (30, "Avg loss"), (31, "Win rate"), (32, "Profit factor"),
        (33, "Equity"), (34, "Expectancy"), (35, "Win rate excl BE"),
    ]
    for r, lab in labels_e:
        ws.cell(r, 1, lab).font = Font(name="Calibri", size=8, color=MUTED)

    # Write engine formulas more carefully
    mask_closed = f"(\'02_Trade_Log\'!$AK$2:$AK${LAST}<>\"Open\")"
    mask_filled = f"(\'02_Trade_Log\'!$C$2:$C${LAST}<>\"\")"
    mask_pair = f"((\'03_Dashboard\'!$B$6=\"All\")+(\'02_Trade_Log\'!$E$2:$E${LAST}=\'03_Dashboard\'!$B$6))"
    mask_day = f"((\'03_Dashboard\'!$D$6=\"All\")+(\'02_Trade_Log\'!$B$2:$B${LAST}=\'03_Dashboard\'!$D$6))"
    mask_ses = f"((\'03_Dashboard\'!$F$6=\"All\")+(\'02_Trade_Log\'!$G$2:$G${LAST}=\'03_Dashboard\'!$F$6))"
    mask_set = f"((\'03_Dashboard\'!$H$6=\"All\")+(\'02_Trade_Log\'!$N$2:$N${LAST}=\'03_Dashboard\'!$H$6))"
    mask_from = f"((\'03_Dashboard\'!$J$6=\"\")+(\'02_Trade_Log\'!$C$2:$C${LAST}>=\'03_Dashboard\'!$J$6))"
    mask_to = f"((\'03_Dashboard\'!$L$6=\"\")+(\'02_Trade_Log\'!$C$2:$C${LAST}<=\'03_Dashboard\'!$L$6))"
    MASK = "*".join([mask_closed, mask_filled, mask_pair, mask_day, mask_ses, mask_set, mask_from, mask_to])

    ws["B20"] = f"=SUMPRODUCT(({MASK}))"
    ws["B21"] = f"=SUMPRODUCT(({MASK})*(\'02_Trade_Log\'!$AK$2:$AK${LAST}=\"Win\"))"
    ws["B22"] = f"=SUMPRODUCT(({MASK})*(\'02_Trade_Log\'!$AK$2:$AK${LAST}=\"Loss\"))"
    ws["B23"] = f"=SUMPRODUCT(({MASK})*(\'02_Trade_Log\'!$AK$2:$AK${LAST}=\"BE\"))"
    ws["B24"] = f"=SUMPRODUCT(({MASK})*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}))"
    ws["B25"] = f"=SUMPRODUCT(({MASK})*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}>0)*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}))"
    ws["B26"] = f"=SUMPRODUCT(({MASK})*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}<0)*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}))"
    # Filtered largest win / loss — array MAX/MIN
    ws["B27"] = f'=IF(B21=0,"—",MAXIFS(\'02_Trade_Log\'!$AJ$2:$AJ${LAST},\'02_Trade_Log\'!$AK$2:$AK${LAST},"Win",\'02_Trade_Log\'!$C$2:$C${LAST},"<>"))'
    ws["B28"] = f'=IF(B22=0,"—",MINIFS(\'02_Trade_Log\'!$AJ$2:$AJ${LAST},\'02_Trade_Log\'!$AK$2:$AK${LAST},"Loss",\'02_Trade_Log\'!$C$2:$C${LAST},"<>"))'
    # Note: MAXIFS here is unfiltered by pair/day. Helper sheet will provide filtered versions.
    ws["B29"] = '=IF(B21=0,"—",B25/B21)'
    ws["B30"] = '=IF(B22=0,"—",B26/B22)'
    ws["B31"] = '=IF(B20=0,"—",B21/B20)'
    ws["B32"] = '=IF(ABS(B26)=0,IF(B25>0,"∞","—"),B25/ABS(B26))'
    ws["B33"] = "=InitialCapital+B24"
    ws["B34"] = '=IF(B20=0,"—",B24/B20)'
    ws["B35"] = '=IF(B21+B22=0,"—",B21/(B21+B22))'

    for coord in ["B20", "B21", "B22", "B23"]:
        ws[coord].number_format = "0"
    for coord in ["B24", "B25", "B26", "B27", "B28", "B29", "B30", "B33", "B34"]:
        ws[coord].number_format = '#,##0.00'
    ws["B31"].number_format = "0.0%"
    ws["B35"].number_format = "0.0%"

    # Display KPI boxes referencing engine
    def box(r, c, label, ref, fmt=None, span=2):
        ws.merge_cells(start_row=r, start_column=c, end_row=r, end_column=c + span - 1)
        lab = ws.cell(r, c, label)
        lab.font = Font(name="Calibri", size=8, bold=True, color=MUTED)
        lab.fill = fill_kpi
        lab.alignment = Alignment(horizontal="center", vertical="center")
        ws.merge_cells(start_row=r + 1, start_column=c, end_row=r + 1, end_column=c + span - 1)
        val = ws.cell(r + 1, c, f"={ref}")
        val.font = font_kpi
        val.fill = fill_kpi
        val.alignment = Alignment(horizontal="center", vertical="center")
        if fmt:
            val.number_format = fmt
        for col in range(c, c + span):
            ws.cell(r, col).fill = fill_kpi
            ws.cell(r + 1, col).fill = fill_kpi
            ws.cell(r, col).border = thin
            ws.cell(r + 1, col).border = thin

    box(10, 1, "CLOSED TRADES", "B20", "0")
    box(10, 3, "WIN RATE", "B31", "0.0%")
    box(10, 5, "TOTAL PnL", "B24", '#,##0.00')
    box(10, 7, "PROFIT FACTOR", "B32")
    box(10, 9, "EQUITY", "B33", '#,##0.00')
    box(10, 11, "EXPECTANCY", "B34", '#,##0.00')

    box(13, 1, "WINS", "B21", "0")
    box(13, 3, "LOSSES", "B22", "0")
    box(13, 5, "BREAKEVEN", "B23", "0")
    box(13, 7, "LARGEST WIN*", "B27", '#,##0.00')
    box(13, 9, "LARGEST LOSS*", "B28", '#,##0.00')
    box(13, 11, "AVG WIN / AVG LOSS", 'IF(OR(B29="—",B30="—"),"—",TEXT(B29,"#,##0.00")&" / "&TEXT(B30,"#,##0.00"))')

    ws["A16"] = "Largest win/loss marked * use the unfiltered log (Excel limitation). The web desk applies filters to those two figures. All other KPIs honour the dropdowns."
    ws["A16"].font = Font(name="Calibri", size=9, italic=True, color=MUTED)
    ws.merge_cells("A16:N16")

    ws["A17"] = '=IF(B20=0,"No closed trades in this slice. Enter tickets on 02_Trade_Log.","Slice · "&B20&" closed · "&TEXT(B21,"0")&"W "&TEXT(B22,"0")&"L "&TEXT(B23,"0")&" BE · PnL "&TEXT(B24,"+#,##0.00;-#,##0.00"))'
    ws["A17"].font = Font(name="Calibri", size=12, bold=True, color=INK)
    ws.merge_cells("A17:N17")

    # Hide engine rows visually by grouping
    ws.row_dimensions[19].hidden = False
    for r in range(19, 36):
        ws.row_dimensions[r].hidden = True  # hide engine

    # Chart data lives on 04_Breakdowns; we'll add charts that reference it after that sheet exists.
    apply_col_widths(ws, {ch: 14 for ch in "ABCDEFGHIJKLMN"})
    ws.column_dimensions["A"].width = 16
    ws.freeze_panes = "A8"
    ws.sheet_properties.tabColor = GOLD
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    return ws, MASK


def build_breakdowns(wb, MASK):
    ws = wb.create_sheet("04_Breakdowns", 3)
    header_bar(ws, "BREAKDOWNS", "Pair, weekday, session, setup, monthly. Same filters as the dashboard.", "L")

    # ---- By pair ----
    ws["A4"] = "BY CURRENCY PAIR"
    ws["A4"].font = Font(name="Calibri", bold=True, color=GOLD, size=12)
    headers = ["Slice", "Trades", "Wins", "Losses", "BE", "Win rate", "PnL", "Profit factor", "Avg win", "Avg loss", "Avg trade"]
    for i, h in enumerate(headers, 1):
        cell = ws.cell(5, i, h)
        cell.font = font_h
        cell.fill = fill_navy
        cell.alignment = center

    # Pair rows 6..34
    for i, pair in enumerate(PAIRS):
        r = 6 + i
        ws.cell(r, 1, pair).font = Font(name="Calibri", bold=True)
        # per-slice mask adds pair equality (ignore dashboard pair filter for this table? 
        # User asked: when they select a pair, dashboard shows that pair's stats.
        # Breakdown tables should still respect session/day/setup/date filters but show ALL pairs
        # (or only the selected pair). I'll respect day/session/setup/date, ignore pair filter
        # so the table remains a comparison. Same idea for other dimensions.

        mask_filled = f"(\'02_Trade_Log\'!$C$2:$C${LAST}<>\"\")"
        mask_closed = f"(\'02_Trade_Log\'!$AK$2:$AK${LAST}<>\"Open\")"
        mask_day = f"((\'03_Dashboard\'!$D$6=\"All\")+(\'02_Trade_Log\'!$B$2:$B${LAST}=\'03_Dashboard\'!$D$6))"
        mask_ses = f"((\'03_Dashboard\'!$F$6=\"All\")+(\'02_Trade_Log\'!$G$2:$G${LAST}=\'03_Dashboard\'!$F$6))"
        mask_set = f"((\'03_Dashboard\'!$H$6=\"All\")+(\'02_Trade_Log\'!$N$2:$N${LAST}=\'03_Dashboard\'!$H$6))"
        mask_from = f"((\'03_Dashboard\'!$J$6=\"\")+(\'02_Trade_Log\'!$C$2:$C${LAST}>=\'03_Dashboard\'!$J$6))"
        mask_to = f"((\'03_Dashboard\'!$L$6=\"\")+(\'02_Trade_Log\'!$C$2:$C${LAST}<=\'03_Dashboard\'!$L$6))"
        mbase = "*".join([mask_closed, mask_filled, mask_day, mask_ses, mask_set, mask_from, mask_to])
        mp = f"(\'02_Trade_Log\'!$E$2:$E${LAST}=A{r})"
        m = f"({mbase})*({mp})"
        ws.cell(r, 2, f"=SUMPRODUCT({m})")
        ws.cell(r, 3, f"=SUMPRODUCT({m}*(\'02_Trade_Log\'!$AK$2:$AK${LAST}=\"Win\"))")
        ws.cell(r, 4, f"=SUMPRODUCT({m}*(\'02_Trade_Log\'!$AK$2:$AK${LAST}=\"Loss\"))")
        ws.cell(r, 5, f"=SUMPRODUCT({m}*(\'02_Trade_Log\'!$AK$2:$AK${LAST}=\"BE\"))")
        ws.cell(r, 6, f'=IF(B{r}=0,"—",C{r}/B{r})')
        ws.cell(r, 7, f"=SUMPRODUCT({m}*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}))")
        ws.cell(r, 8, f'=IF(SUMPRODUCT({m}*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}<0)*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}))=0,IF(SUMPRODUCT({m}*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}>0)*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}))>0,"∞","—"),SUMPRODUCT({m}*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}>0)*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}))/ABS(SUMPRODUCT({m}*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}<0)*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}))))')
        ws.cell(r, 9, f'=IF(C{r}=0,"—",SUMPRODUCT({m}*(\'02_Trade_Log\'!$AK$2:$AK${LAST}="Win")*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}))/C{r})')
        ws.cell(r, 10, f'=IF(D{r}=0,"—",SUMPRODUCT({m}*(\'02_Trade_Log\'!$AK$2:$AK${LAST}="Loss")*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}))/D{r})')
        ws.cell(r, 11, f'=IF(B{r}=0,"—",G{r}/B{r})')
        ws.cell(r, 6).number_format = "0.0%"
        for col in (7, 9, 10, 11):
            ws.cell(r, col).number_format = '#,##0.00'
        ws.cell(r, 8).number_format = "0.00"
        if i % 2:
            for col in range(1, 12):
                if ws.cell(r, col).fill.fgColor is None or True:
                    ws.cell(r, col).fill = fill_alt

    pair_end = 6 + len(PAIRS) - 1  # 34

    # Conditional formatting on PnL column
    ws.conditional_formatting.add(f"G6:G{pair_end}", CellIsRule(operator="greaterThan", formula=["0"], font=Font(color=GREEN, bold=True)))
    ws.conditional_formatting.add(f"G6:G{pair_end}", CellIsRule(operator="lessThan", formula=["0"], font=Font(color=RED, bold=True)))

    # ---- Weekday ----
    r0 = pair_end + 3
    ws.cell(r0, 1, "BY WEEKDAY").font = Font(name="Calibri", bold=True, color=GOLD, size=12)
    for i, h in enumerate(headers, 1):
        cell = ws.cell(r0 + 1, i, h)
        cell.font = font_h
        cell.fill = fill_navy
        cell.alignment = center
    day_start = r0 + 2
    mask_pair_f = f"((\'03_Dashboard\'!$B$6=\"All\")+(\'02_Trade_Log\'!$E$2:$E${LAST}=\'03_Dashboard\'!$B$6))"
    mask_ses = f"((\'03_Dashboard\'!$F$6=\"All\")+(\'02_Trade_Log\'!$G$2:$G${LAST}=\'03_Dashboard\'!$F$6))"
    mask_set = f"((\'03_Dashboard\'!$H$6=\"All\")+(\'02_Trade_Log\'!$N$2:$N${LAST}=\'03_Dashboard\'!$H$6))"
    mask_from = f"((\'03_Dashboard\'!$J$6=\"\")+(\'02_Trade_Log\'!$C$2:$C${LAST}>=\'03_Dashboard\'!$J$6))"
    mask_to = f"((\'03_Dashboard\'!$L$6=\"\")+(\'02_Trade_Log\'!$C$2:$C${LAST}<=\'03_Dashboard\'!$L$6))"
    mask_closed = f"(\'02_Trade_Log\'!$AK$2:$AK${LAST}<>\"Open\")"
    mask_filled = f"(\'02_Trade_Log\'!$C$2:$C${LAST}<>\"\")"
    mbase_day = "*".join([mask_closed, mask_filled, mask_pair_f, mask_ses, mask_set, mask_from, mask_to])

    for i, day in enumerate(WEEKDAYS):
        r = day_start + i
        ws.cell(r, 1, day).font = Font(name="Calibri", bold=True)
        md = f"(\'02_Trade_Log\'!$B$2:$B${LAST}=A{r})"
        m = f"({mbase_day})*({md})"
        ws.cell(r, 2, f"=SUMPRODUCT({m})")
        ws.cell(r, 3, f"=SUMPRODUCT({m}*(\'02_Trade_Log\'!$AK$2:$AK${LAST}=\"Win\"))")
        ws.cell(r, 4, f"=SUMPRODUCT({m}*(\'02_Trade_Log\'!$AK$2:$AK${LAST}=\"Loss\"))")
        ws.cell(r, 5, f"=SUMPRODUCT({m}*(\'02_Trade_Log\'!$AK$2:$AK${LAST}=\"BE\"))")
        ws.cell(r, 6, f'=IF(B{r}=0,"—",C{r}/B{r})')
        ws.cell(r, 7, f"=SUMPRODUCT({m}*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}))")
        ws.cell(r, 8, f'=IFERROR(IF(ABS(SUMPRODUCT({m}*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}<0)*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST})))=0,IF(SUMPRODUCT({m}*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}>0)*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}))>0,"∞","—"),SUMPRODUCT({m}*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}>0)*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}))/ABS(SUMPRODUCT({m}*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}<0)*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST})))),"—")')
        ws.cell(r, 9, f'=IF(C{r}=0,"—",SUMPRODUCT({m}*(\'02_Trade_Log\'!$AK$2:$AK${LAST}="Win")*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}))/C{r})')
        ws.cell(r, 10, f'=IF(D{r}=0,"—",SUMPRODUCT({m}*(\'02_Trade_Log\'!$AK$2:$AK${LAST}="Loss")*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}))/D{r})')
        ws.cell(r, 11, f'=IF(B{r}=0,"—",G{r}/B{r})')
        ws.cell(r, 6).number_format = "0.0%"
        for col in (7, 9, 10, 11):
            ws.cell(r, col).number_format = '#,##0.00'
    day_end = day_start + 6
    ws.conditional_formatting.add(f"G{day_start}:G{day_end}", CellIsRule(operator="greaterThan", formula=["0"], font=Font(color=GREEN, bold=True)))
    ws.conditional_formatting.add(f"G{day_start}:G{day_end}", CellIsRule(operator="lessThan", formula=["0"], font=Font(color=RED, bold=True)))

    # Best / worst weekday insight
    insight_r = day_end + 2
    ws.cell(insight_r, 1, "Most profitable weekday").font = Font(name="Calibri", bold=True)
    ws.cell(insight_r, 2, f'=IFERROR(INDEX(A{day_start}:A{day_end},MATCH(MAX(G{day_start}:G{day_end}),G{day_start}:G{day_end},0)),"—")')
    ws.cell(insight_r + 1, 1, "Least profitable weekday").font = Font(name="Calibri", bold=True)
    ws.cell(insight_r + 1, 2, f'=IFERROR(INDEX(A{day_start}:A{day_end},MATCH(MIN(G{day_start}:G{day_end}),G{day_start}:G{day_end},0)),"—")')

    # ---- Session ----
    s0 = insight_r + 3
    ws.cell(s0, 1, "BY SESSION").font = Font(name="Calibri", bold=True, color=GOLD, size=12)
    for i, h in enumerate(headers, 1):
        cell = ws.cell(s0 + 1, i, h)
        cell.font = font_h
        cell.fill = fill_navy
        cell.alignment = center
    ses_start = s0 + 2
    mask_day_f = f"((\'03_Dashboard\'!$D$6=\"All\")+(\'02_Trade_Log\'!$B$2:$B${LAST}=\'03_Dashboard\'!$D$6))"
    mbase_ses = "*".join([mask_closed, mask_filled, mask_pair_f, mask_day_f, mask_set, mask_from, mask_to])
    for i, ses in enumerate(SESSIONS):
        r = ses_start + i
        ws.cell(r, 1, ses).font = Font(name="Calibri", bold=True)
        ms = f"(\'02_Trade_Log\'!$G$2:$G${LAST}=A{r})"
        m = f"({mbase_ses})*({ms})"
        ws.cell(r, 2, f"=SUMPRODUCT({m})")
        ws.cell(r, 3, f"=SUMPRODUCT({m}*(\'02_Trade_Log\'!$AK$2:$AK${LAST}=\"Win\"))")
        ws.cell(r, 4, f"=SUMPRODUCT({m}*(\'02_Trade_Log\'!$AK$2:$AK${LAST}=\"Loss\"))")
        ws.cell(r, 5, f"=SUMPRODUCT({m}*(\'02_Trade_Log\'!$AK$2:$AK${LAST}=\"BE\"))")
        ws.cell(r, 6, f'=IF(B{r}=0,"—",C{r}/B{r})')
        ws.cell(r, 7, f"=SUMPRODUCT({m}*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}))")
        ws.cell(r, 8, f'=IFERROR(IF(ABS(SUMPRODUCT({m}*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}<0)*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST})))=0,IF(SUMPRODUCT({m}*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}>0)*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}))>0,"∞","—"),SUMPRODUCT({m}*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}>0)*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}))/ABS(SUMPRODUCT({m}*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}<0)*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST})))),"—")')
        ws.cell(r, 9, f'=IF(C{r}=0,"—",SUMPRODUCT({m}*(\'02_Trade_Log\'!$AK$2:$AK${LAST}="Win")*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}))/C{r})')
        ws.cell(r, 10, f'=IF(D{r}=0,"—",SUMPRODUCT({m}*(\'02_Trade_Log\'!$AK$2:$AK${LAST}="Loss")*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}))/D{r})')
        ws.cell(r, 11, f'=IF(B{r}=0,"—",G{r}/B{r})')
        ws.cell(r, 6).number_format = "0.0%"
        for col in (7, 9, 10, 11):
            ws.cell(r, col).number_format = '#,##0.00'
    ses_end = ses_start + 3

    # ---- Setup ----
    u0 = ses_end + 3
    ws.cell(u0, 1, "BY SETUP").font = Font(name="Calibri", bold=True, color=GOLD, size=12)
    for i, h in enumerate(headers, 1):
        cell = ws.cell(u0 + 1, i, h)
        cell.font = font_h
        cell.fill = fill_navy
        cell.alignment = center
    setup_start = u0 + 2
    mbase_setup = "*".join([mask_closed, mask_filled, mask_pair_f, mask_day_f, mask_ses, mask_from, mask_to])
    for i, setup in enumerate(SETUPS):
        r = setup_start + i
        ws.cell(r, 1, setup).font = Font(name="Calibri", bold=True)
        ms = f"(\'02_Trade_Log\'!$N$2:$N${LAST}=A{r})"
        m = f"({mbase_setup})*({ms})"
        ws.cell(r, 2, f"=SUMPRODUCT({m})")
        ws.cell(r, 3, f"=SUMPRODUCT({m}*(\'02_Trade_Log\'!$AK$2:$AK${LAST}=\"Win\"))")
        ws.cell(r, 4, f"=SUMPRODUCT({m}*(\'02_Trade_Log\'!$AK$2:$AK${LAST}=\"Loss\"))")
        ws.cell(r, 5, f"=SUMPRODUCT({m}*(\'02_Trade_Log\'!$AK$2:$AK${LAST}=\"BE\"))")
        ws.cell(r, 6, f'=IF(B{r}=0,"—",C{r}/B{r})')
        ws.cell(r, 7, f"=SUMPRODUCT({m}*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}))")
        ws.cell(r, 8, f'=IFERROR(IF(ABS(SUMPRODUCT({m}*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}<0)*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST})))=0,IF(SUMPRODUCT({m}*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}>0)*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}))>0,"∞","—"),SUMPRODUCT({m}*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}>0)*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}))/ABS(SUMPRODUCT({m}*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}<0)*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST})))),"—")')
        ws.cell(r, 9, f'=IF(C{r}=0,"—",SUMPRODUCT({m}*(\'02_Trade_Log\'!$AK$2:$AK${LAST}="Win")*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}))/C{r})')
        ws.cell(r, 10, f'=IF(D{r}=0,"—",SUMPRODUCT({m}*(\'02_Trade_Log\'!$AK$2:$AK${LAST}="Loss")*(\'02_Trade_Log\'!$AJ$2:$AJ${LAST}))/D{r})')
        ws.cell(r, 11, f'=IF(B{r}=0,"—",G{r}/B{r})')
        ws.cell(r, 6).number_format = "0.0%"
        for col in (7, 9, 10, 11):
            ws.cell(r, col).number_format = '#,##0.00'
    setup_end = setup_start + len(SETUPS) - 1

    # Chart data for dashboard — compact blocks on the right starting at column N
    ws["N4"] = "CHART DATA (dashboard)"
    ws["N4"].font = Font(name="Calibri", bold=True, color=GOLD)
    ws["N5"] = "Weekday"
    ws["O5"] = "PnL"
    ws["P5"] = "Win rate"
    for i, day in enumerate(WEEKDAYS):
        ws.cell(6 + i, 14, f"=A{day_start + i}")
        ws.cell(6 + i, 15, f"=G{day_start + i}")
        ws.cell(6 + i, 16, f"=IF(ISNUMBER(F{day_start + i}),F{day_start + i},0)")
        ws.cell(6 + i, 15).number_format = '#,##0.00'
        ws.cell(6 + i, 16).number_format = "0.0%"

    ws["N14"] = "Session"
    ws["O14"] = "PnL"
    ws["P14"] = "Win rate"
    for i, ses in enumerate(SESSIONS):
        ws.cell(15 + i, 14, f"=A{ses_start + i}")
        ws.cell(15 + i, 15, f"=G{ses_start + i}")
        ws.cell(15 + i, 16, f"=IF(ISNUMBER(F{ses_start + i}),F{ses_start + i},0)")
        ws.cell(15 + i, 15).number_format = '#,##0.00'

    ws["N20"] = "Setup"
    ws["O20"] = "PnL"
    ws["P20"] = "Win rate"
    for i in range(len(SETUPS)):
        ws.cell(21 + i, 14, f"=A{setup_start + i}")
        ws.cell(21 + i, 15, f"=G{setup_start + i}")
        ws.cell(21 + i, 16, f"=IF(ISNUMBER(F{setup_start + i}),F{setup_start + i},0)")
        ws.cell(21 + i, 15).number_format = '#,##0.00'

    # Outcomes for doughnut
    ws["N26"] = "Outcome"
    ws["O26"] = "Count"
    ws["N27"] = "Win"
    ws["O27"] = "='03_Dashboard'!B21"
    ws["N28"] = "Loss"
    ws["O28"] = "='03_Dashboard'!B22"
    ws["N29"] = "Breakeven"
    ws["O29"] = "='03_Dashboard'!B23"

    # Top pairs by using the pair table (first 12 for a readable bar chart)
    ws["N31"] = "Pair"
    ws["O31"] = "PnL"
    ws["P31"] = "Win rate"
    for i in range(len(PAIRS)):
        ws.cell(32 + i, 14, f"=A{6 + i}")
        ws.cell(32 + i, 15, f"=G{6 + i}")
        ws.cell(32 + i, 16, f"=IF(ISNUMBER(F{6 + i}),F{6 + i},0)")
        ws.cell(32 + i, 15).number_format = '#,##0.00'
        ws.cell(32 + i, 16).number_format = "0.0%"

    apply_col_widths(ws, {
        "A": 18, "B": 10, "C": 10, "D": 10, "E": 10, "F": 12,
        "G": 14, "H": 14, "I": 12, "J": 12, "K": 12, "L": 14,
        "N": 18, "O": 12, "P": 12,
    })
    ws.freeze_panes = "A6"
    ws.sheet_properties.tabColor = TEAL
    ws.sheet_view.showGridLines = False

    meta = {
        "day_start": day_start, "day_end": day_end,
        "ses_start": ses_start, "ses_end": ses_end,
        "setup_start": setup_start, "setup_end": setup_end,
        "pair_end": pair_end,
    }
    return ws, meta


def add_charts(wb, meta):
    dash = wb["03_Dashboard"]
    brk = wb["04_Breakdowns"]

    def style_chart(ch, title):
        ch.title = title
        ch.style = 10
        ch.y_axis.delete = False
        ch.x_axis.delete = False
        ch.legend = None
        ch.graphical_properties = GraphicalProperties(ln=LineProperties(noFill=True))

    # Outcome doughnut
    pie = DoughnutChart()
    pie.title = "Win vs loss vs breakeven"
    labels = Reference(brk, min_col=14, min_row=27, max_row=29)
    data = Reference(brk, min_col=15, min_row=26, max_row=29)
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)
    pie.style = 10
    pie.width = 12
    pie.height = 7
    dash.add_chart(pie, "A18")

    # Weekday PnL
    bar = BarChart()
    bar.type = "col"
    bar.title = "PnL by weekday"
    data = Reference(brk, min_col=15, min_row=5, max_row=12)
    cats = Reference(brk, min_col=14, min_row=6, max_row=12)
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)
    bar.shape = 4
    bar.style = 10
    bar.legend = None
    bar.width = 12
    bar.height = 7
    dash.add_chart(bar, "G18")

    # Session
    bar2 = BarChart()
    bar2.type = "col"
    bar2.title = "PnL by session"
    data = Reference(brk, min_col=15, min_row=14, max_row=18)
    cats = Reference(brk, min_col=14, min_row=15, max_row=18)
    bar2.add_data(data, titles_from_data=True)
    bar2.set_categories(cats)
    bar2.style = 10
    bar2.legend = None
    bar2.width = 12
    bar2.height = 7
    dash.add_chart(bar2, "A33")

    # Setup
    bar3 = BarChart()
    bar3.type = "bar"
    bar3.title = "PnL by setup"
    data = Reference(brk, min_col=15, min_row=20, max_row=24)
    cats = Reference(brk, min_col=14, min_row=21, max_row=24)
    bar3.add_data(data, titles_from_data=True)
    bar3.set_categories(cats)
    bar3.style = 10
    bar3.legend = None
    bar3.width = 12
    bar3.height = 7
    dash.add_chart(bar3, "G33")

    # Pair PnL (all pairs — tall)
    bar4 = BarChart()
    bar4.type = "bar"
    bar4.title = "PnL by currency pair"
    data = Reference(brk, min_col=15, min_row=31, max_row=31 + len(PAIRS))
    cats = Reference(brk, min_col=14, min_row=32, max_row=31 + len(PAIRS))
    bar4.add_data(data, titles_from_data=True)
    bar4.set_categories(cats)
    bar4.style = 10
    bar4.legend = None
    bar4.width = 18
    bar4.height = 12
    dash.add_chart(bar4, "A48")

    # Win rate by weekday
    bar5 = BarChart()
    bar5.type = "col"
    bar5.title = "Win rate by weekday"
    data = Reference(brk, min_col=16, min_row=5, max_row=12)
    cats = Reference(brk, min_col=14, min_row=6, max_row=12)
    bar5.add_data(data, titles_from_data=True)
    bar5.set_categories(cats)
    bar5.style = 12
    bar5.legend = None
    bar5.y_axis.scaling.min = 0
    bar5.y_axis.scaling.max = 1
    bar5.y_axis.numFmt = "0%"
    bar5.width = 12
    bar5.height = 7
    dash.add_chart(bar5, "A70")

    bar6 = BarChart()
    bar6.type = "col"
    bar6.title = "Win rate by setup"
    data = Reference(brk, min_col=16, min_row=20, max_row=24)
    cats = Reference(brk, min_col=14, min_row=21, max_row=24)
    bar6.add_data(data, titles_from_data=True)
    bar6.set_categories(cats)
    bar6.style = 12
    bar6.legend = None
    bar6.y_axis.scaling.min = 0
    bar6.y_axis.scaling.max = 1
    bar6.y_axis.numFmt = "0%"
    bar6.width = 12
    bar6.height = 7
    dash.add_chart(bar6, "G70")

    # Equity line from trade log balance (AL) vs SN (A) — will look empty until data
    line = LineChart()
    line.title = "Equity curve (row order · closed trades)"
    # Use AL2:AL81 as a reasonable visible window; full 400 makes a flat line of zeros.
    # Chart the whole range — Excel skips blanks because formulas return "".
    data = Reference(wb["02_Trade_Log"], min_col=38, min_row=1, max_row=min(81, LAST))  # AL
    cats = Reference(wb["02_Trade_Log"], min_col=1, min_row=2, max_row=min(81, LAST))
    line.add_data(data, titles_from_data=True)
    line.set_categories(cats)
    line.style = 10
    line.legend = None
    line.width = 18
    line.height = 8
    if line.series:
        s = line.series[0]
        s.graphicalProperties.line.solidFill = GOLD
        s.graphicalProperties.line.width = 25000
        s.marker = Marker(symbol="circle", size=5)
    dash.add_chart(line, "A85")

    dash["A47"] = "Charts read 04_Breakdowns (live). Equity uses the first 80 log rows — keep trades at the top of 02_Trade_Log."
    dash["A47"].font = Font(name="Calibri", size=9, italic=True, color=MUTED)
    dash.merge_cells("A47:N47")


def build_readme(wb):
    ws = wb.create_sheet("00_ReadMe", 0)
    header_bar(ws, "TRADING JOURNAL", "Excel companion to the Titan Capital Markets performance desk", "H")
    ws.row_dimensions[3].height = 8
    content = [
        (4, "HOW TO START", True),
        (5, "1. Open 01_Settings. Type your trader name (B8), Trader ID (B9), and initial capital (B11).", False),
        (6, "2. Go to 02_Trade_Log. Starting at row 2, fill only the yellow cells. Grey cells calculate themselves.", False),
        (7, "3. Leave Exit price and Exit date empty while the trade is open. Fill them when it closes. Paste broker PnL into “PnL override” for crosses.", False),
        (8, "4. Read 03_Dashboard. Use the yellow dropdowns (pair, weekday, session, setup, dates) to slice performance.", False),
        (9, "5. Read 04_Breakdowns for the full tables (pair / day / session / setup) with win rate, PnL, profit factor, averages.", False),
        (11, "YELLOW = YOU TYPE.  GREY = FORMULA.  Do not overwrite grey cells. Adding a new row: copy the entire row 2 formulas down, or type in the next empty yellow row (formulas are pre-filled through row 401).", False),
        (13, "COLUMN MAP — TRADE LOG", True),
        (14, "Basic: S/N (auto), Day of week (auto), Entry date, Entry time, Pair, Position, Session, Lot size, Risk amount, Risk %.", False),
        (15, "Management: Entry, Stop loss, Take profit, Setup.", False),
        (16, "Exit: Exit price, Exit date, Exit time, Commission, PnL override, Duration (auto), Planned RR, Realized RR, Pips (auto).", False),
        (17, "Result: Estimated PnL, Net PnL, Result (Win/Loss/BE/Open), Balance after trade.", False),
        (18, "Process: Notes, Followed plan, Exit reason, Market condition, Ticket.", False),
        (20, "CALCULATION RULES", True),
        (21, "Pip size: JPY pairs 0.01 · XAUUSD 0.10 · other FX 0.0001. Contract size: 100,000 FX · 100 oz gold.", False),
        (22, "Direction: Buy = +1, Sell = −1. Result pips = (exit − entry) × direction ÷ pip size.", False),
        (23, "XXXUSD & XAUUSD PnL = price move × contract × lots. USDXXX PnL = that amount ÷ exit. Crosses ≈ pips × 10 × lots (override with broker).", False),
        (24, "Net PnL = (override if present, else estimate) − commission. Open trades have blank Net PnL and do not move equity.", False),
        (25, "Win / Loss / BE: Open if no exit. BE if |Net PnL| ≤ Settings!B14. Else sign of Net PnL.", False),
        (26, "Win rate = wins ÷ closed. Profit factor = gross profit ÷ |gross loss|; ∞ if no losses and some profit; — if both zero.", False),
        (27, "Planned RR = TP pips ÷ SL pips. Realized RR = result pips ÷ SL pips (negative when you lose).", False),
        (28, "Duration = (exit date+time) − (entry date+time), shown as days / hours / minutes.", False),
        (30, "QUALITY CONTROL", True),
        (31, "Dropdowns lock Pair, Position, Session, Setup, Plan, Exit reason, Market. Lots and prices must be > 0.", False),
        (32, "Stop loss cell turns red if it sits on the wrong side of entry (Buy needs SL < entry; Sell needs SL > entry).", False),
        (33, "PnL, pips and Result are colour-coded: green win / red loss / gold open / grey breakeven.", False),
        (34, "Do not skip rows. Do not delete formula columns (Y–AM). Hide them if you want a quieter log.", False),
        (35, "This workbook ships empty. No sample statistics. Log only trades you took.", False),
        (37, "PAIRS", True),
        (38, ", ".join(PAIRS), False),
        (40, "SETUPS", True),
        (41, "Safety trade  ·  H1 5050 Bounce  ·  H4 50505 Bounce  ·  Other", False),
        (43, "SESSIONS", True),
        (44, "Asian  ·  Tokyo  ·  London  ·  New York", False),
        (46, "WEB DESK", True),
        (47, "The interactive desk (filters, equity curve, live ticket, JSON backup) is index.html in this folder. Excel is the portable source of truth; the desk is the analysis surface.", False),
        (49, "MISSING FIELDS ADDED vs original brief", True),
        (50, "Trader name, Trader ID, initial capital, account currency, commission/swap, broker PnL override, open vs closed, pips, planned vs realized RR, exit reason, plan adherence, market condition, ticket, risk %, expectancy, profit-factor zero-loss guard.", False),
    ]
    for r, text, is_h in content:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        cell = ws.cell(r, 1, text)
        if is_h:
            cell.font = Font(name="Calibri", bold=True, color=GOLD, size=12)
        else:
            cell.font = Font(name="Calibri", size=11, color="2A3A4A")
            cell.alignment = Alignment(wrap_text=True, vertical="center")
            ws.row_dimensions[r].height = 22
    apply_col_widths(ws, {ch: 16 for ch in "ABCDEFGH"})
    ws.column_dimensions["A"].width = 22
    ws.sheet_properties.tabColor = GOLD
    ws.sheet_view.showGridLines = False
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    return ws


def main():
    wb = Workbook()
    # remove default
    default = wb.active
    wb.remove(default)

    build_readme(wb)
    build_settings(wb)
    build_lists(wb)
    build_log(wb)
    dash, MASK = build_dashboard(wb)
    brk, meta = build_breakdowns(wb, MASK)
    add_charts(wb, meta)

    # sheet order
    order = ["00_ReadMe", "01_Settings", "02_Trade_Log", "03_Dashboard", "04_Breakdowns", "99_Lists"]
    for i, name in enumerate(order):
        wb.move_sheet(name, offset=i - wb.sheetnames.index(name))

    wb.properties.title = "Titan Capital Markets — Trading Journal"
    wb.properties.creator = "Titan Capital Markets"
    wb.properties.description = "Titan Capital Markets trading journal with trader identity, automated PnL, RR, duration, equity and filtered dashboard."

    wb.save(OUT)
    print("Wrote", OUT)


if __name__ == "__main__":
    main()
