"""Builds Veolia_capacity_2027.xlsx: investment capacity to end-2027 under the <=3x leverage commitment.

Every hard-coded number lives on the Inputs sheet with its source (document + PDF page).
All results are Excel formulas; recalculate (Excel, or LibreOffice) after editing an input.
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.comments import Comment

OUT = "Veolia_capacity_2027.xlsx"
F = "Arial"
BLUE, GREEN, BLACK = "0000FF", "008000", "000000"
YELLOW = PatternFill("solid", fgColor="FFFF00")
HEAD = PatternFill("solid", fgColor="1F3864")
SUB = PatternFill("solid", fgColor="D9E1F2")
EUR = '#,##0;(#,##0);"-"'
PCT = '0.0%;(0.0%);"-"'
X = '0.00"x"'

D_H126 = "Amendment to 2025 URD - Half-year report 30/06/2026"
D_FY25 = "Consolidated financial statements 31/12/2025"
D_H124 = "Amendment to 2023 URD - Half-year report 30/06/2024"
D_BRF = "Capstone Project Briefing"
D_WEB = "Moody's Credit Opinion 04/05/2026 (veolia.com) - NOT YET VERIFIED"

wb = Workbook()

def font(c, color=BLACK, bold=False, size=10, italic=False):
    c.font = Font(name=F, color=color, bold=bold, size=size, italic=italic)

def header(ws, row, labels, widths=None):
    for i, t in enumerate(labels, 1):
        c = ws.cell(row=row, column=i, value=t)
        c.font = Font(name=F, bold=True, color="FFFFFF", size=10)
        c.fill = HEAD
        c.alignment = Alignment(wrap_text=True, vertical="center")
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[L(i)].width = w

# ---------------------------------------------------------------- README
ws = wb.active
ws.title = "README"
ws.column_dimensions["A"].width = 120
lines = [
    ("Veolia - investment capacity to 31/12/2027 under the leverage commitment", True),
    ("Capstone Subject 2 - perimeter: Financial capacity. Workshop 2 working model.", False),
    ("", False),
    ("Question: how much can Veolia commit to acquisitions between 1 July 2026 and 31 December 2027 while ending 2027 at net debt / EBITDA <= 3.0x?", False),
    ("", False),
    ("How to read it", True),
    ("Inputs     - every hard-coded number, with document, PDF page and how it was derived. Blue = input, yellow = judgement call to defend.", False),
    ("Spent      - Block 1: what has gone out against the EUR 4bn GreenUp envelope, gross and net of disposals.", False),
    ("Start      - Block 2: starting position (net debt, EBITDA, leverage, maturities, liquidity) and the two net-debt figures that do not reconcile.", False),
    ("Step-by-step - the base case laid out line by line: EBITDA build, net debt roll-forward, leverage test, capacity. Start here.", False),
    ("Engine     - Blocks 3-4: one column per case. Green cells link to Inputs; blue cells in a case column are that case's overrides.", False),
    ("Summary    - the capacity range by case and the three sources of capacity (retained FCF, disposals, leverage headroom).", False),
    ("Formulas   - the four standard formulas (leverage, headroom, available cash, practical capacity) applied and corrected.", False),
    ("Tornado    - Block 5: one-way sensitivities ranked by swing, computed from the Engine.", False),
    ("Rating     - cross-check against the agency trigger (Moody's FFO/net debt). Low confidence until the PDF is read.", False),
    ("", False),
    ("Method (Engine)", True),
    ("NFD Dec-26 = NFD Jun-26 - net FCF H2-26 + dividends/coupons H2-26 + buyback - disposals H2-26 + FX/other", False),
    ("NFD Dec-27 before new M&A = NFD Dec-26 - net FCF 2027 + all dividends & hybrid coupons 2027 + buyback - disposals 2027 + FX/other", False),
    ("Max NFD Dec-27 = leverage cap x EBITDA 2027 (EBITDA net of what the disposals take away)", False),
    ("Capacity = (Max NFD - NFD before M&A) / (1 - cap x f / deal multiple), f = share of acquired EBITDA counted in the 2027 ratio (0 = none, 1 = full pro forma)", False),
    ("Definitions follow Veolia's own (H1-2026 report, PDF p.31): leverage = NFD incl. IFRS 16 / EBITDA incl. IFRS 16; net FCF is before dividends and financial investments.", False),
    ("", False),
    ("Rule followed: no number in this workbook was typed in as a result. Results are formulas over sourced inputs.", True),
]
for i, (t, b) in enumerate(lines, 1):
    c = ws.cell(row=i, column=1, value=t)
    font(c, bold=b, size=12 if i == 1 else 10)

# ---------------------------------------------------------------- INPUTS
wi = wb.create_sheet("Inputs")
header(wi, 1, ["Key", "Input", "Value", "Unit", "Source document", "PDF page", "Derivation / note", "Confidence"],
       [16, 46, 12, 9, 40, 9, 80, 11])
INPUTS = [
    # key, label, value, unit, source, page, note, confidence, judgement?
    ("sec", "Starting position"),
    ("NFD_J26", "Net financial debt 30/06/2026", 24548, "EUR m", D_H126, "26", "Closing NFD, Veolia definition (excl. Suez PPA remeasurement). Bridge table 3.3.1.", "High", 0),
    ("NFD_D25", "Net financial debt 31/12/2025 (bridge basis)", 19657, "EUR m", D_H126, "25-26", "Opening NFD in the H1-2026 bridge.", "High", 0),
    ("NFD_D25s", "Net financial debt 31/12/2025 (statutory note)", 19821, "EUR m", D_FY25, "66", "Note 8.3.2.1. Also shown as 12/31/2025 in H1-2026 key figures (p.6). Gap vs bridge basis = Suez PPA remeasurement (to confirm).", "High", 0),
    ("NFD_D24", "Net financial debt 31/12/2024", 17819, "EUR m", D_H126, "6", "Key figures and opening of H1-2025 bridge (p.26).", "High", 0),
    ("NFD_J25", "Net financial debt 30/06/2025", 20764, "EUR m", D_H126, "26", "", "High", 0),
    ("NFD_D23", "Net financial debt 31/12/2023", 17903, "EUR m", D_H124, "6", "Key figures, 12/31/2023 column.", "High", 0),
    ("EBITDA24", "EBITDA FY2024", 6788, "EUR m", D_H126, "6", "", "High", 0),
    ("EBITDA25", "EBITDA FY2025", 7050, "EUR m", D_H126, "6", "", "High", 0),
    ("EBITDA_H124", "EBITDA H1 2024", 3266, "EUR m", D_H124, "6", "As published in 2024 (not restated).", "Medium", 0),
    ("EBITDA_H125", "EBITDA H1 2025", 3367, "EUR m", D_H126, "6", "", "High", 0),
    ("EBITDA_H126", "EBITDA H1 2026", 3552, "EUR m", D_H126, "6", "", "High", 0),
    ("sec", "Net free cash flow (Veolia definition: before dividends, financial investments and divestitures)"),
    ("FCF23", "Net FCF FY2023", 1143, "EUR m", D_H124, "6", "", "High", 0),
    ("FCF24", "Net FCF FY2024", 1156, "EUR m", D_H126, "6", "", "High", 0),
    ("FCF25", "Net FCF FY2025", 1178, "EUR m", D_H126, "6", "", "High", 0),
    ("FCF_H125", "Net FCF H1 2025", -451, "EUR m", D_H126, "26", "", "High", 0),
    ("FCF_H126", "Net FCF H1 2026 (actual)", -288, "EUR m", D_H126, "26", "", "High", 0),
    ("FCF_g", "Net FCF growth per year 2026-27", "=(C_FCF25/C_FCF23)^(1/2)-1", "%", "Derived", "", "Base = 2023-25 CAGR of net FCF (formula). Net FCF has been flat while EBITDA grew: capex and WCR absorb the growth.", "Medium", 1),
    ("CE_FCF27", "Clean Earth incremental net FCF 2027", 0, "EUR m", "Judgement", "", "EBITDA ~EUR 0.2bn less capex, cash tax and interest on ~EUR 2.8bn of acquisition debt (net FCF is after interest) ~ zero. Replace with Enviri 10-K data.", "Low", 1),
    ("sec", "Shareholder and hybrid outflows"),
    ("DivPar26", "Parent dividend paid 2026 (on FY2025)", 1099, "EUR m", D_H126, "14", "EUR 1.50/share paid 13/05/2026.", "High", 0),
    ("DivGrowth", "Parent dividend growth 2027", 0.08, "%", D_H126, "29", "Guidance: dividend growth in line with current EPS, current net income +8% minimum.", "Medium", 1),
    ("DivMin", "Dividends to minorities per year", "=C_DivCF25-C_DivPar25", "EUR m", D_FY25 + " / " + D_H126, "7-8 / 6", "FY2025 dividends paid (cash flow statement, 1,280) less parent dividend (H1-2025 key figures, 1,023).", "Medium", 0),
    ("DivCF25", "Dividends paid FY2025, cash flow statement", 1280, "EUR m", D_FY25, "8", "Parent + minorities.", "High", 0),
    ("DivPar25", "Parent dividend paid 2025", 1023, "EUR m", D_H126, "6", "H1-2025 column. NB FY2025 column of the same table shows 895: inconsistency to raise.", "Medium", 0),
    ("Coup25", "Hybrid coupons paid FY2025", 121, "EUR m", D_FY25, "8", "", "High", 0),
    ("Hyb25", "Hybrid stock end-2025", 4100, "EUR m", D_FY25, "74", "Note 9.4.2, 'EUR 4.1 billion'.", "High", 0),
    ("HybRed", "Hybrid redeemed Feb-2026", 850, "EUR m", D_H126, "14", "Already reclassified as debt at 31/12/2025 (FY25 p.74), so no further NFD impact.", "High", 0),
    ("DivBridgeH125", "Dividends + coupons in H1-2025 NFD bridge", 1267, "EUR m", D_H126, "26", "Equals 1,173 dividends + 94 coupons (cash flow statement p.36). Gives the H2 seasonal residual.", "High", 0),
    ("Buyback", "Net buyback per year (capital reduction less share issues)", "=C_CapRed25-C_ShIss25", "EUR m", D_FY25, "8", "Guidance: buyback offsets employee plan dilution. FY2025: 402 - 318.", "Medium", 0),
    ("CapRed25", "Share capital reduction FY2025", 402, "EUR m", D_FY25, "8", "", "High", 0),
    ("ShIss25", "Proceeds on issue of shares FY2025", 318, "EUR m", D_FY25, "8", "", "High", 0),
    ("sec", "Disposals (asset rotation)"),
    ("DispTotal", "Disposal programme", 2000, "EUR m", D_H126, "29", "'EUR 2bn+ disposal program will be delivered by mid 2028.' Gross or net, and whether H1-26 disposals count, is not stated.", "Medium", 0),
    ("DispShare", "Share of programme cashed by 31/12/2027", 0.75, "%", "Judgement", "", "Straight-line to mid-2028 from late 2025 gives ~80%; trimmed for back-loading of large disposals.", "Low", 1),
    ("DispShareH2", "Of which in H2 2026", 0.20, "%", "Judgement", "", "Only EUR 73m of financial disposals in H1-2026 (p.27).", "Low", 1),
    ("DispMult", "EV/EBITDA multiple of assets sold", 10, "x", "Judgement", "", "Each EUR of disposal also removes EBITDA from the ratio's denominator. No published multiple.", "Low", 1),
    ("sec", "EBITDA trajectory"),
    ("Org26", "Organic EBITDA growth 2026", 0.055, "%", D_H126, "29", "Guidance +5% to +6%, mid-point.", "Medium", 0),
    ("FX26", "FX effect on EBITDA 2026", -0.009, "%", D_H126, "22", "H1-2026 FX effect -0.9%, extrapolated to the year.", "Low", 1),
    ("Org27", "Organic EBITDA growth 2027", 0.05, "%", D_BRF, "20", "GreenUp 'organic growth of around 5% a year'.", "Medium", 1),
    ("CE_EV", "Clean Earth price", 3040, "USD m", D_BRF, "21", "Briefing table. H1-2026 report (p.14) gives USD 2,989m closing price.", "Medium", 0),
    ("CE_Mult", "Clean Earth EV / 2026e EBITDA post-synergies", 9.8, "x", D_BRF, "21", "", "Medium", 0),
    ("CE_Syn", "Clean Earth run-rate synergies (year 4)", 120, "USD m", D_BRF, "21", "", "Medium", 0),
    ("CE_SynPh27", "Share of synergies achieved in 2027", 0.25, "%", "Judgement", "", "Year 1 of 4; H1-2026 report says synergies begin in 2027 (p.13).", "Low", 1),
    ("USD_EUR", "USD per EUR at closing", "=2989/2542", "USD/EUR", D_H126, "14", "USD 2,989m = EUR 2,542m on 01/06/2026.", "High", 0),
    ("CE_Months26", "Clean Earth months consolidated in 2026", 7, "months", D_H126, "14", "Closed 1 June 2026.", "High", 0),
    ("EBITDA27_ovr", "EBITDA 2027 override (0 = use trajectory)", 0, "EUR m", D_BRF, "20", "Set to 8,000 to test the GreenUp target ('at least EUR 8bn in 2027').", "-", 1),
    ("sec", "Constraint and deal assumptions"),
    ("Lmax", "Leverage cap at 31/12/2027", 3.0, "x", D_BRF + " / " + D_H124, "20 / 12", "GreenUp: 'financial leverage less than or equal to 3x'.", "High", 0),
    ("Lguid26", "Leverage guidance end-2026", 3.0, "x", D_H126, "29", "'3x or slightly above, Clean Earth included' - used as a validation check only.", "High", 0),
    ("DealMult", "EV/EBITDA paid on new acquisitions", 10, "x", D_BRF, "21", "Anchored on ~9.8x (Clean Earth) and ~11x (WTS minorities).", "Medium", 1),
    ("ProForma", "Share of acquired EBITDA counted in 2027 ratio", 0, "%", "Judgement", "", "0 = conservative (deal closes late 2027 or ratio uses reported EBITDA). Ask Veolia whether the covenant/guidance ratio is pro forma.", "Low", 1),
    ("FX_H2", "FX / other on NFD, H2 2026", 0, "EUR m", "Judgement", "", "H1-2026: -260 (p.26). USD share of debt has risen with Clean Earth.", "Low", 1),
    ("FX_27", "FX / other on NFD, 2027", 0, "EUR m", "Judgement", "", "Positive = debt increases.", "Low", 1),
    ("sec", "Liquidity (for the cash-availability formula)"),
    ("Cash_J26", "Cash, cash equivalents and liquid assets 30/06/2026", 9332, "EUR m", D_H126, "14", "'Cash stands at EUR 9,332 million'. Equals 7,288 cash + 2,044 liquid assets (p.28).", "High", 0),
    ("Undrawn_J26", "Undrawn committed credit lines 30/06/2026", "=4500+1373", "EUR m", D_H126, "14", "Syndicated loan 4,500 (matures 2030) + bilateral lines 1,373. Counted as 'permitted borrowing'; new bond issues are not counted.", "High", 0),
    ("CurDebt_J26", "Current debt incl. overdrafts 30/06/2026 (due within 12 months)", 11374, "EUR m", D_H126, "28", "Mostly commercial paper and bonds due by June 2027. Treated as repaid, not rolled over (prudent).", "High", 0),
    ("Mat27", "Contractual debt flows due in calendar 2027", 3891, "EUR m", D_FY25, "66", "Undiscounted, incl. interest, at 31/12/2025.", "Medium", 0),
    ("Mat27H2", "Share of 2027 flows falling in H2 2027 (not already in current debt)", 0.5, "%", "Judgement", "", "Timing within 2027 not published.", "Low", 1),
    ("MinCash", "Minimum cash reserve", 2000, "EUR m", "Judgement", "", "No published floor. Veolia has held EUR 7-10bn of cash at every reporting date; EUR 2bn is a working floor to debate.", "Low", 1),
    ("Committed", "Investments already committed, not yet paid", 0, "EUR m", D_FY25, "18", "The only securities purchase commitment at 31/12/2025 was Clean Earth, paid June 2026. Industrial capex is already inside net FCF.", "Medium", 0),
    ("sec", "Rating cross-check (unverified - read the PDF before using)"),
    ("Moody_Fcst27", "Moody's forecast FFO / net debt 2027", 0.195, "%", D_WEB, "", "Search-result extract only; the PDF could not be downloaded from this environment.", "Low", 1),
    ("Moody_Trig", "Moody's downgrade threshold FFO / net debt ('below the high teens')", 0.18, "%", D_WEB, "", "'High teens' read as 18%. Agency wording, not a number: judgement.", "Low", 1),
]
REF = {}
r = 2
for row in INPUTS:
    if row[0] == "sec":
        c = wi.cell(row=r, column=1, value=row[1]); font(c, bold=True)
        for col in range(1, 9):
            wi.cell(row=r, column=col).fill = SUB
        r += 1
        continue
    key, label, val, unit, src, page, note, conf, judg = row
    REF[key] = f"Inputs!$C${r}"
    wi.cell(row=r, column=1, value=key); wi.cell(row=r, column=2, value=label)
    wi.cell(row=r, column=3, value=val)
    for col, v in ((4, unit), (5, src), (6, page), (7, note), (8, conf)):
        wi.cell(row=r, column=col, value=v)
    for col in range(1, 9):
        font(wi.cell(row=r, column=col))
        wi.cell(row=r, column=col).alignment = Alignment(wrap_text=True, vertical="top")
    vc = wi.cell(row=r, column=3)
    font(vc, color=BLACK if isinstance(val, str) and val.startswith("=") and src == "Derived" else BLUE)
    vc.number_format = {"%": PCT, "x": X}.get(unit, EUR if unit in ("EUR m", "USD m") else "0.000")
    if judg:
        vc.fill = YELLOW
    r += 1
# resolve C_key placeholders inside input formulas
for row in wi.iter_rows(min_row=2, min_col=3, max_col=3):
    c = row[0]
    if isinstance(c.value, str) and "C_" in c.value:
        v = c.value
        for k in sorted(REF, key=len, reverse=True):
            v = v.replace("C_" + k, REF[k].replace("Inputs!", ""))
        c.value = v
wi.freeze_panes = "A2"
n = r + 1
font(wi.cell(row=n, column=1, value="Legend: blue = sourced input; black = derived from other inputs; yellow fill = judgement call the group must defend."), italic=True)

def I(k):
    return REF[k]

# ---------------------------------------------------------------- SPENT (block 1)
sp = wb.create_sheet("Spent")
header(sp, 1, ["Period", "Item", "Gross acquisitions (EUR m)", "Disposals (EUR m)", "Basis", "Source", "PDF page", "Booster?"],
       [10, 52, 16, 14, 34, 44, 9, 10])
rows = [
    ("A. Named deals (net financial investment as published)", None),
    ("2024", "Friedrich Hofmann (recycling, Germany)", 315, None, "Deal price", D_H124, "29", "No"),
    ("2025", "Danubius (electricity flexibility, Hungary)", 271, None, "Net financial investment", D_FY25, "17", "Yes"),
    ("2025", "WTS 30% minority buy-out from CDPQ", 1500, None, "Purchase price (USD 1.75bn)", D_FY25, "17", "Yes"),
    ("2025", "US hazardous waste tuck-ins + Chameleon", 247, None, "Net financial investment", D_FY25, "17", "Yes"),
    ("2025", "Zeeklite (hazardous waste, Japan)", 85, None, "Net financial investment", D_FY25, "17", "Yes"),
    ("H1-26", "Enviropacific Services (Australia)", 137, None, "Net financial investment", D_H126, "27", "Yes"),
    ("H1-26", "Clean Earth (US) incl. IFRS 16 debt and costs", 2778, None, "Net financial investment", D_H126, "27", "Yes"),
    ("B. All financial flows (cash flow statement basis, comparable across years)", None),
    ("2024", "Purchases of investments + NCI partial purchases / disposals + NCI sales", "=482+32", "=949+1", "Cash flow statement", D_FY25, "7-8", ""),
    ("2025", "Purchases of investments + NCI partial purchases / disposals + NCI sales", "=702+1563", "=59+37", "Cash flow statement", D_FY25, "7-8", ""),
    ("H1-26", "Financial acquisitions / financial disposals (Veolia NFI basis)", 2942, 73, "Net financial investments", D_H126, "27", ""),
    ("2023", "FY2023 not covered by the documents loaded (only H1-2023: -241 / +183)", None, None, "GAP", D_H124, "29", ""),
]
r = 2
first_A = last_A = first_B = last_B = None
for row in rows:
    if row[1] is None:
        c = sp.cell(row=r, column=1, value=row[0]); font(c, bold=True)
        for col in range(1, 9):
            sp.cell(row=r, column=col).fill = SUB
        r += 1
        continue
    for col, v in enumerate(row, 1):
        c = sp.cell(row=r, column=col, value=v); font(c, color=BLUE if col in (3, 4) and v is not None else BLACK)
        if col in (3, 4):
            c.number_format = EUR
    if row[0] != "2023":
        if first_B is None and "flows" in str(sp.cell(row=r-1, column=1).value or "") or (first_B is None and row[4] in ("Cash flow statement",)):
            pass
    r += 1
# locate blocks
A0, A1 = 3, 9
B0, B1 = 11, 13
r += 1
summ = [
    ("Envelope announced (GreenUp 'EUR 4bn in growth investments')", "=4000", "Briefing p.20 says 'acquisitions over 2023-2027'; the 2024 GreenUp text (H1-24 p.12) says '2024-2027 growth investments'. The scope is not the same thing."),
    ("Named deals, gross", f"=SUM(C{A0}:C{A1})", ""),
    ("Named deals in boosters", f'=SUMIF(H{A0}:H{A1},"Yes",C{A0}:C{A1})', "Envelope says EUR 2bn prioritised for boosters."),
    ("All financial acquisitions 2024-H1-26, gross", f"=SUM(C{B0}:C{B1})", "Mixes CF-statement basis (2024-25) and NFI basis (H1-26): conservative order of magnitude."),
    ("All financial disposals 2024-H1-26", f"=SUM(D{B0}:D{B1})", "2024 includes SADE and Haikou (pre-programme)."),
    ("All financial acquisitions, net of disposals", "=C{g}-C{d}", ""),
    ("Envelope remaining - gross reading", "=C{e}-C{g}", "Negative = envelope already exceeded."),
    ("Envelope remaining - net reading", "=C{e}-C{n}", ""),
]
base = r
for i, (lab, f, note) in enumerate(summ):
    rr = base + i
    font(sp.cell(row=rr, column=2, value=lab), bold=True)
    c = sp.cell(row=rr, column=3, value=f.format(e=base, g=base+3, d=base+4, n=base+5)); font(c); c.number_format = EUR
    font(sp.cell(row=rr, column=5, value=note), italic=True)
sp.cell(row=base, column=3).font = Font(name=F, color=BLUE)
SPENT_GROSS = f"Spent!$C${base+6}"
SPENT_NET = f"Spent!$C${base+7}"

# ---------------------------------------------------------------- START (block 2)
st = wb.create_sheet("Start")
st.column_dimensions["A"].width = 58
for col in "BCDE":
    st.column_dimensions[col].width = 14
st.column_dimensions["F"].width = 70
header(st, 1, ["Metric (EUR m)", "Dec-24", "Jun-25", "Dec-25", "Jun-26", "Note"])
srows = [
    ("Net financial debt (Veolia definition)", I("NFD_D24"), I("NFD_J25"), I("NFD_D25"), I("NFD_J26"), "Dec-25 statutory note shows 19,821: the 164 gap must be explained in the report."),
    ("EBITDA - fiscal year / LTM", I("EBITDA24"), f"={I('EBITDA24')}-{I('EBITDA_H124')}+{I('EBITDA_H125')}", I("EBITDA25"), f"={I('EBITDA25')}-{I('EBITDA_H125')}+{I('EBITDA_H126')}", "June columns = last twelve months (FY - previous H1 + current H1)."),
    ("Leverage NFD / EBITDA", "=B2/B3", "=C2/C3", "=D2/D3", "=E2/E3", "June ratios are seasonally high (dividend and WCR outflows land in H1) - do not compare with the year-end 3x cap."),
]
for i, row in enumerate(srows, 2):
    font(st.cell(row=i, column=1, value=row[0]))
    for j, v in enumerate(row[1:5], 2):
        val = v if str(v).startswith("=") else "=" + v
        c = st.cell(row=i, column=j, value=val); font(c, color=GREEN if "Inputs" in val and "/" not in val else BLACK)
        c.number_format = X if i == 4 else EUR
    font(st.cell(row=i, column=6, value=row[5]), italic=True)
r = 7
font(st.cell(row=r, column=1, value="Debt maturity profile at 31/12/2025 (undiscounted contractual flows incl. interest)"), bold=True)
header(st, r+1, ["Line (EUR m)", "2026", "2027", "2028", "2029", "Source"])
mats = [("Bond issues", 1924, 3043, 2447, 1781), ("Other liabilities and overdrafts", 7219, 465, 728, 302), ("IFRS 16 lease debt", 499, 384, 317, 198)]
for i, m in enumerate(mats):
    rr = r + 2 + i
    font(st.cell(row=rr, column=1, value=m[0]))
    for j, v in enumerate(m[1:], 2):
        c = st.cell(row=rr, column=j, value=v); font(c, color=BLUE); c.number_format = EUR
    font(st.cell(row=rr, column=6, value=f"{D_FY25}, PDF p.66"), italic=True)
rr = r + 5
font(st.cell(row=rr, column=1, value="Total"), bold=True)
for j in range(2, 6):
    c = st.cell(row=rr, column=j, value=f"=SUM({L(j)}{r+2}:{L(j)}{r+4})"); font(c, bold=True); c.number_format = EUR
r = rr + 2
font(st.cell(row=r, column=1, value="Liquidity and rating at 30/06/2026"), bold=True)
liq = [("Total liquid assets", 15205, "H1-26 p.28"), ("Current debt incl. overdrafts", 11374, "H1-26 p.28"),
       ("Net liquidity", f"=B{r+1}-B{r+2}", "Matches 3,831 published (p.14)"),
       ("2026 bond issues (Jan EUR 2.5bn + Apr EUR 1.0bn + Jun EUR 0.4bn convertible)", "=2500+1000+400", "H1-26 p.14: 2027 maturities pre-financed in part"),
       ("Rating", "BBB / Baa1, stable", "S&P 27/04/2026, Moody's 04/05/2026 (H1-26 p.14)"),
       ("Bond/bank documentation financial covenants at parent", "None", "FY25 p.67: the 3x is a public commitment, not a covenant - the binding constraint is the rating")]
for i, (lab, v, src) in enumerate(liq, 1):
    font(st.cell(row=r+i, column=1, value=lab))
    c = st.cell(row=r+i, column=2, value=v); font(c, color=BLUE if isinstance(v, int) else BLACK); c.number_format = EUR
    font(st.cell(row=r+i, column=6, value=src), italic=True)

# ---------------------------------------------------------------- ENGINE
en = wb.create_sheet("Engine")
DRIVERS = ["NFD_J26", "FCF25", "FCF_g", "FCF_H126", "CE_FCF27", "DivPar26", "DivGrowth", "DivMin", "DivCF25", "Coup25",
           "Hyb25", "HybRed", "DivBridgeH125", "Buyback", "DispTotal", "DispShare", "DispShareH2", "DispMult", "Org26", "FX26",
           "Org27", "EBITDA25", "CE_EV", "CE_Mult", "CE_Syn", "CE_SynPh27", "USD_EUR", "CE_Months26", "EBITDA27_ovr",
           "Lmax", "Lguid26", "DealMult", "ProForma", "FX_H2", "FX_27"]
CASES = [
    ("Base", "Trajectory", {}),
    ("GreenUp met", "EBITDA 2027 = EUR 8.0bn", {"EBITDA27_ovr": 8000}),
    ("Downside", "Slower growth, half the disposals, USD/FX drag", {"Org27": 0.04, "FCF_g": 0.0, "DispShare": 0.5, "FX_27": 300, "DispMult": 8}),
    ("Upside", "Full disposals by 2027, pro forma ratio", {"Org27": 0.06, "DispShare": 1.0, "ProForma": 1.0, "DispMult": 12}),
]
SENS = [
    ("EBITDA 2027 level", "EBITDA27_ovr", 7750, 8250),
    ("Organic EBITDA growth 2027", "Org27", 0.04, 0.06),
    ("Net FCF growth p.a.", "FCF_g", -0.03, 0.05),
    ("Disposals cashed by end-2027", "DispShare", 0.5, 1.0),
    ("Multiple on assets sold", "DispMult", 7, 13),
    ("Parent dividend growth 2027", "DivGrowth", 0.04, 0.10),
    ("FX / other on NFD 2027 (EUR m)", "FX_27", 400, -400),
    ("Acquired EBITDA in 2027 ratio", "ProForma", 0.0, 1.0),
    ("Multiple paid on acquisitions", "DealMult", 8, 12),
    ("Clean Earth EV/EBITDA (sets its EBITDA)", "CE_Mult", 11.5, 8.5),
    ("Leverage cap", "Lmax", 2.9, 3.1),
]
for name, k, lo, hi in SENS:
    CASES.append((f"{name} - low", "Sensitivity", {k: lo}))
    CASES.append((f"{name} - high", "Sensitivity", {k: hi}))

en.column_dimensions["A"].width = 52
en["A1"] = "Engine - one column per case (EUR m unless stated)"; font(en["A1"], bold=True, size=12)
en["A2"] = "Case"; en["A3"] = "Description"
for i, (cn, desc, _) in enumerate(CASES):
    col = 2 + i
    en.column_dimensions[L(col)].width = 15
    c = en.cell(row=2, column=col, value=cn); c.font = Font(name=F, bold=True, color="FFFFFF"); c.fill = HEAD
    c.alignment = Alignment(wrap_text=True)
    c = en.cell(row=3, column=col, value=desc); font(c, italic=True, size=8); c.alignment = Alignment(wrap_text=True)
font(en["A2"], bold=True); font(en["A3"], bold=True)
row = 5
font(en.cell(row=4, column=1, value="Drivers (green = linked to Inputs, blue = case override)"), bold=True)
DR = {}
labels = {r_[0]: r_[1] for r_ in INPUTS if r_[0] != "sec"}
for k in DRIVERS:
    DR[k] = row
    font(en.cell(row=row, column=1, value=labels[k]))
    unit = [x for x in INPUTS if x[0] == k][0][3]
    for i, (_, _, ov) in enumerate(CASES):
        c = en.cell(row=row, column=2 + i)
        if k in ov:
            c.value = ov[k]; font(c, color=BLUE); c.fill = YELLOW
        else:
            c.value = "=" + I(k); font(c, color=GREEN)
        c.number_format = {"%": PCT, "x": X}.get(unit, EUR if unit in ("EUR m", "USD m") else "0.000")
    row += 1
row += 1
font(en.cell(row=row, column=1, value="Calculation"), bold=True); row += 1
CALC = [
    # key, label, formula template using {k} -> cell in same column, fmt, bold
    ("CE_E26", "Clean Earth EBITDA 2026e, full year (EUR m)", "=({CE_EV}/{CE_Mult}-{CE_Syn})/{USD_EUR}", EUR, 0),
    ("Syn27", "Clean Earth synergies in 2027 (EUR m)", "={CE_Syn}*{CE_SynPh27}/{USD_EUR}", EUR, 0),
    ("FCF26", "Net FCF FY2026e", "={FCF25}*(1+{FCF_g})", EUR, 0),
    ("FCF_H2", "Net FCF H2 2026e (FY less H1 actual)", "={FCF26}-{FCF_H126}", EUR, 0),
    ("FCF27", "Net FCF FY2027e", "={FCF26}*(1+{FCF_g})+{CE_FCF27}", EUR, 0),
    ("DivH2", "Dividends + coupons H2 2026 (FY25 residual)", "={DivCF25}+{Coup25}-{DivBridgeH125}", EUR, 0),
    ("Coup27", "Hybrid coupons 2027 (after Feb-26 redemption)", "={Coup25}*({Hyb25}-{HybRed})/{Hyb25}", EUR, 0),
    ("DivPar27", "Parent dividend 2027", "={DivPar26}*(1+{DivGrowth})", EUR, 0),
    ("Out27", "Total dividends, coupons, buyback 2027", "={DivPar27}+{DivMin}+{Coup27}+{Buyback}", EUR, 0),
    ("DispBy27", "Disposals cashed Jul-26 to Dec-27", "={DispTotal}*{DispShare}", EUR, 0),
    ("DispH2", "  of which H2 2026", "={DispBy27}*{DispShareH2}", EUR, 0),
    ("Disp27", "  of which 2027", "={DispBy27}-{DispH2}", EUR, 0),
    ("NFD26", "NFD 31/12/2026e", "={NFD_J26}-{FCF_H2}+{DivH2}+{Buyback}-{DispH2}+{FX_H2}", EUR, 1),
    ("E26", "EBITDA 2026e", "={EBITDA25}*(1+{Org26}+{FX26})+{CE_E26}*{CE_Months26}/12", EUR, 0),
    ("Lev26", "Leverage 31/12/2026e", "={NFD26}/{E26}", X, 1),
    ("Chk26", "Check vs guidance '3x or slightly above'", '=IF(ABS({Lev26}-{Lguid26})<=0.15,"consistent","REVIEW")', None, 0),
    ("NFD27", "NFD 31/12/2027e before any new acquisition", "={NFD26}-{FCF27}+{Out27}-{Disp27}+{FX_27}", EUR, 1),
    ("E27g", "EBITDA 2027e before disposals (trajectory or override)", "=IF({EBITDA27_ovr}>0,{EBITDA27_ovr},{EBITDA25}*(1+{Org26}+{FX26})*(1+{Org27})+{CE_E26}*(1+{Org27})+{Syn27})", EUR, 0),
    ("Lost", "EBITDA removed by disposals (H2-26 full year, 2027 half year)", "=({DispH2}+0.5*{Disp27})/{DispMult}", EUR, 0),
    ("E27", "EBITDA 2027e", "={E27g}-{Lost}", EUR, 1),
    ("Lev27", "Leverage 31/12/2027e before new acquisitions", "={NFD27}/{E27}", X, 1),
    ("MaxNFD", "Maximum NFD at cap", "={Lmax}*{E27}", EUR, 0),
    ("Head", "Leverage headroom 31/12/2027 (debt room)", "={MaxNFD}-{NFD27}", EUR, 1),
    ("Cap", "ACQUISITION CAPACITY Jul-26 to Dec-27 (EV, EUR m)", "={Head}/(1-{Lmax}*{ProForma}/{DealMult})", EUR, 1),
    ("s1", "Source 1 - retained FCF (FCF less dividends, coupons, buyback, FX)", "={FCF_H2}+{FCF27}-{DivH2}-{Buyback}-{Out27}-{FX_H2}-{FX_27}", EUR, 0),
    ("s2", "Source 2 - disposals, net of the debt room they cost", "={DispBy27}-{Lmax}*{Lost}", EUR, 0),
    ("s3", "Source 3 - leverage room vs Jun-26 debt (cap x EBITDA before disposals - NFD Jun-26)", "={Lmax}*{E27g}-{NFD_J26}", EUR, 0),
    ("s4", "Source 4 - acquired EBITDA counted in ratio (pro forma uplift)", "={Cap}-{Head}", EUR, 0),
    ("chk", "Check: sources sum to capacity", "=ROUND({s1}+{s2}+{s3}+{s4}-{Cap},6)", EUR, 0),
]
CR = {}
for key, lab, tmpl, fmt, b in CALC:
    CR[key] = row
    font(en.cell(row=row, column=1, value=lab), bold=bool(b))
    for i in range(len(CASES)):
        colL = L(2 + i)
        refs = {k: f"{colL}{v}" for k, v in {**DR, **CR}.items()}
        c = en.cell(row=row, column=2 + i, value=tmpl.format(**refs))
        font(c, bold=bool(b))
        if fmt:
            c.number_format = fmt
        if key == "Cap":
            c.fill = YELLOW
    row += 1
en.freeze_panes = "B4"

def ecell(key, i):
    return f"Engine!${L(2+i)}${CR[key]}"

# ---------------------------------------------------------------- SUMMARY
sm = wb.create_sheet("Summary", 1)
sm.column_dimensions["A"].width = 60
for col in "BCDE":
    sm.column_dimensions[col].width = 16
sm["A1"] = "Investment capacity to 31/12/2027 - summary (EUR m)"; font(sm["A1"], bold=True, size=12)
header(sm, 3, ["Line"] + [c[0] for c in CASES[:4]])
lines = [("NFD 30/06/2026 (actual)", "NFD_J26", EUR, "dr"), ("NFD 31/12/2026e", "NFD26", EUR, ""), ("Leverage 31/12/2026e", "Lev26", X, ""),
         ("Check vs 2026 guidance", "Chk26", None, ""), ("EBITDA 2027e", "E27", EUR, ""), ("NFD 31/12/2027e before new M&A", "NFD27", EUR, ""),
         ("Leverage 31/12/2027e before new M&A", "Lev27", X, ""), ("Maximum NFD at cap", "MaxNFD", EUR, ""),
         ("Leverage headroom (debt room)", "Head", EUR, ""), ("ACQUISITION CAPACITY (EV)", "Cap", EUR, ""),
         ("  Source 1 - retained FCF", "s1", EUR, ""), ("  Source 2 - disposals net of EBITDA lost", "s2", EUR, ""),
         ("  Source 3 - leverage room vs Jun-26 debt", "s3", EUR, ""), ("  Source 4 - pro forma uplift", "s4", EUR, "")]
for j, (lab, k, fmt, kind) in enumerate(lines, 4):
    font(sm.cell(row=j, column=1, value=lab), bold=k in ("Cap",))
    for i in range(4):
        ref = f"Engine!${L(2+i)}${DR[k]}" if kind == "dr" else ecell(k, i)
        c = sm.cell(row=j, column=2 + i, value="=" + ref); font(c, color=GREEN, bold=k == "Cap")
        if fmt:
            c.number_format = fmt
        if k == "Cap":
            c.fill = YELLOW
j += 2
font(sm.cell(row=j, column=1, value="Range across the four cases"), bold=True)
font(sm.cell(row=j+1, column=1, value="Low"))
c = sm.cell(row=j+1, column=2, value=f"=MIN(B{4+9}:E{4+9})"); font(c, bold=True); c.number_format = EUR
font(sm.cell(row=j+2, column=1, value="High"))
c = sm.cell(row=j+2, column=2, value=f"=MAX(B{4+9}:E{4+9})"); font(c, bold=True); c.number_format = EUR
font(sm.cell(row=j+4, column=1, value="Block 1 cross-reference: EUR 4bn envelope remaining (gross / net readings)"), bold=True)
c = sm.cell(row=j+5, column=2, value="=" + SPENT_GROSS); font(c, color=GREEN); c.number_format = EUR
c = sm.cell(row=j+5, column=3, value="=" + SPENT_NET); font(c, color=GREEN); c.number_format = EUR
font(sm.cell(row=j+5, column=1, value="Gross / net of disposals"))

# ---------------------------------------------------------------- TORNADO
tn = wb.create_sheet("Tornado")
tn.column_dimensions["A"].width = 42
for col in "BCDEFGHIJ":
    tn.column_dimensions[col].width = 14
tn["A1"] = "Block 5 - one-way sensitivity of acquisition capacity (EUR m)"; font(tn["A1"], bold=True, size=12)
tn["A2"] = "Base capacity"; font(tn["A2"])
tn["B2"] = "=" + ecell("Cap", 0); font(tn["B2"], color=GREEN, bold=True); tn["B2"].number_format = EUR
header(tn, 4, ["Driver", "Low input", "High input", "Capacity @ low", "Capacity @ high", "Swing", "Rank key"])
n0 = 5
for s, (name, k, lo, hi) in enumerate(SENS):
    rr = n0 + s
    ilo, ihi = 4 + 2 * s, 5 + 2 * s
    font(tn.cell(row=rr, column=1, value=name))
    unit = [x for x in INPUTS if x[0] == k][0][3]
    nf = {"%": PCT, "x": X}.get(unit, EUR)
    for col, v in ((2, f"=Engine!{L(2+ilo)}{DR[k]}"), (3, f"=Engine!{L(2+ihi)}{DR[k]}")):
        c = tn.cell(row=rr, column=col, value=v); font(c, color=GREEN); c.number_format = nf
    for col, idx in ((4, ilo), (5, ihi)):
        c = tn.cell(row=rr, column=col, value="=" + ecell("Cap", idx)); font(c, color=GREEN); c.number_format = EUR
    c = tn.cell(row=rr, column=6, value=f"=ABS(E{rr}-D{rr})"); font(c); c.number_format = EUR
    c = tn.cell(row=rr, column=7, value=f"=F{rr}+ROW()/1000000"); font(c, size=8)
n1 = n0 + len(SENS) - 1
rr = n1 + 3
font(tn.cell(row=rr - 1, column=1, value="Ranked by swing"), bold=True)
header(tn, rr, ["Rank / driver", "Capacity @ low", "Capacity @ high", "Swing"])
for s in range(len(SENS)):
    q = rr + 1 + s
    m = f"MATCH(LARGE($G${n0}:$G${n1},{s+1}),$G${n0}:$G${n1},0)"
    font(tn.cell(row=q, column=1, value=f'="{s+1}. "&INDEX($A${n0}:$A${n1},{m})'))
    for col, src in ((2, "D"), (3, "E"), (4, "F")):
        c = tn.cell(row=q, column=col, value=f"=INDEX(${src}${n0}:${src}${n1},{m})"); font(c); c.number_format = EUR

# ---------------------------------------------------------------- STEP-BY-STEP (base case, every line visible)
sb = wb.create_sheet("Step-by-step", 2)
for col, w in zip("ABCDEFG", [6, 50, 15, 15, 15, 70, 34]):
    sb.column_dimensions[col].width = w
sb["A1"] = "Step-by-step calculation, base case (EUR m unless stated)"; font(sb["A1"], bold=True, size=12)
sb["A2"] = "Every number is a formula over the Inputs sheet. Change an input there and this sheet recalculates. Column C holds 2026: EBITDA is the full year, cash flows are July-December only (H1 is already in the June debt)."
font(sb["A2"], italic=True)
header(sb, 4, ["Step", "Line", "2026", "2027", "Total Jul-26 to Dec-27", "How it is calculated", "Source / input"])
SROWS = [
    ("h", "A. Operating profit (EBITDA) - the denominator of the leverage test"),
    ("a1", "A1", "EBITDA of the existing business, prior year", f"={I('EBITDA25')}", "=C{a4}", None, EUR, "2026: FY2025 EBITDA. 2027: the 2026 figure from A4.", "Inputs: EBITDA FY2025 (H1-26 report p.6)"),
    ("a2", "A2", "Organic growth", f"={I('Org26')}", f"={I('Org27')}", None, PCT, "2026: guidance mid-point. 2027: GreenUp ~5%.", "Inputs: Org26, Org27"),
    ("a3", "A3", "Currency effect", f"={I('FX26')}", "=0", None, PCT, "2026: H1 effect extrapolated. 2027: none assumed.", "Inputs: FX26"),
    ("a4", "A4", "EBITDA of the existing business", "=C{a1}*(1+C{a2}+C{a3})", "=D{a1}*(1+D{a2}+D{a3})", None, EUR, "A1 x (1 + A2 + A3)", ""),
    ("a5", "A5", "Clean Earth EBITDA, full year", f"=({I('CE_EV')}/{I('CE_Mult')}-{I('CE_Syn')})/{I('USD_EUR')}", f"=C{{a5}}*(1+{I('Org27')})", None, EUR, "(Price / EV-to-EBITDA multiple - run-rate synergies) / USD per EUR. 2027 grows with A2.", "Inputs: CE_EV, CE_Mult, CE_Syn, USD_EUR (briefing p.21)"),
    ("a6", "A6", "Months of Clean Earth consolidated", f"={I('CE_Months26')}", "=12", None, "0", "Closed 1 June 2026.", "Inputs: CE_Months26"),
    ("a7", "A7", "Clean Earth contribution", "=C{a5}*C{a6}/12", "=D{a5}*D{a6}/12", None, EUR, "A5 x A6 / 12", ""),
    ("a8", "A8", "Clean Earth synergies", "=0", f"={I('CE_Syn')}*{I('CE_SynPh27')}/{I('USD_EUR')}", None, EUR, "Run-rate synergies x share achieved in 2027 / USD per EUR", "Inputs: CE_Syn, CE_SynPh27"),
    ("a9", "A9", "EBITDA lost through asset sales", "=0", f"=-(C{{b8}}+0.5*D{{b8}})/{I('DispMult')}", None, EUR, "-(H2-26 sales + half of 2027 sales) / multiple of assets sold", "Inputs: DispMult"),
    ("a10", "A10", "GROUP EBITDA", "=C{a4}+C{a7}+C{a8}+C{a9}", "=D{a4}+D{a7}+D{a8}+D{a9}", None, EUR, "A4 + A7 + A8 + A9", ""),
    ("h", "B. Net debt roll-forward - the numerator of the leverage test"),
    ("b1", "B1", "Net debt, opening", f"={I('NFD_J26')}", "=C{b11}", None, EUR, "2026: actual at 30 June. 2027: closing 2026 from B11.", "Inputs: NFD_J26 (H1-26 report p.26)"),
    ("b2", "B2", "Net free cash flow, full year", f"={I('FCF25')}*(1+{I('FCF_g')})", f"=C{{b2}}*(1+{I('FCF_g')})+{I('CE_FCF27')}", None, EUR, "Prior year x (1 + growth). Net FCF is after capex, interest and tax, before dividends.", "Inputs: FCF25, FCF_g, CE_FCF27"),
    ("b3", "B3", "Less: net FCF already in the June debt (H1 2026)", f"=-{I('FCF_H126')}", "=0", None, EUR, "H1 2026 net FCF was negative (-288), so H2 is the full year plus 288.", "Inputs: FCF_H126 (H1-26 report p.26)"),
    ("b4", "B4", "Net free cash flow in the period", "=C{b2}+C{b3}", "=D{b2}+D{b3}", "=C{b4}+D{b4}", EUR, "B2 + B3", ""),
    ("b5", "B5", "Parent dividend", "=0", f"={I('DivPar26')}*(1+{I('DivGrowth')})", "=C{b5}+D{b5}", EUR, "Paid in May. 2027 = 2026 dividend x (1 + growth).", "Inputs: DivPar26, DivGrowth"),
    ("b6", "B6", "Minority dividends and hybrid coupons", f"={I('DivCF25')}+{I('Coup25')}-{I('DivBridgeH125')}", f"={I('DivMin')}+{I('Coup25')}*({I('Hyb25')}-{I('HybRed')})/{I('Hyb25')}", "=C{b6}+D{b6}", EUR, "H2-26: what FY2025 paid beyond H1-2025. 2027: minorities + coupons scaled to the smaller hybrid stock.", "Inputs: DivCF25, Coup25, DivBridgeH125, DivMin, Hyb25, HybRed"),
    ("b7", "B7", "Net share buyback", f"={I('Buyback')}", f"={I('Buyback')}", "=C{b7}+D{b7}", EUR, "Capital reduction less employee share issue (FY2025 pattern).", "Inputs: Buyback"),
    ("b8", "B8", "Asset sale proceeds", f"={I('DispTotal')}*{I('DispShare')}*{I('DispShareH2')}", f"={I('DispTotal')}*{I('DispShare')}*(1-{I('DispShareH2')})", "=C{b8}+D{b8}", EUR, "Programme x share cashed by end-2027, split H2-26 / 2027.", "Inputs: DispTotal, DispShare, DispShareH2"),
    ("b9", "B9", "Currency and other effects on debt", f"={I('FX_H2')}", f"={I('FX_27')}", "=C{b9}+D{b9}", EUR, "Positive = debt goes up.", "Inputs: FX_H2, FX_27"),
    ("b10", "B10", "Change in net debt", "=-C{b4}+C{b5}+C{b6}+C{b7}-C{b8}+C{b9}", "=-D{b4}+D{b5}+D{b6}+D{b7}-D{b8}+D{b9}", "=C{b10}+D{b10}", EUR, "- B4 + B5 + B6 + B7 - B8 + B9", ""),
    ("b11", "B11", "NET DEBT, CLOSING (before any new acquisition)", "=C{b1}+C{b10}", "=D{b1}+D{b10}", None, EUR, "B1 + B10", ""),
    ("h", "C. Leverage test and acquisition capacity"),
    ("c1", "C1", "Leverage = net debt / EBITDA", "=C{b11}/C{a10}", "=D{b11}/D{a10}", None, X, "B11 / A10", ""),
    ("c2", "C2", "Leverage limit", f"={I('Lguid26')}", f"={I('Lmax')}", None, X, "2026: guidance '3x or slightly above'. 2027: the <=3x commitment.", "Inputs: Lguid26, Lmax"),
    ("c3", "C3", "Maximum net debt allowed", "=C{c2}*C{a10}", "=D{c2}*D{a10}", None, EUR, "C2 x A10", ""),
    ("c4", "C4", "Debt headroom", "=C{c3}-C{b11}", "=D{c3}-D{b11}", None, EUR, "C3 - B11. Negative in 2026 = slightly above 3x, as guided.", ""),
    ("c5", "C5", "Share of acquired EBITDA counted in the ratio", None, f"={I('ProForma')}", None, PCT, "0 = conservative; 100% if the ratio is pro forma.", "Inputs: ProForma"),
    ("c6", "C6", "EV/EBITDA paid on acquisitions", None, f"={I('DealMult')}", None, X, "", "Inputs: DealMult"),
    ("c7", "C7", "ACQUISITION CAPACITY, Jul-26 to Dec-27 (enterprise value)", None, "=D{c4}/(1-D{c2}*D{c5}/D{c6})", None, EUR, "C4 / (1 - C2 x C5 / C6): every EUR of deal adds debt, and C5/C6 of it in EBITDA.", ""),
    ("h", "D. Checks"),
    ("d1", "D1", "Check: equals Engine base case (should be 0)", None, "=ROUND(D{c7}-" + ecell("Cap", 0) + ",6)", None, EUR, "Two independent layouts of the same calculation.", ""),
    ("d2", "D2", "Check: 2026 leverage vs guidance", None, '=IF(ABS(C{c1}-C{c2})<=0.15,"consistent with guidance","REVIEW")', None, None, "Guidance: '3x or slightly above, Clean Earth included' (H1-26 p.29).", ""),
]
r = 5; SBR = {}
for row in SROWS:
    if row[0] == "h":
        r += 1 if r > 5 else 0
        SBR.setdefault("_h", []).append(r); r += 1; continue
    SBR[row[0]] = r; r += 1
r = 5; hi = 0
for row in SROWS:
    if row[0] == "h":
        r = SBR["_h"][hi]; hi += 1
        c = sb.cell(row=r, column=1, value=row[1]); font(c, bold=True)
        for col in range(1, 8):
            sb.cell(row=r, column=col).fill = SUB
        continue
    k, step, lab, fc, fd, fe, fmt, how, src = row
    rr = SBR[k]
    big = lab.isupper() or lab.startswith(("GROUP", "NET DEBT", "ACQUISITION"))
    font(sb.cell(row=rr, column=1, value=step)); font(sb.cell(row=rr, column=2, value=lab), bold=big)
    for col, f in ((3, fc), (4, fd), (5, fe)):
        if f is None:
            continue
        c = sb.cell(row=rr, column=col, value=f.format(**SBR))
        font(c, color=GREEN if "Inputs!" in f and not any(o in f.replace("Inputs!", "") for o in "+-*/(") else BLACK, bold=big)
        if fmt:
            c.number_format = fmt
        if k == "c7" and col == 4:
            c.fill = YELLOW
    font(sb.cell(row=rr, column=6, value=how), italic=True)
    font(sb.cell(row=rr, column=7, value=src), size=9)
    for col in (6, 7):
        sb.cell(row=rr, column=col).alignment = Alignment(wrap_text=True, vertical="top")
sb.freeze_panes = "C5"

# ---------------------------------------------------------------- FORMULAS (the four standard formulas, applied)
fm = wb.create_sheet("Formulas", 3)
fm.column_dimensions["A"].width = 64; fm.column_dimensions["B"].width = 14; fm.column_dimensions["C"].width = 90
fm["A1"] = "Capacity with the four standard formulas (base case, EUR m)"; font(fm["A1"], bold=True, size=12)
def E0(k):
    return f"Engine!$B${DR[k]}" if k in DR else ecell(k, 0)
FROWS = [
    ("h", "1-2. Net leverage and debt headroom TODAY (30/06/2026, last-12-months EBITDA)"),
    ("lev_now", "Net leverage = net debt / EBITDA", f"={I('NFD_J26')}/Start!E3", X, "June debt peak / LTM EBITDA."),
    ("head_now", "Debt headroom = Lmax x EBITDA - net debt", f"={I('Lmax')}*Start!E3-{I('NFD_J26')}", EUR, "Negative: on today's figures there is no room. Misleading: June is the seasonal peak and no future cash is counted."),
    ("h", "1-2. Same formulas at the TEST DATE (31/12/2027, base-case forecast before new deals)"),
    ("lev_27", "Net leverage = net debt / EBITDA", "=" + E0("Lev27"), X, "The 3x is tested at year-end on Veolia's own definition (no bank covenant at parent level: FY25 p.67)."),
    ("head_27", "Debt headroom = Lmax x EBITDA 2027 - net debt end-2027", "=" + E0("Head"), EUR, "Net debt end-2027 already includes cash from operations, dividends and disposals."),
    ("h", "3. Cash available for investment, Jul-2026 to Dec-2027"),
    ("c_open", "+ Opening cash (incl. liquid assets)", f"={I('Cash_J26')}", EUR, ""),
    ("c_gen", "+ Cash generated: net FCF H2-26 + 2027", "=" + E0("FCF_H2") + "+" + E0("FCF27"), EUR, "Net FCF is after capex, interest and tax."),
    ("c_div", "- Dividends, hybrid coupons, buyback", "=-(" + E0("DivH2") + "+" + E0("Buyback") + "+" + E0("Out27") + ")", EUR, ""),
    ("c_disp", "+ Disposal proceeds", "=" + E0("DispBy27"), EUR, ""),
    ("c_borrow", "+ Permitted borrowing (undrawn committed lines)", f"={I('Undrawn_J26')}", EUR, "Excludes new bond issues: Veolia raised EUR 3.9bn in H1-2026, so this is prudent."),
    ("c_rep", "- Repayments (current debt + H2-2027 maturities)", f"=-({I('CurDebt_J26')}+{I('Mat27')}*{I('Mat27H2')})", EUR, "Assumes nothing is refinanced."),
    ("c_comm", "- Committed spending", f"=-{I('Committed')}", EUR, ""),
    ("c_res", "- Minimum cash reserve", f"=-{I('MinCash')}", EUR, ""),
    ("c_avail", "Available cash (total)", "=SUM(B{c_open}:B{c_res})", EUR, "Before any refinancing in the bond market."),
    ("h", "4. Practical investment capacity"),
    ("f4_written", "As written: min(available cash, headroom + cash from operations) - committed", "=MIN(B{c_avail},B{head_27}+B{c_gen}+B{c_div})-" + I("Committed"), EUR, "Overstates: the end-2027 headroom already contains the cash from operations, so it is counted twice."),
    ("f4_corr", "Corrected: min(available cash, end-2027 headroom) - committed", "=MIN(B{c_avail},B{head_27})-" + I("Committed"), EUR, "Headroom taken from the full net-debt forecast at the test date, as the note to formula 4 recommends."),
    ("bind", "Binding constraint", '=IF(B{c_avail}<B{head_27},"Cash","Leverage (3x)")', None, ""),
    ("f4_pf", "If acquired EBITDA counts in the 2027 ratio (deal at the input multiple)", "=MIN(B{c_avail},B{head_27}/(1-" + I("Lmax") + "/" + I("DealMult") + "))-" + I("Committed"), EUR, "Each EUR 1 of EV buys 1/multiple of EBITDA, which raises the limit by 3/multiple."),
]
fr = 3; FR = {}
for row in FROWS:
    if row[0] == "h":
        c = fm.cell(row=fr, column=1, value=row[1]); font(c, bold=True)
        for col in range(1, 4):
            fm.cell(row=fr, column=col).fill = SUB
        fr += 1; continue
    FR[row[0]] = fr; fr += 1
for row in FROWS:
    if row[0] == "h":
        continue
    k, lab, f, fmt, note = row
    rr = FR[k]
    font(fm.cell(row=rr, column=1, value=lab), bold=k in ("c_avail", "f4_corr"))
    c = fm.cell(row=rr, column=2, value=f.format(**FR)); font(c, bold=k in ("c_avail", "f4_corr"))
    if fmt:
        c.number_format = fmt
    if k == "f4_corr":
        c.fill = YELLOW
    font(fm.cell(row=rr, column=3, value=note), italic=True)

# ---------------------------------------------------------------- RATING
ra = wb.create_sheet("Rating")
ra.column_dimensions["A"].width = 70; ra.column_dimensions["B"].width = 16; ra.column_dimensions["C"].width = 70
ra["A1"] = "Rating cross-check - is the agency trigger tighter than 3x? (LOW CONFIDENCE until PDF read)"; font(ra["A1"], bold=True, size=12)
rl = [
    ("Moody's forecast FFO/net debt 2027", f"={I('Moody_Fcst27')}", PCT, "Moody's metric, Moody's-adjusted debt: not comparable with Veolia NFD."),
    ("Downgrade threshold (judgement on 'high teens')", f"={I('Moody_Trig')}", PCT, ""),
    ("Extra debt tolerable before trigger, as % of Moody's forecast debt", "=B3/B4-1", PCT, "FFO held constant: the acquired company's FFO would add back some room."),
    ("Applied to base-case NFD 31/12/2027 before M&A (proxy)", "=B5*" + ecell("NFD27", 0), EUR, "Proxy only: assumes Moody's forecast already embeds Veolia's plan."),
    ("Compare: leverage headroom at 3.0x, base case", "=" + ecell("Head", 0), EUR, ""),
    ("Which constraint binds?", '=IF(B6<B7,"Rating trigger","3x commitment")', None, "To be redone once the Moody's and S&P reports are attached."),
]
for i, (lab, f, fmt, note) in enumerate(rl, 3):
    font(ra.cell(row=i, column=1, value=lab))
    c = ra.cell(row=i, column=2, value=f); font(c, color=GREEN if "Inputs" in f or "Engine" in f else BLACK)
    if fmt:
        c.number_format = fmt
    font(ra.cell(row=i, column=3, value=note), italic=True)

for wsx in wb.worksheets:
    wsx.sheet_view.showGridLines = False
    wsx.page_setup.orientation = "landscape"
    wsx.page_setup.paperSize = wsx.PAPERSIZE_A4
    wsx.page_setup.fitToWidth = 1
    wsx.page_setup.fitToHeight = 0
    wsx.sheet_properties.pageSetUpPr.fitToPage = True
# Print only the four headline cases of the Engine; sensitivity columns are summarised on Tornado
en.print_area = f"A1:E{en.max_row}"
wb.save(OUT)
print("saved", OUT, "cases:", len(CASES))
