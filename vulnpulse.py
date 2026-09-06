"""
VulnPulse — vulnpulse.py (formula-driven edition)
=====================================================
Same report structure/polish as the original Jun-Jul exec workbook —
Exec Summary + charts, Raw Data, Tag Leaderboard, Aging Analysis, Top Risks,
Methodology — but every number is an Excel formula (COUNTIFS/SUMIFS/
AVERAGEIFS) referencing the raw data sheet, so the workbook recalculates
if you edit or extend the underlying rows.

Sourced from 1-10 (or any number of) per-tag detailed CrowdStrike exports,
each containing BOTH open and closed findings with exact Created/Closed
dates and Status in {Open, Closed, Reopened}.

USAGE:
    python vulnpulse.py --input-dir raw_exports_by_tag \
                         --period-start 2026-07-01 --period-end 2026-07-31 \
                         --out-dir vulnpulse_reports
"""
import argparse, os, sys, glob
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import ColorScaleRule

REQUIRED = ["Hostname", "HostType", "OSVersion", "Product", "CVE ID", "Status",
            "Severity", "Created Date", "Closed Date", "Base Score", "Tags",
            "Asset Criticality", "Vulnerability ID", "Recommended Remediations"]

SLA_DAYS = {"Critical": 15, "High": 30, "Medium": 90, "Low": 180}

ARIAL = "Arial"
TITLE_FONT = Font(name=ARIAL, size=16, bold=True, color="1F4E78")
SUBTITLE_FONT = Font(name=ARIAL, size=10, italic=True, color="595959")
SECTION_FONT = Font(name=ARIAL, size=12, bold=True, color="FFFFFF")
SECTION_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(name=ARIAL, size=10, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
BODY_FONT = Font(name=ARIAL, size=10)
BOLD_BODY = Font(name=ARIAL, size=10, bold=True)
NOTE_FONT = Font(name=ARIAL, size=9, italic=True, color="808080")
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def header_row(ws, row, c1, c2):
    for c in range(c1, c2 + 1):
        cellx = ws.cell(row=row, column=c)
        cellx.font = HEADER_FONT; cellx.fill = HEADER_FILL
        cellx.alignment = Alignment(horizontal="center", wrap_text=True)
        cellx.border = BORDER

def banner(ws, row, text, span=6):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    cellx = ws.cell(row=row, column=1, value=text)
    cellx.font = SECTION_FONT
    for c in range(1, span + 1):
        ws.cell(row=row, column=c).fill = SECTION_FILL
    cellx.alignment = Alignment(horizontal="left", indent=1, vertical="center")
    ws.row_dimensions[row].height = 20

def cell(ws, r, c, v, font=BODY_FONT, fmt=None):
    cc = ws.cell(row=r, column=c, value=v)
    cc.font = font; cc.border = BORDER
    if fmt: cc.number_format = fmt
    return cc

# ---------------------------------------------------------------------------
def parse_args():
    p = argparse.ArgumentParser()
    p.add_argument("--input-dir", required=True)
    p.add_argument("--period-start", required=True)
    p.add_argument("--period-end", required=True)
    p.add_argument("--out-dir", default="vulnpulse_reports")
    return p.parse_args()

args = parse_args()
PERIOD_START = pd.Timestamp(args.period_start)
PERIOD_END = pd.Timestamp(args.period_end)
os.makedirs(args.out_dir, exist_ok=True)

files = sorted(glob.glob(os.path.join(args.input_dir, "*.csv")) +
                glob.glob(os.path.join(args.input_dir, "*.xlsx")))
if not files:
    print(f"ERROR: no .csv/.xlsx files found in {args.input_dir}")
    sys.exit(1)

frames = []
for f in files:
    d = pd.read_csv(f) if f.lower().endswith(".csv") else pd.read_excel(f)
    missing = [c for c in REQUIRED if c not in d.columns]
    if missing:
        print(f"ERROR: '{f}' missing columns: {missing}\nFound: {list(d.columns)}")
        sys.exit(1)
    if "Tags" not in d.columns or d["Tags"].isna().all():
        d["Tags"] = os.path.splitext(os.path.basename(f))[0]
    frames.append(d)

df = pd.concat(frames, ignore_index=True)
df["Created Date"] = pd.to_datetime(df["Created Date"], errors="coerce", utc=True).dt.tz_localize(None)
df["Closed Date"] = pd.to_datetime(df["Closed Date"], errors="coerce", utc=True).dt.tz_localize(None)
n_rows = len(df)
LAST_ROW = n_rows + 1 + 500  # header at row1, data from row2, generous buffer

print(f"Loaded {n_rows} findings from {len(files)} file(s) ({df['Tags'].nunique()} tag(s), {df['Hostname'].nunique()} hosts)")
print(f"Period: {args.period_start} to {args.period_end}\n")

# ---------------------------------------------------------------------------
wb = openpyxl.Workbook()
wb.remove(wb.active)

# ===========================================================================
# RAW DATA SHEET with formula-driven helper columns
# ===========================================================================
raw = wb.create_sheet("All_Findings_Raw")
RAW_COLS = ["Hostname", "HostType", "OSVersion", "Product", "CVE ID", "Status",
            "Severity", "Created Date", "Closed Date", "Base Score", "Tags",
            "Asset Criticality", "Vulnerability ID", "Recommended Remediations",
            "Age Days", "Days to Close", "SLA Days", "Is Currently Open",
            "Is New (Period)", "Is Remediated (Period)", "Is Reopened",
            "Breached SLA (Open)", "Closed Within SLA"]
for i, h in enumerate(RAW_COLS, start=1):
    raw.cell(row=1, column=i, value=h)
header_row(raw, 1, 1, len(RAW_COLS))

for r_off, (_, row_) in enumerate(df.iterrows()):
    r = r_off + 2
    vals = [row_[c] for c in REQUIRED]
    for c, v in enumerate(vals, start=1):
        cc = raw.cell(row=r, column=c, value=v); cc.font = BODY_FONT; cc.border = BORDER
        if c in (8, 9):
            cc.number_format = "yyyy-mm-dd"
    # Helper formulas (columns O..W = 15..23)
    cell(raw, r, 15, f"='Exec Summary'!$E$2-H{r}", fmt="0")                       # Age Days
    cell(raw, r, 16, f'=IF(I{r}="","",I{r}-H{r})', fmt="0")                       # Days to Close
    cell(raw, r, 17, f"=VLOOKUP(G{r},$AA$2:$AB$5,2,FALSE)", fmt="0")              # SLA Days
    cell(raw, r, 18, f'=IF(OR(F{r}="Open",F{r}="Reopened"),1,0)')                 # Is Currently Open
    cell(raw, r, 19, f"=IF(AND(H{r}>='Exec Summary'!$E$1,H{r}<='Exec Summary'!$E$2),1,0)")  # Is New
    cell(raw, r, 20, f'=IF(AND(F{r}="Closed",I{r}<>"",I{r}>=\'Exec Summary\'!$E$1,I{r}<=\'Exec Summary\'!$E$2),1,0)')  # Is Remediated
    cell(raw, r, 21, f'=IF(F{r}="Reopened",1,0)')                                 # Is Reopened
    cell(raw, r, 22, f"=IF(AND(R{r}=1,O{r}>Q{r}),1,0)")                           # Breached SLA Open
    cell(raw, r, 23, f'=IF(AND(T{r}=1,P{r}<>"",P{r}<=Q{r}),1,0)')                 # Closed Within SLA

# Small SLA lookup table (used by VLOOKUP above)
raw["AA1"] = "Severity"; raw["AB1"] = "SLA Days"
for i, (sev, days) in enumerate(SLA_DAYS.items(), start=2):
    raw.cell(row=i, column=27, value=sev)
    raw.cell(row=i, column=28, value=days)
header_row(raw, 1, 27, 28)

widths = [16, 12, 18, 20, 14, 10, 10, 13, 13, 10, 14, 14, 16, 34, 10, 12, 10, 10, 10, 12, 10, 12, 12]
for i, w in enumerate(widths, start=1):
    raw.column_dimensions[get_column_letter(i)].width = w
raw.freeze_panes = "A2"
raw.auto_filter.ref = f"A1:{get_column_letter(len(RAW_COLS))}{n_rows+1}"

RAWCOL = {name: get_column_letter(i+1) for i, name in enumerate(RAW_COLS)}
def rc(name, end=LAST_ROW):
    return f"All_Findings_Raw!${RAWCOL[name]}$2:${RAWCOL[name]}${end}"

# ===========================================================================
# EXEC SUMMARY
# ===========================================================================
es = wb.create_sheet("Exec Summary")
wb.move_sheet("Exec Summary", offset=-(len(wb.sheetnames) - 1))

es["A1"] = "VulnPulse — Executive Summary"; es["A1"].font = TITLE_FONT
es["A2"] = f"Source: {len(files)} tag export file(s)  |  {df['Hostname'].nunique()} hosts  |  {df['Tags'].nunique()} tags"
es["A2"].font = SUBTITLE_FONT
es.merge_cells("A1:C1")

es["D1"] = "Period Start:"; es["D1"].font = BOLD_BODY
es["E1"] = PERIOD_START; es["E1"].number_format = "yyyy-mm-dd"; es["E1"].font = BOLD_BODY
es["D2"] = "Period End:"; es["D2"].font = BOLD_BODY
es["E2"] = PERIOD_END; es["E2"].number_format = "yyyy-mm-dd"; es["E2"].font = BOLD_BODY

row = 4
banner(es, row, "1. Flow — New / Remediated / Backlog (exact, from Created/Closed dates)"); row += 1
cell(es, row, 1, "Metric", HEADER_FONT); cell(es, row, 2, "Count", HEADER_FONT)
header_row(es, row, 1, 2); row += 1
flow_start = row
flow_defs = [
    ("New Findings This Period", f"=SUM({rc('Is New (Period)')})"),
    ("Remediated This Period (true closes only)", f"=SUM({rc('Is Remediated (Period)')})"),
    ("Currently Open (includes Reopened)", f"=SUM({rc('Is Currently Open')})"),
    ("Backlog (open, predates this period)", f"=SUMPRODUCT(({rc('Is Currently Open')}=1)*({rc('Created Date')}<$E$1))"),
    ("Currently Reopened (regressed findings)", f"=SUM({rc('Is Reopened')})"),
]
for label, formula in flow_defs:
    cell(es, row, 1, label); cell(es, row, 2, formula)
    row += 1
net_row = row
cell(es, row, 1, "Net Change (New − Remediated)")
cell(es, row, 2, f"=B{flow_start}-B{flow_start+1}")
row += 2

banner(es, row, "2. True Remediation Velocity & SLA Compliance"); row += 1
for i, h in enumerate(["Severity", "Avg Days to Close", "% Closed Within SLA", "Open Count", "SLA Breaches"], start=1):
    cell(es, row, i, h, HEADER_FONT)
header_row(es, row, 1, 5); row += 1
sev_start = row
for sev in ["Critical", "High", "Medium", "Low"]:
    cell(es, row, 1, sev)
    cell(es, row, 2, f'=IFERROR(AVERAGEIFS({rc("Days to Close")},{rc("Is Remediated (Period)")},1,{rc("Severity")},A{row}),0)', fmt="0.0")
    cell(es, row, 3, f'=IFERROR(SUMIFS({rc("Closed Within SLA")},{rc("Is Remediated (Period)")},1,{rc("Severity")},A{row})/SUMIFS({rc("Is Remediated (Period)")},{rc("Severity")},A{row})*100,0)', fmt="0.0")
    cell(es, row, 4, f'=SUMIFS({rc("Is Currently Open")},{rc("Severity")},A{row})')
    cell(es, row, 5, f'=SUMIFS({rc("Breached SLA (Open)")},{rc("Severity")},A{row})')
    row += 1
sev_end = row - 1
row += 1

chart1 = BarChart()
chart1.title = "New vs. Remediated vs. Reopened"
chart1.y_axis.title = "Findings"
data1 = Reference(es, min_col=2, min_row=flow_start-1, max_row=flow_start+1)
cats1 = Reference(es, min_col=1, min_row=flow_start, max_row=flow_start+1)
chart1.add_data(data1, titles_from_data=True)
chart1.set_categories(cats1)
chart1.width = 12; chart1.height = 7
es.add_chart(chart1, "G4")

chart2 = BarChart()
chart2.title = "Open Findings & SLA Breaches by Severity"
chart2.y_axis.title = "Findings"
data2 = Reference(es, min_col=4, max_col=5, min_row=sev_start-1, max_row=sev_end)
cats2 = Reference(es, min_col=1, min_row=sev_start, max_row=sev_end)
chart2.add_data(data2, titles_from_data=True)
chart2.set_categories(cats2)
chart2.width = 12; chart2.height = 7
es.add_chart(chart2, "G20")

for col, w in zip("ABCDE", [42, 20, 20, 14, 14]):
    es.column_dimensions[col].width = w

# ===========================================================================
# TAG LEADERBOARD
# ===========================================================================
lb = wb.create_sheet("Tag Leaderboard")
lb["A1"] = "Performance by Tag / Project"; lb["A1"].font = TITLE_FONT
lb.merge_cells("A1:J1")

tags = sorted(df["Tags"].dropna().unique().tolist())
host_counts = df.groupby("Tags")["Hostname"].nunique().to_dict()

cols = ["Tags", "Hosts", "New", "Remediated", "Remediation Rate %", "Currently Open",
        "Reopened", "Open Critical", "SLA Breaches", "Avg Days to Close"]
hr = 3
for i, h in enumerate(cols, start=1):
    cell(lb, hr, i, h, HEADER_FONT)
header_row(lb, hr, 1, len(cols))
r = hr + 1
first_r = r
for tag in tags:
    cell(lb, r, 1, tag)
    cell(lb, r, 2, host_counts.get(tag, 0))  # unique host count computed at generation time (see Methodology)
    cell(lb, r, 3, f'=SUMIFS({rc("Is New (Period)")},{rc("Tags")},A{r})')
    cell(lb, r, 4, f'=SUMIFS({rc("Is Remediated (Period)")},{rc("Tags")},A{r})')
    cell(lb, r, 5, f'=IFERROR(D{r}/(D{r}+F{r})*100,0)', fmt="0.0")
    cell(lb, r, 6, f'=SUMIFS({rc("Is Currently Open")},{rc("Tags")},A{r})')
    cell(lb, r, 7, f'=SUMIFS({rc("Is Reopened")},{rc("Tags")},A{r})')
    cell(lb, r, 8, f'=SUMIFS({rc("Is Currently Open")},{rc("Tags")},A{r},{rc("Severity")},"Critical")')
    cell(lb, r, 9, f'=SUMIFS({rc("Breached SLA (Open)")},{rc("Tags")},A{r})')
    cell(lb, r, 10, f'=IFERROR(AVERAGEIFS({rc("Days to Close")},{rc("Is Remediated (Period)")},1,{rc("Tags")},A{r}),0)', fmt="0.0")
    r += 1
last_r = r - 1
rule = ColorScaleRule(start_type="min", start_color="FFC7CE", mid_type="percentile", mid_value=50,
                       mid_color="FFEB9C", end_type="max", end_color="C6EFCE")
lb.conditional_formatting.add(f"E{first_r}:E{last_r}", rule)
for i, w in enumerate([16, 10, 8, 12, 16, 14, 10, 12, 12, 16], start=1):
    lb.column_dimensions[get_column_letter(i)].width = w

# ===========================================================================
# AGING ANALYSIS
# ===========================================================================
ag = wb.create_sheet("Aging Analysis")
ag["A1"] = "Aging of Currently-Open Findings"; ag["A1"].font = TITLE_FONT
ag.merge_cells("A1:E1")
hr = 3
for i, h in enumerate(["Age Bucket", "All Open", "Critical", "High"], start=1):
    cell(ag, hr, i, h, HEADER_FONT)
header_row(ag, hr, 1, 4)
buckets = [("0-30 days", 0, 30), ("31-60 days", 31, 60), ("61-90 days", 61, 90), ("90+ days", 91, None)]
r = hr + 1
for label, lo, hi in buckets:
    cell(ag, r, 1, label)
    if hi is not None:
        cell(ag, r, 2, f'=COUNTIFS({rc("Is Currently Open")},1,{rc("Age Days")},">={lo}",{rc("Age Days")},"<={hi}")')
        cell(ag, r, 3, f'=COUNTIFS({rc("Is Currently Open")},1,{rc("Age Days")},">={lo}",{rc("Age Days")},"<={hi}",{rc("Severity")},"Critical")')
        cell(ag, r, 4, f'=COUNTIFS({rc("Is Currently Open")},1,{rc("Age Days")},">={lo}",{rc("Age Days")},"<={hi}",{rc("Severity")},"High")')
    else:
        cell(ag, r, 2, f'=COUNTIFS({rc("Is Currently Open")},1,{rc("Age Days")},">={lo}")')
        cell(ag, r, 3, f'=COUNTIFS({rc("Is Currently Open")},1,{rc("Age Days")},">={lo}",{rc("Severity")},"Critical")')
        cell(ag, r, 4, f'=COUNTIFS({rc("Is Currently Open")},1,{rc("Age Days")},">={lo}",{rc("Severity")},"High")')
    r += 1
last_ag = r - 1

chart3 = BarChart()
chart3.title = "Open Findings by Age"
chart3.y_axis.title = "Findings"
data3 = Reference(ag, min_col=2, max_col=4, min_row=hr, max_row=last_ag)
cats3 = Reference(ag, min_col=1, min_row=hr+1, max_row=last_ag)
chart3.add_data(data3, titles_from_data=True)
chart3.set_categories(cats3)
chart3.width = 13; chart3.height = 7
ag.add_chart(chart3, "F3")
for i, w in enumerate([16, 12, 10, 10], start=1):
    ag.column_dimensions[get_column_letter(i)].width = w

# ===========================================================================
# TOP RISKS (top hosts, top recurring CVEs) + remediation guidance
# ===========================================================================
tr = wb.create_sheet("Top Risks")
tr["A1"] = "Risk Concentration"; tr["A1"].font = TITLE_FONT
tr.merge_cells("A1:E1")

tr["A3"] = "Top Hosts by Open Critical + High Findings"; tr["A3"].font = BOLD_BODY
hr1 = 4
for i, h in enumerate(["Hostname", "Tags", "Open Critical", "Open High", "Total"], start=1):
    cell(tr, hr1, i, h, HEADER_FONT)
header_row(tr, hr1, 1, 5)

open_df = df[df["Status"].str.lower().isin(["open", "reopened"])]
top_hosts = (open_df[open_df["Severity"].isin(["Critical", "High"])]
             .groupby("Hostname").size().sort_values(ascending=False).head(15).index.tolist())
host_tag = df.drop_duplicates("Hostname").set_index("Hostname")["Tags"].to_dict()

r = hr1 + 1
for h in top_hosts:
    cell(tr, r, 1, h)
    cell(tr, r, 2, host_tag.get(h, ""))
    cell(tr, r, 3, f'=SUMIFS({rc("Is Currently Open")},{rc("Hostname")},A{r},{rc("Severity")},"Critical")')
    cell(tr, r, 4, f'=SUMIFS({rc("Is Currently Open")},{rc("Hostname")},A{r},{rc("Severity")},"High")')
    cell(tr, r, 5, f"=C{r}+D{r}")
    r += 1

r += 2
cell(tr, r, 1, "Most Recurring CVEs (currently open, across hosts)", BOLD_BODY)
r += 1
hr2 = r
for i, h in enumerate(["CVE ID", "Severity", "Product", "# Hosts Affected"], start=1):
    cell(tr, r, i, h, HEADER_FONT)
header_row(tr, hr2, 1, 4)
r += 1
top_cves = open_df["CVE ID"].value_counts().head(10).index.tolist()
cve_meta = df.drop_duplicates("CVE ID").set_index("CVE ID")
for c in top_cves:
    cell(tr, r, 1, c)
    cell(tr, r, 2, cve_meta.loc[c, "Severity"])
    cell(tr, r, 3, cve_meta.loc[c, "Product"])
    cell(tr, r, 4, f'=SUMIFS({rc("Is Currently Open")},{rc("CVE ID")},A{r})')
    r += 1

for i, w in enumerate([18, 16, 22, 16, 12], start=1):
    tr.column_dimensions[get_column_letter(i)].width = w

# ===========================================================================
# METHODOLOGY
# ===========================================================================
mn = wb.create_sheet("Methodology & Assumptions")
mn["A1"] = "Methodology & Assumptions"; mn["A1"].font = TITLE_FONT
notes = [
    f"Data source: {len(files)} detailed CrowdStrike finding-level export file(s), one per tag/project, combined on the 'All_Findings_Raw' sheet.",
    f"Reporting period: {args.period_start} to {args.period_end} — set in cells E1/E2 on the Exec Summary sheet; every formula in this workbook references those two cells, so changing them recalculates the whole report.",
    "New = Created Date within the period. Remediated = Status is Closed (not Reopened) AND Closed Date within the period — both exact, from CrowdStrike's own timestamps.",
    "Reopened findings count as currently OPEN (active risk), not resolved. A finding closed then reopened does NOT count as a remediation win, even if Closed Date falls in this period.",
    f"SLA targets: Critical={SLA_DAYS['Critical']}d, High={SLA_DAYS['High']}d, Medium={SLA_DAYS['Medium']}d, Low={SLA_DAYS['Low']}d — see the small lookup table on All_Findings_Raw (columns AA:AB). Edit there to change SLA policy; formulas update automatically.",
    "All Exec Summary, Leaderboard, Aging, and Top Risks figures are LIVE FORMULAS (COUNTIFS/SUMIFS/AVERAGEIFS) referencing All_Findings_Raw — edit or add rows there and the rest of the workbook recalculates.",
    "Exception: 'Hosts' (unique host count) per tag on the Leaderboard is computed once at report-generation time, not as a live formula — true unique-count formulas are unreliable across Excel/LibreOffice versions at this scale. Regenerate the report to refresh this figure after adding new hosts.",
    "This report expects each input file to include BOTH open and closed findings (not just currently-open) — if a file only has open findings, New/Remediated totals will be understated.",
]
r = 3
for n in notes:
    cell(mn, r, 1, f"• {n}", NOTE_FONT)
    mn.cell(row=r, column=1).alignment = Alignment(wrap_text=True, vertical="top")
    mn.row_dimensions[r].height = 32
    r += 1
mn.column_dimensions["A"].width = 130

out_path = os.path.join(args.out_dir, f"VulnPulse_Report_{args.period_start}_to_{args.period_end}.xlsx")
wb.save(out_path)
print(f"✅ VulnPulse report: {out_path}")
