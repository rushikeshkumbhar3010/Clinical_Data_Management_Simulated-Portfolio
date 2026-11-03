"""Build the four Excel deliverables. Run after 01 and 02:  python scripts/03_build_excel.py"""
import pandas as pd, numpy as np, os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.chart import BarChart, Reference

ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")
disc = pd.read_csv(f"{ROOT}/scripts/discrepancies_detected.csv")
rng = np.random.default_rng(77)

NAVY, TEAL, LIGHT, GREY = "1F3864", "0E7C86", "EAF2F8", "F2F2F2"
F = lambda **k: Font(name="Arial", size=k.pop("size", 10), **k)
thin = Side(style="thin", color="BFBFBF"); BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
DATEFMT = "DD-MMM-YYYY"

def title(ws, text, sub=None, width=8):
    ws["A1"] = text; ws["A1"].font = F(size=16, bold=True, color=NAVY)
    if sub: ws["A2"] = sub; ws["A2"].font = F(size=10, italic=True, color="595959")
    ws.sheet_view.showGridLines = False

def header(ws, row, cols, fill=NAVY):
    for i, c in enumerate(cols, 1):
        x = ws.cell(row=row, column=i, value=c)
        x.font = F(bold=True, color="FFFFFF"); x.fill = PatternFill("solid", fgColor=fill)
        x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); x.border = BORDER
    ws.row_dimensions[row].height = 28

def body(ws, r0, rows, wrap_cols=(), zebra=True):
    for k, row in enumerate(rows):
        for j, v in enumerate(row, 1):
            c = ws.cell(row=r0 + k, column=j, value=v)
            c.font = F(); c.border = BORDER
            c.alignment = Alignment(vertical="top", wrap_text=(j in wrap_cols))
            if zebra and k % 2 == 1: c.fill = PatternFill("solid", fgColor=GREY)

def widths(ws, ws_widths):
    for i, w in enumerate(ws_widths, 1): ws.column_dimensions[L(i)].width = w

# =====================================================================
# 1. Study visit matrix + eCRF specs
# =====================================================================
wb = Workbook(); ws = wb.active; ws.title = "Visit_Schedule"
title(ws, "Study Visit Schedule Matrix", "Protocol NSCLC-P2-001 | Phase II Open-Label Study in Advanced NSCLC | N = 350")
header(ws, 4, ["Visit #", "Visit Name", "Target Day / Window", "Window Start (Day)", "Window End (Day)", "Forms / Domains Required"])
visits = [
    (1, "Screening", "Day -14 to Day -1", -14, -1, "Demographics (DM), Inclusion/Exclusion (IE), Medical History (MH), Vital Signs (VS), Baseline Labs (LB)"),
    (2, "Baseline / Cycle 1", "Day 1 (±1 day)", 0, 2, "Vital Signs (VS), Study Drug Administration (EC), Concomitant Medications (CM), Adverse Events (AE)"),
    (3, "Treatment Evaluation", "Day 28 (±3 days)", 25, 31, "Vital Signs (VS), Lab Tests (LB), Adverse Events (AE), Concomitant Meds (CM)"),
    (4, "End of Study / Safety Follow-Up", "30 days post-last dose", 30, 30, "Physical Exam (PE), Vital Signs (VS), AE Resolution (AE), Study Exit Form (DS)"),
]
body(ws, 5, visits, wrap_cols=(2, 3, 6))
ws["A10"] = "Form-by-visit matrix  (X = form collected)"; ws["A10"].font = F(bold=True, size=12, color=TEAL)
header(ws, 11, ["Form", "Domain", "Screening", "Baseline / C1D1", "Day 28 (V3)", "EOS / Follow-up"], fill=TEAL)
forms = [("Demographics", "DM", "X", "", "", ""), ("Inclusion/Exclusion", "IE", "X", "", "", ""), ("Medical History", "MH", "X", "", "", ""),
         ("Vital Signs", "VS", "X", "X", "X", "X"), ("Laboratory", "LB", "X", "", "X", ""), ("Study Drug Administration", "EC", "", "X", "", ""),
         ("Concomitant Medications", "CM", "", "X", "X", ""), ("Adverse Events", "AE", "", "X", "X", "X"),
         ("Serious AE Report (event-driven)", "SAE", "", "X", "X", "X"), ("Physical Exam", "PE", "", "", "", "X"), ("Study Exit", "DS", "", "", "", "X")]
body(ws, 12, forms)
for r in range(12, 12 + len(forms)):
    for c in range(3, 7):
        x = ws.cell(row=r, column=c); x.alignment = Alignment(horizontal="center")
        if x.value == "X": x.fill = PatternFill("solid", fgColor="C6E0B4"); x.font = F(bold=True, color="375623")
ws["A25"] = "Total forms per visit"; ws["A25"].font = F(bold=True)
for c in range(3, 7):
    ws.cell(row=25, column=c, value=f'=COUNTIF({L(c)}12:{L(c)}22,"X")').font = F(bold=True)
    ws.cell(row=25, column=c).alignment = Alignment(horizontal="center")
widths(ws, [30, 30, 22, 20, 20, 80]); ws.freeze_panes = "A5"

def ecrf(name, ttl, rows):
    s = wb.create_sheet(name); title(s, ttl, "Modular eCRF specification | CDISC-aligned variable names")
    header(s, 4, ["Variable", "Label", "Data Type", "Format / Units", "Controlled Terms / Codelist", "Required", "Derivation / Linked Edit Check"])
    body(s, 5, rows, wrap_cols=(2, 5, 7)); widths(s, [14, 34, 14, 20, 52, 11, 44]); s.freeze_panes = "A5"

ecrf("eCRF_DM", "Demographics (DM) Form", [
    ("SUBJID", "Subject ID", "Text", "SITE-XXX (e.g., S01-001)", "", "Y", "Unique; site prefix must match Site_ID"),
    ("BRTHDTC", "Date of Birth", "Date", "YYYY-MM-DD", "", "Y", ""),
    ("AGE", "Age in Years", "Numeric", "Years", "", "Y", "Calculated: floor((RFICDTC - BRTHDTC)/365.25); CK_DM_01 (18-75)"),
    ("SEX", "Sex", "Radio", "", "Male | Female | Other", "Y", ""),
    ("RACE", "Race", "Checkbox", "", "Asian | White | Black/African American | Native | Other", "Y", "CK_DM_02 (not blank)"),
    ("RFICDTC", "Informed Consent Date", "Date", "YYYY-MM-DD", "", "Y", "Anchor date for AE/VS chronology checks")])
ecrf("eCRF_VS", "Vital Signs (VS) Form", [
    ("VSDAT", "Assessment Date", "Date", "YYYY-MM-DD", "", "Y", "CK_VS_04: must be on/after consent date"),
    ("SYSBP", "Systolic Blood Pressure", "Integer", "mmHg", "", "Y", "CK_VS_01: SYSBP > DIABP"),
    ("DIABP", "Diastolic Blood Pressure", "Integer", "mmHg", "", "Y", "CK_VS_01"),
    ("PULSE", "Pulse Rate", "Integer", "beats/min", "", "Y", "CK_VS_02: 40-140"),
    ("TEMP", "Body Temperature", "Decimal", "°C", "", "Y", "CK_VS_05: 34-41 °C"),
    ("WEIGHT", "Weight", "Decimal", "kg", "", "Y", "Visit completeness: CK_VS_03")])
ecrf("eCRF_AE", "Adverse Events (AE) Form", [
    ("AETERM", "Adverse Event Verbatim Term", "Free text", "", "", "Y", "Coded to MedDRA PT (AEDECOD)"),
    ("AESTDAT", "Start Date/Time", "Date", "YYYY-MM-DD", "", "Y", "CK_AE_03: not before consent"),
    ("AEENDAT", "End Date/Time", "Date", "YYYY-MM-DD", "Or tick 'Ongoing'", "Cond.", "CK_AE_01: not before start; CK_AE_04"),
    ("AESEV", "Severity", "Radio", "", "Mild | Moderate | Severe", "Y", ""),
    ("AESER", "Serious Adverse Event?", "Radio", "", "Yes | No", "Y", "CK_AE_02: Yes requires SAE form"),
    ("AEREL", "Causality to Investigational Product", "Radio", "", "Unrelated | Unlikely | Possible | Probable | Definite", "Y", ""),
    ("AEOUT", "Outcome", "Radio", "", "Recovered/Resolved | Recovering/Resolving | Not Recovered | Fatal | Unknown", "Y", "CK_AE_04")])
ecrf("eCRF_LB", "Laboratory (LB) Form", [
    ("LBDAT", "Collection Date", "Date", "YYYY-MM-DD", "", "Y", ""),
    ("HGB", "Hemoglobin", "Decimal", "g/dL", "", "Y", "CK_LB_01: 5-20 g/dL"),
    ("WBC", "White Blood Cells", "Decimal", "10^9/L", "", "Y", ""),
    ("PLT", "Platelets", "Integer", "10^9/L", "", "Y", ""),
    ("ALT", "Alanine Aminotransferase", "Integer", "U/L", "", "Y", ""),
    ("CREAT", "Creatinine", "Decimal", "mg/dL", "", "Y", "")])
wb.save(f"{ROOT}/Study_Visit_Matrix_and_eCRF_Specifications.xlsx")

# =====================================================================
# 2. Data Validation Plan & Edit checks
# =====================================================================
wb = Workbook(); ws = wb.active; ws.title = "Edit_Check_Specifications"
title(ws, "Edit Check Specifications (Data Validation Plan)", "Logic rules executed against Demographics, Vital Signs, Laboratory and Adverse Event data")
header(ws, 4, ["Check ID", "Form Name", "Variable(s)", "Check Type", "Logic Condition / Rule", "Discrepancy Query Text", "Execution", "Priority"])
checks = [
    ("CK_DM_01", "Demographics", "AGE", "Univariate / Boundary", "AGE < 18 OR AGE > 75", "Subject age is [AGE]. Protocol requires enrolled subjects to be aged 18–75 years. Please verify date of birth or record protocol deviation.", "System", "High", 12),
    ("CK_DM_02", "Demographics", "RACE", "Missing data", "RACE == NULL", "Race is blank. Please provide the subject's race per source or confirm it was not collected.", "Manual", "Low", 8),
    ("CK_VS_01", "Vital Signs", "SYSBP, DIABP", "Multivariate / Range", "SYSBP <= DIABP", "Systolic BP ([SYSBP]) cannot be less than or equal to Diastolic BP ([DIABP]). Please confirm and correct values.", "System", "High", 15),
    ("CK_VS_02", "Vital Signs", "PULSE", "Univariate / Range", "PULSE < 40 OR PULSE > 140", "Pulse rate ([PULSE] bpm) is outside the clinically plausible physiological range (40–140 bpm). Please verify source record.", "System", "Medium", 20),
    ("CK_VS_03", "Vital Signs", "VS (all fields)", "Missing visit data", "Visit 3 record exists AND all VS fields == NULL", "Vital signs are missing for Visit 3 (Day 28). Please provide the assessment or confirm that it was not done.", "System", "High", 25),
    ("CK_VS_04", "Vital Signs", "VSDAT, RFICDTC", "Cross-form / Date", "VSDAT < RFICDTC", "Vital signs date precedes the informed consent date. Please verify and correct the date or provide clarification.", "System", "High", 10),
    ("CK_VS_05", "Vital Signs", "TEMP", "Univariate / Range", "TEMP < 34 OR TEMP > 41 (°C)", "Body temperature ([TEMP] °C) is outside the plausible range. Please verify value and unit against source.", "System", "Medium", 10),
    ("CK_LB_01", "Laboratory", "HGB", "Univariate / Range", "HGB < 5 OR HGB > 20 (g/dL)", "Hemoglobin result of [HGB] g/dL is outside the physiologically plausible range. Please verify the value and unit against the lab report.", "System", "High", 15),
    ("CK_AE_01", "Adverse Events", "AESTDAT, AEENDAT", "Logic / Sequence", "AEENDAT < AESTDAT", "Adverse event stop date cannot precede start date. Please confirm and resolve date chronology.", "System", "High", 15),
    ("CK_AE_02", "Adverse Events", "AESER, SAE_FORM", "Cross-form Logic", "AESER == 'Yes' AND SAE_FORM == NULL", "Adverse Event is marked Serious (SAE), but the dedicated SAE Reporting Form is missing. Please complete the SAE notification form.", "System", "Critical", 15),
    ("CK_AE_03", "Adverse Events", "AESTDAT, RFICDTC", "Cross-form / Date", "AESTDAT < RFICDTC", "AE start date precedes the informed consent date. Please verify and correct the date or provide clarification.", "System", "High", 10),
    ("CK_AE_04", "Adverse Events", "AEENDAT, AEOUT, AEONGO", "Missing data / Conditional", "AEONGO != 'Y' AND (AEENDAT == NULL OR AEOUT == NULL)", "AE has no stop date and/or outcome and is not marked ongoing. Please provide the missing information.", "Manual", "Medium", 20),
    ("CK_AE_05", "Adverse Events", "AETERM, AESTDAT", "Duplicate record", "Same SUBJID + AETERM + AESTDAT occurs > 1 time", "AE appears more than once for this subject. Please confirm whether these are separate events or a duplicate entry.", "Manual", "Medium", 8),
]
body(ws, 5, [c[:8] for c in checks], wrap_cols=(3, 4, 5, 6))
widths(ws, [11, 15, 20, 22, 36, 70, 11, 10]); ws.freeze_panes = "A5"
for r in range(5, 5 + len(checks)):
    p = ws.cell(row=r, column=8); col = {"Critical": "F8CBAD", "High": "FFE699", "Medium": "DDEBF7", "Low": "E2EFDA"}[p.value]
    p.fill = PatternFill("solid", fgColor=col); p.alignment = Alignment(horizontal="center", vertical="top")

# discrepancy listing
dl = wb.create_sheet("Discrepancy_Listing"); title(dl, "Discrepancy Listing (output of validation run)", "Generated by scripts/02_run_validation.py")
header(dl, 4, ["Check_ID", "Site_ID", "Subject_ID", "Form_Name", "Field_Name", "Record_Ref", "Check_Type", "Query_Text"])
dsort = disc.sort_values(["Check_ID", "Subject_ID"]).fillna("")
body(dl, 5, dsort[["Check_ID", "Site_ID", "Subject_ID", "Form_Name", "Field_Name", "Record_Ref", "Check_Type", "Query_Text"]].values.tolist(), wrap_cols=(8,))
widths(dl, [11, 8, 12, 15, 26, 18, 11, 100]); dl.freeze_panes = "A5"; dl.auto_filter.ref = f"A4:H{4 + len(dsort)}"
last = 4 + len(dsort)

# results
rs = wb.create_sheet("Validation_Results", 1); title(rs, "Validation Run Summary", "Counts calculated live from Discrepancy_Listing; seeded error target in blue (from dataset generator)")
header(rs, 4, ["Check ID", "Form", "Check Type", "Priority", "Records Flagged", "Seeded Target", "Match?"])
for i, c in enumerate(checks):
    r = 5 + i
    vals = [c[0], c[1], c[3], c[7], f'=COUNTIF(Discrepancy_Listing!$A$5:$A${last},A{r})', c[8], f'=IF(E{r}=F{r},"✔","CHECK")']
    for j, v in enumerate(vals, 1):
        x = rs.cell(row=r, column=j, value=v); x.font = F(color="0000FF" if j == 6 else "000000"); x.border = BORDER
        if j >= 5: x.alignment = Alignment(horizontal="center")
t = 5 + len(checks)
rs.cell(row=t, column=1, value="TOTAL").font = F(bold=True)
for col in (5, 6):
    x = rs.cell(row=t, column=col, value=f"=SUM({L(col)}5:{L(col)}{t-1})"); x.font = F(bold=True); x.alignment = Alignment(horizontal="center")
rs.cell(row=t, column=7, value=f'=IF(E{t}=F{t},"✔","CHECK")').alignment = Alignment(horizontal="center")
rs.cell(row=t + 2, column=1, value="Detection rate (flagged / seeded)").font = F(bold=True)
rs.cell(row=t + 2, column=5, value=f"=E{t}/F{t}").number_format = "0.0%"
widths(rs, [12, 16, 26, 11, 16, 14, 10])
rs.conditional_formatting.add(f"G5:G{t}", CellIsRule(operator="equal", formula=['"✔"'], fill=PatternFill("solid", bgColor="C6EFCE")))
ch = BarChart(); ch.type = "bar"; ch.title = "Discrepancies by Check"; ch.style = 10
ch.add_data(Reference(rs, min_col=5, min_row=4, max_row=t - 1), titles_from_data=True); ch.set_categories(Reference(rs, min_col=1, min_row=5, max_row=t - 1))
ch.height, ch.width, ch.legend = 9, 16, None; rs.add_chart(ch, "I4")
wb.move_sheet("Validation_Results", offset=0)
wb.save(f"{ROOT}/Data_Validation_Plan_and_Edit_Checks.xlsx")

# =====================================================================
# 3. Master query tracker
# =====================================================================
CUT = pd.Timestamp("2026-08-31")
q = disc.sample(frac=1, random_state=5).reset_index(drop=True)
n = len(q)
issued = np.sort(pd.to_datetime("2026-07-06") + pd.to_timedelta(rng.integers(0, 52, n), unit="D"))
q["Date_Issued"] = issued
q["SLA"] = np.where(q.Check_ID == "CK_AE_02", 1, 5)
status = np.array(["Closed"] * n, dtype=object)
status[-19:] = "Open"
status[rng.choice(np.arange(n - 55, n - 19), 20, replace=False)] = "Answered"
status[rng.choice(np.arange(60, n - 55), 13, replace=False)] = "Re-issued"
q["Status"] = status
RESP = {
 "CK_DM_01": ["Age verified against source; date of birth transcription error corrected.", "Value confirmed correct per source. Protocol deviation recorded and PI notified."],
 "CK_DM_02": ["Race collected per source and entered.", "Race not collected at site per local regulation; marked as not collected."],
 "CK_VS_01": ["Values were transposed at entry; corrected per source document.", "Source verified; SBP/DBP entered in the wrong fields and corrected."],
 "CK_VS_02": ["Entry error; corrected per source.", "Value confirmed per source; investigator assessed as not clinically significant."],
 "CK_VS_03": ["Assessment performed; values entered from source.", "Visit not performed (subject missed visit); marked as not done."],
 "CK_VS_04": ["Date entered incorrectly; corrected per source.", "Year typed incorrectly; corrected."],
 "CK_VS_05": ["Temperature was recorded in °F; converted and corrected to °C.", "Unit error confirmed; value corrected per source."],
 "CK_LB_01": ["Value transcribed from the lab report with wrong unit; corrected.", "Decimal entry error; corrected per central lab report."],
 "CK_AE_01": ["Dates were swapped at entry; corrected per source.", "Stop date typed incorrectly; corrected."],
 "CK_AE_02": ["SAE form completed and submitted to safety.", "SAE form was sent separately; now uploaded to the system."],
 "CK_AE_03": ["Event onset confirmed after consent; start date corrected.", "Pre-consent event reclassified to medical history."],
 "CK_AE_04": ["Stop date and outcome entered per source.", "AE is ongoing; flag ticked and outcome updated."],
 "CK_AE_05": ["Duplicate entry; second record deleted.", "Both events are distinct occurrences; confirmed separate."],
}
bd = lambda d, k: pd.Timestamp(np.busday_offset(d.date(), k, roll="forward"))
resp, ans, aud, ini = [], [], [], []
for i, r in q.iterrows():
    d = r.Date_Issued
    if r.Status == "Open":
        resp.append(""); ans.append(pd.NaT); aud.append(d + pd.Timedelta(hours=9, minutes=int(rng.integers(0, 600)))); ini.append(""); continue
    k = int(rng.integers(0, 2)) if r.SLA == 1 else int(rng.choice([1, 2, 2, 3, 3, 4, 5, 6, 8], p=[.08, .17, .17, .17, .15, .1, .08, .05, .03]))
    if r.SLA == 1 and rng.random() < .1: k = 2
    a = min(bd(d, k), CUT)
    ans.append(a); text = RESP[r.Check_ID][int(rng.integers(0, 2))]
    if r.Status == "Re-issued": text = "Please see source document; investigator reviewing. " + ("(Response incomplete – query re-issued.)")
    resp.append(text)
    if r.Status == "Closed":
        c = min(bd(a, int(rng.integers(0, 3))), CUT); aud.append(c + pd.Timedelta(hours=int(rng.integers(9, 18)), minutes=int(rng.integers(0, 60)))); ini.append("RK")
    else:
        aud.append(a + pd.Timedelta(hours=int(rng.integers(9, 18)), minutes=int(rng.integers(0, 60)))); ini.append("")
q["Site_Response"], q["Date_Answered"], q["Audit"], q["Init"] = resp, ans, aud, ini
q["Query_ID"] = [f"QRY-{i+1:03d}" for i in range(n)]
q["Mode"] = q.Check_Type.map({"System": "System automated edit check", "Manual": "Manual CDA review"})

wb = Workbook(); sm = wb.active; sm.title = "Summary"; ql = wb.create_sheet("Query_Log")
title(ql, "Master Discrepancy & Query Tracker", "Query lifecycle: Open → Answered → (Re-issued) → Closed  |  Issued by Rushikesh Kumbhar (CDA)")
cols = ["Query_ID", "Check_ID", "Site_ID", "Subject_ID", "Form_Name", "Field_Name", "Record_Ref", "Check_Type", "Query_Text", "Date_Issued", "Issued_By",
        "Site_Response", "Date_Answered", "Query_Status", "Audit_Trail_Timestamp", "Closure_Initials", "SLA_Target_Days", "TAT_Days", "SLA_Met", "Days_Open_At_Cut"]
header(ql, 4, cols)
for i, r in q.iterrows():
    x = 5 + i
    vals = [r.Query_ID, r.Check_ID, r.Site_ID, r.Subject_ID, r.Form_Name, r.Field_Name, r.Record_Ref, r.Mode, r.Query_Text, r.Date_Issued.to_pydatetime(),
            "Rushikesh Kumbhar (CDA)", r.Site_Response, None if pd.isna(r.Date_Answered) else r.Date_Answered.to_pydatetime(), r.Status,
            r.Audit.to_pydatetime(), r.Init, int(r.SLA),
            f'=IF(M{x}="","",NETWORKDAYS(J{x},M{x})-1)', f'=IF(M{x}="","",IF(R{x}<=Q{x},"Yes","No"))',
            f'=IF(N{x}="Closed","",Summary!$C$4-J{x})']
    for j, v in enumerate(vals, 1):
        c = ql.cell(row=x, column=j, value=v); c.font = F(); c.border = BORDER; c.alignment = Alignment(vertical="top", wrap_text=j in (9, 12))
        if j in (10, 13): c.number_format = DATEFMT
        if j == 15: c.number_format = "DD-MMM-YYYY HH:MM"
        if i % 2: c.fill = PatternFill("solid", fgColor=GREY)
end = 4 + n
widths(ql, [10, 10, 8, 11, 14, 22, 15, 24, 78, 13, 22, 44, 13, 12, 20, 10, 9, 9, 9, 11]); ql.freeze_panes = "E5"; ql.auto_filter.ref = f"A4:T{end}"
dv = DataValidation(type="list", formula1='"Open,Answered,Re-issued,Closed"', allow_blank=False); ql.add_data_validation(dv); dv.add(f"N5:N{end}")
for s, col in {"Open": "F8CBAD", "Answered": "FFE699", "Re-issued": "D9C3E8", "Closed": "C6E0B4"}.items():
    ql.conditional_formatting.add(f"N5:N{end}", CellIsRule(operator="equal", formula=[f'"{s}"'], fill=PatternFill("solid", bgColor=col)))
ql.conditional_formatting.add(f"S5:S{end}", CellIsRule(operator="equal", formula=['"No"'], font=Font(name="Arial", color="C00000", bold=True)))

# summary
title(sm, "Query Management Summary", "All figures are live formulas over Query_Log")
sm["B4"] = "Database cut date"; sm["B4"].font = F(bold=True)
sm["C4"] = CUT.to_pydatetime(); sm["C4"].number_format = DATEFMT; sm["C4"].font = F(color="0000FF"); sm["D4"] = "← assumption: simulated data cut-off (edit to re-age open queries)"; sm["D4"].font = F(italic=True, color="7F7F7F")
header(sm, 6, ["", "Status", "Queries", "% of Total"]); sm.cell(row=6, column=1).fill = PatternFill(None)
rngQ = f"Query_Log!$N$5:$N${end}"
for i, s in enumerate(["Open", "Answered", "Re-issued", "Closed"]):
    r = 7 + i
    sm.cell(row=r, column=2, value=s).font = F(); sm.cell(row=r, column=3, value=f'=COUNTIF({rngQ},B{r})').font = F()
    sm.cell(row=r, column=4, value=f"=C{r}/$C$11").number_format = "0.0%"
sm["B11"] = "Total"; sm["C11"] = "=SUM(C7:C10)"; sm["B11"].font = sm["C11"].font = F(bold=True)
sm["B13"] = "Closure rate"; sm["C13"] = "=C10/C11"; sm["C13"].number_format = "0.0%"
sm["B14"] = "SLA met (answered queries)"; sm["C14"] = f'=COUNTIF(Query_Log!$S$5:$S${end},"Yes")/(COUNTIF(Query_Log!$S$5:$S${end},"Yes")+COUNTIF(Query_Log!$S$5:$S${end},"No"))'; sm["C14"].number_format = "0.0%"
sm["B15"] = "Avg turnaround (business days)"; sm["C15"] = f"=AVERAGE(Query_Log!$R$5:$R${end})"; sm["C15"].number_format = "0.0"
sm["B16"] = "Oldest open query (days)"; sm["C16"] = f"=MAX(Query_Log!$T$5:$T${end})"
for r in (13, 14, 15, 16): sm[f"B{r}"].font = F(bold=True); sm[f"C{r}"].font = F()
# by site
header(sm, 18, ["", "Site", "Total", "Open", "Closed"]); sm.cell(row=18, column=1).fill = PatternFill(None)
for i, s in enumerate([f"S{k:02d}" for k in range(1, 11)]):
    r = 19 + i; sm.cell(row=r, column=2, value=s).font = F()
    sm.cell(row=r, column=3, value=f"=COUNTIF(Query_Log!$C$5:$C${end},B{r})"); sm.cell(row=r, column=4, value=f'=COUNTIFS(Query_Log!$C$5:$C${end},B{r},{rngQ},"Open")')
    sm.cell(row=r, column=5, value=f'=COUNTIFS(Query_Log!$C$5:$C${end},B{r},{rngQ},"Closed")')
widths(sm, [3, 30, 14, 14, 12]); sm.column_dimensions["D"].width = 14
c1 = BarChart(); c1.type = "col"; c1.title = "Queries by status"; c1.add_data(Reference(sm, min_col=3, min_row=6, max_row=10), titles_from_data=True)
c1.set_categories(Reference(sm, min_col=2, min_row=7, max_row=10)); c1.legend = None; c1.height, c1.width = 7.5, 12; sm.add_chart(c1, "G4")
c2 = BarChart(); c2.type = "col"; c2.title = "Queries by site"; c2.add_data(Reference(sm, min_col=3, min_row=18, max_row=28), titles_from_data=True)
c2.set_categories(Reference(sm, min_col=2, min_row=19, max_row=28)); c2.legend = None; c2.height, c2.width = 7.5, 12; sm.add_chart(c2, "G20")

# GCP guidance sheet
gp = wb.create_sheet("GCP_Query_Writing"); title(gp, "GCP-Compliant Query Writing", "Queries must be neutral, specific and must not lead the investigator")
header(gp, 4, ["Type", "Example", "Why"])
body(gp, 5, [("Non-compliant (leading)", "Did the patient have headache on 12-Jun-2026? Please change date to 12-Jun-2026.", "Suggests the answer / dictates the correction — violates GCP data-integrity principles."),
             ("Compliant", "Adverse event start date is recorded as 12-May-2026, which precedes the Informed Consent date of 01-Jun-2026. Please verify and correct the date or provide clarification.", "States the discrepancy factually; site decides the resolution from source.")], wrap_cols=(2, 3), zebra=False)
widths(gp, [26, 80, 70]); gp.row_dimensions[5].height = 45; gp.row_dimensions[6].height = 60
wb.save(f"{ROOT}/Master_Discrepancy_and_Query_Tracker.xlsx")
q.drop(columns=["Check_Type"]).rename(columns={"Mode": "Check_Type"}).to_csv(f"{ROOT}/scripts/query_log_export.csv", index=False)

# =====================================================================
# 4. SDTM mapping
# =====================================================================
wb = Workbook(); ws = wb.active; ws.title = "Overview"
title(ws, "CDISC SDTM Mapping Specification", "STUDYID = NSCLC-P2-001  |  SDTM-IG 3.4  |  Domains: DM, VS, AE, LB")
ws["A4"] = "Conventions"; ws["A4"].font = F(bold=True, size=12, color=TEAL)
for i, t in enumerate(["USUBJID = STUDYID || '-' || raw Subject_ID (e.g., NSCLC-P2-001-S01-001)", "All dates → ISO 8601 (YYYY-MM-DD)",
                        "Vital signs and labs are transposed from wide (raw) to tall (SDTM) — one row per test per visit", "Controlled terminology per NCI CDISC CT; AEDECOD = MedDRA Preferred Term",
                        "Rows failing edit checks are queried and resolved BEFORE mapping (see Master_Discrepancy_and_Query_Tracker.xlsx)"]):
    ws.cell(row=5 + i, column=1, value="• " + t).font = F()
ws.column_dimensions["A"].width = 120
def dom(name, ttl, rows):
    s = wb.create_sheet(name); title(s, ttl, "Raw eCRF variable → SDTM variable")
    header(s, 4, ["Source Dataset", "Raw Variable", "SDTM Variable", "SDTM Label", "Type", "Controlled Terminology / Format", "Mapping Rule / Derivation"])
    body(s, 5, rows, wrap_cols=(4, 6, 7)); widths(s, [14, 22, 14, 34, 8, 36, 64]); s.freeze_panes = "A5"
dom("DM", "Demographics (DM)", [
    ("DM.csv", "—", "STUDYID", "Study Identifier", "Char", "", "Constant 'NSCLC-P2-001'"),
    ("DM.csv", "—", "DOMAIN", "Domain Abbreviation", "Char", "DM", "Constant"),
    ("DM.csv", "Subject_ID", "USUBJID", "Unique Subject Identifier", "Char", "STUDYID-SITE-NNN", "STUDYID || '-' || Subject_ID  → NSCLC-P2-001-S01-001"),
    ("DM.csv", "Subject_ID", "SUBJID", "Subject Identifier for the Study", "Char", "", "Direct copy"),
    ("DM.csv", "Site_ID", "SITEID", "Study Site Identifier", "Char", "", "Direct copy"),
    ("DM.csv", "Age", "AGE", "Age", "Num", "", "Direct copy (after CK_DM_01 resolved)"),
    ("DM.csv", "—", "AGEU", "Age Units", "Char", "YEARS", "Constant"),
    ("DM.csv", "Sex", "SEX", "Sex", "Char", "M | F | U", "Male→M; Female→F; Other/blank→U"),
    ("DM.csv", "Race", "RACE", "Race", "Char", "ASIAN | WHITE | BLACK OR AFRICAN AMERICAN | AMERICAN INDIAN OR ALASKA NATIVE | OTHER", "Asian→ASIAN; White→WHITE; Black/African American→BLACK OR AFRICAN AMERICAN; Native→AMERICAN INDIAN OR ALASKA NATIVE; blank→ (null)"),
    ("DM.csv", "Informed_Consent_Date", "RFICDTC", "Date/Time of Informed Consent", "Char", "ISO 8601 YYYY-MM-DD", "Format date as ISO 8601"),
    ("VS.csv", "VSDAT (Baseline/C1D1)", "RFSTDTC", "Subject Reference Start Date/Time", "Char", "ISO 8601", "Date of first dose = Baseline/C1D1 visit date")])
dom("VS", "Vital Signs (VS)", [
    ("VS.csv", "—", "STUDYID / DOMAIN", "Study ID / Domain", "Char", "VS", "Constants"),
    ("VS.csv", "Subject_ID", "USUBJID", "Unique Subject Identifier", "Char", "", "STUDYID || '-' || Subject_ID"),
    ("VS.csv", "(row sequence)", "VSSEQ", "Sequence Number", "Num", "", "Increment per subject, ordered by VSDTC then VSTESTCD"),
    ("VS.csv", "Systolic_BP", "VSTESTCD / VSTEST", "Vital Signs Test Short Name / Name", "Char", "SYSBP / Systolic Blood Pressure", "One record per test: VSTESTCD='SYSBP'"),
    ("VS.csv", "Systolic_BP", "VSORRES / VSORRESU", "Result as Collected / Unit", "Char", "mmHg", "VSORRES = value; VSORRESU = 'mmHg'"),
    ("VS.csv", "Diastolic_BP", "VSTESTCD / VSORRES", "Diastolic BP", "Char", "DIABP; mmHg", "VSTESTCD='DIABP'; VSORRES = value; VSORRESU = 'mmHg'"),
    ("VS.csv", "Pulse", "VSTESTCD / VSORRES", "Pulse Rate", "Char", "PULSE; beats/min", "VSTESTCD='PULSE'; VSORRES = value; VSORRESU = 'beats/min'"),
    ("VS.csv", "Temp_C", "VSTESTCD / VSORRES", "Temperature", "Char", "TEMP; C", "VSTESTCD='TEMP'; VSORRESU = 'C'"),
    ("VS.csv", "Weight_kg", "VSTESTCD / VSORRES", "Weight", "Char", "WEIGHT; kg", "VSTESTCD='WEIGHT'; VSORRESU = 'kg'"),
    ("VS.csv", "Visit", "VISIT / VISITNUM", "Visit Name / Number", "Char/Num", "SCREENING=1; BASELINE=2; DAY 28=3; END OF STUDY=4", "Map visit name to number"),
    ("VS.csv", "VSDAT", "VSDTC", "Date/Time of Measurements", "Char", "ISO 8601", "Format as YYYY-MM-DD"),
    ("VS.csv", "—", "VSSTRESC / VSSTRESN / VSSTRESU", "Standardized result", "Char/Num", "", "Copy of original (units already standard)"),
    ("VS.csv", "—", "VSBLFL", "Baseline Flag", "Char", "Y", "Y on last non-missing pre-dose result (Screening/Baseline)")])
dom("AE", "Adverse Events (AE)", [
    ("AE.csv", "—", "STUDYID / DOMAIN", "Study ID / Domain", "Char", "AE", "Constants"),
    ("AE.csv", "Subject_ID", "USUBJID", "Unique Subject Identifier", "Char", "", "STUDYID || '-' || Subject_ID"),
    ("AE.csv", "AE_ID", "AESEQ", "Sequence Number", "Num", "", "Sequential per subject"),
    ("AE.csv", "AE_Term", "AETERM", "Reported Term for the Adverse Event", "Char", "", "Verbatim, direct copy"),
    ("AE.csv", "MedDRA_PT", "AEDECOD", "Dictionary-Derived Term", "Char", "MedDRA PT", "Coded term from MedDRA autocoder (Medical Coding sign-off)"),
    ("AE.csv", "Severity", "AESEV", "Severity/Intensity", "Char", "MILD | MODERATE | SEVERE", "Upper case"),
    ("AE.csv", "Serious", "AESER", "Serious Event", "Char", "Y | N", "Yes→Y; No→N"),
    ("AE.csv", "Causality", "AEREL", "Causality", "Char", "NOT RELATED | UNLIKELY | POSSIBLE | PROBABLE | DEFINITE", "Unrelated→NOT RELATED; others upper case"),
    ("AE.csv", "Outcome", "AEOUT", "Outcome of Adverse Event", "Char", "RECOVERED/RESOLVED | RECOVERING/RESOLVING | NOT RECOVERED/NOT RESOLVED | FATAL | UNKNOWN", "Map to CT"),
    ("AE.csv", "AE_Start_Date", "AESTDTC", "Start Date/Time of Adverse Event", "Char", "ISO 8601", "YYYY-MM-DD"),
    ("AE.csv", "AE_End_Date", "AEENDTC", "End Date/Time of Adverse Event", "Char", "ISO 8601", "YYYY-MM-DD; null if ongoing"),
    ("AE.csv", "Ongoing", "AEENRF", "End Relative to Reference Period", "Char", "ONGOING", "Ongoing = 'Y' → 'ONGOING'"),
    ("AE.csv", "AE_Start_Date", "AESTDY", "Study Day of Start", "Num", "", "AESTDTC − RFSTDTC + 1 (if ≥ RFSTDTC), else AESTDTC − RFSTDTC")])
dom("LB", "Laboratory (LB)", [
    ("LB.csv", "Subject_ID", "USUBJID", "Unique Subject Identifier", "Char", "", "STUDYID || '-' || Subject_ID"),
    ("LB.csv", "HGB_gdL", "LBTESTCD / LBORRES / LBORRESU", "Hemoglobin", "Char", "HGB; g/dL", "One record per analyte"),
    ("LB.csv", "WBC_10e9L", "LBTESTCD / LBORRES / LBORRESU", "Leukocytes", "Char", "WBC; 10^9/L", ""),
    ("LB.csv", "PLT_10e9L", "LBTESTCD / LBORRES / LBORRESU", "Platelets", "Char", "PLAT; 10^9/L", ""),
    ("LB.csv", "ALT_UL", "LBTESTCD / LBORRES / LBORRESU", "Alanine Aminotransferase", "Char", "ALT; U/L", ""),
    ("LB.csv", "CREAT_mgdL", "LBTESTCD / LBORRES / LBORRESU", "Creatinine", "Char", "CREAT; mg/dL", ""),
    ("LB.csv", "LBDAT", "LBDTC", "Date/Time of Specimen Collection", "Char", "ISO 8601", "YYYY-MM-DD")])
wb.save(f"{ROOT}/CDISC_SDTM_Mapping_Specification.xlsx")
print("Excel files built")
