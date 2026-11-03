"""Build the Power BI star-schema CSVs + dashboard data JSON.  Run after 01-03."""
import pandas as pd, numpy as np, os, json
ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")
D = f"{ROOT}/Simulated_Clinical_Trial_Dataset"; M = f"{ROOT}/PowerBI/data_model"; os.makedirs(M, exist_ok=True)
dm = pd.read_csv(f"{D}/DM.csv", parse_dates=["Informed_Consent_Date"]); vs = pd.read_csv(f"{D}/VS.csv", parse_dates=["VSDAT"])
lb = pd.read_csv(f"{D}/LB.csv", parse_dates=["LBDAT"]); ae = pd.read_csv(f"{D}/AE.csv", parse_dates=["AE_Start_Date", "AE_End_Date"])
sites = pd.read_csv(f"{D}/Sites.csv")
q = pd.read_csv(f"{ROOT}/scripts/query_log_export.csv", parse_dates=["Date_Issued", "Date_Answered", "Audit"])
CUT = pd.Timestamp("2026-08-31")

dim_site = sites.copy(); dim_site.to_csv(f"{M}/Dim_Site.csv", index=False)
dm.rename(columns={"Informed_Consent_Date": "Consent_Date"}).to_csv(f"{M}/Dim_Subject.csv", index=False)
pd.DataFrame({"Visit": ["Screening", "Baseline/C1D1", "Visit 3 (Day 28)", "End of Study"], "Visit_Num": [1, 2, 3, 4]}).to_csv(f"{M}/Dim_Visit.csv", index=False)
dates = pd.date_range("2026-01-01", "2026-12-31")
pd.DataFrame({"Date": dates, "Year": dates.year, "Month_Num": dates.month, "Month": dates.strftime("%b"), "Week_Start": (dates - pd.to_timedelta(dates.weekday, unit="D"))}).to_csv(f"{M}/Dim_Date.csv", index=False)
checks = pd.read_csv(f"{ROOT}/scripts/discrepancies_detected.csv")[["Check_ID", "Form_Name"]].drop_duplicates()
prio = {"CK_AE_02": "Critical", "CK_VS_02": "Medium", "CK_VS_05": "Medium", "CK_AE_04": "Medium", "CK_AE_05": "Medium", "CK_DM_02": "Low"}
checks["Priority"] = checks.Check_ID.map(prio).fillna("High")
checks.sort_values("Check_ID").to_csv(f"{M}/Dim_Check.csv", index=False)
vs.to_csv(f"{M}/Fact_VitalSigns.csv", index=False); lb.to_csv(f"{M}/Fact_Labs.csv", index=False); ae.to_csv(f"{M}/Fact_AdverseEvents.csv", index=False)
q["TAT_Days"] = [np.nan if pd.isna(a) else np.busday_count(i.date(), a.date()) for i, a in zip(q.Date_Issued, q.Date_Answered)]
q["Days_Open"] = np.where(q.Status == "Closed", np.nan, (CUT - q.Date_Issued).dt.days)
q[["Query_ID", "Check_ID", "Site_ID", "Subject_ID", "Form_Name", "Field_Name", "Check_Type", "Query_Text", "Date_Issued", "Date_Answered", "Status", "SLA", "TAT_Days", "Days_Open"]] \
    .rename(columns={"Status": "Query_Status", "SLA": "SLA_Target_Days"}).to_csv(f"{M}/Fact_Queries.csv", index=False)

# ---- dashboard JSON ----
dj = dict(
    sites=[dict(id=r.Site_ID, name=r.Site_Name.replace(", ", " – ")) for r in sites.itertuples()],
    subj=[dict(s=r.Site_ID, m=r.Informed_Consent_Date.strftime("%Y-%m")) for r in dm.itertuples()],
    q=[dict(id=r.Query_ID, c=r.Check_ID, s=r.Site_ID, f=r.Form_Name, st=r.Status, i=r.Date_Issued.strftime("%Y-%m-%d"),
            a=None if pd.isna(r.Date_Answered) else r.Date_Answered.strftime("%Y-%m-%d"), sla=int(r.SLA),
            tat=None if pd.isna(r.TAT_Days) else int(r.TAT_Days), subj=r.Subject_ID) for r in q.itertuples()],
    ae=[dict(s=r.Site_ID, t=r.MedDRA_PT, sev=r.Severity, ser=r.Serious, m=r.AE_Start_Date.strftime("%Y-%m"), rel=r.Causality) for r in ae.itertuples()],
    vsn=[dict(s=r.Site_ID, v=r.Visit, sbp=None if pd.isna(r.Systolic_BP) else int(r.Systolic_BP)) for r in vs.itertuples()],
    cut="2026-08-31")
json.dump(dj, open(f"{ROOT}/PowerBI/dashboard_data.json", "w"), separators=(",", ":"))
print("model ok", {f: len(pd.read_csv(f'{M}/{f}')) for f in os.listdir(M)})
