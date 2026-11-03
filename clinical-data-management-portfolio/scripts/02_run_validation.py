"""Execute the Data Validation Plan (edit checks) with pandas and write the discrepancy listing.
Run:  python scripts/02_run_validation.py
"""
import pandas as pd, numpy as np, os
D = os.path.join(os.path.dirname(__file__), "..", "Simulated_Clinical_Trial_Dataset")
dm = pd.read_csv(f"{D}/DM.csv", parse_dates=["Informed_Consent_Date"])
vs = pd.read_csv(f"{D}/VS.csv", parse_dates=["VSDAT"])
lb = pd.read_csv(f"{D}/LB.csv", parse_dates=["LBDAT"])
ae = pd.read_csv(f"{D}/AE.csv", parse_dates=["AE_Start_Date", "AE_End_Date"])
cons = dm.set_index("Subject_ID").Informed_Consent_Date
f = lambda d: "" if pd.isna(d) else pd.Timestamp(d).strftime("%d-%b-%Y")
out = []

def add(cid, form, field, mode, r, text, rec=""):
    out.append(dict(Check_ID=cid, Site_ID=r.Site_ID, Subject_ID=r.Subject_ID, Form_Name=form, Field_Name=field,
                    Check_Type=mode, Record_Ref=rec, Query_Text=text))

for _, r in dm[(dm.Age < 18) | (dm.Age > 75)].iterrows():
    add("CK_DM_01", "Demographics", "Age", "System", r,
        f"Subject age is recorded as {r.Age}. Protocol requires enrolled subjects to be aged 18-75 years. "
        "Please verify the date of birth/age or document a protocol deviation.")
for _, r in dm[dm.Race.isna()].iterrows():
    add("CK_DM_02", "Demographics", "Race", "Manual", r, "Race is blank. Please provide the subject's race per source or confirm it was not collected.")

m = vs[vs.Systolic_BP.isna() & vs.Diastolic_BP.isna() & vs.Pulse.isna()]
for _, r in m[m.Visit.str.startswith("Visit 3")].iterrows():
    add("CK_VS_03", "Vital Signs", "VS (all fields)", "System", r, "Vital signs are missing for Visit 3 (Day 28). Please provide the assessment or confirm that it was not done.", r.Visit)
for _, r in vs[vs.Systolic_BP <= vs.Diastolic_BP].iterrows():
    add("CK_VS_01", "Vital Signs", "Systolic_BP / Diastolic_BP", "System", r,
        f"Systolic BP ({int(r.Systolic_BP)} mmHg) is less than or equal to diastolic BP ({int(r.Diastolic_BP)} mmHg) at {r.Visit}. Please verify against source and correct if necessary.", r.Visit)
for _, r in vs[(vs.Pulse < 40) | (vs.Pulse > 140)].iterrows():
    add("CK_VS_02", "Vital Signs", "Pulse", "System", r,
        f"Pulse rate ({int(r.Pulse)} bpm) at {r.Visit} is outside the expected range of 40-140 bpm. Please verify against source.", r.Visit)
for _, r in vs[(vs.Temp_C < 34) | (vs.Temp_C > 41)].iterrows():
    add("CK_VS_05", "Vital Signs", "Temp_C", "System", r,
        f"Body temperature ({r.Temp_C} C) at {r.Visit} is outside the plausible range. Please verify value and unit against source.", r.Visit)
for _, r in vs[vs.VSDAT.notna() & (vs.VSDAT < vs.Subject_ID.map(cons))].iterrows():
    add("CK_VS_04", "Vital Signs", "VSDAT", "System", r,
        f"Vital signs date ({f(r.VSDAT)}) at {r.Visit} precedes the informed consent date ({f(cons[r.Subject_ID])}). Please verify and correct the date or provide clarification.", r.Visit)
for _, r in lb[(lb.HGB_gdL < 5) | (lb.HGB_gdL > 20)].iterrows():
    add("CK_LB_01", "Laboratory", "HGB_gdL", "System", r,
        f"Hemoglobin result of {r.HGB_gdL:g} g/dL at {r.Visit} is outside the physiologically plausible range. Please verify the value and unit against the lab report.", r.Visit)

for _, r in ae[ae.AE_End_Date < ae.AE_Start_Date].iterrows():
    add("CK_AE_01", "Adverse Events", "AE_End_Date", "System", r,
        f"AE '{r.AE_Term}': stop date ({f(r.AE_End_Date)}) precedes start date ({f(r.AE_Start_Date)}). Please verify and correct the dates or provide clarification.", r.AE_ID)
for _, r in ae[ae.AE_Start_Date < ae.Subject_ID.map(cons)].iterrows():
    add("CK_AE_03", "Adverse Events", "AE_Start_Date", "System", r,
        f"AE '{r.AE_Term}': start date is recorded as {f(r.AE_Start_Date)}, which precedes the informed consent date of {f(cons[r.Subject_ID])}. Please verify and correct the date or provide clarification.", r.AE_ID)
for _, r in ae[(ae.Ongoing != "Y") & (ae.AE_End_Date.isna() | ae.Outcome.isna())].iterrows():
    add("CK_AE_04", "Adverse Events", "AE_End_Date / Outcome", "Manual", r,
        f"AE '{r.AE_Term}' (start {f(r.AE_Start_Date)}) has no stop date and no outcome recorded, and is not marked as ongoing. Please provide the missing information.", r.AE_ID)
for _, r in ae[(ae.Serious == "Yes") & (ae.SAE_Form_Received != "Y")].iterrows():
    add("CK_AE_02", "Adverse Events", "Serious / SAE Form", "System", r,
        f"AE '{r.AE_Term}' (start {f(r.AE_Start_Date)}) is marked serious, but no SAE report form has been received. Please complete the SAE form or clarify the seriousness criteria.", r.AE_ID)
dd = ae[ae.duplicated(["Subject_ID", "AE_Term", "AE_Start_Date"], keep="first")]
for _, r in dd.iterrows():
    add("CK_AE_05", "Adverse Events", "AE_Term / AE_Start_Date", "Manual", r,
        f"AE '{r.AE_Term}' with start date {f(r.AE_Start_Date)} appears more than once for this subject. Please confirm whether these are separate events or a duplicate entry.", r.AE_ID)

disc = pd.DataFrame(out)
disc.to_csv(f"{D}/../scripts/discrepancies_detected.csv", index=False)
print(disc.Check_ID.value_counts().sort_index().to_string())
print("TOTAL discrepancies:", len(disc))
