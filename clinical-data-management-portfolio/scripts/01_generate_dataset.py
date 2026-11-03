"""Simulate a 350-subject Phase II oncology dataset (DM, VS, LB, AE) and inject ~183 deliberate errors.
Run:  python scripts/01_generate_dataset.py
"""
import numpy as np, pandas as pd, os, datetime as dt

rng = np.random.default_rng(2026)
OUT = os.path.join(os.path.dirname(__file__), "..", "Simulated_Clinical_Trial_Dataset")
os.makedirs(OUT, exist_ok=True)
CUT = pd.Timestamp("2026-08-31")           # database cut-off date
N = 350

SITES = {
    "S01": "Tata Memorial Centre, Mumbai", "S02": "AIIMS, New Delhi", "S03": "Kidwai Memorial, Bengaluru",
    "S04": "CMC Vellore", "S05": "Cancer Institute, Chennai", "S06": "Ruby Hall Clinic, Pune",
    "S07": "Apollo Hospitals, Hyderabad", "S08": "Tata Medical Center, Kolkata",
    "S09": "Gujarat Cancer Research Inst., Ahmedabad", "S10": "SGPGI, Lucknow",
}
site_ids = list(SITES)
weights = np.array([.16, .14, .12, .11, .10, .09, .08, .08, .06, .06]); weights /= weights.sum()

# ---------------- DM ----------------
site = rng.choice(site_ids, N, p=weights)
counter = {s: 0 for s in site_ids}
subj = []
for s in site:
    counter[s] += 1
    subj.append(f"{s}-{counter[s]:03d}")
consent = pd.to_datetime("2026-01-15") + pd.to_timedelta(rng.integers(0, 106, N), unit="D")  # to 30-Apr
dm = pd.DataFrame({
    "Subject_ID": subj, "Site_ID": site, "Informed_Consent_Date": consent,
    "Age": np.clip(rng.normal(58, 9, N).round(), 22, 74).astype(int),
    "Sex": rng.choice(["Male", "Female"], N, p=[.62, .38]),
    "Race": rng.choice(["Asian", "White", "Black/African American", "Native", "Other"], N, p=[.82, .08, .04, .02, .04]),
})
dm = dm.sort_values(["Site_ID", "Subject_ID"]).reset_index(drop=True)

# ---------------- VS ----------------
VISITS = ["Screening", "Baseline/C1D1", "Visit 3 (Day 28)", "End of Study"]
rows = []
base_w = dict(zip(dm.Subject_ID, np.clip(rng.normal(66, 12, N), 42, 115)))
for _, r in dm.iterrows():
    c = r.Informed_Consent_Date
    d_scr = c + pd.Timedelta(days=int(rng.integers(0, 4)))
    d_bl = c + pd.Timedelta(days=int(rng.integers(14, 21)))
    d_v3 = d_bl + pd.Timedelta(days=28 + int(rng.integers(-3, 4)))
    d_eos = d_v3 + pd.Timedelta(days=30 + int(rng.integers(0, 4)))
    for v, d in zip(VISITS, [d_scr, d_bl, d_v3, d_eos]):
        sys_ = int(np.clip(rng.normal(124, 12), 96, 165))
        dia = int(np.clip(rng.normal(78, 8), 58, 98))
        if sys_ - dia < 20: sys_ = dia + 25
        rows.append(dict(Subject_ID=r.Subject_ID, Site_ID=r.Site_ID, Visit=v, VSDAT=d, Systolic_BP=sys_,
                         Diastolic_BP=dia, Pulse=int(np.clip(rng.normal(76, 10), 52, 110)),
                         Temp_C=round(float(np.clip(rng.normal(36.7, .3), 35.8, 37.8)), 1),
                         Weight_kg=round(float(base_w[r.Subject_ID] + rng.normal(0, .8)), 1)))
vs = pd.DataFrame(rows)

# ---------------- LB ----------------
rows = []
for _, r in dm.iterrows():
    v = vs[(vs.Subject_ID == r.Subject_ID)]
    for visit, vname in [("Screening", "Screening"), ("Visit 3 (Day 28)", "Visit 3 (Day 28)")]:
        d = v[v.Visit == vname].VSDAT.iloc[0]
        rows.append(dict(Subject_ID=r.Subject_ID, Site_ID=r.Site_ID, Visit=visit, LBDAT=d,
                         HGB_gdL=round(float(np.clip(rng.normal(12.8, 1.4), 9.5, 16.8)), 1),
                         WBC_10e9L=round(float(np.clip(rng.normal(7.2, 1.8), 3.5, 12.5)), 1),
                         PLT_10e9L=int(np.clip(rng.normal(255, 55), 140, 420)),
                         ALT_UL=int(np.clip(rng.normal(28, 10), 8, 70)),
                         CREAT_mgdL=round(float(np.clip(rng.normal(0.9, .18), .5, 1.4)), 2)))
lb = pd.DataFrame(rows)

# ---------------- AE (clean) ----------------
TERMS = [("Nausea", "Nausea"), ("Fatigue", "Fatigue"), ("Headache", "Headache"), ("Diarrhoea", "Diarrhoea"),
         ("Skin rash", "Rash"), ("Vomiting", "Vomiting"), ("Decreased appetite", "Decreased appetite"),
         ("Cough", "Cough"), ("Neutropenia", "Neutropenia"), ("Anaemia", "Anaemia"),
         ("Increased ALT", "Alanine aminotransferase increased"), ("Dyspnoea", "Dyspnoea"),
         ("Constipation", "Constipation"), ("Peripheral neuropathy", "Peripheral sensory neuropathy"),
         ("Pyrexia", "Pyrexia"), ("Pneumonitis", "Pneumonitis")]
tp = np.array([12, 14, 8, 9, 7, 7, 8, 6, 6, 5, 4, 4, 3, 2, 3, 2], float); tp /= tp.sum()
ae_rows = []
bl_dates = vs[vs.Visit == "Baseline/C1D1"].set_index("Subject_ID").VSDAT
for _, r in dm.iterrows():
    for _ in range(rng.poisson(1.45)):
        ti = rng.choice(len(TERMS), p=tp)
        start = bl_dates[r.Subject_ID] + pd.Timedelta(days=int(rng.integers(0, 80)))
        if start > CUT: start = CUT - pd.Timedelta(days=int(rng.integers(1, 10)))
        sev = rng.choice(["Mild", "Moderate", "Severe"], p=[.55, .33, .12])
        serious = "Yes" if (sev == "Severe" and rng.random() < .45) else "No"
        ongoing = "Y" if rng.random() < .10 else "N"
        end, out = start + pd.Timedelta(days=int(rng.integers(1, 21))), "Recovered/Resolved"
        if end > CUT: ongoing = "Y"
        if ongoing == "Y": end, out = pd.NaT, rng.choice(["Not Recovered", "Recovering/Resolving"])
        elif rng.random() < .015: out = "Fatal"
        elif rng.random() < .10: out = "Recovering/Resolving"
        ae_rows.append(dict(Subject_ID=r.Subject_ID, Site_ID=r.Site_ID, AE_Term=TERMS[ti][0], MedDRA_PT=TERMS[ti][1],
                            AE_Start_Date=start, AE_End_Date=end, Ongoing=ongoing, Severity=sev, Serious=serious,
                            Causality=rng.choice(["Unrelated", "Unlikely", "Possible", "Probable", "Definite"],
                                                 p=[.2, .2, .3, .22, .08]),
                            Outcome=out, SAE_Form_Received="Y" if serious == "Yes" else "NA"))
ae = pd.DataFrame(ae_rows)

# ---------------- Error injection ----------------
errs = []                                   # answer key
used_vs, used_ae, used_lb, used_dm = set(), set(), set(), set()

def pick(pool, n, used):
    pool = [p for p in pool if p not in used]
    ch = list(rng.choice(pool, n, replace=False))
    used.update(ch); return ch

v3 = vs.index[vs.Visit == "Visit 3 (Day 28)"].tolist()
base = vs.index[vs.Visit != "Visit 3 (Day 28)"].tolist()
allvs = vs.index.tolist()

for i in pick(v3, 25, used_vs):                                   # 25 missing V3 vitals
    vs.loc[i, ["VSDAT", "Systolic_BP", "Diastolic_BP", "Pulse", "Temp_C", "Weight_kg"]] = np.nan
    errs.append(("CK_VS_03", i, "VS"))
for i in pick(allvs, 15, used_vs):                                # 15 BP inversions
    s, d = vs.loc[i, "Systolic_BP"], vs.loc[i, "Diastolic_BP"]
    vs.loc[i, "Systolic_BP"], vs.loc[i, "Diastolic_BP"] = d, s
    errs.append(("CK_VS_01", i, "VS"))
for i in pick(allvs, 20, used_vs):                                # 20 pulse out of range
    vs.loc[i, "Pulse"] = int(rng.choice([28, 31, 34, 36, 152, 160, 171, 188]))
    errs.append(("CK_VS_02", i, "VS"))
for i in pick(allvs, 10, used_vs):                                # 10 temp in Fahrenheit
    vs.loc[i, "Temp_C"] = round(float(rng.uniform(97.4, 101.2)), 1)
    errs.append(("CK_VS_05", i, "VS"))
for i in pick(allvs, 10, used_vs):                                # 10 VS dates before consent
    c = dm.set_index("Subject_ID").Informed_Consent_Date[vs.loc[i, "Subject_ID"]]
    vs.loc[i, "VSDAT"] = c - pd.Timedelta(days=int(rng.integers(2, 20)))
    errs.append(("CK_VS_04", i, "VS"))

for i in pick(lb.index.tolist(), 15, used_lb):                    # 15 impossible Hb
    lb.loc[i, "HGB_gdL"] = float(rng.choice([220, 210, 180, 142, 130, 125, 1.3, 0.9]))
    errs.append(("CK_LB_01", i, "LB"))

for i in pick(dm.index.tolist(), 12, used_dm):                    # 12 ages out of range
    dm.loc[i, "Age"] = int(rng.choice([15, 16, 17, 79, 81, 84, 88]))
    errs.append(("CK_DM_01", i, "DM"))
for i in pick(dm.index.tolist(), 8, used_dm):                     # 8 missing race
    dm.loc[i, "Race"] = np.nan
    errs.append(("CK_DM_02", i, "DM"))

cons = dm.set_index("Subject_ID").Informed_Consent_Date
aeall = ae.index.tolist()
for i in pick(aeall, 10, used_ae):                                # 10 AE start before consent
    c = cons[ae.loc[i, "Subject_ID"]]
    ae.loc[i, "AE_Start_Date"] = c - pd.Timedelta(days=int(rng.integers(3, 31)))
    errs.append(("CK_AE_03", i, "AE"))
for i in pick(aeall, 20, used_ae):                                # 20 no outcome/end date, not ongoing
    ae.loc[i, ["AE_End_Date", "Outcome"]] = [pd.NaT, np.nan]
    ae.loc[i, "Ongoing"] = "N"
    errs.append(("CK_AE_04", i, "AE"))
for i in pick(aeall, 15, used_ae):                                # 15 stop before start
    ae.loc[i, "Ongoing"] = "N"; ae.loc[i, "Outcome"] = "Recovered/Resolved"
    ae.loc[i, "AE_End_Date"] = ae.loc[i, "AE_Start_Date"] - pd.Timedelta(days=int(rng.integers(1, 11)))
    errs.append(("CK_AE_01", i, "AE"))
for i in pick(aeall, 15, used_ae):                                # 15 SAE without SAE form
    ae.loc[i, ["Serious", "Severity", "SAE_Form_Received"]] = ["Yes", "Severe", "N"]
    errs.append(("CK_AE_02", i, "AE"))
dup_src = pick(aeall, 8, used_ae)                                 # 8 duplicate AE records
dups = ae.loc[dup_src].copy()
ae = pd.concat([ae, dups], ignore_index=True)
for k, i in enumerate(range(len(ae) - 8, len(ae))):
    errs.append(("CK_AE_05", i, "AE"))

ae = ae.sort_values(["Subject_ID", "AE_Start_Date"], kind="stable").reset_index(drop=True)
# NOTE: sorting changed indices; rebuild answer key by re-detection in 02_run_validation.py, so keep only counts here.
ae.insert(0, "AE_ID", [f"AE-{k+1:04d}" for k in range(len(ae))])

def fmt(df, cols):
    df = df.copy()
    for c in cols: df[c] = pd.to_datetime(df[c]).dt.strftime("%Y-%m-%d")
    return df

fmt(dm, ["Informed_Consent_Date"]).to_csv(f"{OUT}/DM.csv", index=False)
fmt(vs, ["VSDAT"]).to_csv(f"{OUT}/VS.csv", index=False)
fmt(lb, ["LBDAT"]).to_csv(f"{OUT}/LB.csv", index=False)
fmt(ae, ["AE_Start_Date", "AE_End_Date"]).to_csv(f"{OUT}/AE.csv", index=False)
pd.DataFrame([{"Site_ID": k, "Site_Name": v, "PI_Name": f"Dr. Investigator {k}"} for k, v in SITES.items()]) \
    .to_csv(f"{OUT}/Sites.csv", index=False)
print("Rows -> DM", len(dm), "VS", len(vs), "LB", len(lb), "AE", len(ae))
print("Injected:", len(errs))
