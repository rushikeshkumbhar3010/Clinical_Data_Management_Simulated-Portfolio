const fs = require("fs");
const d = require("docx");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle, AlignmentType,
  HeadingLevel, LevelFormat, Header, Footer, PageNumber, PageBreak } = d;

const NAVY = "1F3864", TEAL = "0E7C86", W = 9360 + 1440 - 1440; // A4 content width ≈ 9638 with 1" margins -> use 9638
const CW = 9638;
const border = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const borders = { top: border, bottom: border, left: border, right: border };

const p = (text, o = {}) => new Paragraph({ spacing: { after: 100, line: 276 }, ...o.para,
  children: [new TextRun({ text, font: "Arial", size: 21, ...o.run })] });
const rich = (parts, para = {}) => new Paragraph({ spacing: { after: 100, line: 276 }, ...para,
  children: parts.map(([t, b]) => new TextRun({ text: t, bold: !!b, font: "Arial", size: 21 })) });
const h1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(t)] });
const h2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)] });
const bullet = (t, boldLead) => new Paragraph({ numbering: { reference: "b", level: 0 }, spacing: { after: 60, line: 264 },
  children: boldLead ? [new TextRun({ text: boldLead, bold: true, font: "Arial", size: 21 }), new TextRun({ text: t, font: "Arial", size: 21 })]
                     : [new TextRun({ text: t, font: "Arial", size: 21 })] });

function cell(text, w, o = {}) {
  return new TableCell({ width: { size: w, type: WidthType.DXA }, borders,
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: 70, bottom: 70, left: 100, right: 100 },
    children: String(text).split("\n").map(t => new Paragraph({ spacing: { after: 20 },
      children: [new TextRun({ text: t, font: "Arial", size: 19, bold: !!o.bold, color: o.color })] })) });
}
function table(widths, head, rows) {
  return new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths,
    rows: [new TableRow({ tableHeader: true, children: head.map((h, i) => cell(h, widths[i], { fill: NAVY, bold: true, color: "FFFFFF" })) }),
      ...rows.map((r, k) => new TableRow({ cantSplit: true, children: r.map((c, i) => cell(c, widths[i], { fill: k % 2 ? "F2F6FA" : undefined, bold: i === 0 })) }))] });
}
const gap = () => new Paragraph({ spacing: { after: 80 }, children: [] });

const doc = new Document({
  creator: "Rushikesh Kumbhar", title: "Data Management Plan",
  styles: {
    default: { document: { run: { font: "Arial", size: 21 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: "Arial", size: 28, bold: true, color: NAVY },
        paragraph: { spacing: { before: 280, after: 120 }, keepNext: true, outlineLevel: 0,
          border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: TEAL, space: 2 } } } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { font: "Arial", size: 23, bold: true, color: TEAL }, paragraph: { spacing: { before: 180, after: 80 }, keepNext: true, outlineLevel: 1 } },
    ] },
  numbering: { config: [{ reference: "b", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
    style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] }] },
  sections: [{
    properties: { page: { margin: { top: 1300, bottom: 1200, left: 1100, right: 1100 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT,
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF", space: 4 } },
      children: [new TextRun({ text: "Data Management Plan  |  Protocol NSCLC-P2-001  |  Version 1.0", font: "Arial", size: 16, color: "7F7F7F" })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "Confidential – Simulated portfolio document   |   Page ", font: "Arial", size: 16, color: "7F7F7F" }),
        new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 16, color: "7F7F7F" }),
        new TextRun({ text: " of ", font: "Arial", size: 16, color: "7F7F7F" }),
        new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "Arial", size: 16, color: "7F7F7F" })] })] }) },
    children: [
      new Paragraph({ spacing: { before: 200, after: 40 }, children: [new TextRun({ text: "DATA MANAGEMENT PLAN", font: "Arial", size: 44, bold: true, color: NAVY })] }),
      new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "Phase II Open-Label Oncology Study in Advanced Non-Small Cell Lung Cancer", font: "Arial", size: 26, color: TEAL })] }),
      table([2600, 7038], ["Document Control", ""], [
        ["Protocol Number", "NSCLC-P2-001"], ["Document Version / Date", "v1.0  |  03-Nov-2026"],
        ["Author", "Rushikesh Kumbhar, Clinical Data Associate (CDA)"], ["Status", "Draft for Clinical Data Lead review"]]),
      h1("1. Study Overview"),
      table([2600, 7038], ["Item", "Description"], [
        ["Study Title", "A Phase II, Open-Label, Multicentre Study of an Investigational Product in Subjects with Advanced Non-Small Cell Lung Cancer"],
        ["Protocol Number", "NSCLC-P2-001"], ["Phase / Design", "Phase II, single-arm, open-label"],
        ["Sample Size", "N = 350 subjects across 10 investigator sites"],
        ["Study Population", "Adults aged 18–75 years with histologically confirmed advanced NSCLC"],
        ["Data Capture", "Electronic CRF (EDC) with real-time system edit checks; central laboratory data transferred as external data"],
        ["Visit Schedule", "Screening (Day -14 to -1) → Baseline/C1D1 (Day 1 ±1) → Treatment Evaluation (Day 28 ±3) → End of Study / Safety Follow-Up (30 days post last dose)"]]),
      h2("1.1 Primary Objectives"),
      bullet("Evaluate the preliminary efficacy of the investigational product in subjects with advanced NSCLC."),
      bullet("Characterise the safety and tolerability profile, including adverse events (AEs) and serious adverse events (SAEs)."),
      bullet("Ensure that collected data are complete, accurate, consistent and verifiable (ALCOA+) and ready for CDISC SDTM tabulation."),
      h2("1.2 Scope"),
      p("This plan covers data collection, validation, query management, medical coding, external data reconciliation, database lock and data transfer for DM, IE, MH, VS, LB, EC, CM, AE, PE and DS domains. It complies with ICH-GCP E6(R2), 21 CFR Part 11 and applicable data protection regulations."),

      h1("2. Roles & Responsibilities"),
      table([2300, 7338], ["Role", "Responsibilities"], [
        ["Clinical Data Lead (CDL)", "Owns the DMP; approves eCRF design, edit checks and database lock; escalates data quality risks; reviews metrics with the study team."],
        ["Clinical Data Associate (CDA)", "Builds and tests edit checks (UAT); performs manual data review; issues, tracks and closes queries; reconciles SAE and external data; maintains the discrepancy log and query metrics."],
        ["Principal Investigator (PI)", "Accountable for accuracy of site data; responds to queries within timelines; provides electronic signature on completed casebooks before lock."],
        ["Clinical Research Associate (CRA / Monitor)", "Performs source data verification (SDV) and monitoring; supports sites in query resolution; reports protocol deviations."],
        ["Medical Coder / Safety Physician", "Codes AEs and medications (MedDRA / WHODrug); reviews SAE narratives and medical plausibility."],
        ["Biostatistician / Programmer", "Reviews data structure; consumes SDTM datasets; provides the statistical analysis plan inputs."]]),

      h1("3. Data Handling Procedures"),
      h2("3.1 eCRF Completion Guidelines"),
      bullet("Data are entered by authorised site staff within 5 business days of the visit; source documents are the primary record.", "Timeliness: "),
      bullet("Dates use YYYY-MM-DD format; units are fixed on the form; unknown dates are recorded using partial dates only where allowed.", "Format: "),
      bullet("No blank mandatory fields; missing values are explained using a 'Not Done' flag and a reason.", "Completeness: "),
      bullet("Verbatim AE terms are recorded as spoken (no abbreviations); one event per record.", "AE terms: "),
      bullet("Every change is captured in the audit trail with user, date/time and reason for change.", "Changes: "),
      h2("3.2 Source Data Verification (SDV) Criteria"),
      table([3000, 3100, 3538], ["Data Category", "SDV Level", "Rationale"], [
        ["Informed consent, eligibility (IE), SAEs, primary endpoint data", "100% SDV", "Subject safety and critical to quality"],
        ["Vital signs, laboratory, study drug administration", "Risk-based: 100% for first 3 subjects per site, then 30% random sample", "Moderate impact; escalate to 100% if error rate > 5%"],
        ["Concomitant medications, non-serious AEs", "Targeted SDV on key fields", "Lower risk; centrally reviewed"]]),
      h2("3.3 Data Validation & Query Resolution Timelines"),
      p("Edit checks are specified in Data_Validation_Plan_and_Edit_Checks.xlsx (13 checks). System checks fire at data entry; manual checks run weekly on the cleaning listings. Queries are written in a neutral, non-leading way and never suggest a value."),
      table([3000, 2300, 4338], ["Query Type", "Site Response Target", "Escalation"], [
        ["Standard data queries", "5 business days", "Reminder at day 3; CRA follow-up at day 5; CDL review at day 10"],
        ["SAE-related queries", "24 hours", "Same-day call to site and safety physician notification"],
        ["Re-issued queries", "3 business days", "Escalate to PI if unresolved after two re-issues"]]),
      p("Query lifecycle: Open → Answered → (Re-issued) → Closed. Every query is logged in Master_Discrepancy_and_Query_Tracker.xlsx with date issued, response, closure initials and audit trail timestamp."),
      h2("3.4 External Data, SAE Reconciliation and Medical Coding"),
      bullet("Central laboratory data are loaded per the data transfer agreement and reconciled against the eCRF at least monthly."),
      bullet("SAE data in the safety database are reconciled against the eCRF AE form before each interim review and prior to lock."),
      bullet("AEs are coded to MedDRA (Preferred Term and SOC); medications to WHODrug. Coding queries go through the CDA before the coder."),

      h2("3.5 Data Flow and Systems"),
      table([2300, 7338], ["Stage", "Description"], [
        ["1. Capture", "Sites enter source data into the EDC; system edit checks fire in real time and raise automatic queries."],
        ["2. Review", "CDA reviews cleaning listings (missing visits, date logic, duplicates, lab ranges); CRA performs SDV."],
        ["3. Resolve", "Queries are answered by the site, reviewed by the CDA and closed or re-issued."],
        ["4. Reconcile", "SAE, central lab and coding reconciliation completed against the clinical database."],
        ["5. Tabulate", "Cleaned raw data are mapped to CDISC SDTM (DM, VS, AE, LB) per the mapping specification."],
        ["6. Report", "Query metrics and data quality indicators refresh in the Power BI dashboard each week."]]),
      h2("3.6 Protocol Deviations and Data Quality Risks"),
      p("Deviations identified during data review (for example, ages outside 18–75 years, assessments outside the visit window, or events dated before informed consent) are documented in the deviation log and shared with the CRA. Data management does not classify deviations as major or minor; this is a clinical team decision."),
      table([3000, 3300, 3338], ["Risk", "Mitigation", "Indicator"], [
        ["Missing Day 28 vital signs", "Missing-visit listing reviewed weekly", "% of expected VS records present"],
        ["Late or incorrect AE dates", "Date-sequence checks CK_AE_01 and CK_AE_03", "Number of AE date discrepancies per site"],
        ["Unreconciled SAEs", "Monthly SAE reconciliation", "SAEs without an SAE form (CK_AE_02)"],
        ["Implausible laboratory values", "Range checks and unit verification", "Lab queries per 100 results"]]),
      h2("3.7 Standards and Dictionaries"),
      bullet("CDISC CDASH (eCRF design), SDTM-IG 3.4 (tabulation) and NCI Controlled Terminology."),
      bullet("MedDRA (current version at time of coding) for AEs and medical history; WHODrug Global for concomitant medications."),
      bullet("21 CFR Part 11 compliant audit trail, electronic signatures and role-based access."),

      new Paragraph({ children: [new PageBreak()] }),
      h1("4. Database Lock Procedures"),
      table([700, 2700, 6238], ["Step", "Activity", "Criteria / Owner"], [
        ["1", "Pre-lock data review", "All expected visits entered; no open queries; missing pages resolved. Owner: CDA"],
        ["2", "Soft lock", "Database becomes read-only for all users except Data Management; late changes need CDL approval. Owner: CDL"],
        ["3", "Final audit trail reconciliation", "Audit trail reviewed for unexplained or late changes; SAE and external data reconciliation completed and signed. Owner: CDA / CDL"],
        ["4", "Medical coding sign-off", "100% of AEs and medications coded and approved by medical coder / safety physician."],
        ["5", "PI signature sign-off", "Each PI electronically signs the final casebook for all subjects at their site."],
        ["6", "Hard lock", "Complete database freeze after written approval from CDL, Biostatistics and Clinical Lead; all user access revoked; archive created."],
        ["7", "Post-lock", "Final SDTM datasets, define.xml and data transfer to the sponsor. Unlock only via documented change control."]]),
      h2("4.1 Quality Metrics Reported to the Study Team"),
      bullet("Query rate (queries per 100 data points), open query ageing, and median response time per site."),
      bullet("Percentage of queries closed within SLA; percentage of SDV completed; SAE reconciliation status."),
      bullet("Current status is visualised in the Power BI Clinical Data Management dashboard (see PowerBI folder)."),

      h2("4.2 Abbreviations"),
      table([1700, 7938], ["Term", "Meaning"], [
        ["AE / SAE", "Adverse Event / Serious Adverse Event"], ["CDA / CDL", "Clinical Data Associate / Clinical Data Lead"],
        ["CRA", "Clinical Research Associate (Monitor)"], ["EDC / eCRF", "Electronic Data Capture / electronic Case Report Form"],
        ["GCP", "Good Clinical Practice"], ["SDV", "Source Data Verification"], ["SDTM", "Study Data Tabulation Model (CDISC)"]]),
      h1("5. Approvals"),
      table([2800, 2800, 2200, 1838], ["Role", "Name", "Signature", "Date"], [
        ["Author (CDA)", "Rushikesh Kumbhar", "", ""], ["Clinical Data Lead", "", "", ""], ["Principal Investigator (representative)", "", "", ""]]),
    ] }],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(process.argv[2], b); console.log("DMP written"); });
