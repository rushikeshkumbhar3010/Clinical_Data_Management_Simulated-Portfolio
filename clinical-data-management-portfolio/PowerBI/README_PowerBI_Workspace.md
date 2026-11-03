# Power BI Workspace – Clinical Data Management

A ready-to-import star schema, DAX measures, Power Query scripts and a theme for the NSCLC-P2-001 simulated Phase II study.
`Clinical_DM_Dashboard.html` is a working preview of the finished report (same measures, same four pages).

## 1. Folder contents
| File | Purpose |
|---|---|
| `data_model/` | Nine CSVs forming the star schema (below) |
| `PowerQuery_Load.m` | M queries with typed columns |
| `DAX_Measures.dax` | 30+ measures grouped by Enrolment / Queries / Safety / Completeness |
| `Clinical_DM_Theme.json` | Report theme (View → Themes → Browse for themes) |
| `Clinical_DM_Dashboard.html` | Interactive preview of the report |

## 2. Data model (star schema)
```
Dim_Site (Site_ID) 1──* Dim_Subject (Subject_ID) 1──* Fact_VitalSigns | Fact_Labs | Fact_AdverseEvents | Fact_Queries
Dim_Visit (Visit) 1──* Fact_VitalSigns, Fact_Labs       Dim_Check (Check_ID) 1──* Fact_Queries
Dim_Date (Date) 1──* Fact_Queries[Date_Issued], Fact_AdverseEvents[AE_Start_Date], Dim_Subject[Consent_Date]   (first = active, others inactive)
```
Relationships: all one-to-many, single direction (dimension → fact). Mark `Dim_Date[Date]` as the date table. Hide key columns on fact tables.

## 3. Build steps
1. Power BI Desktop → Get data → Blank query → set parameter `FolderPath`, then paste each query from `PowerQuery_Load.m`.
2. Model view → create the relationships above.
3. Enter Data → empty table `_Measures` → paste measures from `DAX_Measures.dax`.
4. View → Themes → Browse → `Clinical_DM_Theme.json`.
5. Build the four pages below. Add a **Site** slicer (Dim_Site[Site_ID]) and sync it across pages (View → Sync slicers).

## 4. Report pages
| Page | Visuals |
|---|---|
| Study overview | 6 KPI cards (Subjects, Total Queries, Closure Rate %, SLA Met %, Backlog, Serious AEs); donut Query_Status; bar Check_ID × Total Queries; column Enrolment by month; line Issued vs Answered by Week_Start; bar Queries by Site |
| Query management | Stacked column Site × Query_Status; ageing bar (bins of Days_Open); TAT histogram; table of open queries sorted by Days_Open |
| Safety (AEs) | Bar top 10 MedDRA_PT; donut Severity; column AEs by month; bar Serious AEs by Site |
| Site data quality | Scorecard table (Subjects, Total Queries, Queries per 100 Subjects with data bars, Backlog, Closure %, SLA Met % with `SLA Colour` conditional formatting); matrix Check_ID × Site_ID with background colour scale (white → #0D366B) |

Enable cross-filtering and drill-through from Site to a subject-level query list (drill-through field: Subject_ID).

## 5. Refresh
Re-run `scripts/01`–`05` to regenerate data, then Home → Refresh in Power BI.
