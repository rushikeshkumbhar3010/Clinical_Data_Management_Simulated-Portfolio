// Power Query (M) – paste into Advanced Editor, one query per table.
// Set the parameter FolderPath to the full path of PowerBI\data_model before refreshing.
// Parameter:  FolderPath = "C:\clinical-data-management-portfolio\PowerBI\data_model"  (Text, required)

// ---- Fact_Queries ----
let
    Source  = Csv.Document(File.Contents(FolderPath & "\Fact_Queries.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed   = Table.TransformColumnTypes(Headers, {{"Date_Issued", type date}, {"Date_Answered", type date}, {"SLA_Target_Days", Int64.Type}, {"TAT_Days", Int64.Type}, {"Days_Open", Int64.Type}}, "en-US")
in
    Typed

// ---- Fact_AdverseEvents ----
let
    Source  = Csv.Document(File.Contents(FolderPath & "\Fact_AdverseEvents.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed   = Table.TransformColumnTypes(Headers, {{"AE_Start_Date", type date}, {"AE_End_Date", type date}}, "en-US")
in
    Typed

// ---- Fact_VitalSigns ----
let
    Source  = Csv.Document(File.Contents(FolderPath & "\Fact_VitalSigns.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed   = Table.TransformColumnTypes(Headers, {{"VSDAT", type date}, {"Systolic_BP", Int64.Type}, {"Diastolic_BP", Int64.Type}, {"Pulse", Int64.Type}, {"Temp_C", type number}, {"Weight_kg", type number}}, "en-US")
in
    Typed

// ---- Fact_Labs ----
let
    Source  = Csv.Document(File.Contents(FolderPath & "\Fact_Labs.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed   = Table.TransformColumnTypes(Headers, {{"LBDAT", type date}, {"HGB_gdL", type number}, {"WBC_10e9L", type number}, {"PLT_10e9L", Int64.Type}, {"ALT_UL", Int64.Type}, {"CREAT_mgdL", type number}}, "en-US")
in
    Typed

// ---- Dim_Subject ----   (same pattern for Dim_Site, Dim_Visit, Dim_Check)
let
    Source  = Csv.Document(File.Contents(FolderPath & "\Dim_Subject.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed   = Table.TransformColumnTypes(Headers, {{"Consent_Date", type date}, {"Age", Int64.Type}}, "en-US")
in
    Typed

// ---- Dim_Date ----  (mark as date table on [Date])
let
    Source  = Csv.Document(File.Contents(FolderPath & "\Dim_Date.csv"), [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]),
    Headers = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    Typed   = Table.TransformColumnTypes(Headers, {{"Date", type date}, {"Week_Start", type date}, {"Year", Int64.Type}, {"Month_Num", Int64.Type}}, "en-US")
in
    Typed
