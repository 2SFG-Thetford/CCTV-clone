let
    // 1. Load all files from folder
    Source = Folder.Files //("\\xxxx.net\filestore\Thetford Filestore\...."),

    // 2. Filter Excel files (any case)
    ExcelFiles = Table.SelectRows(Source, each Text.Lower(Text.Trim([Extension])) = ".xlsx"),

    // 3. Ensure there is at least one file and get latest
    LatestFileRecord =
        if Table.IsEmpty(ExcelFiles) then
            error "No Excel files found in folder."
        else
            Table.First(Table.Sort(ExcelFiles, {{"Date modified", Order.Descending}})),

    // 4. Load workbook
    Workbook = Excel.Workbook(LatestFileRecord[Content], null, true),

    // 5. Take first sheet
    FirstSheet = Workbook{0}[Data],

    // 6. Promote headers
    PromotedHeaders = Table.PromoteHeaders(FirstSheet, [PromoteAllScalars=true]),

    // 7. Clean column names
    CleanHeaders = Table.TransformColumnNames(PromotedHeaders, Text.Trim),

    // 8. Load Table1 and get existing Refs
    RefTable = Excel.CurrentWorkbook(){[Name = "Table1"]}[Content],
    RefList = List.Transform(RefTable[Ref], each Text.From(_)),

    // 9. Filter rows where Ref not in Table1
    FilterNewRefs = Table.SelectRows(CleanHeaders, each not List.Contains(RefList, Text.From([Ref]))),

    // 10. Convert main [Date] column (not duplicates) to dd/MM/yyyy
    FixDate = if List.Contains(Table.ColumnNames(FilterNewRefs), "Date") then
        Table.TransformColumns(
            FilterNewRefs,
            {{"Date", each try Date.From(DateTime.FromText(_)) otherwise null, type date}}
        )
    else FilterNewRefs,

    // 11. Identify fixed columns (do not collapse)
    FixedColumns = List.Select(
        Table.ColumnNames(FixDate),
        each List.Contains({"Ref","Date","Questionnaire Type","Site","Location","Questionnaire","Auditor","Audit Score","Audit Score Comments"}, _)
    ),

    // 12. Extra columns (duplicates to collapse)
    ExtraColumns = List.Difference(Table.ColumnNames(FixDate), FixedColumns),

    // 13. Optimized collapse of duplicate columns
    DuplicateColumnNames = List.Distinct(ExtraColumns),
    Collapsed =
        List.Accumulate(
            DuplicateColumnNames,
            FixDate,
            (state, colName) =>
                let
                    Matches = List.Select(Table.ColumnNames(state), each _ = colName),
                    Result = if List.Count(Matches) <= 1 then state
                             else
                                Table.AddColumn(
                                    Table.RemoveColumns(state, Matches),
                                    colName,
                                    (r) => List.First(List.RemoveNulls(List.Transform(Matches, each Record.Field(r, _)))),
                                    type any
                                )
                in
                    Result
        ),

    // 14. Remove fully empty columns
    RemoveEmptyColumns = Table.SelectColumns(
        Collapsed,
        List.Select(
            Table.ColumnNames(Collapsed),
            (col) => List.NonNullCount(Table.Column(Collapsed, col)) > 0
        )
    ),

    // 15. Replace Audit Score Comments based on Audit Score
    ReplaceAuditComments =
        Table.RenameColumns(
            Table.RemoveColumns(
                Table.AddColumn(
                    RemoveEmptyColumns,
                    "Audit Score Comments_NEW",
                    each
                        if [Audit Score] = "Green" then "Pass"
                        else if [Audit Score] = "Amber" then "Caution"
                        else if [Audit Score] = "Red" then "Fail"
                        else [Audit Score Comments],
                    type text
                ),
                {"Audit Score Comments"}
            ),
            {{"Audit Score Comments_NEW", "Audit Score Comments"}}
        ),

    // 16. Reorder columns so Audit Score Comments appears after Audit Score
    ReorderColumns =
        let
            Cols = Table.ColumnNames(ReplaceAuditComments),
            Pos = List.PositionOf(Cols, "Audit Score"),
            ColsWithoutComments = List.RemoveItems(Cols, {"Audit Score Comments"}),
            NewOrder = List.InsertRange(ColsWithoutComments, Pos + 1, {"Audit Score Comments"})
        in
            Table.ReorderColumns(ReplaceAuditComments, NewOrder),

    // 17. Sort by Ref ascending
    FinalOutput = Table.Sort(ReorderColumns, {{"Ref", Order.Ascending}})
in
    FinalOutput
