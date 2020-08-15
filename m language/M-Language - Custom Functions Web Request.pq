(https_url) =>

let
    // Grab the URL
    Source = Json.Document(
        Web.Contents(
            https_url,
            [Headers=[Accept="text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9", #"Accept-Language"="en-US,en;q=0.9"]]
        )
    ),

    // Convert it to a Table.
    ConvertToTable = Record.ToTable(Source),

    // Expand the List.
    ExpandColumns = Table.ExpandListColumn(ConvertToTable, "Value"),

    // Grab the Record Field Names.
    RecordFieldNames = Record.FieldNames(ExpandColumns{0}[Value]),

    // Expand the Record Column.
    ExpandedTable = Table.ExpandRecordColumn(ExpandColumns, "Value", RecordFieldNames),

    // Grab the First Column Name.
    FirstColumnName = Table.ColumnNames(ExpandedTable){0},
    
    // Rename the Table.
    RenamedTable = Table.RenameColumns(ExpandedTable, {FirstColumnName, "folderName"}),

    // Parse the formattedLastModifiedTime column.
    TableWithDateTime = Table.AddColumn(RenamedTable, "lastModifiedDatetime", each DateTime.From([formattedLastModifiedTime]), type datetime)

in
    TableWithDateTime

// Invoke it.

let

    // Call the function we just built to pull the data.
    Source = GrabResource("https://www.govinfo.gov/bulkdata/json/CFR/2008"),

    // Sort the data on the Modify Datetime.
    SortedData = Table.Sort(Source,{{"lastModifiedDatetime", Order.Ascending}})

in
    SortedData