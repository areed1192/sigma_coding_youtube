(resourceColumn as text , recordTable as table) =>

// https://docs.microsoft.com/en-us/powerquery-m/m-spec-sections

let

    // Grab the Column Names.
    TableColumnNames = Table.ColumnNames(recordTable),

    // Define the Colulmn Names we want to keep.
    TableColumnNamesWeWant = List.Select(TableColumnNames, each List.Contains({"Label", "URL_Link", "Modify_Datetime", "displayLabel", "link", "lastModifiedDatetime"}, _)),

    // Define the Function to aggregate by. Here we want all the rows, but only for the columns we specify.
    GroupFunction = {
        "table_records",
        each Table.RenameColumns(Table.SelectColumns(_, TableColumnNamesWeWant), {
                {"Label","Resource_Label"},
                {"URL_Link", "Resource_Link"},
                {"Modify_Datetime","Resource_Modify_Datetime"},
                {"displayLabel", "Folder_Label"},
                {"link", "Folder_Link"},
                {"lastModifiedDatetime", "Folder_Modify_Datetime"}
            }
        ), 
        type table
    },

    // Group the Tables.
    ResourceNames = Table.Group(recordTable, resourceColumn, GroupFunction),

    // Rename the Folder Name.
    RenamedTables = Table.RenameColumns(ResourceNames, {{"Folder_Name", "Resource_Name"}}),

    // Define the columns we want as records.
    RecordColumns = Table.AddColumn(
        RenamedTables, 
        "table_records_select", 
        each Table.SelectColumns([table_records], {"Resource_Label", "Resource_Link", "Resource_Modify_Datetime","Folder_Label", "Folder_Link", "Folder_Modify_Datetime"})
    ),

    // Convert the table_records column to a JSON string.
    ConvertToJSON = Table.AddColumn(RecordColumns, "JSON_Record", each Text.FromBinary(Json.FromValue(RecordColumns)), type text),

    // Keep the columns we want.
    KeepSpecificColumns = Table.SelectColumns(ConvertToJSON, {"Resource_Name", "JSON_Record"})

in

    KeepSpecificColumns

// Invoke it.
let

    // Convert the data sources to a JSON string.
    Source = ConvertToJson("Folder_Name", Data_Sources_With_Request)
    
in
    Source

    