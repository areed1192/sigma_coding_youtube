let

    // Grab the Source.
    Source = Json.Document(File.Contents("C:\Users\Alex\OneDrive\Desktop\Sigma\Repo - API Client US Congress\united-states-congress-python-api\samples\responses\data_sources.jsonc")),

    // Convert it to a Table.
    ConvertToTable = Record.ToTable(Source),

    // Expand it.
    ExpandColumns = Table.ExpandListColumn(ConvertToTable, "Value"),

    // Expand the Records.
    ExpandRecords = Table.ExpandRecordColumn(
        ExpandColumns,
        "Value",
        {"formattedLastModifiedTime", "name", "folder", "displayLabel", "link", "justFileName", "fileExtension"}, {"formattedLastModifiedTime", "name.1", "folder", "displayLabel", "link", "justFileName", "fileExtension"}
    ),

    // Rename the columns
    RenameTable = Table.RenameColumns(
        ExpandRecords,
        {{"Name", "Category"}, {"name.1", "Folder_Name"}, {"folder", "Is_Folder_Flag"}, {"displayLabel", "Label"}, {"link", "URL_Link"}, {"justFileName", "File_Name"}, {"fileExtension", "File_Extension"}}
    ),

    // Parse the Datetime.
    ParseDatetime = Table.AddColumn(RenameTable, "Modify_Datetime", each DateTime.From([formattedLastModifiedTime]), type datetime),

    // Remove the old Column.
    Cleanup = Table.RemoveColumns(ParseDatetime,{"formattedLastModifiedTime", "File_Extension"}),

    // Change the types.
    TypeCleanUp = Table.TransformColumnTypes(
        Cleanup,
        {{"Category", type text}, {"Folder_Name", type text}, {"Label", type text}, {"URL_Link", type text}, {"File_Name", type text}, {"Is_Folder_Flag", type logical}}
    ),
    
    // Make a request for each URL.
    GrabTheRequests = Table.AddColumn(TypeCleanUp, "Request_Result", each GrabResource([URL_Link])),

    // Expand the Table Column.
    ExpandTableColumns = Table.ExpandTableColumn(
        GrabTheRequests,
        "Request_Result", 
        {"folderName", "formattedLastModifiedTime", "name", "folder", "displayLabel", "link", "justFileName", "lastModifiedDatetime"},
        {"folderName", "formattedLastModifiedTime", "name", "folder", "displayLabel", "link", "justFileName", "lastModifiedDatetime"}
    ),

    // Clean up and Change the Data Types.
    ChangeTypes = Table.TransformColumnTypes(
        ExpandTableColumns,
        {
            {"justFileName", type text}, {"link", type text}, {"displayLabel", type text}, {"name", type text}, {"folderName", type text}, {"lastModifiedDatetime", type datetime}, {"formattedLastModifiedTime", type datetime}, {"folder", type logical}
        }
    )
    
in
    ChangeTypes