
// 
// 
// PART ONE: LOAD THE TABLE DATA.
// 
// 

let
    
    // Grab the StateTable.
    Source = Excel.CurrentWorkbook(){[Name="StateTable"]}[Content],

    // Change the Data Types of that Table.
    ChangeTableTypes = Table.TransformColumnTypes(
        Source,
        {
            {"Field", type text},
            {"Value", type text}
        }
    ),

    // Transpose the table.
    TransposeTheTable = Table.Transpose(ChangeTableTypes),

    // Promote the First Row to headers.
    PromoteTheFirstRow = Table.PromoteHeaders(
        TransposeTheTable,
        [PromoteAllScalars=true]
    ),

    // Change More Column Types.
    ChangeMoreColumnTypes = Table.TransformColumnTypes(
        PromoteTheFirstRow,
        {
            {"Access_Token", type text},
            {"Refresh_Token", type text}
        }
    ),

    // Grab the Table Records.
    TableRecords = Table.ToRecords(ChangeMoreColumnTypes){0},

    // Change to JSON Records, returns Binary Data.
    JSONRecords = Json.FromValue(TableRecords)

in

    // Load the Records.
    JSONRecords

// 
// 
// PART TWO: LOAD THE URL
// 
// 

let

    // Load the Content Table.
    Source = Excel.CurrentWorkbook(){[Name="Credentials"]}[Content],

    // Change the Data Types.
    ChangeColumnTypes = Table.TransformColumnTypes(
        Source,
        {
            {"Field", type text},
            {"Value", type text}
        }
    ),

    // Drill Down to the URL Value.
    URLValue = ChangeColumnTypes{0}[Value]

in

    // Load the URL Value.
    URLValue

// 
// 
// PART THREE: MAKE THE POST REQUEST.
// 
// 

let
    
    // Define the Source, in this case, we are grabbing content from the result of a POST request.
    Source = Json.Document(
        Web.Contents(
            Url_Credentials, 
            [
                Headers=[#"Content-Type"="application/json"], 
                Content= Credentials
            ]

        )
    ),

    // Take our Table and convert it to records.
    ConvertToRecords = Record.ToTable(Source),

    // Change the Data Types.
    ChangeContentType = Table.TransformColumnTypes(
        ConvertToRecords,
        {
            {"Value", type text}
        }
    )

in
    // Load the Data.
    ChangeContentType