Attribute VB_Name = "Module2"
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    Range("Table1[[#Headers],[Name]]").Select
    Application.WindowState = xlNormal
    ActiveWorkbook.Queries.Add Name:="Table1", Formula:= _
        "let" & Chr(13) & "" & Chr(10) & "    Source = Excel.CurrentWorkbook(){[Name=""Table1""]}[Content]," & Chr(13) & "" & Chr(10) & "    #""Changed Type"" = Table.TransformColumnTypes(Source,{{""Name"", type text}, {""Age"", Int64.Type}, {""Salary"", Int64.Type}})" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "    #""Changed Type"""
    Workbooks("Book1").Connections.Add2 "Query - Table1", _
        "Connection to the 'Table1' query in the workbook.", _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=Table1;Extended Properties=" _
        , """Table1""", 6, True, False
    Range("B7").Select
End Sub
