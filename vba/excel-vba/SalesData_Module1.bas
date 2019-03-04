Attribute VB_Name = "Module1"
Sub ImportTextFilePQ()

'Declare my variables
Dim QueryName, SourceFormula, ConnStr As String


'Create a new query, and define the query formula.
QueryName = "SalesDataPull"
SourceFormula = "let Source = Csv.Document(File.Contents(""C:\Users\Alex\Desktop\SalesData.txt""), " & _
                                               "[Delimiter=""#(tab)"", " & _
                                               "Columns=5, " & _
                                               "Encoding=1200, " & _
                                               "QuoteStyle=QuoteStyle.None]), " & _
                "#""Promoted Headers"" = Table.PromoteHeaders(Source, [PromoteAllScalars=true]), " & _
                "#""Changed Type"" = Table.TransformColumnTypes(#""Promoted Headers"",{{""FY"", Int64.Type}, " & _
                "{""Per"", Int64.Type}, {""Jrnl Dt"", Int64.Type}, {""Amount"", type number}, {""Acct Descr"", type text}}) in #""Changed Type"""
                
'Add a new query to my workbook
ActiveWorkbook.Queries.Add Name:=QueryName, _
                           Formula:=SourceFormula, _
                           Description:="This pulls in my sales data from the text file."
                           
'Create connection string
ConnStr = "OLEDB;" & _
          "Provider=Microsoft.Mashup.OleDb.1;" & _
          "Data Source=$Workbook$;" & _
          "Location=""SalesDataPull"";" & _
          "Extended Properties="""""
          
'Add a table to my worksheet and pull the query in.
With ActiveSheet.ListObjects.Add(SourceType:=xlSrcExternal, _
                                 LinkSource:=True, _
                                 xlListObjectHasHeaders:=xlYes, _
                                 Source:=ConnStr, _
                                 TableStyleName:="TableStyleMedium8", _
                                 Destination:=Range("$A$1")).QueryTable
                 .CommandType = xlCmdSql
                 .CommandText = Array("SELECT * FROM [SalesDataPull]")
                 .Refresh BackgroundQuery = False
End With
                          

End Sub
