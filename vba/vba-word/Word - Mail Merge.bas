Sub MailMergeTest()

'Define our Variables.
Dim wrdApp As Application
Dim wrdActiveDocument As Document
Dim wrdMailMerge As MailMerge

'Some Variables Specific to a Connection String.
Dim xlDataSourceFile As String
Dim xlDataSourceConnectionString As String

'Define the Application
Set wrdApp = Application

'Define the Word Document
Set wrdActiveDocument = wrdApp.ActiveDocument

'Lets Grab the Documents MailMerge Object.
Set wrdMailMerge = wrdActiveDocument.MailMerge
    
    'Step 1: Define What type of Mail Merge we are doing.
    '        In this case, we are going to create some invoices (Letters)
    '        using are excel data.
    wrdMailMerge.MainDocumentType = wdFormLetters
    
    'Step 2: Define the data source (Excel File).
    
    'Define the File Path of the Data Source.
    xlDataSourceFile = "C:\Users\Alex\OneDrive\Growth - Tutorial Videos\Lessons - VBA\VBA - Word\Working With Mail Merge.xlsm"
    
    'Define the Connection String for the Data Source.
    xlDataSourceConnection = "Provider=Microsoft.ACE.OLEDB.12.0;" _
                             & "User ID=Admin;" _
                             & "Data Source=" + xlDataSourceFile + ";" _
                             & "Mode=Read;" _
                             & "Extended Properties=""" _
                             & "HDR=YES;" _
                             & "IMEX=1;"";" _
                             & "Jet OLEDB:System database="""";" _
                             & "Jet OLEDB:Regist"

    'Step 3: Open the Data Source.
    wrdMailMerge.OpenDataSource Name:=xlDataSourceFile, _
                                LinkToSource:=True, _
                                Connection:=xlDataSourceConnection, _
                                SQLStatement:="SELECT * FROM `Invoices$`"
                                
End Sub