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

Sub WorkWithDataSource()

'Define our Variables.
Dim wrdApp As Application
Dim wrdActiveDocument As Document
Dim wrdMailMerge As MailMerge

Dim wrdMailMergeField As MailMergeDataField
Dim wrdMailMergeFields As MailMergeDataFields

Dim wrdMailMergeFieldName As MailMergeFieldName
Dim wrdMailMergeFieldNames As MailMergeFieldNames

'WARNING INTELLISENSE ERROR.
'Dim wrdMailMergeTest As MailMergeDataSource <<< THIS WILL NOT WORK

'Define the Application
Set wrdApp = Application

'Define the Word Document
Set wrdActiveDocument = wrdApp.ActiveDocument

'Lets Grab the Documents MailMerge Object.
Set wrdMailMerge = wrdActiveDocument.MailMerge

'Now I have a Data Source Already, so it's okay to grab the DataSource Property. MAKE SURE YOU HAVE A DATA SOURCE!
With wrdMailMerge.DataSource
    
    'Print Some Details.
    Debug.Print "Data Source Name Is: " + .Name
    Debug.Print "Data Source Connection String Is: " + .ConnectString
    Debug.Print "Data Source Query String Is: " + .QueryString
    Debug.Print "Data Source Table Name Is: " + .TableName
    Debug.Print "Data Source Record Count Is: " + CStr(.RecordCount)
    Debug.Print "Data Source Active Record Index Is: " + CStr(.ActiveRecord)
    Debug.Print "++++++++++++++++++++"
    
End With

'We can also find a record in our Data Source.
If wrdMailMerge.DataSource.FindRecord(FindText:="Reed", Field:="Last_Name") = True Then
    Debug.Print "was Found"
End If

'Loop through each Data Field.
For Each wrdMailMergeField In wrdMailMerge.DataSource.DataFields

    'Print Some Details.
    Debug.Print "Field Name: " + wrdMailMergeField.Name
    Debug.Print "Field Value: " + wrdMailMergeField.Value
    Debug.Print "Field Index: " + CStr(wrdMailMergeField.Index)
    Debug.Print "++++++++++++++++++++"

Next

'Loop through each Field Name.
For Each wrdMailMergeFieldName In wrdMailMerge.DataSource.FieldNames

    'Print Some Details.
    Debug.Print "Field Name: " + wrdMailMergeFieldName.Name
    Debug.Print "Field Index: " + CStr(wrdMailMergeFieldName.Index)
    Debug.Print "++++++++++++++++++++"

Next

End Sub