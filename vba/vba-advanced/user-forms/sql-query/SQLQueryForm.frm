VERSION 5.00

Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SQLQueryForm 
   Caption         =   "SQL Manager"
   ClientHeight    =   10560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9855.001
   OleObjectBlob   =   "SQLQueryForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End

Attribute VB_Name = "SQLQueryForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ExecuteButton_Click()

' This function does a lot but in essence, it grabs each of the TextBoxs in our form, and grabs their values.
' With those values, we have the info we need to create a connection string and with a connection string we can
' create a connection to a database using the "ADODB" library.
'
' If we can create a successful connection, we pass through a query and then take the results that are returned to us
' and dump them in our ListView object. Where the ListView object is data grid we see when we launch the form.

'Define variables related to Form
Dim QueryTextbox As MSForms.TextBox
Dim ServerTextbox As MSForms.TextBox
Dim DatabaseTextbox As MSForms.TextBox
Dim QueryResults As ListView

'Define variables related to Connection Object
Dim ConnObj As ADODB.Connection
Dim RecSet As ADODB.Recordset
Dim ConnCmd As ADODB.Command
Dim ColNames As ADODB.Fields
Dim Count As Integer

'Define some strings we will use
Dim DataSource As String
Dim DatabaseName As String
Dim ServerName As String
Dim SQLQuery As String

'Grab the Query Textbox
Set QueryTextbox = Me.Controls.Item("QueryTextbox")

'Grab the Server Textbox
Set ServerTextbox = Me.Controls.Item("ServerTextbox")

'Grab the Database Textbox
Set DatabaseTextbox = Me.Controls.Item("DatabaseTextbox")

'Grab the List View
Set QueryResults = Me.Controls.Item("QueryResults")

'Grab the Server name
ServerName = ServerTextbox.Value

'Grab the database name
DatabaseName = DatabaseTextbox.Value

'Grab the database query
SQLQuery = QueryTextbox.Value

'Define the connection string, this takes all the items in our textbox and joins them together
DataSource = "Server=" + ServerName + ";Database=" + DatabaseName + ";Trusted_connection=yes;"

'Create a new connection object & a new command object
Set ConnObj = New ADODB.Connection
Set ConnCmd = New ADODB.Command

'Create a new connection
With ConnObj
    .Provider = "MSOLEDBSQL"    'For SQL Database
    .ConnectionString = DataSource
    .Open
End With

'This will allow the command object to use the active connection
ConnCmd.ActiveConnection = ConnObj

'Define the query string (Comes from our textbox) & the command type
ConnCmd.CommandText = SQLQuery
ConnCmd.CommandType = adCmdText

'Exectue the query & get the column names (Fields)
Set RecSet = ConnCmd.Execute
Set ColNames = RecSet.Fields

'Write the results to query section
With QueryResults
    
    'Lets clear the old results
    .ColumnHeaders.Clear
    .ListItems.Clear

    'For each field name, add it to the list view as a column header
    For Each ColName In ColNames
        
        'Go to the column headers collection, and use the add method to add a new header
        .ColumnHeaders.Add Text:=ColName.Name
        
    Next
    
   ' If the recordset is empty, display a message
   If RecSet.EOF Then
        
        'Display the message
        MsgBox Prompt:="The query returned no results. No Data to display.", Buttons:=vbInformation, Title:="Query Status"
        
        'Close the connection
        ConnObj.Close
        
        'Exit the sub
        Exit Sub
        
   Else 'Otherwise begin populating the list view
        
        'Grab the recordset
        With RecSet
        
            'Keep going until you've reach the end of the recordset
            Do Until .EOF
            
            'Initalize a count, this will help to determine whether to add a new row vice a new column
            Count = 1
            
                'Loop through all the fields in the recordset
                For Each fld In .Fields
                
                    'If it's the first field of the recordset, that means we have the first column of a new row
                    If Count = 1 Then
                    
                        'If it's a new row, then we will add a new ListItems (ROW) object
                        Set ListItm = QueryResults.ListItems.Add(Text:=fld.Value)
                    
                    Else
                    
                        'If it's not a new row, then add a ListSubItem (ELEMENT) instead
                        ListItm.ListSubItems.Add Text:=fld.Value
                    
                    End If
                    
                    'Make sure to increment the count, or else EVERYONE will be a "New Row"
                    Count = Count + 1
                
                Next
            
            'Move to the next recordset
            .MoveNext
            
            Loop
        End With
        
        'When you're done with all the recordsets close the connection
        ConnObj.Close

    End If
    
End With

End Sub
