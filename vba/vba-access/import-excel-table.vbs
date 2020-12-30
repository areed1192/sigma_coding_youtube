Option Compare Database

Sub ImportDataFromExcel()

'Declare my access variables.
Dim AccessApp As Application
Dim AccessDatabase As Database
Dim AccessTable As TableDef
Dim AccessImportTable As TableDef
Dim AccessTableField As Field
Dim AccessRecord As Recordset

'Declare my Excel Variables.
Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim xlDataTable As Excel.ListObject
Dim xlDataColumn As Excel.ListColumn
Dim xlDataRow As Excel.ListRow

Dim wasFound As Boolean
Dim fieldWasFound As Boolean

'Grab the Access Application.
Set AccessApp = Application

'Grab the Database inside of my application.
Set AccessDatabase = AccessApp.CurrentDb

'Set the Excel Application, I AM ASSUMING THE APPLICATION IS OPEN!!!
Set xlApp = GetObject(, "Excel.Application")

'Grab the workbook.
Set xlBook = xlApp.ActiveWorkbook

'Grab the Database Worksheet.
Set xlSheet = xlBook.Worksheets("Database")

'Grab the Snake Database Object.
Set xlDataTable = xlSheet.ListObjects("SnakeDatabase")

For Each AccessTable In AccessDatabase.TableDefs

    'Check if the name already exists.
    If AccessTable.Name = xlDataTable.Name Then
    
        'If we found it, then grab the table definition.
        Set AccessImportTable = AccessDatabase.TableDefs(xlDataTable.Name)
        
        Debug.Print "Table " + AccessImportTable.Name + " was found."
        
        'Set the flag.
        wasCreated = False
        
        'Exit the loop
        Exit For
        
    End If
Next

'If we didn't find the table then create it.
If AccessImportTable Is Nothing Then

    'Create a new table definition.
    Set AccessImportTable = AccessDatabase.CreateTableDef(Name:=xlDataTable.Name)
    Debug.Print "New Table was created, the table name is: " + AccessImportTable.Name
    wasCreated = True
    
    'Loop through each of the columns in my excel table.
    For Each xlDataColumn In xlDataTable.ListColumns
    
        'Print out some info for the user.
        Debug.Print "Column Field " + xlDataColumn.Name + " is going to be created."
        Debug.Print "Column Field Data Type will be " + CStr(CellType(xlDataColumn.DataBodyRange.Item(1)))
        
        'Take the Acces Table.
        With AccessImportTable
        
            'Add the Field to it.
            .Fields.Append .CreateField(Name:=xlDataColumn.Name, Type:=CellType(xlDataColumn.DataBodyRange.Item(1)))
        
        End With
    Next
    
    'Add to the Table Definitions Collection.
    AccessDatabase.TableDefs.Append Object:=AccessImportTable
    
    'Refresh the Window.
    AccessApp.RefreshDatabaseWindow

End If

'Once we add the table we will open up a new recordset to add the data.
Set AccessRecord = AccessDatabase.OpenRecordset(Name:=xlDataTable.Name)

'Loop through each row in my Excel Datatable.
For Each xlDataRow In xlDataTable.ListRows

    'Define Item count.
    itemCount = 1
    
    'Use the `AddNew` method to add a new record.
    AccessRecord.AddNew
    
    'Loop through the Header Row.
    For Each itemHeader In xlDataTable.HeaderRowRange
        
        'Grab the Field Name.
        FieldName = itemHeader.Value
        
        'Grab the Field Value.
        FieldValue = xlDataRow.Range.Item(itemCount).Value
        
        If FieldValue = "" Then
        
            'Specify that the value is null.
            AccessRecord(FieldName).Value = "null"
            
        Else
        
            'If it's not null add it to the field record.
            AccessRecord(FieldName).Value = FieldValue
            
        End If
        
        itemCount = itemCount + 1
    
    Next
    
    'Update the Recordset.
    AccessRecord.Update

Next

End Sub

Function CellType(pRange As Excel.Range)

Select Case True

    Case VBA.IsEmpty(pRange): CellType = vbNull
    Case Excel.Application.IsText(pRange): CellType = dbText
    Case Excel.Application.IsLogical(pRange): CellType = dbBoolean
    Case Excel.Application.IsErr(pRange): CellType = vbNull
    Case VBA.IsDate(pRange): CellType = dbDate
    Case VBA.InStr(1, pRange.Text, ":") <> 0: CellType = dbTime
    Case VBA.IsNumeric(pRange): CellType = dbNumeric
    
End Select

End Function
