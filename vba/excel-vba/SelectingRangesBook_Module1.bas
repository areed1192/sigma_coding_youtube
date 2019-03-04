Attribute VB_Name = "Module1"
Sub SelectingRanges_ObjectHierarchy()
    
    'The most explicit way to reference a range of cells.
    Application.Workbooks("SelectingRangesBook.xlsm").Worksheets("Sheet1").Range("A1").Select
    
    'Still very explict but can be confusing for sure because we need to keep in mind whether we have a personal macro workbook or not
    Application.Workbooks(2).Worksheets(1).Range("B1").Select
    
    'Less Explicit, ThisWorkbook is referring to the workbook that contains our subroutine
    Application.ThisWorkbook.Worksheets(1).Range("C1").Select
     
    'Less Explicit, ActiveWorkbook is referring to the workbook that is open and we see.
    Application.ActiveWorkbook.Worksheets(1).Range("B1").Select
    

    'We can even be less explicit in our code, but then VBA will fill in the blanks. Which is a good and bad.
    
    
    'Forget the application
    Workbooks("SelectingRangesBook.xlsm").Worksheets("Sheet1").Range("A2").Select '<<< This is how it looks in VBA's eyes Application.Workbooks("SelectingRangesBook.xlsm").Worksheets("Sheet1").Range("B1").Select
    
    'Forget the workbook
    Worksheets("Sheet1").Range("A2").Select '<<< This is how it looks in VBA's eyes Application.ActiveWorkbook.Worksheets("Sheet1").Range("B1").Select
    
    'Forget the worksheet
    'Range("A2").Select   '<<< This is how it looks in VBA's eyes Application.ActiveWorkbook.ActiveWorksheet.Range("B1").Select
    
    
    'If you use the Select method to select cells, be aware that Select works only on the active worksheet.
    
End Sub


Sub SelectingWorkbooks()

    ' The application object has a collection that houses all of our workbooks, its called "WORKBOOKS"
    ' To access this collection, we call the Application and then the property Workbooks.
    ' This collection does not contain add-in workbooks (.xla) and workbooks in protected view are not a member of this collection.
    
    ' --- "Application.Workbooks"
    
    ' If we don't use an object qualifier for this property it's equivalent to using "Application.Workbooks"
    
    ' --- "Workbooks" is the same as "Application.Workbooks" we just are dropping the "Application" object.
    
    
    ' SELECTING INDIVIDUAL WORKBOOKS FROM THE WORKBOOK COLLECTIONS

    ' The Explict way - Using the KEY METHOD
      Application.Workbooks("SelectingWorkbooks.xlsx").Activate
      Application.Workbooks("SelectingWorkbooks2.xlsx").Activate
    '             Workbooks("SelectingWorkbooks.xlsx").Activate  <<< Also Works
    '             Workbooks("SelectingWorkbooks2.xlsx").Activate <<< Also Works
    
    ' The Less Explict way - Using the INDEX METHOD.
    ' NOTE ONE: Index is determined by the order in which my workbooks were open.
    ' NOTE TWO: If you have a PERSONAL.XSLB then this is ALSO a workbook in your collection and it is ALWAYS OPENED FIRST
    
      Application.Workbooks(1).Activate '<<< This is my Personal Macro Workbook and this workbook is ALWAYS OPENED FIRST
      Application.Workbooks(2).Activate '<<< This is my workbook that was opened SECOND and it is named "SelectingWorkbooks.xlsx"
      Application.Workbooks(3).Activate '<<< This is my workbook that was opened THIRD and it is named "SelectingWorkbooks2.xlsx"
    '             Workbooks(1).Activate '<<< Also Works
    '             Workbooks(2).Activate '<<< Also Works
    '             Workbooks(3).Activate '<<< Also Works

    
    ' Using the ActiveWorkbook Method
    ' NOTE ONE: The ActiveWorkbook is the one that is opened and we can see on our screen.
      ActiveWorkbook.Worksheets("Sheet1").Range("A1").Value = 7000
    
    ' Using the ThisWorkbook Method
    ' NOTE ONE: The ThisWorkbook is the one that houses our code.
      ThisWorkbook.Worksheets("Sheet1").Range("A1").Value = 7000
    

End Sub

