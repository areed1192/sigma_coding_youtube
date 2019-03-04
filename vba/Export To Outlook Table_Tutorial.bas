Attribute VB_Name = "Tutorial"
Sub TableToOutlook_Multi()

    'Declare Outlook Variables
    Dim oLookApp As Outlook.Application
    Dim oLookItm As Outlook.MailItem
    Dim oLookIns As Outlook.Inspector
    
    'Declare Word Variables
    Dim oWrdDoc As Word.Document
    Dim oWrdRng As Word.Range
    Dim oWrdTbl As Word.Table
    
    'Declare Excel Variables
    Dim ExcTbl As ListObject
    
    On Error Resume Next
    
    'Get the active instance of Outlook
    Set oLookApp = GetObject(, "Outlook.Application")
    
        'If error create a new instance
        If Err.Number = 429 Then
           Set oLookApp = New Outlook.Application
        End If
        
    'Create a new email
    Set oLookItm = oLookApp.CreateItem(olMailItem)
    
    With oLookItm
    
        'Define some basic info
        .To = "abc@xyz.com"
        .CC = "abc@xyz.com"
        .Subject = "My Excel Tables"
        .Body = "Here are my tables."
        
        'Display the email
        .Display
        
        'Get the inspector
        Set oLookIns = .GetInspector
        
        'Get the Word Editor
        Set oWrdDoc = oLookIns.WordEditor
        
        'Loop through tables on the ACTIVE SHEET
        For Each ExcTbl In ActiveSheet.ListObjects
        
            'Copy the table
            ExcTbl.Range.Copy
            
            'Define the range
            Set oWrdRng = oWrdDoc.Application.ActiveDocument.Content
                oWrdRng.Collapse Direction:=wdCollapseEnd
                
            'Add a new paragraph
            Set oWrdRng = oWrdDoc.Paragraphs.Add
                oWrdRng.InsertBreak
                
            'Paste the table
            oWrdRng.PasteExcelTable LinkedToExcel:=False, WordFormatting:=True, RTF:=True
        
        Next
        
    End With
End Sub



Sub TableToOutlook_Multi_Workbook()

    'Declare Outlook Variables
    Dim oLookApp As Outlook.Application
    Dim oLookItm As Outlook.MailItem
    Dim oLookIns As Outlook.Inspector
    
    'Declare Word Variables
    Dim oWrdDoc As Word.Document
    Dim oWrdRng As Word.Range
    Dim oWrdTbl As Word.Table
    
    'Declare Excel Variables
    Dim ExcTbl As ListObject
    Dim WrkSht As Worksheet
    
    On Error Resume Next
    
    'Get the active instance of Outlook
    Set oLookApp = GetObject(, "Outlook.Application")
    
        'If error create a new instance
        If Err.Number = 429 Then
           Set oLookApp = New Outlook.Application
        End If
        
    'Create a new email
    Set oLookItm = oLookApp.CreateItem(olMailItem)
    
    With oLookItm
    
        'Define some basic info
        .To = "abc@xyz.com"
        .CC = "abc@xyz.com"
        .Subject = "My Excel Tables"
        .Body = "Here are my tables."
        
        'Display the email
        .Display
        
        'Get the inspector
        Set oLookIns = .GetInspector
        
        'Get the Word Editor
        Set oWrdDoc = oLookIns.WordEditor
        
        'Loop through each worksheet in the ACTIVE WORKBOOK
        For Each WrkSht In ActiveWorkbook.Worksheets
        
            WrkSht.Activate
        
            'Loop through tables on the ACTIVE SHEET
            For Each ExcTbl In WrkSht.ListObjects
            
                'Copy the table
                ExcTbl.Range.Copy
                
                'Define the range
                Set oWrdRng = oWrdDoc.Application.ActiveDocument.Content
                    oWrdRng.Collapse Direction:=wdCollapseEnd
                    
                'Add a new paragraph
                Set oWrdRng = oWrdDoc.Paragraphs.Add
                    oWrdRng.InsertBreak
                    
                'Paste the table
                oWrdRng.PasteExcelTable LinkedToExcel:=False, WordFormatting:=True, RTF:=True
                
                'Create a refernce to the table
                Set oWrdTbl = oWrdDoc.Tables(oWrdDoc.Tables.Count)
                    oWrdTbl.AllowAutoFit = True
                    oWrdTbl.AutoFitBehavior (wdAutoFitWindow)
                    oWrdTbl.Style = wdStyleTableDarkList
                    oWrdTbl.BottomPadding = PixelsToPoints(10, True)
                    oWrdTbl.TopPadding = PixelsToPoints(10, True)
            
            Next
        Next
        
    End With
End Sub










