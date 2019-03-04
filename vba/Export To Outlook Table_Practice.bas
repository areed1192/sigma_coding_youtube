Attribute VB_Name = "Practice"
Sub TableToOutlook_Single()

    'Declare Outlook Variables
    Dim oLookApp As Outlook.Application
    Dim oLookItm As Outlook.MailItem
    Dim oLookIns As Outlook.Inspector
    
    'Declare Word Variables
    Dim oWrdDoc As Word.Document
    Dim oWrdRng As Word.Range
    Dim oWrdTbl As Word.Table
    
    'Delcare Excel Variables
    Dim ExcTbl As ListObject
    
    On Error Resume Next
    
    'Get the Active instance of Outlook if there is one
    Set oLookApp = GetObject(, "Outlook.Application")
    
        'If Outlook isn't open then create a new instance of Outlook
        If Err.Number = 429 Then
        
            'Clear Error
            Err.Clear
        
            'Create a new instance of Outlook
            Set oLookApp = New Outlook.Application
            
        End If
        
    'Create a new email
    Set oLookItm = oLookApp.CreateItem(olMailItem)
          
    'Create a reference to the Excel Table
    Set ExcTbl = Sheet1.ListObjects(1)

    With oLookItm
    
        'Define some basic info of our email
        .To = "xyz@abc.com"
        .CC = "xyz@abc.com"
        .Subject = "Here are all of my Ranges"
        .Body = "Here are all the Ranges from my worksheet."
        
        'Display the email
        .Display
        
        'Get the Active Inspector
        Set oLookIns = .GetInspector
        
        'Get the document within the inspector
        Set oWrdDoc = oLookIns.WordEditor

        'Copy the table
        ExcTbl.Range.Copy
        
        'Define the range, insert a blank line, collapse the selection.
        Set oWrdRng = oWrdDoc.Application.ActiveDocument.Content
            oWrdRng.Collapse Direction:=wdCollapseEnd
            
        'Add a new paragragp and then a break
        Set oWrdRng = oWdEditor.Paragraphs.Add
            oWrdRng.InsertBreak
                    
        'Paste the object.
        oWrdRng.PasteExcelTable LinkedToExcel:=True, WordFormatting:=True, RTF:=True
        
        'Create a reference to the Word Table
        Set oWrdTbl = oWrdDoc.Tables(oWrdDoc.Tables.Count)
        
            'Make sure it fits to the email length
            oWrdTbl.AllowAutoFit = True
            oWrdTbl.AutoFitBehavior (wdAutoFitWindow)
    
    End With
        
        
End Sub


Sub TableToOutlook_Multi_Sheet()

    'Declare Outlook Variables
    Dim oLookApp As Outlook.Application
    Dim oLookItm As Outlook.MailItem
    Dim oLookIns As Outlook.Inspector
    
    'Declare Word Variables
    Dim oWrdDoc As Word.Document
    Dim oWrdRng As Word.Range
    Dim oWrdTbl As Word.Table
    
    'Delcare Excel Variables
    Dim ExcTbl As ListObject
    Dim WrkSht As Worksheet
    
    On Error Resume Next
    
    'Get the Active instance of Outlook if there is one
    Set oLookApp = GetObject(, "Outlook.Application")
    
        'If Outlook isn't open then create a new instance of Outlook
        If Err.Number = 429 Then
        
            'Clear Error
            Err.Clear
        
            'Create a new instance of Outlook
            Set oLookApp = New Outlook.Application
            
        End If
        
    'Create a new email
    Set oLookItm = oLookApp.CreateItem(olMailItem)

    With oLookItm
    
        'Define some basic info of our email
        .To = "xyz@abc.com"
        .CC = "xyz@abc.com"
        .Subject = "Here are all of my Ranges"
        .Body = "Here are all the Ranges from my worksheet."
        
        'Display the email
        .Display
        
        'Get the Active Inspector
        Set oLookIns = .GetInspector
        
        'Get the document within the inspector
        Set oWrdDoc = oLookIns.WordEditor
        
        For Each ExcTbl In ActiveSheet.ListObjects
        
            'Copy the table
            ExcTbl.Range.Copy
            
            'Define the range, insert a blank line, collapse the selection.
            Set oWrdRng = oWrdDoc.Application.ActiveDocument.Content
                oWrdRng.Collapse Direction:=wdCollapseEnd
                
            'Add a new paragragp and then a break
            Set oWrdRng = oWdEditor.Paragraphs.Add
                oWrdRng.InsertBreak
                        
            'Paste the object.
            oWrdRng.PasteExcelTable LinkedToExcel:=True, WordFormatting:=True, RTF:=True
            
            'Create a reference to the Word Table
            Set oWrdTbl = oWrdDoc.Tables(oWrdDoc.Tables.Count)
        
                'Make sure it fits to the email length
                oWrdTbl.AllowAutoFit = True
                oWrdTbl.AutoFitBehavior (wdAutoFitWindow)
        
        Next
    
    End With
        
End Sub
Sub TableToOutlook_Multi_Book()

    'Declare Outlook Variables
    Dim oLookApp As Outlook.Application
    Dim oLookItm As Outlook.MailItem
    Dim oLookIns As Outlook.Inspector
    
    'Declare Word Variables
    Dim oWrdDoc As Word.Document
    Dim oWrdRng As Word.Range
    Dim oWrdTbl As Word.Table
    
    'Delcare Excel Variables
    Dim ExcTbl As ListObject
    Dim WrkSht As Worksheet
    
    On Error Resume Next
    
    'Get the Active instance of Outlook if there is one
    Set oLookApp = GetObject(, "Outlook.Application")
    
        'If Outlook isn't open then create a new instance of Outlook
        If Err.Number = 429 Then
        
            'Clear Error
            Err.Clear
        
            'Create a new instance of Outlook
            Set oLookApp = New Outlook.Application
            
        End If
        
    'Create a new email
    Set oLookItm = oLookApp.CreateItem(olMailItem)

    With oLookItm
    
        'Define some basic info of our email
        .To = "xyz@abc.com"
        .CC = "xyz@abc.com"
        .Subject = "Here are all of my Ranges"
        .Body = "Here are all the Ranges from my worksheet."
        
        'Display the email
        .Display
        
        'Get the Active Inspector
        Set oLookIns = .GetInspector
        
        'Get the document within the inspector
        Set oWrdDoc = oLookIns.WordEditor
        
        For Each WrkSht In ActiveWorkbook.Worksheets
            For Each ExcTbl In WrkSht.ListObjects
            
                'Copy the table
                ExcTbl.Range.Copy
                
                'Define the range, insert a blank line, collapse the selection.
                Set oWrdRng = oWrdDoc.Application.ActiveDocument.Content
                    oWrdRng.Collapse Direction:=wdCollapseEnd
                    
                'Add a new paragragp and then a break
                Set oWrdRng = oWdEditor.Paragraphs.Add
                    oWrdRng.InsertBreak
                            
                'Paste the object.
                oWrdRng.PasteExcelTable LinkedToExcel:=True, WordFormatting:=True, RTF:=True
                
                'Create a reference to the Word Table
                Set oWrdTbl = oWrdDoc.Tables(oWrdDoc.Tables.Count)
            
                    'Make sure it fits to the email length
                    oWrdTbl.AllowAutoFit = True
                    oWrdTbl.AutoFitBehavior (wdAutoFitWindow)
                    oWrdTbl.Style = wdStyleTableDarkList
                    oWrdTbl.BottomPadding = PixelsToPoints(10, True)
                    oWrdTbl.TopPadding = PixelsToPoints(10, True)
            Next
        Next
    
    End With
        
End Sub

