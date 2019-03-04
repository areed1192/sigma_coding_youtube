Attribute VB_Name = "Module2"
Sub RangeToOutlook_Single()

    'Declare Outlook Variables
    Dim oLookApp As Outlook.Application
    Dim oLookItm As Outlook.MailItem
    Dim oLookIns As Outlook.Inspector
    
    'Declare Word Variables
    Dim oWrdDoc As Word.Document
    Dim oWrdRng As Word.Range
    
    'Delcare Excel Variables
    Dim ExcRng As Range
    
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
    
          
    'Create an array to hold ranges
    ExcRng = Sheet1.Range("B2:C7")

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

        
        ExcRng.Copy
        
        'Define the range, insert a blank line, collapse the selection.
        Set oWrdRng = oWrdDoc.Application.ActiveDocument.Content
            oWrdRng.Collapse Direction:=wdCollapseEnd
            
        'Add a new paragragp and then a break
        Set oWrdRng = oWdEditor.Paragraphs.Add
            oWrdRng.InsertBreak
                    
        'Paste the object.
        oWrdRng.PasteSpecial DataType:=wdPasteMetafilePicture

    
    End With
        
        
End Sub


Sub RangeToOutlook_Multi()

    'Declare Outlook Variables
    Dim oLookApp As Outlook.Application
    Dim oLookItm As Outlook.MailItem
    Dim oLookIns As Outlook.Inspector
    
    'Declare Word Variables
    Dim oWrdDoc As Word.Document
    Dim oWrdRng As Word.Range
    
    'Delcare Excel Variables
    Dim RngArray As Variant
    
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
    
          
    'Create an array to hold ranges
    RngArray = Array(Sheet1.Range("B2:C7"), Sheet2.Range("A1:B6"))

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
        
      For Each Item In RngArray
        
            Item.Copy
            
            'Define the range, insert a blank line, collapse the selection.
            Set oWrdRng = oWrdDoc.Application.ActiveDocument.Content
                oWrdRng.Collapse Direction:=wdCollapseEnd
                
            'Add a new paragragp and then a break
            Set oWrdRng = oWdEditor.Paragraphs.Add
                oWrdRng.InsertBreak
                        
            'Paste the object.
            oWrdRng.PasteSpecial DataType:=wdPasteMetafilePicture
        
     Next
    
    End With        
        
End Sub
