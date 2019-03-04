Attribute VB_Name = "Module3"
Sub CopyRangesToOutlook_multi()

    'Declare Outlook Variables
    Dim oLookApp As Outlook.Application
    Dim oLookItm As Outlook.MailItem
    Dim oLookIns As Outlook.Inspector
    
    'Declare Word Variables
    Dim oWrdDoc As Word.Document
    Dim oWrdRng As Word.Range
    
    'Declare Excel Variables
    Dim RngArray As Variant
    
    On Error Resume Next
    
    'Try And get an active instance of Outlook
    Set oLookApp = GetObject(, "Outlook.Application")
        
        'If error create a new instance
        If Err.Number = 429 Then
        
            'Create a new instance of Outlook
            Set oLookApp = New Outlook.Application
        
        End If
        
    'Create a new email
    Set oLookItm = oLookApp.CreateItem(olMailItem)
    
    'Create array to house all of our ranges
    RngArray = Array(Sheet1.Range("B2:C7"), Sheet2.Range("A1:B6"))
    
    With oLookItm
    
        'Fillout some basic info
        .To = "abc@xyz.com"
        .CC = "abc@xyz.com"
        .Subject = "Here are all my ranges"
        .Body = "Here are all the ranges from my workbook. Don't they look nice"
        
        'Display the email
        .Display
        
        'Get the active inspector
        Set oLookIns = .GetInspector
        
        'Get the Word Editor
        Set oWrdDoc = oLookIns.WordEditor
        
        'Loop through the Array
        For Each Item In RngArray
            
            'Copy the range
            Item.Copy
            
            'Define the range
            Set oWrdRng = oWrdDoc.Application.ActiveDocument.Content
                oWrdRng.Collapse Direction:=wdCollapseEnd
                
            'Insert a new paragraph
            Set oWrdRng = oWrdDoc.Paragraphs.Add
                oWrdRng.InsertBreak
                
            'Paste the Object
            oWrdRng.PasteSpecial DataType:=wdPasteMetafilePicture
        
        Next
    
    End With
    
End Sub


















