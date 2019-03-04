Attribute VB_Name = "Module3"
Sub ChartToOutlook_Multi()

    'Declare Outlook Variables
    Dim oLookApp As Outlook.Application
    Dim oLookItm As Outlook.MailItem
    Dim oLookIns As Outlook.Inspector
    
    'Declare Word Variables
    Dim oWrdDoc As Word.Document
    Dim oWrdRng As Word.Range
    
    'Delcare Excel Variables
    Dim ChrObj As ChartObject
    
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
        .Subject = "Here are all of my charts"
        .Body = "Here are all the charts from my worksheet."
        
        'Display the email
        .Display
        
        'Get the Active Inspector
        Set oLookIns = .GetInspector
        
        'Get the document within the inspector
        Set oWrdDoc = oLookIns.WordEditor
        
        'looping through each chart
        For Each ChrObj In ActiveSheet.ChartObjects
            
            'copy the chart
            ChrObj.Chart.ChartArea.Copy
            
            'Define the range we want to paste it in.
            Set oLookRng = oWrdDoc.Application.ActiveDocument.Content
                oLookRng.InsertAfter " " & vbNewLine
                oLookRng.Collapse Direction:=wdCollapseEnd
                
            'Paste the range
            oLookRng.Paste
        
        Next
    
    End With
        
        
End Sub



















