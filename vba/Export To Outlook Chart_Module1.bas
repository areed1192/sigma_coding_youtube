Attribute VB_Name = "Module1"
Sub ChartToOutlook_Multi()

    'Declare Outlook Variables
    Dim oLookApp As Outlook.Application
    Dim oLookFdr As Outlook.Folder
    Dim oLookNsp As Outlook.Namespace
    Dim oLookItm As Outlook.MailItem
    
    'Declare Excel Variables
    Dim ChrObj As ChartObject

    On Error Resume Next
    
    'Test if Outlook is Open
    Set oLookApp = GetObject(, "Outlook.Application")

        'If the Application isn't open it will return a 429 error
        If Err.Number = 429 Then

          'If it is not open then clear the error and create a new instance of Outlook
           Err.Clear
           Set oLookApp = New Outlook.Application

        End If
           
    'Create a mail item in outlook.
    Set oLookItm = oLookApp.CreateItem(olMailItem)

    'With the new email we just created.
    With oLookItm
        
        'Define basic infromation about the email
        .To = "xyz@anc.com"
        .CC = "abc@xyz.com"
        .Subject = "Test"
        .Body = "Dear Mr Lee" & vbNewLine
        
        'Show the new email.
        .Display
        
        'Get the Word Editor
        Set oWdEditor = .GetInspector.WordEditor
        
        'Loop through each chart in the active sheet
        For Each ChrObj In ActiveSheet.ChartObjects
            
            'Copy the Chart
            ChrObj.Chart.ChartArea.Copy
            
            'Define the range, insert a blank line, collapse the selection.
            Set oWdRng = oWdEditor.Application.ActiveDocument.Content
                oWdRng.InsertAfter " " & vbNewLine
                oWdRng.Collapse Direction:=wdCollapseEnd
                        
            'Paste the object.
            oWdRng.Paste
             
        Next
 
    End With

End Sub

Sub ChartToOutlook_single()

    'Declare Outlook Variables
    Dim oLookApp As Outlook.Application
    Dim oLookFdr As Outlook.Folder
    Dim oLookNsp As Outlook.Namespace
    Dim oLookItm As Outlook.MailItem
    
    
    'Declare Excel Variables
    Dim ChrObj As ChartObject

    On Error Resume Next
    
    'Test if Outlook is Open
    Set oLookApp = GetObject(, "Outlook.Application")

        'If the Application isn't open it will return a 429 error
        If Err.Number = 429 Then
        
          'If it is not open then clear the error and create a new instance of Outlook
           Err.Clear
           Set oLookApp = New Outlook.Application
           
'           'Create a NameSpace
'           Set oLookNsp = oLookApp.GetNamespace("MAPI")
'
'           'Create an Outlook Session and get the default folder.
'           Set oLookFdr = oLookApp.Session.GetDefaultFolder(olFolderInbox)
'               oLookFdr.Display
               
        End If
        
    'Create a reference to the chart and copy it.
    Set ChrObj = ActiveSheet.ChartObjects(1)
        ChrObj.Chart.ChartArea.Copy
           
    'Create a mail item in outlook.
    Set oLookItm = oLookApp.CreateItem(olMailItem)

    'With the new email we just created.
    With oLookItm
        
        'Define basic infromation about the email
        .To = "xyz@anc.com"
        .CC = "abc@xyz.com"
        .Subject = "Test"
        .Body = "Dear Mr Lee" & vbNewLine
        
        'Show the new email.
        .Display
        
        'Get the Word Editor
        Set oWdEditor = .GetInspector.WordEditor
            
        'Define the range, insert a blank line, collapse the selection.
        Set oWdRng = oWdEditor.Application.ActiveDocument.Content
            oWdRng.InsertAfter " " & vbNewLine
            oWdRng.Collapse Direction:=wdCollapseEnd
                    
        'Paste the object.
        oWdRng.Paste
 
    End With

End Sub


