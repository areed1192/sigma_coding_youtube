Attribute VB_Name = "Module2"
Sub ExportToOutlook()

'Declare Outlook Variables
 Dim oLookApp As Outlook.Application
 Dim oLookItm As Outlook.MailItem

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

        'Create the Outlook item
         With oLookItm

             'Pass through the necessary info
             .To = "Someone"
             .Subject = "Test"
             .Display

             'Get the word editor
              Set oLookInsp = .GetInspector
              Set oWdEditor = oLookInsp.WordEditor

             'Define the content area
              Set oWdContent = oWdEditor.Content
                  oWdContent.InsertParagraphBefore

             'Define the range where we want to paste.
              Set oWdRng = oWdEditor.Paragraphs(1).Range

    End With




End Sub
