Attribute VB_Name = "HyperlinksVBA"
Option Explicit

Sub WorkingWithHyperlinks()

Dim path, Img_Path As String
Dim Hyp As Hyperlink

path = ThisWorkbook.path
Img_Path = "https://petco.sharepoint.com/sites/HR/LD/Resources/All%20About%20Pets%202018-%20Style%20Guide.pdf"

Application.DisplayAlerts = False

'Add a Hyperlink to the Worksheet.
Worksheets("Sheet1").Hyperlinks.Add Anchor:=Range("A5"), _
                     Address:="https://example.microsoft.com", _
                     ScreenTip:="Microsoft Web Site", _
                     TextToDisplay:="Microsoft"

Worksheets("Sheet1").Hyperlinks.Add Anchor:=Range("A6"), _
                     Address:=Img_Path, _
                     ScreenTip:="This is an image of a dog.", _
                     TextToDisplay:="Dog"
                     
Application.Wait Now() + #12:00:01 AM#
'Delete a Hyperlink to the Worksheet.
'Worksheets("Sheet1").Hyperlinks(1).Delete

'Follow a Hyperlink & open a new window.
'Worksheets("Sheet1").Hyperlinks(1).Follow NewWindow:=True

For Each Hyp In ActiveSheet.Hyperlinks
    MsgBox Hyp.Address
Next


'If the hyperlink is a document then save the file attached to link.
Worksheets("Sheet1").Hyperlinks(4).CreateNewDocument Filename:=path & "\Myfile.pdf", _
                                          EditNow:=False, _
                                          Overwrite:=True

Application.DisplayAlerts = True

'This will add the hyperlink to the system favorites folder.
'Worksheets("Sheet1").Hyperlinks(1).AddToFavorites

'With Worksheets("Sheet1").Hyperlinks(1)
'
'    .Address
'    .Application
'    .EmailSubject
'    .Name
'    .Parent
'    .Range
'    .ScreenTip
'    .Shape
'    .SubAddress
'    .TextToDisplay
'    .Type
'
'End With

End Sub
