Sub FormatInputCell()

' NAME:
' --------
' FormatInputCell
'
' OVERVIEW:
' ---------
' Formats an input cell to meet financial modeling standards.

'Declare Variables.
Dim xlApp As Application

'Set the Application.
Set xlApp = Application

'Grab the selection.
With xlApp.Selection
    
    'Set the border style.
    .Borders.LineStyle = xlContinuous
    .Borders.ThemeColor = 1
    .Borders.TintAndShade = -0.349986266670736
    .Borders.Weight = xlThin
    
    'Set the fill color.
    .Interior.Color = 13434879
    
    'Set the font color.
    .Font.Color = -65536
    
End With

End Sub
