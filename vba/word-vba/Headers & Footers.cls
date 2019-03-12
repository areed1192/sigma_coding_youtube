Sub WorkWithHeaderFooter()

'Declare variables
Dim myHeader As HeaderFooter
Dim myFooter As HeaderFooter

'Define the header or footer
Set myHeader = ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary)
Set myFooter = ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary)

'wdHeaderFooterPrimary
'wdHeaderFooterEvenPages
'wdHeaderFooterFirstPage

'Set the header text
myHeader.Range.Text = "This is my header"
myHeader.Range.Bold = True

'Set the header text
myFooter.Range.Text = "This is my footer"
myFooter.Range.Italic = True

'Add a page number and align it.
myFooter.PageNumbers.Add PageNumberAlignment:=wdAlignPageNumberRight

'Print the index of my footer
Debug.Print myFooter.Index

'Does myFooter Exist? True/False
Debug.Print myFooter.Exists

End Sub
