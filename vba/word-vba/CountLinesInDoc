Sub CountLinsInDoc()

'Declare our Range Object
Dim LinCountDoc, LinCountPara, LinCountSec As Integer
Dim CountParas, CountPages, CountChar, CountWords As Integer

'Count the lines in the Active Document
LinCountDoc = ActiveDocument.BuiltInDocumentProperties(wdPropertyLines)

'Count the lines in a paragraph
LinCountPara = Paragraphs(1).Range.ComputeStatistics(wdStatisticLines)

'Count the lines on the page you have selected.
LinCountSec = ActiveDocument.Bookmarks("\page").Range.ComputeStatistics(wdStatisticLines)

'Count the number of pages in a document
CountPages = ActiveDocument.BuiltInDocumentProperties(wdPropertyPages)

'Count the number of paragraphs in a document
CountParas = ActiveDocument.BuiltInDocumentProperties(wdPropertyParas)

'Count the number of characters in a document
CountChar = ActiveDocument.BuiltInDocumentProperties(wdPropertyCharacters)

'Count the number of words in a document
CountWords = ActiveDocument.BuiltInDocumentProperties(wdPropertyWords)

End Sub
