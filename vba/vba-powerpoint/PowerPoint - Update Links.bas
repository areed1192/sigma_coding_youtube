Sub UpdateLinkNonOneDrive()

    'Declare PowerPoint Variables
    Dim PPTSlide As Slide
    Dim PPTShape As Shape
    Dim SourceFile, FileName As String
    Dim Position As Integer

    'Declare Excel Variables
    Dim xlApp As Excel.Application
    Dim xlWorkBook As Excel.Workbook

    'Create a new Excel Application, make it invisible, set the Excel Display alerts to False.
    Set xlApp = New Excel.Application
        xlApp.Visible = False
        xlApp.DisplayAlerts = False

    'Loop through each slide in the Presentation.
    For Each PPTSlide In ActivePresentation.Slides

        'Loop through Each Shape in the slide
        For Each PPTShape In PPTSlide.Shapes

            'If the Shape is a linked OLEObject.
            If PPTShape.Type = msoLinkedOLEObject Then

                'Get the Source File of the shape.
                SourceFile = PPTShape.LinkFormat.SourceFullName

                'We may need to parse the Source file because if it's linked to a chart, for example, we can get the following:
                'C:\Users\NAME\ExcelBook.xlsx!Chart_One!
                'We want it to look like the following:
                'C:\Users\NAME\ExcelBook.xlsx

                'This will parse the source file so that it only includes the file name.
                Position = InStr(1, SourceFile, "!", vbTextCompare)
                FileName = Left(SourceFile, Position - 1)

                'This will open the file as read-only, and will not update the links in the Excel file.
                Set xlWorkBook = xlApp.Workbooks.Open(FileName, False, True)

                    'Update the link
                    PPTShape.LinkFormat.Update

                'Close the workbook and release it from memory.
                xlWorkBook.Close
                Set xlWorkBook = Nothing

            End If

        Next PPTShape
    Next PPTSlide

    'Close the Excel App & release it from memory
    xlApp.Quit
    Set xlApp = Nothing

End Sub
