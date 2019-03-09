Sub WorkWithSlide()

Dim PPTSlide As Slide

'Get the slide & print the name.
Set PPTSlide = ActivePresentation.Slides(2)
    Debug.Print PPTSlide.Name

'Add a tag to our slide
PPTSlide.Tags.Add Name:="Tag1", Value:="MyNewSlide"

'Add a comment to our slide.
PPTSlide.Comments.Add Left:=100, Top:=100, Author:="Alex", AuthorInitials:="AR", Text:="This slide looks good."

'Print the design of our slide, this should be the default Office theme
Debug.Print PPTSlide.Design.Name

'Lets add a date time to our footer
PPTSlide.HeadersFooters.DateAndTime.Visible = True

'Lets also add a slide number.
PPTSlide.HeadersFooters.SlideNumber.Visible = True

'Export the slide. VERY UNSTABLE AND SEEMS TO FAIL A LOT
PPTSlide.Export FileName:="C:\Users\Alex\Desktop\MySlide", FilterName:="PNG", ScaleWidth:=100, ScaleHeight:=100

End Sub
