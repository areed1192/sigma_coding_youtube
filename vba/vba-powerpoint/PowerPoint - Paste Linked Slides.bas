Option Explicit

Sub PasteLinkedSlide()

Dim PPTPres As Presentation
Dim PPTSld As Slide
Dim PPTLnkSld As Slide
Dim PPTOLEShape As Shape
Dim PPTSec As SectionProperties

Dim SecIndex As Integer
Dim SecFirstSldIndex As Integer
Dim SecLastSldIndex As Integer
Dim SldIndex As Integer

'Grab the Active Presentation.
Set PPTPres = Application.ActivePresentation

'In my Presentation I have sections, so let's grab the SectionProperties Collection.
Set PPTSec = PPTPres.SectionProperties

'Define the Slide I'm Pasting to. In this case it's Slide 2.
Set PPTLnkSld = PPTPres.Slides.Item(Index:=2)
  
'Loop through each Section in the Presentation.
For SecIndex = 1 To PPTSec.Count
    
    'If the Section Name matches the one we are looking for, then continue.
    If PPTSec.Name(sectionIndex:=SecIndex) = "Presentation Slides" Then
        
        'Grab the Index of the First Slide in that Section.
        SecFirstSldIndex = PPTSec.FirstSlide(sectionIndex:=SecIndex)
        
        'Grab the Index of the Last Slide in that Section.
        SecLastSldIndex = SecFirstSldIndex + PPTSec.SlidesCount(sectionIndex:=SecIndex) - 1
        
        'Print out some info about that section.
        Debug.Print "Section Name: " + PPTSec.Name(sectionIndex:=SecIndex)
        Debug.Print "First Slide Index: " + CStr(SecFirstSldIndex)
        Debug.Print "Last Slide Index: " + CStr(SecLastSldIndex)
        
        'Loop through the Indexes.
        For SldIndex = SecFirstSldIndex To SecLastSldIndex
            
            'Set the Slide while in the Loop.
            Set PPTSld = PPTPres.Slides.Item(SldIndex)
                
                'Copy that Slide
                PPTSld.Copy
                
                'Paste the Slide to Linked Slide (Slide 2), this returns a ShapeRange which contains 1 shape.
                'Use the Item method to grab the 1 item in that range, so we have a Shape Object.
                Set PPTOLEShape = PPTLnkSld.Shapes.PasteSpecial(DataType:=ppPasteOLEObject, Link:=True).Item(1)
                    
                    'Take the Shape Object and Set the Height and Width.
                    PPTOLEShape.Height = 145
                    PPTOLEShape.Width = 259
            
        Next
    End If
Next

End Sub

Sub GetShapeDimensions()

'Declare our Variables.
Dim PPTSelec As Selection
Dim PPTShape As Shape

'Grab the Current Selection in the Active POWERPOINT Window.
Set PPTSelec = Application.ActiveWindow.Selection

'The Selection object has a ShapeRange Collection that we can use to grab the Item we want.
Set PPTShape = PPTSelec.ShapeRange.Item(1)
    
    'Print some details.
    Debug.Print "Shape Height is: " + CStr(Round(PPTShape.Height, 2))
    Debug.Print "Shape Width is: " + CStr(Round(PPTShape.Width, 2))

End Sub
