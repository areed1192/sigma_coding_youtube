Option Explicit

Sub WorkingWithMediaFormatObjects()

Dim PPTPres As Presentation
Dim PPTSld As Slide
Dim PPTShape As Shape
Dim PPTMediaFormat As MediaFormat

Set PPTPres = Application.ActivePresentation
Set PPTSld = PPTPres.Slides(1)

Set PPTShape = PPTSld.Shapes.Item(1)
    'PPTShape.LinkFormat.Update

'Set PPTMediaFormat = PPTShape.MediaFormat
'
'    Debug.Print PPTMediaFormat.IsEmbedded
'    Debug.Print PPTMediaFormat.VideoFrameRate
'    Debug.Print PPTMediaFormat.EndPoint
'    Debug.Print PPTMediaFormat.ResamplingStatus
'    Debug.Print PPTMediaFormat.AudioCompressionType
'    Debug.Print PPTMediaFormat.IsLinked

End Sub

Sub AddVideoFormatObjects()

Dim PPTPres As Presentation
Dim PPTSld As Slide
Dim PPTShape As Shape
Dim PPTMediaFormat As MediaFormat

'Grab the presentation
Set PPTPres = Application.ActivePresentation

'Grab a slide.
Set PPTSld = PPTPres.Slides(3)

'Insert the Media Format Object
Set PPTShape = PPTSld.Shapes.AddMediaObject2(FileName:="C:\Users\Alex\OneDrive\Growth - Tutorial Videos\Videos - Final\classes_pt_3.mp4", _
                                             LinkToFile:=True, _
                                             SaveWithDocument:=False, _
                                             Left:=200, _
                                             Top:=200, _
                                             Width:=200, _
                                             Height:=200)
'Grab the Object
Set PPTMediaFormat = PPTShape.MediaFormat

End Sub

Sub WorkingWithMediaFormatObject()

Dim PPTPres As Presentation
Dim PPTSld As Slide
Dim PPTShape As Shape
Dim PPTMediaFormat As MediaFormat

'Grab the presentation.
Set PPTPres = Application.ActivePresentation

'Grab the Slide.
Set PPTSld = PPTPres.Slides(1)

'Grab the Shape.
Set PPTShape = PPTSld.Shapes.Item(2)

    ' Name                Value   Description
    ' ----                -----   ------
    ' ppMediaTypeMixed    -2      Mixed
    ' ppMediaTypeMovie     3      Movie
    ' ppMediaTypeOther     1      Others
    ' ppMediaTypeSound     2      Sound
    
    'Print Media Type
    Debug.Print "Media Type is: " + CStr(PPTShape.MediaType)
  
'Grab the Shape MediaFormat Object.
Set PPTMediaFormat = PPTShape.MediaFormat

    'Media Boolean Properties.
    Debug.Print "Is the Media Embedded? " + CStr(PPTMediaFormat.IsEmbedded)
    Debug.Print "Is the Media Linked? " + CStr(PPTMediaFormat.IsLinked)
    
    'Media Properties.
    Debug.Print "Video Frame Rate: " + CStr(PPTMediaFormat.VideoFrameRate)
    Debug.Print "Video Endpoint: " + CStr(PPTMediaFormat.EndPoint)
    Debug.Print "Video Resampling Status: " + CStr(PPTMediaFormat.ResamplingStatus)
    Debug.Print "Audio Compression Rate: " + CStr(PPTMediaFormat.AudioCompressionType)
    
    'Change the start point
    PPTMediaFormat.StartPoint = 3000
    
    'Change the End Point
    PPTMediaFormat.EndPoint = 60000
    
End Sub

Sub WorkingWithMediaFormatLinkObject()

Dim PPTPres As Presentation
Dim PPTSld As Slide
Dim PPTShape As Shape
Dim PPTLinkFormat As LinkFormat

'Grab the presentation.
Set PPTPres = Application.ActivePresentation

'Grab the Slide.
Set PPTSld = PPTPres.Slides(2)

'Grab the Shape.
Set PPTShape = PPTSld.Shapes.Item(1)

    ' Name                Value   Description
    ' ----                -----   ------
    ' ppMediaTypeMixed    -2      Mixed
    ' ppMediaTypeMovie     3      Movie
    ' ppMediaTypeOther     1      Others
    ' ppMediaTypeSound     2      Sound
    
    'Print Media Type
    Debug.Print "Media Type is: " + CStr(PPTShape.MediaType)
  
'Grab the Shape MediaFormat Object.
Set PPTLinkFormat = PPTShape.LinkFormat
    
    'Grab the Source Full Name.
    Debug.Print "Link Source: " + PPTLinkFormat.SourceFullName
    
End Sub

