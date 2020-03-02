Option Explicit

Sub ControlSlideShow()

Dim pptPres As Presentation
Dim pptSide As Slide

'The actual slide show components.
Dim pptSldShowSettings As SlideShowSettings
Dim pptSldShowWindow As SlideShowWindow
Dim pptSldShowNav As SlideNavigation
Dim pptSldShowView As SlideShowView

'Grab the active presentation
Set pptPres = ActivePresentation

'Grab the SlideShowSettings Object
Set pptSldShowSettings = pptPres.SlideShowSettings
    
    'Specify the starting slide
    pptSldShowSettings.StartingSlide = 1
    
    'Specify the ending slide
    pptSldShowSettings.EndingSlide = 1
    
    'Keep looping
    pptSldShowSettings.LoopUntilStopped = msoFalse
    
    'Change the pointer color
    pptSldShowSettings.PointerColor.RGB = RGB(0, 0, 255)
    
'Start the slide Show, this returns a SlideShowWindow Object
Set pptSldShowWindow = pptSldShowSettings.Run

    'Print out some properties about our Slide Show Window
    With pptSldShowWindow
    
        Debug.Print .Height
        Debug.Print .IsFullScreen
        Debug.Print .Width
        Debug.Print .Top
    
    End With
    
    'The SlideShowWindow Object has a SlideNavigation Object.
    Set pptSldShowNav = pptSldShowWindow.SlideNavigation
    
    'Let's make it visible
    pptSldShowNav.Visible = False

'If I grab the view property it returns a SlideShowView Object.
Set pptSldShowView = pptSldShowWindow.View
    
    'Enable the Laser Pointer
    pptSldShowView.LaserPointerEnabled = True
    
    'Set the Pointer to be a pen
    pptSldShowView.PointerType = ppSlideShowPointerPen
    
    'Go to the next slide
    pptSldShowView.Next
    
    'Go backwards
    pptSldShowView.Previous
    
    'Go to the first slide
    pptSldShowView.Last
    
    'Go to the last slide
    pptSldShowView.Next
    
    'Exit the slide Show
    pptSldShowView.Exit

End Sub
