Sub ChangeSlideSize()

With ActivePresentation.PageSetup
    
    'Set to 16|9
    .SlideSize = ppSlideSizeCustom
    .SlideWidth = 10 * 72
    .SlideHeight = 5.625 * 72

    'Set to 4|3
    .SlideSize = ppSlideSizeCustom
    .SlideWidth = 10 * 72
    .SlideHeight = 7.5 * 72
    
End With

End Sub
