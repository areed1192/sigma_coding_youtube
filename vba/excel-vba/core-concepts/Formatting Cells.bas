Attribute VB_Name = "Module1"
Option Explicit

Sub Formats()

'Change the Number Format of a Cell
Range("A1").NumberFormat = "General"
Range("A1").NumberFormat = "$#,##0.00"
Range("A1").NumberFormat = "$#,##0.00_);[Red]($#,##0.00)"

'Change the Horizontal alignment of a cell.
Range("A1").HorizontalAlignment = xlHAlignCenter

            'Name                               Value   Description
            'xlHAlignCenter                     -4108   Center.
            'xlHAlignCenterAcrossSelection      7       Center across selection.
            'xlHAlignDistributed                -4117   Distribute.
            'xlHAlignFill                       5       Fill.
            'xlHAlignGeneral                    1       Align according to data type.
            'xlHAlignJustify                    -4130   Justify.
            'xlHAlignLeft                       -4131   Left.
            'xlHAlignRight                      -4152   Right.

'Change the Vertical alignment of a cell.
Range("A1").VerticalAlignment = xlVAlignTop

            'Name                  Value   Description
            'xlVAlignBottom        -4107   Bottom
            'xlVAlignCenter        -4108   Center
            'xlVAlignDistributed   -4117   Distributed
            'xlVAlignJustify       -4130   Justify
            'xlVAlignTop           -4160   Top

'Wrape the text content of a cell.
Range("A1").WrapText = True

'If set to true it will automatically reduce the size of the font in order to fit the content in the column width.
Rows(1).ShrinkToFit = True

'Merge a Range of Cells
Range("A1:A4").MergeCells = True

'Change the reading order of a Cell.
Range("A1").ReadingOrder = xlRTL

            'Possible Values
            'xlRTL
            'xlLTR
            'xlContext

'Change the orientation of the cell content.
Range("A1").Orientation = xlHorizontal

            'Possible Values
            'xlDownward
            'xlHorizontal
            'xlUpward
            'xlVertical


'Change the font of a cell.
Range("A1:A5").Font.Name = "Calibri"

'Change the font style of a cell.
Range("A1:A5").Font.FontStyle = "Italic"

            'Possible Values
            'Regular
            'Bold
            'Italic
            'Bold Italic

'Change the Font Size - Range 1 to 409
Range("A1").Font.Size = 14

'Change the Underline Style of a Font.
Range("A1").Font.Underline = xlUnderlineStyleDouble

            'Name                                Value   Description
            'xlUnderlineStyleDouble              -4119   Double thick underline.
            'xlUnderlineStyleDoubleAccounting     5      Two thin underlines placed close together.
            'xlUnderlineStyleNone                -4142   No underlining.
            'xlUnderlineStyleSingle               2      Single underlining.
            
            
'Change the font color, using different methods.
Range("A1").Font.Color = vbBlack
Range("A1").Font.Color = 0
Range("A1").Font.Color = RGB(0, 0, 0)

'Add Font, Strikethrough, subscript, or superscript.
Range("A1").Font.Strikethrough = True
Range("A1").Font.Subscript = True
Range("A1").Font.Superscript = True

'Add a Cell Border, and select a border style.
Range("A1").Borders(xlEdgeBottom).LineStyle = xlContinuous
Range("A1").Borders(xlEdgeBottom).LineStyle = xlNone

            'Border Constant        Value     Description
            'xlDiagonalDown          5        (Border running from the upper left-hand corner to the lower right of each cell in the range).
            'xlDiagonalUp            6        (Border running from the lower left-hand corner to the upper right of each cell in the range).
            'xlEdgeBottom            9        (Border at the bottom of the range).
            'xlEdgeLeft              7        (Border at the left-hand edge of the range).
            'xlEdgeRight             10       (Border at the right-hand edge of the range).
            'xlEdgeTop               8        (Border at the top of the range).
            'xlInsideHorizontal      12       (Horizontal borders for all cells in the range except borders on the outside of the range).
            'xlInsideVertical        11       (Vertical borders for all the cells in the range except borders on the outside of the range).


            'Name                Value      Description
            'xlContinuous        1          Continuous line.
            'xlDash             -4115       Dashed line.
            'xlDashDot           4          Alternating dashes and dots.
            'xlDashDotDot        5          Dash followed by two dots.
            'xlDot              -4118       Dotted line.
            'xlDouble           -4119       Double line.
            'xlLineStyleNone    -4142       No line.
            'xlSlantDashDot      13         Slanted dashes.


'Change the border weight
Range("A1").Borders(xlEdgeBottom).Weight = xlThin

            'Name           Value   Description
            'xlHairline         1   Hairline (thinnest border).
            'xlMedium       -4138   Medium.
            'xlThick            4   Thick (widest border).
            'xlThin             2   Thin.

'Change the Border Color.
Range("A1").Borders(xlEdgeBottom).Color = vbGreen
Range("A1").Borders(xlEdgeBottom).Color = RGB(255, 0, 0)

'Change the Pattern of the cell interior.
Range("A1").Interior.Pattern = xlPatternCrissCross

'xlPatternAutomatic (Excel controls the pattern.)
'xlPatternChecker (Checkerboard.)
'xlPatternCrissCross (Criss-cross lines.)
'xlPatternDown (Dark diagonal lines running from the upper left to the lower right.)
'xlPatternGray16 (16% gray.)
'xlPatternGray25 (25% gray.)
'xlPatternGray50 (50% gray.)
'xlPatternGray75 (75% gray.)
'xlPatternGray8 (8% gray.)
'xlPatternGrid (Grid.)
'xlPatternHorizontal (Dark horizontal lines.)
'xlPatternLightDown (Light diagonal lines running from the upper left to the lower right.)
'xlPatternLightHorizontal (Light horizontal lines.)
'xlPatternLightUp (Light diagonal lines running from the lower left to the upper right.)
'xlPatternLightVertical (Light vertical bars.)
'xlPatternNone (No pattern.)
'xlPatternSemiGray75 (75% dark moiré.)
'xlPatternSolid (Solid color.)
'xlPatternUp (Dark diagonal lines running from the lower left to the upper right.)

'Change the interior color of the cell.
Range("A1").Interior.Color = RGB(255, 0, 0)

'You can enter a number from -1 (darkest) to 1 (lightest) for the TintAndShade property. Zero (0) is neutral.
Range("A1").Interior.TintAndShade

'Change the Cell interior to a theme color.
Range("A1").Interior.ThemeColor = xlThemeColorDark1



End Sub


Sub Test()

   'Add Gradients
        With Range("C1").Interior
            .Pattern = xlPatternLinearGradient
            .Gradient.Degree = 180
            
            'Adjust Color Stops
                'Clear Default Color Stops
                    .Gradient.ColorStops.Clear
                
                'Add A Color Stop
                    With .Gradient.ColorStops.Add(0)
                        .Color = RGB(255, 255, 255)
                    End With
                
                'Add Another Color Stop
                    With .Gradient.ColorStops.Add(1)
                        .Color = RGB(141, 180, 227)
                    End With
        End With




End Sub

Sub row()

Rows(5).ShrinkToFit = True

End Sub
