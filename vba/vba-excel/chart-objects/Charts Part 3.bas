Sub WorkWithChartArea()

'Declare some variables
Dim ChtSer As Series
Dim ChtSerColl As SeriesCollection
Dim Chrt As ChartObject
Dim ChrtAxs As Axis

'Create a reference to the chart
Set Chrt = ActiveSheet.ChartObjects(1)

    'Add some lines to the chart area
    Chrt.Chart.ChartArea.Format.Line.Visible = True
    
    'Add some 3D effects to the chart area.
    Chrt.Chart.ChartArea.Format.ThreeD.BevelTopType = msoBevelCircle
    
    'Copy the chart area
    Chrt.Chart.ChartArea.Copy
    
    'Clear the formats
    Chrt.Chart.ChartArea.ClearFormats
    Chrt.Chart.ChartArea.Clear
    Chrt.Chart.ChartArea.ClearContents
        
    'Add a shadow format and then do some additional formatting.
    With Chrt.Chart.ChartArea.Format.Shadow
         .Visible = True
         .Style = msoShadowStyleOuterShadow
         .Transparency = 0.4
         .ForeColor.RGB = RGB(34, 60, 252)
    End With
    
    'Add some rounded corners
    Chrt.Chart.ChartArea.RoundedCorners = True
    
    'Select the Chart Area
    Chrt.Chart.ChartArea.Select
    
    'Set the Chart Axis
    Set ChrtAxs = Chrt.Chart.Axes(Type:=xlValue, AxisGroup:=xlPrimary)
    
    'Change the major and minor unit of the Axes
    ChrtAxs.MajorUnit = 1000
    ChrtAxs.MinorUnit = 400
    
    'Change the scale type
    ChrtAxs.ScaleType = xlLogarithmic 'or xlScaleLinear
    
    'Change the Tick Label Position
    ChrtAxs.TickLabelPosition = xlTickLabelPositionHigh

    'Plot by row or column
    Chrt.Chart.PlotBy = xlRows
    Chrt.Chart.PlotBy = xlColumns
    
    'Change where the Axis Crosses at.
    ChrtAxs.Crosses = xlMaximum
    ChrtAxs.CrossesAt = 1
    
    'Set the Chart Axis
    Set ChrtAxs = Chrt.Chart.Axes(Type:=xlCategories, AxisGroup:=xlPrimary)
    
    'Change the Tick Label Spacing - Applies only to category and series axes
    ChrtAxs.TickLabelSpacing = 10 '1 to 31999
    
End Sub
