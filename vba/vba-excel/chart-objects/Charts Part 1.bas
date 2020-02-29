Sub CreateChart()

'Declare some variables
Dim Chrt As ChartObject
Dim DataRng As Range

'Add a chart object, this would be an empty shell
Set Chrt = ActiveSheet.ChartObjects.Add(Left:=400, _
                                        Width:=400, _
                                        Height:=400, _
                                        Top:=50)

    'Define the data to be used in the chart.
    Set DataRng = Range("B3:E7")
    Chrt.Chart.SetSourceData Source:=DataRng
    
    'Define the type of chart it is.
    Chrt.Chart.ChartType = xlBarClustered

'Lets add a title
Chrt.Chart.HasTitle = True

'Create a reference to that title
Dim ChrtTitle As ChartTitle
Set ChrtTitle = Chrt.Chart.ChartTitle
    
    'Do some formatting with the title.
    ChrtTitle.Text = "Performance"
    ChrtTitle.Shadow = False
    ChrtTitle.Characters.Font.Bold = False
    ChrtTitle.Characters.Font.Name = "Arial Nova"

'Add a legend to the chart
Chrt.Chart.HasLegend = True

'Create a reference to that legend
Dim ChrtLeg As Legend
Set ChrtLeg = Chrt.Chart.Legend

    'Do some formatting
    ChrtLeg.Position = xlLegendPositionTop
    ChrtLeg.Height = 20

'Remove the gridlines
Chrt.Chart.SetElement msoElementPrimaryCategoryGridLinesNone
Chrt.Chart.SetElement msoElementPrimaryValueGridLinesNone

'Make sure the chart has some axes, it's usually true by default
Chrt.Chart.HasAxis(xlCategory, xlPrimary) = True
Chrt.Chart.HasAxis(xlValue, xlPrimary) = True

'Make sure each axis has a title
Chrt.Chart.Axes(xlValue, xlPrimary).HasTitle = True
Chrt.Chart.Axes(xlCategory, xlPrimary).HasTitle = True

'Take the newly created title and create a reference to it.
Dim AxisTitle As AxisTitle
Set AxisTitle = Chrt.Chart.Axes(xlCategory, xlPrimary).AxisTitle

    'Do some formatting.
    AxisTitle.Text = "Years"
    AxisTitle.HorizontalAlignment = xlCenter
    AxisTitle.Characters.Font.Color = vbRed
    
End Sub
