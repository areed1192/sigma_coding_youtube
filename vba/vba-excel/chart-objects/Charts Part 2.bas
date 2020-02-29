Sub WorkWithSeries()

'Declare some variables
Dim ChtSer As Series
Dim ChtSerColl As SeriesCollection
Dim Chrt As ChartObject

'Create a reference to the chart
Set Chrt = ActiveSheet.ChartObjects(1)

'Get the series collection from the chart
Set ChtSerColl = Chrt.Chart.SeriesCollection

'Print the name of each series in the collection
For Each ChtSer In ChtSerColl
    Debug.Print ChtSer.Name
Next

'Select one series using the item method.
Set ChtSer = ChtSerColl.Item("Profit")

'With thr profit series
With ChtSer
     
     'Add some data labels
     .HasDataLabels = True
     .ApplyDataLabels Type:=xlValue
     
     'Change the fill color of the series.
     .Format.Fill.ForeColor.RGB = RGB(34, 60, 252)
     
     'Add some error bars
     .HasErrorBars = True
     
     'Add and format a leader line for just that series
     .HasLeaderLines = True
     .LeaderLines.Border.Color = vbRed
     
     'Format the series border
     With .Format
          .Line.Visible = msoCTrue
          .Line.Weight = 3
          .Line.DashStyle = msoLineDashDot
          .Line.ForeColor.TintAndShade = 1
          '.ThreeD.BevelBottomDepth = msoBevelConvex
     End With
     
     'Print the Axis it belongs to, 1 is primary 2 is secondary
     Debug.Print ChtSer.AxisGroup
     
     'Is this Series filtered?
     Debug.Print ChtSer.IsFiltered

End With

End Sub
