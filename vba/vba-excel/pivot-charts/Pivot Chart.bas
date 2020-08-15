Option Explicit

Sub CreatePivotChart()

Dim xlDataWorkbook As Workbook
Dim xlDataSourceSheet As Worksheet
Dim xlDataSourceTable As ListObject
Dim xlDataChartSheet As Worksheet

Dim xlPvtCache As PivotCache
Dim xlPvtTable As PivotTable
Dim xlPvtField As PivotField
Dim xlPvtShape As Shape
Dim xlPvtChart As Chart

Dim xlChartAxis As Axis

'Define the Workbook
Set xlDataWorkbook = ThisWorkbook

'Define the data source sheet.
Set xlDataSourceSheet = xlDataWorkbook.Worksheets("USA_Daily_All")

'Define where we want the chart to go.
Set xlDataChartSheet = xlDataWorkbook.Worksheets("USA_Daily_All_Chart")

'Set the Data Table Source
Set xlDataSourceTable = xlDataSourceSheet.ListObjects("DailyUnitedStatesAll")

'Create the Pivot Cache.
Set xlPvtCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=xlDataSourceTable.Name, Version:=6)

'Create the Pivot Table.
Set xlPvtTable = xlPvtCache.CreatePivotTable(TableDestination:=xlDataChartSheet.Range("A2"))

'Create the Pivot Shape.
Set xlPvtShape = xlDataChartSheet.Shapes.AddChart2(XlChartType:=XlChartType.xlColumnStacked100, _
                                                   Left:=100, _
                                                   Top:=100, _
                                                   Width:=900, _
                                                   Height:=300)
                                                   
'Create the Pivot Chart.
Set xlPvtChart = xlPvtShape.Chart

    'Set the Data Source.
    xlPvtChart.SetSourceData Source:=xlPvtTable.TableRange1
    
    'Could also reference the table like this.
    'xlPvtChart.PivotLayout.PivotTable
    
    
    'Grab the `DateParsed` field.
    Set xlPvtField = xlPvtTable.PivotFields("DateParsed")
    
    'Add the `DateParsed` field.
    With xlPvtField
        .Orientation = xlColumnField
        .Position = 1
    End With
    
    'Auto Group the Dates.
    xlPvtField.AutoGroup
    
    
    'Grab the `positive` field.
    Set xlPvtField = xlPvtTable.PivotFields("positive")
    
    'Add the `positive` field.
    With xlPvtField
        .Orientation = xlDataField
        .Position = 1
    End With
    
    
    'Grab the `dataQualityGrade` field.
    Set xlPvtField = xlPvtTable.PivotFields("dataQualityGrade")
    
    'Add the `dataQualityGrade` field.
    With xlPvtField
        .Orientation = xlPageField
        .Position = 1
    End With
    
    
    'Grab the `Months` field.
    Set xlPvtField = xlPvtTable.PivotFields("Months")
    
    'Add the `Months` field.
    With xlPvtField
        .Orientation = xlRowField
        .Position = 1
    End With
    
    'Change the color.
    xlPvtChart.ChartColor = 18
    
    'Delete the Gridlines.
    xlPvtChart.Axes(xlValue).MajorGridlines.Delete
    
    'Change the Font.
    With xlPvtChart.ChartArea.Format.TextFrame2.TextRange.Font
        .NameComplexScript = "Roboto"
        .NameFarEast = "Roboto"
        .Name = "Roboto"
        .Size = 10
    End With
    
    'Grab the Date Field.
    Set xlPvtField = xlPvtTable.PivotFields("DateParsed")
        
        'Add a Filter, that only selects this Quarter.
        xlPvtField.PivotFilters.Add2 Type:=xlDateThisQuarter
        
       
    'Set the Data Labels.
    xlPvtChart.SetElement Element:=msoElementDataLabelCenter
    
    'Set the Legend.
    xlPvtChart.SetElement Element:=msoElementLegendTop
    
    'Remove the Primary Value Axis.
    xlPvtChart.SetElement Element:=msoElementPrimaryValueAxisNone
    
    'Set the Axis Title.
    xlPvtChart.SetElement Element:=msoElementPrimaryValueAxisTitleAdjacentToAxis
    
    'Grab the Value Axis.
    Set xlChartAxis = xlPvtChart.Axes(xlValue)
        
        'Set the caption
        'xlChartAxis.AxisTitle.Text = "Positive Cases"
        
        'Set the Font to Bold.
        'xlChartAxis.AxisTitle.Format.TextFrame2.TextRange.Font.Bold = True
    
    'xlPvtChart.Axes(xlValue).AxisTitle.Text = "Positive Cases"
    
    'Set the Axis Title
    xlPvtChart.SetElement Element:=msoElementPrimaryCategoryAxisTitleAdjacentToAxis
    
    'Grab the Axis Object.
    Set xlChartAxis = xlPvtChart.Axes(xlCategory)
        
        'Set the caption
        xlChartAxis.AxisTitle.Caption = "Month"
        
        'Set the Font to Bold.
        xlChartAxis.AxisTitle.Format.TextFrame2.TextRange.Font.Bold = True
    
    'Change the `Chart` name.
    xlPvtShape.Name = "CovidAnalysis"
        
End Sub

