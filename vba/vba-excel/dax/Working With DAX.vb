Sub WorkingWithRelationships()

'Declare our Variables.
Dim xlBook As Workbook
Dim xlTableSheet As Worksheet
Dim xlPivotSheet As Worksheet
Dim xlBookConnections As Connections
Dim xlbookConnection As WorkbookConnection

Dim xlModel As Model
Dim xlProductColumn As ModelTableColumn
Dim xlProductNameColumn As ModelTableColumn

Dim xlListObjects As ListObjects
Dim xlListObjectSales As ListObject
Dim xlListObjectProducts As ListObject

Dim xlMeasures As ModelMeasures
Dim xlMeasure As ModelMeasure

Dim xlPivotCache As PivotCache
Dim xlPivotTableModel As PivotTable
Dim xlCubeFields As CubeFields
Dim xlCubeField As CubeField

'Set the Book.
Set xlBook = ThisWorkbook

'Grab the Connections Collection.
Set xlBookConnections = xlBook.Connections

'Grab the Workbook Model.
Set xlModel = xlBook.Model

'Grab the Sheets.
Set xlTableSheet = xlBook.Worksheets("Tables")
Set xlPivotSheet = xlBook.Worksheets("Pivot")

'Grab the List Object Collection.
Set xlListObjects = xlTableSheet.ListObjects

'Grab the Sales Table.
Set xlListObjectSales = xlListObjects.Item("fSales")

'Grab the Product Table.
Set xlListObjectProducts = xlListObjects.Item("dProducts")
    
    'Add the Sales Table.
    xlBookConnections.Add2 Name:="fSalesTableConnection", _
                           Description:="Represents our Sales Data Table", _
                           ConnectionString:="WORKSHEET;" + xlBook.Path, _
                           CommandText:=xlBook.Name + "!fSales", _
                           lCmdType:=7, _
                           CreateModelConnection:=True, _
                           ImportRelationships:=False

    'Add the Products Table.
    xlBookConnections.Add2 Name:="dProductsTableConnection", _
                           Description:="Represents our Product Data Table", _
                           ConnectionString:="WORKSHEET;" + xlBook.Path, _
                           CommandText:=xlBook.Name + "!dProducts", _
                           lCmdType:=7, _
                           CreateModelConnection:=True, _
                           ImportRelationships:=False
                           
    
    'Grab the Model Columns.
    Set xlProductColumn = xlModel.ModelTables("fSales").ModelTableColumns("Product")
    Set xlProductNameColumn = xlModel.ModelTables("dProducts").ModelTableColumns("Product Name")
    
    'Lets Add a relationship to our model.
    xlModel.ModelRelationships.Add ForeignKeyColumn:=xlProductColumn, PrimaryKeyColumn:=xlProductNameColumn
    
    'Create a Pivot Cache.
    Set xlPivotCache = xlBook.PivotCaches.Create(SourceType:=xlExternal, SourceData:=xlBookConnections("ThisWorkbookDataModel"), Version:=6)

    'Create a Pivot Table from that Cache.
    Set xlPivotTableModel = xlPivotCache.CreatePivotTable(TableDestination:=xlPivotSheet.Range("A3"), TableName:="SalesAnalysis", DefaultVersion:=6)

    'Lets add a Measure.
    Set xlMeasure = xlModel.ModelMeasures.Add(MeasureName:="Total Sales", _
                                              AssociatedTable:=xlModel.ModelTables("fSales"), _
                                              Formula:=" SUMX(fSales, fSales[Units Sold] * RELATED(dProducts[Price]))", _
                                              FormatInformation:=xlModel.ModelFormatCurrency("Default", 2), _
                                              Description:="The number of units sold times the unit price.")


    'Grab the Cube Fields.
    Set xlCubeFields = xlPivotTableModel.CubeFields

    'Add the Fields to specific Orientation.
    With xlCubeFields("[fSales].[Region]")
        .Orientation = xlRowField
        .Position = 1
    End With

    With xlCubeFields("[dProducts].[Product Name]")
        .Orientation = xlRowField
        .Position = 2
    End With

    'Grab the Sum of Price Measure.
    Set xlCubeField = xlPivotTableModel.CubeFields.GetMeasure(AttributeHierarchy:="[dProducts].[Price]", _
                                                              Function:=xlSum, _
                                                              Caption:="PriceSum")

    'Add the Sum of Price Measure.
    xlPivotTableModel.AddDataField Field:=xlCubeField
    
    'Add the Total Sales Measure.
    xlPivotTableModel.AddDataField Field:=xlPivotTableModel.CubeFields("[Measures].[Total Sales]")
    
    'Change it to Light Green.
    xlPivotTableModel.TableStyle2 = "PivotStyleLight14"
    
    'Change the Row Height.
    xlPivotSheet.Cells.RowHeight = 18

End Sub