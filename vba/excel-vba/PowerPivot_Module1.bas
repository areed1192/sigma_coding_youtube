Attribute VB_Name = "Module1"
Sub CreateDataModel()

    Dim DataTbl As ListObject
        
    Set DataTbl = Worksheets("Sheet1").ListObjects.Add(SourceType:=xlSrcRange, _
    Source:=Range("A1:C3"), XlListObjectHasHeaders:=xlYes, TableStyleName:="TableStyleLight9")


'Set ModelConn = Model.AddConnection(


End Sub
