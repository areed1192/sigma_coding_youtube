Sub CreateListObject()

    Dim ListObj As ListObject
    Set ListObj = ActiveSheet.ListObjects.Add(SourceType:=xlSrcRange, _
                                              Source:=Range("$A$1:$C$5"), _
                                              XlListObjectHasHeaders:=xlNo, _
                                              Destination:=Range("$I$1"), _
                                              TableStyleName:="TableStyleLight1")
                                              
    'PARMETER ONE:
    'SourceType:= xlSrcExternal, (External Data Connection) _
                  xlSrcModel (PowerPivot Model), _
                  xlSrcQuery (Query), _
                  xlSrcRange (Range), _
                  xlSrcXml (XML)
    
    'PARAMTER TWO:
    'Source:=  When SourceType equal xlSrcRange we need to pass through a range object. _
               If Omitted, then the source will default to the range returned by the list range _
               detection code. When SourceType equals xlSrcExternal an array of string values must be _
               passed through with the following three elements: _
                    0 - URL to SharePoint Site _
                    1 - ListName _
                    2 - View GUID
                    
    'PARAMETER THREE:
    'LinkSource:= Indicates whether an external data source is to be linked to the ListObject object. _
                  If SourceType is xlSrcExternal, default is True. Invalid if SourceType is xlSrcRange , _
                  and will return an error if not omitted.
                  
    'PARAMETER FOUR:
    'XlListObjectHasHeaders:= An xlYesNoGuess constant that indicates whether the data being imported has column labels. _
                              If the Source does not contain headers, Excel will automatically generate headers. _
                                  Value Default: xlGuess. _
                                  Value Yes: xlYes. _
                                  Value No: xlNo.
                  
    
    'PARAMETER FIVE:
    'Destination:= A Range object specifying a single-cell reference as the destination for the top-left corner of the new list object. _
                   If the Range object refers to more than one cell, an error is generated. The Destination argument must be specified when SourceType is set to xlSrcExternal. _
                   The Destination argument is ignored if SourceType is set to xlSrcRange. The destination range must be on the worksheet that _
                   contains the ListObjects collection specified by expression. New columns will be inserted at the Destination to fit the new list. _
                   Therefore, existing data will not be overwritten.
    
    'PARAMETER SIX:
    'TableStyleName:= This is the name of a table style that we can apply to the table, in this example we used TableStyleLight1.


End Sub

Sub CreateTableFromSharePointList()
    
   'Declare Variables
   Dim Server As String
   Dim ListName As String
   Dim ViewName As String
   Dim ListObj As ListObject
   
   'Get Connection String Paramaters
   Server = "https://petco.sharepoint.com/sites/AnalyticsCommunity/_vti_bin" '<<< THIS IS YOUR SHAREPOINT SITE REPLACE WIT THE PROPER URL & MAKE SURE Vti_bin IS AT THE END.
   ListName = "{6DC036F2-05F3-4EB9-A732-43E1C6827682}" '<<< This is the list name
   ViewName = "{5D8F3E3B-2425-4CCB-953C-22E82242EC05}" '<<< This is the view name
   
   
   'TO GET THE LIST NAME/ID
   'Go To the List in SharePoin, click Settings & then right click "Audience Target Settings"
   
   'Click Copy link address & Paste it in VBA.
   'https://petco.sharepoint.com/sites/AnalyticsCommunity/_layouts/ListEnableTargeting.aspx?List={6dc036f2-05f3-4eb9-a732-43e1c6827682}
   
   'Delete Everything Before List & this is the List ID/Name:
   'List={6dc036f2-05f3-4eb9-a732-43e1c6827682}
   
   
   'TO GET THE VIEW ID
   'Go to the list in SharePoint, then the list section at the top of the ribbon, click modify view, & copy the URL on the new page.
   'https://petco.sharepoint.com/sites/AnalyticsCommunity/_layouts/15/ViewEdit.aspx?List=6dc036f2-05f3-4eb9-a732-43e1c6827682&View=%7B5D8F3E3B-2425-4CCB-953C-22E82242EC05%7D&Source=https%3A%2F%2Fpetco%2Esharepoint%2Ecom%2Fsites%2FAnalyticsCommunity%2FLists%2FPlatforms_Master%2FAllItems%2Easpx

   'Keeo only the View Section
   'View=%7B5D8F3E3B-2425-4CCB-953C-22E82242EC05%7D
   
   'Replace %7B with "{" & %7D with "}", you now have the View ID.
   'View={5D8F3E3B-2425-4CCB-953C-22E82242EC05}
   

  Set ListObj = ActiveSheet.ListObjects.Add(SourceType:=xlSrcExternal, _
                                            Source:=Array(Server, ListName, ViewName), _
                                            XlListObjectHasHeaders:=xlYes, _
                                            LinkSource:=xlYes, _
                                            Destination:=Range("$A$21"), _
                                            TableStyleName:="TableStyleLight1")
  
  
End Sub


Sub SelectListObject()

    Dim LisObj As ListObject

    Set LisObj = ActiveSheet.ListObjects(2)
        LisObj.Range.Select
        LisObj.HeaderRowRange.Select
        LisObj.DataBodyRange.Select
        LisObj.TotalsRowRange.Select

End Sub



