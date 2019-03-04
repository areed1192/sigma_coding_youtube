Option Explicit

Sub GetInfoAboutMyQueryTable()

With ActiveSheet.ListObjects(1)

    'Use the SourceDataFile Property to get the path to the source data file.
    MsgBox .QueryTable.SourceDataFile

    'Use the SourceConnectionFile Property to get the path to the connection file you exported.
    MsgBox .QueryTable.SourceConnectionFile
    
    'Use the Connection Property to find out the connection string used in this QueryTable.
    MsgBox .QueryTable.Connection
    
    'Use the Creator Property -- USED FOR MACS -- To see which application created this QueryTable.
    MsgBox .QueryTable.Creator
    
    'Get the workbook connection used in this QueryTable
    MsgBox .QueryTable.WorkbookConnection
    
    'Get the Result Range of My Query.
    MsgBox .QueryTable.ResultRange.Address
    
    'Get the QueryType of My Query
    MsgBox .QueryTable.QueryType
    
        'xlADORecordset  7   Based on an ADO recordset query
        'xlDAORecordset  2   Based on a DAO recordset query, for query tables only
        'xlODBCQuery     1   Based on an ODBC data source
        'xlOLEDBQuery    5   Based on an OLE DB query, including OLAP data sources
        'xlTextImport    6   Based on a text file, for query tables only
        'xlWebQuery      4   Based on a Web page, for query tables only

End With

End Sub
