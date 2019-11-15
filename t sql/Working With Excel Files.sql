
/*

  UNDERSTANDING THE sp_MSset_oledb_prop
  ------------------------------------------------------------------------------------------------------------------
  This is a stored procedure in the Microsoft SQL Server that will help manipulate existing OLEDB providers in the
  server. In this example we are just pulling all of the providers we currently have in our server. There is no manipulation
  at this point. This stored procedure exists under the Master databse.

*/
EXEC sp_MSset_oledb_prop
GO;


/*

  UNDERSTANDING THE Show Advanced Options
  ------------------------------------------------------------------------------------------------------------------
  Some configuration options, such as affinity mask and recovery interval, are designated as advanced options. By 
  default, these options are not available for viewing and changing. To make them available, set the ShowAdvancedOptions 
  configuration option to 1. In this case we want to see the "Ad Hoc Distributed Queries" which is an advanced option.

*/
EXEC sp_configure 'Show Advanced Options', 1
RECONFIGURE -- This will Permentantly Reconfigure the Server MAY NOT WANT THIS
GO;


/*

  UNDERSTANDING THE Ad Hoc Distributed Queries
  ------------------------------------------------------------------------------------------------------------------
   By default, SQL Server does not allow ad hoc distributed queries using OPENROWSET and OPENDATASOURCE. When this 
   option is set to 1, SQL Server allows ad hoc access. When this option is not set or is set to 0, SQL Server does 
   not allow ad hoc access. Ad hoc distributed queries use the OPENROWSET and OPENDATASOURCE functions to connect to 
   remote data sources that use OLE DB. OPENROWSET and OPENDATASOURCE should be used only to reference OLE DB data 
   sources that are accessed infrequently. For any data sources that will be accessed more than several times, define
   a linked server.

*/ 
EXEC sp_configure 'Ad Hoc Distributed Queries', 1
RECONFIGURE -- This will Permentantly Reconfigure the Server MAY NOT WANT THIS
GO;


/*

  UNDERSTANDING THE AllowInProcess
  ------------------------------------------------------------------------------------------------------------------
  By selecting Allow Inprocess, SQL Server allows the provider to be instantiated or allows the provider to run as 
  an In Process server. When the option is not set, the default behavior is to allow the provider to run outside the 
  SQL Server process.

*/
EXEC sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.12.0', N'AllowInProcess', 1 
GO;


/*

  UNDERSTANDING THE DynamicParameters
  ------------------------------------------------------------------------------------------------------------------
  Allows SQL placeholders (represented by '?') in parameterized queries.

*/
EXEC sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.12.0', N'DynamicParameters', 1 
GO;


/*
DRIVER INFO:
------------------------------------------

LINK TO DRIVERs: https://www.microsoft.com/en-us/download/details.aspx?id=13255

NAME:
Microsoft.Jet.OLEDB.4.0

SQL Version Used With:
32-bit SQL Server

Excel File Version Used With: 
Excel 2003 files.

--------------------------------------

NAME:
Microsoft.ACE.OLEDB.12.0

SQL Version Used With:
64-bit SQL Server or 32-bit SQL Server

Excel File Version Used With: 
64-bit SQL Server any Excel files.
32-bit SQL Server for Excel 2007 files.


Step Num	Step										SQL Server x86 for Excel 2003		SQL Server x86 for Excel 2007		SQL Server x86 for Excel Any Version
--------	--------------------------------------		-----------------------------		-----------------------------		------------------------------------
1			Install Microsoft.ACE.OLEDB.12.0 driver		Not Needed							x86									x64
2			Configure Ad Hoc Distributed Queries		Yes									Yes									Yes
3			Grant rights to TEMP directory				Yes									Yes									Not Needed
4			Configure ACE OLE DB properties				Not Needed							Yes									Yes


 POSSIBLE ERRORS WHEN IMPORTING EXCEL FILES:
 ------------------------------------------------------------------------------------------------------------------
 ------------------------------------------------------------------------------------------------------------------

 Link to Microsoft Documentation: https://docs.microsoft.com/en-us/sql/relational-databases/import-export/import-data-from-excel-to-sql?view=sql-server-2017

 ------------------------------------------------------------------------------------------------------------------
 ------------------------------------------------------------------------------------------------------------------

 ERROR NAME:
 Ad Hoc Distributed Queries is turned off.
 
 ERROR DESCRIPTION:
 You are trying to use OPENROWSET without enabling 'Ad Hoc Distributed Queries'. Run the following code to resolve the issue:

 EXEC sp_configure 'Ad Hoc Distributed Queries', 1


 ERROR RESPONSE:
 1. Msg 15281, Level 16, State 1, Line 1
	SQL Server blocked access to STATEMENT 'OpenRowset/OpenDatasource' of component 'Ad Hoc Distributed Queries' because this component is turned off as part of
	the security configuration for this server. A system administrator can enable the use of 'Ad Hoc Distributed Queries' by using sp_configure. For more information 
	about enabling 'Ad Hoc Distributed Queries',see "Surface Area Configuration" in SQL Server Books Online.

 ------------------------------------------------------------------------------------------------------------------

 ERROR NAME:
 "Microsoft.ACE.OLEDB.12.0" has not been registered

 ERROR DESCRIPTION:
 This error occurs because the OLEDB provider is not installed. Install it from Microsoft Access Database Engine 2010 Redistributable. 
 Be sure to install the 64-bit version if Windows and SQL Server are both 64-bit.

 ERROR RESPONSE:
 1. Msg 7403, Level 16, State 1, Line 3
    The OLE DB provider "Microsoft.ACE.OLEDB.12.0" has not been registered.

 ------------------------------------------------------------------------------------------------------------------

 ERROR NAME:
 Cannot create an instance of OLE DB provider "Microsoft.ACE.OLEDB.12.0" for linked server "(null)".

 ERROR DESCRIPTION:
 This indicates that the Microsoft OLEDB has not been configured properly. Run the following Transact-SQL code to resolve this:

 EXEC sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.12.0', N'AllowInProcess', 1   
 EXEC sp_MSset_oledb_prop N'Microsoft.ACE.OLEDB.12.0', N'DynamicParameters', 1

 ERROR RESPONSE:
 1. Msg 7302, Level 16, State 1, Line 3
    Cannot create an instance of OLE DB provider "Microsoft.ACE.OLEDB.12.0" for linked server "(null)".

 ------------------------------------------------------------------------------------------------------------------

 ERROR NAME:
 The 32-bit OLE DB provider "Microsoft.ACE.OLEDB.12.0" cannot be loaded in-process on a 64-bit SQL Server.

 ERROR DESCRIPTION:
 This occurs when a 32-bit version of the OLD DB provider is installed with a 64-bit SQL Server. To resolve this 
 issue, uninstall the 32-bit version and install the 64-bit version of the OLE DB provider instead.

 ERROR RESPONSE:
 1. Msg 7438, Level 16, State 1, Line 3
    The 32-bit OLE DB provider "Microsoft.ACE.OLEDB.12.0" cannot be loaded in-process on a 64-bit SQL Server.

 ------------------------------------------------------------------------------------------------------------------

 ERROR NAME:
 The OLE DB provider "Microsoft.ACE.OLEDB.12.0" for linked server "(null)" reported an error. The provider did not give 
 any information about the error.

 ERROR DESCRIPTION:
 Both of these errors typically indicate a permissions issue between the SQL Server process and the file. Ensure that the 
 account that is running the SQL Server service has full access permission to the file. We recommend against trying to import 
 files from the desktop.

 ERROR RESPONSE:
 1. Msg 7399, Level 16, State 1, Line 3
    The OLE DB provider "Microsoft.ACE.OLEDB.12.0" for linked server "(null)" reported an error. The provider did not give any information about the error.

 2. Msg 7303, Level 16, State 1, Line 3
    Cannot initialize the data source object of OLE DB provider "Microsoft.ACE.OLEDB.12.0" for linked server "(null)".

 ------------------------------------------------------------------------------------------------------------------

 ERROR NAME:
 OLE DB provider "Microsoft.Jet.OLEDB.4.0" for linked server "(null)" returned message "Unspecified error". ONLY APPLIES TO 32-BIT SQL SERVER

 ERROR DESCRIPTION:
 The main problem is that an OLE DB provider creates a temporary file during the query in the SQL Server temp directory using credentials of a user who run the query.
 If the SQL user doesn't have access to the temp folder than SQL can't create the temporary folder. Perform the following steps to fix the issue:

 Step 1: Go to the Temp Folder
	
	If SQL Server is run under Network Service account the temp directory is like: C:\Windows\ServiceProfiles\NetworkService\AppData\Local\Temp
	
	If SQL Server is run under Local Service account the temp directory is like: C:\Windows\ServiceProfiles\LocalService\AppData\Local\Temp

 Step 2: Grant Read/Write Access to the Temp folder for the specified user. This can be done using icals.

	icals C:\Windows\ServiceProfiles\NetworkService\AppData\Local\Temp /grant <USER>:(R,W)


 ERROR RESPONSE:
 1. Msg 7303, Level 16, State 1, Line 1
	Cannot initialize the data source object of OLE DB provider "Microsoft.Jet.OLEDB.4.0" for linked server "(null)".

*/



USE SigmaCodingDatabase
GO

-- Select all of the data on Sheet "Data_A"
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0','Excel 12.0;Database=C:\Users\Alex\OneDrive\Desktop\Sales_Data.xlsx', [Data_A$])

-- Select a portion of the data on Sheet "Data_B"
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0','Excel 12.0;Database=C:\Users\Alex\OneDrive\Desktop\Sales_Data.xlsx', [Data_B$A1:C2])

-- Select a portion of the data on Sheet "Data_C"
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0','Excel 12.0;Database=C:\Users\Alex\OneDrive\Desktop\Sales_Data.xlsx', [Data_C$])

-- Select the named range found on Sheet "Data_D"
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0','Excel 12.0;Database=C:\Users\Alex\OneDrive\Desktop\Sales_Data.xlsx', CountryData) -- [CountryData]  Also works.

-- Select the named range found on Sheet "Data_D"
SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0','Excel 12.0;Database=C:\Users\Alex\OneDrive\Desktop\Sales_Data.xlsx', 'SELECT * FROM [Data_A$] WHERE Sales > 20000') -- [CountryData]  Also works.


-- Insert into a Temporary Table called MyData
SELECT * INTO MyData FROM OPENROWSET('Microsoft.ACE.OLEDB.12.0','Excel 12.0;Database=C:\Users\Alex\OneDrive\Desktop\Sales_Data.xlsx', CountryData)

/* 
	On a side note, SQL Server drops a temporary table automatically when you close the connection that created it.
*/

-- Select the data from the new table.
SELECT * FROM MyData


-- We can also export to an Excel File. It's almost identical to importing except now we do an INSERT INTO.
INSERT INTO OPENROWSET('Microsoft.ACE.OLEDB.12.0','Excel 12.0;Database=C:\Users\Alex\OneDrive\Desktop\MyExport.xlsx', 'SELECT * FROM [Data$]') SELECT * FROM MyData
