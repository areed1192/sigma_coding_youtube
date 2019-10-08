/*

  UNDERSTANDING THE Show Advanced Options
  ------------------------------------------------------------------------------------------------------------------
  Some configuration options, such as affinity mask and recovery interval, are designated as advanced options. By 
  default, these options are not available for viewing and changing. To make them available, set the ShowAdvancedOptions 
  configuration option to 1.

*/

EXEC sp_configure 'show advanced options', 1
RECONFIGURE
GO

/*

  UNDERSTANDING THE OLE Automation Procedue
  ------------------------------------------------------------------------------------------------------------------
  Use the Ole Automation Procedures option to specify whether OLE Automation objects can be instantiated within 
  Transact-SQL batches. This option can also be configured using the Policy-Based Management or the sp_configure stored 
  procedure. The Ole Automation Procedures option can be set to the following values.

  Value: 0
  Definition: OLE Automation Procedures are disabled. Default for new instances of SQL Server.

  Value: 1
  Definition: OLE Automation Procedures are enabled.
  
  When OLE Automation Procedures are enabled, a call to sp_OACreate will start the OLE shared execution environment. The current 
  value of the Ole Automation Procedures option can be viewed and changed by using the sp_configure system stored procedure.

*/

EXEC sp_configure 'Ole Automation Procedures', 1
RECONFIGURE
GO;



-- Variable declaration related to the Object.
DECLARE @token INT;
DECLARE @ret INT;

-- Variable declaration related to the Request.
DECLARE @url NVARCHAR(MAX);
DECLARE @authHeader NVARCHAR(64);
DECLARE @contentType NVARCHAR(64);
DECLARE @apiKey NVARCHAR(32);

-- Variable declaration related to the JSON string.
DECLARE @json AS TABLE(Json_Table NVARCHAR(MAX))

-- Set Authentications
--SET @authHeader = 'BASIC 0123456789ABCDEF0123456789ABCDEF';
--SET @contentType = 'application/x-www-form-urlencoded';

-- Set the API Key, I'm just grabbing it from another table in my Database.
SET @apiKey = (SELECT api_key FROM [SigmaCodingDatabase].[dbo].[API_Services] WHERE service_name = 'Alpha Vantage')

-- Define the URL
SET @url = 'https://www.alphavantage.co/query?function=TIME_SERIES_INTRADAY&symbol=MSFT&interval=5min&datatype=json&apikey=' + @apikey

-- This creates the new object.
EXEC @ret = sp_OACreate 'MSXML2.XMLHTTP', @token OUT;
IF @ret <> 0 RAISERROR('Unable to open HTTP connection.', 10, 1);

-- This calls the necessary methods.
EXEC @ret = sp_OAMethod @token, 'open', NULL, 'GET', @url, 'false';
--EXEC @ret = sp_OAMethod @token, 'setRequestHeader', NULL, 'Authentication', @authHeader;
--EXEC @ret = sp_OAMethod @token, 'setRequestHeader', NULL, 'Content-type', @contentType;
EXEC @ret = sp_OAMethod @token, 'send'

-- Grab the responseText property, and insert the JSON string into a table temporarily. This is very important, if you don't do this step you'll run into problems.
INSERT into @json (Json_Table) EXEC sp_OAGetProperty @token, 'responseText'

-- Select the JSON string from the Table we just inserted it into. You'll also be able to see the entire string with this statement.
SELECT * FROM @json


/*

	Okay, at this point we have a JSON string we can begin to parse. To parse this JSON string we will use the OPENJSON() function used in T-SQL.
	From here, we will begin unraveling the JSON one level at a time. This is often the most confusing part as JSON strings can be extremely nested.
	The way to approach this is to start from the TOP of the string and work your way in.

	Additionally, we will leverage the CROSS APPLY function in order to merge all the results of our parsing to a single table. When looking at the code
	below it might look a little backwards. For example, the top SELECT statement is referencing all the columns we unpacked and each individual WITH
	statment below it seems to be referencing a certain scetion of the JSON string.

*/


-- Display all the data we just parsed, keep in mind you can negate certain columns we parsed. There is no requirement to display all the columns.
SELECT 

	[Time_Series_Metadata].[2. Symbol] AS Ticker_Symbol,
	[Time_Series_Metadata].[4. Interval] AS Time_Interval,
	[Time_Series_Date].[key] AS Date_ID,
	[PriceData].[1. open] AS Open_Price,
    [PriceData].[2. high] AS High_Price,
    [PriceData].[3. low] AS Low_Price,
    [PriceData].[4. close] AS Close_Price,
    [PriceData].[5. volume] AS Voume

FROM OPENJSON((SELECT * FROM @json))  -- USE OPENJSON to begin the parse.

-- At the highest level we have two parts, the `Meta Data` and the `Time Series (5min)` dictionaries.
WITH (
	[Time Series (5min)] NVARCHAR(MAX) AS JSON,
	[Meta Data] NVARCHAR(MAX) AS JSON
) AS  MetaData


/*
	Okay at this point we have two pieces of the bigger document. Let's start with the `Meta Data` portion 
	because that's easiest to parse.
*/

-- Parse the Metadata
CROSS APPLY OPENJSON([MetaData].[Meta Data])
WITH(
    [1. Information] NVARCHAR(MAX),
    [2. Symbol] NVARCHAR(MAX),
    [3. Last Refreshed"] NVARCHAR(MAX),
    [4. Interval] NVARCHAR(MAX),
    [5. Output Size] NVARCHAR(MAX),
    [6. Time Zone] NVARCHAR(MAX)
) AS Time_Series_Metadata


/*
	With the `Meta Data` out of the way we can proceed to the `Time Series (5 min)` portion. This is
	a little challenging as each key inside this dictionary of dictionaries is unique. They are timestamps.
	Normally, we just would specify the portion we want by defining the section `Time Series (5 min)` followed
	by the date. In essence it would look like this:

		`Time Series (5 min)`.`20190101110012`

	The problem is we could have a 100 of these dictionaries each one having a different time stamp. It's not
	realistic to write each one out. What we could do is first grab all the all keys and then grab all the values
	for each key using CROSS APPLY.

*/


-- The tricky part, we need to go through each time series JSON array, and grab the key, value, and type.
CROSS APPLY OPENJSON([MetaData].[Time Series (5min)]) AS Time_Series_Date

-- Then take ONLY THE VALUE part and parse it further, as this is where the data is.
CROSS APPLY OPENJSON([Time_Series_Date].[value])
WITH(
	
    [1. open] NVARCHAR(MAX),
    [2. high] NVARCHAR(MAX),
    [3. low] NVARCHAR(MAX),
    [4. close] NVARCHAR(MAX),
    [5. volume] NVARCHAR(MAX)

) AS PriceData
