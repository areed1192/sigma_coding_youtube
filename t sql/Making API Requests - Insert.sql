DECLARE @token INT;
DECLARE @ret INT;
DECLARE @url NVARCHAR(MAX);
DECLARE @apiKey NVARCHAR(32);
DECLARE @json AS TABLE(Json_Table NVARCHAR(MAX))

-- Set the API Key, I'm just grabbing it from another table in my Database.
SET @apiKey = (SELECT api_key FROM [SigmaCodingDatabase].[dbo].[API_Services] WHERE service_name = 'Alpha Vantage')
SET @url = 'https://www.alphavantage.co/query?function=TIME_SERIES_INTRADAY&symbol=MSFT&interval=5min&datatype=json&apikey=' + @apikey

-- This creates the new object.
EXEC @ret = sp_OACreate 'MSXML2.XMLHTTP', @token OUT;
IF @ret <> 0 RAISERROR('Unable to open HTTP connection.', 10, 1);

-- This calls the necessary methods.
EXEC @ret = sp_OAMethod @token, 'open', NULL, 'GET', @url, 'false';
EXEC @ret = sp_OAMethod @token, 'send'

-- Grab the responseText property, and insert the JSON string into a table temporarily. This is very important, if you don't do this step you'll run into problems.
INSERT into @json (Json_Table) EXEC sp_OAGetProperty @token, 'responseText'

-- What this section does is take the same query from the last tutorial, `Making API Requests` and does an insert into a Non-Existing table.
SELECT

	[Time_Series_Metadata].[2. Symbol] AS Ticker_Symbol,
	[Time_Series_Metadata].[4. Interval] AS Time_Interval,
	[Time_Series_Date].[key] AS Date_ID,
	[PriceData].[1. open] AS Open_Price,
    [PriceData].[2. high] AS High_Price,
    [PriceData].[3. low] AS Low_Price,
    [PriceData].[4. close] AS Close_Price,
    [PriceData].[5. volume] AS Voume

-- This is all I have to add to insert into a non-existing table.
INTO [dbo].[JSONData]
FROM OPENJSON((SELECT * FROM @json))

-- At the highest level we have two parts, the `Meta Data` and the `Time Series (5min)` dictionaries.
WITH (
	[Time Series (5min)] NVARCHAR(MAX) AS JSON,
	[Meta Data] NVARCHAR(MAX) AS JSON
) AS  MetaData

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