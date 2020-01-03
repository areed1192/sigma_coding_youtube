-- This will select all the columns in our table.
SELECT [Year]
      ,[Month]
      ,[Items_Sold]
      ,[Price]
      ,[Revenue]
      ,[Cost]
FROM [SigmaCodingDatabase].[dbo].[SalesData]


/**

	PIVOT:

	Definition Provided By Microsoft:
	---------------------------------
	PIVOT rotates a table-valued expression by turning the unique values from "one column" in the expression 
	into "multiple columns" in the output. And PIVOT runs aggregations where they're required on any remaining 
	column values that are wanted in the final output. 

	Like most SQL problems work from the inner query outwards.

	EXAMPLE 1: BASIC PIVOT

**/


SELECT 

-- STEP 7: 
-- Select the columns from the "Sales_Pivot_By_Year" table. I also do additional 
-- data type transformations and gave the column headers new names.
CAST([2008] AS DECIMAL(18, 2)) AS '2008_Avg',
CAST([2009] AS DECIMAL(18, 2)) AS '2009_Avg',
CAST([2010] AS DECIMAL(18, 2)) AS '2010_Avg'

FROM
(
	
	-- STEP 1:
	-- Select the columns from the table we want to pivot. We will call this our Source Data Query.
	SELECT 
		[Year],
		[Revenue]
	FROM [SigmaCodingDatabase].[dbo].[SalesData]

) AS Sales_Data -- STEP 2: Give the source query an Alias name.


-- STEP 3:
-- Use the PIVOT function to specify which values will become column headers 
-- and the aggregation we want to perform. The result of this will become our
-- new "Pivot Table".
PIVOT
(
	
	-- STEP 4:
	-- Let's do an average of the Revenue.
	AVG([Revenue])

	-- STEP 5: 
	-- For the specified years in our [Year] column.
	FOR [Year] in ([2008], [2009], [2010])

) AS Sales_Pivot_By_Year -- STEP 6: Give the "Pivot Table" an alias name.


/**

	PIVOT:

	EXAMPLE 2: PIVOT - MULTIPLE AGGREGATIONS

	When you want to do multiple aggregations on different columns, things become a little more complex. Honestly,
	usually using PIVOT isn't the most optimal solution but in some cases you'll need to do it. Let's explore how
	to aggregate multiple columns and pivot them.

**/


SELECT

-- STEP 10:
-- Select the columns you want to show. Keep in mind they will have NULL values. To fix the null values, use the
-- COALESCE function.

COALESCE([2008_1], 0) AS 'Avg_Rev_2008',
COALESCE([2009_1], 0) AS 'Avg_Rev_2009',
COALESCE([2010_1], 0) AS 'Avg_Rev_2010',
COALESCE([2008_2], 0) AS 'Avg_Items_2008',
COAlESCE([2009_2], 0) AS 'Avg_Items_2009',
COALESCE([2010_2], 0) AS 'Avg_Items_2010'

FROM

(

	-- STEP 3:
	-- Select all the columns again, this time perform an AVG aggreation on the [Revenue] and [Items_Sold]
	-- column. Now, because we performed an aggregation, we need to group the data. Group the data by [Year],
	-- [Year_1], and [Year_2].
	--
	-- For readability I've filtered the rows so that they only include the years we want, in this case
	-- (2008, 2009, 2010).

	SELECT
		[Year],
		[Year_1],
		[Year_2],
		AVG([Revenue]) AS Revenue,
		AVG([Items_Sold]) AS Items_Sold

	FROM
	(

	-- STEP 1:
	-- Select the columns from the table we want to pivot. We will call this our Source Data Query.
	-- Notice here I have copies of the [Year] column, this is because we will need it to do multiple pivots.

		SELECT
			[Year],
			CAST([Year] AS NVARCHAR(5)) + '_1' AS Year_1,
			CAST([Year] AS NVARCHAR(5)) + '_2' AS Year_2,
			[Revenue],
			[Items_Sold]
		FROM [SigmaCodingDatabase].[dbo].[SalesData]

	) AS Source_Data -- STEP 2: Give source data an alias name.

	WHERE [Year] IN (2008, 2009, 2010)
	GROUP BY [Year], [Year_1], [Year_2]

) AS Source_Data_Grouped -- STEP 4: Give this table an alias name.


-- STEP 5:
-- Define the first pivot. Now the pivot requires that we perform an aggregation, however, we can see that
-- up above we already performed that aggregation. That means there is no need to perform it twice. Instead we
-- can "cheat" a little and just sum the results. By summing it we satisfy the requirement of having an aggregation
-- function in the PIVOT function.


PIVOT
(	
	-- Let's just SUM the [Revenue] column.
	SUM([Revenue])

	-- STEP 6: 
	-- Remember that we copied the [Year] column? Well, the reason why is becasue we can't use the same column in TWO different
	-- PIVOT functions. The work around is to copy each column and simply select the rows we want. As a warning this becomes, very
	-- complex as the number of rows you wish to pivot increases.
	FOR [Year_1] in ([2008_1], [2009_1], [2010_1])

) AS Sales_Pivot_By_Year -- STEP 7: Give the 1ST PIVOT TABLE an alias name.


-- STEP 8:
-- Repeat step 7, but this time just change the column we wish to aggregate and make sure to
-- select the other copied column we created.


PIVOT
(
	
	-- Let's just SUM the [Items_Sold] column.
	SUM([Items_Sold])

	-- Again be careful to select the right column!
	FOR [Year_2] in ([2008_2], [2009_2], [2010_2])

) AS Items_Pivot_By_Year -- STEP 9: Give the 2ND PIVOT TABLE an alias name.



