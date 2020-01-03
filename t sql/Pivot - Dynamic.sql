
/**

	DYNAMIC PIVOT QUERIES:
	
	We've seen that using PIVOT can make the transformation of our data very easy. However, we also, hopefully,
	recognized some challenges with it. For example, it's great but you have to SELECT each column we want to PIVOT.
	With small data sets this isn't an issue. However, for larger or more dynamic data sets this won't work. In this
	section, we will cover DYNAMIC PIVOTING or in other words how to make our queries dynamic so that they PIVOT all
	the values in a particular column.

	GOAL:
	Take the SalesData table and PIVOT all the DISTINCT VALUES in the [Year] column.

	Ideally, we want something like the following:

	Avg_Rev_2008	Avg_Rev_2009	    ***	        Avg_Rev_2018	Avg_Rev_2019
	------------	------------	------------	------------	------------
	  $1000.00		  $2000.00			***			  $2000.00		  $3000.00

**/

-- Let's take a look at our table again!
SELECT [Year]
      ,[Month]
      ,[Items_Sold]
      ,[Price]
      ,[Revenue]
      ,[Cost]
FROM [SigmaCodingDatabase].[dbo].[SalesData]

-- STEP 1: Declare Variables
-- We will need one to store our column names and one to store the entire query we are going to build.
-- These will all be NVARCHAR.
DECLARE @cols AS NVARCHAR(MAX)
DECLARE	@query AS NVARCHAR(MAX)

/**
	
	BUILDING OUR COLUMNS STRING:
	----------------------------
	For the next few steps I'll be going step by step how we need build our query. Now at the end it won't be multiple steps
	like I show below instead it will just be a single query. However, for educational purposes I want to break it out into it's
	different steps so you understand clearly what has to be done.

**/

-- STEP 2: Grab all the Distinct values in the [Year] column.
-- Remember we don't need duplicate values just distinct values.


	-- Here is the distinct query.
	SELECT DISTINCT [Year] FROM [SigmaCodingDatabase].[dbo].[SalesData]


-- STEP 3: Convert Values to Proper Format.
-- Additionally, we need everything to be the proper format when building our query. For example, 2008 should look like
-- this [2008]. Now T-SQL already has a built-in function called QUOTENAME that will wrap each value in "[]". Additionally,
-- each value needs to be sperated by a "," so we will add that component as well.


	-- Here is the distinct query with the addition of QUOTENAME and the "," character.
	SELECT DISTINCT ',' + QUOTENAME([Year]) FROM [SigmaCodingDatabase].[dbo].[SalesData]


-- STEP 4: Building an XML Type.
-- At this point, things get more complex because we need to multiple steps at once. However, the first thing is understanding
-- what we are trying to get. We have our query it's "properly formatted" at least to the point it should be a this stage. What
-- we do from here is take that query and use the "FOR XML" clause.
--
-- The "FOR XML" clause will take a SELECT statement (which returns a ROWSET) and convert it to an XML result set. The "FOR XML"
-- clause has different "MODES" that specify how the result should be returned in this case, we will use the "PATH" mode and set
-- this to blank. I'm sure I confused enough people at this point so let's take it one step at a time.

	
	-- Use FOR XML to return the ROWSET as XML content, in this example let's see what happens when we just specify PATH.
	SELECT DISTINCT ',' + QUOTENAME([Year]) FROM [SigmaCodingDatabase].[dbo].[SalesData] FOR XML PATH

	-- Same query but now set PATH to 'Year'
	SELECT DISTINCT ',' + QUOTENAME([Year]) FROM [SigmaCodingDatabase].[dbo].[SalesData] FOR XML PATH('Year')

	-- Same query but now set PATH to ''. Ultimately this is the one we want.
	SELECT DISTINCT ',' + QUOTENAME([Year]) FROM [SigmaCodingDatabase].[dbo].[SalesData] FOR XML PATH('')


-- Okay hopefully, now see how the XML PATH clause works. It's just wrapping each "row value" in a tag. There is one extra argument
-- we need to specify, "TYPE". XML enables you to optionally request that the result of a FOR XML query be returned as xml data type by 
-- specifying the TYPE directive


-- STEP 5: Access the values.
-- Now the XML Type is great but it's not helpful if we can access the values. However, TSQL makes this easy by providing the
-- value() method. With the value() method you can perform an XQuery against XML and return SQL values. Here are the arguments for
-- the value() method.
--
--		XQuery
--		Is the XQuery expression, a string literal, that retrieves data inside the XML instance. The XQuery must return at most one 
--      value. Otherwise, an error is returned.
--
--		SQLType
--		Is the preferred SQL type, a string literal, to be returned. The return type of this method matches the SQLType parameter. 
--		SQLType cannot be an xml data type, a common language runtime (CLR) user-defined type, image, text, ntext, 
--		or sql_variant data type. SQLType can be an SQL, user-defined data type.
-- 
-- Why are we doing all this you may ask? Well it's to convert the entire XML content to something that TSQL can handle. In this case,
-- we want NVARCHAR.


	-- Access the values of our XML content and conert it to NVARCHARMAX.
    SELECT (SELECT DISTINCT ',' + QUOTENAME([Year]) FROM [SigmaCodingDatabase].[dbo].[SalesData] FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)')


-- STEP 6: USE the STUFF FUNCTION
-- With our XML Type, we can move the next portion, the STUFF Function. The STUFF function inserts a string into another string. 
-- It deletes a specified length of characters in the first string at the start position and then inserts the second string 
-- into the first string at the start position. The only reason we do this step is to remvoe that leading "," from the value that
-- is returned from the XML content. (PROBLEM) >>> ,[2008]

	-- Remove the leading "," by using the STUFF function. In our example we want to insert a '' space. 
	-- Then insert it into the "@cols" variable.
	SET @cols = STUFF((SELECT DISTINCT ',' + QUOTENAME([Year]) FROM [SigmaCodingDatabase].[dbo].[SalesData] FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'),1,1,'')

	-- Let's see what the final result is!
	SELECT @cols
 

/**
	
	BUILDING OUR QUERY STRING:
	----------------------------
	Okay, we have our columns all select that was the hard part, now comes the easy part building the query. For this one,
	we can use our exisiting query, convert it to a string, and just concatenate the pieces.

**/


SET @query = 
'SELECT ' + @cols + ' FROM (SELECT [Year], [Revenue] FROM [SigmaCodingDatabase].[dbo].[SalesData]) 
AS Sales_Data PIVOT (AVG([Revenue]) FOR [Year] in (' + @cols + ')) AS Sales_Pivot_By_Year'

-- Let's see what the final result is!
SELECT @query

/**
	
	EXECUTING OUR QUERY STRING:
	----------------------------
	I told you that part would be easy! The next part is even easier, we just have to execute the query. That's simple,
	just use the built-in "EXECUTE" command. The "EXECUTE" command executes a command string or character string within a 
	Transact-SQL batch, or one of the following modules: system stored procedure, user-defined stored procedure, 
	CLR stored procedure, scalar-valued user-defined function, or extended stored procedure.

**/

-- Execute the final query.
EXECUTE(@query)