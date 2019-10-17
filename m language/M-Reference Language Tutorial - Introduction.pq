
/* 
    
    What is the M Formula Langague?
    ----------------------------------------
    Microsoft Power Query provides a powerful data import experience that encompasses many features. Power Query works with
    Analysis Services, Excel, and Power BI workbooks. A core capability of Power Query is to filter and combine, that is, to mash-up
    data from one or more of a rich collection of supported data sources. Any such data mashup is expressed using the Power Query
    M Formula Language. It's a functional, case sensitive language similar to F#.

    What is a M Formula Language Query?
    ------------------------------------
    A Power Query M formula language query is composed of formula expression steps that create a mashup query. A formula expression 
    can be evaluated (computed), yielding a value. The let expression encapsulates a set of values to be computed, assigned names, 
    and then used in a subsequent expression that follows the in statement. 

    Below is an example of a M Formula Language Query.

*/

    /* 
    Comment Structure
    ------------------------------------ 
    */

    /* Multi
        Line
        Comment */ 

    // Single Line Comment 


let
    Source = Text.Proper("hello world.")
in
    Source



/* 
    
    What is a Primitive Value?
    ----------------------------------------
    A primitive value is single-part value, such as a number, logical, text, or null. 
    A null value can be used to indicate the absence of any data.

    Below is an example of the different primitive values.

*/



let

    myBinary = #binary({0x30, 0x31, 0x32}),
    myDate = #date(2015, 5, 23),
    myDateTime = #datetime(2015, 5, 23, 0, 0 ,0),
    myDateTimeZone = #datetimezone(2015, 5, 23, 0, 0 ,0, -8, 0),
    myDuration = #duration(0, 15, 35, 0),
    myLogical = true,
    myNull = null,
    myNumber = 0, 
    myText = "abc",
    myTime = #time(12, 34, 12)

in
    myTime



/* 
    
    What is a Functional Value?
    ----------------------------------------
    A Function is a value which, when invoked with arguments, produces a new value. Functions are written by 
    listing the function’s parameters in parentheses, followed by the goes-to symbol =>, followed by the expression defining the function. 

    Below is an example of a functional value.

*/


let 
    // Define the function 
    MyFunction = (parameter1, parameter2) => (parameter1 + parameter2) / 2 ,

    // call my function
    Source = MyFunction(2, 4)  
 
in 
    Source

/* 
    
    What are Structured Data Values?
    ----------------------------------------

    List
    --------------
    A List is a zero-based ordered sequence of values enclosed in curly brace characters { }. The curly brace characters { } 
    are also used to retrieve an item from a List by index position. See [List value](#_List_value).

    Power Query M supports an infinite list size, but if a list is written as a literal, the list has a fixed length. 
    For example, {1, 2, 3} has a fixed length of 3.


    Record
    -------------
    A Record is a set of fields. A field is a name/value pair where the name is a text value that is unique within the field’s 
    record. The syntax for record values allows the names to be written without quotes, a form also referred to as identifiers. 
    An identifier can take the following two forms:

        1. identifier_name such as OrderID.
        2. #"identifier name" such as #"Today's data is: ".


    Table
    ------------
    A Table is a set of values organized into named columns and rows. The column type can be implicit or explicit. You can use 
    #table to create a list of column names and list of rows. A Table of values is a List in a List. The curly brace characters { } 
    are also used to retrieve a row from a Table by index position.


*/


let

    myList = {"a", 100, false},
    myListOfNumbers = {1, 2, 3},
    myListOfLists = {
        {1, 2, 3},
        {"a", "b", "c"},
        {true, false, true},
        {"a", 100, false}
    },
    myListOfRecords = {
        [CustomerID = 1, Name = "Bob", Phone = "123-4567"],
        [CustomerID = 2, Name = "Jim", Phone = "987-6543"]
    },
    myListIndexing = {"a", 100, false}{0}, /* returns "a" */
    myListIndexingNested = {
        {1, 2, 3},
        {"a", "b", "c"},
        {true, false, true},
        {"a", 100, false}
    }{0}{1}  /* returns 2 */

in
    myListIndexingNested


let 
    myRecord = [Name = "Alex", ID = 1111, Hometown = "San Diego"],
    myRecordsNested = [Name = "Alex", ID = 1111, Hometown = "San Diego", Order = [ProductID = "ABC", Amount = 100.00]],
    myName = myRecord[Name],
    myAmount = myRecordsNested[Order][Amount]

in
    myAmount


let  
  myTable = #table(   
                    {"OrderID", "CustomerID", "Item", "Price"},   /* My Column Names */
                    {{1, 1, "Fishing rod", 100.00},   /* My Row 1 */
                     {2, 1, "1 lb. worms", 5.00}}     /* My Row 2 */
                   ),

  myTableWithTypes = #table( 
                             type table [OrderID = number, CustomerID = number, Item = text, Price = number], /* My Column Names With Defined Types */ 
                             {{1, 1, "Fishing rod", 100.00},   /* My Row 1 */
                             {2, 1, "1 lb. worms", 5.00}}     /* My Row 2 */
                           ),
  
  myTableRowOne = myTable{0}

in  
  myTableRowOne

/*

TYPE CONVERSION
------------------------------------------------------------------------------------------------------------------------------------------------------

NUMBER
------
Type conversion	                            Description
---------------                             -----------
Number.FromText(text as text) as number	    Returns a number value from a text value.
Number.ToText(number as number) as text	    Returns a text value from a number value.
Number.From(value as any) as number	        Returns a number value from a value.
Int32.From(value as any) as number	        Returns a 32-bit integer number value from the given value.
Int64.From(value as any) as number	        Returns a 64-bit integer number value from the given value.
Single.From(value as any) as number	        Returns a Single number value from the given value.
Double.From(value as any) as number	        Returns a Double number value from the given value.
Decimal.From(value as any) as number	    Returns a Decimal number value from the given value.
Currency.From(value as any) as number	    Returns a Currency number value from the given value.


TEXT
------
Type conversion	                            Description
---------------                             -----------
Text.From(value as any) as text	            Returns the text representation of a number, date, time, datetime, datetimezone, logical, duration or binary value.


LOGICAL
------
Type conversion	                            Description
---------------                             -----------
Logical.FromText(text as text) as logical	Returns a logical value of true or false from a text value.
Logical.ToText(logical as logical) as text	Returns a text value from a logical value.
Logical.From(value as any) as logical	    Returns a logical value from a value.


DATE, TIME, DATETIME, AND DATE TIMEZONE
---------------------------------------
Type conversion	                                                                                    Description
---------------                                                                                     -----------
.FromText(text as text) as date, time, datetime, or datetimezone	                                Returns a date, time, datetime, or datetimezone value from a set of date formats and culture value.
.ToText(date, time, dateTime, or dateTimeZone as date, time, datetime, or datetimezone) as text	    Returns a text value from a date, time, datetime, or datetimezone value.
.From(value as any)	                                                                                Returns a date, time, datetime, or datetimezone value from a value.
.ToRecord(date, time, dateTime, or dateTimeZone as date, time, datetime, or datetimezone)	        Returns a record containing parts of a date, time, datetime, or datetimezone value.


*/



/*

OPERATORS
------------------------------------------------------------------------------------------------------------------------------------------------------


Plus operator (+)
------------------------------
Expression	                            Equals
1 + 2	                                Numeric addition: 3
#time(12,23,0) + #duration(0,0,2,0)	    Time arithmetic: #time(12,25,0)


Combination operator (&)
------------------------------
Function	                Equals
"A" & "BC"	                Text concatenation: "ABC"
{1} & {2, 3}	            List concatenation: {1, 2, 3}
[ a = 1 ] & [ b = 2 ]	    Record merge: [ a = 1, b = 2 ]


List of M operators
------------------------------


COMMON OPERATORS
-----------------
Common operators which apply to null, logical, number, time, date, datetime, datetimezone, duration, text, binary)

Operator	Description
--------    ------------
>	        Greater than
>=	        Greater than or equal
<	        Less than
<=	        Less than or equal
=	        Equal
<>	        Not equal


LOGICAL OPERATORS
-----------------

Operator	Description
--------    ------------
or	        Conditional logical OR
and	        Conditional logical AND
not	        Logical NOT


NUMBER OPERATORS
-----------------

Operator	Description
--------    ------------
+	        Sum
-	        Difference
*	        Product
/	        Quotient
+x	        Unary plus
-x	        Negation


TEXT OPERATORS
-----------------

Operator	Description
--------    ------------
&	        Concatenation


LIST, RECORD, AND TABLE OPERATORS
-----------------

Operator	Description
--------    ------------
=	        Equal
<>	        Not equal
&	        Concatenation
[]	        Access the fields of a record by name.
{}	        Access an item in a list by its zero-based numeric index.


TYPE COMPATIABILITY AND ASSERTION OPERATORS
-----------------

Operator	Description
--------    ------------
is	        The expression x is y returns true if the type of x is compatible with y, and returns false if the type of x is not compatible with y.
as	        The expression x as y asserts that the value x is compatible with y as per the is operator.



DATE OPERATORS
-----------------

Operator	Left Operand	Right Operand	Meaning
---------   ------------    -------------   -------
x + y	    time	        duration	    Date offset by duration
x + y	    duration	    time	        Date offset by duration
x - y	    time	        duration	    Date offset by negated duration
x - y	    time	        time	        Duration between dates
x & y	    date	        time	        Merged datetime


DATETIME OPERATORS
------------------

Operator	Left Operand	Right Operand	Meaning
---------   ------------    -------------   -------
x + y	    datetime	    duration	    Datetime offset by duration
x + y	    duration	    datetime	    Datetime offset by duration
x - y	    datetime	    duration	    Datetime offset by negated duration
x - y	    datetime	    datetime	    Duration between datetimes


DATETIMEZONE OPERATORS
----------------------

Operator	Left Operand	Right Operand	Meaning
---------   ------------    -------------   -------
x + y	    datetimezone	duration	    Datetimezone offset by duration
x + y	    duration	    datetimezone	Datetimezone offset by duration
x - y	    datetimezone	duration	    Datetimezone offset by negated duration
x - y	    datetimezone	datetimezone	Duration between datetimezones


DURATION OPERATORS
------------------

Operator	Left Operand	Right Operand	Meaning
---------   ------------    -------------   -------
x + y	    datetime	    duration	    Datetime offset by duration
x + y	    duration	    datetime	    Datetime offset by duration
x + y	    duration	    duration	    Sum of durations
x - y	    datetime	    duration	    Datetime offset by negated duration
x - y	    datetime	    datetime	    Duration between datetimes
x - y	    duration	    duration	    Difference of durations
x * y	    duration	    number	        N times a duration
x * y	    number  	    duration	    N times a duration
x / y	    duration	    number	        Fraction of a duration

*/