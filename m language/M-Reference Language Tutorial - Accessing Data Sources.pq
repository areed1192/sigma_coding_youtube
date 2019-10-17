
/* 
    
    Working with data sources.
    ----------------------------------------
    Inside Power Query and Power BI you'll be working with a variety of data sources, so it's important to know how to take your different files
    and load them into these applications in the correct way. Depending on the type of file it is, the steps you take could be different, but 
    we will cover a few common scenarios for loading data.

*/



let

    // The most basic thing we can do is grab all the BINARY CONTENT of a file.
    myCSVFileContent = File.Contents("C:\Users\Alex\OneDrive\Desktop\korea_data.csv"),

    // Depending on the type of file it's coming from, we then pass it through the propert document function. In this case a Csv.Document function.
    myCsvDocument = Csv.Document(myCSVFileContent, [Delimiter=",", Columns=11, Encoding=TextEncoding.Utf8, QuoteStyle=QuoteStyle.None]),

    /*
        Normally you'll see the two steps up above in a single command, where the "File.Contents" function is found inside the document function.

        Source = Csv.Document(File.Contents("C:\Users\Alex\OneDrive\Desktop\korea_data.csv"),[Delimiter=",", Columns=11, Encoding=65001, QuoteStyle=QuoteStyle.None])
    */

    //Let's grab an Excel workbook this time, first grab the content.
    myExcelFileContent = File.Contents("C:\Users\Alex\OneDrive\Growth - Tutorial Videos\Lessons - VBA\balance_sheet_test.xlsx"),

    // Then pass it through the Excel.Workbook function. This particular document does not use a record to house paramaters.
    myExcelFileDocument = Excel.Workbook(myExcelFileContent, true, true),

    /*
        If you want the current workbook you're in, then do the following.

        Source = Excel.CurrentWorkbook()

        No arguments are needed.
    */


    // However, content comes in many forms, we also have web pages. First we grab the web content.
    myWebContent = Web.Contents("https://en.wikipedia.org/wiki/Star_Wars", [Timeout=#duration(0, 1, 40, 0)]),

    // Then pass it through the Web.Page function. Again, normally, these two steps are merged into one.
    myWebPageContent = Web.Page(myWebPageContent),



    // Let's grab an acccess database. Because it's technially a file, we use the File.Contents first.
    myAccessDatabaseContent = File.Contents("C:\Users\Alex\OneDrive\Career - Work Items\Petco\Financial_Analyst\Petco_Financial_Data.accdb"),

    // Then pass it through the proper document.
    myAccessDatabase =  Access.Database(myAccessDatabaseContent ,[CreateNavigationProperties=true]),

    // From here a new table is returned that has a structural representation of the database. If I want to grab a specific table I would do the following.
    myAccessTable = myAccessDatabase{[Item = "ACTUALS_CAPITAL"]}[Data], // You can read this as return the row that has the column "Item" equal to "ACTUALS_CAPITAL" and then grab the "Data" column from that row.


    // Let's grab an SQL Database, here I specify a Server, a Database, and set the HierarchicalNavigation to true and pass through a query.
    mySQLDatabase= Sql.Database("DESKTOP-MOA8O6M\SQLEXPRESS", "stock_database", [HierarchicalNavigation=true, Query="SELECT * FROM [dbo].[td_service_data]"]),

    // Let's grab all the content of a folder. Where each row is either a file (binary) or folder (table)
    myFolderContents = Folder.Contents("C:\Users\Alex\OneDrive\Growth - Tutorial Videos"),

    // Let's grab all the content of a JSON document.
    myJsonDocument = Json.Document(File.Contents("C:\Users\Alex\OneDrive\Desktop\MyJsonDocument.json")),

    // Doesn't seem to be released yet? 
    myHTMLTable = Html.Table("<a href=""/test.html"">Test</a>", {{"Link", "a", each [Attributes][href]}}),

    // Grab all the Tables in PDF file
    myPDFTables = Pdf.Tables(File.Contents("C:\Users\Alex\OneDrive\Desktop\Linear Regression Model.pdf"))

    // Let's wrap things up and make an API request.
    myApiRequest = Json.Document(Web.Contents("https://www.alphavantage.co/query?function=TIME_SERIES_INTRADAY&symbol=MSFT&interval=5min&apikey=demo"))


in
    myApiRequest 

    