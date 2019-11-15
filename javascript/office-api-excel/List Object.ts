
    $("#createtable").click(() => tryCatch(CreateTable));
    $("#adddata").click(() => tryCatch(AddData));
    $("#filter").click(() => tryCatch(Filters));
    $("#clearfilter").click(() => tryCatch(ClearFilters));
    $("#tableselection").click(() => tryCatch(TableSelection));
    $("#jsontable").click(() => tryCatch(JSONToTable));

    async function CreateTable() {

      await Excel.run(async (context) => {

        // define the sheet where the table will live
        let sheet = context.workbook.worksheets.getItem('MyTable');

        // add the table
        let expensesTable = sheet.tables.add('A1:D1', true)

        // create a table name
        expensesTable.name = "ExpenseTable"

        // set the header values
        expensesTable.getHeaderRowRange().values = [["Product","Color","Amount","Price"]];

        // add the data to it
        expensesTable.rows.add(null /*add rows to the end of the table*/, [
          ["Laptop", "Red", "2", "$1200"],
          ["Computer", "Red", "3", "$1400"],
          ["TV", "Blue", "1", "$2700"],
          ["TV", "Blue", "1", "$3300"],
          ["Laptop", "Yellow", "2", "$3500"],
          ["Computer", "Red", "9", "$13500"],
          ["TV", "Blue", "2", "$97000"]
        ]);

        await context.sync();
      });
    }
    async function AddData() {

      await Excel.run(async (context) => {

        // define the sheet where the table will live
        let sheet = context.workbook.worksheets.getItem('MyTable');

        // grab the table
        let expensesTable = sheet.tables.getItem("ExpenseTable");

        // define some new data to add to the table
        let newData = [["Phone", "Purple", "9", "$1000"], ["Phone", "Purple", "9", "$1000"]]

        // add the data to row 7, remember we start at 0 and we don't include the header row.
        expensesTable.rows.add(6, newData);

        // add a new column, with a formula
        expensesTable.columns.add(null, [
          ["Total Revenue"],
          ["=[Amount] * [Price]"],
          ["=[Amount] * [Price]"],
          ["=[Amount] * [Price]"],
          ["=[Amount] * [Price]"],
          ["=[Amount] * [Price]"],
          ["=[Amount] * [Price]"],
          ["=[Amount] * [Price]"],
          ["=[Amount] * [Price]"],
          ["=[Amount] * [Price]"]
        ]);

        await context.sync();
      });
    }

    async function Filters() {
      await Excel.run(async (context) => {
        
        // define the sheet where the table will live
        let sheet = context.workbook.worksheets.getItem('MyTable');

        // grab the table
        let expensesTable = sheet.tables.getItem("ExpenseTable");

        // define a filter for the color column
        let colorFilter = expensesTable.columns.getItem('Color').filter;
        colorFilter.apply({
          filterOn: Excel.FilterOn.values,
          values:['Red','Purple']
        })

        // define a filter for the price column
        let priceFilter = expensesTable.columns.getItem('Price').filter;
        priceFilter.apply({
          filterOn: Excel.FilterOn.topItems,
          criterion1:"2"
        }) 

        // method two with Top Items
        priceFilter.applyTopItemsFilter(2);
        await context.sync();
      });
    }

    async function TableSelection() {
      await Excel.run(async (context) => {

        // define the sheet where the table will live
        let sheet = context.workbook.worksheets.getItem('MyTable');

        // grab the table
        let expensesTable = sheet.tables.getItem("ExpenseTable");

        // grab the header range
        //expensesTable.getHeaderRowRange().select();

        // grab the data body range
        //expensesTable.getDataBodyRange().select();    

        // get the entire table
        // expensesTable.getRange().select();

        // grab a specific column
        expensesTable.columns.getItem("Color").getRange().select();

        // grab a specific row
        expensesTable.rows.getItemAt(2).getRange().select();

        // grab a total row
        expensesTable.getTotalRowRange().select();

        await context.sync();

      });
    }

    async function JSONToTable() {
      await Excel.run(async (context) => {

        // define the sheet where the table will live
        let sheet = context.workbook.worksheets.getItem('MyTable');

        // add the table
        let jsonTable = sheet.tables.add('A20:D20', true)

        // create a table name
        jsonTable.name = "JsonTable"

        // add the headers
        jsonTable.getHeaderRowRange().values = [["Date","Merchant","Category","Amount"]]

        // define some transactions
        var transactions = [
          {
            DATE: "1/1/2017",
            MERCHANT: "The Phone Company",
            CATEGORY: "Communications",
            AMOUNT: "$120"
          },
          {
            DATE: "1/1/2017",
            MERCHANT: "Southridge Video",
            CATEGORY: "Entertainment",
            AMOUNT: "$40"
          }
        ];

        // convert our JSON Object to an array
        var newData = transactions.map((item => 
        [item.DATE, item.MERCHANT,item.CATEGORY, item.AMOUNT]))

        // add the array to the table
        jsonTable.rows.add(null, newData)
        await context.sync();
      });
    }

    async function ClearFilters() {
      await Excel.run(async (context) => {

        // define the sheet where the table will live
        let sheet = context.workbook.worksheets.getItem('MyTable');

        // grab the table
        let expensesTable = sheet.tables.getItem("ExpenseTable");

        // clear the filter
        expensesTable.clearFilters();
        
        await context.sync();
      });
    }

    /** Default helper for invoking an action and handling errors. */
    async function tryCatch(callback) {
      try {
        await callback();
      } catch (error) {
        // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
        console.error(error);
      }
    }