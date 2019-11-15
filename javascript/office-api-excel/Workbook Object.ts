  $("#run1").click(() => tryCatch(workbookObjects));
  $("#run2").click(() => tryCatch(myWorksheetsColl));

  async function workbookObjects() {

    await Excel.run(async (context) => {

      // grab the workbook
      const xlWorkbook = context.workbook;

      // grab the chart
      const xlChart = xlWorkbook.getActiveChart();

      // grab the PivotTable
      //const xlPivotTable = xlWorkbook.pivotTables.getItem('PivotTable1');
      // grab the chart name property and load it.
      //xlPivotTable.load(["name", "dataHierarchies"]);
      // grab the chart name property and load it.
      xlChart.load(["name", "chartType", "id"]);

      // sync it
      await context.sync();

      // print the results
      console.log(`The chart name is ${xlChart.name}`);
      console.log(`The chart type is ${xlChart.chartType}`);
      console.log(`The chart ID is ${xlChart.id}`);

      // print the results
      //console.log(`The Pivot Table name is ${xlPivotTable.name}`);
      //console.log(`The Pivot table has ${xlPivotTable.dataHierarchies} hierarchies`);

    });
  }

  async function myWorksheetsColl() {
    
    await Excel.run(async (context) => {

      // grab the workbook
      const xlWorkbook = context.workbook;

      // Define the worksheets collection
      const xlWorksheets = xlWorkbook.worksheets;

      // define the properties we want to load
      xlWorkbook.load(["properties/"]);

      // define the properties we want to load
      xlWorksheets.load(["name", "items", "id"]);

      // get the count, this requires a load.
      let worksheetCount = xlWorksheets.getCount();

      // sync the pivot table properties, the worksheet properties, and the getCount.
      await context.sync();

      // let's loop through all the worksheets, using a "For Each" loop
      for (let i in xlWorksheets.items) {

        console.log(xlWorksheets.items[i].name);
      }

      // let's loop through all the worksheets using a "For" loop
      for (let i = 0; i < worksheetCount.value; i++) {
        console.log(xlWorksheets.items[i].id);
      }

      // print some properties.
      console.log(xlWorkbook.properties.lastAuthor);
      console.log(xlWorkbook.properties.creationDate);
      console.log(Excel.CalculationState);
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