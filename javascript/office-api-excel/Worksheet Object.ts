  $("#run").click(() => tryCatch(worksheetObject));

  async function worksheetObject() {
    await Excel.run(async (context) => {

      // define our workbook
      const xlWorkbook = context.workbook;

      // define our worksheet
      //const xlWorksheet = xlWorkbook.worksheets.getItem("Workbook Object");
      const xlWorksheet = xlWorkbook.worksheets.getActiveWorksheet();

      // select the used range.
      xlWorksheet.getUsedRange().select();

      // select the last column in the used range.
      xlWorksheet
        .getUsedRange()
        .getLastColumn()
        .select();

      // select the last row in the used range.
      xlWorksheet
        .getUsedRange()
        .getLastRow()
        .select();

      // select a single cell
      //xlWorksheet.getCell(7, 3).select();
      //load a property of our worksheet.
      xlWorksheet.load("name");
      await context.sync();

      // print it out
      console.log(xlWorksheet.name);

      // copy our worksheet
      //xlWorksheet.copy();
      // turn off the gridlines
      xlWorksheet.showGridlines = false;

      await context.sync();
      const xlRange = xlWorksheet.getRange("B15:D17");

      xlRange.load("cellCount");      
      await context.sync();

      for (var cell = 0; cell < xlRange.cellCount; cell++) {

        xlWorksheet.getCell(7, 2).select();
        console.log(cell);

      }

      // Select a sheet using key.
      //xlWorkbook.worksheets.getItem("Range Object").activate();

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