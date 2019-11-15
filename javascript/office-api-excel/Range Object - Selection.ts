  $("#run1").click(() => tryCatch(workingWithRanges));

  async function workingWithRanges() {
    await Excel.run(async (context) => {

      // define the sheet
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      // Grab a range using Get range
      sheet.getRange("A1:C7").select();

      // Grab a range using Get range
      sheet.getRange("myNewRange").select();

      // Grab a range using Indexes, remember it is 0-based
      sheet.getRangeByIndexes(0,0,7,3).select();

      // Grab the used range
      sheet.getUsedRange().select();

      // Grab the used range
      sheet.getUsedRangeOrNullObject().select();

      // Grab the entire worksheet range
      sheet.getRange().select();

      // Define a multi range
      let multiRange = sheet.getRanges('A1:A5, B1:B5');
      multiRange.load('cellCount');

      // Grab a range using Get range
      let myRange = sheet.getRange("A1:C7");
        
        // Get the entire columns
        myRange.getEntireColumn().select();

        // Get the entrie Rows
        myRange.getEntireRow().select();

        // Get Columns After, positive outside the range, negative inside the range
        myRange.getColumnsAfter(-2).select();

        // Get Rows Below
        myRange.getRowsBelow(-1).select();

        // Get Rows Above
        myRange.getRowsAbove(-1).select();
        
        // Get the last Cell
        myRange.getLastCell().select();

        // Get the last row
        myRange.getLastRow().select();

        // Get the last Column
        myRange.getLastColumn().select();

        // Get a specific row, remember 0-based
        myRange.getRow(1).select();

        // Get a specific column, remember 0-based
        myRange.getColumn(0).select();

        // Get Offset Range
        myRange.getOffsetRange(1,1).select();

        // Get Resized Range
        myRange.getResizedRange(2, 3).select();

        // Get the intersection, takes a second range as an argument.
        myRange.getIntersection('C7:E11').select();

        // Get a specific cell in that range.
        myRange.getCell(2,1).select();

        // Get resized range, I like to think the inner one but it's not exact.
        myRange.getAbsoluteResizedRange(2,2).select();

        // Get bounded range, it's like creating a new rectangle.
        myRange.getBoundingRect("E13:E15").select();

      await context.sync();

      console.log(multiRange.cellCount);
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