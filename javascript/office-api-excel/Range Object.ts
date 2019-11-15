
  $("#run1").click(() => tryCatch(grabbingValues));
  $("#run2").click(() => tryCatch(puttingValues));
  $("#run3").click(() => tryCatch(puttingFormulas));

  async function grabbingValues() {

    await Excel.run(async (context) => {

      // Grab the sheet
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      // Grab a range using Get Range
      let salesRng = sheet.getRange("B2:E7");

      // load the values property & the text properties
      salesRng.load(["values", "text", "formulas", "formulasR1C1"]);      
      await context.sync();

      // let's grab the values
      let myValues = salesRng.values;

      // access the first row
      console.log(myValues[0]);

      // access the first column of the first row
      console.log(myValues[0][0]);

      // loop through the entire array
      // let's start with the rows
      myValues.forEach(function(row, index) {

        // print each row
        console.log(row);

        // then each column in that row
        row.forEach(function(col, index2) {

          // print each column
          console.log(col);

        });
      });

      // stringify the values:
      // the first para is the values
      // the second para is what you want missing values to be
      // the third is how many spaces you want
      console.log(JSON.stringify(salesRng.values, null, 1));
      console.log(JSON.stringify(salesRng.text, null, 1));
      console.log(JSON.stringify(salesRng.formulas, null, 1));

    });
  }

  async function puttingValues() {

    await Excel.run(async (context) => {

      // Grab the sheet
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      // Grab a range using Get Range
      let salesRng = sheet.getRange("B2:E7");

      // load the values property & the text properties
      salesRng.load(["values", "text", "formulas", "formulasR1C1"]);      
      await context.sync();

      // let's grab the values
      let myValues = salesRng.values;

      // Let's loop through the entire range, manipulate the values
      // and store the new values in a different range.
      // First, we need to define an empty array, keep in mind this will be
      // a matrix so it will serve as the outer array
      //
      // ROW 1 [[COL1, COL2, COL3]
      // ROW 2  [COL1, COL2, COL3]
      // ROW 3  [COL1, COL2, COL3]]
      let myOuterArray = [];

      // loop through each row
      myValues.forEach(function(row) {

        // define the inner array
        let myInnerArray = [];

        // then loop through each column
        row.forEach(function(col) {

          // take the current value and add 2 to it or append "2" to the
          // end of it
          myInnerArray.push(col + 2);

        });

        // append the inner array to the outer array
        myOuterArray.push(myInnerArray);

      });

      // define the new sales Range that will house the values.
      let newSalesRange = sheet.getRange("B9:E14");

      // append the values to the new range.
      newSalesRange.values = myOuterArray;

    });
  }

  async function puttingFormulas() {

    await Excel.run(async (context) => {

      // Grab the sheet
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      // Grab a range using Get Range
      let salesRng = sheet.getRange("B2:E7");

      // load the values property & the text properties
      salesRng.load(["values", "text", "formulas", "formulasR1C1"]);
      await context.sync();

      // let's grab the values
      let myValues = salesRng.values;

      // now the previous example, assumed we wanted values. What if we want
      // formulas? Well we have a few different options.
      // formulas - Standard A1:C1
      // formulasLocal - Standard A1:C1, but language specific. For example, in German it would be '=SUMME('A1:C1')'
      // formulasR1C1 - Standard R1C1 which is cell A1.
      // I want R1C1 so that my formulas adjust based on their new position
      let myFormulas = salesRng.formulasR1C1;

      // Again define the outer array
      let myArrayFormula = [];

      // let's start with the rows
      myFormulas.slice(1).forEach(function(row) {

        // define the innery array
        let myInnerArray = [];

        // only modify the second colun
        myInnerArray.push(row[0]);
        myInnerArray.push(row[1] + 2);
        myInnerArray.push(row[2]);
        myInnerArray.push(row[3]);

        // push to the outer array
        myArrayFormula.push(myInnerArray);

      });

      // define new range & append values to it.
      let salesRngFormula = sheet.getRange("B17:E21");
      salesRngFormula.formulas = myArrayFormula;
      
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