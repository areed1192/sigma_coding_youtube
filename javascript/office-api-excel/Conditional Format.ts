$("#run").click(() => tryCatch(run));

async function run() {
  await Excel.run(async (context) => {

    // Define the sheet
    const xlSheet = context.workbook.worksheets.getItem("Data");

    /*
      CONDITIONAL FORMAT: COLOR SCALE
    */

    // Define the range of cells
    const salesRange = xlSheet.getRange("C2:C7");

    // Add a Color Scale Conditional Format Object to the range
    const salesCondFormat = salesRange.conditionalFormats.add(Excel.ConditionalFormatType.colorScale)

    // Define our Criteria for our Color Scale
    const salesCriteria = {
      minimum:{
        type: Excel.ConditionalFormatColorCriterionType.lowestValue,
        color: "blue"
      },
      midpoint: {
        formula: "50",
        type: Excel.ConditionalFormatColorCriterionType.percent,
        color: "orange"
      },
      maximum: {
        type: Excel.ConditionalFormatColorCriterionType.highestValue,
        color: "yellow"
      },     
    }

    // Add our criteria to our Color Scale
    salesCondFormat.colorScale.criteria = salesCriteria

    /*
      CONDITIONAL FORMAT: CUSTOM
    */

    // Define the range of cells
    const costRange = xlSheet.getRange("D2:D7");

    // Add a Custom Conditional Format Object to the range
    const costCondFormat = costRange.conditionalFormats.add(Excel.ConditionalFormatType.custom)

    // First define the rule
    costCondFormat.custom.rule.formula = "=IF(D2 > AVERAGE($D$2:$D$7), TRUE)"

    // Define the format: change font to white
    costCondFormat.custom.format.font.color = "white";

    // change the fill to red
    costCondFormat.custom.format.fill.color = "red";

    // change the font to italic
    costCondFormat.custom.format.font.italic = true;

    // change the font to bold
    costCondFormat.custom.format.font.bold = true;

    /*
      CONDITIONAL FORMAT: TEXT
    */

    // Define the range of cells
    const itemRange = xlSheet.getRange("B2:B7");

    // Add a ContainsText Conditional Format Object to the range
    const itemCondFormat = itemRange.conditionalFormats.add(Excel.ConditionalFormatType.containsText)

    // First define the rule
    itemCondFormat.textComparison.rule = {
      
      // Look for the text "New Item"
      text:"New Item",

      // Look for cells that CONTAIN "New Item"
      operator:Excel.ConditionalTextOperator.contains

    }

    // Second define the format
    itemCondFormat.textComparison.format.font.color = "white";
    itemCondFormat.textComparison.format.fill.color = "green";

  /*
    CONDITIONAL FORMAT: DATA BAR
  */

    // Define the range of cells
    const profitRange = xlSheet.getRange("E2:E7");

    // Add a ContainsText Conditional Format Object to the range
    const profitCondFormat = profitRange.conditionalFormats.add(Excel.ConditionalFormatType.dataBar)

    // Specify the axist format
    profitCondFormat.dataBar.axisFormat = Excel.ConditionalDataBarAxisFormat.cellMidPoint

    // Specify the direction of the Data Bar.
    profitCondFormat.dataBar.barDirection = Excel.ConditionalDataBarDirection.leftToRight

    // Change the fill color of the databar
    profitCondFormat.dataBar.positiveFormat.fillColor = "blue"

    // Change the Gradient
    profitCondFormat.dataBar.positiveFormat.gradientFill = false;

    // do update.
    context.application.calculate(Excel.CalculationType.full)

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