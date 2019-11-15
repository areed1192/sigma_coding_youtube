$("#createWithNames").click(() => tryCatch(createWithNames));
async function createWithNames() {
  await Excel.run(async (context) => {
    // define the sheet that contains the data and the sheet where you want the data to be.
    const worksheetData = context.workbook.worksheets.getItem("Data");
    const worksheetPivt = context.workbook.worksheets.getItem("Pivot");

    const dataRange = worksheetData.getRange("A1:E21");
    const dataTable = worksheetData.tables.getItem("Table1").getRange();

    const pivtRange = worksheetPivt.getRange("A2");
    worksheetPivt.pivotTables.add("Farm Sales", dataTable, pivtRange);

    await context.sync();
  });
}

$("#deletePivot").click(() => tryCatch(deletePivot));
async function deletePivot() {
  await Excel.run(async (context) => {
    context.workbook.worksheets
      .getItem("Pivot")
      .pivotTables.getItem("Farm Sales")
      .delete();

    await context.sync();
  });
}

$("#addRow").click(() => tryCatch(addRow));
async function addRow() {
  await Excel.run(async (context) => {
    const worksheetPivt = context.workbook.worksheets.getItem("Pivot");
    // grab the pivot table
    const pivotTable = worksheetPivt.pivotTables.getItem("Farm Sales");

    // define the hiearchies collection
    const rowHeir = pivotTable.rowHierarchies;

    // check if the PivotTable already has rows
    const farmRow = rowHeir.getItemOrNullObject("Farm");
    const typeRow = rowHeir.getItemOrNullObject("Type");
    const classRow = rowHeir.getItemOrNullObject("Classification");

    // load the object
    rowHeir.load();
    await context.sync();

    // check if its null and if it is then add that row to the pivot table
    // this little setup will allow you to "drill down"
    if (farmRow.isNullObject) {
      rowHeir.add(pivotTable.hierarchies.getItem("Farm"));
    } else if (typeRow.isNullObject) {
      rowHeir.add(pivotTable.hierarchies.getItem("Type"));
    } else if (classRow.isNullObject) {
      rowHeir.add(pivotTable.hierarchies.getItem("Classification"));
    }
    await context.sync();
  });
}

$("#rowDetail").click(() => tryCatch(rowDetails));
async function rowDetails() {
  await Excel.run(async (context) => {
    // you have a hierarchy
    // you have fields in the hierarchy
    // you have items in the fields

    const worksheetPivt = context.workbook.worksheets.getItem("Pivot");

    // grab the pivot table
    const pivotTable = worksheetPivt.pivotTables.getItem("Farm Sales");

    // define the hiearchies collection
    const rowHeir = pivotTable.rowHierarchies;

    // get the count of items in the hierarhcy
    let rowCount = rowHeir.getCount();
    await context.sync();

    // print the details
    console.log(`There are '${rowCount.value}' items in the Pivot Table`);

    // define a Pivot Field in the row hierarchy
    const typeRow = rowHeir.getItemOrNullObject("Type");

    // load the object details
    rowHeir.load();
    typeRow.load(["fields"]);
    let fieldCount = typeRow.fields.getCount();
    await context.sync();

    // display the number of fields in the Type
    console.log(`There is/are '${fieldCount.value}' items in the Type field`);

    // this is confusing but you get field again, then the item from that field
    let typeField = typeRow.fields.getItem("Type");
    let lemonItem = typeField.items.getItem("Lemon");

    lemonItem.load("name");
    await context.sync();

    console.log(`You've selected the '${lemonItem.name}' item in the Pivot Table`);
  });
}

$("#addAggs").click(() => tryCatch(addAggs));
async function addAggs() {
  await Excel.run(async (context) => {
    // grab the pivot table
    const worksheetPivt = context.workbook.worksheets.getItem("Pivot");
    const pivotTable = worksheetPivt.pivotTables.getItem("Farm Sales");

    // // add some data hierarchies, by default these will be sums for numeric values
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold at Farm"));
    pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Crates Sold Wholesale"));

    // load the items for the next step
    pivotTable.dataHierarchies.load("items");
    await context.sync();

    // redfine the aggregation so we can calculate the max and the count
    pivotTable.dataHierarchies.items[0].summarizeBy = Excel.AggregationFunction.count;
    pivotTable.dataHierarchies.items[1].summarizeBy = Excel.AggregationFunction.max;

    await context.sync();
  });
}
