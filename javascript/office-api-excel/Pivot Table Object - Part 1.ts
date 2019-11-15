$("#removeRow").click(() => tryCatch(removeRow));
$("#toggleColumn").click(() => tryCatch(toggleColumn));
$("#changeHierarchyNames").click(() => tryCatch(changeHierarchyNames));
$("#changeLayout").click(() => tryCatch(changeLayout));
$("#setup").click(() => tryCatch(setup));

async function removeRow() {
  await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    // define the hiearchies collection
    const rowHeir = pivotTable.rowHierarchies;

    // check if the PivotTable already has rows
    const farmRow = rowHeir.getItemOrNullObject("Farm");
    const typeRow = rowHeir.getItemOrNullObject("Type");
    const classRow = rowHeir.getItemOrNullObject("Classification");

    // load the object
    rowHeir.load();
    await context.sync();

    if (!classRow.isNullObject) {
      pivotTable.rowHierarchies.remove(classRow);
    } else if (!typeRow.isNullObject) {
      pivotTable.rowHierarchies.remove(typeRow);
    } else if (!farmRow.isNullObject) {
      pivotTable.rowHierarchies.remove(farmRow);
    }

    await context.sync();
  });
}

async function toggleColumn() {
  await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");

    // check if the PivotTable already has a column
    const column = pivotTable.columnHierarchies.getItemOrNullObject("Farm");
    column.load("id");
    await context.sync();

    if (column.isNullObject) {
      // ading the farm column to the column hierarchy automatically removes it from the row hierarchy
      pivotTable.columnHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    } else {
      pivotTable.columnHierarchies.remove(column);
    }

    await context.sync();
  });
}

async function changeHierarchyNames() {
  await Excel.run(async (context) => {
    const dataHierarchies = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales")
      .dataHierarchies;
    dataHierarchies.load("no-properties-needed");
    await context.sync();

    dataHierarchies.items[0].name = "Farm Sales";
    dataHierarchies.items[1].name = "Wholesale";
    await context.sync();
  });
}

async function changeLayout() {
  await Excel.run(async (context) => {
    const pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.layout.load("layoutType");
    await context.sync();

    // cycle between the three layout types
    if (pivotTable.layout.layoutType === "Compact") {
      pivotTable.layout.layoutType = "Outline";
    } else if (pivotTable.layout.layoutType === "Outline") {
      pivotTable.layout.layoutType = "Tabular";
    } else {
      pivotTable.layout.layoutType = "Compact";
    }
    await context.sync();
    console.log("Pivot layout is now " + pivotTable.layout.layoutType);
  });
}
