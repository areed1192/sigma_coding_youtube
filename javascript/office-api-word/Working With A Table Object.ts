https://appsforoffice.microsoft.com/lib/1/hosted/office.js 


$("#run").click(() => tryCatch(run));

async function run() {
  await Word.run(async (context) => {

    // Grab the document body
    const body = context.document.body;

    // Insert a table into the document body, 7 rows, 4 columns
    const wrdTbl = body.insertTable(7, 4, Word.RangeLocation.start)

    // Grab the rows of the table.
    const wrdRows = wrdTbl.rows

    // Grab the first row
    const row_1 = wrdRows.getFirst()

    // Fill the header row so it's navy blue
    row_1.shadingColor = '#003366'

    // Do some additional formatting
    row_1.preferredHeight = 40
    row_1.font.name = "Roboto Medium"
    row_1.font.size = 12
    row_1.horizontalAlignment = "Centered"
    row_1.verticalAlignment = "Center"

    // Add some values to the header row
    row_1.values = [['Name','Age','Sales','Cost']]

    // Load the items property for the rows collection
    wrdRows.load('items')
    await context.sync();

    // Looping through the rows
    for (var i = 1; i < wrdRows.items.length; i++){

      // Add some values
      wrdRows.items[i].values = [['Alex','70','40000','60000']]

      // Change the color to a dark grey
      wrdRows.items[i].font.color = '#666666'

      // Change the vertical alignment
      wrdRows.items[i].verticalAlignment = Word.VerticalAlignment.center

      // Change the height
      wrdRows.items[i].preferredHeight = 35

      // Grab the cells
      wrdRows.items[i].cells.load('items');
      await context.sync();

      // Set the alignment of certain cells
      wrdRows.items[i].cells.items[1].horizontalAlignment = 'Right'
      wrdRows.items[i].cells.items[2].horizontalAlignment = 'Right'
      wrdRows.items[i].cells.items[3].horizontalAlignment = Word.Alignment.justified

      // Set the cell padding.
      wrdRows.items[i].cells.items[3].setCellPadding(Word.CellPaddingLocation.top, 10)
    }

    // Grab all the borders
    const tblBorders = wrdTbl.getBorder(Word.BorderLocation.all)

    // Change the color to a dark blue
    tblBorders.color = "#000044"

    // Change the width
    tblBorders.width = 1
    
    // Change the border type.
    tblBorders.type = Word.BorderType.wave

    const tblBordersIV = wrdTbl.getBorder(Word.BorderLocation.insideVertical)
    tblBordersIV.type = "None"

    const tblBordersOT = wrdTbl.getBorder(Word.BorderLocation.outside)
    tblBordersOT.type = "None"

    wrdTbl.load('headerRowCount');
    wrdTbl.styleBuiltIn = Word.Style.gridTable6Colorful_Accent1
    wrdTbl.styleLastColumn = true;

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
