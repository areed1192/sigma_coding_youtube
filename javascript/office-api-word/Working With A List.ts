$("#insert-controls").click(() => tryCatch(insertList));
$("#setup").click(() => tryCatch(setup));
$("#format").click(() => tryCatch(formatList));

async function insertList(){
  await Word.run(async (context) => {

    // grab all the paragraphs
    let paragraph2 = context.document.body.paragraphs;
    paragraph2.load('items');
    await context.sync();

    // Indicates new list to be started in the second paragraph
    let list = paragraph2.items[1].startNewList();

    // insert the list at the start location
    let myFirstItem = list.insertParagraph("My 0 item at start", Word.InsertLocation.start)

    // insert at the end of our list
    list.insertParagraph("My 1 item at end", Word.InsertLocation.end).listItem.level = 3

    // insert at before location, not part of our list.
    list.insertParagraph("My 2 item at before", Word.InsertLocation.before)

    // insert at after location, not part of our list.
    list.insertParagraph("My 3 item at after", Word.InsertLocation.after)

    // delete the first item in our list.
    myFirstItem.delete();

    await context.sync();
  });
}


async function formatList(){
  await Word.run(async(context)=> {

    // grab the lists collection
    let wordLists = context.document.body.lists

    // grab the first list
    let firstList = wordLists.getFirstOrNullObject();

    // set the bullet design for the first level
    firstList.setLevelBullet(0, Word.ListBullet.hollow);

    // set indent level for the first level to 50 points, 20 points for images
    firstList.setLevelIndents(0, 50, 20);

    await context.sync();

  })
}

async function setup() {
  await Word.run(async (context) => {
    context.document.body.clear();
    context.document.body.insertParagraph(
      "Themes and styles also help keep your document coordinated. When you click design and choose a new Theme, the pictures, charts, and SmartArt graphics change to match your new theme. When you apply styles, your headings change to match the new theme. ",
      "Start"
    );
    context.document.body.insertParagraph(
      "Save time in Word with new buttons that show up where you need them. To change the way a picture fits in your document, click it and a button for layout options appears next to it. When you work on a table, click where you want to add a row or a column, and then click the plus sign. ",
      "Start"
    );
    context.document.body.insertParagraph(
      "Video provides a powerful way to help you prove your point. When you click Online Video, you can paste in the embed code for the video you want to add. You can also type a keyword to search online for the video that best fits your document.",
      "Start"
    );
    context.document.body.paragraphs
      .getLast()
      .insertText(
        "To make your document look professionally produced, Word provides header, footer, cover page, and text box designs that complement each other. For example, you can add a matching cover page, header, and sidebar. Click Insert and then choose the elements you want from the different galleries. ",
        "Replace"
      );
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
