  $("#run").click(() => tryCatch(run));

  async function run() {
    await Word.run(async (context) => {

      // grab the document
      let doc = context.document;

      // insert some text into our document
      doc.body.insertText("This is me inserting some text into my document. \n", "Start")

      // select the content in the document
      doc.body.getRange('Content').select();

      // search for a particular word
      doc.body.search("document",{matchCase:true, matchWholeWord:true}).getFirst().select();

      // define a variable that contains all the results
      let myResults = doc.body.search("document", { matchCase: true, matchWholeWord: true })

      // load the properties
      myResults.load(['items','text']);
      await context.sync();

      // loop through the results
      myResults.items.forEach(function(rng){

        console.log(rng.text)

      })

      // change the style
      doc.body.load(['style']);
      await context.sync();
      
      doc.body.style = 'Heading 1'

      // save it
      doc.save();

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