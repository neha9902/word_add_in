/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    // document.getElementById("run").onclick = run;
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported("WordApi", "1.3")) {
      console.log("Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.");
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("extract-insights").onclick = extractInsights;
  }
});

// export async function run() {
//   return Word.run(async (context) => {
//     /**
//      * Insert your Word code here
//      */

//     // insert a paragraph at the end of the document.
//     const paragraph = context.document.body.insertParagraph("Hello world", Word.InsertLocation.end);

//     // change the paragraph color to blue.
//     paragraph.font.color = "blue";

//     await context.sync();
//   });
// }

// function downloadPlotCSV() {
//   fetch('https://jsonplaceholder.typicode.com/todos/1').then(function (response) {
//           return response.json();
//       })
//     }



async function extractInsights() {
  console.log("testing")
  await Word.run(async (context) => {
    // TODO1: Queue commands to insert a paragraph into the document.
    const docBody = context.document.body;
    
  

    // let fetchRes = fetch(
    //   "http://localhost:1000/testing");
    //   await fetchRes.then(res =>
    //     res.json()).then(d => {
    //       const docBody = context.document.body;
    //       const myJSON = JSON.stringify(d);
    //       docBody.insertParagraph(
    //           myJSON,
    //           "Start"
    //         );
    //     })
        // console.log(data)
        // console.log(typeof data)
    // y = json.dumps(fetchRes)
    var json_dummy = require('./sample.json'); //(with path)
    console.log(json_dummy)
    console.log(typeof json_dummy)
    for (let key in json_dummy) {
      if (json_dummy.hasOwnProperty(key))
      {
          value = json_dummy[key];
          // console.log(key, value);
          // console.log(typeof key)
          // console.log(typeof value)
          docBody.insertParagraph(
            key+":-  "+value,
            "End"
          );
          const k = docBody.search(key, { matchCase: true, matchWholeWord: true })
          k.load('font');
          await context.sync();
          // console.log(k.text)
          k.items[0].font.color = 'purple';
          // k.items[0].font.highlightColor = '#FFFF00'; //Yellow
          k.items[0].font.bold = true;
          k.items[0].font.size = 18;
      }
  }
    // docBody.insertParagraph(
    //   "hello",
    //   "End"
    // );

    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}
