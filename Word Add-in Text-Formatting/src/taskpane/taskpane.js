/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";
import { base64Image } from "./base64Image";

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported("WordApi", "1.3")) {
      // eslint-disable-next-line no-undef
      console.log("Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.");
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("insert-paragraph").onclick = insertParagraph;
    document.getElementById("apply-style").onclick = appyStyle;
    document.getElementById("apply-custom-style").onclick = applyCustomStyle;
    document.getElementById("change-font").onclick = changeFont;
    document.getElementById("insert-text-into-range").onclick = insertTextIntoRange;
    document.getElementById("replace-text").onclick = replaceText;
    document.getElementById("insert-text-outside-range").onclick = insertTextBeforeRange;
    document.getElementById("insert-image").onclick = insertImage;
    document.getElementById("insert-html").onclick = insertHTML;
    document.getElementById("add-table").onclick = addTable;
    document.getElementById("create-content-control").onclick = createContentControl;
    document.getElementById("replace-content-in-control").onclick = replaceContentInControl;
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

const replaceContentInControl = () => {
  Word.run((context) => {
    var serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
    serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", "Replace");
    return context.sync();
  }).catch(handleError);
};
const createContentControl = () => {
  Word.run((context) => {
    var doc = context.document;
    var serviceNameRange = doc.getSelection();
    var serviceNameContentControl = serviceNameRange.insertContentControl();
    serviceNameContentControl.title = "Service Name";
    serviceNameContentControl.tag = "serviceName";
    serviceNameContentControl.appearance = "Tags";
    serviceNameContentControl.color = "blue";
    return context.sync();
  }).catch(handleError);
};

const insertImage = () => {
  Word.run((context) => {
    context.document.body.insertInlinePictureFromBase64(base64Image, "End");
    return context.sync();
  }).catch(handleError);
};

const insertHTML = () => {
  Word.run((context) => {
    var blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", "After");
    blankParagraph.insertHtml(
      '<p style="font-family: verdana;color: red">Inserted HTML.</p><p>Another paragraph</p>',
      "End"
    );
    return context.sync();
  }).catch(handleError);
};

const addTable = () => {
  console.log("Beginning Insert Table function");
  Word.run((context) => {
    var secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    var tableData = [
      ["Name", "ID", "Birth City"],
      ["Bob", "434", "Chicago"],
      ["Sue", "719", "Havana"],
    ];
    secondParagraph.insertTable(3, 3, "After", tableData);

    return context.sync();
  }).catch(handleError);
};

function insertParagraph() {
  Word.run(function (context) {
    var docBody = context.document.body;
    docBody.insertParagraph(
      "Office has several versions, including Office 2016, Microsoft 365 Click-to-Run, and Office on the web.",
      "Start"
    );

    return context.sync();
  }).catch(function (error) {
    // eslint-disable-next-line no-undef
    handleError(error);
  });
}

function handleError(error) {
  console.log("Error: " + error);
  if (error instanceof OfficeExtension.Error) {
    // eslint-disable-next-line no-undef
    console.log("Debug info: " + JSON.stringify(error.debugInfo));
  }
}

function appyStyle() {
  Word.run((context) => {
    var firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.styleBuiltIn = Word.Style.intenseReference;
    return context.sync();
  }).catch((error) => {
    handleError(error);
  });
}

function applyCustomStyle() {
  Word.run((context) => {
    insertParagraph();
    var lastParagraph = context.document.body.paragraphs.getLast();
    lastParagraph.font.set({
      bold: true,
      name: "Arial Rounded MT Bold",
      size: 22,
    });
    return context.sync();
  }).catch((error) => handleError(error));
}

function changeFont() {
  Word.run((context) => {
    var secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    secondParagraph.font.set({
      name: "Arial",
      bold: true,
      size: 18,
    });
    return context.sync();
  }).catch((error) => handleError(error));
}

function insertTextIntoRange() {
  Word.run((context) => {
    // Queue commands to insert text into a selected range.
    var doc = context.document;
    var originalRange = doc.getSelection();
    originalRange.insertText(" (C2R)", "End");

    //Load the text of the range and sync so that the
    //        current range text can be read.

    //step 1: queue a command to load(fetch) the properties the code needs to read.
    originalRange.load("text");
    //step 2: call the context's sync method to send the queued command to the document
    //for execution and return the requested information
    return context
      .sync()
      .then(() => {
        //step 3: B/C the sync() method is asynchronous,ensure it has completed before your code calls the properties that were fetched.
        doc.body.insertParagraph("Original range: " + originalRange.text, "End");
      })
      .then(context.sync);
  }).catch((error) => handleError(error));
}

function replaceText() {
  Word.run((context) => {
    var doc = context.document;
    var originalRange = doc.getSelection();
    originalRange.insertText("many", "Replace");
    return context.sync();
  }).catch(handleError);
}

function insertTextBeforeRange() {
  Word.run((context) => {
    var doc = context.document;
    var originalRange = doc.getSelection();
    originalRange.insertText("Office 2019, ", "Before");
    originalRange.load();
    return context
      .sync()
      .then(() => {
        doc.body.insertParagraph("Current text of original range: " + originalRange.text, "End");
      })
      .then(context.sync);
  }).catch(handleError);
}
