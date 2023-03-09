
//Prompt users to select a file to open
var myFile = File.openDialog('Please select an excel file to open', "*.xlsx", false);
if (myFile != null) {
    //Do something with the file
    alert( "You selected: " + myFile.fsName);
} else {
    //No file was selected
    alert("No file was selected.");
}

//Prompt user 'Would you like to cancel the script?'
var myResponse = confirm("Would you like to cancel the script?", false, "Cancel Script");
if (myResponse == true) {
    //User clicked 'OK'
    alert("The script was canceled.");
} else {
    //User clicked 'Cancel'
    alert("The script will continue.");
    //stop the script
    exit();
}

// Prompt users to select a directory to which to add files
var destinationFolder = Folder.selectDialog( 'Please select the folder in which new files will be saved' );

//Create a new text frame in the document
var questionTextFrame = myDocument.textFrames.add();

//Set the text frame's position and size
questionTextFrame.geometricBounds = [ 72, 72, 144, 540 ];

//Define the paragraph style as a variable

var questionParagraphStyle = myDocument.paragraphStyles.itemByName("Question Text");
var answerParagraphStyle = myDocument.paragraphStyles.itemByName("body text");

//Loop through the rows of the excel file, add content from Column A to a new text frame using the questionParagraphStyle, and add content from Column B to a new text frame using the answerParagraphStyle
for (var i = 0; i < myFile.rows.length; i++) {
    questionTextFrame.contents = myFile.rows[i].cells[0].contents;
    questionTextFrame.paragraphs[0].applyParagraphStyle(questionParagraphStyle);
    questionTextFrame.contents = myFile.rows[i].cells[1].contents;
    questionTextFrame.paragraphs[0].applyParagraphStyle(answerParagraphStyle);
    //Check if the cell value is a file path with page range information
    if (myFile.rows[i].cells[2].contents != "") {
        //Split the cell value into an array of file path and page range
        var fileArray = myFile.rows[i].cells[2].contents.split("|");
        //Check if the file exists
        if (File(fileArray[0]).exists) {
            //open the file indicated in the cell value
            var myFile = app.open(File(fileArray[0]));
            //Loop through the pages in the file
            for (var j = 0; j < myFile.pages.length; j++) {
                //Check if the page is in the page range
                if (j >= fileArray[1].split("-")[0] && j <= fileArray[1].split("-")[1]) {
                    //Duplicate the page into the document
                    myFile.pages[j].duplicate(LocationOptions.AT_END, myDocument);
                }
            }
            
        }
    }


}

