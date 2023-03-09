//Define the open inDesign document as a variable
var myRFP = app.activeDocument;

//Prompt users to select a file to open
var rfpMappingFile = File.openDialog('Please select an excel file to open', "*.xlsx", false);
if (rfpMappingFile != null) {
    //Do something with the file
    alert( "You selected: " + rfpMappingFile.fsName);
} else {
    //No file was selected
    alert("No file was selected.");
}

//Prompt user 'Would you like to cancel the script? if no file is selected'
if (rfpMappingFile == null) {
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
}

// Prompt users to select a directory to which to add files
var destinationFolder = Folder.selectDialog( 'Please select the folder in which new files will be saved' );

//Create a new text frame in the document
var questionTextFrame = myRFP.textFrames.add();

//define the size of the text frame (7.5 inches x 10 inches)
var questionTextFrameWidth = 7.5 * 72;
var questionTextFrameHeight = 10 * 72;
