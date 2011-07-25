// This script makes a copy merged based on selection, calculate the right size
// create a new doc based on the selection size, and paste it into the doc
// written by Allen Jeng, Adobe Systems Inc., Photoshop QE Team

// enable double clicking from the Macintosh Finder or the Windows Explorer
#target photoshop

// in case we double clicked the file
app.bringToFront();

// get the width and height of selection
var calcWidth  = app.activeDocument.selection.bounds[2] - app.activeDocument.selection.bounds[0];
var calcHeight = app.activeDocument.selection.bounds[3] - app.activeDocument.selection.bounds[1];

// get the document resolution
var docResolution = app.activeDocument.resolution;

// copy merged
app.activeDocument.selection.copy(true);

// create a new doc based on selection size
var myNewDoc = app.documents.add(calcWidth, calcHeight, docResolution);

// paste in from clipboard
myNewDoc.paste();

//myNewDoc.export();