// enable double clicking from the Macintosh Finder or the Windows Explorer
#target photoshop

// in case we double clicked the file
app.bringToFront();
//app.notifiers.add(“add”, eventFile)
//app.notifiersEnabled = true;

// debug level: 0-2 (0:disable, 1:break on error, 2:break at beginning)
// $.level = 1;
// debugger; // launch debugger on next line
var __docRef;
var __docName;
var __LayersComps;
var re = RegExp(/_v\d+\.psd/gi);

if (app.documents.length == 0)
{
	
}
else
{
	__docRef = app.activeDocument;
	__LayersComps = __docRef.layerComps;
	
	if (__LayersComps.length > 0)
	{
		for (var __j = 0; __j < __LayersComps.length; __j++)
		{
			__docName = __docRef.name.replace(re, "");
			
			if (__LayersComps[__j].name.indexOf("Layer Comp") > -1)
			{
				__LayersComps[__j].name = __docRef.name.replace(re, "_" + (__j + 1));
			}
			else if (__LayersComps[__j].name.indexOf(__docName) == -1)
			{
				__LayersComps[__j].name = __docRef.name.replace(re, "_" + __LayersComps[__j].name);
			}
		}
	}
}


docRef = null;
textColor = null;
newTextLayer = null;