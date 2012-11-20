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
var __LayersComps;
var __LayerComp;

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
			__LayerComp = __LayersComps[__j];
			__LayerComp.apply();
			__LayerComp.recapture();
		}
	}
}


__docRef = null;
__LayersComps = null;
__LayerComp = null;