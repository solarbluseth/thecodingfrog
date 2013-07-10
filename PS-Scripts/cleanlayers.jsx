// enable double clicking from the Macintosh Finder or the Windows Explorer
#target photoshop

// in case we double clicked the file
app.bringToFront();
//app.notifiers.add(“add”, eventFile)
//app.notifiersEnabled = true;

// debug level: 0-2 (0:disable, 1:break on error, 2:break at beginning)
// $.level = 1;
// debugger; // launch debugger on next line

const ns = 'http://www.smartobjectlinks.com/1.0/';
var docRef;
var re = RegExp(/(\+)+\s*/gi);

if (app.documents.length == 0)
{
	
}
else
{
	docRef = app.activeDocument;
	ProcessCleanLayersName(docRef, 1);
}
//__docRef = __appRef.ActiveDocument

function ProcessCleanLayersName(__ActiveDocument, __idx)
{
	var __Layers;
	var __Layer;
	var __isVisible;
	var __j;
	var __FistLayerText = "";
	var __reg = new RegExp("Group\s*\d*","gim");
	
	__Layers = __ActiveDocument.layers;
	if (__Layers == undefined)
		return;
		
	if (__Layers.length < 1)
		return;
	
	for (__j = 0; __j < __Layers.length; __j++)
	{
		__Layer = __Layers[__j];
		__isVisible = __Layer.visible;
		app.activeDocument.activeLayer = __Layer;
		if (__Layer.typename == "LayerSet")
		{
			__Layer.name = repeat("+", __idx) + " " + __Layer.name.replace(re, "");
			var __NewLayerName = ProcessCleanLayersName(__Layer, __idx + 1);
			if (__NewLayerName != "" && __reg.test(__Layer.name) == true)
			{
				__Layer.name = repeat("+", __idx) + " " + __NewLayerName;
			}
		}
		else if (__Layer.typename == "ArtLayer")
		{
			if (__Layer.kind == LayerKind.SMARTOBJECT)
			{
				try
				{
					var xmp = new XMPMeta( __Layer.xmpMetadata.rawData );
					if (xmp.getProperty(ns,"linkDate").toString() != "")
					{
						__Layer.name = repeat("+", __idx) + " " + __Layer.name.replace(re, "");
						layerColour("violet");
					}
				}
				catch(e)
				{
				}
			}
			else if (__Layer.kind == LayerKind.TEXT)
			{
				if (__FistLayerText == "")
					__FistLayerText = __Layer.name;
			}
		}
		else
		{
			
		}
		
		app.activeDocument.activeLayer = __Layer;
		app.activeDocument.activeLayer.visible = __isVisible;
	}
	return __FistLayerText;
}

function repeat(pattern, count)
{
	if (count < 1)
		return '';
	var result = '';
	while (count > 0)
	{
		if (count & 1)
			result += pattern;
		count >>= 1, pattern += pattern;
	};
	return result;
};

function layerColour(colour)
{
    switch (colour.toLocaleLowerCase())
	{
		case 'red': colour = 'Rd  '; break;
		case 'orange' : colour = 'Orng'; break;
		case 'yellow' : colour = 'Ylw '; break;
		case 'yellow' : colour = 'Ylw '; break;
		case 'green' : colour = 'Grn '; break;
		case 'blue' : colour = 'Bl  '; break;
		case 'violet' : colour = 'Vlt '; break;
		case 'gray' : colour = 'Gry '; break;
		case 'none' : colour = 'None'; break;
		default : colour = 'None'; break;
    }
	
    var desc = new ActionDescriptor();
    var ref = new ActionReference();
    ref.putEnumerated( charIDToTypeID('Lyr '), charIDToTypeID('Ordn'), charIDToTypeID('Trgt') );
    desc.putReference( charIDToTypeID('null'), ref );
    var desc2 = new ActionDescriptor();
    desc2.putEnumerated( charIDToTypeID('Clr '), charIDToTypeID('Clr '), charIDToTypeID(colour) );
    desc.putObject( charIDToTypeID('T   '), charIDToTypeID('Lyr '), desc2 );
    executeAction( charIDToTypeID('setd'), desc, DialogModes.NO );
};


docRef = null;
textColor = null;
newTextLayer = null;