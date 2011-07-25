// enable double clicking from the Macintosh Finder or the Windows Explorer
#target photoshop

// in case we double clicked the file
app.bringToFront();
//app.notifiers.add(“add”, eventFile)
//app.notifiersEnabled = true;

// debug level: 0-2 (0:disable, 1:break on error, 2:break at beginning)
// $.level = 1;
// debugger; // launch debugger on next line
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
			 ProcessCleanLayersName(__Layer, __idx + 1);
		}
		 else if (__Layer.typename == "ArtLayer")
		{
			
		}
		else
		{
			
		}
		app.activeDocument.activeLayer.visible = __isVisible;
	}
	return;
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

class ImageRight
{
	    public function Parse(__name)
        var __mc;
        var __gc;
        var __cc;

        __mc = Regex.Matches(__name, "(\w{2})\-\w{2}\-(DA|€€)\-(.*)", RegexOptions.Multiline And RegexOptions.IgnoreCase)

        If __mc.Count > 0 Then
            'MessageBox.Show(__name)
            __gc = __mc.Item(0).Groups

            __cc = __gc.Item(1).Captures
            __bankcode = __cc(0).Value

            __cc = __gc.Item(3).Captures
            __imagecode = __cc(0).Value

            __isValidCode = True

            Call setURL()
        Else
            __isValidCode = False
            __bank = vbNullString
            __imagecode = vbNullString
        End If
    End Sub
}


docRef = null;
textColor = null;
newTextLayer = null;