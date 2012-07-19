/*
 * Grid Generator by Jean-Louis PERSAT @TheCodingFrog
 * Based on Can Berk Güder's 960gs.jsx
 * Copyright (c) 2012 Can Berk Güder
 */

#target photoshop

var doc;

var resList = [['960',960],['1003',1003],['1024',1024],['1280',1280]];

var defaultRes = 1;
var res = defaultRes;

var rsmList = [['2',2],
	['3',3],
	['4',4],
	['5',5],
	['6',6],
	['12',12],
	['16',16],
	['24',24]];

var defaultRsm = 5;
var rsm = defaultRsm;

var w = new Window(
"dialog {\
	text: 'Grid Generator',\
	alignChildren: 'fill',\
	vertical: Panel {\
		text: 'Vertical Guides',\
		alignChildren: 'right',\
		screenRes: Group {\
			st: StaticText { text: 'Screen size:' },\
			et: DropDownList { alignment:'left', characters: 5 }\
		},\
		numColumns: Group {\
			st: StaticText { text: 'Number of columns:' },\
			et: DropDownList { alignment:'left', characters: 5 }\
		},\
		marginWidth: Group {\
			st: StaticText { text: 'Margin width:' },\
			et: EditText { text: '20', characters: 5 }\
		},\
		gutterWidth: Group {\
			st: StaticText { text: 'Gutter width:' },\
			et: EditText { text: '20', characters: 5 }\
		},\
		colWidth: Group {\
			st: StaticText { text: 'Column width:' },\
			et: StaticText { text: '', characters: 5, enabled: false }\
		}\
	},\
	horizontal: Panel {\
		text: 'Horizontal Guides',\
		alignChildren: 'right',\
		rowHeight: Group {\
			st: StaticText { text: 'Row height:' },\
			et: EditText { text: '20', characters: 5 }\
		}\
	},\
	opts: Panel {\
		text: 'Options',\
		alignChildren: 'left',\
		addVertical: Checkbox { text: 'Add vertical guides', value: true },\
		addHorizontal: Checkbox { text: 'Add horizontal guides' },\
		centerHorizontal: Checkbox { text: 'Center horizontally', value: true },\
		createFill: Checkbox { text: 'Create filled columns', value: true },\
		clearGuides: Checkbox { text: 'Clear existing guides', value: false }\
	},\
	buttons: Group {\
		alignment: 'center',\
		okBtn: Button { text: 'OK', properties: { name: 'ok' }},\
		cancelBtn: Button { text: 'Cancel', properties: { name: 'cancel' }}\
	}\
}");

w.buttons.okBtn.onClick = function()
{
	var fillColor = new SolidColor();
	fillColor.rgb.red  = 255;
	fillColor.rgb.green = 0;
	fillColor.rgb.blue = 0;
	
	if(w.opts.addVertical.value)
	{
		var gutterWidth = Number(w.vertical.gutterWidth.et.text);
		var marginWidth = Number(w.vertical.marginWidth.et.text);
		var numColumns = Number(rsmList[rsm-0][1]);
		var colWidth = Math.ceil((resList[res-0][1] - (((numColumns-1) * gutterWidth) + (2 * marginWidth))) / numColumns);
		
		if(colWidth > 0 && numColumns > 0)
		{
			if(w.opts.clearGuides.value)
			{
				clearAllGuides();
			}
			
			var x = 0;
			if(w.opts.centerHorizontal.value)
			{
				var docWidth = doc.width.as('px');
				var gridWidth = (numColumns * colWidth) + ((numColumns-1) * gutterWidth);
				x = Math.ceil((docWidth - gridWidth) / 2);
			}
			else
			{
				x = marginWidth;
			}
			
			//drawStrokes(0, 0, colWidth, doc.height, marginWidth, gutterWidth, "h", app.foregroundColor, numColumns);
			//var cols = doc.activeLayer;
			//cols.name = "Cols";
		
			if (gutterWidth > 0)
			{
				if(w.opts.createFill.value)
				{
					var gridH = doc.artLayers.add("GridH");
					gridH.name = numColumns + " columns grid";
				}
				
				for(var i = 0; i < numColumns; i++)
				{
					if(w.opts.createFill.value)
					{
						doc.selection.select([[x,0],[x,doc.height],[x + colWidth,doc.height],[x + colWidth,0]]);
						doc.selection.fill(fillColor, ColorBlendMode.NORMAL, 100, false);
					}
					
					doc.guides.add(Direction.VERTICAL, UnitValue(x, 'px'));
					x += colWidth;
					doc.guides.add(Direction.VERTICAL, UnitValue(x, 'px'));
					x += gutterWidth;
					
					
				}
			}
			else
			{
				for(var i = 0; i <= numColumns; i++)
				{
					doc.guides.add(Direction.VERTICAL, UnitValue(x, 'px'));
					x += colWidth;
				}
			}
		
		}
	}
	
	if (w.opts.addHorizontal.value)
	{
		var rowHeight = Number(w.horizontal.rowHeight.et.text);

		if(rowHeight > 0)
		{
			var docHeight = doc.height.as('px');
			var y = rowHeight;
			var i = 0
			
			if(w.opts.createFill.value)
			{
				//var gridV = doc.artLayers.add("GridV");
				//gridV.name = numColumns + " columns grid";
			}
			
			while(y < docHeight)
			{
				if(w.opts.createFill.value && i % 2 == 0)
				{
					//doc.selection.select([[0,y],[doc.width,y],[doc.width,y + rowHeight],[0,y + rowHeight]]);
					//doc.selection.fill(fillColor, ColorBlendMode.NORMAL, 40, false);
				}
				i++;
				doc.guides.add(Direction.HORIZONTAL, UnitValue(y, 'px'));
				y += rowHeight;
			}
		}
	}
		
	if(w.opts.createFill.value)
	{
		doc.selection.deselect();
		//var grid = gridV.merge();
		gridH.opacity = 20;
		gridH.positionLocked = true;
		gridH.pixelsLocked = true;
	}

	w.close();
}

w.vertical.numColumns.et.onChange = function()
{
	if ( this.selection != null )
	{
		rsm = this.selection;
		updateColWidth();
	}
}

w.vertical.screenRes.et.onChange = function()
{
	if ( this.selection != null )
	{
		res = this.selection;
		updateColWidth();
	}
}

w.vertical.gutterWidth.et.onChanging = function()
{
	updateColWidth();
}

w.vertical.marginWidth.et.onChanging = function()
{
	updateColWidth();
}

function main()
{
	try
	{
		doc = app.activeDocument;
	}
	catch(e)
	{
		alert('No active document', '960gs', true);
		return;
	}
	
	for (i in resList)
	{
		w.vertical.screenRes.et.add('item',resList[i][0])
	}
	w.vertical.screenRes.et.selection = defaultRes;
	
	for (i in rsmList)
	{
		w.vertical.numColumns.et.add('item',rsmList[i][0])
	}
	w.vertical.numColumns.et.selection = defaultRsm;
	
	updateColWidth();
	w.show();
}

function updateColWidth()
{
	var gutterWidth = Number(w.vertical.gutterWidth.et.text);
	var marginWidth = Number(w.vertical.marginWidth.et.text);
	var numColumns = Number(rsmList[rsm-0][1]);
	var colWidth = (resList[res-0][1] - (((numColumns-1) * gutterWidth) + (2 * marginWidth))) / numColumns;
	w.vertical.colWidth.et.text = Math.ceil(colWidth);
}

function clearAllGuides()
{
	var desc = new ActionDescriptor();
	var ref = new ActionReference();
	ref.putEnumerated( charIDToTypeID( "Gd  " ), charIDToTypeID( "Ordn" ), charIDToTypeID( "Al  " ) );
	desc.putReference( charIDToTypeID( "null" ), ref );
	executeAction( charIDToTypeID( "Dlt " ), desc, DialogModes.NO );
}

function drawStrokes(x, y, w, h, m, g , d, c, n)
{
	var prevColor = app.foregroundColor;
	app.foregroundColor = c;
	
	// =======================================================
	var id2631 = charIDToTypeID( "Mk  " );
    var desc192 = new ActionDescriptor();
    var id2632 = charIDToTypeID( "null" );
	var ref77 = new ActionReference();
	var id2633 = stringIDToTypeID( "contentLayer" );
	ref77.putClass( id2633 );
    desc192.putReference( id2632, ref77 );
    var id2634 = charIDToTypeID( "Usng" );
	var desc193 = new ActionDescriptor();
	var id2635 = charIDToTypeID( "Type" );
	var id2636 = stringIDToTypeID( "solidColorLayer" );
	desc193.putClass( id2635, id2636 );
	var id2637 = charIDToTypeID( "Shp " );
	var desc194 = new ActionDescriptor();
	var id2638 = charIDToTypeID( "Top " );
	var id2639 = charIDToTypeID( "#Pxl" );
	if (d == "h"){
		desc194.putUnitDouble( id2638, id2639, y );
	}
	else {
		desc194.putUnitDouble( id2638, id2639, y+m );
	}
            
	var id2640 = charIDToTypeID( "Left" );
	var id2641 = charIDToTypeID( "#Pxl" );
	if (d == "h"){
		desc194.putUnitDouble( id2640, id2641, x+m );
	}
	else {
		desc194.putUnitDouble( id2640, id2641, x );
	}
            
	var id2642 = charIDToTypeID( "Btom" );
	var id2643 = charIDToTypeID( "#Pxl" );
	if (d == "h"){
		desc194.putUnitDouble( id2642, id2643, y+h );
	}
	else {
		desc194.putUnitDouble( id2642, id2643, y+h+m );
	}
	
	var id2644 = charIDToTypeID( "Rght" );
	var id2645 = charIDToTypeID( "#Pxl" );
	if (d == "h"){
		desc194.putUnitDouble( id2644, id2645, x+w+m );
	}
	else {
		desc194.putUnitDouble( id2644, id2645, x+w );
	}
            
	var id2646 = charIDToTypeID( "Rctn" );
	desc193.putObject( id2637, id2646, desc194 );
    var id2647 = stringIDToTypeID( "contentLayer" );
    desc192.putObject( id2634, id2647, desc193 );
	executeAction( id2631, desc192, DialogModes.NO );

	var it=0;
	if (d == "h"){
		it = n-1
	}
	else {
		it = Math.round(h/(h+g)) - 1;
	}

	for (var i= 0; i < it; i++){

	// =======================================================
		var id2648 = charIDToTypeID( "AddT" );
		var desc195 = new ActionDescriptor();
		var id2649 = charIDToTypeID( "null" );
		var ref78 = new ActionReference();
		var id2650 = charIDToTypeID( "Path" );
		var id2651 = charIDToTypeID( "Ordn" );
		var id2652 = charIDToTypeID( "Trgt" );
		ref78.putEnumerated( id2650, id2651, id2652 );
		desc195.putReference( id2649, ref78 );
		var id2653 = charIDToTypeID( "T   " );
		var desc196 = new ActionDescriptor();
		var id2654 = charIDToTypeID( "Top " );
		var id2655 = charIDToTypeID( "#Pxl" );
		if (d == "h"){
			desc196.putUnitDouble( id2654, id2655, y );
		}
		else{
			desc196.putUnitDouble( id2654, id2655, h*(i+1)+g*(i+1)+m );
		}
		var id2656 = charIDToTypeID( "Left" );
		var id2657 = charIDToTypeID( "#Pxl" );
		if (d == "h"){
			desc196.putUnitDouble( id2656, id2657, ( m + (w*(i+1)) + (g*(i+1))) );
		}
		else{
			desc196.putUnitDouble( id2656, id2657, (x));
		}
		var id2658 = charIDToTypeID( "Btom" );
		var id2659 = charIDToTypeID( "#Pxl" );
		if (d == "h"){
			desc196.putUnitDouble( id2658, id2659, y+h );
		}
		else{
			desc196.putUnitDouble( id2658, id2659, h*(i+1)+g*(i+1)+h+m );
		}
		
		var id2660 = charIDToTypeID( "Rght" );
		var id2661 = charIDToTypeID( "#Pxl" );
		if (d == "h"){
			desc196.putUnitDouble( id2660, id2661, ( (m + (w*(i+1)) + (g*(i+1)))+w ));
		}
		else{
			desc196.putUnitDouble( id2660, id2661, (x+w) );
		}
			
		var id2662 = charIDToTypeID( "Rctn" );
		desc195.putObject( id2653, id2662, desc196 );
		executeAction( id2648, desc195, DialogModes.NO );
	}
	app.foregroundColor = prevColor;
	doc.activeLayer.opacity = 10;

}

function addToSelection(layer){
	var id549 = charIDToTypeID( "slct" );
    var desc107 = new ActionDescriptor();
    var id550 = charIDToTypeID( "null" );
        var ref93 = new ActionReference();
        var id551 = charIDToTypeID( "Lyr " );
        ref93.putName( id551, layer.name );
    desc107.putReference( id550, ref93 );
    var id552 = stringIDToTypeID( "selectionModifier" );
    var id553 = stringIDToTypeID( "selectionModifierType" );
    var id554 = stringIDToTypeID( "addToSelection" );
    desc107.putEnumerated( id552, id553, id554 );
    var id555 = charIDToTypeID( "MkVs" );
    desc107.putBoolean( id555, false );
executeAction( id549, desc107, DialogModes.NO );
}

function convertToSmart(){
	var id583 = stringIDToTypeID( "newPlacedLayer" );
	executeAction( id583, undefined, DialogModes.NO );
}

main();