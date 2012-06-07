/*
 * Grid Generator by Jean-Louis PERSAT @TheCodingFrog
 * Based on Can Berk Güder's 960gs.jsx
 * Copyright (c) 2012 Can Berk Güder
 */

#target photoshop

var doc;

var resList = [['1024',1024],
	['1280',1280]];

var defaultRes = 0;
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

			if (gutterWidth > 0)
			{
				for(var i = 0; i < numColumns; i++)
				{
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

			while(y < docHeight)
			{
				doc.guides.add(Direction.HORIZONTAL, UnitValue(y, 'px'));
				y += rowHeight;
			}
		}
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

main();