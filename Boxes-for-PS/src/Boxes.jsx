﻿/**************************************************************************	ADOBE SYSTEMS INCORPORATED	 Copyright 2008 Adobe Systems Incorporated	 All Rights Reserved.	NOTICE:  Adobe permits you to use, modify, and distribute this file	in accordance with the terms of the Adobe license agreement accompanying	it.  If you have received this file from a source other than Adobe, then	your use, modification, or distribution of it requires the prior written	permission of Adobe.**************************************************************************//**	Name: Boxes.jsx	@author Jean-Louis PERSAT @TheCodingFrog	@fileoverview ...*/#target photoshopvar fillColor = new SolidColor();var doc = app.activeDocument;function getColor(){	xml = "<object>";	xml += convertToXML(app.foregroundColor.rgb.hexValue.toLowerCase(), "foregroundColor");	xml += "</object>"	return xml;}function hasOpenedDoc(){	//alert(app.documents.length);	xml = "<object>";	xml += convertToXML((app.documents.length > 0), "exists");	xml += "</object>"	return xml;}function convertToXML(property, identifier){	var type = typeof property;	var xml = '<property id = "' + identifier + '" >';		switch(type)	{		case "number":			xml += "<number>";			xml += property.toString();			xml += "</number>";			break;		case "boolean":			xml += "<" + property.toString() + "/>";			break;		case "string":			xml += "<string>";			xml += property.toString();			xml += "</string>";			break;		case "object":			// Object case is currently not supported			alert("Object case is currently not supported");			//xml += "<object>";			//for(var i in property)			//	xml += convertToXML(property[i], 			//xml += "</object>";			break;		case "undefined":			xml += "<string>undefined</string>";			break;		default:			alert("Type " + type + " is unknown.");			return "";	}	xml += '</property>';	return xml;}function DrawBox(w, h, x, y, r, lt, rt, lb, rb, c , o){	//alert(lt);	doc = app.activeDocument;		var xPos = 0;	var yPos = 0;	var __w = parseInt(w);	var __h = parseInt(h);		var __lt = parseFloat(lt);	var __rt = parseFloat(rt);	var __lb = parseFloat(lb);	var __rb = parseFloat(rb);		var __r = parseInt(c.substring(0,2), 16);    var __g = parseInt(c.substring(2,4), 16);    var __b = parseInt(c.substring(4,6), 16);	var __o = parseInt(o);		fillColor.rgb.red  = __r;	fillColor.rgb.green = __g;	fillColor.rgb.blue = __b;		var __alayer = doc.activeLayer;		if (x == "-1" && r == "0")	{		xPos = Math.round((doc.width.as('px') - __w) / 2);	}	else	{		xPos = parseInt(x);	}	if (y == "-1" && r == "0")	{		yPos = Math.round((doc.height.as('px') - __h) / 2);	}	else	{		yPos = parseInt(y);	}		var __i = 0;		if (__lt != 0 || __rt != 0 || __lb != 0 || __rb != 0)	{		//alert("shape");		var lineArray = new Array();				lineArray[__i] = new PathPointInfo;		lineArray[__i].kind = PointKind.CORNERPOINT;		if (__lt != 0)		{			lineArray[__i].anchor = Array(xPos + __lt,yPos);			lineArray[__i].leftDirection = lineArray[__i].anchor;			lineArray[__i].rightDirection = Array(xPos + (__lt/2), yPos);		}		else		{			lineArray[__i].anchor = Array(xPos,yPos);			lineArray[__i].leftDirection = lineArray[__i].anchor;			lineArray[__i].rightDirection = lineArray[__i].anchor;		}				__i++;				lineArray[__i] = new PathPointInfo;		lineArray[__i].kind = PointKind.SMOOTHPOINT;		if (__rt != 0)		{			lineArray[__i].anchor = Array(xPos + __w - __rt, yPos);			lineArray[__i].leftDirection = Array(xPos + __w - (__rt/2), yPos);			lineArray[__i].rightDirection = lineArray[__i].anchor;						__i++;						lineArray[__i] = new PathPointInfo;			lineArray[__i].kind = PointKind.SMOOTHPOINT;			lineArray[__i].anchor = Array(xPos + __w, yPos + __rt);			lineArray[__i].leftDirection = lineArray[__i].anchor;			lineArray[__i].rightDirection = Array(xPos + __w, yPos + (__rt/2));		}		else		{			lineArray[__i].anchor = Array(xPos + __w, yPos);			lineArray[__i].leftDirection = lineArray[__i].anchor;			lineArray[__i].rightDirection = lineArray[__i].anchor;		}				__i++;						lineArray[__i] = new PathPointInfo;		lineArray[__i].kind = PointKind.SMOOTHPOINT;		if (__rb != 0)		{			lineArray[__i].anchor = Array(xPos + __w, yPos + __h - __rb);			lineArray[__i].leftDirection = Array(xPos + __w, yPos + __h - (__rb/2));			lineArray[__i].rightDirection = lineArray[__i].anchor;						__i++;						lineArray[__i] = new PathPointInfo;			lineArray[__i].kind = PointKind.SMOOTHPOINT;			lineArray[__i].anchor = Array(xPos + __w - __rb, yPos + __h);			lineArray[__i].leftDirection = lineArray[__i].anchor;			lineArray[__i].rightDirection = Array(xPos + __w - (__rb/2), yPos + __h);		}		else		{			lineArray[__i].anchor = Array(xPos + __w, yPos + __h);			lineArray[__i].leftDirection = lineArray[__i].anchor;			lineArray[__i].rightDirection = lineArray[__i].anchor;		}				__i++;						lineArray[__i] = new PathPointInfo;		lineArray[__i].kind = PointKind.SMOOTHPOINT;		if (__lb != 0)		{			lineArray[__i].anchor = Array(xPos + __lb, yPos + __h);			lineArray[__i].leftDirection = Array(xPos + (__lb/2), yPos + __h);			lineArray[__i].rightDirection = lineArray[__i].anchor;						__i++;						lineArray[__i] = new PathPointInfo;			lineArray[__i].kind = PointKind.SMOOTHPOINT;			lineArray[__i].anchor = Array(xPos, yPos + __h - __lb);			lineArray[__i].leftDirection = lineArray[__i].anchor;			lineArray[__i].rightDirection = Array(xPos, yPos + __h - (__lb/2));		}		else		{			lineArray[__i].anchor = Array(xPos, yPos + __h);			lineArray[__i].leftDirection = lineArray[__i].anchor;			lineArray[__i].rightDirection = lineArray[__i].anchor;		}				__i++;				if (__lt != 0)		{			lineArray[__i] = new PathPointInfo;			lineArray[__i].kind = PointKind.SMOOTHPOINT;			lineArray[__i].anchor = Array(xPos, yPos + __lt);			lineArray[__i].leftDirection = Array(xPos, yPos + (__lt/2));			lineArray[__i].rightDirection = lineArray[__i].anchor;		}				var lineSubPathArray = new Array();		lineSubPathArray[0] = new SubPathInfo();		lineSubPathArray[0].operation = ShapeOperation.SHAPEADD;		lineSubPathArray[0].closed = true;		lineSubPathArray[0].entireSubPath = lineArray;						var myPathItem = doc.pathItems.add("A Line", lineSubPathArray);				var layerTypeRef = new ActionReference();        layerTypeRef.putClass( stringIDToTypeID( "contentLayer" )  );        var newFillLayer = new ActionDescriptor();        newFillLayer.putReference( charIDToTypeID( "null" ) , layerTypeRef );        var colorValues = new ActionDescriptor();		colorValues.putDouble( charIDToTypeID( "Rd  " ) , __r );		colorValues.putDouble( charIDToTypeID( "Grn " ) , __g );		colorValues.putDouble( charIDToTypeID( "Bl  " ) , __b );        var rgbColor = new ActionDescriptor();        rgbColor.putObject( charIDToTypeID( "Clr " ) , charIDToTypeID( "RGBC" ) , colorValues );        var fillType = new ActionDescriptor();        fillType.putObject( charIDToTypeID( "Type" ) , stringIDToTypeID( "solidColorLayer" ) , rgbColor );        newFillLayer.putObject( charIDToTypeID( "Usng" ) , stringIDToTypeID( "contentLayer" ) , fillType );        executeAction( charIDToTypeID( "Mk  " ) , newFillLayer, DialogModes.NO );        myPathItem.remove();		        doc.activeLayer.opacity = __o;				if (r == "1")		{			var box = doc.activeLayer;			//alert(box.name);			positionLayer(box, __alayer, xPos, yPos, 0);		}					}	else	{		//alert("pixel");		var box = doc.artLayers.add("Box");		doc.selection.select([[xPos,yPos],[xPos + __w,yPos],[xPos + __w,yPos + __h],[xPos,yPos + __h]]);		doc.selection.fill(fillColor, ColorBlendMode.NORMAL, 100, false);		doc.selection.deselect();		if (r == "1")		{			positionLayer(box, __alayer, xPos, yPos, 1);		}		box.move(__alayer, ElementPlacement.PLACEBEFORE);		doc.selection.deselect();		box.transparentPixelsLocked = true;		doc.activeLayer = box;		doc.activeLayer.opacity = __o;	}	}function positionLayer(lyr, adoc, xPos, yPos, c){// layerObject, Number, Number   // if can not move layer return   if (adoc.isBackgroundLayer || lyr.isBackgroundLayer || lyr.positionLocked) return;      // get the layer bounds   var __lyrB = lyr.bounds;   // get top left position   var __lyrBX = __lyrB[0].value;   var __lyrBY = __lyrB[1].value;      // get the layer bounds   var __adocB = adoc.bounds;   // get top left position   var __adocBX = __adocB[0].value;   var __adocBY = __adocB[1].value;      // the difference between where layer needs to be and is now   var deltaX = __adocBX - __lyrBX + c;   var deltaY =  __adocBY - __lyrBY + c;   // move the layer into position   //alert(deltaX + ":" + xPos);   lyr.translate (deltaX + xPos, deltaY + yPos);}