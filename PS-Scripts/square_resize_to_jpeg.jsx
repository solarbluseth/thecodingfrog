// 
//  resize canvas  & export JPG 
// 

#target photoshop 
app.bringToFront(); 

// Save current dialog preferences 
var startDisplayDialogs = app.displayDialogs;      

// Save current unit preferences 
app.displayDialogs = DialogModes.NO;
var couleur = new SolidColor();
couleur.rgb.red = 255;
couleur.rgb.green = 255;
couleur.rgb.blue = 255;
app.backgroundColor = couleur;

// Source and Destination Folders 
var size = 400;
var inputFolder = Folder.selectDialog("Selectionnez le répertoire source"); 
var outputFolder = Folder.selectDialog("Selectionnez le répertoire de destination"); 

ProcessImages(); 

function ProcessImages() 
{ 
	var filesOpened = 0; 

	if ( inputFolder == null || outputFolder == null)
	{ 
		if ( inputFolder == null)
		{ 
			alert("Pas de répertoire source sélectionné"); 
		} 
		if ( outputFolder == null)
		{ 
			alert("Pas de répertoire de destination sélectionné"); 
		} 
	}
	else
	{ 
		var fileList = inputFolder.getFiles(); 
		for ( var i = 0; i < fileList.length; i++ ) 
		{ 
			if ( fileList[i] instanceof File && ! fileList[i].hidden) 
			{ 
				open( fileList[i] ); 
				ResizeImage();
				filesOpened++;  
			} 
		} 
	} 
	return filesOpened; 
} 

function ExportJpeg(filePrefix, fileSuffix)
{ 
	try 
	{ 
		var docRef = app.activeDocument; 
		var docName = app.activeDocument.name.slice(0,-4); 
		var jpegOptions = new JPEGSaveOptions(); 
		jpegOptions.quality = 8 
		docRef.flatten() 
		docRef.bitsPerChannel = BitsPerChannelType.EIGHT 
		jpegFile = new File( outputFolder + "//"  + filePrefix + docName + fileSuffix ); 
		//Save Document As 
		docRef.saveAs(jpegFile, JPEGSaveOptions, true, Extension.LOWERCASE); 
	} 
	catch (e) 
	{ 
		alert("Erreur lors de la sauvegarde de l'image. \r\r" + e); 
		return; 
	} 
}; 

function ResizeImage() 
{ 
      var docRef = app.activeDocument; 
      var docWidth = docRef.width.as("px"); 
      var docHeight = docRef.height.as("px");
	  var coeff = docWidth / docHeight;
	  
      if (docWidth < docHeight) 
      { 
         docRef.resizeImage(size * coeff, size, 72, ResampleMethod.BICUBIC ); 
      }       
      else if (docWidth > docHeight) 
      { 
         docRef.resizeImage(size, size / coeff, 72, ResampleMethod.BICUBIC ); 
      } 
      else if (docWidth == docHeight) 
      { 
		docRef.resizeImage(size, size, 72, ResampleMethod.BICUBIC ); 
      }
  
	docRef.resizeCanvas(size, size, AnchorPosition.MIDDLECENTER);
	
	app.displayDialogs = DialogModes.NO; 
	ExportJpeg("", ".jpg"); 

	var savedState = docRef.activeHistoryState; 
	docRef.activeHistoryState = savedState; 
	docRef.close(SaveOptions.DONOTSAVECHANGES); 
	docRef = null; 
}; 

// Reset preferences 
app.displayDialogs = startDisplayDialogs; 

//alert("Operation Complete!" + "\n" + "Images were successfully exported to:" + "\n" + "\n" + outputFolder.toString().match(/([^\.]+)/)[1] + "/");