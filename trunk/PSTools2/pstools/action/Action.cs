using System;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.IO;

namespace PSTools
{
	class Action
	{
		public enum Actions
		{
			SAVE = 1,
			CLEAN = 2,
			IMAGE_RIGHTS = 3,
			EXPORT_SO = 4,
			SAVE_SELECTION = 5
		}

		public enum Colors
		{
			NONE,
			RED,
			ORANGE,
			YELLOW,
			GREEN,
			BLUE,
			VIOLET,
			GRAY
		}

		public enum Format
		{
			JPG,
			PNG,
			GIF
		}

		private Photoshop.Application __appRef = new Photoshop.Application();
		private bool __openDoc = true;
		private bool __stayOpen = false;
		private string[] __cmdargs;
		private Form __form;

		public Action(Form _form)
		{
			__form = _form;
		}

		public void execute(Actions __action, string[] __args)
		{
			execute(__action, "", __args);
		}

		public void execute(Actions __action, string __subcommand, string[] __args)
		{
			Photoshop.Document __docRef;

			__cmdargs = __args;
			//MessageBox.Show(__args[1].ToString());
			__docRef = this.openDocument();

			switch (__action)
			{
				case Actions.SAVE:
					saveToFile(__docRef, __args);
					break;
				case Actions.CLEAN:
					cleanLayersName(__docRef.Layers, 1);
					break;
				case Actions.IMAGE_RIGHTS:
					exportImagesRights(__docRef, __docRef.Layers);
					break;
				case Actions.EXPORT_SO:
					exportSmartObjects(__docRef, __docRef.Layers);
					break;
				case Actions.SAVE_SELECTION:
					saveToFile(__docRef, __args, true);
					break;
			}

			this.closeDocument(__docRef);

			// End program
			Application.Exit();
		}

		private Photoshop.Document openDocument()
		{
			Photoshop.Document __docRef = null;

			try
			{
				if (__appRef.Documents.Count > 0)
				{
					for (int i = 1; i <= __appRef.Documents.Count; i++)
					{
						if (__cmdargs[1 + __form.idx] == __appRef.Documents[i].FullName)
						{
							__appRef.ActiveDocument = __appRef.Documents[i];
							__docRef = __appRef.ActiveDocument;
							__openDoc = false;
							__stayOpen = true;
						}
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

			if (__openDoc)
			{
				try
				{
					__docRef = __appRef.Open(__cmdargs[1 + __form.idx], null, null);
				}
				catch (Exception ex)
				{
					MessageBox.Show(ex.Message);
				}
			}
			return __docRef;
		}

		private void closeDocument(Photoshop.Document __docRef)
		{
			if (!__stayOpen)
			{
				__docRef.Close(2);
			}
		}

		private bool exportSmartObjects(Photoshop.Document __docRef, object _layers)
		{
			Photoshop.Layers __layers;
			Photoshop.ArtLayer __alayer = null;
			Photoshop.LayerSet __slayer;
			bool __isArtLayer = false;
			object __layer;
			int __j;
			string __soType;

			__layers = (Photoshop.Layers)_layers;

			for (__j = 1; __j <= __layers.Count; __j++)
			{
				__layer = __layers[__j];

				try
				{
					__alayer = (Photoshop.ArtLayer)__layer;
					__isArtLayer = true;
				}
				catch
				{
					__isArtLayer = false;
				}

				if (__isArtLayer) // Everything as Layer goes here
				{
					__appRef.ActiveDocument.ActiveLayer = __layer;

					if (__alayer.Kind == Photoshop.PsLayerKind.psSmartObjectLayer)
					{

						int __idplacedLayerExportContents;
						__idplacedLayerExportContents = __appRef.StringIDToTypeID("placedLayerExportContents");


						Photoshop.ActionDescriptor __desc4;
						__desc4 = new Photoshop.ActionDescriptor();

						int __idnull;
						__idnull = __appRef.CharIDToTypeID("null");

						if (!Directory.Exists(__docRef.Path + "+ Assets\\"))
						{
							Directory.CreateDirectory(__docRef.Path + "+ Assets\\");
						}

						__soType = getSmartObjectType(__appRef);
						if (__soType != "")
						{
							__desc4.PutPath(__idnull, __docRef.Path + "+ Assets\\" + wipeName(__alayer.Name) + __soType);
							__appRef.ExecuteAction(__idplacedLayerExportContents, __desc4, Photoshop.PsDialogModes.psDisplayNoDialogs);
						}
					}
				}
				else // Everything as LayerSet goes here
				{
					__slayer = (Photoshop.LayerSet)__layer;
					__appRef.ActiveDocument.ActiveLayer = __layer;

					if (__slayer.LayerType == Photoshop.PsLayerType.psLayerSet)
					{
						bool __test = exportSmartObjects(__docRef, __slayer.Layers);
					}
				}
			}
			return true;
		}

		private string getSmartObjectType(Photoshop.Application __appRef)
		{
			string __value;
			Photoshop.ActionReference __aRef = new Photoshop.ActionReference();
			__aRef.PutEnumerated(__appRef.CharIDToTypeID("Lyr "), __appRef.CharIDToTypeID("Ordn"), __appRef.CharIDToTypeID("Trgt"));

			Photoshop.ActionDescriptor __desc;
			__desc = __appRef.ExecuteActionGet(__aRef);

			if (__desc.HasKey(__appRef.StringIDToTypeID("smartObject")))
			{
				int __desc2 = __appRef.ExecuteActionGet(__aRef).GetObjectValue(__appRef.StringIDToTypeID("smartObject")).GetEnumerationValue(__appRef.StringIDToTypeID("placed"));

				if ((string)(__appRef.TypeIDToStringID(__desc2)) == "vectorData")
				{
					__value = ".ai";
				}
				else if ((string)(__appRef.TypeIDToStringID(__desc2)) == "rasterizeContent")
				{
					__value = ".psd";
				}
				else
				{
					__value = "";
				}
			}
			else
			{
				__value = "";
			}
			return __value;
		}

		private string wipeName(string __name)
		{
			string __value;
			Regex __re = new Regex("(\\+)+\\s", RegexOptions.IgnoreCase | RegexOptions.Multiline);
			__value = __re.Replace(__name, "");

			__re = new Regex("\\s(copy)\\s(\\d)*", RegexOptions.IgnoreCase | RegexOptions.Multiline);
			__value = __re.Replace(__value, "");

			__re = new Regex("(vector)\\s*|^v:", RegexOptions.IgnoreCase | RegexOptions.Multiline);
			__value = __re.Replace(__value, "");

			__re = new Regex(":", RegexOptions.IgnoreCase | RegexOptions.Multiline);
			__value = __re.Replace(__value, "");

			__re = new Regex("\\s", RegexOptions.IgnoreCase | RegexOptions.Multiline);
			__value = __re.Replace(__value, "_");
			return __value;
		}


		private bool exportImagesRights(Photoshop.Document __docRef, object _layers)
		{
			Photoshop.Layers __layers;
			Photoshop.ArtLayer __alayer = null;
			Photoshop.LayerSet __slayer;
			bool __isArtLayer = false;
			object __layer;
			int __j;
			string __soType;
			ImageRight __ir;

			__ir = new ImageRight();

			__layers = (Photoshop.Layers)_layers;

			for (__j = 1; __j <= __layers.Count; __j++)
			{
				__layer = __layers[__j];

				try
				{
					__alayer = (Photoshop.ArtLayer)__layer;
					__isArtLayer = true;
				}
				catch
				{
					__isArtLayer = false;
				}

				if (__isArtLayer) // Everything as Layer goes here
				{
					__appRef.ActiveDocument.ActiveLayer = __layer;

					if (__alayer.Kind == Photoshop.PsLayerKind.psNormalLayer)
					{
						__ir.parse(__alayer.Name);
						if (__ir.isValidURL)
						{
							__ir.createLink(__docRef.Path);
						}
					}
					else if (__alayer.Kind == Photoshop.PsLayerKind.psSmartObjectLayer)
					{
						__ir.parse(__alayer.Name);
						if (__ir.Code != null)
						{
							if (__ir.isValidURL)
							{
								__ir.createLink(__docRef.Path);
							}
						}
						else
						{
							__soType = getSmartObjectType(__appRef);
							if (__soType == ".psd")
							{

								int __opn;
								__opn = __appRef.StringIDToTypeID("placedLayerEditContents");

								Photoshop.ActionDescriptor __desc4;
								__desc4 = new Photoshop.ActionDescriptor();

								try
								{
									__appRef.ExecuteAction(__opn, __desc4, Photoshop.PsDialogModes.psDisplayNoDialogs);
								}
								catch (InvalidOperationException ex)
								{
									MessageBox.Show(ex.Message);
								}
								bool __test = exportImagesRights(__docRef, __appRef.ActiveDocument.Layers);
								__appRef.ActiveDocument.Close(2);
							}
						}
					}

					
				}
				else // Everything as LayerSet goes here
				{
					__slayer = (Photoshop.LayerSet)__layer;
					__appRef.ActiveDocument.ActiveLayer = __layer;

					if (__slayer.LayerType == Photoshop.PsLayerType.psLayerSet)
					{
						bool __test = exportImagesRights(__docRef, __slayer.Layers);
					}
				}
			}

			/*Photoshop.ArtLayers __layers;
			Photoshop.ArtLayer __layer;
			bool __isVisible;
			int __j;
			ImageRight __ir;
			string __soType;

			__ir = new ImageRight();
			__layers = __docRef.ArtLayers;

			for (__j = 1; __j <= __layers.Count; __j++)
			{
				__layer = __layers[__j];
				__isVisible = __layer.Visible;
				__docRef.ActiveLayer = __layer;

				//MessageBox.Show(__Layer.Name)
				if (__layer.Kind == Photoshop.PsLayerKind.psNormalLayer)
					{
						__ir.parse(__layer.Name);
						if (__ir.isValidURL)
						{
							__ir.createLink(__docRef.Path);
						}
					}
					else if (__layer.Kind == Photoshop.PsLayerKind.psSmartObjectLayer)
					{
						__ir.parse(__layer.Name);
						if (__ir.Code != null)
						{
							if (__ir.isValidURL)
							{
								__ir.createLink(__docRef.Path);
							}
						}
						else
						{
							__soType = getSmartObjectType(__appRef);
							if (__soType == ".psd")
							{

								int __opn;
								__opn = __appRef.StringIDToTypeID("placedLayerEditContents");

								Photoshop.ActionDescriptor __desc4;
								__desc4 = new Photoshop.ActionDescriptor();

								try
								{
									__appRef.ExecuteAction(__opn, __desc4, Photoshop.PsDialogModes.psDisplayNoDialogs);
								}
								catch (InvalidOperationException ex)
								{
									MessageBox.Show(ex.Message);
								}
								exportImagesRights(__appRef.ActiveDocument);
								__appRef.ActiveDocument.Close(2);
							}
						}
				}
				__layer.Visible = __isVisible;
				__appRef.ActiveDocument.ActiveLayer = __layer;
			}*/
			return true;
		}

		private string cleanLayersName(object _layers, int __idx)
		{
			string returnValue;
			Photoshop.Layers __layers = null;
			object __layer;
			Photoshop.LayerSet __slayer;
			Photoshop.ArtLayer __alayer = null;
			bool __isVisible;
			int __j;
			ImageRight __ir;
			bool __isArtLayer = false;

			string __xmlDoc;
			string __FistLayerText = "";

			Regex __reg = new Regex("Group\\s*\\d*", RegexOptions.IgnoreCase);

			__ir = new ImageRight();

			//MessageBox.Show(_layers.GetType().ToString());
			__layers = (Photoshop.Layers)_layers;
			for (__j = 1; __j <= __layers.Count; __j++)
			{
				__layer = __layers[__j];

				try
				{
					__alayer = (Photoshop.ArtLayer)__layer;
					__isArtLayer = true;
				}
				catch
				{
					__isArtLayer = false;
				}

				if (__isArtLayer) // Everything as Layer goes here
				{
					__isVisible = __alayer.Visible;
					__appRef.ActiveDocument.ActiveLayer = __layer;

					__ir.parse(__alayer.Name);
					if (__ir.isValidCode)
					{
						__alayer.Name = "#" + Regex.Replace(__alayer.Name, "#", "");
					}
					if (__alayer.Kind == Photoshop.PsLayerKind.psSmartObjectLayer) //SMARTOBJECT
					{
						try
						{
							__xmlDoc = __appRef.ActiveDocument.XMPMetadata.RawData;
							if (__xmlDoc != "")
							{
								__alayer.Name = new string(Convert.ToChar("+"), __idx) + " " + Regex.Replace(__alayer.Name, "(\\+)+\\s*", "");
								ChangeLayerColour(Colors.VIOLET);
							}
						}
						catch
						{
						}
					}
					else if (__alayer.Kind == Photoshop.PsLayerKind.psTextLayer)
					{
						if (__FistLayerText == "")
						{
							__FistLayerText = __alayer.Name;
						}
					}

					__alayer.Visible = __isVisible;
				}
				else // Everything as LayerSet goes here
				{
					__slayer = (Photoshop.LayerSet)__layer;
					__isVisible = __slayer.Visible;
					__appRef.ActiveDocument.ActiveLayer = __layer;

					if (__slayer.LayerType == Photoshop.PsLayerType.psLayerSet)
					{
						__slayer.Name = new string(Convert.ToChar("+"), __idx) + " " + Regex.Replace(__slayer.Name, "(\\+)+\\s*", "");
						//MessageBox.Show(__slayer.Name);
						string __NewLayerName = cleanLayersName(__slayer.Layers, __idx + 1);
						if (__NewLayerName != "" && __reg.IsMatch(__slayer.Name))
						{
							__slayer.Name = new string(Convert.ToChar("+"), __idx) + " " + __NewLayerName;
						}
					}

					__slayer.Visible = __isVisible;
				}
			}
			returnValue = __FistLayerText;
			return returnValue;
		}

		private void ChangeLayerColour(Colors __col)
		{
			string __colour;
			Photoshop.ActionDescriptor __desc;
			Photoshop.ActionReference __ref;
			Photoshop.ActionDescriptor __desc2;

			switch (__col)
			{
				case Colors.RED:
					__colour = "Rd  ";
					break;
				case Colors.ORANGE:
					__colour = "Orng";
					break;
				case Colors.YELLOW:
					__colour = "Ylw ";
					break;
				case Colors.GREEN:
					__colour = "Grn ";
					break;
				case Colors.BLUE:
					__colour = "Bl  ";
					break;
				case Colors.VIOLET:
					__colour = "Vlt ";
					break;
				case Colors.GRAY:
					__colour = "Gry ";
					break;
				case Colors.NONE:
					__colour = "None";
					break;
				default:
					__colour = "None";
					break;
			}

			__desc = new Photoshop.ActionDescriptor();
			__ref = new Photoshop.ActionReference();
			__ref.PutEnumerated(__appRef.CharIDToTypeID("Lyr "), __appRef.CharIDToTypeID("Ordn"), __appRef.CharIDToTypeID("Trgt"));
			__desc.PutReference(__appRef.CharIDToTypeID("null"), __ref);

			__desc2 = new Photoshop.ActionDescriptor();
			__desc2.PutEnumerated(__appRef.CharIDToTypeID("Clr "), __appRef.CharIDToTypeID("Clr "), __appRef.CharIDToTypeID(__colour));
			__desc.PutObject(__appRef.CharIDToTypeID("T   "), __appRef.CharIDToTypeID("Lyr "), __desc2);
			__appRef.ExecuteAction(__appRef.CharIDToTypeID("setd"), __desc, Photoshop.PsDialogModes.psDisplayNoDialogs);
		}


		private void saveToFile(Photoshop.Document __docRef, string[] __args)
		{
			saveToFile(__docRef, __args, false);
		}

		private void saveToFile(Photoshop.Document __docRef, string[] __args, bool __selectionOnly)
		{
			bool __exportLayerComps = false;
			int __compsCount;
			int __compsIndex;
			Photoshop.LayerComp __compRef;
			Photoshop.Document __duppedDocument;
			string __fileNameBody = null;
			bool __hasSelection;
			Format __imageFormat;

			__exportLayerComps = __form.ExportLayerComps.Checked;
			//MessageBox.Show(__args.ToString());

			bool __isNamedLayerComp = false;

			Photoshop.JPEGSaveOptions __jpgSaveOptions = null;

			try
			{
				__jpgSaveOptions = new Photoshop.JPEGSaveOptions();
				__jpgSaveOptions.EmbedColorProfile = false;
				__jpgSaveOptions.FormatOptions = Photoshop.PsFormatOptionsType.psStandardBaseline; // 1 psStandardBaseline
				__jpgSaveOptions.Matte = Photoshop.PsMatteType.psNoMatte; // 1 psNoMatte
			}
			catch (Exception __e)
			{
				DialogResult __dr = MessageBox.Show("Photoshop is busy with open dialog or something." + "\r\n" + "\r\n" + "Please switch to Photoshop then close open dialogs or leave editing state", "Photoshop not ready", MessageBoxButtons.OK);
				if (__dr == DialogResult.OK)
				{
					Application.Exit();
				}
			}

			switch (__args[2 + __form.idx])
			{
				case "jpg":
					__jpgSaveOptions.Quality = int.Parse(__args[4 + __form.idx]);
					__imageFormat = Format.JPG;
					__exportLayerComps = true; //force using this mode
					break;
				case "png":
					__imageFormat = Format.PNG;
					__exportLayerComps = true;
					break;
				case "gif":
					__imageFormat = Format.GIF;
					break;
				case "sc":
					__jpgSaveOptions.Quality = 12;
					__imageFormat = Format.JPG;
					__exportLayerComps = true; //force using this mode
					break;
				default: // Export each layer by its name
					__jpgSaveOptions.Quality = int.Parse(__args[4 + __form.idx]);
					__imageFormat = Format.JPG;
					break;
			}

			switch (__args[3 + __form.idx])
			{
				case "name":
					__isNamedLayerComp = true;
					break;
				case "index":
					__isNamedLayerComp = false;
					break;
				default:
					__isNamedLayerComp = false;
					break;
			}

			

			Photoshop.ExportOptionsSaveForWeb __gifExportOptionsSaveForWeb = new Photoshop.ExportOptionsSaveForWeb();
			//gifExportOptionsSaveForWeb.MatteColor = 255
			__gifExportOptionsSaveForWeb.Format = Photoshop.PsSaveDocumentType.psCompuServeGIFSave; // 3;
			__gifExportOptionsSaveForWeb.ColorReduction = Photoshop.PsColorReductionType.psAdaptive; //1;
			__gifExportOptionsSaveForWeb.Colors = 256;
			__gifExportOptionsSaveForWeb.Dither = Photoshop.PsDitherType.psNoise; //3;
			__gifExportOptionsSaveForWeb.DitherAmount = 100;
			__gifExportOptionsSaveForWeb.Quality = 100;
			__gifExportOptionsSaveForWeb.Transparency = true;
			__gifExportOptionsSaveForWeb.TransparencyAmount = 100;
			__gifExportOptionsSaveForWeb.TransparencyDither = Photoshop.PsDitherType.psNoDither; //2;
			__gifExportOptionsSaveForWeb.IncludeProfile = false;
			__gifExportOptionsSaveForWeb.Lossy = 0;
			__gifExportOptionsSaveForWeb.WebSnap = 0;

			Photoshop.ExportOptionsSaveForWeb __pngExportOptionsSaveForWeb = new Photoshop.ExportOptionsSaveForWeb();
			__pngExportOptionsSaveForWeb.Format = Photoshop.PsSaveDocumentType.psPNGSave; // 13;
			__pngExportOptionsSaveForWeb.PNG8 = false;
			__pngExportOptionsSaveForWeb.Transparency = true;

			//MessageBox.Show(__args(3))
			
			//MessageBox.Show(__imageType)


			__compsCount = __docRef.LayerComps.Count;

			__hasSelection = false;
			__hasSelection = hasScreenSelection(__docRef);

			// Exporting layercomps by index or name
			if (__exportLayerComps)
			{
				if (__compsCount <= 1)
				{
					//Set textItemRef = appRef.ActiveDocument.Layers(1)

					//textItemRef.TextItem.Contents = Args.Item(1)

					//outFileName = Args.Item(1)
					if (__hasSelection)
						saveScreenSelection(__docRef, __docRef, __jpgSaveOptions); 
					
					if (!__selectionOnly) // IF screen selection then save crop
					{
						//MessageBox.Show(__args[1]);
						if (__imageFormat == Format.JPG)
						{
							__docRef.SaveAs(__args[1 + __form.idx], __jpgSaveOptions, true, null);
						}
						else if (__imageFormat == Format.PNG)
						{
							__fileNameBody = __docRef.Name.Substring(0, __docRef.Name.LastIndexOf(".")) + ".png";
							__docRef.Export(__docRef.Path + __fileNameBody, 2, __pngExportOptionsSaveForWeb);
						}
						else
						{
							__fileNameBody = __docRef.Name.Substring(0, __docRef.Name.LastIndexOf(".")) + ".gif";
							__docRef.Export(__docRef.Path + __fileNameBody, 2, __gifExportOptionsSaveForWeb);
						}
					}

				}
				else
				{
					//msgbox("comps!")
					for (__compsIndex = 1; __compsIndex <= __compsCount; __compsIndex++)
					{
						//MsgBox(docRef.LayerComps.Count)
						//End
						__compRef = __docRef.LayerComps[__compsIndex];
						//if (exportInfo.selectionOnly && !compRef.selected) continue; // selected only
						__compRef.Apply();

						//msgbox(compRef.Name)
						if (__hasSelection)
						{
							__duppedDocument = __docRef.Duplicate(null, null);
							saveScreenSelection(__docRef, __duppedDocument, __compsIndex, __jpgSaveOptions);
						}

						if (!__selectionOnly)
						{

							__duppedDocument = __docRef.Duplicate(null, null);

							if (!__isNamedLayerComp)
							{
								__fileNameBody = __docRef.Name.Substring(0, __docRef.Name.LastIndexOf(".")) + "." + __compsIndex;
							}
							else
							{
								__fileNameBody = __compRef.Name;
							}


							if (__imageFormat == Format.JPG)
							{
								__fileNameBody += ".jpg";
								__duppedDocument.SaveAs(__docRef.Path + __fileNameBody, __jpgSaveOptions, true, null);
							}
							else if (__imageFormat == Format.PNG)
							{
								__fileNameBody += ".png";
								__duppedDocument.Export(__docRef.Path + __fileNameBody, 2, __pngExportOptionsSaveForWeb);
							}
							else
							{
								__fileNameBody += ".gif";
								__duppedDocument.Export(__docRef.Path + __fileNameBody, 2, __gifExportOptionsSaveForWeb);
							}
							__duppedDocument.Close(2);

						}
					}
					__compRef = __docRef.LayerComps[1];
					__compRef.Apply();
				}

				//MsgBox(Me.AutoArchive.Checked)
				archiveFiles(__docRef);
			}
			else //Exporting each layers by name
			{
				Photoshop.ArtLayer __layer;
				for (__compsIndex = 1; __compsIndex <= __docRef.ArtLayers.Count; __compsIndex++)
				{
					__layer = __docRef.ArtLayers[__compsIndex];
					//isVisible = oLayer.visible
					__appRef.ActiveDocument.ActiveLayer = __layer;
					//oLayer.Apply()
					//duppedDocument = docRef.Duplicate()
					//msgbox(compRef.Name)
					__fileNameBody = (string)(__layer.Name + ".jpg");
					//msgbox(fileNameBody)
					__docRef.SaveAs(__docRef.Path + __fileNameBody, __jpgSaveOptions, true, null);
					
					__layer.Visible = false;
					__appRef.ActiveDocument.ActiveLayer = __layer;
					//duppedDocument.Close(2)
				}
			}

			//ExportImagesRights()
			if (__compsCount > 0)
			{
				__compRef = __docRef.LayerComps[1];
				__compRef.Apply();
			}
		}

		private void saveScreenSelection(Photoshop.Document __docRef, Photoshop.Document __doc, Photoshop.JPEGSaveOptions __options)
		{
			saveScreenSelection(__docRef, __doc, -1, __options);
		}

		private void saveScreenSelection(Photoshop.Document __docRef, Photoshop.Document __doc, int __idx, Photoshop.JPEGSaveOptions __options)
		{
			Photoshop.Channel __selChannel;
			object __selBounds;
			string __fileNameBody;

			try
			{
				__selChannel = __doc.Channels["screen"];
				//MessageBox.Show(__selChannel.Name)
				__doc.Selection.Load(__selChannel, null, null);
				__selBounds = __doc.Selection.Bounds;
				__doc.Crop(__selBounds, null, null, null);
			}
			catch (Exception)
			{
				MessageBox.Show("You have to create a selection named \"screen\"", "No selection found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				Application.Exit();
			}

			__fileNameBody = (__docRef.Name.LastIndexOf(".") > -1) ? __docRef.Name.Substring(0, __docRef.Name.LastIndexOf(".")) : __docRef.Name;
			__fileNameBody += (__idx <= -1) ? "_screen" : "." + __idx + "_screen";
			__fileNameBody += ".jpg";
			__doc.SaveAs(__docRef.Path + __fileNameBody, __options, true, null);

			if (__idx > 0)
				__doc.Close(2);
		}

		private void archiveFiles(Photoshop.Document __docRef)
		{
			DirectoryInfo __di;
			try
			{
				__di = new DirectoryInfo(__docRef.Path);
			}
			catch (Exception)
			{
				goto finish;
			}


			//MsgBox(di.Name)
			if (__form.AutoArchive.Checked && !isExcludeDirectory(__di.Name))
			{

				//Dim di As DirectoryInfo
				FileInfo[] __afi;


				//MsgBox(Directory.Exists(docRef.Path & "\" & Me.ArchiveDirectory.Text & "\"))
				//If Not Directory.Exists(__docRef.Path & "\" & Me.ArchiveDirectory.Text & "\") Then
				//Directory.CreateDirectory(__docRef.Path & "\" & Me.ArchiveDirectory.Text & "\")
				//End If

				//di = New DirectoryInfo(docRef.Path)
				string __currentFileName = __docRef.Name.Substring(0, __docRef.Name.LastIndexOf("."));

				//Dim __RegexObj As Regex = New Regex("\d*$")
				Regex __RegexObj = new Regex("(\\d*)\\.*\\d*$");
				Match __myMatches;
				string __currentVersion;
				string __cleanFileName;

				if (__RegexObj.IsMatch(__currentFileName))
				{
					__myMatches = __RegexObj.Match(__currentFileName);
					__currentVersion = __myMatches.Value;

					if (__currentVersion.Length < 1)
					{
						goto finish;
					}

					__cleanFileName = __RegexObj.Replace(__currentFileName, "");
					__cleanFileName = __cleanFileName.Replace("+", "\\+");
					__cleanFileName = __cleanFileName.Replace(" ", "\\s");
					//MsgBox(__cleanFileName)
					Regex __RegexObj2 = new Regex("^" + __cleanFileName + "(\\d+|\\.)");

					__afi = __di.GetFiles("*.*");
					//MsgBox(__currentVersion)
					foreach (FileInfo __fi in __afi)
					{
						if (__RegexObj2.IsMatch(__fi.Name))
						{
							//MsgBox(__fi.Name)
							if (isOldFileVersion(__fi.Name, __currentVersion.ToString())) //And Directory.Exists(__docRef.Path & "\" & Me.ArchiveDirectory.Text & "\") Then
							{
								//MsgBox(__fi.Name)
								if (!Directory.Exists(__docRef.Path + "\\" + __form.ArchiveDirectory.Text + "\\"))
								{
									Directory.CreateDirectory(__docRef.Path + "\\" + __form.ArchiveDirectory.Text + "\\");
								}
								try
								{
									File.Copy(__docRef.Path + __fi.Name, __docRef.Path + "\\" + __form.ArchiveDirectory.Text + "\\" + __fi.Name, true);
								}
								catch (Exception ex)
								{
									MessageBox.Show(ex.Message);
								}

								try
								{
									File.Delete(__docRef.Path + __fi.Name);
								}
								catch { }

							}
						}
					}
				}
			}

		finish:{ }
		}

		private bool hasScreenSelection(Photoshop.Document __doc)
		{
			bool __value;
			Photoshop.Channel __selChannel;

			try
			{
				__selChannel = __doc.Channels["screen"];
				__value = true;
			}
			catch (Exception)
			{
				__value = false;
			}

			return __value;
		}

		private bool isExcludeDirectory(string __dirName)
		{
			Array __ExcludeDirectories;

			
			if (__form.ExcludeDirectories.Text != "")
			{
				__ExcludeDirectories = (__form.ExcludeDirectories.Text + ";" + __form.ArchiveDirectory.Text).Split(';');
				foreach (string __ExcludeDirectory in __ExcludeDirectories)
				{
					if (__ExcludeDirectory == __dirName)
					{
						return true;
					}
				}
				return false;
			}
			else
			{
				return false;
			}
		}

		private bool isOldFileVersion(string __filename, string _version)
		{
			string __newFileName = __filename.Substring(0, __filename.LastIndexOf("."));
			string[] __version = _version.Split(Convert.ToChar("."));
			//MsgBox(__version(0))
			Regex __RegexObj = new Regex("(\\d*)(\\.*\\d*)*(_screen)*$");
			try
			{
				if (__RegexObj.IsMatch(__newFileName))
				{
					//MsgBox("> " & __RegexObj.Match(__newFileName).Groups(1).Value & ":" & version)
					if (int.Parse(__RegexObj.Match(__newFileName).Groups[1].Value) < int.Parse(__version[0]))
					{
						return true;
					}
					else
					{
						return false;
					}
				}
				else
				{
					return false;
				}
			}
			catch (Exception)
			{
				//MsgBox(ex.Message)
				return false;
			}
		}

	}	
}
