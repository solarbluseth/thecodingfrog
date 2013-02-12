using System;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.IO;
using System.Reflection;

namespace PSTools
{
	class Action
	{
		public enum Actions
		{
			SAVE = 1,
			CLEAN = 2,
			IMAGE_RIGHTS = 3,
			EXPORT_SMARTOBJECTS = 4,
			SAVE_SELECTION = 5,
			EXPORT_BASE64 = 6,
			EXPORT_ASSETS = 7
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
		private bool __keepOpen = false;
		private string[] __cmdargs;
		private Form __form;

		/// <summary>
		/// Initializes a new instance of the <see cref="Action"/> class.
		/// </summary>
		/// <param name="_form">Form</param>
		public Action(Form _form)
		{
			__form = _form;
		}

		/// <summary>
		/// Executes the specified __action.
		/// </summary>
		/// <param name="__action">Action</param>
		/// <param name="__args">Arguments</param>
		public void execute(Actions __action, string[] __args)
		{
			execute(__action, "", __args);
		}

		/// <summary>
		/// Executes the specified __action.
		/// </summary>
		/// <param name="__action">Action</param>
		/// <param name="__subcommand">Subcommand</param>
		/// <param name="__args">Arguments</param>
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
				case Actions.EXPORT_SMARTOBJECTS:
					exportSmartObjects(__docRef, __docRef.Layers);
					break;
				case Actions.SAVE_SELECTION:
					saveToFile(__docRef, __args, true);
					break;
				case Actions.EXPORT_BASE64:
					exportBase64();
					break;
				case Actions.EXPORT_ASSETS:
					exportAssets(__docRef, __docRef.Layers);
					break;
			}

			this.closeDocument(__docRef);

			// End program
			Application.Exit();
		}

		/// <summary>
		/// Opens the document.
		/// </summary>
		/// <returns></returns>
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
							__keepOpen = true;
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

		/// <summary>
		/// Closes the document.
		/// </summary>
		/// <param name="__docRef">Document refrence</param>
		private void closeDocument(Photoshop.Document __docRef)
		{
			if (!__keepOpen)
			{
				__docRef.Close(2);
			}
		}

		/// <summary>
		/// Exports the smart objects.
		/// </summary>
		/// <param name="__docRef">Document refrence</param>
		/// <param name="_layers">Layers</param>
		/// <returns></returns>
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

						if (!Directory.Exists(__docRef.Path + "+ SmartObjects\\"))
						{
							Directory.CreateDirectory(__docRef.Path + "+ SmartObjects\\");
						}

						__soType = getSmartObjectType(__appRef);
						if (__soType != "")
						{
							__desc4.PutPath(__idnull, __docRef.Path + "+ SmartObjects\\" + wipeName(__alayer.Name) + __soType);
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

		/// <summary>
		/// Gets the type of the smart object.
		/// </summary>
		/// <param name="__appRef">Application Reference</param>
		/// <returns></returns>
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

		/// <summary>
		/// Wipes the name.
		/// </summary>
		/// <param name="__name">Name</param>
		/// <returns></returns>
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


		/// <summary>
		/// Exports the images rights.
		/// </summary>
		/// <param name="__docRef">Document</param>
		/// <param name="_layers">Layers</param>
		/// <returns></returns>
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
			return true;
		}

		/// <summary>
		/// Cleans the name of the layers.
		/// </summary>
		/// <param name="_layers">layers</param>
		/// <param name="__idx">Depth Index</param>
		/// <returns></returns>
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

		/// <summary>
		/// Changes the layer colour.
		/// </summary>
		/// <param name="__col">Color</param>
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


		/// <summary>
		/// Saves to file.
		/// </summary>
		/// <param name="__docRef">Document</param>
		/// <param name="__args">Arguments</param>
		private void saveToFile(Photoshop.Document __docRef, string[] __args)
		{
			saveToFile(__docRef, __args, false);
		}

		/// <summary>
		/// Saves to file.
		/// </summary>
		/// <param name="__docRef">Document</param>
		/// <param name="__args">Arguments</param>
		/// <param name="__selectionOnly">if set to <c>true</c> Export Selection Only</param>
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
			catch //(Exception __e)
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

		/// <summary>
		/// Saves the screen selection.
		/// </summary>
		/// <param name="__docRef">Document Refrence</param>
		/// <param name="__doc">Document</param>
		/// <param name="__options">JPEG options</param>
		private void saveScreenSelection(Photoshop.Document __docRef, Photoshop.Document __doc, Photoshop.JPEGSaveOptions __options)
		{
			saveScreenSelection(__docRef, __doc, -1, __options);
		}

		/// <summary>
		/// Saves the screen selection.
		/// </summary>
		/// <param name="__docRef">Document Refrence</param>
		/// <param name="__doc">Document</param>
		/// <param name="__idx">Layer Comp Index</param>
		/// <param name="__options">JPEG options</param>
		private void saveScreenSelection(Photoshop.Document __docRef, Photoshop.Document __doc, int __idx, Photoshop.JPEGSaveOptions __options)
		{
			Photoshop.Channel __selChannel;
			object __selBounds;
			string __fileNameBody;
			bool __hasSelection = false;
			Photoshop.ActionDescriptor __desc = new Photoshop.ActionDescriptor();
			Photoshop.ActionReference __ref = new Photoshop.ActionReference();
			Photoshop.ArtLayer __layer;
			int __j;

			try
			{
				__selChannel = __doc.Channels["screen"];
				__hasSelection = true;
				//MessageBox.Show(__selChannel.Name)
				__doc.Selection.Load(__selChannel, null, null);
				__selBounds = __doc.Selection.Bounds;
				__doc.Crop(__selBounds, null, null, null);
			}
			catch (Exception)
			{
				__hasSelection = false;
				
			}
			if (!__hasSelection)
			{
				for (__j = 1; __j <= __docRef.Layers.Count; __j++)
				{
					try
					{
						__layer = (Photoshop.ArtLayer)__docRef.Layers[__j];

						if (__layer.Name == "@screen")
						{
							__selBounds = __layer.Bounds;
							__doc.Crop(__selBounds, null, null, null);
							__hasSelection = true;
							break;
						}
					}
					catch
					{
						__hasSelection = false;
					}
				}
			}

			if (__hasSelection)
			{
				__fileNameBody = (__docRef.Name.LastIndexOf(".") > -1) ? __docRef.Name.Substring(0, __docRef.Name.LastIndexOf(".")) : __docRef.Name;
				__fileNameBody += (__idx <= -1) ? "_screen" : "." + __idx + "_screen";
				__fileNameBody += ".jpg";
				__doc.SaveAs(__docRef.Path + __fileNameBody, __options, true, null);
			}
			else
			{
				MessageBox.Show("You have to create a selection named \"screen\" or a layer named \"@screen\"", "No selection found", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
				Application.Exit();
			}

			if (__idx > 0)
				__doc.Close(2);
		}

		/// <summary>
		/// Archives the files.
		/// </summary>
		/// <param name="__docRef">The __doc ref.</param>
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

		/// <summary>
		/// Determines whether [has screen selection] [the specified __doc].
		/// </summary>
		/// <param name="__doc">The __doc.</param>
		/// <returns>
		///   <c>true</c> if [has screen selection] [the specified __doc]; otherwise, <c>false</c>.
		/// </returns>
		private bool hasScreenSelection(Photoshop.Document __doc)
		{
			bool __value;
			Photoshop.Channel __selChannel;
			Photoshop.ActionDescriptor __desc = new Photoshop.ActionDescriptor();
			Photoshop.ActionReference __ref = new Photoshop.ActionReference();

			try
			{
				__selChannel = __doc.Channels["screen"];
				__value = true;
			}
			catch (Exception)
			{
				__value = false;
			}

			if (!__value)
			{
				try
				{
					__ref.PutName(__appRef.CharIDToTypeID("Lyr "), "@screen");
					__desc.PutReference(__appRef.CharIDToTypeID("null"), __ref);
					__desc.PutEnumerated(__appRef.StringIDToTypeID("selectionModifier"), __appRef.StringIDToTypeID("selectionModifierType"), __appRef.StringIDToTypeID("removeFromSelection"));
					__desc.PutBoolean(__appRef.CharIDToTypeID("MkVs"), true);
					__appRef.ExecuteAction(__appRef.CharIDToTypeID("slct"), __desc, Photoshop.PsDialogModes.psDisplayNoDialogs);
					__value = true;
				}
				catch (Exception)
				{
					__value = false;
				}
			}

			return __value;
		}

		/// <summary>
		/// Determines whether [is exclude directory] [the specified __dir name].
		/// </summary>
		/// <param name="__dirName">Name of the __dir.</param>
		/// <returns>
		///   <c>true</c> if [is exclude directory] [the specified __dir name]; otherwise, <c>false</c>.
		/// </returns>
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

		/// <summary>
		/// Determines whether [is old file version] [the specified __filename].
		/// </summary>
		/// <param name="__filename">The __filename.</param>
		/// <param name="_version">The _version.</param>
		/// <returns>
		///   <c>true</c> if [is old file version] [the specified __filename]; otherwise, <c>false</c>.
		/// </returns>
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

		/// <summary>
		/// Exports image to base64 encoding.
		/// </summary>
		private void exportBase64()
		{
			//MessageBox.Show(__args(2))
			string __path = null;
			FileInfo __fi;
			string __fn;
			string __ext;

			string __filepath = __cmdargs[1 + __form.idx];
			
			//MessageBox.Show(__filepath);
			if (__filepath.LastIndexOf("\\") > -1)
			{
				__path = __filepath.Substring(0, Convert.ToInt32(__filepath.LastIndexOf("\\") + 1));
			}


			__fi = new FileInfo(__filepath);
			__fn = __fi.Name.Substring(0, __fi.Name.Length - __fi.Extension.Length);
			__ext = __fi.Extension.Substring(1);
			//MessageBox.Show(__ext)

			byte[] __bytes = File.ReadAllBytes(__filepath);

			string __b64String = Convert.ToBase64String(__bytes);
			string __dataUrl = "<html><body><img src=\"data:image/" + __ext + ";base64," + __b64String + "\"/></body></html>";

			File.WriteAllText(__path + __fn + ".html", __dataUrl);
		}

		private bool exportAssets(Photoshop.Document __docRef, object _layers)
		{
			Photoshop.Layers __layers;
			Photoshop.ArtLayer __alayer = null;
			Photoshop.LayerSet __slayer;
			bool __isArtLayer = false;
			object __layer;
			int __j;


			__layers = (Photoshop.Layers)_layers;
			for (__j = 1; __j <= __layers.Count; __j++)
			{
				__appRef.ActiveDocument = __docRef;
				__layer = __layers[__j];

				try
				{
					__alayer = (Photoshop.ArtLayer)__layer;
					__isArtLayer = true;
				}
				catch //(Exception __e)
				{
					__isArtLayer = false;
				}

				if (__isArtLayer) // Everything as Layer goes here
				{
					__appRef.ActiveDocument.ActiveLayer = __layer;
					if (__alayer.Name.IndexOf(".png") > -1)
					{
						saveAsset(__docRef, __alayer.Name);
					}
				}
				else
				{
					try
					{
						__slayer = (Photoshop.LayerSet)__layer;
						__appRef.ActiveDocument.ActiveLayer = __layer;
						
						if (__slayer.LayerType == Photoshop.PsLayerType.psLayerSet)
						{
							
							if (__slayer.Name.IndexOf(".png") > -1)
							{
								saveAsset(__docRef, __slayer.Name, __slayer.Layers);
							}
							else
							{
								exportAssets(__docRef, __slayer.Layers);
							}
						}
					}
					catch //(Exception __e)
					{
						//MessageBox.Show("error>" + __e.Message);
					}
				}
			}
			return true;
		}

		/*private void createSO()
		{
			Photoshop.ActionDescriptor __desc = __appRef.ExecuteAction(__appRef.StringIDToTypeID("newPlacedLayer"), null, Photoshop.PsDialogModes.psDisplayNoDialogs);
			
			Photoshop.ActionReference __ref = new Photoshop.ActionReference();
			Photoshop.ActionDescriptor __desc2 = new Photoshop.ActionDescriptor();
			__ref.PutEnumerated(__appRef.CharIDToTypeID("Lyr "), __appRef.CharIDToTypeID("Ordn"), __appRef.CharIDToTypeID("Trgt"));
			__desc2.PutReference(__appRef.CharIDToTypeID("null"), __ref);
			__appRef.ExecuteAction(__appRef.StringIDToTypeID("rasterizeLayer"), __desc2, Photoshop.PsDialogModes.psDisplayNoDialogs);

			
			Photoshop.ActionDescriptor __desc3 = new Photoshop.ActionDescriptor();
			Photoshop.ActionReference __ref3 = new Photoshop.ActionReference();
			__ref3.PutProperty(__appRef.CharIDToTypeID("Chnl"), __appRef.CharIDToTypeID("fsel"));
			__desc3.PutReference(__appRef.CharIDToTypeID("null"), __ref3);
			__desc3.PutEnumerated(__appRef.CharIDToTypeID("T   "), __appRef.CharIDToTypeID("Ordn"), __appRef.CharIDToTypeID("Al  "));
			__appRef.ExecuteAction(__appRef.CharIDToTypeID("setd"), __desc3, Photoshop.PsDialogModes.psDisplayNoDialogs);
			
			__appRef.ExecuteAction(__appRef.CharIDToTypeID("copy"), null, Photoshop.PsDialogModes.psDisplayNoDialogs);
				
			//return true;
		}*/

		private void saveAsset(Photoshop.Document __docRef, string __name)
		{
			saveAsset(__docRef, __name, null);
		}

		private void saveAsset(Photoshop.Document __docRef, string __name, object __layers)
		{
			Photoshop.Document __duppedDocument;
			object __bounds = null;

			Photoshop.ExportOptionsSaveForWeb __pngExportOptionsSaveForWeb = new Photoshop.ExportOptionsSaveForWeb();
			__pngExportOptionsSaveForWeb.Format = Photoshop.PsSaveDocumentType.psPNGSave; // 13;
			__pngExportOptionsSaveForWeb.PNG8 = false;
			__pngExportOptionsSaveForWeb.Transparency = true;

			if (!Directory.Exists(__docRef.Path + "+ Assets\\"))
			{
				Directory.CreateDirectory(__docRef.Path + "+ Assets\\");
			}

			if (__layers != null)
				__bounds = checkBounds(__layers);

			Photoshop.ActionDescriptor __desc = __appRef.ExecuteAction(__appRef.StringIDToTypeID("newPlacedLayer"), null, Photoshop.PsDialogModes.psDisplayNoDialogs);
			
			__duppedDocument = __docRef.Duplicate(__name, null);
					
			__appRef.ActiveDocument = __duppedDocument;

			moveLayer(__name);
			selectAllLayers();
			deselectLayer(__name);
			hideAllLayers();
			deselectLayers();
			//System.Threading.Thread.Sleep(2000);
			if (__bounds == null)
			{
				__duppedDocument.Trim(Photoshop.PsTrimType.psTransparentPixels, true, true, true, true);
			}
			else
			{
				__duppedDocument.Crop(__bounds, null, null, null);
			}

			string __tempName = (__name.IndexOf("@1x") > -1) ? Regex.Replace(__name, "@1x", "") : __name;
			__duppedDocument.Export(__docRef.Path + "+ Assets\\" + __tempName, 2, __pngExportOptionsSaveForWeb);

			if (__name.IndexOf("@2x") > -1)
			{
				//__duppedDocument.ResizeImage(__duppedDocument.Width / 2, __duppedDocument.Width / 2, null, Photoshop.PsResampleMethod.psBicubicSmoother);
				resizeImage(__duppedDocument.Width / 2);
				try
				{
					__duppedDocument.Export(__docRef.Path + "+ Assets\\" + Regex.Replace(__name, "@2x", ""), 2, __pngExportOptionsSaveForWeb);
				}
				catch
				{
					MessageBox.Show("Please allow to save all slices in Save For Web options");
				}
			}
			else if (__name.IndexOf("@1x") > -1)
			{
				//__duppedDocument.ResizeImage(__duppedDocument.Width * 2, __duppedDocument.Width * 2, null, Photoshop.PsResampleMethod.psBicubicSmoother);
				resizeImage(__duppedDocument.Width * 2);
				try
				{
					__duppedDocument.Export(__docRef.Path + "+ Assets\\" + Regex.Replace(__name, "@1x", "@2x"), 2, __pngExportOptionsSaveForWeb);
				}
				catch
				{
					MessageBox.Show("Please allow to save all slices in Save For Web options");
				}
			}
			//__duppedDocument.Close(2);
		}

		private void deselectLayers()
		{
			Photoshop.ActionDescriptor __desc = new Photoshop.ActionDescriptor();
			Photoshop.ActionReference __ref = new Photoshop.ActionReference();

			__ref.PutEnumerated(__appRef.CharIDToTypeID("Lyr "), __appRef.CharIDToTypeID("Ordn"), __appRef.CharIDToTypeID("Trgt"));
			__desc.PutReference(__appRef.CharIDToTypeID("null"), __ref);
			__appRef.ExecuteAction(__appRef.StringIDToTypeID("selectNoLayers"), __desc, Photoshop.PsDialogModes.psDisplayNoDialogs);
		}

		private void selectAllLayers()
		{
			Photoshop.ActionDescriptor __desc = new Photoshop.ActionDescriptor();
			Photoshop.ActionReference __ref = new Photoshop.ActionReference();

			__ref.PutEnumerated(__appRef.CharIDToTypeID("Lyr "), __appRef.CharIDToTypeID("Ordn"), __appRef.CharIDToTypeID("Trgt"));
			__desc.PutReference(__appRef.CharIDToTypeID("null"), __ref);
			__appRef.ExecuteAction(__appRef.StringIDToTypeID("selectAllLayers"), __desc, Photoshop.PsDialogModes.psDisplayNoDialogs);
		}

		private void deselectLayer(string __name)
		{
			Photoshop.ActionDescriptor __desc = new Photoshop.ActionDescriptor();
			Photoshop.ActionReference __ref = new Photoshop.ActionReference();
			
			__ref.PutName(__appRef.CharIDToTypeID("Lyr "), __name);
			__desc.PutReference(__appRef.CharIDToTypeID("null"), __ref);
			__desc.PutEnumerated(__appRef.StringIDToTypeID("selectionModifier"), __appRef.StringIDToTypeID("selectionModifierType"), __appRef.StringIDToTypeID("removeFromSelection"));
			__desc.PutBoolean(__appRef.CharIDToTypeID("MkVs"), false);
			__appRef.ExecuteAction(__appRef.CharIDToTypeID("slct"), __desc, Photoshop.PsDialogModes.psDisplayNoDialogs);

		}

		private void moveLayer(string __name)
		{
			Photoshop.ActionReference __ref = new Photoshop.ActionReference();
			Photoshop.ActionDescriptor __desc = new Photoshop.ActionDescriptor();
			__ref.PutProperty( __appRef.CharIDToTypeID("Prpr") , __appRef.CharIDToTypeID("NmbL"));
			__ref.PutEnumerated(__appRef.CharIDToTypeID("Dcmn"), __appRef.CharIDToTypeID("Ordn"), __appRef.CharIDToTypeID("Trgt"));
			int __nblayers = __appRef.ExecuteActionGet(__ref).GetInteger(__appRef.CharIDToTypeID("NmbL"));
   
			Photoshop.ActionDescriptor __desc1 = new Photoshop.ActionDescriptor();
			Photoshop.ActionReference __ref1 = new Photoshop.ActionReference();
			Photoshop.ActionReference __ref2 = new Photoshop.ActionReference();

			__ref1.PutEnumerated(__appRef.CharIDToTypeID("Lyr "), __appRef.CharIDToTypeID("Ordn"), __appRef.CharIDToTypeID("Trgt"));
			__desc1.PutReference(__appRef.CharIDToTypeID("null"), __ref1);

			__ref2.PutIndex(__appRef.CharIDToTypeID("Lyr "), __nblayers);
			__desc1.PutReference(__appRef.CharIDToTypeID("T   "), __ref2);

			__desc1.PutBoolean(__appRef.CharIDToTypeID("Adjs"), false);
			__desc1.PutInteger(__appRef.CharIDToTypeID("Vrsn"), 5);
			__appRef.ExecuteAction(__appRef.CharIDToTypeID("move"), __desc1, Photoshop.PsDialogModes.psDisplayNoDialogs);
		}

		private void hideAllLayers()
		{
			Photoshop.ActionReference __ref = new Photoshop.ActionReference();
			__ref.PutEnumerated(__appRef.CharIDToTypeID("Lyr "), __appRef.CharIDToTypeID("Ordn"), __appRef.CharIDToTypeID("Trgt"));
			Photoshop.ActionList __list = new Photoshop.ActionList();
			__list.PutReference(__ref);
			Photoshop.ActionDescriptor __desc = new Photoshop.ActionDescriptor();
			__desc.PutList(__appRef.CharIDToTypeID("null"), __list);
			__appRef.ExecuteAction(__appRef.CharIDToTypeID("Hd  "), __desc, Photoshop.PsDialogModes.psDisplayNoDialogs);
		}

		private object checkBounds(object _layers)
		{
			Photoshop.Layers __layers;
			object __layer;
			Photoshop.ArtLayer __alayer;
			object __res = null;
			int __j;

			__layers = (Photoshop.Layers)_layers;
			for (__j = 1; __j <= __layers.Count; __j++)
			{
				__layer = __layers[__j];
				try
				{
					__alayer = (Photoshop.ArtLayer)__layer;

					if (__alayer.Name.IndexOf("@bounds", StringComparison.InvariantCultureIgnoreCase) > -1)
					{
						__res = __alayer.Bounds;
						break;
					}
				}
				catch
				{
				}
				
			}
			return __res;
		}

		private void resizeImage(double __width)
		{
			Photoshop.ActionDescriptor __desc = new Photoshop.ActionDescriptor();

			__desc.PutUnitDouble(__appRef.CharIDToTypeID("Wdth"), __appRef.CharIDToTypeID("#Pxl"), __width);
			__desc.PutBoolean(__appRef.StringIDToTypeID("scaleStyles"), true);
			__desc.PutBoolean(__appRef.CharIDToTypeID("CnsP"), true);
			__desc.PutEnumerated(__appRef.CharIDToTypeID("Intr"), __appRef.CharIDToTypeID("Intp"), __appRef.CharIDToTypeID("Bcbc"));
			__appRef.ExecuteAction(__appRef.CharIDToTypeID("ImgS"), __desc, Photoshop.PsDialogModes.psDisplayNoDialogs);
		}

	}	
}
