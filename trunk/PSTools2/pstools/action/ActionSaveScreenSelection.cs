using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace PSTools
{
	class ActionSaveScreenSelection
	{
		private Photoshop.Document __doc;
		private Photoshop.JPEGSaveOptions __saveOptions;
		private List<List<string>> __channelsSelectionList = new List<List<string>> ();
		private List<List<string>> __boundsSelectionList = new List<List<string>> ();
		private int __countTotal = 0;
		private int __countChannels = 0;
		private int __countBounds = 0;
		private char[] __delimitersOne = new char[] { '=' };
		private char[] __delimitersTwo = new char[] { ',' };

		public ActionSaveScreenSelection(Photoshop.Document __docRef, Photoshop.JPEGSaveOptions __jpgSaveOptions)
		{
			__doc = __docRef;
			__saveOptions = __jpgSaveOptions;

			countSelections(__doc);
		}

		private void countSelections(Photoshop.Document __doc)
		{
			Photoshop.ArtLayer __layer;
			
			__countTotal = 0;
			__countChannels = 0;
			__countBounds = 0;
			
			foreach (Photoshop.Channel __channel in __doc.Channels)
			{
				if (__channel.Name.Contains("screen"))
				{
					__countChannels++;
					addToCompList(__channel.Name, __channelsSelectionList);			
				}
			}

			for (int __j = 0; __j < __doc.Layers.Count; __j++)
			{
				try
				{
					__layer = (Photoshop.ArtLayer)__doc.Layers[__j];
					if (__layer.Name.Contains("@screen"))
					{
						__countBounds++;
						addToCompList(__layer.Name, __boundsSelectionList);
					}
				}
				catch
				{
				}
			}

			__countTotal = __countChannels + __countBounds;
		}

		private void addToCompList(string __name, List<List<string>> __list)
		{
			List<string> __sublist = new List<string>();
			string[] __indexes = __name.Split(__delimitersOne, StringSplitOptions.RemoveEmptyEntries);
			if (__indexes.Length > 1)
			{
				string[] __indexesOne = __indexes[1].Split(__delimitersTwo, StringSplitOptions.RemoveEmptyEntries);

				foreach (string __index in __indexesOne)
				{
					//MessageBox.Show(__index.ToString());
					if (!__sublist.Contains(__index.ToString()))
						__sublist.Add(__index.ToString());
				}
			}
			__list.Add(__sublist);
			//MessageBox.Show(__name + ">" + __list.Count);
		}

		public void saveAll()
		{
			saveAll(0, 0);
		}

		public void saveAll(int __count, int __index)
		{
			Photoshop.Document __duppedDocument = null;
			Photoshop.ArtLayer __layer;
			object __selBounds;
			int __idx = 0;

			if (__countChannels > 0)
			{
				foreach (Photoshop.Channel __channel in __doc.Channels)
				{
					if (__channel.Name.Contains("screen"))
					{
						__idx++;
						try
						{
							__duppedDocument = __doc.Duplicate(null, null);
						}
						catch (Exception)
						{
							MessageBox.Show("Essayez de redémarrer UTC FMCore");
						}

						__duppedDocument.Selection.Load(__channel, null, null);
						__selBounds = __duppedDocument.Selection.Bounds;
						__duppedDocument.Crop(__selBounds, null, null, null);

						if (__count > 0)
						{
							// check if index is to be captured
							if (__channelsSelectionList[__idx - 1].Count > 0)
							{
								if (__channelsSelectionList[__idx - 1].Contains(__index.ToString()))
								{
									if (__countTotal > 1)
										saveSelection(__duppedDocument, __count, __index, __idx);
									else
										saveSelection(__duppedDocument, __count, __index, 0);
								}									
							}
							else
							{
								if (__countTotal > 1)
									saveSelection(__duppedDocument, __count, __index, __idx);
								else
									saveSelection(__duppedDocument, __count, __index, 0);
							}
						}
						else
						{
							saveSelection(__duppedDocument, __countTotal, __idx, 0);
						}
						__duppedDocument.Close(2);

					}
				}
			}

			int __idx2 = -1;
			if (__countBounds > 0)
			{
				for (int __j = 1; __j <= __doc.Layers.Count; __j++)
				{
					try
					{
						__layer = (Photoshop.ArtLayer)__doc.Layers[__j];

						if (__layer.Name.Contains("@screen"))
						{
							__idx++;
							__idx2++;
							try
							{
								//System.Threading.Thread.Sleep(1000); 
								__duppedDocument = __doc.Duplicate(null, null);
							}
							catch (Exception)
							{
								MessageBox.Show("Essayez de redémarrer UTC FMCore");
							}
							
							__selBounds = __layer.Bounds;
							__duppedDocument.Crop(__selBounds, null, null, null);

							if (__count > 0)
							{
								//MessageBox.Show("Layer comp " + __index + ":" + __boundsSelectionList[__idx2].Count);
								// check if index is to be captured
								//MessageBox.Show("array " + __idx2 + ":" + __boundsSelectionList[__idx2].Count);
								if (__boundsSelectionList[__idx2].Count > 0)
								{
									//MessageBox.Show("array " + __idx2 + ":" + __boundsSelectionList[__idx2].Count);
									if (__boundsSelectionList[__idx2].Contains(__index.ToString()))
									{
										if (__countTotal > 1)
											saveSelection(__duppedDocument, __count, __index, __idx);
										else
											saveSelection(__duppedDocument, __count, __index, 0);
									}
								}
								else
								{
									if (__countTotal > 1)
										saveSelection(__duppedDocument, __count, __index, __idx);
									else
										saveSelection(__duppedDocument, __count, __index, 0);
								}
							}
							else
							{
								saveSelection(__duppedDocument, __countTotal, __idx, 0);
							}

							__duppedDocument.Close(2);
							
						}
					}
					catch
					{
						
					}
				}
			}
		}

		private void saveSelection(Photoshop.Document __duppedDocument, int __count, int __idx)
		{
			saveSelection(__duppedDocument, __count, __idx, 0);
		}

		private void saveSelection(Photoshop.Document __duppedDocument, int __count, int __idx, int __subindex)
		{
			string __fileNameBody;
			string __sidx = string.Empty;

			if (__subindex > 0)
				__sidx = "." + __subindex;
			__fileNameBody = (__doc.Name.LastIndexOf(".") > -1) ? __doc.Name.Substring(0, __doc.Name.LastIndexOf(".")) : __doc.Name;
			__fileNameBody += (__count <= 1) ? "" : "." + __idx + __sidx + "";
			__fileNameBody += ".jpg";
			if (!Directory.Exists(__doc.Path + "/+ Screens/"))
				Directory.CreateDirectory(__doc.Path + "/+ Screens/");

			__duppedDocument.SaveAs(__doc.Path + "/+ Screens/" + __fileNameBody, __saveOptions, true, null);
		}

		public bool hasSelection
		{
			get
			{
				return (__countTotal > 0);
			}
		}

		public int count
		{
			get {
				return __countTotal;
			}
		}
	}
}
