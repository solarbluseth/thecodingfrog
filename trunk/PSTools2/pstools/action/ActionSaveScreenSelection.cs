using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PSTools
{
	class ActionSaveScreenSelection
	{
		private Photoshop.Document __doc;
		private Photoshop.JPEGSaveOptions __saveOptions;
		private int __count = 0;

		public ActionSaveScreenSelection(Photoshop.Document __docRef, Photoshop.JPEGSaveOptions __jpgSaveOptions)
		{
			__doc = __docRef;
			__saveOptions = __jpgSaveOptions;

			countScreenSelection(__doc);
		}

		private void countScreenSelection(Photoshop.Document __doc)
		{
			__count = 0;
			
			foreach (Photoshop.Channel __channel in __doc.Channels)
			{
				if (__channel.Name.Contains("screen"))
				{
					__count++;
				}
			}
		}

		public void saveAll()
		{
			Photoshop.Document __duppedDocument;
			object __selBounds;
			int __idx = 0;
			string __fileNameBody;

			foreach (Photoshop.Channel __channel in __doc.Channels)
			{
				if (__channel.Name.Contains("screen"))
				{
					__idx++;
					__duppedDocument = __doc.Duplicate(null, null);
					
					__duppedDocument.Selection.Load(__channel, null, null);
					__selBounds = __duppedDocument.Selection.Bounds;
					__duppedDocument.Crop(__selBounds, null, null, null);

					__fileNameBody = (__doc.Name.LastIndexOf(".") > -1) ? __doc.Name.Substring(0, __doc.Name.LastIndexOf(".")) : __doc.Name;
					__fileNameBody += (__idx <= -1) ? "_screen" : "." + __idx + "_screen";
					__fileNameBody += ".jpg";
					__duppedDocument.SaveAs(__doc.Path + __fileNameBody, __saveOptions, true, null);

					__duppedDocument.Close(2);
				}
			}
		}

		public bool hasSelection
		{
			get
			{
				return (__count > 0);
			}
		}

		public int count
		{
			get {
				return __count;
			}
		}
	}
}
