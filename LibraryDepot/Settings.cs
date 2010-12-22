using System;
using System.IO;
using System.Text;
using System.Collections;
using System.ComponentModel;
using System.Drawing.Design;
using System.Collections.Generic;
using System.Windows.Forms.Design;
using System.Windows.Forms;
using System.Xml.Serialization;
using PluginCore.Localization;
using PluginCore.Managers;
using PluginCore;

namespace LibraryDepot
{
	public delegate void SettingChangeHandler(string setting);
	
	[Serializable]
    public class Settings
    {
		public event SettingChangeHandler Changed;

		const string DEFAULT_LIBRARY_PATH = "";

		private string __librarypath = DEFAULT_LIBRARY_PATH;

		/// <summary> 
		/// Get and sets the __librarypath
		/// </summary>
		[DisplayName("Library path")]
		[DefaultValue(DEFAULT_LIBRARY_PATH)]
		[Editor(typeof(FolderNameEditor), typeof(UITypeEditor))]
		public string LibraryPath
		{
			get { return this.__librarypath; }
			set
			{
				this.__librarypath = value;
				FireChanged("LibraryPath");
			}
		}

		[Browsable(false)]
		private void FireChanged(string setting)
		{
			if (Changed != null)
				Changed(setting);
		}
    }

}
