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
		const string DEFAULT_SRC_PATH = "/src/";
		const string DEFAULT_SWC_PATH = "/lib/";

		private string __librarypath = DEFAULT_LIBRARY_PATH;
		private string __srcpath = DEFAULT_SRC_PATH;
		private string __swcpath = DEFAULT_SWC_PATH;

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

		/// <summary> 
		/// Get and sets the __librarypath
		/// </summary>
		[DisplayName("Src path")]
		[DefaultValue(DEFAULT_SRC_PATH)]
		public string SrcPath
		{
			get { return this.__srcpath; }
			set
			{
				this.__srcpath = value;
				FireChanged("SrcPath");
			}
		}

		/// <summary> 
		/// Get and sets the __librarypath
		/// </summary>
		[DisplayName("Swc path")]
		[DefaultValue(DEFAULT_SWC_PATH)]
		public string SwcPath
		{
			get { return this.__swcpath; }
			set
			{
				this.__swcpath = value;
				FireChanged("SwcPath");
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
