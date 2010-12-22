using System;
using System.ComponentModel;
using System.Collections.Specialized;
using System.Windows.Forms;
using System.Text;
using PluginCore.Localization;
using System.Collections;

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
