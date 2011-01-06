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
	/*public enum TimerCheckInterval
	{
		Never = 0,
		Every30sec = 30000,
		Every60sec = 60000,
		Every2min = 120000,
		Every5min = 300000
	}*/

	[Serializable]
    public class Settings
    {
		public event SettingChangeHandler Changed;

		const string DEFAULT_LIBRARY_PATH = "";
		const string DEFAULT_SRC_PATH = "/src/";
		const string DEFAULT_SWC_PATH = "/lib/";
		//const TimerCheckInterval DEFAULT_TIMER = TimerCheckInterval.Never;
		const bool DEFAULT_ENABLE_WATCHER = true;

		private string __librarypath = DEFAULT_LIBRARY_PATH;
		private string __srcpath = DEFAULT_SRC_PATH;
		private string __swcpath = DEFAULT_SWC_PATH;
		//private TimerCheckInterval __checktimer = DEFAULT_TIMER;
		private bool __enablewatcher = DEFAULT_ENABLE_WATCHER;

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
		/// Get and sets the __srcpath
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
		/// Get and sets the __swcpath
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

		/// <summary> 
		/// Get and sets the __checktimer
		/// </summary>
		/*[DisplayName("Check library path")]
		[DefaultValue(DEFAULT_TIMER)]
		public TimerCheckInterval CheckTimer
		{
			get { return this.__checktimer; }
			set
			{
				this.__checktimer = value;
				FireChanged("CheckTimer");
			}
		}*/

		/// <summary> 
		/// Get and sets the __checktimer
		/// </summary>
		[DisplayName("Enable library watcher")]
		[DefaultValue(DEFAULT_ENABLE_WATCHER)]
		public bool EnableWatcher
		{
			get { return this.__enablewatcher; }
			set
			{
				this.__enablewatcher = value;
				FireChanged("EnableWatcher");
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
