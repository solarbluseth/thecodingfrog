using System;
using System.IO;
using System.Drawing;
using System.Windows.Forms;
using System.ComponentModel;
using WeifenLuo.WinFormsUI.Docking;
using LibraryDepot.Resources;
using PluginCore.Localization;
using PluginCore.Utilities;
using PluginCore.Managers;
using PluginCore.Helpers;
using PluginCore;
using ProjectManager.Projects;
using ProjectManager.Actions;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using ProjectManager.Controls.TreeView;
using System.Collections;
using System.Collections.Specialized;
using System.Reflection;
using System.Resources;
using ICSharpCode.SharpZipLib.Core;
using ICSharpCode.SharpZipLib.Zip;
//using ICSharpCode.SharpZipLib.Zip.Compression.Streams;
//using System.Timers;

namespace LibraryDepot
{
	

    public class PluginMain : IPlugin
	{
		private String pluginName = "LibraryDepot";
		private String pluginGuid = "01B1E6B4-DA6E-4d3a-A1E7-AD65E4238C37";
        private String pluginHelp = "www.flashdevelop.org/community/";
        private String pluginDesc = "Library Depot utility plugin for FlashDevelop 3";
        private String pluginAuth = "Jean-Louis PERSAT";
        private String settingFilename;
        private Settings settingObject;
        private Image pluginImage;
		private string __LibraryPath;
		private ProjectTreeView __ptv;
		private string[] __DisallowedDirectories = { ".svn", "_svn" };
		private string[] __AllowedFiles = { ".as" };
		//private System.Timers.Timer __timer;
		private bool _stop;
		private FileSystemWatcher watcher;
        
	    #region Required Properties

		/// <summary>
		/// Api level of the plugin
		/// </summary>
		public Int32 Api
		{
			get { return 1; }
		}
		
		/// <summary>
        /// Name of the plugin
        /// </summary> 
        public String Name
		{
			get { return this.pluginName; }
		}

        /// <summary>
        /// GUID of the plugin
        /// </summary>
        public String Guid
		{
			get { return this.pluginGuid; }
		}

        /// <summary>
        /// Author of the plugin
        /// </summary> 
        public String Author
		{
			get { return this.pluginAuth; }
		}

        /// <summary>
        /// Description of the plugin
        /// </summary> 
        public String Description
		{
			get { return this.pluginDesc; }
		}

        /// <summary>
        /// Web address for help
        /// </summary> 
        public String Help
		{
			get { return this.pluginHelp; }
		}

        /// <summary>
        /// Object that contains the settings
        /// </summary>
        [Browsable(false)]
        public Object Settings
        {
            get { return settingObject; }
        }
		
		#endregion
		
		#region Required Methods
		
		/// <summary>
		/// Initializes the plugin
		/// </summary>
		public void Initialize()
		{
            this.InitBasics();
            this.LoadSettings();
        }
		
		/// <summary>
		/// Disposes the plugin
		/// </summary>
		public void Dispose()
		{
            this.SaveSettings();
		}
		
		/// <summary>
		/// Handles the incoming events
		/// </summary>
		public void HandleEvent(Object sender, NotifyEvent e, HandlingPriority prority)
        {
            switch (e.Type)
            {
                case EventType.Command:
                    string cmd = (e as DataEvent).Action;
					if (cmd == "ProjectManager.TreeSelectionChanged")
					{
						__ptv = (ProjectTreeView)sender;
					    BuildContextMenu(sender);
					}
                    break;
            }
		}

        /// <summary>
        /// Nodes the selected.
        /// </summary>
        /// <param name="sender">The sender.</param>
        private void BuildContextMenu(Object sender)
        {
			//MessageBox.Show("BuildContextMenu");
			int __idx = 0;
			int __savedpos = 0;
			try
			{
				ToolStripItemCollection __tsi = __ptv.ContextMenuStrip.Items;
				foreach (ToolStripItem __Item in __tsi)
				{

					if (__Item.Text == LocaleHelper.GetString("LibraryDepot.Menu.Title"))
					{
						// removing menu
						__tsi.Remove(__Item);

						// removing separator
						__tsi.RemoveAt(__idx);

						// saving position for recreating treemenu
						__savedpos = __idx;
					}
					__idx++;
				}
			}
			catch (Exception ex)
			{
				//MessageBox.Show(ex.Message);
			}

			ToolStripMenuItem tsmi = new ToolStripMenuItem(LocaleHelper.GetString("LibraryDepot.Menu.Title"), PluginBase.MainForm.FindImage("205"));

			BuildDirectoryMenu(tsmi, __LibraryPath);
			BuildFileMenu(tsmi, __LibraryPath);

			(sender as Control).ContextMenuStrip.Items.Insert(__savedpos, tsmi);
			(sender as Control).ContextMenuStrip.Items.Insert(__savedpos + 1, new ToolStripSeparator());

			/*if (settingObject.CheckTimer != TimerCheckInterval.Never)
			{
				__timer.Start();
			}*/
        }

		#endregion

        #region Custom Methods
       
        /// <summary>
        /// Initializes important variables
        /// </summary>
        public void InitBasics()
        {
            String dataPath = Path.Combine(PathHelper.DataDir, "LibraryDepot");
            if (!Directory.Exists(dataPath)) Directory.CreateDirectory(dataPath);
            this.settingFilename = Path.Combine(dataPath, "Settings.fdb");
            this.pluginImage = PluginBase.MainForm.FindImage("106");
            
        }

        /// <summary>
        /// Adds the required event handlers
        /// </summary> 
        public void AddEventHandlers()
        {
            // Set events you want to listen (combine as flags)
			EventManager.AddEventHandler(this, EventType.Command | EventType.FileSwitch | EventType.UIStarted);
        }

        /// <summary>
        /// Initializes the localization of the plugin
        /// </summary>
        public void InitLocalization()
        {
            LocaleVersion locale = PluginBase.MainForm.Settings.LocaleVersion;
            switch (locale)
            {
                /*
                case LocaleVersion.fi_FI : 
                    // We have Finnish available... or not. :)
                    LocaleHelper.Initialize(LocaleVersion.fi_FI);
                    break;
                */
                default : 
                    // Plugins should default to English...
                    LocaleHelper.Initialize(LocaleVersion.en_US);
                    break;
            }
            this.pluginDesc = LocaleHelper.GetString("Info.Description");

			/*__timer = new System.Timers.Timer();
			__timer.Elapsed += new ElapsedEventHandler(HandleTimerComplete);
			if (settingObject.CheckTimer != TimerCheckInterval.Never)
			{
				__timer.Interval = (int)settingObject.CheckTimer;
			}*/
        }

		/*private void HandleTimerComplete(object sender, ElapsedEventArgs e)
		{
			//MessageBox.Show("HandleTimerComplete:" + __ptv.ContextMenuStrip.Visible.ToString());
			if (__ptv.ContextMenuStrip.Visible)
			{
				__timer.Stop();
				__timer.Start();
			}
			else
			{
				__timer.Stop();
				BuildContextMenu(__ptv);
			}
		}*/

        /// <summary>
        /// Loads the plugin settings
        /// </summary>
        public void LoadSettings()
        {
            settingObject = new Settings();
			if (!File.Exists(this.settingFilename))
			{
				this.SaveSettings(); 
				
				FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
				folderBrowserDialog.Description = "Select your library depot folder";
				if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
				{
					settingObject.LibraryPath = folderBrowserDialog.SelectedPath;
					__LibraryPath = folderBrowserDialog.SelectedPath;
				}
				else
				{
					__LibraryPath = GetPath(PathHelper.AppDir) + "\\FlashDevelop\\Depot\\";
					if (!Directory.Exists(__LibraryPath))
						Directory.CreateDirectory(__LibraryPath);
					settingObject.LibraryPath = __LibraryPath;
				}								
			}
			else
			{
				Object obj = ObjectSerializer.Deserialize(this.settingFilename, settingObject);
				settingObject = (Settings)obj;
				__LibraryPath = settingObject.LibraryPath;
			}

			//__LibraryPath = settingObject.LibraryPath;
			//if (__LibraryPath == "")
			//{
			//    __LibraryPath = GetPath(PathHelper.AppDir) + "\\FlashDevelop\\Depot\\";
			//    settingObject.LibraryPath = __LibraryPath;
			//}
            settingObject.Changed += SettingChanged;

			watcher = new FileSystemWatcher();
			watcher.Path = __LibraryPath;
			watcher.NotifyFilter = NotifyFilters.FileName | NotifyFilters.Size;

			watcher.Created += new FileSystemEventHandler(OnFileEvent);
			watcher.Deleted += new FileSystemEventHandler(OnFileEvent);
			watcher.Renamed += new RenamedEventHandler(OnRenameEvent);
			watcher.IncludeSubdirectories = true;
			watcher.EnableRaisingEvents = true; 

			this.AddEventHandlers();
			this.InitLocalization();
        }

		public void OnFileEvent(object sender, FileSystemEventArgs fsea)
		{
			BuildContextMenu(__ptv);
		}

		public void OnRenameEvent(Object source, RenamedEventArgs rea)
		{
			BuildContextMenu(__ptv);
		}



        /// <summary>
        /// Saves the plugin settings
        /// </summary>
        public void SaveSettings()
        {
            settingObject.Changed -= SettingChanged;
			ObjectSerializer.Serialize(this.settingFilename, settingObject);
        }

        /// <summary>
        /// Settings the changed.
        /// </summary>
        /// <param name="type">The type.</param>
        /// <param name="add">if set to <c>true</c> [add].</param>
        private void SettingChanged(string type)
        {
			// recreate menu on LibraryPath change
			switch(type)
			{
				case "LibraryPath":
					__LibraryPath = settingObject.LibraryPath; 
					BuildContextMenu(__ptv);
					break;
				/*case "CheckTimer":
					if (settingObject.CheckTimer != TimerCheckInterval.Never)
					{
						__timer.Interval = (int)settingObject.CheckTimer;
						__timer.Start();
					}
					else
					{
						__timer.Stop();
					}
					break;*/
				case "EnableWatcher":
					watcher.EnableRaisingEvents = settingObject.EnableWatcher;
					break;
			}
        }

		private void BuildDirectoryMenu(ToolStripMenuItem __menu, string __path)
		{
			if (!Directory.Exists(__path)) return;
			string[] __Directories = Directory.GetDirectories(__path);
			ToolStripMenuItem tsmiDir = null;
			foreach (string __DirectoryPath in __Directories)
			{
				DirectoryInfo __DirectoryInfo = new DirectoryInfo(__DirectoryPath);
				tsmiDir = new ToolStripMenuItem(__DirectoryInfo.Name, PluginBase.MainForm.FindImage("203"));
				BuildDirectoryMenu(tsmiDir, __DirectoryPath);
				BuildFileMenu(tsmiDir, __DirectoryPath, true);
				
				__menu.DropDownItems.Add(tsmiDir);				
			}
		}

		private void BuildFileMenu(ToolStripMenuItem __menu, string __path)
		{
			BuildFileMenu(__menu, __path, false);
		}

		private void BuildFileMenu(ToolStripMenuItem __menu, string __path, bool __checkexistingdir)
		{
			bool __hasItems = false;
			string[] __Files = Directory.GetFiles(__path);
			bool __isAllowed = false;
			foreach (string __FilePath in __Files)
			{
				FileInfo __FileInfo = new FileInfo(__FilePath);
				Image __picto = null;
				switch (__FileInfo.Extension.ToLower())
				{
					case ".zip":
						__isAllowed = true;
						__hasItems = true;
						//__picto = PluginBase.MainForm.FindImage("133");
						__picto = PluginBase.MainForm.FindImage("274");
						break;
					case ".swc":
						__isAllowed = true;
						__hasItems = true;
						//__picto = PluginBase.MainForm.FindImage("274");
						Assembly assembly = Assembly.GetExecutingAssembly();
						try
						{
							__picto = new Bitmap(assembly.GetManifestResourceStream("LibraryDepot.Resources.SwcFile.png"));
						}
						catch (Exception ex)
						{
							//MessageBox.Show(ex.Message);
						}
						break;
					default:
						__isAllowed = false;
						break;
				}
				if (__isAllowed)
					__menu.DropDownItems.Add(new ToolStripMenuItem(__FileInfo.Name, __picto, new EventHandler(delegate
						{
							OutputLog(__FileInfo.Name);
							ChooseCopyType(__FileInfo.FullName);
						}
					)));
			}
			if (__checkexistingdir)
				if (!__hasItems)
					__menu.Enabled = false;
		}

		private void ChooseCopyType(string __file)
		{
			FileInfo __FileInfo = new FileInfo(__file);
			switch (__FileInfo.Extension.ToLower())
			{
				case ".zip":
					ExtractLibrary(__file);
					break;
				case ".swc":
					CopyLibrary(__file);
					break;
			}
		}

		private void ExtractLibrary(string __file)
		{
			bool __foundSRC = false;

			using (ZipInputStream s = new ZipInputStream(File.OpenRead(__file)))
			{

				ZipEntry theEntry;
				while ((theEntry = s.GetNextEntry()) != null)
				{
					if (Path.GetDirectoryName(theEntry.Name).ToLower() == "src")
						__foundSRC = true;
				}
			}
			FastZipEvents events = new FastZipEvents();
			events.ProcessFile = ProcessFileMethod;
			FastZip __fastZip = new FastZip(events);
			string __fileFilter = @"+\.as$";
			string __directoryFilter;

			if (__foundSRC)
			{
				__directoryFilter = @"src";
				__fastZip.ExtractZip(__file, GetPath(), FastZip.Overwrite.Prompt, OverwritePrompt, __fileFilter, __directoryFilter, true);

				//MessageBox.Show(GetPath() + "/classes/");
				if (Directory.Exists(GetPath() + "/classes/"))
				{
					Directory.Move(GetPath() + "/src/", GetPath() + "/classes/");
				}
				else if (settingObject.SrcPath.Replace("/", "").Replace("\\", "").ToLower() != "src")
				{
					Directory.Move(GetPath() + "/src/", GetPath() + settingObject.SrcPath);
				}
			}
			else
			{
				__directoryFilter = null;
				//MessageBox.Show(Directory.Exists(GetPath() + "/classes/").ToString()); 
				if (Directory.Exists(GetPath() + "/classes/"))
				{
					__fastZip.ExtractZip(__file, GetPath() + "/classes/", FastZip.Overwrite.Prompt, OverwritePrompt, __fileFilter, __directoryFilter, true);
				}
				else
				{
					__fastZip.ExtractZip(__file, GetPath() + settingObject.SrcPath, FastZip.Overwrite.Prompt, OverwritePrompt, __fileFilter, __directoryFilter, true);
				}
			}
			


			//PluginBase.MainForm.EditorMenu.Hide();
			/*using (ZipInputStream s = new ZipInputStream(File.OpenRead(__file)))
			{

				ZipEntry theEntry;
				while ((theEntry = s.GetNextEntry()) != null)
				{

					string directoryName = Path.GetDirectoryName(theEntry.Name);
					if (isDirectoryAllowed(directoryName))
					{
						string fileName = Path.GetFileName(theEntry.Name);
						if (isFileAllowed(Path.GetExtension(theEntry.Name)))
						{

							// create directory
							if (directoryName.Length > 0)
							{
								Directory.CreateDirectory(GetPath() + settingObject.SrcPath + directoryName);
							}

							if (fileName != String.Empty)
							{
								using (FileStream streamWriter = File.Create(GetPath() + settingObject.SrcPath + theEntry.Name))
								{

									int size = 2048;
									byte[] data = new byte[2048];
									while (true)
									{
										size = s.Read(data, 0, data.Length);
										if (size > 0)
										{
											streamWriter.Write(data, 0, size);
										}
										else
										{
											break;
										}
									}
								}
							}
						}
					}
				}
			}*/
		}

		private bool OverwritePrompt(string fileName)
		{

			// In this method you can choose whether to overwrite a file.
			DialogResult dr = MessageBox.Show("Overwrite " + fileName, "Overwrite?", MessageBoxButtons.YesNoCancel);
			if (dr == DialogResult.Cancel)
			{
				_stop = true;
				// Must return true if we want to abort processing, so that the ProcessFileMethod will be called.
				// When the ProcessFileMethod sets ContinueRunning false, processing will immediately stop.
				return true;
			}
			return dr == DialogResult.Yes;
		}

		private void ProcessFileMethod(object sender, ScanEventArgs args)
		{

			string fileName = args.Name;
			// To stop all further processing, set args.ContinueRunning = false
			if (_stop)
			{
				args.ContinueRunning = false;
			}
		}


		private void CopyLibrary(string __file)
		{
			FileInfo __FileInfo = new FileInfo(__file);
			if (!Directory.Exists(GetPath() + settingObject.SwcPath))
				Directory.CreateDirectory(GetPath() + settingObject.SwcPath);
			File.Copy(__file, GetPath() + settingObject.SwcPath + __FileInfo.Name);
		}

		private void OutputLog(string __file)
		{
			TraceManager.Add(LocaleHelper.GetString("LibraryDepot.Log.Info") + __file);
		}

		private string GetPath()
		{
			return GetPath(PluginBase.CurrentProject.ProjectPath);
		}

		private string GetPath(string __path)
		{
			int __pos = __path.LastIndexOf("\\");
			string __realpath = __path.Substring(0, __pos + 1);
			return __realpath;
		}

		private bool isDirectoryAllowed(string __directory)
		{
			if (__directory.Length == 0)
				return false;

			foreach (string __item in __DisallowedDirectories)
			{
				if (__item.ToLower() == __directory.ToLower())
				{
					return false;
				}
			}
			return true;
		}

		private bool isFileAllowed(string __fileext)
		{
			foreach (string __item in __AllowedFiles)
			{
				if (__item.ToLower() == __fileext.ToLower())
				{
					return true;
				}
			}
			return false;
		}

		#endregion

	}
	
}
