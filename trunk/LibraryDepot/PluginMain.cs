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
using System.Text.RegularExpressions;
using System.Collections.Generic;
using ProjectManager.Controls.TreeView;
using ProjectManager.Projects;
using System.Collections;
using System.Collections.Specialized;
using System.Reflection;
using System.Resources;
using ICSharpCode.SharpZipLib.Zip;
using ICSharpCode.SharpZipLib.Zip.Compression.Streams;

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
        
	    #region Required Properties

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
            this.AddEventHandlers();
            this.InitLocalization();
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
                case EventType.UIStarted:
                    //MessageBox.Show(sender.GetType().ToString());
                    //EditorMenu(sender);
                    CheckEditorMenuRequirement(sender);
                    break;
                case EventType.FileSwitch:
                    //string cmd2 = (e as DataEvent).Action;
                    //MessageBox.Show(cmd2);
                    CheckEditorMenuRequirement(sender);
                    break;
                case EventType.Command:
                    string cmd = (e as DataEvent).Action;
                    //MessageBox.Show(cmd);
					if (cmd == "ProjectManager.TreeSelectionChanged")
					{
					    NodeSelected(sender);
					}
                    break;
            }
		}

        /// <summary>
        /// Nodes the selected.
        /// </summary>
        /// <param name="sender">The sender.</param>
        private void NodeSelected(Object sender)
        {
            //int _Idx = 3;

			ToolStripMenuItem tsmi = new ToolStripMenuItem(LocaleHelper.GetString("LibraryDepot.Menu.Title"), PluginBase.MainForm.FindImage("198"));

			BuildDirectoryMenu(tsmi, __LibraryPath);
			BuildFileMenu(tsmi, __LibraryPath);

			(sender as Control).ContextMenuStrip.Items.Insert(0, tsmi);
			//(sender as Control).ContextMenuStrip.Items.Insert(1, new ToolStripSeparator());
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
            this.pluginImage = PluginBase.MainForm.FindImage("99");
            
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
        }

        /// <summary>
        /// Loads the plugin settings
        /// </summary>
        public void LoadSettings()
        {
            settingObject = new Settings();
            if (!File.Exists(this.settingFilename)) this.SaveSettings();
            else
            {
                Object obj = ObjectSerializer.Deserialize(this.settingFilename, settingObject);
                settingObject = (Settings)obj;
            }
            
            //settingObject.Changed += SettingChanged;
            
        }

        /// <summary>
        /// Saves the plugin settings
        /// </summary>
        public void SaveSettings()
        {
            //settingObject.Changed -= SettingChanged;
			//if (settingObject.LibraryPath == "")
			//{
			//    MessageBox.Show(GetPath());
			//    //settingObject.LibraryPath = GetPath() + "/libraries/";
			//}
            ObjectSerializer.Serialize(this.settingFilename, settingObject);
        }

        /// <summary>
        /// Settings the changed.
        /// </summary>
        /// <param name="type">The type.</param>
        /// <param name="add">if set to <c>true</c> [add].</param>
        private void SettingChanged(string type, bool add)
        {
            
        }

        private void CheckEditorMenuRequirement(Object sender)
        {
			bool __foundMenu = false;
			try
			{
				ToolStripItemCollection __tsi = (sender as IMainForm).EditorMenu.Items;//.Find(LocaleHelper.GetString("TraceUtil.Menu.Title"), true);
				//MessageBox.Show("moo");
				foreach (ToolStripItem __Item in __tsi)
				{
					if (__Item.Text == LocaleHelper.GetString("LibraryDepot.Menu.Title"))
					{
						__foundMenu = true;
						//__tsi.Remove(__Item);
					}
				}
				//MessageBox.Show(__foundMenu.ToString());
				if (!__foundMenu)
					AddMenu(sender, "EditorMenu");
			}
			catch (Exception ex)
			{
				//MessageBox.Show(ex.Message);
			}
        }

        private void AddMenu(Object sender, string type)
        {
			//MessageBox.Show(type);
			ToolStripMenuItem tsmi = new ToolStripMenuItem(LocaleHelper.GetString("LibraryDepot.Menu.Title"), PluginBase.MainForm.FindImage("198"));

			__LibraryPath = settingObject.LibraryPath;
			if (__LibraryPath == "")
			{
				__LibraryPath = GetPath(PathHelper.AppDir) + "/FlashDevelop/Depot/";
				settingObject.LibraryPath = __LibraryPath;
			}
			//MessageBox.Show(__LibraryPath);
			BuildDirectoryMenu(tsmi, __LibraryPath);
			BuildFileMenu(tsmi, __LibraryPath);

			switch (type)
			{
				case "EditorMenu":
					(sender as IMainForm).EditorMenu.Items.Insert(0, tsmi);
					//(sender as IMainForm).EditorMenu.Items.Insert(1, new ToolStripSeparator());
					break;
				case "EditMenu":
					ToolStripMenuItem editMenu = (ToolStripMenuItem)PluginBase.MainForm.FindMenuItem("EditMenu");
					//MessageBox.Show(editMenu.Text);
					editMenu.DropDownItems.Insert(0, tsmi);
					//editMenu.DropDownItems.Insert(1, new ToolStripSeparator());
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
				BuildFileMenu(tsmiDir, __DirectoryPath);
				//tsmi.DropDownItems.Add(new ToolStripMenuItem(__DirectoryInfo.Name, null));
				__menu.DropDownItems.Add(tsmiDir);				
			}
		}

		private void BuildFileMenu(ToolStripMenuItem __menu, string __path)
		{
			string[] __Files = Directory.GetFiles(__path);
			foreach (string __FilePath in __Files)
			{
				//MessageBox.Show(__FilePath);
				FileInfo __FileInfo = new FileInfo(__FilePath);
				__menu.DropDownItems.Add(new ToolStripMenuItem(__FileInfo.Name, PluginBase.MainForm.FindImage("274"), new EventHandler(delegate
						{
							OutputLog(__FileInfo.Name);
							CopyLibrary(__FileInfo.FullName);
						}
					)));
			}
		}

		private void CopyLibrary(string __file)
		{
			//PluginBase.MainForm.EditorMenu.Hide();
			using (ZipInputStream s = new ZipInputStream(File.OpenRead(__file)))
			{

				ZipEntry theEntry;
				while ((theEntry = s.GetNextEntry()) != null)
				{

					string directoryName = Path.GetDirectoryName(theEntry.Name);
					string fileName = Path.GetFileName(theEntry.Name);

					// create directory
					if (directoryName.Length > 0)
					{
						Directory.CreateDirectory(GetPath() + "/src/" + directoryName);
					}

					if (fileName != String.Empty)
					{
						using (FileStream streamWriter = File.Create(GetPath() + "/src/" + theEntry.Name))
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

		#endregion

	}
	
}
