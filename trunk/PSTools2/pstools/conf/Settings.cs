using System;
using Microsoft.Win32;

namespace PSTools
{
	class Settings
	{
		private bool __doExportLayerComps = false;
		private RegistryKey __key = null;
		private bool __settingsLoaded = false;
		private Version __version = new Version();

		/// <summary>
		/// Initializes a new instance of the <see cref="Settings"/> class.
		/// </summary>
		public Settings()
		{

		}

		/// <summary>
		/// Loads settings
		/// </summary>
		/// <param name="__form">Form</param>
		public void Load(Form __form)
		{
			try
			{
				__key = Registry.CurrentUser.OpenSubKey("Software\\\\PSTools\\\\");
				if (__key.GetValue("AutoArchive").ToString() == "1")
				{
					__form.AutoArchive.Checked = true;
				}
				else
				{
					__form.AutoArchive.Checked = false;
				}

				//__key = Registry.CurrentUser.OpenSubKey("Software\\\\PSTools\\\\");
				if (__key.GetValue("ArchiveDirectory").ToString() != "")
				{
					__form.ArchiveDirectory.Text = __key.GetValue("ArchiveDirectory").ToString();
				}
				else
				{
					__form.ArchiveDirectory.Text = "Archives";
				}

				//__key = Registry.CurrentUser.OpenSubKey("Software\\\\PSTools\\\\");
				if (__key.GetValue("ExcludeDirectories").ToString() != "")
				{
					__form.ExcludeDirectories.Text = __key.GetValue("ExcludeDirectories").ToString();
				}
				else
				{
					__form.ExcludeDirectories.Text = "";
				}

				//__key = Registry.CurrentUser.OpenSubKey("Software\\\\PSTools\\\\");
				//MessageBox.Show(">>> " & key.GetValue("NamedExportQuality"))
				if (__key.GetValue("NamedExportQuality").ToString() != "")
				{
					__form.NamedExportQuality.Value = Convert.ToDecimal(__key.GetValue("NamedExportQuality"));
				}
				else
				{
					__form.NamedExportQuality.Value = 6;
				}

				//__key = Registry.CurrentUser.OpenSubKey("Software\\\\PSTools\\\\");
				//MessageBox.Show(">>> " & key.GetValue("ExportLayerComps"))
				if (__key.GetValue("ExportLayerComps").ToString() == "1")
				{
					__form.ExportLayerComps.Checked = true;
					__doExportLayerComps = true;
				}
				else
				{
					__form.ExportLayerComps.Checked = false;
					__doExportLayerComps = false;
				}

				if (__version.isInstalled())
				{
					__form.Install.Text = "Install";
				}
				else
				{
					__form.Install.Text = "Uninstall";
				}

				__settingsLoaded = true;
			}
			catch (Exception)
			{
			}
		}

		/// <summary>
		/// Gets a value indicating whether [settings loaded].
		/// </summary>
		/// <value>
		///   <c>true</c> if [settings loaded]; otherwise, <c>false</c>.
		/// </value>
		public bool SettingsLoaded
		{
			get
			{
				return __settingsLoaded;
			}
		}

		/// <summary>
		/// Gets or sets a value indicating whether [do export layer comps].
		/// </summary>
		/// <value>
		///   <c>true</c> if [do export layer comps]; otherwise, <c>false</c>.
		/// </value>
		public bool doExportLayerComps
		{
			get
			{
				return __doExportLayerComps;
			}
			set
			{
				__doExportLayerComps = value;
			}
		}
	}
}
