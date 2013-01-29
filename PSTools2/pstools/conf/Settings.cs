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

		public Settings()
		{

		}

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

		public bool SettingsLoaded
		{
			get
			{
				return __settingsLoaded;
			}
		}

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
