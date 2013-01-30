using System;
using Microsoft.Win32;

namespace PSTools
{
	class Installer
	{
		private RegistryKey __newKey;
		private Version __version = new Version();

		public void install()
		{
			installCommons();

			int __i = 0; 
			foreach (Version.Versions __item in Enum.GetValues(typeof(Version.Versions)))
			{
				if (__version.check(__item))
				{
					//MessageBox.Show((int)__item + " > ");
					installVersion((int)__item, Version.IllustratorVersions[__i]);
				}
				__i++;
			}
		}

		public void installCommons()
		{
			// JPEG BASE64
			__newKey = Registry.ClassesRoot.CreateSubKey("ACDSee Pro 4.jpg\\\\shell\\\\PSTools");
			__newKey.SetValue("MUIVerb", "Photoshop action...", RegistryValueKind.String);
			__newKey.SetValue("Icon", "\"" + System.Reflection.Assembly.GetExecutingAssembly().Location + "\",0", RegistryValueKind.String);
			__newKey.SetValue("SubCommands", "PSTools.Base64;PSTools.Config", RegistryValueKind.String);
			__newKey.Close();

			__newKey = Registry.ClassesRoot.CreateSubKey("jpegfile\\\\shell\\\\PSTools");
			__newKey.SetValue("MUIVerb", "Photoshop action...", RegistryValueKind.String);
			__newKey.SetValue("Icon", "\"" + System.Reflection.Assembly.GetExecutingAssembly().Location + "\",0", RegistryValueKind.String);
			__newKey.SetValue("SubCommands", "PSTools.Base64;PSTools.Config", RegistryValueKind.String);
			__newKey.Close();

			// GIF BASE64
			__newKey = Registry.ClassesRoot.CreateSubKey("ACDSee Pro 4.gif\\\\shell\\\\PSTools");
			__newKey.SetValue("MUIVerb", "Photoshop action...", RegistryValueKind.String);
			__newKey.SetValue("Icon", "\"" + System.Reflection.Assembly.GetExecutingAssembly().Location + "\",0", RegistryValueKind.String);
			__newKey.SetValue("SubCommands", "PSTools.Base64;PSTools.Config", RegistryValueKind.String);
			__newKey.Close();

			__newKey = Registry.ClassesRoot.CreateSubKey("giffile\\\\shell\\\\PSTools");
			__newKey.SetValue("MUIVerb", "Photoshop action...", RegistryValueKind.String);
			__newKey.SetValue("Icon", "\"" + System.Reflection.Assembly.GetExecutingAssembly().Location + "\",0", RegistryValueKind.String);
			__newKey.SetValue("SubCommands", "PSTools.Base64;PSTools.Config", RegistryValueKind.String);
			__newKey.Close();

			// PNG BASE64
			__newKey = Registry.ClassesRoot.CreateSubKey("ACDSee Pro 4.png\\\\shell\\\\PSTools");
			__newKey.SetValue("MUIVerb", "Photoshop action...", RegistryValueKind.String);
			__newKey.SetValue("Icon", "\"" + System.Reflection.Assembly.GetExecutingAssembly().Location + "\",0", RegistryValueKind.String);
			__newKey.SetValue("SubCommands", "PSTools.Base64;PSTools.Config", RegistryValueKind.String);
			__newKey.Close();

			__newKey = Registry.ClassesRoot.CreateSubKey("pngfile\\\\shell\\\\PSTools");
			__newKey.SetValue("MUIVerb", "Photoshop action...", RegistryValueKind.String);
			__newKey.SetValue("Icon", "\"" + System.Reflection.Assembly.GetExecutingAssembly().Location + "\",0", RegistryValueKind.String);
			__newKey.SetValue("SubCommands", "PSTools.Base64;PSTools.Config", RegistryValueKind.String);
			__newKey.Close();

			__newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.Base64");
			__newKey.SetValue("MUIVerb", "Export Base64 URI", RegistryValueKind.String);
			__newKey.SetValue("Icon", "shell32.dll,43", RegistryValueKind.String);
			__newKey.Close();

			__newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.Base64\\\\Command");
			__newKey.SetValue("", "\"" + System.Reflection.Assembly.GetExecutingAssembly().Location + "\" \"-b64\" \"%1\" \"jpg\" \"index\" \"12\"", RegistryValueKind.String);
			__newKey.Close();
		}

		public void installVersion(int __psVersion, string __aiVersion)
		{
			// Commands
			__newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.Image." + __psVersion + "\\\\shell\\\\PSTools");
			__newKey.SetValue("MUIVerb", "Photoshop action...", RegistryValueKind.String);
			__newKey.SetValue("Icon", "\"" + System.Reflection.Assembly.GetExecutingAssembly().Location + "\",0", RegistryValueKind.String);
			__newKey.SetValue("SubCommands", "PSTools.JPEGByIndex100;PSTools.Screen;PSTools.JPEGByIndex60;PSTools.JPEGByName100;PSTools.JPEGByName60;PSTools.PNGByIndex;PSTools.PNGByName;PSTools.GIFByIndex;PSTools.ImagesRights;PSTools.SO;PSTools.Clean;PSTools.Config", RegistryValueKind.String);
			__newKey.Close();

			__newKey = Registry.ClassesRoot.CreateSubKey("Photoshop.PSBFile." + __psVersion + "\\\\shell\\\\PSTools");
			__newKey.SetValue("MUIVerb", "Photoshop action...", RegistryValueKind.String);
			__newKey.SetValue("Icon", "\"" + System.Reflection.Assembly.GetExecutingAssembly().Location + "\",0", RegistryValueKind.String);
			__newKey.SetValue("SubCommands", "PSTools.JPEGByIndex100;PSTools.Screen;PSTools.JPEGByIndex60;PSTools.JPEGByName100;PSTools.JPEGByName60;PSTools.PNGByIndex;PSTools.PNGByName;PSTools.GIFByIndex;PSTools.ImagesRights;PSTools.SO;PSTools.Clean;PSTools.Config", RegistryValueKind.String);
			__newKey.Close();

			__newKey = Registry.ClassesRoot.CreateSubKey("Adobe.Illustrator.EPS\\\\shell\\\\PSTools");
			__newKey.SetValue("MUIVerb", "Photoshop action...", RegistryValueKind.String);
			__newKey.SetValue("Icon", "\"" + System.Reflection.Assembly.GetExecutingAssembly().Location + "\",0", RegistryValueKind.String);
			__newKey.SetValue("SubCommands", "PSTools.JPEGByIndex100;PSTools.JPEGByIndex60;PSTools.JPEGByName100;PSTools.JPEGByName60;PSTools.PNGByIndex;PSTools.PNGByName;PSTools.GIFByIndex;PSTools.Config", RegistryValueKind.String);
			__newKey.Close();

			__newKey = Registry.ClassesRoot.CreateSubKey("Adobe.Illustrator." + __aiVersion + "\\\\shell\\\\PSTools");
			__newKey.SetValue("MUIVerb", "Photoshop action...", RegistryValueKind.String);
			__newKey.SetValue("Icon", "\"" + System.Reflection.Assembly.GetExecutingAssembly().Location + "\",0", RegistryValueKind.String);
			__newKey.SetValue("SubCommands", "PSTools.JPEGByIndex100;PSTools.JPEGByIndex60;PSTools.JPEGByName100;PSTools.JPEGByName60;PSTools.PNGByIndex;PSTools.PNGByName;PSTools.GIFByIndex;PSTools.Config", RegistryValueKind.String);
			__newKey.Close();


			// PSTools.JPEGByIndex100
			__newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.JPEGByIndex100");
			__newKey.SetValue("MUIVerb", "Save Layer Comps As JPEG 100% (by index)", RegistryValueKind.String);
			__newKey.SetValue("Icon", "shell32.dll,43", RegistryValueKind.String);
			__newKey.Close();

			__newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.JPEGByIndex100\\\\command");
			__newKey.SetValue("", "\"" + System.Reflection.Assembly.GetExecutingAssembly().Location + "\" \"-s\" \"%1\" \"jpg\" \"index\" \"12\"", RegistryValueKind.String);
			__newKey.Close();

			// PSTools.JPEGByIndex60
			__newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.JPEGByIndex60");
			__newKey.SetValue("MUIVerb", "Save Layer Comps As JPEG 60% (by index)", RegistryValueKind.String);
			__newKey.SetValue("Icon", "shell32.dll,301", RegistryValueKind.String);
			__newKey.SetValue("CommandFlags", "32", RegistryValueKind.DWord);
			__newKey.Close();

			__newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.JPEGByIndex60\\\\command");
			__newKey.SetValue("", "\"" + System.Reflection.Assembly.GetExecutingAssembly().Location + "\" \"-s\" \"%1\" \"jpg\" \"index\" \"6\"", RegistryValueKind.String);
			__newKey.Close();

			// PSTools.JPEGByName100
			__newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.JPEGByName100");
			__newKey.SetValue("MUIVerb", "Save Layer Comps As JPEG 100% (by name)", RegistryValueKind.String);
			__newKey.SetValue("Icon", "shell32.dll,301", RegistryValueKind.String);
			//newKey.SetValue("Icon", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """,0", RegistryValueKind.String)
			__newKey.Close();

			__newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.JPEGByName100\\\\command");
			__newKey.SetValue("", "\"" + System.Reflection.Assembly.GetExecutingAssembly().Location + "\" \"-s\" \"%1\" \"jpg\" \"name\" \"12\"", RegistryValueKind.String);
			__newKey.Close();

			// PSTools.JPEGByName60
			__newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.JPEGByName60");
			__newKey.SetValue("MUIVerb", "Save Layer Comps As JPEG 60% (by name)", RegistryValueKind.String);
			__newKey.SetValue("Icon", "shell32.dll,301", RegistryValueKind.String);
			//newKey.SetValue("Icon", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """,0", RegistryValueKind.String)
			__newKey.Close();

			__newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.JPEGByName60\\\\command");
			__newKey.SetValue("", "\"" + System.Reflection.Assembly.GetExecutingAssembly().Location + "\" \"-s\" \"%1\" \"jpg\" \"name\" \"6\"", RegistryValueKind.String);
			__newKey.Close();

			// PSTools.PNGByIndex
			__newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.PNGByIndex");
			__newKey.SetValue("MUIVerb", "Save Layer Comps As PNG (by index)", RegistryValueKind.String);
			__newKey.SetValue("Icon", "shell32.dll,301", RegistryValueKind.String);
			//newKey.SetValue("Icon", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """,0", RegistryValueKind.String)
			__newKey.Close();

			__newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.PNGByIndex\\\\command");
			__newKey.SetValue("", "\"" + System.Reflection.Assembly.GetExecutingAssembly().Location + "\" \"-s\" \"%1\" \"png\" \"index\"", RegistryValueKind.String);
			__newKey.Close();

			// PSTools.PNGByName
			__newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.PNGByName");
			__newKey.SetValue("MUIVerb", "Save Layer Comps As PNG (by name)", RegistryValueKind.String);
			__newKey.SetValue("Icon", "shell32.dll,301", RegistryValueKind.String);
			//newKey.SetValue("Icon", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """,0", RegistryValueKind.String)
			__newKey.Close();

			__newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.PNGByName\\\\command");
			__newKey.SetValue("", "\"" + System.Reflection.Assembly.GetExecutingAssembly().Location + "\" \"-s\" \"%1\" \"png\" \"name\"", RegistryValueKind.String);
			__newKey.Close();

			// PSTools.GIFByIndex
			__newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.GIFByIndex");
			__newKey.SetValue("MUIVerb", "Save Layer Comps As GIF (by index)", RegistryValueKind.String);
			__newKey.SetValue("Icon", "shell32.dll,301", RegistryValueKind.String);
			//newKey.SetValue("Icon", """" + System.Reflection.Assembly.GetExecutingAssembly.Location + """,0", RegistryValueKind.String)
			//__newKey.SetValue("CommandFlags", "32", RegistryValueKind.DWord)
			__newKey.Close();

			__newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.GIFByIndex\\\\command");
			__newKey.SetValue("", "\"" + System.Reflection.Assembly.GetExecutingAssembly().Location + "\" \"-s\" \"%1\" \"gif\" \"index\"", RegistryValueKind.String);
			__newKey.Close();

			// PSTools.ImagesRights
			__newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.ImagesRights");
			__newKey.SetValue("MUIVerb", "List Images Rights", RegistryValueKind.String);
			__newKey.SetValue("Icon", "shell32.dll,54", RegistryValueKind.String);
			__newKey.SetValue("CommandFlags", "32", RegistryValueKind.DWord);
			__newKey.Close();

			__newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.ImagesRights\\\\command");
			__newKey.SetValue("", "\"" + System.Reflection.Assembly.GetExecutingAssembly().Location + "\" \"-r\" \"%1\"", RegistryValueKind.String);
			__newKey.Close();

			// PSTools.SO
			__newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.SO");
			__newKey.SetValue("MUIVerb", "Export Smart Objects", RegistryValueKind.String);
			__newKey.SetValue("Icon", "shell32.dll,132", RegistryValueKind.String);
			__newKey.Close();

			__newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.SO\\\\command");
			__newKey.SetValue("", "\"" + System.Reflection.Assembly.GetExecutingAssembly().Location + "\" \"-so\" \"%1\"", RegistryValueKind.String);
			__newKey.Close();

			// PSTools.Clean
			__newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.Clean");
			__newKey.SetValue("MUIVerb", "Clean Layers Name", RegistryValueKind.String);
			//__newKey.SetValue("Icon", "shell32.dll,238", RegistryValueKind.String)
			//newKey.SetValue("CommandFlags ", "20", RegistryValueKind.DWord)
			__newKey.Close();

			__newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.Clean\\\\command");
			__newKey.SetValue("", "\"" + System.Reflection.Assembly.GetExecutingAssembly().Location + "\" \"-w\" \"%1\"", RegistryValueKind.String);
			__newKey.Close();

			// PSTools.Screen
			__newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.Screen");
			__newKey.SetValue("MUIVerb", "Save Screen Selection As JPEG", RegistryValueKind.String);
			__newKey.SetValue("Icon", "shell32.dll,196", RegistryValueKind.String);
			__newKey.Close();

			__newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.Screen\\\\command");
			__newKey.SetValue("", "\"" + System.Reflection.Assembly.GetExecutingAssembly().Location + "\" \"-sc\" \"%1\" \"sc\" \"index\"", RegistryValueKind.String);
			__newKey.Close();

			// PSTools.Config
			__newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.Config");
			__newKey.SetValue("MUIVerb", "Configuration", RegistryValueKind.String);
			__newKey.SetValue("Icon", "shell32.dll,21", RegistryValueKind.String);
			__newKey.SetValue("CommandFlags", "32", RegistryValueKind.DWord);
			__newKey.Close();

			__newKey = Registry.LocalMachine.CreateSubKey("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.Config\\\\command");
			__newKey.SetValue("", "\"" + System.Reflection.Assembly.GetExecutingAssembly().Location + "\" \"-c\"", RegistryValueKind.String);
			__newKey.Close();
		}

		public void uninstall()
		{
			uninstallCommons();
			
			int __i = 0;
			foreach (Version.Versions __item in Enum.GetValues(typeof(Version.Versions)))
			{
				if (__version.check(__item))
				{
					//MessageBox.Show((int)__item + " > " + Version.IllustratorVersions[__i].ToString());
					uninstallVersion((int)__item, Version.IllustratorVersions[__i]);
				}
				__i++;
			}
		}

		public void uninstallCommons()
		{
			// JPEG BASE64
			try
			{
				Registry.ClassesRoot.DeleteSubKeyTree("ACDSee Pro 4.jpg\\\\shell\\\\PSTools");
			}
			catch { }

			try
			{
				Registry.ClassesRoot.DeleteSubKeyTree("jpegfile\\\\shell\\\\PSTools");
			}
			catch { }

			// GIF BASE64
			try
			{
				Registry.ClassesRoot.DeleteSubKeyTree("ACDSee Pro 4.gif\\\\shell\\\\PSTools");
			}
			catch { }

			try
			{
				Registry.ClassesRoot.DeleteSubKeyTree("giffile\\\\shell\\\\PSTools");
			}
			catch { }

			// PNG BASE64
			try
			{
				Registry.ClassesRoot.DeleteSubKeyTree("ACDSee Pro 4.png\\\\shell\\\\PSTools");
			}
			catch { }

			try
			{
				Registry.ClassesRoot.DeleteSubKeyTree("pngfile\\\\shell\\\\PSTools");
			}
			catch { }

			try
			{
				Registry.LocalMachine.DeleteSubKeyTree("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.Base64");
			}
			catch { }
		}

		public void uninstallVersion(int __psVersion, string __aiVersion)
		{
			try
			{
				Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.Image." + __psVersion + "\\\\shell\\\\PSTools");
			}
			catch { }

			try
			{
				Registry.ClassesRoot.DeleteSubKeyTree("Photoshop.PSBFile." + __psVersion + "\\\\shell\\\\PSTools");
			}
			catch { }

			try
			{
				Registry.ClassesRoot.DeleteSubKeyTree("Adobe.Illustrator.EPS\\\\shell\\\\PSTools");
			}
			catch { }

			try
			{
				Registry.ClassesRoot.DeleteSubKeyTree("Adobe.Illustrator." + __aiVersion + "\\\\shell\\\\PSTools");
			}
			catch { }

			try
			{
				Registry.LocalMachine.DeleteSubKeyTree("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.JPEGByIndex100");
			}
			catch { }

			try
			{
				Registry.LocalMachine.DeleteSubKeyTree("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.JPEGByIndex60");
			}
			catch { }

			try
			{
				Registry.LocalMachine.DeleteSubKeyTree("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.JPEGByName100");
			}
			catch { }

			try
			{
				Registry.LocalMachine.DeleteSubKeyTree("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.JPEGByName60");
			}
			catch { }

			try
			{
				Registry.LocalMachine.DeleteSubKeyTree("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.PNGByIndex");
			}
			catch { }

			try
			{
				Registry.LocalMachine.DeleteSubKeyTree("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.PNGByName");
			}
			catch { }

			try
			{
				Registry.LocalMachine.DeleteSubKeyTree("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.GIFByIndex");
			}
			catch { }

			try
			{
				Registry.LocalMachine.DeleteSubKeyTree("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.ImagesRights");
			}
			catch { }

			try
			{
				Registry.LocalMachine.DeleteSubKeyTree("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.SO");
			}
			catch { }

			try
			{
				Registry.LocalMachine.DeleteSubKeyTree("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.Clean");
			}
			catch { }

			try
			{
				Registry.LocalMachine.DeleteSubKeyTree("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.Screen");
			}
			catch{}

			try
			{
				Registry.LocalMachine.DeleteSubKeyTree("SOFTWARE\\\\Microsoft\\\\Windows\\\\CurrentVersion\\\\Explorer\\\\CommandStore\\\\shell\\\\PSTools.Config");
			}
			catch { }
		}

	}
}
