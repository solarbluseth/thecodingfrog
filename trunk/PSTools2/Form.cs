using System.Diagnostics;
using System;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Reflection;


namespace PSTools
{
	public partial class Form
	{
		[STAThread]
		static void Main()
		{
			if (PriorProcess() != null)
			{

				MessageBox.Show("Another instance of the app is already running.");
				return;
			}

			AppDomain.CurrentDomain.AssemblyResolve += delegate(object sender, ResolveEventArgs e)
			{
				AssemblyName requestedName = new AssemblyName(e.Name);
				
				if (requestedName.Name == "Interop.Photoshop")
				{
					// Put code here to load whatever version of the assembly you actually have
					string __appPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
					//MessageBox.Show(__appPath); 
					Assembly __assembly = null;
					try
					{
						__assembly = Assembly.LoadFile(__appPath + "\\Interop.Photoshop.CS55.dll");
					}
					catch
					{
						try
						{
							__assembly = Assembly.LoadFile(__appPath + "\\Interop.Photoshop.CS5.dll");
						}
						catch
						{

							try
							{
								__assembly = Assembly.LoadFile(__appPath + "\\Interop.Photoshop.CS4.dll");
							}
							catch
							{
								MessageBox.Show("Can't load any assembly file : Interop.Photoshop<#>.dll");
								Application.Exit();
								Process.GetCurrentProcess().Kill();
							}
						}
					}
					return __assembly;
				}
				else
				{
					return null;
				}
			};

			System.Windows.Forms.Application.Run(new Form());
		}
		[DllImport("kernel32",EntryPoint="GetConsoleTitleA", ExactSpelling=true, CharSet=CharSet.Ansi, SetLastError=true)]
		public static extern int GetConsoleTitle(string lpConsoleTitle, int nSize);
		[DllImport("user32",EntryPoint="FindWindowA", ExactSpelling=true, CharSet=CharSet.Ansi, SetLastError=true)]
		public static extern int FindWindow(string lpClassName, string lpWindowName);
		[DllImport("user32",EntryPoint="ShowWindow", ExactSpelling=true, CharSet=CharSet.Ansi, SetLastError=true)]
		public static extern int ShowWindow(int hwnd, int nCmdShow);

		public const int SW_HIDE = 0;
		public const int SW_SHOWNORMAL = 1;
		public const int SW_NORMAL = 1;
		public const int SW_SHOWMINIMIZED = 2;
		private string[] __args;
		
		const string NS = "http://www.smartobjectlinks.com/1.0/";
		const string FOUND = "Found";
		const string NOT_FOUND = "Not found";
		private Settings __settings = new Settings();
		private Version __version = new Version();
		private ActionDispatcher __actionDispatcher;
		private Installer __installer = new Installer();
		private int __idx = 0;

		public static Process PriorProcess()
		// Returns a System.Diagnostics.Process pointing to
		// a pre-existing process with the same name as the
		// current one, if any; or null if the current process
		// is unique.
		{
			Process curr = Process.GetCurrentProcess();
			Process[] procs = Process.GetProcessesByName(curr.ProcessName);
			foreach (Process p in procs)
			{
				if ((p.Id != curr.Id) &&
					(p.MainModule.FileName == curr.MainModule.FileName))
					return p;
			}
			return null;
		}

		public Form()
		{
			// This call is required by the Windows Form Designer.
			InitializeComponent();

			//MessageBox.Show("debug");
			/*try
			{*/
				string __strTitle;
				int __rtnLen;
				int __hwnd;

				__strTitle = new StringBuilder().Append(' ', 256).ToString();

				__rtnLen = GetConsoleTitle(__strTitle, 256);
				if (__rtnLen > 0)
				{
					__strTitle = __strTitle.Substring(0, __rtnLen);
				}

				__hwnd = FindWindow(null, __strTitle);

				this.Text = (string)(System.Reflection.Assembly.GetExecutingAssembly().GetName().Name.ToString());
				this.LabelCompiled.Text = (string)("V " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Major.ToString() + "." + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Minor.ToString() + "." + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Build.ToString() + ", Compiled " + CompileDate.BuildDate.ToString());


				__settings.Load(this);

				if (__version.isInstalled())
				{
					__args = Environment.GetCommandLineArgs();
					// 0 if debug mode
					__idx = 1;
                    //__args = new string[] { "-s", "V:\\OTCLYON_OfficeTourismeCongrèsLyon\\LCC_LyonCityCard\\04_Design\\02_Desktop\\+ Elements\\etoiles.psb", "png", "index", "12" };
					//__args = new string[] { "-s", "E:\\AI\\02_liste_v04.psd", "jpg", "index", "12" };
					/*#if DEBUG
					#else
									MessageBox.Show("debug");
				
				
					#endif*/
					//MessageBox.Show(__args[0].ToString());
					int __num = __args.Length - 1;

					__actionDispatcher = new ActionDispatcher(this);
					int __windowState = __actionDispatcher.command(__args);
					//MessageBox.Show(__windowState.ToString());
					ShowWindow(__hwnd, __windowState);
					if (__windowState == SW_HIDE)
					{
						this.Visible = false;
						this.ShowInTaskbar = false;
						this.WindowState = FormWindowState.Minimized;
						Application.Exit();
						Process.GetCurrentProcess().Kill();
					}
					else
					{
						//MessageBox.Show("here");
						SyncUI();
					}
				}
				else
				{
					//MessageBox.Show("there");
					SyncUI();
				}
			/*}
			catch (Exception __e)
			{
				FileLogger.Instance.Open(@"C:\pstools.log", true); 
				FileLogger.Instance.CreateEntry(__e.Message);
				FileLogger.Instance.Close();
			}*/

			//MessageBox.Show("ok");
			
		}


		private void SyncUI()
		{
			if (__version.isInstalled())
			{
				this.Install.Text = "Uninstall";
			}
			else
			{
				this.Install.Text = "Install";
			}

			foreach (Version.Versions __item in Enum.GetValues(typeof(Version.Versions)))
			{
				string __name = Enum.GetName(typeof(Version.Versions),__item);
				//MessageBox.Show(__name);
				Control[] __ctrl = this.Controls.Find("Label" + __name, true);
				//__ctrl[0].Text = "test";
				if (__version.check(__item))
				{
					__ctrl[0].Text = FOUND.ToString();
				}
				else
				{
					__ctrl[0].Text = NOT_FOUND.ToString();
				}
			}
		}

		public void Install_Click(System.Object sender, System.EventArgs e)
		{

			if (__version.isInstalled())
			{
				__installer.uninstall();
				Install.Text = "Install";
			}
			else
			{
				__installer.install();
				Install.Text = "Uninstall";
			}
		}
		
		public void AutoArchive_CheckedChanged(System.Object sender, System.EventArgs e)
		{
			RegistryKey __newKey;
			__newKey = Registry.CurrentUser.CreateSubKey("Software\\\\PSTools");
			if (this.AutoArchive.Checked)
			{
				__newKey.SetValue("AutoArchive", "1", RegistryValueKind.String);
			}
			else
			{
				__newKey.SetValue("AutoArchive", "0", RegistryValueKind.String);
			}
			__newKey.Close();
		}
		
		public void ArchiveDirectory_TextChanged(System.Object sender, System.EventArgs e)
		{
			RegistryKey __newKey;
			__newKey = Registry.CurrentUser.CreateSubKey("Software\\\\PSTools");
			if (this.ArchiveDirectory.Text.Trim() != "")
			{
				__newKey.SetValue("ArchiveDirectory", this.ArchiveDirectory.Text, RegistryValueKind.String);
			}
			else
			{
				this.ArchiveDirectory.Text = "Archives";
				__newKey.SetValue("ArchiveDirectory", "Archives", RegistryValueKind.String);
			}
			__newKey.Close();
		}
		
		public void NamedExportQuality_ValueChanged(System.Object sender, System.EventArgs e)
		{
			RegistryKey __newKey;
			__newKey = Registry.CurrentUser.CreateSubKey("Software\\\\PSTools");
			if (__settings.SettingsLoaded)
			{
				//MessageBox.Show(Me.NamedExportQuality.Value.ToString())
				__newKey.SetValue("NamedExportQuality", this.NamedExportQuality.Value, RegistryValueKind.String);
			}
			__newKey.Close();
		}
		
		public void ExportLayerComps_CheckedChanged(System.Object sender, System.EventArgs e)
		{
			RegistryKey __newKey;
			__newKey = Registry.CurrentUser.CreateSubKey("Software\\\\PSTools");
			if (this.ExportLayerComps.Checked)
			{
				__newKey.SetValue("ExportLayerComps", "1", RegistryValueKind.String);
				__settings.doExportLayerComps = true;
			}
			else
			{
				__newKey.SetValue("ExportLayerComps", "0", RegistryValueKind.String);
				__settings.doExportLayerComps = false;
			}
			__newKey.Close();
		}
		
		public void ToolTip1_Popup(System.Object sender, System.Windows.Forms.PopupEventArgs e)
		{
			
		}
		
		public void ExportLayerComps_MouseHover(object sender, System.EventArgs e)
		{
			ToolTip1.Show("Export Layer Comps if checked, else it will save each layers", ExportLayerComps);
		}
		
		public void ExcludeDirectories_TextChanged(System.Object sender, System.EventArgs e)
		{
			RegistryKey __newKey;
			__newKey = Registry.CurrentUser.CreateSubKey("Software\\\\PSTools");
			if (this.ExcludeDirectories.Text.Trim() != "")
			{
				__newKey.SetValue("ExcludeDirectories", this.ExcludeDirectories.Text, RegistryValueKind.String);
			}
			else
			{
				this.ExcludeDirectories.Text = "";
				__newKey.SetValue("ExcludeDirectories", "", RegistryValueKind.String);
			}
			__newKey.Close();
		}

		public int idx
		{
			get { return __idx; }
		}

		private void label4_Click(object sender, EventArgs e)
		{

		}
		
		
	}
}
