using System;
using System.Collections.Generic;
using Microsoft.Win32;
using System.Windows.Forms;

namespace PSTools
{
	class Version
	{
		private RegistryKey __key = null;
		private bool __isInstalled = false;
		private List<int> __versionsInstalled = new List<int>();

		public enum Versions
		{
		    CS3 = 10,
		    CS4 = 11,
		    CS5 = 12,
			CS55 = 55,
			CS6 = 13,
			CC = 14
		}

        public static List<string> OrderVersions = new List<string>() { "CS3", "CS4", "CS5", "CS55", "CS6", "CC" };
		public static List<string> IllustratorVersions = new List<string>() { "12", "13", "14", "15.1", "16", "17" };
		public static List<string> PSBVersions = new List<string>() { "11", "12", "13", "55", "60", "70" };
		//public decimal IllustratorVersions = {"12", "13", "14", "15.1"};

		/// <summary>
		/// Determines if PSTools is installed.
		/// </summary>
		/// <returns>
		///   <c>true</c> if it is installed; otherwise, <c>false</c>.
		/// </returns>
		public bool isInstalled()
		{
			__isInstalled = false;
			foreach (Versions __version in Enum.GetValues(typeof(Versions)))
			{
				__key = Registry.ClassesRoot.OpenSubKey("Photoshop.Image." + (int)__version + "\\\\shell\\\\PSTools");
				if (__key != null)
				{
					//MessageBox.Show(__key.Name);
					__isInstalled = true;
					__versionsInstalled.Add((int)__version);
				}

				/*try
				{
					MessageBox.Show("Photoshop.Image." + (int)__version + "\\\\shell\\\\Save as JPEG");

					
				}
				catch (Exception __e)
				{
					MessageBox.Show((int)__version + ">>> bad" + __e.Message); 
					//__isInstalled = false;
				}*/
			}
			return __isInstalled;
		}

		/// <summary>
		/// Checks the specified version.
		/// </summary>
		/// <param name="__version">Version.</param>
		/// <returns></returns>
		public bool check(Versions __version)
		{
            //MessageBox.Show(">" + ((int)__version).ToString());
			__key = Registry.ClassesRoot.OpenSubKey("Photoshop.Image." + (int)__version);
            //MessageBox.Show(__key.Name);
			if (__key != null)
			{
				//MessageBox.Show(__key.Name);
				return true;
			}
			else
			{
				return false;
			}
		}

		/// <summary>
		/// Gets the installed versions.
		/// </summary>
		/// <value>
		/// The installed versions.
		/// </value>
		public List<int> installedVersions
		{
			get { return __versionsInstalled;}
		}
	}
}
