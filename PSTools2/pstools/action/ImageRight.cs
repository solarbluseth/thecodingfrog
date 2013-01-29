using System.Text.RegularExpressions;
using System.IO;

namespace PSTools
{
	public class ImageRight
	{
		private string __bank;
		private string __bankcode;
		private string __imagecode;
		private string __type;
		private string __aquired;
		private string __url;
		private bool __isValidCode = false;


		/// <summary>
		/// Parses the specified Name.
		/// </summary>
		/// <param name="__name">Name</param>
		public void parse(string __name)
		{
			MatchCollection __mc;
			GroupCollection __gc;
			CaptureCollection __cc;
			
			__mc = Regex.Matches(__name, "(\\w{2})\\-(\\w{2})\\-(DA|€€)\\-(.*)", RegexOptions.Multiline & RegexOptions.IgnoreCase);
			
			if (__mc.Count > 0)
			{
				//MessageBox.Show(__name)
				__gc = __mc[0].Groups;
				
				__cc = __gc[1].Captures;
				__bankcode = __cc[0].Value;

				__cc = __gc[2].Captures;
				__type = __cc[0].Value;

				__cc = __gc[3].Captures;
				__aquired = __cc[0].Value;
				
				__cc = __gc[4].Captures;
				__imagecode = __cc[0].Value;
				
				__isValidCode = true;
				
				setURL();
			}
			else
			{
				__isValidCode = false;
				__bank = null;
				__imagecode = null;
			}
		}

		/// <summary>
		/// Sets the URL.
		/// </summary>
		private void setURL()
		{
			switch (__bankcode)
			{
				case "GI":
					__url = "www.gettyimages.fr/detail/{0}";
					__bank = "Getty Images";
					break;
				case "IS":
					__url = "www.istockphoto.com/file_thumbview_approve/{0}/2/";
					__bank = "iStockPhoto";
					break;
				case "CO":
					__url = "www.corbisimages.com/stock-photo/rights-managed/{0}/e/?tab=details&caller=search";
					__bank = "Corbis Images";
					break;
				case "GO":
					__url = "www.graphicobsession.com/recherche.php?LibRecherche={0}";
					__bank = "GraphicObsession";
					break;
				case "PA":
					__url = "www.photoalto.fr/recherche?idsearch=4vxi&bs_string={0}&searchWithin=0";
					__bank = "PhotoAlto";
					break;
				case "SX":
					__url = "http://www.sxc.hu/photo/{0}";
					__bank = "stock.xchng";
					break;
				case "FK":
					__url = null;
					__bank = null;
					break;
				default:
					__bank = null;
					__imagecode = null;
					__url = null;
					break;
			}
		}

		/// <summary>
		/// Creates the link.
		/// </summary>
		/// <param name="__path">Path</param>
		public void createLink(string __path)
		{
			StreamWriter __sw;
			
			if (! Directory.Exists(__path + "+ Rights\\"))
			{
				Directory.CreateDirectory(__path + "+ Rights\\");
			}
			File.Delete(__path + "+ Rights\\" + __bank + "-" + __type + "-" + __aquired + "-" + __imagecode + ".url");
			__sw = File.CreateText(__path + "+ Rights\\" + __bank + "-" + __type + "-" + __aquired + "-" + __imagecode + ".url");
			__sw.WriteLine("[InternetShortcut]");
			__sw.WriteLine("URL=http://" + string.Format(__url, __imagecode));
			__sw.Close();
		}


		/// <summary>
		/// Gets the code.
		/// </summary>
		/// <value>
		/// The code.
		/// </value>
		public string Code
		{
			get
			{
				return __imagecode;
			}
		}

		/// <summary>
		/// Gets the image bank.
		/// </summary>
		/// <value>
		/// The image bank.
		/// </value>
		public string ImageBank
		{
			get
			{
				return __bank;
			}
		}

		/// <summary>
		/// Gets the URL.
		/// </summary>
		/// <value>
		/// The URL.
		/// </value>
		public string URL
		{
			get
			{
				return __url;
			}
		}

		/// <summary>
		/// Gets a value indicating whether this instance is valid URL.
		/// </summary>
		/// <value>
		/// <c>true</c> if this instance is valid URL; otherwise, <c>false</c>.
		/// </value>
		public bool isValidURL
		{
			get
			{
				return (__url != null && __imagecode != null);
			}
		}

		/// <summary>
		/// Gets a value indicating whether this instance is valid code.
		/// </summary>
		/// <value>
		/// <c>true</c> if this instance is valid code; otherwise, <c>false</c>.
		/// </value>
		public bool isValidCode
		{
			get
			{
				return __isValidCode;
			}
		}
		
	}
	
}
