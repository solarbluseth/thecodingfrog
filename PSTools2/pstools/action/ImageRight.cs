using System.Text.RegularExpressions;
using System.Diagnostics;
using System;
using System.Windows.Forms;
using System.Collections;
using System.Data;
using System.Collections.Generic;
using System.IO;

namespace PSTools
{
	public class ImageRight
	{
		private string __bank;
		private string __bankcode;
		private string __imagecode;
		private string __url;
		private bool __isValidCode = false;
		
		
		public void parse(string __name)
		{
			MatchCollection __mc;
			GroupCollection __gc;
			CaptureCollection __cc;
			
			__mc = Regex.Matches(__name, "(\\w{2})\\-\\w{2}\\-(DA|€€)\\-(.*)", RegexOptions.Multiline & RegexOptions.IgnoreCase);
			
			if (__mc.Count > 0)
			{
				//MessageBox.Show(__name)
				__gc = __mc[0].Groups;
				
				__cc = __gc[1].Captures;
				__bankcode = (string) (__cc[0].Value);
				
				__cc = __gc[3].Captures;
				__imagecode = (string) (__cc[0].Value);
				
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
		
		private void setURL()
		{
			switch (__bankcode)
			{
				case "GI":
					__url = "www.gettyimages.fr/detail/{0}";
					__bank = "Getty Images";
					break;
				case "IS":
					__url = "www.istockphoto.com/file_thumbview_approve/{0}";
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
		
		public void createLink(string __path)
		{
			StreamWriter __sw;
			
			if (! Directory.Exists(__path + "+ Rights\\"))
			{
				Directory.CreateDirectory(__path + "+ Rights\\");
			}
			__sw = File.CreateText(__path + ("+ Rights\\" + __bank + " " + __imagecode + ".url"));
			__sw.WriteLine("[InternetShortcut]");
			__sw.WriteLine("URL=http://" + string.Format(__url, __imagecode));
			__sw.Close();
		}
		
		
		public string Code
		{
			get
			{
				return __imagecode;
			}
		}
		
		public string ImageBank
		{
			get
			{
				return __bank;
			}
		}
		
		public string URL
		{
			get
			{
				return __url;
			}
		}
		
		public bool isValidURL
		{
			get
			{
				return (__url != null && __imagecode != null);
			}
		}
		
		public bool isValidCode
		{
			get
			{
				return __isValidCode;
			}
		}
		
	}
	
}
