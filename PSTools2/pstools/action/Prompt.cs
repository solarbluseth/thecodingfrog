using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Text;
using System.Drawing;

namespace PSTools
{
	public static class Prompt
	{
		public static string ShowDialog(string __text, string __caption)
		{
			System.Windows.Forms.Form __prompt = new System.Windows.Forms.Form();
			__prompt.Width = 440;
			__prompt.Height = 120;
			__prompt.Text = __caption;
			__prompt.FormBorderStyle = FormBorderStyle.FixedSingle;
			//__prompt.BackColor = Color.Transparent;

			Label __label = new Label() { Left = 20, Top = 10, Width = 350, Text = __text };

			TextBox __path = new TextBox() { Left = 20, Top = 30, Width = 400 };

			Button __confirmation = new Button() { Text = "Ok", Left = 20, Width = 100, Top = 60 };
			Button __cancel = new Button() { Text = "Cancel", Left = 130, Width = 100, Top = 60 };
			__confirmation.Click += (sender, e) => { __prompt.Close(); };
			__cancel.Click += (sender, e) => { __prompt.Close(); };
			__prompt.Controls.Add(__confirmation);
			__prompt.Controls.Add(__cancel);
			__prompt.Controls.Add(__path);
			__prompt.Controls.Add(__label);
			__prompt.ShowDialog();
			return __path.Text;
		}
	}
}
