using System.Text.RegularExpressions;
using System.Diagnostics;
using System;
using System.Windows.Forms;
using System.Collections;
using Microsoft.VisualBasic;
using System.Data;
using System.Collections.Generic;
using System.IO;

namespace PSTools
{
	[global::Microsoft.VisualBasic.CompilerServices.DesignerGenerated()]
	public partial class Form : System.Windows.Forms.Form
	{
		
		//Form overrides dispose to clean up the component list.
		[System.Diagnostics.DebuggerNonUserCode()]
		protected override void Dispose(bool disposing)
		{
			try
			{
				if (disposing && components != null)
				{
					components.Dispose();
				}
			}
			finally
			{
				base.Dispose(disposing);
			}
		}

		//Required by the Windows Form Designer
		
		//NOTE: The following procedure is required by the Windows Form Designer
		//It can be modified using the Windows Form Designer.
		//Do not modify it using the code editor.
		[System.Diagnostics.DebuggerStepThrough()]
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form));
			this.Install = new System.Windows.Forms.Button();
			this.GroupBox1 = new System.Windows.Forms.GroupBox();
			this.Label1 = new System.Windows.Forms.Label();
			this.ArchiveDirectory = new System.Windows.Forms.TextBox();
			this.AutoArchive = new System.Windows.Forms.CheckBox();
			this.GroupBox2 = new System.Windows.Forms.GroupBox();
			this.CS6 = new System.Windows.Forms.Label();
			this.LabelCS55 = new System.Windows.Forms.Label();
			this.CS55 = new System.Windows.Forms.Label();
			this.LabelCS5 = new System.Windows.Forms.Label();
			this.LabelCS4 = new System.Windows.Forms.Label();
			this.LabelCS3 = new System.Windows.Forms.Label();
			this.CS5 = new System.Windows.Forms.Label();
			this.CS4 = new System.Windows.Forms.Label();
			this.CS3 = new System.Windows.Forms.Label();
			this.GroupBox3 = new System.Windows.Forms.GroupBox();
			this.ExportLayerComps = new System.Windows.Forms.CheckBox();
			this.NamedExportQuality = new System.Windows.Forms.NumericUpDown();
			this.Label2 = new System.Windows.Forms.Label();
			this.ToolTip1 = new System.Windows.Forms.ToolTip(this.components);
			this.Label3 = new System.Windows.Forms.Label();
			this.GroupBox4 = new System.Windows.Forms.GroupBox();
			this.ExcludeDirectories = new System.Windows.Forms.TextBox();
			this.LabelCompiled = new System.Windows.Forms.Label();
			this.LabelCS6 = new System.Windows.Forms.Label();
			this.GroupBox1.SuspendLayout();
			this.GroupBox2.SuspendLayout();
			this.GroupBox3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.NamedExportQuality)).BeginInit();
			this.GroupBox4.SuspendLayout();
			this.SuspendLayout();
			// 
			// Install
			// 
			this.Install.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
						| System.Windows.Forms.AnchorStyles.Right)));
			this.Install.Location = new System.Drawing.Point(6, 138);
			this.Install.Name = "Install";
			this.Install.Size = new System.Drawing.Size(247, 23);
			this.Install.TabIndex = 0;
			this.Install.Text = "Install";
			this.Install.UseVisualStyleBackColor = true;
			this.Install.Click += new System.EventHandler(this.Install_Click);
			// 
			// GroupBox1
			// 
			this.GroupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
						| System.Windows.Forms.AnchorStyles.Right)));
			this.GroupBox1.Controls.Add(this.Label1);
			this.GroupBox1.Controls.Add(this.ArchiveDirectory);
			this.GroupBox1.Controls.Add(this.AutoArchive);
			this.GroupBox1.Location = new System.Drawing.Point(12, 185);
			this.GroupBox1.Name = "GroupBox1";
			this.GroupBox1.Size = new System.Drawing.Size(260, 74);
			this.GroupBox1.TabIndex = 1;
			this.GroupBox1.TabStop = false;
			this.GroupBox1.Text = "Archive";
			// 
			// Label1
			// 
			this.Label1.AutoSize = true;
			this.Label1.Location = new System.Drawing.Point(7, 22);
			this.Label1.Name = "Label1";
			this.Label1.Size = new System.Drawing.Size(49, 13);
			this.Label1.TabIndex = 2;
			this.Label1.Text = "Directory";
			// 
			// ArchiveDirectory
			// 
			this.ArchiveDirectory.Location = new System.Drawing.Point(75, 19);
			this.ArchiveDirectory.Name = "ArchiveDirectory";
			this.ArchiveDirectory.Size = new System.Drawing.Size(178, 20);
			this.ArchiveDirectory.TabIndex = 1;
			this.ArchiveDirectory.TextChanged += new System.EventHandler(this.ArchiveDirectory_TextChanged);
			// 
			// AutoArchive
			// 
			this.AutoArchive.AutoSize = true;
			this.AutoArchive.Location = new System.Drawing.Point(75, 45);
			this.AutoArchive.Name = "AutoArchive";
			this.AutoArchive.Size = new System.Drawing.Size(166, 17);
			this.AutoArchive.TabIndex = 0;
			this.AutoArchive.Text = "Auto archive previous version";
			this.AutoArchive.UseVisualStyleBackColor = true;
			this.AutoArchive.CheckedChanged += new System.EventHandler(this.AutoArchive_CheckedChanged);
			// 
			// GroupBox2
			// 
			this.GroupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
						| System.Windows.Forms.AnchorStyles.Left)
						| System.Windows.Forms.AnchorStyles.Right)));
			this.GroupBox2.Controls.Add(this.LabelCS6);
			this.GroupBox2.Controls.Add(this.CS6);
			this.GroupBox2.Controls.Add(this.LabelCS55);
			this.GroupBox2.Controls.Add(this.CS55);
			this.GroupBox2.Controls.Add(this.LabelCS5);
			this.GroupBox2.Controls.Add(this.LabelCS4);
			this.GroupBox2.Controls.Add(this.LabelCS3);
			this.GroupBox2.Controls.Add(this.CS5);
			this.GroupBox2.Controls.Add(this.CS4);
			this.GroupBox2.Controls.Add(this.CS3);
			this.GroupBox2.Controls.Add(this.Install);
			this.GroupBox2.Location = new System.Drawing.Point(12, 12);
			this.GroupBox2.Name = "GroupBox2";
			this.GroupBox2.Size = new System.Drawing.Size(260, 167);
			this.GroupBox2.TabIndex = 2;
			this.GroupBox2.TabStop = false;
			this.GroupBox2.Text = "Install";
			// 
			// CS6
			// 
			this.CS6.AutoSize = true;
			this.CS6.Location = new System.Drawing.Point(8, 112);
			this.CS6.Name = "CS6";
			this.CS6.Size = new System.Drawing.Size(90, 13);
			this.CS6.TabIndex = 9;
			this.CS6.Text = "Photoshop CS6 : ";
			// 
			// LabelCS55
			// 
			this.LabelCS55.AutoSize = true;
			this.LabelCS55.Location = new System.Drawing.Point(102, 90);
			this.LabelCS55.Name = "LabelCS55";
			this.LabelCS55.Size = new System.Drawing.Size(54, 13);
			this.LabelCS55.TabIndex = 8;
			this.LabelCS55.Text = "Not found";
			// 
			// CS55
			// 
			this.CS55.AutoSize = true;
			this.CS55.Location = new System.Drawing.Point(8, 90);
			this.CS55.Name = "CS55";
			this.CS55.Size = new System.Drawing.Size(99, 13);
			this.CS55.TabIndex = 7;
			this.CS55.Text = "Photoshop CS5.5 : ";
			// 
			// LabelCS5
			// 
			this.LabelCS5.AutoSize = true;
			this.LabelCS5.Location = new System.Drawing.Point(102, 68);
			this.LabelCS5.Name = "LabelCS5";
			this.LabelCS5.Size = new System.Drawing.Size(54, 13);
			this.LabelCS5.TabIndex = 6;
			this.LabelCS5.Text = "Not found";
			// 
			// LabelCS4
			// 
			this.LabelCS4.AutoSize = true;
			this.LabelCS4.Location = new System.Drawing.Point(102, 44);
			this.LabelCS4.Name = "LabelCS4";
			this.LabelCS4.Size = new System.Drawing.Size(54, 13);
			this.LabelCS4.TabIndex = 5;
			this.LabelCS4.Text = "Not found";
			// 
			// LabelCS3
			// 
			this.LabelCS3.AutoSize = true;
			this.LabelCS3.Location = new System.Drawing.Point(102, 20);
			this.LabelCS3.Name = "LabelCS3";
			this.LabelCS3.Size = new System.Drawing.Size(54, 13);
			this.LabelCS3.TabIndex = 4;
			this.LabelCS3.Text = "Not found";
			// 
			// CS5
			// 
			this.CS5.AutoSize = true;
			this.CS5.Location = new System.Drawing.Point(7, 68);
			this.CS5.Name = "CS5";
			this.CS5.Size = new System.Drawing.Size(99, 13);
			this.CS5.TabIndex = 3;
			this.CS5.Text = "Photoshop CS5    : ";
			// 
			// CS4
			// 
			this.CS4.AutoSize = true;
			this.CS4.Location = new System.Drawing.Point(7, 44);
			this.CS4.Name = "CS4";
			this.CS4.Size = new System.Drawing.Size(99, 13);
			this.CS4.TabIndex = 2;
			this.CS4.Text = "Photoshop CS4    : ";
			// 
			// CS3
			// 
			this.CS3.AutoSize = true;
			this.CS3.Location = new System.Drawing.Point(7, 20);
			this.CS3.Name = "CS3";
			this.CS3.Size = new System.Drawing.Size(99, 13);
			this.CS3.TabIndex = 1;
			this.CS3.Text = "Photoshop CS3    : ";
			// 
			// GroupBox3
			// 
			this.GroupBox3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
						| System.Windows.Forms.AnchorStyles.Right)));
			this.GroupBox3.Controls.Add(this.ExportLayerComps);
			this.GroupBox3.Controls.Add(this.NamedExportQuality);
			this.GroupBox3.Controls.Add(this.Label2);
			this.GroupBox3.Location = new System.Drawing.Point(12, 327);
			this.GroupBox3.Name = "GroupBox3";
			this.GroupBox3.Size = new System.Drawing.Size(260, 54);
			this.GroupBox3.TabIndex = 3;
			this.GroupBox3.TabStop = false;
			this.GroupBox3.Text = "Export By Name";
			// 
			// ExportLayerComps
			// 
			this.ExportLayerComps.AutoSize = true;
			this.ExportLayerComps.Checked = true;
			this.ExportLayerComps.CheckState = System.Windows.Forms.CheckState.Checked;
			this.ExportLayerComps.Location = new System.Drawing.Point(149, 21);
			this.ExportLayerComps.Name = "ExportLayerComps";
			this.ExportLayerComps.Size = new System.Drawing.Size(84, 17);
			this.ExportLayerComps.TabIndex = 4;
			this.ExportLayerComps.Text = "LayerComps";
			this.ExportLayerComps.UseVisualStyleBackColor = true;
			this.ExportLayerComps.CheckedChanged += new System.EventHandler(this.ExportLayerComps_CheckedChanged);
			this.ExportLayerComps.MouseHover += new System.EventHandler(this.ExportLayerComps_MouseHover);
			// 
			// NamedExportQuality
			// 
			this.NamedExportQuality.Location = new System.Drawing.Point(75, 19);
			this.NamedExportQuality.Maximum = new decimal(new int[] {
            12,
            0,
            0,
            0});
			this.NamedExportQuality.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
			this.NamedExportQuality.Name = "NamedExportQuality";
			this.NamedExportQuality.Size = new System.Drawing.Size(58, 20);
			this.NamedExportQuality.TabIndex = 3;
			this.NamedExportQuality.Value = new decimal(new int[] {
            6,
            0,
            0,
            0});
			this.NamedExportQuality.ValueChanged += new System.EventHandler(this.NamedExportQuality_ValueChanged);
			// 
			// Label2
			// 
			this.Label2.AutoSize = true;
			this.Label2.Location = new System.Drawing.Point(8, 21);
			this.Label2.Name = "Label2";
			this.Label2.Size = new System.Drawing.Size(39, 13);
			this.Label2.TabIndex = 1;
			this.Label2.Text = "Quality";
			// 
			// ToolTip1
			// 
			this.ToolTip1.Popup += new System.Windows.Forms.PopupEventHandler(this.ToolTip1_Popup);
			// 
			// Label3
			// 
			this.Label3.AutoSize = true;
			this.Label3.Location = new System.Drawing.Point(7, 24);
			this.Label3.Name = "Label3";
			this.Label3.Size = new System.Drawing.Size(57, 13);
			this.Label3.TabIndex = 6;
			this.Label3.Text = "Directories";
			this.ToolTip1.SetToolTip(this.Label3, "List of directories seperated by comma");
			// 
			// GroupBox4
			// 
			this.GroupBox4.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)
						| System.Windows.Forms.AnchorStyles.Right)));
			this.GroupBox4.Controls.Add(this.Label3);
			this.GroupBox4.Controls.Add(this.ExcludeDirectories);
			this.GroupBox4.Location = new System.Drawing.Point(12, 266);
			this.GroupBox4.Name = "GroupBox4";
			this.GroupBox4.Size = new System.Drawing.Size(260, 55);
			this.GroupBox4.TabIndex = 5;
			this.GroupBox4.TabStop = false;
			this.GroupBox4.Text = "Excluded from auto archiving";
			// 
			// ExcludeDirectories
			// 
			this.ExcludeDirectories.Location = new System.Drawing.Point(75, 21);
			this.ExcludeDirectories.Name = "ExcludeDirectories";
			this.ExcludeDirectories.Size = new System.Drawing.Size(178, 20);
			this.ExcludeDirectories.TabIndex = 5;
			this.ExcludeDirectories.TextChanged += new System.EventHandler(this.ExcludeDirectories_TextChanged);
			// 
			// LabelCompiled
			// 
			this.LabelCompiled.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.LabelCompiled.AutoSize = true;
			this.LabelCompiled.Enabled = false;
			this.LabelCompiled.Location = new System.Drawing.Point(13, 393);
			this.LabelCompiled.Name = "LabelCompiled";
			this.LabelCompiled.Size = new System.Drawing.Size(54, 13);
			this.LabelCompiled.TabIndex = 6;
			this.LabelCompiled.Text = "Version v.";
			// 
			// LabelCS6
			// 
			this.LabelCS6.AutoSize = true;
			this.LabelCS6.Location = new System.Drawing.Point(102, 112);
			this.LabelCS6.Name = "LabelCS6";
			this.LabelCS6.Size = new System.Drawing.Size(54, 13);
			this.LabelCS6.TabIndex = 10;
			this.LabelCS6.Text = "Not found";
			// 
			// Form
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
			this.ClientSize = new System.Drawing.Size(284, 418);
			this.Controls.Add(this.LabelCompiled);
			this.Controls.Add(this.GroupBox3);
			this.Controls.Add(this.GroupBox2);
			this.Controls.Add(this.GroupBox1);
			this.Controls.Add(this.GroupBox4);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Name = "Form";
			this.Text = "SaveAsJPEG";
			this.GroupBox1.ResumeLayout(false);
			this.GroupBox1.PerformLayout();
			this.GroupBox2.ResumeLayout(false);
			this.GroupBox2.PerformLayout();
			this.GroupBox3.ResumeLayout(false);
			this.GroupBox3.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.NamedExportQuality)).EndInit();
			this.GroupBox4.ResumeLayout(false);
			this.GroupBox4.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();

		}
		internal System.Windows.Forms.Button Install;
		internal System.Windows.Forms.GroupBox GroupBox1;
		internal System.Windows.Forms.CheckBox AutoArchive;
		internal System.Windows.Forms.GroupBox GroupBox2;
		internal System.Windows.Forms.Label CS3;
		internal System.Windows.Forms.Label CS5;
		internal System.Windows.Forms.Label CS4;
		internal System.Windows.Forms.Label LabelCS5;
		internal System.Windows.Forms.Label LabelCS4;
		internal System.Windows.Forms.Label LabelCS3;
		internal System.Windows.Forms.Label Label1;
		internal System.Windows.Forms.TextBox ArchiveDirectory;
		internal System.Windows.Forms.GroupBox GroupBox3;
		internal System.Windows.Forms.Label Label2;
		internal System.Windows.Forms.NumericUpDown NamedExportQuality;

		internal System.Windows.Forms.CheckBox ExportLayerComps;
		internal System.Windows.Forms.ToolTip ToolTip1;
		internal System.Windows.Forms.GroupBox GroupBox4;
		internal System.Windows.Forms.Label Label3;
		internal System.Windows.Forms.TextBox ExcludeDirectories;
		internal System.Windows.Forms.Label LabelCS55;
		internal System.Windows.Forms.Label CS55;
		internal System.Windows.Forms.Label LabelCompiled;
		private System.ComponentModel.IContainer components;
		internal Label CS6;
		internal Label LabelCS6;
	}
	
}
