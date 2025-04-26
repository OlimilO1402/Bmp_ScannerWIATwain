using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace TwainGui
{
	/// <summary>
	/// Summary description for InfoForm.
	/// </summary>
	internal class InfoForm : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.TextBox txtWidth;
		private System.Windows.Forms.TextBox txtHeight;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.TextBox txtBitsPix;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.TextBox txtSize;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Button btnOK;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public InfoForm( BITMAPINFOHEADER bmi )
		{
			InitializeComponent();
			txtWidth.Text = bmi.biWidth.ToString();
			txtHeight.Text = bmi.biHeight.ToString();
			txtBitsPix.Text = bmi.biBitCount.ToString();
			int skb = bmi.biSizeImage >> 10;
			txtSize.Text = skb.ToString();
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.btnOK = new System.Windows.Forms.Button();
			this.label4 = new System.Windows.Forms.Label();
			this.txtWidth = new System.Windows.Forms.TextBox();
			this.txtHeight = new System.Windows.Forms.TextBox();
			this.txtSize = new System.Windows.Forms.TextBox();
			this.txtBitsPix = new System.Windows.Forms.TextBox();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// btnOK
			// 
			this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
			this.btnOK.Location = new System.Drawing.Point(104, 136);
			this.btnOK.Name = "btnOK";
			this.btnOK.Size = new System.Drawing.Size(96, 24);
			this.btnOK.TabIndex = 2;
			this.btnOK.Text = "OK";
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(8, 107);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(88, 16);
			this.label4.TabIndex = 0;
			this.label4.Text = "Size KB";
			// 
			// txtWidth
			// 
			this.txtWidth.Location = new System.Drawing.Point(104, 8);
			this.txtWidth.Name = "txtWidth";
			this.txtWidth.ReadOnly = true;
			this.txtWidth.Size = new System.Drawing.Size(160, 20);
			this.txtWidth.TabIndex = 1;
			this.txtWidth.Text = "";
			// 
			// txtHeight
			// 
			this.txtHeight.Location = new System.Drawing.Point(104, 40);
			this.txtHeight.Name = "txtHeight";
			this.txtHeight.ReadOnly = true;
			this.txtHeight.Size = new System.Drawing.Size(160, 20);
			this.txtHeight.TabIndex = 1;
			this.txtHeight.Text = "";
			// 
			// txtSize
			// 
			this.txtSize.Location = new System.Drawing.Point(104, 104);
			this.txtSize.Name = "txtSize";
			this.txtSize.ReadOnly = true;
			this.txtSize.Size = new System.Drawing.Size(160, 20);
			this.txtSize.TabIndex = 1;
			this.txtSize.Text = "";
			// 
			// txtBitsPix
			// 
			this.txtBitsPix.Location = new System.Drawing.Point(104, 72);
			this.txtBitsPix.Name = "txtBitsPix";
			this.txtBitsPix.ReadOnly = true;
			this.txtBitsPix.Size = new System.Drawing.Size(160, 20);
			this.txtBitsPix.TabIndex = 1;
			this.txtBitsPix.Text = "";
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(8, 13);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(88, 16);
			this.label1.TabIndex = 0;
			this.label1.Text = "Width";
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(8, 45);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(88, 16);
			this.label2.TabIndex = 0;
			this.label2.Text = "Height";
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(8, 76);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(88, 16);
			this.label3.TabIndex = 0;
			this.label3.Text = "Bits/Pixel";
			// 
			// InfoForm
			// 
			this.AcceptButton = this.btnOK;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(274, 175);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.btnOK,
																		  this.txtWidth,
																		  this.label1,
																		  this.txtHeight,
																		  this.label2,
																		  this.txtBitsPix,
																		  this.label3,
																		  this.txtSize,
																		  this.label4});
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "InfoForm";
			this.ShowInTaskbar = false;
			this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
			this.Text = "InfoForm";
			this.ResumeLayout(false);

		}
		#endregion


	}
}
