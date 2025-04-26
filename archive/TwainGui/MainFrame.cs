using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using TwainLib;

namespace TwainGui
{
public class MainFrame : System.Windows.Forms.Form, IMessageFilter
	{
	private System.Windows.Forms.MdiClient mdiClient1;
	private System.Windows.Forms.MenuItem menuMainFile;
	private System.Windows.Forms.MenuItem menuItemScan;
	private System.Windows.Forms.MenuItem menuItemSelSrc;
	private System.Windows.Forms.MenuItem menuMainWindow;
	private System.Windows.Forms.MenuItem menuItemExit;
	private System.Windows.Forms.MenuItem menuItemSepr;
	private System.Windows.Forms.MainMenu mainFrameMenu;

	private System.ComponentModel.Container components = null;

	public MainFrame()
		{
		InitializeComponent();
		tw = new Twain();
		tw.Init( this.Handle );
		}

	protected override void Dispose( bool disposing )
		{
		if( disposing )
			{
			tw.Finish();
			if (components != null) 
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
			this.menuMainFile = new System.Windows.Forms.MenuItem();
			this.menuItemSelSrc = new System.Windows.Forms.MenuItem();
			this.menuItemScan = new System.Windows.Forms.MenuItem();
			this.menuItemSepr = new System.Windows.Forms.MenuItem();
			this.menuItemExit = new System.Windows.Forms.MenuItem();
			this.mainFrameMenu = new System.Windows.Forms.MainMenu();
			this.menuMainWindow = new System.Windows.Forms.MenuItem();
			this.mdiClient1 = new System.Windows.Forms.MdiClient();
			this.SuspendLayout();
			// 
			// menuMainFile
			// 
			this.menuMainFile.Index = 0;
			this.menuMainFile.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																						 this.menuItemSelSrc,
																						 this.menuItemScan,
																						 this.menuItemSepr,
																						 this.menuItemExit});
			this.menuMainFile.MergeType = System.Windows.Forms.MenuMerge.MergeItems;
			this.menuMainFile.Text = "&File";
			// 
			// menuItemSelSrc
			// 
			this.menuItemSelSrc.Index = 0;
			this.menuItemSelSrc.MergeOrder = 11;
			this.menuItemSelSrc.Text = "&Select Source...";
			this.menuItemSelSrc.Click += new System.EventHandler(this.menuItemSelSrc_Click);
			// 
			// menuItemScan
			// 
			this.menuItemScan.Index = 1;
			this.menuItemScan.MergeOrder = 12;
			this.menuItemScan.Text = "&Acquire...";
			this.menuItemScan.Click += new System.EventHandler(this.menuItemScan_Click);
			// 
			// menuItemSepr
			// 
			this.menuItemSepr.Index = 2;
			this.menuItemSepr.MergeOrder = 19;
			this.menuItemSepr.Text = "-";
			// 
			// menuItemExit
			// 
			this.menuItemExit.Index = 3;
			this.menuItemExit.MergeOrder = 21;
			this.menuItemExit.Text = "&Exit";
			this.menuItemExit.Click += new System.EventHandler(this.menuItemExit_Click);
			// 
			// mainFrameMenu
			// 
			this.mainFrameMenu.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																						  this.menuMainFile,
																						  this.menuMainWindow});
			// 
			// menuMainWindow
			// 
			this.menuMainWindow.Index = 1;
			this.menuMainWindow.MdiList = true;
			this.menuMainWindow.Text = "&Window";
			// 
			// mdiClient1
			// 
			this.mdiClient1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.mdiClient1.Name = "mdiClient1";
			this.mdiClient1.TabIndex = 0;
			// 
			// MainFrame
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(600, 345);
			this.Controls.AddRange(new System.Windows.Forms.Control[] {
																		  this.mdiClient1});
			this.IsMdiContainer = true;
			this.Menu = this.mainFrameMenu;
			this.Name = "MainFrame";
			this.Text = "TWAIN GUI";
			this.ResumeLayout(false);

		}
		#endregion



	private void menuItemExit_Click(object sender, System.EventArgs e)
		{
		Close();
		}

	private void menuItemScan_Click(object sender, System.EventArgs e)
		{
		if( ! msgfilter )
			{
			this.Enabled = false;
			msgfilter = true;
			Application.AddMessageFilter( this );
			}
		tw.Acquire();
		}

	private void menuItemSelSrc_Click(object sender, System.EventArgs e)
		{
		tw.Select();
		}


	bool IMessageFilter.PreFilterMessage( ref Message m )
		{
		TwainCommand cmd = tw.PassMessage( ref m );
		if( cmd == TwainCommand.Not )
			return false;

		switch( cmd )
			{
			case TwainCommand.CloseRequest:
				{
				EndingScan();
				tw.CloseSrc();
				break;
				}
			case TwainCommand.CloseOk:
				{
				EndingScan();
				tw.CloseSrc();
				break;
				}
			case TwainCommand.DeviceEvent:
				{
				break;
				}
			case TwainCommand.TransferReady:
				{
				ArrayList pics = tw.TransferPictures();
				EndingScan();
				tw.CloseSrc();
				picnumber++;
				for( int i = 0; i < pics.Count; i++ )
					{
					IntPtr img = (IntPtr) pics[ i ];
					PicForm newpic = new PicForm( img );
					newpic.MdiParent = this;
					int picnum = i + 1;
					newpic.Text = "ScanPass" + picnumber.ToString() + "_Pic" + picnum.ToString();
					newpic.Show();
					}
				break;
				}
			}

		return true;
		}

	private void EndingScan()
		{
		if( msgfilter )
			{
			Application.RemoveMessageFilter( this );
			msgfilter = false;
			this.Enabled = true;
			this.Activate();
			}
		}

	private bool	msgfilter;
	private Twain	tw;
	private int		picnumber = 0;








	[STAThread]
	static void Main() 
		{
		if( Twain.ScreenBitDepth < 15 )
			{
			MessageBox.Show( "Need high/true-color video mode!", "Screen Bit Depth", MessageBoxButtons.OK, MessageBoxIcon.Information );
			return;
			}

		MainFrame mf = new MainFrame();
		Application.Run( mf );
		}

	} // class MainFrame

} // namespace TwainGui
