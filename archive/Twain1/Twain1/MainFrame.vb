Imports System
Imports System.Drawing
Imports System.Collections
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Data

Imports Twain1.TwainLib

Namespace TwainGui

	Public Class MainFrame
		Inherits System.Windows.Forms.Form
		Implements IMessageFilter

		Private msgfilter As Boolean
		Private tw As Twain
		Private picnumber As Integer = 0

		<STAThread()> Shared Sub Main()
			If (Twain.ScreenBitDepth < 15) Then
				MessageBox.Show("Need high/true-color video mode!", "Screen Bit Depth", MessageBoxButtons.OK, MessageBoxIcon.Information)
				Return
			End If

			Dim mf As MainFrame = New MainFrame
			Application.Run(mf)
		End Sub

		Public Sub New()
			InitializeComponent()
			tw = New Twain
			tw.Init(Me.Handle)
		End Sub

		Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
			If disposing Then
				If Not (components Is Nothing) Then
					components.Dispose()
				End If
			End If
			MyBase.Dispose(disposing)
		End Sub

#Region " Windows Form Designer generated code "

		Private components As System.ComponentModel.Container = Nothing
		Private mdiClient1 As System.Windows.Forms.MdiClient
		Private menuMainFile As System.Windows.Forms.MenuItem
		Private menuItemScan As System.Windows.Forms.MenuItem
		Private menuItemSelSrc As System.Windows.Forms.MenuItem
		Private menuMainWindow As System.Windows.Forms.MenuItem
		Private menuItemExit As System.Windows.Forms.MenuItem
		Private menuItemSepr As System.Windows.Forms.MenuItem
		Private mainFrameMenu As System.Windows.Forms.MainMenu
		Private Sub InitializeComponent()
			Me.menuMainFile = New System.Windows.Forms.MenuItem
			Me.menuItemSelSrc = New System.Windows.Forms.MenuItem
			Me.menuItemScan = New System.Windows.Forms.MenuItem
			Me.menuItemSepr = New System.Windows.Forms.MenuItem
			Me.menuItemExit = New System.Windows.Forms.MenuItem
			Me.mainFrameMenu = New System.Windows.Forms.MainMenu
			Me.menuMainWindow = New System.Windows.Forms.MenuItem
			Me.mdiClient1 = New System.Windows.Forms.MdiClient
			Me.SuspendLayout()

			Me.menuMainFile.Index = 0
			Me.menuMainFile.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.menuItemSelSrc, Me.menuItemScan, Me.menuItemSepr, Me.menuItemExit})
			Me.menuMainFile.MergeType = System.Windows.Forms.MenuMerge.MergeItems
			Me.menuMainFile.Text = "&File"

			Me.menuItemSelSrc.Index = 0
			Me.menuItemSelSrc.MergeOrder = 11
			Me.menuItemSelSrc.Text = "&Select Source..."
			'Me.menuItemSelSrc.Click += New System.EventHandler(this.menuItemSelSrc_Click)
			AddHandler menuItemSelSrc.Click, AddressOf Me.menuItemSelSrc_Click

			Me.menuItemScan.Index = 1
			Me.menuItemScan.MergeOrder = 12
			Me.menuItemScan.Text = "&Acquire..."
			'Me.menuItemScan.Click += New System.EventHandler(this.menuItemScan_Click)
			AddHandler menuItemScan.Click, AddressOf Me.menuItemScan_Click

			Me.menuItemSepr.Index = 2
			Me.menuItemSepr.MergeOrder = 19
			Me.menuItemSepr.Text = "-"

			Me.menuItemExit.Index = 3
			Me.menuItemExit.MergeOrder = 21
			Me.menuItemExit.Text = "&Exit"
			'Me.menuItemExit.Click += New System.EventHandler(this.menuItemExit_Click)
			AddHandler menuItemExit.Click, AddressOf menuItemExit_Click

			Me.mainFrameMenu.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.menuMainFile, Me.menuMainWindow})

			Me.menuMainWindow.Index = 1
			Me.menuMainWindow.MdiList = True
			Me.menuMainWindow.Text = "&Window"

			Me.mdiClient1.Dock = System.Windows.Forms.DockStyle.Fill
			Me.mdiClient1.Name = "mdiClient1"
			Me.mdiClient1.TabIndex = 0

			Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
			Me.ClientSize = New System.Drawing.Size(600, 345)
			Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.mdiClient1})
			Me.IsMdiContainer = True
			Me.Menu = Me.mainFrameMenu
			Me.Name = "MainFrame"
			Me.Text = "My Twain"
			Me.ResumeLayout(False)
		End Sub

#End Region

		Private Sub menuItemExit_Click(ByVal sender As Object, ByVal e As System.EventArgs)
			Close()
		End Sub

		Private Sub menuItemScan_Click(ByVal sender As Object, ByVal e As System.EventArgs)
			If (Not msgfilter) Then
				Me.Enabled = False
				msgfilter = True
				Application.AddMessageFilter(Me)
			End If
			tw.Acquire()
		End Sub

		Private Sub menuItemSelSrc_Click(ByVal sender As Object, ByVal e As System.EventArgs)
			tw.Select()
		End Sub


		Public Function PreFilterMessage(ByRef m As Message) As Boolean Implements IMessageFilter.PreFilterMessage
			Dim cmd As TwainCommand = tw.PassMessage(m)
			If (cmd = TwainCommand.Not) Then
				Return False
			End If

			Select Case cmd
				Case TwainCommand.CloseRequest
					EndingScan()
					tw.CloseSrc()
				Case TwainCommand.CloseOk
					EndingScan()
					tw.CloseSrc()
				Case TwainCommand.DeviceEvent
				Case TwainCommand.TransferReady
					Dim pics As ArrayList = tw.TransferPictures()
					EndingScan()
					tw.CloseSrc()
					picnumber += 1
					Dim i As Integer
					For i = 0 To pics.Count - 1 Step 1
						Dim img As IntPtr = CType(pics(i), IntPtr)
						Dim newpic As PicForm = New PicForm(img)
						newpic.MdiParent = Me
						Dim picnum As Integer = i + 1
						'Hier Name des Scans angeben -> Nummer, Datum, Bezeichnung, wie auch immer...
						newpic.Text = "TestVGNR" + picnumber.ToString() + "_Pic" + picnum.ToString()
						newpic.Show()
					Next
			End Select

			Return True
		End Function

		Private Sub EndingScan()
			If (msgfilter) Then
				Application.RemoveMessageFilter(Me)
				msgfilter = False
				Me.Enabled = True
				Me.Activate()
			End If
		End Sub
	End Class

End Namespace