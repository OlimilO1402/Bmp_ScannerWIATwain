Imports System
Imports System.Drawing
Imports System.Collections
Imports System.ComponentModel
Imports System.Windows.Forms

Namespace TwainGui
	Friend Class InfoForm
		Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "
		Public Sub New(ByVal bmi As BITMAPINFOHEADER)
			InitializeComponent()
			txtWidth.Text = bmi.biWidth.ToString()
			txtHeight.Text = bmi.biHeight.ToString()
			txtBitsPix.Text = bmi.biBitCount.ToString()
			Dim skb As Integer = Int(bmi.biSizeImage / 2 ^ 10)
			txtSize.Text = skb.ToString()
		End Sub

		Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
			If disposing Then
				If Not (components Is Nothing) Then
					components.Dispose()
				End If
			End If
			MyBase.Dispose(disposing)
		End Sub

		Private components As System.ComponentModel.Container = Nothing
		Private label1 As System.Windows.Forms.Label
		Private txtWidth As System.Windows.Forms.TextBox
		Private txtHeight As System.Windows.Forms.TextBox
		Private label2 As System.Windows.Forms.Label
		Private txtBitsPix As System.Windows.Forms.TextBox
		Private label3 As System.Windows.Forms.Label
		Private txtSize As System.Windows.Forms.TextBox
		Private label4 As System.Windows.Forms.Label
		Private btnOK As System.Windows.Forms.Button

		Private Sub InitializeComponent()
			Me.btnOK = New System.Windows.Forms.Button
			Me.label4 = New System.Windows.Forms.Label
			Me.txtWidth = New System.Windows.Forms.TextBox
			Me.txtHeight = New System.Windows.Forms.TextBox
			Me.txtSize = New System.Windows.Forms.TextBox
			Me.txtBitsPix = New System.Windows.Forms.TextBox
			Me.label1 = New System.Windows.Forms.Label
			Me.label2 = New System.Windows.Forms.Label
			Me.label3 = New System.Windows.Forms.Label
			Me.SuspendLayout()

			Me.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK
			Me.btnOK.Location = New System.Drawing.Point(104, 136)
			Me.btnOK.Name = "btnOK"
			Me.btnOK.Size = New System.Drawing.Size(96, 24)
			Me.btnOK.TabIndex = 2
			Me.btnOK.Text = "OK"

			Me.label4.Location = New System.Drawing.Point(8, 107)
			Me.label4.Name = "label4"
			Me.label4.Size = New System.Drawing.Size(88, 16)
			Me.label4.TabIndex = 0
			Me.label4.Text = "Size KB"

			Me.txtWidth.Location = New System.Drawing.Point(104, 8)
			Me.txtWidth.Name = "txtWidth"
			Me.txtWidth.ReadOnly = True
			Me.txtWidth.Size = New System.Drawing.Size(160, 20)
			Me.txtWidth.TabIndex = 1
			Me.txtWidth.Text = ""

			Me.txtHeight.Location = New System.Drawing.Point(104, 40)
			Me.txtHeight.Name = "txtHeight"
			Me.txtHeight.ReadOnly = True
			Me.txtHeight.Size = New System.Drawing.Size(160, 20)
			Me.txtHeight.TabIndex = 1
			Me.txtHeight.Text = ""

			Me.txtSize.Location = New System.Drawing.Point(104, 104)
			Me.txtSize.Name = "txtSize"
			Me.txtSize.ReadOnly = True
			Me.txtSize.Size = New System.Drawing.Size(160, 20)
			Me.txtSize.TabIndex = 1
			Me.txtSize.Text = ""

			Me.txtBitsPix.Location = New System.Drawing.Point(104, 72)
			Me.txtBitsPix.Name = "txtBitsPix"
			Me.txtBitsPix.ReadOnly = True
			Me.txtBitsPix.Size = New System.Drawing.Size(160, 20)
			Me.txtBitsPix.TabIndex = 1
			Me.txtBitsPix.Text = ""

			Me.label1.Location = New System.Drawing.Point(8, 13)
			Me.label1.Name = "label1"
			Me.label1.Size = New System.Drawing.Size(88, 16)
			Me.label1.TabIndex = 0
			Me.label1.Text = "Width"

			Me.label2.Location = New System.Drawing.Point(8, 45)
			Me.label2.Name = "label2"
			Me.label2.Size = New System.Drawing.Size(88, 16)
			Me.label2.TabIndex = 0
			Me.label2.Text = "Height"

			Me.label3.Location = New System.Drawing.Point(8, 76)
			Me.label3.Name = "label3"
			Me.label3.Size = New System.Drawing.Size(88, 16)
			Me.label3.TabIndex = 0
			Me.label3.Text = "Bits/Pixel"

			Me.AcceptButton = Me.btnOK
			Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
			Me.ClientSize = New System.Drawing.Size(274, 175)
			Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.btnOK, Me.txtWidth, Me.label1, Me.txtHeight, Me.label2, Me.txtBitsPix, Me.label3, Me.txtSize, Me.label4})
			Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
			Me.MaximizeBox = False
			Me.MinimizeBox = False
			Me.Name = "InfoForm"
			Me.ShowInTaskbar = False
			Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
			Me.Text = "InfoForm"
			Me.ResumeLayout(False)
		End Sub
#End Region

	End Class
End Namespace
