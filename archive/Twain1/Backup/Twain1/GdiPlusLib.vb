Imports System
Imports System.IO
Imports System.Collections
Imports System.Runtime.InteropServices
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Windows.Forms

Namespace GdiPlusLib

	Public Class Gdip
		<DllImport("gdiplus.dll", ExactSpelling:=True)> Friend Shared Function GdipCreateBitmapFromGdiDib(ByVal bminfo As IntPtr, ByVal pixdat As IntPtr, ByRef image As IntPtr) As Integer
		End Function
		<DllImport("gdiplus.dll", ExactSpelling:=True, CharSet:=CharSet.Unicode)> Friend Shared Function GdipSaveImageToFile(ByVal image As IntPtr, ByVal filename As String, <[In]()> ByRef clsid As Guid, ByVal encparams As IntPtr) As Integer
		End Function
		<DllImport("gdiplus.dll", ExactSpelling:=True)> Friend Shared Function GdipDisposeImage(ByVal image As IntPtr) As Integer
		End Function

		'private static ImageCodecInfo[] codecs = ImageCodecInfo.GetImageEncoders();
		Private Shared codecs() As ImageCodecInfo = ImageCodecInfo.GetImageEncoders()

		'private static bool GetCodecClsid( string filename, out Guid clsid )
		Private Shared Function GetCodecClsid(ByVal filename As String, ByRef clsid As Guid) As Boolean
			clsid = Guid.Empty
			Dim ext As String = Path.GetExtension(filename)
			'Checking string for null
			If IsNothing(ext) Then
				Return False
			End If
			ext = "*" + ext.ToUpper()
			Dim codec As ImageCodecInfo
			For Each codec In codecs
				If (codec.FilenameExtension.IndexOf(ext) >= 0) Then
					clsid = codec.Clsid
					Return True
				End If
			Next
			Return False
		End Function

		'public static bool SaveDIBAs( string picname, IntPtr bminfo, IntPtr pixdat )
		Public Shared Function SaveDIBAs(ByVal picname As String, ByVal bminfo As IntPtr, ByVal pixdat As IntPtr) As Boolean

			Dim fbd As New MyFolderBrowserDialog
			Dim clsid As Guid
			Dim img As IntPtr = IntPtr.Zero
			Dim st As Integer = GdipCreateBitmapFromGdiDib(bminfo, pixdat, img)
			Dim picPath As String

			If fbd.ShowDialogSave() = DialogResult.OK Then
				MsgBox("so, hier kann der Datenbankeintrag erfolgen")
				picPath = fbd.ReturnPath + "\" + picname + ".tif"
				If Not GetCodecClsid(picPath, clsid) Then
					MessageBox.Show("Unbekanntes Bildformat " + Path.GetExtension(picPath), "Image Codec", MessageBoxButtons.OK, MessageBoxIcon.Information)
					Return False
				End If

				If (st <> 0) Or (Equals(img, IntPtr.Zero)) Then
					Return False
				End If
				st = GdipSaveImageToFile(img, picPath, clsid, IntPtr.Zero)
				GdipDisposeImage(img)
				Return st = 0
			Else
				Return False
			End If

		End Function

	End Class

End Namespace
