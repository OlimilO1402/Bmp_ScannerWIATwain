Public Class MyFolderBrowserDialog
	Inherits Windows.Forms.Design.FolderNameEditor

	Private fb As New Windows.Forms.Design.FolderNameEditor.FolderBrowser
	Private fbReturnPath As String
	Public fbDescription As String

	Public ReadOnly Property ReturnPath() As String
		Get
			Return fbReturnPath
		End Get
	End Property

	Public Function ShowDialogChoose() As DialogResult

		Dim Result As DialogResult
		fb.Description = Me.fbDescription
		fb.StartLocation = System.Windows.Forms.Design.FolderNameEditor.FolderBrowserFolder.MyComputer
		fb.Style = System.Windows.Forms.Design.FolderNameEditor.FolderBrowserStyles.BrowseForEverything
		Result = fb.ShowDialog
		If (Result = DialogResult.OK) Then
			Me.fbReturnPath = fb.DirectoryPath
		Else
			Me.fbReturnPath = String.Empty
		End If
		Return Result

	End Function

	Public Function ShowDialogSave() As DialogResult

		Dim Result As DialogResult
		fb.Description = Me.fbDescription
		fb.StartLocation = System.Windows.Forms.Design.FolderNameEditor.FolderBrowserFolder.MyComputer
		fb.Style = System.Windows.Forms.Design.FolderNameEditor.FolderBrowserStyles.BrowseForComputer
		Result = fb.ShowDialog
		If (Result = DialogResult.OK) Then
			Me.fbReturnPath = fb.DirectoryPath
		Else
			Me.fbReturnPath = String.Empty
		End If
		Return Result

	End Function

End Class
