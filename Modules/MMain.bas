Attribute VB_Name = "MMain"
Option Explicit

Sub Main()
    Dim aPFN As String: aPFN = App.Path & "\eztw32.dll"
    If Dir(aPFN) = "" Then
        MsgBox "File not found, trying to write the file: " & vbCrLf & aPFN
        Dim bin() As Byte: bin = LoadResData(101, "CUSTOM")
        If Not WriteFile(bin, aPFN) Then
            MsgBox "Could not write the file: " & vbCrLf & aPFN
        End If
    End If
    FMain.Show
End Sub

Function WriteFile(bytes() As Byte, PFN As String) As Boolean
Try: On Error GoTo Catch
    Dim FNr As Integer: FNr = FreeFile
    Open PFN For Binary Access Write As FNr
    Put FNr, , bytes
    WriteFile = True
    GoTo Finally
Catch:
    MsgBox "Error writing the file: " & vbCrLf & PFN
Finally:
    Close FNr
End Function
