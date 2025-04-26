VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScannerWIA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_DeviceManager As WIA.DeviceManager
Private m_CommonDialog  As WIA.CommonDialog
Private m_DeviceInfo    As WIA.DeviceInfo
Private m_Device        As WIA.Device
Private m_Image         As WIA.ImageFile

Private Sub Class_Initialize()
    Set m_DeviceManager = New WIA.DeviceManager
    Set m_CommonDialog = New WIA.CommonDialog
End Sub

Public Function TryGetDeviceNames(DeviceNames_out() As String) As Boolean
Try: On Error GoTo Catch
    Dim di As DeviceInfo
    Dim s As String
    Dim i As Integer, u As Integer: u = m_DeviceManager.DeviceInfos.Count - 1
    If u < 0 Then Exit Function
    ReDim DeviceNames_out(0 To u) As String
    For i = 0 To u
        Set di = m_DeviceManager.DeviceInfos.Item(i + 1)
        DeviceNames_out(i) = di.Properties("Name") & "; " & di.DeviceID & "; " & WIADeviceType_ToStr(di.Type)
    Next
    TryGetDeviceNames = True
    Exit Function
Catch:
    If Err.Number = &H80210015 Then
        MsgBox "WIA device not found!"
    End If
End Function

Private Function WIADeviceType_ToStr(e As WiaDeviceType) As String
    Dim s As String
    Select Case e
    Case WiaDeviceType.CameraDeviceType:      s = "Camera"
    Case WiaDeviceType.ScannerDeviceType:     s = "Scanner"
    Case WiaDeviceType.VideoDeviceType:       s = "Video"
    Case WiaDeviceType.UnspecifiedDeviceType: s = "Unspecified"
    Case Else:                                s = "Unspecified?"
    End Select
    WIADeviceType_ToStr = s
End Function

Public Function DeviceInfos_ToStr() As String
    If m_DeviceInfo Is Nothing Then
        MsgBox "Select device first!"
        Exit Function
    End If
    Dim s As String
    Dim i As Integer, c As Integer: c = m_DeviceInfo.Properties.Count
    Debug.Print c
    Dim pp As WIA.Property
    For i = 1 To c
        Set pp = m_DeviceInfo.Properties.Item(i)
        s = s & pp.Name & ": " & pp.Value & vbCrLf
    Next
    DeviceInfos_ToStr = s
End Function

Public Sub SelectDevice(ByVal DeviceNameAndID As String)
Try: On Error GoTo Catch
    Dim sn() As String: sn = Split(DeviceNameAndID, "; ")
    Dim di As DeviceInfo
    Dim i As Integer, c As Integer: c = m_DeviceManager.DeviceInfos.Count
    If c <= 0 Then Exit Sub
    For i = 1 To c
        Set di = m_DeviceManager.DeviceInfos.Item(i)
        If di.Properties("Name") = sn(0) Then
            Set m_DeviceInfo = di
            Set m_Device = m_DeviceInfo.Connect
            Exit Sub
        End If
    Next
Catch:
    MsgBox "WIA device not found!"
End Sub

Public Sub ShowDevicePropertiesDialog()
    If m_Device Is Nothing Then
        MsgBox "Select device first!"
        Exit Sub
    End If
    m_CommonDialog.ShowDeviceProperties m_Device
End Sub

Public Function Scan() As StdPicture
Try: On Error GoTo Catch
    Dim Image As WIA.ImageFile: Set Image = m_CommonDialog.ShowAcquireImage
    If Image Is Nothing Then Exit Function
    Dim tmpPFN As String: tmpPFN = Environ("tmp") & "\WIAImage.bmp"
    'Kill tmpPFN
    Image.SaveFile tmpPFN
    Set Scan = LoadPicture(tmpPFN)
    Kill tmpPFN
    Exit Function
Catch:
    Select Case Err.Number
    Case &H80210015
        MsgBox "WIA device not found!"
    Case &H80070050
        MsgBox "Could not save image file, maybe file already exists!"
    Case 53
        MsgBox Err.Description
    End Select
    Debug.Print Err.Number
End Function

