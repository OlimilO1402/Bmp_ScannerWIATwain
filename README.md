# Bmp_ScannerWIATwain  
## Shows how to use your scanning device using TWAIN or WIA  

[![GitHub](https://img.shields.io/github/license/OlimilO1402/Bmp_ScannerWIATwain?style=plastic)](https://github.com/OlimilO1402/Bmp_ScannerWIATwain/blob/master/LICENSE) 
[![GitHub release (latest by date)](https://img.shields.io/github/v/release/OlimilO1402/Bmp_ScannerWIATwain?style=plastic)](https://github.com/OlimilO1402/Bmp_ScannerWIATwain/releases/latest)
[![Github All Releases](https://img.shields.io/github/downloads/OlimilO1402/Bmp_ScannerWIATwain/total.svg)](https://github.com/OlimilO1402/Bmp_ScannerWIATwain/releases/download/v1.0.0/ScannerWIATwain_v2025.1.30.zip)
![GitHub followers](https://img.shields.io/github/followers/OlimilO1402?style=social)


Project started in january 2025.  
This example shows the different possibilities we have in windows to use your imaging scanner devices.  
Either using the mighty TWAIN-interface with the, old but gold, EZTW32.dll or via the new WIA way with the windows dll "wiaaut.dll", and of course how to implement it in VB.  
First we have to select the scanner driver, then we can read a picture from the device.  
There are 2 classes ScannerTwain and ScannerWIA. There is no explicit Interface but both use of course the same main function names "Sub SelectDevice()" and "Function Scan() As StdPicture".  

Function Scan of the class ScannerTWAIN:  
```vba  
Public Function Scan() As StdPicture
    Dim hr As Long: hr = TWAIN_AcquireToClipboard(m_hWnd, 0)    ' show the dialog of your imaging scanner device
    Set Scan = Clipboard.GetData(ClipBoardConstants.vbCFBitmap) ' insert the scanned picture as a bmp-file from Clipboard
End Function
```

Function Scan of the class ScannerWIA:  
```vba  
Private m_DeviceManager As WIA.DeviceManager
Private m_CommonDialog  As WIA.CommonDialog
Private m_DeviceInfo    As WIA.DeviceInfo
Private m_Device        As WIA.Device
Private m_Image         As WIA.ImageFile

Private Sub Class_Initialize()
    Set m_DeviceManager = New WIA.DeviceManager
    Set m_CommonDialog = New WIA.CommonDialog
End Sub
'...
Public Function Scan() As StdPicture
Try: On Error GoTo Catch
    Dim Image As WIA.ImageFile: Set Image = m_CommonDialog.ShowAcquireImage '
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

```


[Twain.org](https://twain.org)  


![IEnumVarImpl Image](Resources/IEnumVarImpl.png "IEnumVarImpl Image")
