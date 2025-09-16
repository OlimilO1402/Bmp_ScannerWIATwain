VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScannerTwain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function TWAIN_AcquireToClipboard Lib "eztw32" (ByVal hwndApp As Long, ByVal wPixTypes As Long) As Long
Private Declare Function TWAIN_SelectImageSource Lib "eztw32" (ByVal hwndApp As Long) As Long

Private m_hWnd As Long 'Ptr

Friend Sub New_(Owner As Form)
    m_hWnd = Owner.hWnd
End Sub

Public Function SelectDevice() As Boolean
Try: On Error GoTo Catch
    Dim hr As Long: hr = TWAIN_SelectImageSource(m_hWnd) ' Scanquelle Auswahlfenster anzeigen
    'Debug.Print "hr: " & hr
    SelectDevice = hr <> 0
    Exit Function
Catch:
End Function

Public Function Scan() As StdPicture
Try: On Error GoTo Catch
    Dim hr As Long: hr = TWAIN_AcquireToClipboard(m_hWnd, 0) ' show the dialog of your imaging scanner device
    Set Scan = Clipboard.GetData(ClipBoardConstants.vbCFBitmap) ' insert BMP-picture from Clipboard
Catch:
End Function

