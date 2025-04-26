VERSION 5.00
Begin VB.Form frmTwainScanner 
   Caption         =   "Scanner und Visual Basic"
   ClientHeight    =   7275
   ClientLeft      =   90
   ClientTop       =   660
   ClientWidth     =   10905
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   10905
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox picScanBild 
      Align           =   1  'Oben ausrichten
      AutoSize        =   -1  'True
      Height          =   7290
      Left            =   0
      ScaleHeight     =   7230
      ScaleWidth      =   10845
      TabIndex        =   0
      Top             =   0
      Width           =   10905
   End
   Begin VB.Menu mnuDatei 
      Caption         =   "Datei"
      Begin VB.Menu mnuQuelle 
         Caption         =   "&Quelle auswählen"
      End
      Begin VB.Menu mnuScannen 
         Caption         =   "&Scan starten"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuStrich 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEnd 
         Caption         =   "&Beenden"
      End
   End
End
Attribute VB_Name = "frmTwainScanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function TWAIN_AcquireToClipboard Lib "EZTW32.DLL" (ByVal hwndApp As Long, ByVal wPixTypes As Long) As Long

Private Declare Function TWAIN_SelectImageSource Lib "EZTW32.DLL" (ByVal hwndApp As Long) As Long


Private Sub mnuEnd_Click()
    Unload Me ' Programmende
End Sub

Private Sub mnuQuelle_Click()
    Dim Ret As Long: Ret = TWAIN_SelectImageSource(Me.hWnd) ' Scanquelle Auswahlfenster anzeigen
    If Ret <> 0 Then mnuScannen.Enabled = True ' Nur wenn etwas gewählt wurde, Scan ermöglichen
End Sub

Private Sub mnuScannen_Click()
    Dim Ret As Long: Ret = TWAIN_AcquireToClipboard(Me.hWnd, 0) ' Scandialog des Scanners anzeigen
    picScanBild.Picture = Clipboard.GetData(ClipBoardConstants.vbCFDIB) ' BMP-Bild aus dem Clipboard einfügen
End Sub
