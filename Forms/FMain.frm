VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "ScannerWIATwain"
   ClientHeight    =   8475
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13095
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8475
   ScaleWidth      =   13095
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnProperties 
      Caption         =   "WIA DeviceInfo"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   0
      Width           =   1815
   End
   Begin VB.ComboBox CmbZoom 
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   0
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   0
      ScaleHeight     =   6495
      ScaleWidth      =   12975
      TabIndex        =   0
      Top             =   360
      Width           =   12975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Zoom:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   555
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileImport 
         Caption         =   "Import"
         Begin VB.Menu mnuFileImportTwain 
            Caption         =   "Twain"
            Begin VB.Menu mnuFileImportTwainSelectSource 
               Caption         =   "Select Source..."
            End
            Begin VB.Menu mnuFileImportTwainProperties 
               Caption         =   "Properties..."
               Visible         =   0   'False
            End
            Begin VB.Menu mnuFileImportTwainRead 
               Caption         =   "Read..."
            End
         End
         Begin VB.Menu mnuFileImportWIA 
            Caption         =   "WIA"
            Begin VB.Menu mnuFileImportWIASelectSource 
               Caption         =   "Select Source..."
            End
            Begin VB.Menu mnuFileImportWIAProperties 
               Caption         =   "Properties..."
            End
            Begin VB.Menu mnuFileImportWIARead 
               Caption         =   "Read..."
            End
         End
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuViewZoomNormal 
         Caption         =   "Normal 1:1"
      End
      Begin VB.Menu mnuViewZoomIn_ 
         Caption         =   "Zoom In"
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "2:1"
            Index           =   2
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "3:1"
            Index           =   3
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "4:1"
            Index           =   4
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "5:1"
            Index           =   5
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "6:1"
            Index           =   6
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "7:1"
            Index           =   7
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "8:1"
            Index           =   8
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "9:1"
            Index           =   9
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "10:1"
            Index           =   10
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "11:1"
            Index           =   11
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "12:1"
            Index           =   12
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "13:1"
            Index           =   13
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "14:1"
            Index           =   14
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "15:1"
            Index           =   15
         End
         Begin VB.Menu mnuViewZoomIn 
            Caption         =   "16:1"
            Index           =   16
         End
      End
      Begin VB.Menu mnuViewZoomOut_ 
         Caption         =   "Zoom Out"
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:2"
            Index           =   2
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:3"
            Index           =   3
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:4"
            Index           =   4
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:5"
            Index           =   5
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:6"
            Index           =   6
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:7"
            Index           =   7
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:8"
            Index           =   8
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:9"
            Index           =   9
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:10"
            Index           =   10
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:11"
            Index           =   11
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:12"
            Index           =   12
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:13"
            Index           =   13
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:14"
            Index           =   14
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:15"
            Index           =   15
         End
         Begin VB.Menu mnuViewZoomOut 
            Caption         =   "1:16"
            Index           =   16
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   " ? "
      Begin VB.Menu mnuHelpInfo 
         Caption         =   "Info"
      End
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'https://access-im-unternehmen.de/Dokumente_scannen_mit_WIA/
'Verweise:
'Microsoft Windows Image Acquisition Library v2.0
'C:\Windows\System32\wiaaut.dll
Private m_ScanTwain As ScannerTwain
Private m_ScanWIA   As ScannerWIA
Private m_Image     As IPictureDisp 'StdPicture
Private m_PBZoom    As PictureBoxZoom

Private Sub Form_Load()
    Set m_ScanTwain = MNew.ScannerTwain(Me)
    Set m_ScanWIA = New ScannerWIA
    InitZoom
    Set m_PBZoom = MNew.PictureBoxZoom(Me, Me.Picture1, m_Image)
End Sub

Private Sub Form_Resize()
    Dim L As Single: L = Picture1.Left
    Dim T As Single: T = Picture1.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight
    If W > 0 And H > 0 Then Picture1.Move L, T, W, H: m_PBZoom.Refresh
End Sub

Sub InitZoom()
    Dim i As Long
    For i = 16 To 1 Step -1: CmbZoom.AddItem CStr(i) & ":1": Next
    For i = 2 To 16:         CmbZoom.AddItem "1:" & CStr(i): Next
    If m_PBZoom Is Nothing Then CmbZoom.Text = "1:1" Else CmbZoom.Text = m_PBZoom.PropZoomToStr
    mnuViewZoomNormal.Checked = True
End Sub

Private Sub BtnProperties_Click()
    Dim s As String: s = m_ScanWIA.DeviceInfos_ToStr
    If Len(s) = 0 Then Exit Sub
    MsgBox s
End Sub

Private Sub mnuFileImportTwainSelectSource_Click()
    m_ScanTwain.SelectDevice
End Sub

Private Sub mnuFileImportTwainRead_Click()
    GetScannedImage m_ScanTwain
End Sub

Private Sub mnuFileImportWIASelectSource_Click()
    Dim DeviceNames() As String
    If Not m_ScanWIA.TryGetDeviceNames(DeviceNames) Then
        MsgBox "No devices found!"
        Exit Sub
    End If
    Dim SelDevice As String
    If FSelect.ShowDialog(Me, "Select Source", DeviceNames, SelDevice) = vbCancel Then Exit Sub
    m_ScanWIA.SelectDevice SelDevice
End Sub

Private Sub mnuFileImportWIAProperties_Click()
    m_ScanWIA.ShowDevicePropertiesDialog
End Sub

Private Sub mnuFileImportWIARead_Click()
    GetScannedImage m_ScanWIA
End Sub

Private Sub GetScannedImage(ImageScanner)
    Dim img As IPictureDisp: Set img = ImageScanner.Scan
    If img Is Nothing Then MsgBox "Image not found!": Exit Sub
    Set m_Image = img
    If m_PBZoom Is Nothing Then
        Set m_PBZoom = MNew.PictureBoxZoom(Me, Me.Picture1, m_Image)
    Else
        Set m_PBZoom.Image = m_Image
    End If
End Sub

Private Sub mnuEditCut_Click()
    Clipboard.SetData Picture1.Picture, ClipBoardConstants.vbCFBitmap
    Set Picture1.Picture = Nothing
    Picture1.Cls
End Sub

Private Sub mnuEditCopy_Click()
    Clipboard.SetData Picture1.Picture, ClipBoardConstants.vbCFBitmap
End Sub

Private Sub mnuEditPaste_Click()
    If Not Clipboard.GetFormat(ClipBoardConstants.vbCFBitmap) Then Exit Sub
    Set m_Image = Clipboard.GetData(ClipBoardConstants.vbCFBitmap)
    Set m_PBZoom = MNew.PictureBoxZoom(Me, Me.Picture1, m_Image)
    CmbZoom.Text = m_PBZoom.PropZoomToStr
    mnuViewZoom_UnCheckAll
    mnuViewZoomNormal.Checked = True
End Sub

Private Sub mnuHelpInfo_Click()
    MsgBox App.CompanyName & " " & App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & _
           App.FileDescription
End Sub

Private Sub mnuViewZoom_UnCheckAll()
    Dim i As Integer
    For i = mnuViewZoomIn.LBound To mnuViewZoomIn.UBound:   mnuViewZoomIn(i).Checked = False:  Next
    For i = mnuViewZoomOut.LBound To mnuViewZoomOut.UBound: mnuViewZoomOut(i).Checked = False: Next
    mnuViewZoomNormal.Checked = False
End Sub

Private Sub mnuViewZoomNormal_Click()
    If m_PBZoom Is Nothing Then Exit Sub
    m_PBZoom.ZoomFactor = 1
    CmbZoom.Text = m_PBZoom.PropZoomToStr
    mnuViewZoom_UnCheckAll
    mnuViewZoomNormal.Checked = True
End Sub

Private Sub mnuViewZoomIn_Click(Index As Integer)
    If m_Image Is Nothing Then Exit Sub
    m_PBZoom.ZoomFactor = Index
    CmbZoom.Text = m_PBZoom.PropZoomToStr
    mnuViewZoom_UnCheckAll
    mnuViewZoomIn(Index).Checked = True
End Sub

Private Sub mnuViewZoomOut_Click(Index As Integer)
    If m_PBZoom Is Nothing Then Exit Sub
    m_PBZoom.ZoomFactor = 1 / Index
    CmbZoom.Text = m_PBZoom.PropZoomToStr
    mnuViewZoom_UnCheckAll
    mnuViewZoomOut(Index).Checked = True
End Sub

Private Sub CmbZoom_Click()
    If m_PBZoom Is Nothing Then Exit Sub
    Dim li As Long:     li = CmbZoom.ListIndex - 15
    Dim z As Double: z = IIf(li = 0, 1, IIf(li < 1, Abs(li) + 1, 1))
    Dim n As Double: n = IIf(li = 0, 1, IIf(li < 1, 1, Abs(li) + 1))
    m_PBZoom.ZoomFactor = z / n
    mnuViewZoom_UnCheckAll
    If z = 1 And n = 1 Then
        mnuViewZoomNormal.Checked = True
    ElseIf z = 1 Then
        mnuViewZoomOut(CLng(n)).Checked = True
    ElseIf n = 1 Then
        mnuViewZoomIn(CLng(z)).Checked = True
    End If
End Sub
'
'Sub UpdateView()
'    m_PBZoom.Refresh
'End Sub

