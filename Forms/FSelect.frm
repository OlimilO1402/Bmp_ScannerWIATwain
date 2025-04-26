VERSION 5.00
Begin VB.Form FSelect 
   BorderStyle     =   5  'Änderbares Werkzeugfenster
   Caption         =   "Select"
   ClientHeight    =   2175
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9855
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnCancel 
      Cancel          =   -1  'True
      Caption         =   "Abbrechen"
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   1590
      ItemData        =   "FSelect.frx":0000
      Left            =   0
      List            =   "FSelect.frx":0002
      TabIndex        =   0
      Top             =   0
      Width           =   9855
   End
End
Attribute VB_Name = "FSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Items() As String
Private m_ItemSel As String
Private m_Result  As VbMsgBoxResult

Private Sub Form_Load()
    m_Result = VbMsgBoxResult.vbCancel
End Sub

Private Sub Form_Resize()
    Dim brdr As Single: brdr = 8 * Screen.TwipsPerPixelX
    Dim L As Single, T As Single
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight - BtnOK.Height - 2 * brdr
    If 0 < W And 0 < H Then
        List1.Move L, T, W, H
    End If
    L = Me.ScaleWidth / 2 - BtnOK.Width - brdr
    T = Me.ScaleHeight - BtnOK.Height - brdr
    W = BtnOK.Width: H = BtnOK.Height
    If 0 < W And 0 < H Then
        BtnOK.Move L, T, W, H
        L = Me.ScaleWidth / 2 + brdr
        BtnCancel.Move L, T, W, H
    End If
End Sub

Public Function ShowDialog(Owner As Form, Caption As String, Items() As String, ItemSelected_out As String) As VbMsgBoxResult
    Me.Caption = Caption
    m_Items = Items
    UpdateView
    Me.Show vbModal, Owner
    ItemSelected_out = m_ItemSel
    ShowDialog = m_Result
End Function

Private Sub UpdateView()
    List1.Clear
    Dim s 'As String
    For Each s In m_Items
        List1.AddItem s
    Next
End Sub

Private Function GetItemSelected() As String
    If List1.ListIndex < 0 Then List1.ListIndex = 0
    Dim li As Integer: li = List1.ListIndex
    GetItemSelected = List1.List(li)
End Function

Private Sub BtnOK_Click()
    m_ItemSel = GetItemSelected
    m_Result = VbMsgBoxResult.vbOK: Unload Me
End Sub
Private Sub BtnCancel_Click()
    m_Result = VbMsgBoxResult.vbCancel: Unload Me
End Sub

Private Sub List1_DblClick()
    m_ItemSel = GetItemSelected
    If Len(m_ItemSel) = 0 Then Exit Sub
    m_Result = VbMsgBoxResult.vbOK: Unload Me
End Sub
