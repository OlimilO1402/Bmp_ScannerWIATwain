VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save As..."
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileImport 
         Caption         =   "Import"
         Begin VB.Menu mnuFileImportTwain 
            Caption         =   "Twain"
            Begin VB.Menu mnuFileImportTwainAcquire 
               Caption         =   "Acquire..."
            End
            Begin VB.Menu mnuFileImportTwainSelectsource 
               Caption         =   "Select Source..."
            End
         End
      End
      Begin VB.Menu mnuFileExport 
         Caption         =   "Export"
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuFileImportTwainAcquire_Click()
    '
End Sub

Private Sub mnuFileImportTwainSelectsource_Click()
    'start DataSourceManager
    
End Sub
