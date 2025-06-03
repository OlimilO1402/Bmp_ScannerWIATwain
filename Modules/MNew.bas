Attribute VB_Name = "MNew"
Option Explicit

Public Function PictureBoxZoom(Window As Form, Canvas As PictureBox, aImage As StdPicture) As PictureBoxZoom
    Set PictureBoxZoom = New PictureBoxZoom: PictureBoxZoom.New_ Window, Canvas, aImage
End Function

Public Function ScannerTwain(Owner As Form) As ScannerTwain
    Set ScannerTwain = New ScannerTwain: ScannerTwain.New_ Owner
End Function
