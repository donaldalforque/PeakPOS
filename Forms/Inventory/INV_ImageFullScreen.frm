VERSION 5.00
Begin VB.Form INV_ImageFullScreen 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   8415
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image picMain 
      Height          =   8415
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8415
   End
End
Attribute VB_Name = "INV_ImageFullScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    picMain.width = Me.width
    picMain.Height = Me.Height
End Sub

