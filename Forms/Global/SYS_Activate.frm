VERSION 5.00
Begin VB.Form SYS_ActivateFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   750
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnActivate 
      Caption         =   "ACTIVATE"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "SYS_ActivateFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnActivate_Click()
    SaveSetting "PeakPOS", "Data", "Default", GetSerialNumber(App.Path)
    MsgBox "Success!", vbInformation
End Sub
