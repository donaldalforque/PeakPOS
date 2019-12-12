VERSION 5.00
Begin VB.Form SYS_FormPassFrm 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   630
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   630
   ScaleWidth      =   5370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPass 
      BackColor       =   &H00C0C0FF&
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "A"
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "SYS_FormPassFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txtPass_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtPass.Text = "PeakPOS2015" Then
            SYS_WorkstationFrm.Show (1)
            Unload Me
        Else
            MsgBox "Invalid code.", vbCritical
        End If
    End If
End Sub
