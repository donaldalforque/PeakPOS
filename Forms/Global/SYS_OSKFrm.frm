VERSION 5.00
Begin VB.Form SYS_OSKFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3960
   Icon            =   "SYS_OSKFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   3960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCancel 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7560
      Width           =   3975
   End
   Begin VB.CommandButton btnAccept 
      BackColor       =   &H0080FF80&
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6240
      Width           =   3975
   End
   Begin VB.CommandButton btnDot 
      BackColor       =   &H00C0FFFF&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox txtValue 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3975
   End
   Begin VB.CommandButton btn0 
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton btnClear 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4920
      Width           =   1335
   End
   Begin VB.CommandButton btn9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton btn8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton btn7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton btn6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton btn5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton btn4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton btn3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton btn2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton btn1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "SYS_OSKFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public txtControl As TextBox

Private Sub btn0_Click()
    txtValue.text = txtValue.text & "0"
End Sub

Private Sub btn1_Click()
    txtValue.text = txtValue.text & "1"
End Sub

Private Sub btn2_Click()
    txtValue.text = txtValue.text & "2"
End Sub

Private Sub btn3_Click()
    txtValue.text = txtValue.text & "3"
End Sub

Private Sub btn4_Click()
    txtValue.text = txtValue.text & "4"
End Sub

Private Sub btn5_Click()
    txtValue.text = txtValue.text & "5"
End Sub

Private Sub btn6_Click()
    txtValue.text = txtValue.text & "6"
End Sub

Private Sub btn7_Click()
    txtValue.text = txtValue.text & "7"
End Sub

Private Sub btn8_Click()
    txtValue.text = txtValue.text & "8"
End Sub

Private Sub btn9_Click()
    txtValue.text = txtValue.text & "9"
End Sub

Private Sub btnAccept_Click()
    If txtValue.text = "" Then txtValue.text = "1.00"
    If IsNumeric(txtValue.text) = False Then txtValue.text = "1.00"
    txtControl.text = txtValue.text
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnClear_Click()
    If txtValue.text = "" Then Exit Sub
    txtValue.text = Mid(txtValue.text, 1, Len(txtValue.text) - 1)
End Sub

Private Sub btnDot_Click()
    txtValue.text = txtValue.text & "."
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKey1, vbKeyNumpad1
            btn1_Click
        Case vbKey2, vbKeyNumpad2
            btn2_Click
        Case vbKey3, vbKeyNumpad3
            btn3_Click
        Case vbKey4, vbKeyNumpad4
            btn4_Click
        Case vbKey5, vbKeyNumpad5
            btn5_Click
        Case vbKey6, vbKeyNumpad6
            btn6_Click
        Case vbKey7, vbKeyNumpad7
            btn7_Click
        Case vbKey8, vbKeyNumpad8
            btn8_Click
        Case vbKey9, vbKeyNumpad9
            btn9_Click
        Case vbKey0, vbKeyNumpad0
            btn0_Click
        Case vbKeyBack
            btnClear_Click
        Case vbKeyEscape
            btnCancel_Click
        Case vbKeyReturn
            btnAccept_Click
    End Select
End Sub
'Private Sub btnClear_Click()
'    If txtPin.text = "" Then Exit Sub
'    txtPin.text = Mid(txtPin.text, 1, Len(txtPin.text) - 1)
'
'End Sub


