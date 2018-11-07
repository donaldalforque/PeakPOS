VERSION 5.00
Begin VB.Form POS_BaggerFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4590
   Icon            =   "POS_BaggerFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCancel 
      Caption         =   "ESC: Cancel"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   2880
      Picture         =   "POS_BaggerFrm.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton btnAccept 
      Caption         =   "ENTER: Accept"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   1200
      Picture         =   "POS_BaggerFrm.frx":239B
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
   End
   Begin VB.TextBox txtBagger 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   4215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Please input Bagger's name:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   2460
   End
   Begin VB.Label lblCaption_Title 
      AutoSize        =   -1  'True
      Caption         =   "Bagger"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   900
   End
   Begin VB.Image picModuleImage 
      Height          =   480
      Left            =   240
      Picture         =   "POS_BaggerFrm.frx":476F
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "POS_BaggerFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAccept_Click()
If Trim(txtBagger.text) <> "" Then
    frmCashier.Bagger = txtBagger.text
    'frmCashier.lblBagger.Caption = txtBagger.Text
    Unload Me
Else
    MsgBox "Please input a bagger's name.", vbCritical, "CS Point of Sale"
    txtBagger.SetFocus
End If
End Sub

Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case 13 'Enter
        btnAccept_Click
    Case vbKeyEscape
        btnCancel_Click
End Select
End Sub

Private Sub Form_Load()
txtBagger.SelStart = 0
txtBagger.SelLength = Len(txtBagger.text)
End Sub
