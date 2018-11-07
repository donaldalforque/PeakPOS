VERSION 5.00
Begin VB.Form SYS_DateTimeCheckerFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnChange 
      Caption         =   "Change Date and Time"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   2640
      Width           =   2655
   End
   Begin VB.CommandButton btnContinue 
      Caption         =   "Continue"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   2
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Timer timer_date 
      Interval        =   1000
      Left            =   10560
      Top             =   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "System Date and Time Settings"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   0
      Picture         =   "SYS_DateTimeCheckerFrm.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7485
   End
   Begin VB.Label lbldate 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   6900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current system Date and Time:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   390
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   3960
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BorderColor     =   &H00808080&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1815
      Left            =   120
      Top             =   720
      Width           =   7215
   End
End
Attribute VB_Name = "SYS_DateTimeCheckerFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnChange_Click()
    SYS_ChangeDateTimeFrm.Show (1)
End Sub

Private Sub btnContinue_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    lbldate.Caption = Format(Now, "MMMM DD, YYYY hh:mm:ss ampm")
End Sub

Private Sub timer_date_Timer()
lbldate.Caption = Format(Now, "MMMM DD, YYYY hh:mm:ss ampm")
End Sub
