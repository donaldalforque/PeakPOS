VERSION 5.00
Begin VB.Form POS_CashBreakDownFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6870
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   6870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCheckQty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   1920
      TabIndex        =   10
      Text            =   "0"
      Top             =   7080
      Width           =   735
   End
   Begin VB.TextBox txtCheckAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2760
      TabIndex        =   11
      Text            =   "0"
      Top             =   7080
      Width           =   975
   End
   Begin VB.CommandButton btnAccept 
      Caption         =   "Enter: Print"
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
      Left            =   3480
      Picture         =   "POS_CashBreakDownFrm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   8640
      Width           =   1575
   End
   Begin VB.TextBox txtCents 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2760
      TabIndex        =   9
      Text            =   "0"
      Top             =   6480
      Width           =   975
   End
   Begin VB.TextBox txt1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2760
      TabIndex        =   8
      Text            =   "0"
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox txt5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2760
      TabIndex        =   7
      Text            =   "0"
      Top             =   5280
      Width           =   975
   End
   Begin VB.TextBox txt10 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2760
      TabIndex        =   6
      Text            =   "0"
      Top             =   4680
      Width           =   975
   End
   Begin VB.TextBox txt20 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2760
      TabIndex        =   5
      Text            =   "0"
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox txt50 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2760
      TabIndex        =   4
      Text            =   "0"
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox txt100 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2760
      TabIndex        =   3
      Text            =   "0"
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox txt200 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2760
      TabIndex        =   2
      Text            =   "0"
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txt500 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2760
      TabIndex        =   1
      Text            =   "0"
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox txt1000 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2760
      TabIndex        =   0
      Text            =   "0"
      Top             =   1080
      Width           =   975
   End
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
      Left            =   5160
      Picture         =   "POS_CashBreakDownFrm.frx":2228
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8640
      Width           =   1575
   End
   Begin VB.Label lblchecks 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   58
      Top             =   7080
      Width           =   2055
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   57
      Top             =   7080
      Width           =   195
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Checks"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   56
      Top             =   7080
      Width           =   1515
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cents"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   54
      Top             =   6480
      Width           =   1515
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   53
      Top             =   6480
      Width           =   195
   End
   Begin VB.Label lblCents 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   52
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   51
      Top             =   7920
      Width           =   3255
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   6600
      Y1              =   8520
      Y2              =   8520
   End
   Begin VB.Label lbltotal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      TabIndex        =   50
      Top             =   7920
      Width           =   3255
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6600
      Y1              =   7800
      Y2              =   7800
   End
   Begin VB.Label Label36 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   49
      Top             =   5880
      Width           =   180
   End
   Begin VB.Label lbl1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   48
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   47
      Top             =   5880
      Width           =   195
   End
   Begin VB.Label Label33 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "1.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   46
      Top             =   5880
      Width           =   1515
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   45
      Top             =   5280
      Width           =   180
   End
   Begin VB.Label lbl5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   44
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   43
      Top             =   5280
      Width           =   195
   End
   Begin VB.Label Label29 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "5.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   42
      Top             =   5280
      Width           =   1515
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   41
      Top             =   4680
      Width           =   180
   End
   Begin VB.Label lbl10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   40
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   39
      Top             =   4680
      Width           =   195
   End
   Begin VB.Label Label25 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "10.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   38
      Top             =   4680
      Width           =   1515
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   37
      Top             =   4080
      Width           =   180
   End
   Begin VB.Label lbl20 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   36
      Top             =   4080
      Width           =   2055
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   35
      Top             =   4080
      Width           =   195
   End
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "20.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   34
      Top             =   4080
      Width           =   1515
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   33
      Top             =   3480
      Width           =   180
   End
   Begin VB.Label lbl50 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   32
      Top             =   3480
      Width           =   2055
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   31
      Top             =   3480
      Width           =   195
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "50.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   30
      Top             =   3480
      Width           =   1515
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   29
      Top             =   2880
      Width           =   180
   End
   Begin VB.Label lbl100 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   28
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   27
      Top             =   2880
      Width           =   195
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "100.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   26
      Top             =   2880
      Width           =   1515
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   25
      Top             =   2280
      Width           =   180
   End
   Begin VB.Label lbl200 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   24
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   23
      Top             =   2280
      Width           =   195
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "200.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   22
      Top             =   2280
      Width           =   1515
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   21
      Top             =   1680
      Width           =   180
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   20
      Top             =   1080
      Width           =   180
   End
   Begin VB.Label lbl500 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   19
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   18
      Top             =   1680
      Width           =   195
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "500.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   1680
      Width           =   1515
   End
   Begin VB.Label lbl1000 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   16
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   15
      Top             =   1080
      Width           =   195
   End
   Begin VB.Image picModuleImage 
      Height          =   480
      Left            =   240
      Picture         =   "POS_CashBreakDownFrm.frx":45B7
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CASH BREAKDOWN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   840
      TabIndex        =   14
      Top             =   195
      Width           =   2925
   End
   Begin VB.Label lblCash 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1,000.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   13
      Top             =   1080
      Width           =   1470
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   -120
      Picture         =   "POS_CashBreakDownFrm.frx":4BDC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "POS_CashBreakDownFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAccept_Click()
    Dim x As Variant
    x = MsgBox("Are you sure all information are correct?", vbQuestion + vbYesNo)
    If x = vbYes Then
        Dim CashBreakDownId As Long
        Dim con As New ADODB.Connection
        Set rec = New ADODB.Recordset
        Dim cmd As New ADODB.Command
        
        con.ConnectionString = ConnString
        con.Open
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "POS_CashBreakDown_Insert"
        cmd.Parameters.Append cmd.CreateParameter("@POS_CashBreakDownId", adInteger, adParamInputOutput, , 0)
        cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
        cmd.Parameters.Append cmd.CreateParameter("@WorkStationId", adInteger, adParamInput, , WorkstationId)
        cmd.Parameters.Append cmd.CreateParameter("@Total", adDecimal, adParamInput, , NVAL(lbltotal.Caption))
                              cmd.Parameters("@Total").NumericScale = 2
                              cmd.Parameters("@Total").Precision = 18
        cmd.Execute
        
        CashBreakDownId = cmd.Parameters("@POS_CashBreakDownId")
        
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "POS_CashBreakDownLine_Insert"
        
        cmd.Parameters.Append cmd.CreateParameter("@POS_CashBreakDownId", adInteger, adParamInput, , CashBreakDownId)
        cmd.Parameters.Append cmd.CreateParameter("@1000", adDecimal, adParamInput, , NVAL(txt1000.Text))
                              cmd.Parameters("@1000").NumericScale = 2
                              cmd.Parameters("@1000").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@500", adDecimal, adParamInput, , NVAL(txt500.Text))
                              cmd.Parameters("@500").NumericScale = 2
                              cmd.Parameters("@500").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@200", adDecimal, adParamInput, , NVAL(txt200.Text))
                              cmd.Parameters("@200").NumericScale = 2
                              cmd.Parameters("@200").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@100", adDecimal, adParamInput, , NVAL(txt100.Text))
                              cmd.Parameters("@100").NumericScale = 2
                              cmd.Parameters("@100").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@50", adDecimal, adParamInput, , NVAL(txt50.Text))
                              cmd.Parameters("@50").NumericScale = 2
                              cmd.Parameters("@50").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@20", adDecimal, adParamInput, , NVAL(txt20.Text))
                              cmd.Parameters("@20").NumericScale = 2
                              cmd.Parameters("@20").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@10", adDecimal, adParamInput, , NVAL(txt10.Text))
                              cmd.Parameters("@10").NumericScale = 2
                              cmd.Parameters("@10").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@5", adDecimal, adParamInput, , NVAL(txt5.Text))
                              cmd.Parameters("@5").NumericScale = 2
                              cmd.Parameters("@5").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@1", adDecimal, adParamInput, , NVAL(txt1.Text))
                              cmd.Parameters("@1").NumericScale = 2
                              cmd.Parameters("@1").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@Cents", adDecimal, adParamInput, , NVAL(txtCents.Text) / 100)
                              cmd.Parameters("@Cents").NumericScale = 2
                              cmd.Parameters("@Cents").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@CheckQty", adDecimal, adParamInput, , NVAL(txtCheckQty.Text))
                              cmd.Parameters("@CheckQty").NumericScale = 2
                              cmd.Parameters("@CheckQty").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , NVAL(txtCheckAmount.Text))
                              cmd.Parameters("@CheckAmount").NumericScale = 2
                              cmd.Parameters("@CheckAmount").Precision = 18
        cmd.Execute
        con.Close
        
       
        '**PRINT RECEIPT******
        Dim crxApp As New CRAXDRT.Application
        Dim crxRpt As New CRAXDRT.Report
        Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\POS_CashBreakDown.rpt")
        crxRpt.DiscardSavedData
        crxRpt.EnableParameterPrompting = False
        crxRpt.ParameterFields.GetItemByName("@CashBreakDownId").AddCurrentValue CashBreakDownId

        Call ResetRptDB(crxRpt)
        crxRpt.PrintOut False

        '**END PRINT RECEIPT**
        
        MsgBox "Press enter to close window", vbInformation
        Unload Me
    End If
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Public Sub CountTotal()
    Dim onek, fiveh, twoh, oneh, fivef, twof, tenf, fivec, onec, cents, Total, checks As Double
    onek = NVAL(txt1000) * 1000
    lbl1000.Caption = FormatNumber(onek, 2, vbTrue, vbFalse)
    
    fiveh = NVAL(txt500) * 500
    lbl500.Caption = FormatNumber(fiveh, 2, vbTrue, vbFalse)
    
    twoh = NVAL(txt200) * 200
    lbl200.Caption = FormatNumber(twoh, 2, vbTrue, vbFalse)
    
    oneh = NVAL(txt100) * 100
    lbl100.Caption = FormatNumber(oneh, 2, vbTrue, vbFalse)
    
    fivef = NVAL(txt50) * 50
    lbl50.Caption = FormatNumber(fivef, 2, vbTrue, vbFalse)
    
    twof = NVAL(txt20) * 20
    lbl20.Caption = FormatNumber(twof, 2, vbTrue, vbFalse)
    
    tenf = NVAL(txt10) * 10
    lbl10.Caption = FormatNumber(tenf, 2, vbTrue, vbFalse)
    
    fivec = NVAL(txt5) * 5
    lbl5.Caption = FormatNumber(fivec, 2, vbTrue, vbFalse)
    
    onec = NVAL(txt1) * 1
    lbl1.Caption = FormatNumber(onec, 2, vbTrue, vbFalse)
    
    cents = NVAL(txtCents.Text) / 100
    lblCents.Caption = FormatNumber(cents, 2, vbTrue, vbFalse)
    
    checks = NVAL(txtCheckAmount.Text)
    lblchecks.Caption = FormatNumber(checks, 2, vbTrue, vbFalse)
    
    Total = onek + fiveh + twoh + oneh + fivef + twof + tenf + fivec + onec + cents + checks
    lbltotal.Caption = FormatNumber(Total, 2, vbTrue, vbFalse)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            btnAccept_Click
        Case vbKeyEscape
            btnCancel_Click
    End Select
End Sub

Private Sub txt1_Change()
    If IsNumeric(txt1.Text) = False Then
        txt1.Text = "0"
        selectText txt1
    Else
        CountTotal
    End If
End Sub

Private Sub txt1_GotFocus()
    selectText txt1
End Sub

Private Sub txt10_Change()
    If IsNumeric(txt10.Text) = False Then
        txt10.Text = "0"
        selectText txt10
    Else
        CountTotal
    End If
End Sub

Private Sub txt10_GotFocus()
    selectText txt10
End Sub

Private Sub txt100_Change()
    If IsNumeric(txt100.Text) = False Then
        txt100.Text = "0"
        selectText txt100
    Else
        CountTotal
    End If
End Sub

Private Sub txt100_GotFocus()
    selectText txt100
End Sub

Private Sub txt1000_Change()
    If IsNumeric(txt1000.Text) = False Then
        txt1000.Text = "0"
        selectText txt1000
    Else
        CountTotal
    End If
End Sub

Private Sub txt1000_GotFocus()
    selectText txt1000
End Sub

Private Sub txt20_Change()
    If IsNumeric(txt20.Text) = False Then
        txt20.Text = "0"
        selectText txt20
    Else
        CountTotal
    End If
End Sub

Private Sub txt20_GotFocus()
    selectText txt20
End Sub

Private Sub txt200_Change()
    If IsNumeric(txt200.Text) = False Then
        txt200.Text = "0"
        selectText txt200
    Else
        CountTotal
    End If
End Sub

Private Sub txt200_GotFocus()
    selectText txt200
End Sub

Private Sub txt5_Change()
    If IsNumeric(txt5.Text) = False Then
        txt5.Text = "0"
        selectText txt5
    Else
        CountTotal
    End If
End Sub

Private Sub txt5_GotFocus()
    selectText txt5
End Sub

Private Sub txt50_Change()
    If IsNumeric(txt50.Text) = False Then
        txt50.Text = "0"
        selectText txt50
    Else
        CountTotal
    End If
End Sub

Private Sub txt50_GotFocus()
    selectText txt50
End Sub

Private Sub txt500_Change()
    If IsNumeric(txt500.Text) = False Then
        txt500.Text = "0"
        selectText txt500
    Else
        CountTotal
    End If
End Sub

Private Sub txt500_GotFocus()
    selectText txt500
End Sub

Private Sub txtCents_Change()
    If IsNumeric(txtCents.Text) = False Then
        txtCents.Text = "0"
        selectText txtCents
    Else
        CountTotal
    End If
End Sub

Private Sub txtCents_GotFocus()
    selectText txtCents
End Sub

Private Sub txtCheckAmount_Change()
    If IsNumeric(txtCheckAmount.Text) = False Then
        txtCheckAmount.Text = "0"
        selectText txtCheckAmount
    Else
        CountTotal
    End If
End Sub

Private Sub txtCheckAmount_GotFocus()
    selectText txtCheckAmount
End Sub

Private Sub txtCheckQty_Change()
    If IsNumeric(txtCheckQty.Text) = False Then
        txtCheckQty.Text = "0"
        selectText txtCheckQty
    Else
        CountTotal
    End If
End Sub

Private Sub txtCheckQty_GotFocus()
     selectText txtCheckQty
End Sub
