VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form SO_CashAdvance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cash Advance"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   15090
   Begin VB.PictureBox picPaid 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1860
      Left            =   -99999
      Picture         =   "SO_CashAdvanceFrm.frx":0000
      ScaleHeight     =   1860
      ScaleWidth      =   5250
      TabIndex        =   0
      Top             =   2640
      Visible         =   0   'False
      Width           =   5250
   End
   Begin VB.PictureBox pic_Cancelled 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   -99999
      Picture         =   "SO_CashAdvanceFrm.frx":8F91
      ScaleHeight     =   2295
      ScaleWidth      =   6195
      TabIndex        =   18
      Top             =   2640
      Width           =   6195
   End
   Begin VB.PictureBox picCompleted 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   -99999
      Picture         =   "SO_CashAdvanceFrm.frx":19C2D
      ScaleHeight     =   2295
      ScaleWidth      =   6195
      TabIndex        =   19
      Top             =   2640
      Width           =   6195
   End
   Begin VB.CommandButton btnStatus 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Complete Order"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9165
      TabIndex        =   15
      Top             =   6720
      Width           =   2535
   End
   Begin VB.CommandButton btnPaid 
      BackColor       =   &H0080FF80&
      Caption         =   "PAY"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9170
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   7200
      Width           =   2535
   End
   Begin VB.Frame Body_Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   4640
      TabIndex        =   29
      Top             =   0
      Width           =   10455
      Begin MSComctlLib.ListView lvCustomer 
         Height          =   2655
         Left            =   -9999
         TabIndex        =   30
         Top             =   930
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CustomerId"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Balance"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Contact"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Address"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Frame Frame_Body 
         BackColor       =   &H00FFFFFF&
         Height          =   3780
         Left            =   120
         TabIndex        =   56
         Top             =   2685
         Width           =   10215
         Begin VB.TextBox txtInterestRate 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   8400
            TabIndex        =   11
            Top             =   600
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.TextBox txtDays 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   8400
            TabIndex        =   10
            Top             =   240
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.ComboBox cmbAccount 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   2760
            Visible         =   0   'False
            Width           =   5535
         End
         Begin VB.ComboBox cmbBank 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   2400
            Visible         =   0   'False
            Width           =   5535
         End
         Begin VB.TextBox txtCheckAmount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1320
            TabIndex        =   8
            Text            =   "0.00"
            Top             =   960
            Width           =   5535
         End
         Begin VB.TextBox txtCheckNumber 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1320
            TabIndex        =   7
            Top             =   600
            Width           =   5535
         End
         Begin VB.TextBox txtCash 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1320
            TabIndex        =   6
            Text            =   "0.00"
            Top             =   240
            Width           =   5535
         End
         Begin MSComCtl2.DTPicker dtCheckDate 
            Height          =   330
            Left            =   1320
            TabIndex        =   9
            Top             =   1320
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   6619137
            CurrentDate     =   41509
         End
         Begin VB.Label lblPreview 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   6750
            TabIndex        =   68
            Top             =   3240
            Width           =   45
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "[For Check Issuance, deduct from]"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   67
            Top             =   1920
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Int(%)/Mo."
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   7200
            TabIndex        =   65
            Top             =   600
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Days"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   7200
            TabIndex        =   64
            Top             =   240
            Visible         =   0   'False
            Width           =   435
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account #"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   63
            Top             =   2760
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   62
            Top             =   2400
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Check Date"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   61
            Top             =   1320
            Width           =   1035
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Check Amt."
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   60
            Top             =   960
            Width           =   1050
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Check #"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   59
            Top             =   600
            Width           =   705
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cash"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   58
            Top             =   240
            Width           =   435
         End
      End
      Begin VB.Frame Frame_Footer 
         BackColor       =   &H00FFFFFF&
         Height          =   2415
         Left            =   120
         TabIndex        =   46
         Top             =   6480
         Width           =   10215
         Begin VB.TextBox txtRemarks 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2010
            Left            =   1320
            TabIndex        =   14
            Top             =   240
            Width           =   3015
         End
         Begin VB.TextBox txtDiscount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   8400
            TabIndex        =   17
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Particulars"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   66
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lblInterest 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   8400
            TabIndex        =   55
            Top             =   885
            Width           =   1620
         End
         Begin VB.Label lblCaption_AR 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "INTEREST"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   7200
            TabIndex        =   54
            Top             =   885
            Width           =   855
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   8160
            TabIndex        =   53
            Top             =   1200
            Width           =   1860
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7200
            TabIndex        =   52
            Top             =   1200
            Width           =   660
         End
         Begin VB.Label lblSubTotal 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   8400
            TabIndex        =   51
            Top             =   600
            Width           =   1620
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SUB-TOTAL"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   7200
            TabIndex        =   50
            Top             =   600
            Width           =   1035
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "REM. BAL."
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   7200
            TabIndex        =   49
            Top             =   1485
            Width           =   1035
         End
         Begin VB.Label lblRemainingBalance 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   8280
            TabIndex        =   48
            Top             =   1485
            Width           =   1740
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Discount"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7200
            TabIndex        =   47
            Top             =   240
            Width           =   900
         End
      End
      Begin VB.Frame Frame_Header2 
         BackColor       =   &H00FFFFFF&
         Height          =   1755
         Left            =   7170
         TabIndex        =   38
         Top             =   360
         Width           =   3165
         Begin VB.TextBox txtOrderNumber 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   960
            TabIndex        =   5
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox txtStatus 
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   960
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   1320
            Width           =   2055
         End
         Begin MSComCtl2.DTPicker dtOrder 
            Height          =   330
            Left            =   960
            TabIndex        =   40
            Top             =   600
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   6619137
            CurrentDate     =   41509
         End
         Begin MSComCtl2.DTPicker dtDue 
            Height          =   330
            Left            =   960
            TabIndex        =   41
            Top             =   960
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   6619137
            CurrentDate     =   41509
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Invc #"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   45
            Top             =   240
            Width           =   525
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   44
            Top             =   600
            Width           =   435
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   43
            Top             =   1320
            Width           =   570
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Due"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   42
            Top             =   960
            Width           =   375
         End
      End
      Begin VB.Frame Frame_Header1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2055
         Left            =   240
         TabIndex        =   31
         Top             =   600
         Width           =   6735
         Begin VB.ComboBox cmbCustomer 
            BackColor       =   &H00C0C0FF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1200
            TabIndex        =   1
            Top             =   0
            Width           =   3015
         End
         Begin VB.TextBox txtContact 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1200
            TabIndex        =   2
            Top             =   720
            Width           =   3015
         End
         Begin VB.TextBox txtAddress 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   930
            Left            =   1200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   1080
            Width           =   3015
         End
         Begin VB.ComboBox cmbTerms 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            ItemData        =   "SO_CashAdvanceFrm.frx":222BB
            Left            =   5520
            List            =   "SO_CashAdvanceFrm.frx":222BD
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Customer"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   0
            TabIndex        =   37
            Top             =   0
            Width           =   900
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   0
            TabIndex        =   36
            Top             =   720
            Width           =   600
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   0
            TabIndex        =   35
            Top             =   1080
            Width           =   750
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Terms"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   4320
            TabIndex        =   34
            Top             =   0
            Width           =   570
         End
         Begin VB.Label lblBalance 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1200
            TabIndex        =   33
            Top             =   390
            Width           =   3015
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Balance"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   0
            TabIndex        =   32
            Top             =   360
            Width           =   720
         End
      End
      Begin MSComctlLib.Toolbar tb_Standard 
         Height          =   330
         Left            =   0
         TabIndex        =   57
         Top             =   0
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   582
         ButtonWidth     =   1561
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "New"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Save"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cancel"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "Print"
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame LayoutFrame_Search 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   4575
      Begin VB.TextBox txtSearch_OrderNumber 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1440
         TabIndex        =   24
         Top             =   480
         Width           =   3015
      End
      Begin VB.ComboBox cmbSearch_Status 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   840
         Width           =   3015
      End
      Begin VB.CommandButton btnSearch 
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   22
         Top             =   1320
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvSearch 
         Height          =   7095
         Left            =   120
         TabIndex        =   25
         Top             =   1800
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   12515
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "PurchaseOrderId"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Order #"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order #"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   28
         Top             =   480
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   27
         Top             =   80
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   240
         TabIndex        =   26
         Top             =   840
         Width           =   570
      End
   End
   Begin MSComctlLib.ListView lvItemList 
      Height          =   2655
      Left            =   -10080
      TabIndex        =   20
      Top             =   3240
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ProductId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Item Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Uom"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   14325
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_CashAdvanceFrm.frx":222BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_CashAdvanceFrm.frx":28B21
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_CashAdvanceFrm.frx":2F383
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_CashAdvanceFrm.frx":35BE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_CashAdvanceFrm.frx":35E5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_CashAdvanceFrm.frx":364CB
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "SO_CashAdvance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public StatusId, CustomerId, SalesOrderId, id As Integer
Dim TotalSacks As Double
Dim OrderLine(10000) As Integer
Dim ctrOrderLine As Integer

Public Sub Initialize()
    Dim txtControl As Control
    For Each txtControl In Me.Controls
        If TypeOf txtControl Is TextBox And txtControl.Name <> "txtSearch_Order" Then
            txtControl.text = ""
            txtStatus.text = "Open"
        End If
    Next
    
    isNotCompleted (True)
    lblBalance.Caption = "0.00"
    cmbCustomer.text = ""
    cmbCustomer.BackColor = &HC0C0FF
    'cmbTerms.BackColor = &HC0C0FF
    dtOrder.value = Format(Now, "MM/DD/YY")
    dtDue.value = Format(Now, "MM/DD/YY")
    dtCheckDate.value = Format(Now, "MM/DD/YY")
    btnStatus.Caption = "Complete Order"
    lvCustomer.ListItems.Clear
    lvCustomer.Visible = False
    lvItemList.Visible = False
    lvCustomer.Left = -9999
    lvItemList.Left = -9999
    'picCompleted.Left = -9999
    ''pic_Cancelled.Left = -9999
    'picPaid.Left = -9999
    btnPaid.Visible = False
    
    lblPreview.Caption = ""
    lblSubTotal.Caption = "0.00"
    lblTotal.Caption = "0.00"
    lblRemainingBalance.Caption = "0.00"
    
    id = 1
    StatusId = 1
    CustomerId = 0
    SalesOrderId = 0
    TotalSacks = 0
    
    ctrOrderLine = 0
    
    On Error Resume Next
    cmbCustomer.SetFocus
    isModify = False
End Sub
Private Sub Save(ByVal StatusId As Integer, Optional isReopen As Boolean = False)
    If Validated = True Then
        On Error GoTo ErrorHandler
        Set con = New ADODB.Connection
        Set rec = New ADODB.Recordset
        Set cmd = New ADODB.Command
        
        con.ConnectionString = ConnString
        con.Open
        cmd.ActiveConnection = con
        con.BeginTrans
        cmd.CommandType = adCmdStoredProc
        cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInputOutput, , SalesOrderId)
        If Trim(txtOrderNumber.text) = "" Then
            cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInputOutput, 50, Null)
        Else
            cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInputOutput, 50, txtOrderNumber.text)
        End If
        cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtOrder.value)
        cmd.Parameters.Append cmd.CreateParameter("@DueDate", adDate, adParamInput, , dtDue.value)
        cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , StatusId)
        cmd.Parameters.Append cmd.CreateParameter("@TermId", adInteger, adParamInput, , cmbTerms.ItemData(cmbTerms.ListIndex))
        cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInput, , CustomerId)
        cmd.Parameters.Append cmd.CreateParameter("@Days", adDecimal, adParamInput, , Val(Replace(txtDays.text, ",", "")))
                              cmd.Parameters("@Days").Precision = 18
                              cmd.Parameters("@Days").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@InterestRate", adDecimal, adParamInput, , Val(Replace(txtInterestRate.text, ",", "")))
                              cmd.Parameters("@InterestRate").Precision = 18
                              cmd.Parameters("@InterestRate").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@Cash", adDecimal, adParamInput, , Val(Replace(txtCash.text, ",", "")))
                              cmd.Parameters("@Cash").Precision = 18
                              cmd.Parameters("@Cash").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.text, ",", "")))
                              cmd.Parameters("@CheckAmount").Precision = 18
                              cmd.Parameters("@CheckAmount").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.text)
        cmd.Parameters.Append cmd.CreateParameter("@checkDate", adDate, adParamInput, , dtCheckDate.value)
        cmd.Parameters.Append cmd.CreateParameter("@Interest", adDecimal, adParamInput, , Val(Replace(lblInterest.Caption, ",", "")))
                              cmd.Parameters("@Interest").Precision = 18
                              cmd.Parameters("@Interest").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@Subtotal", adDecimal, adParamInput, , Val(Replace(lblSubTotal.Caption, ",", "")))
                              cmd.Parameters("@Subtotal").Precision = 18
                              cmd.Parameters("@Subtotal").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@Total", adDecimal, adParamInput, , Val(Replace(lblTotal.Caption, ",", "")))
                              cmd.Parameters("@Total").Precision = 18
                              cmd.Parameters("@Total").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 500, txtRemarks.text)
        cmd.Parameters.Append cmd.CreateParameter("@Discount", adDecimal, adParamInput, , Val(Replace(txtDiscount.text, ",", "")))
                              cmd.Parameters("@Discount").Precision = 18
                              cmd.Parameters("@Discount").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , 1) 'NOT SET!
        If (Trim(txtCheckNumber.text) <> "" Or Val(Replace(txtCheckAmount.text, ",", "")) > 0) And isReopen = False Then
            cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , cmbAccount.ItemData(cmbAccount.ListIndex))
        Else
            cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , Null)
        End If
        
        If SalesOrderId = 0 Then
            cmd.CommandText = "SO_CashAdvance_Insert"
        Else
            cmd.Parameters.Append cmd.CreateParameter("@isReopen", adBoolean, adParamInput, , isReopen)
            cmd.CommandText = "SO_CashAdvance_Update"
        End If
        
        cmd.Execute
        SalesOrderId = cmd.Parameters("@SalesOrderId")
        txtOrderNumber.text = cmd.Parameters("@OrderNumber")
        
        'INFLOW AND OUTFLOW
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "FIN_OutFlow_Insert"
        cmd.Parameters.Append cmd.CreateParameter("@Cash", adDecimal, adParamInput, , Val(Replace(txtCash.text, ",", "")))
                              cmd.Parameters("@Cash").Precision = 18
                              cmd.Parameters("@Cash").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.text, ",", "")))
                              cmd.Parameters("@CheckAmount").Precision = 18
                              cmd.Parameters("@CheckAmount").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.text)
        cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
        cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , Format(Now, "MM/DD/YY"))
        cmd.Parameters.Append cmd.CreateParameter("@Particulars", adVarChar, adParamInput, 250, "Cash Advance: " & txtOrderNumber.text & " " & txtRemarks.text)
        cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SalesOrderId)
        cmd.Execute
        
        'UPDATE FUND
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "FIN_Funds_Deduct"
        cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , 1) 'DEFAULT VALUE
        cmd.Parameters.Append cmd.CreateParameter("@Cash", adDecimal, adParamInput, , Val(Replace(txtCash.text, ",", "")))
                              cmd.Parameters("@Cash").Precision = 18
                              cmd.Parameters("@Cash").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.text, ",", "")))
                              cmd.Parameters("@CheckAmount").Precision = 18
                              cmd.Parameters("@CheckAmount").NumericScale = 2
        cmd.Execute
        
        'CHECK REGISTRY
        If Trim(txtCheckNumber.text) <> "" Or Val(Replace(txtCheckAmount.text, ",", "")) > 0 Then
            If StatusId = 2 Then
                'Set con = New ADODB.Connection
                Set cmd = New ADODB.Command
                'con.ConnectionString = ConnString
                'con.Open
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "FIN_CheckRegistry_Insert"
                cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.text)
                cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
                cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.text, ",", "")))
                                      cmd.Parameters("@Amount").NumericScale = 2
                                      cmd.Parameters("@Amount").Precision = 18
                cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, txtRemarks.text)
                cmd.Parameters.Append cmd.CreateParameter("@isReceivable", adBoolean, adParamInput, , "False")
                cmd.Parameters.Append cmd.CreateParameter("@CheckStatusId", adInteger, adParamInput, , 1) '--open
                cmd.Parameters.Append cmd.CreateParameter("@CheckRegistryId", adInteger, adParamInputOutput, , Null)
                cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SalesOrderId)
                cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , Null)
                cmd.Parameters.Append cmd.CreateParameter("@ExpenseId", adInteger, adParamInput, , Null)
                cmd.Parameters.Append cmd.CreateParameter("@POS_SalesId", adInteger, adParamInput, , Null)
                cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , cmbAccount.ItemData(cmbAccount.ListIndex))
                
                cmd.Execute
                'con.Close
            End If
        End If
        
        con.CommitTrans
        con.Close
        
        Dim item As MSComctlLib.ListItem
        
        If StatusId = 2 Then
            isNotCompleted (False)
            txtStatus.text = "Complete"
            'picCompleted.Left = 6360
            'picCompleted.Visible = False
            Me.StatusId = 2
        End If
        
        Dim isFound As Boolean
        isFound = False
        For Each item In lvSearch.ListItems
            If SalesOrderId = item.text Then
                item.SubItems(1) = txtOrderNumber.text
                item.SubItems(2) = txtStatus.text
                isFound = True
                item.Selected = True
                item.EnsureVisible
                Exit For
            End If
        Next
        If isFound = False Then
            Set item = lvSearch.ListItems.add(, , SalesOrderId)
                item.SubItems(1) = txtOrderNumber.text
                item.SubItems(2) = txtStatus.text
                item.Selected = True
                item.EnsureVisible
        End If
        
        
        
'        'PRINT PREVIEW
'        If StatusId = 2 Then
'            Screen.MousePointer = vbHourglass
'            BASE_PrintPreviewFrm.Show
'            Dim crxApp As New CRAXDRT.Application
'            Dim crxRpt As New CRAXDRT.Report
'            Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\SO_SalesOrder.rpt")
'            crxRpt.RecordSelectionFormula = "{SO_SalesOrder.SalesOrderId}= " & SalesOrderId & ""
'            crxRpt.DiscardSavedData
'
'            Call ResetRptDB(crxRpt)
'
'            BASE_PrintPreviewFrm.CRViewer.ReportSource = crxRpt
'            BASE_PrintPreviewFrm.CRViewer.ViewReport
'            BASE_PrintPreviewFrm.CRViewer.Zoom 1
'            Screen.MousePointer = vbDefault
'        End If
    End If
    Exit Sub
ErrorHandler:
    con.RollbackTrans
    con.Close
    If IsNumeric(Err.Description) = True Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(Val(Err.Description))
    Else
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & Err.Description
    End If
    GLOBAL_MessageFrm.Show (1)
End Sub

Private Function Validated() As Boolean
    If CustomerId = 0 Then
        Validated = False
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(13)
        GLOBAL_MessageFrm.Show (1)
        cmbCustomer.SetFocus
    ElseIf (Trim(txtCheckNumber.text) <> "" Or Val(Replace(txtCheckAmount.text, ",", ""))) And (cmbAccount.text = "" And Frame_Body.Enabled = True) Then
        Validated = False
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(21)
        GLOBAL_MessageFrm.Show (1)
        cmbAccount.SetFocus

'    ElseIf cmbTerms.ItemData(cmbTerms.ListIndex) <= 0 Then
'        Validated = False
'        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(14)
'        GLOBAL_MessageFrm.Show (1)
'        cmbTerms.SetFocus
    Else
        Validated = True
    End If
End Function

Public Sub isNotCompleted(ByVal a As Boolean)
    Frame_Header1.Enabled = a
    Frame_Header2.Enabled = a
    Frame_Body.Enabled = a
    Frame_Footer.Enabled = a
End Sub
Public Sub CountTotal()
    Dim Total, subtotal, discount, subtotal1, interestrate, Interest, days, cash, checkamount As Double
    Dim item As MSComctlLib.ListItem
    
    If IsNumeric(txtDays.text) = False Then
        days = 0
    Else
        days = Val(Replace(txtDays.text, ",", ""))
    End If
    
    If IsNumeric(txtInterestRate.text) = False Then
        interestrate = 0
    Else
        interestrate = Val(Replace(txtInterestRate.text, ",", ""))
    End If
    
    If IsNumeric(txtCash.text) = False Then
        cash = 0
    Else
        cash = Val(Replace(txtCash.text, ",", ""))
    End If
    
    If IsNumeric(txtCheckAmount.text) = False Then
        checkamount = 0
    Else
        checkamount = Val(Replace(txtCheckAmount.text, ",", ""))
    End If
    
    If IsNumeric(txtDiscount.text) = False Then
        discount = 0
    Else
        discount = Val(Replace(txtDiscount.text, ",", ""))
    End If
    
    subtotal1 = (subtotal1 + cash + checkamount) - discount
    lblSubTotal.Caption = FormatNumber(subtotal1, 2, vbTrue, vbFalse)
    
    'COMPUTE INTEREST
    Interest = ((days / 30) * (subtotal1 * (interestrate / 100)))
    lblInterest.Caption = FormatNumber(Interest, 2, vbTrue, vbFalse)
    
    Total = Interest + subtotal1
    lblTotal.Caption = FormatNumber(Total, 2, vbTrue, vbFalse)
    lblRemainingBalance.Caption = lblTotal.Caption
End Sub
Public Sub Populate(ByVal data As String)
    Select Case data
        Case "Bank"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Bank")
            cmbBank.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        cmbBank.AddItem rec!bankname
                        cmbBank.ItemData(cmbBank.NewIndex) = rec!BankId
                    End If
                    rec.MoveNext
                Loop
            End If
            On Error Resume Next
            cmbBank.ListIndex = 0
        Case "Terms"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Terms")
            cmbTerms.Clear
            cmbTerms.AddItem ""
            cmbTerms.ItemData(cmbTerms.NewIndex) = -1
            cmbTerms.AddItem "<Add New..>"
            cmbTerms.ItemData(cmbTerms.NewIndex) = 0
            If Not rec.EOF Then
                Do Until rec.EOF
                    cmbTerms.AddItem rec!Terms
                    'cmbTerms.Tag = rec!daysdue
                    cmbTerms.ItemData(cmbTerms.NewIndex) = rec!TermId
                    rec.MoveNext
                Loop
            End If
            cmbTerms.ListIndex = 0
            'cmbTerms.BackColor = &HC0C0FF
        Case "Status"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Status")
            cmbSearch_Status.Clear
            cmbSearch_Status.AddItem ""
            cmbSearch_Status.ItemData(cmbSearch_Status.NewIndex) = 0
            cmbSearch_Status.ListIndex = 0
            If Not rec.EOF Then
                Do Until rec.EOF
                    cmbSearch_Status.AddItem rec!Status
                    cmbSearch_Status.ItemData(cmbSearch_Status.NewIndex) = rec!StatusId
                    rec.MoveNext
                Loop
            End If
        Case "Customer"
            Set rec = New ADODB.Recordset
            Set cmd = New ADODB.Command
            Set con = New ADODB.Connection
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "BASE_Customer_Get"
            cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInput, , CustomerId)
            Set rec = cmd.Execute
            If Not rec.EOF Then
                Do Until rec.EOF
                    lblBalance.Caption = FormatNumber(rec!balance, 2, vbTrue, vbFalse)
                    rec.MoveNext
                Loop
            End If
            con.Close
        Case "SalesOrderGet"
            Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            Set rec = New ADODB.Recordset
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "SO_SalesOrder_Get"
            cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SalesOrderId)
            Set rec = cmd.Execute
            If Not rec.EOF Then
                If rec!isCashAdvance = "True" Then
                    CustomerId = rec!CustomerId
                    'lblBalance.Caption = FormatNumber(rec!balance, 2, vbTrue, vbFalse)
                    cmbCustomer.BackColor = vbWhite
                    StatusId = rec!StatusId
                    txtStatus.text = rec!Status
                    txtOrderNumber.text = rec!OrderNumber
                    txtAddress.text = rec!Address
                    txtContact.text = rec!Phone
                    txtRemarks.text = rec!Remarks
                    dtOrder.value = Format(rec!Date, "MM/DD/YY")
                    dtDue.value = Format(rec!DueDate, "MM/DD/YY")
                    lblSubTotal.Caption = FormatNumber(rec!subtotal, 2, vbTrue)
                    txtOrderNumber.text = rec!OrderNumber
                    txtDays.text = rec!days
                    txtInterestRate.text = rec!interestrate
                    txtCash.text = FormatNumber(rec!cash, 2, vbTrue, vbFalse)
                    lblRemainingBalance.Caption = FormatNumber(rec!balance, 2, vbTrue, vbFalse)
                    On Error Resume Next
                    txtCheckNumber.text = rec!CheckNumber
                    txtCheckAmount.text = FormatNumber(rec!checkamount, 2, vbTrue, vbFalse)
                    dtCheckDate.value = Format(rec!CheckDate, "MM/DD/YY")
                    If IsNull(rec!discount) = True Then
                        txtDiscount.text = ""
                    Else
                        txtDiscount.text = FormatNumber(rec!discount, 2, vbTrue, vbFalse)
                    End If
                    
                    If IsNull(rec!Terms) Then
                        cmbTerms.ListIndex = 0
                    Else
                        cmbTerms.text = rec!Terms
                    End If
                    
                    If rec!StatusId = 2 Then 'COMPLETED
                        isNotCompleted (False)
                        'picCompleted.Left = 6360
                        ''pic_Cancelled.Left = -9999
                        'picPaid.Left = -9999
                        'picCompleted.Visible = False
                        ''pic_Cancelled.Visible = False
                        'picPaid.Visible = False
                        tb_Standard.Buttons(4).Caption = "Cancel"
                        tb_Standard.Buttons(4).Image = 3
                        btnPaid.Visible = True
                        btnStatus.Caption = "Reopen Order"
                    ElseIf rec!StatusId = 4 Then 'Cancelled
                        isNotCompleted (False)
                        'picCompleted.Left = -9999
                        ''pic_Cancelled.Left = 6360
                        'picPaid.Left = -9999
                        'picPaid.Visible = False
                        'picCompleted.Visible = False
                        ''pic_Cancelled.Visible = True
                        tb_Standard.Buttons(4).Caption = "Activate"
                        tb_Standard.Buttons(4).Image = 6
                        btnPaid.Visible = False
                        btnStatus.Caption = "Complete Order"
                    ElseIf rec!StatusId = 3 Then 'PAID
                        isNotCompleted (False)
                        'picPaid.Left = 6890
                        'picCompleted.Left = -9999
                        ''pic_Cancelled.Left = -9999
                        'picPaid.Visible = True
                        'picCompleted.Visible = False
                        ''pic_Cancelled.Visible = False
                        tb_Standard.Buttons(4).Caption = "Cancel"
                        tb_Standard.Buttons(4).Image = 3
                        btnPaid.Visible = False
                        btnStatus.Caption = "Reopen Order"
                    Else
                        isNotCompleted (True)
                        'picCompleted.Left = -9999
                        ''pic_Cancelled.Left = -9999
                        'picPaid.Left = -9999
                        'picCompleted.Visible = False
                        ''pic_Cancelled.Visible = False
                        'picPaid.Visible = False
                        tb_Standard.Buttons(4).Caption = "Cancel"
                        tb_Standard.Buttons(4).Image = 3
                        btnPaid.Visible = False
                        btnStatus.Caption = "Complete Order"
                    End If
                    lblRemainingBalance.Caption = FormatNumber(rec!balance, 2, vbTrue, vbFalse)
                    'PREVENT ERROR ON DISPLAY FOR cmbCustomer_Change Event
                    On Error Resume Next
                    cmbCustomer.text = rec!Name
                    lvCustomer.Left = -9999
                    lvCustomer.Visible = False
                End If
            End If
            con.Close
        Case "SalesOrderLoad"
            Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            Set rec = New ADODB.Recordset
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "SO_SalesOrder_Get"
            Set rec = cmd.Execute
            Dim item As MSComctlLib.ListItem
            lvSearch.ListItems.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isCashAdvance = "True" Then
                        Set item = lvSearch.ListItems.add(, , rec!SalesOrderId)
                            item.SubItems(1) = rec!OrderNumber
                            item.SubItems(2) = rec!Status
                    End If
                    rec.MoveNext
                Loop
            End If
            con.Close
    End Select
End Sub


Private Sub btnItemSearch_Click()
'    Set con = New ADODB.Connection
'    Set rec = New ADODB.Recordset
'    Set cmd = New ADODB.Command
'    Dim Item As MSComctlLib.ListItem
'
'    con.ConnectionString = ConnString
'    con.Open
'    cmd.ActiveConnection = con
'    cmd.CommandType = adCmdStoredProc
'    cmd.CommandText = "BASE_Product_Search"
'    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 500, txtItemSearch.text)
'
'    Set rec = cmd.Execute
'    If Not rec.EOF Then
'        lvItemList.ListItems.Clear
'        Do Until rec.EOF
'            If rec!isActive = "True" Then
'                Set Item = lvItemList.ListItems.add(, , rec!ProductId)
'                    Item.SubItems(1) = rec!Itemcode
'                    Item.SubItems(2) = rec!Name
'                    Item.SubItems(3) = FormatNumber(rec!UnitPrice, 2, vbTrue, vbFalse)
'                    Item.SubItems(4) = rec!Uom
'                lvItemList.Visible = True
'                lvItemList.Left = 6070
'            End If
'            rec.MoveNext
'        Loop
'    Else
'        lvItemList.Visible = False
'        lvItemList.Left = -9999
'    End If
'    'DistinctList lvItemList
'    con.Close
End Sub

Private Sub btnPaid_Click()
    Unload FIN_PaymentDialogFrm
    With FIN_PaymentDialogFrm
        .isBatchPO = False
        .isBatchSO = False
        .isPO = False
        .isSO = True
        .isSOAR = False
        .chkOnline.Visible = True
        .lblCustomerName.Caption = cmbCustomer.text
        .lblOrderNumber.Caption = txtOrderNumber.text
        .lblBalance.Caption = lblRemainingBalance.Caption
        .lblBalanceForwarded.Caption = lblRemainingBalance.Caption
        .Show (1)
    End With
End Sub

Private Sub btnSearch_Click()
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "SO_SalesOrder_Get"
    cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , Null)
    If cmbSearch_Status.ItemData(cmbSearch_Status.ListIndex) = 0 Then
        cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , Null)
    Else
        cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , cmbSearch_Status.ItemData(cmbSearch_Status.ListIndex))
    End If
    cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInput, 50, txtSearch_OrderNumber.text)
    Dim item As MSComctlLib.ListItem
    Set rec = cmd.Execute
    lvSearch.ListItems.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            If rec!isCashAdvance = "True" Then
                Set item = lvSearch.ListItems.add(, , rec!SalesOrderId)
                    item.SubItems(1) = rec!OrderNumber
                    item.SubItems(2) = rec!Status
            End If
            rec.MoveNext
        Loop
    End If
    con.Close
End Sub

Private Sub btnStatus_Click()
    If StatusId = 1 Then
        Save (2)
        If Me.StatusId = 2 Then
            btnStatus.Caption = "Reopen Order"
            btnPaid.Visible = True
        End If
    ElseIf StatusId = 2 Or StatusId = 3 Then
        Dim x As Variant
        x = MsgBox("Are you sure you want to reopen this order? This will revert all connected " & _
                    "transactions such as payments and inventory.", vbCritical + vbYesNo, "WARNING")
        If x = vbYes Then
            Save 1, True
            btnStatus.Caption = "Complete Order"
            txtStatus.text = "Open"
            btnPaid.Visible = False
            isNotCompleted (True)
            'picCompleted.Visible = False
            'picCompleted.Left = -9999
            'picPaid.Left = -9999
            'picPaid.Visible = False
            Me.StatusId = 1
        Else
        End If
    End If
End Sub

Private Sub cmbBank_Click()
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_BankAccount_Load"
    
    cmd.Parameters.Append cmd.CreateParameter("@BankId", adInteger, adParamInput, , cmbBank.ItemData(cmbBank.ListIndex))
    Set rec = cmd.Execute
    cmbAccount.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            If rec!isActive = "True" Then
                cmbAccount.AddItem rec!accountnumber & " - " & rec!AccountName
                cmbAccount.ItemData(cmbAccount.NewIndex) = rec!AccountId
            End If
            rec.MoveNext
        Loop
    End If
    con.Close
End Sub

Private Sub cmbCustomer_Change()
    If Trim(cmbCustomer.text) <> "" Then
        Dim item As MSComctlLib.ListItem
        Set con = New ADODB.Connection
        Set rec = New ADODB.Recordset
        Set cmd = New ADODB.Command
        'Dim item As MSComctlLib.ListItem
    
        con.ConnectionString = ConnString
        con.Open
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "BASE_Customer_Search"
        cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, Trim(cmbCustomer.text))
        Set rec = cmd.Execute
        If Not rec.EOF Then
            lvCustomer.ListItems.Clear
            lvCustomer.Left = 1440
            lvCustomer.Visible = True
            Do Until rec.EOF
                If rec!isActive = "True" Then
                    Set item = lvCustomer.ListItems.add(, , rec!CustomerId)
                        item.SubItems(1) = rec!CustomerCode
                        item.SubItems(2) = rec!Name
                        item.SubItems(3) = FormatNumber(rec!balance, 2, vbTrue, vbFalse)
                        item.SubItems(4) = rec!Phone
                        item.SubItems(5) = rec!Address
                End If
                rec.MoveNext
            Loop
        Else
            lvCustomer.Visible = False
            lvCustomer.Left = -9999
        End If
        con.Close
    
'        If Trim(cmbCustomer.text) = "" Then
'            cmbCustomer.BackColor = &HC0C0FF
'        Else
'            cmbCustomer.BackColor = vbWhite
'        End If
    End If
End Sub

Private Sub cmbCustomer_GotFocus()
    selectText cmbCustomer
End Sub

Private Sub cmbCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Set con = New ADODB.Connection
            Set rec = New ADODB.Recordset
            Set cmd = New ADODB.Command
            Dim item As MSComctlLib.ListItem
            
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "BASE_Customer_Search"
            cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 500, "")
            cmd.Parameters.Append cmd.CreateParameter("@CustomerCode", adVarChar, adParamInput, 50, cmbCustomer.text)
            Set rec = cmd.Execute
            If Not rec.EOF Then
                lvCustomer.ListItems.Clear
                lvCustomer.Left = 1440
                lvCustomer.Visible = True
                Do Until rec.EOF
                    Set item = lvCustomer.ListItems.add(, , rec!CustomerId)
                        item.SubItems(1) = rec!CustomerCode
                        item.SubItems(2) = rec!Name
                        item.SubItems(3) = FormatNumber(rec!balance, 2, vbTrue, vbFalse)
                        item.SubItems(4) = rec!Phone
                        item.SubItems(5) = rec!Address
                    rec.MoveNext
                Loop
            Else
                lvCustomer.Visible = False
                lvCustomer.Left = -9999
            End If
            con.Close
        Case vbKeyUp, vbKeyDown
            If lvCustomer.Visible = True Then
                lvCustomer.SetFocus
            End If
    End Select
End Sub

Private Sub cmbTerms_Click()
'    If cmbTerms.ListIndex > 1 Then
'        txtDays.text = cmbTerms.Tag
'    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            If lvCustomer.Visible = True Then
                lvCustomer.Visible = False
                lvCustomer.Left = -9999
                cmbCustomer.SetFocus
'            ElseIf lvItemList.Visible = True Then
'                lvItemList.Visible = False
'                lvItemList.Left = -9999
'                txtItemSearch.SetFocus
            End If
    End Select
End Sub

Private Sub Form_Load()
    Initialize
    Populate "Terms"
    Populate "Status"
    Populate "Bank"
    'Populate "SalesOrderLoad"
    
    On Error Resume Next
    cmbSearch_Status.ListIndex = 2
    btnSearch_Click
    
    lvSearch.ColumnHeaders(2).width = lvSearch.width * 0.48
    lvSearch.ColumnHeaders(3).width = lvSearch.width * 0.45
    
    lvItemList.ColumnHeaders(2).width = lvItemList.width * 0.35
    lvItemList.ColumnHeaders(3).width = lvItemList.width * 0.4
    lvItemList.ColumnHeaders(4).width = lvItemList.width * 0.2
    
    lvCustomer.ColumnHeaders(2).width = lvCustomer.width * 0.25
    lvCustomer.ColumnHeaders(3).width = lvCustomer.width * 0.42
    lvCustomer.ColumnHeaders(4).width = lvCustomer.width * 0.28
End Sub





Private Sub lblGrossAmount_Click()

End Sub

Private Sub lblGrossKilos_Click()

End Sub

Private Sub lvCustomer_DblClick()
    If lvCustomer.ListItems.Count > 0 Then
        CustomerId = lvCustomer.SelectedItem.text
        cmbCustomer.text = lvCustomer.SelectedItem.SubItems(2)
        lblBalance.Caption = lvCustomer.SelectedItem.SubItems(3)
        txtContact.text = lvCustomer.SelectedItem.SubItems(4)
        txtAddress.text = lvCustomer.SelectedItem.SubItems(5)
        lvCustomer.Visible = False
        lvCustomer.Left = -9999
        txtCash.SetFocus
        'txtItemSearch.SetFocus
        cmbCustomer.BackColor = vbWhite
    End If
End Sub

Private Sub lvCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            lvCustomer_DblClick
    End Select
End Sub

Private Sub lvItemList_DblClick()
    SO_ItemSelectOptionFrm.txtPrice.text = lvItemList.SelectedItem.SubItems(3)
    SO_ItemSelectOptionFrm.Show (1)
End Sub

Private Sub lvItemList_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Not lvItemList.SelectedItem = -1 Then
                SO_ItemSelectOptionFrm.txtPrice.text = lvItemList.SelectedItem.SubItems(3)
                SO_ItemSelectOptionFrm.Show (1)
            End If
    End Select
End Sub




Private Sub lvSearch_ItemClick(ByVal item As MSComctlLib.ListItem)
    If lvSearch.ListItems.Count > 0 Then
        SalesOrderId = lvSearch.SelectedItem.text
        Populate "SalesOrderLineGet"
        CountTotal
        Populate "SalesOrderGet"
        Populate "Customer"
    End If
End Sub



Private Sub tb_Standard_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 ' New
            Initialize
        Case 2 'Save
            If StatusId <= 1 Then
                Save (1) 'Status Open
            End If
        Case 4 'Cancel
            If SalesOrderId <> 0 Then
                If StatusId = 1 Then 'Status Open > Cancelled
                    StatusId = 4
                    txtStatus.text = "Cancelled"
                    Save (4)
                    isNotCompleted (False)
                    ''pic_Cancelled.Left = 6360
                    ''pic_Cancelled.Visible = True
                    tb_Standard.Buttons(4).Caption = "Activate"
                    tb_Standard.Buttons(4).Image = 6
                ElseIf StatusId = 4 Then 'Status Cancelled > Activate
                    StatusId = 1
                    txtStatus.text = "Open"
                    Save (1)
                    isNotCompleted (True)
                    ''pic_Cancelled.Left = -9999
                    ''pic_Cancelled.Visible = False
                    tb_Standard.Buttons(4).Caption = "Cancel"
                    tb_Standard.Buttons(4).Image = 3
                End If
            End If
        Case 6 'PRINT PREVIEW
            If SalesOrderId <> 0 Then
                Screen.MousePointer = vbHourglass
                BASE_PrintPreviewFrm.Show
                Dim crxApp As New CRAXDRT.Application
                Dim crxRpt As New CRAXDRT.Report
                Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\SO_SalesOrder.rpt")
                crxRpt.RecordSelectionFormula = "{SO_SalesOrder.SalesOrderId}= " & SalesOrderId & ""
                crxRpt.DiscardSavedData
                
                Call ResetRptDB(crxRpt)
                
                BASE_PrintPreviewFrm.CRViewer.ReportSource = crxRpt
                BASE_PrintPreviewFrm.CRViewer.ViewReport
                BASE_PrintPreviewFrm.CRViewer.Zoom 1
                Screen.MousePointer = vbDefault
            End If
    End Select
End Sub

Private Sub txtCash_Change()
    If IsNumeric(txtCash.text) = False Then
        txtCash.text = "0.00"
        selectText txtCash
    Else
'        txtCash.text = FormatNumber(txtCash.text, 0)
'        txtCash.SelStart = Len(txtCash.text)
        lblPreview.Caption = FormatNumber(txtCash.text, 2, vbTrue)
        CountTotal
    End If
End Sub

Private Sub txtCash_GotFocus()
    selectText txtCash
End Sub

Private Sub txtCash_LostFocus()
    lblPreview.Caption = ""
    txtCash.text = FormatNumber(txtCash.text, 2, vbTrue)
End Sub

Private Sub txtCheckAmount_Change()
    If IsNumeric(txtCheckAmount.text) = False Then
        txtCheckAmount.text = "0.00"
        selectText txtCheckAmount
    Else
'        txtCheckAmount.text = FormatNumber(txtCheckAmount.text, 0)
'        txtCheckAmount.SelStart = Len(txtCheckAmount.text)
        lblPreview.Caption = FormatNumber(txtCheckAmount.text, 2, vbTrue)
        CountTotal
    End If
End Sub

Private Sub txtCheckAmount_LostFocus()
    lblPreview.Caption = ""
    txtCheckAmount.text = FormatNumber(txtCheckAmount.text, 2, vbTrue)
End Sub

Private Sub txtDays_Change()
    If IsNumeric(txtDays.text) = False Then
        txtDays.text = ""
    End If
    CountTotal
End Sub

Private Sub txtDays_GotFocus()
    selectText txtDays
End Sub

Private Sub txtDiscount_Change()
    If IsNumeric(txtDiscount.text) = False Then
        txtDiscount.text = ""
    End If
    CountTotal
End Sub

Private Sub txtDiscount_GotFocus()
    selectText txtDiscount
End Sub

Private Sub txtInterestRate_Change()
    If IsNumeric(txtInterestRate.text) = False Then
        txtInterestRate.text = ""
    End If
    CountTotal
End Sub

Private Sub txtInterestRate_GotFocus()
    selectText txtInterestRate
End Sub

Private Sub txtItemSearch_Change()
    'btnItemSearch_Click
End Sub

Private Sub txtItemSearch_GotFocus()
    'selectText txtItemSearch
End Sub



Private Sub txtSearch_OrderNumber_Change()
    btnSearch_Click
End Sub

Private Sub txtSearch_OrderNumber_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = 13 Then btnSearch_Click
End Sub


