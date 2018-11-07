VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form INV_AuditStockFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   15090
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picStatus 
      BorderStyle     =   0  'None
      Height          =   1055
      Left            =   7680
      Picture         =   "INV_AuditStockFrm.frx":0000
      ScaleHeight     =   1050
      ScaleWidth      =   3750
      TabIndex        =   47
      Top             =   1800
      Width           =   3755
   End
   Begin VB.PictureBox picCompleted 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   -9999
      Picture         =   "INV_AuditStockFrm.frx":7D72
      ScaleHeight     =   2295
      ScaleWidth      =   6195
      TabIndex        =   1
      Top             =   2520
      Width           =   6195
   End
   Begin VB.PictureBox pic_Cancelled 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   -9999
      Picture         =   "INV_AuditStockFrm.frx":10400
      ScaleHeight     =   2295
      ScaleWidth      =   6195
      TabIndex        =   0
      Top             =   2520
      Width           =   6195
   End
   Begin VB.Frame LayoutFrame_Search 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   3615
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
         Left            =   2280
         TabIndex        =   29
         Top             =   2880
         Width           =   1215
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
         TabIndex        =   28
         Top             =   840
         Width           =   2055
      End
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
         TabIndex        =   27
         Top             =   480
         Width           =   2055
      End
      Begin MSComctlLib.ListView lvSearch 
         Height          =   5535
         Left            =   120
         TabIndex        =   30
         Top             =   3360
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   9763
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
      Begin MSComCtl2.DTPicker DateTo 
         Height          =   345
         Left            =   1440
         TabIndex        =   31
         Top             =   2040
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   88604673
         CurrentDate     =   41686
      End
      Begin MSComCtl2.DTPicker DateFrom 
         Height          =   345
         Left            =   1440
         TabIndex        =   32
         Top             =   1200
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   88604673
         CurrentDate     =   41686
      End
      Begin MSComCtl2.DTPicker TimeFrom 
         Height          =   345
         Left            =   1440
         TabIndex        =   43
         Top             =   1560
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   88604674
         UpDown          =   -1  'True
         CurrentDate     =   41686
      End
      Begin MSComCtl2.DTPicker TimeTo 
         Height          =   345
         Left            =   1440
         TabIndex        =   45
         Top             =   2400
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   88604674
         UpDown          =   -1  'True
         CurrentDate     =   41686
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time To"
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
         TabIndex        =   46
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time From"
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
         TabIndex        =   44
         Top             =   1560
         Width           =   990
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date From"
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
         TabIndex        =   37
         Top             =   1200
         Width           =   960
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date To"
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
         TabIndex        =   36
         Top             =   2040
         Width           =   705
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
         TabIndex        =   35
         Top             =   840
         Width           =   570
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
         TabIndex        =   34
         Top             =   80
         Width           =   795
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Audit #"
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
         TabIndex        =   33
         Top             =   480
         Width           =   660
      End
   End
   Begin VB.Frame Body_Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   3600
      TabIndex        =   2
      Top             =   0
      Width           =   11535
      Begin VB.CommandButton btnComplete 
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
         Height          =   355
         Left            =   240
         TabIndex        =   3
         Top             =   1800
         Width           =   2895
      End
      Begin VB.Frame Frame_Header1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   735
         Left            =   120
         TabIndex        =   19
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
            Left            =   -9999
            TabIndex        =   21
            Top             =   0
            Visible         =   0   'False
            Width           =   4695
         End
         Begin VB.TextBox txtSalesman 
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
            Left            =   -9999
            TabIndex        =   20
            Top             =   360
            Visible         =   0   'False
            Width           =   4695
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
            Left            =   -9999
            TabIndex        =   24
            Top             =   0
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Salesman"
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
            Left            =   -9999
            TabIndex        =   23
            Top             =   360
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Stock Audit"
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
            Left            =   120
            TabIndex        =   22
            Top             =   120
            Width           =   1845
         End
      End
      Begin VB.Frame Frame_Body 
         BackColor       =   &H00FFFFFF&
         Height          =   5700
         Left            =   120
         TabIndex        =   17
         Top             =   2160
         Width           =   11295
         Begin VB.CommandButton btnAdd 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Add Products to Audit"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   355
            Left            =   8160
            TabIndex        =   42
            Top             =   240
            Width           =   3015
         End
         Begin VB.ComboBox cmbAuditLocation 
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
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   240
            Width           =   3855
         End
         Begin MSComctlLib.ListView lvItems 
            Height          =   4815
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   8493
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
            NumItems        =   9
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ProductId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "AuditStockId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "AuditStockLineId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Code"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Unit"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "System Qty"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Text            =   "Physical Qty"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   8
               Text            =   "Adjustment"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Audit Location"
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
            TabIndex        =   41
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame_Footer 
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   120
         TabIndex        =   12
         Top             =   7800
         Width           =   11295
         Begin VB.TextBox txtRemarks 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   1200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   240
            Width           =   4095
         End
         Begin VB.TextBox txtCash 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            Height          =   390
            Left            =   -9999
            TabIndex        =   13
            Top             =   240
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
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
            TabIndex        =   16
            Top             =   240
            Width           =   795
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CASH"
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
            Left            =   -9999
            TabIndex        =   15
            Top             =   240
            Visible         =   0   'False
            Width           =   540
         End
      End
      Begin VB.Frame Frame_Header2 
         BackColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   8250
         TabIndex        =   5
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
            TabIndex        =   7
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
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   1320
            Width           =   2055
         End
         Begin MSComCtl2.DTPicker dtOrder 
            Height          =   330
            Left            =   960
            TabIndex        =   8
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
            Format          =   88604673
            CurrentDate     =   41509
         End
         Begin MSComCtl2.DTPicker dtTime 
            Height          =   330
            Left            =   960
            TabIndex        =   38
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
            Format          =   88604674
            UpDown          =   -1  'True
            CurrentDate     =   41509
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Time"
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
            TabIndex        =   39
            Top             =   960
            Width           =   465
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Audit #"
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
            TabIndex        =   11
            Top             =   240
            Width           =   660
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
            TabIndex        =   10
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
            TabIndex        =   9
            Top             =   1320
            Width           =   570
         End
      End
      Begin MSComctlLib.Toolbar tb_Standard 
         Height          =   330
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   582
         ButtonWidth     =   1667
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
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cancel"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Print"
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvCustomer 
         Height          =   2655
         Left            =   -9999
         TabIndex        =   25
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
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   14400
      Top             =   120
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
            Picture         =   "INV_AuditStockFrm.frx":2109C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_AuditStockFrm.frx":278FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_AuditStockFrm.frx":2E160
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_AuditStockFrm.frx":349C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_AuditStockFrm.frx":34C37
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_AuditStockFrm.frx":352A8
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "INV_AuditStockFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AuditStockId As Long
Dim StatusId As Integer
Dim item As MSComctlLib.ListItem
Private Sub Initialize()
    cmbAuditLocation.Enabled = True
    picStatus.Visible = False
    dtOrder.value = Format(Now, "MM/DD/YY")
    'picCompleted.Visible = False
    ''pic_Cancelled.Visible = False
    
    StatusId = 0
    lvItems.ListItems.Clear
    txtStatus.Text = ""
    txtRemarks.Text = ""
    txtOrderNumber.Text = ""
'    lvItemList.Visible = False
    AuditStockId = 0
    
    On Error Resume Next
    txtItemSearch.SetFocus
End Sub
Private Sub ComputeAdjustment()
    For Each item In lvItems.ListItems
        If item.SubItems(7) <> "" Then
            item.SubItems(8) = NVAL(item.SubItems(7)) - NVAL(item.SubItems(6))
            item.SubItems(8) = FormatNumber(item.SubItems(8))
        End If
    Next
End Sub
Public Sub CountTotal()
'    Dim item As MSComctlLib.ListItem
'    Dim total As Double
'
'    For Each item In lvItems.ListItems
'        item.SubItems(11) = FormatNumber(NVAL(item.SubItems(6)) * NVAL(item.SubItems(10)), 2, vbTrue, vbFalse)
'        total = total + NVAL(item.SubItems(11))
'    Next
'    lblTotal.Caption = FormatNumber(total, 2, vbTrue, vbFalse)
End Sub

Private Sub btnAdd_Click()
    INV_ProductSelectionFrm.Show (1)
    DisplayProducts
    GetSystemQuantity
End Sub
Private Sub GetSystemQuantity()
    Dim con As New ADODB.Connection
    con.ConnectionString = ConnString
    con.Open
    
    For Each item In lvItems.ListItems
        Set cmd = New ADODB.Command
        Set rec = New ADODB.Recordset
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "INV_ProductQuantity_Get"
        cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , Val(item.Text))
        cmd.Parameters.Append cmd.CreateParameter("LocationId", adInteger, adParamInput, , cmbAuditLocation.ItemData(cmbAuditLocation.ListIndex))
        Set rec = cmd.Execute
        If Not rec.EOF Then
            Do Until rec.EOF
                item.SubItems(6) = FormatNumber(rec!quantity, 2, vbTrue, vbFalse)
                rec.MoveNext
            Loop
        End If
    Next
    con.Close
End Sub
Public Sub DisplayProducts()
    On Error Resume Next
    
    Dim exists As Boolean
    Dim item As MSComctlLib.ListItem
    
    If ProductSet.RecordCount <= 0 Then Exit Sub

    'Dim item As MSComctlLib.ListItem
    If Not ProductSet.EOF Then
        ProductSet.MoveFirst
        Do Until ProductSet.EOF
            For Each item In lvItems.ListItems
                If item.Text = ProductSet!ProductId Then
                    exists = True
                    Exit For
                End If
            Next

            If exists = False Then
                Set item = lvItems.ListItems.add(, , ProductSet!ProductId)
                item.SubItems(3) = ProductSet!itemcode
                item.SubItems(4) = ProductSet!Name
                item.SubItems(5) = ProductSet!unit
            End If
            ProductSet.MoveNext
        Loop
    End If
    cmbAuditLocation.Enabled = False
End Sub

Private Sub btnComplete_Click()
    If EditAccessRights(28) = False Then
        MsgBox ErrorCodes(74), vbCritical, "Limited Rights"
        Exit Sub
    End If
    If StatusId = 2 Then
        MsgBox "Save failed. No changes made. Order is already complete.", vbCritical
        Exit Sub
    End If
    Dim x As Variant
    x = MsgBox("This will complete the transaction. Product inventories will now be updated. Proceed?", vbExclamation + vbOKCancel)
    If x = vbOK Then
        For Each item In lvItems.ListItems
            If item.SubItems(7) = "" Then item.SubItems(7) = item.SubItems(6)
        Next
        ComputeAdjustment
        Save (1)
        Save (2)
        LoadImageStatus picStatus, GetStatus(StatusId)
    End If
End Sub

Private Sub btnItemSearch_Click()
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    Dim item As MSComctlLib.ListItem
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "POS_ItemSearch_Name"
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 500, txtItemSearch.Text)
    
    Set rec = cmd.Execute
    lvItemList.ListItems.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            If rec!isActive = "True" Then
                Set item = lvItemList.ListItems.add(, , rec!ProductId)
                    item.SubItems(1) = rec!itemcode
                    item.SubItems(2) = rec!Name
                    item.SubItems(3) = rec!Uom
                    'Item.SubItems(4) = FormatNumber(rec!totalQty, 2, vbTrue, vbFalse)
                lvItemList.Visible = True
                lvItemList.Left = 5280
                lvItemList.Top = 3090
                LastProductId = rec!ProductId
                rec.MoveNext
            Else
                rec.MoveNext
            End If
        Loop
    Else
        lvItemList.Visible = False
        lvItemList.Left = -9999
    End If
    'DistinctList lvItemList
    con.Close
End Sub

Private Sub btnSearch_Click()
    Dim con As New ADODB.Connection
    con.ConnectionString = ConnString
    con.Open
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    cmd.ActiveConnection = con
    cmd.CommandText = "INV_AuditStock_Search"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInput, 50, txtSearch_OrderNumber.Text)
    cmd.Parameters.Append cmd.CreateParameter("@DateFrom", adDate, adParamInput, , DateFrom.value)
    cmd.Parameters.Append cmd.CreateParameter("@TimeFrom", adVarChar, adParamInput, 50, Str(TimeFrom.value))
    cmd.Parameters.Append cmd.CreateParameter("@DateTo", adDate, adParamInput, , DateTo.value)
    cmd.Parameters.Append cmd.CreateParameter("@TimeTo", adVarChar, adParamInput, 50, Str(TimeTo.value))
    If cmbSearch_Status.ListIndex = 0 Then
        cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , Null)
    Else
        cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , cmbSearch_Status.ItemData(cmbSearch_Status.ListIndex))
    End If
    Set rec = cmd.Execute
    lvSearch.ListItems.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            Set item = lvSearch.ListItems.add(, , rec!AuditStockId)
                item.SubItems(1) = rec!OrderNumber
                item.SubItems(2) = rec!Status
            rec.MoveNext
        Loop
    End If
    con.Close
End Sub

Private Sub cmbAuditLocation_Click()
    GetSystemQuantity
End Sub

Private Sub cmbSearch_Status_Click()
    'btnSearch_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case vbKeyEscape
            If lvItemList.Visible = True Then
                lvItemList.Visible = False
                lvItemList.Left = -9999
                txtItemSearch.SetFocus
            End If
        Case vbKeyN
            If Shift = vbCtrlMask Then
                tb_Standard_ButtonClick tb_Standard.Buttons(1)
            End If
        Case vbKeyS
            If Shift = vbCtrlMask Then
                tb_Standard_ButtonClick tb_Standard.Buttons(2)
            End If
        Case vbKeyO
            If Shift = vbCtrlMask Then
                tb_Standard_ButtonClick tb_Standard.Buttons(4)
            End If
        Case vbKeyP
            If Shift = vbCtrlMask Then
                tb_Standard_ButtonClick tb_Standard.Buttons(6)
            End If
        Case vbKeyF4
            txtItemSearch.SetFocus
    End Select
End Sub

Private Sub Form_Load()
    lvItems.ColumnHeaders(4).width = lvItems.width * 0.1383
    lvItems.ColumnHeaders(5).width = lvItems.width * 0.3183
    lvItems.ColumnHeaders(6).width = lvItems.width * 0.0983
    lvItems.ColumnHeaders(7).width = lvItems.width * 0.1383
    lvItems.ColumnHeaders(8).width = lvItems.width * 0.1383
    lvItems.ColumnHeaders(8).width = lvItems.width * 0.1383

    
    lvSearch.ColumnHeaders(2).width = lvSearch.width * 0.48
    lvSearch.ColumnHeaders(3).width = lvSearch.width * 0.45
    
'    lvItemList.ColumnHeaders(2).width = lvItemList.width * 0.285
'    lvItemList.ColumnHeaders(3).width = lvItemList.width * 0.685
    
    DateFrom.value = Format(Now, "mm/dd/yy")
    DateTo.value = Format(Now, "mm/dd/yy")
    
    dtTime.value = Format("00:00:00", "hh:mm:ss")
    TimeFrom.value = Format("00:00:00", "hh:mm:ss")
    TimeTo.value = Format("23:59:59", "hh:mm:ss")
    
    Populate "Status"
    Populate "AuditStockLoad"
    Populate "Location"
    Initialize
    
    cmbAuditLocation.ListIndex = 0
End Sub

Private Sub lvItemList_DblClick()
    INV_AuditStockOptFrm.Show (1)
End Sub

Private Sub lvItemList_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            lvItemList_DblClick
    End Select
End Sub

Private Sub lvItems_DblClick()
    If lvItems.ListItems.Count <= 0 Then Exit Sub
    INV_AuditStockOptFrm.txtSystem.Text = lvItems.SelectedItem.SubItems(6)
    INV_AuditStockOptFrm.txtPhysical.Text = lvItems.SelectedItem.SubItems(6)
    selectText INV_AuditStockOptFrm.txtPhysical
    INV_AuditStockOptFrm.Show (1)
End Sub

Private Sub lvItems_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            lvItems_DblClick
        Case vbKeyDelete
            Dim x As Variant
            x = MsgBox("This will permanently delete the selected record. Proceed?", vbExclamation + vbOKCancel)
            If x = vbOK Then
                If NVAL(lvItems.SelectedItem.SubItems(2)) <> 0 Then
                    Dim con As New ADODB.Connection
                    con.ConnectionString = ConnString
                    con.Open
                    Set cmd = New ADODB.Command
                    cmd.ActiveConnection = con
                    cmd.CommandType = adCmdStoredProc
                    cmd.CommandText = "INV_AuditStockLine_Delete"
                    cmd.Parameters.Append cmd.CreateParameter("@AuditStockLineId", adInteger, adParamInput, , NVAL(lvItems.SelectedItem.SubItems(2)))
                    cmd.Execute
                    con.Close
                    
                    SavePOSAuditTrail UserId, WorkstationId, "", "Removed item: " & lvItems.SelectedItem.SubItems(4) & " from audit stock order: " & txtOrderNumber.Text
                End If
                lvItems.ListItems.Remove (lvItems.SelectedItem.Index)
                MsgBox "Item(s) removed.", vbInformation
            End If
            If lvItems.ListItems.Count <= 0 Then cmbAuditLocation.Enabled = True
    End Select
End Sub
Private Sub SelectOrders()
    Initialize
    
    Dim con As New ADODB.Connection
    con.ConnectionString = ConnString
    con.Open
    
    'Get AuditStock
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "INV_AuditStock_Get"
    cmd.ActiveConnection = con
    cmd.Parameters.Append cmd.CreateParameter("@AuditStockId", adInteger, adParamInput, , lvSearch.SelectedItem.Text)
    
    Set rec = cmd.Execute
    If Not rec.EOF Then
        txtOrderNumber.Text = rec!OrderNumber
        dtOrder.value = rec!Date
        dtTime.value = TimeValue(rec!Time)
        txtRemarks.Text = rec!Remarks
        StatusId = rec!StatusId
        txtStatus.Text = rec!Status
        AuditStockId = rec!AuditStockId
    End If
    
    'Get Line
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandText = "INV_AuditStockLine_Get"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("@AuditStockId", adInteger, adParamInput, , lvSearch.SelectedItem.Text)
    Set rec = cmd.Execute
    
    lvItems.ListItems.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            Set item = lvItems.ListItems.add(, , rec!ProductId)
                item.SubItems(1) = rec!AuditStockId
                item.SubItems(2) = rec!AuditStocklineid
                item.SubItems(3) = rec!itemcode
                item.SubItems(4) = rec!Name
                item.SubItems(5) = rec!unit
                item.SubItems(6) = FormatNumber(rec!systemquantity, 2, vbTrue, vbFalse)
                If IsNull(rec!physicalquantity) Then
                    item.SubItems(7) = ""
                Else
                    item.SubItems(7) = FormatNumber(rec!physicalquantity, 2, vbTrue, vbFalse)
                End If
                If IsNull(rec!adjustment) Then
                    item.SubItems(8) = ""
                Else
                    item.SubItems(8) = FormatNumber(rec!adjustment, 2, vbTrue, vbFalse)
                End If
            rec.MoveNext
        Loop
    End If
    
    If StatusId = 1 Then
        GetSystemQuantity
        ComputeAdjustment
        MsgBox "System quantity has been updated to ensure data accuracy.", vbInformation
    End If
    
    If StatusId = 2 Then
        txtStatus.Text = "Completed"
        'picCompleted.Left = 5760
        'picCompleted.Visible = True
        ''pic_Cancelled.Visible = False
    End If
    
    If StatusId = 7 Then
        ''pic_Cancelled.Left = 5760
        ''pic_Cancelled.Visible = True
        'picCompleted.Visible = False
        txtStatus.Text = "Cancelled"
    End If
    CountTotal
    con.Close
End Sub

Private Sub lvSearch_ItemClick(ByVal item As MSComctlLib.ListItem)
    SelectOrders
    LoadImageStatus picStatus, GetStatus(StatusId)
End Sub

Private Sub tb_Standard_ButtonClick(ByVal Button As MSComctlLib.Button)
    If EditAccessRights(28) = False Then
        MsgBox ErrorCodes(74), vbCritical, "Limited Rights"
        Exit Sub
    End If
    Select Case Button.Index
        Case 1 ' New
            Initialize
        Case 2 'Save
            If StatusId <= 1 Then
                Save (1) 'Status Open
                LoadImageStatus picStatus, GetStatus(StatusId)
            ElseIf StatusId = 2 Then
                MsgBox "Save failed. No changes made. Order is already complete.", vbCritical
            ElseIf StatusId = 7 Then
                MsgBox "Save failed. No changes made. Order is already cancelled.", vbCritical
            End If
        Case 4 'Cancel
            If StatusId = 7 Or StatusId = 2 Then
                MsgBox "Save failed. No changes made. Order is already cancelled/completed.", vbCritical
                Exit Sub
            End If
            If AuditStockId <> 0 Then
                Dim x As Variant
                x = MsgBox("This will cancel the transaction. Proceed?", vbCritical + vbOKCancel)
                If x = vbOK Then
                    Save (7)
                    LoadImageStatus picStatus, GetStatus(StatusId)
                End If
            End If
        Case 6 'PRINT PREVIEW
            If AuditStockId <> 0 Then
                Screen.MousePointer = vbHourglass
                BASE_PrintPreviewFrm.Show
                Dim crxApp As New CRAXDRT.Application
                Dim crxRpt As New CRAXDRT.Report
                Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\INV_AuditStock.rpt")
                Call ResetRptDB(crxRpt)
                crxRpt.DiscardSavedData
                crxRpt.RecordSelectionFormula = "{INV_AuditStock.AuditStockId}= " & AuditStockId & ""
                
                'crxRpt.PrintOut False

                BASE_PrintPreviewFrm.CRViewer.ReportSource = crxRpt
                BASE_PrintPreviewFrm.CRViewer.ViewReport
                BASE_PrintPreviewFrm.CRViewer.Zoom 1
                Screen.MousePointer = vbDefault
                
                SavePOSAuditTrail UserId, WorkstationId, "", "Generated print preview on audit stock order: " & txtOrderNumber.Text, "INVENTORY"
            End If
    End Select
End Sub

Private Sub txtItemSearch_Change()
    btnItemSearch_Click
End Sub


Private Sub Save(ByVal iStatusId As Integer, Optional isReopen As Variant)
    If lvItems.ListItems.Count > 0 Then
        
        'On Error GoTo ErrorHandler
        
        Dim item As MSComctlLib.ListItem
        Set con = New ADODB.Connection
        Set rec = New ADODB.Recordset
        Set cmd = New ADODB.Command
        
        con.ConnectionString = ConnString
        con.Open
        con.BeginTrans
        
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.Parameters.Append cmd.CreateParameter("@AuditStockId", adInteger, adParamInputOutput, , AuditStockId)
        If Trim(txtOrderNumber.Text) = "" Then
            cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInputOutput, 50, Null)
        Else
            cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInputOutput, 50, txtOrderNumber.Text)
        End If
        cmd.Parameters.Append cmd.CreateParameter("@AuditLocationId", adInteger, adParamInput, , cmbAuditLocation.ItemData(cmbAuditLocation.ListIndex))
        cmd.Parameters.Append cmd.CreateParameter("@Date", adDBDate, adParamInput, , dtOrder.value)
        cmd.Parameters.Append cmd.CreateParameter("@Time", adDBTime, adParamInput, , dtTime.value)
        cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 500, txtRemarks.Text)
        cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , iStatusId)
        cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
        
        If AuditStockId = 0 Then
            cmd.CommandText = "INV_AuditStock_Insert"
        Else
            cmd.CommandText = "INV_AuditStock_Update"
        End If
        
        cmd.Execute
        
        If AuditStockId = 0 Then 'ADD TO SEARCH
            Set item = lvSearch.ListItems.add(, , cmd.Parameters("@AuditStockId"))
                item.SubItems(1) = cmd.Parameters("@OrderNumber")
                If iStatusId = 1 Then item.SubItems(2) = "Open"
                If iStatusId = 2 Then item.SubItems(2) = "Completed"
                item.Selected = True
                item.EnsureVisible
            
            AuditStockId = cmd.Parameters("@AuditStockId")
            txtOrderNumber.Text = cmd.Parameters("@OrderNumber")
            
            SavePOSAuditTrail UserId, WorkstationId, "", "Created new audit stock order: " & txtOrderNumber.Text, "INVENTORY"
        Else
            For Each item In lvSearch.ListItems
                If Val(item.Text) = Val(cmd.Parameters("@AuditStockId")) Then
                    item.SubItems(1) = cmd.Parameters("@OrderNumber")
                    If iStatusId = 1 Then item.SubItems(2) = "Open"
                    If iStatusId = 2 Then item.SubItems(2) = "Completed"
                    item.Selected = True
                    item.EnsureVisible
                    Exit For
                End If
            Next
            
            Dim auditstatus As String
            If iStatusId = 1 Then auditstatus = "Open"
            If iStatusId = 2 Then auditstatus = "Completed"
            If iStatusId = 7 Then auditstatus = "Cancelled"
            SavePOSAuditTrail UserId, WorkstationId, "", "Updated audit stock order: " & txtOrderNumber.Text & " - Status: " & auditstatus, "INVENTORY"
        End If
        
        'SAVE ORDER LINE
        For Each item In lvItems.ListItems
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc

            cmd.Parameters.Append cmd.CreateParameter("@AuditStockLineId", adInteger, adParamInputOutput, , NVAL(item.SubItems(2)))
            cmd.Parameters.Append cmd.CreateParameter("@AuditStockId", adInteger, adParamInput, , AuditStockId)
            cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , item.Text)
            cmd.Parameters.Append cmd.CreateParameter("@Uom", adVarChar, adParamInput, 250, item.SubItems(5))
            cmd.Parameters.Append cmd.CreateParameter("@SystemQuantity", adDecimal, adParamInput, , NVAL(item.SubItems(6)))
                                  cmd.Parameters("@SystemQuantity").Precision = 18
                                  cmd.Parameters("@SystemQuantity").NumericScale = 2
            If item.SubItems(7) = "" Then
                cmd.Parameters.Append cmd.CreateParameter("@PhysicalQuantity", adDecimal, adParamInput, , Null)
                                  cmd.Parameters("@PhysicalQuantity").Precision = 18
                                  cmd.Parameters("@PhysicalQuantity").NumericScale = 2
            Else
                cmd.Parameters.Append cmd.CreateParameter("@PhysicalQuantity", adDecimal, adParamInput, , NVAL(item.SubItems(7)))
                                  cmd.Parameters("@PhysicalQuantity").Precision = 18
                                  cmd.Parameters("@PhysicalQuantity").NumericScale = 2
            End If
            If item.SubItems(8) = "" Then
                cmd.Parameters.Append cmd.CreateParameter("@Adjustment", adDecimal, adParamInput, , Null)
                                  cmd.Parameters("@Adjustment").Precision = 18
                                  cmd.Parameters("@Adjustment").NumericScale = 2
            Else
                cmd.Parameters.Append cmd.CreateParameter("@Adjustment", adDecimal, adParamInput, , NVAL(item.SubItems(8)))
                                  cmd.Parameters("@Adjustment").Precision = 18
                                  cmd.Parameters("@Adjustment").NumericScale = 2
            End If
            cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , iStatusId)
            cmd.Parameters.Append cmd.CreateParameter("@AuditLocationId", adInteger, adParamInput, , cmbAuditLocation.ItemData(cmbAuditLocation.ListIndex))
            
            If item.SubItems(2) = "" Then
                cmd.CommandText = "INV_AuditStockLine_Insert"
                cmd.Execute
            Else
                cmd.CommandText = "INV_AuditStockLine_Update"
                cmd.Execute
            End If
            'cmd.Execute
            item.SubItems(2) = cmd.Parameters("@AuditStockLineId")
        Next
        
        con.CommitTrans
        con.Close
        
        txtStatus.Text = "Open"
        StatusId = 1
        
        If iStatusId = 2 Then
            txtStatus.Text = "Completed"
            'picCompleted.Left = 5760
            'picCompleted.Visible = True
            ''pic_Cancelled.Visible = False
            StatusId = 2
        End If
        
        If iStatusId = 7 Then
            ''pic_Cancelled.Left = 5760
            ''pic_Cancelled.Visible = True
            'picCompleted.Visible = False
            txtStatus.Text = "Cancelled"
            StatusId = 7
            On Error Resume Next
            lvSearch.SelectedItem.SubItems(2) = "Cancelled"
        End If
    Else
        MsgBox "Save failed. No data found.", vbCritical, "PeakPOS"
    End If
    Exit Sub
'ErrorHandler:
'    con.RollbackTrans
'    con.Close
'    If IsNumeric(Err.Description) = True Then
'        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(Val(Err.Description))
'    Else
'        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & Err.Description
'    End If
'    GLOBAL_MessageFrm.Show (1)
End Sub

Public Sub Populate(ByVal data As String)
    Select Case data
        Case "AuditStockLoad"
            Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            Set rec = New ADODB.Recordset
            
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "INV_AuditStock_Get"
            cmd.Parameters.Append cmd.CreateParameter("@AuditStockId", adInteger, adParamInput, , Null)
            Set rec = cmd.Execute
            lvSearch.ListItems.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    Set item = lvSearch.ListItems.add(, , rec!AuditStockId)
                        item.SubItems(1) = rec!OrderNumber
                        item.SubItems(2) = rec!Status
                    rec.MoveNext
                Loop
            End If
            con.Close
        Case "Location"
            Set rec = Global_Data("Location")
            cmbAuditLocation.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    cmbAuditLocation.AddItem rec!Location
                    cmbAuditLocation.ItemData(cmbAuditLocation.NewIndex) = rec!LocationId
                    rec.MoveNext
                Loop
            End If
            On Error Resume Next
            cmbAuditLocation.Text = "STORE"
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
    End Select
End Sub

Private Sub txtItemSearch_GotFocus()
    selectText txtItemSearch
End Sub

Private Sub txtItemSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            If lvItemList.Visible = True Then
                lvItemList.SetFocus
            Else
                lvItems.SetFocus
            End If
    End Select
    
End Sub

Private Sub txtSearch_OrderNumber_Change()
    btnSearch_Click
End Sub
