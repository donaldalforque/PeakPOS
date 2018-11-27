VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form SO_SalesOrderFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Order"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15090
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   15090
   Visible         =   0   'False
   Begin VB.PictureBox picStatus 
      BorderStyle     =   0  'None
      Height          =   1055
      Left            =   8160
      Picture         =   "SO_SalesOrderFrm.frx":0000
      ScaleHeight     =   1050
      ScaleWidth      =   3750
      TabIndex        =   82
      Top             =   2160
      Width           =   3755
   End
   Begin VB.CommandButton btnInvoice 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Invoice"
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
      Left            =   9240
      TabIndex        =   71
      Top             =   7440
      Width           =   2535
   End
   Begin MSComctlLib.ListView lvItemList 
      Height          =   3735
      Left            =   6120
      TabIndex        =   45
      Top             =   3360
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   6588
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
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Cost"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.CommandButton btnPick 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Pick Up Order"
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
      Left            =   9240
      TabIndex        =   70
      Top             =   6960
      Width           =   2535
   End
   Begin VB.PictureBox picPaid 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1860
      Left            =   -9999
      Picture         =   "SO_SalesOrderFrm.frx":7D72
      ScaleHeight     =   1860
      ScaleWidth      =   5250
      TabIndex        =   59
      Top             =   2640
      Visible         =   0   'False
      Width           =   5250
   End
   Begin VB.CommandButton btnStatus 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Pick && Invoice"
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
      Left            =   9240
      TabIndex        =   14
      Top             =   7920
      Visible         =   0   'False
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
      Left            =   -9999
      Style           =   1  'Graphical
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   7440
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   14400
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
            Picture         =   "SO_SalesOrderFrm.frx":10D03
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_SalesOrderFrm.frx":17565
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_SalesOrderFrm.frx":1DDC7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_SalesOrderFrm.frx":24629
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_SalesOrderFrm.frx":2489E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_SalesOrderFrm.frx":24F0F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame LayoutFrame_Search 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   4575
      Begin VB.TextBox txtSearch_Name 
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
         TabIndex        =   20
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
         TabIndex        =   24
         Top             =   2400
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
         TabIndex        =   21
         Top             =   1200
         Width           =   3015
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
         TabIndex        =   19
         Top             =   480
         Width           =   3015
      End
      Begin MSComctlLib.ListView lvSearch 
         Height          =   6015
         Left            =   120
         TabIndex        =   25
         Top             =   2880
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   10610
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
         NumItems        =   4
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
            Text            =   "Ref #"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComCtl2.DTPicker DateTo 
         Height          =   345
         Left            =   1440
         TabIndex        =   23
         Top             =   1920
         Width           =   3015
         _ExtentX        =   5318
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
         Format          =   143785985
         CurrentDate     =   41686
      End
      Begin MSComCtl2.DTPicker DateFrom 
         Height          =   345
         Left            =   1440
         TabIndex        =   22
         Top             =   1560
         Width           =   3015
         _ExtentX        =   5318
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
         Format          =   143785985
         CurrentDate     =   41686
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         TabIndex        =   80
         Top             =   840
         Width           =   555
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
         TabIndex        =   69
         Top             =   1560
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
         TabIndex        =   68
         Top             =   1920
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
         TabIndex        =   30
         Top             =   1200
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
         TabIndex        =   29
         Top             =   80
         Width           =   795
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order/Ref #"
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
         Width           =   1095
      End
   End
   Begin VB.Frame Body_Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   4640
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      Begin MSComctlLib.ListView lvCustomer 
         Height          =   3200
         Left            =   -9999
         TabIndex        =   39
         Top             =   930
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   5636
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
         NumItems        =   10
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
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "CreditLine"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "PricingScheme"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Sales Rep"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CheckBox chkWithdraw 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Auto Withdrawal Slip"
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
         Left            =   7200
         TabIndex        =   63
         Top             =   3000
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Frame Frame_Header1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   2175
         Left            =   240
         TabIndex        =   46
         Top             =   600
         Width           =   6735
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
            Left            =   1200
            TabIndex        =   81
            Top             =   360
            Width           =   3015
         End
         Begin VB.TextBox txtInterestRate 
            Alignment       =   1  'Right Justify
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
            Left            =   5520
            TabIndex        =   4
            Top             =   720
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtDays 
            Alignment       =   1  'Right Justify
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
            Left            =   5520
            TabIndex        =   3
            Top             =   360
            Visible         =   0   'False
            Width           =   1215
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
            ItemData        =   "SO_SalesOrderFrm.frx":2B771
            Left            =   5520
            List            =   "SO_SalesOrderFrm.frx":2B773
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   0
            Width           =   1215
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
            Height          =   405
            Left            =   1200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            Top             =   1800
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
            TabIndex        =   15
            Top             =   1440
            Width           =   3015
         End
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
         Begin VB.Label lblCreditLine 
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
            TabIndex        =   79
            Top             =   1095
            Width           =   3015
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Credit Line"
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
            TabIndex        =   78
            Top             =   1095
            Width           =   1005
         End
         Begin VB.Label Label18 
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
            TabIndex        =   64
            Top             =   345
            Width           =   900
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
            TabIndex        =   55
            Top             =   720
            Width           =   720
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
            TabIndex        =   54
            Top             =   750
            Width           =   3015
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
            Left            =   4320
            TabIndex        =   53
            Top             =   720
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
            Left            =   4320
            TabIndex        =   52
            Top             =   360
            Visible         =   0   'False
            Width           =   435
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
            TabIndex        =   50
            Top             =   0
            Width           =   570
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
            TabIndex        =   49
            Top             =   1800
            Width           =   750
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
            TabIndex        =   48
            Top             =   1440
            Width           =   600
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Business"
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
            TabIndex        =   47
            Top             =   0
            Width           =   810
         End
      End
      Begin VB.Frame Frame_Header2 
         BackColor       =   &H00FFFFFF&
         Height          =   2115
         Left            =   7170
         TabIndex        =   33
         Top             =   360
         Width           =   3165
         Begin VB.TextBox txtReferenceNumber 
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
            TabIndex        =   65
            Top             =   1680
            Width           =   2055
         End
         Begin VB.TextBox txtStatus 
            BackColor       =   &H00E0E0E0&
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
            Height          =   330
            Left            =   960
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   1320
            Width           =   2055
         End
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
         Begin MSComCtl2.DTPicker dtOrder 
            Height          =   330
            Left            =   960
            TabIndex        =   6
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
            Format          =   143785985
            CurrentDate     =   41509
         End
         Begin MSComCtl2.DTPicker dtDue 
            Height          =   330
            Left            =   960
            TabIndex        =   7
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
            Format          =   143785985
            CurrentDate     =   41509
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ref #"
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
            TabIndex        =   66
            Top             =   1680
            Width           =   465
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
            TabIndex        =   56
            Top             =   960
            Width           =   375
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
            TabIndex        =   37
            Top             =   1320
            Width           =   570
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
            TabIndex        =   35
            Top             =   600
            Width           =   435
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Order #"
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
            TabIndex        =   34
            Top             =   240
            Width           =   690
         End
      End
      Begin VB.Frame Frame_Footer 
         BackColor       =   &H00FFFFFF&
         Height          =   2175
         Left            =   120
         TabIndex        =   32
         Top             =   6720
         Width           =   10215
         Begin VB.ComboBox cmbPricingScheme 
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
            ItemData        =   "SO_SalesOrderFrm.frx":2B775
            Left            =   1200
            List            =   "SO_SalesOrderFrm.frx":2B785
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1680
            Width           =   3135
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
            TabIndex        =   18
            Top             =   240
            Visible         =   0   'False
            Width           =   1695
         End
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
            Height          =   1410
            Left            =   1200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Top             =   240
            Width           =   3135
         End
         Begin VB.TextBox txtDiscount 
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
            Height          =   330
            Left            =   2520
            TabIndex        =   13
            Top             =   360
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label lblTotalPayment 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   8460
            TabIndex        =   77
            Top             =   1160
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Paid"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7200
            MouseIcon       =   "SO_SalesOrderFrm.frx":2B7C3
            MousePointer    =   99  'Custom
            TabIndex        =   76
            Top             =   1155
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Invoiced "
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   7200
            TabIndex        =   75
            Top             =   885
            Width           =   870
         End
         Begin VB.Label lblInvoice 
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
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   8460
            TabIndex        =   74
            Top             =   885
            Width           =   1620
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
            Left            =   8460
            TabIndex        =   73
            Top             =   1440
            Width           =   1620
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            Height          =   285
            Left            =   8460
            TabIndex        =   72
            Top             =   555
            Width           =   1620
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pricing"
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
            Top             =   1680
            Width           =   630
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Discount"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            TabIndex        =   61
            Top             =   360
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "BALANCE"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   7200
            TabIndex        =   60
            Top             =   1455
            Width           =   915
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
            TabIndex        =   57
            Top             =   240
            Visible         =   0   'False
            Width           =   540
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
            TabIndex        =   51
            Top             =   240
            Width           =   795
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SUB-TOTAL"
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
            TabIndex        =   44
            Top             =   240
            Width           =   1020
         End
         Begin VB.Label lblSubTotal 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            Height          =   285
            Left            =   8460
            TabIndex        =   43
            Top             =   240
            Width           =   1620
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   7200
            TabIndex        =   42
            Top             =   555
            Width           =   495
         End
         Begin VB.Label lblCaption_AR 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "INTEREST"
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
            TabIndex        =   41
            Top             =   645
            Visible         =   0   'False
            Width           =   855
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
            Left            =   2460
            TabIndex        =   40
            Top             =   645
            Visible         =   0   'False
            Width           =   1620
         End
      End
      Begin VB.Frame Frame_Body 
         BackColor       =   &H00FFFFFF&
         Height          =   3900
         Left            =   120
         TabIndex        =   31
         Top             =   2805
         Width           =   10215
         Begin MSComctlLib.ListView lvItems 
            Height          =   3135
            Left            =   120
            TabIndex        =   10
            Top             =   600
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   5530
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
            NumItems        =   15
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "SalesOrderLineId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "SalesOrderId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Item Code"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Quantity"
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
               Text            =   "Price"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Text            =   "Sub-Total"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   8
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   9
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "LocationId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Text            =   "ProductId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Text            =   "PickQty"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Text            =   "ActualQuantity"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   14
               Text            =   "ReserveId"
               Object.Width           =   0
            EndProperty
         End
         Begin VB.TextBox txtCode 
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
            Top             =   600
            Visible         =   0   'False
            Width           =   3015
         End
         Begin VB.CommandButton btnItemSearch 
            Height          =   330
            Left            =   4320
            Picture         =   "SO_SalesOrderFrm.frx":2B915
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   240
            Width           =   375
         End
         Begin VB.TextBox txtItemSearch 
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
            TabIndex        =   9
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Code"
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
            Top             =   600
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
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
            TabIndex        =   38
            Top             =   240
            Width           =   555
         End
      End
      Begin MSComctlLib.Toolbar tb_Standard 
         Height          =   330
         Left            =   0
         TabIndex        =   26
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
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "SalesOrder"
                     Text            =   "Sales Order"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Object.Visible         =   0   'False
                     Key             =   "PickList"
                     Text            =   "Pick List"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "DeliveryReceipt"
                     Text            =   "Delivery Receipt"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Object.Visible         =   0   'False
                     Key             =   "Invoice"
                     Text            =   "Sales Invoice"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "SO_SalesOrderFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public StatusId, CustomerId, SalesOrderId, id As Long
Dim TotalSacks As Double
Dim OrderLine(10000) As Long
Dim ctrOrderLine As Long
Public CostPricing As Boolean
Public global_remarks As String
Dim identifier As Boolean

Public Sub Initialize()
    Dim txtControl As Control
    For Each txtControl In Me.Controls
        If TypeOf txtControl Is TextBox And txtControl.Name <> "txtSearch_Order" Then
            txtControl.Text = ""
            txtStatus.Text = "Open"
        End If
    Next
    

    global_remarks = ""
    txtRemarks.Text = ""
    txtReferenceNumber.Text = ""
    picStatus.Visible = False
    identifier = False
    isNotCompleted (True)
    lblBalance.Caption = "0.00"
    lblInvoice.Caption = "0.00"
    cmbCustomer.Text = ""
    cmbCustomer.BackColor = &HC0C0FF
    'cmbTerms.BackColor = &HC0C0FF
    dtOrder.value = Format(Now, "MM/DD/YY hh:mm:ss")
    dtDue.value = Format(Now, "MM/DD/YY hh:mm:ss")
    btnStatus.Caption = "Pick && Invoice"
    lvItems.ListItems.Clear
    lvCustomer.ListItems.Clear
    lvCustomer.Visible = False
    lvItemList.Visible = False
    lvCustomer.Left = -9999
    lvItemList.Left = -9999
    btnPaid.Visible = False
    
    lblSubTotal.Caption = "0.00"
    lbltotal.Caption = "0.00"
    lblRemainingBalance.Caption = "0.00"
    
    id = 1
    StatusId = 1
    CustomerId = 0
    SalesOrderId = 0
    TotalSacks = 0
    CostPricing = False
    
    ctrOrderLine = 0
    
    'DeleteReserves UserId, False, True, False
    
    On Error Resume Next
    cmbCustomer.SetFocus
    isModify = False
    cmbTerms.ListIndex = 0
    Unload SO_InvoiceFrm
End Sub
Public Sub GetPricing()
    Dim con As New ADODB.Connection
    Dim pRec As New ADODB.Recordset
    Dim item As MSComctlLib.ListItem
    
    Set cmd = New ADODB.Command
    con.ConnectionString = ConnString
    con.Open
    For Each item In lvItems.ListItems
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "INV_ProductPricing_Get"
        cmd.Parameters.Append cmd.CreateParameter("@PricingSchemeId", adInteger, adParamInput, , cmbPricingScheme.ItemData(cmbPricingScheme.ListIndex))
        cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , Val(item.SubItems(9)))
        Set pRec = cmd.Execute
        If Not pRec.EOF Then
            item.SubItems(6) = FormatNumber(pRec!price, 2, vbTrue, vbFalse)
        End If
    Next
    
    For Each item In lvItemList.ListItems
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "INV_ProductPricing_Get"
        cmd.Parameters.Append cmd.CreateParameter("@PricingSchemeId", adInteger, adParamInput, , cmbPricingScheme.ItemData(cmbPricingScheme.ListIndex))
        cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , item.Text)
        Set pRec = cmd.Execute
        If Not pRec.EOF Then
            item.SubItems(3) = FormatNumber(pRec!price, 2, vbTrue, vbFalse)
        End If
    Next
    con.Close
    CountTotal
End Sub

Private Sub Save(ByVal StatusId As Integer, Optional isReopen As Boolean)
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
        If Trim(txtOrderNumber.Text) = "" Then
            cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInputOutput, 50, Null)
        Else
            cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInputOutput, 50, txtOrderNumber.Text)
        End If
        cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtOrder.value)
        cmd.Parameters.Append cmd.CreateParameter("@DueDate", adDate, adParamInput, , dtDue.value)
        cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , StatusId)
        cmd.Parameters.Append cmd.CreateParameter("@TermId", adInteger, adParamInput, , cmbTerms.ItemData(cmbTerms.ListIndex))
        cmd.Parameters.Append cmd.CreateParameter("@PricingSchemeId", adInteger, adParamInput, , cmbPricingScheme.ItemData(cmbPricingScheme.ListIndex))
        cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInput, , CustomerId)
        cmd.Parameters.Append cmd.CreateParameter("@Days", adDecimal, adParamInput, , Val(Replace(txtDays.Text, ",", "")))
                              cmd.Parameters("@Days").Precision = 18
                              cmd.Parameters("@Days").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@InterestRate", adDecimal, adParamInput, , Val(Replace(txtInterestRate.Text, ",", "")))
                              cmd.Parameters("@InterestRate").Precision = 18
                              cmd.Parameters("@InterestRate").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@Cash", adDecimal, adParamInput, , Val(Replace(txtCash.Text, ",", "")))
                              cmd.Parameters("@Cash").Precision = 18
                              cmd.Parameters("@Cash").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@Interest", adDecimal, adParamInput, , Val(Replace(lblInterest.Caption, ",", "")))
                              cmd.Parameters("@Interest").Precision = 18
                              cmd.Parameters("@Interest").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@Subtotal", adDecimal, adParamInput, , Val(Replace(lblSubTotal.Caption, ",", "")))
                              cmd.Parameters("@Subtotal").Precision = 18
                              cmd.Parameters("@Subtotal").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@Total", adDecimal, adParamInput, , Val(Replace(lbltotal.Caption, ",", "")))
                              cmd.Parameters("@Total").Precision = 18
                              cmd.Parameters("@Total").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 500, txtRemarks.Text)
        cmd.Parameters.Append cmd.CreateParameter("@Salesman", adVarChar, adParamInput, 250, txtSalesman.Text)
        cmd.Parameters.Append cmd.CreateParameter("@ReferenceNumber", adVarChar, adParamInput, 250, txtReferenceNumber.Text)
        cmd.Parameters.Append cmd.CreateParameter("@Discount", adDecimal, adParamInput, , Val(Replace(txtDiscount.Text, ",", "")))
                              cmd.Parameters("@Discount").Precision = 18
                              cmd.Parameters("@Discount").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , 1) 'NOT SET!
        cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@isReopen", adBoolean, adParamInput, , isReopen)
        cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
        cmd.Parameters.Append cmd.CreateParameter("@WorkStationId", adInteger, adParamInput, , WorkstationId)
        
        If SalesOrderId = 0 Then
            cmd.CommandText = "SO_SalesOrder_Insert"
            cmd.Execute
            
            SalesOrderId = cmd.Parameters("@SalesOrderId")
            txtOrderNumber.Text = cmd.Parameters("@OrderNumber")
           
            SavePOSAuditTrail UserId, WorkstationId, "", "Created new sales order: " & txtOrderNumber.Text, "SALES"
        Else
            cmd.CommandText = "SO_SalesOrder_Update"
            cmd.Execute
            
            Dim auditstatus As String
            If StatusId = 1 Then auditstatus = "Open"
            If StatusId = 2 Then auditstatus = "Completed"
            If StatusId = 7 Then auditstatus = "Cancelled"
            
            SavePOSAuditTrail UserId, WorkstationId, "", "Updated sales order: " & txtOrderNumber.Text & " - Status: " & auditstatus, "SALES"
        End If
        
        
        
        'SAVE ORDER LINE
        Dim item As MSComctlLib.ListItem

        For Each item In lvItems.ListItems
            If AllowNegativeInventory = False Then
                'UPDATE ITEM RESERVES
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = con
                cmd.CommandText = "INV_ProductReserveByModule_Update"
                cmd.CommandType = adCmdStoredProc
                cmd.Parameters.Append cmd.CreateParameter("@ReserveId", adInteger, adParamInput, , Val(item.SubItems(14)))
                cmd.Parameters.Append cmd.CreateParameter("@ModuleId", adInteger, adParamInput, , 2)
                cmd.Parameters.Append cmd.CreateParameter("@Id", adInteger, adParamInput, , SalesOrderId)
                cmd.Parameters.Append cmd.CreateParameter("@WorkStationId", adInteger, adParamInput, , WorkstationId)
                cmd.Execute
            End If
        
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            
            cmd.Parameters.Append cmd.CreateParameter("@SalesOrderLineId", adInteger, adParamInputOutput, , Val(item.Text))
            cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SalesOrderId)
            cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInput, , CustomerId)
            cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtOrder.value)
            cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , Val(item.SubItems(9)))
            cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, item.SubItems(3))
            cmd.Parameters.Append cmd.CreateParameter("@Quantity", adDecimal, adParamInput, , Val(Replace(item.SubItems(4), ",", "")))
                                  cmd.Parameters("@Quantity").Precision = 18
                                  cmd.Parameters("@Quantity").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@Uom", adVarChar, adParamInput, 250, item.SubItems(5))
            cmd.Parameters.Append cmd.CreateParameter("@Price", adDecimal, adParamInput, , Val(Replace(item.SubItems(6), ",", "")))
                                  cmd.Parameters("@Price").Precision = 18
                                  cmd.Parameters("@Price").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@Subtotal", adDecimal, adParamInput, , Val(Replace(item.SubItems(7), ",", "")))
                                  cmd.Parameters("@Subtotal").Precision = 18
                                  cmd.Parameters("@Subtotal").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@LocationId", adInteger, adParamInput, , Val(item.SubItems(8)))
            cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , StatusId)
            cmd.Parameters.Append cmd.CreateParameter("@ReserveId", adInteger, adParamInput, , Val(item.SubItems(14)))
            cmd.Parameters.Append cmd.CreateParameter("@ActualQuantity", adDecimal, adParamInput, , Val(Replace(item.SubItems(13), ",", "")))
                                  cmd.Parameters("@ActualQuantity").Precision = 18
                                  cmd.Parameters("@ActualQuantity").NumericScale = 2
                                  
            If item.Text = "" Then
                cmd.CommandText = "SO_SalesOrderLine_Insert"
            Else
                cmd.Parameters.Append cmd.CreateParameter("@isReopen", adBoolean, adParamInput, , isReopen)
                cmd.CommandText = "SO_SalesOrderLine_Update"
            End If
            cmd.Execute
            item.Text = cmd.Parameters("@SalesOrderLineId")
        Next

        'DELETE ORDERLINE IF ANY
        Dim ctr As Integer
        For ctr = 0 To ctrOrderLine
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc

            If OrderLine(ctr) <> 0 Then
                cmd.Parameters.Append cmd.CreateParameter("@SalesOrderLineId", adInteger, adParamInput, , OrderLine(ctr))
                cmd.CommandText = "SO_SalesOrderLine_Delete"
                cmd.Execute
                
                
            Else
                Exit For
            End If
        Next

        con.CommitTrans
        con.Close
        
        If StatusId = 2 Then
            isNotCompleted (False)
            txtStatus.Text = "Complete"
            Me.StatusId = 2
        End If
        
        Dim isFound As Boolean
        isFound = False
        For Each item In lvSearch.ListItems
            If SalesOrderId = item.Text Then
                item.SubItems(1) = txtOrderNumber.Text
                item.SubItems(2) = txtReferenceNumber.Text
                item.SubItems(3) = txtStatus.Text
                isFound = True
                item.Selected = True
                item.EnsureVisible
                Exit For
            End If
        Next
        If isFound = False Then
            Set item = lvSearch.ListItems.add(, , SalesOrderId)
                item.SubItems(1) = txtOrderNumber.Text
                item.SubItems(2) = txtReferenceNumber.Text
                item.SubItems(3) = txtStatus.Text
                item.Selected = True
                item.EnsureVisible
        End If
        
'        'PRINT WITHDRAWAL
'
'        'Screen.MousePointer = vbHourglass
'        'BASE_PrintPreviewFrm.Show
'        Dim crxApp As New CRAXDRT.Application
'        Dim crxRpt As New CRAXDRT.Report
'        Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\INV_WarehouseWithdrawal.rpt")
'        crxRpt.RecordSelectionFormula = "{INV_WarehouseWithdrawal.WarehouseWithdrawalId}= " & WarehouseWithdrawalId & ""
'        crxRpt.DiscardSavedData
'
'        Call ResetRptDB(crxRpt)
'        crxRpt.PrintOut False

        '****************
        
        'PRINT PREVIEW
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
    Dim Total, subtotal, discount, subtotal1, interestrate, Interest, days, cash As Double
    Dim item As MSComctlLib.ListItem
    
    If IsNumeric(txtDays.Text) = False Then
        days = 0
    Else
        days = txtDays.Text
    End If
    
    If IsNumeric(txtInterestRate.Text) = False Then
        interestrate = 0
    Else
        interestrate = txtInterestRate.Text
    End If
    
    If IsNumeric(txtCash.Text) = False Then
        cash = 0
    Else
        cash = txtCash.Text
    End If
    
    If IsNumeric(txtDiscount.Text) = False Then
        discount = 0
    Else
        discount = txtDiscount.Text
    End If
    
    For Each item In lvItems.ListItems
        subtotal = Val(Replace(item.SubItems(4), ",", "")) * Val(Replace(item.SubItems(6), ",", ""))
        item.SubItems(7) = FormatNumber(subtotal, 2, vbTrue, vbFalse)
        subtotal1 = subtotal1 + subtotal
    Next
    subtotal1 = subtotal1 + cash - discount
    lblSubTotal.Caption = FormatNumber(subtotal1, 2, vbTrue, vbFalse)
    
    'COMPUTE INTEREST
    Interest = ((days / 30) * (subtotal1 * (interestrate / 100)))
    lblInterest.Caption = FormatNumber(Interest, 2, vbTrue, vbFalse)
    
    Total = Interest + subtotal1
    lbltotal.Caption = FormatNumber(Total, 2, vbTrue, vbFalse)
    'lblRemainingBalance.Caption = lblTotal.Caption
    lblTotalPayment.Caption = FormatNumber(NVAL(lblBalance.Caption) - NVAL(lblInvoice.Caption), 2, vbTrue)
End Sub
Public Sub Populate(ByVal data As String)
    Select Case data
        Case "Terms"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Terms")
            cmbTerms.Clear
            cmbTerms.AddItem ""
'            cmbTerms.ItemData(cmbTerms.NewIndex) = -1
'            cmbTerms.AddItem "<Add New..>"
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
                If rec!isCashAdvance = "False" Then
                    CustomerId = rec!CustomerId
                    cmbCustomer.BackColor = vbWhite
                    StatusId = rec!StatusId
                    txtStatus.Text = rec!Status
                    txtOrderNumber.Text = rec!OrderNumber
                    txtAddress.Text = rec!Address
                    txtContact.Text = rec!Phone
                    dtOrder.value = Format(rec!Date, "MM/DD/YY")
                    dtDue.value = Format(rec!DueDate, "MM/DD/YY")
                    lblSubTotal.Caption = FormatNumber(rec!subtotal, 2, vbTrue)
                    txtDays.Text = rec!days
                    txtInterestRate.Text = rec!interestrate
                    txtCash.Text = FormatNumber(rec!cash, 2, vbTrue, vbFalse)
                    lblRemainingBalance.Caption = FormatNumber(rec!balance, 2, vbTrue, vbFalse)
                    'On Error Resume Next
                    lblInvoice.Caption = FormatNumber(rec!invoiced, 2, vbTrue, vbFalse)
                    txtReferenceNumber.Text = rec!ReferenceNumber
                    If IsNull(rec!discount) = True Then
                        txtDiscount.Text = ""
                    Else
                        txtDiscount.Text = FormatNumber(rec!discount, 2, vbTrue, vbFalse)
                    End If
                    
                    If rec!StatusId = 2 Then 'COMPLETED
                        isNotCompleted (False)
                        tb_Standard.Buttons(4).Caption = "Cancel"
                        tb_Standard.Buttons(4).Image = 3
                        btnPaid.Visible = True
                        btnStatus.Caption = "Pick && Invoice"
                    ElseIf rec!StatusId = 7 Then 'Cancelled
                        isNotCompleted (False)
                        'tb_Standard.Buttons(4).Caption = "Re-Open"
                        'tb_Standard.Buttons(4).Image = 6
                        btnPaid.Visible = False
                        btnStatus.Caption = "Pick && Invoice"
                    ElseIf rec!StatusId = 6 Then 'PAID
                        isNotCompleted (False)
                        tb_Standard.Buttons(4).Caption = "Cancel"
                        tb_Standard.Buttons(4).Image = 3
                        btnPaid.Visible = False
                        'btnStatus.Caption = "Reopen Order"
                    ElseIf ((rec!StatusId = 4) Or (rec!StatusId = 8) Or (rec!StatusId = 9) Or (StatusId = 5)) Then
                        isNotCompleted (False)
                        tb_Standard.Buttons(4).Caption = "Cancel"
                        tb_Standard.Buttons(4).Image = 3
                        btnPaid.Visible = False
                    Else
                        isNotCompleted (True)
                        tb_Standard.Buttons(4).Caption = "Cancel"
                        tb_Standard.Buttons(4).Image = 3
                        btnPaid.Visible = False
                        btnStatus.Caption = "Pick && Invoice"
                    End If
                    lblRemainingBalance.Caption = FormatNumber(rec!balance, 2, vbTrue, vbFalse)
                    lbltotal.Caption = FormatNumber(rec!Total, 2, vbTrue, vbFalse)
                    If rec!PricingScheme = "" Then
                        cmbPricingScheme.ListIndex = 0
                    Else
                        cmbPricingScheme.Text = rec!PricingScheme
                    End If
                    
                    'PREVENT ERROR ON DISPLAY FOR cmbCustomer_Change Event
                    'On Error Resume Next
                    cmbCustomer.Text = rec!Name
                    txtSalesman.Text = rec!SalesRep
                    lvCustomer.Left = -9999
                    lvCustomer.Visible = False
                    If rec!Terms = "" Then cmbTerms.ListIndex = 0 Else cmbTerms.Text = rec!Terms
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
                    If rec!isCashAdvance = "False" Then
                        Set item = lvSearch.ListItems.add(, , rec!SalesOrderId)
                            item.SubItems(1) = rec!OrderNumber
                            item.SubItems(2) = rec!Status
                    End If
                    rec.MoveNext
                Loop
            End If
            con.Close
        Case "SalesOrderLineGet"
            On Error Resume Next
            Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            Set rec = New ADODB.Recordset
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "SO_SalesOrderLine_Get"
            cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SalesOrderId)
            Set rec = cmd.Execute
            lvItems.ListItems.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    Set item = lvItems.ListItems.add(, , rec!SalesOrderLineId)
                        item.SubItems(1) = rec!SalesOrderId
                        item.SubItems(2) = rec!itemcode
                        item.SubItems(3) = rec!Name
                        item.SubItems(4) = FormatNumber(rec!quantity, 2, vbTrue)
                        item.SubItems(5) = rec!Uom
                        item.SubItems(6) = FormatNumber(rec!price, 2, vbTrue)
                        item.SubItems(7) = FormatNumber(rec!subtotal, 2, vbTrue)
                        item.SubItems(8) = rec!LocationId
                        item.SubItems(9) = rec!ProductId
                        item.SubItems(13) = rec!ActualQuantity
                        On Error Resume Next
                        item.SubItems(14) = rec!ReserveId
                        item.SubItems(12) = rec!pickedquantity
                    rec.MoveNext
                Loop
            End If
            con.Close
        Case "PricingScheme"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("PricingScheme")
            
            cmbPricingScheme.Clear
            cmbPricingScheme.AddItem ""
            cmbPricingScheme.ItemData(cmbPricingScheme.NewIndex) = 0
            
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        cmbPricingScheme.AddItem rec!PricingScheme
                        cmbPricingScheme.ItemData(cmbPricingScheme.NewIndex) = rec!PricingSchemeId
                    End If
                    rec.MoveNext
                Loop
            End If
            
            cmbPricingScheme.ListIndex = 0
    End Select
End Sub


Private Sub btnInvoice_Click()
    If SalesOrderId = 0 Then Exit Sub
    SO_InvoiceFrm.Show '(1)
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
    cmd.CommandText = "BASE_Product_Search1"
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 500, txtItemSearch.Text)
    Dim LastProductId As Long
    Set rec = cmd.Execute
    If Not rec.EOF Then
        lvItemList.ListItems.Clear
        Do Until rec.EOF
            If rec!isActive = "True" Then
                If LastProductId <> rec!ProductId Then
                    Set item = lvItemList.ListItems.add(, , rec!ProductId)
                        item.SubItems(1) = rec!itemcode
                        item.SubItems(2) = rec!Name
                        item.SubItems(3) = FormatNumber(rec!unitprice, 2, vbTrue)
                        item.SubItems(4) = rec!Uom
                        item.SubItems(5) = rec!unitcost
                    lvItemList.Visible = True
                    lvItemList.Left = 6070
                    'lvItemList.Top = 3720
                    LastProductId = rec!ProductId
                    rec.MoveNext
                Else
                    rec.MoveNext
                End If
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

Private Sub btnPaid_Click()
'    Unload FIN_PaymentDialogFrm
'    With FIN_PaymentDialogFrm
'        .isBatchPO = False
'        .isBatchSO = False
'        .isPO = False
'        .isSO = False
'        .isSOAR = True
'        .chkOnline.Visible = True
'        .lblCustomerName.Caption = cmbCustomer.text
'        .lblOrderNumber.Caption = txtOrderNumber.text
'        .lblBalance.Caption = lblRemainingBalance.Caption
'        .lblBalanceForwarded.Caption = lblRemainingBalance.Caption
'        .Show (1)
'    End With
End Sub

Private Sub btnPick_Click()
    If EditAccessRights(35) = False Then
        MsgBox ErrorCodes(74), vbCritical, "Limited Rights"
        Exit Sub
    End If
    
    If SalesOrderId = 0 Then Exit Sub
    If StatusId >= 6 Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(63)
        GLOBAL_MessageFrm.Show (1)
        Exit Sub
    End If
    
    Dim totalReceived, totalOrdered As Double
    Dim item As MSComctlLib.ListItem
    
    Unload SO_PickOrderFrm
    
    For Each item In lvItems.ListItems
        totalOrdered = totalOrdered + Val(Replace(item.SubItems(4), ",", ""))
        totalReceived = totalReceived + Val(Replace(item.SubItems(12), ",", ""))
    Next
    
    'Validate if All orders already fullfilled
    If totalOrdered <= totalReceived And StatusId = 5 Then
'        Dim X As Variant
'        X = MsgBox("All orders have already been picked. Would you like to complete this order?", vbYesNo + vbQuestion)
'        If X = vbYes Then
'            'UPDATE STATUS
'            Set con = New ADODB.Connection
'            con.ConnectionString = ConnString
'            con.Open
'            Set cmd = New ADODB.Command
'                cmd.ActiveConnection = con
'                cmd.CommandType = adCmdStoredProc
'                cmd.CommandText = "SO_SalesOrderStatus_Update"
'                cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SO_SalesOrderFrm.SalesOrderId)
'                cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , 2) 'COMPLETE
'                cmd.Execute
'            con.Close
'            MsgBox "Order completed.", vbInformation
'
'            lvSearch.SelectedItem.SubItems(2) = "Completed"
'            StatusId = 2
'            'Requery Data
'            'Dim item As MSComctlLib.ListItem
'            Set item = lvSearch.SelectedItem
'            lvSearch_ItemClick item
'        Else
            SO_PickOrderFrm.txtOrderNumber.Text = txtOrderNumber.Text
            'PO_ReceiveOrderFrm.SalesOrderIdx = Me.SalesOrderId
            SO_PickOrderFrm.Show
'        End If
    Else
        SO_PickOrderFrm.txtOrderNumber.Text = txtOrderNumber.Text
        'PO_ReceiveOrderFrm.SalesOrderIdx = Me.SalesOrderId
        SO_PickOrderFrm.Show
    End If
End Sub

Public Sub btnSearch_Click()
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
    cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInput, 50, txtSearch_OrderNumber.Text)
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 50, txtSearch_name.Text)
    Dim item As MSComctlLib.ListItem
    Set rec = cmd.Execute
    lvSearch.ListItems.Clear
    
    Dim strdatefrom, strdateto As String
    strdatefrom = datefrom.value & " 00:00:00"
    strdateto = dateto.value & " 23:59:59"
    
    If Not rec.EOF Then
        Do Until rec.EOF
            If rec!isCashAdvance = "False" Then
                'If rec!Date >= datefrom.value And rec!Date <= dateto.value Then
                If rec!Date >= CDate(strdatefrom) And rec!Date <= CDate(strdateto) Then
                    Set item = lvSearch.ListItems.add(, , rec!SalesOrderId)
                        item.SubItems(1) = rec!OrderNumber
                        item.SubItems(2) = rec!ReferenceNumber
                        item.SubItems(3) = rec!Status
                End If
            End If
            rec.MoveNext
        Loop
    End If
    con.Close
End Sub

Private Sub btnStatus_Click()
    'CHECK STATUS OF ORDER
    

'    If StatusId = 1 Then
'        Save (2)
'        If Me.StatusId = 2 Then
'            'btnStatus.Caption = "Reopen Order"
'            btnPaid.Visible = True
'        End If
'    ElseIf StatusId = 2 Or StatusId = 3 Then
''        Dim X As Variant
'        X = MsgBox("Are you sure you want to reopen this order? This will revert all connected " & _
'                    "transactions such as payments.", vbCritical + vbYesNo, "WARNING")
'        If X = vbYes Then
'            Save 1, True
'            btnStatus.Caption = "Pick && Invoice"
'            txtStatus.text = "Open"
'            btnPaid.Visible = False
'            isNotCompleted (True)
'            'picCompleted.Visible = False
'            'picCompleted.Left = -9999
'            'picPaid.Left = -9999
'            'picPaid.Visible = False
'            Me.StatusId = 1
        'Else
    'End If
   ' End If
End Sub

Private Sub cmbCustomer_Change()
    If identifier = True Then Exit Sub
    If Trim(cmbCustomer.Text) <> "" Then
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
        cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, Trim(cmbCustomer.Text))
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
                        item.SubItems(6) = rec!PricingSchemeId
                        item.SubItems(7) = rec!creditlimit
                        item.SubItems(8) = rec!PricingScheme
                        item.SubItems(9) = rec!SalesRep
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
            cmd.Parameters.Append cmd.CreateParameter("@CustomerCode", adVarChar, adParamInput, 50, cmbCustomer.Text)
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
                        item.SubItems(6) = rec!PricingSchemeId
                        item.SubItems(7) = rec!creditlimit
                        item.SubItems(8) = rec!PricingScheme
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

Private Sub cmbPricing_Click()
'    Set con = New ADODB.Connection
'    Set rec = New ADODB.Recordset
'    Dim item As MSComctlLib.ListItem
'
'    con.ConnectionString = ConnString
'    con.Open
'    For Each item In lvItems.ListItems
'        Set cmd = New ADODB.Command
'        cmd.ActiveConnection = con
'        cmd.CommandText = "BASE_Product_Get"
'        cmd.CommandType = adCmdStoredProc
'        cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , item.SubItems(9))
'        Set rec = cmd.Execute
'        If Not rec.EOF Then
'            Do Until rec.EOF
'                If cmbPricing.ListIndex = 0 Then
'                    item.SubItems(6) = FormatNumber(rec!unitprice, 2, vbTrue)
'                ElseIf cmbPricing.ListIndex = 1 Then '30 Days
'                    item.SubItems(6) = FormatNumber(rec!price1, 2, vbTrue)
'                ElseIf cmbPricing.ListIndex = 2 Then '60 Days
'                    item.SubItems(6) = FormatNumber(rec!price2, 2, vbTrue)
'                ElseIf cmbPricing.ListIndex = 3 Then '90 Days
'                    item.SubItems(6) = FormatNumber(rec!price3, 2, vbTrue)
'                End If
'                rec.MoveNext
'            Loop
'        End If
'    Next
'    con.Close
End Sub

Private Sub cmbPricingScheme_Click()
    GetPricing
End Sub

Private Sub cmbTerms_Click()
    If cmbTerms.ListIndex > 0 Then
        dtDue.value = Format(Now, "MM/DD/YY")
        dtDue.value = dtDue.value + GetTermDays(cmbTerms.ItemData(cmbTerms.ListIndex))
    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            If lvCustomer.Visible = True Then
                lvCustomer.Visible = False
                lvCustomer.Left = -9999
                cmbCustomer.SetFocus
            ElseIf lvItemList.Visible = True Then
                lvItemList.Visible = False
                lvItemList.Left = -9999
                txtItemSearch.SetFocus
                'txtCode.SetFocus
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
        Case vbKeyF1
            Unload INV_ProductSearch
            INV_ProductSearch.isPO = False
            INV_ProductSearch.isWithdraw = False
            INV_ProductSearch.isSO = True
            INV_ProductSearch.isAS = False
            INV_ProductSearch.Show (1)
    End Select
End Sub

Private Sub Form_Load()
    '****** REGION Listview Columns *********
    lvSearch.ColumnHeaders(2).width = lvSearch.width * 0.32
    lvSearch.ColumnHeaders(3).width = lvSearch.width * 0.32
    lvSearch.ColumnHeaders(4).width = lvSearch.width * 0.32
    
    lvItemList.ColumnHeaders(2).width = lvItemList.width * 0.13
    lvItemList.ColumnHeaders(3).width = lvItemList.width * 0.7
    lvItemList.ColumnHeaders(4).width = lvItemList.width * 0.13
    
    lvCustomer.ColumnHeaders(2).width = lvCustomer.width * 0.18
    lvCustomer.ColumnHeaders(3).width = lvCustomer.width * 0.36
    lvCustomer.ColumnHeaders(4).width = lvCustomer.width * 0.18
    lvCustomer.ColumnHeaders(10).width = lvCustomer.width * 0.24
    
    lvItems.ColumnHeaders(3).width = lvItems.width * 0.14
    lvItems.ColumnHeaders(4).width = lvItems.width * 0.29
    lvItems.ColumnHeaders(5).width = lvItems.width * 0.12
    lvItems.ColumnHeaders(5).width = lvItems.width * 0.12
    lvItems.ColumnHeaders(6).width = lvItems.width * 0.09
    lvItems.ColumnHeaders(7).width = lvItems.width * 0.13
    lvItems.ColumnHeaders(8).width = lvItems.width * 0.2
    
    'StatusBarWidth Me, statusBar_Main
    '****************************************
    
    Initialize
    Populate "Terms"
    Populate "Status"
    Populate "PricingScheme"
    'Populate "SalesOrderLoad"
    
    datefrom.value = Format(Now - 30, "MM/DD/YY")
    dateto.value = Format(Now, "MM/DD/YY")
    
    On Error Resume Next
    cmbSearch_Status.ListIndex = 1
    cmbPricingScheme.ListIndex = 0
    btnSearch_Click
    
    UpdateCustomerOrderDues
End Sub






Private Sub lblGrossAmount_Click()

End Sub

Private Sub lblGrossKilos_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'DELETE PENDING RESERVES UNDER WORKSTATION
    DeleteReserves WorkstationId, 2
End Sub

Private Sub lblRemainingBalance_Change()
    If IsNumeric(Val(Replace(lblRemainingBalance.Caption, ",", ""))) = False Then
        lblRemainingBalance.Caption = "0.00"
    End If
End Sub

Private Sub lblTotal_Change()
    If IsNumeric(Val(Replace(lbltotal.Caption, ",", ""))) = False Then
        lbltotal.Caption = "0.00"
    Else
        'lblRemainingBalance.Caption = lblTotal.Caption
    End If
End Sub

Private Sub lvCustomer_DblClick()
    If lvCustomer.ListItems.Count > 0 Then
        CustomerId = lvCustomer.SelectedItem.Text
        cmbCustomer.Text = lvCustomer.SelectedItem.SubItems(2)
        lblBalance.Caption = lvCustomer.SelectedItem.SubItems(3)
        lblCreditLine.Caption = FormatNumber(lvCustomer.SelectedItem.SubItems(7), 2, vbTrue, vbFalse)
        txtContact.Text = lvCustomer.SelectedItem.SubItems(4)
        txtAddress.Text = lvCustomer.SelectedItem.SubItems(5)
        txtSalesman.Text = lvCustomer.SelectedItem.SubItems(9)
        CostPricing = lvCustomer.SelectedItem.SubItems(6)
        lvCustomer.Visible = False
        lvCustomer.Left = -9999
        txtItemSearch.SetFocus
        'txtCode.SetFocus
        cmbCustomer.BackColor = vbWhite
        If lvCustomer.SelectedItem.SubItems(8) = "" Then
            cmbPricingScheme.ListIndex = 0
        Else
            cmbPricingScheme.Text = lvCustomer.SelectedItem.SubItems(8)
        End If
    End If
End Sub

Private Sub lvCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            lvCustomer_DblClick
    End Select
End Sub

Private Sub lvItemList_DblClick()
    
    SO_ItemSelectOptionFrm.txtDescription.Text = lvItemList.SelectedItem.SubItems(2)
    SO_ItemSelectOptionFrm.Show (1)
    'SO_ItemSelectOptionFrm.txtPrice.text = lvItemList.SelectedItem.SubItems(3)
    
End Sub

Private Sub lvItemList_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Not lvItemList.SelectedItem = -1 Then
                SO_ItemSelectOptionFrm.txtDescription.Text = lvItemList.SelectedItem.SubItems(2)
                SO_ItemSelectOptionFrm.Show (1)
                'SO_ItemSelectOptionFrm.txtPrice.text = lvItemList.SelectedItem.SubItems(3)
            End If
    End Select
End Sub

Private Sub lvItems_DblClick()
    If lvItems.ListItems.Count > 0 Then
        isModify = True
        With SO_ItemSelectOptionFrm
            .txtQuantity.Text = lvItems.SelectedItem.SubItems(4)
            .txtPrice.Text = lvItems.SelectedItem.SubItems(6)
            .txtDescription.Text = lvItems.SelectedItem.SubItems(3)
            .Show (1)
        End With
    End If
End Sub

Private Sub lvItems_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp
            If lvItems.ListItems.Count > 0 Then
                If lvItems.SelectedItem.Index = 1 Then
                    txtItemSearch.SetFocus
                    'txtCode.SetFocus
                End If
            End If
        Case vbKeyDelete
            If lvItems.ListItems.Count > 0 Then
                DeleteReserveLine lvItems.SelectedItem.SubItems(14)
                If lvItems.SelectedItem.Text <> "" Then
                    OrderLine(ctrOrderLine) = Val(lvItems.SelectedItem.Text)
                    ctrOrderLine = ctrOrderLine + 1
                    
                    SavePOSAuditTrail UserId, WorkstationId, "", "Removed item: " & lvItems.SelectedItem.SubItems(3) & " from sales order: " & txtOrderNumber.Text
                    
                    lvItems.ListItems.Remove (lvItems.SelectedItem.Index)
                Else
                    lvItems.ListItems.Remove (lvItems.SelectedItem.Index)
                End If
            End If
        Case vbKeyReturn
            lvItems_DblClick
    End Select
    CountTotal
End Sub

Private Sub lvSearch_ItemClick(ByVal item As MSComctlLib.ListItem)
    If lvSearch.ListItems.Count > 0 Then
        'Initialize
        If AllowNegativeInventory = False Then
            DeleteReserves WorkstationId, 2
        End If
        
        identifier = True
        SalesOrderId = lvSearch.SelectedItem.Text
        Populate "SalesOrderGet"
        Populate "SalesOrderLineGet"
        Populate "Customer"
        CountTotal
        identifier = False
        
        LoadImageStatus picStatus, GetStatus(StatusId)
    End If
End Sub



Private Sub tb_Standard_ButtonClick(ByVal Button As MSComctlLib.Button)
    If EditAccessRights(13) = False Then
        MsgBox ErrorCodes(74), vbCritical, "Limited Rights"
        Exit Sub
    End If
    Select Case Button.Index
        Case 1 ' New
            Initialize
        Case 2 'Save
            If StatusId <= 3 Then
                Save (StatusId)
                LoadImageStatus picStatus, GetStatus(StatusId)
            Else
                MsgBox "Save failed. Cannot save a record when it is Invoiced, Picked, Paid or Cancelled.", vbCritical, "Error Saving.."
            End If
        Case 4 'Cancel
            If StatusId = 7 Then
                MsgBox "Order already cancelled.", vbCritical
                Exit Sub
            End If
            If SalesOrderId <> 0 Then
                Dim x As Variant
                x = MsgBox("Are you sure you want to cancel this order?", vbQuestion + vbYesNo)
                If x = vbNo Then Exit Sub
            
                If StatusId <= 6 Then 'Status Open > Cancelled
                    StatusId = 7
                    txtStatus.Text = "Cancelled"
                    Save (7)
                    isNotCompleted (False)
                ElseIf StatusId = 7 Then 'Status Cancelled > Activate
                    StatusId = 1
                    txtStatus.Text = "Open"
                    Save (1)
                    isNotCompleted (True)
                    tb_Standard.Buttons(4).Caption = "Cancel"
                    tb_Standard.Buttons(4).Image = 3
                End If
                LoadImageStatus picStatus, GetStatus(StatusId)
            End If
        Case 6 'PRINT PREVIEW
            If SalesOrderId <> 0 Then
                Screen.MousePointer = vbHourglass
                BASE_PrintPreviewFrm.isInvoice = False
                BASE_PrintPreviewFrm.Show
                Dim crxApp As New CRAXDRT.Application
                Dim crxRpt As New CRAXDRT.Report
                Set crxRpt = crxApp.OpenReport(App.path & "\Reports\SO_SalesOrder.rpt")
                crxRpt.EnableParameterPrompting = False
                'crxRpt.RecordSelectionFormula = "{SO_SalesOrder.SalesOrderId}= " & SalesOrderId & ""
                
                crxRpt.DiscardSavedData

                Call ResetRptDB(crxRpt)
                crxRpt.ParameterFields.GetItemByName("@SalesOrderId").AddCurrentValue Int(SalesOrderId)
                
                BASE_PrintPreviewFrm.CRViewer.ReportSource = crxRpt
                BASE_PrintPreviewFrm.CRViewer.ViewReport
                BASE_PrintPreviewFrm.CRViewer.Zoom 1
                Screen.MousePointer = vbDefault
                
                SavePOSAuditTrail UserId, WorkstationId, "", "Generated print preview on sales order: " & txtOrderNumber.Text, "SALES"
            End If
    End Select
End Sub

Private Sub Text1_Change()

End Sub

Private Sub tb_Standard_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "DeliveryReceipt"
            If SalesOrderId <> 0 Then
                Screen.MousePointer = vbHourglass
                BASE_PrintPreviewFrm.Show
                Dim crxApp As New CRAXDRT.Application
                Dim crxRpt As New CRAXDRT.Report
                Set crxRpt = crxApp.OpenReport(App.path & "\Reports\SO_DeliveryReceipt.rpt")
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
    If IsNumeric(txtCash.Text) = False Then
        txtCash.Text = Trim(txtCash.Text)
    Else
        CountTotal
    End If
End Sub

Private Sub txtCash_GotFocus()
    selectText txtCash
End Sub

Private Sub txtCode_Change()
    If Trim(txtCode.Text) = "" Then
        lvItemList.Visible = False
        Exit Sub
    End If
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    Dim item As MSComctlLib.ListItem
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_Product_Search1"
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 500, Null)
    cmd.Parameters.Append cmd.CreateParameter("@ItemCode", adVarChar, adParamInput, 50, txtCode.Text)
    Dim LastProductId As Long
    Set rec = cmd.Execute
    If Not rec.EOF Then
        lvItemList.ListItems.Clear
        Do Until rec.EOF
            If rec!isActive = "True" Then
                If LastProductId <> rec!ProductId Then
                    Set item = lvItemList.ListItems.add(, , rec!ProductId)
                        item.SubItems(1) = rec!itemcode
                        item.SubItems(2) = rec!Name
                        item.SubItems(3) = FormatNumber(rec!unitprice, 2, vbTrue)
                    lvItemList.Visible = True
                    lvItemList.Left = 6070
                    lvItemList.Top = 3360
                    LastProductId = rec!ProductId
                    rec.MoveNext
                Else
                    rec.MoveNext
                End If
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

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp, vbKeyDown
            If lvItemList.Visible = True Then
                lvItemList.SetFocus
            Else
                lvItems.SetFocus
            End If
    End Select
End Sub

Private Sub txtDays_Change()
    If IsNumeric(txtDays.Text) = False Then
        txtDays.Text = ""
    End If
    CountTotal
End Sub

Private Sub txtDays_GotFocus()
    selectText txtDays
End Sub

Private Sub txtDiscount_Change()
    If IsNumeric(txtDiscount.Text) = False Then
        txtDiscount.Text = ""
    End If
    CountTotal
End Sub

Private Sub txtDiscount_GotFocus()
    selectText txtDiscount
End Sub

Private Sub txtInterestRate_Change()
    If IsNumeric(txtInterestRate.Text) = False Then
        txtInterestRate.Text = ""
    End If
    CountTotal
End Sub

Private Sub txtInterestRate_GotFocus()
    selectText txtInterestRate
End Sub

Private Sub txtItemSearch_Change()
    If Trim(txtItemSearch.Text) = "" Then Exit Sub
    btnItemSearch_Click
    GetPricing
End Sub

Private Sub txtItemSearch_GotFocus()
    selectText txtItemSearch
End Sub

Private Sub txtItemSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If txtItemSearch.Text = "" Then Exit Sub
            
'            If Not lvItemList.SelectedItem = -1 Then
'                SO_ItemSelectOptionFrm.Show (1)
'                SO_ItemSelectOptionFrm.txtDescription.Text = lvItemList.SelectedItem.SubItems(2)
'                'SO_ItemSelectOptionFrm.txtPrice.text = lvItemList.SelectedItem.SubItems(3)
'            End If

            On Error Resume Next
            If Trim(txtItemSearch.Text) = "" Then Exit Sub

            Dim item As MSComctlLib.ListItem
            Set rec = New ADODB.Recordset
            Set rec = ProductBarcode(txtItemSearch.Text)

            Dim isExisting As Boolean
            isExisting = False

            'CHECK AVAILABILITY
            Dim Available As Double
            Dim ReserveId As String
            Dim avaialble As Double
            Available = checkAvailableQuantity(rec!ProductId)


            If Not rec.EOF Then 'Item found display in Lvitems
                If isExisting = False Then
                    If AllowNegativeInventory = False Then
                        If Available < 1 Then
                            MsgBox "Insufficient quantity. " & "Remaining quantity: " & Available, vbCritical, "Insufficient Quantity"
                            Exit Sub
                        End If
                    End If

                    'INSERT RESERVE
                    ReserveId = ReserveProduct(0, rec!ProductId, 1, UserId, WorkstationId, False, 2)

                    Set item = lvItems.ListItems.add(, , "")
                    item.SubItems(1) = ""
                    item.SubItems(2) = rec!itemcode 'ItemCode
                    item.SubItems(3) = rec!Name 'Name
                    item.SubItems(4) = "1.00"
                    item.SubItems(5) = rec!Uom
                    item.SubItems(6) = FormatNumber(rec!unitprice, 2, vbTrue, vbFalse)
                    item.SubItems(8) = ""
                    item.SubItems(9) = rec!ProductId
                    item.SubItems(13) = 1
                    item.SubItems(14) = ReserveId
                Else
                    If Available + (NVAL(item.SubItems(4)) * NVAL(item.SubItems(13))) < (NVAL(item.SubItems(4)) * NVAL(item.SubItems(13))) + NVAL(item.SubItems(13)) Then
                        MsgBox "Insufficient quantity. " & "Remaining quantity: " & Available, vbCritical, "Insufficient Quantity"
                        Exit Sub
                    Else
                        'UPDATE RESERVE
                        ReserveId = ReserveProduct(item.SubItems(14), item.SubItems(9), (NVAL(item.SubItems(4)) + 1) * item.SubItems(13), UserId, WorkstationId, False, 2)
                        item.SubItems(4) = FormatNumber(Val(Replace(item.SubItems(4), ",", "")) + 1, 2, vbTrue, vbFalse)
                    End If
                End If

                CountTotal
                selectText txtItemSearch
            Else
                MsgBox "Item not found.", vbCritical, "Not Found"
                selectText txtItemSearch
            End If
            GetPricing
            
        Case vbKeyUp, vbKeyDown
            If lvItemList.Visible = True Then
                lvItemList.SetFocus
            Else
                lvItems.SetFocus
            End If
    End Select
End Sub

Private Sub txtSearch_Name_Change()
    btnSearch_Click
End Sub

Private Sub txtSearch_OrderNumber_Change()
    btnSearch_Click
End Sub

Private Sub txtSearch_OrderNumber_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then btnSearch_Click
End Sub
