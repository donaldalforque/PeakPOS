VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form SO_InvoiceFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10410
   Icon            =   "SO_InvoiceFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   10410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnReceiveOrder 
      Caption         =   "Auto Fill"
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
      Left            =   6840
      TabIndex        =   28
      Top             =   2520
      Width           =   3405
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
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
            Picture         =   "SO_InvoiceFrm.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_InvoiceFrm.frx":686E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_InvoiceFrm.frx":D0D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_InvoiceFrm.frx":13932
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_InvoiceFrm.frx":13BA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_InvoiceFrm.frx":14218
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tb_Standard 
      Height          =   330
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   582
      ButtonWidth     =   1376
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
            Object.Visible         =   0   'False
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
   Begin VB.Frame Body_Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   9135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      Begin VB.CommandButton btnSalesReturns 
         Caption         =   "Add Sales Returns"
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
         Left            =   240
         TabIndex        =   40
         Top             =   1920
         Width           =   2805
      End
      Begin VB.Frame Frame_Header2 
         BackColor       =   &H00FFFFFF&
         Height          =   1515
         Left            =   7080
         TabIndex        =   29
         Top             =   480
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
            TabIndex        =   30
            Top             =   960
            Width           =   2055
         End
         Begin MSComCtl2.DTPicker dtOrder 
            Height          =   330
            Left            =   960
            TabIndex        =   31
            Top             =   240
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
            Format          =   112656385
            CurrentDate     =   41509
         End
         Begin MSComCtl2.DTPicker dtDue 
            Height          =   330
            Left            =   960
            TabIndex        =   32
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
            Format          =   112656385
            CurrentDate     =   41509
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
            Top             =   240
            Width           =   435
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
            TabIndex        =   34
            Top             =   600
            Width           =   375
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
            TabIndex        =   33
            Top             =   960
            Width           =   465
         End
      End
      Begin VB.Frame Frame_Body 
         BackColor       =   &H00FFFFFF&
         Height          =   4380
         Left            =   120
         TabIndex        =   20
         Top             =   2325
         Width           =   10215
         Begin MSComctlLib.ListView lvItems 
            Height          =   3495
            Left            =   120
            TabIndex        =   21
            Top             =   720
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   6165
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
               Text            =   "InvoiceLineId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "InvoiceId"
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
               Alignment       =   1
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
               SubItemIndex    =   8
               Text            =   "ProductId"
               Object.Width           =   0
            EndProperty
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
            TabIndex        =   24
            Top             =   240
            Visible         =   0   'False
            Width           =   3015
         End
         Begin VB.CommandButton btnItemSearch 
            Height          =   330
            Left            =   4320
            Picture         =   "SO_InvoiceFrm.frx":1AA7A
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   240
            Visible         =   0   'False
            Width           =   375
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
            TabIndex        =   22
            Top             =   600
            Visible         =   0   'False
            Width           =   3015
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
            TabIndex        =   26
            Top             =   240
            Visible         =   0   'False
            Width           =   555
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
            TabIndex        =   25
            Top             =   600
            Visible         =   0   'False
            Width           =   480
         End
      End
      Begin VB.Frame Frame_Footer 
         BackColor       =   &H00FFFFFF&
         Height          =   2295
         Left            =   120
         TabIndex        =   4
         Top             =   6720
         Width           =   10215
         Begin VB.TextBox txtRefunds 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
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
            Left            =   8760
            TabIndex        =   9
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox txtDiscountPercent 
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
            Left            =   8760
            TabIndex        =   37
            Top             =   600
            Width           =   1335
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
            Left            =   8760
            TabIndex        =   8
            Top             =   960
            Width           =   1335
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
            TabIndex        =   7
            Top             =   240
            Width           =   5295
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
            TabIndex        =   6
            Top             =   240
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.ComboBox cmbPricing 
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
            ItemData        =   "SO_InvoiceFrm.frx":1AC9E
            Left            =   -99999
            List            =   "SO_InvoiceFrm.frx":1ACAE
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1680
            Visible         =   0   'False
            Width           =   3135
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Refunds"
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
            TabIndex        =   39
            Top             =   1320
            Width           =   795
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Others (%)"
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
            TabIndex        =   38
            Top             =   600
            Width           =   1035
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
            TabIndex        =   19
            Top             =   645
            Visible         =   0   'False
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1200
            TabIndex        =   18
            Top             =   645
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
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
            Left            =   7200
            TabIndex        =   17
            Top             =   1755
            Width           =   600
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
            TabIndex        =   16
            Top             =   240
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   7200
            TabIndex        =   15
            Top             =   240
            Width           =   1020
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
            TabIndex        =   14
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
            TabIndex        =   13
            Top             =   240
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Others"
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
            TabIndex        =   12
            Top             =   960
            Width           =   660
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
            Left            =   -99999
            TabIndex        =   11
            Top             =   1680
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
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
            Left            =   7995
            TabIndex        =   10
            Top             =   1755
            Width           =   2085
         End
      End
      Begin VB.Frame Frame_Header1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   735
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   6735
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Invoice"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   435
            Left            =   120
            TabIndex        =   36
            Top             =   120
            Width           =   1890
         End
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
         TabIndex        =   1
         Top             =   3000
         Visible         =   0   'False
         Width           =   3015
      End
      Begin MSComctlLib.ListView lvCustomer 
         Height          =   2655
         Left            =   -9999
         TabIndex        =   2
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
End
Attribute VB_Name = "SO_InvoiceFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public InvoiceId As String
Dim OrderLine(10000) As Long
Dim ctrOrderLine As Long
Public Sub Populate(ByVal data As String)
    Select Case data
        Case "InvoiceGet"
            Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            Set rec = New ADODB.Recordset
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "SO_Invoice_Get"
            cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SO_SalesOrderFrm.SalesOrderId)
            Set rec = cmd.Execute
            If Not rec.EOF Then
                dtOrder.value = Format(rec!Date, "MM/DD/YY")
                dtDue.value = Format(rec!DueDate, "MM/DD/YY")
                lblSubTotal.Caption = FormatNumber(rec!subtotal, 2, vbTrue)
                lblTotal.Caption = FormatNumber(rec!Total, 2, vbTrue, vbFalse)
                On Error Resume Next
                txtReferenceNumber.Text = rec!ReferenceNumber
                txtRefunds.Text = FormatNumber(rec!refunds, 2, vbTrue, vbFalse)
                If IsNull(rec!discount) = True Then
                    txtDiscount.Text = ""
                Else
                    txtDiscount.Text = FormatNumber(rec!discount, 2, vbTrue, vbFalse)
                End If
                InvoiceId = rec!InvoiceId
            End If
            con.Close
            
        Case "InvoiceLineGet"
            Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            Set rec = New ADODB.Recordset
            Dim item As MSComctlLib.ListItem
            
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "SO_InvoiceLine_Get"
            cmd.Parameters.Append cmd.CreateParameter("@InvoiceId", adInteger, adParamInput, , Val(InvoiceId))
            Set rec = cmd.Execute
            lvItems.ListItems.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    Set item = lvItems.ListItems.add(, , rec!InvoiceLineId)
                        item.SubItems(1) = rec!InvoiceId
                        item.SubItems(2) = rec!itemcode
                        item.SubItems(3) = rec!Name
                        item.SubItems(4) = FormatNumber(rec!quantity, 2, vbTrue)
                        item.SubItems(5) = rec!Uom
                        item.SubItems(6) = FormatNumber(rec!price, 2, vbTrue)
                        item.SubItems(7) = FormatNumber(rec!subtotal, 2, vbTrue)
                        'item.SubItems(8) = rec!LocationId
                        item.SubItems(8) = rec!ProductId
                    rec.MoveNext
                Loop
            End If
            con.Close
    End Select
End Sub
Public Sub CountTotal()
    Dim Total, subtotal, discount, subtotal1, interestrate, Interest, days, cash As Double
    Dim item As MSComctlLib.ListItem
    subtotal1 = 0
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
    'subtotal1 = subtotal1
    lblSubTotal.Caption = FormatNumber(subtotal1, 2, vbTrue, vbFalse)
    
    'COMPUTE INTEREST
    'Interest = ((days / 30) * (subtotal1 * (interestrate / 100)))
    'lblInterest.Caption = FormatNumber(Interest, 2, vbTrue, vbFalse)
    
    Total = subtotal1 - discount - NVAL(txtRefunds.Text)
    lblTotal.Caption = FormatNumber(Total, 2, vbTrue, vbFalse)
    'lblRemainingBalance.Caption = lblTotal.Caption
End Sub
Private Sub Label4_Click()

End Sub

Private Sub lblInvoice_Click()

End Sub

Private Sub btnReceiveOrder_Click()
    Set con = New ADODB.Connection
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    
    Dim item As MSComctlLib.ListItem
    Dim itemx As MSComctlLib.ListItem
    
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "SO_Invoice_AutoFill"
    cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SO_SalesOrderFrm.SalesOrderId)
    Set rec = cmd.Execute
    'lvItems.ListItems.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            Dim isFound As Boolean
            'Check if Current Record exists in the list
'            For Each itemx In lvItems.ListItems
'                If itemx.SubItems(8) = rec!ProductId And itemx.SubItems(5) = rec!Uom Then
'                    isFound = True
'                    'itemx.SubItems(4) = FormatNumber(rec!pickedquantity, 2, vbTrue)
'                    'Exit For
'                End If
'            Next
            
            If isFound = False Then
                Set item = lvItems.ListItems.add(, , "")
                    item.SubItems(1) = ""
                    item.SubItems(2) = rec!itemcode
                    item.SubItems(3) = rec!Name
                    item.SubItems(4) = FormatNumber(rec!pickedquantity, 2, vbTrue)
                    item.SubItems(5) = rec!Uom
                    item.SubItems(6) = FormatNumber(rec!price, 2, vbTrue)
                    item.SubItems(7) = FormatNumber(rec!subtotal, 2, vbTrue)
                    'item.SubItems(8) = ""
                    item.SubItems(8) = rec!ProductId
            End If
            rec.MoveNext
        Loop
    End If
    con.Close
    CountTotal
End Sub

Private Sub Command1_Click()
    
End Sub

Private Sub btnSalesReturns_Click()
    SO_InvoiceSalesReturnFrm.cmbCustomer.Text = SO_SalesOrderFrm.cmbCustomer.Text
    SO_InvoiceSalesReturnFrm.Show (1)
End Sub

Private Sub Form_Load()
    lvItems.ColumnHeaders(3).width = lvItems.width * 0.14
    lvItems.ColumnHeaders(4).width = lvItems.width * 0.29
    lvItems.ColumnHeaders(5).width = lvItems.width * 0.12
    lvItems.ColumnHeaders(5).width = lvItems.width * 0.12
    lvItems.ColumnHeaders(6).width = lvItems.width * 0.09
    lvItems.ColumnHeaders(7).width = lvItems.width * 0.13
    lvItems.ColumnHeaders(8).width = lvItems.width * 0.2
    
    dtOrder.value = SO_SalesOrderFrm.dtOrder.value
    dtDue.value = SO_SalesOrderFrm.dtDue.value
    txtReferenceNumber.Text = SO_SalesOrderFrm.txtReferenceNumber.Text
    
    InvoiceId = "0"
    
    Populate "InvoiceGet"
    Populate "InvoiceLineGet"
        
    CountTotal
    
    On Error Resume Next
    txtDiscountPercent.Text = Val(Replace(txtDiscount.Text, ",", "")) / Val(Replace(lblSubTotal.Caption, ",", "")) * 100
    txtRemarks.Text = SO_SalesOrderFrm.global_remarks
    
    SO_InvoiceSalesReturnFrm.Show
    SO_InvoiceSalesReturnFrm.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload SO_InvoiceSalesReturnFrm
End Sub

Private Sub lvItems_DblClick()
    If lvItems.ListItems.count > 0 Then
        With SO_InvoiceDialogFrm
            .txtQuantity.Text = lvItems.SelectedItem.SubItems(4)
            .txtPrice.Text = lvItems.SelectedItem.SubItems(6)
            
            .isModify = True
            .Show (1)
        End With
    End If
End Sub

Private Sub lvItems_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
'        Case vbKeyUp
'            If lvItems.ListItems.Count > 0 Then
'                If lvItems.SelectedItem.Index = 1 Then
'                    txtItemSearch.SetFocus
'                    'txtCode.SetFocus
'                End If
'            End If
        Case vbKeyDelete
            If lvItems.ListItems.count > 0 Then
                If lvItems.SelectedItem.Text <> "" Then
                    OrderLine(ctrOrderLine) = Val(lvItems.SelectedItem.Text)
                    ctrOrderLine = ctrOrderLine + 1
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

Private Sub tb_Standard_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 2 'SAVE
            If EditAccessRights(32) = False Then
                MsgBox ErrorCodes(74), vbCritical, "Limited Rights"
                Exit Sub
            End If
            If SO_SalesOrderFrm.StatusId = 7 Then
                GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(64)
                GLOBAL_MessageFrm.Show (1)
                Exit Sub
            End If
            
            If IsNumeric(txtDiscountPercent.Text) = False Then
                txtDiscountPercent.Text = 0
            End If
            
            Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            
            con.ConnectionString = ConnString
            con.Open
            con.BeginTrans
            
            'SAVE INVOICE
            cmd.CommandType = adCmdStoredProc
            cmd.ActiveConnection = con
            cmd.Parameters.Append cmd.CreateParameter("@InvoiceId", adInteger, adParamInputOutput, , Val(InvoiceId))
            cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SO_SalesOrderFrm.SalesOrderId)
            cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtOrder.value)
            cmd.Parameters.Append cmd.CreateParameter("@DueDate", adDate, adParamInput, , dtDue.value)
            cmd.Parameters.Append cmd.CreateParameter("@Discount", adDecimal, adParamInput, , Val(Replace(txtDiscount.Text, ",", "")))
                                  cmd.Parameters("@Discount").Precision = 18
                                  cmd.Parameters("@Discount").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@Refunds", adDecimal, adParamInput, , Val(Replace(txtRefunds.Text, ",", "")))
                                  cmd.Parameters("@Refunds").Precision = 18
                                  cmd.Parameters("@Refunds").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@SubTotal", adDecimal, adParamInput, , Val(Replace(lblSubTotal.Caption, ",", "")))
                                  cmd.Parameters("@SubTotal").Precision = 18
                                  cmd.Parameters("@SubTotal").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@Total", adDecimal, adParamInput, , Val(Replace(lblTotal.Caption, ",", "")))
                                  cmd.Parameters("@Total").Precision = 18
                                  cmd.Parameters("@Total").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@ReferenceNumber", adVarChar, adParamInput, 50, txtReferenceNumber.Text)
            cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, txtRemarks.Text)
            
            If Val(InvoiceId) = 0 Then
                cmd.CommandText = "SO_Invoice_Insert"
                cmd.Execute
                InvoiceId = cmd.Parameters("@InvoiceId")
                
                SavePOSAuditTrail UserId, WorkstationId, "", "Invoiced sales order ref: " & SO_SalesOrderFrm.txtOrderNumber.Text, "SALES"
            Else
                cmd.CommandText = "SO_Invoice_Update"
                cmd.Execute
                
                SavePOSAuditTrail UserId, WorkstationId, "", "Updated invoice sales order ref: " & SO_SalesOrderFrm.txtOrderNumber.Text, "SALES"
            End If
            
            
            
            'Update Status to INVOICED
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "SO_SalesOrderStatus_Update"
            cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SO_SalesOrderFrm.SalesOrderId)
            cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , 4) 'Invoiced
            cmd.Execute
            
            'SAVE LINE
            Dim item As MSComctlLib.ListItem
            For Each item In lvItems.ListItems
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                
                cmd.Parameters.Append cmd.CreateParameter("@InvoiceLineId", adInteger, adParamInputOutput, , Val(item.Text))
                cmd.Parameters.Append cmd.CreateParameter("@InvoiceId", adInteger, adParamInput, , InvoiceId)
                cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtOrder.value)
                cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , Val(item.SubItems(8)))
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
                If item.Text = "" Then
                    cmd.CommandText = "SO_InvoiceLine_Insert"
                Else
                    cmd.CommandText = "SO_InvoiceLine_Update"
                End If
                cmd.Execute
                item.Text = cmd.Parameters("@InvoiceLineId")
            Next
            
            'UPDATE RESERVERS
            
            
            'UPDATE SO REMAINING BALANCE
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "SO_Balance_Update"
            cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SO_SalesOrderFrm.SalesOrderId)
            cmd.Execute
            
            
            'INSERT LEDGER
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandText = "FIN_CustomerLedger_Insert"
            cmd.CommandType = adCmdStoredProc
        
            cmd.Parameters.Append cmd.CreateParameter("@SOPaymentId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInput, , SO_SalesOrderFrm.CustomerId)
            cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SO_SalesOrderFrm.SalesOrderId)
            cmd.Parameters.Append cmd.CreateParameter("@Particulars", adVarChar, adParamInput, 250, "Sales")
            cmd.Parameters.Append cmd.CreateParameter("@Debit", adDecimal, adParamInput, , Val(Replace(lblTotal.Caption, ",", "")))
                                  cmd.Parameters("@Debit").Precision = 18
                                  cmd.Parameters("@Debit").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@Credit", adDecimal, adParamInput, , Null)
                                  cmd.Parameters("@Credit").Precision = 18
                                  cmd.Parameters("@Credit").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtOrder.value)
            cmd.Parameters.Append cmd.CreateParameter("@DueDate", adDate, adParamInput, , dtDue.value)
            cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInput, , Null)
            cmd.Execute
            
            
            With SO_SalesOrderFrm
                .StatusId = 4
                'On Error Resume Next
                .lvSearch.SelectedItem.SubItems(3) = "Invoiced"
                .txtStatus.Text = "Invoiced"
            End With
            
            'DELETE ORDERLINE IF ANY
            Dim ctr As Integer
            For ctr = 0 To ctrOrderLine
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
    
                If OrderLine(ctr) <> 0 Then
                    cmd.Parameters.Append cmd.CreateParameter("@InvoiceLineId", adInteger, adParamInput, , OrderLine(ctr))
                    cmd.CommandText = "SO_InvoiceLine_Delete"
                    cmd.Execute
                Else
                    Exit For
                End If
            Next
            
            
            'UPDATE SALESRETURNSTATUS
            For Each item In SO_InvoiceSalesReturnFrm.lvModules.ListItems
                If item.Checked = True Then
                    Set cmd = New ADODB.Command
                    cmd.ActiveConnection = con
                    cmd.CommandType = adCmdStoredProc
                    cmd.Parameters.Append cmd.CreateParameter("@SalesReturnId", adInteger, adParamInput, , NVAL(item.SubItems(5)))
                    cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , NVAL(item.SubItems(6)))
                    cmd.CommandText = "SO_SalesReturnStatus_Update"
                    cmd.Execute
                End If
            Next
            
            con.CommitTrans
            con.Close
            
            MsgBox "Invoice saved.", vbInformation
            
            LoadImageStatus SO_SalesOrderFrm.picStatus, GetStatus(SO_SalesOrderFrm.StatusId)
            UpdateCustomerOrderDues
        Case 6 'PRINT
            If InvoiceId <> 0 Then
                Screen.MousePointer = vbHourglass
                BASE_PrintPreviewFrm.isInvoice = True
                BASE_PrintPreviewFrm.Show '(1)
                Dim crxApp As New CRAXDRT.Application
                Dim crxRpt As New CRAXDRT.Report
                Set crxRpt = crxApp.OpenReport(App.path & "\Reports\SO_SalesInvoice.rpt")
                crxRpt.RecordSelectionFormula = "{SO_Invoice.InvoiceId}= " & Val(InvoiceId) & ""
                crxRpt.DiscardSavedData

                Call ResetRptDB(crxRpt)

                BASE_PrintPreviewFrm.CRViewer.ReportSource = crxRpt
                BASE_PrintPreviewFrm.CRViewer.ViewReport
                BASE_PrintPreviewFrm.CRViewer.Zoom 1
                Screen.MousePointer = vbDefault
                
                SavePOSAuditTrail UserId, WorkstationId, "", "Generated print preview on sales invoice: " & SO_SalesOrderFrm.txtOrderNumber.Text, "SALES"
            End If
    End Select
End Sub

Private Sub txtDiscount_Change()
    If IsNumeric(txtDiscount.Text) = False Then
        txtDiscount.Text = ""
    End If
    CountTotal
End Sub

Private Sub txtDiscountPercent_Change()
    If IsNumeric(txtDiscountPercent.Text) = False Then
        'txtDiscountPercent.text = ""
        'txtDiscount.text = ""
    Else
        'compute percentage
        Dim discounted As Double
        discounted = (Val(Replace(lblSubTotal.Caption, ",", "")) * Val(Replace(txtDiscountPercent.Text, ",", ""))) / 100
        txtDiscount.Text = FormatNumber(discounted, 2, vbTrue, vbFalse)
    End If
    CountTotal
End Sub

Private Sub txtDiscountPercent_LostFocus()
    If IsNumeric(txtDiscountPercent.Text) = False Then
        txtDiscountPercent.Text = Val(txtDiscountPercent.Text)
    End If
End Sub
