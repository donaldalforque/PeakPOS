VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form PO_PurchaseInvoiceFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14940
   Icon            =   "PO_PurchaseInvoiceFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   14940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame LayoutFrame_Search 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   -120
      TabIndex        =   57
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
         TabIndex        =   12
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
         TabIndex        =   13
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
         TabIndex        =   16
         Top             =   2040
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvSearch 
         Height          =   6255
         Left            =   195
         TabIndex        =   17
         Top             =   2520
         Width           =   4230
         _ExtentX        =   7461
         _ExtentY        =   11033
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
            Text            =   "Invoice #"
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
         TabIndex        =   15
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
         Format          =   112001025
         CurrentDate     =   41686
      End
      Begin MSComCtl2.DTPicker DateFrom 
         Height          =   345
         Left            =   1440
         TabIndex        =   14
         Top             =   1200
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
         Format          =   112001025
         CurrentDate     =   41686
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice #"
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
         TabIndex        =   62
         Top             =   480
         Width           =   825
      End
      Begin VB.Label Label15 
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
         Left            =   240
         TabIndex        =   61
         Top             =   75
         Width           =   795
      End
      Begin VB.Label Label17 
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
         TabIndex        =   60
         Top             =   840
         Width           =   570
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
         TabIndex        =   59
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
         TabIndex        =   58
         Top             =   1560
         Width           =   705
      End
   End
   Begin VB.Frame Body_Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   4500
      TabIndex        =   18
      Top             =   0
      Width           =   10455
      Begin VB.PictureBox picStatus 
         BorderStyle     =   0  'None
         Height          =   1055
         Left            =   3120
         Picture         =   "PO_PurchaseInvoiceFrm.frx":000C
         ScaleHeight     =   1050
         ScaleWidth      =   3750
         TabIndex        =   74
         Top             =   2160
         Visible         =   0   'False
         Width           =   3755
      End
      Begin VB.CommandButton btnSalesReturns 
         Caption         =   "Add Purchase Returns"
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
         Left            =   120
         TabIndex        =   73
         Top             =   2640
         Width           =   2805
      End
      Begin MSComctlLib.ListView lvDelivery 
         Height          =   2415
         Left            =   5640
         TabIndex        =   5
         Top             =   2740
         Visible         =   0   'False
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   4260
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
            Text            =   "ReceiveOrderId"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Delivery #"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Frame Frame_Header1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   495
         Left            =   240
         TabIndex        =   46
         Top             =   600
         Width           =   4335
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Purchase Invoice"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   435
            Left            =   0
            TabIndex        =   47
            Top             =   0
            Width           =   2475
         End
      End
      Begin VB.Frame Frame_Footer 
         BackColor       =   &H00FFFFFF&
         Height          =   2655
         Left            =   120
         TabIndex        =   29
         Top             =   6120
         Width           =   10215
         Begin VB.TextBox txtFreight 
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
            TabIndex        =   10
            Top             =   1320
            Width           =   1335
         End
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
            Left            =   -9999
            TabIndex        =   54
            Top             =   2040
            Visible         =   0   'False
            Width           =   3405
         End
         Begin VB.TextBox txtVAT 
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
            TabIndex        =   9
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtFees 
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
            TabIndex        =   31
            Top             =   2040
            Visible         =   0   'False
            Width           =   1335
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
            ItemData        =   "PO_PurchaseInvoiceFrm.frx":7D7E
            Left            =   -9999
            List            =   "PO_PurchaseInvoiceFrm.frx":7D8E
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   1680
            Visible         =   0   'False
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
            TabIndex        =   32
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
            Height          =   1650
            Left            =   1200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   240
            Width           =   5295
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
            Top             =   600
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
            Height          =   390
            Left            =   -9999
            TabIndex        =   30
            Top             =   600
            Visible         =   0   'False
            Width           =   1335
         End
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
            TabIndex        =   11
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label lblVatpercent 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   8475
            MouseIcon       =   "PO_PurchaseInvoiceFrm.frx":7DCC
            MousePointer    =   99  'Custom
            TabIndex        =   64
            Top             =   960
            Width           =   180
         End
         Begin VB.Label lblDiscountPercent 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   8475
            MouseIcon       =   "PO_PurchaseInvoiceFrm.frx":7F1E
            MousePointer    =   99  'Custom
            TabIndex        =   63
            Top             =   600
            Width           =   180
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Freight"
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
            Left            =   6960
            TabIndex        =   55
            Top             =   1320
            Width           =   675
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tax (VAT)"
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
            Left            =   6960
            TabIndex        =   53
            Top             =   960
            Width           =   945
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Other Deductions"
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
            Left            =   -9999
            TabIndex        =   52
            Top             =   2400
            Visible         =   0   'False
            Width           =   1710
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
            TabIndex        =   45
            Top             =   2100
            Width           =   2085
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
            Left            =   -9999
            TabIndex        =   44
            Top             =   1680
            Visible         =   0   'False
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
            Left            =   6960
            TabIndex        =   43
            Top             =   600
            Width           =   855
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
            TabIndex        =   42
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
            TabIndex        =   41
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
            Left            =   6960
            TabIndex        =   40
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
            TabIndex        =   39
            Top             =   240
            Width           =   1620
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
            Left            =   6960
            TabIndex        =   38
            Top             =   2100
            Width           =   600
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
            TabIndex        =   37
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
            TabIndex        =   36
            Top             =   645
            Visible         =   0   'False
            Width           =   1620
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
            Left            =   -9999
            TabIndex        =   35
            Top             =   600
            Visible         =   0   'False
            Width           =   1035
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
            Left            =   6960
            TabIndex        =   34
            Top             =   1680
            Width           =   795
         End
      End
      Begin VB.Frame Frame_Body 
         BackColor       =   &H00FFFFFF&
         Height          =   3030
         Left            =   120
         TabIndex        =   23
         Top             =   3045
         Width           =   10215
         Begin MSComctlLib.ListView lvItems 
            Height          =   2655
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   4683
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
            NumItems        =   10
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
               Text            =   "Cost"
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
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "ReceiveOrderLineId"
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
            Left            =   -9999
            TabIndex        =   26
            Top             =   600
            Visible         =   0   'False
            Width           =   3015
         End
         Begin VB.CommandButton btnItemSearch 
            Height          =   330
            Left            =   -9999
            Picture         =   "PO_PurchaseInvoiceFrm.frx":8070
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   240
            Visible         =   0   'False
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
            Left            =   -9999
            TabIndex        =   24
            Top             =   240
            Visible         =   0   'False
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
            Left            =   -9999
            TabIndex        =   28
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
            Left            =   -9999
            TabIndex        =   27
            Top             =   240
            Visible         =   0   'False
            Width           =   555
         End
      End
      Begin VB.Frame Frame_Header2 
         BackColor       =   &H00FFFFFF&
         Height          =   2535
         Left            =   6210
         TabIndex        =   19
         Top             =   360
         Width           =   4125
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
            Left            =   1320
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   1680
            Width           =   2655
         End
         Begin VB.ComboBox cmbReferenceNumber 
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
            TabIndex        =   4
            Top             =   2040
            Width           =   2655
         End
         Begin VB.TextBox txtInvoiceNumber 
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
            TabIndex        =   0
            Top             =   240
            Width           =   2655
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
            ItemData        =   "PO_PurchaseInvoiceFrm.frx":8294
            Left            =   1320
            List            =   "PO_PurchaseInvoiceFrm.frx":8296
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   960
            Width           =   2655
         End
         Begin MSComCtl2.DTPicker dtOrder 
            Height          =   330
            Left            =   1320
            TabIndex        =   1
            Top             =   600
            Width           =   2655
            _ExtentX        =   4683
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
            Format          =   112001025
            CurrentDate     =   41509
         End
         Begin MSComCtl2.DTPicker dtDue 
            Height          =   330
            Left            =   1320
            TabIndex        =   3
            Top             =   1320
            Width           =   2655
            _ExtentX        =   4683
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
            Format          =   112001025
            CurrentDate     =   41509
         End
         Begin VB.Label Label18 
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
            TabIndex        =   72
            Top             =   1680
            Width           =   570
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice #"
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
            Top             =   240
            Width           =   825
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Terms"
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
            TabIndex        =   51
            Top             =   960
            Width           =   570
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DR/Ref #"
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
            TabIndex        =   22
            Top             =   2040
            Width           =   810
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
            TabIndex        =   21
            Top             =   1320
            Width           =   375
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
            TabIndex        =   20
            Top             =   600
            Width           =   435
         End
      End
      Begin MSComctlLib.ListView lvCustomer 
         Height          =   2655
         Left            =   -9999
         TabIndex        =   49
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
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4800
         Top             =   480
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
               Picture         =   "PO_PurchaseInvoiceFrm.frx":8298
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PO_PurchaseInvoiceFrm.frx":EAFA
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PO_PurchaseInvoiceFrm.frx":1535C
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PO_PurchaseInvoiceFrm.frx":1BBBE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PO_PurchaseInvoiceFrm.frx":1BE33
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "PO_PurchaseInvoiceFrm.frx":1C4A4
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tb_Standard 
         Height          =   330
         Left            =   0
         TabIndex        =   50
         Top             =   0
         Width           =   13215
         _ExtentX        =   23310
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
            EndProperty
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
         TabIndex        =   48
         Top             =   3000
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label lblPORef 
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
         Left            =   1320
         TabIndex        =   70
         Top             =   1560
         Width           =   4680
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO Ref #:"
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
         Left            =   360
         TabIndex        =   69
         Top             =   1560
         Width           =   840
      End
      Begin VB.Label lblSupplierName 
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
         Left            =   1320
         TabIndex        =   68
         Top             =   1200
         Width           =   4680
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier:"
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
         Left            =   360
         TabIndex        =   67
         Top             =   1200
         Width           =   840
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier:"
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
         Left            =   100
         TabIndex        =   66
         Top             =   0
         Width           =   840
      End
      Begin VB.Label lblSupplier 
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
         Height          =   855
         Left            =   100
         TabIndex        =   65
         Top             =   0
         Width           =   4560
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "PO_PurchaseInvoiceFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim InvoiceId As Long
Dim OrderLine(10000) As Long
Dim ctrOrderLine As Long
Public PurchaseOrderId As Long
Public ReceiveOrderId As Long
Public StatusId As Long

Public Sub Initialize()
    ReceiveOrderId = 0
    InvoiceId = 0
    StatusId = 1
    lvDelivery.Visible = False
    
    lvItems.ListItems.Clear
    
    Dim txtControl As Control
    For Each txtControl In Me.Controls
        If TypeOf txtControl Is TextBox And txtControl.Name <> "txtSearch_Order" And txtControl.Name <> "txtSearch_ReferenceNumber" Then
            txtControl.Text = ""
            txtStatus.Text = "Open"
        End If
    Next
    cmbReferenceNumber.Text = ""
    
    On Error Resume Next
    txtInvoiceNumber.SetFocus
End Sub
Private Function Validated() As Boolean
    If txtInvoiceNumber.Text = "" Then
        Validated = False
        GLOBAL_MessageFrm.lblErrorMessage.Caption = "Invoice number is required."
        GLOBAL_MessageFrm.Show (1)
        txtInvoiceNumber.SetFocus
    ElseIf ReceiveOrderId = 0 Then
        Validated = False
        GLOBAL_MessageFrm.lblErrorMessage.Caption = "Please select a delivery receipt to invoice."
        GLOBAL_MessageFrm.Show (1)
        cmbReferenceNumber.SetFocus
    Else
        Validated = True
    End If
End Function
Public Sub Populate(ByVal data As String)
    Select Case data
        Case "Terms"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Terms")
            cmbTerms.Clear
            cmbTerms.AddItem ""
            cmbTerms.ItemData(cmbTerms.NewIndex) = 0
            If Not rec.EOF Then
                Do Until rec.EOF
                    cmbTerms.AddItem rec!Terms
                    cmbTerms.ItemData(cmbTerms.NewIndex) = rec!TermId
                    rec.MoveNext
                Loop
            End If
            cmbTerms.ListIndex = 0
        Case "Status"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data(data)
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
        Case "InvoiceGet"
            Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            Set rec = New ADODB.Recordset
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "PO_PurchaseInvoice_Get"
            cmd.Parameters.Append cmd.CreateParameter("@PurchaseInvoiceId", adInteger, adParamInput, , InvoiceId)
            Set rec = cmd.Execute
            If Not rec.EOF Then
                StatusId = rec!StatusId
                ReceiveOrderId = rec!ReceiveOrderId
                txtStatus.Text = rec!Status
                dtOrder.value = Format(rec!Date, "MM/DD/YY")
                dtDue.value = Format(rec!DueDate, "MM/DD/YY")
                If rec!Terms = "" Then cmbTerms.ListIndex = 0 Else cmbTerms.Text = rec!Terms
                txtInvoiceNumber.Text = rec!OrderNumber
                lblSubTotal.Caption = FormatNumber(rec!subtotal, 2, vbTrue)
                lblTotal.Caption = FormatNumber(rec!Total, 2, vbTrue, vbFalse)
                cmbReferenceNumber.Text = rec!ReferenceNumber
                txtRefunds.Text = FormatNumber(rec!refunds, 2, vbTrue, vbFalse)
                txtFreight.Text = FormatNumber(rec!freight, 2, vbTrue, vbFalse)
                txtVAT.Text = FormatNumber(rec!tax, 2, vbTrue, vbFalse)
                txtRemarks.Text = rec!Remarks
                If IsNull(rec!discount) = True Then txtDiscount.Text = "" Else txtDiscount.Text = FormatNumber(rec!discount, 2, vbTrue, vbFalse)
                InvoiceId = rec!PurchaseInvoiceId
            End If
            con.Close
            
        Case "InvoiceLineGet"
            Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            Set rec = New ADODB.Recordset
            Dim Item As MSComctlLib.ListItem
            
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "PO_PurchaseInvoiceLine_Get"
            cmd.Parameters.Append cmd.CreateParameter("@InvoiceId", adInteger, adParamInput, , Val(InvoiceId))
            Set rec = cmd.Execute
            lvItems.ListItems.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    Set Item = lvItems.ListItems.add(, , rec!PurchaseInvoiceLineId)
                        Item.SubItems(1) = rec!PurchaseInvoiceId
                        Item.SubItems(2) = rec!itemcode
                        Item.SubItems(3) = rec!Name
                        Item.SubItems(4) = FormatNumber(rec!quantity, 2, vbTrue)
                        Item.SubItems(5) = rec!Uom
                        Item.SubItems(6) = FormatNumber(rec!cost, 2, vbTrue)
                        Item.SubItems(7) = FormatNumber(rec!subtotal, 2, vbTrue)
                        Item.SubItems(8) = rec!ProductId
                    rec.MoveNext
                Loop
            End If
            con.Close
        Case "DeliveryReceipt"
            Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            Set rec = New ADODB.Recordset
            
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "PO_ReceiveOrder_Get"
            cmd.Parameters.Append cmd.CreateParameter("@ReceiveOrderId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , PurchaseOrderId)
            cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , 2) 'SELECT ALL COMPLETED DELIVERY
            
            Set rec = cmd.Execute
            If Not rec.EOF Then
                lvDelivery.ListItems.Clear
                Do Until rec.EOF
                    Set Item = lvDelivery.ListItems.add(, , rec!ReceiveOrderId)
                        Item.SubItems(1) = rec!DeliveryNumber
                        Item.SubItems(2) = Format(rec!receiveddate, "mm/dd/yy")
                    rec.MoveNext
                Loop
            End If
            con.Close
    End Select
End Sub
Public Sub CountTotal()
    Dim Total, subtotal, discount, subtotal1, interestrate, Interest, days, cash As Double
    Dim Item As MSComctlLib.ListItem
    subtotal1 = 0
    If IsNumeric(txtDiscount.Text) = False Then
        discount = 0
    Else
        discount = txtDiscount.Text
    End If
    
    For Each Item In lvItems.ListItems
        subtotal = Val(Replace(Item.SubItems(4), ",", "")) * Val(Replace(Item.SubItems(6), ",", ""))
        Item.SubItems(7) = FormatNumber(subtotal, 2, vbTrue, vbFalse)
        subtotal1 = subtotal1 + subtotal
    Next

    lblSubTotal.Caption = FormatNumber(subtotal1, 2, vbTrue, vbFalse)
    
    Total = (subtotal1 - discount - NVAL(txtRefunds.Text) - NVAL(txtFees.Text)) + NVAL(txtVAT.Text) + NVAL(txtFreight.Text)
    lblTotal.Caption = FormatNumber(Total, 2, vbTrue, vbFalse)
End Sub
Private Sub lblInvoice_Click()

End Sub

Private Sub btnComplete_Click()

End Sub

Private Sub btnReceiveOrder_Click()
    Set con = New ADODB.Connection
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    
    Dim Item As MSComctlLib.ListItem
    Dim itemx As MSComctlLib.ListItem
    
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "PO_PurchaseInvoice_AutoFill"
    cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , PO_PurchaseOrderFrm.PurchaseOrderId)
    Set rec = cmd.Execute
    'lvItems.ListItems.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            Dim isFound As Boolean
'            'Check if Current Record exists in the list
'            For Each itemx In lvItems.ListItems
'                If itemx.SubItems(8) = rec!ProductId And itemx.SubItems(5) = rec!Uom Then
'                    isFound = True
'                    'itemx.SubItems(4) = FormatNumber(rec!pickedquantity, 2, vbTrue)
'                    'Exit For
'                End If
'            Next
            
            If isFound = False Then
                Set Item = lvItems.ListItems.add(, , "")
                    Item.SubItems(1) = ""
                    Item.SubItems(2) = rec!itemcode
                    Item.SubItems(3) = rec!Name
                    Item.SubItems(4) = FormatNumber(rec!receivedquantity, 2, vbTrue)
                    Item.SubItems(5) = rec!Uom
                    Item.SubItems(6) = FormatNumber(rec!cost, 2, vbTrue)
                    Item.SubItems(7) = FormatNumber(rec!subtotal, 2, vbTrue)
                    'item.SubItems(8) = ""
                    Item.SubItems(8) = rec!ProductId
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
    PO_PurchaseReturnInvoiceFrm.cmbVendor.Text = PO_PurchaseOrderFrm.cmbVendor.Text
    PO_PurchaseReturnInvoiceFrm.Show (1)
End Sub

Public Sub SelectDelivery()
    Set con = New ADODB.Connection
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    Dim Item As MSComctlLib.ListItem

    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "PO_Invoice_Generate"
    cmd.Parameters.Append cmd.CreateParameter("@ReceiveOrderId", adInteger, adParamInput, , ReceiveOrderId)
    cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , PurchaseOrderId)
    Set rec = cmd.Execute

    If Not rec.EOF Then
        lvItems.ListItems.Clear
        Do Until rec.EOF
            Set Item = lvItems.ListItems.add(, , "")
                Item.SubItems(1) = ""
                Item.SubItems(2) = rec!itemcode
                Item.SubItems(3) = rec!Name
                Item.SubItems(4) = FormatNumber(rec!quantity, 2, vbTrue, vbFalse)
                Item.SubItems(5) = rec!Uom
                Item.SubItems(6) = FormatNumber(rec!cost, 2, vbTrue, vbFalse)
                Item.SubItems(8) = rec!ProductId
                Item.SubItems(9) = rec!ReceiveOrderLineId
            rec.MoveNext
        Loop
    End If

    con.Close
    cmbReferenceNumber.Text = lvDelivery.SelectedItem.SubItems(1)
    txtRemarks.SetFocus
    CountTotal
End Sub
Public Sub GetDeliveryCost()
    Dim con As New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    
    Dim Item As MSComctlLib.ListItem
    For Each Item In lvItems.ListItems
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "PO_InvoiceCost_Get"
        cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , Item.SubItems(8))
        cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , PurchaseOrderId)
        Set rec = cmd.Execute
        If Not rec.EOF Then
            Item.SubItems(6) = FormatNumber(rec!cost, 2, vbTrue, vbFalse)
        End If
    Next
    con.Close
End Sub


Public Sub btnSearch_Click()
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "PO_PurchaseInvoice_Search"
    cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInput, 50, txtSearch_OrderNumber.Text)
    cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , PurchaseOrderId)
    cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , cmbSearch_Status.ItemData(cmbSearch_Status.ListIndex))
    cmd.Parameters.Append cmd.CreateParameter("@DateFrom", adDate, adParamInput, , DateFrom.value)
    cmd.Parameters.Append cmd.CreateParameter("@DateTo", adDate, adParamInput, , DateTo.value)
    
    Dim Item As MSComctlLib.ListItem
    Set rec = cmd.Execute
    lvSearch.ListItems.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            Set Item = lvSearch.ListItems.add(, , rec!PurchaseInvoiceId)
                Item.SubItems(1) = rec!OrderNumber
                Item.SubItems(2) = rec!ReferenceNumber
                Item.SubItems(3) = rec!Status
            rec.MoveNext
        Loop
    End If
    con.Close
End Sub

Private Sub cmbReferenceNumber_GotFocus()
    lvDelivery.Visible = True
    Populate "DeliveryReceipt"
    lvDelivery.Left = 5715
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
            lvDelivery.Visible = False
    End Select
End Sub

Private Sub Form_Load()
    lvSearch.ColumnHeaders(2).width = lvSearch.width * 0.32
    lvSearch.ColumnHeaders(3).width = lvSearch.width * 0.32
    lvSearch.ColumnHeaders(4).width = lvSearch.width * 0.32

    lvItems.ColumnHeaders(3).width = lvItems.width * 0.14
    lvItems.ColumnHeaders(4).width = lvItems.width * 0.29
    lvItems.ColumnHeaders(5).width = lvItems.width * 0.12
    lvItems.ColumnHeaders(5).width = lvItems.width * 0.12
    lvItems.ColumnHeaders(6).width = lvItems.width * 0.09
    lvItems.ColumnHeaders(7).width = lvItems.width * 0.13
    lvItems.ColumnHeaders(8).width = lvItems.width * 0.2
    
    lvDelivery.ColumnHeaders(2).width = lvDelivery.width * 0.47
    lvDelivery.ColumnHeaders(3).width = lvDelivery.width * 0.47
    
    Initialize
    
    dtOrder.value = PO_PurchaseOrderFrm.dtOrder.value
    dtDue.value = Format(Now, "mm/dd/yy")
    
    txtRemarks.Text = PO_PurchaseOrderFrm.global_remarks
    
    Populate "Terms"
    Populate "InvoiceGet"
    Populate "Status"
'    Populate "InvoiceLineGet"
    
    DateFrom.value = Format(Now - 30, "MM/DD/YY")
    DateTo.value = Format(Now, "MM/DD/YY")
    CountTotal
    
    PO_PurchaseReturnInvoiceFrm.Show
    PO_PurchaseReturnInvoiceFrm.Hide
    
    On Error Resume Next
    txtDiscountPercent.Text = Val(Replace(txtDiscount.Text, ",", "")) / Val(Replace(lblSubTotal.Caption, ",", "")) * 100
End Sub

Private Sub lblDiscountPercent_Click()
    Dim x As Variant
    x = InputBox("Please input discount in percentage.")
    If IsNumeric(x) = False And Trim(x) <> "" Then
        MsgBox "Invalid value.", vbCritical
    Else
        Dim discount As Double
        discount = NVAL(lblSubTotal.Caption) * (NVAL(x) / 100)
        txtDiscount.Text = FormatNumber(discount, 2, vbTrue, vbFalse)
    End If
End Sub

Private Sub lblVatpercent_Click()
    Dim x As Variant
    x = InputBox("Please input tax in percentage.")
    If IsNumeric(x) = False And Trim(x) <> "" Then
        MsgBox "Invalid value.", vbCritical
    Else
        Dim tax As Double
        tax = (NVAL(lblSubTotal.Caption) - NVAL(txtDiscount.Text)) * (NVAL(x) / 100)
        txtVAT.Text = FormatNumber(tax, 2, vbTrue, vbFalse)
    End If
End Sub

Private Sub lvDelivery_ItemClick(ByVal Item As MSComctlLib.ListItem)
    ReceiveOrderId = lvDelivery.SelectedItem.Text
    SelectDelivery
    GetDeliveryCost
    
    lvDelivery.Visible = False
    lvDelivery.Left = -9999
    CountTotal
End Sub

Private Sub lvItems_DblClick()
    If lvItems.ListItems.count > 0 Then
        With PO_PurchaseInvoiceDialogFrm
            .txtQuantity.Text = lvItems.SelectedItem.SubItems(4)
            .txtPrice.Text = lvItems.SelectedItem.SubItems(6)
            .isModify = True
            .Show (1)
        End With
    End If
End Sub

Private Sub lvItems_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
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

Private Sub lvSearch_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lvSearch.ListItems.count > 0 Then
        Initialize
        InvoiceId = lvSearch.SelectedItem.Text
        Populate "InvoiceGet"
        Populate "InvoiceLineGet"
        CountTotal
        LoadImageStatus picStatus, GetStatus(StatusId)
    End If
End Sub

Private Sub tb_Standard_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 ' NEW
            Initialize
        Case 2 'SAVE
            If Validated = False Then Exit Sub
            If EditAccessRights(32) = False Then
                MsgBox ErrorCodes(74), vbCritical, "Limited Rights"
                Exit Sub
            End If
            If StatusId = 7 Then
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
            cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , PurchaseOrderId)
            cmd.Parameters.Append cmd.CreateParameter("@ReceiveOrderId", adInteger, adParamInput, , ReceiveOrderId)
            cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInput, 250, txtInvoiceNumber.Text)
            cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtOrder.value)
            cmd.Parameters.Append cmd.CreateParameter("@DueDate", adDate, adParamInput, , dtDue.value)
            cmd.Parameters.Append cmd.CreateParameter("@TermId", adInteger, adParamInput, , cmbTerms.ItemData(cmbTerms.ListIndex))
            cmd.Parameters.Append cmd.CreateParameter("@SubTotal", adDecimal, adParamInput, , Val(Replace(lblSubTotal.Caption, ",", "")))
                                  cmd.Parameters("@SubTotal").Precision = 18
                                  cmd.Parameters("@SubTotal").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@Total", adDecimal, adParamInput, , Val(Replace(lblTotal.Caption, ",", "")))
                                  cmd.Parameters("@Total").Precision = 18
                                  cmd.Parameters("@Total").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@Discount", adDecimal, adParamInput, , Val(Replace(txtDiscount.Text, ",", "")))
                                  cmd.Parameters("@Discount").Precision = 18
                                  cmd.Parameters("@Discount").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@Refunds", adDecimal, adParamInput, , Val(Replace(txtRefunds.Text, ",", "")))
                                  cmd.Parameters("@Refunds").Precision = 18
                                  cmd.Parameters("@Refunds").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, txtRemarks.Text)
            cmd.Parameters.Append cmd.CreateParameter("@ReferenceNumber", adVarChar, adParamInput, 50, cmbReferenceNumber.Text)
            cmd.Parameters.Append cmd.CreateParameter("@Freight", adDecimal, adParamInput, , NVAL(txtFreight.Text))
                                  cmd.Parameters("@Freight").Precision = 18
                                  cmd.Parameters("@Freight").NumericScale = 2
             cmd.Parameters.Append cmd.CreateParameter("@Tax", adDecimal, adParamInput, , NVAL(txtVAT.Text))
                                  cmd.Parameters("@Tax").Precision = 18
                                  cmd.Parameters("@Tax").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , StatusId)
            cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
            cmd.Parameters.Append cmd.CreateParameter("@WorkStationId", adInteger, adParamInput, , WorkstationId)
            
            If Val(InvoiceId) = 0 Then
                cmd.CommandText = "PO_PurchaseInvoice_Insert"
                cmd.Execute
                InvoiceId = cmd.Parameters("@InvoiceId")
                
                SavePOSAuditTrail UserId, WorkstationId, "", "Invoiced delivery ref #: " & cmbReferenceNumber.Text, "PURCHASING"
            Else
                cmd.CommandText = "PO_PurchaseInvoice_Update"
                cmd.Execute
                InvoiceId = cmd.Parameters("@InvoiceId")
                
                SavePOSAuditTrail UserId, WorkstationId, "", "Updated invoice order ref: " & PO_PurchaseOrderFrm.txtOrderNumber.Text, "PURCHASING"
            End If
            
            'SAVE LINE
            Dim Item As MSComctlLib.ListItem
            For Each Item In lvItems.ListItems
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc

                cmd.Parameters.Append cmd.CreateParameter("@InvoiceLineId", adInteger, adParamInputOutput, , Val(Item.Text))
                cmd.Parameters.Append cmd.CreateParameter("@InvoiceId", adInteger, adParamInput, , InvoiceId)
                cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , PurchaseOrderId)
                cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtOrder.value)
                cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , Val(Item.SubItems(8)))
                cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, Item.SubItems(3))
                cmd.Parameters.Append cmd.CreateParameter("@Quantity", adDecimal, adParamInput, , Val(Replace(Item.SubItems(4), ",", "")))
                                      cmd.Parameters("@Quantity").Precision = 18
                                      cmd.Parameters("@Quantity").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@Uom", adVarChar, adParamInput, 250, Item.SubItems(5))
                cmd.Parameters.Append cmd.CreateParameter("@Cost", adDecimal, adParamInput, , Val(Replace(Item.SubItems(6), ",", "")))
                                      cmd.Parameters("@Cost").Precision = 18
                                      cmd.Parameters("@Cost").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@Subtotal", adDecimal, adParamInput, , Val(Replace(Item.SubItems(7), ",", "")))
                                      cmd.Parameters("@Subtotal").Precision = 18
                                      cmd.Parameters("@Subtotal").NumericScale = 2
                If Item.Text = "" Then
                    cmd.CommandText = "PO_PurchaseInvoiceLine_Insert"
                Else
                    cmd.CommandText = "PO_PurchaseInvoiceLine_Update"
                End If
                cmd.Execute
                Item.Text = cmd.Parameters("@InvoiceLineId")
            Next

            With PO_PurchaseOrderFrm
                .StatusId = 4
                .lvSearch.SelectedItem.SubItems(2) = "Invoiced"
                .txtStatus.Text = "Invoiced"
                LoadImageStatus .picStatus, GetStatus(.StatusId)
            End With

            'DELETE ORDERLINE IF ANY
            Dim ctr As Integer
            For ctr = 0 To ctrOrderLine
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc

                If OrderLine(ctr) <> 0 Then
                    cmd.Parameters.Append cmd.CreateParameter("@InvoiceLineId", adInteger, adParamInput, , OrderLine(ctr))
                    cmd.CommandText = "PO_PurchaseInvoiceLine_Delete"
                    cmd.Execute
                Else
                    Exit For
                End If
            Next

            'UPDATE PURCHASERETURNSTATUS
            For Each Item In PO_PurchaseReturnInvoiceFrm.lvModules.ListItems
                If Item.Checked = True Then
                    Set cmd = New ADODB.Command
                    cmd.ActiveConnection = con
                    cmd.CommandType = adCmdStoredProc
                    cmd.Parameters.Append cmd.CreateParameter("@PurchaseReturnId", adInteger, adParamInput, , NVAL(Item.SubItems(5)))
                    cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , NVAL(Item.SubItems(6)))
                    cmd.CommandText = "PO_PurchaseReturnStatus_Update"
                    cmd.Execute
                End If
            Next
            
            con.CommitTrans
            con.Close
            
            MsgBox "Invoice saved.", vbInformation
        Case 4 'CANCEL
            
        Case 6 'PRINT
            If InvoiceId <> 0 Then
                Screen.MousePointer = vbHourglass
                BASE_PrintPreviewFrm.Show '(1)
                Dim crxApp As New CRAXDRT.Application
                Dim crxRpt As New CRAXDRT.Report
                Set crxRpt = crxApp.OpenReport(App.path & "\Reports\PO_PurchaseInvoice.rpt")
                crxRpt.RecordSelectionFormula = "{PO_PurchaseInvoice.PurchaseInvoiceId}= " & Val(InvoiceId) & ""
                crxRpt.DiscardSavedData

                Call ResetRptDB(crxRpt)

                BASE_PrintPreviewFrm.CRViewer.ReportSource = crxRpt
                BASE_PrintPreviewFrm.CRViewer.ViewReport
                BASE_PrintPreviewFrm.CRViewer.Zoom 1
                Screen.MousePointer = vbDefault
                
                SavePOSAuditTrail UserId, WorkstationId, "", "Generated print preview on purchase invoice: " & PO_PurchaseOrderFrm.txtOrderNumber.Text, "PURCHASING"
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

Private Sub txtFees_Change()
    If IsNumeric(txtFees.Text) = False Then
        txtFees.Text = "0.00"
    End If
    CountTotal
End Sub

Private Sub txtFreight_Change()
    If IsNumeric(txtFreight.Text) = False Then
        txtFreight.Text = "0.00"
    End If
    CountTotal
End Sub

Private Sub txtVAT_Change()
    If IsNumeric(txtVAT.Text) = False And txtVAT.Text <> "" Then
        txtVAT.Text = "0.00"
    End If
    CountTotal
End Sub
