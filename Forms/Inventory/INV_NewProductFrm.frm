VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form INV_NewProductFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Product"
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
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_NewProductFrm.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_NewProductFrm.frx":6862
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_NewProductFrm.frx":D0C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_NewProductFrm.frx":13926
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_NewProductFrm.frx":1A188
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   4640
      TabIndex        =   33
      Top             =   0
      Width           =   10455
      Begin MSComDlg.CommonDialog DialogBox 
         Left            =   9240
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Select Image"
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1935
         Left            =   6120
         TabIndex        =   73
         Top             =   720
         Width           =   4215
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            Height          =   1575
            Left            =   120
            TabIndex        =   76
            Top             =   360
            Width           =   4095
            Begin VB.Image picMain 
               Height          =   1455
               Left            =   1080
               Stretch         =   -1  'True
               Top             =   120
               Width           =   2055
            End
         End
         Begin VB.Label btnFullScreen 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "View Full Screen"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   195
            Left            =   1275
            MouseIcon       =   "INV_NewProductFrm.frx":209EA
            MousePointer    =   99  'Custom
            TabIndex        =   78
            Top             =   120
            Width           =   1140
         End
         Begin VB.Label lblClear 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Clear"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   195
            Left            =   2640
            MouseIcon       =   "INV_NewProductFrm.frx":20B3C
            MousePointer    =   99  'Custom
            TabIndex        =   77
            Top             =   120
            Width           =   375
         End
         Begin VB.Label lblSelectImage 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Select Image"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   195
            Left            =   3240
            MouseIcon       =   "INV_NewProductFrm.frx":20C8E
            MousePointer    =   99  'Custom
            TabIndex        =   75
            Top             =   120
            Width           =   900
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Picture"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   0
            TabIndex        =   74
            Top             =   0
            Width           =   810
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1575
         Left            =   6120
         TabIndex        =   70
         Top             =   6960
         Width           =   4215
         Begin VB.CommandButton btnExtraInfo 
            Caption         =   "Add Extra Product Info"
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
            Left            =   1680
            TabIndex        =   17
            Top             =   1080
            Width           =   2535
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Extra Info"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   0
            TabIndex        =   72
            Top             =   0
            Width           =   1125
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "You can add extra details on your products such as expiry dates and Lot Numbers."
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   71
            Top             =   480
            Width           =   3900
         End
      End
      Begin VB.TextBox txtAlias2 
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
         Left            =   2280
         MaxLength       =   500
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   2640
         Width           =   3495
      End
      Begin VB.TextBox txtAlias1 
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
         Left            =   2280
         MaxLength       =   500
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   2280
         Width           =   3495
      End
      Begin VB.Frame FRAME_ProductDetails3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   975
         Left            =   6120
         TabIndex        =   61
         Top             =   5880
         Width           =   4215
         Begin VB.ComboBox cmbTaxInfo_Tax 
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   480
            Width           =   2535
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tax"
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
            Top             =   480
            Width           =   315
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tax Info"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   0
            TabIndex        =   62
            Top             =   0
            Width           =   930
         End
      End
      Begin VB.TextBox txtSalesInfoSRPMarkUp 
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
         Left            =   5280
         TabIndex        =   9
         Top             =   4560
         Width           =   495
      End
      Begin VB.TextBox txtSalesInfoDPMarkUp 
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
         Left            =   -9999
         TabIndex        =   19
         Top             =   4320
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtSalesInfoMSMarkUp 
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
         Left            =   -9999
         TabIndex        =   21
         Top             =   4680
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtSalesInfoBCMarkUp 
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
         Left            =   -9999
         TabIndex        =   23
         Top             =   5040
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtBasicInfo_Barcode 
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
         Left            =   2280
         MaxLength       =   500
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox txtSalesInfo_BCP 
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
         Left            =   -9999
         TabIndex        =   22
         Top             =   5040
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtSalesInfo_SP 
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
         Left            =   -9999
         TabIndex        =   20
         Top             =   4680
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtSalesInfo_DP 
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
         Left            =   -9999
         TabIndex        =   18
         Top             =   4320
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Frame Frame_ProductDetails2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   3135
         Left            =   6120
         TabIndex        =   47
         Top             =   2760
         Width           =   4215
         Begin VB.TextBox txtReorderQuantity 
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
            Left            =   1680
            TabIndex        =   15
            Top             =   2640
            Width           =   2535
         End
         Begin VB.TextBox txtReorderPoint 
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
            Left            =   1680
            TabIndex        =   14
            Top             =   2280
            Width           =   2535
         End
         Begin VB.ComboBox cmbVendor 
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   840
            Width           =   2535
         End
         Begin VB.ComboBox cmbStorageInfo_Uom 
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
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1920
            Width           =   2535
         End
         Begin VB.TextBox txtCostingInfo_AverageCost 
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
            Left            =   1680
            TabIndex        =   11
            Top             =   480
            Width           =   2535
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reorder Qty"
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
            Top             =   2640
            Width           =   1125
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reorder Point"
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
            Top             =   2280
            Width           =   1290
         End
         Begin VB.Label lblShowConversion 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Show Product Conversion"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   195
            Left            =   2400
            MouseIcon       =   "INV_NewProductFrm.frx":20DE0
            MousePointer    =   99  'Custom
            TabIndex        =   65
            Top             =   1680
            Width           =   1785
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier"
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
            Top             =   840
            Width           =   780
         End
         Begin VB.Label lblStorageInfoTitle 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Storage Info"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   0
            TabIndex        =   54
            Top             =   1440
            Width           =   1425
         End
         Begin VB.Label lblUom 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Unit of Measure"
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
            TabIndex        =   53
            Top             =   1920
            Width           =   1485
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Costing Info"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   0
            TabIndex        =   49
            Top             =   0
            Width           =   1395
         End
         Begin VB.Label lblCostingInfo_Cost 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Current Cost"
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
            TabIndex        =   48
            Top             =   480
            Width           =   1155
         End
      End
      Begin VB.TextBox txtSalesInfo_Price 
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
         Left            =   3050
         TabIndex        =   8
         Top             =   4560
         Width           =   2175
      End
      Begin VB.Frame Frame_ProductDetails1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2775
         Left            =   240
         TabIndex        =   40
         Top             =   5880
         Width           =   5535
         Begin MSComctlLib.ListView lvInventory 
            Height          =   1815
            Left            =   240
            TabIndex        =   10
            Top             =   480
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   3201
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FlatScrollBar   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   6
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "InventoryId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "LocationId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "ProductId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Location"
               Object.Width           =   6115
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Sublocation"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Quantity"
               Object.Width           =   2806
            EndProperty
         End
         Begin VB.Label lblInventory_MoreLocations 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Specify Locations"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   195
            Left            =   4320
            MouseIcon       =   "INV_NewProductFrm.frx":20F32
            MousePointer    =   99  'Custom
            TabIndex        =   44
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblInventory_QtyOnHand 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   4440
            TabIndex        =   43
            Top             =   2400
            Width           =   1065
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Quantity on Hand:"
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
            TabIndex        =   42
            Top             =   2400
            Width           =   1680
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Inventory"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   0
            TabIndex        =   41
            Top             =   0
            Width           =   1140
         End
      End
      Begin VB.ComboBox cmbBasicInfo_Type 
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
         ItemData        =   "INV_NewProductFrm.frx":21084
         Left            =   2280
         List            =   "INV_NewProductFrm.frx":21086
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3360
         Width           =   3495
      End
      Begin VB.TextBox txtBasicInfo_ItemCode 
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
         Left            =   2280
         MaxLength       =   50
         TabIndex        =   1
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox txtBasicInfo_Name 
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
         Left            =   2280
         MaxLength       =   500
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1560
         Width           =   3495
      End
      Begin VB.ComboBox cmbBasicInfo_Category 
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
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   3000
         Width           =   3495
      End
      Begin MSComctlLib.Toolbar tb_Standard 
         Height          =   330
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   582
         ButtonWidth     =   1588
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
               Caption         =   "Delete"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Copy"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alias 2"
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
         Left            =   480
         TabIndex        =   69
         Top             =   2640
         Width           =   600
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alias 1"
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
         Left            =   480
         TabIndex        =   68
         Top             =   2280
         Width           =   600
      End
      Begin VB.Label lblShowMorePrice 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show more Prices"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   195
         Left            =   4500
         MouseIcon       =   "INV_NewProductFrm.frx":21088
         MousePointer    =   99  'Custom
         TabIndex        =   64
         Top             =   4920
         Width           =   1245
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mark-Up (%)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4740
         TabIndex        =   60
         Top             =   4200
         Width           =   1020
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Barcode"
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
         Left            =   480
         TabIndex        =   58
         Top             =   1920
         Width           =   750
      End
      Begin VB.Label lblDiscount3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mega Discount"
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
         TabIndex        =   57
         Top             =   5040
         Width           =   1365
      End
      Begin VB.Label lblDiscount2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Special Discount"
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
         TabIndex        =   56
         Top             =   4680
         Width           =   1515
      End
      Begin VB.Label lblDiscount1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Standard Discount"
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
         TabIndex        =   55
         Top             =   4320
         Width           =   1680
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Info"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   240
         TabIndex        =   46
         Top             =   3960
         Width           =   1125
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Standard Retail Price"
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
         Left            =   480
         TabIndex        =   45
         Top             =   4560
         Width           =   1920
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
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
         Left            =   480
         TabIndex        =   39
         Top             =   3360
         Width           =   450
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Basic Info"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   240
         TabIndex        =   38
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Code"
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
         Left            =   480
         TabIndex        =   37
         Top             =   1200
         Width           =   960
      End
      Begin VB.Label Label8 
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
         Left            =   480
         TabIndex        =   36
         Top             =   1560
         Width           =   555
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
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
         Left            =   480
         TabIndex        =   35
         Top             =   3000
         Width           =   825
      End
   End
   Begin VB.Frame LayoutFrame_Search 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   9390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.TextBox txtSearch_Supplier 
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
         TabIndex        =   26
         Top             =   1200
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
         TabIndex        =   51
         Top             =   1935
         Width           =   3015
      End
      Begin VB.TextBox txtSearch_ItemCode 
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
      Begin VB.ComboBox cmbSearch_Category 
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
         TabIndex        =   27
         Top             =   1560
         Width           =   3015
      End
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
         TabIndex        =   25
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
         TabIndex        =   28
         Top             =   2400
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvSearch 
         Height          =   6015
         Left            =   120
         TabIndex        =   29
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
            Object.Width           =   6526
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Category"
            Object.Width           =   2893
         EndProperty
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
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
         TabIndex        =   79
         Top             =   1200
         Width           =   780
      End
      Begin VB.Label Label16 
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
         TabIndex        =   52
         Top             =   1935
         Width           =   570
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item code"
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
         TabIndex        =   50
         Top             =   480
         Width           =   930
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
         TabIndex        =   32
         Top             =   80
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
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
         TabIndex        =   31
         Top             =   1560
         Width           =   825
      End
      Begin VB.Label Label2 
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
         TabIndex        =   30
         Top             =   840
         Width           =   555
      End
   End
End
Attribute VB_Name = "INV_NewProductFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ProductId As Long
Public isService, isInsert, isActive As Boolean
Dim deleteCtr(10000) As Long
Dim ctr As Long
Dim PicPath As String

Private Sub Initialize()
    Dim txtControl As Control
    isService = False
    ProductId = 0
    ctr = 0
    isActive = True
    isActivated (True)
    
    For Each txtControl In Me.Controls
        If TypeOf txtControl Is TextBox And txtControl.Name <> "txtSearch_Name" Then
            txtControl.Text = ""
        End If
    Next
    
    txtBasicInfo_ItemCode.BackColor = &HC0C0FF
    txtBasicInfo_Name.BackColor = &HC0C0FF
    
    cmbBasicInfo_Category.ListIndex = 0
    cmbSearch_Category.ListIndex = 0
    cmbBasicInfo_Type.ListIndex = 0
    cmbStorageInfo_Uom.ListIndex = 0
    cmbVendor.ListIndex = 0
    
    tb_Standard.Buttons(4).Image = 3
    tb_Standard.Buttons(4).Caption = "Delete"
    
    lvInventory.ListItems.Clear
    Populate ("Location")
    
    On Error Resume Next
    txtBasicInfo_ItemCode.SetFocus
    CountQuantity
    cmbStorageInfo_Uom.ListIndex = 6 'pcs
    picMain.Picture = Nothing
    DialogBox.filename = ""
    PicPath = ""
End Sub
Public Sub CountQuantity()
    Dim item As MSComctlLib.ListItem
    Dim Total As Double
    For Each item In lvInventory.ListItems
        If item.SubItems(1) < 4 Then 'SHOULD HAVE LIKE <...
            Total = Total + Val(Replace(item.SubItems(5), ",", ""))
        End If
    Next
    lblInventory_QtyOnHand.Caption = FormatNumber(Total, 2, vbTrue, vbFalse)
End Sub
Private Sub isActivated(value As Boolean)
    'BASIC INFO
    txtBasicInfo_ItemCode.Enabled = value
    txtBasicInfo_Name.Enabled = value
    txtBasicInfo_Barcode.Enabled = value
    cmbBasicInfo_Category.Enabled = value
    cmbBasicInfo_Type.Enabled = value
    
    'SALES INFO
    txtSalesInfo_Price.Enabled = value
    txtSalesInfoSRPMarkUp.Enabled = value
    txtSalesInfo_DP.Enabled = value
    txtSalesInfoDPMarkUp.Enabled = value
    txtSalesInfo_SP.Enabled = value
    txtSalesInfoMSMarkUp.Enabled = value
    txtSalesInfo_BCP.Enabled = value
    txtSalesInfoBCMarkUp.Enabled = value
    
    'INVENTORY INFO
    lvInventory.Enabled = value
    
    'COSTING INFO
    txtCostingInfo_AverageCost.Enabled = value
    cmbVendor.Enabled = value
    
    'STORAGE INFO
    cmbStorageInfo_Uom.Enabled = value
End Sub

Public Sub Populate(ByVal data As String)
    Set rec = New ADODB.Recordset
    Select Case data
        Case "Category"
            Set rec = Global_Data("Category")
            cmbBasicInfo_Category.Clear
            cmbSearch_Category.Clear
            cmbSearch_Category.AddItem ""
            cmbSearch_Category.ItemData(cmbSearch_Category.NewIndex) = 0
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        cmbSearch_Category.AddItem rec!Category
                        cmbSearch_Category.ItemData(cmbSearch_Category.NewIndex) = rec!CategoryId
                        
                        cmbBasicInfo_Category.AddItem rec!Category
                        cmbBasicInfo_Category.ItemData(cmbBasicInfo_Category.NewIndex) = rec!CategoryId
                    End If
                    rec.MoveNext
                Loop
            End If
        Case "Vendor"
            Set rec = Global_Data("Vendor")
            cmbVendor.Clear
            cmbVendor.AddItem ""
            cmbVendor.ItemData(cmbVendor.NewIndex) = 0
            If Not rec.EOF Then
                Do Until rec.EOF
                    cmbVendor.AddItem rec!Name
                    cmbVendor.ItemData(cmbVendor.NewIndex) = rec!VendorId
                    rec.MoveNext
                Loop
            End If
        Case "Status"
            cmbSearch_Status.Clear
            cmbSearch_Status.AddItem ""
            cmbSearch_Status.ItemData(cmbSearch_Status.NewIndex) = -1
            cmbSearch_Status.AddItem "Active"
            cmbSearch_Status.ItemData(cmbSearch_Status.NewIndex) = 1
            cmbSearch_Status.AddItem "Deactivated"
            cmbSearch_Status.ItemData(cmbSearch_Status.NewIndex) = 0
            cmbSearch_Status.ListIndex = 1
        Case "Type"
            Set rec = Global_Data("Type")
            cmbBasicInfo_Type.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    cmbBasicInfo_Type.AddItem rec!Type
                    cmbBasicInfo_Type.ItemData(cmbBasicInfo_Type.NewIndex) = rec!TypeId
                    rec.MoveNext
                Loop
            End If
        Case "Uom"
            Set rec = Global_Data("Uom")
            cmbStorageInfo_Uom.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        cmbStorageInfo_Uom.AddItem rec!Uom
                        cmbStorageInfo_Uom.ItemData(cmbStorageInfo_Uom.NewIndex) = rec!UomId
                    End If
                    rec.MoveNext
                Loop
            End If
            On Error Resume Next
            cmbStorageInfo_Uom.ListIndex = 6 'pcs
        Case "Tax"
            Set rec = Global_Data("Tax")
            cmbTaxInfo_Tax.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        cmbTaxInfo_Tax.AddItem rec!TaxName
                        cmbTaxInfo_Tax.ItemData(cmbTaxInfo_Tax.NewIndex) = rec!TaxId
                    End If
                    rec.MoveNext
                Loop
            End If
            On Error Resume Next
            cmbTaxInfo_Tax.ListIndex = 1
        Case "Location"
            Set rec = Global_Data("Location")
            Dim item As MSComctlLib.ListItem
            lvInventory.ListItems.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!LocationId = 1 Then
                        Set item = lvInventory.ListItems.add(, , "")
                            item.SubItems(1) = rec!LocationId 'LocationId
                            item.SubItems(3) = rec!Location 'Location
                            item.SubItems(5) = "0.00" 'Quantity
                        Exit Do
                    End If
                    rec.MoveNext
                Loop
            End If
        Case "Product"
            Set rec = Global_Data("Product")
            lvSearch.ListItems.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        Set item = lvSearch.ListItems.add(, , rec!ProductId)
                            item.SubItems(1) = rec!itemcode
                            item.SubItems(2) = rec!Name
                            item.SubItems(3) = rec!Category
                    End If
                    rec.MoveNext
                Loop
            End If
        Case "ProductSelect"
            Set con = New ADODB.Connection
            Set rec = New ADODB.Recordset
            Set cmd = New ADODB.Command
            
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "BASE_Product_Get"
            cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , ProductId)
            Set rec = cmd.Execute
            If Not rec.EOF Then
                txtBasicInfo_ItemCode.Text = rec!itemcode
                txtBasicInfo_Name.Text = rec!Name
                txtBasicInfo_Barcode.Text = rec!Barcode
                txtAlias1.Text = rec!alias1
                txtAlias2.Text = rec!alias2
                cmbBasicInfo_Category.Text = rec!Category
                cmbBasicInfo_Type.Text = rec!Type
                cmbTaxInfo_Tax.Text = rec!TaxName
                On Error Resume Next
                cmbStorageInfo_Uom.Text = rec!Uom
                txtReorderPoint.Text = FormatNumber(rec!reorderpoint, 2, vbTrue, vbFalse)
                txtReorderQuantity.Text = FormatNumber(rec!ReorderQuantity, 2, vbTrue, vbFalse)
                picMain.Picture = LoadPicture(rec!Image)
                PicPath = rec!Image
                If IsNull(rec!VendorId) Then
                    cmbVendor.ListIndex = 0
                Else
                    On Error Resume Next
                    cmbVendor.Text = rec!Vendor
                End If
                If IsNull(rec!UnitPriceMarkUp) = True Then txtSalesInfoSRPMarkUp.Text = "" Else txtSalesInfoSRPMarkUp.Text = rec!UnitPriceMarkUp
                If IsNull(rec!Price1MarkUp) = True Then txtSalesInfoDPMarkUp.Text = "" Else txtSalesInfoDPMarkUp.Text = rec!Price1MarkUp
                If IsNull(rec!Price2MarkUp) = True Then txtSalesInfoMSMarkUp.Text = "" Else txtSalesInfoMSMarkUp.Text = rec!Price2MarkUp
                If IsNull(rec!Price3MarkUp) = True Then txtSalesInfoBCMarkUp.Text = "" Else txtSalesInfoBCMarkUp.Text = rec!Price3MarkUp
                txtSalesInfo_Price.Text = FormatNumber(rec!unitprice, 2, vbTrue, vbFalse)
                txtSalesInfo_DP.Text = FormatNumber(rec!price1, 2, vbTrue, vbFalse)
                txtSalesInfo_SP.Text = FormatNumber(rec!price2, 2, vbTrue, vbFalse)
                txtSalesInfo_BCP.Text = FormatNumber(rec!price3, 2, vbTrue, vbFalse)
                txtCostingInfo_AverageCost.Text = FormatNumber(rec!unitcost, 2, vbTrue, vbFalse)
            End If
            isActive = rec!isActive
            If rec!isActive = "False" Then
                tb_Standard.Buttons(4).Caption = "Activate"
                tb_Standard.Buttons(4).Image = 4
                isActivated (False)
            Else
                tb_Standard.Buttons(4).Caption = "Delete"
                tb_Standard.Buttons(4).Image = 3
                isActivated (True)
            End If
            con.Close
        Case "InventoryLoad"
            Set con = New ADODB.Connection
            Set rec = New ADODB.Recordset
            Set cmd = New ADODB.Command
            
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "BASE_Inventory_Get"
            cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , ProductId)
            Set rec = cmd.Execute
            lvInventory.ListItems.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    Set item = lvInventory.ListItems.add(, , rec!inventoryId)
                        item.SubItems(1) = rec!LocationId
                        item.SubItems(2) = rec!ProductId
                        item.SubItems(3) = rec!Location
                        item.SubItems(5) = FormatNumber(rec!quantity, 2, vbTrue, vbFalse)
                    rec.MoveNext
                Loop
            End If
            con.Close
    End Select
End Sub
Private Function ValidateData() As Boolean
    ValidateData = False
    
    'CHECK EMPTY FIELDS
    If Trim(txtSalesInfo_Price.Text) = "" Then txtSalesInfo_Price.Text = "0.00"
    If Trim(txtSalesInfo_DP.Text) = "" Then txtSalesInfo_DP.Text = "0.00"
    If Trim(txtSalesInfo_SP.Text) = "" Then txtSalesInfo_SP.Text = "0.00"
    If Trim(txtSalesInfo_BCP.Text) = "" Then txtSalesInfo_BCP.Text = "0.00"
    
    If Trim(txtCostingInfo_AverageCost.Text) = "" Then txtCostingInfo_AverageCost.Text = "0.00"
    
    If Trim(txtBasicInfo_ItemCode.Text) = "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(1)
        GLOBAL_MessageFrm.Show (1)
        txtBasicInfo_ItemCode.SetFocus
    ElseIf Trim(txtBasicInfo_Name.Text) = "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(2)
        GLOBAL_MessageFrm.Show (1)
        txtBasicInfo_Name.SetFocus
    ElseIf cmbBasicInfo_Category.ListIndex = -1 Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(6)
        GLOBAL_MessageFrm.Show (1)
        cmbBasicInfo_Category.SetFocus
    ElseIf IsNumeric(txtSalesInfo_Price.Text) = False And Trim(txtSalesInfo_Price.Text) <> "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(7)
        GLOBAL_MessageFrm.Show (1)
        txtSalesInfo_Price.SetFocus
    ElseIf IsNumeric(txtSalesInfoSRPMarkUp.Text) = False And Trim(txtSalesInfoSRPMarkUp.Text) <> "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(41)
        GLOBAL_MessageFrm.Show (1)
        txtSalesInfoSRPMarkUp.SetFocus
    ElseIf IsNumeric(txtSalesInfo_DP.Text) = False And Trim(txtSalesInfo_DP.Text) <> "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(7)
        GLOBAL_MessageFrm.Show (1)
        txtSalesInfo_DP.SetFocus
    ElseIf IsNumeric(txtSalesInfoDPMarkUp.Text) = False And Trim(txtSalesInfoDPMarkUp.Text) <> "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(41)
        GLOBAL_MessageFrm.Show (1)
        txtSalesInfoDPMarkUp.SetFocus
    ElseIf IsNumeric(txtSalesInfo_SP.Text) = False And Trim(txtSalesInfo_SP.Text) <> "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(7)
        GLOBAL_MessageFrm.Show (1)
        txtSalesInfo_SP.SetFocus
    ElseIf IsNumeric(txtSalesInfoMSMarkUp.Text) = False And Trim(txtSalesInfoMSMarkUp.Text) <> "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(41)
        GLOBAL_MessageFrm.Show (1)
        txtSalesInfoMSMarkUp.SetFocus
    ElseIf IsNumeric(txtSalesInfo_BCP.Text) = False And Trim(txtSalesInfo_BCP.Text) <> "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(7)
        GLOBAL_MessageFrm.Show (1)
        txtSalesInfo_BCP.SetFocus
    ElseIf IsNumeric(txtSalesInfoBCMarkUp.Text) = False And Trim(txtSalesInfoBCMarkUp.Text) <> "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(41)
        GLOBAL_MessageFrm.Show (1)
        txtSalesInfoBCMarkUp.SetFocus
    ElseIf Trim(cmbStorageInfo_Uom.Text) = "" And isService = False Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(10)
        GLOBAL_MessageFrm.Show (1)
        cmbStorageInfo_Uom.SetFocus
    ElseIf IsNumeric(txtCostingInfo_AverageCost.Text) = False And Trim(txtCostingInfo_AverageCost.Text) <> "" And isService = False Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(9)
        GLOBAL_MessageFrm.Show (1)
        txtCostingInfo_AverageCost.SetFocus
    ElseIf IsNumeric(txtReorderPoint.Text) = False And Trim(txtReorderPoint.Text) <> "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(72)
        GLOBAL_MessageFrm.Show (1)
        txtReorderPoint.SetFocus
    ElseIf IsNumeric(txtReorderQuantity.Text) = False And Trim(txtReorderQuantity.Text) <> "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(73)
        GLOBAL_MessageFrm.Show (1)
        txtReorderQuantity.SetFocus
    ElseIf Trim(cmbTaxInfo_Tax.Text) = "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(50)
        GLOBAL_MessageFrm.Show (1)
        cmbTaxInfo_Tax.SetFocus
    Else
        ValidateData = True
    End If
End Function

Private Sub btnExtraInfo_Click()
    If ProductId = 0 Then
        MsgBox "Please select or create a product first.", vbCritical
    Else
        INV_ProductExtraInfoFrm.Show (1)
    End If
End Sub

Private Sub btnFullScreen_Click()
    INV_ImageFullScreen.picMain.Picture = LoadPicture(PicPath)
    INV_ImageFullScreen.Show
End Sub

Public Sub btnSearch_Click()
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_Product_Search1"
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 500, txtSearch_name.Text)
    If Trim(txtSearch_ItemCode.Text) <> "" Then
        cmd.Parameters.Append cmd.CreateParameter("@ItemCode", adVarChar, adParamInput, 50, txtSearch_ItemCode.Text)
    Else
        cmd.Parameters.Append cmd.CreateParameter("@ItemCode", adVarChar, adParamInput, 50, Null)
    End If
    cmd.Parameters.Append cmd.CreateParameter("@Supplier", adVarChar, adParamInput, 500, txtSearch_Supplier.Text)
    If cmbSearch_Category.ListIndex <> 0 Then
        cmd.Parameters.Append cmd.CreateParameter("@CategoryId", adInteger, adParamInput, , cmbSearch_Category.ItemData(cmbSearch_Category.ListIndex))
    Else
        cmd.Parameters.Append cmd.CreateParameter("@CategoryId", adInteger, adParamInput, , Null)
    End If
    cmd.Parameters.Append cmd.CreateParameter("@LocationId", adInteger, adParamInput, , Null)
    cmd.Parameters.Append cmd.CreateParameter("@TypeId", adInteger, adParamInput, , Null)
    If cmbSearch_Status.ListIndex <> 0 Then
        cmd.Parameters.Append cmd.CreateParameter("@isActive", adBoolean, adParamInput, , cmbSearch_Status.ItemData(cmbSearch_Status.ListIndex))
    End If
    
    
    Set rec = cmd.Execute
    lvSearch.ListItems.Clear
    Dim item As MSComctlLib.ListItem
    If Not rec.EOF Then
        Do Until rec.EOF
            'If rec!isActive = "True" Then
                Set item = lvSearch.ListItems.add(, , rec!ProductId)
                    item.SubItems(1) = rec!itemcode
                    item.SubItems(2) = rec!Name
                    item.SubItems(3) = rec!Category
            'End If
            rec.MoveNext
        Loop
    End If
    'DistinctList lvSearch
    con.Close
    BASE_ContainerFrm.statusBar_Main.Panels(1).Text = "Total Items: " & lvSearch.ListItems.Count
End Sub

Private Sub cmbBasicInfo_Type_Click()
    If cmbBasicInfo_Type.ListIndex <> 0 Then
        Frame_ProductDetails1.Visible = False
        'Frame_ProductDetails2.Visible = False
        isService = True
    Else
        Frame_ProductDetails1.Visible = True
        Frame_ProductDetails2.Visible = True
        isService = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyS
            If Shift = 2 Then
                tb_Standard_ButtonClick tb_Standard.Buttons(2)
            End If
        Case vbKeyN
            If Shift = 2 Then
                tb_Standard_ButtonClick tb_Standard.Buttons(1)
            End If
        Case vbKeyQ
            If Shift = 2 Then
                tb_Standard_ButtonClick tb_Standard.Buttons(4)
            End If
        Case vbKeyD
            If Shift = 2 Then
                tb_Standard_ButtonClick tb_Standard.Buttons(6)
            End If
    End Select
End Sub

Private Sub Form_Load()
    Populate ("Category")
    Populate ("Type")
    Populate ("Uom")
    Populate ("Location")
    Populate ("Product")
    Populate ("Status")
    Populate ("Vendor")
    Populate ("Tax")
    Initialize
    btnSearch_Click
    'StatusBarWidth Me, statusBar_Main
    
    'Access Rights
    txtCostingInfo_AverageCost.Enabled = EditAccessRights(2)
    txtCostingInfo_AverageCost.Visible = ViewAccessRights(2)
    lblCostingInfo_Cost.Visible = ViewAccessRights(2)
End Sub

Private Sub lblClear_Click()
    DialogBox.filename = ""
    picMain.Picture = Nothing
    PicPath = ""
End Sub

Private Sub lblInventory_MoreLocations_Click()
    CenterChildForm INV_LocationFrm
    INV_LocationFrm.Show
End Sub

Private Sub lblSelectImage_Click()
    DialogBox.Flags = cdlOFNHideReadOnly
    DialogBox.Filter = "JPG (*.jpg)"
    DialogBox.ShowOpen
    picMain.Picture = LoadPicture(DialogBox.filename)
    PicPath = DialogBox.filename
End Sub

Private Sub lblShowConversion_Click()
    If EditAccessRights(1) = False Then
        MsgBox ErrorCodes(74), vbCritical, "Limited Rights"
        Exit Sub
    End If
    If ProductId <> 0 Then
        INV_UomConversionFrm.Show (1)
    End If
End Sub

Private Sub lblShowMorePrice_Click()
    If EditAccessRights(1) = False Then
        MsgBox ErrorCodes(74), vbCritical, "Limited Rights"
        Exit Sub
    End If
    If ProductId <> 0 Then
        INV_UomPricingFrm.Show (1)
    End If
End Sub

Private Sub lvInventory_DblClick()
'    With lvInventory
'        If .ListItems.Count > 0 Then
'            Dim i As String
'            i = InputBox("Input quantity.", "Quantity", lvInventory.SelectedItem.SubItems(5))
'            If i = "" Then
'                Exit Sub
'            ElseIf IsNumeric(i) = False Then
'                Exit Sub
'            Else
'                .SelectedItem.SubItems(5) = FormatNumber(i, 2, vbFalse, vbFalse)
'                .SetFocus
'                CountQuantity
'            End If
'        End If
'    End With
End Sub

Private Sub lvInventory_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyDelete
        If lvInventory.ListItems.Count > 0 Then
            If lvInventory.SelectedItem.SubItems(1) <> "1" Then 'NOT Default Location
                If lvInventory.SelectedItem.Text <> "" Then 'Existing data
                        deleteCtr(ctr) = Val(lvInventory.SelectedItem.Text)
                        ctr = ctr + 1
                        lvInventory.ListItems.Remove (lvInventory.SelectedItem.Index)
                Else
                    lvInventory.ListItems.Remove (lvInventory.SelectedItem.Index)
                End If
            End If
        End If
    Case 13
        Call lvInventory_DblClick
    End Select
End Sub

Public Sub lvSearch_ItemClick(ByVal item As MSComctlLib.ListItem)
    With lvSearch
        If .ListItems.Count > 0 Then
            DialogBox.filename = ""
            picMain.Picture = Nothing
            ProductId = .SelectedItem.Text
            Populate "ProductSelect"
            Populate "InventoryLoad"
            CountQuantity
            isInsert = False
            'Access Rights
            txtCostingInfo_AverageCost.Enabled = EditAccessRights(2)
        End If
    End With
End Sub

Private Sub tb_Standard_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error GoTo ErrHandler
    If EditAccessRights(1) = False Then
        MsgBox ErrorCodes(74), vbCritical, "Limited Rights"
        Exit Sub
    End If
    Dim item As MSComctlLib.ListItem
    Select Case Button.Index
        Case 1 'New
            Initialize
            BASE_ContainerFrm.statusBar_Main.Panels(1).Text = MessageCodes(6) & " " & MessageCodes(1)
        Case 2 'Save
            If isActive = False Then Exit Sub
            If ValidateData = True Then
                On Error GoTo ErrHandler
                Set con = New ADODB.Connection
                Set cmd = New ADODB.Command
                
                'SAVE MAIN PRODUCT DETAILS
                con.ConnectionString = ConnString
                con.Open
                con.BeginTrans
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                
                'Price Mark Up variables
                Dim UnitPriceMarkUp, Price1MarkUp, Price2MarkUp, Price3MarkUp As Variant
                If Trim(txtSalesInfoSRPMarkUp.Text) = "" Then UnitPriceMarkUp = Null Else UnitPriceMarkUp = txtSalesInfoSRPMarkUp.Text
                If Trim(txtSalesInfoDPMarkUp.Text) = "" Then Price1MarkUp = Null Else Price1MarkUp = txtSalesInfoDPMarkUp.Text
                If Trim(txtSalesInfoMSMarkUp.Text) = "" Then Price2MarkUp = Null Else Price2MarkUp = txtSalesInfoMSMarkUp.Text
                If Trim(txtSalesInfoBCMarkUp.Text) = "" Then Price3MarkUp = Null Else Price3MarkUp = txtSalesInfoBCMarkUp.Text
                
                cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInputOutput, , ProductId)
                cmd.Parameters.Append cmd.CreateParameter("@ItemCode", adVarChar, adParamInput, 50, txtBasicInfo_ItemCode.Text)
                cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 500, txtBasicInfo_Name.Text)
                cmd.Parameters.Append cmd.CreateParameter("@Barcode", adVarChar, adParamInput, 50, txtBasicInfo_Barcode.Text)
                cmd.Parameters.Append cmd.CreateParameter("@Alias1", adVarChar, adParamInput, 50, txtAlias1.Text)
                cmd.Parameters.Append cmd.CreateParameter("@Alias2", adVarChar, adParamInput, 50, txtAlias2.Text)
                cmd.Parameters.Append cmd.CreateParameter("@CategoryId", adInteger, adParamInput, , cmbBasicInfo_Category.ItemData(cmbBasicInfo_Category.ListIndex))
                cmd.Parameters.Append cmd.CreateParameter("@TypeId", adInteger, adParamInput, , cmbBasicInfo_Type.ItemData(cmbBasicInfo_Type.ListIndex))
                cmd.Parameters.Append cmd.CreateParameter("@TaxId", adInteger, adParamInput, , cmbTaxInfo_Tax.ItemData(cmbTaxInfo_Tax.ListIndex))
                cmd.Parameters.Append cmd.CreateParameter("@UnitPrice", adDecimal, adParamInput, , Val(Replace(txtSalesInfo_Price.Text, ",", "")))
                                      cmd.Parameters("@UnitPrice").Precision = 18
                                      cmd.Parameters("@UnitPrice").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@Price1", adDecimal, adParamInput, , Val(Replace(txtSalesInfo_DP.Text, ",", "")))
                                      cmd.Parameters("@Price1").Precision = 18
                                      cmd.Parameters("@Price1").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@Price2", adDecimal, adParamInput, , Val(Replace(txtSalesInfo_SP.Text, ",", "")))
                                      cmd.Parameters("@Price2").Precision = 18
                                      cmd.Parameters("@Price2").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@Price3", adDecimal, adParamInput, , Val(Replace(txtSalesInfo_BCP.Text, ",", "")))
                                      cmd.Parameters("@Price3").Precision = 18
                                      cmd.Parameters("@Price3").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@UnitPriceMarkUp", adDecimal, adParamInput, , UnitPriceMarkUp)
                                      cmd.Parameters("@UnitPriceMarkUp").Precision = 18
                                      cmd.Parameters("@UnitPriceMarkUp").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@Price1MarkUp", adDecimal, adParamInput, , Price1MarkUp)
                                      cmd.Parameters("@Price1MarkUp").Precision = 18
                                      cmd.Parameters("@Price1MarkUp").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@Price2MarkUp", adDecimal, adParamInput, , Price2MarkUp)
                                      cmd.Parameters("@Price2MarkUp").Precision = 18
                                      cmd.Parameters("@Price2MarkUp").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@Price3MarkUp", adDecimal, adParamInput, , Price3MarkUp)
                                      cmd.Parameters("@Price3MarkUp").Precision = 18
                                      cmd.Parameters("@Price3MarkUp").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@UnitCost", adDecimal, adParamInput, , Val(Replace(txtCostingInfo_AverageCost.Text, ",", "")))
                                      cmd.Parameters("@UnitCost").Precision = 18
                                      cmd.Parameters("@UnitCost").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@ReorderPoint", adDecimal, adParamInput, , Val(Replace(txtReorderPoint.Text, ",", "")))
                                      cmd.Parameters("@ReorderPoint").Precision = 18
                                      cmd.Parameters("@ReorderPoint").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@ReorderQuantity", adDecimal, adParamInput, , Val(Replace(txtReorderQuantity.Text, ",", "")))
                                      cmd.Parameters("@ReorderQuantity").Precision = 18
                                      cmd.Parameters("@ReorderQuantity").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@Uom", adVarChar, adParamInput, 50, cmbStorageInfo_Uom.Text)
                If cmbVendor.ListIndex = 0 Then
                    cmd.Parameters.Append cmd.CreateParameter("@VendorId", adInteger, adParamInput, , Null)
                Else
                    cmd.Parameters.Append cmd.CreateParameter("@VendorId", adInteger, adParamInput, , cmbVendor.ItemData(cmbVendor.ListIndex))
                End If
                If PicPath <> "" Then
                    cmd.Parameters.Append cmd.CreateParameter("@Image", adVarChar, adParamInput, 400, App.path & "\images\" & txtBasicInfo_Name.Text & ".jpg")
                Else
                    cmd.Parameters.Append cmd.CreateParameter("@Image", adVarChar, adParamInput, 400, "")
                End If
                         
                If ProductId = 0 Then
                    cmd.CommandText = "BASE_Product_Insert"
                    cmd.Execute
                    ProductId = cmd.Parameters("@ProductId")
                    isInsert = True
                    If PicPath <> "" Then
                        FileCopy DialogBox.filename, App.path & "\images\" & txtBasicInfo_Name.Text & ".jpg"
                    End If
                    
                    'Audit Trail
                    SavePOSAuditTrail UserId, WorkstationId, "", "Created product: " & txtBasicInfo_Name.Text, "INVENTORY"
                Else
                    cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
                    cmd.CommandText = "BASE_Product_Update"
                    cmd.Execute
                    isInsert = False
                    If PicPath <> "" Then
                        Dim filename As String
                        filename = GetFileNameFromPath(PicPath)
                        
                        If Left(filename, Len(filename) - 4) <> txtBasicInfo_Name.Text Then
                            FileCopy PicPath, App.path & "\images\" & txtBasicInfo_Name.Text & ".jpg"
                        End If
                    End If
                    
                    'Audit Trail
                    SavePOSAuditTrail UserId, WorkstationId, "", "Updated product: " & txtBasicInfo_Name.Text, "INVENTORY"
                End If
                
                
                'UNIT OF MEASURE
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.Parameters.Append cmd.CreateParameter("@Uom", adVarChar, adParamInput, 50, cmbStorageInfo_Uom.Text)
                cmd.CommandText = "BASE_Uom_Insert"
                cmd.Execute
                
                'INV_UomPricing/Conversion
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.Parameters.Append cmd.CreateParameter("@UomConversionId", adInteger, adParamInputOutput, , 0)
                cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , ProductId)
                cmd.Parameters.Append cmd.CreateParameter("@UomId", adInteger, adParamInput, , cmbStorageInfo_Uom.ItemData(cmbStorageInfo_Uom.ListIndex))
                cmd.Parameters.Append cmd.CreateParameter("@ToQty", adDecimal, adParamInput, , Null)
                                      cmd.Parameters("@ToQty").NumericScale = 2
                                      cmd.Parameters("@ToQty").Precision = 18
                cmd.Parameters.Append cmd.CreateParameter("@Price", adDecimal, adParamInput, , Val(Replace(txtSalesInfo_Price.Text, ",", "")))
                                      cmd.Parameters("@Price").NumericScale = 2
                                      cmd.Parameters("@Price").Precision = 18
                cmd.Parameters.Append cmd.CreateParameter("@byQuantity", adDecimal, adParamInput, , 1)
                                      cmd.Parameters("@byQuantity").NumericScale = 2
                                      cmd.Parameters("@byQuantity").Precision = 18
                cmd.CommandText = "INV_UomConversion_Insert"
                cmd.Execute
                
                'INVENTORY
                For Each item In lvInventory.ListItems
                    Set cmd = New ADODB.Command
                    cmd.ActiveConnection = con
                    cmd.CommandType = adCmdStoredProc

                    cmd.Parameters.Append cmd.CreateParameter("@InventoryId", adInteger, adParamInputOutput, , Val(item.Text))
                    cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , ProductId)
                    cmd.Parameters.Append cmd.CreateParameter("@LocationId", adInteger, adParamInput, , Val(item.SubItems(1)))
                    cmd.Parameters.Append cmd.CreateParameter("@Quantity", adDecimal, adParamInput, , Val(Replace(item.SubItems(5), ",", "")))
                                          cmd.Parameters("@Quantity").Precision = 18
                                          cmd.Parameters("@Quantity").NumericScale = 2
                    If Trim(item.Text) = "" Then
                        cmd.CommandText = "BASE_Inventory_Insert"
                        cmd.Execute
                        item.Text = cmd.Parameters("@InventoryId")
                        item.SubItems(2) = ProductId
                    Else
'                        cmd.CommandText = "BASE_Inventory_Update"
'                        cmd.Execute
                    End If
                Next
                
                'DELETE INVENTORY ITEMS
                Dim i As Long
                For i = 0 To 10000
                    If deleteCtr(i) = 0 Then Exit For
                    Set cmd = New ADODB.Command
                    cmd.ActiveConnection = con
                    cmd.CommandType = adCmdStoredProc
                    cmd.Parameters.Append cmd.CreateParameter("@InventoryId", adInteger, adParamInput, , deleteCtr(i))
                    cmd.CommandText = "BASE_Inventory_Delete"
                    cmd.Execute
                Next i
                                
                Dim isFound As Boolean
                isFound = False
                For Each item In lvSearch.ListItems
                    If ProductId = item.Text Then
                        item.SubItems(1) = txtBasicInfo_ItemCode.Text
                        item.SubItems(2) = txtBasicInfo_Name.Text
                        isFound = True
                        item.Selected = True
                        item.EnsureVisible
                        Exit For
                    End If
                Next
                If isFound = False Then
                    Set item = lvSearch.ListItems.add(, , ProductId)
                        item.SubItems(1) = txtBasicInfo_ItemCode.Text
                        item.SubItems(2) = txtBasicInfo_Name.Text
                        item.Selected = True
                        item.EnsureVisible
                End If
                
                con.CommitTrans
                BASE_ContainerFrm.statusBar_Main.Panels(1).Text = MessageCodes(1) & " " & MessageCodes(0)
                con.Close
            End If
        Case 4 ' Delete
            If ProductId <> 0 Then
                Set con = New ADODB.Connection
                Set cmd = New ADODB.Command
                con.ConnectionString = ConnString
                con.Open
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "BASE_Product_Delete"
                cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , ProductId)
                If isActive = True Then
                    cmd.Parameters.Append cmd.CreateParameter("@isActive", adBoolean, adParamInput, , "False")
                Else
                    cmd.Parameters.Append cmd.CreateParameter("@isActive", adBoolean, adParamInput, , "True")
                End If
                con.BeginTrans
                cmd.Execute
                con.CommitTrans
                con.Close
                If isActive = True Then
                    BASE_ContainerFrm.statusBar_Main.Panels(1).Text = MessageCodes(1) & " " & MessageCodes(4)
                    isActive = False
                    tb_Standard.Buttons(4).Caption = "Activate"
                    tb_Standard.Buttons(4).Image = 4
                    isActivated (False)
                Else
                    BASE_ContainerFrm.statusBar_Main.Panels(1).Text = MessageCodes(1) & " " & MessageCodes(5)
                    isActive = True
                    tb_Standard.Buttons(4).Caption = "Delete"
                    tb_Standard.Buttons(4).Image = 3
                    isActivated (True)
                End If
            End If
        Case 6 ' COPY
            On Error Resume Next
            If ProductId = 0 Then Exit Sub
            txtBasicInfo_ItemCode.SetFocus
            'txtBasicInfo_ItemCode.text = ""
            txtBasicInfo_ItemCode.BackColor = &HC0C0FF
            ProductId = 0
            
            For Each item In lvInventory.ListItems
                item.Text = ""
            Next
    End Select
    Exit Sub
ErrHandler:
    con.RollbackTrans
    If IsNumeric(Err.Description) = True Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(Err.Description)
        BASE_ContainerFrm.statusBar_Main.Panels(1).Text = ErrorCodes(0) & " " & ErrorCodes(Err.Description)
    Else
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & Err.Description
        BASE_ContainerFrm.statusBar_Main.Panels(1).Text = ErrorCodes(0) & " " & Err.Description
    End If
        GLOBAL_MessageFrm.Show (1)
End Sub

Private Sub txtBasicInfo_ItemCode_Change()
    If Trim(txtBasicInfo_ItemCode.Text) = "" Then
        txtBasicInfo_ItemCode.BackColor = &HC0C0FF
    Else
        txtBasicInfo_ItemCode.BackColor = vbWhite
    End If
End Sub

Private Sub txtBasicInfo_Name_Change()
    If Trim(txtBasicInfo_Name.Text) = "" Then
        txtBasicInfo_Name.BackColor = &HC0C0FF
    Else
        txtBasicInfo_Name.BackColor = vbWhite
    End If
End Sub

Private Sub txtCostingInfo_AverageCost_Change()
    If IsNumeric(txtCostingInfo_AverageCost.Text) = False And Trim(txtCostingInfo_AverageCost.Text) <> "" Then
        txtCostingInfo_AverageCost.BackColor = &HC0C0FF
    Else
        txtCostingInfo_AverageCost.BackColor = vbWhite
    End If
End Sub

Private Sub txtReorderPoint_GotFocus()
    selectText txtReorderPoint
End Sub

Private Sub txtReorderQuantity_GotFocus()
    selectText txtReorderQuantity
End Sub

Private Sub txtSalesInfo_BCP_Change()
    If IsNumeric(txtSalesInfo_BCP.Text) = False And Trim(txtSalesInfo_BCP.Text) <> "" Then
        txtSalesInfo_BCP.BackColor = &HC0C0FF
    Else
        txtSalesInfo_BCP.BackColor = vbWhite
    End If
End Sub

Private Sub txtSalesInfo_BCP_GotFocus()
    selectText txtSalesInfo_BCP
End Sub

Private Sub txtSalesInfo_DP_Change()
    If IsNumeric(txtSalesInfo_DP.Text) = False And Trim(txtSalesInfo_DP.Text) <> "" Then
        txtSalesInfo_DP.BackColor = &HC0C0FF
    Else
        txtSalesInfo_DP.BackColor = vbWhite
    End If
End Sub

Private Sub txtSalesInfo_DP_GotFocus()
    selectText txtSalesInfo_DP
End Sub

Private Sub txtSalesInfo_Price_Change()
    If IsNumeric(txtSalesInfo_Price.Text) = False And Trim(txtSalesInfo_Price.Text) <> "" Then
        txtSalesInfo_Price.BackColor = &HC0C0FF
    Else
        txtSalesInfo_Price.BackColor = vbWhite
    End If
End Sub

Private Sub txtSalesInfo_Price_GotFocus()
    selectText txtSalesInfo_Price
End Sub

Private Sub txtSalesInfo_SP_Change()
    If IsNumeric(txtSalesInfo_SP.Text) = False And Trim(txtSalesInfo_SP.Text) <> "" Then
        txtSalesInfo_SP.BackColor = &HC0C0FF
    Else
        txtSalesInfo_SP.BackColor = vbWhite
    End If
End Sub

Private Sub txtSalesInfo_SP_GotFocus()
    selectText txtSalesInfo_SP
End Sub

Private Sub txtSalesInfoBCMarkUp_Change()
    If IsNumeric(txtSalesInfoBCMarkUp.Text) = False And Trim(txtSalesInfoBCMarkUp.Text) <> "" Then
        txtSalesInfoBCMarkUp.BackColor = &HC0C0FF
    Else
        txtSalesInfoBCMarkUp.BackColor = vbWhite
        Dim price As Double
        price = Val(Replace(txtCostingInfo_AverageCost.Text, ",", "")) * Val(Replace(txtSalesInfoBCMarkUp.Text, ",", ""))
        price = (price / 100) + Val(Replace(txtCostingInfo_AverageCost.Text, ",", ""))
        txtSalesInfo_BCP.Text = FormatNumber(price, 2, vbTrue, vbFalse)
    End If
End Sub

Private Sub txtSalesInfoBCMarkUp_GotFocus()
    selectText txtSalesInfoBCMarkUp
End Sub

Private Sub txtSalesInfoDPMarkUp_Change()
    If IsNumeric(txtSalesInfoDPMarkUp.Text) = False And Trim(txtSalesInfoDPMarkUp.Text) <> "" Then
        txtSalesInfoDPMarkUp.BackColor = &HC0C0FF
    Else
        txtSalesInfoDPMarkUp.BackColor = vbWhite
        Dim price As Double
        price = Val(Replace(txtCostingInfo_AverageCost.Text, ",", "")) * Val(Replace(txtSalesInfoDPMarkUp.Text, ",", ""))
        price = (price / 100) + Val(Replace(txtCostingInfo_AverageCost.Text, ",", ""))
        txtSalesInfo_DP.Text = FormatNumber(price, 2, vbTrue, vbFalse)
    End If
End Sub

Private Sub txtSalesInfoDPMarkUp_GotFocus()
    selectText txtSalesInfoDPMarkUp
End Sub

Private Sub txtSalesInfoMSMarkUp_Change()
    If IsNumeric(txtSalesInfoMSMarkUp.Text) = False And Trim(txtSalesInfoMSMarkUp.Text) <> "" Then
        txtSalesInfoMSMarkUp.BackColor = &HC0C0FF
    Else
        txtSalesInfoMSMarkUp.BackColor = vbWhite
        Dim price As Double
        price = Val(Replace(txtCostingInfo_AverageCost.Text, ",", "")) * Val(Replace(txtSalesInfoMSMarkUp.Text, ",", ""))
        price = (price / 100) + Val(Replace(txtCostingInfo_AverageCost.Text, ",", ""))
        txtSalesInfo_SP.Text = FormatNumber(price, 2, vbTrue, vbFalse)
    End If
End Sub

Private Sub txtSalesInfoMSMarkUp_GotFocus()
    selectText txtSalesInfoMSMarkUp
End Sub

Private Sub txtSalesInfoSRPMarkUp_Change()
    If IsNumeric(txtSalesInfoSRPMarkUp.Text) = False And Trim(txtSalesInfoSRPMarkUp.Text) <> "" Then
        txtSalesInfoSRPMarkUp.BackColor = &HC0C0FF
    Else
        txtSalesInfoSRPMarkUp.BackColor = vbWhite
        'compute SRP Mark-up
        Dim price As Double
        price = Val(Replace(txtCostingInfo_AverageCost.Text, ",", "")) * Val(Replace(txtSalesInfoSRPMarkUp.Text, ",", ""))
        price = (price / 100) + Val(Replace(txtCostingInfo_AverageCost.Text, ",", ""))
        txtSalesInfo_Price.Text = FormatNumber(price, 2, vbTrue, vbFalse)
    End If
End Sub

Private Sub txtSalesInfoSRPMarkUp_GotFocus()
    selectText txtSalesInfoSRPMarkUp
End Sub

Private Sub txtSearch_ItemCode_Change()
    btnSearch_Click
End Sub

Private Sub txtSearch_Name_Change()
    btnSearch_Click
End Sub
