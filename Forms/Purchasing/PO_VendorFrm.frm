VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form PO_VendorFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Supplier"
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
      Left            =   -9960
      Picture         =   "PO_VendorFrm.frx":0000
      ScaleHeight     =   1860
      ScaleWidth      =   5250
      TabIndex        =   37
      Top             =   2640
      Visible         =   0   'False
      Width           =   5250
   End
   Begin VB.PictureBox pic_Cancelled 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   -9960
      Picture         =   "PO_VendorFrm.frx":8F91
      ScaleHeight     =   2295
      ScaleWidth      =   6195
      TabIndex        =   34
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
      Left            =   -9960
      Picture         =   "PO_VendorFrm.frx":19C2D
      ScaleHeight     =   2295
      ScaleWidth      =   6195
      TabIndex        =   33
      Top             =   2640
      Width           =   6195
   End
   Begin VB.Frame LayoutFrame_Search 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   4575
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
         TabIndex        =   27
         Top             =   1320
         Width           =   1215
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
         TabIndex        =   26
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
         TabIndex        =   25
         Top             =   840
         Width           =   3015
      End
      Begin MSComctlLib.ListView lvSearch 
         Height          =   7095
         Left            =   120
         TabIndex        =   28
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
            Text            =   "CustomerId"
            Object.Width           =   2540
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
      End
      Begin VB.TextBox txtSearch_Code 
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
         Top             =   2520
         Visible         =   0   'False
         Width           =   3015
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
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name/Code"
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
         Top             =   480
         Width           =   1125
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
         Top             =   840
         Width           =   570
      End
      Begin VB.Label Label14 
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
         TabIndex        =   29
         Top             =   2520
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.Frame Body_Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   4635
      TabIndex        =   0
      Top             =   0
      Width           =   10485
      Begin VB.ComboBox cmbOrderStatus 
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
         ItemData        =   "PO_VendorFrm.frx":222BB
         Left            =   1680
         List            =   "PO_VendorFrm.frx":222BD
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   5520
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton btnSearch_OrderHistory 
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
         Height          =   345
         Left            =   3600
         TabIndex        =   17
         Top             =   5500
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame FRE_Info 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   4335
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   10095
         Begin VB.TextBox txtTIN 
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
            Left            =   1440
            MaxLength       =   500
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   1200
            Width           =   3375
         End
         Begin VB.TextBox txtAddress 
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
            Height          =   1050
            Left            =   1440
            MaxLength       =   50
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   3120
            Width           =   3375
         End
         Begin VB.TextBox txtMobile 
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
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   6
            Top             =   2760
            Width           =   3375
         End
         Begin VB.TextBox txtPhone 
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
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   5
            Top             =   2400
            Width           =   3375
         End
         Begin VB.TextBox txtName 
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
            Left            =   1440
            MaxLength       =   500
            ScrollBars      =   2  'Vertical
            TabIndex        =   3
            Top             =   840
            Width           =   3375
         End
         Begin VB.TextBox txtCode 
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
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   2
            Top             =   480
            Width           =   3375
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TIN"
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
            Left            =   360
            TabIndex        =   40
            Top             =   1200
            Width           =   315
         End
         Begin VB.Label Label10 
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
            Left            =   360
            TabIndex        =   16
            Top             =   3120
            Width           =   750
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile"
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
            Left            =   360
            TabIndex        =   15
            Top             =   2760
            Width           =   660
         End
         Begin VB.Label Label5 
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
            Left            =   360
            TabIndex        =   14
            Top             =   2400
            Width           =   600
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact Info"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   345
            Left            =   120
            TabIndex        =   13
            Top             =   1920
            Width           =   1425
         End
         Begin VB.Label Label2 
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
            Left            =   360
            TabIndex        =   12
            Top             =   1560
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
            Left            =   1440
            TabIndex        =   11
            Top             =   1560
            Width           =   3375
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
            Left            =   360
            TabIndex        =   10
            Top             =   840
            Width           =   555
         End
         Begin VB.Label Label7 
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
            Left            =   360
            TabIndex        =   9
            Top             =   480
            Width           =   480
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
            ForeColor       =   &H000080FF&
            Height          =   345
            Left            =   120
            TabIndex        =   8
            Top             =   0
            Width           =   1125
         End
      End
      Begin MSComctlLib.ListView lvCustomer 
         Height          =   2655
         Left            =   -9999
         TabIndex        =   19
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
      Begin MSComctlLib.Toolbar tb_Standard 
         Height          =   330
         Left            =   0
         TabIndex        =   20
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
            NumButtons      =   4
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
         EndProperty
      End
      Begin MSComctlLib.ListView lvOrderHistory 
         Height          =   2895
         Left            =   360
         TabIndex        =   21
         Top             =   6000
         Visible         =   0   'False
         Width           =   9900
         _ExtentX        =   17463
         _ExtentY        =   5106
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "SalesOrderId"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Order #"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Due"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Interest"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Total"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Balance"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order History"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   345
         Left            =   240
         TabIndex        =   39
         Top             =   4920
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Label Label12 
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
         Left            =   600
         TabIndex        =   22
         Top             =   5520
         Visible         =   0   'False
         Width           =   570
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   14445
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
            Picture         =   "PO_VendorFrm.frx":222BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PO_VendorFrm.frx":28B21
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PO_VendorFrm.frx":2F383
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PO_VendorFrm.frx":35BE5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PO_VendorFrm.frx":35E5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PO_VendorFrm.frx":364CB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvItemList 
      Height          =   2655
      Left            =   -9960
      TabIndex        =   38
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
   Begin VB.CommandButton btnPaid 
      BackColor       =   &H0080FF80&
      Caption         =   "PAID IN FULL"
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
      Left            =   9285
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   7605
      Visible         =   0   'False
      Width           =   2535
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
      Left            =   9285
      TabIndex        =   36
      Top             =   7200
      Width           =   2535
   End
End
Attribute VB_Name = "PO_VendorFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public VendorId As Long
Dim isActive As Boolean
Private Sub Initialize()
    Dim txtControl As Control

    For Each txtControl In Me.Controls
        If TypeOf txtControl Is TextBox And txtControl.Name <> "txtSearch_Name" Then
            txtControl.Text = ""
        End If
    Next
    
    txtCode.BackColor = &HC0C0FF
    txtName.BackColor = &HC0C0FF
    
    isActivated (True)
    
    lblBalance.Caption = "0.00"
    lvOrderHistory.ListItems.Clear
    
    isActive = True
    tb_Standard.Buttons(4).Image = 3
    tb_Standard.Buttons(4).Caption = "Delete"
    
    VendorId = 0
    
    On Error Resume Next
    txtCode.SetFocus
    
    
End Sub
Private Sub isActivated(value As Boolean)
    txtCode.Enabled = value
    txtName.Enabled = value
    txtPhone.Enabled = value
    txtMobile.Enabled = value
    txtAddress.Enabled = value
    FRE_Info.Enabled = value
End Sub

Public Sub Populate(ByVal data As String)
    Set rec = New ADODB.Recordset
    Dim item As MSComctlLib.ListItem
    Select Case data
        Case "Vendor"
            Set rec = Global_Data("Vendor")
            lvSearch.ListItems.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    Set item = lvSearch.ListItems.add(, , rec!VendorId)
                        item.SubItems(1) = rec!VendorCode
                        item.SubItems(2) = rec!Name
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
        Case "OrderStatus"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Status")
            cmbOrderStatus.Clear
            cmbOrderStatus.AddItem ""
            cmbOrderStatus.ItemData(cmbOrderStatus.NewIndex) = 0
            cmbOrderStatus.ListIndex = 0
            If Not rec.EOF Then
                Do Until rec.EOF
                    cmbOrderStatus.AddItem rec!Status
                    cmbOrderStatus.ItemData(cmbOrderStatus.NewIndex) = rec!StatusId
                    rec.MoveNext
                Loop
            End If
        Case "VendorGet"
            Set con = New ADODB.Connection
            Set rec = New ADODB.Recordset
            Set cmd = New ADODB.Command
            
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "BASE_Vendor_Get"
            cmd.Parameters.Append cmd.CreateParameter("@VendorId", adInteger, adParamInput, , VendorId)
            Set rec = cmd.Execute
            If Not rec.EOF Then
                On Error Resume Next
                txtCode.Text = rec!VendorCode
                txtName.Text = rec!Name
                lblBalance.Caption = FormatNumber(rec!OutStandingBalance, 2, vbTrue, vbFalse)
                txtMobile.Text = rec!Mobile
                txtPhone.Text = rec!Phone
                txtAddress.Text = rec!Address
                isActive = rec!isActive
                txtTIN.Text = rec!TIN
                If rec!isActive = "False" Then
                    tb_Standard.Buttons(4).Caption = "Activate"
                    tb_Standard.Buttons(4).Image = 6
                    isActivated (False)
                Else
                    tb_Standard.Buttons(4).Caption = "Delete"
                    tb_Standard.Buttons(4).Image = 3
                    isActivated (True)
                End If
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
            cmd.Parameters.Append cmd.CreateParameter("@VendorId", adInteger, adParamInput, , VendorId)
            Set rec = cmd.Execute
            lvOrderHistory.ListItems.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    Set item = lvOrderHistory.ListItems.add(, , rec!inventoryId)
                        item.SubItems(1) = rec!LocationId
                        item.SubItems(2) = rec!VendorId
                        item.SubItems(3) = rec!Location
                        item.SubItems(5) = FormatNumber(rec!quantity, 2, vbTrue, vbFalse)
                    rec.MoveNext
                Loop
            End If
            con.Close
    End Select
End Sub
Private Function isValidated() As Boolean
    isValidated = False
    'CHECK EMPTY FIELDS
    If Trim(txtCode.Text) = "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(1)
        GLOBAL_MessageFrm.Show (1)
        txtCode.SetFocus
    ElseIf Trim(txtName.Text) = "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(2)
        GLOBAL_MessageFrm.Show (1)
        txtName.SetFocus
    Else
        isValidated = True
    End If
End Function

Public Sub btnSearch_Click()
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_Vendor_Search"
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 500, txtSearch_name.Text)
    If Trim(txtSearch_Code.Text) <> "" Then
        cmd.Parameters.Append cmd.CreateParameter("@VendorCode", adVarChar, adParamInput, 50, txtSearch_Code.Text)
    Else
        cmd.Parameters.Append cmd.CreateParameter("@VendorCode", adVarChar, adParamInput, 50, Null)
    End If
    If cmbSearch_Status.ListIndex <> 0 Then
        cmd.Parameters.Append cmd.CreateParameter("@isActive", adInteger, adParamInput, , cmbSearch_Status.ItemData(cmbSearch_Status.ListIndex))
    End If
    Set rec = cmd.Execute
    lvSearch.ListItems.Clear
    Dim item As MSComctlLib.ListItem
    If Not rec.EOF Then
        Do Until rec.EOF
            'If rec!isactive = "True" Then
                Set item = lvSearch.ListItems.add(, , rec!VendorId)
                    item.SubItems(1) = rec!VendorCode
                    item.SubItems(2) = rec!Name
                    
            'End If
            rec.MoveNext
        Loop
    End If
    'DistinctList lvSearch
    con.Close
End Sub

'Private Sub cmbBasicInfo_Type_Click()
'    If cmbBasicInfo_Type.ListIndex <> 0 Then
'        Frame_VendorDetails1.Visible = False
'        Frame_VendorDetails2.Visible = False
'        isService = True
'    Else
'        Frame_VendorDetails1.Visible = True
'        Frame_VendorDetails2.Visible = True
'        isService = False
'    End If
'End Sub

Private Sub btnSearch_OrderHistory_Click()
'    If VendorId = 0 Then Exit Sub
'
'    Set con = New ADODB.Connection
'    Set rec = New ADODB.Recordset
'    Set cmd = New ADODB.Command
'
'    con.ConnectionString = ConnString
'    con.Open
'    cmd.ActiveConnection = con
'    cmd.CommandType = adCmdStoredProc
'    cmd.CommandText = "PO_PurchaseOrder_Get"
'    cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , Null)
'    If cmbOrderStatus.ItemData(cmbOrderStatus.ListIndex) = 0 Then
'        cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , Null)
'    Else
'        cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , cmbOrderStatus.ItemData(cmbOrderStatus.ListIndex))
'    End If
'    cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInput, 50, Null)
'    cmd.Parameters.Append cmd.CreateParameter("@VendorId", adInteger, adParamInput, , VendorId)
'    Dim item As MSComctlLib.ListItem
'    Set rec = cmd.Execute
'    lvOrderHistory.ListItems.Clear
'    If Not rec.EOF Then
'        Do Until rec.EOF
'            Set item = lvOrderHistory.ListItems.add(, , rec!PurchaseOrderId)
'                item.SubItems(1) = rec!OrderNumber
'                item.SubItems(2) = Format(rec!Date, "MM/DD/YY")
''                item.SubItems(3) = Format(rec!DueDate, "MM/DD/YY")
'                item.SubItems(4) = FormatNumber(rec!subtotal, 2, vbTrue, vbFalse)
'                item.SubItems(5) = FormatNumber(rec!Interest, 2, vbTrue, vbFalse)
'                item.SubItems(6) = FormatNumber(rec!total, 2, vbTrue, vbFalse)
'                item.SubItems(7) = FormatNumber(rec!balance, 2, vbTrue, vbFalse)
'                item.SubItems(8) = rec!Status
'            rec.MoveNext
'        Loop
'    End If
'    con.Close
End Sub

Private Sub Form_Load()
    Initialize
    
    '****** REGION Listview Columns *********
    lvSearch.ColumnHeaders(1).width = lvSearch.width * 0
    lvSearch.ColumnHeaders(2).width = lvSearch.width * 0.33
    lvSearch.ColumnHeaders(3).width = lvSearch.width * 0.62
        
    lvOrderHistory.ColumnHeaders(1).width = lvOrderHistory.width * 0
    lvOrderHistory.ColumnHeaders(2).width = lvOrderHistory.width * 0.125
    lvOrderHistory.ColumnHeaders(3).width = lvOrderHistory.width * 0.105
    lvOrderHistory.ColumnHeaders(4).width = lvOrderHistory.width * 0.105
    lvOrderHistory.ColumnHeaders(5).width = lvOrderHistory.width * 0.13
    lvOrderHistory.ColumnHeaders(6).width = lvOrderHistory.width * 0.125
    lvOrderHistory.ColumnHeaders(7).width = lvOrderHistory.width * 0.13
    lvOrderHistory.ColumnHeaders(8).width = lvOrderHistory.width * 0.13
    lvOrderHistory.ColumnHeaders(9).width = lvOrderHistory.width * 0.123
    '****************************************
    
    Populate "Vendor"
    Populate "Status"
    Populate "OrderStatus"
    btnSearch_Click
End Sub

Private Sub Label18_Click()

End Sub

Private Sub lblInventory_MoreLocations_Click()
    CenterChildForm INV_LocationFrm
    INV_LocationFrm.Show
End Sub

'Private Sub lvOrderHistory_DblClick()
'    With lvOrderHistory
'        If .ListItems.Count > 0 Then
'            Dim i As String
'            i = InputBox("Input quantity.", "Quantity", lvOrderHistory.SelectedItem.SubItems(5))
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
'End Sub

'Private Sub lvOrderHistory_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'    Case vbKeyDelete
'        If lvOrderHistory.ListItems.Count > 0 Then
'            If lvOrderHistory.SelectedItem.SubItems(1) <> "1" Then 'NOT Default Location
'                If lvOrderHistory.SelectedItem.text <> "" Then 'Existing data
'                        deleteCtr(ctr) = Val(lvOrderHistory.SelectedItem.text)
'                        ctr = ctr + 1
'                        lvOrderHistory.ListItems.Remove (lvOrderHistory.SelectedItem.Index)
'                Else
'                    lvOrderHistory.ListItems.Remove (lvOrderHistory.SelectedItem.Index)
'                End If
'            End If
'        End If
'    Case 13
'        Call lvOrderHistory_DblClick
'    End Select
'End Sub

Private Sub lvSearch_ItemClick(ByVal item As MSComctlLib.ListItem)
    With lvSearch
        If .ListItems.Count > 0 Then
            VendorId = .SelectedItem.Text
            Populate "VendorGet"
            btnSearch_OrderHistory_Click
        End If
    End With
End Sub

Private Sub tb_Standard_ButtonClick(ByVal Button As MSComctlLib.Button)
    If EditAccessRights(11) = False Then
        MsgBox ErrorCodes(74), vbCritical, "Limited Rights"
        Exit Sub
    End If
    Dim item As MSComctlLib.ListItem
    Select Case Button.Index
        Case 1 'New
            Initialize
        Case 2 'Save
            If isActive = False Then Exit Sub
            If isValidated = True Then
                On Error GoTo ErrHandler
                Set con = New ADODB.Connection
                Set cmd = New ADODB.Command
                
                'SAVE MAIN Vendor DETAILS
                con.ConnectionString = ConnString
                con.Open
                con.BeginTrans
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                
                cmd.Parameters.Append cmd.CreateParameter("@VendorId", adInteger, adParamInputOutput, , VendorId)
                cmd.Parameters.Append cmd.CreateParameter("@VendorCode", adVarChar, adParamInput, 50, txtCode.Text)
                cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 500, txtName.Text)
                cmd.Parameters.Append cmd.CreateParameter("@TIN", adVarChar, adParamInput, 500, txtTIN.Text)
                cmd.Parameters.Append cmd.CreateParameter("@Phone", adVarChar, adParamInput, 50, txtPhone.Text)
                cmd.Parameters.Append cmd.CreateParameter("@Mobile", adVarChar, adParamInput, 50, txtMobile.Text)
                cmd.Parameters.Append cmd.CreateParameter("@Address", adVarChar, adParamInput, 500, txtAddress.Text)
                                
                If VendorId = 0 Then
                    cmd.CommandText = "BASE_Vendor_Insert"
                    cmd.Execute
                    VendorId = cmd.Parameters("@VendorId")
                    
                    SavePOSAuditTrail UserId, WorkstationId, "", "Created new supplier: " & txtName.Text, "PURCHASING"
                Else
                    'cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , 1) 'NOT SET
                    cmd.CommandText = "BASE_Vendor_Update"
                    cmd.Execute
                    
                    SavePOSAuditTrail UserId, WorkstationId, "", "Updated supplier: " & txtName.Text, "PURCHASING"
                End If
                
                con.CommitTrans
                con.Close
                
                BASE_ContainerFrm.statusBar_Main.Panels(1).Text = MessageCodes(1) & " " & MessageCodes(0)
        
                Dim isFound As Boolean
                isFound = False
                For Each item In lvSearch.ListItems
                    If VendorId = item.Text Then
                        item.SubItems(1) = txtCode.Text
                        item.SubItems(2) = txtName.Text
                        isFound = True
                        item.Selected = True
                        item.EnsureVisible
                        Exit For
                    End If
                Next
                If isFound = False Then
                    Set item = lvSearch.ListItems.add(, , VendorId)
                        item.SubItems(1) = txtCode.Text
                        item.SubItems(2) = txtName.Text
                        item.Selected = True
                        item.EnsureVisible
                End If
            End If
        Case 4 ' Delete
            If VendorId <> 0 Then
                Set con = New ADODB.Connection
                Set cmd = New ADODB.Command
                con.ConnectionString = ConnString
                con.Open
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "BASE_Vendor_Delete"
                cmd.Parameters.Append cmd.CreateParameter("@VendorId", adInteger, adParamInput, , VendorId)
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
                    tb_Standard.Buttons(4).Image = 6
                    isActivated (False)
                Else
                    BASE_ContainerFrm.statusBar_Main.Panels(1).Text = MessageCodes(1) & " " & MessageCodes(5)
                    isActive = True
                    tb_Standard.Buttons(4).Caption = "Delete"
                    tb_Standard.Buttons(4).Image = 3
                    isActivated (True)
                End If
                
                SavePOSAuditTrail UserId, WorkstationId, "", "Deactivated supplier: " & txtName.Text, "PURCHASING"
            End If
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

Private Sub txtCode_Change()
    If Trim(txtCode.Text) = "" Then
        txtCode.BackColor = &HC0C0FF
    Else
        txtCode.BackColor = vbWhite
    End If
End Sub

Private Sub txtName_Change()
    If Trim(txtName.Text) = "" Then
        txtName.BackColor = &HC0C0FF
    Else
        txtName.BackColor = vbWhite
    End If
End Sub

Private Sub txtSearch_VendorCode_Change()
    btnSearch_Click
End Sub

Private Sub txtSearch_Name_Change()
    btnSearch_Click
End Sub


