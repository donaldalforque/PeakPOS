VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form SO_CustomerFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Customer"
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
            Picture         =   "SO_CustomerFrm.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_CustomerFrm.frx":6862
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_CustomerFrm.frx":D0C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_CustomerFrm.frx":13926
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_CustomerFrm.frx":13B9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_CustomerFrm.frx":1420C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Body_Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   4640
      TabIndex        =   30
      Top             =   0
      Width           =   10485
      Begin VB.Frame FRE_Info 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   4695
         Left            =   120
         TabIndex        =   37
         Top             =   600
         Width           =   10095
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
            Left            =   7080
            MaxLength       =   50
            TabIndex        =   11
            Top             =   2640
            Width           =   2655
         End
         Begin VB.CheckBox chkMember 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Member of an organization"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7080
            TabIndex        =   12
            Top             =   3600
            Width           =   2895
         End
         Begin VB.TextBox txtSalesman 
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
            Top             =   1520
            Width           =   3375
         End
         Begin VB.ComboBox cmbCity 
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
            TabIndex        =   54
            Top             =   4200
            Width           =   3375
         End
         Begin VB.TextBox txtAvailablePoints 
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
            Left            =   7080
            MaxLength       =   50
            TabIndex        =   9
            Top             =   840
            Width           =   2655
         End
         Begin VB.TextBox txtCreditLine 
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
            Left            =   7080
            MaxLength       =   50
            TabIndex        =   10
            Top             =   2280
            Width           =   2655
         End
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
            ItemData        =   "SO_CustomerFrm.frx":1AA6E
            Left            =   7080
            List            =   "SO_CustomerFrm.frx":1AA78
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1200
            Width           =   2655
         End
         Begin VB.ComboBox cmbCollector 
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
            ItemData        =   "SO_CustomerFrm.frx":1AA89
            Left            =   -9999
            List            =   "SO_CustomerFrm.frx":1AA99
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1560
            Visible         =   0   'False
            Width           =   2655
         End
         Begin VB.TextBox txtCardNumber 
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
            Left            =   7080
            MaxLength       =   50
            TabIndex        =   8
            Top             =   480
            Width           =   2655
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
            TabIndex        =   1
            Top             =   480
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
            TabIndex        =   2
            Top             =   840
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
            TabIndex        =   4
            Top             =   2400
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
            TabIndex        =   5
            Top             =   2760
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
            TabIndex        =   6
            Top             =   3120
            Width           =   3375
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Interest Rate"
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
            Left            =   5400
            TabIndex        =   58
            Top             =   2640
            Width           =   1200
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Other Info"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   345
            Left            =   5160
            TabIndex        =   57
            Top             =   3240
            Width           =   1215
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Rep"
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
            TabIndex        =   56
            Top             =   1515
            Width           =   885
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "City"
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
            TabIndex        =   55
            Top             =   4200
            Width           =   360
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Credit Info"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   345
            Left            =   5160
            TabIndex        =   53
            Top             =   1800
            Width           =   1230
         End
         Begin VB.Label Label20 
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
            Left            =   5400
            TabIndex        =   52
            Top             =   2280
            Width           =   1005
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pricing Scheme"
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
            Left            =   5400
            TabIndex        =   51
            Top             =   1200
            Width           =   1410
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Available Points"
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
            Left            =   5400
            TabIndex        =   50
            Top             =   840
            Width           =   1500
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Card Number"
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
            Left            =   5400
            TabIndex        =   49
            Top             =   480
            Width           =   1230
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Collector"
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
            TabIndex        =   48
            Top             =   1560
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.Label Label15 
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
            ForeColor       =   &H0000C000&
            Height          =   345
            Left            =   5160
            TabIndex        =   47
            Top             =   0
            Width           =   1125
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
            ForeColor       =   &H0000C000&
            Height          =   345
            Left            =   120
            TabIndex        =   46
            Top             =   0
            Width           =   1125
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
            TabIndex        =   45
            Top             =   480
            Width           =   480
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
            TabIndex        =   44
            Top             =   840
            Width           =   555
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
            TabIndex        =   43
            Top             =   1200
            Width           =   3375
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
            TabIndex        =   42
            Top             =   1200
            Width           =   720
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
            ForeColor       =   &H0000C000&
            Height          =   345
            Left            =   120
            TabIndex        =   41
            Top             =   1920
            Width           =   1425
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
            TabIndex        =   40
            Top             =   2400
            Width           =   600
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
            TabIndex        =   39
            Top             =   2760
            Width           =   660
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
            TabIndex        =   38
            Top             =   3120
            Width           =   750
         End
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
         TabIndex        =   15
         Top             =   5745
         Width           =   1095
      End
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
         ItemData        =   "SO_CustomerFrm.frx":1AAA9
         Left            =   1680
         List            =   "SO_CustomerFrm.frx":1AAAB
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   5760
         Width           =   1815
      End
      Begin MSComctlLib.ListView lvCustomer 
         Height          =   2655
         Left            =   -9999
         TabIndex        =   31
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
         TabIndex        =   32
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
         Height          =   2655
         Left            =   360
         TabIndex        =   16
         Top             =   6240
         Width           =   9900
         _ExtentX        =   17463
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
         TabIndex        =   34
         Top             =   5760
         Width           =   570
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
         ForeColor       =   &H0000C000&
         Height          =   345
         Left            =   360
         TabIndex        =   33
         Top             =   5280
         Width           =   1590
      End
   End
   Begin VB.Frame LayoutFrame_Search 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   4575
      Begin MSComctlLib.ListView lvSearch 
         Height          =   7095
         Left            =   120
         TabIndex        =   21
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
         TabIndex        =   17
         Top             =   2520
         Visible         =   0   'False
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
         TabIndex        =   19
         Top             =   840
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
         TabIndex        =   18
         Top             =   480
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
         TabIndex        =   20
         Top             =   1320
         Width           =   1215
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
         TabIndex        =   36
         Top             =   2520
         Visible         =   0   'False
         Width           =   480
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
         TabIndex        =   29
         Top             =   480
         Width           =   1125
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
         TabIndex        =   28
         Top             =   80
         Width           =   795
      End
   End
   Begin VB.PictureBox picCompleted 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   -9960
      Picture         =   "SO_CustomerFrm.frx":1AAAD
      ScaleHeight     =   2295
      ScaleWidth      =   6195
      TabIndex        =   25
      Top             =   2640
      Width           =   6195
   End
   Begin VB.PictureBox pic_Cancelled 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   -9960
      Picture         =   "SO_CustomerFrm.frx":2313B
      ScaleHeight     =   2295
      ScaleWidth      =   6195
      TabIndex        =   24
      Top             =   2640
      Width           =   6195
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
      TabIndex        =   23
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
      TabIndex        =   22
      Top             =   7200
      Width           =   2535
   End
   Begin VB.PictureBox picPaid 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1860
      Left            =   -9960
      Picture         =   "SO_CustomerFrm.frx":33DD7
      ScaleHeight     =   1860
      ScaleWidth      =   5250
      TabIndex        =   0
      Top             =   2640
      Visible         =   0   'False
      Width           =   5250
   End
   Begin MSComctlLib.ListView lvItemList 
      Height          =   2655
      Left            =   -9960
      TabIndex        =   26
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
End
Attribute VB_Name = "SO_CustomerFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CustomerId As Integer
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
    txtAvailablePoints.Text = "0.00"
    lvOrderHistory.ListItems.Clear
    
    isActive = True
    tb_Standard.Buttons(4).Image = 3
    tb_Standard.Buttons(4).Caption = "Delete"
    
    CustomerId = 0
    
    cmbCollector.ListIndex = 0
    cmbPricingScheme.ListIndex = 0
    
    On Error Resume Next
    txtCode.SetFocus
    cmbCity.ListIndex = 0
    chkMember.value = Checked
End Sub
Private Sub isActivated(value As Boolean)
    txtCode.Enabled = value
    txtName.Enabled = value
    txtPhone.Enabled = value
    txtMobile.Enabled = value
    txtAddress.Enabled = value
    FRE_Info.Enabled = value
    cmbPricingScheme.Enabled = value
End Sub

Public Sub Populate(ByVal data As String)
    Set rec = New ADODB.Recordset
    Dim Item As MSComctlLib.ListItem
    Select Case data
        Case "Customer"
            Set rec = Global_Data("Customer")
            lvSearch.ListItems.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    Set Item = lvSearch.ListItems.add(, , rec!CustomerId)
                        Item.SubItems(1) = rec!CustomerCode
                        Item.SubItems(2) = rec!Name
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
        Case "City"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("City")
            cmbCity.Clear
            cmbCity.AddItem ""
            cmbCity.ItemData(cmbCity.NewIndex) = 0
            cmbCity.ListIndex = 0
            If Not rec.EOF Then
                Do Until rec.EOF
                    cmbCity.AddItem rec!City
                    cmbCity.ItemData(cmbCity.NewIndex) = rec!CityId
                    rec.MoveNext
                Loop
            End If
            cmbCity.ListIndex = 0
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
        Case "CustomerGet"
            Set con = New ADODB.Connection
            Set rec = New ADODB.Recordset
            Set cmd = New ADODB.Command
            
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "BASE_Customer_Get"
            cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInput, , CustomerId)
            Set rec = cmd.Execute
            If Not rec.EOF Then
                txtCode.Text = rec!CustomerCode
                txtName.Text = rec!Name
                txtCardnumber.Text = rec!CardNumber
                lblBalance.Caption = FormatNumber(rec!balance, 2, vbTrue, vbFalse)
                txtMobile.Text = rec!Mobile
                txtPhone.Text = rec!Phone
                txtAddress.Text = rec!Address
                isActive = rec!isActive
                txtSalesman.Text = rec!SalesRep
                txtCreditLine.Text = FormatNumber(rec!creditlimit, 2, vbTrue, vbFalse)
                txtInterestRate.Text = FormatNumber(rec!interestrate, 2, vbTrue, vbFalse)
                
                If rec!City = "" Then
                    cmbCity.ListIndex = 0
                Else
                    cmbCity.Text = rec!City
                End If
                If rec!PricingScheme = "" Then
                    cmbPricingScheme.ListIndex = 0
                Else
                    cmbPricingScheme.Text = rec!PricingScheme
                End If
                
                If rec!memberid = 1 Then chkMember.value = Checked
                If rec!memberid = 2 Then chkMember.value = Unchecked
                
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
            'cmbCost.text = rec!atcostprice
            On Error Resume Next
                cmbCollector.Text = rec!Collector
                
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
            cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInput, , CustomerId)
            Set rec = cmd.Execute
            lvOrderHistory.ListItems.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    Set Item = lvOrderHistory.ListItems.add(, , rec!inventoryId)
                        Item.SubItems(1) = rec!LocationId
                        Item.SubItems(2) = rec!CustomerId
                        Item.SubItems(3) = rec!Location
                        Item.SubItems(5) = FormatNumber(rec!quantity, 2, vbTrue, vbFalse)
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
Private Function isValidated() As Boolean
    isValidated = False
    'CHECK EMPTY FIELDS
'    If Trim(txtCode.Text) = "" Then
''        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(1)
''        GLOBAL_MessageFrm.Show (1)
''        txtCode.SetFocus
'        isValidated = True
   If Trim(txtName.Text) = "" Then
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
    cmd.CommandText = "BASE_Customer_Search"
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 500, txtSearch_Name.Text)
    If Trim(txtSearch_Code.Text) <> "" Then
        cmd.Parameters.Append cmd.CreateParameter("@CustomerCode", adVarChar, adParamInput, 50, txtSearch_Code.Text)
    Else
        cmd.Parameters.Append cmd.CreateParameter("@CustomerCode", adVarChar, adParamInput, 50, Null)
    End If
    If cmbSearch_Status.ListIndex <> 0 Then
        cmd.Parameters.Append cmd.CreateParameter("@isActive", adInteger, adParamInput, , cmbSearch_Status.ItemData(cmbSearch_Status.ListIndex))
    End If
    Set rec = cmd.Execute
    lvSearch.ListItems.Clear
    Dim Item As MSComctlLib.ListItem
    If Not rec.EOF Then
        Do Until rec.EOF
            'If rec!isactive = "True" Then
                Set Item = lvSearch.ListItems.add(, , rec!CustomerId)
                    Item.SubItems(1) = rec!CustomerCode
                    Item.SubItems(2) = rec!Name
                    
            'End If
            rec.MoveNext
        Loop
    End If
    'DistinctList lvSearch
    con.Close
End Sub

'Private Sub cmbBasicInfo_Type_Click()
'    If cmbBasicInfo_Type.ListIndex <> 0 Then
'        Frame_CustomerDetails1.Visible = False
'        Frame_CustomerDetails2.Visible = False
'        isService = True
'    Else
'        Frame_CustomerDetails1.Visible = True
'        Frame_CustomerDetails2.Visible = True
'        isService = False
'    End If
'End Sub

Private Sub btnSearch_OrderHistory_Click()
    If CustomerId = 0 Then Exit Sub
    
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "SO_SalesOrder_Get"
    cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , Null)
    If cmbOrderStatus.ItemData(cmbOrderStatus.ListIndex) = 0 Then
        cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , Null)
    Else
        cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , cmbOrderStatus.ItemData(cmbOrderStatus.ListIndex))
    End If
    cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInput, 50, "")
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 50, txtName.Text)
    cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInput, , CustomerId)
    Dim Item As MSComctlLib.ListItem
    Set rec = cmd.Execute
    lvOrderHistory.ListItems.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            Set Item = lvOrderHistory.ListItems.add(, , rec!SalesOrderId)
                Item.SubItems(1) = rec!OrderNumber
                Item.SubItems(2) = Format(rec!Date, "MM/DD/YY")
                Item.SubItems(3) = Format(rec!DueDate, "MM/DD/YY")
                Item.SubItems(4) = FormatNumber(rec!subtotal, 2, vbTrue, vbFalse)
                Item.SubItems(5) = FormatNumber(rec!Interest, 2, vbTrue, vbFalse)
                Item.SubItems(6) = FormatNumber(rec!Total, 2, vbTrue, vbFalse)
                Item.SubItems(7) = FormatNumber(rec!balance, 2, vbTrue, vbFalse)
                Item.SubItems(8) = rec!Status
            rec.MoveNext
        Loop
    End If
    con.Close
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
    
    Populate "Customer"
    Populate "Status"
    Populate "OrderStatus"
    Populate "PricingScheme"
    Populate "City"
    
    btnSearch_Click
    
    cmbPricingScheme.ListIndex = 0
End Sub

Private Sub lblInventory_MoreLocations_Click()
    CenterChildForm INV_LocationFrm
    INV_LocationFrm.Show
End Sub
Private Sub lvSearch_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With lvSearch
        If .ListItems.count > 0 Then
            CustomerId = .SelectedItem.Text
            Populate "CustomerGet"
            btnSearch_OrderHistory_Click
            
            'GET Available POINTS
            Set con = New ADODB.Connection
            Set rec = New ADODB.Recordset
            Set cmd = New ADODB.Command
            
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "POS_LoyaltyCard_Search"
            cmd.Parameters.Append cmd.CreateParameter("@CardNumber", adVarChar, adParamInput, 250, txtCardnumber.Text)
            Set rec = cmd.Execute
            txtAvailablePoints.Text = "0.00"
            If Not rec.EOF Then
                Do Until rec.EOF
                    txtAvailablePoints.Text = FormatNumber(rec!points, 2, vbTrue, vbFalse)
                    rec.MoveNext
                Loop
            End If
            con.Close
        End If
    End With
End Sub

Private Sub tb_Standard_ButtonClick(ByVal Button As MSComctlLib.Button)
    If EditAccessRights(16) = False Then
        MsgBox ErrorCodes(74), vbCritical, "Limited Rights"
        Exit Sub
    End If
    Dim Item As MSComctlLib.ListItem
    Select Case Button.Index
        Case 1 'New
            Initialize
        Case 2 'Save
            If isActive = False Then Exit Sub
            If isValidated = True Then
                On Error GoTo ErrHandler
                Set con = New ADODB.Connection
                Set cmd = New ADODB.Command
                
                'SAVE MAIN Customer DETAILS
                con.ConnectionString = ConnString
                con.Open
                con.BeginTrans
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                
                cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInputOutput, , CustomerId)
                cmd.Parameters.Append cmd.CreateParameter("@CustomerCode", adVarChar, adParamInputOutput, 50, txtCode.Text)
                cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 500, txtName.Text)
                cmd.Parameters.Append cmd.CreateParameter("@SalesRep", adVarChar, adParamInput, 250, txtSalesman.Text)
                cmd.Parameters.Append cmd.CreateParameter("@Mobile", adVarChar, adParamInput, 50, txtMobile.Text)
                cmd.Parameters.Append cmd.CreateParameter("@Phone", adVarChar, adParamInput, 50, txtPhone.Text)
                cmd.Parameters.Append cmd.CreateParameter("@Address", adVarChar, adParamInput, 500, txtAddress.Text)
                cmd.Parameters.Append cmd.CreateParameter("@CityId", adInteger, adParamInput, , cmbCity.ItemData(cmbCity.ListIndex))
                cmd.Parameters.Append cmd.CreateParameter("@CardNumber", adVarChar, adParamInput, 250, txtCardnumber.Text)
                cmd.Parameters.Append cmd.CreateParameter("@CreditLimit", adDecimal, adParamInput, , NVAL(txtCreditLine.Text))
                                      cmd.Parameters("@CreditLimit").NumericScale = 2
                                      cmd.Parameters("@CreditLimit").Precision = 18
                cmd.Parameters.Append cmd.CreateParameter("@CardPoints", adDecimal, adParamInput, , Val(Replace(txtAvailablePoints.Text, ",", "")))
                                      cmd.Parameters("@CardPoints").NumericScale = 2
                                      cmd.Parameters("@CardPoints").Precision = 18
                cmd.Parameters.Append cmd.CreateParameter("@PricingSchemeId", adInteger, adParamInput, , cmbPricingScheme.ItemData(cmbPricingScheme.ListIndex))
                If chkMember.value = Checked Then
                    cmd.Parameters.Append cmd.CreateParameter("@MemberId", adInteger, adParamInput, , 1)
                Else
                    cmd.Parameters.Append cmd.CreateParameter("@MemberId", adInteger, adParamInput, , 2)
                End If
                cmd.Parameters.Append cmd.CreateParameter("@InterestRate", adDecimal, adParamInput, , NVAL(txtInterestRate.Text))
                                      cmd.Parameters("@InterestRate").NumericScale = 2
                                      cmd.Parameters("@InterestRate").Precision = 18
                If CustomerId = 0 Then
                    cmd.CommandText = "BASE_Customer_Insert"
                    cmd.Execute
                    CustomerId = cmd.Parameters("@CustomerId")
                    
                    SavePOSAuditTrail UserId, WorkstationId, "", "Created new customer: " & txtName.Text, "SALES"
                Else
                    'cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId) 'NOT SET
                    cmd.CommandText = "BASE_Customer_Update"
                    cmd.Execute
                    
                    SavePOSAuditTrail UserId, WorkstationId, "", "Updated customer: " & txtName.Text, "SALES"
                End If
                
                con.CommitTrans
                con.Close
                
                BASE_ContainerFrm.statusBar_Main.Panels(1).Text = MessageCodes(7) & " " & MessageCodes(0)
        
                Dim isFound As Boolean
                isFound = False
                For Each Item In lvSearch.ListItems
                    If CustomerId = Item.Text Then
                        Item.SubItems(1) = txtCode.Text
                        Item.SubItems(2) = txtName.Text
                        isFound = True
                        Item.Selected = True
                        Item.EnsureVisible
                        Exit For
                    End If
                Next
                If isFound = False Then
                    Set Item = lvSearch.ListItems.add(, , CustomerId)
                        Item.SubItems(1) = txtCode.Text
                        Item.SubItems(2) = txtName.Text
                        Item.Selected = True
                        Item.EnsureVisible
                End If
            End If
        Case 4 ' Delete
            If CustomerId <> 0 Then
                Set con = New ADODB.Connection
                Set cmd = New ADODB.Command
                con.ConnectionString = ConnString
                con.Open
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "BASE_Customer_Delete"
                cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInput, , CustomerId)
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
                
                SavePOSAuditTrail UserId, WorkstationId, "", "Deactivated customer: " & txtName.Text, "SALES"
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

Private Sub Text1_Change()

End Sub

Private Sub txtAvailablePoints_Change()
    If IsNumeric(txtAvailablePoints.Text) = False And Trim(txtAvailablePoints.Text) <> "" Then
        txtAvailablePoints.Text = "0.00"
        selectText txtAvailablePoints
    End If
End Sub

Private Sub txtCardNumber_GotFocus()
    selectText txtCardnumber
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

Private Sub txtSearch_CustomerCode_Change()
    btnSearch_Click
End Sub

Private Sub txtSearch_Name_Change()
    btnSearch_Click
End Sub

