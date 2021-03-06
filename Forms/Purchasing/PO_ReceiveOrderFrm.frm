VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form PO_ReceiveOrderFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14955
   Icon            =   "PO_ReceiveOrderFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   14955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picStatus 
      BorderStyle     =   0  'None
      Height          =   1055
      Left            =   7920
      Picture         =   "PO_ReceiveOrderFrm.frx":000C
      ScaleHeight     =   1050
      ScaleWidth      =   3750
      TabIndex        =   17
      Top             =   2280
      Width           =   3755
   End
   Begin MSComctlLib.ListView lvItemList 
      Height          =   3135
      Left            =   6090
      TabIndex        =   18
      Top             =   3090
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   5530
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
   Begin VB.Frame Body_Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   4560
      TabIndex        =   30
      Top             =   0
      Width           =   10455
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
         Left            =   120
         TabIndex        =   54
         Top             =   2160
         Width           =   2895
      End
      Begin VB.Frame Frame_Header1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1935
         Left            =   240
         TabIndex        =   44
         Top             =   480
         Width           =   6495
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
            Left            =   1080
            TabIndex        =   50
            Top             =   720
            Width           =   4560
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label9 
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
            Left            =   120
            TabIndex        =   49
            Top             =   720
            Width           =   840
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Receive Order"
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
            Left            =   0
            TabIndex        =   46
            Top             =   0
            Width           =   2280
         End
      End
      Begin VB.Frame Frame_Header2 
         BackColor       =   &H00FFFFFF&
         Height          =   2115
         Left            =   6930
         TabIndex        =   40
         Top             =   360
         Width           =   3405
         Begin VB.ComboBox cmbLocation 
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
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1680
            Width           =   2055
         End
         Begin VB.TextBox txtReferenceNumber 
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
            Left            =   1200
            TabIndex        =   1
            Top             =   600
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
            Left            =   1200
            TabIndex        =   3
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
            Left            =   1200
            TabIndex        =   0
            Top             =   240
            Width           =   2055
         End
         Begin MSComCtl2.DTPicker dtOrder 
            Height          =   330
            Left            =   1200
            TabIndex        =   2
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
            Format          =   111869953
            CurrentDate     =   41509
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Location"
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
            TabIndex        =   48
            Top             =   1680
            Width           =   780
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PO/Ref #"
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
            TabIndex        =   47
            Top             =   600
            Width           =   825
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
            TabIndex        =   42
            Top             =   960
            Width           =   435
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Delivery #"
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
            TabIndex        =   41
            Top             =   240
            Width           =   930
         End
      End
      Begin VB.Frame Frame_Footer 
         BackColor       =   &H00FFFFFF&
         Height          =   1935
         Left            =   120
         TabIndex        =   34
         Top             =   6840
         Width           =   10215
         Begin VB.TextBox txtReceivedBy 
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
            TabIndex        =   35
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
            Height          =   810
            Left            =   1320
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   600
            Width           =   5295
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Received by"
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
            Width           =   1125
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
            TabIndex        =   39
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
            TabIndex        =   38
            Top             =   600
            Width           =   795
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Items"
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
            TabIndex        =   37
            Top             =   240
            Width           =   1035
         End
         Begin VB.Label lblTotalItems 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
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
            Left            =   8520
            TabIndex        =   36
            Top             =   240
            Width           =   1500
         End
      End
      Begin VB.Frame Frame_Body 
         BackColor       =   &H00FFFFFF&
         Height          =   4260
         Left            =   120
         TabIndex        =   32
         Top             =   2520
         Width           =   10215
         Begin VB.CommandButton btnAutoFill 
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
            TabIndex        =   53
            Top             =   150
            Width           =   3285
         End
         Begin VB.CommandButton btnItemSearch 
            Height          =   330
            Left            =   4320
            Picture         =   "PO_ReceiveOrderFrm.frx":7D7E
            Style           =   1  'Graphical
            TabIndex        =   6
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
            TabIndex        =   5
            Top             =   240
            Width           =   3015
         End
         Begin MSComctlLib.ListView lvItems 
            Height          =   3495
            Left            =   120
            TabIndex        =   7
            Top             =   600
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
            NumItems        =   8
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ReceiveOrderLineId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "PurchaseOrderId"
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
               SubItemIndex    =   6
               Text            =   "ProductId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "ReceiveOrderId"
               Object.Width           =   0
            EndProperty
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
            TabIndex        =   33
            Top             =   240
            Width           =   555
         End
      End
      Begin MSComctlLib.ListView lvVendor 
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
         TabIndex        =   45
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
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame LayoutFrame_Search 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   -75
      TabIndex        =   24
      Top             =   0
      Width           =   4575
      Begin VB.TextBox txtSearch_ReferenceNumber 
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
         Left            =   1440
         TabIndex        =   11
         Top             =   840
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
         TabIndex        =   10
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
         TabIndex        =   12
         Top             =   1200
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
         TabIndex        =   15
         Top             =   2400
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvSearch 
         Height          =   5895
         Left            =   200
         TabIndex        =   16
         Top             =   2880
         Width           =   4235
         _ExtentX        =   7461
         _ExtentY        =   10398
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
            Text            =   "Delivery #"
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
         TabIndex        =   14
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
         Format          =   111869953
         CurrentDate     =   41686
      End
      Begin MSComCtl2.DTPicker DateFrom 
         Height          =   345
         Left            =   1440
         TabIndex        =   13
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
         Format          =   111869953
         CurrentDate     =   41686
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PO/Ref #"
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
         Top             =   840
         Width           =   825
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery #"
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
         Left            =   240
         TabIndex        =   28
         Top             =   75
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
         TabIndex        =   27
         Top             =   1200
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
         TabIndex        =   26
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
         TabIndex        =   25
         Top             =   1920
         Width           =   705
      End
   End
   Begin VB.PictureBox picCompleted 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   -10080
      Picture         =   "PO_ReceiveOrderFrm.frx":7FA2
      ScaleHeight     =   2295
      ScaleWidth      =   6195
      TabIndex        =   23
      Top             =   2640
      Width           =   6195
   End
   Begin VB.PictureBox pic_Cancelled 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   -10080
      Picture         =   "PO_ReceiveOrderFrm.frx":10630
      ScaleHeight     =   2295
      ScaleWidth      =   6195
      TabIndex        =   22
      Top             =   2640
      Width           =   6195
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
      Height          =   375
      Left            =   8745
      TabIndex        =   21
      Top             =   7200
      Width           =   2895
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
      Left            =   -10080
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   7200
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.PictureBox picPaid 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1860
      Left            =   -10080
      Picture         =   "PO_ReceiveOrderFrm.frx":212CC
      ScaleHeight     =   1860
      ScaleWidth      =   5250
      TabIndex        =   19
      Top             =   2640
      Visible         =   0   'False
      Width           =   5250
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
            Picture         =   "PO_ReceiveOrderFrm.frx":2A25D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PO_ReceiveOrderFrm.frx":30ABF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PO_ReceiveOrderFrm.frx":37321
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PO_ReceiveOrderFrm.frx":3DB83
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PO_ReceiveOrderFrm.frx":3DDF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PO_ReceiveOrderFrm.frx":3E469
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "PO_ReceiveOrderFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public StatusId, VendorId, ReceiveOrderId, PurchaseOrderId, id As Long
Dim TotalSacks As Double
Dim OrderLine(10000) As Integer
Dim ctrOrderLine As Integer
Public global_remarks As String

Public Sub Initialize()
    Dim txtControl As Control
    For Each txtControl In Me.Controls
        If TypeOf txtControl Is TextBox And txtControl.Name <> "txtSearch_Order" And txtControl.Name <> "txtReferenceNumber" Then
            txtControl.Text = ""
            txtStatus.Text = "Open"
        End If
    Next
    
    global_remarks = ""
    picStatus.Visible = False
    isNotCompleted (True)
    dtOrder.value = Format(Now, "MM/DD/YY hh:mm:ss")
    txtStatus.Text = "Open"
    lvItems.ListItems.Clear
    lvVendor.Visible = False
    lvItemList.Visible = False
    lvVendor.Left = -9999
    lvItemList.Left = -9999
    txtSearch_ReferenceNumber.Text = txtReferenceNumber.Text
    btnPaid.Visible = False
    tb_Standard.Buttons(4).Caption = "Cancel"
    tb_Standard.Buttons(4).Image = 3
    
    lblTotalItems.Caption = "0.00"
    
    id = 1
    StatusId = 1
    VendorId = 0
    ReceiveOrderId = 0
    TotalSacks = 0
    
    ctrOrderLine = 0
    
    On Error Resume Next
    isModify = False
End Sub
Private Sub Save(ByVal StatusId As Integer, Optional isReopen As Variant)
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
        cmd.Parameters.Append cmd.CreateParameter("@ReceiveOrderId", adInteger, adParamInputOutput, , ReceiveOrderId)
        cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , PurchaseOrderId)
        cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtOrder.value)
        cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInput, 50, txtOrderNumber.Text)
        cmd.Parameters.Append cmd.CreateParameter("@ReferenceNumber", adVarChar, adParamInput, 50, txtReferenceNumber.Text)
        cmd.Parameters.Append cmd.CreateParameter("@LocationId", adInteger, adParamInput, , cmbLocation.ItemData(cmbLocation.ListIndex))
        cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , StatusId)
        cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, txtRemarks.Text)
        cmd.Parameters.Append cmd.CreateParameter("@ReceivedBy", adVarChar, adParamInput, 250, txtReceivedBy.Text)
        cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
        
        If ReceiveOrderId = 0 Then
            cmd.CommandText = "PO_ReceiveOrder_Insert"
            cmd.Execute
            ReceiveOrderId = cmd.Parameters("@ReceiveOrderId")
            
            SavePOSAuditTrail UserId, WorkstationId, "", "Created new receive order: " & txtOrderNumber.Text, "PURCHASING"
        Else
            cmd.CommandText = "PO_ReceiveOrder_Update"
            cmd.Execute
            
            Dim auditstatus As String
            If StatusId = 1 Then auditstatus = "Open"
            If StatusId = 2 Then auditstatus = "Completed"
            If StatusId = 7 Then auditstatus = "Cancelled"
            
            SavePOSAuditTrail UserId, WorkstationId, "", "Updated receive order details: " & txtOrderNumber.Text & " - Status: " & auditstatus, "PURCHASING"
        End If

        'SAVE ORDER LINE
        Dim Item As MSComctlLib.ListItem

        For Each Item In lvItems.ListItems
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc

            cmd.Parameters.Append cmd.CreateParameter("@ReceiveOrderLineId", adInteger, adParamInputOutput, , Val(Item.Text))
            cmd.Parameters.Append cmd.CreateParameter("@ReceiveOrderId", adInteger, adParamInput, , ReceiveOrderId)
            cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , Val(Item.SubItems(6)))
            cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, Item.SubItems(3))
            cmd.Parameters.Append cmd.CreateParameter("@Quantity", adDecimal, adParamInput, , Val(Replace(Item.SubItems(4), ",", "")))
                                  cmd.Parameters("@Quantity").Precision = 18
                                  cmd.Parameters("@Quantity").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@Uom", adVarChar, adParamInput, 250, Item.SubItems(5))
            cmd.Parameters.Append cmd.CreateParameter("@LocationId", adInteger, adParamInput, , cmbLocation.ItemData(cmbLocation.ListIndex))
            cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , StatusId)
            
            If Item.Text = "" Then
                cmd.CommandText = "PO_ReceiveOrderLine_Insert"
            Else
                cmd.CommandText = "PO_ReceiveOrderLine_Update"
            End If
            cmd.Execute
            Item.Text = cmd.Parameters("@ReceiveOrderLineId")
            Item.SubItems(7) = ReceiveOrderId
        Next

        'DELETE ORDERLINE IF ANY
        Dim ctr As Integer
        For ctr = 0 To ctrOrderLine
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc

            If OrderLine(ctr) <> 0 Then
                cmd.Parameters.Append cmd.CreateParameter("@ReceiveOrderLineId", adInteger, adParamInput, , OrderLine(ctr))
                cmd.CommandText = "PO_ReceiveOrderLine_Delete"
                cmd.Execute
            Else
                Exit For
            End If
        Next

        con.CommitTrans
        con.Close

        If StatusId = 2 Then
            isNotCompleted (False)
            txtStatus.Text = "Completed"
            Me.StatusId = 2
            
            With PO_PurchaseOrderFrm
                .StatusId = 3
                LoadImageStatus PO_PurchaseOrderFrm.picStatus, GetStatus(PO_PurchaseOrderFrm.StatusId)
                On Error Resume Next
                .lvSearch.SelectedItem.SubItems(2) = "Receiving"
                .txtStatus.Text = "Receiving"
            End With
        End If

        Dim isFound As Boolean
        isFound = False
        For Each Item In lvSearch.ListItems
            If ReceiveOrderId = Item.Text Then
                Item.SubItems(1) = txtOrderNumber.Text
                Item.SubItems(3) = txtStatus.Text
                isFound = True
                Item.Selected = True
                Item.EnsureVisible
                Exit For
            End If
        Next
        If isFound = False Then
            Set Item = lvSearch.ListItems.add(, , ReceiveOrderId)
                Item.SubItems(1) = txtOrderNumber.Text
                Item.SubItems(3) = txtStatus.Text
                Item.Selected = True
                Item.EnsureVisible
        End If
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
    If txtOrderNumber.Text = "" Then
        Validated = False
        GLOBAL_MessageFrm.lblErrorMessage.Caption = "Delivery number is required."
        GLOBAL_MessageFrm.Show (1)
        txtOrderNumber.SetFocus
    ElseIf cmbLocation.ListIndex = 0 Then
        Validated = False
        GLOBAL_MessageFrm.lblErrorMessage.Caption = "Receive location is required."
        GLOBAL_MessageFrm.Show (1)
        cmbLocation.SetFocus
    Else
        Validated = True
    End If
End Function

Public Sub isNotCompleted(ByVal a As Boolean)
    Frame_Header1.enabled = a
    Frame_Header2.enabled = a
    Frame_Body.enabled = a
    Frame_Footer.enabled = a
End Sub
Public Sub CountTotal()
    Dim Total As Double
    Dim Item As MSComctlLib.ListItem
    
    For Each Item In lvItems.ListItems
        Total = Total + NVAL(Item.SubItems(4))
    Next
    lblTotalItems.Caption = FormatNumber(Total, 0, vbTrue, vbFalse)
End Sub
Public Sub Populate(ByVal data As String)
    Select Case data
        Case "Location"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data(data)
            cmbLocation.Clear
            cmbLocation.AddItem ""
            cmbLocation.ItemData(cmbLocation.NewIndex) = 0
            cmbLocation.ListIndex = 0
            If Not rec.EOF Then
                Do Until rec.EOF
                    cmbLocation.AddItem rec!Location
                    cmbLocation.ItemData(cmbLocation.NewIndex) = rec!LocationId
                    rec.MoveNext
                Loop
            End If
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
        Case "ReceiveOrderGet"
            Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            Set rec = New ADODB.Recordset
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "PO_ReceiveOrder_Get"
            cmd.Parameters.Append cmd.CreateParameter("@ReceiveOrderId", adInteger, adParamInput, , ReceiveOrderId)
            cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , PurchaseOrderId)
            Set rec = cmd.Execute
            If Not rec.EOF Then
                StatusId = rec!StatusId
                txtStatus.Text = rec!Status
                txtOrderNumber.Text = rec!DeliveryNumber
                dtOrder.value = Format(rec!receiveddate, "MM/DD/YY")
                txtRemarks.Text = rec!Remarks
                txtReceivedBy.Text = rec!receivedby
                
                On Error Resume Next
                cmbLocation.Text = rec!Location
                
                If rec!StatusId = 2 Then 'COMPLETED
                    isNotCompleted (False)
                    tb_Standard.Buttons(4).Caption = "Cancel"
                    tb_Standard.Buttons(4).Image = 3
                ElseIf rec!StatusId = 3 Then 'IN PROGRESS
                    isNotCompleted (False)
                    tb_Standard.Buttons(4).Caption = "Cancel"
                    tb_Standard.Buttons(4).Image = 3
                ElseIf rec!StatusId = 7 Then 'Cancelled
                    isNotCompleted (False)
                    tb_Standard.Buttons(4).Caption = "Activate"
                    tb_Standard.Buttons(4).Image = 6
                ElseIf rec!StatusId = 6 Then 'PAID
                    isNotCompleted (False)
                    tb_Standard.Buttons(4).Caption = "Cancel"
                    tb_Standard.Buttons(4).Image = 3
                Else
                    isNotCompleted (True)
                    tb_Standard.Buttons(4).Caption = "Cancel"
                    tb_Standard.Buttons(4).Image = 3
                End If
            End If
            con.Close
        Case "ReceiveOrderLoad"
            Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            Set rec = New ADODB.Recordset
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "PO_ReceiveOrder_Get"
            cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , PurchaseOrderId)
            Set rec = cmd.Execute
            Dim Item As MSComctlLib.ListItem
            lvSearch.ListItems.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isCashAdvance = "False" Then
                        Set Item = lvSearch.ListItems.add(, , rec!ReceiveOrderId)
                            Item.SubItems(1) = rec!OrderNumber
                            Item.SubItems(2) = rec!ReferenceNumber
                            Item.SubItems(3) = rec!Status
                    End If
                    rec.MoveNext
                Loop
            End If
            con.Close
        Case "ReceiveOrderLineGet"
            Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            Set rec = New ADODB.Recordset
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "PO_ReceiveOrderLine_Get"
            cmd.Parameters.Append cmd.CreateParameter("@ReceiveOrderId", adInteger, adParamInput, , ReceiveOrderId)
            Set rec = cmd.Execute
            lvItems.ListItems.Clear
            'On Error Resume Next
            If Not rec.EOF Then
                Do Until rec.EOF
                    Set Item = lvItems.ListItems.add(, , rec!ReceiveOrderLineId)
                        Item.SubItems(1) = PurchaseOrderId
                        Item.SubItems(2) = rec!itemcode
                        Item.SubItems(3) = rec!Name
                        Item.SubItems(4) = FormatNumber(rec!quantity, 2, vbTrue)
                        Item.SubItems(5) = rec!Uom
                        Item.SubItems(6) = rec!ProductId
                        Item.SubItems(7) = rec!ReceiveOrderId
                    rec.MoveNext
                Loop
            End If
            con.Close
    End Select
End Sub


Private Sub btnAutoFill_Click()
    'AutoFill
    Dim Item, dritem As MSComctlLib.ListItem
    Dim rItem As MSComctlLib.ListItem
    Dim x As Integer
    Dim isFound As Boolean
    
    For Each Item In PO_PurchaseOrderFrm.lvItems.ListItems
        isFound = False
        With PO_ReceiveOrderFrm
        
'            For Each dritem In .lvItems.ListItems
'                If dritem.SubItems(6) = item.SubItems(9) Then
'                    isFound = True
'                    Exit For
'                End If
'            Next
            
            If isFound = False Then
                Set rItem = lvItems.ListItems.add(, , "")
                    rItem.SubItems(2) = Item.SubItems(2)
                    rItem.SubItems(3) = Item.SubItems(3)
                    rItem.SubItems(4) = Item.SubItems(4)
                    rItem.SubItems(5) = Item.SubItems(5)
                    rItem.SubItems(6) = Item.SubItems(9)
            End If
        End With
    Next
    
    txtItemSearch.SetFocus
    lvItemList.Visible = False
    CountTotal
End Sub

Private Sub btnComplete_Click()
    If StatusId = 2 Then
        MsgBox "Save failed. No changes made. Order is already complete.", vbCritical
        Exit Sub
    End If
    Dim x As Variant
    x = MsgBox("This will complete the transaction. Product inventories will now be updated. Proceed?", vbExclamation + vbOKCancel)
    If x = vbOK Then
        If Validated = True Then
            Save (1)
            Save (2)
            LoadImageStatus picStatus, GetStatus(StatusId)
        End If
    End If
End Sub

Private Sub btnInvoice_Click()
    If EditAccessRights(34) = False Then
        MsgBox ErrorCodes(74), vbCritical, "Limited Rights"
        Exit Sub
    End If
    
    If ReceiveOrderId = 0 Then Exit Sub
    PO_PurchaseInvoiceFrm.Show '(1)
End Sub

Private Sub btnItemSearch_Click()
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    Dim Item As MSComctlLib.ListItem
    
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
                    Set Item = lvItemList.ListItems.add(, , rec!ProductId)
                        Item.SubItems(1) = rec!itemcode
                        Item.SubItems(2) = rec!Name
                        Item.SubItems(3) = FormatNumber(rec!unitcost, 2, vbTrue, vbFalse)
                        Item.SubItems(4) = rec!Uom
                    lvItemList.Visible = True
                    lvItemList.Left = 6070
                    'lvItemList.Top = 3600
                    LastProductId = rec!ProductId
                    rec.MoveNext
                Else
                    rec.MoveNext
                End If
            Else
                rec.MoveNext
            End If
            'rec.MoveNext
        Loop
    Else
        lvItemList.Visible = False
        lvItemList.Left = -9999
    End If
    'DistinctList lvItemList
    con.Close
End Sub

Private Sub btnReceiveOrder_Click()
    If ReceiveOrderId = 0 Then Exit Sub
    
    If EditAccessRights(33) = False Then
        MsgBox ErrorCodes(74), vbCritical, "Limited Rights"
        Exit Sub
    End If
    
    If (StatusId = 2) Or (StatusId = 7) Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(60)
        GLOBAL_MessageFrm.Show (1)
        Exit Sub
    End If
    
    
'    Dim totalReceived, totalOrdered As Double
'    Dim item As MSComctlLib.ListItem
'
'    For Each item In lvItems.ListItems
'        totalOrdered = totalOrdered + Val(Replace(item.SubItems(4), ",", ""))
'        totalReceived = totalReceived + Val(Replace(item.SubItems(11), ",", ""))
'    Next
    
    'Validate if All orders already fullfilled
'    If totalOrdered <= totalReceived Then
'        Dim X As Variant
'        X = MsgBox("All orders have already been received. Would you like to complete this order?", vbYesNo + vbQuestion)
'        If X = vbYes Then
'            'UPDATE STATUS
'            Set con = New ADODB.Connection
'            con.ConnectionString = ConnString
'            con.Open
'            Set cmd = New ADODB.Command
'                cmd.ActiveConnection = con
'                cmd.CommandType = adCmdStoredProc
'                cmd.CommandText = "PO_ReceiveOrderStatus_Update"
'                cmd.Parameters.Append cmd.CreateParameter("@ReceiveOrderId", adInteger, adParamInput, , PO_ReceiveOrderFrm.ReceiveOrderId)
'                cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , 2) 'IN PROGRESS
'                cmd.Execute
'            con.Close
'            MsgBox "Order completed.", vbInformation
'
'            Set item = lvSearch.SelectedItem
'            lvSearch_ItemClick item
'        Else
'            PO_ReceiveOrderFrm.txtOrderNumber.text = txtOrderNumber.text
'            'PO_ReceiveOrderFrm.ReceiveOrderIdx = Me.ReceiveOrderId
'            PO_ReceiveOrderFrm.Show
'        End If
'    Else
        'PO_ReceiveOrderFrm.txtOrderNumber.Text = txtOrderNumber.Text
        'PO_ReceiveOrderFrm.ReceiveOrderIdx = Me.ReceiveOrderId
        PO_ReceiveOrderFrm.Show
'    End If
    
    
    
End Sub

Public Sub btnSearch_Click()
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "PO_ReceiveOrder_Search"
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
            Set Item = lvSearch.ListItems.add(, , rec!ReceiveOrderId)
                Item.SubItems(1) = rec!OrderNumber
                Item.SubItems(2) = rec!ReferenceNumber
                Item.SubItems(3) = rec!Status
            rec.MoveNext
        Loop
    End If
    con.Close
End Sub

Private Sub btnStatus_Click()
    If StatusId = 1 Or StatusId = 3 Then
        Dim x As Variant
        x = MsgBox("Are you sure you want to complete this order?", vbQuestion + vbYesNo)
        If x = vbYes Then
            Save (2)
            If Me.StatusId = 2 Then
                'btnStatus.Caption = "Reopen Order"
                ''picCompleted.Left = 6840
                ''picCompleted.Visible = True
                '''pic_Cancelled.Visible = False
                '''pic_Cancelled.Left = -9999
                ''picPaid.Visible = False
                ''picPaid.Left = -9999
                'btnPaid.Visible = True
            End If
        End If
    ElseIf StatusId = 2 Or StatusId = 4 Then
        'Dim x As Variant
        x = MsgBox("Are you sure you want to reopen this order? This will revert all connected " & _
                    "transactions for this order.", vbCritical + vbYesNo, "WARNING")
        If x = vbYes Then
            Save 1, True
            'btnStatus.Caption = "Complete Order"
            txtStatus.Text = "Open"
            btnPaid.Visible = False
            isNotCompleted (True)
            ''picCompleted.Visible = False
            ''picCompleted.Left = -9999
            ''picPaid.Left = -9999
            ''picPaid.Visible = False
            Me.StatusId = 1
        Else
        End If
    End If
End Sub

Private Sub cmbVendor_Change()
'    If Trim(cmbVendor.Text) <> "" Then
'        Dim Item As MSComctlLib.ListItem
'        Set con = New ADODB.Connection
'        Set rec = New ADODB.Recordset
'        Set cmd = New ADODB.Command
'        'Dim item As MSComctlLib.ListItem
'
'        con.ConnectionString = ConnString
'        con.Open
'        cmd.ActiveConnection = con
'        cmd.CommandType = adCmdStoredProc
'        cmd.CommandText = "BASE_Vendor_Search"
'        cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, Trim(cmbVendor.Text))
'        Set rec = cmd.Execute
'        If Not rec.EOF Then
'            lvVendor.ListItems.Clear
'            lvVendor.Left = 1440
'            lvVendor.Visible = True
'            Do Until rec.EOF
'                If rec!isActive = "True" Then
'                    Set Item = lvVendor.ListItems.add(, , rec!VendorId)
'                        Item.SubItems(1) = rec!VendorCode
'                        Item.SubItems(2) = rec!Name
'                        Item.SubItems(3) = FormatNumber(rec!OutStandingBalance, 2, vbTrue, vbFalse)
'                        Item.SubItems(4) = rec!Phone
'                        Item.SubItems(5) = rec!Address
'                End If
'                rec.MoveNext
'            Loop
'        Else
'            lvVendor.Visible = False
'            lvVendor.Left = -9999
'        End If
'        con.Close
'
''        If Trim(cmbVendor.text) = "" Then
''            cmbVendor.BackColor = &HC0C0FF
''        Else
''            cmbVendor.BackColor = vbWhite
''        End If
'    End If
'End Sub
'
'Private Sub cmbVendor_GotFocus()
'    selectText cmbVendor
'End Sub
'
'Private Sub cmbVendor_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'        Case vbKeyReturn
'            Set con = New ADODB.Connection
'            Set rec = New ADODB.Recordset
'            Set cmd = New ADODB.Command
'            Dim Item As MSComctlLib.ListItem
'
'            con.ConnectionString = ConnString
'            con.Open
'            cmd.ActiveConnection = con
'            cmd.CommandType = adCmdStoredProc
'            cmd.CommandText = "BASE_Vendor_Search"
'            cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 500, "")
'            cmd.Parameters.Append cmd.CreateParameter("@VendorCode", adVarChar, adParamInput, 50, cmbVendor.Text)
'            Set rec = cmd.Execute
'            If Not rec.EOF Then
'                lvVendor.ListItems.Clear
'                lvVendor.Left = 1440
'                lvVendor.Visible = True
'                Do Until rec.EOF
'                    Set Item = lvVendor.ListItems.add(, , rec!VendorId)
'                        Item.SubItems(1) = rec!VendorCode
'                        Item.SubItems(2) = rec!Name
'                        Item.SubItems(3) = FormatNumber(rec!OutStandingBalance, 2, vbTrue, vbFalse)
'                        Item.SubItems(4) = rec!Phone
'                        Item.SubItems(5) = rec!Address
'                    rec.MoveNext
'                Loop
'            Else
'                lvVendor.Visible = False
'                lvVendor.Left = -9999
'            End If
'            con.Close
'        Case vbKeyUp, vbKeyDown
'            If lvVendor.Visible = True Then
'                lvVendor.SetFocus
'            End If
'    End Select
End Sub

Private Sub cmbTerms_Click()
'    If cmbTerms.ListIndex > 1 Then
'        txtDays.text = cmbTerms.Tag
'    End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            If lvVendor.Visible = True Then
                lvVendor.Visible = False
                lvVendor.Left = -9999
                'cmbVendor.SetFocus
            ElseIf lvItemList.Visible = True Then
                lvItemList.Visible = False
                lvItemList.Left = -9999
                txtItemSearch.SetFocus
                'txtCode.SetFocus
            End If
        Case vbKeyF1
            Unload INV_ProductSearch
            INV_ProductSearch.isPO = True
            INV_ProductSearch.isWithdraw = False
            INV_ProductSearch.isSO = False
            INV_ProductSearch.isAS = False
            INV_ProductSearch.Show (1)
        Case vbKeyF4
            txtItemSearch.SetFocus
            'txtCode.SetFocus
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
    
    lvVendor.ColumnHeaders(2).width = lvVendor.width * 0.25
    lvVendor.ColumnHeaders(3).width = lvVendor.width * 0.42
    lvVendor.ColumnHeaders(4).width = lvVendor.width * 0.28
    
    lvItems.ColumnHeaders(3).width = lvItems.width * 0.14
    lvItems.ColumnHeaders(4).width = lvItems.width * 0.62
    lvItems.ColumnHeaders(5).width = lvItems.width * 0.12
    lvItems.ColumnHeaders(6).width = lvItems.width * 0.09
    lvItems.ColumnHeaders(7).width = lvItems.width * 0
    
    'StatusBarWidth Me, statusBar_Main
    '****************************************
    
    Initialize
    Populate "Terms"
    Populate "Status"
    Populate "Location"
    DateFrom.value = Format(Now - 30, "MM/DD/YY")
    DateTo.value = Format(Now, "MM/DD/YY")
End Sub





Private Sub lblGrossAmount_Click()

End Sub

Private Sub lblGrossKilos_Click()

End Sub

Private Sub lblCaption_AR_Click()

End Sub

Private Sub lvVendor_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'lvVendor_DblClick
    End Select
End Sub

Private Sub lvItemList_DblClick()
    'PO_ReceiveItemOptFrm.txtCost.Text = lvItemList.SelectedItem.SubItems(3)
    PO_ReceiveItemOptFrm.txtDescription.Text = lvItemList.SelectedItem.SubItems(2)
    isModify = False
    PO_ReceiveItemOptFrm.Show (1)
End Sub

Private Sub lvItemList_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If lvItemList.ListItems.count > 0 Then
                isModify = False
                PO_ReceiveItemOptFrm.txtDescription.Text = lvItemList.SelectedItem.SubItems(2)
                PO_ReceiveItemOptFrm.Show (1)
            End If
    End Select
End Sub

Private Sub lvItems_DblClick()
    If lvItems.ListItems.count > 0 Then
        isModify = True
        With PO_ReceiveOrderDialogFrm
            .txtQuantity.Text = lvItems.SelectedItem.SubItems(4)
            .txtDescription.Text = lvItems.SelectedItem.SubItems(3)
            On Error Resume Next
            .cmbUnit.Text = lvItems.SelectedItem.SubItems(5)
            .Show (1)
        End With
    End If
End Sub

Private Sub lvItems_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp
            If lvItems.ListItems.count > 0 Then
                If lvItems.SelectedItem.Index = 1 Then
                    txtItemSearch.SetFocus
                    'txtCode.SetFocus
                End If
            End If
        Case vbKeyDelete
            If lvItems.ListItems.count > 0 Then
                If lvItems.SelectedItem.Text <> "" Then
                    OrderLine(ctrOrderLine) = Val(lvItems.SelectedItem.Text)
                    ctrOrderLine = ctrOrderLine + 1
                    
                    SavePOSAuditTrail UserId, WorkstationId, "", "Removed item: " & lvItems.SelectedItem.SubItems(3) & " from purchase order: " & txtOrderNumber.Text
                    
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

Public Sub lvSearch_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lvSearch.ListItems.count > 0 Then
        Initialize
        ReceiveOrderId = lvSearch.SelectedItem.Text
        Populate "ReceiveOrderLineGet"
        CountTotal
        Populate "ReceiveOrderGet"
       ' Populate "Vendor"
        
        LoadImageStatus picStatus, GetStatus(StatusId)
    End If
End Sub



Private Sub tb_Standard_ButtonClick(ByVal Button As MSComctlLib.Button)
    If EditAccessRights(9) = False Then
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
            ElseIf StatusId = 2 Then 'Receiving
                MsgBox "Save failed. Order is already complete. No changes made.", vbCritical
                Exit Sub
            End If
        Case 4 'Cancel
            If StatusId = 7 Then
                MsgBox "Order already cancelled.", vbCritical
                Exit Sub
            ElseIf StatusId = 2 Then
                MsgBox "Cancel failed. Order is already complete. No changes made.", vbCritical
                Exit Sub
            ElseIf StatusId = 4 Then
                MsgBox "Cannot cancel an invoiced order.", vbCritical
                Exit Sub
            End If
            If ReceiveOrderId <> 0 Then
                Dim x As Variant
                x = MsgBox("Are you sure you want to cancel this order?", vbQuestion + vbYesNo)
                If x = vbNo Then Exit Sub
                
                If StatusId = 1 Then 'Status Open > Cancelled
                    StatusId = 7
                    txtStatus.Text = "Cancelled"
                    Save (7)
                    isNotCompleted (False)
                    '''pic_Cancelled.Left = 6360
                    '''pic_Cancelled.Visible = True
                    tb_Standard.Buttons(4).Caption = "Activate"
                    tb_Standard.Buttons(4).Image = 6
                End If
                LoadImageStatus picStatus, GetStatus(StatusId)
            End If
        Case 6 'PRINT PREVIEW
            If ReceiveOrderId <> 0 Then
                Screen.MousePointer = vbHourglass
                BASE_PrintPreviewFrm.Show
                Dim crxApp As New CRAXDRT.Application
                Dim crxRpt As New CRAXDRT.Report
                Set crxRpt = crxApp.OpenReport(App.path & "\Reports\PO_ReceiveOrder.rpt")
                crxRpt.RecordSelectionFormula = "{PO_ReceiveOrder.ReceiveOrderId}= " & ReceiveOrderId & ""
                crxRpt.DiscardSavedData

                Call ResetRptDB(crxRpt)

                BASE_PrintPreviewFrm.CRViewer.ReportSource = crxRpt
                BASE_PrintPreviewFrm.CRViewer.ViewReport
                BASE_PrintPreviewFrm.CRViewer.Zoom 1
                Screen.MousePointer = vbDefault
                
                SavePOSAuditTrail UserId, WorkstationId, "", "Generated print preview on purchase order: " & txtOrderNumber.Text, "PURCHASING"
            End If
    End Select
End Sub

Private Sub txtAdjustment_Change()
'    If IsNumeric(txtAdjustment.Text) = False Then
'        txtAdjustment.Text = "0.00"
'    Else
'        CountTotal
'    End If
End Sub

Private Sub txtAdjustment_GotFocus()
    'selectText txtAdjustment
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
'    If Trim(txtCode.Text) = "" Then
'        lvItemList.Visible = False
'        Exit Sub
'    End If
'    Set con = New ADODB.Connection
'    Set rec = New ADODB.Recordset
'    Set cmd = New ADODB.Command
'    Dim Item As MSComctlLib.ListItem
'
'    con.ConnectionString = ConnString
'    con.Open
'    cmd.ActiveConnection = con
'    cmd.CommandType = adCmdStoredProc
'    cmd.CommandText = "BASE_Product_Search1"
'    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 500, Null)
'    cmd.Parameters.Append cmd.CreateParameter("@ItemCode", adVarChar, adParamInput, 50, txtCode.Text)
'    Dim LastProductId As Long
'    Set rec = cmd.Execute
'    If Not rec.EOF Then
'        lvItemList.ListItems.Clear
'        Do Until rec.EOF
'            If rec!isActive = "True" Then
'                If LastProductId <> rec!ProductId Then
'                    Set Item = lvItemList.ListItems.add(, , rec!ProductId)
'                        Item.SubItems(1) = rec!itemcode
'                        Item.SubItems(2) = rec!Name
'                        Item.SubItems(3) = FormatNumber(rec!unitcost, 2, vbTrue, vbFalse)
'                        Item.SubItems(4) = rec!Uom
'                    lvItemList.Visible = True
'                    lvItemList.Left = 6070
'                    lvItemList.Top = 3240
'                    LastProductId = rec!ProductId
'                    rec.MoveNext
'                Else
'                    rec.MoveNext
'                End If
'            Else
'                rec.MoveNext
'            End If
'        Loop
'    Else
'        lvItemList.Visible = False
'        lvItemList.Left = -9999
'    End If
'    'DistinctList lvItemList
'    con.Close
End Sub

Private Sub txtCode_GotFocus()
    'selectText txtCode
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

Private Sub txtFreight_Change()
   
End Sub

Private Sub txtFreight_GotFocus()

End Sub

Private Sub txtItemSearch_Change()
    If Trim(txtItemSearch.Text) = "" Then Exit Sub
    btnItemSearch_Click
End Sub

Private Sub txtItemSearch_GotFocus()
    selectText txtItemSearch
End Sub

Private Sub txtItemSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(txtItemSearch.Text) = "" Then Exit Sub
            'Barcode
            Dim Item As MSComctlLib.ListItem
            Set rec = New ADODB.Recordset
            Set rec = ProductBarcode(txtItemSearch.Text)
            
            Dim isExisting As Boolean
            isExisting = False
            
            For Each Item In lvItems.ListItems
                If Not rec.EOF Then
                    If Item.SubItems(9) = rec!ProductId And Item.SubItems(5) = rec!Uom Then
                        isExisting = True
                        Exit For
                    End If
                End If
            Next
            
            If Not rec.EOF Then 'Item found display in Lvitems
                If isExisting = False Then
                    Set Item = lvItems.ListItems.add(, , "")
                    Item.SubItems(1) = ""
                    Item.SubItems(2) = rec!itemcode 'ItemCode
                    Item.SubItems(3) = rec!Name 'Name
                    Item.SubItems(4) = "1.00"
                    Item.SubItems(5) = rec!Uom
                    Item.SubItems(6) = FormatNumber(rec!unitcost, 2, vbTrue, vbFalse)
                    Item.SubItems(8) = ""
                    Item.SubItems(9) = rec!ProductId
                    'item.SubItems(13) = 1
                Else
                    Item.SubItems(4) = FormatNumber(Val(Replace(Item.SubItems(4), ",", "")) + 1, 2, vbTrue, vbFalse)
                End If
                
                CountTotal
                selectText txtItemSearch
            Else
                MsgBox "Item not found.", vbCritical, "Not Found"
                selectText txtItemSearch
            End If
        Case vbKeyUp, vbKeyDown
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

Private Sub txtSearch_OrderNumber_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then btnSearch_Click
End Sub


Private Sub txtSearch_Supplier_Change()
     btnSearch_Click
End Sub


