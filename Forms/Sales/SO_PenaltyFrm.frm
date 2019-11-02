VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form SO_SalesAdjustmentFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Adjustment"
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
   Begin VB.PictureBox picStatus 
      BorderStyle     =   0  'None
      Height          =   1055
      Left            =   7320
      Picture         =   "SO_PenaltyFrm.frx":0000
      ScaleHeight     =   1050
      ScaleWidth      =   3750
      TabIndex        =   0
      Top             =   1800
      Width           =   3755
   End
   Begin VB.Frame Body_Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   3600
      TabIndex        =   15
      Top             =   0
      Width           =   11535
      Begin VB.Frame Frame_Header2 
         BackColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   7770
         TabIndex        =   31
         Top             =   360
         Width           =   3645
         Begin VB.ComboBox cmbType 
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
            ItemData        =   "SO_PenaltyFrm.frx":7D72
            Left            =   1440
            List            =   "SO_PenaltyFrm.frx":7D7C
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   960
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
            Left            =   1440
            TabIndex        =   33
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
            Left            =   1440
            TabIndex        =   32
            Top             =   240
            Width           =   2055
         End
         Begin MSComCtl2.DTPicker dtOrder 
            Height          =   330
            Left            =   1440
            TabIndex        =   34
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
            Format          =   123928577
            CurrentDate     =   41509
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Type"
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
            Top             =   960
            Width           =   450
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
            TabIndex        =   38
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
            TabIndex        =   37
            Top             =   600
            Width           =   435
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Reference #"
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
            TabIndex        =   36
            Top             =   240
            Width           =   1110
         End
      End
      Begin VB.Frame Frame_Footer 
         BackColor       =   &H00FFFFFF&
         Height          =   1335
         Left            =   120
         TabIndex        =   26
         Top             =   7560
         Width           =   11295
         Begin VB.TextBox txtAdjustmentAmount 
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
            Left            =   9120
            TabIndex        =   44
            Top             =   600
            Width           =   2055
         End
         Begin VB.TextBox txtAdjustmentPercent 
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
            Left            =   9120
            TabIndex        =   41
            Top             =   240
            Width           =   2055
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
            TabIndex        =   28
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
            Height          =   930
            Left            =   1200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   27
            Top             =   240
            Width           =   4095
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Adjustment in Amt"
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
            Left            =   7320
            TabIndex        =   43
            Top             =   600
            Width           =   1755
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Adjustment in %"
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
            Left            =   7560
            TabIndex        =   42
            Top             =   240
            Width           =   1530
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
            TabIndex        =   30
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
            TabIndex        =   29
            Top             =   240
            Width           =   795
         End
      End
      Begin VB.Frame Frame_Body 
         BackColor       =   &H00FFFFFF&
         Height          =   5340
         Left            =   120
         TabIndex        =   23
         Top             =   2160
         Width           =   11295
         Begin VB.CommandButton btnAdd 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Select orders to adjust"
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
            TabIndex        =   24
            Top             =   240
            Width           =   3015
         End
         Begin MSComctlLib.ListView lvItems 
            Height          =   4455
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Width           =   11055
            _ExtentX        =   19500
            _ExtentY        =   7858
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
            NumItems        =   12
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "PenaltyLineId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "PenaltyId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "SalesOrderId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "CustomerId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Customer"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Order #"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Due Date"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   7
               Text            =   "Days Due"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   8
               Text            =   "Total"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   9
               Text            =   "Balance"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   10
               Text            =   "Adjustment"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   11
               Text            =   "Total Balance"
               Object.Width           =   2540
            EndProperty
         End
      End
      Begin VB.Frame Frame_Header1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   735
         Left            =   120
         TabIndex        =   17
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
            Left            =   -9999
            TabIndex        =   19
            Top             =   360
            Visible         =   0   'False
            Width           =   4695
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
            Left            =   -9999
            TabIndex        =   18
            Top             =   0
            Visible         =   0   'False
            Width           =   4695
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Adjustment (Debit/Credit Memo)"
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
            Width           =   6345
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
            TabIndex        =   21
            Top             =   360
            Visible         =   0   'False
            Width           =   885
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
            TabIndex        =   20
            Top             =   0
            Visible         =   0   'False
            Width           =   900
         End
      End
      Begin VB.CommandButton btnComplete 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Complete Adjustment"
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
         TabIndex        =   16
         Top             =   1800
         Width           =   2895
      End
      Begin MSComctlLib.Toolbar tb_Standard 
         Height          =   330
         Left            =   0
         TabIndex        =   39
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
         TabIndex        =   40
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
   Begin VB.Frame LayoutFrame_Search 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3615
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
         TabIndex        =   6
         Top             =   480
         Width           =   2055
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
         TabIndex        =   5
         Top             =   840
         Width           =   2055
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
         Left            =   2280
         TabIndex        =   4
         Top             =   2040
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvSearch 
         Height          =   6375
         Left            =   120
         TabIndex        =   7
         Top             =   2520
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   11245
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
         TabIndex        =   8
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
         Format          =   123928577
         CurrentDate     =   41686
      End
      Begin MSComCtl2.DTPicker DateFrom 
         Height          =   345
         Left            =   1440
         TabIndex        =   9
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
         Format          =   123928577
         CurrentDate     =   41686
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reference #"
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
         TabIndex        =   14
         Top             =   480
         Width           =   1110
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
         TabIndex        =   13
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
         TabIndex        =   12
         Top             =   840
         Width           =   570
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
         TabIndex        =   11
         Top             =   1560
         Width           =   705
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
         TabIndex        =   10
         Top             =   1200
         Width           =   960
      End
   End
   Begin VB.PictureBox pic_Cancelled 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   -9960
      Picture         =   "SO_PenaltyFrm.frx":7D8F
      ScaleHeight     =   2295
      ScaleWidth      =   6195
      TabIndex        =   2
      Top             =   2520
      Width           =   6195
   End
   Begin VB.PictureBox picCompleted 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   45
      Picture         =   "SO_PenaltyFrm.frx":18A2B
      ScaleHeight     =   2295
      ScaleWidth      =   6195
      TabIndex        =   1
      Top             =   2520
      Width           =   6195
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   14445
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
            Picture         =   "SO_PenaltyFrm.frx":210B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_PenaltyFrm.frx":2791B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_PenaltyFrm.frx":2E17D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_PenaltyFrm.frx":349DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_PenaltyFrm.frx":34C54
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_PenaltyFrm.frx":352C5
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "SO_SalesAdjustmentFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AdjustmentId As Long
Dim StatusId As Integer
Dim Item As MSComctlLib.ListItem
Private Sub Initialize()
    picStatus.Visible = False
    dtOrder.value = Format(Now, "MM/DD/YY hh:mm:ss")

    StatusId = 0
    lvItems.ListItems.Clear
    txtStatus.Text = ""
    txtRemarks.Text = ""
    txtOrderNumber.Text = ""
    txtAdjustmentAmount.Text = ""
    txtAdjustmentPercent.Text = ""
    AdjustmentId = 0

    On Error Resume Next
    txtItemSearch.SetFocus
    cmbType.ListIndex = 0
End Sub
Public Sub CountTotal()
    Dim totalPercent As Double
    For Each Item In lvItems.ListItems
        totalPercent = (NVAL(txtAdjustmentPercent.Text) / 100) * Item.SubItems(8)
        If cmbType = "DEBIT" Then
            Item.SubItems(10) = FormatNumber(totalPercent + NVAL(txtAdjustmentAmount.Text), 2, vbTrue, vbFalse)
            Item.SubItems(11) = FormatNumber(NVAL(Item.SubItems(9)) + NVAL(Item.SubItems(9)), 2, vbTrue, vbFalse)
        Else
            Item.SubItems(10) = FormatNumber(NVAL(txtAdjustmentAmount.Text) + totalPercent, 2, vbTrue, vbFalse)
            Item.SubItems(11) = FormatNumber(NVAL(Item.SubItems(9)) - NVAL(Item.SubItems(10)), 2, vbTrue, vbFalse)
        End If
    Next
End Sub
Public Sub Populate(ByVal data As String)
    Select Case data
        Case "AdjustmentLoad"
            Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            Set rec = New ADODB.Recordset

            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "SO_Adjustment_Get"
            cmd.Parameters.Append cmd.CreateParameter("@AdjustmentId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInput, 250, Null)
            cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , Null)
            Set rec = cmd.Execute
            lvSearch.ListItems.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    Set Item = lvSearch.ListItems.add(, , rec!AdjustmentId)
                        Item.SubItems(1) = rec!OrderNumber
                        Item.SubItems(2) = rec!Status
                    rec.MoveNext
                Loop
            End If
            con.Close
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

Private Sub btnAdd_Click()
    SO_SelectOrdersFrm.Show (1)
    DisplayProducts
End Sub

Public Sub DisplayProducts()
    On Error Resume Next

    Dim exists As Boolean
    Dim Item As MSComctlLib.ListItem

    If OrderSet.RecordCount <= 0 Then Exit Sub

    'Dim item As MSComctlLib.ListItem
    If Not OrderSet.EOF Then
        OrderSet.MoveFirst
        Do Until OrderSet.EOF
            For Each Item In lvItems.ListItems
                If NVAL(Item.SubItems(2)) = NVAL(OrderSet!SalesOrderId) Then
                    exists = True
                    Exit For
                End If
            Next

            If exists = False Then
                Set Item = lvItems.ListItems.add(, , "")
                Item.SubItems(2) = OrderSet!SalesOrderId
                Item.SubItems(3) = OrderSet!CustomerId
                Item.SubItems(4) = OrderSet!Customer
                Item.SubItems(5) = OrderSet!OrderNumber
                Item.SubItems(6) = OrderSet!DueDate
                Item.SubItems(7) = OrderSet!DaysDue
                Item.SubItems(8) = OrderSet!Total
                Item.SubItems(9) = OrderSet!balance
            End If
            OrderSet.MoveNext
        Loop
    End If
    'cmbAuditLocation.Enabled = False
End Sub

Private Sub cmbType_Click()
    CountTotal
End Sub

Private Sub Form_Load()
    DateFrom.value = Format(Now, "mm/dd/yy")
    DateTo.value = Format(Now, "mm/dd/yy")

    Populate "Status"
    'Populate "AdjustmentLoad"
    Initialize
End Sub

Private Sub lvSearch_ItemClick(ByVal Item As MSComctlLib.ListItem)
    SelectOrders
    LoadImageStatus picStatus, GetStatus(StatusId)
End Sub

Private Sub SelectOrders()
    Initialize

    Dim con As New ADODB.Connection
    con.ConnectionString = ConnString
    con.Open

    'Get Adjustment
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset

    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "SO_Adjustment_Get"
    cmd.ActiveConnection = con
    cmd.Parameters.Append cmd.CreateParameter("@AdjustmentId", adInteger, adParamInput, , lvSearch.SelectedItem.Text)
    
    Set rec = cmd.Execute
    If Not rec.EOF Then
        txtOrderNumber.Text = rec!OrderNumber
        dtOrder.value = rec!Date
        txtRemarks.Text = rec!Remarks
        StatusId = rec!StatusId
        txtStatus.Text = rec!Status
        AdjustmentId = rec!AdjustmentId
        txtAdjustmentPercent.Text = FormatNumber(rec!Percent, 2, vbTrue, vbFalse)
        txtAdjustmentAmount.Text = FormatNumber(rec!Amount, 2, vbTrue, vbFalse)
        cmbType.Text = rec!Type
    End If

    'Get Line
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandText = "SO_AdjustmentLine_Get"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("@AdjustmentId", adInteger, adParamInput, , lvSearch.SelectedItem.Text)
    Set rec = cmd.Execute

    lvItems.ListItems.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            Set Item = lvItems.ListItems.add(, , rec!AdjustmentLineId)
                Item.SubItems(1) = rec!AdjustmentId
                Item.SubItems(2) = rec!SalesOrderId
                Item.SubItems(3) = rec!CustomerId
                Item.SubItems(4) = rec!Customer
                Item.SubItems(5) = rec!OrderNumber
                Item.SubItems(6) = Format(rec!DueDate, "mm/dd/yy")
                Item.SubItems(7) = rec!DaysDue
                Item.SubItems(8) = FormatNumber(rec!Total, 2, vbTrue, vbFalse)
                Item.SubItems(9) = FormatNumber(rec!balance, 2, vbTrue, vbFalse)
                Item.SubItems(10) = FormatNumber(rec!lineamount, 2, vbTrue, vbFalse)
            rec.MoveNext
        Loop
    End If

    If StatusId = 2 Then
        txtStatus.Text = "Completed"
    End If

    If StatusId = 7 Then
        txtStatus.Text = "Cancelled"
    End If

    CountTotal
    con.Close
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
            If AdjustmentId <> 0 Then
                Dim x As Variant
                x = MsgBox("This will cancel the transaction. Proceed?", vbCritical + vbOKCancel)
                If x = vbOK Then
                    Save (7)
                    LoadImageStatus picStatus, GetStatus(StatusId)
                End If
            End If
        Case 6 'PRINT PREVIEW
            If AdjustmentId <> 0 Then
                Screen.MousePointer = vbHourglass
                BASE_PrintPreviewFrm.Show
                Dim crxApp As New CRAXDRT.Application
                Dim crxRpt As New CRAXDRT.Report
                Set crxRpt = crxApp.OpenReport(App.path & "\Reports\SO_Adjustment.rpt")
                Call ResetRptDB(crxRpt)
                crxRpt.DiscardSavedData
                crxRpt.RecordSelectionFormula = "{SO_Adjustment.AdjustmentId}= " & AdjustmentId & ""

                'crxRpt.PrintOut False

                BASE_PrintPreviewFrm.CRViewer.ReportSource = crxRpt
                BASE_PrintPreviewFrm.CRViewer.ViewReport
                BASE_PrintPreviewFrm.CRViewer.Zoom 1
                Screen.MousePointer = vbDefault
            End If
    End Select
End Sub
Private Sub Save(ByVal iStatusId As Integer, Optional isReopen As Variant)
    If lvItems.ListItems.count > 0 Then
        'On Error GoTo ErrorHandler
        Dim Item As MSComctlLib.ListItem
        Set con = New ADODB.Connection
        Set rec = New ADODB.Recordset
        Set cmd = New ADODB.Command
        
        con.ConnectionString = ConnString
        con.Open
        con.BeginTrans
        
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.Parameters.Append cmd.CreateParameter("@AdjustmentId", adInteger, adParamInputOutput, , AdjustmentId)
        If Trim(txtOrderNumber.Text) = "" Then
            cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInputOutput, 50, Null)
        Else
            cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInputOutput, 50, txtOrderNumber.Text)
        End If
        cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtOrder.value)
        cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , iStatusId)
        cmd.Parameters.Append cmd.CreateParameter("@Percent", adDecimal, adParamInput, , NVAL(txtAdjustmentPercent.Text))
                              cmd.Parameters("@Percent").NumericScale = 2
                              cmd.Parameters("@Percent").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , NVAL(txtAdjustmentAmount.Text))
                              cmd.Parameters("@Amount").NumericScale = 2
                              cmd.Parameters("@Amount").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@Type", adVarChar, adParamInput, 450, cmbType.Text)
        cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 4000, txtRemarks.Text)
        cmd.Parameters.Append cmd.CreateParameter("@WorkStationId", adInteger, adParamInput, , WorkstationId)
        cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
        
        If AdjustmentId = 0 Then
            cmd.CommandText = "SO_Adjustment_Insert"
            cmd.Execute
        Else
            cmd.CommandText = "SO_Adjustment_Update"
            cmd.Execute
        End If
        
        If AdjustmentId = 0 Then 'ADD TO SEARCH
            Set Item = lvSearch.ListItems.add(, , cmd.Parameters("@AdjustmentId"))
                Item.SubItems(1) = cmd.Parameters("@OrderNumber")
                If iStatusId = 1 Then Item.SubItems(2) = "Open"
                If iStatusId = 2 Then Item.SubItems(2) = "Completed"
                Item.Selected = True
                Item.EnsureVisible
            
            AdjustmentId = cmd.Parameters("@AdjustmentId")
            txtOrderNumber.Text = cmd.Parameters("@OrderNumber")
            
            SavePOSAuditTrail UserId, WorkstationId, "", "Created new order Adjustment: " & txtOrderNumber.Text, "SALES"
        Else
            For Each Item In lvSearch.ListItems
                If Val(Item.Text) = Val(cmd.Parameters("@AdjustmentId")) Then
                    Item.SubItems(1) = cmd.Parameters("@OrderNumber")
                    If iStatusId = 1 Then Item.SubItems(2) = "Open"
                    If iStatusId = 2 Then Item.SubItems(2) = "Completed"
                    Item.Selected = True
                    Item.EnsureVisible
                    Exit For
                End If
            Next
            
            Dim auditstatus As String
            If iStatusId = 1 Then auditstatus = "Open"
            If iStatusId = 2 Then auditstatus = "Completed"
            If iStatusId = 7 Then auditstatus = "Cancelled"
            SavePOSAuditTrail UserId, WorkstationId, "", "Updated order Adjustment: " & txtOrderNumber.Text & " - Status: " & auditstatus, "SALES"
        End If
        
        
        
        'SAVE ORDER LINE
        For Each Item In lvItems.ListItems
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc

            cmd.Parameters.Append cmd.CreateParameter("AdjustmentLineId", adInteger, adParamInputOutput, , NVAL(Item.Text))
            cmd.Parameters.Append cmd.CreateParameter("@AdjustmentId", adInteger, adParamInput, , AdjustmentId)
            cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , NVAL(Item.SubItems(2)))
            cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInput, , NVAL(Item.SubItems(3)))
            cmd.Parameters.Append cmd.CreateParameter("@Customer", adVarChar, adParamInput, 400, Item.SubItems(4))
            cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInput, 50, Item.SubItems(5))
            cmd.Parameters.Append cmd.CreateParameter("@DueDate", adDate, adParamInput, , Item.SubItems(6))
            cmd.Parameters.Append cmd.CreateParameter("@DaysDue", adInteger, adParamInput, , Item.SubItems(7))
            cmd.Parameters.Append cmd.CreateParameter("@Total", adDecimal, adParamInput, , NVAL(Item.SubItems(8)))
                                  cmd.Parameters("@Total").NumericScale = 2
                                  cmd.Parameters("@Total").Precision = 18
            cmd.Parameters.Append cmd.CreateParameter("@Balance", adDecimal, adParamInput, , NVAL(Item.SubItems(9)))
                                  cmd.Parameters("@Balance").NumericScale = 2
                                  cmd.Parameters("@Balance").Precision = 18
            cmd.Parameters.Append cmd.CreateParameter("@Adjustment", adDecimal, adParamInput, , NVAL(Item.SubItems(10)))
                                  cmd.Parameters("@Adjustment").NumericScale = 2
                                  cmd.Parameters("@Adjustment").Precision = 18
            cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , iStatusId)
            cmd.Parameters.Append cmd.CreateParameter("@AdjustmentDate", adDate, adParamInput, , dtOrder.value)
            cmd.Parameters.Append cmd.CreateParameter("@Type", adVarChar, adParamInput, 450, cmbType.Text)
            If Item.Text = "" Then
                cmd.CommandText = "SO_AdjustmentLine_Insert"
                cmd.Execute
            Else
                cmd.CommandText = "SO_AdjustmentLine_Update"
                cmd.Execute
            End If
            'cmd.Execute
            Item.Text = cmd.Parameters("AdjustmentLineId")
        Next
        
        con.CommitTrans
        con.Close
        
        txtStatus.Text = "Open"
        StatusId = 1
        
        If iStatusId = 2 Then
            txtStatus.Text = "Completed"
            StatusId = 2
        End If
        
        If iStatusId = 7 Then
            txtStatus.Text = "Cancelled"
            StatusId = 7
            On Error Resume Next
            lvSearch.SelectedItem.SubItems(2) = "Cancelled"
        End If
    Else
        MsgBox "Save failed. No data found.", vbCritical, "PeakPOS"
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
    x = MsgBox("This will complete the transaction. Orders will now be updated. Proceed?", vbExclamation + vbOKCancel)
    If x = vbOK Then
        For Each Item In lvItems.ListItems
            If Item.SubItems(7) = "" Then Item.SubItems(7) = Item.SubItems(6)
        Next
        'ComputeAdjustment
        Save (1)
        Save (2)
        LoadImageStatus picStatus, GetStatus(StatusId)
    End If
End Sub

Private Sub btnSearch_Click()
    Dim con As New ADODB.Connection
    con.ConnectionString = ConnString
    con.Open
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    cmd.ActiveConnection = con
    cmd.CommandText = "SO_Adjustment_Get"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("@AdjustmentId", adInteger, adParamInput, , Null)
    cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInput, 50, txtSearch_OrderNumber.Text)
    If cmbSearch_Status.ListIndex = 0 Then
        cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , Null)
    Else
        cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , cmbSearch_Status.ItemData(cmbSearch_Status.ListIndex))
    End If
    Set rec = cmd.Execute
    lvSearch.ListItems.Clear
    
    Dim strdatefrom, strdateto As String
    strdatefrom = DateFrom.value & " 00:00:00"
    strdateto = DateTo.value & " 23:59:59"
    
    If Not rec.EOF Then
        Do Until rec.EOF
            If rec!Date >= CDate(strdatefrom) And rec!Date <= CDate(strdateto) Then
                Set Item = lvSearch.ListItems.add(, , rec!AdjustmentId)
                    Item.SubItems(1) = rec!OrderNumber
                    Item.SubItems(2) = rec!Status
            End If
            rec.MoveNext
        Loop
    End If
    con.Close
End Sub

Private Sub txtAdjustmentAmount_Change()
    CountTotal
End Sub

Private Sub txtAdjustmentPercent_Change()
    CountTotal
End Sub
