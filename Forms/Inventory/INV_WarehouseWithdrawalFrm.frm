VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form INV_WarehouseWithdrawalFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transfer Stock"
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
   Begin MSComctlLib.ListView lvItemList 
      Height          =   4815
      Left            =   6120
      TabIndex        =   20
      Top             =   2520
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   8493
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
      NumItems        =   7
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
         SubItemIndex    =   3
         Text            =   "Unit"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "On Hand"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Reserved"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Available"
         Object.Width           =   2540
      EndProperty
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
      Left            =   8760
      TabIndex        =   10
      Top             =   7320
      Width           =   2535
   End
   Begin VB.Frame Body_Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   4640
      TabIndex        =   25
      Top             =   0
      Width           =   10455
      Begin MSComctlLib.ListView lvCustomer 
         Height          =   2655
         Left            =   -9999
         TabIndex        =   26
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
      Begin VB.Frame Frame_Header1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1215
         Left            =   120
         TabIndex        =   40
         Top             =   600
         Width           =   6735
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
            Left            =   1200
            TabIndex        =   2
            Top             =   360
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
            Left            =   1200
            TabIndex        =   1
            Top             =   0
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
            Left            =   1200
            TabIndex        =   3
            Top             =   720
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
            Left            =   0
            TabIndex        =   43
            Top             =   0
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
            Left            =   0
            TabIndex        =   42
            Top             =   720
            Width           =   885
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ref #"
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
            TabIndex        =   41
            Top             =   360
            Width           =   465
         End
      End
      Begin VB.Frame Frame_Body 
         BackColor       =   &H00FFFFFF&
         Height          =   5100
         Left            =   120
         TabIndex        =   36
         Top             =   1965
         Width           =   10215
         Begin VB.ComboBox cmbDestination 
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
            TabIndex        =   53
            Top             =   960
            Width           =   3015
         End
         Begin VB.ComboBox cmbOrigin 
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
            TabIndex        =   51
            Top             =   600
            Width           =   3015
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
            Left            =   7080
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   240
            Visible         =   0   'False
            Width           =   3015
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
            TabIndex        =   7
            Top             =   240
            Width           =   3015
         End
         Begin VB.CommandButton btnItemSearch 
            Height          =   330
            Left            =   4320
            Picture         =   "INV_WarehouseWithdrawalFrm.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   240
            Width           =   375
         End
         Begin MSComctlLib.ListView lvItems 
            Height          =   3495
            Left            =   120
            TabIndex        =   8
            Top             =   1440
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
               Text            =   "WarehouseWithdrawalLineId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "WarehouseWithdrawalId"
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
               Text            =   "PrevQuantity"
               Object.Width           =   0
            EndProperty
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Destination"
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
            TabIndex        =   54
            Top             =   960
            Width           =   1080
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Origin"
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
            TabIndex        =   52
            Top             =   600
            Width           =   570
         End
         Begin VB.Label Label11 
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
            Left            =   5880
            TabIndex        =   46
            Top             =   240
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Item Name"
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
            Width           =   1035
         End
      End
      Begin VB.Frame Frame_Footer 
         BackColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   120
         TabIndex        =   32
         Top             =   7080
         Width           =   10215
         Begin VB.TextBox txtWarehouseManager 
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
            Left            =   6840
            TabIndex        =   12
            Top             =   1320
            Width           =   3255
         End
         Begin VB.TextBox txtDriver 
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
            Left            =   6840
            TabIndex        =   11
            Top             =   600
            Width           =   3255
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
            TabIndex        =   9
            Top             =   240
            Width           =   2655
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
            TabIndex        =   33
            Top             =   240
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Warehouse Mngr."
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
            Left            =   6840
            TabIndex        =   45
            Top             =   960
            Width           =   1650
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Driver"
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
            Left            =   6840
            TabIndex        =   44
            Top             =   240
            Width           =   570
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
            TabIndex        =   35
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
            TabIndex        =   34
            Top             =   240
            Visible         =   0   'False
            Width           =   540
         End
      End
      Begin VB.Frame Frame_Header2 
         BackColor       =   &H00FFFFFF&
         Height          =   1395
         Left            =   7170
         TabIndex        =   27
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
            TabIndex        =   4
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
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   960
            Width           =   2055
         End
         Begin MSComCtl2.DTPicker dtOrder 
            Height          =   330
            Left            =   960
            TabIndex        =   5
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
            Format          =   74121217
            CurrentDate     =   41509
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
            TabIndex        =   31
            Top             =   240
            Width           =   690
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
            TabIndex        =   30
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
            TabIndex        =   29
            Top             =   960
            Width           =   570
         End
      End
      Begin MSComctlLib.Toolbar tb_Standard 
         Height          =   330
         Left            =   0
         TabIndex        =   39
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
      BackColor       =   &H00FFC0C0&
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
         TabIndex        =   13
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
         TabIndex        =   14
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
         TabIndex        =   15
         Top             =   2040
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvSearch 
         Height          =   6375
         Left            =   120
         TabIndex        =   16
         Top             =   2520
         Width           =   4335
         _ExtentX        =   7646
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
         TabIndex        =   47
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
         Format          =   74121217
         CurrentDate     =   41686
      End
      Begin MSComCtl2.DTPicker DateFrom 
         Height          =   345
         Left            =   1440
         TabIndex        =   48
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
         Format          =   74121217
         CurrentDate     =   41686
      End
      Begin VB.Label Label15 
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
         TabIndex        =   50
         Top             =   1560
         Width           =   705
      End
      Begin VB.Label Label14 
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
         TabIndex        =   49
         Top             =   1200
         Width           =   960
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
         Top             =   840
         Width           =   570
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
      Picture         =   "INV_WarehouseWithdrawalFrm.frx":0224
      ScaleHeight     =   2295
      ScaleWidth      =   6195
      TabIndex        =   19
      Top             =   2640
      Width           =   6195
   End
   Begin VB.PictureBox pic_Cancelled 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   -9960
      Picture         =   "INV_WarehouseWithdrawalFrm.frx":88B2
      ScaleHeight     =   2295
      ScaleWidth      =   6195
      TabIndex        =   18
      Top             =   2640
      Width           =   6195
   End
   Begin VB.CommandButton btnPaid 
      BackColor       =   &H0080FF80&
      Caption         =   "PAY"
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
      TabIndex        =   17
      Top             =   7200
      Visible         =   0   'False
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
      Picture         =   "INV_WarehouseWithdrawalFrm.frx":1954E
      ScaleHeight     =   1860
      ScaleWidth      =   5250
      TabIndex        =   0
      Top             =   2640
      Visible         =   0   'False
      Width           =   5250
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
            Picture         =   "INV_WarehouseWithdrawalFrm.frx":224DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_WarehouseWithdrawalFrm.frx":28D41
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_WarehouseWithdrawalFrm.frx":2F5A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_WarehouseWithdrawalFrm.frx":35E05
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_WarehouseWithdrawalFrm.frx":3607A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_WarehouseWithdrawalFrm.frx":366EB
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "INV_WarehouseWithdrawalFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public StatusId, CustomerId, WarehouseWithdrawalId, id As Integer
Dim TotalSacks As Double
Dim OrderLine(100000) As Long
Dim ctrOrderLine As Long

Public Sub Initialize()
    Dim txtControl As Control
    For Each txtControl In Me.Controls
        If TypeOf txtControl Is TextBox And txtControl.Name <> "txtSearch_Order" Then
            txtControl.text = ""
            txtStatus.text = "Open"
        End If
    Next
    
    isNotCompleted (True)
    'lblBalance.Caption = "0.00"
    cmbCustomer.text = ""
    cmbCustomer.BackColor = &HC0C0FF
    'cmbTerms.BackColor = &HC0C0FF
    dtOrder.value = Format(Now, "MM/DD/YY")
    'dtDue.value = Format(Now, "MM/DD/YY")
    btnStatus.Caption = "Complete Order"
    lvItems.ListItems.Clear
    lvCustomer.ListItems.Clear
    lvCustomer.Visible = False
    lvItemList.Visible = False
    lvCustomer.Left = -9999
    lvItemList.Left = -9999
    picCompleted.Left = -9999
    pic_Cancelled.Left = -9999
    picPaid.Left = -9999
    btnPaid.Visible = False
    txtRemarks.text = "Withdraw"
    tb_Standard.Buttons(4).Caption = "Cancel"
    tb_Standard.Buttons(4).Image = 3
    'lblSubTotal.Caption = "0.00"
    'lblTotal.Caption = "0.00"
    'lblRemainingBalance.Caption = "0.00"
    
    id = 1
    StatusId = 1
    CustomerId = 0
    WarehouseWithdrawalId = 0
    TotalSacks = 0
    
    ctrOrderLine = 0
    
    On Error Resume Next
    cmbCustomer.SetFocus
    SO_ItemSelectOptionFrm.isModify = False
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
        cmd.Parameters.Append cmd.CreateParameter("@WarehouseWithdrawalId", adInteger, adParamInputOutput, , WarehouseWithdrawalId)
        If Trim(txtOrderNumber.text) = "" Then
            cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInputOutput, 50, Null)
        Else
            cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInputOutput, 50, txtOrderNumber.text)
        End If
        cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtOrder.value)
        cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , StatusId)
        cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInput, , CustomerId)
        cmd.Parameters.Append cmd.CreateParameter("@ReferenceNumber", adVarChar, adParamInput, 500, txtReferenceNumber.text)
        cmd.Parameters.Append cmd.CreateParameter("@Salesman", adVarChar, adParamInput, 500, txtSalesman.text)
        cmd.Parameters.Append cmd.CreateParameter("@Driver", adVarChar, adParamInput, 500, txtDriver.text)
        cmd.Parameters.Append cmd.CreateParameter("@WarehouseManager", adVarChar, adParamInput, 500, txtWarehouseManager.text)
        cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
        cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 500, txtRemarks.text)
        cmd.Parameters.Append cmd.CreateParameter("@FromLocationId", adInteger, adParamInput, , cmbOrigin.ItemData(cmbOrigin.ListIndex))
        cmd.Parameters.Append cmd.CreateParameter("@ToLocationId", adInteger, adParamInput, , cmbDestination.ItemData(cmbDestination.ListIndex))
        
        If WarehouseWithdrawalId = 0 Then
            cmd.CommandText = "INV_WarehouseWithdrawal_Insert"
        Else
            'cmd.Parameters.Append cmd.CreateParameter("@isReopen", adBoolean, adParamInput, , isReopen)
            cmd.CommandText = "INV_WarehouseWithdrawal_Update"
        End If
        
        cmd.Execute
        WarehouseWithdrawalId = cmd.Parameters("@WarehouseWithdrawalId")
        txtOrderNumber.text = cmd.Parameters("@OrderNumber")
        
        'SAVE ORDER LINE
        Dim item As MSComctlLib.ListItem

        For Each item In lvItems.ListItems
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc

            cmd.Parameters.Append cmd.CreateParameter("@WarehouseWithdrawalLineId", adInteger, adParamInputOutput, , Val(item.text))
            cmd.Parameters.Append cmd.CreateParameter("@WarehouseWithdrawalId", adInteger, adParamInput, , WarehouseWithdrawalId)
            cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , Val(item.SubItems(6)))
            cmd.Parameters.Append cmd.CreateParameter("@Quantity", adDecimal, adParamInput, , Val(Replace(item.SubItems(4), ",", "")))
                                  cmd.Parameters("@Quantity").Precision = 18
                                  cmd.Parameters("@Quantity").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@Uom", adVarChar, adParamInput, 250, item.SubItems(5))
            cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , StatusId)
            cmd.Parameters.Append cmd.CreateParameter("@FromLocationId", adInteger, adParamInput, , cmbOrigin.ItemData(cmbOrigin.ListIndex))
            cmd.Parameters.Append cmd.CreateParameter("@ToLocationId", adInteger, adParamInput, , cmbDestination.ItemData(cmbDestination.ListIndex))
            cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, txtRemarks.text)

            If item.text = "" Then
                cmd.CommandText = "INV_WarehouseWithdrawalLine_Insert"
            Else
                cmd.Parameters.Append cmd.CreateParameter("@PrevQuantity", adDecimal, adParamInput, , Val(Replace(item.SubItems(7), ",", "")))
                                  cmd.Parameters("@PrevQuantity").Precision = 18
                                  cmd.Parameters("@PrevQuantity").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@isReopen", adBoolean, adParamInput, , isReopen)
                cmd.CommandText = "INV_WarehouseWithdrawalLine_Update"
            End If
            cmd.Execute
            item.text = cmd.Parameters("@WarehouseWithdrawalLineId")
        Next
'
        'DELETE ORDERLINE IF ANY
        Dim ctr As Integer
        For ctr = 0 To ctrOrderLine
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc

            If OrderLine(ctr) <> 0 Then
                cmd.Parameters.Append cmd.CreateParameter("@WarehouseWithdrawalLineId", adInteger, adParamInput, , OrderLine(ctr))
                cmd.CommandText = "INV_WarehouseWithdrawalLine_Delete"
                cmd.Execute
            Else
                Exit For
            End If
        Next
'
        con.CommitTrans
        con.Close

        If StatusId = 2 Then
            isNotCompleted (False)
            txtStatus.text = "Completed"
            picCompleted.Left = 6360
            picCompleted.Visible = False
            Me.StatusId = 2
        End If
'
        Dim isFound As Boolean
        isFound = False
        For Each item In lvSearch.ListItems
            If WarehouseWithdrawalId = item.text Then
                item.SubItems(1) = txtOrderNumber.text
                item.SubItems(2) = txtStatus.text
                isFound = True
                item.Selected = True
                item.EnsureVisible
                Exit For
            End If
        Next
        If isFound = False Then
            Set item = lvSearch.ListItems.add(, , WarehouseWithdrawalId)
                item.SubItems(1) = txtOrderNumber.text
                item.SubItems(2) = txtStatus.text
                item.Selected = True
                item.EnsureVisible
        End If
        
'        'PRINT PREVIEW
'        If StatusId = 2 Then
'            'Screen.MousePointer = vbHourglass
'            'BASE_PrintPreviewFrm.Show
'            Dim crxApp As New CRAXDRT.Application
'            Dim crxRpt As New CRAXDRT.Report
'            Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\INV_WarehouseWithdrawal.rpt")
'            crxRpt.RecordSelectionFormula = "{INV_WarehouseWithdrawal.WarehouseWithdrawalId}= " & WarehouseWithdrawalId & ""
'            crxRpt.DiscardSavedData
'
'            Call ResetRptDB(crxRpt)
'            crxRpt.PrintOut False
'
''            BASE_PrintPreviewFrm.CRViewer.ReportSource = crxRpt
''            BASE_PrintPreviewFrm.CRViewer.ViewReport
''            BASE_PrintPreviewFrm.CRViewer.Zoom 1
''            Screen.MousePointer = vbDefault
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
'    Dim total, subtotal, discount, subtotal1, interestrate, Interest, days, cash As Double
'    Dim item As MSComctlLib.ListItem
'
'    If IsNumeric(txtDays.text) = False Then
'        days = 0
'    Else
'        days = txtDays.text
'    End If
'
'    If IsNumeric(txtInterestRate.text) = False Then
'        interestrate = 0
'    Else
'        interestrate = txtInterestRate.text
'    End If
'
'    If IsNumeric(txtCash.text) = False Then
'        cash = 0
'    Else
'        cash = txtCash.text
'    End If
'
'    If IsNumeric(txtDiscount.text) = False Then
'        discount = 0
'    Else
'        discount = txtDiscount.text
'    End If
'
'    For Each item In lvItems.ListItems
'        subtotal = Val(Replace(item.SubItems(4), ",", "")) * Val(Replace(item.SubItems(6), ",", ""))
'        item.SubItems(7) = FormatNumber(subtotal, 2, vbTrue, vbFalse)
'        subtotal1 = subtotal1 + subtotal
'    Next
'    subtotal1 = subtotal1 + cash - discount
'    lblSubTotal.Caption = FormatNumber(subtotal1, 2, vbTrue, vbFalse)
'
'    'COMPUTE INTEREST
'    Interest = ((days / 30) * (subtotal1 * (interestrate / 100)))
'    lblInterest.Caption = FormatNumber(Interest, 2, vbTrue, vbFalse)
'
'    total = Interest + subtotal1
'    lblTotal.Caption = FormatNumber(total, 2, vbTrue, vbFalse)
'    lblRemainingBalance.Caption = lblTotal.Caption
End Sub
Public Sub Populate(ByVal data As String)
    Select Case data
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
                    'lblBalance.Caption = FormatNumber(rec!balance, 2, vbTrue, vbFalse)
                    rec.MoveNext
                Loop
            End If
            con.Close
        Case "WarehouseWithdrawalGet"
            Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            Set rec = New ADODB.Recordset
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "INV_WarehouseWithdrawal_Get"
            cmd.Parameters.Append cmd.CreateParameter("@WarehouseWithdrawalId", adInteger, adParamInput, , WarehouseWithdrawalId)
            Set rec = cmd.Execute
            If Not rec.EOF Then
                'If rec!isCashAdvance = "False" Then
                    CustomerId = rec!CustomerId
                    'lblBalance.Caption = FormatNumber(rec!balance, 2, vbTrue, vbFalse)
                    cmbCustomer.BackColor = vbWhite
                    StatusId = rec!StatusId
                    txtStatus.text = rec!Status
                    txtOrderNumber.text = rec!OrderNumber
                    'txtAddress.text = rec!address
                    'txtContact.text = rec!phone
                    dtOrder.value = Format(rec!Date, "MM/DD/YY")
                    'dtDue.value = Format(rec!duedate, "MM/DD/YY")
                    'lblSubTotal.Caption = FormatNumber(rec!subtotal, 2, vbTrue)
                    txtOrderNumber.text = rec!OrderNumber
                    txtRemarks.text = rec!Remarks
                    txtReferenceNumber.text = rec!ReferenceNumber
                    txtSalesman.text = rec!salesman
                    'txtDays.text = rec!days
                    'txtInterestRate.text = rec!interestrate
                    'txtCash.text = FormatNumber(rec!cash, 2, vbTrue, vbFalse)
                    'lblRemainingBalance.Caption = FormatNumber(rec!balance, 2, vbTrue, vbFalse)
                    'On Error Resume Next
'                    If IsNull(rec!discount) = True Then
'                        'txtDiscount.text = ""
'                    Else
'                        'txtDiscount.text = FormatNumber(rec!discount, 2, vbTrue, vbFalse)
'                    End If
'
'                    If IsNull(rec!Terms) Then
'                        'cmbTerms.ListIndex = 0
'                    Else
'                        'cmbTerms.text = rec!Terms
'                    End If
                    
                    If rec!StatusId = 2 Then 'COMPLETED
                        isNotCompleted (False)
                        picCompleted.Left = 6360
                        pic_Cancelled.Left = -9999
                        picPaid.Left = -9999
                        picCompleted.Visible = False
                        pic_Cancelled.Visible = False
                        picPaid.Visible = False
                        tb_Standard.Buttons(4).Caption = "Cancel"
                        tb_Standard.Buttons(4).Image = 3
                        btnPaid.Visible = True
                        btnStatus.Caption = "Reopen Order"
                    ElseIf rec!StatusId = 4 Then 'Cancelled
                        isNotCompleted (False)
                        picCompleted.Left = -9999
                        pic_Cancelled.Left = 6360
                        picPaid.Left = -9999
                        picPaid.Visible = False
                        picCompleted.Visible = False
                        pic_Cancelled.Visible = True
                        tb_Standard.Buttons(4).Caption = "Activate"
                        tb_Standard.Buttons(4).Image = 6
                        btnPaid.Visible = False
                        btnStatus.Caption = "Complete Order"
                    ElseIf rec!StatusId = 3 Then 'PAID
                        isNotCompleted (False)
                        picPaid.Left = 6890
                        picCompleted.Left = -9999
                        pic_Cancelled.Left = -9999
                        picPaid.Visible = True
                        picCompleted.Visible = False
                        pic_Cancelled.Visible = False
                        tb_Standard.Buttons(4).Caption = "Cancel"
                        tb_Standard.Buttons(4).Image = 3
                        btnPaid.Visible = False
                        btnStatus.Caption = "Reopen Order"
                    Else
                        isNotCompleted (True)
                        picCompleted.Left = -9999
                        pic_Cancelled.Left = -9999
                        picPaid.Left = -9999
                        picCompleted.Visible = False
                        pic_Cancelled.Visible = False
                        picPaid.Visible = False
                        tb_Standard.Buttons(4).Caption = "Cancel"
                        tb_Standard.Buttons(4).Image = 3
                        btnPaid.Visible = False
                        btnStatus.Caption = "Complete Order"
                    End If
                    'lblRemainingBalance.Caption = FormatNumber(rec!balance, 2, vbTrue, vbFalse)
                    'PREVENT ERROR ON DISPLAY FOR cmbCustomer_Change Event
                    On Error Resume Next
                    cmbCustomer.text = rec!Name
                    lvCustomer.Left = -9999
                    lvCustomer.Visible = False
                'End If
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
                        Set item = lvSearch.ListItems.add(, , rec!WarehouseWithdrawalId)
                            item.SubItems(1) = rec!OrderNumber
                            item.SubItems(2) = rec!Status
                    End If
                    rec.MoveNext
                Loop
            End If
            con.Close
        Case "WarehouseWithdrawalLineGet"
            Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            Set rec = New ADODB.Recordset
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "INV_WarehouseWithdrawalLine_Get"
            cmd.Parameters.Append cmd.CreateParameter("@WarehouseWithdrawalId", adInteger, adParamInput, , WarehouseWithdrawalId)
            Set rec = cmd.Execute
            lvItems.ListItems.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    Set item = lvItems.ListItems.add(, , rec!WarehouseWithdrawalLineId)
                        item.SubItems(1) = rec!WarehouseWithdrawalId
                        item.SubItems(2) = rec!itemcode
                        item.SubItems(3) = rec!Name
                        item.SubItems(4) = FormatNumber(rec!quantity, 2, vbTrue)
                        item.SubItems(5) = rec!Uom
                        item.SubItems(6) = rec!ProductId
                        item.SubItems(7) = rec!quantity
                    rec.MoveNext
                Loop
            End If
            con.Close
        Case "Location"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Location")
            cmbOrigin.Clear
            cmbDestination.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        cmbOrigin.AddItem rec!Location
                        cmbOrigin.ItemData(cmbOrigin.NewIndex) = rec!LocationId
                        cmbDestination.AddItem rec!Location
                        cmbDestination.ItemData(cmbDestination.NewIndex) = rec!LocationId
                    End If
                    rec.MoveNext
                Loop
            End If
            On Error Resume Next
            cmbOrigin.ListIndex = 1
            cmbDestination.ListIndex = 2
    End Select
End Sub
Private Sub btnAR_Click()
    'PO_AccountsReceivableFrm.txtAR.text = lblSubTotal.Caption
    'PO_AccountsReceivableFrm.Show (1)
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
    cmd.CommandText = "BASE_Product_Search2"
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 500, txtItemSearch.text)
    
    Dim QtyonHand, Reserved As Double
    Dim PrevProductId As Long
    
    Set rec = cmd.Execute
    If Not rec.EOF Then
        lvItemList.ListItems.Clear
        Do Until rec.EOF
            If rec!isActive = "True" Then
                If PrevProductId <> rec!ProductId Then
                    QtyonHand = 0
                    Reserved = 0
                    Set item = lvItemList.ListItems.add(, , rec!ProductId)
                        item.SubItems(1) = rec!itemcode
                        item.SubItems(2) = rec!Name
                        If rec!Location = "<RESERVED>" Then
                            Reserved = Reserved + rec!quantity
                        Else
                            QtyonHand = QtyonHand + rec!quantity
                        End If
                        item.SubItems(3) = rec!Uom
                        item.SubItems(4) = FormatNumber(QtyonHand, 2, vbTrue, vbFalse)
                        item.SubItems(5) = FormatNumber(Reserved, 2, vbTrue, vbFalse)
                        item.SubItems(6) = FormatNumber(QtyonHand - Reserved, 2, vbTrue, vbFalse)
                Else
                        If rec!Location = "<RESERVED>" Then
                            Reserved = Reserved + rec!quantity
                        Else
                            QtyonHand = QtyonHand + rec!quantity
                        End If
                        lvItemList.ListItems(lvItemList.ListItems.Count).SubItems(4) = FormatNumber(QtyonHand, 2, vbTrue, vbFalse)
                        lvItemList.ListItems(lvItemList.ListItems.Count).SubItems(5) = FormatNumber(Reserved, 2, vbTrue, vbFalse)
                        lvItemList.ListItems(lvItemList.ListItems.Count).SubItems(6) = FormatNumber(QtyonHand - Reserved, 2, vbTrue, vbFalse)
                End If
                PrevProductId = rec!ProductId
                lvItemList.Top = 2520
                lvItemList.Visible = True
                lvItemList.Left = 6070
            End If
            rec.MoveNext
        Loop
    Else
        lvItemList.Visible = False
        lvItemList.Left = -9999
    End If
    'DistinctList lvItemList
    con.Close
End Sub

'Private Sub btnPaid_Click()
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
'End Sub

Private Sub btnSearch_Click()
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "INV_WarehouseWithdrawal_Get"
    cmd.Parameters.Append cmd.CreateParameter("@WarehouseWithdrawalId", adInteger, adParamInput, , Null)
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
            If rec!Date >= DateFrom.value And rec!Date <= DateTo.value Then
                Set item = lvSearch.ListItems.add(, , rec!WarehouseWithdrawalId)
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
        Dim X As Variant
        X = MsgBox("Are you sure you want to reopen this order? This will revert all connected " & _
                    "transactions such as inventory movements.", vbCritical + vbYesNo, "WARNING")
        If X = vbYes Then
            Save 1, True
            btnStatus.Caption = "Complete Order"
            txtStatus.text = "Open"
            btnPaid.Visible = False
            isNotCompleted (True)
            picCompleted.Visible = False
            picCompleted.Left = -9999
            picPaid.Left = -9999
            picPaid.Visible = False
            Me.StatusId = 1
        Else
        End If
    End If
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
On Error Resume Next
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
        Case vbKeyF4
            txtItemSearch.SetFocus
        Case vbKeyF1
            Unload INV_ProductSearch
            INV_ProductSearch.isWithdraw = True
            INV_ProductSearch.isPO = False
            INV_ProductSearch.isSO = False
            INV_ProductSearch.Show (1)
    End Select
End Sub

Private Sub Form_Load()
    '****** REGION Listview Columns *********
    lvSearch.ColumnHeaders(2).width = lvSearch.width * 0.48
    lvSearch.ColumnHeaders(3).width = lvSearch.width * 0.45
    
    'lvItemList.ColumnHeaders(1).width = lvItemList.width * 0.249
    lvItemList.ColumnHeaders(2).width = lvItemList.width * 0.1633
    lvItemList.ColumnHeaders(3).width = lvItemList.width * 0.4433
    lvItemList.ColumnHeaders(4).width = lvItemList.width * 0.1233
    lvItemList.ColumnHeaders(5).width = lvItemList.width * 0.1233
    lvItemList.ColumnHeaders(6).width = lvItemList.width * 0.1233
    lvItemList.ColumnHeaders(7).width = lvItemList.width * 0.1233
    
    lvCustomer.ColumnHeaders(2).width = lvCustomer.width * 0.25
    lvCustomer.ColumnHeaders(3).width = lvCustomer.width * 0.42
    lvCustomer.ColumnHeaders(4).width = lvCustomer.width * 0.28
    
    lvItems.ColumnHeaders(3).width = lvItems.width * 0.135
    lvItems.ColumnHeaders(4).width = lvItems.width * 0.605
    lvItems.ColumnHeaders(5).width = lvItems.width * 0.125
    lvItems.ColumnHeaders(6).width = lvItems.width * 0.105
'    lvItems.ColumnHeaders(8).width = lvItems.width * 0.2
    
    'StatusBarWidth Me, statusBar_Main
    '****************************************
    
    Initialize
    Populate "Terms"
    Populate "Status"
    Populate "Location"
    'Populate "SalesOrderLoad"
    
    DateFrom.value = Format(Now, "MM/DD/YY")
    DateTo.value = Format(Now, "MM/DD/YY")
    
    On Error Resume Next
    cmbSearch_Status.ListIndex = 1
    btnSearch_Click
End Sub






Private Sub lblGrossAmount_Click()

End Sub

Private Sub lblGrossKilos_Click()

End Sub

Private Sub lblSubTotal_Click()
End Sub

Private Sub lblInterest_Click()

End Sub

Private Sub Label23_Click()

End Sub

Private Sub lvCustomer_DblClick()
    If lvCustomer.ListItems.Count > 0 Then
        CustomerId = lvCustomer.SelectedItem.text
        cmbCustomer.text = lvCustomer.SelectedItem.SubItems(2)
        'lblBalance.Caption = lvCustomer.SelectedItem.SubItems(3)
        'txtContact.text = lvCustomer.SelectedItem.SubItems(4)
        'txtAddress.text = lvCustomer.SelectedItem.SubItems(5)
        lvCustomer.Visible = False
        lvCustomer.Left = -9999
        'txtItemSearch.SetFocus
        txtReferenceNumber.SetFocus
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
    'SO_ItemSelectOptionFrm.txtPrice.text = lvItemList.SelectedItem.SubItems(3)
    'SO_ItemSelectOptionFrm.Show (1)
    If lvItemList.ListItems.Count > 0 Then
        Unload INV_WarehouseWithdrawalOptFrm
        INV_WarehouseWithdrawalOptFrm.isModify = False
        INV_WarehouseWithdrawalOptFrm.isFormSearch = False
        INV_WarehouseWithdrawalOptFrm.lblAvailable.Caption = lvItemList.SelectedItem.SubItems(6)
        INV_WarehouseWithdrawalOptFrm.Show (1)
       
    End If
End Sub

Private Sub lvItemList_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If lvItemList.ListItems.Count > 0 Then
                'SO_ItemSelectOptionFrm.txtPrice.text = lvItemList.SelectedItem.SubItems(3)
                'SO_ItemSelectOptionFrm.Show (1)
                Unload INV_WarehouseWithdrawalOptFrm
                INV_WarehouseWithdrawalOptFrm.isModify = False
                INV_WarehouseWithdrawalOptFrm.isFormSearch = False
                INV_WarehouseWithdrawalOptFrm.lblAvailable.Caption = lvItemList.SelectedItem.SubItems(6)
                INV_WarehouseWithdrawalOptFrm.Show (1)
            End If
        
    End Select
End Sub

Private Sub lvItems_DblClick()
    Unload INV_WarehouseWithdrawalOptFrm
    If lvItems.ListItems.Count > 0 Then
        With INV_WarehouseWithdrawalOptFrm
            .txtQuantity.text = lvItems.SelectedItem.SubItems(4)
            '.txtPrice.text = lvItems.SelectedItem.SubItems(6)
            .isModify = True
            .Show (1)
        End With
    End If
End Sub

Private Sub lvItems_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp
            If lvItems.ListItems.Count > 0 Then
                If lvItems.SelectedItem.Index = 1 Then
                    'txtItemSearch.SetFocus
                    txtCode.SetFocus
                End If
            End If
        Case vbKeyDelete
            If lvItems.ListItems.Count > 0 Then
                If lvItems.SelectedItem.text <> "" Then
                    OrderLine(ctrOrderLine) = Val(lvItems.SelectedItem.text)
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

Private Sub lvSearch_ItemClick(ByVal item As MSComctlLib.ListItem)
    If lvSearch.ListItems.Count > 0 Then
        WarehouseWithdrawalId = lvSearch.SelectedItem.text
        Populate "WarehouseWithdrawalLineGet"
        CountTotal
        Populate "WarehouseWithdrawalGet"
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
            If WarehouseWithdrawalId <> 0 Then
                If StatusId = 1 Then 'Status Open > Cancelled
                    StatusId = 4
                    txtStatus.text = "Cancelled"
                    Save (4)
                    isNotCompleted (False)
                    pic_Cancelled.Left = 6360
                    pic_Cancelled.Visible = True
                    tb_Standard.Buttons(4).Caption = "Activate"
                    tb_Standard.Buttons(4).Image = 6
                ElseIf StatusId = 4 Then 'Status Cancelled > Activate
                    StatusId = 1
                    txtStatus.text = "Open"
                    Save (1)
                    isNotCompleted (True)
                    pic_Cancelled.Left = -9999
                    pic_Cancelled.Visible = False
                    tb_Standard.Buttons(4).Caption = "Cancel"
                    tb_Standard.Buttons(4).Image = 3
                End If
            End If
        Case 6 'PRINT PREVIEW
'            If WarehouseWithdrawalId <> 0 Then
'                Screen.MousePointer = vbHourglass
'                BASE_PrintPreviewFrm.Show
'                Dim crxApp As New CRAXDRT.Application
'                Dim crxRpt As New CRAXDRT.Report
'                Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\INV_WarehouseWithdrawal.rpt")
'                crxRpt.RecordSelectionFormula = "{INV_WarehouseWithdrawal.WarehouseWithdrawalId}= " & WarehouseWithdrawalId & ""
'                crxRpt.DiscardSavedData
'
'                Call ResetRptDB(crxRpt)
'                'crxRpt.PrintOut False
'
'                BASE_PrintPreviewFrm.CRViewer.ReportSource = crxRpt
'                BASE_PrintPreviewFrm.CRViewer.ViewReport
'                BASE_PrintPreviewFrm.CRViewer.Zoom 1
'                Screen.MousePointer = vbDefault
'            End If
    End Select
End Sub



Private Sub Text1_Change()

End Sub

Private Sub txtCode_Change()
    If Trim(txtCode.text) = "" Then
        Exit Sub
        lvItemList.Visible = False
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
    cmd.Parameters.Append cmd.CreateParameter("@ItemCode", adVarChar, adParamInput, 50, txtCode.text)
    Dim QtyonHand, Reserved As Double
    Dim PrevProductId As Long
    
    Set rec = cmd.Execute
    If Not rec.EOF Then
        lvItemList.ListItems.Clear
        Do Until rec.EOF
            If rec!isActive = "True" Then
                If PrevProductId <> rec!ProductId Then
                    QtyonHand = 0
                    Reserved = 0
                    Set item = lvItemList.ListItems.add(, , rec!ProductId)
                        item.SubItems(1) = rec!itemcode
                        item.SubItems(2) = rec!Name
                        If rec!Location = "<RESERVED>" Then
                            Reserved = Reserved + rec!quantity
                        Else
                            QtyonHand = QtyonHand + rec!quantity
                        End If
                        item.SubItems(3) = rec!Uom
                        item.SubItems(4) = FormatNumber(QtyonHand, 2, vbTrue, vbFalse)
                        item.SubItems(5) = FormatNumber(Reserved, 2, vbTrue, vbFalse)
                        item.SubItems(6) = FormatNumber(QtyonHand - Reserved, 2, vbTrue, vbFalse)
                Else
                        If rec!Location = "<RESERVED>" Then
                            Reserved = Reserved + rec!quantity
                        Else
                            QtyonHand = QtyonHand + rec!quantity
                        End If
                        lvItemList.ListItems(lvItemList.ListItems.Count).SubItems(4) = FormatNumber(QtyonHand, 2, vbTrue, vbFalse)
                        lvItemList.ListItems(lvItemList.ListItems.Count).SubItems(5) = FormatNumber(Reserved, 2, vbTrue, vbFalse)
                        lvItemList.ListItems(lvItemList.ListItems.Count).SubItems(6) = FormatNumber(QtyonHand - Reserved, 2, vbTrue, vbFalse)
                End If
                PrevProductId = rec!ProductId
                lvItemList.Visible = True
                lvItemList.Left = 6070
                lvItemList.Top = 2520
            End If
            rec.MoveNext
        Loop
'        For Each item In lvItemList.ListItems
'            If item.SubItems(1) Like txtCode.text Then
'                item.Selected = True
'                item.EnsureVisible
'                Exit For
'            End If
'        Next
    Else
        lvItemList.Visible = False
        lvItemList.Left = -9999
    End If
    'DistinctList lvItemList
    con.Close
End Sub

Private Sub txtCode_GotFocus()
    selectText txtCode
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

Private Sub txtItemSearch_GotFocus()
    selectText txtItemSearch
End Sub

Private Sub txtItemSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp, vbKeyDown
            If lvItemList.Visible = True Then
                lvItemList.SetFocus
            Else
                lvItems.SetFocus
            End If
        Case vbKeyReturn
            If Trim(txtItemSearch.text) = "" Then
                Exit Sub
                lvItemList.Visible = False
            Else
                btnItemSearch_Click
            End If
    End Select
End Sub

Private Sub txtReferenceNumber_GotFocus()
    selectText txtReferenceNumber
End Sub

Private Sub txtRemarks_GotFocus()
    selectText txtRemarks
End Sub

Private Sub txtSalesman_GotFocus()
    selectText txtSalesman
End Sub

Private Sub txtSearch_OrderNumber_Change()
    btnSearch_Click
End Sub

Private Sub txtSearch_OrderNumber_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then btnSearch_Click
End Sub


