VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form SO_PickOrderFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13215
   Icon            =   "SO_PickupOrderFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   13215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvItemList 
      Height          =   3735
      Left            =   1200
      TabIndex        =   0
      Top             =   2835
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
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
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Uom"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Frame Fre_Body 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7815
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   12735
      Begin VB.TextBox txtreceivedby 
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
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   7440
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.ComboBox cmblocation 
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
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1320
         Width           =   3015
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
         Height          =   1050
         Left            =   1200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   6360
         Width           =   5895
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
         Left            =   0
         TabIndex        =   15
         Top             =   1320
         Width           =   3405
      End
      Begin VB.Frame Frame_Body 
         BackColor       =   &H00FFFFFF&
         Height          =   4620
         Left            =   0
         TabIndex        =   8
         Top             =   1680
         Width           =   12735
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
            Left            =   9600
            TabIndex        =   11
            Top             =   240
            Visible         =   0   'False
            Width           =   3015
         End
         Begin VB.CommandButton btnItemSearch 
            Height          =   330
            Left            =   5880
            Picture         =   "SO_PickupOrderFrm.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   10
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
            Left            =   960
            TabIndex        =   9
            Top             =   240
            Visible         =   0   'False
            Width           =   4935
         End
         Begin MSComctlLib.ListView lvItems 
            Height          =   3855
            Left            =   120
            TabIndex        =   12
            Top             =   600
            Width           =   12495
            _ExtentX        =   22040
            _ExtentY        =   6800
            View            =   3
            LabelEdit       =   1
            MultiSelect     =   -1  'True
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
            NumItems        =   17
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "ReceiveOrderLineId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "ReceiveOrderId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "ProductId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Itemcode"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Qty"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Unit"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   7
               Text            =   "Location"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   8
               Text            =   "Date"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   9
               Text            =   "Reference"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "LocationId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Text            =   "ReceivedQty"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Text            =   "OldQty"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   13
               Text            =   "PurchaseOrderLineId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   14
               Text            =   "ActualQuantity"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   15
               Text            =   "ReserveId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   16
               Text            =   "SalesOrderLineId"
               Object.Width           =   0
            EndProperty
         End
         Begin VB.Label Label9 
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
            Left            =   8400
            TabIndex        =   14
            Top             =   240
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
            Left            =   240
            TabIndex        =   13
            Top             =   240
            Visible         =   0   'False
            Width           =   555
         End
      End
      Begin VB.Frame Frame_Header2 
         BackColor       =   &H00FFFFFF&
         Height          =   1275
         Left            =   9120
         TabIndex        =   3
         Top             =   0
         Width           =   3525
         Begin VB.TextBox txtDeliveryNumber 
            Appearance      =   0  'Flat
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
            Top             =   600
            Width           =   2055
         End
         Begin VB.TextBox txtOrderNumber 
            Appearance      =   0  'Flat
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
            TabIndex        =   4
            Top             =   240
            Width           =   2055
         End
         Begin VB.Label Label1 
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
            TabIndex        =   7
            Top             =   600
            Width           =   465
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SO #"
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
            TabIndex        =   6
            Top             =   240
            Width           =   405
         End
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Picked By"
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
         TabIndex        =   27
         Top             =   7440
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
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
         Left            =   3720
         TabIndex        =   26
         Top             =   1320
         Width           =   780
      End
      Begin VB.Label lblShowHistory 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Show Received History"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF8080&
         Height          =   225
         Left            =   10815
         MouseIcon       =   "SO_PickupOrderFrm.frx":0230
         MousePointer    =   99  'Custom
         TabIndex        =   25
         Top             =   1440
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.Label lblReceived 
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
         Left            =   10920
         TabIndex        =   24
         Top             =   6765
         Width           =   1740
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Picked"
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
         Left            =   9120
         TabIndex        =   23
         Top             =   6765
         Width           =   1185
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Ordered"
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
         Left            =   9120
         TabIndex        =   22
         Top             =   6480
         Width           =   1350
      End
      Begin VB.Label lblTotal 
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
         Height          =   285
         Left            =   10800
         TabIndex        =   21
         Top             =   6480
         Width           =   1860
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
         Left            =   0
         TabIndex        =   20
         Top             =   6360
         Width           =   795
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "SO_PickupOrderFrm.frx":0382
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pick-up Orders"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   720
         TabIndex        =   19
         Top             =   240
         Width           =   1725
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   12480
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
            Picture         =   "SO_PickupOrderFrm.frx":2C88
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_PickupOrderFrm.frx":94EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_PickupOrderFrm.frx":FD4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_PickupOrderFrm.frx":165AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_PickupOrderFrm.frx":16823
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_PickupOrderFrm.frx":16E94
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tb_Standard 
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   582
      ButtonWidth     =   1349
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
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Print"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   8175
      Left            =   120
      Top             =   360
      Width           =   12975
   End
End
Attribute VB_Name = "SO_PickOrderFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public PurchaseOrderId As Integer
Dim PickOrderId As Long
Public Sub CountTotalQty()
    Dim totalQty As Double
    Dim totalReceived As Double
    Dim item As MSComctlLib.ListItem
    With SO_SalesOrderFrm
        For Each item In .lvItems.ListItems
            totalQty = totalQty + Val(Replace(item.SubItems(4), ",", ""))
            totalReceived = totalReceived + Val(Replace(item.SubItems(12), ",", ""))
        Next
    End With
    lblTotal.Caption = FormatNumber(totalQty, 2, vbTrue, vbFalse)
    lblReceived.Caption = FormatNumber(totalReceived, 2, vbTrue, vbFalse)
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
                        item.SubItems(3) = FormatNumber(rec!unitcost, 2, vbTrue, vbFalse)
                        item.SubItems(4) = rec!Uom
                    lvItemList.Visible = True
                    lvItemList.Left = 1200
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
    'AutoFill
    Dim item As MSComctlLib.ListItem
    Dim rItem As MSComctlLib.ListItem
    Dim x As Integer
    Dim isFound As Boolean
    
    For x = 1 To lvItems.ListItems.Count
        If x > lvItems.ListItems.Count Then Exit For
        If lvItems.ListItems(x).Text = "" Then
            lvItems.ListItems.Remove (x)
            x = x - 1
        End If
    Next
    
    If SO_SalesOrderFrm.StatusId = 5 Then Exit Sub 'Status of order is PICKED
    With SO_SalesOrderFrm
        For Each item In .lvItems.ListItems
            If Val(Replace(item.SubItems(4), ",", "")) - Val(item.SubItems(12)) > 0 Then
                Set rItem = lvItems.ListItems.add(, , "")
                    rItem.SubItems(2) = item.SubItems(9)
                    rItem.SubItems(3) = item.SubItems(2)
                    rItem.SubItems(4) = item.SubItems(3)
                    rItem.SubItems(5) = FormatNumber(Val(Replace(item.SubItems(4), ",", "")) - Val(item.SubItems(12)), 2, vbTrue, vbFalse)
                    rItem.SubItems(6) = item.SubItems(5)
                    rItem.SubItems(7) = cmbLocation.Text
                    rItem.SubItems(8) = Format(Now, "MM/DD/YY")
                    rItem.SubItems(9) = txtDeliveryNumber.Text
                    rItem.SubItems(10) = cmbLocation.ItemData(cmbLocation.ListIndex)
                    rItem.SubItems(14) = item.SubItems(13)
                    rItem.SubItems(15) = item.SubItems(14)
                    rItem.SubItems(16) = item.Text
                    rItem.Selected = False
            End If
        Next
    End With
'    txtItemSearch.SetFocus
'    lvItemList.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            If lvItemList.Visible = True Then
                lvItemList.Visible = False
                lvItemList.Left = -9999
                txtItemSearch.SetFocus
                'txtCode.SetFocus
            End If
        Case vbKeyF4
            txtItemSearch.SetFocus
            'txtCode.SetFocus
        Case vbKeyS
            If Shift = vbCtrlMask Then
                tb_Standard_ButtonClick tb_Standard.Buttons(2)
            End If
        Case vbKeyP
            If Shift = vbCtrlMask Then
                tb_Standard_ButtonClick tb_Standard.Buttons(6)
            End If
    End Select
End Sub

Private Sub Form_Load()
    lvItems.ColumnHeaders(4).width = lvItems.width * 0.109
    lvItems.ColumnHeaders(5).width = lvItems.width * 0.32
    lvItems.ColumnHeaders(6).width = lvItems.width * 0.085
    lvItems.ColumnHeaders(7).width = lvItems.width * 0.092
    lvItems.ColumnHeaders(8).width = lvItems.width * 0.141
    lvItems.ColumnHeaders(9).width = lvItems.width * 0.113
    lvItems.ColumnHeaders(10).width = lvItems.width * 0.109

    lvItemList.ColumnHeaders(2).width = lvItemList.width * 0.2
    lvItemList.ColumnHeaders(3).width = lvItemList.width * 0.76
    
    Populate "Location"
    Populate "PickOrder"
    Populate "PickOrderLine"
    txtRemarks.Text = SO_SalesOrderFrm.txtRemarks.Text
    txtDeliveryNumber.Text = SO_SalesOrderFrm.txtReferenceNumber.Text
    CountTotalQty
    
    On Error Resume Next
    cmbLocation.Text = "Store"
    cmbLocation.ListIndex = 0
End Sub
Public Sub Populate(ByVal data As String)
    Dim item As MSComctlLib.ListItem
    Select Case data
        Case "Location"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Location")
            cmbLocation.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        cmbLocation.AddItem rec!Location
                        cmbLocation.ItemData(cmbLocation.NewIndex) = rec!LocationId
                    End If
                    rec.MoveNext
                Loop
            End If
            On Error Resume Next
            Dim x As Integer
            For x = 1 To cmbLocation.ListCount
                If cmbLocation.ItemData(x) = 2 Then 'warehouse
                    cmbLocation.ListIndex = x
                    Exit For
                End If
            Next
        Case "PickOrder"
            Set rec = New ADODB.Recordset
            Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandText = "SO_PickOrder_Get"
            cmd.CommandType = adCmdStoredProc
            cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SO_SalesOrderFrm.SalesOrderId)
            
            Set rec = cmd.Execute
            If Not rec.EOF Then
                PickOrderId = rec!PickOrderId
                txtRemarks.Text = rec!Remarks
            Else
                txtRemarks.Text = ""
                PickOrderId = 0
            End If
            con.Close
        Case "PickOrderLine"
            Set rec = New ADODB.Recordset
            Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            
            'Dim item As MSComctlLib.ListItem
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "SO_PickOrderLine_Get"
            cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SO_SalesOrderFrm.SalesOrderId)
            Set rec = cmd.Execute
            lvItems.ListItems.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    Set item = lvItems.ListItems.add(, , rec!PickOrderLineId)
                        item.SubItems(1) = rec!PickOrderId
                        item.SubItems(2) = rec!ProductId
                        item.SubItems(3) = rec!itemcode
                        item.SubItems(4) = rec!Name
                        item.SubItems(5) = FormatNumber(rec!quantity, 2, vbTrue, vbFalse)
                        item.SubItems(6) = rec!Uom
                        item.SubItems(7) = rec!Location
                        item.SubItems(8) = Format(rec!Date, "MM/DD/YY")
                        item.SubItems(9) = rec!Reference
                        item.SubItems(10) = rec!LocationId
                        item.SubItems(12) = rec!quantity
                        item.Selected = True
                    rec.MoveNext
                Loop
            End If
            con.Close
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim x As Variant
    If lvItemList.ListItems.Count Then
        x = MsgBox("Are you sure you want to close this window?", vbYesNo + vbQuestion)
        If x = vbYes Then
            Call SO_SalesOrderFrm.Populate("SalesOrderLineGet")
            Unload Me
        Else
            Cancel = 1
        End If
    End If
End Sub

Private Sub lvItemList_DblClick()
    Dim item As MSComctlLib.ListItem
    Dim isFound As Boolean
    For Each item In SO_SalesOrderFrm.lvItems.ListItems
        If item.SubItems(9) = lvItemList.SelectedItem.Text Then
            isFound = True
            Exit For
        End If
    Next
    If isFound = True Then
        SO_PickOrderDialogFrm.isModify = False
        SO_PickOrderDialogFrm.Show (1)
    Else
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(59)
        GLOBAL_MessageFrm.Show (1)
    End If
End Sub

Private Sub lvItemList_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            lvItemList_DblClick
    End Select
End Sub

Private Sub lvItems_DblClick()
'    If lvItems.ListItems.Count > 0 Then
'        If lvItems.SelectedItem.text <> "" Then Exit Sub
'        With SO_PickOrderDialogFrm
'            .txtQuantity.text = lvItems.SelectedItem.SubItems(5)
'            .dtDate.value = lvItems.SelectedItem.SubItems(8)
'            .txtReference.text = lvItems.SelectedItem.SubItems(9)
'            On Error Resume Next
'            .cmbLocation.text = lvItems.SelectedItem.SubItems(7)
'            .isModify = True
'            .Show (1)
'            SO_PickOrderFrm.ZOrder 0
'        End With
'    End If
End Sub

Private Sub lvItems_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
            If lvItems.SelectedItem.Text = "" Then
                'lvItems.ListItems.Remove (lvItems.SelectedItem.Index)
            End If
        Case vbKeyReturn
            lvItems_DblClick
    End Select
End Sub

Private Sub tb_Standard_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 2 'Save
            Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            
            con.ConnectionString = ConnString
            con.Open
            'con.BeginTrans
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "SO_PickOrder_Insert"
            cmd.Parameters.Append cmd.CreateParameter("@PickOrderId", adInteger, adParamInputOutput, , PickOrderId)
            cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SO_SalesOrderFrm.SalesOrderId)
            cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, txtRemarks.Text)
            cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
            cmd.Execute
            
            PickOrderId = cmd.Parameters("@PickOrderId")
            
            'AUDIT TRAIL
            SavePOSAuditTrail UserId, WorkstationId, "", "Picked order with ref #: " & txtOrderNumber.Text, "SALES"
            
            'save line
            Dim item As MSComctlLib.ListItem
            Dim PickOrderLineId As Integer
            Dim isFound As Boolean
            Dim Available As Double
            
            For Each item In lvItems.ListItems
                If item.Text = "" Then
                    Set cmd = New ADODB.Command
                    cmd.ActiveConnection = con
                    cmd.CommandType = adCmdStoredProc
                    cmd.CommandText = "SO_PickOrderLine_Insert"
                    cmd.Parameters.Append cmd.CreateParameter("@PickOrderLineId", adInteger, adParamInputOutput, , PickOrderLineId)
                    cmd.Parameters.Append cmd.CreateParameter("@PickOrderId", adInteger, adParamInput, , PickOrderId)
                    cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SO_SalesOrderFrm.SalesOrderId)
                    cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , item.SubItems(2))
                    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, item.SubItems(4))
                    cmd.Parameters.Append cmd.CreateParameter("@Uom", adVarChar, adParamInput, 50, item.SubItems(6))
                    cmd.Parameters.Append cmd.CreateParameter("@LocationId", adInteger, adParamInput, , item.SubItems(10))
                    cmd.Parameters.Append cmd.CreateParameter("@Quantity", adDecimal, adParamInput, , Val(Replace(item.SubItems(5), ",", "")))
                                          cmd.Parameters("@Quantity").Precision = 18
                                          cmd.Parameters("@Quantity").NumericScale = 2
                    cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , item.SubItems(8))
                    cmd.Parameters.Append cmd.CreateParameter("@Reference", adVarChar, adParamInput, 250, item.SubItems(9))
                    cmd.Parameters.Append cmd.CreateParameter("@ReserveId", adInteger, adParamInput, , Val(item.SubItems(15)))
                    cmd.Parameters.Append cmd.CreateParameter("@SalesOrderLineId", adInteger, adParamInput, , Val(item.SubItems(16)))
                    cmd.Execute
                    
                    item.SubItems(1) = PickOrderId
                    item.Text = cmd.Parameters("@PickOrderLineId")
                    item.Selected = True
                    isFound = True
                    
                    'UPDATE RESERVED
'                    Dim iQty As Double
'                    iQty = Val(Replace(item.SubItems(5), ",", "")) * Val(Replace(item.SubItems(14), ",", ""))
'                    UpdateReserveQuantity item.SubItems(15), iQty * -1, item.SubItems(2), SO_SalesOrderFrm.SalesOrderId
                End If
            Next
            
            If isFound = True Then 'update status of PO to inProgress
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "SO_SalesOrderStatus_Update"
                cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SO_SalesOrderFrm.SalesOrderId)
                cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , 5) 'Picking
                cmd.Execute
                
                With SO_SalesOrderFrm
                    .txtStatus.Text = "Picking"
                    .lvSearch.SelectedItem.SubItems(3) = "Picking"
                    .StatusId = 5 'Picking
                End With
                
            End If
                                     
            'con.CommitTrans
            con.Close
            
            SO_SalesOrderFrm.SalesOrderId = SO_SalesOrderFrm.lvSearch.SelectedItem.Text
            SO_SalesOrderFrm.Populate "SalesOrderLineGet"
            SO_SalesOrderFrm.isNotCompleted (False)
            
            CountTotalQty
            
            MsgBox "Order saved.", vbInformation
            LoadImageStatus SO_SalesOrderFrm.picStatus, GetStatus(SO_SalesOrderFrm.StatusId)
        Case 6 'PRINT PREVIEW
'            If PurchaseOrderId <> 0 Then
'                Screen.MousePointer = vbHourglass
'                BASE_PrintPreviewFrm.Show
'                Dim crxApp As New CRAXDRT.Application
'                Dim crxRpt As New CRAXDRT.Report
'                Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\PO_WarehouseEntryForm.rpt")
'                crxRpt.RecordSelectionFormula = "{PO_PurchaseOrder.PurchaseOrderId}= " & PurchaseOrderId & ""
'                crxRpt.DiscardSavedData
'
'                Call ResetRptDB(crxRpt)
'
'                BASE_PrintPreviewFrm.CRViewer.ReportSource = crxRpt
'                BASE_PrintPreviewFrm.CRViewer.ViewReport
'                BASE_PrintPreviewFrm.CRViewer.Zoom 1
'                Screen.MousePointer = vbDefault
'            End If
    End Select
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
            btnItemSearch_Click
        Case vbKeyUp, vbKeyDown
            If lvItemList.Visible = True Then
                lvItemList.SetFocus
            Else
                lvItems.SetFocus
            End If
    End Select
End Sub

Private Sub txtRemarks_Change()
    SO_SalesOrderFrm.global_remarks = txtRemarks.Text
End Sub
