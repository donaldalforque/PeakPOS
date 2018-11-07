VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form PO_ReceiveOrderFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13215
   Icon            =   "PO_ReceiveOrderFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8700
   ScaleWidth      =   13215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvItemList 
      Height          =   3735
      Left            =   1200
      TabIndex        =   6
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
         Object.Width           =   0
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
            Picture         =   "PO_ReceiveOrderFrm.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PO_ReceiveOrderFrm.frx":686E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PO_ReceiveOrderFrm.frx":D0D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PO_ReceiveOrderFrm.frx":13932
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PO_ReceiveOrderFrm.frx":13BA7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "PO_ReceiveOrderFrm.frx":14218
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tb_Standard 
      Height          =   330
      Left            =   0
      TabIndex        =   5
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
   Begin VB.Frame Fre_Body 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   7815
      Left            =   240
      TabIndex        =   9
      Top             =   600
      Width           =   12735
      Begin VB.Frame Frame_Header2 
         BackColor       =   &H00FFFFFF&
         Height          =   1275
         Left            =   9120
         TabIndex        =   15
         Top             =   0
         Width           =   3525
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
            TabIndex        =   16
            Top             =   240
            Width           =   2055
         End
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
            TabIndex        =   0
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PO #"
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
            TabIndex        =   18
            Top             =   240
            Width           =   420
         End
         Begin VB.Label Label1 
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
            TabIndex        =   17
            Top             =   600
            Width           =   930
         End
      End
      Begin VB.Frame Frame_Body 
         BackColor       =   &H00FFFFFF&
         Height          =   4620
         Left            =   0
         TabIndex        =   11
         Top             =   1680
         Width           =   12735
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
            TabIndex        =   3
            Top             =   240
            Width           =   4935
         End
         Begin VB.CommandButton btnItemSearch 
            Height          =   330
            Left            =   5880
            Picture         =   "PO_ReceiveOrderFrm.frx":1AA7A
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   240
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
            Left            =   9600
            TabIndex        =   12
            Top             =   240
            Visible         =   0   'False
            Width           =   3015
         End
         Begin MSComctlLib.ListView lvItems 
            Height          =   3855
            Left            =   120
            TabIndex        =   7
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
            NumItems        =   15
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
               Text            =   "Cost"
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
            Left            =   240
            TabIndex        =   14
            Top             =   240
            Width           =   555
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
            TabIndex        =   13
            Top             =   240
            Visible         =   0   'False
            Width           =   480
         End
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
         TabIndex        =   1
         Top             =   1320
         Width           =   3405
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
         TabIndex        =   8
         Top             =   6360
         Width           =   5895
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
         TabIndex        =   2
         Top             =   1320
         Width           =   3015
      End
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
         TabIndex        =   10
         Top             =   7440
         Visible         =   0   'False
         Width           =   5895
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Receive Orders"
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
         TabIndex        =   27
         Top             =   240
         Width           =   1740
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "PO_ReceiveOrderFrm.frx":1AC9E
         Top             =   120
         Width           =   480
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
         TabIndex        =   26
         Top             =   6360
         Width           =   795
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
         TabIndex        =   25
         Top             =   6480
         Visible         =   0   'False
         Width           =   1860
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
         TabIndex        =   24
         Top             =   6480
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Received"
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
         Visible         =   0   'False
         Width           =   1440
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
         TabIndex        =   22
         Top             =   6765
         Visible         =   0   'False
         Width           =   1740
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
         MouseIcon       =   "PO_ReceiveOrderFrm.frx":1D07E
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   1440
         Visible         =   0   'False
         Width           =   1845
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
         TabIndex        =   20
         Top             =   1320
         Width           =   780
      End
      Begin VB.Label Label2 
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
         Left            =   0
         TabIndex        =   19
         Top             =   7440
         Visible         =   0   'False
         Width           =   1125
      End
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
Attribute VB_Name = "PO_ReceiveOrderFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public PurchaseOrderId As Integer
Dim ReceiveOrderId As Integer
Public Sub CountTotalQty()
    Dim totalQty As Double
    Dim totalReceived As Double
    Dim item As MSComctlLib.ListItem
    With PO_PurchaseOrderFrm
        For Each item In .lvItems.ListItems
            totalQty = totalQty + Val(Replace(item.SubItems(4), ",", ""))
            totalReceived = totalReceived + Val(Replace(item.SubItems(11), ",", ""))
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
                        item.SubItems(5) = rec!unitcost
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
    
    With PO_PurchaseOrderFrm
        For Each item In .lvItems.ListItems
            If Val(Replace(item.SubItems(4), ",", "")) - Val(item.SubItems(11)) > 0 Then
                Set rItem = lvItems.ListItems.add(, , "")
                    rItem.SubItems(2) = item.SubItems(9)
                    rItem.SubItems(3) = item.SubItems(2)
                    rItem.SubItems(4) = item.SubItems(3)
                    rItem.SubItems(5) = FormatNumber(Val(Replace(item.SubItems(4), ",", "")) - Val(item.SubItems(11)), 2, vbTrue, vbFalse)
                    rItem.SubItems(6) = item.SubItems(5)
                    rItem.SubItems(7) = cmbLocation.Text
                    rItem.SubItems(8) = Format(Now, "MM/DD/YY")
                    rItem.SubItems(9) = txtDeliveryNumber.Text
                    rItem.SubItems(10) = cmbLocation.ItemData(cmbLocation.ListIndex)
                    rItem.SubItems(13) = item.Text
                    rItem.SubItems(14) = item.SubItems(6) 'cost
                    rItem.Selected = False
            End If
        Next
    End With
    txtItemSearch.SetFocus
    lvItemList.Visible = False
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
    Populate "ReceivedOrder"
    Populate "ReceivedOrderLine"
    
    CountTotalQty
    
    On Error Resume Next
    'cmbLocation.text = "Store"
    cmbLocation.ListIndex = 0
    txtRemarks.Text = PO_PurchaseOrderFrm.txtRemarks.Text
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
        Case "ReceivedOrder"
            Set rec = New ADODB.Recordset
            Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandText = "PO_ReceiveOrder_Get"
            cmd.CommandType = adCmdStoredProc
            cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , PO_PurchaseOrderFrm.PurchaseOrderId)
            
            Set rec = cmd.Execute
            If Not rec.EOF Then
                ReceiveOrderId = rec!ReceiveOrderId
                txtRemarks.Text = rec!Remarks
            Else
                txtRemarks.Text = ""
                ReceiveOrderId = 0
            End If
            con.Close
        Case "ReceivedOrderLine"
            Set rec = New ADODB.Recordset
            Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            
            'Dim item As MSComctlLib.ListItem
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "PO_ReceiveOrderLine_Get"
            cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , PO_PurchaseOrderFrm.PurchaseOrderId)
            Set rec = cmd.Execute
            lvItems.ListItems.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    Set item = lvItems.ListItems.add(, , rec!ReceiveOrderLineId)
                        item.SubItems(1) = rec!ReceiveOrderId
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
'    Dim x As Variant
'    If lvItemList.ListItems.Count Then
'        x = MsgBox("Are you sure you want to close this window?", vbYesNo + vbQuestion)
'        If x = vbYes Then
            Call PO_PurchaseOrderFrm.Populate("PurchaseOrderLineGet")
            Unload PO_ReceiveOrderDialogFrm
            Unload Me
'        Else
'            Cancel = 1
'        End If
'    End If
End Sub

Private Sub lvItemList_DblClick()
    Dim item As MSComctlLib.ListItem
    Dim isFound As Boolean
    For Each item In PO_PurchaseOrderFrm.lvItems.ListItems
        If item.SubItems(9) = lvItemList.SelectedItem.Text Then
            isFound = True
            Exit For
        End If
    Next
'    If isFound = True Then
        isModify = False
        PO_ReceiveOrderDialogFrm.txtDescription.Text = lvItemList.SelectedItem.SubItems(2)
        PO_ReceiveOrderDialogFrm.Show '(1)
'    Else
'        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(59)
'        GLOBAL_MessageFrm.Show (1)
'    End If
End Sub

Private Sub lvItemList_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            lvItemList_DblClick
    End Select
End Sub

Private Sub lvItems_DblClick()
    If lvItems.ListItems.Count > 0 Then
        If lvItems.SelectedItem.Text <> "" Then Exit Sub
        isModify = True
        With PO_ReceiveOrderDialogFrm
            .txtQuantity.Text = lvItems.SelectedItem.SubItems(5)
            .txtDescription.Text = lvItems.SelectedItem.SubItems(4)
            .dtDate.value = lvItems.SelectedItem.SubItems(8)
            .txtReference.Text = lvItems.SelectedItem.SubItems(9)
            .cmbLocation.Text = lvItems.SelectedItem.SubItems(7)
            .cmbUnit.Text = lvItems.SelectedItem.SubItems(6)
            .Show
        End With
    End If
End Sub

Private Sub lvItems_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
            If lvItems.SelectedItem.Text = "" Then
                lvItems.ListItems.Remove (lvItems.SelectedItem.Index)
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
            con.BeginTrans
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            
            cmd.Parameters.Append cmd.CreateParameter("@ReceiveOrderId", adInteger, adParamInputOutput, , ReceiveOrderId)
            cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , PO_PurchaseOrderFrm.PurchaseOrderId)
            cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, txtRemarks.Text)
            cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
            
            If Val(ReceiveOrderId) = 0 Then
                cmd.CommandText = "PO_ReceiveOrder_Insert"
                cmd.Execute
            Else
                cmd.CommandText = "PO_ReceiveOrder_Update"
                cmd.Execute
            End If
            
            ReceiveOrderId = cmd.Parameters("@ReceiveOrderId")
            
            'AUDIT TRAIL
            SavePOSAuditTrail UserId, WorkstationId, "", "Received order with ref #: " & txtOrderNumber.Text, "PURCHASING"
            
            'save line
            Dim item As MSComctlLib.ListItem
            Dim ReceiveOrderLineId As Integer
            Dim isFound As Boolean
            
            For Each item In lvItems.ListItems
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                If item.Text = "" Then
                    cmd.CommandText = "PO_ReceiveOrderLine_Insert"
                    cmd.Parameters.Append cmd.CreateParameter("@ReceiveOrderLineId", adInteger, adParamInputOutput, , ReceiveOrderLineId)
                    cmd.Parameters.Append cmd.CreateParameter("@ReceiveOrderId", adInteger, adParamInput, , ReceiveOrderId)
                    cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , PO_PurchaseOrderFrm.PurchaseOrderId)
                    cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , item.SubItems(2))
                    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, item.SubItems(4))
                    cmd.Parameters.Append cmd.CreateParameter("@Uom", adVarChar, adParamInput, 250, item.SubItems(6))
                    cmd.Parameters.Append cmd.CreateParameter("@ReceivedCost", adDecimal, adParamInput, , item.SubItems(14))
                                          cmd.Parameters("@ReceivedCost").Precision = 18
                                          cmd.Parameters("@ReceivedCost").NumericScale = 2
                    cmd.Parameters.Append cmd.CreateParameter("@LocationId", adInteger, adParamInput, , item.SubItems(10))
                    cmd.Parameters.Append cmd.CreateParameter("@Quantity", adDecimal, adParamInput, , item.SubItems(5))
                                          cmd.Parameters("@Quantity").Precision = 18
                                          cmd.Parameters("@Quantity").NumericScale = 2
                    cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , item.SubItems(8))
                    cmd.Parameters.Append cmd.CreateParameter("@Reference", adVarChar, adParamInput, 250, item.SubItems(9))
                    cmd.Execute
                    
                    item.SubItems(1) = ReceiveOrderId
                    item.Text = cmd.Parameters("@ReceiveOrderLineId")
                    item.Selected = True
                    isFound = True
                End If
            Next
            
            If isFound = True Then 'update status of PO to inProgress
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "PO_PurchaseOrderStatus_Update"
                cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , PO_PurchaseOrderFrm.PurchaseOrderId)
                cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , 3) 'IN PROGRESS
                cmd.Execute
            End If
                                     
            con.CommitTrans
            con.Close
            
            PO_PurchaseOrderFrm.PurchaseOrderId = PO_PurchaseOrderFrm.lvSearch.SelectedItem.Text
            PO_PurchaseOrderFrm.lvSearch.SelectedItem.SubItems(2) = "In progress"
            PO_PurchaseOrderFrm.txtStatus.Text = "In progress"
            PO_PurchaseOrderFrm.StatusId = 3
            PO_PurchaseOrderFrm.isNotCompleted (False)
            
            'PO_PurchaseOrderFrm.Populate "PurchaseOrderLineGet"
            
            MsgBox "Order saved.", vbInformation
            
            LoadImageStatus PO_PurchaseOrderFrm.picStatus, GetStatus(PO_PurchaseOrderFrm.StatusId)
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
    PO_PurchaseOrderFrm.global_remarks = txtRemarks.Text
End Sub
