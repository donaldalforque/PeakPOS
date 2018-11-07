VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form INV_ProductConversionFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13575
   Icon            =   "INV_ProductConversionFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   13575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnSearchCode 
      Height          =   330
      Left            =   6720
      Picture         =   "INV_ProductConversionFrm.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
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
      Left            =   1560
      TabIndex        =   1
      Top             =   1440
      Width           =   5175
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
      Left            =   1560
      TabIndex        =   3
      Top             =   1800
      Width           =   5175
   End
   Begin VB.CommandButton btnItemSearch 
      Height          =   330
      Left            =   6720
      Picture         =   "INV_ProductConversionFrm.frx":0230
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton btnSearch 
      Caption         =   "Add Products"
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
      Left            =   10800
      TabIndex        =   5
      Top             =   1800
      Width           =   2415
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
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
      Left            =   12120
      TabIndex        =   8
      Top             =   8880
      Width           =   1335
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Save && Close"
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
      Left            =   10680
      TabIndex        =   7
      Top             =   8880
      Width           =   1335
   End
   Begin MSComctlLib.ListView lvSearch 
      Height          =   6255
      Left            =   360
      TabIndex        =   6
      Top             =   2280
      Width           =   12855
      _ExtentX        =   22675
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
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "FromProductId"
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
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Unit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Itemcode"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "To Qty"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Unit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "ToProductId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "ProductConversionId"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Itemcode"
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
      TabIndex        =   11
      Top             =   1440
      Width           =   885
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
      Left            =   360
      TabIndex        =   10
      Top             =   1800
      Width           =   555
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "INV_ProductConversionFrm.frx":0454
      Top             =   360
      Width           =   480
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product Conversion"
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
      Left            =   960
      TabIndex        =   9
      Top             =   405
      Width           =   2280
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "This tool allows you to setup and configure how you sell and store your products from wholesale to retail and vice versa."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   450
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   10215
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   8655
      Left            =   120
      Top             =   120
      Width           =   13335
   End
End
Attribute VB_Name = "INV_ProductConversionFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnItemSearch_Click()
    Dim item As MSComctlLib.ListItem
    For Each item In lvSearch.ListItems
        If UCase(item.SubItems(2)) Like UCase(txtItemSearch.text) & "*" Then
            item.Selected = True
            item.EnsureVisible
            'lvItemList.SetFocus
            Exit For
        End If
    Next
End Sub

Private Sub btnSave_Click()
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    
    Dim item As MSComctlLib.ListItem
    Dim ConversionId As Integer
    
    con.ConnectionString = ConnString
    con.Open
    
    For Each item In lvSearch.ListItems
        If item.SubItems(9) <> "" Then
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            
            cmd.Parameters.Append cmd.CreateParameter("@ConversionId", adInteger, adParamInput, , Val(item.SubItems(10)))
            cmd.Parameters.Append cmd.CreateParameter("@FromProductId", adInteger, adParamInput, , item.text)
            cmd.Parameters.Append cmd.CreateParameter("@FromQty", adDecimal, adParamInput, , Val(Replace(item.SubItems(3), ",", "")))
                                  cmd.Parameters("@FromQty").NumericScale = 2
                                  cmd.Parameters("@FromQty").Precision = 18
            cmd.Parameters.Append cmd.CreateParameter("@ToProductId", adInteger, adParamInput, , Val(item.SubItems(9)))
            cmd.Parameters.Append cmd.CreateParameter("@ToQty", adDecimal, adParamInput, , Val(Replace(item.SubItems(7), ",", "")))
                                  cmd.Parameters("@ToQty").NumericScale = 2
                                  cmd.Parameters("@ToQty").Precision = 18
            If item.SubItems(10) = "" Then
                cmd.CommandText = "INV_ProductConversion_Insert"
            Else
                cmd.CommandText = "INV_ProductConversion_Update"
            End If
            cmd.Execute
        End If
    Next
    con.Close
    MsgBox "Product conversion details saved.", vbInformation
    Unload Me
End Sub

Private Sub btnSearch_Click()
    INV_AddProductFrm.isPriceMarkUp = False
    INV_AddProductFrm.Show (1)
End Sub

Private Sub btnSearchCode_Click()
    Dim item As MSComctlLib.ListItem
    For Each item In lvSearch.ListItems
        If UCase(item.SubItems(1)) Like UCase(txtCode.text) & "*" Then
            item.Selected = True
            item.EnsureVisible
            'lvItemList.SetFocus
            Exit For
        End If
    Next
End Sub

Private Sub Form_Load()
    lvSearch.ColumnHeaders(2).width = lvSearch.width * 0.1
    lvSearch.ColumnHeaders(3).width = lvSearch.width * 0.206
    lvSearch.ColumnHeaders(4).width = lvSearch.width * 0.09
    lvSearch.ColumnHeaders(5).width = lvSearch.width * 0.09
    lvSearch.ColumnHeaders(6).width = lvSearch.width * 0.1
    lvSearch.ColumnHeaders(7).width = lvSearch.width * 0.206
    lvSearch.ColumnHeaders(8).width = lvSearch.width * 0.09
    lvSearch.ColumnHeaders(9).width = lvSearch.width * 0.09
    
    Set con = New ADODB.Connection
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "INV_ProductConversion_Get"
    Set rec = cmd.Execute
    Dim item As MSComctlLib.ListItem
    If Not rec.EOF Then
        Do Until rec.EOF
            Set item = lvSearch.ListItems.add(, , rec!fromProductId)
                item.SubItems(1) = rec!fromItemCode
                item.SubItems(2) = rec!fromName
                item.SubItems(4) = rec!fromUom
                item.SubItems(3) = FormatNumber(rec!fromqty, 2, vbTrue, vbFalse)
                item.SubItems(5) = rec!toItemCode
                item.SubItems(6) = rec!toName
                item.SubItems(7) = FormatNumber(rec!toqty, 2, vbTrue, vbFalse)
                item.SubItems(8) = rec!toUom
                item.SubItems(9) = rec!toProductId
                item.SubItems(10) = rec!ConversionId
            rec.MoveNext
        Loop
    End If
    
    con.Close
End Sub

Private Sub lvSearch_DblClick()
    If lvSearch.ListItems.Count = 0 Then Exit Sub
    If lvSearch.SelectedItem.SubItems(10) = "" Then
        'INV_ProductConversionDialogFrm.isModify = False
        INV_ProductConversionDialogFrm.Show (1)
    Else
        With INV_ProductConversionDialogFrm
            .lblSelectedProduct.Caption = lvSearch.SelectedItem.SubItems(6)
            .lblUom.Caption = lvSearch.SelectedItem.SubItems(8)
            .txtQuantity.text = lvSearch.SelectedItem.SubItems(7)
            .Show (1)
        End With
    End If
End Sub

Private Sub lvSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If lvSearch.ListItems.Count > 0 Then
                lvSearch_DblClick
            End If
        Case vbKeyDelete
            If lvSearch.SelectedItem.SubItems(10) = "" Then
                lvSearch.ListItems.Remove (lvSearch.SelectedItem.Index)
            End If
    End Select
End Sub

Private Sub txtCode_Change()
    btnSearchCode_Click
End Sub

Private Sub txtCode_GotFocus()
    selectText txtCode
End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            If lvSearch.ListItems.Count > 0 Then lvSearch.SetFocus
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
        Case vbKeyDown
            If lvSearch.ListItems.Count > 0 Then lvSearch.SetFocus
    End Select
End Sub
