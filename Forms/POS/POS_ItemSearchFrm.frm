VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form POS_ItemSearchFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13170
   Icon            =   "POS_ItemSearchFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   13170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnDetails 
      Caption         =   "F1: Details"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   8040
      Picture         =   "POS_ItemSearchFrm.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "ESC: Cancel"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   11400
      Picture         =   "POS_ItemSearchFrm.frx":05E3
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton btnAccept 
      Caption         =   "ENTER: Accept"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   9720
      Picture         =   "POS_ItemSearchFrm.frx":2972
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7200
      Width           =   1575
   End
   Begin MSComctlLib.ListView lvItemSearch 
      Height          =   5295
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   9340
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   15478
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Price"
         Object.Width           =   15478
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Code"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "QTY on Hand"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "AllowNegativeInv"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "ProductId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Price"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Price1"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Price2"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Price3"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "UnitCost"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Percentage"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Unit"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.CommandButton btnBarcode 
      Caption         =   "F5"
      Height          =   375
      Left            =   12480
      TabIndex        =   3
      Top             =   1320
      Width           =   495
   End
   Begin VB.TextBox txtBarcode 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1320
      Width           =   11175
   End
   Begin VB.CommandButton btnName 
      Caption         =   "F4"
      Height          =   375
      Left            =   12480
      TabIndex        =   1
      Top             =   840
      Width           =   495
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   0
      Top             =   840
      Width           =   11175
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ITEM SEARCH"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   120
      Width           =   3135
   End
   Begin VB.Image picModuleImage 
      Height          =   480
      Left            =   120
      Picture         =   "POS_ItemSearchFrm.frx":4D46
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   240
      Picture         =   "POS_ItemSearchFrm.frx":5360
      Stretch         =   -1  'True
      Top             =   120
      Width           =   12765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   240
      TabIndex        =   6
      Top             =   1320
      Width           =   540
   End
   Begin VB.Label lblBarcode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   630
   End
End
Attribute VB_Name = "POS_ItemSearchFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ProductRec As New ADODB.Recordset
Private Sub btnAccept_Click()
    If lvItemSearch.ListItems.Count > 0 Then
        POS_QuantityFrm.isChangeQuantity = False
        POS_QuantityFrm.Show (1)
    End If
End Sub

Private Sub btnBarcode_Click()
    txtBarcode.SetFocus
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Public Sub btnReturn_Click()
    selectText txtSearch
End Sub

Private Sub btnDetails_Click()
    POS_ItemDetailsFrm.ProductId = lvItemSearch.SelectedItem.SubItems(5)
    POS_ItemDetailsFrm.Show (1)
End Sub

Private Sub btnName_Click()
    txtName.SetFocus
End Sub
Private Sub Form_Activate()
    On Error Resume Next
    txtName.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            btnDetails_Click
        Case 115 'F4 Return
            btnName_Click
        Case 116 'F5
            btnBarcode_Click
        Case vbKeyEscape 'ESC
            btnCancel_Click
        Case 13 'Enter
            btnAccept_Click
        Case vbKeyF10
            'show UOM Menu
            If lvItemSearch.ListItems.Count > 0 Then
                POS_UomFrm.ProductId = lvItemSearch.SelectedItem.SubItems(5)
                POS_UomFrm.Show (1)
            End If
    End Select
End Sub

Private Sub Form_Load()
    If PharmacyMode = "ON" Then
        lvItemSearch.ColumnHeaders(1).width = lvItemSearch.width * 0.63 'Name
        lvItemSearch.ColumnHeaders(2).width = (lvItemSearch.width * 0.17)   'Price
        lvItemSearch.ColumnHeaders(4).width = lvItemSearch.width * 0.17 'Quantity
    Else
        lvItemSearch.ColumnHeaders(1).width = lvItemSearch.width * 0.8 'Name
        lvItemSearch.ColumnHeaders(2).width = (lvItemSearch.width * 0.17)   'Price
        lvItemSearch.ColumnHeaders(4).width = lvItemSearch.width * 0 'Quantity
    End If
    
    'Set rec = ProductName("")
End Sub

Private Sub txtSearch_Change()
    
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If lvItemSearch.ListItems.Count > 0 Then
        If KeyCode = 40 Or KeyCode = vbKeyUp Then 'Down/Up Arrow
            lvItemSearch.SetFocus
        End If
    End If
End Sub

Private Sub lvItemSearch_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'        Case vbKeyReturn
'            btnAccept_Click
'    End Select
End Sub

Private Sub txtBarcode_Change()
    On Error Resume Next
    If Trim(txtBarcode.text) = "" Then Exit Sub
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    Dim item As MSComctlLib.ListItem
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "POS_ItemSearch_Code"
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, txtBarcode.text)
    Set rec = cmd.Execute
    lvItemSearch.ListItems.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            If rec!isActive = "True" Then
                Set item = lvItemSearch.ListItems.add(, , rec!Name)
                    item.SubItems(5) = rec!ProductId
                    'item.SubItems(3) = rec!Quantity
                    item.SubItems(2) = rec!itemcode
                    item.SubItems(6) = rec!unitprice
                    'On Error Resume Next
                    item.SubItems(7) = rec!price1
                    item.SubItems(8) = rec!price2
                    item.SubItems(9) = rec!price3
                    item.SubItems(10) = rec!unitcost
                    item.SubItems(11) = rec!Percentage
                    item.SubItems(12) = rec!Uom
                    
'                    If UCase(POS_CashierFrm.lblDiscount.Caption) = UCase("| DISCOUNT TYPE: NONE") Then
                        item.SubItems(1) = FormatNumber(rec!unitprice, 2, vbTrue)
'                    ElseIf UCase(POS_CashierFrm.lblDiscount.Caption) = UCase("| DISCOUNT TYPE: DP") Then
'                        item.SubItems(1) = FormatNumber(rec!price1, 2, vbTrue)
'                    ElseIf UCase(POS_CashierFrm.lblDiscount.Caption) = UCase("| DISCOUNT TYPE: MS") Then
'                        item.SubItems(1) = FormatNumber(rec!price2, 2, vbTrue)
'                    ElseIf UCase(POS_CashierFrm.lblDiscount.Caption) = UCase("| DISCOUNT TYPE: BC") Then
'                        item.SubItems(1) = FormatNumber(rec!price3, 2, vbTrue)
'                    End If
            End If
            rec.MoveNext
        Loop
    End If
    con.Close
End Sub

Private Sub txtBarcode_GotFocus()
    selectText txtBarcode
End Sub

Private Sub txtBarcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
'            Set con = New ADODB.Connection
'            Set rec = New ADODB.Recordset
'            Set cmd = New ADODB.Command
'            Dim item As MSComctlLib.ListItem
'
'            con.ConnectionString = ConnString
'            con.Open
'            cmd.ActiveConnection = con
'            cmd.CommandType = adCmdStoredProc
'            cmd.CommandText = "POS_ItemSearch"
'            cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, Null)
'            cmd.Parameters.Append cmd.CreateParameter("@Barcode", adVarChar, adParamInput, 50, Null)
'            cmd.Parameters.Append cmd.CreateParameter("@LocationId", adInteger, adParamInput, , POS_CashierFrm.POSLocationId)
'            cmd.Parameters.Append cmd.CreateParameter("@ItemCode", adVarChar, adParamInput, 50, txtBarcode.text)
'            Set rec = cmd.Execute
'            lvItemSearch.ListItems.Clear
'            If Not rec.EOF Then
'                Do Until rec.EOF
'                    If rec!isActive = "True" Then
'                        Set item = lvItemSearch.ListItems.add(, , rec!Name)
'                            item.SubItems(5) = rec!ProductId
'                            item.SubItems(3) = rec!quantity
'                            item.SubItems(2) = rec!itemcode
'                            item.SubItems(6) = rec!unitprice
'                            item.SubItems(7) = rec!price1
'                            item.SubItems(8) = rec!price2
'                            item.SubItems(9) = rec!price3
'                            item.SubItems(10) = rec!unitcost
'                            item.SubItems(11) = rec!Percentage
'
''                            If POS_CashierFrm.discount = UCase("NONE") Then
'                                item.SubItems(1) = FormatNumber(rec!unitprice, 2, vbTrue)
''                            ElseIf POS_CashierFrm.discount = UCase("Distributor's Price") Then
''                                item.SubItems(1) = FormatNumber(rec!price1, 2, vbTrue)
''                            ElseIf POS_CashierFrm.discount = UCase("Satellite Price") Then
''                                item.SubItems(1) = FormatNumber(rec!price2, 2, vbTrue)
''                            ElseIf POS_CashierFrm.discount = UCase("Business Center Price") Then
''                                item.SubItems(1) = FormatNumber(rec!price3, 2, vbTrue)
''                            End If
'                    End If
'                    rec.MoveNext
'                Loop
'            End If
'            con.Close
        Case vbKeyDown
            If lvItemSearch.ListItems.Count > 0 Then
                lvItemSearch.SetFocus
            End If
    End Select
End Sub

Private Sub txtName_Change()
    On Error GoTo ErrMessage
    If Trim(txtName.text) = "" Then Exit Sub
    'Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    'Set cmd = New ADODB.Command
    Dim item As MSComctlLib.ListItem

'    con.ConnectionString = ConnString
'    con.Open
'    cmd.ActiveConnection = con
'    cmd.CommandType = adCmdStoredProc
'    cmd.CommandText = "POS_ItemSearch_Name"
'    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, txtName.text)

    Set rec = ProductName(txtName.text)

    If Not rec.EOF Then
        lvItemSearch.ListItems.Clear
        Do Until rec.EOF
            If rec!isActive = "True" Then
                Set item = lvItemSearch.ListItems.add(, , rec!Name)
                    item.SubItems(5) = rec!ProductId
                    item.SubItems(3) = FormatNumber(rec!Quantity, 2, vbTrue, vbFalse)
                    item.SubItems(2) = rec!itemcode
                    item.SubItems(6) = rec!unitprice
                    'On Error Resume Next
                    item.SubItems(7) = rec!price1
                    item.SubItems(8) = rec!price2
                    item.SubItems(9) = rec!price3
                    item.SubItems(10) = rec!unitcost
                    item.SubItems(11) = rec!Percentage
                    item.SubItems(12) = rec!Uom

'                    If UCase(POS_CashierFrm.lblDiscount.Caption) = UCase("| DISCOUNT TYPE: NONE") Then
                        item.SubItems(1) = FormatNumber(rec!unitprice, 2, vbTrue)
'                    ElseIf UCase(POS_CashierFrm.lblDiscount.Caption) = UCase("| DISCOUNT TYPE: DP") Then
'                        item.SubItems(1) = FormatNumber(rec!price1, 2, vbTrue)
'                    ElseIf UCase(POS_CashierFrm.lblDiscount.Caption) = UCase("| DISCOUNT TYPE: MS") Then
'                        item.SubItems(1) = FormatNumber(rec!price2, 2, vbTrue)
'                    ElseIf UCase(POS_CashierFrm.lblDiscount.Caption) = UCase("| DISCOUNT TYPE: BC") Then
'                        item.SubItems(1) = FormatNumber(rec!price3, 2, vbTrue)
'                    End If
            End If
            rec.MoveNext
        Loop
    Else

    End If
    Exit Sub
ErrMessage:
    MsgBox "A problem occured while processing your request. " & Err.Description & " Please try again.", vbCritical
    SYS_ErrorLog UserId, WorkstationId, Err.Description
    'con.Close
End Sub

Private Sub txtName_GotFocus()
    selectText txtName
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
'        Case vbKeyReturn
'            If Trim(txtName.text) = "" Then Exit Sub
'            'Set con = New ADODB.Connection
'            Set rec = New ADODB.Recordset
'            'Set cmd = New ADODB.Command
'            Dim item As MSComctlLib.ListItem
'
'        '    con.ConnectionString = ConnString
'        '    con.Open
'        '    cmd.ActiveConnection = con
'        '    cmd.CommandType = adCmdStoredProc
'        '    cmd.CommandText = "POS_ItemSearch_Name"
'        '    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, txtName.text)
'
'            Set rec = ProductName(txtName.text)
'            lvItemSearch.ListItems.Clear
'            If Not rec.EOF Then
'
'                Do Until rec.EOF
'                    If rec!isActive = "True" Then
'                        Set item = lvItemSearch.ListItems.add(, , rec!Name)
'                            item.SubItems(5) = rec!ProductId
'                            'item.SubItems(3) = rec!Quantity
'                            item.SubItems(2) = rec!itemcode
'                            item.SubItems(6) = rec!unitprice
'                            'On Error Resume Next
'                            item.SubItems(7) = rec!price1
'                            item.SubItems(8) = rec!price2
'                            item.SubItems(9) = rec!price3
'                            item.SubItems(10) = rec!unitcost
'                            item.SubItems(11) = rec!Percentage
'                            item.SubItems(12) = rec!Uom
'
'        '                    If UCase(POS_CashierFrm.lblDiscount.Caption) = UCase("| DISCOUNT TYPE: NONE") Then
'                                item.SubItems(1) = FormatNumber(rec!unitprice, 2, vbTrue)
'        '                    ElseIf UCase(POS_CashierFrm.lblDiscount.Caption) = UCase("| DISCOUNT TYPE: DP") Then
'        '                        item.SubItems(1) = FormatNumber(rec!price1, 2, vbTrue)
'        '                    ElseIf UCase(POS_CashierFrm.lblDiscount.Caption) = UCase("| DISCOUNT TYPE: MS") Then
'        '                        item.SubItems(1) = FormatNumber(rec!price2, 2, vbTrue)
'        '                    ElseIf UCase(POS_CashierFrm.lblDiscount.Caption) = UCase("| DISCOUNT TYPE: BC") Then
'        '                        item.SubItems(1) = FormatNumber(rec!price3, 2, vbTrue)
'        '                    End If
'                    End If
'                    rec.MoveNext
'                Loop
'            Else
'
'            End If
        Case vbKeyDown
            If lvItemSearch.ListItems.Count > 0 Then
                lvItemSearch.SetFocus
            End If
    End Select
End Sub
