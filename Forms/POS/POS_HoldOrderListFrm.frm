VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form POS_HoldListFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hold List"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8640
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   8640
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnPrint 
      Caption         =   "F1:Print History"
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
      Left            =   1800
      Picture         =   "POS_HoldOrderListFrm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6000
      Width           =   1935
   End
   Begin VB.CommandButton btnDelete 
      Caption         =   "DEL: Delete"
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
      Left            =   120
      Picture         =   "POS_HoldOrderListFrm.frx":2228
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6000
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
      Left            =   6960
      Picture         =   "POS_HoldOrderListFrm.frx":286D
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6000
      Width           =   1575
   End
   Begin VB.CommandButton btnAccept 
      Caption         =   "ENTER: Select"
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
      Left            =   5280
      Picture         =   "POS_HoldOrderListFrm.frx":4BFC
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6000
      Width           =   1575
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   5595
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   9869
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
         Name            =   "Segoe UI"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "POSOrderId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name/Ref #"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Amount"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   5835
      Left            =   120
      Top             =   120
      Width           =   8415
   End
End
Attribute VB_Name = "POS_HoldListFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAccept_Click()
    If lvList.ListItems.Count <= 0 Then Exit Sub
    If POS_CashierFrm.Visible = False Then Exit Sub
    'load order
    Dim x As Variant
    x = MsgBox("This will remove your current POS transaction and will load the selected order. Are you sure you want to continue?", vbQuestion + vbYesNo)
    If x = vbYes Then
        'On Error GoTo errhandler
        'load order
        Set con = New ADODB.Connection
        Set cmd = New ADODB.Command
        Set rec = New ADODB.Recordset
        Dim item As MSComctlLib.ListItem
        con.ConnectionString = ConnString
        con.Open
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "POS_OrderLine_Get"
        cmd.Parameters.Append cmd.CreateParameter("@POS_OrderId", adInteger, adParamInput, , Val(lvList.SelectedItem.Text))
        Set rec = cmd.Execute
        If Not rec.EOF Then
            'clear list
            POS_CashierFrm.lvList.ListItems.Clear
            Do Until rec.EOF
                Set item = POS_CashierFrm.lvList.ListItems.add(, , rec!Name)
                    item.SubItems(1) = FormatNumber(rec!quantity, 2, vbTrue, vbFalse)
                    item.SubItems(2) = rec!unit
                    item.SubItems(3) = FormatNumber(rec!price, 2, vbTrue, vbFalse)
                    item.SubItems(4) = FormatNumber(rec!discount, 2, vbTrue, vbFalse)
                    item.SubItems(5) = rec!price
                    item.SubItems(6) = rec!unitcost
                    item.SubItems(7) = 0
                    item.SubItems(8) = rec!ProductId
                    item.SubItems(9) = rec!price
                    item.SubItems(10) = 0
                    item.SubItems(11) = 0
                    item.SubItems(12) = 0
                    item.SubItems(13) = rec!Percentage
                    item.SubItems(14) = rec!tax
                    item.SubItems(16) = rec!ActualQuantity
                rec.MoveNext
            Loop
        End If
        con.Close
        
        POS_CashierFrm.CountTotal
        POS_CashierFrm.CountTax
        
        POS_CashierFrm.POSOrderId = lvList.SelectedItem.Text
        POS_CashierFrm.POSHoldOrderReference = lvList.SelectedItem.SubItems(1)
        
        Unload Me
        Unload POS_OptionsFrm
    Else
    End If
    Exit Sub
'errhandler:
'    MsgBox "An error occured while loading order. Please try again.", vbCritical, "Error loading.."
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    
End Sub

Private Sub btnDelete_Click()
    If lvList.ListItems.Count <= 0 Then Exit Sub
    
    Dim x As Variant
    x = MsgBox("This will delete the order. Proceed?", vbQuestion + vbYesNo)
    If x = vbYes Then
        Set con = New ADODB.Connection
        Set rec = New ADODB.Recordset
        Set cmd = New ADODB.Command
        
        Dim item As MSComctlLib.ListItem
        
        con.ConnectionString = ConnString
        con.Open
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "POS_OrderStatus_Update"
        cmd.Parameters.Append cmd.CreateParameter("@POS_OrderId", adInteger, adParamInput, , NVAL(lvList.SelectedItem.Text))
        cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, "MANUAL")
        Set rec = cmd.Execute
        con.Close
    End If
    
    SavePOSAuditTrail UserId, WorkstationId, "", "Removed order ref: " & lvList.SelectedItem.SubItems(1)
    
    lvList.ListItems.Remove (lvList.SelectedItem.Index)
End Sub

Private Sub btnPrint_Click()
    POS_PrintHistoryHoldOrderListFrm.Show (1)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            btnAccept_Click
        Case vbKeyDelete
            btnDelete_Click
        Case vbKeyEscape
            btnCancel_Click
        Case vbKeyF1
            btnPrint_Click
    End Select
End Sub

Private Sub Form_Load()
    lvList.ColumnHeaders(2).width = lvList.width * 0.48
    lvList.ColumnHeaders(3).width = lvList.width * 0.48

    
    Populate
End Sub

Private Sub Populate()
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    Dim item As MSComctlLib.ListItem
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "POS_Order_Get"
    cmd.Parameters.Append cmd.CreateParameter("@POS_OrderId", adInteger, adParamInput, , Null)
    cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , Null)
    cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
    Set rec = cmd.Execute
    lvList.ListItems.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            Set item = lvList.ListItems.add(, , rec!POS_OrderId)
                item.SubItems(1) = rec!ReferenceNumber
                item.SubItems(2) = FormatNumber(rec!Total, 2, vbTrue, vbFalse)
            rec.MoveNext
        Loop
    End If
    con.Close
End Sub

