VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form PO_PurchaseReturnInvoiceFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Returns"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8775
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   8775
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbVendor 
      BackColor       =   &H00FFFFFF&
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
      Height          =   345
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   7335
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Add"
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
      Left            =   5760
      TabIndex        =   1
      Top             =   6480
      Width           =   1455
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
      Left            =   7320
      TabIndex        =   0
      Top             =   6480
      Width           =   1335
   End
   Begin MSComctlLib.ListView lvModules 
      Height          =   5775
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   10186
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Return #"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "SalesReturnId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "StatusId"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vendor"
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
      TabIndex        =   4
      Top             =   120
      Width           =   690
   End
End
Attribute VB_Name = "PO_PurchaseReturnInvoiceFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub btnSave_Click()
    getrefunds
    PO_PurchaseInvoiceFrm.CountTotal
    Me.Hide
End Sub
Private Sub getrefunds()
    Dim item As MSComctlLib.ListItem
    Dim refunds As Double
'
'    'SAVE DATA
'    Dim con As New ADODB.Connection
'    Set cmd = New ADODB.Command
'    Set rec = New ADODB.Recordset
'
'    con.ConnectionString = ConnString
'    con.Open
'
'    For Each item In lvModules.ListItems
'        Set cmd = New ADODB.Command
'        cmd.ActiveConnection = con
'        cmd.CommandText = "SO_InvoiceRefunds_Insert"
'        cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , SO_PurchaseOrderFrm.PurchaseOrderId)
'        cmd.Parameters.Append cmd.CreateParameter("@PurchaseReturnId", adInteger, adParamInput, , Val(lvModules.SelectedItem.SubItems(5)))
'        cmd.Execute
'    Next
'
'    con.Close
    
    For Each item In lvModules.ListItems
        If item.Checked = True Then
            refunds = refunds + NVAL(item.SubItems(3))
        End If
    Next
    refunds = refunds
    PO_PurchaseInvoiceFrm.txtRefunds.text = FormatNumber(refunds, 2, vbTrue, vbFalse)
End Sub
Private Sub cmbVendor_Change()
    SearchReturns
End Sub

Private Sub cmbVendor_Click()
    SearchReturns
End Sub

Private Sub Form_Load()
    '****** REGION Listview Columns *********
    
    lvModules.ColumnHeaders(1).width = lvModules.width * 0.03
    lvModules.ColumnHeaders(2).width = lvModules.width * 0.342
    lvModules.ColumnHeaders(3).width = lvModules.width * 0.192
    lvModules.ColumnHeaders(4).width = lvModules.width * 0.192
    lvModules.ColumnHeaders(5).width = lvModules.width * 0.192
    
    Populate "Vendor"
End Sub
Private Sub Form_Unload(Cancel As Integer)
'    Cancel = True
'    Me.Hide
End Sub

Private Sub SearchReturns()
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "PO_PurchaseReturn_Get"
    cmd.Parameters.Append cmd.CreateParameter("@PurchaseReturnId", adInteger, adParamInput, , Null)
    cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , 2)
    cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInput, 50, Null)
    cmd.Parameters.Append cmd.CreateParameter("@VendorId", adInteger, adParamInput, , cmbVendor.ItemData(cmbVendor.ListIndex))
    Dim item As MSComctlLib.ListItem
    Set rec = cmd.Execute
    lvModules.ListItems.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            'If rec!isCashAdvance = "False" Then
                    Set item = lvModules.ListItems.add(, , rec!PurchaseReturnId)
                        item.SubItems(1) = rec!OrderNumber
                        item.SubItems(2) = Format(rec!Date, "MM/DD/YY")
                        item.SubItems(3) = FormatNumber(rec!Total, 2, vbTrue, vbFalse)
                        item.SubItems(4) = rec!Status
                        item.SubItems(5) = rec!PurchaseReturnId
                        item.SubItems(6) = rec!StatusId
            'End If
            rec.MoveNext
        Loop
    End If
    con.Close
End Sub
Public Sub Populate(ByVal data As String)
    Dim item As MSComctlLib.ListItem
    Select Case data
        Case "Vendor"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Vendor")
            cmbVendor.Clear
            cmbVendor.AddItem ""
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        cmbVendor.AddItem rec!Name
                        cmbVendor.ItemData(cmbVendor.NewIndex) = rec!VendorId
                    End If
                    rec.MoveNext
                Loop
            End If
            cmbVendor.ListIndex = 0
    End Select
End Sub

Private Sub lvModules_DblClick()
    If lvModules.ListItems.Count > 0 Then
        BASE_StatusFrm.Show (1)
    End If
End Sub
