VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form SO_OrderPenaltiesFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11550
   Icon            =   "SO_OrderPenaltiesFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   11550
   ShowInTaskbar   =   0   'False
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
      Left            =   9600
      TabIndex        =   9
      Top             =   2400
      Width           =   1695
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
      Left            =   1200
      TabIndex        =   1
      Top             =   1320
      Width           =   10095
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
      TabIndex        =   3
      Top             =   2040
      Width           =   10095
   End
   Begin VB.ComboBox cmbCustomer 
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
      Left            =   1200
      TabIndex        =   2
      Top             =   1680
      Width           =   10095
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   7080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_OrderPenaltiesFrm.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_OrderPenaltiesFrm.frx":686E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_OrderPenaltiesFrm.frx":D0D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_OrderPenaltiesFrm.frx":13932
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SO_OrderPenaltiesFrm.frx":1A194
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvOrders 
      Height          =   5055
      Left            =   240
      TabIndex        =   4
      Top             =   2880
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   8916
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
         Text            =   "Customer"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Due Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Days OD"
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
         Text            =   "InterestRate"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "CustomerId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Subtotal"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "View Transactions"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   9720
      MouseIcon       =   "SO_OrderPenaltiesFrm.frx":209F6
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label3 
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
      Left            =   240
      TabIndex        =   8
      Top             =   1320
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "You can issue penalties or count late payments for orders that are passed due dates."
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
      Left            =   240
      TabIndex        =   7
      Top             =   765
      Width           =   7455
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Penalty and Late Payments"
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
      Left            =   840
      TabIndex        =   6
      Top             =   285
      Width           =   3090
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "SO_OrderPenaltiesFrm.frx":20B48
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label2 
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
      TabIndex        =   5
      Top             =   2040
      Width           =   690
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
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   900
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   7935
      Left            =   120
      Top             =   120
      Width           =   11295
   End
End
Attribute VB_Name = "SO_OrderPenaltiesFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btnSearch_Click()
    Dim Code, Name As Variant
    Dim item As MSComctlLib.ListItem
    
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "FIN_AccountsReceivable_Search1"
    
    Code = txtCode.text
    Name = cmbcustomer.text
    If Trim(txtCode.text) = "" Then Code = Null
    If Trim(cmbcustomer.text) = "" Then Name = Null
    
    cmd.Parameters.Append cmd.CreateParameter("@CustomerCode", adVarChar, adParamInput, 50, Code)
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 500, Name)
    cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInput, 50, txtOrderNumber.text)
    cmd.Parameters.Append cmd.CreateParameter("@Sort", adVarChar, adParamInput, 50, "Date")
    cmd.Parameters.Append cmd.CreateParameter("@Order", adVarChar, adParamInput, 50, "ASC")
    
    Set rec = cmd.Execute
    lvOrders.ListItems.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            'If DateDiff("d", rec!duedate, Now) > 0 Then
                Set item = lvOrders.ListItems.add(, , rec!SalesOrderId)
                    item.SubItems(1) = rec!OrderNumber
                    item.SubItems(2) = rec!CustomerCode & "-" & rec!Name
                    item.SubItems(3) = Format(rec!Date, "MM/DD/YY")
                    item.SubItems(4) = Format(rec!DueDate, "MM/DD/YY")
                    If DateDiff("d", rec!DueDate, Now) < 0 Then
                        item.SubItems(5) = ""
                    Else
                        item.SubItems(5) = DateDiff("d", rec!DueDate, Now)
                    End If
                    item.SubItems(6) = FormatNumber(rec!total, 2, vbTrue, vbFalse)
                    item.SubItems(7) = FormatNumber(rec!balance, 2, vbTrue, vbFalse)
                    item.SubItems(8) = rec!interestrate
                    item.SubItems(9) = rec!CustomerId
                    item.SubItems(10) = FormatNumber(rec!subtotal, 2, vbTrue, vbFalse)
            'End If
            rec.MoveNext
        Loop
    End If
    
    con.Close
End Sub

Private Sub cmbCustomer_Change()
   btnSearch_Click
End Sub

Private Sub Form_Load()
    lvOrders.ColumnHeaders(2).width = lvOrders.width * 0.13
    lvOrders.ColumnHeaders(3).width = lvOrders.width * 0.2
    lvOrders.ColumnHeaders(4).width = lvOrders.width * 0.12
    lvOrders.ColumnHeaders(5).width = lvOrders.width * 0.12
    lvOrders.ColumnHeaders(6).width = lvOrders.width * 0.13
    lvOrders.ColumnHeaders(7).width = lvOrders.width * 0.1342
    lvOrders.ColumnHeaders(8).width = lvOrders.width * 0.1342
End Sub

Public Sub Populate(ByVal data As String)
'    Select Case data
'        Case "Status"
'            Set rec = New ADODB.Recordset
'            Set rec = Global_Data("Status")
'            cmbSearch_Status.Clear
'            cmbSearch_Status.AddItem ""
'            cmbSearch_Status.ItemData(cmbSearch_Status.NewIndex) = 0
'            cmbSearch_Status.ListIndex = 0
'            If Not rec.EOF Then
'                Do Until rec.EOF
'                    cmbSearch_Status.AddItem rec!Status
'                    cmbSearch_Status.ItemData(cmbSearch_Status.NewIndex) = rec!StatusId
'                    rec.MoveNext
'                Loop
'            End If
'        Case "Customer"
'            Set rec = New ADODB.Recordset
'            Set cmd = New ADODB.Command
'            Set con = New ADODB.Connection
'            con.ConnectionString = ConnString
'            con.Open
'            cmd.ActiveConnection = con
'            cmd.CommandType = adCmdStoredProc
'            cmd.CommandText = "BASE_Customer_Get"
'            cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInput, , CustomerId)
'            Set rec = cmd.Execute
'            If Not rec.EOF Then
'                Do Until rec.EOF
'                    lblBalance.Caption = FormatNumber(rec!balance, 2, vbTrue, vbFalse)
'                    rec.MoveNext
'                Loop
'            End If
'            con.Close
'    End Select
End Sub

Private Sub lvOrders_DblClick()
    If lvOrders.ListItems.Count > 0 Then
        With SO_PenaltyFrm
            .lblCustomer.Caption = lvOrders.SelectedItem.SubItems(2)
            .lblOrderNumber.Caption = lvOrders.SelectedItem.SubItems(1)
            .lblDate.Caption = lvOrders.SelectedItem.SubItems(3)
            .lblDueDate.Caption = lvOrders.SelectedItem.SubItems(4)
            .lblDaysOverdue.Caption = lvOrders.SelectedItem.SubItems(5)
            .lblTotal.Caption = lvOrders.SelectedItem.SubItems(6)
            .lblBalance.Caption = lvOrders.SelectedItem.SubItems(7)
            .lblSubTotal.Caption = lvOrders.SelectedItem.SubItems(10)
            .txtBalance.text = lvOrders.SelectedItem.SubItems(7)
            .txtDays.text = lvOrders.SelectedItem.SubItems(5)
            .txtInterestRate.text = lvOrders.SelectedItem.SubItems(8)
            .SalesOrderId = lvOrders.SelectedItem.text
            .CustomerId = lvOrders.SelectedItem.SubItems(9)
            .Show (1)
        End With
    End If
End Sub

Private Sub txtCode_Change()
    btnSearch_Click
End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            btnSearch_Click
    End Select
End Sub

Private Sub txtOrderNumber_Change()
    btnSearch_Click
End Sub

Private Sub txtOrderNumber_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            btnSearch_Click
    End Select
End Sub
