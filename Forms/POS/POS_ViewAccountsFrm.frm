VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form POS_ViewAccountsFrm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9525
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10590
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9525
   ScaleWidth      =   10590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnPrintByCustomer 
      Caption         =   "F2: Print by Customer"
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
      Left            =   2400
      Picture         =   "POS_ViewAccountsFrm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8520
      Width           =   2175
   End
   Begin VB.TextBox txtSearch_name 
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
      Left            =   1560
      TabIndex        =   0
      Top             =   1800
      Width           =   8655
   End
   Begin VB.TextBox txtSearch_Code 
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
      Left            =   1560
      TabIndex        =   1
      Top             =   1320
      Width           =   8655
   End
   Begin VB.CommandButton btnSummary 
      Caption         =   "F1: Print Summary"
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
      Picture         =   "POS_ViewAccountsFrm.frx":0625
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8520
      Width           =   2175
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
      Height          =   495
      Left            =   8520
      TabIndex        =   3
      Top             =   2760
      Width           =   1695
   End
   Begin VB.ComboBox cmbCity 
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
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2280
      Width           =   8655
   End
   Begin VB.CommandButton btnAccept 
      Caption         =   "F3: Pay Accounts"
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
      Left            =   4680
      Picture         =   "POS_ViewAccountsFrm.frx":0C4A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8520
      Width           =   1935
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
      Left            =   8880
      Picture         =   "POS_ViewAccountsFrm.frx":126A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8520
      Width           =   1575
   End
   Begin MSComctlLib.ListView lvSearch 
      Height          =   4335
      Left            =   360
      TabIndex        =   4
      Top             =   3360
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   7646
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
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "CustomerId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Code"
         Object.Width           =   0
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
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Left            =   480
      TabIndex        =   13
      Top             =   1800
      Width           =   570
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL: 0.00"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5640
      TabIndex        =   12
      Top             =   7800
      Width           =   4380
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "City"
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
      Left            =   480
      TabIndex        =   11
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
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
      Left            =   480
      TabIndex        =   10
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SEARCH:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   915
   End
   Begin VB.Label lblCaption_Title 
      AutoSize        =   -1  'True
      Caption         =   "POS Accounts Receivable"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   6
      Top             =   120
      Width           =   3180
   End
   Begin VB.Image picModuleImage 
      Height          =   480
      Left            =   120
      Picture         =   "POS_ViewAccountsFrm.frx":35F9
      Top             =   120
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   7695
      Left            =   120
      Top             =   720
      Width           =   10335
   End
End
Attribute VB_Name = "POS_ViewAccountsFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub Populate(ByVal data As String)
    Set rec = New ADODB.Recordset
    Dim item As MSComctlLib.ListItem
    Select Case data
        
        Case "City"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("City")
            cmbCity.Clear
            cmbCity.AddItem ""
            cmbCity.ItemData(cmbCity.NewIndex) = 0
            cmbCity.ListIndex = 0
            If Not rec.EOF Then
                Do Until rec.EOF
                    cmbCity.AddItem rec!City
                    cmbCity.ItemData(cmbCity.NewIndex) = rec!CityId
                    rec.MoveNext
                Loop
            End If
            cmbCity.ListIndex = 0
        End Select
End Sub

Private Sub CountTotal()
    Dim item As MSComctlLib.ListItem
    Dim Total As Double
    Total = 0
    For Each item In lvSearch.ListItems
        Total = Total + Val(Replace(item.SubItems(3), ",", ""))
    Next
    lbltotal.Caption = "TOTAL: " & FormatNumber(Total, 2, vbTrue, vbFalse)
End Sub

Private Sub btnAccept_Click()
    lvSearch_DblClick
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Function isValidated() As Boolean
    Dim hasPayment As Boolean
    Dim item As MSComctlLib.ListItem
    
    For Each item In lvSearch.ListItems
        If item.SubItems(10) <> "" Then
            If Val(Replace(item.SubItems(10), ",", "")) > 0 Then
                hasPayment = True
                Exit For
            End If
        End If
    Next
    If hasPayment = True Then
        isValidated = True
    Else
        isValidated = False
    End If
End Function

Private Sub btnSave_Click()
'    On Error GoTo ErrorHandler
'    If isValidated = True Then
'        Dim Item As MSComctlLib.ListItem
'
'        Set con = New ADODB.Connection
'        Set rec = New ADODB.Recordset
'
'        con.ConnectionString = ConnString
'        con.Open
'        con.BeginTrans
'
'        For Each Item In lvSearch.ListItems
'            If Val(Item.SubItems(10)) > 0 Then
'                'SAVE PAYMENT
'                Set cmd = New ADODB.Command
'                cmd.ActiveConnection = con
'                cmd.CommandType = adCmdStoredProc
'                cmd.CommandText = "SO_Payment_Insert"
'
'                cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , Item.text)
'                cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Val(Replace(Item.SubItems(10), ",", "")))
'                                      cmd.Parameters("@Amount").NumericScale = 2
'                                      cmd.Parameters("@Amount").Precision = 18
'                cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , Item.SubItems(13))
'                cmd.Parameters.Append cmd.CreateParameter("@PaymentType", adVarChar, adParamInput, 250, Item.SubItems(11))
'                cmd.Parameters.Append cmd.CreateParameter("@ChequeNumber", adVarChar, adParamInput, 250, Item.SubItems(14))
'                If Item.SubItems(15) <> "" Then
'                    cmd.Parameters.Append cmd.CreateParameter("@ChequeDate", adDate, adParamInput, , Item.SubItems(15))
'                Else
'                    cmd.Parameters.Append cmd.CreateParameter("@ChequeDate", adDate, adParamInput, , Null)
'                End If
'                If Item.SubItems(16) = "BANK" Then
'                    cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , Item.SubItems(18))
'                    cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , Null)
'                Else
'                    cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , Null)
'                    cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , Item.SubItems(19))
'                End If
'                cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 500, Item.SubItems(12))
'                cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInput, 250, Item.SubItems(3))
'                If Trim(Item.SubItems(18)) <> "" Then
'                    cmd.Parameters.Append cmd.CreateParameter("@RefAccountId", adInteger, adParamInput, , Item.SubItems(18))
'                End If
'                cmd.Execute
'            End If
'        Next
'
'        con.CommitTrans
'        con.Close
'
'        btnSearch_Click
'        MsgBox MessageCodes(3) & " " & MessageCodes(0), vbInformation, ""
'    Else
'        GLOBAL_MessageFrm.lblErrorMessage = ErrorCodes(0) & " " & ErrorCodes(17)
'        GLOBAL_MessageFrm.Show (1)
'    End If
'    Exit Sub
'ErrorHandler:
'    con.RollbackTrans
'    con.Close
'    If IsNumeric(Err.Description) = True Then
'        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(Val(Err.Description))
'    Else
'        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & Err.Description
'    End If
'    GLOBAL_MessageFrm.Show (1)
'
End Sub

Private Sub btnPrintByCustomer_Click()
    Dim item As MSComctlLib.ListItem
    Dim x As Variant
    x = MsgBox("This will print receivable by customer. Continue?", vbCritical + vbYesNo)
    If x = vbNo Then Exit Sub
    For Each item In lvSearch.ListItems
        Dim crxApp As New CRAXDRT.Application
        Dim crxRpt As New CRAXDRT.Report
        
        Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\POS_AccountsReceivableSummary.rpt")
        crxRpt.DiscardSavedData
        crxRpt.EnableParameterPrompting = False
        crxRpt.ParameterFields.GetItemByName("ReportTitle").AddCurrentValue "Accounts Receivable Summary"
        crxRpt.ParameterFields.GetItemByName("@CustomerId").AddCurrentValue Val(item.Text)
        crxRpt.ParameterFields.GetItemByName("@UserId").AddCurrentValue UserId
        crxRpt.ParameterFields.GetItemByName("@WorkStationId").AddCurrentValue WorkstationId
        Call ResetRptDB(crxRpt)
        crxRpt.EnableParameterPrompting = False
        crxRpt.PrintOut False
    Next
End Sub


Public Sub btnSearch_Click()
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "FIN_AccountsReceivable_Search"
    Dim Code, Name, Order, OrderNumber, Sort As Variant
    
    If Trim(txtSearch_Code.Text) = "" Then Code = Null Else Code = txtSearch_Code.Text
    If Trim(txtSearch_name.Text) = "" Then Name = Null Else Name = txtSearch_name.Text
    'If Trim(txtSearch_Order.text) = "" Then OrderNumber = Null Else OrderNumber = txtSearch_Order.text
'    If optCustomerName.value = True Then Sort = "Name"
'    'If optOrderNumber.value = True Then Sort = "Order"
'    'If optOrderDate.value = True Then Sort = "Date"
'    'If optDueDate.value = True Then Sort = "DueDate"
'    If optBalance.value = True Then Sort = "Balance"
'    If optAscending.value = True Then Order = "ASC"
'    If optDescending.value = True Then Order = "DESC"
    
    cmd.Parameters.Append cmd.CreateParameter("@CustomerCode", adVarChar, adParamInput, 50, Code)
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 50, Name)
    cmd.Parameters.Append cmd.CreateParameter("@City", adInteger, adParamInput, , cmbCity.ItemData(cmbCity.ListIndex))
    cmd.Parameters.Append cmd.CreateParameter("@Sort", adVarChar, adParamInput, 250, "Name")
    cmd.Parameters.Append cmd.CreateParameter("@Order", adVarChar, adParamInput, 50, "ASC")
    
    Dim item As MSComctlLib.ListItem
    Set rec = cmd.Execute
    lvSearch.ListItems.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            Set item = lvSearch.ListItems.add(, , rec!CustomerId)
                item.SubItems(1) = rec!CustomerCode
                item.SubItems(2) = rec!Name
                item.SubItems(3) = FormatNumber(rec!balance, 2, vbTrue, vbFalse)
            rec.MoveNext
        Loop
    End If
    con.Close
    CountTotal
End Sub

Private Sub btnSummary_Click()
    Dim crxApp As New CRAXDRT.Application
    Dim crxRpt As New CRAXDRT.Report
    
    Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\POS_AccountsReceivableSummary.rpt")
    crxRpt.DiscardSavedData
    crxRpt.EnableParameterPrompting = False
    crxRpt.ParameterFields.GetItemByName("ReportTitle").AddCurrentValue "Accounts Receivable Summary"
    crxRpt.ParameterFields.GetItemByName("@CustomerId").AddCurrentValue 0
    Call ResetRptDB(crxRpt)
    crxRpt.EnableParameterPrompting = False
    crxRpt.PrintOut False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            btnSummary_Click
        Case vbKeyF2
            btnPrintByCustomer_Click
        Case vbKeyF3
            btnAccept_Click
        Case vbKeyEscape
            btnCancel_Click
    End Select
End Sub

Private Sub Form_Load()
    'StatusBarWidth Me, statusBar_Main
    
    lvSearch.ColumnHeaders(2).width = lvSearch.width * 0.15 'Code
    lvSearch.ColumnHeaders(3).width = lvSearch.width * 0.51 'Name
    lvSearch.ColumnHeaders(4).width = lvSearch.width * 0.32 'Balance
'    lvSearch.ColumnHeaders(5).width = lvSearch.width * 0.08 'Date
'    lvSearch.ColumnHeaders(6).width = lvSearch.width * 0.08 'Due
'    lvSearch.ColumnHeaders(7).width = lvSearch.width * 0.1  'Amount
'    lvSearch.ColumnHeaders(8).width = lvSearch.width * 0.1  'Interest
'    lvSearch.ColumnHeaders(9).width = lvSearch.width * 0.11 'Total
'    lvSearch.ColumnHeaders(10).width = lvSearch.width * 0.11 'Balance
    'lvSearch.ColumnHeaders(11).width = lvSearch.width * 0.1  'Payment
    'lvSearch.ColumnHeaders(12).width = lvSearch.width * 0.08 'Mode
    'lvSearch.ColumnHeaders(13).width = lvSearch.width * 0.11 'Remarks
    
    Populate "City"
    
    btnSearch_Click
End Sub

Private Sub lvSearch_DblClick()
    If EditAccessRights(20) = False Then
        MsgBox ErrorCodes(74), vbCritical, "Limited Rights"
        Exit Sub
    End If
    If lvSearch.ListItems.Count <= 0 Then Exit Sub
    
    FIN_CustomerPaymentFrm.CustomerId = lvSearch.SelectedItem.Text
    FIN_CustomerPaymentFrm.Show (1)
End Sub

Private Sub txtSearch_Code_Change()
    btnSearch_Click
End Sub

Private Sub txtSearch_Code_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            btnSearch_Click
    End Select
End Sub

Private Sub txtSearch_Name_Change()
    btnSearch_Click
End Sub

Private Sub txtSearch_Order_Change()
    btnSearch_Click
End Sub

Private Sub txtSearch_Name_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            btnSearch_Click
    End Select
End Sub

