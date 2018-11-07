VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form POS_AccountPaymentFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8310
   Icon            =   "POS_AccountPaymentFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPayment 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   -9999
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   2640
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.ComboBox cmbTerms 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      ItemData        =   "POS_AccountPaymentFrm.frx":000C
      Left            =   1680
      List            =   "POS_AccountPaymentFrm.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1560
      Width           =   3375
   End
   Begin VB.ComboBox cmbCustomer 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1680
      TabIndex        =   0
      Top             =   1080
      Width           =   6255
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
      Left            =   4920
      Picture         =   "POS_AccountPaymentFrm.frx":0010
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3960
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
      Left            =   6600
      Picture         =   "POS_AccountPaymentFrm.frx":23E4
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3960
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker dtDue 
      Height          =   465
      Left            =   1680
      TabIndex        =   2
      Top             =   2040
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   820
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   96337921
      CurrentDate     =   41686
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Payment"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   -9999
      TabIndex        =   12
      Top             =   2760
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label lblAmountDue 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   29.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      TabIndex        =   11
      Top             =   2805
      Width           =   4695
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL DUE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   480
      TabIndex        =   10
      Top             =   3000
      Width           =   1590
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   360
      Top             =   2760
      Width           =   7575
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Due"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   360
      TabIndex        =   9
      Top             =   2040
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Terms"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   360
      TabIndex        =   8
      Top             =   1560
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   360
      TabIndex        =   7
      Top             =   1080
      Width           =   1125
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Sales"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   840
      TabIndex        =   5
      Top             =   360
      Width           =   1335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   240
      X2              =   8040
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "POS_AccountPaymentFrm.frx":4773
      Top             =   240
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   3735
      Left            =   120
      Top             =   120
      Width           =   8055
   End
End
Attribute VB_Name = "POS_AccountPaymentFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_Change()

End Sub

Private Sub Text1_GotFocus()
    
End Sub

Private Sub btnAccept_Click()
    Dim x As Variant
    Dim SalesOrderId As Long
    
    x = MsgBox("Are you sure all information are correct?", vbQuestion + vbYesNo)
    If x = vbNo Then Exit Sub
    
    POS_SavingFrm.pbSaving.Min = 0
    POS_SavingFrm.pbSaving.Max = POS_CashierFrm.lvList.ListItems.Count
    POS_SavingFrm.Show
    
    Dim OrderNumber As String
    
    Dim con As New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    con.BeginTrans
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "SO_SalesOrder_Insert"
    cmd.ActiveConnection = con
    
    cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInputOutput, , SalesOrderId)
    cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInputOutput, 50, Null)
    cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , Format(Now, "MM/DD/YY"))
    cmd.Parameters.Append cmd.CreateParameter("@DueDate", adDate, adParamInput, , dtDue.value)
    cmd.Parameters.Append cmd.CreateParameter("StatusId", adInteger, adParamInput, , 4) 'INVOICED
    cmd.Parameters.Append cmd.CreateParameter("@TermId", adInteger, adParamInput, , cmbTerms.ItemData(cmbTerms.ListIndex))
    cmd.Parameters.Append cmd.CreateParameter("@PricingSchemeId", adInteger, adParamInput, , 0)
    cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInput, , POS_CashierFrm.POSCustomerId)
    cmd.Parameters.Append cmd.CreateParameter("@Days", adDecimal, adParamInput, , 0)
                              cmd.Parameters("@Days").Precision = 18
                              cmd.Parameters("@Days").NumericScale = 2
    cmd.Parameters.Append cmd.CreateParameter("@InterestRate", adDecimal, adParamInput, , 0)
                          cmd.Parameters("@InterestRate").Precision = 18
                          cmd.Parameters("@InterestRate").NumericScale = 2
    cmd.Parameters.Append cmd.CreateParameter("@Cash", adDecimal, adParamInput, , 0)
                          cmd.Parameters("@Cash").Precision = 18
                          cmd.Parameters("@Cash").NumericScale = 2
    cmd.Parameters.Append cmd.CreateParameter("@Interest", adDecimal, adParamInput, , 0)
                          cmd.Parameters("@Interest").Precision = 18
                          cmd.Parameters("@Interest").NumericScale = 2
    cmd.Parameters.Append cmd.CreateParameter("@Subtotal", adDecimal, adParamInput, , NVAL(POS_PayFrm.lblAmountDue.Caption))
                          cmd.Parameters("@Subtotal").Precision = 18
                          cmd.Parameters("@Subtotal").NumericScale = 2
    cmd.Parameters.Append cmd.CreateParameter("@Total", adDecimal, adParamInput, , NVAL(POS_PayFrm.lblAmountDue.Caption))
                          cmd.Parameters("@Total").Precision = 18
                          cmd.Parameters("@Total").NumericScale = 2
    cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 500, "")
    cmd.Parameters.Append cmd.CreateParameter("@Salesman", adVarChar, adParamInput, 250, "")
    cmd.Parameters.Append cmd.CreateParameter("@ReferenceNumber", adVarChar, adParamInput, 250, "")
    cmd.Parameters.Append cmd.CreateParameter("@Discount", adDecimal, adParamInput, , 0)
                          cmd.Parameters("@Discount").Precision = 18
                          cmd.Parameters("@Discount").NumericScale = 2
    cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , 1) 'NOT SET!
    cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , Null)
    cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
    cmd.Execute
    
    OrderNumber = cmd.Parameters("@OrderNumber")
    SalesOrderId = cmd.Parameters("@SalesOrderId")
    
    'SAVE ORDER LINE
    Dim item As MSComctlLib.ListItem

    For Each item In POS_CashierFrm.lvList.ListItems
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        
        cmd.Parameters.Append cmd.CreateParameter("@SalesOrderLineId", adInteger, adParamInputOutput, , 0)
        cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SalesOrderId)
        cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInput, , POS_CashierFrm.POSCustomerId)
        cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , Format(Now, "mm/dd/yy"))
        cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , item.SubItems(8))
        cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, item.Text)
        cmd.Parameters.Append cmd.CreateParameter("@Quantity", adDecimal, adParamInput, , NVAL(item.SubItems(1)))
                              cmd.Parameters("@Quantity").Precision = 18
                              cmd.Parameters("@Quantity").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@Uom", adVarChar, adParamInput, 250, item.SubItems(2))
        cmd.Parameters.Append cmd.CreateParameter("@Price", adDecimal, adParamInput, , NVAL(item.SubItems(3)))
                              cmd.Parameters("@Price").Precision = 18
                              cmd.Parameters("@Price").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@Subtotal", adDecimal, adParamInput, , NVAL(item.SubItems(5)))
                              cmd.Parameters("@Subtotal").Precision = 18
                              cmd.Parameters("@Subtotal").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@LocationId", adInteger, adParamInput, , 1)
        cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , 4)
        cmd.Parameters.Append cmd.CreateParameter("@ReserveId", adInteger, adParamInput, , 0)
        cmd.Parameters.Append cmd.CreateParameter("@ActualQuantity", adDecimal, adParamInput, , NVAL(item.SubItems(16)))
                              cmd.Parameters("@ActualQuantity").Precision = 18
                              cmd.Parameters("@ActualQuantity").NumericScale = 2
                              
        cmd.CommandText = "SO_SalesOrderLine_Insert"
        cmd.Execute
    Next
    
    'PICK UP
    Dim PickOrderId As Long
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "SO_PickOrder_Insert"
    cmd.Parameters.Append cmd.CreateParameter("@PickOrderId", adInteger, adParamInputOutput, , PickOrderId)
    cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SalesOrderId)
    cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, "")
    cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
    cmd.Execute
    
    PickOrderId = cmd.Parameters("@PickOrderId")
    
    'Pick Order Line
    Dim PickOrderLineId As Long
    
    For Each item In POS_CashierFrm.lvList.ListItems
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "SO_PickOrderLine_Insert"
        cmd.Parameters.Append cmd.CreateParameter("@PickOrderLineId", adInteger, adParamInputOutput, , PickOrderLineId)
        cmd.Parameters.Append cmd.CreateParameter("@PickOrderId", adInteger, adParamInput, , PickOrderId)
        cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SalesOrderId)
        cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , item.SubItems(8))
        cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, item.Text)
        cmd.Parameters.Append cmd.CreateParameter("@Uom", adVarChar, adParamInput, 50, item.SubItems(2))
        cmd.Parameters.Append cmd.CreateParameter("@LocationId", adInteger, adParamInput, , 1)
        cmd.Parameters.Append cmd.CreateParameter("@Quantity", adDecimal, adParamInput, , NVAL(item.SubItems(1)))
                              cmd.Parameters("@Quantity").Precision = 18
                              cmd.Parameters("@Quantity").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , Format(Now, "mm/dd/yy"))
        cmd.Parameters.Append cmd.CreateParameter("@Reference", adVarChar, adParamInput, 250, "")
        cmd.Parameters.Append cmd.CreateParameter("@ReserveId", adInteger, adParamInput, , 0)
        cmd.Parameters.Append cmd.CreateParameter("@SalesOrderLineId", adInteger, adParamInput, , 0)
        cmd.Execute
    Next
    
    'SAVE INVOICE
    Dim InvoiceId As Long
    Set cmd = New ADODB.Command
    cmd.CommandType = adCmdStoredProc
    cmd.ActiveConnection = con
    cmd.Parameters.Append cmd.CreateParameter("@InvoiceId", adInteger, adParamInputOutput, , Val(InvoiceId))
    cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SalesOrderId)
    cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , Format(Now, "mm/dd/yy"))
    cmd.Parameters.Append cmd.CreateParameter("@DueDate", adDate, adParamInput, , dtDue.value)
    cmd.Parameters.Append cmd.CreateParameter("@Discount", adDecimal, adParamInput, , 0)
                          cmd.Parameters("@Discount").Precision = 18
                          cmd.Parameters("@Discount").NumericScale = 2
    cmd.Parameters.Append cmd.CreateParameter("@Refunds", adDecimal, adParamInput, , 0)
                          cmd.Parameters("@Refunds").Precision = 18
                          cmd.Parameters("@Refunds").NumericScale = 2
    cmd.Parameters.Append cmd.CreateParameter("@SubTotal", adDecimal, adParamInput, , NVAL(POS_PayFrm.lblAmountDue.Caption))
                          cmd.Parameters("@SubTotal").Precision = 18
                          cmd.Parameters("@SubTotal").NumericScale = 2
    cmd.Parameters.Append cmd.CreateParameter("@Total", adDecimal, adParamInput, , NVAL(POS_PayFrm.lblAmountDue.Caption))
                          cmd.Parameters("@Total").Precision = 18
                          cmd.Parameters("@Total").NumericScale = 2
    cmd.Parameters.Append cmd.CreateParameter("@ReferenceNumber", adVarChar, adParamInput, 50, "")
    cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, "")
    cmd.CommandText = "SO_Invoice_Insert"
    cmd.Execute
    InvoiceId = cmd.Parameters("@InvoiceId")
    
    'SAVE INVOICE LINE
    Dim InvoiceLineId As Long
    For Each item In POS_CashierFrm.lvList.ListItems
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        
        cmd.Parameters.Append cmd.CreateParameter("@InvoiceLineId", adInteger, adParamInputOutput, , 0)
        cmd.Parameters.Append cmd.CreateParameter("@InvoiceId", adInteger, adParamInput, , InvoiceId)
        cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , Format(Now, "mm/dd/yy"))
        cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , Val(item.SubItems(8)))
        cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, item.Text)
        cmd.Parameters.Append cmd.CreateParameter("@Quantity", adDecimal, adParamInput, , NVAL(item.SubItems(1)))
                              cmd.Parameters("@Quantity").Precision = 18
                              cmd.Parameters("@Quantity").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@Uom", adVarChar, adParamInput, 250, item.SubItems(2))
        cmd.Parameters.Append cmd.CreateParameter("@Price", adDecimal, adParamInput, , NVAL(item.SubItems(3)))
                              cmd.Parameters("@Price").Precision = 18
                              cmd.Parameters("@Price").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@Subtotal", adDecimal, adParamInput, , NVAL(item.SubItems(5)))
                              cmd.Parameters("@Subtotal").Precision = 18
                              cmd.Parameters("@Subtotal").NumericScale = 2
        cmd.CommandText = "SO_InvoiceLine_Insert"
        cmd.Execute
        
        POS_SavingFrm.pbSaving.value = POS_SavingFrm.pbSaving.value + 1
    Next
    
    'UPDATE SO REMAINING BALANCE
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "SO_Balance_Update"
    cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SalesOrderId)
    cmd.Execute
    
    'PAYMENT
    If NVAL(txtPayment.Text) > 0 Then
        'TRANSACTION ID
        Dim TransactionId As Long
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "BASE_TransactionId_Insert"
        cmd.Parameters.Append cmd.CreateParameter("@TransactionId", adInteger, adParamInputOutput, , 0)
        cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, "Customer Payment")
        cmd.Execute
        TransactionId = cmd.Parameters("@TransactionId")
        
        'PAYMENT
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "SO_Payment_Insert"
        cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SalesOrderId)
        cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , NVAL(txtPayment.Text))
                              cmd.Parameters("@Amount").NumericScale = 2
                              cmd.Parameters("@Amount").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , Format(Now, "mm/dd/yy"))
        cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , 0)
                      cmd.Parameters("@CheckAmount").NumericScale = 2
                      cmd.Parameters("@CheckAmount").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, "")
        cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , Format(Now, "mm/dd/yy"))
        cmd.Parameters.Append cmd.CreateParameter("@SalesReturn", adDecimal, adParamInput, , Null)
                              cmd.Parameters("@SalesReturn").NumericScale = 2
                              cmd.Parameters("@SalesReturn").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , 1)
        cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, "")
        cmd.Parameters.Append cmd.CreateParameter("@IssuingBank", adVarChar, adParamInput, 250, "")
        cmd.Parameters.Append cmd.CreateParameter("@isOnline", adBoolean, adParamInput, , False)
        cmd.Parameters.Append cmd.CreateParameter("@SOPaymentId", adInteger, adParamInputOutput, , 0)
        cmd.Parameters.Append cmd.CreateParameter("@TransactionId", adInteger, adParamInput, , TransactionId)
        cmd.Execute
        
        
         'PAYMENT HISTORY
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "SO_PaymentHistory_Insert"
        cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInput, , POS_CashierFrm.POSCustomerId)
        cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , Format(Now, "mm/dd/yy"))
        cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , NVAL(txtPayment.Text))
                              cmd.Parameters("@Amount").NumericScale = 2
                              cmd.Parameters("@Amount").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , 0)
                              cmd.Parameters("@CheckAmount").NumericScale = 2
                              cmd.Parameters("@CheckAmount").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 50, "")
        cmd.Parameters.Append cmd.CreateParameter("@SalesDiscount", adDecimal, adParamInput, , 0)
                              cmd.Parameters("@SalesDiscount").NumericScale = 2
                              cmd.Parameters("@SalesDiscount").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , Now)
        cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 4000, OrderNumber)
        cmd.Parameters.Append cmd.CreateParameter("@IssuingBank", adVarChar, adParamInput, 250, "")
        cmd.Parameters.Append cmd.CreateParameter("@TransactionId", adInteger, adParamInput, , TransactionId)
        cmd.Execute
    End If
    
    'SAVE POS AUDIT
    SavePOSAuditTrail UserId, WorkstationId, 0, "PROCESSED CREDIT SALES"
    
    con.CommitTrans
    con.Close
    
    UpdateCustomerOrderDues
    
    Unload POS_SavingFrm
    'Dim x As Variant
    x = MsgBox("Do you want to print a receipt?", vbQuestion + vbYesNo)
    If x = vbYes Then
        '**PRINT RECEIPT******
        Dim crxApp As New CRAXDRT.Application
        Dim crxRpt As New CRAXDRT.Report
        Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\POS_Receipt_Account.rpt")
        'crxRpt.RecordSelectionFormula = "{POS_Sales.POS_SalesId}= " & Val(POS_SalesId) & ""
        crxRpt.DiscardSavedData
        crxRpt.EnableParameterPrompting = False
        crxRpt.ParameterFields.GetItemByName("Notice").AddCurrentValue ""
        crxRpt.RecordSelectionFormula = "{SO_SalesOrder.SalesOrderId}= " & Val(SalesOrderId) & ""
        'crxRpt.ParameterFields.GetItemByName("@POS_SalesId").AddCurrentValue Val(POS_SalesId)

        Call ResetRptDB(crxRpt)
        crxRpt.PrintOut False
        
        '**END PRINT RECEIPT**
    End If
    
    POS_CashierFrm.Initialize
    ClearClassData (0)
    ClearClassData (1)
    ClearClassData (2)
    ClearClassData (3)
    
    Unload Me
    Unload POS_PayFrm
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub cmbTerms_Click()
    If cmbTerms.ListIndex > 0 Then
        dtDue.value = Format(Now, "MM/DD/YY")
        dtDue.value = dtDue.value + GetTermDays(cmbTerms.ItemData(cmbTerms.ListIndex))
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            btnAccept_Click
        Case vbKeyEscape
            btnCancel_Click
    End Select
End Sub

Private Sub Form_Load()
    lblAmountDue.Caption = POS_PayFrm.lblAmountDue.Caption
    dtDue.value = Format(Now, "mm/dd/yy")
    Populate "Terms"
    
    On Error Resume Next
    cmbTerms.Text = "15 Days"
End Sub

Public Sub Populate(ByVal data As String)
    Select Case data
        Case "Terms"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Terms")
            cmbTerms.Clear
            cmbTerms.AddItem ""
            cmbTerms.ItemData(cmbTerms.NewIndex) = 0
            If Not rec.EOF Then
                Do Until rec.EOF
                    cmbTerms.AddItem rec!Terms
                    cmbTerms.ItemData(cmbTerms.NewIndex) = rec!TermId
                    rec.MoveNext
                Loop
            End If
            cmbTerms.ListIndex = 0
    End Select
End Sub

Private Sub txtPayment_Change()
    If NVAL(txtPayment.Text) > NVAL(lblAmountDue.Caption) Then
        txtPayment.Text = lblAmountDue.Caption
    End If
End Sub

Private Sub txtPayment_GotFocus()
    selectText txtPayment
End Sub
