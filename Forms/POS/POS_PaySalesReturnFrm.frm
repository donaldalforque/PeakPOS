VERSION 5.00
Begin VB.Form POS_PaySalesReturnFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRemarks 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   6855
   End
   Begin VB.OptionButton optCredit 
      Caption         =   "CREDIT SALES RETURN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   6855
   End
   Begin VB.OptionButton optCash 
      Caption         =   "CASH SALES RETURN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Value           =   -1  'True
      Width           =   6855
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
      Left            =   5400
      Picture         =   "POS_PaySalesReturnFrm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4680
      Width           =   1575
   End
   Begin VB.CommandButton btnAccept 
      Caption         =   "ENTER: OK"
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
      Left            =   3720
      Picture         =   "POS_PaySalesReturnFrm.frx":238F
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "REASON:"
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
      Left            =   120
      TabIndex        =   6
      Top             =   2880
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PROCEED SALES RETURN?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   4080
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   0
      Picture         =   "POS_PaySalesReturnFrm.frx":4763
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "POS_PaySalesReturnFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAccept_Click()
    If Trim(txtRemarks.Text) = "" Then
        MsgBox "Please enter a valid reason.", vbCritical
        txtRemarks.SetFocus
        Exit Sub
    End If
    
    If optCash.value = True Then
        SaveSalesReturn
    Else
        If POS_CashierFrm.POSCustomerId = 0 Then
            MsgBox "Please select a customer.", vbInformation
            POS_CustomerNameFrm.Show (1)
            If POS_CashierFrm.POSCustomerId <> 0 Then
                'PROCESS RETURN
                SaveSalesReturn
            End If
        Else
            SaveSalesReturn
        End If
    End If
End Sub
Private Sub SaveSalesReturn()
    'SAVE DATA
    Dim POS_SalesReturnId As String
    Dim LoyaltyPointsDiv As Double
    
    Dim rsReceipt As New ADODB.Recordset
'    Set rsReceipt = CreateRecordset
    
    'ComputeTotal SalesTax
    SalesTax = 0
    TaxExempt = 0
    For Each item In POS_CashierFrm.lvList.ListItems
        If item.SubItems(20) = "True" Then
            TaxExempt = TaxExempt + Val(Replace(item.SubItems(5), ",", ""))
        Else
            SalesTax = SalesTax + item.SubItems(14)
        End If
        
        TotalDiscount = TotalDiscount + Val(Replace(item.SubItems(17), ",", ""))
    Next
    
    
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    con.BeginTrans
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "POS_SalesReturn_Insert"
    cmd.Parameters.Append cmd.CreateParameter("@POS_SalesReturnId", adInteger, adParamInputOutput, , Val(POS_SalesReturnId))
    cmd.Parameters.Append cmd.CreateParameter("@Total", adDecimal, adParamInput, , NVAL(POS_CashierFrm.txtTotal.Caption))
                          cmd.Parameters("@Total").NumericScale = 2
                          cmd.Parameters("@Total").Precision = 18
    cmd.Parameters.Append cmd.CreateParameter("@Subtotal", adDecimal, adParamInput, , Null)
                          cmd.Parameters("@Subtotal").NumericScale = 2
                          cmd.Parameters("@Subtotal").Precision = 18
    cmd.Parameters.Append cmd.CreateParameter("@Tendered", adDecimal, adParamInput, , 0)
                          cmd.Parameters("@Tendered").NumericScale = 2
                          cmd.Parameters("@Tendered").Precision = 18
    cmd.Parameters.Append cmd.CreateParameter("@DiscountType", adVarChar, adParamInput, 250, "")
    cmd.Parameters.Append cmd.CreateParameter("@SalesTax", adDecimal, adParamInput, , SalesTax)
                          cmd.Parameters("@SalesTax").NumericScale = 2
                          cmd.Parameters("@SalesTax").Precision = 18
    cmd.Parameters.Append cmd.CreateParameter("@TaxExempt", adDecimal, adParamInput, , TaxExempt)
                          cmd.Parameters("@TaxExempt").NumericScale = 2
                          cmd.Parameters("@TaxExempt").Precision = 18
    cmd.Parameters.Append cmd.CreateParameter("@Discount", adDecimal, adParamInput, , TotalDiscount)
                          cmd.Parameters("@Discount").NumericScale = 2
                          cmd.Parameters("@Discount").Precision = 18
    cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInput, , POS_CashierFrm.POSCustomerId) 'NOT SET
    cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
    cmd.Parameters.Append cmd.CreateParameter("@WorkStationid", adInteger, adParamInput, , WorkstationId)
    cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, Null)
    If optCash.value = True Then
        cmd.Parameters.Append cmd.CreateParameter("@Type", adVarChar, adParamInput, 250, "CASH")
    Else
        cmd.Parameters.Append cmd.CreateParameter("@Type", adVarChar, adParamInput, 250, "CREDIT")
    End If
    cmd.Execute
    
    POS_SalesReturnId = cmd.Parameters("@POS_SalesReturnId")
    
    Dim OrderNumber As String
    Dim POS_OrderId As String
'        OrderNumber = cmd.Parameters("@OrderNumber")
    
    'LINE
    For Each item In POS_CashierFrm.lvList.ListItems
        POS_OrderId = NVAL(item.SubItems(21))
        If item.SubItems(20) = "True" Then
            TaxExempt = Val(Replace(item.SubItems(5), ",", ""))
            SalesTax = 0
        Else
            SalesTax = item.SubItems(14)
            TaxExempt = 0
        End If
    
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "POS_SalesReturnLine_Insert"
        cmd.Parameters.Append cmd.CreateParameter("@POS_SalesReturnId", adInteger, adParamInput, , Val(POS_SalesReturnId))
        cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , item.SubItems(8))
        cmd.Parameters.Append cmd.CreateParameter("@Unit", adVarChar, adParamInput, 50, item.SubItems(2))
        cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, item.Text)
        cmd.Parameters.Append cmd.CreateParameter("@Price", adDecimal, adParamInput, , Val(Replace(item.SubItems(3), ",", "")))
                              cmd.Parameters("@Price").NumericScale = 2
                              cmd.Parameters("@Price").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@UnitCost", adDecimal, adParamInput, , Val(Replace(item.SubItems(6), ",", "")))
                              cmd.Parameters("@UnitCost").NumericScale = 2
                              cmd.Parameters("@UnitCost").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@Quantity", adDecimal, adParamInput, , Val(Replace(item.SubItems(1), ",", "")))
                              cmd.Parameters("@Quantity").NumericScale = 2
                              cmd.Parameters("@Quantity").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@Subtotal", adDecimal, adParamInput, , Val(Replace(item.SubItems(5), ",", "")))
                              cmd.Parameters("@Subtotal").NumericScale = 2
                              cmd.Parameters("@Subtotal").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@Tax", adDecimal, adParamInput, , SalesTax)
                              cmd.Parameters("@Tax").NumericScale = 2
                              cmd.Parameters("@Tax").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@TaxExempt", adDecimal, adParamInput, , TaxExempt)
                              cmd.Parameters("@TaxExempt").NumericScale = 2
                              cmd.Parameters("@TaxExempt").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@ItemDiscount", adDecimal, adParamInput, , Val(Replace(item.SubItems(17), ",", "")))
                              cmd.Parameters("@ItemDiscount").NumericScale = 2
                              cmd.Parameters("@ItemDiscount").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@ActualQuantity", adDecimal, adParamInput, , (Val(Replace(item.SubItems(1), ",", "")) * Val(Replace(item.SubItems(16), ",", ""))))
                              cmd.Parameters("@ActualQuantity").NumericScale = 2
                              cmd.Parameters("@ActualQuantity").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@LocationId", adInteger, adParamInput, , POS_CashierFrm.POSLocationId)
        
        cmd.Execute
        
        'DELETE RESERVE LINE
        DeleteReserveLine item.SubItems(18)
    Next
    
    'SAVE POS_Audit
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "POS_UserAudit_Insert"
    cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
    cmd.Parameters.Append cmd.CreateParameter("@WorkStationId", adInteger, adParamInput, , WorkstationId)
    cmd.Parameters.Append cmd.CreateParameter("@POS_SalesReturnId", adInteger, adParamInput, , Val(POS_SalesReturnId))
    cmd.Parameters.Append cmd.CreateParameter("@Activity", adVarChar, adParamInput, 250, "ACCEPT PAYMENT")
    cmd.Parameters.Append cmd.CreateParameter("@Module", adVarChar, adParamInput, 250, "POS")
    cmd.Execute
    
    'UPDATE ORDERS STATUS
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "POS_OrderStatus_Update"
    cmd.Parameters.Append cmd.CreateParameter("@POS_OrderId", adInteger, adParamInput, , POS_OrderId)
    cmd.Execute
    
    'DELETE ORDERS
    If NVAL(POS_CashierFrm.POSOrderId) <> 0 Then
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "POS_Order_Delete"
        cmd.Parameters.Append cmd.CreateParameter("@POS_OrderId", adInteger, adParamInput, , NVAL(POS_CashierFrm.POSOrderId))
        Set rec = cmd.Execute
    End If
    
    con.CommitTrans
    con.Close
    
    Unload POS_SavingFrm
    Dim x As Variant
    x = MsgBox("Do you want to print a receipt?", vbQuestion + vbYesNo)
    If x = vbYes Then
        '**PRINT RECEIPT******
        Dim crxApp As New CRAXDRT.Application
        Dim crxRpt As New CRAXDRT.Report
        Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\POS_Receipt_Return.rpt")
        crxRpt.DiscardSavedData
        crxRpt.EnableParameterPrompting = False
        crxRpt.ParameterFields.GetItemByName("Notice").AddCurrentValue ""
        crxRpt.ParameterFields.GetItemByName("@POS_SalesReturnId").AddCurrentValue Val(POS_SalesReturnId)

        Call ResetRptDB(crxRpt)
        crxRpt.PrintOut False
        
        '**END PRINT RECEIPT**
    End If
    
    POS_CashierFrm.Initialize
    
    MsgBox "Sales Return complete.", vbInformation
    Unload Me
    POS_CashierFrm.Initialize
End Sub

Private Sub btnCancel_Click()
    Dim x As Variant
    x = MsgBox("Disable sales return mode?", vbQuestion + vbYesNo)
    If x = vbYes Then
        POS_CashierFrm.SalesReturn = False
        MsgBox "SALES RETURN MODE DISABLED. You can now process a regular sale.", vbInformation
        Unload Me
    Else
        Unload Me
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

