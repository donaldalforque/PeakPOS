VERSION 5.00
Begin VB.Form POS_CashPayFrm 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6600
   Icon            =   "POS_TenderFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnPaymentOption 
      BackColor       =   &H0080FF80&
      Caption         =   "F1: Payment Option"
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton btnCancel 
      BackColor       =   &H008080FF&
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
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2880
      Width           =   1935
   End
   Begin VB.TextBox txtCash 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   39.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   1
      Text            =   "0"
      Top             =   1680
      Width           =   5175
   End
   Begin VB.Label lblAmountDue 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00404040&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   705
      Left            =   1320
      TabIndex        =   6
      Top             =   840
      Width           =   5175
   End
   Begin VB.Label lblCaption_AmountDue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DUE"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   675
   End
   Begin VB.Label lblCaption_Cash 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CASH"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   900
   End
   Begin VB.Label lblCaption_Title 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PAYMENT"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   1590
   End
   Begin VB.Image picModuleImage 
      Height          =   480
      Left            =   120
      Picture         =   "POS_TenderFrm.frx":000C
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "POS_CashPayFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnPaymentOption_Click()
    'payment options
    POS_PaymentOptionsFrm.Show
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            btnPaymentOption_Click
        Case vbKeyEscape
            btnCancel_Click
        Case vbKeyReturn
            Dim due As Double
            Dim cash As Double
            due = Val(Replace(lblAmountDue.Caption, ",", ""))
            cash = Val(Replace(txtCash.text, ",", ""))
            If cash < due Then
                GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(34)
                GLOBAL_MessageFrm.Show (1)
                Exit Sub
            Else
                'SAVE DATA
                Dim POS_SalesId As Integer
                
                POS_SavingFrm.pbSaving.Min = 0
                POS_SavingFrm.pbSaving.Max = POS_CashierFrm.lvList.ListItems.Count
                POS_SavingFrm.Show
                
                Set con = New ADODB.Connection
                Set rec = New ADODB.Recordset
                Set cmd = New ADODB.Command
                
                con.ConnectionString = ConnString
                con.Open
                cmd.ActiveConnection = con
                con.BeginTrans
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "POS_Sales_Insert"
                cmd.Parameters.Append cmd.CreateParameter("@POS_SalesId", adInteger, adParamInputOutput, , POS_SalesId)
                cmd.Parameters.Append cmd.CreateParameter("@Total", adDecimal, adParamInput, , Val(Replace(lblAmountDue.Caption, ",", "")))
                                      cmd.Parameters("@Total").NumericScale = 2
                                      cmd.Parameters("@Total").Precision = 18
                cmd.Parameters.Append cmd.CreateParameter("@Subtotal", adDecimal, adParamInput, , Null)
                                      cmd.Parameters("@Subtotal").NumericScale = 2
                                      cmd.Parameters("@Subtotal").Precision = 18
                cmd.Parameters.Append cmd.CreateParameter("@Tendered", adDecimal, adParamInput, , Val(Replace(txtCash.text, ",", "")))
                                      cmd.Parameters("@Tendered").NumericScale = 2
                                      cmd.Parameters("@Tendered").Precision = 18
                cmd.Parameters.Append cmd.CreateParameter("@DiscountType", adVarChar, adParamInput, 250, "")
                'cmd.Parameters.Append cmd.CreateParameter("@Discount", adDecimal, adParamInput, , POS_CashierFrm.discount)
                                      'cmd.Parameters("@Discount").NumericScale = 2
                                      'cmd.Parameters("@Discount").Precision = 18
                cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInput, , POS_CashierFrm.POSCustomerId) 'NOT SET
                cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
                cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, Null)
                cmd.Execute
                
                POS_SalesId = cmd.Parameters("@POS_SalesId")
                
                'LINE
                Dim item As MSComctlLib.ListItem
                For Each item In POS_CashierFrm.lvList.ListItems
                    Set cmd = New ADODB.Command
                    cmd.ActiveConnection = con
                    cmd.CommandType = adCmdStoredProc
                    cmd.CommandText = "POS_SalesLine_Insert"
                    cmd.Parameters.Append cmd.CreateParameter("@POS_SalesId", adInteger, adParamInput, , POS_SalesId)
                    cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , item.SubItems(7))
                    cmd.Parameters.Append cmd.CreateParameter("@ItemCode", adVarChar, adParamInput, 50, item.text)
                    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, item.SubItems(1))
                    cmd.Parameters.Append cmd.CreateParameter("@Price", adDecimal, adParamInput, , Val(Replace(item.SubItems(3), ",", "")))
                                          cmd.Parameters("@Price").NumericScale = 2
                                          cmd.Parameters("@Price").Precision = 18
                    cmd.Parameters.Append cmd.CreateParameter("@UnitCost", adDecimal, adParamInput, , Val(Replace(item.SubItems(5), ",", "")))
                                          cmd.Parameters("@UnitCost").NumericScale = 2
                                          cmd.Parameters("@UnitCost").Precision = 18
                    cmd.Parameters.Append cmd.CreateParameter("@Quantity", adDecimal, adParamInput, , Val(Replace(item.SubItems(2), ",", "")))
                                          cmd.Parameters("@Quantity").NumericScale = 2
                                          cmd.Parameters("@Quantity").Precision = 18
                    cmd.Parameters.Append cmd.CreateParameter("@Subtotal", adDecimal, adParamInput, , Val(Replace(item.SubItems(4), ",", "")))
                                          cmd.Parameters("@Subtotal").NumericScale = 2
                                          cmd.Parameters("@Subtotal").Precision = 18
                    cmd.Parameters.Append cmd.CreateParameter("@isVatable", adBoolean, adParamInput, , Null)
                    cmd.Parameters.Append cmd.CreateParameter("@VAT", adDecimal, adParamInput, , Null)
                                          cmd.Parameters("@VAT").NumericScale = 2
                                          cmd.Parameters("@VAT").Precision = 18
                    cmd.Parameters.Append cmd.CreateParameter("@LocationId", adInteger, adParamInput, , POS_CashierFrm.POSLocationId)
                    cmd.Execute
                    
                    POS_SavingFrm.pbSaving.value = POS_SavingFrm.pbSaving.value + 1
                Next
                
'                'Add Fund Balance
'                Set cmd = New ADODB.Command
'                cmd.ActiveConnection = con
'                cmd.CommandType = adCmdStoredProc
'                cmd.CommandText = "FIN_Fund_AddBalance"
'                cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , 1) 'CASHIER's FUND
'                cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , 0)
'                                      cmd.Parameters("@CheckAmount").NumericScale = 2
'                                      cmd.Parameters("@CheckAmount").Precision = 18
'                cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Val(Replace(txtCash.text, ",", "")))
'                                      cmd.Parameters("@Amount").NumericScale = 2
'                                      cmd.Parameters("@Amount").Precision = 18
'                cmd.Execute
'
'                'INFLOW AND OUTFLOW
'                Set cmd = New ADODB.Command
'                cmd.ActiveConnection = con
'                cmd.CommandType = adCmdStoredProc
'                cmd.CommandText = "FIN_InFlow_Insert"
'                cmd.Parameters.Append cmd.CreateParameter("@Cash", adDecimal, adParamInput, , Val(Replace(txtCash.text, ",", "")))
'                                      cmd.Parameters("@Cash").Precision = 18
'                                      cmd.Parameters("@Cash").NumericScale = 2
'                cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , 0)
'                                      cmd.Parameters("@CheckAmount").Precision = 18
'                                      cmd.Parameters("@CheckAmount").NumericScale = 2
'                cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, "")
'                cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , Null)
'                cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , Format(Now, "MM/DD/YY"))
'                cmd.Parameters.Append cmd.CreateParameter("@Particulars", adVarChar, adParamInput, 250, "Cash Sales")
'                cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , Null)
'                cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , Null)
'                cmd.Parameters.Append cmd.CreateParameter("@ExpenseId", adInteger, adParamInput, , Null)
'                cmd.Parameters.Append cmd.CreateParameter("@POS_SalesId", adInteger, adParamInput, , POS_SalesId)
'                cmd.Execute
                                          
                con.CommitTrans
                con.Close
                
                Unload POS_SavingFrm
                
                X = MsgBox("Do you want to print a receipt?", vbQuestion + vbYesNo)
                If X = vbYes Then
                '**PRINT RECEIPT******
                    Dim crxApp As New CRAXDRT.Application
                    Dim crxRpt As New CRAXDRT.Report
                    Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\POS_Receipt.rpt")
                    crxRpt.RecordSelectionFormula = "{POS_Sales.POS_SalesId}= " & POS_SalesId & ""
                    crxRpt.DiscardSavedData
                    crxRpt.EnableParameterPrompting = False
                    crxRpt.ParameterFields(1).AddCurrentValue ""
    
                    Call ResetRptDB(crxRpt)
                    crxRpt.PrintOut False
                    '**END PRINT RECEIPT**
                End If
                
                POS_ChangeFrm.txtChange.text = FormatNumber(cash - due, 2, vbTrue, vbFalse)
                POS_ChangeFrm.Show (1)
                POS_CashierFrm.Initialize
                Unload Me
            End If
    End Select
End Sub

Private Sub txtCash_Change()
    If IsNumeric(txtCash.text) = False Then
        txtCash.text = "0.00"
        selectText txtCash
    Else
'        txtCash.text = FormatNumber(txtCash.text, 0)
'        txtCash.SelStart = Len(txtCash.text)
    End If
End Sub

Private Sub txtCash_GotFocus()
    selectText txtCash
End Sub

Private Sub txtCash_LostFocus()
    txtCash.text = FormatNumber(txtCash.text, 2, vbTrue)
End Sub
