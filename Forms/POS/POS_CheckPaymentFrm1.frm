VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form POS_CheckPaymentFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   Icon            =   "POS_CheckPaymentFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnAccept 
      BackColor       =   &H0080FF80&
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
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4440
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
      TabIndex        =   6
      Top             =   4440
      Width           =   1935
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1560
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   2520
      Width           =   4935
   End
   Begin VB.TextBox txtRemarks 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   3240
      Width           =   4935
   End
   Begin MSComCtl2.DTPicker dtCheckDate 
      Height          =   435
      Left            =   1560
      TabIndex        =   2
      Top             =   2040
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   51838977
      CurrentDate     =   41776
   End
   Begin VB.TextBox txtCheckNumber 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1560
      TabIndex        =   1
      Top             =   1560
      Width           =   4935
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "AMOUNT"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   2520
      Width           =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "REMARKS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Top             =   3240
      Width           =   1020
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
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   675
   End
   Begin VB.Label lblAmountDue 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
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
      ForeColor       =   &H0000FF00&
      Height          =   705
      Left            =   1560
      TabIndex        =   9
      Top             =   720
      Width           =   4935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "CHECK DATE"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "CHECK #"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   870
   End
   Begin VB.Label lblCaption_Title 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check Payment"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
   Begin VB.Image picModuleImage 
      Height          =   480
      Left            =   120
      Picture         =   "POS_CheckPaymentFrm.frx":000C
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "POS_CheckPaymentFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAccept_Click()
    Dim due As Double
    Dim cash As Double
    due = Val(Replace(lblAmountDue.Caption, ",", ""))
    cash = Val(Replace(txtAmount.text, ",", ""))
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
        cmd.Parameters.Append cmd.CreateParameter("@DiscountType", adVarChar, adParamInput, 250, "")
        cmd.Parameters.Append cmd.CreateParameter("@Discount", adDecimal, adParamInput, , POS_CashierFrm.discount)
                              cmd.Parameters("@Discount").NumericScale = 2
                              cmd.Parameters("@Discount").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInput, , 1) 'NOT SET
        cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
        cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, txtRemarks.text)
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
        
        'SAVE CHECK PAYMENT
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "POS_Payment_Insert"
        cmd.Parameters.Append cmd.CreateParameter("@POS_SalesId", adInteger, adParamInput, , POS_SalesId)
        cmd.Parameters.Append cmd.CreateParameter("@PaymentTypeId", adInteger, adParamInput, , 2) 'CHECK
        cmd.Parameters.Append cmd.CreateParameter("@Payment", adDecimal, adParamInput, , Val(Replace(txtAmount.text, ",", "")))
                              cmd.Parameters("@Payment").NumericScale = 2
                              cmd.Parameters("@Payment").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.text)
        cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
        cmd.Parameters.Append cmd.CreateParameter("@CardNumber", adVarChar, adParamInput, 250, Null)
        cmd.Parameters.Append cmd.CreateParameter("@ReferenceNumber", adVarChar, adParamInput, 250, Null)
        cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, txtRemarks.text)
        cmd.Execute
        
        'SAVE CHECK REGISTRY
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "FIN_CheckRegistry_Insert"
        cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.text)
        cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
        cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Val(Replace(txtAmount.text, ",", "")))
                              cmd.Parameters("@Amount").NumericScale = 2
                              cmd.Parameters("@Amount").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, txtRemarks.text)
        cmd.Parameters.Append cmd.CreateParameter("@isReceivable", adBoolean, adParamInput, , "True")
        cmd.Parameters.Append cmd.CreateParameter("@CheckStatusId", adInteger, adParamInput, , 1)
        cmd.Parameters.Append cmd.CreateParameter("@CheckRegistryId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@ExpenseId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@POS_SalesId", adInteger, adParamInput, , POS_SalesId)
        cmd.Execute
        
        'Add Fund Balance
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "FIN_Fund_AddBalance"
        cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , 1) 'CASHIER's FUND
        cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(txtAmount.text, ",", "")))
                              cmd.Parameters("@CheckAmount").NumericScale = 2
                              cmd.Parameters("@CheckAmount").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , 0)
                              cmd.Parameters("@Amount").NumericScale = 2
                              cmd.Parameters("@Amount").Precision = 18
        cmd.Execute
        
        
        'INFLOW AND OUTFLOW
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "FIN_InFlow_Insert"
        cmd.Parameters.Append cmd.CreateParameter("@Cash", adDecimal, adParamInput, , 0)
                              cmd.Parameters("@Cash").Precision = 18
                              cmd.Parameters("@Cash").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(txtAmount.text, ",", "")))
                              cmd.Parameters("@CheckAmount").Precision = 18
                              cmd.Parameters("@CheckAmount").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.text)
        cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
        cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , Format(Now, "MM/DD/YY"))
        cmd.Parameters.Append cmd.CreateParameter("@Particulars", adVarChar, adParamInput, 250, "Cash Sales")
        cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@ExpenseId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@POS_SalesId", adInteger, adParamInput, , POS_SalesId)
        cmd.Execute
        
        con.CommitTrans
        con.Close
        
        Unload POS_SavingFrm
        
        Dim x As Variant
        x = MsgBox("Do you want to print a receipt?", vbQuestion + vbYesNo)
        If x = vbYes Then
        '**PRINT RECEIPT******
            Dim crxApp As New CRAXDRT.Application
            Dim crxRpt As New CRAXDRT.Report
            Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\POS_Receipt.rpt")
            crxRpt.RecordSelectionFormula = "{POS_Sales.POS_SalesId}= " & POS_SalesId & ""
            crxRpt.DiscardSavedData
            
            Call ResetRptDB(crxRpt)
            crxRpt.PrintOut False
            '**END PRINT RECEIPT**
        End If
        
        Unload Me
        Unload POS_PaymentOptionsFrm
        Unload POS_CashPayFrm
        
        POS_ChangeFrm.txtChange.text = FormatNumber(cash - due, 2, vbTrue, vbFalse)
        POS_ChangeFrm.Show (1)
        POS_CashierFrm.Initialize
        
    End If
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            btnCancel_Click
        Case vbKeyReturn
            btnAccept_Click
    End Select
End Sub

Private Sub Form_Load()
    dtCheckDate.value = Format(Now, "MM/DD/YY")
End Sub

Private Sub txtAmount_Change()
    If IsNumeric(txtAmount.text) = False Then
        txtAmount.text = "0.00"
        selectText txtAmount
    Else
'        txtAmount.text = FormatNumber(txtAmount.text, 0)
'        txtAmount.SelStart = Len(txtAmount.text)
    End If
End Sub

Private Sub txtAmount_GotFocus()
    selectText txtAmount
End Sub

Private Sub txtAmount_LostFocus()
    txtAmount.text = FormatNumber(txtAmount.text, 2, vbTrue)
End Sub

Private Sub txtCheckNumber_GotFocus()
    selectText txtCheckNumber
End Sub
