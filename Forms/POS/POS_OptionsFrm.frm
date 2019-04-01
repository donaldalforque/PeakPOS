VERSION 5.00
Begin VB.Form POS_OptionsFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4470
   Icon            =   "POS_OptionsFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   4470
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnAnalysis 
      Caption         =   "F7: Sales Analysis"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   120
      Picture         =   "POS_OptionsFrm.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4080
      Width           =   4215
   End
   Begin VB.CommandButton btnCashBreakDown 
      Caption         =   "F6: Cash Breakdown"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   2280
      Picture         =   "POS_OptionsFrm.frx":9444
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton btnChangeFund 
      Caption         =   "F5: Change Fund"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   120
      Picture         =   "POS_OptionsFrm.frx":9A69
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      Width           =   2055
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
      Height          =   1200
      Left            =   120
      Picture         =   "POS_OptionsFrm.frx":A08A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5400
      Width           =   4215
   End
   Begin VB.CommandButton btnCashOut 
      Caption         =   "F4: Cash Out"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   2280
      Picture         =   "POS_OptionsFrm.frx":C419
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton btnCashIn 
      Caption         =   "F3: Cash In"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   120
      Picture         =   "POS_OptionsFrm.frx":C924
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton btnHoldList 
      Caption         =   "F2: Hold List"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   2280
      Picture         =   "POS_OptionsFrm.frx":CE46
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.CommandButton btnHoldOrder 
      Caption         =   "F1: Hold Order"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   120
      Picture         =   "POS_OptionsFrm.frx":D440
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "POS_OptionsFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAnalysis_Click()
    AllowAccess = False
    POS_UserPinFrm.Show (1)
    If AllowAccess = False Then Exit Sub
    POS_SalesAnalysisFrm.Show (1)
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnCashBreakDown_Click()
    POS_CashBreakDownFrm.Show (1)
End Sub

Private Sub btnCashIn_Click()
    POS_CashInFrm.Show (1)
End Sub

Private Sub btnCashOut_Click()
    POS_CashOutFrm.Show (1)
End Sub

Private Sub btnChangeFund_Click()
    POS_ChangeFundFrm.Show (1)
End Sub

Private Sub btnHoldList_Click()
    POS_HoldListFrm.Show (1)
End Sub

Private Sub btnHoldOrder_Click()
    Dim x As String
    x = InputBox("Please input name/number #:", , POS_CashierFrm.POSHoldOrderReference)
    If Trim(x) <> "" Then
        Dim con As New ADODB.Connection
        con.ConnectionString = ConnString
        con.Open
        con.BeginTrans
        
        'IF SAME RECORD'
        If x = POS_CashierFrm.POSHoldOrderReference Then
            Set cmd = New ADODB.Command
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "POS_Order_Delete"
            cmd.ActiveConnection = con
            cmd.Parameters.Append cmd.CreateParameter("@POS_OrderId", adInteger, adParamInput, , POS_CashierFrm.POSOrderId)
            cmd.Execute
        End If
        
        With POS_CashierFrm
            'SAVE MAIN DETAILS
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "POS_Order_Insert"
            cmd.Parameters.Append cmd.CreateParameter("@POS_OrderId", adInteger, adParamInputOutput, , POS_OrderId)
            cmd.Parameters.Append cmd.CreateParameter("@Total", adDecimal, adParamInput, , NVAL(.txtTotal.Caption))
                                  cmd.Parameters("@Total").NumericScale = 2
                                  cmd.Parameters("@Total").Precision = 18
            cmd.Parameters.Append cmd.CreateParameter("@UserNumber", adInteger, adParamInput, , NVAL(.txtUserNumber.Text))
            cmd.Parameters.Append cmd.CreateParameter("@WorkStationId", adInteger, adParamInput, , WorkstationId)
            cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInput, , .POSCustomerId)
            cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , Now)
            cmd.Parameters.Append cmd.CreateParameter("@CurrentUserId", adInteger, adParamInput, , UserId)
            cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , 1) 'OPEN
            cmd.Parameters.Append cmd.CreateParameter("@RefereceNumber", adVarChar, adParamInput, 50, x)
            cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 50, "HOLD")
            cmd.Parameters.Append cmd.CreateParameter("@SalesmanId", adInteger, adParamInput, , POS_CashierFrm.SalesmanId)
            cmd.Execute
            POS_OrderId = Val(cmd.Parameters("@POS_OrderId"))
            
            'SAVE LINE
            Dim Item As MSComctlLib.ListItem
            For Each Item In .lvList.ListItems
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "POS_OrderLine_Insert"
                
                cmd.Parameters.Append cmd.CreateParameter("@POS_OrderLineId", adInteger, adParamInputOutput, , NVAL(Item.SubItems(22)))
                cmd.Parameters.Append cmd.CreateParameter("@POS_OrderId", adInteger, adParamInput, , POS_OrderId)
                cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, Item.Text)
                cmd.Parameters.Append cmd.CreateParameter("@Quantity", adDecimal, adParamInput, , NVAL(Item.SubItems(1)))
                                      cmd.Parameters("@Quantity").Precision = 18
                                      cmd.Parameters("@Quantity").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@Unit", adVarChar, adParamInput, 50, Item.SubItems(2))
                cmd.Parameters.Append cmd.CreateParameter("@Price", adDecimal, adParamInput, , NVAL(Item.SubItems(3)))
                                      cmd.Parameters("@Price").Precision = 18
                                      cmd.Parameters("@Price").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@Discount", adDecimal, adParamInput, , NVAL(Item.SubItems(4)))
                                      cmd.Parameters("@Discount").Precision = 18
                                      cmd.Parameters("@Discount").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@Subtotal", adDecimal, adParamInput, , NVAL(Item.SubItems(5)))
                                      cmd.Parameters("@Subtotal").Precision = 18
                                      cmd.Parameters("@Subtotal").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@UnitCost", adDecimal, adParamInput, , NVAL(Item.SubItems(6)))
                                      cmd.Parameters("@UnitCost").Precision = 18
                                      cmd.Parameters("@UnitCost").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@HiddenQuantity", adDecimal, adParamInput, , NVAL(Item.SubItems(7)))
                                      cmd.Parameters("@HiddenQuantity").Precision = 18
                                      cmd.Parameters("@HiddenQuantity").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , NVAL(Item.SubItems(8)))
                cmd.Parameters.Append cmd.CreateParameter("@HiddenPrice", adDecimal, adParamInput, , NVAL(Item.SubItems(9)))
                                      cmd.Parameters("@HiddenPrice").Precision = 18
                                      cmd.Parameters("@HiddenPrice").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@Tax", adDecimal, adParamInput, , NVAL(Item.SubItems(13)))
                                      cmd.Parameters("@Tax").Precision = 18
                                      cmd.Parameters("@Tax").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@TaxComputation", adDecimal, adParamInput, , NVAL(Item.SubItems(14)))
                                      cmd.Parameters("@TaxComputation").Precision = 18
                                      cmd.Parameters("@TaxComputation").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@DiscountType", adVarChar, adParamInput, 50, Item.SubItems(15))
                cmd.Parameters.Append cmd.CreateParameter("@DeductInventory", adDecimal, adParamInput, , NVAL(Item.SubItems(16)))
                                      cmd.Parameters("@DeductInventory").Precision = 18
                                      cmd.Parameters("@DeductInventory").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@Discounted", adDecimal, adParamInput, , NVAL(Item.SubItems(17)))
                                      cmd.Parameters("@Discounted").Precision = 18
                                      cmd.Parameters("@Discounted").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@ReserveId", adInteger, adParamInput, , NVAL(Item.SubItems(18)))
                cmd.Parameters.Append cmd.CreateParameter("@DiscountPercent", adDecimal, adParamInput, , NVAL(Item.SubItems(19)))
                                      cmd.Parameters("@DiscountPercent").Precision = 18
                                      cmd.Parameters("@DiscountPercent").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@isTaxExempt", adVarChar, adParamInput, 50, Item.SubItems(20))
                cmd.Parameters.Append cmd.CreateParameter("@ActualQuantity", adDecimal, adParamInput, , NVAL(Item.SubItems(16)))
                                      cmd.Parameters("@ActualQuantity").Precision = 18
                                      cmd.Parameters("@ActualQuantity").NumericScale = 2
                cmd.Parameters.Append cmd.CreateParameter("@Percentage", adDecimal, adParamInput, , NVAL(Item.SubItems(13)))
                                      cmd.Parameters("@Percentage").Precision = 18
                                      cmd.Parameters("@Percentage").NumericScale = 2
                cmd.Execute
                Item.SubItems(22) = cmd.Parameters("@POS_OrderLineId")
                Item.SubItems(21) = POS_OrderId
            Next
        End With
        
        con.CommitTrans
        con.Close
        
        Unload Me
        POS_CashierFrm.Initialize
    Else
        MsgBox "Invalid name/number", vbCritical, "PeakPOS"
    End If
End Sub

Private Sub Command1_Click()
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            btnHoldOrder_Click
        Case vbKeyF2
            btnHoldList_Click
        Case vbKeyF3 'Cash in
            btnCashIn_Click
        Case vbKeyF4 'cash out
            btnCashOut_Click
        Case vbKeyF5 'change fund
            btnChangeFund_Click
        Case vbKeyF6
            btnCashBreakDown_Click
        Case vbKeyF7
            btnAnalysis_Click
        Case vbKeyEscape
            btnCancel_Click
    End Select
End Sub

