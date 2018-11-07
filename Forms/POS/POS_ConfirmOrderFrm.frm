VERSION 5.00
Begin VB.Form POS_ConfirmOrderFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8175
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnPayCash 
      BackColor       =   &H00C0FFFF&
      Caption         =   "F1: Pay / Charge "
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
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Width           =   2775
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
      Left            =   6480
      Picture         =   "POS_ConfirmOrderFrm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton btnAccept 
      Caption         =   "ENTER: Save Order"
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
      Left            =   3840
      Picture         =   "POS_ConfirmOrderFrm.frx":238F
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   7935
      Begin VB.Label Label2 
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
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1590
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
         Left            =   3000
         TabIndex        =   3
         Top             =   180
         Width           =   4695
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   855
         Left            =   120
         Top             =   120
         Width           =   7695
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONFIRM ORDER"
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
      Left            =   840
      TabIndex        =   5
      Top             =   240
      Width           =   2595
   End
   Begin VB.Image picModuleImage 
      Height          =   480
      Left            =   240
      Picture         =   "POS_ConfirmOrderFrm.frx":4763
      Top             =   240
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   0
      Picture         =   "POS_ConfirmOrderFrm.frx":4D85
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "POS_ConfirmOrderFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim POS_OrderId As Long
Private Sub btnAccept_Click()
    POS_ConfirmPaymentFrm.Show (1)
    If AllowAccess = False Then Exit Sub
    
    Dim con As New ADODB.Connection
    con.ConnectionString = ConnString
    con.Open
    con.BeginTrans
    
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
        cmd.Parameters.Append cmd.CreateParameter("@CurrentUserId", adInteger, adParamInput, , .CurrentUserId)
        cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , 1) 'OPEN
        cmd.Execute
        POS_OrderId = Val(cmd.Parameters("@POS_OrderId"))
        
        'SAVE LINE
        Dim item As MSComctlLib.ListItem
        For Each item In .lvList.ListItems
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "POS_OrderLine_Insert"
            
            cmd.Parameters.Append cmd.CreateParameter("@POS_OrderLineId", adInteger, adParamInputOutput, , NVAL(item.SubItems(22)))
            cmd.Parameters.Append cmd.CreateParameter("@POS_OrderId", adInteger, adParamInput, , POS_OrderId)
            cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, item.Text)
            cmd.Parameters.Append cmd.CreateParameter("@Quantity", adDecimal, adParamInput, , NVAL(item.SubItems(1)))
                                  cmd.Parameters("@Quantity").Precision = 18
                                  cmd.Parameters("@Quantity").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@Unit", adVarChar, adParamInput, 50, item.SubItems(2))
            cmd.Parameters.Append cmd.CreateParameter("@Price", adDecimal, adParamInput, , NVAL(item.SubItems(3)))
                                  cmd.Parameters("@Price").Precision = 18
                                  cmd.Parameters("@Price").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@Discount", adDecimal, adParamInput, , NVAL(item.SubItems(4)))
                                  cmd.Parameters("@Discount").Precision = 18
                                  cmd.Parameters("@Discount").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@Subtotal", adDecimal, adParamInput, , NVAL(item.SubItems(5)))
                                  cmd.Parameters("@Subtotal").Precision = 18
                                  cmd.Parameters("@Subtotal").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@UnitCost", adDecimal, adParamInput, , NVAL(item.SubItems(6)))
                                  cmd.Parameters("@UnitCost").Precision = 18
                                  cmd.Parameters("@UnitCost").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@HiddenQuantity", adDecimal, adParamInput, , NVAL(item.SubItems(7)))
                                  cmd.Parameters("@HiddenQuantity").Precision = 18
                                  cmd.Parameters("@HiddenQuantity").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , NVAL(item.SubItems(8)))
            cmd.Parameters.Append cmd.CreateParameter("@HiddenPrice", adDecimal, adParamInput, , NVAL(item.SubItems(9)))
                                  cmd.Parameters("@HiddenPrice").Precision = 18
                                  cmd.Parameters("@HiddenPrice").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@Tax", adDecimal, adParamInput, , NVAL(item.SubItems(13)))
                                  cmd.Parameters("@Tax").Precision = 18
                                  cmd.Parameters("@Tax").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@TaxComputation", adDecimal, adParamInput, , NVAL(item.SubItems(14)))
                                  cmd.Parameters("@TaxComputation").Precision = 18
                                  cmd.Parameters("@TaxComputation").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@DiscountType", adVarChar, adParamInput, 50, item.SubItems(15))
            cmd.Parameters.Append cmd.CreateParameter("@DeductInventory", adDecimal, adParamInput, , NVAL(item.SubItems(16)))
                                  cmd.Parameters("@DeductInventory").Precision = 18
                                  cmd.Parameters("@DeductInventory").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@Discounted", adDecimal, adParamInput, , NVAL(item.SubItems(17)))
                                  cmd.Parameters("@Discounted").Precision = 18
                                  cmd.Parameters("@Discounted").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@ReserveId", adInteger, adParamInput, , NVAL(item.SubItems(18)))
            cmd.Parameters.Append cmd.CreateParameter("@DiscountPercent", adDecimal, adParamInput, , NVAL(item.SubItems(19)))
                                  cmd.Parameters("@DiscountPercent").Precision = 18
                                  cmd.Parameters("@DiscountPercent").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@isTaxExempt", adVarChar, adParamInput, 50, item.SubItems(20))
            cmd.Parameters.Append cmd.CreateParameter("@ActualQuantity", adDecimal, adParamInput, , NVAL(item.SubItems(16)))
                                  cmd.Parameters("@ActualQuantity").Precision = 18
                                  cmd.Parameters("@ActualQuantity").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@Percentage", adDecimal, adParamInput, , NVAL(item.SubItems(13)))
                                  cmd.Parameters("@Percentage").Precision = 18
                                  cmd.Parameters("@Percentage").NumericScale = 2
            cmd.Execute
            item.SubItems(22) = cmd.Parameters("@POS_OrderLineId")
            item.SubItems(21) = POS_OrderId
        Next
    End With
    
    con.CommitTrans
    con.Close
    
    'PRINT
    Dim x As Variant
    x = MsgBox("Do you want to print a receipt?", vbQuestion + vbYesNo)
    If x = vbYes Then
        '**PRINT RECEIPT******
        Dim crxApp As New CRAXDRT.Application
        Dim crxRpt As New CRAXDRT.Report
        Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\POS_OrderReceipt.rpt")
        'crxRpt.RecordSelectionFormula = "{POS_Sales.POS_SalesId}= " & Val(POS_SalesId) & ""
        crxRpt.DiscardSavedData
        crxRpt.EnableParameterPrompting = False
        'crxRpt.ParameterFields.GetItemByName("Notice").AddCurrentValue ""
        crxRpt.ParameterFields.GetItemByName("@POS_OrderId").AddCurrentValue Val(POS_OrderId)
    
        Call ResetRptDB(crxRpt)
        crxRpt.PrintOut False
        '*** END PRINT
    End If
    
    POS_CashierFrm.Initialize
    If POS_PayFrm.Visible = True Then Unload POS_PayFrm
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnPayCash_Click()
    If DualPharmacyMode = "TRUE" Then
        POS_PayFrm.lblAmountDue.Caption = POS_CashierFrm.txtTotal.Caption
        POS_PayFrm.Show
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            btnAccept_Click
        Case vbKeyEscape
            btnCancel_Click
        Case vbKeyF1
            btnPayCash_Click
    End Select
End Sub

Private Sub Form_Load()
    If DualPharmacyMode = "TRUE" Then
        btnPayCash.Visible = True
    Else
        btnPayCash.Visible = False
    End If
End Sub
