VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FIN_PaymentDialogFrm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   6015
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   4815
         Left            =   120
         TabIndex        =   12
         Top             =   2160
         Width           =   5535
         Begin VB.CheckBox chkOnline 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Online Payment"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1560
            TabIndex        =   35
            Top             =   240
            Visible         =   0   'False
            Width           =   3855
         End
         Begin VB.TextBox txtSalesReturn 
            Alignment       =   1  'Right Justify
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
            Height          =   330
            Left            =   1560
            TabIndex        =   6
            Top             =   2415
            Width           =   3855
         End
         Begin VB.TextBox txtCash 
            Alignment       =   1  'Right Justify
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
            Left            =   1560
            TabIndex        =   2
            Text            =   "0.00"
            Top             =   960
            Width           =   3855
         End
         Begin VB.ComboBox cmbType 
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
            Left            =   -9999
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   960
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.TextBox txtCheckNumber 
            Alignment       =   1  'Right Justify
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
            Height          =   330
            Left            =   1560
            TabIndex        =   4
            Top             =   1695
            Width           =   3855
         End
         Begin VB.TextBox txtRemarks 
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
            Left            =   1560
            MultiLine       =   -1  'True
            TabIndex        =   9
            Top             =   4320
            Width           =   3855
         End
         Begin VB.ComboBox cmbBank 
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
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   3360
            Width           =   3855
         End
         Begin VB.ComboBox cmbAccount 
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
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   3720
            Width           =   3855
         End
         Begin VB.TextBox txtCheckAmount 
            Alignment       =   1  'Right Justify
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
            Height          =   330
            Left            =   1560
            TabIndex        =   3
            Text            =   "0.00"
            Top             =   1335
            Width           =   3855
         End
         Begin MSComCtl2.DTPicker dtCheckDate 
            Height          =   345
            Left            =   1560
            TabIndex        =   5
            Top             =   2040
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   84017153
            CurrentDate     =   41646
         End
         Begin MSComCtl2.DTPicker dtDate 
            Height          =   345
            Left            =   1560
            TabIndex        =   1
            Top             =   600
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   84017153
            CurrentDate     =   41646
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Discount"
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
            TabIndex        =   31
            Top             =   2400
            Width           =   1335
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "[For Check Issuance, deduct from/Save to]"
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
            TabIndex        =   30
            Top             =   2880
            Width           =   3900
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cash"
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
            TabIndex        =   22
            Top             =   960
            Width           =   435
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Method"
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
            Left            =   -9999
            TabIndex        =   21
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Check #"
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
            TabIndex        =   20
            Top             =   1680
            Width           =   705
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
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
            TabIndex        =   19
            Top             =   600
            Width           =   435
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Check Date"
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
            TabIndex        =   18
            Top             =   2040
            Width           =   1035
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Particulars"
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
            TabIndex        =   17
            Top             =   4320
            Width           =   975
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank"
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
            TabIndex        =   16
            Top             =   3360
            Width           =   450
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account"
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
            TabIndex        =   15
            Top             =   3720
            Width           =   750
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Check Amt."
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
            TabIndex        =   14
            Top             =   1320
            Width           =   1050
         End
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
         Left            =   4440
         TabIndex        =   11
         Top             =   7080
         Width           =   1215
      End
      Begin VB.CommandButton btnOK 
         Caption         =   "OK"
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
         Left            =   3120
         TabIndex        =   10
         Top             =   7080
         Width           =   1215
      End
      Begin VB.Label lblPreview 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   5520
         TabIndex        =   34
         Top             =   1800
         Width           =   45
      End
      Begin VB.Label lblBalanceForwarded 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   1680
         TabIndex        =   33
         Top             =   1800
         Width           =   2865
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Balance Fwd"
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
         TabIndex        =   32
         Top             =   1800
         Width           =   1155
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "FIN_PaymentDialogFrm1.frx":0000
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "PAYMENTS"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   29
         Top             =   200
         Width           =   1140
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
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
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
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
         Left            =   120
         TabIndex        =   27
         Top             =   1080
         Width           =   690
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Balance"
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
         TabIndex        =   26
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label lblCustomerName 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1680
         TabIndex        =   25
         Top             =   720
         Width           =   2865
      End
      Begin VB.Label lblOrderNumber 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1680
         TabIndex        =   24
         Top             =   1080
         Width           =   2865
      End
      Begin VB.Label lblBalance 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   1680
         TabIndex        =   23
         Top             =   1440
         Width           =   2865
      End
   End
End
Attribute VB_Name = "FIN_PaymentDialogFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public isPO, isSO, isSOAR, isBatchPO, isBatchSO, isOnline As Boolean
Public Sub Populate(ByVal data As String)
    Select Case data
        Case "Bank"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Bank")
            cmbBank.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        cmbBank.AddItem rec!bankname
                        cmbBank.ItemData(cmbBank.NewIndex) = rec!BankId
                    End If
                    rec.MoveNext
                Loop
            End If
            On Error Resume Next
            cmbBank.ListIndex = 0
    End Select
End Sub
Private Sub CountTotal()
    Dim balance, salesreturn, checkamount, cash As Double
    balance = Val(Replace(lblBalance.Caption, ",", ""))
    salesreturn = Val(Replace(txtSalesReturn.text, ",", ""))
    checkamount = Val(Replace(txtCheckAmount.text, ",", ""))
    cash = Val(Replace(txtCash.text, ",", ""))
    balance = balance - (salesreturn + cash + checkamount)
    lblBalanceForwarded.Caption = FormatNumber(balance, 2, vbTrue, vbFalse)
End Sub
Private Function Validated() As Boolean
    Dim balance, salesreturn, checkamount, cash As Double
    salesreturn = Val(Replace(txtSalesReturn.text, ",", ""))
    checkamount = Val(Replace(txtCheckAmount.text, ",", ""))
    cash = Val(Replace(txtCash.text, ",", ""))
    balance = (salesreturn + cash + checkamount)
    If balance = 0 Then
        Validated = False
    Else
        Validated = True
    End If
End Function
Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOk_Click()
    If Validated = False Then
        Unload Me
        Exit Sub
    End If
    
    Set con = New ADODB.Connection
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    
    If isPO = True Then
        If Val(Replace(txtCheckAmount.text, ",", "")) > 0 Or Trim(txtCheckNumber.text) <> "" Then
            If cmbAccount.text = "" Then
                GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(21)
                GLOBAL_MessageFrm.Show (1)
                cmbAccount.SetFocus
                Exit Sub
            End If
        End If
    
        cmd.CommandText = "PO_Payment_Insert"
        cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , PO_PurchaseOrderFrm.PurchaseOrderId)
        cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Val(Replace(txtCash.text, ",", "")))
                              cmd.Parameters("@Amount").NumericScale = 2
                              cmd.Parameters("@Amount").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtDate.value)
        cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.text, ",", "")))
                              cmd.Parameters("@CheckAmount").NumericScale = 2
                              cmd.Parameters("@CheckAmount").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@SalesReturn", adDecimal, adParamInput, , Val(Replace(txtSalesReturn.text, ",", "")))
                              cmd.Parameters("@SalesReturn").NumericScale = 2
                              cmd.Parameters("@SalesReturn").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.text)
        cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
        If Val(Replace(txtCheckAmount.text, ",", "")) > 0 Or Trim(txtCheckNumber.text) <> "" Then
            cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , cmbAccount.ItemData(cmbAccount.ListIndex))
            cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , Null)
        Else
            cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , 1)
        End If
        cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, txtRemarks.text)
        cmd.Parameters.Append cmd.CreateParameter("@POPaymentId", adInteger, adParamInputOutput, , 0)
        cmd.Execute
        
        Dim POPaymentId As Integer
        POPaymentId = cmd.Parameters("@POPaymentId")
        
        'BANK OUTFLOW
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "FIN_BankOutFlow_Insert"
        cmd.Parameters.Append cmd.CreateParameter("@Cash", adDecimal, adParamInput, , 0)
                              cmd.Parameters("@Cash").Precision = 18
                              cmd.Parameters("@Cash").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.text, ",", "")))
                              cmd.Parameters("@CheckAmount").Precision = 18
                              cmd.Parameters("@CheckAmount").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.text)
        cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
        cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , Format(Now, "MM/DD/YY"))
        cmd.Parameters.Append cmd.CreateParameter("@Particulars", adVarChar, adParamInput, 250, "Purchase Order: " & PO_PurchaseOrderFrm.txtOrderNumber.text & " " & txtRemarks.text)
        cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , PO_PurchaseOrderFrm.PurchaseOrderId)
        cmd.Parameters.Append cmd.CreateParameter("@ExpenseId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@POS_SalesId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@POPaymentId", adInteger, adParamInput, , POPaymentId)
        cmd.Execute
        
        'INFLOW AND OUTFLOW
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "FIN_OutFlow_Insert"
        cmd.Parameters.Append cmd.CreateParameter("@Cash", adDecimal, adParamInput, , Val(Replace(txtCash.text, ",", "")))
                              cmd.Parameters("@Cash").Precision = 18
                              cmd.Parameters("@Cash").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.text, ",", "")))
                              cmd.Parameters("@CheckAmount").Precision = 18
                              cmd.Parameters("@CheckAmount").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.text)
        cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
        cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , Format(Now, "MM/DD/YY"))
        cmd.Parameters.Append cmd.CreateParameter("@Particulars", adVarChar, adParamInput, 250, "Purchase Order: " & PO_PurchaseOrderFrm.txtOrderNumber.text & " " & txtRemarks.text)
        cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , PO_PurchaseOrderFrm.PurchaseOrderId)
        cmd.Parameters.Append cmd.CreateParameter("@ExpenseId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@POS_SalesId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@POPaymentId", adInteger, adParamInput, , POPaymentId)
        cmd.Execute
        
        'CHECK REGISTRY
        If Trim(txtCheckNumber.text) <> "" Or Val(Replace(txtCheckAmount.text, ",", "")) > 0 Then
            'Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            'con.ConnectionString = ConnString
            'con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "FIN_CheckRegistry_Insert"
            cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.text)
            cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
            cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.text, ",", "")))
                                  cmd.Parameters("@Amount").NumericScale = 2
                                  cmd.Parameters("@Amount").Precision = 18
            cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, txtRemarks.text)
            cmd.Parameters.Append cmd.CreateParameter("@isReceivable", adBoolean, adParamInput, , "False")
            cmd.Parameters.Append cmd.CreateParameter("@CheckStatusId", adInteger, adParamInput, , 1) '--open
            cmd.Parameters.Append cmd.CreateParameter("@CheckRegistryId", adInteger, adParamInputOutput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , PO_PurchaseOrderFrm.PurchaseOrderId)
            cmd.Parameters.Append cmd.CreateParameter("@ExpenseId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@POS_SalesId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@SOPaymentId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@POPaymentId", adInteger, adParamInput, , POPaymentId)
            cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , cmbAccount.ItemData(cmbAccount.ListIndex))
            
            cmd.Execute
        End If
        PO_PurchaseOrderFrm.lblRemainingBalance.Caption = lblBalanceForwarded.Caption
    ElseIf isSO = True Then
        If isOnline = True Then
            If cmbAccount.text = "" Then
                GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(21)
                GLOBAL_MessageFrm.Show (1)
                cmbAccount.SetFocus
                Exit Sub
            End If
        End If
    
        cmd.CommandText = "SO_Payment_Insert"
        cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SO_CashAdvance.SalesOrderId)
        cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Val(Replace(txtCash.text, ",", "")))
                              cmd.Parameters("@Amount").NumericScale = 2
                              cmd.Parameters("@Amount").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtDate.value)
        cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.text, ",", "")))
                              cmd.Parameters("@CheckAmount").NumericScale = 2
                              cmd.Parameters("@CheckAmount").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@SalesReturn", adDecimal, adParamInput, , Val(Replace(txtSalesReturn.text, ",", "")))
                              cmd.Parameters("@SalesReturn").NumericScale = 2
                              cmd.Parameters("@SalesReturn").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.text)
        cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
        If isOnline = True Then
            cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , cmbAccount.ItemData(cmbAccount.ListIndex))
            cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , Null)
        Else
            cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , 1)
        End If
        cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, txtRemarks.text)
        cmd.Parameters.Append cmd.CreateParameter("@isOnline", adBoolean, adParamInput, , isOnline)
        cmd.Parameters.Append cmd.CreateParameter("@SOPaymentId", adInteger, adParamInputOutput, , 0)
        cmd.Execute
        
        Dim SOPaymentId As Integer
        SOPaymentId = cmd.Parameters("@SOPaymentId")
        
        'INFLOW AND OUTFLOW
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "FIN_InFlow_Insert"
        cmd.Parameters.Append cmd.CreateParameter("@Cash", adDecimal, adParamInput, , Val(Replace(txtCash.text, ",", "")))
                              cmd.Parameters("@Cash").Precision = 18
                              cmd.Parameters("@Cash").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.text, ",", "")))
                              cmd.Parameters("@CheckAmount").Precision = 18
                              cmd.Parameters("@CheckAmount").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.text)
        cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
        cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , Format(Now, "MM/DD/YY"))
        cmd.Parameters.Append cmd.CreateParameter("@Particulars", adVarChar, adParamInput, 250, "Payment: " & SO_CashAdvance.txtOrderNumber.text & " " & txtRemarks.text)
        cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SO_CashAdvance.SalesOrderId)
        cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@ExpenseId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@POS_SalesId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@SOPaymentId", adInteger, adParamInput, , SOPaymentId)
        cmd.Execute
        
        'CUSTOMER LEDGER
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "FIN_CustomerLedger_Insert"
        cmd.Parameters.Append cmd.CreateParameter("@SOPaymentId", adInteger, adParamInput, , SOPaymentId)
        cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInput, , SO_CashAdvance.CustomerId)
        cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SO_CashAdvance.SalesOrderId)
        cmd.Parameters.Append cmd.CreateParameter("@Particulars", adVarChar, adParamInput, 250, "Payment: " & SO_CashAdvance.txtOrderNumber.text & " " & txtRemarks.text)
        cmd.Parameters.Append cmd.CreateParameter("@Debit", adDecimal, adParamInput, , Null)
                              cmd.Parameters("@Debit").NumericScale = 2
                              cmd.Parameters("@Debit").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@Credit", adDecimal, adParamInput, , (Val(Replace(txtCash.text, ",", "")) + Val(Replace(txtCheckAmount.text, ",", ""))))
                              cmd.Parameters("@Credit").NumericScale = 2
                              cmd.Parameters("@Credit").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtDate.value)
        cmd.Execute
        
        'CHECK REGISTRY
        If (Trim(txtCheckNumber.text) <> "" Or Val(Replace(txtCheckAmount.text, ",", "")) > 0) And isOnline = False Then
            'Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            'con.ConnectionString = ConnString
            'con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "FIN_CheckRegistry_Insert"
            cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.text)
            cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
            cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.text, ",", "")))
                                  cmd.Parameters("@Amount").NumericScale = 2
                                  cmd.Parameters("@Amount").Precision = 18
            cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, txtRemarks.text)
            cmd.Parameters.Append cmd.CreateParameter("@isReceivable", adBoolean, adParamInput, , "True")
            cmd.Parameters.Append cmd.CreateParameter("@CheckStatusId", adInteger, adParamInput, , 1) '--open
            cmd.Parameters.Append cmd.CreateParameter("@CheckRegistryId", adInteger, adParamInputOutput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SO_CashAdvance.SalesOrderId)
            cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@ExpenseId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@POS_SalesId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@SOPaymentId", adInteger, adParamInput, , SOPaymentId)
            cmd.Parameters.Append cmd.CreateParameter("@POPaymentId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , 1)
            
            cmd.Execute
        End If
        SO_CashAdvance.lblRemainingBalance.Caption = lblBalanceForwarded.Caption
    ElseIf isSOAR = True Then
        If isOnline = True Then
            If cmbAccount.text = "" Then
                GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(21)
                GLOBAL_MessageFrm.Show (1)
                cmbAccount.SetFocus
                Exit Sub
            End If
        End If
    
        cmd.CommandText = "SO_Payment_Insert"
        cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SO_SalesOrderFrm.SalesOrderId)
        cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Val(Replace(txtCash.text, ",", "")))
                              cmd.Parameters("@Amount").NumericScale = 2
                              cmd.Parameters("@Amount").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtDate.value)
        cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.text, ",", "")))
                              cmd.Parameters("@CheckAmount").NumericScale = 2
                              cmd.Parameters("@CheckAmount").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@SalesReturn", adDecimal, adParamInput, , Val(Replace(txtSalesReturn.text, ",", "")))
                              cmd.Parameters("@SalesReturn").NumericScale = 2
                              cmd.Parameters("@SalesReturn").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.text)
        cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
        If isOnline = True Then
            cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , cmbAccount.ItemData(cmbAccount.ListIndex))
            cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , Null)
        Else
            cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , 1)
        End If
        cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, txtRemarks.text)
        cmd.Parameters.Append cmd.CreateParameter("@isOnline", adBoolean, adParamInput, , isOnline)
        cmd.Parameters.Append cmd.CreateParameter("@SOPaymentId", adInteger, adParamInputOutput, , 0)
        cmd.Execute
        
        'Dim SOPaymentId As Integer
        SOPaymentId = cmd.Parameters("@SOPaymentId")
        
        'INFLOW AND OUTFLOW
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "FIN_InFlow_Insert"
        cmd.Parameters.Append cmd.CreateParameter("@Cash", adDecimal, adParamInput, , Val(Replace(txtCash.text, ",", "")))
                              cmd.Parameters("@Cash").Precision = 18
                              cmd.Parameters("@Cash").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.text, ",", "")))
                              cmd.Parameters("@CheckAmount").Precision = 18
                              cmd.Parameters("@CheckAmount").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.text)
        cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
        cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , Format(Now, "MM/DD/YY"))
        cmd.Parameters.Append cmd.CreateParameter("@Particulars", adVarChar, adParamInput, 250, "Payment: " & SO_SalesOrderFrm.txtOrderNumber.text & " " & txtRemarks.text)
        cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SO_SalesOrderFrm.SalesOrderId)
        cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@ExpenseId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@POS_SalesId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@SOPaymentId", adInteger, adParamInput, , SOPaymentId)
        cmd.Execute
        
        'CUSTOMER LEDGER
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.ActiveConnection = con
        cmd.CommandText = "FIN_CustomerLedger_Insert"
        cmd.Parameters.Append cmd.CreateParameter("@SOPaymentId", adInteger, adParamInput, , SOPaymentId)
        cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInput, , SO_SalesOrderFrm.CustomerId)
        cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SO_SalesOrderFrm.SalesOrderId)
        cmd.Parameters.Append cmd.CreateParameter("@Particulars", adVarChar, adParamInput, 250, "Payment: " & SO_SalesOrderFrm.txtOrderNumber.text & " " & txtRemarks.text)
        cmd.Parameters.Append cmd.CreateParameter("@Debit", adDecimal, adParamInput, , Null)
                              cmd.Parameters("@Debit").NumericScale = 2
                              cmd.Parameters("@Debit").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@Credit", adDecimal, adParamInput, , (Val(Replace(txtCash.text, ",", "")) + Val(Replace(txtCheckAmount.text, ",", ""))))
                              cmd.Parameters("@Credit").NumericScale = 2
                              cmd.Parameters("@Credit").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtDate.value)
        cmd.Execute
        
        'CHECK REGISTRY
        If (Trim(txtCheckNumber.text) <> "" Or Val(Replace(txtCheckAmount.text, ",", "")) > 0) And isOnline = False Then
            'Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            'con.ConnectionString = ConnString
            'con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "FIN_CheckRegistry_Insert"
            cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.text)
            cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
            cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.text, ",", "")))
                                  cmd.Parameters("@Amount").NumericScale = 2
                                  cmd.Parameters("@Amount").Precision = 18
            cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, txtRemarks.text)
            cmd.Parameters.Append cmd.CreateParameter("@isReceivable", adBoolean, adParamInput, , "True")
            cmd.Parameters.Append cmd.CreateParameter("@CheckStatusId", adInteger, adParamInput, , 1) '--open
            cmd.Parameters.Append cmd.CreateParameter("@CheckRegistryId", adInteger, adParamInputOutput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SO_SalesOrderFrm.SalesOrderId)
            cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@ExpenseId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@POS_SalesId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@SOPaymentId", adInteger, adParamInput, , SOPaymentId)
            cmd.Parameters.Append cmd.CreateParameter("@POPaymentId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , 1)
            
            cmd.Execute
        End If
        SO_SalesOrderFrm.lblRemainingBalance.Caption = lblBalanceForwarded.Caption
    ElseIf isBatchSO = True Then
        If isOnline = True Then
            If cmbAccount.text = "" Then
                GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(21)
                GLOBAL_MessageFrm.Show (1)
                cmbAccount.SetFocus
                Exit Sub
            End If
        End If
    
        cmd.CommandText = "SO_Payment_Insert"
        cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , FIN_AccountsReceivable.lvSearch.SelectedItem.text)
        cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Val(Replace(txtCash.text, ",", "")))
                              cmd.Parameters("@Amount").NumericScale = 2
                              cmd.Parameters("@Amount").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtDate.value)
        cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.text, ",", "")))
                              cmd.Parameters("@CheckAmount").NumericScale = 2
                              cmd.Parameters("@CheckAmount").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@SalesReturn", adDecimal, adParamInput, , Val(Replace(txtSalesReturn.text, ",", "")))
                              cmd.Parameters("@SalesReturn").NumericScale = 2
                              cmd.Parameters("@SalesReturn").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.text)
        cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
        If isOnline = True Then
            cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , cmbAccount.ItemData(cmbAccount.ListIndex))
            cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , Null)
        Else
            cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , 1)
        End If
        cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, txtRemarks.text)
        cmd.Parameters.Append cmd.CreateParameter("@isOnline", adBoolean, adParamInput, , isOnline)
        cmd.Parameters.Append cmd.CreateParameter("@SOPaymentId", adInteger, adParamInputOutput, , 0)
        cmd.Execute
        
        'Dim SOPaymentId As Integer
        SOPaymentId = cmd.Parameters("@SOPaymentId")
        
        'INFLOW AND OUTFLOW
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "FIN_InFlow_Insert"
        cmd.Parameters.Append cmd.CreateParameter("@Cash", adDecimal, adParamInput, , Val(Replace(txtCash.text, ",", "")))
                              cmd.Parameters("@Cash").Precision = 18
                              cmd.Parameters("@Cash").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.text, ",", "")))
                              cmd.Parameters("@CheckAmount").Precision = 18
                              cmd.Parameters("@CheckAmount").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.text)
        cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
        cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , Format(Now, "MM/DD/YY"))
        cmd.Parameters.Append cmd.CreateParameter("@Particulars", adVarChar, adParamInput, 250, "Payment: " & FIN_AccountsReceivable.lvSearch.SelectedItem.SubItems(3) & " " & txtRemarks.text)
        cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , FIN_AccountsReceivable.lvSearch.SelectedItem.text)
        cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , Null)
        cmd.Execute
        
        'CUSTOMER LEDGER
        Set cmd = New ADODB.Command
        cmd.CommandType = adCmdStoredProc
        cmd.ActiveConnection = con
        cmd.CommandText = "FIN_CustomerLedger_Insert"
        cmd.Parameters.Append cmd.CreateParameter("@SOPaymentId", adInteger, adParamInput, , SOPaymentId)
        cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInput, , FIN_AccountsReceivable.lvSearch.SelectedItem.SubItems(10))
        cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , FIN_AccountsReceivable.lvSearch.SelectedItem.text)
        cmd.Parameters.Append cmd.CreateParameter("@Particulars", adVarChar, adParamInput, 250, "Payment: " & FIN_AccountsReceivable.lvSearch.SelectedItem.SubItems(3) & " " & txtRemarks.text)
        cmd.Parameters.Append cmd.CreateParameter("@Debit", adDecimal, adParamInput, , Null)
                              cmd.Parameters("@Debit").NumericScale = 2
                              cmd.Parameters("@Debit").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@Credit", adDecimal, adParamInput, , (Val(Replace(txtCash.text, ",", "")) + Val(Replace(txtCheckAmount.text, ",", ""))))
                              cmd.Parameters("@Credit").NumericScale = 2
                              cmd.Parameters("@Credit").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtDate.value)
        cmd.Execute
        
        'CHECK REGISTRY
        If (Trim(txtCheckNumber.text) <> "" Or Val(Replace(txtCheckAmount.text, ",", "")) > 0) And isOnline = False Then
            'Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            'con.ConnectionString = ConnString
            'con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "FIN_CheckRegistry_Insert"
            cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.text)
            cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
            cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.text, ",", "")))
                                  cmd.Parameters("@Amount").NumericScale = 2
                                  cmd.Parameters("@Amount").Precision = 18
            cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, txtRemarks.text)
            cmd.Parameters.Append cmd.CreateParameter("@isReceivable", adBoolean, adParamInput, , "True")
            cmd.Parameters.Append cmd.CreateParameter("@CheckStatusId", adInteger, adParamInput, , 1) '--open
            cmd.Parameters.Append cmd.CreateParameter("@CheckRegistryId", adInteger, adParamInputOutput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , FIN_AccountsReceivable.lvSearch.SelectedItem.text)
            cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@ExpenseId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@POS_SalesId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@SOPaymentId", adInteger, adParamInput, , SOPaymentId)
            cmd.Parameters.Append cmd.CreateParameter("@POPaymentId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , 1)
            
            cmd.Execute
        End If
        FIN_AccountsReceivable.lvSearch.SelectedItem.SubItems(9) = lblBalanceForwarded.Caption
    ElseIf isBatchPO = True Then
        If Val(Replace(txtCheckAmount.text, ",", "")) > 0 Or Trim(txtCheckNumber.text) <> "" Then
            If cmbAccount.text = "" Then
                GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(21)
                GLOBAL_MessageFrm.Show (1)
                cmbAccount.SetFocus
                Exit Sub
            End If
        End If
    
        cmd.CommandText = "PO_Payment_Insert"
        cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , FIN_AccountsPayable.lvSearch.SelectedItem.text)
        cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Val(Replace(txtCash.text, ",", "")))
                              cmd.Parameters("@Amount").NumericScale = 2
                              cmd.Parameters("@Amount").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtDate.value)
        cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.text, ",", "")))
                              cmd.Parameters("@CheckAmount").NumericScale = 2
                              cmd.Parameters("@CheckAmount").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@SalesReturn", adDecimal, adParamInput, , Val(Replace(txtSalesReturn.text, ",", "")))
                              cmd.Parameters("@SalesReturn").NumericScale = 2
                              cmd.Parameters("@SalesReturn").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.text)
        cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
        If Val(Replace(txtCheckAmount.text, ",", "")) > 0 Or Trim(txtCheckNumber.text) <> "" Then
            cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , cmbAccount.ItemData(cmbAccount.ListIndex))
            cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , Null)
        Else
            cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , 1)
        End If
        cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, txtRemarks.text)
        cmd.Parameters.Append cmd.CreateParameter("@POPaymentId", adInteger, adParamInputOutput, , 0)
        cmd.Execute
        
        'Dim POPaymentId As Integer
        POPaymentId = cmd.Parameters("@POPaymentId")
        
        'BANK OUTFLOW
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "FIN_BankOutFlow_Insert"
        cmd.Parameters.Append cmd.CreateParameter("@Cash", adDecimal, adParamInput, , 0)
                              cmd.Parameters("@Cash").Precision = 18
                              cmd.Parameters("@Cash").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.text, ",", "")))
                              cmd.Parameters("@CheckAmount").Precision = 18
                              cmd.Parameters("@CheckAmount").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.text)
        cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
        cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , Format(Now, "MM/DD/YY"))
        cmd.Parameters.Append cmd.CreateParameter("@Particulars", adVarChar, adParamInput, 250, "Purchase Order: " & PO_PurchaseOrderFrm.txtOrderNumber.text & " " & txtRemarks.text)
        cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , PO_PurchaseOrderFrm.PurchaseOrderId)
        cmd.Parameters.Append cmd.CreateParameter("@ExpenseId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@POS_SalesId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@POPaymentId", adInteger, adParamInput, , POPaymentId)
        cmd.Execute
        
        'INFLOW AND OUTFLOW
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "FIN_OutFlow_Insert"
        cmd.Parameters.Append cmd.CreateParameter("@Cash", adDecimal, adParamInput, , Val(Replace(txtCash.text, ",", "")))
                              cmd.Parameters("@Cash").Precision = 18
                              cmd.Parameters("@Cash").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.text, ",", "")))
                              cmd.Parameters("@CheckAmount").Precision = 18
                              cmd.Parameters("@CheckAmount").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.text)
        cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
        cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , Format(Now, "MM/DD/YY"))
        cmd.Parameters.Append cmd.CreateParameter("@Particulars", adVarChar, adParamInput, 250, "Purchase Order: " & FIN_AccountsPayable.lvSearch.SelectedItem.SubItems(3) & " " & txtRemarks.text)
        cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , FIN_AccountsPayable.lvSearch.SelectedItem.text)
        cmd.Parameters.Append cmd.CreateParameter("@ExpenseId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@POS_SalesId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@POPaymentId", adInteger, adParamInput, , POPaymentId)
        cmd.Execute
        
        'CHECK REGISTRY
        If Trim(txtCheckNumber.text) <> "" Or Val(Replace(txtCheckAmount.text, ",", "")) > 0 Then
            'Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            'con.ConnectionString = ConnString
            'con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "FIN_CheckRegistry_Insert"
            cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.text)
            cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
            cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.text, ",", "")))
                                  cmd.Parameters("@Amount").NumericScale = 2
                                  cmd.Parameters("@Amount").Precision = 18
            cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, txtRemarks.text)
            cmd.Parameters.Append cmd.CreateParameter("@isReceivable", adBoolean, adParamInput, , "False")
            cmd.Parameters.Append cmd.CreateParameter("@CheckStatusId", adInteger, adParamInput, , 1) '--open
            cmd.Parameters.Append cmd.CreateParameter("@CheckRegistryId", adInteger, adParamInputOutput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , FIN_AccountsPayable.lvSearch.SelectedItem.text)
            cmd.Parameters.Append cmd.CreateParameter("@ExpenseId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@POS_SalesId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@SOPaymentId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@POPaymentId", adInteger, adParamInput, , POPaymentId)
            cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , cmbAccount.ItemData(cmbAccount.ListIndex))
            
            cmd.Execute
        End If
        FIN_AccountsPayable.lvSearch.SelectedItem.SubItems(8) = lblBalanceForwarded.Caption
    End If
    con.Close
    
    MsgBox "Payment successful.", vbInformation, "Success"
    Unload Me
End Sub

Private Sub chkOnline_Click()
    If chkOnline.value = Checked Then
        txtRemarks.text = txtRemarks.text & " (Online)"
        isOnline = True
    Else
        txtRemarks.text = Replace(txtRemarks.text, " (Online)", "")
        isOnline = False
    End If
End Sub

Private Sub cmbBank_Click()
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_BankAccount_Load"
    
    cmd.Parameters.Append cmd.CreateParameter("@BankId", adInteger, adParamInput, , cmbBank.ItemData(cmbBank.ListIndex))
    Set rec = cmd.Execute
    cmbAccount.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            If rec!isActive = "True" Then
                cmbAccount.AddItem rec!accountnumber & " - " & rec!AccountName
                cmbAccount.ItemData(cmbAccount.NewIndex) = rec!AccountId
            End If
            rec.MoveNext
        Loop
    End If
    con.Close
End Sub



Private Sub Form_Load()
    Populate "Bank"
    dtDate.value = Format(Now, "MM/DD/YY")
    dtCheckDate.value = Format(Now, "MM/DD/YY")
    lblBalanceForwarded.Caption = lblBalance.Caption
    
    If chkOnline.value = Checked Then isOnline = True
    If chkOnline.value = Unchecked Then isOnline = False
End Sub


Private Sub txtCash_Change()
    txtCash.text = Replace(txtCash.text, "..", ".")
    If IsNumeric(txtCash.text) = False Then
        txtCash.text = "0.00"
        selectText txtCash
        lblPreview.Caption = ""
    Else
        lblPreview.Caption = FormatNumber(txtCash.text, 2, vbTrue)
        CountTotal
    End If
End Sub


Private Sub txtCash_LostFocus()
    lblPreview.Caption = ""
    txtCash.text = FormatNumber(txtCash.text, 2, vbTrue)
End Sub

Private Sub txtCheckAmount_Change()
    If IsNumeric(txtCheckAmount.text) = False Then
        txtCheckAmount.text = "0.00"
        selectText txtCheckAmount
    Else
        lblPreview.Caption = FormatNumber(txtCheckAmount.text, 2, vbTrue)
        CountTotal
    End If
End Sub

Private Sub txtCheckAmount_LostFocus()
    lblPreview.Caption = ""
    txtCheckAmount.text = FormatNumber(txtCheckAmount.text, 2, vbTrue)
End Sub

Private Sub txtSalesReturn_Change()
    If IsNumeric(txtSalesReturn.text) = False Then
        txtSalesReturn.text = "0.00"
        selectText txtSalesReturn
    Else
        CountTotal
        txtSalesReturn.text = FormatNumber(txtSalesReturn.text, 0)
        txtSalesReturn.SelStart = Len(txtSalesReturn.text)
    End If
End Sub

