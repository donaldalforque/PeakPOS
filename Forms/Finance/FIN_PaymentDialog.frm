VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FIN_PaymentDialogFrmBAK 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   8415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
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
         Left            =   2040
         TabIndex        =   10
         Top             =   7920
         Width           =   1215
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
         Left            =   3360
         TabIndex        =   11
         Top             =   7920
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   6015
         Left            =   120
         TabIndex        =   13
         Top             =   1800
         Width           =   4455
         Begin VB.TextBox Text1 
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
            TabIndex        =   30
            Top             =   1215
            Width           =   2775
         End
         Begin VB.ComboBox cmbFund 
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
            TabIndex        =   6
            Top             =   3000
            Width           =   2775
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
            Width           =   2775
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
            Width           =   2775
         End
         Begin VB.ComboBox cmbCashDirection 
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
            ItemData        =   "FIN_PaymentDialog.frx":0000
            Left            =   1560
            List            =   "FIN_PaymentDialog.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   2520
            Width           =   2175
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
            Height          =   1410
            Left            =   1560
            MultiLine       =   -1  'True
            TabIndex        =   9
            Top             =   4440
            Width           =   2775
         End
         Begin VB.TextBox txtChequeNumber 
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
            Top             =   1575
            Width           =   2175
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
            TabIndex        =   2
            Top             =   960
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.TextBox txtAmount 
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
            TabIndex        =   1
            Top             =   240
            Width           =   2775
         End
         Begin MSComCtl2.DTPicker dtChequeDate 
            Height          =   345
            Left            =   1560
            TabIndex        =   4
            Top             =   1920
            Width           =   2175
            _ExtentX        =   3836
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
            Format          =   97517569
            CurrentDate     =   41646
         End
         Begin MSComCtl2.DTPicker dtDate 
            Height          =   345
            Left            =   1560
            TabIndex        =   32
            Top             =   600
            Width           =   2175
            _ExtentX        =   3836
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
            Format          =   97517569
            CurrentDate     =   41646
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
            TabIndex        =   31
            Top             =   1200
            Width           =   1050
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fund"
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
            TabIndex        =   29
            Top             =   3000
            Width           =   465
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
            TabIndex        =   28
            Top             =   3720
            Width           =   750
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
            TabIndex        =   27
            Top             =   3360
            Width           =   450
         End
         Begin VB.Label lblCashLocation 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Save to"
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
            Top             =   2520
            Width           =   675
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
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
            TabIndex        =   25
            Top             =   4440
            Width           =   795
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cheque Date"
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
            TabIndex        =   24
            Top             =   1920
            Width           =   1200
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
            TabIndex        =   23
            Top             =   600
            Width           =   435
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cheque #"
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
            Top             =   1560
            Width           =   870
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
            TabIndex        =   20
            Top             =   240
            Width           =   435
         End
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
         TabIndex        =   19
         Top             =   1440
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
         TabIndex        =   18
         Top             =   1080
         Width           =   2865
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
         TabIndex        =   17
         Top             =   720
         Width           =   2865
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
         TabIndex        =   16
         Top             =   1440
         Width           =   720
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
         TabIndex        =   15
         Top             =   1080
         Width           =   690
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
         TabIndex        =   14
         Top             =   720
         Width           =   900
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
         TabIndex        =   12
         Top             =   200
         Width           =   1140
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "FIN_PaymentDialog.frx":001B
         Top             =   120
         Width           =   480
      End
   End
End
Attribute VB_Name = "FIN_PaymentDialogFrmBAK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public isAR, isCash, isOS, isAP, isAPOOrder, isSOOrder As Boolean
Public Sub Populate(ByVal data As String)
    Select Case data
        Case "Type"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("PaymentType")
            cmbType.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    cmbType.AddItem rec!Type
                    cmbType.ItemData(cmbType.NewIndex) = rec!PaymentTypeId
                    rec.MoveNext
                Loop
            End If
            cmbType.ListIndex = 0
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
        Case "Fund"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Fund")
            cmbFund.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    cmbFund.AddItem rec!FundName
                    cmbFund.ItemData(cmbFund.NewIndex) = rec!FundId
                    rec.MoveNext
                Loop
            End If
            On Error Resume Next
            cmbFund.ListIndex = 0
    End Select
End Sub
Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnSave_Click()
    
End Sub

Private Sub btnOK_Click()
    'On Error Resume Next
    If cmbType.text = UCase("CHECK") Then
        If Trim(txtChequeNumber.text) = "" Then
            GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(32)
            GLOBAL_MessageFrm.Show (1)
            Exit Sub
        End If
    End If
    
    If cmbCashDirection.text = UCase("bank") Then
        If cmbAccount.text = "" Then
            GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(16)
            GLOBAL_MessageFrm.Show (1)
            Exit Sub
        End If
    Else
        If cmbFund.text = "" Then
            GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(20)
            GLOBAL_MessageFrm.Show (1)
            Exit Sub
        End If
    End If

    If Trim(txtAmount.text) = "" Then txtAmount.text = "0"
    If isAR = True Then
        'Data Transfer to Accounts Receivable
        With FIN_AccountsReceivable
            .lvSearch.SelectedItem.SubItems(10) = FormatNumber(txtAmount.text, 2, vbTrue, vbFalse)
            .lvSearch.SelectedItem.SubItems(11) = cmbType.text
            .lvSearch.SelectedItem.SubItems(12) = txtRemarks.text
            .lvSearch.SelectedItem.SubItems(13) = dtDate.value
            .lvSearch.SelectedItem.SubItems(14) = txtChequeNumber.text
            If Trim(txtChequeNumber.text) <> "" Then .lvSearch.SelectedItem.SubItems(15) = dtChequeDate.value
            .lvSearch.SelectedItem.SubItems(16) = cmbCashDirection.text
            .lvSearch.SelectedItem.SubItems(17) = cmbBank.ItemData(cmbBank.ListIndex)
            If cmbAccount.text = "" Then
                .lvSearch.SelectedItem.SubItems(18) = ""
            Else
                .lvSearch.SelectedItem.SubItems(18) = cmbAccount.ItemData(cmbAccount.ListIndex)
            End If
                .lvSearch.SelectedItem.SubItems(19) = cmbFund.ItemData(cmbFund.ListIndex)
            
        End With
        isAR = False
    ElseIf isAP = True Then
        'Data Transfer to Accounts Payable
        With FIN_AccountsPayable
            .lvSearch.SelectedItem.SubItems(9) = FormatNumber(txtAmount.text, 2, vbTrue, vbFalse)
            .lvSearch.SelectedItem.SubItems(10) = cmbType.text
            .lvSearch.SelectedItem.SubItems(11) = txtRemarks.text
            .lvSearch.SelectedItem.SubItems(12) = dtDate.value
            .lvSearch.SelectedItem.SubItems(13) = txtChequeNumber.text
            If Trim(txtChequeNumber.text) <> "" Then .lvSearch.SelectedItem.SubItems(14) = dtChequeDate.value
            .lvSearch.SelectedItem.SubItems(15) = cmbCashDirection.text
            .lvSearch.SelectedItem.SubItems(16) = cmbBank.ItemData(cmbBank.ListIndex)
            If cmbAccount.text = "" Then
                .lvSearch.SelectedItem.SubItems(18) = ""
            Else
                .lvSearch.SelectedItem.SubItems(18) = cmbAccount.ItemData(cmbAccount.ListIndex)
            End If
                .lvSearch.SelectedItem.SubItems(19) = cmbFund.ItemData(cmbFund.ListIndex)
        End With
        isAP = False
    ElseIf isCash = True Then
        Set con = New ADODB.Connection
        Set cmd = New ADODB.Command
        Set rec = New ADODB.Recordset
    
        con.ConnectionString = ConnString
        con.Open
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "PO_Payment_Insert"
        cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , PO_PurchaseOrderFrm.PurchaseOrderId)
        cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Val(Replace(txtAmount.text, ",", "")))
                              cmd.Parameters("@Amount").Precision = 18
                              cmd.Parameters("@Amount").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtDate.value)
        cmd.Parameters.Append cmd.CreateParameter("@PaymentType", adVarChar, adParamInput, 250, cmbType.text)
        cmd.Parameters.Append cmd.CreateParameter("@ChequeNumber", adVarChar, adParamInput, 250, txtChequeNumber.text)
        If cmbType.text = UCase("CHECK") And cmbCashDirection.text = UCase("BANK") Then
            cmd.Parameters.Append cmd.CreateParameter("@ChequeDate", adDate, adParamInput, , dtChequeDate.value)
            cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , cmbAccount.ItemData(cmbAccount.ListIndex))
            cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , Null)
        ElseIf cmbType.text = UCase("CHECK") And cmbCashDirection.text = UCase("FUNDS") Then
            cmd.Parameters.Append cmd.CreateParameter("@ChequeDate", adDate, adParamInput, , dtChequeDate.value)
            cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , cmbFund.ItemData(cmbFund.ListIndex))
        Else
            cmd.Parameters.Append cmd.CreateParameter("@ChequeDate", adDate, adParamInput, , Null)
            If cmbCashDirection.text = UCase("FUNDS") Then
                cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , Null)
                cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , cmbFund.ItemData(cmbFund.ListIndex))
            Else
                cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , cmbAccount.ItemData(cmbAccount.ListIndex))
                cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , Null)
            End If
        End If
        
        cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 500, txtRemarks.text)
        cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInput, 250, PO_PurchaseOrderFrm.txtOrderNumber.text)
        cmd.Execute
        con.Close
    
        'Change to PAID UI View
        
        With PO_PurchaseOrderFrm
            If txtAmount.text = .lblTotal.Caption Then
                .isNotCompleted (False)
                '.picPaid.Left = 6890
                '.picCompleted.Left = -9999
                '.pic_Cancelled.Left = -9999
                '.picPaid.Visible = True
                '.picCompleted.Visible = False
                '.pic_Cancelled.Visible = False
                .tb_Standard.Buttons(4).Caption = "Cancel"
                .tb_Standard.Buttons(4).Image = 3
                .btnPaid.Visible = False
                '.btnStatus.Caption = "Reopen Order"
            End If
        End With
        isCash = False
    ElseIf isOS = True Then
        Set con = New ADODB.Connection
        Set cmd = New ADODB.Command
        Set rec = New ADODB.Recordset
    
        con.ConnectionString = ConnString
        con.Open
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "SO_Payment_Insert"
        cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , SO_SalesOrderFrm.SalesOrderId)
        cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Val(Replace(txtAmount.text, ",", "")))
                              cmd.Parameters("@Amount").Precision = 18
                              cmd.Parameters("@Amount").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtDate.value)
        cmd.Parameters.Append cmd.CreateParameter("@PaymentType", adVarChar, adParamInput, 250, cmbType.text)
        cmd.Parameters.Append cmd.CreateParameter("@ChequeNumber", adVarChar, adParamInput, 250, txtChequeNumber.text)
        If cmbType.text = UCase("CHECK") And cmbCashDirection.text = UCase("BANK") Then
            cmd.Parameters.Append cmd.CreateParameter("@ChequeDate", adDate, adParamInput, , dtChequeDate.value)
            cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , cmbAccount.ItemData(cmbAccount.ListIndex))
            cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , Null)
        ElseIf cmbType.text = UCase("CHECK") And cmbCashDirection.text = UCase("FUNDS") Then
            cmd.Parameters.Append cmd.CreateParameter("@ChequeDate", adDate, adParamInput, , dtChequeDate.value)
            cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , cmbFund.ItemData(cmbFund.ListIndex))
        Else
            cmd.Parameters.Append cmd.CreateParameter("@ChequeDate", adDate, adParamInput, , Null)
            If cmbCashDirection.text = UCase("FUNDS") Then
                cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , Null)
                cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , cmbFund.ItemData(cmbFund.ListIndex))
            Else
                cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , cmbAccount.ItemData(cmbAccount.ListIndex))
                cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , Null)
            End If
        End If
        
        cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 500, txtRemarks.text)
        cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInput, 250, SO_SalesOrderFrm.txtOrderNumber.text)
        cmd.Execute
        con.Close
    
        'Change to PAID UI View
        With SO_SalesOrderFrm
            If txtAmount.text = .lblRemainingBalance.Caption Then
                .isNotCompleted (False)
                '.picPaid.Left = 6890
                '.picCompleted.Left = -9999
                '.pic_Cancelled.Left = -9999
                '.picPaid.Visible = True
                '.picCompleted.Visible = False
                '.pic_Cancelled.Visible = False
                .lblRemainingBalance.Caption = "0.00"
                .tb_Standard.Buttons(4).Caption = "Cancel"
                .tb_Standard.Buttons(4).Image = 3
                .btnPaid.Visible = False
                .btnStatus.Caption = "Reopen Order"
            End If
        End With
        isOS = False
    End If
    Unload Me
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
'    On Error Resume Next
'    cmbAccount.ListIndex = 0
End Sub

Private Sub cmbCashDirection_Click()
'    If cmbCashDirection.text = "FUNDS" Then
'        cmbFund.Enabled = True
'        cmbBank.Enabled = False
'        cmbAccount.Enabled = False
'        cmbBank.BackColor = &HC0C0C0
'        cmbAccount.BackColor = &HC0C0C0
'        cmbFund.BackColor = vbWhite
'    Else
'        cmbFund.Enabled = False
'        cmbBank.Enabled = True
'        cmbAccount.Enabled = True
'        cmbBank.BackColor = vbWhite
'        cmbAccount.BackColor = vbWhite
'        cmbFund.BackColor = &HC0C0C0
'    End If
End Sub

Private Sub cmbType_Click()
    If cmbType.text = "CASH" Then
        txtChequeNumber.Enabled = False
        txtChequeNumber.BackColor = &HC0C0C0
        dtChequeDate.Enabled = False
    Else
        txtChequeNumber.Enabled = True
        txtChequeNumber.BackColor = vbWhite
        dtChequeDate.Enabled = True
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            btnOK_Click
        Case vbKeyEscape
            btnCancel_Click
    End Select
End Sub

Private Sub Form_Load()
    Populate "Type"
    Populate "Bank"
    Populate "Fund"
    cmbCashDirection.ListIndex = 0
    dtDate.value = Format(Now, "MM/DD/YY")
    dtChequeDate.value = Format(Now, "MM/DD/YY")
End Sub

Private Sub txtAmount_Change()
'    If IsNumeric(txtAmount.text) = False Then
'        txtAmount.text = "0"
'    Else
'        If Val(Replace(txtAmount.text, ",", "")) > Val(Replace(lblBalance.Caption, ",", "")) Then
'            txtAmount.text = lblBalance.Caption
'        End If
'    End If
End Sub

Private Sub txtAmount_GotFocus()
    selectText txtAmount
End Sub

Private Sub txtAmount_LostFocus()
    If IsNumeric(txtAmount.text) = False Then
        txtAmount.text = "0"
    Else
        If Val(Replace(txtAmount.text, ",", "")) > Val(Replace(lblBalance.Caption, ",", "")) Then
            txtAmount.text = lblBalance.Caption
        End If
    End If
End Sub
