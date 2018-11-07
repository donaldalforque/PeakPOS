VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FIN_ExpensesFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   Icon            =   "FIN_ExpensesFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnSave 
      Caption         =   "Save"
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
      Left            =   2400
      TabIndex        =   13
      Top             =   5280
      Width           =   1335
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
      Left            =   3840
      TabIndex        =   12
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.TextBox txtCheckNumber 
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
         Left            =   1800
         TabIndex        =   5
         Top             =   3120
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.TextBox txtCheckAmount 
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
         Left            =   1800
         TabIndex        =   4
         Top             =   2760
         Visible         =   0   'False
         Width           =   3015
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
         Height          =   810
         Left            =   1800
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   4200
         Width           =   3015
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
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   5040
         Visible         =   0   'False
         Width           =   3015
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
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   5400
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.ComboBox cmbExpense 
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
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1680
         Width           =   3015
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
         Left            =   1800
         TabIndex        =   3
         Top             =   2415
         Width           =   3015
      End
      Begin MSComCtl2.DTPicker dtFund 
         Height          =   330
         Left            =   1800
         TabIndex        =   2
         Top             =   2040
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   582
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
         Format          =   97320961
         CurrentDate     =   41509
      End
      Begin MSComCtl2.DTPicker dtCheckDate 
         Height          =   330
         Left            =   1800
         TabIndex        =   6
         Top             =   3480
         Visible         =   0   'False
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   582
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
         Format          =   97320961
         CurrentDate     =   41509
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[For Check Issuance, deduct from]"
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
         TabIndex        =   25
         Top             =   4440
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Label Label5 
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
         Left            =   240
         TabIndex        =   24
         Top             =   3480
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label4 
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
         Left            =   240
         TabIndex        =   23
         Top             =   3120
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Check Amount"
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
         TabIndex        =   22
         Top             =   2760
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Label lblPreview 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   4440
         MouseIcon       =   "FIN_ExpensesFrm.frx":000C
         MousePointer    =   99  'Custom
         TabIndex        =   21
         Top             =   3840
         Width           =   315
      End
      Begin VB.Label Label1 
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
         Left            =   3360
         MouseIcon       =   "FIN_ExpensesFrm.frx":015E
         MousePointer    =   99  'Custom
         TabIndex        =   20
         Top             =   1320
         Width           =   1500
      End
      Begin VB.Label Label3 
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
         Left            =   240
         TabIndex        =   19
         Top             =   4200
         Width           =   795
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
         Left            =   240
         TabIndex        =   18
         Top             =   5040
         Visible         =   0   'False
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
         Left            =   240
         TabIndex        =   17
         Top             =   5400
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Expense"
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
         TabIndex        =   16
         Top             =   1680
         Width           =   780
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
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
         TabIndex        =   15
         Top             =   2415
         Width           =   750
      End
      Begin VB.Label Label10 
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
         Left            =   240
         TabIndex        =   14
         Top             =   2040
         Width           =   435
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Record your business expenses to better track cash flow."
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
         TabIndex        =   11
         Top             =   720
         Width           =   4455
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Expenses"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   345
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   1065
      End
   End
End
Attribute VB_Name = "FIN_ExpensesFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub Initialize()
    dtFund.value = Format(Now, "MM/DD/YY")
    txtAmount.Text = "0.00"
    txtCheckAmount.Text = "0.00"
    txtCheckNumber.Text = ""
    dtCheckDate.value = Format(Now, "MM/DD/YY")
    txtRemarks.Text = ""
End Sub
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
       
        Case "Expenses"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Expense")
            cmbExpense.Clear
            cmbExpense.AddItem ""
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        cmbExpense.AddItem rec!Expense
                        cmbExpense.ItemData(cmbExpense.NewIndex) = rec!ExpenseId
                    End If
                    rec.MoveNext
                Loop
            End If
            cmbExpense.ListIndex = 0
    End Select
End Sub
Private Sub Save()
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    con.BeginTrans
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "FIN_Expenses_Insert"
    cmd.Parameters.Append cmd.CreateParameter("@ExpenseId", adInteger, adParamInputOutput, , cmbExpense.ItemData(cmbExpense.ListIndex))
    cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtFund.value)
    cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Val(Replace(txtAmount.Text, ",", "")))
                          cmd.Parameters("@Amount").Precision = 18
                          cmd.Parameters("@Amount").NumericScale = 2
    cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.Text, ",", "")))
                          cmd.Parameters("@CheckAmount").Precision = 18
                          cmd.Parameters("@CheckAmount").NumericScale = 2
    cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 50, txtCheckNumber.Text)
    cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
    cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , 1)
    If Val(Replace(txtCheckAmount.Text, ",", "")) > 0 Or Trim(txtCheckNumber.Text) <> "" Then
        cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , cmbAccount.ItemData(cmbAccount.ListIndex))
    Else
        cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , Null)
    End If
    cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 255, txtRemarks.Text)
    cmd.Parameters.Append cmd.CreateParameter("@WorkStationId", adInteger, adParamInput, , WorkstationId)
    cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
    cmd.Execute
    
    Dim ExpenseId As Integer
    ExpenseId = cmd.Parameters("@ExpenseId")
    
    'BANK OUTFLOW
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "FIN_BankOutFlow_Insert"
    cmd.Parameters.Append cmd.CreateParameter("@Cash", adDecimal, adParamInput, , 0)
                          cmd.Parameters("@Cash").Precision = 18
                          cmd.Parameters("@Cash").NumericScale = 2
    cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.Text, ",", "")))
                          cmd.Parameters("@CheckAmount").Precision = 18
                          cmd.Parameters("@CheckAmount").NumericScale = 2
    cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.Text)
    cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
    cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtFund.value)
    cmd.Parameters.Append cmd.CreateParameter("@Particulars", adVarChar, adParamInput, 250, "Expenses: " & cmbExpense.Text & " " & txtRemarks.Text)
    cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , Null)
    cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , Null)
    cmd.Parameters.Append cmd.CreateParameter("@ExpenseId", adInteger, adParamInput, , ExpenseId)
    cmd.Execute
    
    'INFLOW AND OUTFLOW
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "FIN_OutFlow_Insert"
    cmd.Parameters.Append cmd.CreateParameter("@Cash", adDecimal, adParamInput, , Val(Replace(txtAmount.Text, ",", "")))
                          cmd.Parameters("@Cash").Precision = 18
                          cmd.Parameters("@Cash").NumericScale = 2
    cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.Text, ",", "")))
                          cmd.Parameters("@CheckAmount").Precision = 18
                          cmd.Parameters("@CheckAmount").NumericScale = 2
    cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.Text)
    cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
    cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtFund.value)
    cmd.Parameters.Append cmd.CreateParameter("@Particulars", adVarChar, adParamInput, 250, "Expenses: " & cmbExpense.Text & " " & txtRemarks.Text)
    cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , Null)
    cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , Null)
    cmd.Parameters.Append cmd.CreateParameter("@ExpenseId", adInteger, adParamInput, , ExpenseId)
    cmd.Execute
    
    'CHECK REGISTRY
    If Trim(txtCheckNumber.Text) <> "" Or Val(Replace(txtCheckAmount.Text, ",", "")) > 0 Then
        'Set con = New ADODB.Connection
        Set cmd = New ADODB.Command
        'con.ConnectionString = ConnString
        'con.Open
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "FIN_CheckRegistry_Insert"
        cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.Text)
        cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
        cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.Text, ",", "")))
                              cmd.Parameters("@Amount").NumericScale = 2
                              cmd.Parameters("@Amount").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, txtRemarks.Text)
        cmd.Parameters.Append cmd.CreateParameter("@isReceivable", adBoolean, adParamInput, , "False")
        cmd.Parameters.Append cmd.CreateParameter("@CheckStatusId", adInteger, adParamInput, , 1) '--open
        cmd.Parameters.Append cmd.CreateParameter("@CheckRegistryId", adInteger, adParamInputOutput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@ExpenseId", adInteger, adParamInput, , ExpenseId)
        cmd.Parameters.Append cmd.CreateParameter("@POS_SalesId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@SOPaymentId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@POPaymentId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , cmbAccount.ItemData(cmbAccount.ListIndex))
        
        cmd.Execute
        'con.Close
    End If
    con.CommitTrans
    con.Close
    
    Dim Amount As Double
    Amount = FormatNumber(NVAL(txtAmount.Text) + NVAL(txtCheckAmount.Text), 2, vbTrue, vbFalse)
    SavePOSAuditTrail UserId, WorkstationId, "", "Created new expense: " & cmbExpense.Text & " Amount:" & Amount, "FINANCE"
    
    MsgBox "Record saved.", vbInformation, "Success!"
    Initialize
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnSave_Click()
    If EditAccessRights(18) = False Then
        MsgBox ErrorCodes(74), vbCritical, "Limited Rights"
        Exit Sub
    End If
    If cmbExpense.Text = "" Then
        cmbExpense.SetFocus
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(26)
        GLOBAL_MessageFrm.Show (1)
    ElseIf Trim(txtCheckNumber.Text) <> "" Or Val(Replace(txtCheckAmount.Text, ",", "")) > 0 Then
        If cmbAccount.Text = "" Then
            cmbAccount.SetFocus
            GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(16)
            GLOBAL_MessageFrm.Show (1)
        Else
            Save
        End If
    Else
        Save
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

Private Sub Form_Load()
    Populate "Bank"
    'Populate "Fund"
    Populate "Expenses"
    'cmbCashDirection.ListIndex = 0
    Initialize
End Sub

Private Sub txtFund_Amount_Change()
    
End Sub

Private Sub Label1_Click()
    FIN_ExpensesTransactionsFrm.Show (1)
End Sub

Private Sub txtAmount_Change()
    If IsNumeric(txtAmount.Text) = False Then
        txtAmount.Text = "0.00"
    Else
        lblPreview.Caption = FormatNumber(txtAmount.Text, 2, vbTrue)
    End If
End Sub

Private Sub txtAmount_GotFocus()
    selectText txtAmount
End Sub

Private Sub txtAmount_LostFocus()
    lblPreview.Caption = ""
End Sub

Private Sub txtCheckAmount_Change()
    If IsNumeric(txtCheckAmount.Text) = False Then
        txtCheckAmount.Text = "0.00"
    Else
        lblPreview.Caption = FormatNumber(txtCheckAmount.Text, 2, vbTrue)
    End If
End Sub

Private Sub txtCheckAmount_LostFocus()
    lblPreview.Caption = ""
End Sub
