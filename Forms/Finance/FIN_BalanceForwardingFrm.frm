VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form FIN_BalanceForwardingFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   Icon            =   "FIN_BalanceForwardingFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FRE_Main 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   6615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
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
         Left            =   4200
         TabIndex        =   1
         Top             =   6000
         Width           =   1335
      End
      Begin VB.Frame FRE_Bank 
         BackColor       =   &H00FFFFFF&
         Height          =   5895
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   5415
         Begin VB.OptionButton optCash 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Cash"
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
            Left            =   1440
            TabIndex        =   19
            Top             =   3480
            Width           =   2415
         End
         Begin VB.OptionButton optBank 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Banks"
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
            Left            =   1440
            TabIndex        =   18
            Top             =   1680
            Value           =   -1  'True
            Width           =   2415
         End
         Begin VB.ComboBox cmbFund 
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
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   3840
            Width           =   3015
         End
         Begin VB.ComboBox cmbAccountNumber 
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
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   2400
            Width           =   3015
         End
         Begin VB.ComboBox cmbBank 
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
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   2040
            Width           =   3015
         End
         Begin MSComCtl2.DTPicker dtForward 
            Height          =   330
            Left            =   1440
            TabIndex        =   6
            Top             =   4920
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
            Format          =   73596929
            CurrentDate     =   41509
         End
         Begin VB.Label Label2 
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
            Left            =   3720
            MouseIcon       =   "FIN_BalanceForwardingFrm.frx":000C
            MousePointer    =   99  'Custom
            TabIndex        =   20
            Top             =   1440
            Width           =   1500
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Funds"
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
            Top             =   3840
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   240
            TabIndex        =   16
            Top             =   4200
            Width           =   720
         End
         Begin VB.Label lblFund_Balance 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
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
            Left            =   1440
            TabIndex        =   15
            Top             =   4230
            Width           =   3015
         End
         Begin VB.Label Label7 
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
            TabIndex        =   13
            Top             =   2040
            Width           =   450
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account #"
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
            TabIndex        =   12
            Top             =   2400
            Width           =   900
         End
         Begin VB.Label Label4 
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
            TabIndex        =   11
            Top             =   4920
            Width           =   435
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Use this option to carry over funds left from previous transactions. This will help you better monitor your business cash flow."
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
            TabIndex        =   10
            Top             =   840
            Width           =   5175
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Balance Forward"
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
            TabIndex        =   9
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   240
            TabIndex        =   8
            Top             =   2760
            Width           =   720
         End
         Begin VB.Label lblBank_Balance 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
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
            Left            =   1440
            TabIndex        =   7
            Top             =   2790
            Width           =   3015
         End
      End
      Begin VB.CommandButton btnSave 
         Caption         =   "Forward && Close"
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
         TabIndex        =   2
         Top             =   6000
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FIN_BalanceForwardingFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
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
Private Sub ControlState()
    If optBank.value = True Then
        cmbBank.Enabled = True
        cmbAccountNumber.Enabled = True
        cmbFund.Enabled = False
    Else
        cmbBank.Enabled = False
        cmbAccountNumber.Enabled = False
        cmbFund.Enabled = True
    End If
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnSave_Click()
On Error GoTo ErrorHandler
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    
    If optBank.value = True Then
        If cmbAccountNumber.ListIndex <> -1 Then
            cmd.Parameters.Append cmd.CreateParameter("@Id", adInteger, adParamInput, , cmbAccountNumber.ItemData(cmbAccountNumber.ListIndex))
        End If
    Else
        cmd.Parameters.Append cmd.CreateParameter("@Id", adInteger, adParamInput, , cmbFund.ItemData(cmbFund.ListIndex))
    End If
    
    If optBank.value = True Then
        If cmbAccountNumber.text = "" Then
            GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(21)
            GLOBAL_MessageFrm.Show (1)
            cmbAccountNumber.SetFocus
        Else
            cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Val(Replace(lblBank_Balance.Caption, ",", "")))
                          cmd.Parameters("@Amount").NumericScale = 2
                          cmd.Parameters("@Amount").Precision = 18
            cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtForward.value)
            cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 255, "BALANCE FORWARDED")
            cmd.CommandText = "FIN_BankBalanceForward_Insert"
            cmd.Execute
            MsgBox "Current balance now forwarded.", vbInformation, "BALANCE FORWARDED"
        End If
    Else
        cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Val(Replace(lblFund_Balance.Caption, ",", "")))
                          cmd.Parameters("@Amount").NumericScale = 2
                          cmd.Parameters("@Amount").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtForward.value)
        cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 255, "BALANCE FORWARDED")
        cmd.CommandText = "FIN_FundBalanceForward_Insert"
        cmd.Execute
        MsgBox "Current balance now forwarded.", vbInformation, "BALANCE FORWARDED"
    End If
    con.Close
    Unload Me
    Exit Sub
    
ErrorHandler:
    If IsNumeric(Err.Description) = True Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(Val(Err.Description))
    Else
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & Err.Description
    End If
    GLOBAL_MessageFrm.Show (1)
End Sub

Private Sub cmbAccountNumber_Click()
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_BankAccount_Get"
    
    cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , cmbAccountNumber.ItemData(cmbAccountNumber.ListIndex))
    Set rec = cmd.Execute
    lblBank_Balance.Caption = "0.00"
    If Not rec.EOF Then
        lblBank_Balance.Caption = FormatNumber(rec!balance, 2, vbTrue, vbFalse)
    End If
    con.Close
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
    cmbAccountNumber.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            If rec!isActive = "True" Then
                cmbAccountNumber.AddItem rec!accountnumber & " - " & rec!AccountName
                cmbAccountNumber.ItemData(cmbAccountNumber.NewIndex) = rec!AccountId
            End If
            rec.MoveNext
        Loop
    End If
    con.Close
End Sub

Private Sub cmbFund_Click()
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_Fund_Get"
    
    cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , cmbFund.ItemData(cmbFund.ListIndex))
    Set rec = cmd.Execute
    lblFund_Balance.Caption = "0.00"
    If Not rec.EOF Then
        lblFund_Balance.Caption = FormatNumber(rec!balance, 2, vbTrue, vbFalse)
    End If
    con.Close
End Sub

Private Sub Form_Load()
    ControlState
    Populate "Bank"
    Populate "Fund"
    dtForward.value = Format(Now, "MM/DD/YY 00:00:00")
End Sub

Private Sub Label2_Click()
    FIN_BalanceForwardTransactionsFrm.Show (1)
End Sub

Private Sub optBank_Click()
    ControlState
End Sub

Private Sub optCash_Click()
    ControlState
End Sub
