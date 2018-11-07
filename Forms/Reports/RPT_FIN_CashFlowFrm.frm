VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form RPT_FIN_CashFlowFrm 
   Caption         =   "Cash Flow Report"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15090
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9015
   ScaleWidth      =   15090
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.CheckBox chkCheck 
         BackColor       =   &H00E0E0E0&
         Caption         =   "CHECK"
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
         Left            =   1320
         TabIndex        =   20
         Top             =   3960
         Width           =   2415
      End
      Begin VB.CheckBox chkCash 
         BackColor       =   &H00E0E0E0&
         Caption         =   "CASH"
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
         Left            =   1320
         TabIndex        =   19
         Top             =   3720
         Width           =   2415
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
         ItemData        =   "RPT_FIN_CashFlowFrm.frx":0000
         Left            =   1320
         List            =   "RPT_FIN_CashFlowFrm.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   960
         Width           =   2415
      End
      Begin VB.OptionButton optCash 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cash on Hand"
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
         Left            =   1320
         TabIndex        =   16
         Top             =   600
         Value           =   -1  'True
         Width           =   2415
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
         ItemData        =   "RPT_FIN_CashFlowFrm.frx":002F
         Left            =   1320
         List            =   "RPT_FIN_CashFlowFrm.frx":0039
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2160
         Width           =   2415
      End
      Begin VB.OptionButton optBank 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cash in Bank"
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
         Left            =   1320
         TabIndex        =   13
         Top             =   1440
         Width           =   2415
      End
      Begin VB.CommandButton btnGenerate 
         Caption         =   "Generate Report"
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
         Left            =   1920
         TabIndex        =   3
         Top             =   5640
         Width           =   1815
      End
      Begin VB.TextBox txtTitle 
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
         Left            =   1320
         TabIndex        =   2
         Top             =   4920
         Width           =   2415
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
         ItemData        =   "RPT_FIN_CashFlowFrm.frx":005E
         Left            =   1320
         List            =   "RPT_FIN_CashFlowFrm.frx":0068
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1800
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker DateTo 
         Height          =   345
         Left            =   1320
         TabIndex        =   4
         Top             =   3240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   73990145
         CurrentDate     =   41686
      End
      Begin MSComCtl2.DTPicker DateFrom 
         Height          =   345
         Left            =   1320
         TabIndex        =   5
         Top             =   2880
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   73990145
         CurrentDate     =   41686
      End
      Begin VB.Label Label8 
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
         TabIndex        =   18
         Top             =   960
         Width           =   465
      End
      Begin VB.Label Label6 
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
         Top             =   2160
         Width           =   750
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report Title"
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
         TabIndex        =   11
         Top             =   4920
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Display"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   10
         Top             =   4440
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date From"
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
         TabIndex        =   9
         Top             =   2880
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filter By"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   1005
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date To"
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
         TabIndex        =   7
         Top             =   3240
         Width           =   705
      End
      Begin VB.Label Label3 
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
         TabIndex        =   6
         Top             =   1800
         Width           =   450
      End
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer 
      Height          =   9015
      Left            =   3840
      TabIndex        =   12
      Top             =   0
      Width           =   11295
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "RPT_FIN_CashFlowFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim crxApp As New CRAXDRT.Application
Dim crxRpt As New CRAXDRT.Report
Public Sub Populate(ByVal data As String)
    Dim item As MSComctlLib.ListItem
    Select Case data
        Case "Bank"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Bank")
            cmbBank.Clear
            'cmbBank.AddItem ""
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
            'cmbFund.AddItem ""
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        cmbFund.AddItem rec!FundName
                        cmbFund.ItemData(cmbFund.NewIndex) = rec!FundId
                    End If
                    rec.MoveNext
                Loop
            End If
            On Error Resume Next
            cmbFund.ListIndex = 0
    End Select
End Sub

Private Sub btnGenerate_Click()
    Dim sql, OrderBy As String
    Dim Status, Customer, Account, Fund As Variant
    
    Screen.MousePointer = vbHourglass
    If optCash.value = True Then
        Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\FundCashFlow.rpt")
    Else
        Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\AccountCashFlow.rpt")
    End If
    crxRpt.EnableParameterPrompting = False
    crxRpt.DiscardSavedData
    Call ResetRptDB(crxRpt)
         
    Select Case cmbFund.ListIndex
        Case -1
            Fund = ""
        Case Else
            Fund = "AND BASE_Fund.FundId = '" & cmbFund.ItemData(cmbFund.ListIndex) & "' "
    End Select
    
    Select Case cmbAccount.ListIndex
        Case -1
            Account = ""
        Case Else
            Account = "AND FIN_AccountCashFlow.AccountId = '" & cmbAccount.ItemData(cmbAccount.ListIndex) & "' "
    End Select
    
    Dim method As String
    method = ""
    
    If chkCash.value = Checked Then
        method = "AND PaymentType = 'CASH' "
    End If
    
    If chkCheck.value = Checked Then
        method = "AND PaymentType = 'CHECK' "
    End If
    
'    Select Case cmbSort.ListIndex
'        Case 0
'            OrderBy = "ORDER BY PO_Payment.AccountId ASC, BASE_Bank.BankName ASC"
'        Case 1
'            OrderBy = "ORDER BY PO_Payment.AccountId ASC, PO_Payment.ChequeNumber ASC"
'        Case 2
'            OrderBy = "ORDER BY PO_Payment.AccountId ASC, PO_Payment.ChequeDate ASC"
'        Case 3
'            OrderBy = "ORDER BY PO_Payment.AccountId ASC, PO_Payment.Date ASC"
'        Case 4
'            OrderBy = "ORDER BY PO_Payment.AccountId ASC, PO_Payment.Amount ASC"
'    End Select
    
    If optCash.value = True Then
        OrderBy = "ORDER BY FIN_FundCashFlow.FundId ASC, FIN_FundCashFlow.FundCashFlowId ASC"
        
        sql = "SELECT FIN_FundCashFlow.FundId,FIN_FundCashFlow.Debit,FIN_FundCashFlow.Credit,FIN_FundCashFlow.Date," & _
              "FIN_FundCashFlow.CheckNumber,FIN_FundCashFlow.CheckDate, FIN_FundCashFlow.PaymentType " & _
              "FIN_FundCashFlow.Remarks,BASE_Fund.FundName " & _
              "FROM Peak.dbo.FIN_FundCashFlow FIN_FundCashFlow " & _
              "INNER JOIN Peak.dbo.BASE_Fund BASE_Fund ON FIN_FundCashFlow.FundId = BASE_Fund.FundId " & _
              "WHERE FIN_FundCashFlow.Date >= '" & DateFrom.value & " 00:00:00' " & _
              "AND FIN_FundCashFlow.Date <= '" & dateTo.value & " 23:23:59' " & Fund & method & OrderBy
    Else
        OrderBy = "ORDER BY FIN_AccountCashFlow.AccountId ASC, FIN_AccountCashFlow.AccountCashFlowId ASC"
        
        sql = "SELECT BASE_BankAccount.AccountId,BASE_BankAccount.AccountNumber,BASE_BankAccount.AccountName," & _
              "FIN_AccountCashFlow.Debit,FIN_AccountCashFlow.Credit,FIN_AccountCashFlow.Date," & _
              "FIN_AccountCashFlow.Remarks FROM Peak.dbo.BASE_BankAccount BASE_BankAccount INNER JOIN " & _
              "Peak.dbo.FIN_AccountCashFlow FIN_AccountCashFlow ON BASE_BankAccount.AccountId = " & _
              "FIN_AccountCashFlow.AccountId " & _
              "WHERE FIN_AccountCashFlow.Date >= '" & DateFrom.value & " 00:00:00' " & _
              "AND FIN_AccountCashFlow.Date <= '" & dateTo.value & " 23:23:59' " & Account & OrderBy
    End If
    
    crxRpt.ParameterFields(1).AddCurrentValue txtTitle.text
    crxRpt.ParameterFields(2).AddCurrentValue Str(DateFrom.value)
    crxRpt.ParameterFields(3).AddCurrentValue Str(dateTo.value)
    crxRpt.SQLQueryString = sql
    CRViewer.ReportSource = crxRpt
    CRViewer.ViewReport
    CRViewer.Zoom 1
    Screen.MousePointer = vbDefault
End Sub

Private Sub chkCash_Click()
    If chkCash.value = Checked Then
        chkCheck.value = Unchecked
    End If
End Sub

Private Sub chkCheck_Click()
    If chkCheck.value = Checked Then
        chkCash.value = Unchecked
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

Private Sub CRViewer_PrintButtonClicked(UseDefault As Boolean)
    UseDefault = False
    crxRpt.PrinterSetup Me.hWnd
    crxRpt.PrintOut True
End Sub
Private Sub Form_Load()
    'cmbSort.ListIndex = 0
    Populate "Fund"
    Populate "Bank"
    
    Me.Height = 9390
    Me.width = 15180
    DateFrom.value = Format(Now, "MM/DD/YY")
    dateTo.value = Format(Now, "MM/DD/YY")
    
    txtTitle.text = Me.Caption
    
    If optCash.value = True Then
        cmbFund.Enabled = True
        cmbBank.Enabled = False
        cmbAccount.Enabled = False
    End If
End Sub


Private Sub Form_Resize()
   On Error Resume Next
    CRViewer.width = Me.width - Frame1.width
    CRViewer.Height = Me.Height
    Frame1.Height = Me.Height
End Sub

Private Sub optBank_Click()
    If optBank.value = True Then
        cmbFund.Enabled = False
        cmbBank.Enabled = True
        cmbAccount.Enabled = True
    End If
End Sub

Private Sub optCash_Click()
    If optCash.value = True Then
        cmbFund.Enabled = True
        cmbBank.Enabled = False
        cmbAccount.Enabled = False
    End If
End Sub
