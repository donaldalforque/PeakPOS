VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FIN_ExpensesTransactionsFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9495
   Icon            =   "FIN_ExpensesTransactionsFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
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
         Left            =   1320
         TabIndex        =   9
         Top             =   840
         Width           =   7815
      End
      Begin VB.CommandButton btnSearch 
         Caption         =   "Refresh"
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
         Left            =   7440
         TabIndex        =   2
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CommandButton btnDelete 
         BackColor       =   &H008080FF&
         Caption         =   "Delete"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2040
         Visible         =   0   'False
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   330
         Left            =   1320
         TabIndex        =   3
         Top             =   1200
         Width           =   7815
         _ExtentX        =   13785
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
         Format          =   96337921
         CurrentDate     =   41509
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   330
         Left            =   1320
         TabIndex        =   4
         Top             =   1560
         Width           =   7815
         _ExtentX        =   13785
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
         Format          =   96337921
         CurrentDate     =   41509
      End
      Begin MSComctlLib.ListView lvTransactions 
         Height          =   4335
         Left            =   120
         TabIndex        =   5
         Top             =   2520
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7646
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Expenses"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Check Number"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Check Amount"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Check Date"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label1 
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
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Expenses Transactions"
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
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2595
      End
      Begin VB.Label Label3 
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
         TabIndex        =   7
         Top             =   1200
         Width           =   960
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
         TabIndex        =   6
         Top             =   1560
         Width           =   705
      End
   End
End
Attribute VB_Name = "FIN_ExpensesTransactionsFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub Populate(ByVal data As String)
    Select Case data
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
Private Sub btnDelete_Click()
    Dim x As Variant
    x = MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo, "")
    If x = vbYes Then
        Set con = New ADODB.Connection
        Set cmd = New ADODB.Command
        
        con.ConnectionString = ConnString
        con.Open
        cmd.ActiveConnection = con
        
        cmd.CommandType = adCmdStoredProc
        cmd.Parameters.Append cmd.CreateParameter("@Id", adInteger, adParamInput, , lvTransactions.SelectedItem.Text)
        
        cmd.CommandText = "FIN_Expenses_Delete"
        cmd.Execute
        
        'Save Audit
        SavePOSAuditTrail UserId, WorkstationId, "", "DELETE EXPENSES: " & lvTransactions.SelectedItem.SubItems(2) & ", AMOUNT:" & lvTransactions.SelectedItem.SubItems(3), "FINANCE"
        
        con.Close
        btnSearch_Click
    End If
End Sub

Private Sub btnSearch_Click()
    btnDelete.Visible = False
    
    Dim id As Variant
    
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    Dim item As MSComctlLib.ListItem
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    
    Select Case cmbExpense.ListIndex
        Case -1
            id = Null
        Case 0
            id = Null
        Case Else
            id = cmbExpense.ItemData(cmbExpense.ListIndex)
    End Select
        
    cmd.Parameters.Append cmd.CreateParameter("@ExpenseId", adInteger, adParamInput, , id)
    cmd.Parameters.Append cmd.CreateParameter("@DateFrom", adDate, adParamInput, , dtFrom.value)
    cmd.Parameters.Append cmd.CreateParameter("@DateTo", adDate, adParamInput, , dtTo.value)
    
    lvTransactions.ListItems.Clear
    
    cmd.CommandText = "FIN_Expenses_Search"
    Set rec = cmd.Execute
    If Not rec.EOF Then
        Do Until rec.EOF
            Set item = lvTransactions.ListItems.add(, , rec!ExpensesId)
                item.SubItems(1) = Format(rec!Date, "MM/DD/YY")
                item.SubItems(2) = rec!Expense
                item.SubItems(3) = FormatNumber(rec!amount, 2, vbTrue)
                item.SubItems(4) = rec!CheckNumber
                item.SubItems(5) = FormatNumber(rec!checkamount, 2, vbTrue)
                item.SubItems(6) = Format(rec!Date, "MM/DD/YY")
            rec.MoveNext
        Loop
    End If
    con.Close
End Sub

Private Sub Form_Load()
    lvTransactions.ColumnHeaders(2).width = lvTransactions.width * 0.1633
    lvTransactions.ColumnHeaders(3).width = lvTransactions.width * 0.1633
    lvTransactions.ColumnHeaders(4).width = lvTransactions.width * 0.1633
    lvTransactions.ColumnHeaders(5).width = lvTransactions.width * 0.1633
    lvTransactions.ColumnHeaders(6).width = lvTransactions.width * 0.1633
    lvTransactions.ColumnHeaders(7).width = lvTransactions.width * 0.1633
    
    dtFrom.value = Format(Now, "MM/DD/YY")
    dtTo.value = Format(Now, "MM/DD/YY")
    
    Populate "Expenses"
End Sub

Private Sub lvTransactions_ItemClick(ByVal item As MSComctlLib.ListItem)
    btnDelete.Visible = True
End Sub

Private Sub optBank_Click()
    btnDelete.Visible = False
    btnSearch_Click
End Sub

Private Sub optCash_Click()
    btnDelete.Visible = False
    btnSearch_Click
End Sub

