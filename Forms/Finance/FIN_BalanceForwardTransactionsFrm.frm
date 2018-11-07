VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FIN_BalanceForwardTransactionsFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   Icon            =   "FIN_BalanceForwardTransactionsFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9255
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
         TabIndex        =   7
         Top             =   1920
         Visible         =   0   'False
         Width           =   1695
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
         Left            =   1320
         TabIndex        =   1
         Top             =   720
         Value           =   -1  'True
         Width           =   975
      End
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
         Left            =   2280
         TabIndex        =   2
         Top             =   720
         Width           =   735
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
         TabIndex        =   5
         Top             =   1920
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   330
         Left            =   1320
         TabIndex        =   3
         Top             =   1080
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
         Format          =   71237633
         CurrentDate     =   41509
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   330
         Left            =   1320
         TabIndex        =   4
         Top             =   1440
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
         Format          =   71237633
         CurrentDate     =   41509
      End
      Begin MSComctlLib.ListView lvTransactions 
         Height          =   4455
         Left            =   120
         TabIndex        =   6
         Top             =   2400
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   7858
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
         NumItems        =   5
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
            Text            =   "Account/Fund Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Particulars"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Amount"
            Object.Width           =   2540
         EndProperty
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
         TabIndex        =   10
         Top             =   1440
         Width           =   705
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
         TabIndex        =   9
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Balance Forwarded Transactions"
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
         Width           =   3750
      End
   End
End
Attribute VB_Name = "FIN_BalanceForwardTransactionsFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
        cmd.Parameters.Append cmd.CreateParameter("@Id", adInteger, adParamInput, , lvTransactions.SelectedItem.text)
        
        If optBank.value = True Then
            cmd.CommandText = "FIN_AccountCashFlow_Delete"
            cmd.Execute
        Else
            cmd.CommandText = "FIN_FundCashFlow_Delete"
            cmd.Execute
        End If
        
        con.Close
        btnSearch_Click
    End If
End Sub

Private Sub btnSearch_Click()
    btnDelete.Visible = False
    
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    Dim Item As MSComctlLib.ListItem
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    
    cmd.Parameters.Append cmd.CreateParameter("@DateFrom", adDate, adParamInput, , dtFrom.value)
    cmd.Parameters.Append cmd.CreateParameter("@DateTo", adDate, adParamInput, , dtTo.value)
    
    lvTransactions.ListItems.Clear
    
    If optBank.value = True Then
        cmd.CommandText = "FIN_AccountCashFlow_Search"
        Set rec = cmd.Execute
        If Not rec.EOF Then
            Do Until rec.EOF
                Set Item = lvTransactions.ListItems.Add(, , rec!AccountCashFlowId)
                    Item.SubItems(1) = Format(rec!Date, "MM/DD/YY")
                    Item.SubItems(2) = rec!AccountName
                    Item.SubItems(3) = rec!Remarks
                    Item.SubItems(4) = FormatNumber(rec!debit, 2, vbTrue)
                rec.MoveNext
            Loop
        End If
    Else
        cmd.CommandText = "FIN_FundCashFlow_Search"
        Set rec = cmd.Execute
        If Not rec.EOF Then
            Do Until rec.EOF
                Set Item = lvTransactions.ListItems.Add(, , rec!FundCashFlowId)
                    Item.SubItems(1) = Format(rec!Date, "MM/DD/YY")
                    Item.SubItems(2) = rec!FundName
                    Item.SubItems(3) = rec!Remarks
                    Item.SubItems(4) = FormatNumber(rec!debit, 2, vbTrue)
                rec.MoveNext
            Loop
        End If
    End If
    
    con.Close
End Sub

Private Sub Form_Load()
    lvTransactions.ColumnHeaders(2).width = lvTransactions.width * 0.18
    lvTransactions.ColumnHeaders(3).width = lvTransactions.width * 0.24
    lvTransactions.ColumnHeaders(4).width = lvTransactions.width * 0.3
    lvTransactions.ColumnHeaders(5).width = lvTransactions.width * 0.25
    
    dtFrom.value = Format(Now, "MM/DD/YY")
    dtTo.value = Format(Now, "MM/DD/YY")
End Sub

Private Sub lvTransactions_ItemClick(ByVal Item As MSComctlLib.ListItem)
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
