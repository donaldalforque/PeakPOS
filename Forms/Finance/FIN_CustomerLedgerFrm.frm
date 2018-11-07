VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FIN_CustomerLedgerFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Ledger"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   15090
   Begin MSComCtl2.DTPicker DateFrom 
      Height          =   330
      Left            =   1320
      TabIndex        =   3
      Top             =   1200
      Width           =   13575
      _ExtentX        =   23945
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
      Format          =   69402625
      CurrentDate     =   41843
   End
   Begin VB.Frame LayoutFrame_Search 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   9390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15135
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
         Left            =   13680
         TabIndex        =   5
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtSearch_Name 
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
         Top             =   840
         Width           =   13575
      End
      Begin VB.TextBox txtSearch_Code 
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
         TabIndex        =   1
         Top             =   480
         Width           =   13575
      End
      Begin MSComctlLib.ListView lvSearch 
         Height          =   6375
         Left            =   120
         TabIndex        =   6
         Top             =   2520
         Width           =   14805
         _ExtentX        =   26114
         _ExtentY        =   11245
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CustomerId"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Particulars"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Debit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Credit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Balance"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComCtl2.DTPicker DateTo 
         Height          =   330
         Left            =   1320
         TabIndex        =   4
         Top             =   1560
         Width           =   13575
         _ExtentX        =   23945
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
         Format          =   69402625
         CurrentDate     =   41843
      End
      Begin VB.Label Label3 
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
         Left            =   240
         TabIndex        =   11
         Top             =   1560
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         TabIndex        =   10
         Top             =   840
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search"
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
         TabIndex        =   9
         Top             =   80
         Width           =   795
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label16 
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
         Left            =   240
         TabIndex        =   7
         Top             =   1200
         Width           =   960
      End
   End
End
Attribute VB_Name = "FIN_CustomerLedgerFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnSearch_Click()
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "FIN_CustomerLedger_Search"
    If Trim(txtSearch_Code.text) = "" Then
        cmd.Parameters.Append cmd.CreateParameter("@Code", adVarChar, adParamInput, 50, Null)
    Else
        cmd.Parameters.Append cmd.CreateParameter("@Code", adVarChar, adParamInput, 50, txtSearch_Code.text)
    End If
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, txtSearch_Name.text)
    cmd.Parameters.Append cmd.CreateParameter("@DateFrom", adDate, adParamInput, , DateFrom.value)
    cmd.Parameters.Append cmd.CreateParameter("@DateTo", adDate, adParamInput, , DateTo.value)
    Set rec = cmd.Execute
    Dim item As MSComctlLib.ListItem
    lvSearch.ListItems.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            Set item = lvSearch.ListItems.add(, , rec!CustomerId)
                item.SubItems(1) = Format(rec!Date, "MM/DD/YY")
                item.SubItems(2) = rec!customerCode
                item.SubItems(3) = rec!Name
                If IsNull(rec!OrderNumber) = True Then
                    item.SubItems(4) = rec!Particulars
                Else
                    item.SubItems(4) = rec!Particulars & " :" & rec!OrderNumber
                End If
                If IsNull(rec!debit) = True Then
                    item.SubItems(6) = FormatNumber(rec!credit, 2, vbTrue, vbFalse)
                Else
                    item.SubItems(5) = FormatNumber(rec!debit, 2, vbTrue, vbFalse)
                End If
            rec.MoveNext
        Loop
    End If
    con.Close
    
    'CountRunning Total
    Dim runningtotal As Double
    For Each item In lvSearch.ListItems
        runningtotal = runningtotal + Val(Replace(item.SubItems(5), ",", "")) - Val(Replace(item.SubItems(6), ",", ""))
        item.SubItems(7) = FormatNumber(runningtotal, 2, vbTrue, vbFalse)
    Next
End Sub

Private Sub Form_Load()
    lvSearch.ColumnHeaders(2).width = lvSearch.width * 0.1233
    lvSearch.ColumnHeaders(3).width = lvSearch.width * 0.1233
    lvSearch.ColumnHeaders(4).width = lvSearch.width * 0.2433
    lvSearch.ColumnHeaders(5).width = lvSearch.width * 0.2433
    lvSearch.ColumnHeaders(6).width = lvSearch.width * 0.1233
    lvSearch.ColumnHeaders(7).width = lvSearch.width * 0.1233
    lvSearch.ColumnHeaders(8).width = lvSearch.width * 0.1233
    
    DateFrom.value = Format(Now, "MM/DD/YY")
    DateTo.value = Format(Now, "MM/DD/YY")
End Sub

Private Sub txtSearch_Code_Change()
    btnSearch_Click
End Sub

Private Sub txtSearch_Name_Change()
    btnSearch_Click
End Sub
