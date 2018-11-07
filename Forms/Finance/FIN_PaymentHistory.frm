VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FIN_PaymentHistoryFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Payment History"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15090
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   15090
   Begin VB.CommandButton btnDelete 
      BackColor       =   &H008080FF&
      Caption         =   "Delete Payment"
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
      TabIndex        =   20
      Top             =   8520
      Visible         =   0   'False
      Width           =   1815
   End
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
      Left            =   11760
      TabIndex        =   5
      Top             =   8520
      Visible         =   0   'False
      Width           =   1575
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
      Left            =   13380
      TabIndex        =   6
      Top             =   8520
      Width           =   1575
   End
   Begin VB.Frame LayoutFrame_Search 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   15135
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
         Left            =   1440
         TabIndex        =   15
         Top             =   840
         Width           =   13455
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
         Left            =   1440
         TabIndex        =   14
         Top             =   480
         Width           =   13455
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
         Left            =   13320
         TabIndex        =   3
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1200
         Left            =   1440
         TabIndex        =   7
         Top             =   1960
         Width           =   13455
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Account Type"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   120
            TabIndex        =   8
            Top             =   0
            Width           =   4215
            Begin VB.OptionButton optAccountsReceivable 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Accounts Receivable"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   1
               Top             =   240
               Value           =   -1  'True
               Width           =   3735
            End
            Begin VB.OptionButton optAccountsPayable 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Accounts Payable"
               BeginProperty Font 
                  Name            =   "Calibri"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   2
               Top             =   600
               Width           =   3975
            End
         End
      End
      Begin MSComCtl2.DTPicker dtFrom 
         Height          =   345
         Left            =   1440
         TabIndex        =   16
         Top             =   1200
         Width           =   13455
         _ExtentX        =   23733
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
         Format          =   95944705
         CurrentDate     =   41818
      End
      Begin MSComCtl2.DTPicker dtTo 
         Height          =   345
         Left            =   1440
         TabIndex        =   17
         Top             =   1560
         Width           =   13455
         _ExtentX        =   23733
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
         Format          =   95944705
         CurrentDate     =   41818
      End
      Begin VB.Label Label4 
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
         TabIndex        =   22
         Top             =   1200
         Width           =   960
      End
      Begin VB.Label Label5 
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
         TabIndex        =   21
         Top             =   1560
         Width           =   705
      End
      Begin VB.Label Label13 
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
         TabIndex        =   19
         Top             =   840
         Width           =   555
      End
      Begin VB.Label Label2 
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
         TabIndex        =   18
         Top             =   480
         Width           =   480
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
   End
   Begin MSComctlLib.ListView lvSearch 
      Height          =   4095
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   14820
      _ExtentX        =   26141
      _ExtentY        =   7223
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
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "PaymentHistoryId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "CustomerId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Cash Payment"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Check Payment"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Check Number"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Check Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Issuing Bank"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Adjustment"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Remarks"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "TransactionId"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label lblPreviousPayment 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   9600
      TabIndex        =   13
      Top             =   3600
      Width           =   45
   End
   Begin VB.Label lblPaymentDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   9600
      TabIndex        =   12
      Top             =   3960
      Width           =   45
   End
   Begin VB.Label lblOrderNumber 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   9600
      TabIndex        =   11
      Top             =   3240
      Width           =   45
   End
   Begin VB.Label lblFooterLine 
      Caption         =   "Label5"
      Height          =   15
      Left            =   120
      TabIndex        =   10
      Top             =   8385
      Width           =   14820
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   375
      Left            =   0
      Top             =   -120
      Width           =   15135
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   4995
      Left            =   0
      Top             =   4050
      Width           =   15135
   End
End
Attribute VB_Name = "FIN_PaymentHistoryFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private isAR As Boolean
Private OrderId As Integer

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    If EditAccessRights(22) = False Then
        MsgBox ErrorCodes(74), vbCritical, "Limited Rights"
        Exit Sub
    End If
    Dim x As Variant
    x = MsgBox("Are you sure you want to delete this record?", vbCritical + vbYesNo)
    If x = vbYes Then
        Set con = New ADODB.Connection
        Set cmd = New ADODB.Command
        con.ConnectionString = ConnString
        con.Open
        cmd.ActiveConnection = con
        
        cmd.CommandType = adCmdStoredProc
        cmd.Parameters.Append cmd.CreateParameter("@PaymentId", adInteger, adParamInput, , lvSearch.SelectedItem.Text)
        cmd.Parameters.Append cmd.CreateParameter("@TransactionId", adInteger, adParamInput, , lvSearch.SelectedItem.SubItems(11))
        cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
        cmd.Parameters.Append cmd.CreateParameter("@WorkStationId", adInteger, adParamInput, , WorkstationId)
        
        If isAR = True Then
            cmd.CommandText = "SO_PaymentHistory_delete"
            cmd.Execute
        Else
            cmd.CommandText = "PO_PaymentHistory_delete"
            cmd.Execute
        End If
        
        con.Close
        MsgBox "Record deleted.", vbInformation, "Success"
        
        UpdateCustomerOrderDues
        UpdateVendorOrderDues
        
        btnSearch_Click
    End If
End Sub



Private Sub btnSearch_Click()
    lvSearch.ListItems.Clear
    
    OrderId = 0
    
    'If Trim(txtSearch_Order.text) = "" Then Exit Sub
    
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    Dim item As MSComctlLib.ListItem
    
    con.ConnectionString = ConnString
    con.Open
    
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("@Code", adVarChar, adParamInput, 50, txtSearch_Code.Text)
    cmd.Parameters.Append cmd.CreateParameter("@CustomerName", adVarChar, adParamInput, 250, txtSearch_name.Text)
    cmd.Parameters.Append cmd.CreateParameter("@DateFrom", adDate, adParamInput, , dtFrom.value)
    cmd.Parameters.Append cmd.CreateParameter("@DateTo", adDate, adParamInput, , dtTo.value)
    
    If optAccountsReceivable.value = True Then
        cmd.CommandText = "FIN_SOPaymentHistory_Search"
    Else
        cmd.CommandText = "FIN_POPaymentHistory_Search"
    End If
    
    Set rec = cmd.Execute
    
    If Not rec.EOF Then
        Do Until rec.EOF
            If optAccountsReceivable.value = True Then
                On Error Resume Next
                Set item = lvSearch.ListItems.add(, , rec!SO_PaymentHistoryId)
                    item.SubItems(1) = rec!CustomerId
                    item.SubItems(2) = rec!Name
                    item.SubItems(3) = Format(rec!Date, "MM/DD/YY")
                    item.SubItems(4) = FormatNumber(rec!amount, 2, vbTrue, vbFalse)
                    item.SubItems(5) = FormatNumber(rec!checkamount, 2, vbTrue, vbFalse)
                    item.SubItems(6) = rec!CheckNumber
                    item.SubItems(7) = Format(rec!CheckDate, "MM/DD/YY")
                    item.SubItems(8) = rec!issuingbank
                    item.SubItems(9) = FormatNumber(rec!salesdiscount, 2, vbTrue, vbFalse)
                    item.SubItems(10) = rec!Remarks
                    item.SubItems(11) = rec!TransactionId
                isAR = True
            Else
                Set item = lvSearch.ListItems.add(, , rec!PO_PaymentHistoryId)
                    item.SubItems(1) = rec!VendorId
                    item.SubItems(2) = rec!Name
                    item.SubItems(3) = Format(rec!Date, "MM/DD/YY")
                    item.SubItems(4) = FormatNumber(rec!amount, 2, vbTrue, vbFalse)
                    item.SubItems(5) = FormatNumber(rec!checkamount, 2, vbTrue, vbFalse)
                    item.SubItems(6) = rec!CheckNumber
                    item.SubItems(7) = Format(rec!CheckDate, "MM/DD/YY")
                    item.SubItems(8) = rec!issuingbank
                    item.SubItems(9) = FormatNumber(rec!tax, 2, vbTrue, vbFalse)
                    item.SubItems(10) = rec!Remarks
                    item.SubItems(11) = rec!TransactionId
                isAR = False
            End If
            
            rec.MoveNext
        Loop
    End If
    con.Close
    If lvSearch.ListItems.Count <= 0 Then btnDelete.Visible = False
End Sub

Private Sub Form_Load()
    'StatusBarWidth Me, statusBar_Main
    
    lvSearch.ColumnHeaders(3).width = lvSearch.width * 0.145  'Name
    lvSearch.ColumnHeaders(4).width = lvSearch.width * 0.1  'Date
    lvSearch.ColumnHeaders(5).width = lvSearch.width * 0.1225 'Cash
    lvSearch.ColumnHeaders(6).width = lvSearch.width * 0.1225 'Check
    lvSearch.ColumnHeaders(7).width = lvSearch.width * 0.1225 'Check #
    lvSearch.ColumnHeaders(8).width = lvSearch.width * 0.1 'Check Date
    lvSearch.ColumnHeaders(9).width = lvSearch.width * 0.12 'Check Date
    lvSearch.ColumnHeaders(10).width = lvSearch.width * 0.1225 'Tax
    lvSearch.ColumnHeaders(11).width = lvSearch.width * 0.145 'Remarks
    
    'lvOrderSearch.ColumnHeaders(2).width = lvOrderSearch.width * 0.98
    Dim currdate As String
    currdate = "January 1," & Year(Now)
    
    dtFrom.value = Format(currdate, "MM/DD/YY")
    dtTo.value = Format(Now, "MM/DD/YY")
End Sub

Private Sub Label6_Click()
End Sub

Private Sub lvSearch_ItemClick(ByVal item As MSComctlLib.ListItem)
    btnDelete.Visible = True
End Sub
'
'Private Sub txtSearch_Order_GotFocus()
'    selectText txtSearch_Order
'End Sub

Private Sub txtSearch_Code_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            btnSearch_Click
    End Select
End Sub

Private Sub txtSearch_Name_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            btnSearch_Click
    End Select
End Sub
