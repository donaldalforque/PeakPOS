VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form POS_RecentReceiptsFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Recent Receipts"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7335
   Icon            =   "POS_RecentReceiptsFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCredit 
      Caption         =   "F3:CREDIT"
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
      Left            =   3720
      Picture         =   "POS_RecentReceiptsFrm.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton btnCash 
      Caption         =   "F2:CASH"
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
      Left            =   2160
      Picture         =   "POS_RecentReceiptsFrm.frx":065A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "ESC:Cancel"
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
      Left            =   5280
      Picture         =   "POS_RecentReceiptsFrm.frx":0C7C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4920
      Width           =   1935
   End
   Begin VB.CommandButton btnPrint 
      Caption         =   "F1:Print"
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
      Picture         =   "POS_RecentReceiptsFrm.frx":300B
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4920
      Width           =   1935
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   8281
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "POSSaleId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "OR #"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "POS_RecentReceiptsFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim isCredit As Boolean

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnCash_Click()
    lvList.ListItems.Clear
    Set rec = New ADODB.Recordset
    Set rec = Global_Data("RecentReceipts")
    Dim item As MSComctlLib.ListItem
    If Not rec.EOF Then
        Do Until rec.EOF
            Set item = lvList.ListItems.add(, , rec!POS_SalesId)
                item.SubItems(1) = rec!pos_ordernumber
            rec.MoveNext
        Loop
    End If
    isCredit = False
End Sub

Private Sub btnCredit_Click()
    lvList.ListItems.Clear
    Set rec = New ADODB.Recordset
    Set rec = Global_Data("RecentCreditReceipts")
    Dim item As MSComctlLib.ListItem
    If Not rec.EOF Then
        Do Until rec.EOF
            Set item = lvList.ListItems.add(, , rec!SalesOrderId)
                item.SubItems(1) = rec!OrderNumber
            rec.MoveNext
        Loop
    End If
    isCredit = True
End Sub

Private Sub btnPrint_Click()
    Set con = New ADODB.Connection
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.CommandType = adCmdStoredProc
    cmd.ActiveConnection = con
    cmd.CommandText = "SYSAuditTrail_Insert"
    cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId) '1 DEFAULT
    cmd.Parameters.Append cmd.CreateParameter("@Module", adVarChar, adParamInput, 250, "POS")
    cmd.Parameters.Append cmd.CreateParameter("@Action", adVarChar, adParamInput, 250, "REPRINT")
    cmd.Execute
    con.Close
    
    'Save Audit Trail
    SavePOSAuditTrail VoidUserId, WorkstationId, lvList.SelectedItem.Text, "REPRINT OR#: " & lvList.SelectedItem.SubItems(1)
    
    Dim crxApp As New CRAXDRT.Application
    Dim crxRpt As New CRAXDRT.Report
    
    If isCredit = False Then
        '**PRINT RECEIPT******
        Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\POS_Receipt.rpt")
        crxRpt.DiscardSavedData
        crxRpt.EnableParameterPrompting = False
        crxRpt.ParameterFields.GetItemByName("Notice").AddCurrentValue "***THIS IS A REPRINT***"
        crxRpt.ParameterFields.GetItemByName("@POS_SalesId").AddCurrentValue Val(lvList.SelectedItem.Text)
    
        Call ResetRptDB(crxRpt)
        crxRpt.PrintOut False
    Else
        '**PRINT RECEIPT******
        Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\POS_Receipt_Account.rpt")
        crxRpt.DiscardSavedData
        crxRpt.EnableParameterPrompting = False
        crxRpt.ParameterFields.GetItemByName("Notice").AddCurrentValue ""
        crxRpt.RecordSelectionFormula = "{SO_SalesOrder.SalesOrderId}= " & Val(lvList.SelectedItem.Text) & ""

        Call ResetRptDB(crxRpt)
        crxRpt.PrintOut False
    End If
    
End Sub

Private Sub Command1_Click()
    
End Sub

Private Sub Command2_Click()
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            btnPrint_Click
        Case vbKeyEscape
            Unload Me
        Case vbKeyF2
            btnCash_Click
        Case vbKeyF3
            btnCredit_Click
    End Select
End Sub

Private Sub Form_Load()
    lvList.ColumnHeaders(2).width = lvList.width * 0.95
    btnCash_Click
    
End Sub
