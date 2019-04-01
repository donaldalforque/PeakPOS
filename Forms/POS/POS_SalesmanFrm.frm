VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form POS_SalesmanFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7110
   Icon            =   "POS_SalesmanFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnAll 
      Caption         =   "F2: Print All Sales"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      Picture         =   "POS_SalesmanFrm.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton btnPrint 
      Caption         =   "F1: Print Sales"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Picture         =   "POS_SalesmanFrm.frx":2234
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton btnAccept 
      Caption         =   "ENTER: Accept"
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
      Picture         =   "POS_SalesmanFrm.frx":445C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "ESC: Cancel"
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
      Left            =   5400
      Picture         =   "POS_SalesmanFrm.frx":6830
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5400
      Width           =   1575
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   3375
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   5953
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
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "CustomerId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Salesman"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton btnReturn 
      Caption         =   "F4"
      Height          =   375
      Left            =   6360
      TabIndex        =   3
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtSalesman 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   6015
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select Salesman"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1395
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   4575
      Left            =   120
      Top             =   720
      Width           =   6855
   End
   Begin VB.Label lblCaption_Title 
      AutoSize        =   -1  'True
      Caption         =   "Salesman"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   1200
   End
   Begin VB.Image picModuleImage 
      Height          =   480
      Left            =   120
      Picture         =   "POS_SalesmanFrm.frx":8BBF
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "POS_SalesmanFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub LoadSalesmansOnPOS()
    Dim Item As MSComctlLib.ListItem
    Set rec = New ADODB.Recordset
    Set rec = Global_Data("Salesman")
    If Not rec.EOF Then
        lvList.ListItems.Clear
        Do Until rec.EOF
            Set Item = lvList.ListItems.add(, , rec!SalesmanId)
                Item.SubItems(1) = rec!Salesman
            rec.MoveNext
        Loop
    End If
End Sub

Private Sub btnAccept_Click()
    If lvList.ListItems.count > 0 Then
        POS_CashierFrm.lblSalesman.Caption = "| Salesman: " & lvList.SelectedItem.SubItems(1)
        POS_CashierFrm.SalesmanId = lvList.SelectedItem.Text
        Unload Me
    End If
End Sub

Private Sub btnAccountsReceivable_Click()
    POS_ViewAccountsFrm.Show (1)
End Sub

Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub btnNewSalesman_Click()
    
End Sub

Private Sub btnPrint_Click()
    Dim crxApp As New CRAXDRT.Application
    Dim crxRpt As New CRAXDRT.Report
    
    Screen.MousePointer = vbHourglass
    Set crxRpt = crxApp.OpenReport(App.path & "\Reports\POS_SalesmanReceipt.rpt")
    
    Call ResetRptDB(crxRpt)
    
    crxRpt.EnableParameterPrompting = False
    crxRpt.DiscardSavedData
    
    crxRpt.ParameterFields.GetItemByName("@SalesmanId").AddCurrentValue lvList.SelectedItem.Text
    crxRpt.PrintOut False
    Screen.MousePointer = vbDefault
    
    'POS Audit Trail
    SavePOSAuditTrail VoidUserId, WorkstationId, 0, "Generate Z-Reading Report"
End Sub

Private Sub btnReturn_Click()
    txtSalesman.SelStart = 0
    txtSalesman.SelLength = Len(txtSalesman.Text)
    txtSalesman.SetFocus
End Sub

Private Sub Command1_Click()
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyReturn
        btnAccept_Click
    Case vbKeyEscape
        btnCancel_Click
    Case vbKeyF4
        btnReturn_Click
    Case vbKeyF1
        btnNewSalesman_Click
    Case vbKeyF2
        btnAccountsReceivable_Click
End Select
End Sub

Private Sub Form_Load()
    lvList.ColumnHeaders(2).width = (lvList.width * 0.96)  'Salesman
    LoadSalesmansOnPOS
End Sub

Public Sub txtSalesman_Change()
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_Salesman_Search"
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 500, txtSalesman.Text)

    Set rec = cmd.Execute
    lvList.ListItems.Clear
    Dim Item As MSComctlLib.ListItem
    If Not rec.EOF Then
        Do Until rec.EOF
            If rec!isActive = "True" Then
                Set Item = lvList.ListItems.add(, , rec!SalesmanId)
                    Item.SubItems(1) = rec!Salesman
            End If
            rec.MoveNext
        Loop
    End If
    con.Close
End Sub

Private Sub txtSalesman_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            If lvList.ListItems.count > 0 Then
                lvList.SetFocus
            End If
    End Select
End Sub

Private Sub txtCustomer_Change()

End Sub
