VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form POS_UomFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   Icon            =   "POS_UomFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   4680
      Picture         =   "POS_UomFrm.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6600
      Width           =   1575
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
      Left            =   2880
      Picture         =   "POS_UomFrm.frx":239B
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6600
      Width           =   1695
   End
   Begin MSComctlLib.ListView lvUom 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
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
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "UomId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Unit of Measure"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Price"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Qty"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "POS_UomFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ProductId As String

Private Sub btnAccept_Click()
    If POS_ItemSearchFrm.lvItemSearch.Visible = True Then Exit Sub
    If lvUom.ListItems.Count = 0 Then Exit Sub
    With POS_CashierFrm
        'CHECK INVENTORY
        Dim Available As Double
        If AllowNegativeInventory = False Then
            Available = CheckAvailableQuantity(.lvList.SelectedItem.SubItems(8))
            If Available + NVAL(.lvList.SelectedItem.SubItems(1)) * NVAL(.lvList.SelectedItem.SubItems(16)) _
            < NVAL(.lvList.SelectedItem.SubItems(1)) * NVAL(lvUom.SelectedItem.SubItems(3)) Then
                MsgBox "Insufficient quantity.", vbCritical, "Error!"
                Exit Sub
            End If
        End If
    
        .lvList.SelectedItem.SubItems(2) = lvUom.SelectedItem.SubItems(1)
        .lvList.SelectedItem.SubItems(3) = FormatNumber(lvUom.SelectedItem.SubItems(2), 2, vbTrue, vbFalse)
        .lvList.SelectedItem.SubItems(16) = lvUom.SelectedItem.SubItems(3)
        .CountTotal
        .CountTax
        Unload Me
    End With
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            btnAccept_Click
        Case vbKeyEscape
            btnCancel_Click
    End Select
End Sub

Private Sub Form_Load()
    lvUom.ColumnHeaders(2).width = lvUom.width * 0.6
    lvUom.ColumnHeaders(3).width = lvUom.width * 0.34
    
    'Get Uom Related
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "INV_UomConversion_Get"
    cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , Val(ProductId))
    Set rec = cmd.Execute
    lvUom.ListItems.Clear
    Dim item As MSComctlLib.ListItem
    If Not rec.EOF Then
        Do Until rec.EOF
            Set item = lvUom.ListItems.add(, , rec!UomConversionId)
                item.SubItems(1) = rec!Uom
                If IsNull(rec!price) = True Then
                    item.SubItems(2) = "0.00"
                Else
                    item.SubItems(2) = FormatNumber(rec!price, 2, vbTrue, vbFalse)
                End If
                If IsNull(rec!toqty) = True Then
                    item.SubItems(3) = 1
                Else
                    item.SubItems(3) = FormatNumber(rec!toqty, 2, vbTrue, vbFalse)
                End If
                rec.MoveNext
        Loop
    End If
    con.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ProductId = 0
End Sub
