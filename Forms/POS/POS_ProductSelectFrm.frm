VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form POS_ProductSelectFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14190
   Icon            =   "POS_ProductSelectFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   14190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame_Body 
      BackColor       =   &H00FFFFFF&
      Height          =   8100
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   13935
      Begin VB.TextBox txtCode 
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
         Top             =   240
         Width           =   5175
      End
      Begin VB.CommandButton btnSave 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9960
         TabIndex        =   5
         Top             =   7560
         Width           =   1815
      End
      Begin VB.TextBox txtItemSearch 
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
         Top             =   600
         Width           =   5175
      End
      Begin VB.CommandButton btnItemSearch 
         Height          =   330
         Left            =   6480
         Picture         =   "POS_ProductSelectFrm.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   12000
         TabIndex        =   6
         Top             =   7560
         Width           =   1815
      End
      Begin MSComctlLib.ListView lvItemList 
         Height          =   6495
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   11456
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
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ProductId"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Item Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Unit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Qty on Hand"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Qty Reserved"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Qty Available"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Cost"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Price"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label9 
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
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name/Desc"
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
         Top             =   600
         Width           =   1080
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "POS_ProductSelectFrm.frx":0230
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search Product"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   435
      Left            =   720
      TabIndex        =   9
      Top             =   120
      Width           =   2220
   End
End
Attribute VB_Name = "POS_ProductSelectFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public isWithdraw, isPO, isSO, isAS As Boolean
Public Sub Populate(ByVal data As String)
    Select Case data
'        Case "Status"
'            Set rec = New ADODB.Recordset
'            Set rec = Global_Data("Status")
'            cmbSearch_Status.Clear
'            cmbSearch_Status.AddItem ""
'            cmbSearch_Status.ItemData(cmbSearch_Status.NewIndex) = 0
'            cmbSearch_Status.ListIndex = 0
'            If Not rec.EOF Then
'                Do Until rec.EOF
'                    cmbSearch_Status.AddItem rec!Status
'                    cmbSearch_Status.ItemData(cmbSearch_Status.NewIndex) = rec!StatusId
'                    rec.MoveNext
'                Loop
'            End If
        Case "Product"
            Set con = New ADODB.Connection
            Set rec = New ADODB.Recordset
            Set cmd = New ADODB.Command
            Dim item As MSComctlLib.ListItem
            
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "BASE_Product_Search4"
            cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 500, "")
            
            Dim QtyonHand, Reserved As Double
            Dim PrevProductId As Long
            
            Set rec = cmd.Execute
            If Not rec.EOF Then
                lvItemList.ListItems.Clear
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        If PrevProductId <> rec!ProductId Then
                            QtyonHand = 0
                            Reserved = 0
                            Set item = lvItemList.ListItems.add(, , rec!ProductId)
                                item.SubItems(1) = rec!itemcode
                                item.SubItems(2) = rec!Name
                                If rec!Location = "<RESERVED>" Then
                                    Reserved = Reserved + rec!quantity
                                Else
                                    QtyonHand = QtyonHand + rec!quantity
                                End If
                                item.SubItems(3) = rec!Uom
                                item.SubItems(4) = FormatNumber(QtyonHand, 2, vbTrue, vbFalse)
                                item.SubItems(5) = FormatNumber(Reserved, 2, vbTrue, vbFalse)
                                item.SubItems(6) = FormatNumber(QtyonHand - Reserved, 2, vbTrue, vbFalse)
                                item.SubItems(7) = FormatNumber(rec!unitcost, 2, vbTrue)
                                item.SubItems(8) = FormatNumber(rec!unitprice, 2, vbTrue)
                        Else
                                If rec!LocationId = 4 Then
                                    Reserved = Reserved + rec!quantity
                                Else
                                    QtyonHand = QtyonHand + rec!quantity
                                End If
                                lvItemList.ListItems(lvItemList.ListItems.Count).SubItems(4) = FormatNumber(QtyonHand, 2, vbTrue, vbFalse)
                                lvItemList.ListItems(lvItemList.ListItems.Count).SubItems(5) = FormatNumber(Reserved, 2, vbTrue, vbFalse)
                                lvItemList.ListItems(lvItemList.ListItems.Count).SubItems(6) = FormatNumber(QtyonHand - Reserved, 2, vbTrue, vbFalse)
                        End If
                        PrevProductId = rec!ProductId
'                        lvItemList.Top = 2520
'                        lvItemList.Visible = True
'                        lvItemList.Left = 6070
                    End If
                    rec.MoveNext
                Loop
            Else
'                lvItemList.Visible = False
'                lvItemList.Left = -9999
            End If
            'DistinctList lvItemList
            con.Close
            BASE_ContainerFrm.statusBar_Main.Panels(1).text = "Total Items: " & lvItemList.ListItems.Count
    End Select
End Sub

Public Sub btnItemSearch_Click()
'    Dim itmFound As MSComctlLib.ListItem   ' FoundItem variable.
'
'   Set itmFound = lvItemList.FindItem(txtCode.text, lvwSubItem, , lvwPartial)
'   'Set itmFound = lvItemList.FindItem("Holcim", 2, , lvItemList)
'
'   ' If no ListItem is found, then inform user and exit. If a
'   ' ListItem is found, scroll the control using the EnsureVisible
'   ' method, and select the ListItem.
'   If itmFound Is Nothing Then  ' If no match, inform user and exit.
'      MsgBox "No match found"
'      Exit Sub
'   Else
'       itmFound.EnsureVisible ' Scroll ListView to show found ListItem.
'       itmFound.Selected = True   ' Select the ListItem.
'      ' Return focus to the control to see selection.
'       lvItemList.SetFocus
'   End If
    Dim item As MSComctlLib.ListItem
    For Each item In lvItemList.ListItems
        If UCase(item.SubItems(1)) Like UCase(txtCode.text) & "*" Then
            item.Selected = True
            item.EnsureVisible
            'lvItemList.SetFocus
            Exit For
        End If
    Next
End Sub

Private Sub btnSave_Click()
    If lvItemList.ListItems.Count > 0 Then
        POS_DisplayFrm.lvUoms.SelectedItem.SubItems(2) = lvItemList.SelectedItem.SubItems(2)
        POS_DisplayFrm.lvUoms.SelectedItem.SubItems(3) = lvItemList.SelectedItem.text
        Unload Me
    End If
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
        Case vbKeyF4
            txtCode.SetFocus
        Case vbKeyF5
            txtItemSearch.SetFocus
        Case vbKeyReturn
            btnSave_Click
    End Select
End Sub

Private Sub Form_Load()
    '****** REGION Listview Columns ********
    lvItemList.ColumnHeaders(2).width = lvItemList.width * 0.113
    lvItemList.ColumnHeaders(3).width = lvItemList.width * 0.435
    lvItemList.ColumnHeaders(4).width = lvItemList.width * 0.107
    lvItemList.ColumnHeaders(5).width = lvItemList.width * 0.107
    Populate "Product"
End Sub

Private Sub lvItemList_DblClick()
    If lvItemList.ListItems.Count > 0 Then
        btnSave_Click
    End If
End Sub

Private Sub txtCode_Change()
    btnItemSearch_Click
End Sub

Private Sub txtCode_GotFocus()
    selectText txtCode
End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            lvItemList.SetFocus
    End Select
End Sub

Public Sub txtItemSearch_Change()
    Dim item As MSComctlLib.ListItem
    For Each item In lvItemList.ListItems
        If UCase(item.SubItems(2)) Like UCase(txtItemSearch.text) & "*" Then
            item.Selected = True
            item.EnsureVisible
            'lvItemList.SetFocus
            Exit For
        End If
    Next
End Sub

Private Sub txtItemSearch_GotFocus()
    selectText txtItemSearch
End Sub

Private Sub txtItemSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            lvItemList.SetFocus
    End Select
End Sub

