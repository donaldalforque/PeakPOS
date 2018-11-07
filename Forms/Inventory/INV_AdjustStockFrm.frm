VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form INV_AdjustStockFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14145
   Icon            =   "INV_AdjustStockFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   14145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvItemList 
      Height          =   2655
      Left            =   1560
      TabIndex        =   4
      Top             =   1160
      Visible         =   0   'False
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
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
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Uom"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Frame Frame_Body 
      BackColor       =   &H00FFFFFF&
      Height          =   7620
      Left            =   120
      TabIndex        =   0
      Top             =   600
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
         Left            =   10800
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.CommandButton btnSave 
         Caption         =   "Save && Close"
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
         Left            =   11280
         TabIndex        =   9
         Top             =   7080
         Width           =   1215
      End
      Begin VB.CommandButton btnCancel 
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
         Left            =   12600
         TabIndex        =   10
         Top             =   7080
         Width           =   1215
      End
      Begin VB.TextBox txtRemarks 
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
         TabIndex        =   8
         Top             =   6720
         Width           =   4455
      End
      Begin VB.TextBox txtReferenceNumber 
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
         TabIndex        =   6
         Top             =   6000
         Width           =   4455
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
         Top             =   240
         Width           =   3015
      End
      Begin VB.CommandButton btnItemSearch 
         Height          =   330
         Left            =   4320
         Picture         =   "INV_AdjustStockFrm.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   225
         Width           =   375
      End
      Begin MSComctlLib.ListView lvItems 
         Height          =   5295
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   9340
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
            Text            =   "Location"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Quantity"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Unit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "LocationId"
            Object.Width           =   0
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtAdjust 
         Height          =   330
         Left            =   1320
         TabIndex        =   7
         Top             =   6360
         Width           =   4455
         _ExtentX        =   7858
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
         Format          =   93323265
         CurrentDate     =   41509
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
         Left            =   9600
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
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
         TabIndex        =   14
         Top             =   6720
         Width           =   795
      End
      Begin VB.Label Label6 
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
         Left            =   120
         TabIndex        =   13
         Top             =   6360
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ref #"
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
         TabIndex        =   12
         Top             =   6000
         Width           =   465
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name/Code"
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
         Top             =   240
         Width           =   1125
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "INV_AdjustStockFrm.frx":0230
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Adjust Stock"
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
      Left            =   840
      TabIndex        =   15
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "INV_AdjustStockFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnItemSearch_Click()
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    Dim item As MSComctlLib.ListItem
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_Product_Search1"
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 500, txtItemSearch.text)
    Dim LastProductId As Long
    Set rec = cmd.Execute
    If Not rec.EOF Then
        lvItemList.ListItems.Clear
        Do Until rec.EOF
            If rec!isActive = "True" Then
                If LastProductId <> rec!ProductId Then
                    Set item = lvItemList.ListItems.add(, , rec!ProductId)
                        item.SubItems(1) = rec!itemcode
                        item.SubItems(2) = rec!Name
                        item.SubItems(3) = FormatNumber(rec!unitprice, 2, vbTrue, vbFalse)
                        item.SubItems(4) = rec!Uom
                    lvItemList.Visible = True
                    lvItemList.Left = 1440
                    lvItemList.Top = 1160
                    LastProductId = rec!ProductId
                    rec.MoveNext
                Else
                    rec.MoveNext
                End If
            Else
                rec.MoveNext
            End If
        Loop
    Else
        lvItemList.Visible = False
        lvItemList.Left = -9999
    End If
    'DistinctList lvItemList
    con.Close
End Sub

Private Sub btnSave_Click()
    If EditAccessRights(6) = False Then
        MsgBox ErrorCodes(74), vbCritical, "Limited Rights"
        Exit Sub
    End If
    If lvItems.ListItems.Count <= 0 Then Exit Sub
    Dim x As Variant
    x = MsgBox("Are all information correct?", vbQuestion + vbYesNo, "")
        If x = vbYes Then
        Set con = New ADODB.Connection
        Set rec = New ADODB.Recordset
        
        con.ConnectionString = ConnString
        con.Open
        con.BeginTrans
        
        Dim item As MSComctlLib.ListItem
        
        For Each item In lvItems.ListItems
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "INV_AdjustStock_Insert"
            cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , item.text)
            cmd.Parameters.Append cmd.CreateParameter("@LocationId", adInteger, adParamInput, , item.SubItems(6))
            cmd.Parameters.Append cmd.CreateParameter("@Quantity", adDecimal, adParamInput, , Val(Replace(item.SubItems(4), ",", "")))
                                  cmd.Parameters("@Quantity").NumericScale = 2
                                  cmd.Parameters("@Quantity").Precision = 18
            cmd.Parameters.Append cmd.CreateParameter("@Uom", adVarChar, adParamInput, 250, item.SubItems(5))
            cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtAdjust.value)
            cmd.Parameters.Append cmd.CreateParameter("@ReferenceNumber", adVarChar, adParamInput, 50, txtReferenceNumber.text)
            cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, txtRemarks.text)
            cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
            cmd.Execute
        Next
        
        con.CommitTrans
        con.Close
        MsgBox "Record saved.", vbInformation, ""
        Unload Me
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF4
            txtItemSearch.SetFocus
        Case vbKeyF1
            Unload INV_ProductSearch
            INV_ProductSearch.isPO = False
            INV_ProductSearch.isWithdraw = False
            INV_ProductSearch.isSO = False
            INV_ProductSearch.isAS = True
            INV_ProductSearch.Show (1)
        Case vbKeyEscape
            lvItemList.Visible = False
            txtItemSearch.SetFocus
    End Select
End Sub

Private Sub Form_Load()
    lvItems.ColumnHeaders(2).width = lvItems.width * 0.136
    lvItems.ColumnHeaders(3).width = lvItems.width * 0.441
    lvItems.ColumnHeaders(4).width = lvItems.width * 0.196
    lvItems.ColumnHeaders(5).width = lvItems.width * 0.11
    lvItems.ColumnHeaders(6).width = lvItems.width * 0.095
    
    lvItemList.ColumnHeaders(2).width = lvItemList.width * 0.3
    lvItemList.ColumnHeaders(3).width = lvItemList.width * 0.65
    lvItemList.ColumnHeaders(4).width = 0
    
    lvItemList.Left = -9999
    dtAdjust.value = Format(Now, "MM/DD/YY")
End Sub

Private Sub lvItemList_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If lvItemList.ListItems.Count > 0 Then
                INV_AdjustStockOptFrm.isFormSearch = False
                INV_AdjustStockOptFrm.txtQuantity.text = "1"
                INV_AdjustStockOptFrm.Show (1)
            End If
    End Select
End Sub

Private Sub lvItems_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp
            If lvItems.ListItems.Count > 0 Then
                If lvItems.SelectedItem.Index = 1 Then
                    txtItemSearch.SetFocus
                End If
            End If
        Case vbKeyDown
            If lvItems.ListItems.Count > 0 Then
                If lvItems.SelectedItem.Index = lvItems.ListItems.Count Then
                    txtItemSearch.SetFocus
                End If
            End If
        Case vbKeyDelete
            If lvItems.ListItems.Count > 0 Then
                lvItems.ListItems.Remove (lvItems.SelectedItem.Index)
                txtItemSearch.SetFocus
            End If
    End Select
End Sub

Private Sub txtCode_Change()
    If Trim(txtCode.text) = "" Then
        lvItemList.Visible = False
        Exit Sub
    End If
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    Dim item As MSComctlLib.ListItem
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_Product_Search1"
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 500, Null)
    cmd.Parameters.Append cmd.CreateParameter("@ItemCode", adVarChar, adParamInput, 50, txtCode.text)
    Set rec = cmd.Execute
    Dim LastProductId As Long
    If Not rec.EOF Then
        lvItemList.ListItems.Clear
        Do Until rec.EOF
            If rec!ProductId <> LastProductId Then
                If rec!isActive = "True" Then
                    Set item = lvItemList.ListItems.add(, , rec!ProductId)
                        item.SubItems(1) = rec!itemcode
                        item.SubItems(2) = rec!Name
                        item.SubItems(3) = FormatNumber(rec!unitprice, 2, vbTrue, vbFalse)
                        item.SubItems(4) = rec!Uom
                    lvItemList.Visible = True
                    lvItemList.Left = 1440
                    lvItemList.Top = 1170
                    LastProductId = rec!ProductId
                End If
                rec.MoveNext
            Else
                rec.MoveNext
            End If
        Loop
    Else
        lvItemList.Visible = False
        lvItemList.Left = -9999
    End If
    'DistinctList lvItemList
    con.Close
End Sub

Private Sub txtCode_GotFocus()
    selectText txtCode
End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp, vbKeyDown
            If lvItemList.Visible = True Then
                lvItemList.SetFocus
            Else
                lvItems.SetFocus
            End If
    End Select
End Sub

Private Sub txtItemSearch_Change()
    btnItemSearch_Click
End Sub

Private Sub txtItemSearch_GotFocus()
    selectText txtItemSearch
End Sub

Private Sub txtItemSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(txtItemSearch.text) = "" Then Exit Sub
            btnItemSearch_Click
        Case vbKeyUp, vbKeyDown
            If lvItemList.Visible = True Then
                lvItemList.SetFocus
            Else
                lvItems.SetFocus
            End If
    End Select
End Sub
