VERSION 5.00
Begin VB.Form INV_AdjustStockOptFrm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.ComboBox cmbUnit 
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
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   480
         Width           =   2535
      End
      Begin VB.ComboBox cmbLocation 
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
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox txtQuantity 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   1
         Text            =   "1"
         Top             =   120
         Width           =   2535
      End
      Begin VB.CommandButton btnOk 
         Caption         =   "OK"
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
         Left            =   1440
         TabIndex        =   3
         Top             =   1560
         Width           =   1215
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
         Left            =   2760
         TabIndex        =   4
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
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
         Top             =   480
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
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
         Top             =   840
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
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
         TabIndex        =   5
         Top             =   120
         Width           =   810
      End
   End
End
Attribute VB_Name = "INV_AdjustStockOptFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public isFormSearch As Boolean
Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOk_Click()
    If IsNumeric(Val(Replace(txtQuantity.text, ",", ""))) = False Then txtQuantity.text = "1"
    Dim item As MSComctlLib.ListItem
    With INV_AdjustStockFrm
        If isFormSearch = False Then
            Set item = .lvItems.ListItems.add(, , .lvItemList.SelectedItem.text)
                item.SubItems(1) = .lvItemList.SelectedItem.SubItems(1)
                item.SubItems(2) = .lvItemList.SelectedItem.SubItems(2)
                item.SubItems(3) = cmbLocation.text
                item.SubItems(4) = FormatNumber(txtQuantity.text, 2, vbTrue)
                item.SubItems(5) = cmbUnit.text '.lvItemList.SelectedItem.SubItems(4)
                item.SubItems(6) = cmbLocation.ItemData(cmbLocation.ListIndex)
                
                Unload Me
                .lvItemList.Visible = False
                .lvItemList.Left = -9999
                item.Selected = True
                item.EnsureVisible
                .txtItemSearch.SetFocus
                '.txtCode.SetFocus
        Else
            Set item = .lvItems.ListItems.add(, , INV_ProductSearch.lvItemList.SelectedItem.text)
                item.SubItems(1) = INV_ProductSearch.lvItemList.SelectedItem.SubItems(1)
                item.SubItems(2) = INV_ProductSearch.lvItemList.SelectedItem.SubItems(2)
                item.SubItems(3) = cmbLocation.text
                item.SubItems(4) = FormatNumber(txtQuantity.text, 2, vbTrue)
                item.SubItems(5) = cmbUnit.text 'INV_ProductSearch.lvItemList.SelectedItem.SubItems(3)
                item.SubItems(6) = cmbLocation.ItemData(cmbLocation.ListIndex)
                
                Unload Me
                Unload INV_ProductSearch
                .lvItemList.Visible = False
                .lvItemList.Left = -9999
                item.Selected = True
                item.EnsureVisible
                .txtItemSearch.SetFocus
                '.txtCode.SetFocus
        End If
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            btnOk_Click
        Case vbKeyEscape
            btnCancel_Click
    End Select
End Sub

Private Sub Form_Load()
    'txtQuantity.text = "1.00"
    'selectText txtQuantity
    Set rec = New ADODB.Recordset
    Set rec = Global_Data("Location")
    cmbLocation.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            If rec!isActive = "True" Then
                cmbLocation.AddItem rec!Location
                cmbLocation.ItemData(cmbLocation.NewIndex) = rec!LocationId
                
            End If
            rec.MoveNext
        Loop
    End If
    
    'Get Uom Related
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "INV_UomConversion_Get"
    
    If isModify = True Then
        cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , INV_AdjustStockFrm.lvItems.SelectedItem.SubItems(9))
    Else
        cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , INV_AdjustStockFrm.lvItemList.SelectedItem.text)
    End If
    Set rec = cmd.Execute
    'lvUom.ListItems.Clear
    cmbUnit.Clear
    Dim item As MSComctlLib.ListItem
    If Not rec.EOF Then
        Do Until rec.EOF
            'Set item = lvUom.ListItems.add(, , rec!UomConversionId)
                'item.SubItems(1) = rec!Uom
                cmbUnit.AddItem rec!Uom
                If IsNull(rec!toqty) = True Then
                    'item.SubItems(2) = "0.00"
                    cmbUnit.ItemData(cmbUnit.NewIndex) = 0
                Else
                    cmbUnit.ItemData(cmbUnit.NewIndex) = rec!toqty
                End If
'                If IsNull(rec!toqty) = True Then
'                    item.SubItems(3) = 1
'                Else
'                    item.SubItems(3) = FormatNumber(rec!toqty, 2, vbTrue, vbFalse)
'                End If
                rec.MoveNext
        Loop
    End If
    con.Close
    
    On Error Resume Next
    cmbUnit.ListIndex = 0
    cmbLocation.text = "Store"
End Sub

Private Sub txtQuantity_Change()
    'If IsNumeric(txtQuantity.text) = False Then txtQuantity.text = "1"
End Sub

Private Sub txtQuantity_GotFocus()
    selectText txtQuantity
End Sub
