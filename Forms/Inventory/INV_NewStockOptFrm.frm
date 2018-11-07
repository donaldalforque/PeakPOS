VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form INV_NewStockOptFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4905
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   4905
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCost 
      Alignment       =   1  'Right Justify
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
      Height          =   330
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   5
      Top             =   2040
      Width           =   3135
   End
   Begin VB.ComboBox cmbLocation 
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
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2760
      Width           =   3135
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
      Left            =   2160
      TabIndex        =   11
      Top             =   4320
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
      Left            =   3480
      TabIndex        =   12
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CheckBox chkHasExpiry 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Product has expiry date"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   3360
      Width           =   3015
   End
   Begin VB.TextBox txtLotNumber 
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
      Height          =   330
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   7
      Top             =   2400
      Width           =   3135
   End
   Begin VB.ComboBox cmbUnit 
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
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1680
      Width           =   3135
   End
   Begin VB.TextBox txtQuantity 
      Alignment       =   1  'Right Justify
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
      Height          =   330
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   2
      Text            =   "1"
      Top             =   1320
      Width           =   3135
   End
   Begin MSComCtl2.DTPicker dtExpiry 
      Height          =   345
      Left            =   1440
      TabIndex        =   10
      Top             =   3720
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   91095041
      CurrentDate     =   41686
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cost:"
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
      TabIndex        =   16
      Top             =   2040
      Width           =   465
   End
   Begin VB.Label Label4 
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
      Left            =   240
      TabIndex        =   15
      Top             =   2760
      Width           =   780
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   4680
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Expiry Date:"
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
      TabIndex        =   14
      Top             =   3720
      Width           =   1110
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lot/Batch #:"
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
      TabIndex        =   13
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label9 
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
      Left            =   240
      TabIndex        =   6
      Top             =   1680
      Width           =   390
   End
   Begin VB.Label Label7 
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
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   810
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Use this to input quantity details of your newly arrived products."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   4725
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Quantity Details"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2400
   End
End
Attribute VB_Name = "INV_NewStockOptFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function GetProductLastCost() As Double
    'Get Product Last Purchase Cost
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "INV_ProductLastCost_Get"
    
    If INV_NewStockFrm.lvItemList.Visible = True Then
        cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , INV_NewStockFrm.lvItemList.SelectedItem.Text)
    Else
        cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , INV_NewStockFrm.lvItems.SelectedItem.Text)
    End If
    Set rec = cmd.Execute
    If Not rec.EOF Then
        GetProductLastCost = rec!unitcost
    Else
        GetProductLastCost = 0
    End If
    con.Close
End Function

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    'Transfer data
    With INV_NewStockFrm
        If .lvItemList.Visible = True Then 'New List
            Dim item As MSComctlLib.ListItem
            Set item = .lvItems.ListItems.add(, , .lvItemList.SelectedItem.Text) 'ProductId
                item.SubItems(2) = cmbLocation.ItemData(cmbLocation.ListIndex) 'LocationId
                item.SubItems(4) = .lvItemList.SelectedItem.SubItems(1)
                item.SubItems(5) = .lvItemList.SelectedItem.SubItems(2)
                item.SubItems(6) = FormatNumber(NVAL(txtQuantity.Text), 2, vbTrue, vbFalse)
                item.SubItems(7) = cmbUnit.Text
                item.SubItems(8) = txtLotNumber.Text
                If chkHasExpiry.value = Checked Then
                    item.SubItems(9) = dtExpiry.value
                Else
                    item.SubItems(9) = ""
                End If
                item.SubItems(10) = FormatNumber(NVAL(txtCost.Text), 2, vbTrue, vbFalse)
                item.SubItems(12) = cmbLocation.Text
                item.SubItems(13) = cmbUnit.ItemData(cmbUnit.ListIndex)
                item.Selected = True
                item.EnsureVisible
                Unload Me
                selectText .txtItemSearch
                .txtItemSearch.SetFocus
        Else
            .lvItems.SelectedItem.SubItems(2) = cmbLocation.ItemData(cmbLocation.ListIndex)
            .lvItems.SelectedItem.SubItems(6) = FormatNumber(NVAL(txtQuantity.Text), 2, vbTrue, vbFalse)
            .lvItems.SelectedItem.SubItems(7) = cmbUnit.Text
            .lvItems.SelectedItem.SubItems(8) = txtLotNumber.Text
            If chkHasExpiry.value = Checked Then
                .lvItems.SelectedItem.SubItems(9) = dtExpiry.value
            Else
                .lvItems.SelectedItem.SubItems(9) = ""
            End If
            .lvItems.SelectedItem.SubItems(10) = FormatNumber(NVAL(txtCost.Text), 2, vbTrue, vbFalse)
            .lvItems.SelectedItem.SubItems(12) = cmbLocation.Text
            .lvItems.SelectedItem.SubItems(13) = cmbUnit.ItemData(cmbUnit.ListIndex)
            Unload Me
        End If
        .CountTotal
    End With
End Sub

Private Sub chkHasExpiry_Click()
    If chkHasExpiry.value = Checked Then
        dtExpiry.Enabled = True
    Else
        dtExpiry.Enabled = False
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            btnOK_Click
        Case vbKeyEscape
            btnCancel_Click
    End Select
End Sub

Private Sub Form_Load()
    Populate "Uom"
    Populate "Location"
    txtCost.Text = FormatNumber(GetProductLastCost, 2, vbTrue, vbFalse)
    
    dtExpiry.value = Format(Now, "MM/DD/YY")
    
    If PharmacyMode = "ON" Then
        chkHasExpiry.value = Checked
    Else
        chkHasExpiry.value = Unchecked
    End If
    chkHasExpiry_Click
    
    selectText txtQuantity
    
    With INV_NewStockFrm
        If .lvItemList.Visible = False Then
            cmbLocation.Text = .lvItems.SelectedItem.SubItems(12)
            cmbUnit.Text = .lvItems.SelectedItem.SubItems(7)
            txtCost.Text = .lvItems.SelectedItem.SubItems(10)
            txtLotNumber.Text = .lvItems.SelectedItem.SubItems(8)
            txtQuantity.Text = .lvItems.SelectedItem.SubItems(6)
            If .lvItems.SelectedItem.SubItems(9) <> "" Then
                dtExpiry.value = .lvItems.SelectedItem.SubItems(9)
            End If
            selectText txtQuantity
        End If
    End With
    
    cmbLocation.ListIndex = 0
End Sub

Private Sub txtCost_Change()
    If IsNumeric(txtCost.Text) = False And Not txtCost.Text = "" Then
        txtCost.Text = "0.00"
    End If
End Sub

Private Sub txtQuantity_Change()
'    If IsNumeric(txtQuantity.Text) = False And Not txtQuantity.Text = "" Then
'        txtQuantity.Text = "1"
'    End If
End Sub


Private Sub Populate(ByVal data As String)
    Select Case data
        Case "Uom"
            'Get Uom Related
            Set con = New ADODB.Connection
            Set rec = New ADODB.Recordset
            Set cmd = New ADODB.Command
            
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "INV_UomConversion_Get"
            
            If INV_NewStockFrm.lvItemList.Visible = True Then
                cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , INV_NewStockFrm.lvItemList.SelectedItem.Text)
            Else
                cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , INV_NewStockFrm.lvItems.SelectedItem.Text)
            End If
            Set rec = cmd.Execute
            cmbUnit.Clear
            Dim item As MSComctlLib.ListItem
            If Not rec.EOF Then
                Do Until rec.EOF
                    cmbUnit.AddItem rec!Uom
                    If IsNull(rec!toqty) = True Then
                        cmbUnit.ItemData(cmbUnit.NewIndex) = 0
                    Else
                        cmbUnit.ItemData(cmbUnit.NewIndex) = rec!toqty
                    End If
                    rec.MoveNext
                Loop
            End If
            con.Close
            
            On Error Resume Next
            cmbUnit.ListIndex = 0
        Case "Location"
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
            On Error Resume Next
            cmbLocation.ListIndex = 2
    End Select
End Sub
