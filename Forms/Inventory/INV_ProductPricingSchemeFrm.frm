VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form INV_ProductPricingSchemeFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   14550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnSave 
      Caption         =   "Save && Close"
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
      Left            =   11400
      TabIndex        =   12
      Top             =   8040
      Width           =   1455
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
      Left            =   12960
      TabIndex        =   11
      Top             =   8040
      Width           =   1335
   End
   Begin VB.CommandButton btnAddProduct 
      Caption         =   "Add Products"
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
      Left            =   10080
      TabIndex        =   10
      Top             =   2280
      Width           =   4215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pricing Details"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   14055
      Begin VB.ComboBox cmbPricingScheme 
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
         ItemData        =   "INV_ProductPricingSchemeFrm.frx":0000
         Left            =   120
         List            =   "INV_ProductPricingSchemeFrm.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   840
         Width           =   2535
      End
      Begin VB.ComboBox cmbBasePrice 
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
         ItemData        =   "INV_ProductPricingSchemeFrm.frx":0019
         Left            =   3240
         List            =   "INV_ProductPricingSchemeFrm.frx":0023
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox cmbOperator 
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
         ItemData        =   "INV_ProductPricingSchemeFrm.frx":0032
         Left            =   5160
         List            =   "INV_ProductPricingSchemeFrm.frx":0042
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   840
         Width           =   2415
      End
      Begin VB.ComboBox cmbMode 
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
         ItemData        =   "INV_ProductPricingSchemeFrm.frx":0070
         Left            =   7680
         List            =   "INV_ProductPricingSchemeFrm.frx":007A
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox txtAmount 
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
         Height          =   345
         Left            =   10080
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   840
         Width           =   1815
      End
      Begin VB.CheckBox chkRoundOff 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Decimal Round off"
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
         Left            =   12000
         TabIndex        =   1
         Top             =   840
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Set Pricing Name"
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
         TabIndex        =   17
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Base Price"
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
         Left            =   3240
         TabIndex        =   9
         Top             =   480
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operator"
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
         Left            =   5160
         TabIndex        =   8
         Top             =   480
         Width           =   840
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mode"
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
         Left            =   7680
         TabIndex        =   7
         Top             =   480
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Value"
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
         Left            =   10080
         TabIndex        =   6
         Top             =   480
         Width           =   540
      End
   End
   Begin MSComctlLib.ListView lvSearch 
      Height          =   5175
      Left            =   240
      TabIndex        =   13
      Top             =   2760
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   9128
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
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
         Object.Width           =   6526
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Old Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Cost"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "New Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Cost"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Price"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "INV_ProductPricingSchemeFrm.frx":0093
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product Pricing"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Left            =   960
      TabIndex        =   15
      Top             =   360
      Width           =   1755
   End
   Begin VB.Label lblShowMorePrice 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select/Unselect All"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   225
      Left            =   240
      MouseIcon       =   "INV_ProductPricingSchemeFrm.frx":0724
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   8040
      Width           =   1530
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   8415
      Left            =   120
      Top             =   120
      Width           =   14295
   End
End
Attribute VB_Name = "INV_ProductPricingSchemeFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Populate(ByVal data As String)
    Select Case data
        Case "PricingScheme"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("PricingScheme")
            cmbPricingScheme.Clear
            cmbPricingScheme.AddItem ""
            cmbPricingScheme.ItemData(cmbPricingScheme.NewIndex) = 0
            
            cmbPricingScheme.ListIndex = 0
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        cmbPricingScheme.AddItem rec!PricingScheme
                        cmbPricingScheme.ItemData(cmbPricingScheme.NewIndex) = rec!PricingSchemeId
                    End If
                    rec.MoveNext
                Loop
            End If
        Case "BasePrice"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("PricingScheme")
            cmbBasePrice.Clear
            
            cmbBasePrice.AddItem ""
            cmbBasePrice.ItemData(cmbBasePrice.NewIndex) = -99
            cmbBasePrice.AddItem "Cost"
            cmbBasePrice.ItemData(cmbBasePrice.NewIndex) = -1
            cmbBasePrice.AddItem "Standard Price"
            cmbBasePrice.ItemData(cmbBasePrice.NewIndex) = -2
            
            cmbBasePrice.ListIndex = 0
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        cmbBasePrice.AddItem rec!PricingScheme
                        cmbBasePrice.ItemData(cmbBasePrice.NewIndex) = rec!PricingSchemeId
                    End If
                    rec.MoveNext
                Loop
            End If
    End Select
End Sub
Public Sub ComputePrice()
    Dim item As MSComctlLib.ListItem
    Dim NewPrice As Double
    
    For Each item In lvSearch.ListItems
        Select Case cmbOperator.text
            Case "Add"
                If cmbMode.ListIndex = 0 Then 'Percent
                    NewPrice = NVAL(item.SubItems(3)) + ((NVAL(item.SubItems(3)) * NVAL(txtAmount.text)) / 100)
                    item.SubItems(5) = FormatNumber(NewPrice, 2, vbTrue, vbFalse)
                Else 'Fixed Amount
                    NewPrice = NVAL(item.SubItems(3)) + NVAL(txtAmount.text)
                    item.SubItems(5) = FormatNumber(NewPrice, 2, vbTrue, vbFalse)
                End If
            Case "Subtract"
                If cmbMode.ListIndex = 0 Then 'Percent
                    NewPrice = NVAL(item.SubItems(3)) - ((NVAL(item.SubItems(3)) * NVAL(txtAmount.text)) / 100)
                    item.SubItems(5) = FormatNumber(NewPrice, 2, vbTrue, vbFalse)
                Else 'Fixed Amount
                    NewPrice = NVAL(item.SubItems(3)) - NVAL(txtAmount.text)
                    item.SubItems(5) = FormatNumber(NewPrice, 2, vbTrue, vbFalse)
                End If
            Case "Multiplied by"
                If cmbMode.ListIndex = 0 Then 'Percent
                    NewPrice = NVAL(item.SubItems(3)) * ((NVAL(item.SubItems(3)) * NVAL(txtAmount.text)) / 100)
                    item.SubItems(5) = FormatNumber(NewPrice, 2, vbTrue, vbFalse)
                Else 'Fixed Amount
                    NewPrice = NVAL(item.SubItems(3)) * NVAL(txtAmount.text)
                    item.SubItems(5) = FormatNumber(NewPrice, 2, vbTrue, vbFalse)
                End If
            Case "Divided by"
                If cmbMode.ListIndex = 0 Then 'Percent
                    If NVAL(item.SubItems(3)) = 0 Or NVAL(txtAmount.text) = 0 Then
                        NewPrice = 0
                    Else
                        NewPrice = NVAL(item.SubItems(3)) / ((NVAL(item.SubItems(3)) * NVAL(txtAmount.text)) / 100)
                    End If
                    item.SubItems(5) = FormatNumber(NewPrice, 2, vbTrue, vbFalse)
                Else 'Fixed Amount
                    If NVAL(item.SubItems(3)) = 0 Or NVAL(txtAmount.text) = 0 Then
                        NewPrice = 0
                    Else
                        NewPrice = NVAL(item.SubItems(3)) / NVAL(txtAmount.text)
                    End If
                    item.SubItems(5) = FormatNumber(NewPrice, 2, vbTrue, vbFalse)
                End If
        End Select
    Next
End Sub
Public Sub DisplayProducts()
    On Error Resume Next
    
    Dim exists As Boolean
    Dim item As MSComctlLib.ListItem
    
    If ProductSet.RecordCount <= 0 Then Exit Sub

    'Dim item As MSComctlLib.ListItem
    If Not ProductSet.EOF Then
        ProductSet.MoveFirst
        Do Until ProductSet.EOF
            For Each item In lvSearch.ListItems
                If item.text = ProductSet!ProductId Then
                    exists = True
                    Exit For
                End If
            Next

            If exists = False Then
                Set item = lvSearch.ListItems.add(, , ProductSet!ProductId)
                item.SubItems(1) = ProductSet!itemcode
                item.SubItems(2) = ProductSet!Name
                item.SubItems(4) = FormatNumber(ProductSet!cost, 2, vbTrue, vbFalse)
                item.SubItems(6) = ProductSet!cost
                item.SubItems(7) = ProductSet!price
            End If
            ProductSet.MoveNext
        Loop
    End If
End Sub
    
Public Sub GetPrice()
    Dim con As New ADODB.Connection
    Dim pRec As ADODB.Recordset
    Dim item As MSComctlLib.ListItem
        
    If cmbBasePrice.ItemData(cmbBasePrice.ListIndex) = -1 Then 'COST
        For Each item In lvSearch.ListItems
            item.SubItems(3) = FormatNumber(item.SubItems(6), 2, vbTrue, vbFalse)
        Next
        lvSearch.ColumnHeaders(4).text = "Cost"
        
        lvSearch.ColumnHeaders(2).width = lvSearch.width * 0.12
        lvSearch.ColumnHeaders(3).width = lvSearch.width * 0.547
        lvSearch.ColumnHeaders(4).width = lvSearch.width * 0.175
        lvSearch.ColumnHeaders(5).width = lvSearch.width * 0
        lvSearch.ColumnHeaders(6).width = lvSearch.width * 0.125
    
    ElseIf cmbBasePrice.ItemData(cmbBasePrice.ListIndex) = -2 Then 'SRP
        For Each item In lvSearch.ListItems
            item.SubItems(3) = FormatNumber(item.SubItems(7), 2, vbTrue, vbFalse)
        Next
        lvSearch.ColumnHeaders(4).text = "Suggested Retail Price"
        
        lvSearch.ColumnHeaders(2).width = lvSearch.width * 0.12
        lvSearch.ColumnHeaders(3).width = lvSearch.width * 0.422
        lvSearch.ColumnHeaders(4).width = lvSearch.width * 0.175
        lvSearch.ColumnHeaders(5).width = lvSearch.width * 0.125
        lvSearch.ColumnHeaders(6).width = lvSearch.width * 0.125
        
    ElseIf cmbBasePrice.ItemData(cmbBasePrice.ListIndex) = -99 Then 'Nothing
        For Each item In lvSearch.ListItems
            item.SubItems(3) = ""
        Next
        lvSearch.ColumnHeaders(4).text = "Old Price"
        
        lvSearch.ColumnHeaders(2).width = lvSearch.width * 0.12
        lvSearch.ColumnHeaders(3).width = lvSearch.width * 0.422
        lvSearch.ColumnHeaders(4).width = lvSearch.width * 0.175
        lvSearch.ColumnHeaders(5).width = lvSearch.width * 0.125
        lvSearch.ColumnHeaders(6).width = lvSearch.width * 0.125
    
    Else
        lvSearch.ColumnHeaders(4).text = cmbBasePrice.text
        
        lvSearch.ColumnHeaders(2).width = lvSearch.width * 0.12
        lvSearch.ColumnHeaders(3).width = lvSearch.width * 0.422
        lvSearch.ColumnHeaders(4).width = lvSearch.width * 0.175
        lvSearch.ColumnHeaders(5).width = lvSearch.width * 0.125
        lvSearch.ColumnHeaders(6).width = lvSearch.width * 0.125
    
        con.ConnectionString = ConnString
        con.Open
        For Each item In lvSearch.ListItems
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "INV_ProductPricing_Get"
            cmd.Parameters.Append cmd.CreateParameter("@PricingSchemeId", adInteger, adParamInput, , cmbBasePrice.ItemData(cmbBasePrice.ListIndex))
            cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , item.text)
            Set pRec = cmd.Execute
            If Not pRec.EOF Then
                item.SubItems(3) = FormatNumber(pRec!price, 2, vbTrue, vbFalse)
            Else
                item.SubItems(3) = ""
            End If
        Next
        con.Close
    
    End If
End Sub

Private Sub btnAddProduct_Click()
    INV_ProductSelectionFrm.Show (1)
    DisplayProducts
    ComputePrice
End Sub

Private Sub btnCancel_Click()
   Unload Me
End Sub

Private Sub btnRemoveSelected_Click()
    Dim x As Integer
    For x = 1 To lvSearch.ListItems.Count
        If x > lvSearch.ListItems.Count Then Exit For
        If lvSearch.ListItems(x).Selected = True Then
            lvSearch.ListItems.Remove (x)
            x = x - 1
        End If
    Next
End Sub

Private Sub btnSave_Click()
    If cmbPricingScheme.text = "" Then
        MsgBox "Please select a name for this product pricing.", vbCritical
        cmbPricingScheme.SetFocus
    ElseIf cmbBasePrice.text = "" Then
        MsgBox "Please select a base price.", vbCritical
        cmbBasePrice.SetFocus
    Else
        Dim item As MSComctlLib.ListItem
        Dim con As New ADODB.Connection
        Set rec = New ADODB.Recordset
        
        
        con.ConnectionString = ConnString
        con.Open
        
        For Each item In lvSearch.ListItems
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "INV_ProductPricing_Insert"
            cmd.Parameters.Append cmd.CreateParameter("@PricingSchemeId", adInteger, adParamInput, , cmbPricingScheme.ItemData(cmbPricingScheme.ListIndex))
            cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , item.text)
            cmd.Parameters.Append cmd.CreateParameter("@Price", adDecimal, adParamInput, , NVAL(item.SubItems(5)))
                                  cmd.Parameters("@Price").Precision = 18
                                  cmd.Parameters("@Price").NumericScale = 2
            cmd.Execute
        Next
        
        con.Close
        
        Unload Me
    End If
End Sub


Private Sub chkRoundOff_Click()
    ComputePrice
End Sub

Private Sub cmbBasePrice_Click()
    GetPrice
    ComputePrice
End Sub

Private Sub cmbMode_Click()
    ComputePrice
End Sub

Private Sub cmbOperator_Click()
    ComputePrice
End Sub

Private Sub cmbPricingScheme_Click()
    GetPrice
    ComputePrice
End Sub




Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
            btnRemoveSelected_Click
    End Select
End Sub

Private Sub Form_Load()
    lvSearch.ColumnHeaders(2).width = lvSearch.width * 0.12
    lvSearch.ColumnHeaders(3).width = lvSearch.width * 0.422
    lvSearch.ColumnHeaders(4).width = lvSearch.width * 0.175
    lvSearch.ColumnHeaders(5).width = lvSearch.width * 0.125
    lvSearch.ColumnHeaders(6).width = lvSearch.width * 0.125
    'lvSearch.ColumnHeaders(7).width = lvSearch.width * 0.125
    
    cmbBasePrice.ListIndex = 0
    cmbOperator.ListIndex = 0
    cmbMode.ListIndex = 0
    
    Populate "PricingScheme"
    Populate "BasePrice"
    Populate "Operator"
    
End Sub

Private Sub lblShowMorePrice_Click()
    Dim item As MSComctlLib.ListItem
    For Each item In lvSearch.ListItems
        If item.Selected = True Then
            item.Selected = False
        Else
            item.Selected = True
        End If
    Next
End Sub

Private Sub lvSearch_DblClick()
    Dim x As String
    x = InputBox("Input new price", "Price")
    If IsNumeric(x) = False Then
        MsgBox "Invalid Price.", vbCritical
    Else
        lvSearch.SelectedItem.SubItems(5) = FormatNumber(x, 2, vbTrue, vbFalse)
    End If
End Sub

Private Sub txtAmount_Change()
    If IsNumeric(txtAmount.text) = False Then
        txtAmount.text = "0.00"
    Else
        ComputePrice
    End If
End Sub

Private Sub txtAmount_GotFocus()
    selectText txtAmount
End Sub


