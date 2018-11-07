VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form INV_ProductSelectionFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnAdd 
      Caption         =   "Add"
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
      Left            =   8760
      TabIndex        =   17
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CommandButton btnSelectOne 
      Caption         =   ">"
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
      Left            =   4920
      TabIndex        =   14
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton btnSelectAll 
      Caption         =   ">>"
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
      Left            =   4920
      TabIndex        =   13
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton btnDeselectOne 
      Caption         =   "<"
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
      Left            =   4920
      TabIndex        =   12
      Top             =   4920
      Width           =   615
   End
   Begin VB.CommandButton btnDeselectAll 
      Caption         =   "<<"
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
      Left            =   4920
      TabIndex        =   11
      Top             =   5520
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Product Search"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   10335
      Begin VB.TextBox txtItemCode 
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
         Left            =   1560
         MaxLength       =   50
         TabIndex        =   5
         Top             =   480
         Width           =   8655
      End
      Begin VB.TextBox txtName 
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
         Left            =   1560
         MaxLength       =   500
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   840
         Width           =   8655
      End
      Begin VB.ComboBox cmbCategory 
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
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1200
         Width           =   8655
      End
      Begin VB.CommandButton btnSearch 
         Caption         =   "Search"
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
         Left            =   8520
         TabIndex        =   2
         Top             =   1920
         Width           =   1695
      End
      Begin VB.ComboBox cmbSupplier 
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
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1560
         Width           =   8655
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item Code"
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
         TabIndex        =   9
         Top             =   480
         Width           =   960
      End
      Begin VB.Label Label8 
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
         TabIndex        =   8
         Top             =   840
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
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
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
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
         Top             =   1560
         Width           =   780
      End
   End
   Begin MSComctlLib.ListView lvSearch 
      Height          =   3615
      Left            =   120
      TabIndex        =   15
      Top             =   3240
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   6376
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
         Object.Width           =   6526
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cost"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Srp"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Category"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Unit"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvAdd 
      Height          =   3615
      Left            =   5880
      TabIndex        =   16
      Top             =   3240
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   6376
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
         Object.Width           =   6526
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Cost"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Srp"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "category"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "unit"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product Selection"
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
      Left            =   840
      TabIndex        =   10
      Top             =   120
      Width           =   2010
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "INV_ProductSelectionFrm.frx":0000
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "INV_ProductSelectionFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public isPriceMarkUp As Boolean
Public Sub Populate(ByVal data As String)
    Set rec = New ADODB.Recordset
    Select Case data
        Case "Category"
            Set rec = Global_Data("Category")
            cmbCategory.Clear
            cmbCategory.AddItem ""
            cmbCategory.ItemData(cmbCategory.NewIndex) = 0
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        cmbCategory.AddItem rec!Category
                        cmbCategory.ItemData(cmbCategory.NewIndex) = rec!CategoryId
                    End If
                    rec.MoveNext
                Loop
            End If
            cmbCategory.ListIndex = 0
        Case "Vendor"
            Set rec = Global_Data("Vendor")
            cmbSupplier.AddItem ""
            cmbSupplier.ItemData(cmbSupplier.NewIndex) = 0
            If Not rec.EOF Then
                Do Until rec.EOF
                    cmbSupplier.AddItem rec!Name
                    cmbSupplier.ItemData(cmbSupplier.NewIndex) = rec!VendorId
                    rec.MoveNext
                Loop
            End If
            cmbSupplier.ListIndex = 0
    End Select
End Sub
Private Sub RemoveExisting(ByVal a As MSComctlLib.ListView, ByVal b As MSComctlLib.ListView)
    Dim item As MSComctlLib.ListItem
    Dim aItem As MSComctlLib.ListItem
    Dim x As Integer
    
    For Each item In a.ListItems
        For Each aItem In b.ListItems
            If item.Text = aItem.Text Then
                item.Selected = True
            End If
        Next
    Next
    
    For x = 1 To a.ListItems.Count
        If x > a.ListItems.Count Then Exit For
        If a.ListItems(x).Selected = True Then
            a.ListItems.Remove (x)
            x = x - 1
        End If
    Next
End Sub

Private Sub btnAdd_Click()
    'Pass to ProductSet
    Dim item As MSComctlLib.ListItem
    Set ProductSet = Nothing
    Set ProductSet = New ADODB.Recordset
    
    ProductSet.Fields.Append "ProductId", adInteger
    ProductSet.Fields.Append "ItemCode", adVarChar, 500
    ProductSet.Fields.Append "Name", adVarChar, 4000
    ProductSet.Fields.Append "Cost", adDecimal
               ProductSet.Fields("Cost").Precision = 18
               ProductSet.Fields("Cost").NumericScale = 2
    ProductSet.Fields.Append "Price", adDecimal
               ProductSet.Fields("Price").Precision = 18
               ProductSet.Fields("Price").NumericScale = 2
    ProductSet.Fields.Append "Unit", adVarChar, 250
    ProductSet.Open

    For Each item In lvAdd.ListItems
        ProductSet.AddNew
        ProductSet!ProductId = item.Text
        ProductSet!itemcode = item.SubItems(1)
        ProductSet!Name = item.SubItems(2)
        ProductSet!cost = item.SubItems(3)
        ProductSet!price = item.SubItems(4)
        ProductSet!unit = item.SubItems(6)
        ProductSet.Update
    Next
    
    Unload Me
End Sub

Private Sub btnDeselectAll_Click()
    lvAdd.ListItems.Clear
    btnSearch_Click
End Sub

Private Sub btnDeselectOne_Click()
    Dim x As Integer
    For x = 1 To lvAdd.ListItems.Count
        If x > lvAdd.ListItems.Count Then Exit For
        If lvAdd.ListItems(x).Selected = True Then
            lvAdd.ListItems.Remove (x)
            x = x - 1
        End If
    Next
    
    RemoveExisting lvAdd, lvSearch
    btnSearch_Click
End Sub

Private Sub btnSearch_Click()
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_Product_Search5"
    cmd.Parameters.Append cmd.CreateParameter("@ItemCode", adVarChar, adParamInput, 50, txtItemCode.Text)
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, txtName.Text)
    If cmbCategory.ListIndex = 0 Then
        cmd.Parameters.Append cmd.CreateParameter("@CategoryId", adInteger, adParamInput, , Null)
    Else
        cmd.Parameters.Append cmd.CreateParameter("@CategoryId", adInteger, adParamInput, , cmbCategory.ItemData(cmbCategory.ListIndex))
    End If
    If cmbSupplier.ListIndex = 0 Then
        cmd.Parameters.Append cmd.CreateParameter("@VendorId", adInteger, adParamInput, , Null)
    Else
        cmd.Parameters.Append cmd.CreateParameter("@VendorId", adInteger, adParamInput, , cmbSupplier.ItemData(cmbSupplier.ListIndex))
    End If
    Set rec = cmd.Execute
    Dim item As MSComctlLib.ListItem
    
    lvSearch.ListItems.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            If rec!isActive = "True" Then
                Set item = lvSearch.ListItems.add(, , rec!ProductId)
                    item.SubItems(1) = rec!itemcode
                    item.SubItems(2) = rec!Name
                    item.SubItems(3) = rec!unitcost
                    item.SubItems(4) = rec!unitprice
                    item.SubItems(5) = rec!Category
                    item.SubItems(6) = rec!Uom
            End If
            rec.MoveNext
        Loop
    End If
    con.Close
    
    If lvSearch.ListItems.Count > 0 Then
        lvSearch.ListItems(1).Selected = False
    End If
'
    RemoveExisting lvSearch, lvAdd
End Sub

Private Sub btnSelectAll_Click()
    If lvSearch.ListItems.Count <= 0 Then Exit Sub
    Dim isFound As Boolean
    Dim item As MSComctlLib.ListItem
    Dim aItem As MSComctlLib.ListItem
    Dim nitem As MSComctlLib.ListItem
    
    For Each item In lvSearch.ListItems
        'If Item.Selected = True Then
            For Each aItem In lvAdd.ListItems
                If item.Text = aItem.Text Then
                    isFound = True
                   Exit For
                End If
            Next
            
            If isFound = False Then
                Set nitem = lvAdd.ListItems.add(, , item.Text)
                    nitem.SubItems(1) = item.SubItems(1)
                    nitem.SubItems(2) = item.SubItems(2)
                    nitem.SubItems(3) = item.SubItems(3)
                    nitem.SubItems(4) = item.SubItems(4)
                    nitem.SubItems(5) = item.SubItems(5)
                    nitem.SubItems(6) = item.SubItems(6)
            End If
        'End If
    Next
    
    lvSearch.ListItems.Clear
    'RemoveExisting lvSearch, lvAdd
End Sub

Private Sub btnSelectOne_Click()
    If lvSearch.ListItems.Count <= 0 Then Exit Sub
    Dim isFound As Boolean
    Dim item As MSComctlLib.ListItem
    Dim aItem As MSComctlLib.ListItem
    Dim nitem As MSComctlLib.ListItem
    
    For Each item In lvSearch.ListItems
        If item.Selected = True Then
            For Each aItem In lvAdd.ListItems
                If item.Text = aItem.Text Then
                    isFound = True
                   Exit For
                End If
            Next
            
            If isFound = False Then
                Set nitem = lvAdd.ListItems.add(, , item.Text)
                    nitem.SubItems(1) = item.SubItems(1)
                    nitem.SubItems(2) = item.SubItems(2)
                    nitem.SubItems(3) = item.SubItems(3)
                    nitem.SubItems(4) = item.SubItems(4)
                    nitem.SubItems(5) = item.SubItems(5)
                    nitem.SubItems(6) = item.SubItems(6)
            End If
        End If
    Next
    
    RemoveExisting lvSearch, lvAdd
End Sub

Private Sub Form_Load()
    lvSearch.ColumnHeaders(2).width = lvSearch.width * 0.2
    lvSearch.ColumnHeaders(3).width = lvSearch.width * 0.75
    lvAdd.ColumnHeaders(2).width = lvAdd.width * 0.2
    lvAdd.ColumnHeaders(3).width = lvAdd.width * 0.75
    
    Populate "Vendor"
    Populate "Category"
End Sub

Private Sub txtItemCode_Change()
    btnSearch_Click
End Sub

Private Sub txtName_Change()
    btnSearch_Click
End Sub
