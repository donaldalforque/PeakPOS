VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form INV_AddProductFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Products"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14310
   Icon            =   "INV_AddProductFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   14310
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
      Left            =   12480
      TabIndex        =   12
      Top             =   7920
      Width           =   1695
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
      Left            =   6840
      TabIndex        =   10
      Top             =   5760
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
      Left            =   6840
      TabIndex        =   9
      Top             =   5160
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
      Left            =   6840
      TabIndex        =   8
      Top             =   4560
      Width           =   615
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
      Left            =   6840
      TabIndex        =   7
      Top             =   3960
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
      Top             =   120
      Width           =   14055
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
         TabIndex        =   4
         Top             =   1560
         Width           =   12375
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
         Left            =   12240
         TabIndex        =   5
         Top             =   1920
         Width           =   1695
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
         Width           =   12375
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
         TabIndex        =   2
         Top             =   840
         Width           =   12375
      End
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
         TabIndex        =   1
         Top             =   480
         Width           =   12375
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
         TabIndex        =   16
         Top             =   1560
         Width           =   780
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
         TabIndex        =   15
         Top             =   1200
         Width           =   825
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
         TabIndex        =   14
         Top             =   840
         Width           =   555
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
         TabIndex        =   13
         Top             =   480
         Width           =   960
      End
   End
   Begin MSComctlLib.ListView lvSearch 
      Height          =   5175
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   6615
      _ExtentX        =   11668
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
      Height          =   5175
      Left            =   7560
      TabIndex        =   11
      Top             =   2640
      Width           =   6615
      _ExtentX        =   11668
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
End
Attribute VB_Name = "INV_AddProductFrm"
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
    Dim X As Integer
    
    For Each item In a.ListItems
        For Each aItem In b.ListItems
            If item.text = aItem.text Then
                item.Selected = True
            End If
        Next
    Next
    
    For X = 1 To a.ListItems.Count
        If X > a.ListItems.Count Then Exit For
        If a.ListItems(X).Selected = True Then
            a.ListItems.Remove (X)
            X = X - 1
        End If
    Next
End Sub

Private Sub btnAdd_Click()
    If lvAdd.ListItems.Count <= 0 Then Exit Sub
    
    Dim isFound As Boolean
    Dim item As MSComctlLib.ListItem
    Dim aItem As MSComctlLib.ListItem
    Dim sitem As MSComctlLib.ListItem
    
    If isPriceMarkUp = True Then
        isFound = False
        For Each item In lvAdd.ListItems
            For Each aItem In INV_SRPSetupFrm.lvSearch.ListItems
                If item.text = aItem.text Then
                    isFound = True
                   Exit For
                End If
            Next
            
            If isFound = False Then
                Set sitem = INV_SRPSetupFrm.lvSearch.ListItems.add(, , item.text)
                    sitem.SubItems(1) = item.SubItems(1)
                    sitem.SubItems(2) = item.SubItems(2)
                    sitem.SubItems(3) = item.SubItems(5)
                    sitem.SubItems(4) = FormatNumber(item.SubItems(3), 2, vbTrue, vbFalse)
                    sitem.SubItems(5) = FormatNumber(item.SubItems(4), 2, vbTrue, vbFalse)
            End If
        Next
        INV_SRPSetupFrm.ComputeMarkUp
    Else
        For Each item In lvAdd.ListItems
            isFound = False
            For Each aItem In INV_ProductConversionFrm.lvSearch.ListItems
                If item.text = aItem.text Then
                    isFound = True
                   Exit For
                End If
            Next
            
            If isFound = False Then
                Set sitem = INV_ProductConversionFrm.lvSearch.ListItems.add(, , item.text) 'From ProductId
                    sitem.SubItems(1) = item.SubItems(1)
                    sitem.SubItems(2) = item.SubItems(2)
                    sitem.SubItems(3) = "1"
                    sitem.SubItems(4) = item.SubItems(6)
                    
            End If
        Next
    End If
    
    Unload Me
End Sub

Private Sub btnDeselectAll_Click()
    lvAdd.ListItems.Clear
    btnSearch_Click
End Sub

Private Sub btnDeselectOne_Click()
    Dim X As Integer
    For X = 1 To lvAdd.ListItems.Count
        If X > lvAdd.ListItems.Count Then Exit For
        If lvAdd.ListItems(X).Selected = True Then
            lvAdd.ListItems.Remove (X)
            X = X - 1
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
    cmd.Parameters.Append cmd.CreateParameter("@ItemCode", adVarChar, adParamInput, 50, txtItemCode.text)
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, txtName.text)
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
                If item.text = aItem.text Then
                    isFound = True
                   Exit For
                End If
            Next
            
            If isFound = False Then
                Set nitem = lvAdd.ListItems.add(, , item.text)
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
                If item.text = aItem.text Then
                    isFound = True
                   Exit For
                End If
            Next
            
            If isFound = False Then
                Set nitem = lvAdd.ListItems.add(, , item.text)
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
