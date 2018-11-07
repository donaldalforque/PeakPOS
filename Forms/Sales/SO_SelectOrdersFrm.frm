VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form SO_SelectOrdersFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Orders"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14310
   Icon            =   "SO_SelectOrdersFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   14310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   14055
      Begin VB.TextBox txtOrderNumber 
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
         TabIndex        =   0
         Top             =   480
         Width           =   12375
      End
      Begin VB.TextBox txtCustomer 
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
         TabIndex        =   1
         Top             =   840
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
         TabIndex        =   3
         Top             =   1680
         Width           =   1695
      End
      Begin VB.ComboBox cmbStatus 
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
         ItemData        =   "SO_SelectOrdersFrm.frx":000C
         Left            =   1560
         List            =   "SO_SelectOrdersFrm.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1200
         Width           =   12375
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Order #"
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
         Top             =   480
         Width           =   690
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
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
         Top             =   840
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
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
         TabIndex        =   12
         Top             =   1200
         Width           =   570
      End
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
      TabIndex        =   5
      Top             =   3960
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
      TabIndex        =   6
      Top             =   4560
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
      TabIndex        =   7
      Top             =   5160
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
      Left            =   6840
      TabIndex        =   8
      Top             =   5760
      Width           =   615
   End
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
      TabIndex        =   11
      Top             =   7920
      Width           =   1695
   End
   Begin MSComctlLib.ListView lvSearch 
      Height          =   5175
      Left            =   120
      TabIndex        =   4
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "SalesOrderId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "CustomerId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Customer"
         Object.Width           =   6526
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Order #"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Days Due"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Balance"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Total"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Duedate"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView lvAdd 
      Height          =   5175
      Left            =   7560
      TabIndex        =   9
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
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "SalesOrderId"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "CustomerId"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Customer"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Order #"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Days Due"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Balance"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "DueDate"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "SO_SelectOrdersFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public isPriceMarkUp As Boolean
Dim item As MSComctlLib.ListItem
Public Sub Populate(ByVal data As String)
    Set rec = New ADODB.Recordset
    Select Case data
        Case "Category"

        Case "Status"
            Set rec = Global_Data("Status")
            cmbStatus.AddItem ""
            cmbStatus.ItemData(cmbStatus.NewIndex) = 0
            If Not rec.EOF Then
                Do Until rec.EOF
                    If (rec!StatusId = 4) Or (rec!StatusId = 8) Or (rec!StatusId = 9) Then
                        cmbStatus.AddItem rec!Status
                        cmbStatus.ItemData(cmbStatus.NewIndex) = rec!StatusId
                    End If
                    rec.MoveNext
                Loop
            End If
            cmbStatus.ListIndex = 0
            On Error Resume Next
            cmbStatus.Text = "Invoiced"
    End Select
End Sub

Private Sub RemoveExisting(ByVal a As MSComctlLib.ListView, ByVal b As MSComctlLib.ListView)
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
    'Pass to OrderSet
    Dim item As MSComctlLib.ListItem
    Set OrderSet = Nothing
    Set OrderSet = New ADODB.Recordset
    
    OrderSet.Fields.Append "SalesOrderId", adInteger
    OrderSet.Fields.Append "CustomerId", adInteger
    OrderSet.Fields.Append "Customer", adVarChar, 500
    OrderSet.Fields.Append "OrderNumber", adVarChar, 4000
    OrderSet.Fields.Append "DaysDue", adVarChar, 4000
    OrderSet.Fields.Append "Balance", adVarChar, 4000
    OrderSet.Fields.Append "Total", adVarChar, 4000
    OrderSet.Fields.Append "DueDate", adVarChar, 4000
    OrderSet.Open

    For Each item In lvAdd.ListItems
        OrderSet.AddNew
        OrderSet!SalesOrderId = item.Text
        OrderSet!CustomerId = item.SubItems(1)
        OrderSet!Customer = item.SubItems(2)
        OrderSet!OrderNumber = item.SubItems(3)
        OrderSet!DaysDue = item.SubItems(4)
        OrderSet!balance = item.SubItems(5)
        OrderSet!Total = item.SubItems(6)
        OrderSet!DueDate = item.SubItems(7)
        OrderSet.Update
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
    lvSearch.ListItems.Clear
    
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "SO_SelectOrders"
    cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInput, 50, txtOrderNumber.Text)
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 50, txtCustomer.Text)
    cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , cmbStatus.ItemData(cmbStatus.ListIndex))
    Set rec = cmd.Execute
    If Not rec.EOF Then
        lvSearch.ListItems.Clear
        Do Until rec.EOF
            Set item = lvSearch.ListItems.add(, , rec!SalesOrderId)
                item.SubItems(1) = rec!CustomerId
                item.SubItems(2) = rec!Customer
                item.SubItems(3) = rec!OrderNumber
                item.SubItems(4) = rec!DaysDue
                item.SubItems(5) = FormatNumber(rec!balance, 2, vbTrue, vbFalse)
                item.SubItems(6) = FormatNumber(rec!Total, 2, vbTrue, vbFalse)
                item.SubItems(7) = Format(rec!DueDate, "mm/dd/yy")
            rec.MoveNext
        Loop
    End If
    
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
                nitem.SubItems(7) = item.SubItems(7)
        End If
    Next
    
    lvSearch.ListItems.Clear
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
                    nitem.SubItems(7) = item.SubItems(7)
            End If
        End If
    Next
    
    RemoveExisting lvSearch, lvAdd
End Sub

Private Sub Form_Load()
    lvSearch.ColumnHeaders(1).width = lvSearch.width * 0
    lvSearch.ColumnHeaders(2).width = lvSearch.width * 0
    lvSearch.ColumnHeaders(3).width = lvSearch.width * 0.4
    lvSearch.ColumnHeaders(4).width = lvSearch.width * 0.18
    lvSearch.ColumnHeaders(5).width = lvSearch.width * 0.18
    lvSearch.ColumnHeaders(6).width = lvSearch.width * 0.2
    lvSearch.ColumnHeaders(7).width = lvSearch.width * 0
    
    lvAdd.ColumnHeaders(1).width = lvSearch.width * 0
    lvAdd.ColumnHeaders(2).width = lvSearch.width * 0
    lvAdd.ColumnHeaders(3).width = lvSearch.width * 0.4
    lvAdd.ColumnHeaders(4).width = lvSearch.width * 0.18
    lvAdd.ColumnHeaders(5).width = lvSearch.width * 0.18
    lvAdd.ColumnHeaders(6).width = lvSearch.width * 0.2
    lvAdd.ColumnHeaders(7).width = lvSearch.width * 0

    Populate "Status"
End Sub



Private Sub lvSearch_DblClick()
    If lvSearch.ListItems.Count > 0 Then
        btnSelectOne_Click
    End If
End Sub
