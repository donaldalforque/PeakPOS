VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form INV_ProductListFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Product List"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   15090
   Begin VB.Frame LayoutFrame_Search 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   9390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15135
      Begin VB.ComboBox cmbSearch_Type 
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
         Top             =   1580
         Width           =   13455
      End
      Begin VB.CommandButton btnSearch 
         Caption         =   "Refresh"
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
         Left            =   13680
         TabIndex        =   6
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtSearch_Name 
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
         TabIndex        =   2
         Top             =   840
         Width           =   13455
      End
      Begin VB.ComboBox cmbSearch_Category 
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
         TabIndex        =   3
         Top             =   1200
         Width           =   13455
      End
      Begin VB.TextBox txtSearch_ItemCode 
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
         Top             =   480
         Width           =   13455
      End
      Begin VB.ComboBox cmbSearch_Status 
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
         TabIndex        =   5
         Top             =   1950
         Width           =   13455
      End
      Begin MSComctlLib.ListView lvSearch 
         Height          =   6015
         Left            =   120
         TabIndex        =   7
         Top             =   2880
         Width           =   14800
         _ExtentX        =   26114
         _ExtentY        =   10610
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
         NumItems        =   6
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
            Text            =   "Category"
            Object.Width           =   2893
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
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
         Top             =   1575
         Width           =   450
      End
      Begin VB.Label Label2 
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   1200
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   10
         Top             =   80
         Width           =   795
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Item code"
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
         Width           =   930
      End
      Begin VB.Label Label16 
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
         TabIndex        =   8
         Top             =   1950
         Width           =   570
      End
   End
End
Attribute VB_Name = "INV_ProductListFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ProductId As Long
Public Sub Populate(ByVal data As String)
    Set rec = New ADODB.Recordset
    Select Case data
        Case "Category"
            Set rec = Global_Data("Category")
            cmbSearch_Category.Clear
            cmbSearch_Category.AddItem ""
            cmbSearch_Category.ItemData(cmbSearch_Category.NewIndex) = 0
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        cmbSearch_Category.AddItem rec!Category
                        cmbSearch_Category.ItemData(cmbSearch_Category.NewIndex) = rec!CategoryId
                    End If
                    rec.MoveNext
                Loop
            End If
            cmbSearch_Category.ListIndex = 0
        Case "Status"
            cmbSearch_Status.Clear
            cmbSearch_Status.AddItem ""
            cmbSearch_Status.ItemData(cmbSearch_Status.NewIndex) = -1
            cmbSearch_Status.AddItem "Active"
            cmbSearch_Status.ItemData(cmbSearch_Status.NewIndex) = 1
            cmbSearch_Status.AddItem "Deactivated"
            cmbSearch_Status.ItemData(cmbSearch_Status.NewIndex) = 0
            cmbSearch_Status.ListIndex = 1
        Case "Type"
            Set rec = Global_Data("Type")
            cmbSearch_Type.AddItem ""
            cmbSearch_Type.ItemData(cmbSearch_Type.NewIndex) = 0
            If Not rec.EOF Then
                Do Until rec.EOF
                    cmbSearch_Type.AddItem rec!Type
                    cmbSearch_Type.ItemData(cmbSearch_Type.NewIndex) = rec!TypeId
                    rec.MoveNext
                Loop
            End If
            cmbSearch_Type.ListIndex = 0
    End Select
End Sub
Private Sub btnSearch_Click()
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_Product_Search1"
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 500, txtSearch_Name.text)
    If Trim(txtSearch_ItemCode.text) <> "" Then
        cmd.Parameters.Append cmd.CreateParameter("@ItemCode", adVarChar, adParamInput, 250, txtSearch_ItemCode.text)
    Else
        cmd.Parameters.Append cmd.CreateParameter("@ItemCode", adVarChar, adParamInput, 250, Null)
    End If
    If cmbSearch_Category.ListIndex <> 0 Then
        cmd.Parameters.Append cmd.CreateParameter("@CategoryId", adInteger, adParamInput, , cmbSearch_Category.ItemData(cmbSearch_Category.ListIndex))
    Else
        cmd.Parameters.Append cmd.CreateParameter("@CategoryId", adInteger, adParamInput, , Null)
    End If
    cmd.Parameters.Append cmd.CreateParameter("@LocationId", adInteger, adParamInput, , Null)
    If cmbSearch_Type.ListIndex <> 0 Then
        cmd.Parameters.Append cmd.CreateParameter("@TypeId", adInteger, adParamInput, , cmbSearch_Type.ItemData(cmbSearch_Type.ListIndex))
    Else
        cmd.Parameters.Append cmd.CreateParameter("@TypeId", adInteger, adParamInput, , Null)
    End If
    If cmbSearch_Status.ListIndex <> 0 Then
        cmd.Parameters.Append cmd.CreateParameter("@isActive", adBoolean, adParamInput, , cmbSearch_Status.ItemData(cmbSearch_Status.ListIndex))
    End If

    Set rec = cmd.Execute
    lvSearch.ListItems.Clear
    Dim item As MSComctlLib.ListItem
    If Not rec.EOF Then
        Do Until rec.EOF
            'If rec!isActive = "True" Then
                Set item = lvSearch.ListItems.add(, , rec!ProductId)
                    item.SubItems(1) = rec!itemcode
                    item.SubItems(2) = rec!Name
                    item.SubItems(3) = rec!Category
                    item.SubItems(4) = rec!Type
                    If rec!isActive = "True" Then
                        item.SubItems(5) = "Active"
                    Else
                        item.SubItems(5) = "Deactivated"
                    End If
            'End If
            rec.MoveNext
        Loop
    End If
    'DistinctList lvSearch
    con.Close
    BASE_ContainerFrm.statusBar_Main.Panels(1).text = "Total item(s) found: " & lvSearch.ListItems.Count
End Sub
Private Sub Form_Load()
    Populate ("Category")
    Populate ("Type")
    Populate ("Status")
    
    '****** REGION Listview Columns *********
    lvSearch.ColumnHeaders(2).width = lvSearch.width * 0.13
    lvSearch.ColumnHeaders(3).width = lvSearch.width * 0.41
    lvSearch.ColumnHeaders(4).width = lvSearch.width * 0.16
    lvSearch.ColumnHeaders(5).width = lvSearch.width * 0.14
    lvSearch.ColumnHeaders(6).width = lvSearch.width * 0.14
   
End Sub

Private Sub lvSearch_DblClick()
    If lvSearch.ListItems.Count > 0 Then
        'Access Rights
        If ViewAccessRights(1) = False Then
            MsgBox ErrorCodes(75), vbCritical, "Limited Rights"
            Exit Sub
        End If
        With INV_NewProductFrm
            .txtSearch_ItemCode.text = lvSearch.SelectedItem.SubItems(1)
            .txtSearch_Name.text = lvSearch.SelectedItem.SubItems(2)
            .cmbSearch_Status.ListIndex = cmbSearch_Status.ListIndex
            .cmbSearch_Category.ListIndex = cmbSearch_Category.ListIndex
            .btnSearch_Click
            .lvSearch.ListItems(1).Selected = True
            .ProductId = lvSearch.SelectedItem.text
            .Populate "ProductSelect"
            .Populate "InventoryLoad"
            .CountQuantity
            CornerChildForm INV_NewProductFrm
            .Show
            .ZOrder 0
        End With
    End If
End Sub

Private Sub txtSearch_ItemCode_Change()
    btnSearch_Click
End Sub

Private Sub txtSearch_Name_Change()
    btnSearch_Click
End Sub

