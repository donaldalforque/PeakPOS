VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form SO_CustomerListFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer List"
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
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      Height          =   9390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15135
      Begin VB.TextBox txtSearch_Address 
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
         TabIndex        =   4
         Top             =   1560
         Width           =   13575
      End
      Begin VB.TextBox txtSearch_Contact 
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
         TabIndex        =   3
         Top             =   1200
         Width           =   13575
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
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1920
         Width           =   13575
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
         Left            =   1320
         TabIndex        =   1
         Top             =   480
         Width           =   13575
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
         Left            =   1320
         TabIndex        =   2
         Top             =   840
         Width           =   13575
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
      Begin MSComctlLib.ListView lvSearch 
         Height          =   6015
         Left            =   120
         TabIndex        =   7
         Top             =   2880
         Width           =   14805
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CustomerId"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Name"
            Object.Width           =   6526
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Phone"
            Object.Width           =   2893
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Mobile"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Address"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Balance"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Status"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Top             =   1560
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact"
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
         Width           =   705
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
         TabIndex        =   11
         Top             =   1920
         Width           =   570
      End
      Begin VB.Label Label13 
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
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   480
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
         TabIndex        =   9
         Top             =   80
         Width           =   795
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
         TabIndex        =   8
         Top             =   840
         Width           =   555
      End
   End
End
Attribute VB_Name = "SO_CustomerListFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CustomerId As Long
Public Sub Populate(ByVal data As String)
    Set rec = New ADODB.Recordset
    Select Case data
        Case "Status"
            cmbSearch_Status.Clear
            cmbSearch_Status.AddItem ""
            cmbSearch_Status.ItemData(cmbSearch_Status.NewIndex) = -1
            cmbSearch_Status.AddItem "Active"
            cmbSearch_Status.ItemData(cmbSearch_Status.NewIndex) = 1
            cmbSearch_Status.AddItem "Deactivated"
            cmbSearch_Status.ItemData(cmbSearch_Status.NewIndex) = 0
            cmbSearch_Status.ListIndex = 1
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
    cmd.CommandText = "BASE_Customer_Search"
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 500, txtSearch_Name.text)
    If Trim(txtSearch_ItemCode.text) <> "" Then
        cmd.Parameters.Append cmd.CreateParameter("@ItemCode", adVarChar, adParamInput, 500, txtSearch_ItemCode.text)
    Else
        cmd.Parameters.Append cmd.CreateParameter("@ItemCode", adVarChar, adParamInput, 250, Null)
    End If
    If cmbSearch_Status.ListIndex <> 0 Then
        cmd.Parameters.Append cmd.CreateParameter("@isActive", adBoolean, adParamInput, , cmbSearch_Status.ItemData(cmbSearch_Status.ListIndex))
    Else
        cmd.Parameters.Append cmd.CreateParameter("@isActive", adBoolean, adParamInput, , Null)
    End If
    cmd.Parameters.Append cmd.CreateParameter("@Contact", adVarChar, adParamInput, 50, txtSearch_Contact.text)
    cmd.Parameters.Append cmd.CreateParameter("@Address", adVarChar, adParamInput, 250, txtSearch_Address.text)
    Set rec = cmd.Execute
    lvSearch.ListItems.Clear
    Dim item As MSComctlLib.ListItem
    If Not rec.EOF Then
        Do Until rec.EOF
            'If rec!isActive = "True" Then
                Set item = lvSearch.ListItems.add(, , rec!CustomerId)
                    item.SubItems(1) = rec!customerCode
                    item.SubItems(2) = rec!Name
                    item.SubItems(3) = rec!Phone
                    item.SubItems(4) = rec!Mobile
                    item.SubItems(5) = rec!Address
                    item.SubItems(6) = FormatNumber(rec!balance, 2, vbTrue, vbFalse)
                    If rec!isActive = "True" Then
                        item.SubItems(7) = "Active"
                    Else
                        item.SubItems(7) = "Deactivated"
                    End If
            'End If
            rec.MoveNext
        Loop
    End If
    'DistinctList lvSearch
    con.Close
End Sub

Private Sub Form_Load()
    Populate ("Status")
    
    '****** REGION Listview Columns *********
    lvSearch.ColumnHeaders(2).width = lvSearch.width * 0.13
    lvSearch.ColumnHeaders(3).width = lvSearch.width * 0.24
    lvSearch.ColumnHeaders(4).width = lvSearch.width * 0.11
    lvSearch.ColumnHeaders(5).width = lvSearch.width * 0.11
    lvSearch.ColumnHeaders(6).width = lvSearch.width * 0.16
    lvSearch.ColumnHeaders(7).width = lvSearch.width * 0.13
    lvSearch.ColumnHeaders(8).width = lvSearch.width * 0.1
End Sub

Private Sub lvSearch_DblClick()
    If lvSearch.ListItems.Count > 0 Then
        With SO_CustomerFrm
            .txtSearch_Code.text = lvSearch.SelectedItem.SubItems(1)
            .txtSearch_Name.text = lvSearch.SelectedItem.SubItems(2)
            .cmbSearch_Status.ListIndex = cmbSearch_Status.ListIndex
            .btnSearch_Click
            .lvSearch.ListItems(1).Selected = True
            .CustomerId = lvSearch.SelectedItem.text
            .Populate "CustomerGet"
            CornerChildForm SO_CustomerFrm
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


