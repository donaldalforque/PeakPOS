VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form BASE_UserRolesFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "User Roles"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   12780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkView 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "YES"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   9960
      TabIndex        =   12
      Top             =   1800
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.CheckBox chkEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "YES"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11520
      TabIndex        =   11
      Top             =   1800
      Visible         =   0   'False
      Width           =   1000
   End
   Begin VB.TextBox txtRole 
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
      Left            =   4440
      MaxLength       =   50
      TabIndex        =   7
      Top             =   1440
      Width           =   3015
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   12120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BASE_UserRolesFrm.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BASE_UserRolesFrm.frx":6862
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BASE_UserRolesFrm.frx":D0C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BASE_UserRolesFrm.frx":13926
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BASE_UserRolesFrm.frx":1A188
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame LayoutFrame_Search 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.TextBox txtSearch_OrderNumber 
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
         Left            =   840
         TabIndex        =   2
         Top             =   600
         Width           =   2535
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
         Left            =   2160
         TabIndex        =   1
         Top             =   1080
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvSearch 
         Height          =   6615
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   11668
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "RoleId"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Role"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Role"
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
         TabIndex        =   5
         Top             =   600
         Width           =   420
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
         TabIndex        =   4
         Top             =   80
         Width           =   795
      End
   End
   Begin MSComctlLib.Toolbar tb_Standard 
      Height          =   330
      Left            =   3480
      TabIndex        =   6
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   582
      ButtonWidth     =   1349
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Delete"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Copy"
            ImageIndex      =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvModules 
      Height          =   6015
      Left            =   3720
      TabIndex        =   10
      Top             =   2160
      Width           =   8895
      _ExtentX        =   15690
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "UserRightId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Class"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Module"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "View"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Edit"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Roles"
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
      Left            =   4440
      TabIndex        =   9
      Top             =   720
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   3720
      Picture         =   "BASE_UserRolesFrm.frx":209EA
      Top             =   600
      Width           =   480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Role"
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
      Left            =   3720
      TabIndex        =   8
      Top             =   1440
      Width           =   420
   End
End
Attribute VB_Name = "BASE_UserRolesFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sUserRoleId As Integer
Public Sub Populate(ByVal data As String)
    Dim item As MSComctlLib.ListItem
    Select Case data
        Case "UserRoles"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("UserRoles")
'            cmbRoles.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!UserRoleId = 1 Then
                        rec.MoveNext
                    End If
                    Set item = lvSearch.ListItems.add(, , rec!UserRoleId)
                        item.SubItems(1) = rec!role
                    rec.MoveNext
                Loop
            End If
'            cmbRoles.ListIndex = 0
        Case "UserRights"
            Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            Set rec = New ADODB.Recordset
            
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "BASE_UserRights_Get"
            cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , sUserRoleId)
            Set rec = cmd.Execute
            lvModules.ListItems.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    Set item = lvModules.ListItems.add(, , rec!UserRightsId)
                        item.SubItems(1) = rec!Class
                        item.SubItems(2) = rec!Module
                        If rec!allowedit = "True" Then
                            item.SubItems(4) = "Yes"
                        Else
                            item.SubItems(4) = "No"
                        End If
                        If rec!allowview = "True" Then
                            item.SubItems(3) = "Yes"
                        Else
                            item.SubItems(3) = "No"
                        End If
                    rec.MoveNext
                Loop
            End If
            con.Close
        Case "UserRoleRightsInsert"
            'If no records for ACCESS RIGHTS, insert into base_userrolerights
            Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            Set rec = New ADODB.Recordset
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "BASE_UserRoleRights_Insert"
            cmd.Parameters.Append cmd.CreateParameter("@UserRoleId", adInteger, adParamInput, , sUserRoleId)
            cmd.Execute
            con.Close
    End Select
End Sub
Public Sub Initialize()
    txtRole.Text = ""
    lvModules.ListItems.Clear
    
    sUserRoleId = 0
    
    On Error Resume Next
    txtRole.SetFocus
End Sub
Private Sub Form_Load()
    '****** REGION Listview Columns *********
    lvSearch.ColumnHeaders(1).width = lvSearch.width * 0
    lvSearch.ColumnHeaders(2).width = lvSearch.width * 0.9
    
    lvModules.ColumnHeaders(2).width = lvModules.width * 0.2425
    lvModules.ColumnHeaders(3).width = lvModules.width * 0.4425
    lvModules.ColumnHeaders(4).width = lvModules.width * 0.1425
    lvModules.ColumnHeaders(5).width = lvModules.width * 0.1425
    
    Populate "UserRoles"
    Initialize
End Sub

Private Sub lvModules_Click()
    chkEdit.Visible = True
    chkView.Visible = True

    If lvModules.SelectedItem.SubItems(3) = "Yes" Then
        chkView.value = Checked
    Else
        chkView.value = Unchecked
    End If

    If lvModules.SelectedItem.SubItems(4) = "Yes" Then
        chkEdit.value = Checked
    Else
        chkEdit.value = Unchecked
    End If
End Sub

Private Sub lvSearch_Click()
    If lvSearch.ListItems.Count > 0 Then
        sUserRoleId = Val(lvSearch.SelectedItem.Text)
         
        Dim con As New ADODB.Connection
        Set cmd = New ADODB.Command
        Set rec = New ADODB.Recordset
        
        con.ConnectionString = ConnString
        con.Open
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "BASE_UserRoles_Get"
        cmd.Parameters.Append cmd.CreateParameter("@UserRoleId", adInteger, adParamInput, , Val(lvSearch.SelectedItem.Text))
        Set rec = cmd.Execute
        If Not rec.EOF Then
            txtRole.Text = rec!role
        End If
        con.Close
        
        'If no records for ACCESS RIGHTS, insert into base_userrolerights
        Populate "UserRoleRightsInsert"
        Populate "UserRights"
    End If
End Sub
Private Sub Save()
On Error GoTo ErrHandler:
    If Trim(txtRole.Text) = "" Then
        MsgBox "Role name is required.", vbCritical
        txtRole.SetFocus
        Exit Sub
    End If
    Dim item As MSComctlLib.ListItem
    Dim con As New ADODB.Connection
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    
    'Save Role
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("@RoleId", adInteger, adParamInputOutput, , sUserRoleId)
    cmd.Parameters.Append cmd.CreateParameter("@Role", adVarChar, adParamInput, 250, txtRole.Text)
    cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
    cmd.Parameters.Append cmd.CreateParameter("@WorkStationId", adInteger, adParamInput, , WorkstationId)
    
    If sUserRoleId = 0 Then
        cmd.CommandText = "BASE_UserRole_Insert"
        cmd.Execute
        sUserRoleId = cmd.Parameters("@RoleId")
        
        'Get Access Rights
        Populate "UserRoleRightsInsert"
        Populate "UserRights"
        
        'Add to list
        Set item = lvSearch.ListItems.add(, , sUserRoleId)
            item.SubItems(1) = txtRole.Text
    Else
        cmd.CommandText = "BASE_UserRole_Update"
        cmd.Execute
        
        'Update List
        For Each item In lvSearch.ListItems
            If item.Text = sUserRoleId Then
                item.SubItems(1) = txtRole.Text
                item.EnsureVisible
                Exit For
            End If
        Next
    End If
    
    'Save Rights
    For Each item In lvModules.ListItems
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "BASE_UserRights_Update"
        cmd.Parameters.Append cmd.CreateParameter("@UserRightsId", adInteger, adParamInput, , item.Text)
        If item.SubItems(3) = "No" Then
            cmd.Parameters.Append cmd.CreateParameter("@AllowView", adBoolean, adParamInput, , False)
        Else
            cmd.Parameters.Append cmd.CreateParameter("@AllowView", adBoolean, adParamInput, , True)
        End If
        If item.SubItems(4) = "No" Then
            cmd.Parameters.Append cmd.CreateParameter("@AllowEdit", adBoolean, adParamInput, , False)
        Else
            cmd.Parameters.Append cmd.CreateParameter("@AllowEdit", adBoolean, adParamInput, , True)
        End If
        cmd.Execute
    Next
    con.Close
    
    'update access rights
    GetAccessRights gUserRoleId
    Exit Sub
ErrHandler:
    'con.RollbackTrans
    If IsNumeric(Err.Description) = True Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(Err.Description)
        BASE_ContainerFrm.statusBar_Main.Panels(1).Text = ErrorCodes(0) & " " & ErrorCodes(Err.Description)
    Else
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & Err.Description
        BASE_ContainerFrm.statusBar_Main.Panels(1).Text = ErrorCodes(0) & " " & Err.Description
    End If
        GLOBAL_MessageFrm.Show (1)
End Sub
Private Sub chkEdit_Click()
    If chkEdit.value = Checked Then
        lvModules.SelectedItem.SubItems(4) = "Yes"
    Else
        lvModules.SelectedItem.SubItems(4) = "No"
    End If
End Sub

Private Sub chkView_Click()
    If chkView.value = Checked Then
        lvModules.SelectedItem.SubItems(3) = "Yes"
    Else
        lvModules.SelectedItem.SubItems(3) = "No"
    End If
End Sub

Private Sub tb_Standard_ButtonClick(ByVal Button As MSComctlLib.Button)
    If EditAccessRights(27) = False Then
        MsgBox ErrorCodes(74), vbCritical, "Limited Rights"
        Exit Sub
    End If
    Select Case Button.Index
        Case 1: 'new
            Initialize
        Case 2: 'save
            Save
    End Select
End Sub
