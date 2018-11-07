VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form BASE_UserRightsFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8685
   Icon            =   "BASE_UserRightsFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   8685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   7200
      TabIndex        =   11
      Top             =   8040
      Width           =   1335
   End
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
      Left            =   5760
      TabIndex        =   10
      Top             =   8040
      Width           =   1335
   End
   Begin VB.Frame Frame_Main 
      BackColor       =   &H00FFFFFF&
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      Begin MSComctlLib.ListView lvModules 
         Height          =   4575
         Left            =   240
         TabIndex        =   9
         Top             =   2880
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   8070
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
      Begin VB.TextBox txtUserNumber 
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
         Left            =   6120
         MaxLength       =   4
         TabIndex        =   4
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtPin 
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
         IMEMode         =   3  'DISABLE
         Left            =   6120
         MaxLength       =   4
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   720
         Width           =   2055
      End
      Begin VB.ComboBox cmbRoles 
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
         ItemData        =   "BASE_UserRightsFrm.frx":000C
         Left            =   1050
         List            =   "BASE_UserRightsFrm.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1800
         Width           =   3015
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
         Left            =   6840
         TabIndex        =   19
         Top             =   2880
         Visible         =   0   'False
         Width           =   1000
      End
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
         Left            =   5760
         TabIndex        =   18
         Top             =   2880
         Visible         =   0   'False
         Width           =   1000
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
         Left            =   1050
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1440
         Width           =   3015
      End
      Begin VB.TextBox txtConfirm 
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
         IMEMode         =   3  'DISABLE
         Left            =   6120
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox txtPassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   6120
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox txtUsername 
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
         Left            =   6120
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1440
         Width           =   2055
      End
      Begin VB.CheckBox chkActive 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Active"
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
         Left            =   240
         TabIndex        =   1
         Top             =   1080
         Width           =   1000
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Number"
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
         Left            =   4320
         TabIndex        =   22
         Top             =   360
         Width           =   1230
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pin"
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
         Left            =   4320
         TabIndex        =   21
         Top             =   720
         Width           =   300
      End
      Begin VB.Label Label4 
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
         TabIndex        =   20
         Top             =   1800
         Width           =   420
      End
      Begin VB.Label Label5 
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
         TabIndex        =   17
         Top             =   1440
         Width           =   555
      End
      Begin VB.Label Label3 
         Height          =   15
         Left            =   240
         TabIndex        =   16
         Top             =   2760
         Width           =   7935
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password"
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
         Left            =   4320
         TabIndex        =   15
         Top             =   2160
         Width           =   1680
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Password"
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
         Left            =   4320
         TabIndex        =   14
         Top             =   1800
         Width           =   1350
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
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
         Left            =   4320
         TabIndex        =   13
         Top             =   1440
         Width           =   1020
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "BASE_UserRightsFrm.frx":0010
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User Details"
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
         TabIndex        =   12
         Top             =   405
         Width           =   1380
      End
   End
End
Attribute VB_Name = "BASE_UserRightsFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public CheckUserId As Integer
Public cUserRoleId As Integer
Public Sub Populate(ByVal data As String)
    Dim item As MSComctlLib.ListItem
    Select Case data
        Case "User"
            Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            Set rec = New ADODB.Recordset
            
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "BASE_User_Get"
            cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , CheckUserId)
            Set rec = cmd.Execute
            If Not rec.EOF Then
                txtName.text = rec!Name
                txtUsername.text = rec!UserName
                txtPassword.text = rec!Password
                txtConfirm.text = rec!Password
                On Error Resume Next
                txtPin.text = rec!Pin
                If rec!isActive = "True" Then
                    chkActive.value = 1
                Else
                    chkActive.value = 0
                End If
            End If
        Case "UserRoles"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("UserRoles")
            cmbRoles.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    'If rec!UserRoleId = 1 And UserRoleId <> 1 Then rec.MoveNext
                    If CheckUserId <> 1 Then
                        If rec!UserRoleId = 1 Then
                            rec.MoveNext
                        Else
                            cmbRoles.AddItem rec!role
                            cmbRoles.ItemData(cmbRoles.NewIndex) = rec!UserRoleId
                            rec.MoveNext
                        End If
                    Else
                        If rec!UserRoleId = 1 Then
                            cmbRoles.AddItem rec!role
                            cmbRoles.ItemData(cmbRoles.NewIndex) = rec!UserRoleId
                            rec.MoveNext
                        Else
                            rec.MoveNext
                        End If
                    End If
                    
                Loop
            End If
            cmbRoles.ListIndex = 0
        Case "UserRights"
            Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            Set rec = New ADODB.Recordset
            
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "BASE_UserRights_Get"
            cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , cUserRoleId)
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
    End Select
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnSave_Click()
On Error GoTo ErrHandler
    If Trim(txtName.text) = "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(18)
        GLOBAL_MessageFrm.Show (1)
        txtName.SetFocus
    ElseIf Trim(txtUserNumber.text) = "" Or IsNumeric(txtUserNumber.text) = False Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(44)
        GLOBAL_MessageFrm.Show (1)
        txtUserNumber.SetFocus
    ElseIf IsNumeric(txtPin.text) = False Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(45)
        GLOBAL_MessageFrm.Show (1)
        txtPin.SetFocus
    ElseIf Trim(txtUsername.text) = "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(30)
        GLOBAL_MessageFrm.Show (1)
        txtUsername.SetFocus
    ElseIf txtPassword.text <> txtConfirm.text Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(28)
        GLOBAL_MessageFrm.Show (1)
        txtPassword.SetFocus
    Else
        Set con = New ADODB.Connection
                
        con.ConnectionString = ConnString
        con.Open
                
        Dim item As MSComctlLib.ListItem
        For Each item In lvModules.ListItems
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "BASE_UserRight_Update"
            
            cmd.Parameters.Append cmd.CreateParameter("@UserRightsId", adInteger, adParamInput, , item.text)
            If item.SubItems(3) = "Yes" Then
                cmd.Parameters.Append cmd.CreateParameter("@AllowView", adBoolean, adParamInput, , "True")
            Else
                cmd.Parameters.Append cmd.CreateParameter("@AllowView", adBoolean, adParamInput, , "False")
            End If
            If item.SubItems(4) = "Yes" Then
                cmd.Parameters.Append cmd.CreateParameter("@AllowEdit", adBoolean, adParamInput, , "True")
            Else
                cmd.Parameters.Append cmd.CreateParameter("@AllowEdit", adBoolean, adParamInput, , "False")
            End If
            cmd.Execute
        Next
        
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "BASE_User_Update"
        
        cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , CheckUserId)
        cmd.Parameters.Append cmd.CreateParameter("@UserRoleId", adInteger, adParamInput, , cmbRoles.ItemData(cmbRoles.ListIndex))
        cmd.Parameters.Append cmd.CreateParameter("@UserNumber", adInteger, adParamInput, , Val(txtUserNumber.text))
        cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 50, txtName.text)
        cmd.Parameters.Append cmd.CreateParameter("@Username", adVarChar, adParamInput, 50, txtUsername.text)
        cmd.Parameters.Append cmd.CreateParameter("@Password", adVarChar, adParamInput, 50, txtPassword.text)
        cmd.Parameters.Append cmd.CreateParameter("@Pin", adVarChar, adParamInput, 4, txtPin.text)
        cmd.Parameters.Append cmd.CreateParameter("@isActive", adBoolean, adParamInput, , chkActive.value)
        cmd.Execute
        con.Close
        BASE_GeneralSettingsFrm.Populate "User"
        Unload Me
    End If
    Exit Sub
ErrHandler:
    If IsNumeric(Err.Description) = True Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(Val(Err.Description))
    Else
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & Err.Description
    End If
    GLOBAL_MessageFrm.Show (1)
End Sub

Private Sub chkActive_Click()
    'If UserId = 1 Then chkActive.value = 1
    If CheckUserId = 1 Then chkActive.value = 1
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

Private Sub Form_Load()
    lvModules.ColumnHeaders(2).width = lvModules.width * 0.2425
    lvModules.ColumnHeaders(3).width = lvModules.width * 0.4425
    lvModules.ColumnHeaders(4).width = lvModules.width * 0.1425
    lvModules.ColumnHeaders(5).width = lvModules.width * 0.1425
    
    Populate "UserRights"
    Populate "User"
    Populate "UserRoles"
    
    If CheckUserId = 1 Then
        
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UserId = 0
End Sub

Private Sub lvModules_ItemClick(ByVal item As MSComctlLib.ListItem)
'    chkEdit.Visible = True
'    chkView.Visible = True
'
'    If item.SubItems(3) = "Yes" Then
'        chkView.value = Checked
'    Else
'        chkView.value = Unchecked
'    End If
'
'    If item.SubItems(4) = "Yes" Then
'        chkEdit.value = Checked
'    Else
'        chkEdit.value = Unchecked
'    End If
End Sub

Private Sub txtConfirm_GotFocus()
    selectText txtConfirm
End Sub

Private Sub txtName_GotFocus()
    selectText txtName
End Sub

Private Sub txtPassword_GotFocus()
    selectText txtPassword
End Sub

Private Sub txtPin_GotFocus()
    selectText txtPin
End Sub

Private Sub txtUsername_GotFocus()
    selectText txtUsername
End Sub

Private Sub txtUserNumber_GotFocus()
    selectText txtUserNumber
End Sub
