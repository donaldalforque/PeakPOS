VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FIN_BankAccountsFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank Accounts"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10275
   Icon            =   "FIN_BankAccountsFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   10275
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Fre_Main 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7815
      Left            =   3790
      TabIndex        =   9
      Top             =   0
      Width           =   6450
      Begin VB.ComboBox cmbBank 
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
         TabIndex        =   18
         Top             =   940
         Width           =   3015
      End
      Begin VB.CheckBox chkActive 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Height          =   270
         Left            =   240
         TabIndex        =   6
         Top             =   550
         Value           =   1  'Checked
         Width           =   885
      End
      Begin VB.TextBox txtBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   2280
         Width           =   3015
      End
      Begin VB.TextBox txtName 
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
         TabIndex        =   8
         Top             =   1680
         Width           =   3015
      End
      Begin VB.TextBox txtAccountNumber 
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
         TabIndex        =   7
         Top             =   1320
         Width           =   3015
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   6960
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FIN_BankAccountsFrm.frx":6852
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FIN_BankAccountsFrm.frx":D0B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FIN_BankAccountsFrm.frx":13916
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FIN_BankAccountsFrm.frx":1A178
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FIN_BankAccountsFrm.frx":1A3ED
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FIN_BankAccountsFrm.frx":1AA5E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tb_Standard 
         Height          =   330
         Left            =   0
         TabIndex        =   10
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
               Caption         =   "Cancel"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "Print"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank"
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
         TabIndex        =   19
         Top             =   960
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
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
         Top             =   2280
         Width           =   720
      End
      Begin VB.Label Label4 
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
         Top             =   1690
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account #"
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
         Top             =   1320
         Width           =   900
      End
   End
   Begin VB.Frame Fre_Search 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   7815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
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
         Left            =   2400
         TabIndex        =   4
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtSearch_Accounts 
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
         Top             =   900
         Width           =   2295
      End
      Begin VB.CheckBox chkShow 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Show All"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   3
         Top             =   1320
         Width           =   1000
      End
      Begin VB.ComboBox cmbSearch_Banks 
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
         TabIndex        =   1
         Top             =   525
         Width           =   2295
      End
      Begin MSComctlLib.ListView lvSearch 
         Height          =   4815
         Left            =   120
         TabIndex        =   5
         Top             =   2160
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   8493
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
            Text            =   "AccountId"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Account #"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account #"
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
         Top             =   900
         Width           =   900
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
         TabIndex        =   12
         Top             =   120
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank"
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
         Top             =   525
         Width           =   450
      End
   End
End
Attribute VB_Name = "FIN_BankAccountsFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AccountId As Integer
Public Sub Initialize()
    cmbBank.ListIndex = 0
    txtAccountNumber.text = ""
    txtName.text = ""
    txtBalance.text = "0.00"
    chkActive.Value = 1
    
    AccountId = 0
    
    On Error Resume Next
    cmbBank.SetFocus
End Sub
Private Sub Save()
    On Error GoTo ErrorHandler
    If cmbBank.ItemData(cmbBank.ListIndex) = 0 Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(22)
        GLOBAL_MessageFrm.Show (1)
        cmbBank.SetFocus
    ElseIf Trim(txtAccountNumber.text) = "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(21)
        GLOBAL_MessageFrm.Show (1)
        txtAccountNumber.SetFocus
    ElseIf Trim(txtName.text) = "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(18)
        GLOBAL_MessageFrm.Show (1)
        txtName.SetFocus
    Else
        Set con = New ADODB.Connection
        Set cmd = New ADODB.Command
        Set rec = New ADODB.Recordset
        
        con.ConnectionString = ConnString
        con.Open
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInputOutput, , AccountId)
        cmd.Parameters.Append cmd.CreateParameter("@BankId", adInteger, adParamInput, , cmbBank.ItemData(cmbBank.ListIndex))
        cmd.Parameters.Append cmd.CreateParameter("@AccountNumber", adVarChar, adParamInput, 250, txtAccountNumber.text)
        cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, txtName.text)
        cmd.Parameters.Append cmd.CreateParameter("@isActive", adBoolean, adParamInput, , chkActive.Value)
        
        If AccountId = 0 Then
            cmd.CommandText = "BASE_BankAccount_Insert"
            cmd.Execute
            AccountId = cmd.Parameters("@AccountId")
        Else
            cmd.CommandText = "BASE_BankAccount_Update"
            cmd.Execute
        End If
        con.Close
        
        Dim item As MSComctlLib.ListItem
        Dim isFound As Boolean
        For Each item In lvSearch.ListItems
            If item.text = AccountId Then
                isFound = True
                item.SubItems(1) = txtAccountNumber.text
                item.Selected = True
                item.EnsureVisible
                Exit For
            End If
        Next
        If isFound = False Then
            Set item = lvSearch.ListItems.Add(, , AccountId)
                item.SubItems(1) = txtAccountNumber.text
                item.Selected = True
                item.EnsureVisible
        End If
    End If
    Exit Sub
ErrorHandler:
    If IsNumeric(Err.Description) = True Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(Val(Err.Description))
    Else
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & Err.Description
    End If
    GLOBAL_MessageFrm.Show (1)
End Sub
Public Sub Populate(ByVal data As String)
    Dim item As MSComctlLib.ListItem
    Select Case data
        Case "Bank"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Bank")
            cmbSearch_Banks.Clear
            cmbSearch_Banks.AddItem ""
            cmbSearch_Banks.ItemData(cmbSearch_Banks.NewIndex) = 0
            cmbBank.Clear
            cmbBank.AddItem ""
            cmbBank.ItemData(cmbSearch_Banks.NewIndex) = 0
            
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isactive = "True" Then
                        cmbSearch_Banks.AddItem rec!bankname
                        cmbSearch_Banks.ItemData(cmbSearch_Banks.NewIndex) = rec!BankId
                        cmbBank.AddItem rec!bankname
                        cmbBank.ItemData(cmbBank.NewIndex) = rec!BankId
                    End If
                    rec.MoveNext
                Loop
                cmbSearch_Banks.ListIndex = 0
                cmbBank.ListIndex = 0
            End If
        Case "AccountGet"
            Set con = New ADODB.Connection
            Set rec = New ADODB.Recordset
            Set cmd = New ADODB.Command
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "BASE_BankAccount_Get"
            cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , AccountId)
            Set rec = cmd.Execute
            If Not rec.EOF Then
                txtAccountNumber.text = rec!accountnumber
                txtName.text = rec!accountname
                If rec!isactive = "True" Then chkActive.Value = 1
                If rec!isactive = "False" Then chkActive.Value = 0
                txtBalance.text = FormatNumber(rec!balance, 2, vbTrue, vbFalse)
                cmbBank.text = rec!bankname
            End If
            con.Close
    End Select
End Sub

Private Sub btnSearch_Click()
    Set con = New ADODB.Connection
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    Dim item As MSComctlLib.ListItem
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_BankAccount_Search"
    cmd.Parameters.Append cmd.CreateParameter("@BankId", adInteger, adParamInput, , cmbSearch_Banks.ItemData(cmbSearch_Banks.ListIndex))
    cmd.Parameters.Append cmd.CreateParameter("@AccountNumber", adVarChar, adParamInput, 250, txtSearch_Accounts.text)
    Set rec = cmd.Execute
    lvSearch.ListItems.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            If chkShow.Value = 1 Then
                Set item = lvSearch.ListItems.Add(, , rec!AccountId)
                    item.SubItems(1) = rec!accountnumber
            Else
                If rec!isactive = "True" Then
                    Set item = lvSearch.ListItems.Add(, , rec!AccountId)
                        item.SubItems(1) = rec!accountnumber
                End If
            End If
            rec.MoveNext
        Loop
    End If
    con.Close
End Sub

Private Sub Form_Load()
    lvSearch.ColumnHeaders(2).width = lvSearch.width * 0.93
    
    Populate "Bank"
End Sub

Private Sub lvSearch_ItemClick(ByVal item As MSComctlLib.ListItem)
    If lvSearch.ListItems.Count > 0 Then
        AccountId = item.text
        Populate "AccountGet"
    End If
End Sub

Private Sub tb_Standard_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'NEW
            Initialize
        Case 2 'SAVE
            Save
    End Select
End Sub

Private Sub txtSearch_Accounts_Change()
    btnSearch_Click
End Sub
