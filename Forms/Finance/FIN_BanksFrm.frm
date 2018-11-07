VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FIN_BanksFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6045
   ClientLeft      =   7110
   ClientTop       =   2745
   ClientWidth     =   5295
   Icon            =   "FIN_BanksFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkShow 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Left            =   4050
      TabIndex        =   3
      Top             =   1830
      Width           =   1000
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4680
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIN_BanksFrm.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIN_BanksFrm.frx":686E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIN_BanksFrm.frx":D0D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIN_BanksFrm.frx":13932
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
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
      Left            =   240
      MaxLength       =   50
      TabIndex        =   0
      Top             =   5400
      Width           =   4815
   End
   Begin MSComctlLib.ListView lvBanks 
      Height          =   3135
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "BankId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Bank"
         Object.Width           =   6253
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Balance"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.Toolbar tb_Standard 
      Height          =   540
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   953
      ButtonWidth     =   1349
      ButtonHeight    =   953
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Accounts"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Create banks where your savings/checking accounts can be connected."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   450
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   4815
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Banks"
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
      TabIndex        =   1
      Top             =   840
      Width           =   690
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "FIN_BanksFrm.frx":1A194
      Top             =   720
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   5295
      Left            =   120
      Top             =   600
      Width           =   5055
   End
End
Attribute VB_Name = "FIN_BanksFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BankId As Integer
Public Sub Initialize()
    txtName.Text = ""
    BankId = 0
    txtName.SetFocus
End Sub
Public Sub Populate(ByVal data As String)
    Dim item As MSComctlLib.ListItem
    Select Case data
        Case "Bank"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Bank")
            lvBanks.ListItems.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        Set item = lvBanks.ListItems.add(, , "")
                            item.SubItems(1) = rec!BankId
                            item.SubItems(2) = rec!bankname
                            item.SubItems(3) = FormatNumber(rec!balance, 2, vbTrue, vbFalse)
                            item.Checked = True
                    End If
                    rec.MoveNext
                Loop
            End If
    End Select
End Sub

Private Sub chkShow_Click()
    Dim item As MSComctlLib.ListItem
    Set rec = New ADODB.Recordset
    Set rec = Global_Data("Bank")
    lvBanks.ListItems.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            If chkShow.value = 1 Then
                Set item = lvBanks.ListItems.add(, , "")
                    item.SubItems(1) = rec!BankId
                    item.SubItems(2) = rec!bankname
                    item.SubItems(3) = FormatNumber(rec!balance, 2, vbTrue, vbFalse)
                If rec!isActive = "True" Then item.Checked = True
                lvBanks.ColumnHeaders(1).width = lvBanks.width * 0.06
                lvBanks.ColumnHeaders(3).width = lvBanks.width * 0.44
                lvBanks.ColumnHeaders(4).width = lvBanks.width * 0.44
            Else
                If rec!isActive = "True" Then
                    Set item = lvBanks.ListItems.add(, , "")
                        item.SubItems(1) = rec!BankId
                        item.SubItems(2) = rec!bankname
                        item.SubItems(3) = FormatNumber(rec!balance, 2, vbTrue, vbFalse)
                    If rec!isActive = "True" Then item.Checked = True
                    lvBanks.ColumnHeaders(1).width = lvBanks.width * 0
                    lvBanks.ColumnHeaders(3).width = lvBanks.width * 0.47
                    lvBanks.ColumnHeaders(4).width = lvBanks.width * 0.47
                End If
            End If
            rec.MoveNext
        Loop
    End If
End Sub

Private Sub Form_Load()
    lvBanks.ColumnHeaders(1).width = lvBanks.width * 0
    lvBanks.ColumnHeaders(3).width = lvBanks.width * 0.47
    lvBanks.ColumnHeaders(4).width = lvBanks.width * 0.47
    Populate "Bank"
End Sub


Private Sub lvBanks_ItemClick(ByVal item As MSComctlLib.ListItem)
    BankId = item.SubItems(1)
    txtName.Text = item.SubItems(2)
    txtName.SetFocus
End Sub

Private Sub tb_Standard_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo ErrorHandler:
    Select Case Button.Index
        Case 1 'NEW
            Initialize
        Case 2 'Save
            Dim item As MSComctlLib.ListItem
            Set con = New ADODB.Connection
            Set rec = New ADODB.Recordset
            con.ConnectionString = ConnString
            con.Open
            
            'Check for Deactivate/Activated Lists
            For Each item In lvBanks.ListItems
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "BASE_Bank_Update"
                cmd.Parameters.Append cmd.CreateParameter("@BankId", adInteger, adParamInputOutput, , item.SubItems(1))
                cmd.Parameters.Append cmd.CreateParameter("@BankName", adVarChar, adParamInput, 250, item.SubItems(2))
                cmd.Parameters.Append cmd.CreateParameter("@isActive", adBoolean, adParamInput, , item.Checked)
                cmd.Execute
            Next
            
            If Trim(txtName.Text) = "" Then
                Exit Sub
            End If
        
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.Parameters.Append cmd.CreateParameter("@BankId", adInteger, adParamInputOutput, , BankId)
            cmd.Parameters.Append cmd.CreateParameter("@BankName", adVarChar, adParamInput, 250, txtName.Text)
            cmd.Parameters.Append cmd.CreateParameter("@isActive", adBoolean, adParamInput, , lvBanks.SelectedItem.Checked)
            
            If BankId = 0 Then
                cmd.CommandText = "BASE_Bank_Insert"
                cmd.Execute
                BankId = cmd.Parameters("@BankId")
                Set item = lvBanks.ListItems.add(, , "")
                    item.SubItems(1) = BankId
                    item.SubItems(2) = txtName.Text
                    item.Checked = True
                    item.Selected = True
                    item.EnsureVisible
            Else
                cmd.CommandText = "BASE_Bank_Update"
                cmd.Execute
                For Each item In lvBanks.ListItems
                    If item.SubItems(1) = BankId Then
                        item.SubItems(2) = txtName.Text
                        item.Selected = True
                        item.EnsureVisible
                    End If
                Next
            End If
            con.Close
        Case 4 'Accounts
            FIN_BankAccountsFrm.Show (1)
    End Select
    Exit Sub
ErrorHandler:
    If IsNumeric(Err.Description) = True Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(Val(Err.Description))
    Else
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & Err.Description
    End If
    GLOBAL_MessageFrm.Show (1)
End Sub

Private Sub txtName_GotFocus()
    selectText txtName
End Sub
