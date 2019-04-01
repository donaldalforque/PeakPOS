VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form BASE_SalesmanFrm 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   1
      Top             =   5400
      Width           =   4815
   End
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
      Left            =   4080
      TabIndex        =   0
      Top             =   1800
      Width           =   1000
   End
   Begin MSComctlLib.ListView lvSalesman 
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "CategoryId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Salesman"
         Object.Width           =   6253
      EndProperty
   End
   Begin MSComctlLib.Toolbar tb_Standard 
      Height          =   330
      Left            =   0
      TabIndex        =   3
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
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Accounts"
            ImageIndex      =   4
         EndProperty
      EndProperty
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
            Picture         =   "BASE_SalesmanFrm.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BASE_SalesmanFrm.frx":6862
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BASE_SalesmanFrm.frx":D0C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BASE_SalesmanFrm.frx":13926
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "(Double click a record to show more details.)"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   4815
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "BASE_SalesmanFrm.frx":1A188
      Stretch         =   -1  'True
      Top             =   680
      Width           =   480
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Salesman"
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
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Create salesman and tag them with every sales for better performance tracking, bonus and commission references."
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
      Top             =   1200
      Width           =   4815
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   5415
      Left            =   120
      Top             =   480
      Width           =   5055
   End
End
Attribute VB_Name = "BASE_SalesmanFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SalesmanId As Integer
Public Sub Initialize()
    txtName.Text = ""
    SalesmanId = 0
    txtName.SetFocus
End Sub
Public Sub Populate(ByVal data As String)
    Dim Item As MSComctlLib.ListItem
    Select Case data
        Case "Salesman"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Salesman")
            lvSalesman.ListItems.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        Set Item = lvSalesman.ListItems.add(, , "")
                            Item.SubItems(1) = rec!SalesmanId
                            Item.SubItems(2) = rec!salesman
                            Item.Checked = True
                    End If
                    rec.MoveNext
                Loop
            End If
    End Select
End Sub

Private Sub chkShow_Click()
    Dim Item As MSComctlLib.ListItem
    Set rec = New ADODB.Recordset
    Set rec = Global_Data("Salesman")
    lvSalesman.ListItems.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            If chkShow.value = 1 Then
                Set Item = lvSalesman.ListItems.add(, , "")
                    Item.SubItems(1) = rec!SalesmanId
                    Item.SubItems(2) = rec!salesman
                If rec!isActive = "True" Then Item.Checked = True
                lvSalesman.ColumnHeaders(1).width = lvSalesman.width * 0.06
                lvSalesman.ColumnHeaders(3).width = lvSalesman.width * 0.88
            Else
                If rec!isActive = "True" Then
                    Set Item = lvSalesman.ListItems.add(, , "")
                        Item.SubItems(1) = rec!SalesmanId
                        Item.SubItems(2) = rec!salesman
                    If rec!isActive = "True" Then Item.Checked = True
                    lvSalesman.ColumnHeaders(1).width = lvSalesman.width * 0
                    lvSalesman.ColumnHeaders(3).width = lvSalesman.width * 0.94
                End If
            End If
            rec.MoveNext
        Loop
    End If
End Sub

Private Sub Form_Load()
    lvSalesman.ColumnHeaders(1).width = lvSalesman.width * 0
    lvSalesman.ColumnHeaders(3).width = lvSalesman.width * 0.94
    Populate "Salesman"
End Sub


Private Sub lvSalesman_DblClick()
    If lvSalesman.ListItems.count > 0 Then
        BASE_SalesmanDetailsFrm.lblSalesman.Caption = lvSalesman.SelectedItem.SubItems(2)
        BASE_SalesmanDetailsFrm.SalesmanId = lvSalesman.SelectedItem.SubItems(1)
        BASE_SalesmanDetailsFrm.Show (1)
    End If
End Sub

Private Sub lvSalesman_ItemClick(ByVal Item As MSComctlLib.ListItem)
    SalesmanId = Item.SubItems(1)
    txtName.Text = Item.SubItems(2)
    On Error Resume Next
    txtName.SetFocus
End Sub

Private Sub tb_Standard_ButtonClick(ByVal Button As MSComctlLib.Button)
    If EditAccessRights(4) = False Then
        MsgBox ErrorCodes(74), vbCritical, "Limited Rights"
        Exit Sub
    End If
On Error GoTo ErrorHandler:
    Select Case Button.Index
        Case 1 'NEW
            Initialize
        Case 2 'Save
            Dim Item As MSComctlLib.ListItem
            Set con = New ADODB.Connection
            Set rec = New ADODB.Recordset
            con.ConnectionString = ConnString
            con.Open
            
            'Check for Deactivate/Activated Lists
            For Each Item In lvSalesman.ListItems
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "BASE_Salesman_Update"
                cmd.Parameters.Append cmd.CreateParameter("@SalesmanId", adInteger, adParamInputOutput, , Item.SubItems(1))
                cmd.Parameters.Append cmd.CreateParameter("@Salesman", adVarChar, adParamInput, 250, Item.SubItems(2))
                cmd.Parameters.Append cmd.CreateParameter("@isActive", adBoolean, adParamInput, , Item.Checked)
                cmd.Execute
            Next
            
            If Trim(txtName.Text) = "" Then
                Exit Sub
            End If
        
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.Parameters.Append cmd.CreateParameter("@SalesmanId", adInteger, adParamInputOutput, , SalesmanId)
            cmd.Parameters.Append cmd.CreateParameter("@Salesman", adVarChar, adParamInput, 250, txtName.Text)
            cmd.Parameters.Append cmd.CreateParameter("@isActive", adBoolean, adParamInput, , lvSalesman.SelectedItem.Checked)
            
            If SalesmanId = 0 Then
                cmd.CommandText = "BASE_Salesman_Insert"
                cmd.Execute
                SalesmanId = cmd.Parameters("@SalesmanId")
                Set Item = lvSalesman.ListItems.add(, , "")
                    Item.SubItems(1) = SalesmanId
                    Item.SubItems(2) = txtName.Text
                    Item.Checked = True
                    Item.Selected = True
                    Item.EnsureVisible
            Else
                cmd.CommandText = "BASE_Salesman_Update"
                cmd.Execute
                For Each Item In lvSalesman.ListItems
                    If Item.SubItems(1) = SalesmanId Then
                        Item.SubItems(2) = txtName.Text
                        Item.Selected = True
                        Item.EnsureVisible
                    End If
                Next
            End If
            con.Close
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





