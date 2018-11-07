VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form POS_UserValidationFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkShow 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Require Password"
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
      Left            =   5280
      TabIndex        =   2
      Top             =   1560
      Width           =   1845
   End
   Begin VB.TextBox txtname 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   7320
      Width           =   8415
   End
   Begin MSComctlLib.ListView lvUoms 
      Height          =   5055
      Left            =   240
      TabIndex        =   0
      Top             =   1920
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   8916
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Button"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Require Validation"
         Object.Width           =   2540
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
            Object.Visible         =   0   'False
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
            Picture         =   "POS_Settings_UserValidationFrm.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "POS_Settings_UserValidationFrm.frx":6862
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "POS_Settings_UserValidationFrm.frx":D0C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "POS_Settings_UserValidationFrm.frx":13926
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "POS_Settings_UserValidationFrm.frx":1A188
      Top             =   680
      Width           =   480
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "POS User Validation"
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
      Width           =   2295
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "This tool allows you to enable/disable user validation on buttons present in the POS main screen."
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
      Width           =   5655
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   6615
      Left            =   120
      Top             =   480
      Width           =   7215
   End
End
Attribute VB_Name = "POS_UserValidationFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim UomId As Integer
Public Sub Initialize()
    txtname.text = ""
    UomId = 0
    txtname.SetFocus
End Sub
Public Sub Populate(ByVal data As String)
    Dim item As MSComctlLib.ListItem
    Select Case data
        Case "POSDisplay"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("POSValidation")
            lvUoms.ListItems.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    'If rec!isActive = "True" Then
                        Set item = lvUoms.ListItems.add(, , rec!POS_UserValidationId)
                            item.SubItems(1) = rec!Module
                            'If IsNull(rec!Name) = False Then
                            item.SubItems(2) = rec!isRequired
                            'End If
                            'item.Checked = True
                    'End If
                    rec.MoveNext
                Loop
            End If
    End Select
End Sub

Private Sub chkShow_Click()
    If lvUoms.ListItems.Count > 0 Then
        If chkShow.value = 0 Then
            lvUoms.SelectedItem.SubItems(2) = "False"
        Else
            lvUoms.SelectedItem.SubItems(2) = "True"
        End If
    End If
End Sub

Private Sub Form_Load()
    'lvUoms.ColumnHeaders(1).width = lvUoms.width * 0.3167
    lvUoms.ColumnHeaders(2).width = lvUoms.width * 0.45
    lvUoms.ColumnHeaders(3).width = lvUoms.width * 0.45
    'lvUoms.ColumnHeaders(4).width = lvUoms.width * 0
    Populate "POSDisplay"
End Sub


Private Sub lvUoms_DblClick()
'    If lvUoms.ListItems.Count > 0 Then
'        With POS_ProductSelectFrm
'            .txtItemSearch.text = lvUoms.SelectedItem.SubItems(2)
'            .txtItemSearch_Change
'        End With
'        POS_ProductSelectFrm.Show (1)
'    End If
End Sub

Private Sub lvUoms_ItemClick(ByVal item As MSComctlLib.ListItem)
    'SHOW PRODUCT LIST
    'chkShow.Visible = True
    If item.SubItems(2) = "True" Then
        chkShow.value = 1
    Else
        chkShow.value = 0
    End If
End Sub

Private Sub tb_Standard_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo ErrorHandler:
    Select Case Button.Index
        Case 2 'Save
            Dim item As MSComctlLib.ListItem
            Set con = New ADODB.Connection
            Set rec = New ADODB.Recordset
            con.ConnectionString = ConnString
            con.Open
            
            'Check for Deactivate/Activated Lists
            For Each item In lvUoms.ListItems
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "POS_UserValidation_Update"
                cmd.Parameters.Append cmd.CreateParameter("@POS_UserValidationId", adInteger, adParamInput, , item.text)
                cmd.Parameters.Append cmd.CreateParameter("@isRequired", adBoolean, adParamInput, , item.SubItems(2))
                cmd.Execute
            Next
            con.Close
            MsgBox "Records saved.", vbInformation
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
    selectText txtname
End Sub





