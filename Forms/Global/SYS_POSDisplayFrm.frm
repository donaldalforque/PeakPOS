VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form POS_DisplayFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   Icon            =   "SYS_POSDisplayFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvUoms 
      Height          =   5295
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   9340
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "POS #"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Product"
         Object.Width           =   6253
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "ProductId"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.TextBox txtname 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   7320
      Width           =   8415
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
      Left            =   7560
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   1000
   End
   Begin MSComctlLib.Toolbar tb_Standard 
      Height          =   330
      Left            =   0
      TabIndex        =   2
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
            Picture         =   "SYS_POSDisplayFrm.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SYS_POSDisplayFrm.frx":686E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SYS_POSDisplayFrm.frx":D0D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SYS_POSDisplayFrm.frx":13932
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Provide a fast sales transaction by displaying your frequently ordered products on the POS Screen layout."
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
      Width           =   8295
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P.O.S Product Display"
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
      TabIndex        =   3
      Top             =   720
      Width           =   2505
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "SYS_POSDisplayFrm.frx":1A194
      Top             =   680
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   6615
      Left            =   120
      Top             =   480
      Width           =   8655
   End
End
Attribute VB_Name = "POS_DisplayFrm"
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
            Set rec = Global_Data("POSDisplay")
            lvUoms.ListItems.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    'If rec!isActive = "True" Then
                        Set item = lvUoms.ListItems.add(, , "")
                            item.SubItems(1) = rec!POS_DisplayId
                            If IsNull(rec!Name) = False Then
                                item.SubItems(2) = rec!Name
                                item.SubItems(3) = rec!ProductId
                            End If
                            'item.Checked = True
                    'End If
                    rec.MoveNext
                Loop
            End If
    End Select
End Sub

Private Sub chkShow_Click()
    Dim item As MSComctlLib.ListItem
    Set rec = New ADODB.Recordset
    Set rec = Global_Data("Uom")
    lvUoms.ListItems.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            If chkShow.value = 1 Then
                Set item = lvUoms.ListItems.add(, , "")
                    item.SubItems(1) = rec!UomId
                    item.SubItems(2) = rec!Uom
                If rec!isActive = "True" Then item.Checked = True
                lvUoms.ColumnHeaders(1).width = lvUoms.width * 0.06
                lvUoms.ColumnHeaders(3).width = lvUoms.width * 0.88
            Else
                If rec!isActive = "True" Then
                    Set item = lvUoms.ListItems.add(, , "")
                        item.SubItems(1) = rec!UomId
                        item.SubItems(2) = rec!Uom
                    If rec!isActive = "True" Then item.Checked = True
                    lvUoms.ColumnHeaders(1).width = lvUoms.width * 0
                    lvUoms.ColumnHeaders(3).width = lvUoms.width * 0.94
                End If
            End If
            rec.MoveNext
        Loop
    End If
End Sub

Private Sub Form_Load()
    lvUoms.ColumnHeaders(2).width = lvUoms.width * 0.1
    lvUoms.ColumnHeaders(3).width = lvUoms.width * 0.85
    Populate "POSDisplay"
End Sub


Private Sub lvUoms_DblClick()
    If lvUoms.ListItems.Count > 0 Then
        With POS_ProductSelectFrm
            .txtItemSearch.text = lvUoms.SelectedItem.SubItems(2)
            .txtItemSearch_Change
        End With
        POS_ProductSelectFrm.Show (1)
    End If
End Sub

Private Sub lvUoms_ItemClick(ByVal item As MSComctlLib.ListItem)
    'SHOW PRODUCT LIST
    
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
                cmd.CommandText = "POS_Display_Update"
                cmd.Parameters.Append cmd.CreateParameter("@POSDisplayId", adInteger, adParamInput, , item.SubItems(1))
                cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , Val(item.SubItems(3)))
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




