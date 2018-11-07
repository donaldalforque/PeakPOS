VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form INV_ProductExtraInfoFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   11895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   7335
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   11655
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
         Left            =   10320
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtReference 
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
         Left            =   2400
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DateFrom 
         Height          =   345
         Left            =   5640
         TabIndex        =   4
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   87031809
         CurrentDate     =   41686
      End
      Begin MSComCtl2.DTPicker DateTo 
         Height          =   345
         Left            =   8160
         TabIndex        =   7
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   87031809
         CurrentDate     =   41686
      End
      Begin MSComctlLib.ListView lvExpiry 
         Height          =   6495
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   11456
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ProductExpiryId"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ProductId"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Stocked-In Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Reference #"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Lot #"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Expiry Date"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reference/Lot #"
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
         Left            =   840
         TabIndex        =   8
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
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
         Left            =   7800
         TabIndex        =   6
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Expiry From"
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
         Left            =   4440
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComctlLib.Toolbar tb_Standard 
      Height          =   330
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   582
      ButtonWidth     =   1588
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Add"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Save"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4680
      Top             =   0
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
            Picture         =   "INV_ProductExtraInfoFrm.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_ProductExtraInfoFrm.frx":6862
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_ProductExtraInfoFrm.frx":D0C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_ProductExtraInfoFrm.frx":13926
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product Extra Info"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
End
Attribute VB_Name = "INV_ProductExtraInfoFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSearch_Click()
    Dim item As MSComctlLib.ListItem
    Dim con As New ADODB.Connection
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "INV_ProductExpiry_Search"
    cmd.Parameters.Append cmd.CreateParameter("@Reference", adVarChar, adParamInput, 50, txtReference.text)
    cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , INV_NewProductFrm.ProductId)
    cmd.Parameters.Append cmd.CreateParameter("@DateFrom", adDate, adParamInput, , DateFrom.value)
    cmd.Parameters.Append cmd.CreateParameter("@DateTo", adDate, adParamInput, , DateTo.value)
    Set rec = cmd.Execute
    lvExpiry.ListItems.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            Set item = lvExpiry.ListItems.add(, , rec!ProductExpiryId)
                item.SubItems(1) = rec!ProductId
                item.SubItems(2) = rec!StockedinDate
                item.SubItems(3) = rec!Reference
                item.SubItems(4) = rec!LotNumber
                item.SubItems(5) = Format(rec!expirydate, "MM/DD/YY")
            rec.MoveNext
        Loop
    End If
    con.Close
End Sub

Private Sub Form_Load()
    lvExpiry.ColumnHeaders(3).width = lvExpiry.width * 0.24
    lvExpiry.ColumnHeaders(4).width = lvExpiry.width * 0.24
    lvExpiry.ColumnHeaders(5).width = lvExpiry.width * 0.24
    lvExpiry.ColumnHeaders(6).width = lvExpiry.width * 0.24
    
    DateFrom.value = Format(Now - 30, "MM/DD/YY")
    DateTo.value = Format(Now, "MM/DD/YY")
    
    LoadExpiryData
End Sub

Private Sub lvExpiry_DblClick()
    If lvExpiry.ListItems.Count > 0 Then
        INV_ProductExpiryFrm.ProductExpiryId = lvExpiry.SelectedItem.text
        INV_ProductExpiryFrm.Show (1)
    End If
End Sub

Private Sub tb_Standard_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 ' new
            INV_ProductExpiryFrm.ProductExpiryId = 0
            INV_ProductExpiryFrm.Show (1)
        Case 4 ' delete
            If lvExpiry.ListItems.Count > 0 Then
                x = MsgBox("Are you sure you want to delete the selected record?", vbQuestion + vbYesNo)
                If x = vbYes Then
                    Dim con As New ADODB.Connection
                    con.ConnectionString = ConnString
                    con.Open
                    Set cmd = New ADODB.Command
                    cmd.ActiveConnection = con
                    cmd.CommandType = adCmdStoredProc
                    cmd.CommandText = "INV_ProductExpiry_Delete"
                    cmd.Parameters.Append cmd.CreateParameter("@ProductExpiryId", adInteger, adParamInput, , lvExpiry.SelectedItem.text)
                    cmd.Execute
                    con.Close
                    
                    lvExpiry.ListItems.Remove (lvExpiry.SelectedItem.Index)
                    MsgBox "Record deleted.", vbInformation
                End If
            End If
    End Select
End Sub

Public Sub LoadExpiryData()
    Dim item As MSComctlLib.ListItem
    Dim con As New ADODB.Connection
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "INV_ProductExpiry_Get"
    cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , INV_NewProductFrm.ProductId)
    Set rec = cmd.Execute
    If Not rec.EOF Then
        Do Until rec.EOF
            Set item = lvExpiry.ListItems.add(, , rec!ProductExpiryId)
                item.SubItems(1) = INV_NewProductFrm.ProductId
                item.SubItems(2) = rec!StockedinDate
                item.SubItems(3) = rec!Reference
                item.SubItems(4) = rec!LotNumber
                item.SubItems(5) = Format(rec!expirydate, "MM/DD/YY")
            rec.MoveNext
        Loop
    End If
    con.Close
End Sub

Private Sub txtReference_Change()
    btnSearch_Click
End Sub
