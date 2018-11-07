VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FIN_CheckRegistryFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check Registry"
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
   Begin VB.TextBox txtDateCleared 
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
      Left            =   9600
      TabIndex        =   30
      Top             =   8520
      Width           =   1815
   End
   Begin VB.TextBox dtCleared 
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
      Left            =   4200
      TabIndex        =   29
      Top             =   7560
      Width           =   1695
   End
   Begin VB.OptionButton optAddReceivable 
      BackColor       =   &H00FFFFFF&
      Caption         =   "RECEIVABLE"
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
      Left            =   13440
      TabIndex        =   27
      Top             =   8160
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.OptionButton optAddPayable 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PAYABLE"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   13440
      TabIndex        =   26
      Top             =   8640
      Width           =   1455
   End
   Begin VB.TextBox txtRemarks 
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
      Left            =   11520
      TabIndex        =   24
      Top             =   8520
      Width           =   1815
   End
   Begin VB.CommandButton btnUpdateStatus 
      Caption         =   "Update"
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
      Left            =   6000
      TabIndex        =   22
      Top             =   7560
      Width           =   975
   End
   Begin VB.ComboBox cmbUpdateStatus 
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
      Left            =   840
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   7560
      Width           =   1935
   End
   Begin VB.ComboBox cmbStatus 
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
      Left            =   7200
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   8520
      Width           =   2295
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
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
      Left            =   5040
      TabIndex        =   16
      Top             =   8520
      Width           =   2055
   End
   Begin VB.ComboBox cmbCheckStatus 
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
      Left            =   10200
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1800
      Width           =   2295
   End
   Begin VB.CommandButton btnSearch 
      Caption         =   "SEARCH"
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
      Left            =   12960
      TabIndex        =   13
      Top             =   1800
      Width           =   1935
   End
   Begin VB.OptionButton optPayable 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Payables"
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
      Left            =   8040
      TabIndex        =   12
      Top             =   1800
      Value           =   -1  'True
      Width           =   2175
   End
   Begin VB.OptionButton optReceivable 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Receivables"
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
      Left            =   6000
      TabIndex        =   11
      Top             =   1800
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   330
      Left            =   1320
      TabIndex        =   7
      Top             =   1800
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   96141313
      CurrentDate     =   41779
   End
   Begin VB.TextBox txtCheckNumber 
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
      Left            =   120
      TabIndex        =   0
      Top             =   8520
      Width           =   2655
   End
   Begin MSComctlLib.Toolbar tb_Standard 
      Height          =   330
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   582
      ButtonWidth     =   1535
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
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Print"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   14280
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
            Picture         =   "FIN_CheckRegistryFrm.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIN_CheckRegistryFrm.frx":6862
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIN_CheckRegistryFrm.frx":D0C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIN_CheckRegistryFrm.frx":13926
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIN_CheckRegistryFrm.frx":13B9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FIN_CheckRegistryFrm.frx":1420C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvCheckRegistry 
      Height          =   4935
      Left            =   120
      TabIndex        =   5
      Top             =   2400
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   8705
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
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "CheckRegistryId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Check #"
         Object.Width           =   11783
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Customer"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Ref #"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Due Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Days Remaining"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Remarks"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Date Cleared"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   330
      Left            =   3720
      TabIndex        =   9
      Top             =   1800
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   96141313
      CurrentDate     =   41779
   End
   Begin MSComCtl2.DTPicker dtCheckDate 
      Height          =   330
      Left            =   3000
      TabIndex        =   15
      Top             =   8520
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   96141313
      CurrentDate     =   41779
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Cleared"
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
      Left            =   9600
      TabIndex        =   31
      Top             =   8160
      Width           =   1200
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date Cleared"
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
      Left            =   2880
      TabIndex        =   28
      Top             =   7560
      Width           =   1200
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks"
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
      Left            =   11520
      TabIndex        =   25
      Top             =   8160
      Width           =   795
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   7020
      TabIndex        =   23
      Top             =   7440
      Width           =   4365
   End
   Begin VB.Label Label4 
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
      Left            =   120
      TabIndex        =   21
      Top             =   7560
      Width           =   570
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   14880
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Label Label3 
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
      Left            =   7200
      TabIndex        =   19
      Top             =   8160
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
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
      Left            =   5040
      TabIndex        =   17
      Top             =   8160
      Width           =   750
   End
   Begin VB.Label Label10 
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
      Left            =   3360
      TabIndex        =   10
      Top             =   1800
      Width           =   225
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date From"
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
      Top             =   1800
      Width           =   960
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SEARCH"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   915
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "FIN_CheckRegistryFrm.frx":1AA6E
      Top             =   600
      Width           =   480
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   $"FIN_CheckRegistryFrm.frx":1AF5E
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   840
      TabIndex        =   3
      Top             =   600
      Width           =   6615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Check Number"
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
      Left            =   120
      TabIndex        =   2
      Top             =   8160
      Width           =   1365
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Check Date"
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
      Left            =   3000
      TabIndex        =   1
      Top             =   8160
      Width           =   1035
   End
End
Attribute VB_Name = "FIN_CheckRegistryFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CheckRegistryId As Integer
Private Sub Initialize()
    'lblTotal.Caption = "TOTAL: 0.00"
    cmbStatus.ListIndex = 2
    cmbCheckStatus.ListIndex = 0
    cmbUpdateStatus.ListIndex = 2
    
    txtCheckNumber.text = ""
    txtAmount.text = ""
    txtRemarks.text = ""
    dtCheckDate.value = Format(Now, "MM/DD/YY")
    
    CheckRegistryId = 0
    
    
End Sub
'Public Function Validate() As Boolean
'    If Trim(txtCheckNumber.text) = "" Then
'        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(32)
'        GLOBAL_MessageFrm.Show (1)
'        txtCheckNumber.SetFocus
'        Validate = False
'    End If
'End Function
Public Sub CountTotal()
    Dim item As MSComctlLib.ListItem
    Dim total As Double
    For Each item In lvCheckRegistry.ListItems
        total = total + Val(Replace(item.SubItems(8), ",", ""))
    Next
    lblTotal.Caption = "TOTAL: " & FormatNumber(total, 2, vbTrue, vbFalse)
End Sub

Private Sub btnSearch_Click()
    Set con = New ADODB.Connection
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "FIN_CheckRegistry_Search"
    cmd.Parameters.Append cmd.CreateParameter("@DateFrom", adDate, adParamInput, , dtFrom.value)
    cmd.Parameters.Append cmd.CreateParameter("@DateTo", adDate, adParamInput, , dtTo.value)
    If optReceivable.value = True Then
        cmd.Parameters.Append cmd.CreateParameter("@isReceivable", adBoolean, adParamInput, , "True")
    ElseIf optPayable.value = True Then
        cmd.Parameters.Append cmd.CreateParameter("@isReceivable", adBoolean, adParamInput, , "False")
    Else
        cmd.Parameters.Append cmd.CreateParameter("@isReceivable", adBoolean, adParamInput, , "True")
    End If
    If cmbCheckStatus.ItemData(cmbCheckStatus.ListIndex) = -1 Then
        cmd.Parameters.Append cmd.CreateParameter("@CheckStatusId", adInteger, adParamInput, , Null)
    Else
        cmd.Parameters.Append cmd.CreateParameter("@CheckStatusId", adInteger, adParamInput, , cmbCheckStatus.ItemData(cmbCheckStatus.ListIndex))
    End If
    Set rec = cmd.Execute
    lvCheckRegistry.ListItems.Clear
    Dim item As MSComctlLib.ListItem
    If Not rec.EOF Then
        Do Until rec.EOF
                On Error Resume Next
                If optReceivable.value = True Then
                    If rec!isReceivable = "True" Then
                        Set item = lvCheckRegistry.ListItems.add(, , "")
                            item.SubItems(1) = rec!CheckRegistryId
                            item.SubItems(2) = rec!CheckNumber
                            item.SubItems(3) = rec!Name
                            item.SubItems(4) = rec!OrderNumber
                            item.SubItems(5) = rec!CheckDate
                            If rec!CheckStatusId = 1 Then
                                If DateDiff("d", Now, rec!CheckDate) < 0 Then
                                    item.SubItems(6) = DateDiff("d", Now, rec!CheckDate) * -1 & " DAY(S) OVERDUE"
                                Else
                                    item.SubItems(6) = "DUE IN " & DateDiff("d", Now, rec!CheckDate) & " DAY(S)"
                                End If
                            End If
                            item.SubItems(7) = rec!Remarks
                            item.SubItems(8) = FormatNumber(rec!Amount, 2, vbTrue)
                            item.SubItems(9) = rec!Status
                            If rec!DateCleared = Null Then
                                item.SubItems(10) = ""
                            Else
                                item.SubItems(10) = Format(rec!DateCleared, "MM/DD/YY")
                            End If
                    End If
                Else
                    If rec!isReceivable = "False" Then
                        Set item = lvCheckRegistry.ListItems.add(, , "")
                            item.SubItems(1) = rec!CheckRegistryId
                            item.SubItems(2) = rec!CheckNumber
                            item.SubItems(3) = rec!Name
                            item.SubItems(4) = rec!OrderNumber
                            item.SubItems(5) = rec!CheckDate
                            If rec!CheckStatusId = 1 Then
                                If DateDiff("d", Now, rec!CheckDate) < 0 Then
                                    item.SubItems(6) = DateDiff("d", Now, rec!CheckDate) * -1 & " DAY(S) OVERDUE"
                                Else
                                    item.SubItems(6) = "DUE IN " & DateDiff("d", Now, rec!CheckDate) & " DAY(S)"
                                End If
                            End If
                            item.SubItems(7) = rec!Remarks
                            item.SubItems(8) = FormatNumber(rec!Amount, 2, vbTrue)
                            item.SubItems(9) = rec!Status
                            If rec!DateCleared = Null Then
                                item.SubItems(10) = ""
                            Else
                                item.SubItems(10) = Format(rec!DateCleared, "MM/DD/YY")
                            End If
                    End If
                End If
            rec.MoveNext
        Loop
    End If
    con.Close
    CountTotal
End Sub

Private Sub Command1_Click()

End Sub

Private Sub btnUpdateStatus_Click()
    Dim item As MSComctlLib.ListItem
    Dim hasChecked As Boolean
    hasChecked = False
    For Each item In lvCheckRegistry.ListItems
        If item.Checked = True Then
            hasChecked = True
            Exit For
        End If
    Next
    
    Dim x As Variant
    Dim deduct, add As Boolean
    deduct = False
    add = False
    
    If cmbUpdateStatus.ItemData(cmbUpdateStatus.ListIndex) <> 1 Then  'PDC
        x = MsgBox("Would you like to DEDUCT the total selected check amount from your cashier's check balance?", vbQuestion + vbYesNo)
        If x = vbYes Then
            deduct = True
        Else
            deduct = False
        End If
    Else
        x = MsgBox("Would you like to ADD the total selected check amount to your cashier's check balance?", vbQuestion + vbYesNo)
        If x = vbYes Then
            add = True
        Else
            add = False
        End If
    End If
    
    Set con = New ADODB.Connection
    con.ConnectionString = ConnString
    con.Open
    
    If hasChecked = True Then
        For Each item In lvCheckRegistry.ListItems
            If item.Checked = True Then
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "FIN_CheckRegistryStatus_Update"
                cmd.Parameters.Append cmd.CreateParameter("@CheckRegistryId", adInteger, adParamInput, , item.SubItems(1))
                cmd.Parameters.Append cmd.CreateParameter("@CheckStatusId", adInteger, adParamInput, , cmbUpdateStatus.ItemData(cmbUpdateStatus.ListIndex))
                cmd.Parameters.Append cmd.CreateParameter("@DateCleared", adVarChar, adParamInput, 50, dtCleared.text)
                cmd.Execute
                
                If add = True Then
                    Set cmd = New ADODB.Command
                    cmd.ActiveConnection = con
                    cmd.CommandType = adCmdStoredProc
                    cmd.CommandText = "FIN_Fund_AddBalance"
                    cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , 1)
                    cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(item.SubItems(8), ",", "")))
                                          cmd.Parameters("@CheckAmount").NumericScale = 2
                                          cmd.Parameters("@CheckAmount").Precision = 18
                    cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , 0)
                                          cmd.Parameters("@Amount").NumericScale = 2
                                          cmd.Parameters("@Amount").Precision = 18
                    cmd.Execute
                End If
                
                If deduct = True Then
                    Set cmd = New ADODB.Command
                    cmd.ActiveConnection = con
                    cmd.CommandType = adCmdStoredProc
                    cmd.CommandText = "FIN_Fund_AddBalance"
                    cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , 1)
                    cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(item.SubItems(8), ",", "")) * -1)
                                          cmd.Parameters("@CheckAmount").NumericScale = 2
                                          cmd.Parameters("@CheckAmount").Precision = 18
                    cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , 0)
                                          cmd.Parameters("@Amount").NumericScale = 2
                                          cmd.Parameters("@Amount").Precision = 18
                    cmd.Execute
                    
                    Set cmd = New ADODB.Command
                    cmd.ActiveConnection = con
                    cmd.CommandType = adCmdStoredProc
                    cmd.CommandText = "FIN_FundBank_Add"
                    cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , 7) 'CHINABANK
                    cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(item.SubItems(8), ",", "")))
                                          cmd.Parameters("@CheckAmount").NumericScale = 2
                                          cmd.Parameters("@CheckAmount").Precision = 18
                    cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , 0)
                                          cmd.Parameters("@Amount").NumericScale = 2
                                          cmd.Parameters("@Amount").Precision = 18
                    cmd.Execute
                    
                    If optPayable.value = False Then
                        'INFLOW AND OUTFLOW
                        Set cmd = New ADODB.Command
                        cmd.ActiveConnection = con
                        cmd.CommandType = adCmdStoredProc
                        cmd.CommandText = "FIN_OutFlow_Insert"
                        cmd.Parameters.Append cmd.CreateParameter("@Cash", adDecimal, adParamInput, , 0)
                                              cmd.Parameters("@Cash").Precision = 18
                                              cmd.Parameters("@Cash").NumericScale = 2
                        cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(item.SubItems(8), ",", "")))
                                              cmd.Parameters("@CheckAmount").Precision = 18
                                              cmd.Parameters("@CheckAmount").NumericScale = 2
                        cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, item.SubItems(2))
                        cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , item.SubItems(5))
                        cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , Format(Now, "MM/DD/YY"))
                        cmd.Parameters.Append cmd.CreateParameter("@Particulars", adVarChar, adParamInput, 250, "CHECK OUT " & txtRemarks.text)
                        cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , Null)
                        cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , Null)
                        cmd.Execute
                    
                    
                        'BANK INFLOW
                        Set cmd = New ADODB.Command
                        cmd.ActiveConnection = con
                        cmd.CommandType = adCmdStoredProc
                        cmd.CommandText = "FIN_BankInflow_Insert"
                        cmd.Parameters.Append cmd.CreateParameter("@Cash", adDecimal, adParamInput, , 0)
                                              cmd.Parameters("@Cash").Precision = 18
                                              cmd.Parameters("@Cash").NumericScale = 2
                        cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(item.SubItems(8), ",", "")))
                                              cmd.Parameters("@CheckAmount").Precision = 18
                                              cmd.Parameters("@CheckAmount").NumericScale = 2
                        cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, item.SubItems(2))
                        cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , item.SubItems(5))
                        cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , Format(Now, "MM/DD/YY"))
                        cmd.Parameters.Append cmd.CreateParameter("@Particulars", adVarChar, adParamInput, 250, "CHECK OUT " & txtRemarks.text)
                        cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , Null)
                        cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , Null)
                        cmd.Execute
                    End If
                End If
            End If
        Next
    Else
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(36)
        GLOBAL_MessageFrm.Show (1)
        Exit Sub
    End If
    
    con.Close
    
    MsgBox "Record saved.", vbInformation
    Initialize
    btnSearch_Click
End Sub

Private Sub Form_Load()
    lvCheckRegistry.ColumnHeaders(3).width = lvCheckRegistry.width * 0.1
    lvCheckRegistry.ColumnHeaders(4).width = lvCheckRegistry.width * 0.1088
    lvCheckRegistry.ColumnHeaders(5).width = lvCheckRegistry.width * 0.1
    lvCheckRegistry.ColumnHeaders(6).width = lvCheckRegistry.width * 0.1088
    lvCheckRegistry.ColumnHeaders(7).width = lvCheckRegistry.width * 0.1088
    lvCheckRegistry.ColumnHeaders(8).width = lvCheckRegistry.width * 0.1088
    lvCheckRegistry.ColumnHeaders(9).width = lvCheckRegistry.width * 0.1088
    lvCheckRegistry.ColumnHeaders(10).width = lvCheckRegistry.width * 0.1088
    lvCheckRegistry.ColumnHeaders(11).width = lvCheckRegistry.width * 0.1088
    
    dtFrom.value = Format(Now, "mm/dd/yy")
    dtTo.value = Format(Now, "mm/dd/yy")
    dtCheckDate.value = Format(Now, "mm/dd/yy")
    
    Set rec = New ADODB.Recordset
    Set rec = Global_Data("CheckStatus")
    cmbStatus.Clear
    cmbUpdateStatus.Clear
    cmbCheckStatus.Clear
    cmbCheckStatus.AddItem ""
    cmbCheckStatus.ItemData(cmbCheckStatus.NewIndex) = -1
    Do Until rec.EOF
        cmbCheckStatus.AddItem rec!Status
        cmbCheckStatus.ItemData(cmbCheckStatus.NewIndex) = rec!CheckStatusId
        cmbStatus.AddItem rec!Status
        cmbStatus.ItemData(cmbStatus.NewIndex) = rec!CheckStatusId
        cmbUpdateStatus.AddItem rec!Status
        cmbUpdateStatus.ItemData(cmbUpdateStatus.NewIndex) = rec!CheckStatusId
        rec.MoveNext
    Loop
    cmbUpdateStatus.ListIndex = 2
    cmbStatus.ListIndex = 2
    cmbCheckStatus.ListIndex = 0
    
    lblTotal.Caption = "TOTAL:0.00"
    Initialize
End Sub

Private Sub lvCheckRegistry_ItemClick(ByVal item As MSComctlLib.ListItem)
    CheckRegistryId = item.SubItems(1)
    txtCheckNumber.text = item.SubItems(2)
    dtCheckDate.value = item.SubItems(5)
    txtAmount.text = FormatNumber(item.SubItems(8), 2, vbTrue, vbFalse)
    cmbStatus.text = item.SubItems(9)
    txtRemarks.text = item.SubItems(7)
    txtDateCleared.text = item.SubItems(10)
End Sub

Private Sub tb_Standard_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'NEW
            Initialize
        Case 2 'Save
            If Trim(txtCheckNumber.text) = "" Then
                GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(32)
                GLOBAL_MessageFrm.Show (1)
                txtCheckNumber.SetFocus
                Exit Sub
            End If
            
            Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            Set rec = New ADODB.Recordset
            
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            
            cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.text)
            cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
            cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Val(Replace(txtAmount.text, ",", "")))
                                  cmd.Parameters("@Amount").NumericScale = 2
                                  cmd.Parameters("@Amount").Precision = 18
            cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, txtRemarks.text)
            If optAddReceivable.value = True Then
                cmd.Parameters.Append cmd.CreateParameter("@isReceivable", adBoolean, adParamInput, , "True")
            ElseIf optAddPayable.value = True Then
                cmd.Parameters.Append cmd.CreateParameter("@isReceivable", adBoolean, adParamInput, , "False")
            Else
                cmd.Parameters.Append cmd.CreateParameter("@isReceivable", adBoolean, adParamInput, , "True")
            End If
            cmd.Parameters.Append cmd.CreateParameter("@CheckStatusId", adInteger, adParamInput, , cmbStatus.ItemData(cmbStatus.ListIndex))
            cmd.Parameters.Append cmd.CreateParameter("@CheckRegistryId", adInteger, adParamInputOutput, , CheckRegistryId)
            If CheckRegistryId = 0 Then
                cmd.CommandText = "FIN_CheckRegistry_Insert"
            Else
                cmd.CommandText = "FIN_CheckRegistry_Update"
            End If
            cmd.Execute
            
            If optAddReceivable.value = True Then
                Dim x As Variant
                x = MsgBox("Would you like to update the cashier's check balance?", vbQuestion + vbYesNo)
                If x = vbYes Then
                    'ADD/DEDUCT BALANCE
                    Set cmd = New ADODB.Command
                    cmd.ActiveConnection = con
                    cmd.CommandType = adCmdStoredProc
                    cmd.CommandText = "FIN_Fund_AddBalance"
                    cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , 1)
                    If optAddReceivable.value = True Then
                        cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(txtAmount.text, ",", "")))
                    ElseIf optAddPayable.value = True Then
                        cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(txtAmount.text, ",", "")) * -1)
                    Else
                        cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(txtAmount.text, ",", "")))
                    End If
                                          cmd.Parameters("@CheckAmount").NumericScale = 2
                                          cmd.Parameters("@CheckAmount").Precision = 18
                    cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , 0)
                                          cmd.Parameters("@Amount").NumericScale = 2
                                          cmd.Parameters("@Amount").Precision = 18
                    cmd.Execute
                End If
            Else
                'Dim x As Variant
                x = MsgBox("Would you like to update the bank account balance?", vbQuestion + vbYesNo)
                If x = vbYes Then
                    'ADD/DEDUCT BALANCE
                    Set cmd = New ADODB.Command
                    cmd.ActiveConnection = con
                    cmd.CommandType = adCmdStoredProc
                    cmd.CommandText = "FIN_Bank_AddBalance"
                    cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , 7) 'CHINABANK
                    cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , Format(Now, "MM/DD/YY"))
                    cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Val(Replace(txtAmount.text, ",", "")))
                                          cmd.Parameters("@Amount").NumericScale = 2
                                          cmd.Parameters("@Amount").Precision = 18
                    cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, , txtRemarks.text)
                    cmd.Execute
                End If
            End If
            
            'INFLOW AND OUTFLOW
            Dim note As String
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            If optAddReceivable.value = True Then
                cmd.CommandText = "FIN_InFlow_Insert"
                note = "CHECK IN"
            ElseIf optAddPayable.value = True Then
                cmd.CommandText = "FIN_OutFlow_Insert"
                note = "CHECK OUT"
            Else
                cmd.CommandText = "FIN_InFlow_Insert"
                note = "CHECK IN"
            End If
            cmd.Parameters.Append cmd.CreateParameter("@Cash", adDecimal, adParamInput, , 0)
                                  cmd.Parameters("@Cash").Precision = 18
                                  cmd.Parameters("@Cash").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(txtAmount, ",", "")))
                                  cmd.Parameters("@CheckAmount").Precision = 18
                                  cmd.Parameters("@CheckAmount").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.text)
            cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
            cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , Format(Now, "MM/DD/YY"))
            cmd.Parameters.Append cmd.CreateParameter("@Particulars", adVarChar, adParamInput, 250, note & " " & txtRemarks.text)
            cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , Null)
            cmd.Execute
            
            con.Close
            MsgBox "Record saved.", vbInformation
            Initialize
            btnSearch_Click
        Case 4 ' DELETE
            If CheckRegistryId = 0 Then
                GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(35)
                GLOBAL_MessageFrm.Show (1)
                Exit Sub
            Else
                Set con = New ADODB.Connection
                Set rec = New ADODB.Recordset
                Set cmd = New ADODB.Command
                con.ConnectionString = ConnString
                con.Open
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "FIN_CheckRegistry_Delete"
                cmd.Parameters.Append cmd.CreateParameter("@CheckRegistryId", adInteger, adParamInput, , CheckRegistryId)
                cmd.Execute
                con.Close
                MsgBox "Check deleted.", vbInformation
                Initialize
                btnSearch_Click
            End If
    End Select
End Sub

Private Sub txtAmount_Change()
    If IsNumeric(txtAmount.text) = False Then
        txtAmount.text = "0.00"
        selectText txtAmount
    Else
        txtAmount.text = FormatNumber(txtAmount.text, 0)
        txtAmount.SelStart = Len(txtAmount.text)
    End If
End Sub

Private Sub txtCheckNumber_GotFocus()
    selectText txtCheckNumber
End Sub
