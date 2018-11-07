VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form PO_ProductBySupplierFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Supplier Products"
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
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   9390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15135
      Begin VB.ComboBox cmbSearch_Vendor 
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
         TabIndex        =   3
         Top             =   1200
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
         TabIndex        =   4
         Top             =   1680
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvSearch 
         Height          =   6735
         Left            =   120
         TabIndex        =   5
         Top             =   2160
         Width           =   14805
         _ExtentX        =   26114
         _ExtentY        =   11880
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
            Text            =   "ProductId"
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
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "QTY"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Unit"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Cost"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Last Date Purchase"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Supplier"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
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
         TabIndex        =   9
         Top             =   1200
         Width           =   780
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   840
         Width           =   555
      End
   End
End
Attribute VB_Name = "PO_ProductBySupplierFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub Populate(ByVal data As String)
    'On Error Resume Next
    Select Case data
        Case "Vendor"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Vendor")
            cmbSearch_Vendor.Clear
            cmbSearch_Vendor.AddItem ""
            cmbSearch_Vendor.ItemData(cmbSearch_Vendor.NewIndex) = 0
            cmbSearch_Vendor.ListIndex = 0
            If Not rec.EOF Then
                Do Until rec.EOF
                    cmbSearch_Vendor.AddItem rec!Name
                    cmbSearch_Vendor.ItemData(cmbSearch_Vendor.NewIndex) = rec!VendorId
                    rec.MoveNext
                Loop
            End If
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
    cmd.CommandText = "PO_ProductBySupplier_Search"
    
    If Trim(txtSearch_ItemCode.text) = "" Then
        cmd.Parameters.Append cmd.CreateParameter("@Code", adVarChar, adParamInput, 50, Null)
    Else
        cmd.Parameters.Append cmd.CreateParameter("@Code", adVarChar, adParamInput, 50, txtSearch_ItemCode.text)
    End If
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, txtSearch_Name.text)
    If cmbSearch_Vendor.ListIndex = -1 Or cmbSearch_Vendor.ListIndex = 0 Then
        cmd.Parameters.Append cmd.CreateParameter("@VendorId", adInteger, adParamInput, , Null)
    Else
        cmd.Parameters.Append cmd.CreateParameter("@VendorId", adInteger, adParamInput, , cmbSearch_Vendor.ItemData(cmbSearch_Vendor.ListIndex))
    End If
    Set rec = cmd.Execute
    lvSearch.ListItems.Clear
    Dim item As MSComctlLib.ListItem
    If Not rec.EOF Then
        Do Until rec.EOF
            Set item = lvSearch.ListItems.add(, , rec!ProductId)
                item.SubItems(1) = rec!itemcode
                item.SubItems(2) = rec!Name
                item.SubItems(3) = FormatNumber(rec!quantity, 2, vbTrue, vbFalse)
                item.SubItems(4) = rec!uom
                item.SubItems(5) = FormatNumber(rec!cost, 2, vbTrue, vbFalse)
                item.SubItems(6) = Format(rec!receiveddate, "MM/DD/YY")
                item.SubItems(7) = rec!Vendor
            rec.MoveNext
        Loop
    End If
    con.Close
End Sub

Private Sub cmbSearch_Vendor_Click()
'    btnSearch_Click
End Sub

Private Sub Form_Load()
    lvSearch.ColumnHeaders(2).width = lvSearch.width * 0.1
    lvSearch.ColumnHeaders(3).width = lvSearch.width * 0.21
    lvSearch.ColumnHeaders(4).width = lvSearch.width * 0.11
    lvSearch.ColumnHeaders(5).width = lvSearch.width * 0.11
    lvSearch.ColumnHeaders(6).width = lvSearch.width * 0.11
    lvSearch.ColumnHeaders(7).width = lvSearch.width * 0.14
    lvSearch.ColumnHeaders(8).width = lvSearch.width * 0.2
    
    Populate "Vendor"
End Sub

Private Sub txtSearch_ItemCode_Change()
    btnSearch_Click
End Sub

Private Sub txtSearch_Name_Change()
    btnSearch_Click
End Sub
