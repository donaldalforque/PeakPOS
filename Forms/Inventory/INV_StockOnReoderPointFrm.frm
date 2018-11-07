VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form INV_StockOnReorderPointFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCancel 
      Caption         =   "Close"
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
      Left            =   8160
      TabIndex        =   11
      Top             =   9120
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   9135
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
         Left            =   7680
         TabIndex        =   10
         Top             =   1440
         Width           =   1215
      End
      Begin VB.ComboBox cmbSupplier 
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
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1080
         Width           =   7575
      End
      Begin VB.ComboBox cmbCategory 
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
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   720
         Width           =   7575
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
         Left            =   1320
         TabIndex        =   5
         Top             =   360
         Width           =   7575
      End
      Begin VB.Label Label4 
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
         Top             =   1080
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
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
         TabIndex        =   7
         Top             =   720
         Width           =   825
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
         TabIndex        =   4
         Top             =   360
         Width           =   555
      End
   End
   Begin MSComctlLib.ListView lvReorder 
      Height          =   5535
      Left            =   240
      TabIndex        =   2
      Top             =   3480
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   9763
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
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ProductId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Reorder Point"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Current Qty"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Reorder Quantity"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8880
      Top             =   120
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
            Picture         =   "INV_StockOnReoderPointFrm.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_StockOnReoderPointFrm.frx":6862
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_StockOnReoderPointFrm.frx":D0C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_StockOnReoderPointFrm.frx":13926
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_StockOnReoderPointFrm.frx":13B9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_StockOnReoderPointFrm.frx":1420C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tb_Standard 
      Height          =   330
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   582
      ButtonWidth     =   1376
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Create Purchase Order"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
            Caption         =   "Print"
            ImageIndex      =   4
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SalesOrder"
                  Text            =   "Sales Order"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "PickList"
                  Text            =   "Pick List"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Invoice"
                  Text            =   "Sales Invoice"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin VB.Label lblSelectAll 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select All"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   240
      MouseIcon       =   "INV_StockOnReoderPointFrm.frx":1AA6E
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   9240
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unselect All"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   1200
      MouseIcon       =   "INV_StockOnReoderPointFrm.frx":1ABC0
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   9240
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "You can view here the list of stocks that reached reorder point that may require new purchase."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   8175
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stocks on Reorder Point"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   435
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   3510
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "INV_StockOnReoderPointFrm.frx":1AD12
      Top             =   480
      Width           =   480
   End
End
Attribute VB_Name = "INV_StockOnReorderPointFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VendorId As Integer
Dim CategoryId As Integer

Private Sub btnCancel_Click()
    Unload Me
End Sub
Private Sub Populate(ByVal data As String)
    Set rec = New ADODB.Recordset
    Select Case data
        Case "Supplier"
            Set rec = Global_Data("Vendor")
            
            cmbSupplier.Clear
            cmbSupplier.AddItem ""
            cmbSupplier.ItemData(cmbSupplier.NewIndex) = 0
            
            
            If Not rec.EOF Then
                Do Until rec.EOF
                    cmbSupplier.AddItem rec!Name
                    cmbSupplier.ItemData(cmbSupplier.NewIndex) = rec!VendorId
                    rec.MoveNext
                Loop
            End If
            
            On Error Resume Next
            cmbSupplier.ListIndex = 0
        Case "Category"
            Set rec = Global_Data("Category")
            
            cmbCategory.Clear
            cmbCategory.AddItem ""
            cmbCategory.ItemData(cmbCategory.NewIndex) = 0
            
            
            If Not rec.EOF Then
                Do Until rec.EOF
                    cmbCategory.AddItem rec!Category
                    cmbCategory.ItemData(cmbCategory.NewIndex) = rec!CategoryId
                    rec.MoveNext
                Loop
            End If
            On Error Resume Next
            cmbCategory.ListIndex = 0
    End Select
End Sub
Private Sub LoadData()
    Dim item As MSComctlLib.ListItem
    Dim con As New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "INV_ProductsOnReorderPoint_Get"
    cmd.Parameters.Append cmd.CreateParameter("@CategoryId", adInteger, adParamInput, , CategoryId)
    cmd.Parameters.Append cmd.CreateParameter("@VendorId", adInteger, adParamInput, , VendorId)
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, txtName.Text)
    Set rec = cmd.Execute
    
    lvReorder.ListItems.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            Set item = lvReorder.ListItems.add(, , "")
                item.SubItems(1) = rec!ProductId
                item.SubItems(2) = rec!itemcode
                item.SubItems(3) = rec!Name
                item.SubItems(4) = FormatNumber(rec!reorderpoint, 2, vbTrue, vbFalse)
                item.SubItems(5) = FormatNumber(rec!currentqty, 2, vbTrue, vbFalse)
            rec.MoveNext
        Loop
    End If
    con.Close
End Sub

Private Sub btnSearch_Click()
    CategoryId = cmbCategory.ItemData(cmbCategory.ListIndex)
    VendorId = cmbSupplier.ItemData(cmbSupplier.ListIndex)
    LoadData
End Sub

Private Sub cmbCategory_Click()
'    btnSearch_Click
End Sub

Private Sub cmbSupplier_Click()
'    btnSearch_Click
End Sub

Private Sub Form_Load()
    lvReorder.ColumnHeaders(1).width = lvReorder.width * 0.03
    lvReorder.ColumnHeaders(3).width = lvReorder.width * 0.16
    lvReorder.ColumnHeaders(4).width = lvReorder.width * 0.42
    lvReorder.ColumnHeaders(5).width = lvReorder.width * 0.17
    lvReorder.ColumnHeaders(6).width = lvReorder.width * 0.17
    
    Populate "Category"
    Populate "Supplier"
    
    LoadData
End Sub

Private Sub tb_Standard_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 6 'PRINT PREVIEW
            Screen.MousePointer = vbHourglass
            BASE_PrintPreviewFrm.Show
            Dim crxApp As New CRAXDRT.Application
            Dim crxRpt As New CRAXDRT.Report
            Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\INV_ProductReorderpoint.rpt")
            crxRpt.EnableParameterPrompting = False
            crxRpt.DiscardSavedData
            Call ResetRptDB(crxRpt)
            
            crxRpt.ParameterFields.GetItemByName("ReportTitle").AddCurrentValue "Products for Reorder"
            crxRpt.ParameterFields.GetItemByName("@CategoryId").AddCurrentValue cmbCategory.ItemData(cmbCategory.ListIndex)
            crxRpt.ParameterFields.GetItemByName("@VendorId").AddCurrentValue cmbSupplier.ItemData(cmbSupplier.ListIndex)
            crxRpt.ParameterFields.GetItemByName("@Name").AddCurrentValue txtName.Text
            

            BASE_PrintPreviewFrm.CRViewer.ReportSource = crxRpt
            BASE_PrintPreviewFrm.CRViewer.ViewReport
            BASE_PrintPreviewFrm.CRViewer.Zoom 1
            Screen.MousePointer = vbDefault
    End Select
End Sub

Private Sub txtName_Change()
    btnSearch_Click
End Sub
