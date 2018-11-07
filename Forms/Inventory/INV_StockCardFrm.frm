VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form INV_StockCardFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   14910
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvSearch 
      Height          =   3735
      Left            =   1560
      TabIndex        =   13
      Top             =   1530
      Visible         =   0   'False
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   6588
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
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
         Text            =   "ProductId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Item Code"
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
         Text            =   "Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Uom"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Cost"
         Object.Width           =   0
      EndProperty
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
      Left            =   12600
      TabIndex        =   10
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox txtParticulars 
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
      Left            =   1560
      TabIndex        =   3
      Top             =   1560
      Width           =   13215
   End
   Begin VB.TextBox txtItemCode 
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
      Left            =   1560
      TabIndex        =   1
      Top             =   1200
      Width           =   13215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   13560
      Top             =   2520
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
            Picture         =   "INV_StockCardFrm.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_StockCardFrm.frx":6862
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_StockCardFrm.frx":D0C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_StockCardFrm.frx":13926
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_StockCardFrm.frx":13B9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_StockCardFrm.frx":1420C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tb_Standard 
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15015
      _ExtentX        =   26485
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
            Caption         =   "New"
            ImageIndex      =   1
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
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   345
      Left            =   1560
      TabIndex        =   5
      Top             =   1920
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   609
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
      Format          =   96272385
      CurrentDate     =   41818
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   345
      Left            =   1560
      TabIndex        =   6
      Top             =   2280
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   609
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
      Format          =   96272385
      CurrentDate     =   41818
   End
   Begin MSComctlLib.ListView lvStockCard 
      Height          =   5055
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   14685
      _ExtentX        =   25903
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Ref #"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Customer/Vendor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Unit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Particulars"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "P/S Unit"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Cost"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Qty In"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Qty Out"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "Total Qty"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Product Movement History"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   390
      Left            =   840
      TabIndex        =   12
      Top             =   480
      Width           =   3465
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "INV_StockCardFrm.frx":1AA6E
      Top             =   480
      Width           =   480
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Current Inventory: 0.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      TabIndex        =   11
      Top             =   8400
      Width           =   14655
   End
   Begin VB.Label Label4 
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
      Top             =   1920
      Width           =   960
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date To"
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
      Top             =   2280
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Particulars"
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
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code/Name"
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
      TabIndex        =   2
      Top             =   1200
      Width           =   1125
   End
End
Attribute VB_Name = "INV_StockCardFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ProductId As String
Option Explicit
Dim crxApp As New CRAXDRT.Application
Dim crxRpt As New CRAXDRT.Report
Dim UserId As Long

Private Sub btnSearch_Click()
    Dim item As MSComctlLib.ListItem
    Dim con As New ADODB.Connection
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "INV_StockCard_Search"
    cmd.Parameters.Append cmd.CreateParameter("@DateFrom", adDate, adParamInput, , dtFrom.value)
    cmd.Parameters.Append cmd.CreateParameter("@DateTo", adDate, adParamInput, , dtTo.value)
    cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , Val(ProductId))
    cmd.Parameters.Append cmd.CreateParameter("@Particulars", adVarChar, adParamInput, 400, txtParticulars.Text)
    Set rec = cmd.Execute
    lvStockCard.ListItems.Clear
    'On Error Resume Next
    If Not rec.EOF Then
        Do Until rec.EOF
            If IsNull(rec!Date) = False Then
                Set item = lvStockCard.ListItems.add(, , rec!Date)
            Else
                Set item = lvStockCard.ListItems.add(, , "")
            End If
                item.SubItems(1) = rec!Reference
                If IsNull(rec!custvend) = False Then item.SubItems(2) = rec!custvend
                If IsNull(rec!itemcode) = False Then item.SubItems(3) = rec!itemcode
                If IsNull(rec!ProductName) = False Then item.SubItems(4) = rec!ProductName
                If IsNull(rec!standardunit) = False Then item.SubItems(5) = rec!standardunit
                item.SubItems(6) = rec!Particulars
                If IsNull(rec!sellingunit) = False Then item.SubItems(7) = rec!sellingunit
                If Not IsNull(rec!cost) Then
                    item.SubItems(8) = FormatNumber(rec!cost, 2, vbTrue, vbFalse)
                End If
                If Not IsNull(rec!quantityin) Then
                    item.SubItems(9) = FormatNumber(rec!quantityin, 2, vbTrue, vbFalse)
                End If
                If Not IsNull(rec!quantityout) Then
                    item.SubItems(10) = FormatNumber(rec!quantityout, 2, vbTrue, vbFalse)
                End If
            rec.MoveNext
        Loop
    End If
    con.Close
    
    Dim runningtotal As Double
    For Each item In lvStockCard.ListItems
        runningtotal = runningtotal - NVAL(item.SubItems(10)) + NVAL(item.SubItems(9))
        item.SubItems(11) = FormatNumber(runningtotal, 2, vbTrue, vbFalse)
        lblTotal.Caption = "Current Inventory: " & FormatNumber(runningtotal, 2, vbTrue, vbFalse)
    Next
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            lvSearch.Visible = False
            txtItemCode.SetFocus
    End Select
End Sub

Private Sub Form_Load()
    lvStockCard.ColumnHeaders(1).width = lvStockCard.width * 0.08
    lvStockCard.ColumnHeaders(2).width = lvStockCard.width * 0.1
    lvStockCard.ColumnHeaders(3).width = lvStockCard.width * 0.1
    lvStockCard.ColumnHeaders(4).width = lvStockCard.width * 0.1
    lvStockCard.ColumnHeaders(5).width = lvStockCard.width * 0.2
    lvStockCard.ColumnHeaders(6).width = lvStockCard.width * 0.07
    lvStockCard.ColumnHeaders(7).width = lvStockCard.width * 0.2
    lvStockCard.ColumnHeaders(8).width = lvStockCard.width * 0.07
    lvStockCard.ColumnHeaders(9).width = lvStockCard.width * 0.07
    lvStockCard.ColumnHeaders(10).width = lvStockCard.width * 0.07
    lvStockCard.ColumnHeaders(11).width = lvStockCard.width * 0.07
    lvStockCard.ColumnHeaders(12).width = lvStockCard.width * 0.07
    
    lvSearch.ColumnHeaders(2).width = lvSearch.width * 0.25
    lvSearch.ColumnHeaders(3).width = lvSearch.width * 0.72
    lvSearch.ColumnHeaders(4).width = lvSearch.width * 0
    
    Dim currdate As String
    currdate = "January 1," & Year(Now)
    
    dtFrom.value = Format(currdate, "MM/DD/YY")
    dtTo.value = Format(Now, "MM/DD/YY")
    
    ProductId = "0"
End Sub

Private Sub lvSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If lvSearch.ListItems.Count > 0 Then
                ProductId = lvSearch.SelectedItem.Text
                txtItemCode.Text = lvSearch.SelectedItem.SubItems(2)
                lvSearch.Visible = False
                btnSearch.SetFocus
                btnSearch_Click
            End If
    End Select
    
End Sub

Private Sub tb_Standard_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 6
            If Val(ProductId) = 0 Then Exit Sub
            
            Dim sql, OrderBy As String
            Dim Status, Customer, Terms, DateRange As Variant
            
            Screen.MousePointer = vbHourglass
            Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\INV_StockCard.rpt")
            crxRpt.EnableParameterPrompting = False
            crxRpt.DiscardSavedData
            Call ResetRptDB(crxRpt)
            
            crxRpt.ParameterFields.GetItemByName("ReportTitle").AddCurrentValue "Product Movement History"
            crxRpt.ParameterFields.GetItemByName("DateFrom").AddCurrentValue Str(dtFrom.value)
            crxRpt.ParameterFields.GetItemByName("DateTo").AddCurrentValue Str(dtTo.value)
            crxRpt.ParameterFields.GetItemByName("ProductName").AddCurrentValue txtItemCode.Text
            
            crxRpt.ParameterFields.GetItemByName("@DateFrom").AddCurrentValue dtFrom.value
            crxRpt.ParameterFields.GetItemByName("@DateTo").AddCurrentValue dtTo.value
            crxRpt.ParameterFields.GetItemByName("@ProductId").AddCurrentValue Val(ProductId)
            crxRpt.ParameterFields.GetItemByName("@Particulars").AddCurrentValue txtParticulars.Text
            
            BASE_PrintPreviewFrm.CRViewer.ReportSource = crxRpt
            BASE_PrintPreviewFrm.CRViewer.ViewReport
            BASE_PrintPreviewFrm.CRViewer.Zoom 1
            BASE_PrintPreviewFrm.Show
            Screen.MousePointer = vbDefault
    End Select
End Sub

Private Sub txtItemCode_Change()
    If Trim(txtItemCode.Text) = "" Then
        lvSearch.Visible = False
        Exit Sub
    End If
    
    Dim item As MSComctlLib.ListItem
    Dim con As New ADODB.Connection
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "INV_Product_Search"
    cmd.Parameters.Append cmd.CreateParameter("@String", adVarChar, adParamInput, 400, txtItemCode.Text)
    Set rec = cmd.Execute
    lvSearch.ListItems.Clear
    If Not rec.EOF Then
        lvSearch.Visible = True
        Do Until rec.EOF
            If rec!isActive = "True" Then
                Set item = lvSearch.ListItems.add(, , rec!ProductId)
                    item.SubItems(1) = rec!itemcode
                    item.SubItems(2) = rec!Name
            End If
            rec.MoveNext
        Loop
    End If
    con.Close
End Sub

Private Sub txtItemCode_GotFocus()
    selectText txtItemCode
End Sub

Private Sub txtItemCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            If lvSearch.Visible = True Then
                lvSearch.SetFocus
            End If
    End Select
End Sub
