VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form RPT_POS_SalesByQuantityFrm 
   Caption         =   "Sales by Quantity"
   ClientHeight    =   9015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15090
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9015
   ScaleWidth      =   15090
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3855
      Begin VB.TextBox txtOR 
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
         TabIndex        =   11
         Top             =   7320
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.OptionButton optOrderDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Order Date"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   -9999
         TabIndex        =   10
         Top             =   2040
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox cmbSort 
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
         ItemData        =   "RPT_POS_SalesByQuantityFrm.frx":0000
         Left            =   1320
         List            =   "RPT_POS_SalesByQuantityFrm.frx":001B
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   4800
         Width           =   2415
      End
      Begin VB.CommandButton btnGenerate 
         Caption         =   "Generate Report"
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
         Left            =   1920
         TabIndex        =   8
         Top             =   5880
         Width           =   1815
      End
      Begin VB.TextBox txtTitle 
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
         TabIndex        =   7
         Top             =   5160
         Width           =   2415
      End
      Begin VB.ComboBox cmbProduct 
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
         ItemData        =   "RPT_POS_SalesByQuantityFrm.frx":004D
         Left            =   1320
         List            =   "RPT_POS_SalesByQuantityFrm.frx":004F
         TabIndex        =   5
         Top             =   720
         Width           =   2415
      End
      Begin VB.OptionButton optDetails 
         BackColor       =   &H00E0E0E0&
         Caption         =   "by Invoice per Product"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   4
         Top             =   3360
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.OptionButton optQuantitySold 
         BackColor       =   &H00E0E0E0&
         Caption         =   "by Qty. sold per Product"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   3
         Top             =   3720
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.ComboBox cmbCategory 
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
         ItemData        =   "RPT_POS_SalesByQuantityFrm.frx":0051
         Left            =   1320
         List            =   "RPT_POS_SalesByQuantityFrm.frx":0053
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1080
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker TimeFrom 
         Height          =   375
         Left            =   1320
         TabIndex        =   6
         Top             =   2040
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
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
         Format          =   87621634
         UpDown          =   -1  'True
         CurrentDate     =   42217
      End
      Begin MSComCtl2.DTPicker DateTo 
         Height          =   345
         Left            =   1320
         TabIndex        =   12
         Top             =   2640
         Width           =   2415
         _ExtentX        =   4260
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
         Format          =   87621633
         CurrentDate     =   41686
      End
      Begin MSComCtl2.DTPicker DateFrom 
         Height          =   345
         Left            =   1320
         TabIndex        =   13
         Top             =   1680
         Width           =   2415
         _ExtentX        =   4260
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
         Format          =   87621633
         CurrentDate     =   41686
      End
      Begin MSComCtl2.DTPicker TimeTo 
         Height          =   375
         Left            =   1320
         TabIndex        =   14
         Top             =   3000
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
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
         Format          =   87621634
         UpDown          =   -1  'True
         CurrentDate     =   42217
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "POS OR#"
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
         TabIndex        =   23
         Top             =   7320
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sort By"
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
         TabIndex        =   22
         Top             =   4800
         Width           =   645
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Report Title"
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
         Top             =   5160
         Width           =   1095
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Display"
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
         TabIndex        =   20
         Top             =   4320
         Width           =   870
      End
      Begin VB.Label Label2 
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
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Filter By"
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
         TabIndex        =   18
         Top             =   120
         Width           =   1005
      End
      Begin VB.Label Label7 
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
         Left            =   120
         TabIndex        =   17
         Top             =   2640
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Product"
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
         TabIndex        =   16
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label8 
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
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   825
      End
   End
   Begin MSComctlLib.ListView lvItemList 
      Height          =   3735
      Left            =   -6960
      TabIndex        =   0
      Top             =   1080
      Width           =   7575
      _ExtentX        =   13361
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
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
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer 
      Height          =   9015
      Left            =   3840
      TabIndex        =   24
      Top             =   0
      Width           =   11295
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "RPT_POS_SalesByQuantityFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim crxApp As New CRAXDRT.Application
Dim crxRpt As New CRAXDRT.Report
Dim ProductId As Long

Private Sub btnGenerate_Click()
    Dim sql, OrderBy As String
    Dim Status, Customer, Terms, DateRange As Variant
    
    Screen.MousePointer = vbHourglass
    If optDetails.value = True Then
        Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\POS_SalesByProduct.rpt")
        crxRpt.EnableParameterPrompting = False
        crxRpt.DiscardSavedData
        Call ResetRptDB(crxRpt)
        
        crxRpt.ParameterFields.GetItemByName("ReportTitle").AddCurrentValue txtTitle.text
        crxRpt.ParameterFields.GetItemByName("DateFrom").AddCurrentValue DateFrom.value & " " & TimeFrom.value
        crxRpt.ParameterFields.GetItemByName("DateTo").AddCurrentValue DateTo.value & " " & TimeTo.value
        crxRpt.ParameterFields.GetItemByName("@ProductId").AddCurrentValue ProductId
        crxRpt.ParameterFields.GetItemByName("@CategoryId").AddCurrentValue cmbCategory.ItemData(cmbCategory.ListIndex)
        crxRpt.ParameterFields.GetItemByName("@DateFrom").AddCurrentValue DateFrom.value
        crxRpt.ParameterFields.GetItemByName("@DateTo").AddCurrentValue DateTo.value
        crxRpt.ParameterFields.GetItemByName("@TimeFrom").AddCurrentValue Str(TimeFrom.value)
        crxRpt.ParameterFields.GetItemByName("@TimeTo").AddCurrentValue Str(TimeTo.value)
        crxRpt.ParameterFields.GetItemByName("@Sort").AddCurrentValue cmbSort.text
        
    Else
        Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\POS_byQuantitySold.rpt")
        crxRpt.EnableParameterPrompting = False
        crxRpt.DiscardSavedData
        Call ResetRptDB(crxRpt)
        
        crxRpt.ParameterFields.GetItemByName("ReportTitle").AddCurrentValue txtTitle.text
        crxRpt.ParameterFields.GetItemByName("DateFrom").AddCurrentValue DateFrom.value & " " & TimeFrom.value
        crxRpt.ParameterFields.GetItemByName("DateTo").AddCurrentValue DateTo.value & " " & TimeTo.value
        crxRpt.ParameterFields.GetItemByName("@ProductId").AddCurrentValue ProductId
        crxRpt.ParameterFields.GetItemByName("@DateFrom").AddCurrentValue DateFrom.value
        crxRpt.ParameterFields.GetItemByName("@DateTo").AddCurrentValue DateTo.value
        crxRpt.ParameterFields.GetItemByName("@TimeFrom").AddCurrentValue Str(TimeFrom.value)
        crxRpt.ParameterFields.GetItemByName("@TimeTo").AddCurrentValue Str(TimeTo.value)
        crxRpt.ParameterFields.GetItemByName("@Sort").AddCurrentValue cmbSort.text
    End If
    
    CRViewer.ReportSource = crxRpt
    CRViewer.ViewReport
    CRViewer.Zoom 1
    Screen.MousePointer = vbDefault
End Sub



Private Sub cmbProduct_Change()
    If Trim(cmbProduct.text) = "" Then
        ProductId = 0
        lvItemList.Left = -9999
        Exit Sub
    End If
    
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    Dim item As MSComctlLib.ListItem
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_Product_Search1"
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 500, cmbProduct.text)
    Dim LastProductId As Long
    Set rec = cmd.Execute
    If Not rec.EOF Then
        lvItemList.ListItems.Clear
        Do Until rec.EOF
            If rec!isActive = "True" Then
                'If LastProductId <> rec!ProductId Then
                    Set item = lvItemList.ListItems.add(, , rec!ProductId)
                        item.SubItems(1) = rec!itemcode
                        item.SubItems(2) = rec!Name
                        item.SubItems(3) = FormatNumber(rec!unitprice, 2, vbTrue)
                        item.SubItems(4) = rec!Uom
                    lvItemList.Visible = True
                    lvItemList.Left = 1320
                    'lvItemList.Top = 3720
                    'LastProductId = rec!ProductId
                    rec.MoveNext
                'Else
                    'rec.MoveNext
                'End If
            Else
                rec.MoveNext
            End If
        Loop
    Else
        ProductId = 0
        lvItemList.Visible = False
        lvItemList.Left = -9999
    End If
    
    If Trim(cmbProduct.text) = "" Then ProductId = 0
    
    con.Close
End Sub

Private Sub cmbProduct_GotFocus()
    selectText cmbProduct
End Sub

Private Sub cmbProduct_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyUp, vbKeyDown
            If lvItemList.Visible = True Then
                lvItemList.SetFocus
            End If
    End Select
End Sub



Private Sub CRViewer_PrintButtonClicked(UseDefault As Boolean)
    UseDefault = False
    crxRpt.PrinterSetup Me.hWnd
    crxRpt.PrintOut True
End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            lvItemList.Left = -9999
            cmbProduct.SetFocus
    End Select
End Sub

Private Sub Form_Load()
    cmbSort.ListIndex = 0
    
    Me.Height = 9390
    Me.width = 15180
    DateFrom.value = Format(Now, "MM/DD/YY")
    DateTo.value = Format(Now, "MM/DD/YY")
    
    TimeFrom.value = Format("00:00:00", "hh:mm:ss")
    TimeTo.value = Format("23:59:59", "hh:mm:ss")
    
    txtTitle.text = Me.Caption
    
    lvItemList.ColumnHeaders(1).width = 0
    lvItemList.ColumnHeaders(2).width = lvItemList.width * 0.2
    lvItemList.ColumnHeaders(3).width = lvItemList.width * 0.76
    lvItemList.ColumnHeaders(4).width = lvItemList.width * 0
    
    'CATEGORY
    Set rec = New ADODB.Recordset
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
    
    cmbCategory.ListIndex = 0
End Sub


Private Sub Form_Resize()
   On Error Resume Next
    CRViewer.width = Me.width - Frame1.width
    CRViewer.Height = Me.Height
    Frame1.Height = Me.Height
End Sub

Private Sub lvItemList_DblClick()
    If lvItemList.ListItems.Count > 0 Then
        ProductId = lvItemList.SelectedItem.text
        cmbProduct.text = lvItemList.SelectedItem.SubItems(2)
        lvItemList.Left = -9999
        cmbProduct.SetFocus
    End If
End Sub

Private Sub lvItemList_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            lvItemList_DblClick
    End Select
End Sub

