VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form RPT_INV_InventorySummaryFrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Inventory Summary"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15090
   Icon            =   "RPT_INV_InventorySummary.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9015
   ScaleWidth      =   15090
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.ComboBox cmbOrientation 
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
         ItemData        =   "RPT_INV_InventorySummary.frx":6852
         Left            =   1320
         List            =   "RPT_INV_InventorySummary.frx":685C
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   5280
         Width           =   2415
      End
      Begin VB.CheckBox chkDisplay 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Display 0 Inventory"
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
         Left            =   1320
         TabIndex        =   19
         Top             =   1680
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.ComboBox cmbGroup 
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
         ItemData        =   "RPT_INV_InventorySummary.frx":6875
         Left            =   1320
         List            =   "RPT_INV_InventorySummary.frx":6882
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   3240
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txtDescription 
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
         Top             =   1320
         Width           =   2415
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
         ItemData        =   "RPT_INV_InventorySummary.frx":689D
         Left            =   1320
         List            =   "RPT_INV_InventorySummary.frx":68B3
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   4560
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
         TabIndex        =   9
         Top             =   6000
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
         Text            =   "Inventory Summary"
         Top             =   4920
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
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   960
         Width           =   2415
      End
      Begin VB.ComboBox cmbSupplier 
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
         ItemData        =   "RPT_INV_InventorySummary.frx":68F0
         Left            =   1320
         List            =   "RPT_INV_InventorySummary.frx":68FA
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Orientation"
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
         TabIndex        =   20
         Top             =   5280
         Width           =   1080
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group By"
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
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
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
         Top             =   3240
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Top             =   1320
         Width           =   1065
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
         TabIndex        =   14
         Top             =   4560
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
         TabIndex        =   13
         Top             =   4920
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
         TabIndex        =   12
         Top             =   4080
         Width           =   870
      End
      Begin VB.Label Label2 
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
         TabIndex        =   11
         Top             =   960
         Width           =   825
      End
      Begin VB.Label Label3 
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
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   780
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
         TabIndex        =   1
         Top             =   120
         Width           =   1005
      End
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer 
      Height          =   9015
      Left            =   3840
      TabIndex        =   2
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
Attribute VB_Name = "RPT_INV_InventorySummaryFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim crxApp As New CRAXDRT.Application
Dim crxRpt As New CRAXDRT.Report
Public Sub Populate(ByVal data As String)
    Dim Item As MSComctlLib.ListItem
    Select Case data
        Case "Category"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Category")
            cmbCategory.Clear
            cmbCategory.AddItem ""
            cmbCategory.ItemData(cmbCategory.NewIndex) = 0
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        cmbCategory.AddItem rec!Category
                        cmbCategory.ItemData(cmbCategory.NewIndex) = rec!CategoryId
                    End If
                    rec.MoveNext
                Loop
            End If
            cmbCategory.ListIndex = 0
        Case "Vendor"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Vendor")
            cmbSupplier.Clear
            cmbSupplier.AddItem ""
            cmbSupplier.ItemData(cmbSupplier.NewIndex) = 0
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        cmbSupplier.AddItem rec!Name
                        cmbSupplier.ItemData(cmbSupplier.NewIndex) = rec!VendorId
                    End If
                    rec.MoveNext
                Loop
            End If
            cmbSupplier.ListIndex = 0
    End Select
End Sub

Private Sub btnGenerate_Click()
    Dim sql, OrderBy As String
    Dim Status, Customer, Terms, DateRange As Variant
    Dim DisplayZero As Integer
    
    Screen.MousePointer = vbHourglass
    
    If cmbOrientation.ListIndex = 0 Then 'Landscape
        Select Case cmbGroup.ListIndex
            Case 0 'None
                Set crxRpt = crxApp.OpenReport(App.path & "\Reports\INV_InventoryReport_Simple.rpt")
            Case 1 'Supplier
                Set crxRpt = crxApp.OpenReport(App.path & "\Reports\INV_InventoryReport_Supplier.rpt")
            Case 2 'Category
                Set crxRpt = crxApp.OpenReport(App.path & "\Reports\INV_InventoryReport_Simple_Category.rpt")
        End Select
    Else 'Portrait
        Select Case cmbGroup.ListIndex
            Case 0 'None
                Set crxRpt = crxApp.OpenReport(App.path & "\Reports\INV_InventoryReport_Portrait_Simple.rpt")
            Case 1 'Supplier
                Set crxRpt = crxApp.OpenReport(App.path & "\Reports\INV_InventoryReport_Portrait_Supplier.rpt")
            Case 2 'Category
                Set crxRpt = crxApp.OpenReport(App.path & "\Reports\INV_InventoryReport_Portrait_Category.rpt")
        End Select
    End If
    
    If chkDisplay.value = Checked Then
        DisplayZero = -1
    Else
        DisplayZero = 1
    End If
    
    
    Call ResetRptDB(crxRpt)
    
    crxRpt.EnableParameterPrompting = False
    crxRpt.DiscardSavedData
    
    
    crxRpt.ParameterFields.GetItemByName("ReportTitle").AddCurrentValue txtTitle.Text
    crxRpt.ParameterFields.GetItemByName("@CategoryId").AddCurrentValue cmbCategory.ItemData(cmbCategory.ListIndex)
    crxRpt.ParameterFields.GetItemByName("@SupplierId").AddCurrentValue cmbSupplier.ItemData(cmbSupplier.ListIndex)
    crxRpt.ParameterFields.GetItemByName("@Sort").AddCurrentValue cmbSort.Text
    crxRpt.ParameterFields.GetItemByName("@Description").AddCurrentValue txtDescription.Text
    crxRpt.ParameterFields.GetItemByName("@DisplayZero").AddCurrentValue DisplayZero
    
    CRViewer.ReportSource = crxRpt
    CRViewer.ViewReport
    CRViewer.Zoom 1
    Screen.MousePointer = vbDefault
    
    crxApp.CanClose
    Set crxApp = Nothing
End Sub

Private Sub CRViewer_PrintButtonClicked(UseDefault As Boolean)
    UseDefault = False
    crxRpt.PrinterSetup Me.hWnd
    crxRpt.PrintOut True
End Sub

Private Sub Form_Load()
    cmbSupplier.ListIndex = 0
    cmbSort.ListIndex = 0
    cmbGroup.ListIndex = 0
    cmbOrientation.ListIndex = 0
    
    Populate "Category"
    Populate "Vendor"
    
    Me.Height = 9390
    Me.width = 15180
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    CRViewer.width = Me.width - Frame1.width
    CRViewer.Height = Me.Height
    Frame1.Height = Me.Height
End Sub

