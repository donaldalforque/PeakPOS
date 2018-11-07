VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form RPT_FIN_CheckRegistryFrm1 
   Caption         =   "Check Registry Report"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   345
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
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.OptionButton chkReceivable 
         BackColor       =   &H00E0E0E0&
         Caption         =   "RECEIVABLES"
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
         Left            =   1320
         TabIndex        =   18
         Top             =   960
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.OptionButton chkPayable 
         BackColor       =   &H00E0E0E0&
         Caption         =   "PAYABLES"
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
         Left            =   1320
         TabIndex        =   17
         Top             =   600
         Value           =   -1  'True
         Width           =   2415
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
         ItemData        =   "RPT_FIN_CheckRegistryFrm.frx":0000
         Left            =   1320
         List            =   "RPT_FIN_CheckRegistryFrm.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2040
         Width           =   2415
      End
      Begin VB.ComboBox cmbVendor 
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
         ItemData        =   "RPT_FIN_CheckRegistryFrm.frx":0004
         Left            =   1320
         List            =   "RPT_FIN_CheckRegistryFrm.frx":000E
         TabIndex        =   4
         Text            =   "cmbCustomer"
         Top             =   1440
         Width           =   2415
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
         TabIndex        =   3
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
         TabIndex        =   2
         Top             =   5280
         Width           =   1815
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
         ItemData        =   "RPT_FIN_CheckRegistryFrm.frx":0033
         Left            =   1320
         List            =   "RPT_FIN_CheckRegistryFrm.frx":0046
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   4200
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker DateTo 
         Height          =   345
         Left            =   1320
         TabIndex        =   5
         Top             =   2880
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
         Format          =   95748097
         CurrentDate     =   41686
      End
      Begin MSComCtl2.DTPicker DateFrom 
         Height          =   345
         Left            =   1320
         TabIndex        =   6
         Top             =   2520
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
         Format          =   95748097
         CurrentDate     =   41686
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
         Left            =   120
         TabIndex        =   16
         Top             =   2040
         Width           =   570
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Payee / Customer"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   1170
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
         TabIndex        =   12
         Top             =   2880
         Width           =   705
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
         TabIndex        =   11
         Top             =   120
         Width           =   1005
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
         TabIndex        =   10
         Top             =   2520
         Width           =   960
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
         TabIndex        =   9
         Top             =   3720
         Width           =   870
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
         TabIndex        =   8
         Top             =   4560
         Width           =   1095
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
         TabIndex        =   7
         Top             =   4200
         Width           =   645
      End
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer 
      Height          =   9015
      Left            =   3840
      TabIndex        =   14
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
Attribute VB_Name = "RPT_FIN_CheckRegistryFrm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim crxApp As New CRAXDRT.Application
Dim crxRpt As New CRAXDRT.Report
Public Sub Populate(ByVal data As String)
    Dim item As MSComctlLib.ListItem
    Select Case data
        Case "Vendor"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Vendor")
            cmbVendor.Clear
            cmbVendor.AddItem ""
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        cmbVendor.AddItem rec!Name
                        cmbVendor.ItemData(cmbVendor.NewIndex) = rec!VendorId
                    End If
                    rec.MoveNext
                Loop
            End If
            cmbVendor.ListIndex = 0
        Case "CheckStatus"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("CheckStatus")
            cmbCheckStatus.Clear
            cmbCheckStatus.AddItem ""
            If Not rec.EOF Then
                Do Until rec.EOF
                    'If rec!isActive = "True" Then
                        cmbCheckStatus.AddItem rec!Status
                        cmbCheckStatus.ItemData(cmbCheckStatus.NewIndex) = rec!CheckStatusId
                    'End If
                    rec.MoveNext
                Loop
            End If
            On Error Resume Next
            cmbCheckStatus.ListIndex = 0
    End Select
End Sub

Private Sub btnGenerate_Click()
    Dim sql, OrderBy As String
    Dim Status, Vendor, CheckStatus As Variant
    
    Screen.MousePointer = vbHourglass
    If chkPayable.value = True Then
        Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\CheckRegistry.rpt")
    Else
        Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\CheckRegistryreceivable.rpt")
    End If
    crxRpt.EnableParameterPrompting = False
    crxRpt.DiscardSavedData
    Call ResetRptDB(crxRpt)
      
    Select Case cmbVendor.ListIndex
        Case -1
            Vendor = ""
        Case 0
            Vendor = ""
        Case Else
            Vendor = "AND BASE_Vendor.VendorId = '" & cmbVendor.ItemData(cmbVendor.ListIndex) & "' "
    End Select
    
    Select Case cmbCheckStatus.ListIndex
        Case -1
            CheckStatus = ""
        Case 0
            CheckStatus = ""
        Case Else
            CheckStatus = "AND FIN_CheckStatus.checkStatusId = '" & cmbCheckStatus.ItemData(cmbCheckStatus.ListIndex) & "' "
    End Select
    
    
    Select Case cmbSort.ListIndex
        Case 0
            OrderBy = "ORDER BY FIN_CheckRegistry.CheckDate ASC"
        Case 1
            OrderBy = "ORDER BY FIN_CheckRegistry.CheckNumber ASC"
        Case 2
            OrderBy = "ORDER BY FIN_CheckRegistry.CheckDate ASC"
        Case 3
            OrderBy = "ORDER BY BASE_Vendor.Name ASC"
        Case 4
            OrderBy = "ORDER BY FIN_CheckRegistry.Amount ASC"
    End Select
   
    
    If chkPayable.value = True Then
        sql = "SELECT FIN_CheckRegistry.CheckNumber,FIN_CheckRegistry.CheckDate,FIN_CheckRegistry.Amount,FIN_CheckRegistry.Remarks," & _
              "FIN_CheckRegistry.DateCleared,FIN_CheckStatus.Status,BASE_Vendor.Name FROM " & _
              "((Peak.dbo.FIN_CheckRegistry FIN_CheckRegistry INNER JOIN Peak.dbo.FIN_CheckStatus FIN_CheckStatus ON " & _
              "FIN_CheckRegistry.CheckStatusId = FIN_CheckStatus.CheckStatusId) " & _
              "LEFT OUTER JOIN Peak.dbo.PO_PurchaseOrder PO_PurchaseOrder ON " & _
              "FIN_CheckRegistry.PurchaseOrderId = PO_PurchaseOrder.PurchaseOrderId) " & _
              "LEFT OUTER JOIN Peak.dbo.BASE_Vendor BASE_Vendor ON " & _
              "PO_PurchaseOrder.VendorId = BASE_Vendor.VendorId " & _
              "WHERE FIN_CheckRegistry.checkDate >= '" & DateFrom.value & " 00:00:00' " & _
              "AND FIN_CheckRegistry.checkDate <= '" & DateTo.value & " 23:23:59' AND isReceivable = 'False' " ' & CheckStatus & Vendor & OrderBy
    Else
        sql = "SELECT FIN_CheckRegistry.CheckNumber,FIN_CheckRegistry.CheckDate,FIN_CheckRegistry.Amount,FIN_CheckRegistry.Remarks," & _
              "FIN_CheckRegistry.DateCleared , FIN_CheckStatus.Status, SO_SalesOrder.OrderNumber, BASE_Vendor.Name " & _
              "From ((Peak.dbo.FIN_CheckRegistry FIN_CheckRegistry INNER JOIN Peak.dbo.FIN_CheckStatus FIN_CheckStatus ON " & _
              "FIN_CheckRegistry.CheckStatusId = FIN_CheckStatus.CheckStatusId) LEFT OUTER JOIN Peak.dbo.SO_SalesOrder SO_SalesOrder ON " & _
              "FIN_CheckRegistry.SalesOrderId = SO_SalesOrder.SalesOrderId) " & _
              "LEFT OUTER JOIN Peak.dbo.BASE_Vendor BASE_Vendor ON " & _
              "SO_SalesOrder.VendorId = BASE_Vendor.VendorId " & _
              "WHERE FIN_CheckRegistry.checkDate >= '" & DateFrom.value & " 00:00:00' " & _
              "AND FIN_CheckRegistry.checkDate <= '" & DateTo.value & " 23:23:59' AND isReceivable = 'True' " & CheckStatus & Vendor & OrderBy
    End If
    
    crxRpt.ParameterFields(1).AddCurrentValue txtTitle.Text
    crxRpt.ParameterFields(2).AddCurrentValue Str(DateFrom.value)
    crxRpt.ParameterFields(3).AddCurrentValue Str(DateTo.value)
    crxRpt.SQLQueryString = sql
    CRViewer.ReportSource = crxRpt
    CRViewer.ViewReport
    CRViewer.Zoom 1
    Screen.MousePointer = vbDefault
End Sub

Private Sub chkPayable_Click()
    If chkPayable.value = True Then
        txtTitle.Text = Me.Caption & " (Payable)"
    Else
        txtTitle.Text = Me.Caption & " (Receivable)"
    End If
End Sub

Private Sub chkReceivable_Click()
    If chkReceivable.value = True Then
        txtTitle.Text = Me.Caption & " (Receivable)"
    Else
        txtTitle.Text = Me.Caption & " (Payable)"
    End If
End Sub

Private Sub CRViewer_PrintButtonClicked(UseDefault As Boolean)
    UseDefault = False
    crxRpt.PrinterSetup Me.hWnd
    crxRpt.PrintOut True
End Sub
Private Sub Form_Load()
    cmbSort.ListIndex = 0
    Populate "Vendor"
    Populate "CheckStatus"
    
    Me.Height = 9390
    Me.width = 15180
    DateFrom.value = Format(Now, "MM/DD/YY")
    DateTo.value = Format(Now, "MM/DD/YY")
    
    txtTitle.Text = Me.Caption
    
    chkPayable_Click
End Sub


Private Sub Form_Resize()
   On Error Resume Next
    CRViewer.width = Me.width - Frame1.width
    CRViewer.Height = Me.Height
    Frame1.Height = Me.Height
End Sub






