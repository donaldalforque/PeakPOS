VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form RPT_PO_SupplierLedgerFrm 
   Caption         =   "Supplier Ledger"
   ClientHeight    =   9045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15180
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9045
   ScaleWidth      =   15180
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.TextBox txtCode 
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
         TabIndex        =   4
         Top             =   3480
         Visible         =   0   'False
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
         Top             =   2760
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
         ItemData        =   "RPT_PO_SupplierLedgerFrm.frx":0000
         Left            =   1320
         List            =   "RPT_PO_SupplierLedgerFrm.frx":000A
         TabIndex        =   2
         Text            =   "cmbCustomer"
         Top             =   600
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
         TabIndex        =   1
         Top             =   3960
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DatePeriod 
         Height          =   345
         Left            =   1320
         TabIndex        =   5
         Top             =   960
         Visible         =   0   'False
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
         CustomFormat    =   "MMMM"
         Format          =   142868483
         CurrentDate     =   41686
      End
      Begin MSComCtl2.DTPicker DateTo 
         Height          =   345
         Left            =   1320
         TabIndex        =   6
         Top             =   1800
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
         Format          =   142868481
         CurrentDate     =   41686
      End
      Begin MSComCtl2.DTPicker DateFrom 
         Height          =   345
         Left            =   1320
         TabIndex        =   7
         Top             =   1440
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
         Format          =   142868481
         CurrentDate     =   41686
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
         TabIndex        =   14
         Top             =   1800
         Width           =   705
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
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   960
      End
      Begin VB.Label Label3 
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
         Left            =   120
         TabIndex        =   12
         Top             =   3480
         Visible         =   0   'False
         Width           =   480
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
         TabIndex        =   11
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Month"
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
         Top             =   960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label8 
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
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   120
         Width           =   1005
      End
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer 
      Height          =   9015
      Left            =   3840
      TabIndex        =   15
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
Attribute VB_Name = "RPT_PO_SupplierLedgerFrm"
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
    End Select
End Sub

Private Sub btnGenerate_Click()
    If cmbVendor.ListIndex = -1 Then
        MsgBox "Please select a valid Vendor.", vbExclamation, "Required."
        Exit Sub
    End If
    
    Dim sql, OrderBy As String
    Dim Status, Vendor, Code As Variant
    Dim VendorId As Integer
    
    
    Screen.MousePointer = vbHourglass
    
    Set crxRpt = crxApp.OpenReport(App.path & "\Reports\PO_VendorLedger.rpt")
    
    crxRpt.EnableParameterPrompting = False
    crxRpt.DiscardSavedData
    Call ResetRptDB(crxRpt)
          
    crxRpt.ParameterFields.GetItemByName("ReportTitle").AddCurrentValue txtTitle.Text
    crxRpt.ParameterFields.GetItemByName("@VendorId").AddCurrentValue cmbVendor.ItemData(cmbVendor.ListIndex)
    crxRpt.ParameterFields.GetItemByName("@DateFrom").AddCurrentValue DateFrom.value
    crxRpt.ParameterFields.GetItemByName("@DateTo").AddCurrentValue DateTo.value
    'crxRpt.ParameterFields.GetItemByName("@TemplateId").AddCurrentValue StatementTemplateId
    

    CRViewer.ReportSource = crxRpt
    CRViewer.ViewReport
    CRViewer.Zoom 1
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
'    cmbSort.ListIndex = 0
'    Populate "Status"
    Populate "Vendor"
    
    Me.Height = 9390
    Me.width = 15180
    DatePeriod.value = Now
    txtTitle.Text = Me.Caption
    DateFrom.value = Format(Now, "MM/DD/YY")
    DateTo.value = Format(Now, "MM/DD/YY")
End Sub
Private Sub CRViewer_PrintButtonClicked(UseDefault As Boolean)
    UseDefault = False
    crxRpt.PrinterSetup Me.hWnd
    crxRpt.PrintOut True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    CRViewer.width = Me.width - Frame1.width
    CRViewer.Height = Me.Height
    Frame1.Height = Me.Height
End Sub








