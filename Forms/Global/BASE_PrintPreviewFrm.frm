VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form BASE_PrintPreviewFrm 
   Caption         =   "Print Preview"
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9975
   Icon            =   "BASE_PrintPreviewFrm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9030
   ScaleWidth      =   9975
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar tb_Standard 
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   582
      ButtonWidth     =   2117
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Paper Size"
            ImageIndex      =   5
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "onehalf"
                  Text            =   "5.5 x 8.5 (1/2 LW)"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "onefort"
                  Text            =   "4.25 x 5.5 (1/4)"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9000
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BASE_PrintPreviewFrm.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BASE_PrintPreviewFrm.frx":D0B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BASE_PrintPreviewFrm.frx":13916
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BASE_PrintPreviewFrm.frx":1A178
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BASE_PrintPreviewFrm.frx":209DA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer 
      Height          =   8655
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   9975
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
Attribute VB_Name = "BASE_PrintPreviewFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public isInvoice As Boolean

Private Sub CRViewer_PrintButtonClicked(UseDefault As Boolean)
'    UseDefault = False
'    crxRpt.PrinterSetup Me.hWnd
'    crxRpt.PrintOut True
End Sub

Private Sub Form_Resize()
    CRViewer.Height = Me.Height - 405
    CRViewer.width = Me.width - 120
End Sub

Private Sub tb_Standard_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "onehalf"
            Screen.MousePointer = vbHourglass
            BASE_PrintPreviewFrm.Show
            Dim crxApp As New CRAXDRT.Application
            Dim crxRpt As New CRAXDRT.Report
            If isInvoice = False Then
                Set crxRpt = crxApp.OpenReport(App.path & "\Reports\SO_SalesOrder.rpt")
                crxRpt.RecordSelectionFormula = "{SO_SalesOrder.SalesOrderId}= " & SO_SalesOrderFrm.SalesOrderId & ""
            Else
                Set crxRpt = crxApp.OpenReport(App.path & "\Reports\SO_SalesInvoice.rpt")
                crxRpt.RecordSelectionFormula = "{SO_Invoice.InvoiceId}= " & Val(SO_InvoiceFrm.InvoiceId) & ""
            End If
            crxRpt.DiscardSavedData

            Call ResetRptDB(crxRpt)

            BASE_PrintPreviewFrm.CRViewer.ReportSource = crxRpt
            BASE_PrintPreviewFrm.CRViewer.ViewReport
            BASE_PrintPreviewFrm.CRViewer.Zoom 1
            Screen.MousePointer = vbDefault
        Case "onefort"
            Screen.MousePointer = vbHourglass
                BASE_PrintPreviewFrm.Show
                If isInvoice = False Then
                    Set crxRpt = crxApp.OpenReport(App.path & "\Reports\SO_SalesOrder_.25.rpt")
                    crxRpt.RecordSelectionFormula = "{SO_SalesOrder.SalesOrderId}= " & SO_SalesOrderFrm.SalesOrderId & ""
                Else
                    Set crxRpt = crxApp.OpenReport(App.path & "\Reports\SO_SalesInvoice_.25.rpt")
                    crxRpt.RecordSelectionFormula = "{SO_Invoice.InvoiceId}= " & Val(SO_InvoiceFrm.InvoiceId) & ""
                End If
                crxRpt.DiscardSavedData

                Call ResetRptDB(crxRpt)

                BASE_PrintPreviewFrm.CRViewer.ReportSource = crxRpt
                BASE_PrintPreviewFrm.CRViewer.ViewReport
                BASE_PrintPreviewFrm.CRViewer.Zoom 1
                Screen.MousePointer = vbDefault
    End Select
End Sub
