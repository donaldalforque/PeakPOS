VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form POS_PrintHistoryHoldOrderListFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnPrint 
      Caption         =   "F1: Print"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Picture         =   "POS_PrintHistoryHoldOrderListFrm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   3975
   End
   Begin MSComCtl2.DTPicker dtDate 
      Height          =   450
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   794
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   97452033
      CurrentDate     =   42297
   End
   Begin MSComCtl2.DTPicker dtToDate 
      Height          =   450
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   794
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   97452033
      CurrentDate     =   42297
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FROM DATE"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1185
   End
   Begin VB.Label lblToDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TO DATE"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   870
   End
End
Attribute VB_Name = "POS_PrintHistoryHoldOrderListFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnPrint_Click()
    AllowAccess = False
    POS_UserPinFrm.Show (1)
    If AllowAccess = False Then Exit Sub

    Dim crxApp As New CRAXDRT.Application
    Dim crxRpt As New CRAXDRT.Report
    
    Screen.MousePointer = vbHourglass
    Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\POS_HoldOrderListHistory.rpt")
    crxRpt.EnableParameterPrompting = False
    crxRpt.DiscardSavedData
    Call ResetRptDB(crxRpt)
    crxRpt.ParameterFields.GetItemByName("@DateFrom").AddCurrentValue dtDate.value
    crxRpt.ParameterFields.GetItemByName("@DateTo").AddCurrentValue dtToDate.value
    
    crxRpt.PrintOut False
    Screen.MousePointer = vbDefault
    
    'POS Audit Trail
    SavePOSAuditTrail VoidUserId, WorkstationId, 0, "Print Hold order history."
End Sub

Private Sub Form_Load()
    dtDate.value = Format(Now, "MM/DD/YY")
    dtToDate.value = Format(Now, "MM/DD/YY")
End Sub
