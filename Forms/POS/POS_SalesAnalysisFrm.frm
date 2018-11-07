VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form POS_SalesAnalysisFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4470
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCancel 
      Caption         =   "ESC:Cancel"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   240
      Picture         =   "POS_SalesAnalysisFrm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3840
      Width           =   3975
   End
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
      Left            =   240
      Picture         =   "POS_SalesAnalysisFrm.frx":238F
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2880
      Width           =   3975
   End
   Begin MSComCtl2.DTPicker dtDate 
      Height          =   450
      Left            =   240
      TabIndex        =   1
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
      Format          =   97189889
      CurrentDate     =   42297
   End
   Begin MSComCtl2.DTPicker dtToDate 
      Height          =   450
      Left            =   240
      TabIndex        =   2
      Top             =   2280
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
      Format          =   97189889
      CurrentDate     =   42297
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DATE"
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
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   525
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Analysis"
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
      Left            =   840
      TabIndex        =   4
      Top             =   360
      Width           =   1590
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   240
      X2              =   4200
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "POS_SalesAnalysisFrm.frx":45B7
      Top             =   240
      Width           =   480
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
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   870
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   4695
      Left            =   120
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "POS_SalesAnalysisFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnPrint_Click()
    '**PRINT RECEIPT******
    Dim crxApp As New CRAXDRT.Application
    Dim crxRpt As New CRAXDRT.Report
    Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\POS_SalesAnalysis.rpt")
    crxRpt.DiscardSavedData
    crxRpt.EnableParameterPrompting = False
    
    crxRpt.ParameterFields.GetItemByName("@DateFrom").AddCurrentValue dtDate.value
    crxRpt.ParameterFields.GetItemByName("@DateTo").AddCurrentValue dtToDate.value
    crxRpt.ParameterFields.GetItemByName("@UserId").AddCurrentValue UserId
    crxRpt.ParameterFields.GetItemByName("@WorkStationId").AddCurrentValue WorkstationId
    
    Call ResetRptDB(crxRpt)
    crxRpt.PrintOut False
    
    '**END PRINT RECEIPT**
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            btnPrint_Click
        Case vbKeyEscape
            btnCancel_Click
    End Select
End Sub

Private Sub Form_Load()
    dtDate.value = Format(Now, "MM/DD/YY")
    dtToDate.value = Format(Now, "MM/DD/YY")
    
    
End Sub
