VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form POS_EndOfShiftFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "POS_EndOfShiftFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnEnd 
      Caption         =   "F2: End Shift and Log Out"
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
      Picture         =   "POS_EndOfShiftFrm.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6600
      Width           =   3975
   End
   Begin VB.CommandButton btnPrint 
      Caption         =   "F1: Print X-Reading"
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
      Picture         =   "POS_EndOfShiftFrm.frx":0609
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5640
      Width           =   3975
   End
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
      Picture         =   "POS_EndOfShiftFrm.frx":2831
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7560
      Width           =   3975
   End
   Begin VB.ComboBox cmbCashier 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1320
      Width           =   3975
   End
   Begin MSComCtl2.DTPicker dtDate 
      Height          =   450
      Left            =   240
      TabIndex        =   0
      Top             =   2160
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
      Format          =   53673985
      CurrentDate     =   42297
   End
   Begin MSComCtl2.DTPicker startTime 
      Height          =   450
      Left            =   240
      TabIndex        =   8
      Top             =   4080
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
      Format          =   53673986
      CurrentDate     =   42297
   End
   Begin MSComCtl2.DTPicker EndTime 
      Height          =   450
      Left            =   240
      TabIndex        =   9
      Top             =   4920
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
      Format          =   53673986
      CurrentDate     =   42297
   End
   Begin MSComCtl2.DTPicker dtToDate 
      Height          =   450
      Left            =   240
      TabIndex        =   10
      Top             =   3120
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
      Format          =   53673985
      CurrentDate     =   42297
   End
   Begin VB.Label lblStart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "START TIME"
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
      TabIndex        =   13
      Top             =   3720
      Width           =   1170
   End
   Begin VB.Label lblEnd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "END TIME"
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
      TabIndex        =   12
      Top             =   4560
      Width           =   975
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
      TabIndex        =   11
      Top             =   2760
      Width           =   870
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "POS_EndOfShiftFrm.frx":4BC0
      Top             =   240
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   240
      X2              =   4200
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "End of Shift"
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
      TabIndex        =   7
      Top             =   360
      Width           =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CASHIER"
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
      TabIndex        =   6
      Top             =   960
      Width           =   840
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
      TabIndex        =   2
      Top             =   1800
      Width           =   525
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   8415
      Left            =   120
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "POS_EndOfShiftFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub Populate(ByVal data As String)
    Select Case data
        Case "User"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("User")
            cmbCashier.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then  'Cashier
                        cmbCashier.AddItem rec!Name
                        cmbCashier.ItemData(cmbCashier.NewIndex) = rec!UserId
                    End If
                    rec.MoveNext
                Loop
            End If
            On Error Resume Next
            cmbCashier.ListIndex = 0
    End Select
End Sub
Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnEnd_Click()
    'Confirm End of Shift
    Dim x As Variant
    x = MsgBox("Are you sure you want to end your shift and log out?", vbQuestion + vbYesNo, "")
    If x = vbYes Then
        Set con = New ADODB.Connection
        Set cmd = New ADODB.Command
        
        con.ConnectionString = ConnString
        con.Open
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "POS_UserAudit_Insert"
        cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
        cmd.Parameters.Append cmd.CreateParameter("@WorkstationId", adInteger, adParamInput, , WorkstationId)
        cmd.Parameters.Append cmd.CreateParameter("@POS_SalesId", adInteger, adParamInput, , 0)
        cmd.Parameters.Append cmd.CreateParameter("@Activity", adVarChar, adParamInput, 250, "END OF SHIFT")
        cmd.Parameters.Append cmd.CreateParameter("@Module", adVarChar, adParamInput, 250, "POS")
        cmd.Execute
        con.Close
        Unload Me
        Unload POS_CashierFrm
        POS_UserLoginFrm.Show
    End If
End Sub

Private Sub btnPrint_Click()
    Dim crxApp As New CRAXDRT.Application
    Dim crxRpt As New CRAXDRT.Report
    
    Screen.MousePointer = vbHourglass
    Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\POS_XReading.rpt")
    crxRpt.EnableParameterPrompting = False
    crxRpt.DiscardSavedData
    Call ResetRptDB(crxRpt)
    
    'crxRpt.ParameterFields.GetItemByName("ReportTitle").AddCurrentValue txtTitle.text
    'crxRpt.ParameterFields.GetItemByName("DateFrom").AddCurrentValue DateFrom.value & " " & TimeFrom.value
    'crxRpt.ParameterFields.GetItemByName("DateTo").AddCurrentValue DateTo.value & " " & TimeTo.value
    
    crxRpt.ParameterFields.GetItemByName("@Date").AddCurrentValue dtDate.value
    'crxRpt.ParameterFields.GetItemByName("@DateTo").AddCurrentValue dtToDate.value
    'crxRpt.ParameterFields.GetItemByName("@StartTime").AddCurrentValue Str(startTime.value)
    'crxRpt.ParameterFields.GetItemByName("@EndTime").AddCurrentValue Str(EndTime.value)
    crxRpt.ParameterFields.GetItemByName("@UserId").AddCurrentValue cmbCashier.ItemData(cmbCashier.ListIndex)
    crxRpt.ParameterFields.GetItemByName("@WorkstationId").AddCurrentValue WorkstationId
    
    crxRpt.PrintOut False
    Screen.MousePointer = vbDefault
    
    'POS Audit Trail
    SavePOSAuditTrail VoidUserId, WorkstationId, 0, "Generate X-Reading Report"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            btnPrint_Click
        Case vbKeyF2
            btnEnd_Click
        Case vbKeyEscape
            btnCancel_Click
    End Select
End Sub

Private Sub Form_Load()
    Populate "User"
    Dim x As Integer
    For x = 0 To cmbCashier.ListCount - 1
        If cmbCashier.ItemData(x) = UserId Then cmbCashier.ListIndex = x
    Next x
    dtDate.value = Format(Now, "MM/DD/YY")
    
    Dim zstartTime As String
    Dim zEndTime As String
    dtDate.value = Format(Now, "MM/DD/YY")
    dtToDate.value = Format(Now, "MM/DD/YY")

    'Get Time Setup for BIR
    Open App.Path & "\Resources\Time.txt" For Input As #1
    Input #1, zstartTime, zEndTime
    Close #1
    
    startTime.value = zstartTime
    EndTime.value = zEndTime
End Sub
