VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form POS_ZreadingFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "POS_ZreadingFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOption 
      Alignment       =   1  'Right Justify
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
      TabIndex        =   13
      Top             =   3240
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.ComboBox cmbOption 
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
      ItemData        =   "POS_ZreadingFrm.frx":000C
      Left            =   240
      List            =   "POS_ZreadingFrm.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   2760
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "F2: More Reports"
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
      Picture         =   "POS_ZreadingFrm.frx":0037
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6480
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
      Picture         =   "POS_ZreadingFrm.frx":23C6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7440
      Width           =   3975
   End
   Begin VB.CommandButton btnPrint 
      Caption         =   "F1: Print Z-Reading"
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
      Picture         =   "POS_ZreadingFrm.frx":4755
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5520
      Width           =   3975
   End
   Begin MSComCtl2.DTPicker dtDate 
      Height          =   450
      Left            =   240
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
      Format          =   110755841
      CurrentDate     =   42297
   End
   Begin MSComCtl2.DTPicker startTime 
      Height          =   450
      Left            =   240
      TabIndex        =   5
      Top             =   4080
      Visible         =   0   'False
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
      Format          =   110755842
      CurrentDate     =   42297
   End
   Begin MSComCtl2.DTPicker EndTime 
      Height          =   450
      Left            =   240
      TabIndex        =   7
      Top             =   4920
      Visible         =   0   'False
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
      Format          =   110755842
      CurrentDate     =   42297
   End
   Begin MSComCtl2.DTPicker dtToDate 
      Height          =   450
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
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
      Format          =   110755841
      CurrentDate     =   42297
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
      TabIndex        =   10
      Top             =   1800
      Visible         =   0   'False
      Width           =   870
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
      TabIndex        =   8
      Top             =   4560
      Visible         =   0   'False
      Width           =   975
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
      TabIndex        =   6
      Top             =   3720
      Visible         =   0   'False
      Width           =   1170
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
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1185
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Z Reading Report"
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
      TabIndex        =   3
      Top             =   360
      Width           =   1950
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
      Picture         =   "POS_ZreadingFrm.frx":697D
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   $"POS_ZreadingFrm.frx":93C9
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   600
      TabIndex        =   14
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   8295
      Left            =   120
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "POS_ZreadingFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnPrint_Click()
    Dim crxApp As New CRAXDRT.Application
    Dim crxRpt As New CRAXDRT.Report
    
    Screen.MousePointer = vbHourglass
    Set crxRpt = crxApp.OpenReport(App.path & "\Reports\POS_ZReading.rpt")
    
    Call ResetRptDB(crxRpt)
    
    crxRpt.EnableParameterPrompting = False
    crxRpt.DiscardSavedData
    
    crxRpt.ParameterFields.GetItemByName("@Date").AddCurrentValue dtDate.value
    crxRpt.ParameterFields.GetItemByName("@DateTo").AddCurrentValue dtToDate.value
    If startTime.Visible = True Then
        crxRpt.ParameterFields.GetItemByName("@StartTime").AddCurrentValue Str(startTime.value)
    Else
        crxRpt.ParameterFields.GetItemByName("@StartTime").AddCurrentValue "00:00:00"
    End If
    If EndTime.Visible = True Then
        crxRpt.ParameterFields.GetItemByName("@EndTime").AddCurrentValue Str(EndTime.value)
    Else
        crxRpt.ParameterFields.GetItemByName("@EndTime").AddCurrentValue "23:59:59"
    End If
    
    crxRpt.ParameterFields.GetItemByName("@UserId").AddCurrentValue UserId
    crxRpt.ParameterFields.GetItemByName("@WorkstationId").AddCurrentValue WorkstationId
    crxRpt.ParameterFields.GetItemByName("@Value").AddCurrentValue NVAL(txtOption.Text)
    
    If cmbOption.Visible = True Then
        crxRpt.ParameterFields.GetItemByName("@Mode").AddCurrentValue cmbOption.ListIndex
    Else
        crxRpt.ParameterFields.GetItemByName("@Mode").AddCurrentValue 2
    End If
    
    crxRpt.PrintOut False
    Screen.MousePointer = vbDefault
    
    'POS Audit Trail
    SavePOSAuditTrail VoidUserId, WorkstationId, 0, "Generate Z-Reading Report"
End Sub


Private Sub cmbOption_Click()
    Select Case cmbOption.ListIndex
        Case 0 ' Percentage
            EnableButtons (False)
        Case 1 ' Amount
            EnableButtons (False)
        Case 2 ' Time
            EnableButtons (True)
    End Select
End Sub

Private Sub dtDate_Change()
    dtToDate.value = dtDate.value
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1
            btnPrint_Click
        Case vbKeyEscape
            btnCancel_Click
'        Case vbKeyD
'            If Shift = vbAltMask Then
'                If lblStart.Visible = False Then
'                    ShowHidden (True)
'                Else
'                    ShowHidden (False)
'                End If
'            End If
    End Select
End Sub

Private Sub Form_Load()
    Dim zstartTime As String
    Dim zEndTime As String
    Dim preferredvalue As String
    Dim preferred As String
    
    dtDate.value = Format(Now, "MM/DD/YY")
    dtToDate.value = Format(Now, "MM/DD/YY")

    'Get Time Setup for BIR
    Open App.path & "\Resources\Time.txt" For Input As #1
    Input #1, zstartTime, zEndTime
    Close #1
    
    Open App.path & "\Resources\settings.txt" For Input As #1
    Input #1, preferred
    Input #1, preferred
    Input #1, preferred
    Input #1, preferred
    Input #1, preferred
    Input #1, preferred
    Input #1, preferred 'label
    Input #1, preferred 'label
    Input #1, preferredvalue 'value
    Close #1
    
    
    startTime.value = zstartTime
    EndTime.value = zEndTime
    
    dtDate.Enabled = False
    cmbOption.Text = preferred
    txtOption.Text = preferredvalue
End Sub

Private Sub ShowHidden(ByVal value As Boolean)
    lblStart.Visible = value
    startTime.Visible = value
    dtDate.Enabled = value
    lblEnd.Visible = value
    EndTime.Visible = value
    lblToDate.Visible = value
    dtToDate.Visible = value
    cmbOption.Visible = value
    txtOption.Visible = value
    lblInfo.Visible = Not value
End Sub

Private Sub EnableButtons(ByVal value As Boolean)
    On Error Resume Next
    startTime.Enabled = value
    EndTime.Enabled = value
    txtOption.Enabled = Not value
    If txtOption.Enabled = True Then txtOption.BackColor = vbWhite Else txtOption.BackColor = &H8000000F
    If txtOption.Enabled = True Then txtOption.SetFocus
End Sub

