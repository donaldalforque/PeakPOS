VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form SYS_ChangeDateTimeFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update system Date and Time"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnSave 
      Caption         =   "Save"
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
      Left            =   2880
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
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
      Left            =   4200
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker TimeFrom 
      Height          =   465
      Left            =   1440
      TabIndex        =   0
      Top             =   840
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   820
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   69926914
      UpDown          =   -1  'True
      CurrentDate     =   42217
   End
   Begin MSComCtl2.DTPicker DateFrom 
      Height          =   465
      Left            =   1440
      TabIndex        =   1
      Top             =   360
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   820
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   69926913
      CurrentDate     =   41686
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Set Time:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   880
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Set Date:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   380
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1455
      Left            =   120
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "SYS_ChangeDateTimeFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

End Sub

Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub btnSave_Click()
    Dim currentDate, currentTime As String
    currentDate = DateFrom.value & " " & TimeFrom.value
    currentTime = Hour(TimeFrom.value) & ":" & Minute(TimeFrom.value) & ":" & Second(TimeFrom.value)
    Date = CDate(currentDate)
    Time = CDate(currentTime)
    
'    Dim lReturn As Long
'    Dim lpSystemTime As SYSTEMTIME
'    lpSystemTime.wHour = Hour(TimeFrom.value)
'    lpSystemTime.wMinute = Minute(TimeFrom.value)
'    lpSystemTime.wSecond = Second(TimeFrom.value)
'    MsgBox Minute(TimeFrom.value)
'    lReturn = SetSystemTime(lpSystemTime)
    
    MsgBox "Date and time successfully updated.", vbInformation, "Date and time updated"
    Unload Me
    Unload SYS_DateTimeCheckerFrm
End Sub

Private Sub Form_Load()
    DateFrom.value = Format(Now, "MM/DD/YY")
    TimeFrom.value = Format("00:00:00", "hh:mm:ss")
End Sub
