VERSION 5.00
Begin VB.Form POS_ChangeFundFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCancel 
      Caption         =   "ESC: Cancel"
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
      Left            =   4560
      Picture         =   "POS_ChangeFundFrm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton btnAccept 
      Caption         =   "ENTER: Pay"
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
      Left            =   2880
      Picture         =   "POS_ChangeFundFrm.frx":238F
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1680
      Width           =   1575
   End
   Begin VB.TextBox txtCash 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   1
      Text            =   "0.00"
      Top             =   960
      Width           =   3855
   End
   Begin VB.Label lblCash 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AMOUNT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CHANGE FUND"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   840
      TabIndex        =   0
      Top             =   195
      Width           =   2250
   End
   Begin VB.Image picModuleImage 
      Height          =   480
      Left            =   240
      Picture         =   "POS_ChangeFundFrm.frx":4763
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   0
      Picture         =   "POS_ChangeFundFrm.frx":4D84
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "POS_ChangeFundFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAccept_Click()
    Dim x As Variant
    x = MsgBox("Are you sure all information are correct?", vbQuestion + vbYesNo)
    If x = vbYes Then
        Dim CashFlowId As Long
        Dim con As New ADODB.Connection
        Set rec = New ADODB.Recordset
        Dim cmd As New ADODB.Command
        
        con.ConnectionString = ConnString
        con.Open
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "POS_ChangeFund_Insert"
        cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , NVAL(txtCash.Text))
                              cmd.Parameters("@Amount").NumericScale = 2
                              cmd.Parameters("@Amount").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
        cmd.Parameters.Append cmd.CreateParameter("@WorkStationId", adInteger, adParamInput, , WorkstationId)
        
        cmd.Execute
        
        con.Close
        
        MsgBox "Save successful!", vbInformation
        
        Unload Me
    End If
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            btnAccept_Click
        Case vbKeyEscape
            btnCancel_Click
    End Select
End Sub

Private Sub Form_Load()
    selectText txtCash
End Sub

Private Sub txtCash_Change()
    If IsNumeric(txtCash.Text) = False Then
        txtCash.Text = "0"
        selectText txtCash
    End If
End Sub


