VERSION 5.00
Begin VB.Form POS_SalesClerkNumberFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sales Clerk Number"
   ClientHeight    =   735
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2640
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   2640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnAccept 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtUserNumber 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      IMEMode         =   3  'DISABLE
      Left            =   720
      MaxLength       =   4
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   140
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "POS_SalesClerkNumberFrm.frx":0000
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "POS_SalesClerkNumberFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAccept_Click()
    If IsNumeric(txtUserNumber.text) = False Then
        MsgBox "Invalid user number. Please try again", vbCritical, "Invalid number"
        txtUserNumber.text = ""
        Exit Sub
    End If
    
    Dim con As New ADODB.Connection
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_User_GetByNumber"
    cmd.Parameters.Append cmd.CreateParameter("@UserNumber", adInteger, adParamInput, 10, txtUserNumber.text)
    Set rec = cmd.Execute
    If Not rec.EOF Then
        POS_CashierFrm.CurrentUserId = rec!UserId
        POS_CashierFrm.lblCashier.Caption = UCase("|SALES CLERK: " & rec!Name)
        Unload Me
        POS_ItemSearchFrm.Show (1)
    Else
        MsgBox "Invalid user number. Please try again", vbCritical, "Invalid number"
        txtUserNumber.text = ""
    End If
    con.Close
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            btnAccept_Click
    End Select
End Sub

Private Sub Form_Load()
    If POS_UserLoginFrm.Visible = True Then Unload Me
End Sub

