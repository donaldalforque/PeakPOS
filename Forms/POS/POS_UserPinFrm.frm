VERSION 5.00
Begin VB.Form POS_UserPinFrm 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7020
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4320
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "POS_UserPinFrm.frx":0000
   ScaleHeight     =   7020
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPin 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   120
      MaxLength       =   4
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "INPUT USER ID"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label lblCancel 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   2850
      MouseIcon       =   "POS_UserPinFrm.frx":E7A4
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   5660
      Width           =   1350
   End
   Begin VB.Label lbl9 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   2850
      MouseIcon       =   "POS_UserPinFrm.frx":E8F6
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   4445
      Width           =   1350
   End
   Begin VB.Label lbl6 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   2850
      MouseIcon       =   "POS_UserPinFrm.frx":EA48
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   3220
      Width           =   1350
   End
   Begin VB.Label lbl3 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   2850
      MouseIcon       =   "POS_UserPinFrm.frx":EB9A
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   2000
      Width           =   1350
   End
   Begin VB.Label lbl0 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   1480
      MouseIcon       =   "POS_UserPinFrm.frx":ECEC
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   5660
      Width           =   1350
   End
   Begin VB.Label lbl8 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   1480
      MouseIcon       =   "POS_UserPinFrm.frx":EE3E
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   4445
      Width           =   1350
   End
   Begin VB.Label lbl5 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   1480
      MouseIcon       =   "POS_UserPinFrm.frx":EF90
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   3220
      Width           =   1350
   End
   Begin VB.Label lbl2 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   1480
      MouseIcon       =   "POS_UserPinFrm.frx":F0E2
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2000
      Width           =   1350
   End
   Begin VB.Label lblClear 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   120
      MouseIcon       =   "POS_UserPinFrm.frx":F234
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   5660
      Width           =   1350
   End
   Begin VB.Label lbl7 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   120
      MouseIcon       =   "POS_UserPinFrm.frx":F386
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   4440
      Width           =   1350
   End
   Begin VB.Label lbl4 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   120
      MouseIcon       =   "POS_UserPinFrm.frx":F4D8
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3220
      Width           =   1350
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Height          =   1215
      Left            =   120
      MouseIcon       =   "POS_UserPinFrm.frx":F62A
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   1995
      Width           =   1350
   End
End
Attribute VB_Name = "POS_UserPinFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim isUserNumber As Boolean
Dim UserNumber, Pin As Integer
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKey1, vbKeyNumpad1
            lbl1_Click
        Case vbKey2, vbKeyNumpad2
            lbl2_Click
        Case vbKey3, vbKeyNumpad3
            lbl3_Click
        Case vbKey4, vbKeyNumpad4
            lbl4_Click
        Case vbKey5, vbKeyNumpad5
            lbl5_Click
        Case vbKey6, vbKeyNumpad6
            lbl6_Click
        Case vbKey7, vbKeyNumpad7
            lbl7_Click
        Case vbKey8, vbKeyNumpad8
            lbl8_Click
        Case vbKey9, vbKeyNumpad9
            lbl9_Click
        Case vbKey0, vbKeyNumpad0
            lbl0_Click
        Case vbKeyBack
            lblClear_Click
        Case vbKeyEscape
            lblCancel_Click
    End Select
End Sub

Private Sub Form_Load()
    isUserNumber = True
    VoidUserId = 0
End Sub

Private Sub lbl0_Click()
    txtPin.Text = txtPin.Text & "0"
End Sub

Private Sub lbl0_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lbl0.BackStyle = 1
End Sub

Private Sub lbl0_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lbl0.BackStyle = 0
End Sub

Private Sub lbl1_Click()
    txtPin.Text = txtPin.Text & "1"
End Sub

Private Sub lbl1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lbl1.BackStyle = 1
End Sub

Private Sub lbl1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lbl1.BackStyle = 0
End Sub

Private Sub lbl2_Click()
    txtPin.Text = txtPin.Text & "2"
End Sub

Private Sub lbl2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lbl2.BackStyle = 1
End Sub

Private Sub lbl2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lbl2.BackStyle = 0
End Sub

Private Sub lbl3_Click()
    txtPin.Text = txtPin.Text & "3"
End Sub

Private Sub lbl3_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lbl3.BackStyle = 1
End Sub

Private Sub lbl3_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lbl3.BackStyle = 0
End Sub

Private Sub lbl4_Click()
    txtPin.Text = txtPin.Text & "4"
End Sub

Private Sub lbl4_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lbl4.BackStyle = 1
End Sub

Private Sub lbl4_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lbl4.BackStyle = 0
End Sub

Private Sub lbl5_Click()
    txtPin.Text = txtPin.Text & "5"
End Sub

Private Sub lbl5_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lbl5.BackStyle = 1
End Sub

Private Sub lbl5_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lbl5.BackStyle = 0
End Sub

Private Sub lbl6_Click()
    txtPin.Text = txtPin.Text & "6"
End Sub

Private Sub lbl6_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lbl6.BackStyle = 1
End Sub

Private Sub lbl6_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lbl6.BackStyle = 0
End Sub

Private Sub lbl7_Click()
    txtPin.Text = txtPin.Text & "7"
End Sub

Private Sub lbl7_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lbl7.BackStyle = 1
End Sub

Private Sub lbl7_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lbl7.BackStyle = 0
End Sub

Private Sub lbl8_Click()
    txtPin.Text = txtPin.Text & "8"
End Sub

Private Sub lbl8_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lbl8.BackStyle = 1
End Sub

Private Sub lbl8_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lbl8.BackStyle = 0
End Sub

Private Sub lbl9_Click()
    txtPin.Text = txtPin.Text & "9"
End Sub

Private Sub lbl9_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lbl9.BackStyle = 1
End Sub

Private Sub lbl9_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lbl9.BackStyle = 0
End Sub

Private Sub lblCancel_Click()
    AllowAccess = False
    Unload Me
End Sub

Private Sub lblCancel_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblCancel.BackStyle = 1
End Sub

Private Sub lblCancel_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblCancel.BackStyle = 0
End Sub

Private Sub lblClear_Click()
    If txtPin.Text = "" Then Exit Sub
    txtPin.Text = Mid(txtPin.Text, 1, Len(txtPin.Text) - 1)
    
End Sub

Private Sub lblClear_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblClear.BackStyle = 1
End Sub

Private Sub lblClear_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblClear.BackStyle = 0
End Sub

Private Sub txtPin_Change()
    If Len(txtPin.Text) = txtPin.MaxLength Then
        
            Dim con As New ADODB.Connection
            Dim cmd As New ADODB.Command
            Dim rec As New ADODB.Recordset
            
            If isUserNumber = True Then
                'Check User Number
                con.ConnectionString = ConnString
                con.Open
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "User_Pin_Get"
                cmd.Parameters.Append cmd.CreateParameter("@UserNumber", adInteger, adParamInput, , txtPin.Text)
                Set rec = cmd.Execute
                If Not rec.EOF Then
                    If rec!isActive = "True" Then
                        'if void
                        If POS_CashierFrm.Visible = True Then
                            If GetAccessRightsByModule(rec!UserRoleId, 30) = "False" Then 'ModuleID 30 is for VOID rights
                                GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(66)
                                GLOBAL_MessageFrm.Show (1)
                                txtPin.Text = ""
                                AllowAccess = False
                                Exit Sub
                            End If
                        End If
                        VoidUserId = rec!UserId
                        UserNumber = txtPin.Text
                        If IsNull(rec!Pin) Then
                            GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(65)
                            GLOBAL_MessageFrm.Show (1)
                            AllowAccess = False
                            Exit Sub
                        Else
                            Pin = rec!Pin
                        End If
                        If POS_UserLoginFrm.Visible = True Then
                            UserId = rec!UserId
                            CurrentUser = rec!Name
                            POS_CashierFrm.lblCashier.Caption = UCase("|CASHIER: " & CurrentUser)
                            DeleteReserves WorkstationId, 1
                        End If
                        lblTitle.Caption = "INPUT PIN"
                        txtPin.Text = ""
                        isUserNumber = False
                    Else
                        AllowAccess = False
                        txtPin.Text = ""
                        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(56)
                        GLOBAL_MessageFrm.Show (1)
                    End If
                Else
                    AllowAccess = False
                    txtPin.Text = ""
                    GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(56)
                    GLOBAL_MessageFrm.Show (1)
                End If
                con.Close
                
            Else
                'check pin
                If txtPin.Text = Pin Then
                    AllowAccess = True
                    If POS_UserLoginFrm.Visible = True Then
                        'RECORD LOGIN
                        con.ConnectionString = ConnString
                        con.Open
                        Set cmd = New ADODB.Command
                        cmd.ActiveConnection = con
                        cmd.CommandType = adCmdStoredProc
                        cmd.CommandText = "POS_UserAudit_Insert"
                        cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
                        cmd.Parameters.Append cmd.CreateParameter("@WorkstationId", adInteger, adParamInput, , WorkstationId)
                        cmd.Parameters.Append cmd.CreateParameter("@POS_SalesId", adInteger, adParamInput, , Null)
                        cmd.Parameters.Append cmd.CreateParameter("@Activity", adVarChar, adParamInput, 250, "LOGIN")
                        cmd.Parameters.Append cmd.CreateParameter("@Module", adVarChar, adParamInput, 250, "POS")
                        cmd.Execute
                        con.Close
                    End If
                    
                    Unload Me
                Else
                    AllowAccess = False
                    txtPin.Text = ""
                    GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(57)
                    GLOBAL_MessageFrm.Show (1)
                End If
            End If
    End If
End Sub
