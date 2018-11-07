VERSION 5.00
Begin VB.Form POS_UserLoginFrm 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10770
   Icon            =   "POS_UserLoginFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "POS_UserLoginFrm.frx":000C
   ScaleHeight     =   5790
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblShutDown 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   600
      Left            =   5900
      MouseIcon       =   "POS_UserLoginFrm.frx":1C0DA
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   4155
      Width           =   4360
   End
   Begin VB.Label lblLogin 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Height          =   600
      Left            =   540
      MouseIcon       =   "POS_UserLoginFrm.frx":1C22C
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   4155
      Width           =   4360
   End
End
Attribute VB_Name = "POS_UserLoginFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Image2_Click()

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            lblLogin_Click
        Case vbKeyEscape
            lblShutDown_Click
    End Select
End Sub

Private Sub lblLogin_Click()

        'Check if machine is registered in the Server
        Dim ComputerName As String
        ComputerName = Environ("Computername")
        
        Dim con As New ADODB.Connection
        Set cmd = New ADODB.Command
        Set rec = New ADODB.Recordset
        
        con.ConnectionString = ConnString
        con.Open
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "POS_MachineRegistration_Check"
        cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, ComputerName)
        Set rec = cmd.Execute
        If Not rec.EOF Then
            If rec!isActive = "True" Then
                WorkstationId = rec!WorkstationId
                POS_UserPinFrm.Show (1)
                If AllowAccess = True Then
                    AllowAccess = False
                    Unload Me
                    POS_CashierFrm.Show
                End If
            Else
                AllowAccess = False
                GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(58)
                GLOBAL_MessageFrm.Show (1)
            End If
        Else
            AllowAccess = False
            GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(55)
            GLOBAL_MessageFrm.Show (1)
        End If
        con.Close
End Sub

Private Sub lblLogin_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblLogin.BackStyle = 1
End Sub

Private Sub lblLogin_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblLogin.BackStyle = 0
End Sub

Private Sub lblShutDown_Click()
    'clean up
    Dim formcontrol As Form
    For Each formcontrol In Forms
        'Set formcontrol = Nothing
        Unload formcontrol
    Next
End Sub

Private Sub lblShutDown_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblShutDown.BackStyle = 1
End Sub

Private Sub lblShutDown_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblShutDown.BackStyle = 0
End Sub
