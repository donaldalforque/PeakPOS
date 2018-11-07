VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form BASE_LoyaltyCardSettingsFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtname 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   7320
      Visible         =   0   'False
      Width           =   8415
   End
   Begin VB.TextBox txtPurchaseAmount 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3480
      TabIndex        =   0
      Text            =   "1"
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox txtMinPoints 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3480
      TabIndex        =   1
      Text            =   "1"
      Top             =   2280
      Width           =   2175
   End
   Begin MSComctlLib.Toolbar tb_Standard 
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   582
      ButtonWidth     =   2805
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save && Close"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Accounts"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8280
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BASE_LoyaltyCardSettingsFrm.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BASE_LoyaltyCardSettingsFrm.frx":6862
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BASE_LoyaltyCardSettingsFrm.frx":D0C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BASE_LoyaltyCardSettingsFrm.frx":13926
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "BASE_LoyaltyCardSettingsFrm.frx":1A188
      Top             =   720
      Width           =   480
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "This allows you to setup customer loyalty preferences such as purchase points equivalent."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   450
      Left            =   240
      TabIndex        =   7
      Top             =   1320
      Width           =   5535
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Loyalty Program"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Left            =   840
      TabIndex        =   6
      Top             =   840
      Width           =   3075
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "1 Point Purchase Amount:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   5
      Top             =   1920
      Width           =   2115
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Minimum Points to enable Redeem:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   4
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   2295
      Left            =   120
      Top             =   480
      Width           =   5895
   End
End
Attribute VB_Name = "BASE_LoyaltyCardSettingsFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim con As New ADODB.Connection
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "POS_Settings_Get"
    Set rec = cmd.Execute
    If Not rec.EOF Then
        txtPurchaseAmount.Text = rec!LoyaltyPointsDiv
        txtMinPoints.Text = rec!MinPointsRedeem
    End If
    con.Close
End Sub

Private Sub tb_Standard_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 2 'save
            Dim LoyaltyPointsDiv, MinPointsRedeem As Double
            
            If Val(txtPurchaseAmount.Text) = 0 Then LoyaltyPointsDiv = 200 Else LoyaltyPointsDiv = Val(txtPurchaseAmount.Text)
            If Val(txtMinPoints.Text) = 0 Then MinPointsRedeem = 100 Else MinPointsRedeem = Val(txtMinPoints.Text)
        
            Dim con As New ADODB.Connection
            Set cmd = New ADODB.Command
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "POS_Settings_Update"
            cmd.Parameters.Append cmd.CreateParameter("@LoyaltyPointsDiv", adDecimal, adParamInput, , LoyaltyPointsDiv)
                                  cmd.Parameters("@LoyaltyPointsDiv").NumericScale = 2
                                  cmd.Parameters("@LoyaltyPointsDiv").Precision = 18
            cmd.Parameters.Append cmd.CreateParameter("@MinPointsRedeem", adDecimal, adParamInput, , MinPointsRedeem)
                                  cmd.Parameters("@MinPointsRedeem").NumericScale = 2
                                  cmd.Parameters("@MinPointsRedeem").Precision = 18
            cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
            cmd.Parameters.Append cmd.CreateParameter("@WorkStationId", adInteger, adParamInput, , WorkstationId)
            cmd.Execute
            con.Close
            Unload Me
    End Select
End Sub

Private Sub txtMinPoints_Change()
    If IsNumeric(txtMinPoints.Text) = False And Trim(txtMinPoints.Text) <> "" Then
        txtMinPoints.Text = "100"
    End If
End Sub

Private Sub txtPurchaseAmount_Change()
    If IsNumeric(txtPurchaseAmount.Text) = False And Trim(txtPurchaseAmount.Text) <> "" Then
        txtPurchaseAmount.Text = "200"
    End If
End Sub
