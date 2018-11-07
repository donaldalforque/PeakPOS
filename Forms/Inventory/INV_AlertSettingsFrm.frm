VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form INV_AlertSettingsFrm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Product Expiry Alert"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   14
      Top             =   4200
      Width           =   5535
      Begin VB.TextBox txtDays 
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
         Left            =   2280
         TabIndex        =   17
         Text            =   "30"
         Top             =   720
         Width           =   3015
      End
      Begin VB.CheckBox chkEveryLogout 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Every Log-Out"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   16
         Top             =   360
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox chkEveryLogin 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Every Log-In"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   15
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Alert Schedule:"
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
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Will Expire within (Days):"
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
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   2100
      End
   End
   Begin MSComCtl2.DTPicker TimeFrom 
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   3120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   95813634
      UpDown          =   -1  'True
      CurrentDate     =   42217
   End
   Begin VB.TextBox txtInterval 
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
      Left            =   2640
      TabIndex        =   3
      Text            =   "1"
      Top             =   2760
      Width           =   3015
   End
   Begin VB.ComboBox cmbAllow 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "INV_AlertSettingsFrm.frx":0000
      Left            =   2640
      List            =   "INV_AlertSettingsFrm.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3570
      Width           =   3015
   End
   Begin VB.TextBox txtFrequency 
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
      Left            =   2640
      TabIndex        =   2
      Text            =   "1"
      Top             =   2400
      Width           =   3015
   End
   Begin VB.ComboBox cmbSchedule 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "INV_AlertSettingsFrm.frx":0017
      Left            =   2640
      List            =   "INV_AlertSettingsFrm.frx":0027
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2040
      Width           =   3015
   End
   Begin VB.TextBox txtname 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   7320
      Visible         =   0   'False
      Width           =   8415
   End
   Begin MSComctlLib.Toolbar tb_Standard 
      Height          =   330
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   582
      ButtonWidth     =   1349
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
            Caption         =   "Save"
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
            Picture         =   "INV_AlertSettingsFrm.frx":0049
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_AlertSettingsFrm.frx":68AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_AlertSettingsFrm.frx":D10D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_AlertSettingsFrm.frx":1396F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   5880
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check Start Time:"
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
      TabIndex        =   13
      Top             =   3150
      Width           =   1395
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frequency Interval (min):"
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
      TabIndex        =   12
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Allow Negative Inventory: "
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
      TabIndex        =   11
      Top             =   3600
      Width           =   2130
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check Frequency:"
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
      TabIndex        =   10
      Top             =   2400
      Width           =   1410
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Alert Schedule: "
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
      TabIndex        =   9
      Top             =   2040
      Width           =   1260
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "INV_AlertSettingsFrm.frx":1A1D1
      Top             =   680
      Width           =   480
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inventory Alerts"
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
      TabIndex        =   8
      Top             =   840
      Width           =   1875
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "You can setup how often the system will notify for products that are low in quantity and check for negative inventory."
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
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   5055
      Left            =   120
      Top             =   480
      Width           =   5895
   End
End
Attribute VB_Name = "INV_AlertSettingsFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    cmbSchedule.ListIndex = 0
    cmbAllow.ListIndex = 0
    
    Set rec = New ADODB.Recordset
    Set rec = Global_Data("InventorySettings")
    
    If Not rec.EOF Then
        cmbSchedule.text = rec!ReorderPointCheckScheduled
        txtFrequency.text = FormatNumber(rec!ReorderPointCheckFrequency)
        txtInterval.text = FormatNumber(rec!ReorderPointCheckFrequencyInterval)
        TimeFrom.value = rec!ReorderPointCheckStartDate
        If rec!AllowNegativeInventory = "False" Then
            cmbAllow.text = "NO"
        Else
            cmbAllow.text = "YES"
        End If
        If rec!EveryLogIn = "True" Then
            chkEveryLogin.value = Checked
        Else
            chkEveryLogin.value = Unchecked
        End If
        If rec!EveryLogOut = "True" Then
            chkEveryLogout.value = Checked
        Else
            chkEveryLogout.value = Unchecked
        End If
        txtDays.text = FormatNumber(rec!ExpiryDays, 2, vbTrue, vbFalse)
    End If
End Sub

Private Sub tb_Standard_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim con As New ADODB.Connection
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    Dim allow, EveryLogIn, EveryLogOut As Boolean
    
    EveryLogIn = False
    EveryLogOut = False
    
    If cmbAllow.ListIndex = 0 Then
        allow = True
    Else
        allow = False
    End If
    
    If chkEveryLogin.value = Checked Then EveryLogIn = True
    If chkEveryLogout.value = Checked Then EveryLogOut = True
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "INV_Settings_Update"
    cmd.Parameters.Append cmd.CreateParameter("@ReorderPointCheckScheduled", adVarChar, adParamInput, 50, cmbSchedule.text)
    cmd.Parameters.Append cmd.CreateParameter("@ReorderPointCheckFrequency", adDecimal, adParamInput, , NVAL(txtFrequency.text))
                          cmd.Parameters("@ReorderPointCheckFrequency").NumericScale = 2
                          cmd.Parameters("@ReorderPointCheckFrequency").Precision = 18
    cmd.Parameters.Append cmd.CreateParameter("@ReorderPointCheckFrequencyInterval", adDecimal, adParamInput, , NVAL(txtInterval.text))
                          cmd.Parameters("@ReorderPointCheckFrequencyInterval").NumericScale = 2
                          cmd.Parameters("@ReorderPointCheckFrequencyInterval").Precision = 18
    cmd.Parameters.Append cmd.CreateParameter("@ReorderPointCheckStartDate", adDate, adParamInput, , TimeFrom.value)
    cmd.Parameters.Append cmd.CreateParameter("@AllowNegativeInventory", adBoolean, adParamInput, , allow)
    cmd.Parameters.Append cmd.CreateParameter("@EveryLogin", adBoolean, adParamInput, , EveryLogIn)
    cmd.Parameters.Append cmd.CreateParameter("@EveryLogout", adBoolean, adParamInput, , EveryLogOut)
    cmd.Parameters.Append cmd.CreateParameter("@ExpiryDays", adDecimal, adParamInput, , NVAL(txtDays.text))
                          cmd.Parameters("@ExpiryDays").NumericScale = 2
                          cmd.Parameters("@ExpiryDays").Precision = 18
    cmd.Execute
    con.Close
    
    GetInventorySettings 'Update settings variables
    
    MsgBox "Settings updated.", vbInformation
End Sub

Private Sub txtFrequency_Change()
    If IsNumeric(txtFrequency.text) = False Then
        txtFrequency.text = "1"
    End If
End Sub

Private Sub txtInterval_Change()
    If IsNumeric(txtInterval.text) = False Then
        txtInterval.text = "1"
    End If
End Sub
