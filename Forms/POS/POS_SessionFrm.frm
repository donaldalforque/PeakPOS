VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form POS_SessionFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8175
   Icon            =   "POS_SessionFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   8175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkMin 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Minutes"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   25
      Top             =   4440
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   7560
      Top             =   960
   End
   Begin VB.CommandButton btnCheckOut 
      Caption         =   "Check - Out"
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
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6000
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtAdmissionNumber 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   1800
      Width           =   4695
   End
   Begin VB.CheckBox chkUnli 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Unlimited"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   20
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox txtHours 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6600
      TabIndex        =   5
      Text            =   "0"
      Top             =   4440
      Width           =   1335
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
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6000
      Width           =   1575
   End
   Begin VB.TextBox lblTimeLeft 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Height          =   465
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "TIME LEFT: 1 Hour"
      Top             =   525
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.CommandButton btnItemSearch 
      Height          =   450
      Left            =   7440
      Picture         =   "POS_SessionFrm.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2280
      Width           =   495
   End
   Begin VB.TextBox txtStatus 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Active"
      Top             =   1320
      Width           =   4695
   End
   Begin VB.TextBox txtAttendant 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3240
      TabIndex        =   7
      Top             =   5400
      Width           =   4695
   End
   Begin VB.TextBox txtChild 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3240
      TabIndex        =   3
      Top             =   3240
      Width           =   4695
   End
   Begin VB.TextBox txtCustomer 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3240
      TabIndex        =   2
      Top             =   2760
      Width           =   4695
   End
   Begin VB.TextBox txtOR 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3240
      MaxLength       =   50
      TabIndex        =   1
      Top             =   2280
      Width           =   4215
   End
   Begin MSComCtl2.DTPicker DateFrom 
      Height          =   435
      Left            =   3240
      TabIndex        =   4
      Top             =   3960
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   767
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   75366402
      CurrentDate     =   41686
   End
   Begin MSComCtl2.DTPicker dateTo 
      Height          =   435
      Left            =   3240
      TabIndex        =   6
      Top             =   4920
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   767
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   75366402
      CurrentDate     =   41686
   End
   Begin MSComctlLib.Toolbar tb_Standard 
      Height          =   330
      Left            =   0
      TabIndex        =   23
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
         NumButtons      =   6
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
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Extend"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Print"
            ImageIndex      =   4
            Style           =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   5760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "POS_SessionFrm.frx":0230
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "POS_SessionFrm.frx":6A92
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "POS_SessionFrm.frx":D2F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "POS_SessionFrm.frx":13B56
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "POS_SessionFrm.frx":13DCB
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "POS_SessionFrm.frx":1443C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "POS_SessionFrm.frx":1AC9E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Admission #:"
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
      Left            =   360
      TabIndex        =   22
      Top             =   1800
      Width           =   1470
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hours:"
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
      Left            =   360
      TabIndex        =   19
      Top             =   4440
      Width           =   780
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
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
      Left            =   360
      TabIndex        =   14
      Top             =   1320
      Width           =   795
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Attendant:"
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
      Left            =   360
      TabIndex        =   13
      Top             =   5400
      Width           =   1230
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "End Time:"
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
      Left            =   360
      TabIndex        =   12
      Top             =   4920
      Width           =   1125
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Start Time:"
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
      Left            =   360
      TabIndex        =   11
      Top             =   3960
      Width           =   1245
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Child:"
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
      Left            =   360
      TabIndex        =   10
      Top             =   3240
      Width           =   645
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cust./Parent/Guardian:"
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
      Left            =   360
      TabIndex        =   9
      Top             =   2760
      Width           =   2700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "POS OR #:"
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
      Left            =   360
      TabIndex        =   8
      Top             =   2280
      Width           =   1155
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "POS_SessionFrm.frx":21500
      Top             =   525
      Width           =   480
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Admission Slip"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   390
      Left            =   960
      TabIndex        =   0
      Top             =   570
      Width           =   1860
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   6255
      Left            =   120
      Top             =   360
      Width           =   7935
   End
End
Attribute VB_Name = "POS_SessionFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim POS_SalesId As String
Dim CustomerId As String
Dim SlotId As Integer
Dim RoomId As Integer
Public POS_SessionId As String
Dim isExtend As Boolean

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnCheckOut_Click()
    Dim x As Variant
    x = MsgBox("Are you sure you want to check out this customer?", vbQuestion + vbYesNo)
    If x = vbYes Then
        Set con = New ADODB.Connection
        Set cmd = New ADODB.Command
        
        con.ConnectionString = ConnString
        con.Open
        cmd.ActiveConnection = con
        cmd.CommandText = "POS_Session_StatusUpdate"
        cmd.CommandType = adCmdStoredProc
        cmd.Parameters.Append cmd.CreateParameter("@POS_SessionId", adInteger, adParamInput, , Val(POS_SessionId))
        cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , 4) 'CHECK OUT
        cmd.Execute
        con.Close
        
        MsgBox "Customer checked out.", vbInformation
       
        Unload Me
        Unload POS_PlayHouseFrm
    End If
End Sub

Private Sub btnItemSearch_Click()
    Dim isValid As Boolean
    
    'CHECK IF VALID OR #
    Set con = New ADODB.Connection
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "POS_Session_ORCheck"
    cmd.Parameters.Append cmd.CreateParameter("@OrNumber", adVarChar, adParamInput, 50, txtOR.text)
    Set rec = cmd.Execute
    If Not rec.EOF Then
        If rec!result = "False" Then
            GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(68)
            GLOBAL_MessageFrm.Show (1)
            
            txtCustomer.text = ""
            DateFrom.value = Now
            dateTo.value = Now
            txtChild.text = ""
            txtAttendant.text = ""
            txtHours.text = "0"
            
            txtHours.Enabled = False
            txtCustomer.Enabled = False
            DateFrom.Enabled = False
            dateTo.Enabled = False
            txtChild.Enabled = False
            txtAttendant.Enabled = False
            
            On Error Resume Next
            txtOR.SetFocus
        Else
            txtCustomer.text = rec!Name
            DateFrom.value = rec!Date
            dateTo.value = rec!Date
            POS_SalesId = rec!POS_SalesId
            CustomerId = rec!CustomerId
            
            'enable
            'txtCustomer.Enabled = True
            'DateFrom.Enabled = True
            txtHours.Enabled = True
            dateTo.Enabled = True
            txtChild.Enabled = True
            txtAttendant.Enabled = True
        End If
    End If
    con.Close
End Sub

Private Sub btnSave_Click()
    
End Sub

Private Sub chkMin_Click()
    txtHours_Change
End Sub

Private Sub chkUnli_Click()
    If chkUnli.value = vbChecked Then
        txtHours.Enabled = False
        txtHours.text = "8"
    Else
        txtHours.Enabled = True
    End If
End Sub
Public Sub LoadVar(ByVal cRoomId As Integer, cSlotId As Integer)
    RoomId = cRoomId
    SlotId = cSlotId
End Sub

Private Sub Form_Load()
    'CHECK IF HAS DATA
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "POS_Session_Get"
    cmd.Parameters.Append cmd.CreateParameter("@POS_SessionId", adInteger, adParamInput, , Val(POS_SessionId))
    Set rec = cmd.Execute
    If Not rec.EOF Then
          'Disable and Show Controls
            btnCheckOut.Visible = True
            btnItemSearch.Enabled = False
            txtOR.Enabled = False
            txtChild.Enabled = False
            
            txtHours.Enabled = True
            txtAttendant.Enabled = True
            txtChild.Enabled = True
            dateTo.Enabled = True
            
            txtOR.text = rec!pos_ordernumber
            txtCustomer.text = rec!Name
            txtChild.text = rec!childname
            DateFrom.value = rec!startdate
            dateTo.value = rec!enddate
            txtStatus.text = rec!Status
            txtAdmissionNumber.text = rec!POS_SessionId
            txtHours.text = rec!hours
            If rec!isunlimited = "True" Then
                chkUnli.value = vbChecked
            Else
                chkUnli.value = vbUnchecked
            End If
            If rec!isMinutes = "True" Then
                chkMin.value = vbChecked
            Else
                chkMin.value = vbUnchecked
            End If
            txtAttendant.text = rec!attendant
            POS_SalesId = rec!POS_SalesId
            SlotId = rec!SlotId
            RoomId = rec!RoomId
            CustomerId = rec!CustomerId
            
            'show extend button
            tb_Standard.Buttons(4).Visible = True
            tb_Standard.Buttons(5).Visible = True
    End If
    con.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
    POS_SessionId = "0"
    POS_SalesId = "0"
    SlotId = 0
    RoomId = 0
    CustomerId = "0"
    tb_Standard.Buttons(4).Visible = False
    tb_Standard.Buttons(5).Visible = False
End Sub

Private Sub tb_Standard_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 2 ' SAVE
            If Trim(txtOR.text) = "" Then
                On Error Resume Next
                GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(68)
                GLOBAL_MessageFrm.Show (1)
                txtOR.SetFocus
            ElseIf Trim(txtCustomer.text) = "" Then
                On Error Resume Next
                GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(13)
                GLOBAL_MessageFrm.Show (1)
                txtCustomer.SetFocus
            ElseIf Trim(txtChild.text) = "" Then
                On Error Resume Next
                GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(69)
                GLOBAL_MessageFrm.Show (1)
                txtChild.SetFocus
            ElseIf Val(txtHours.text) = 0 Then
                On Error Resume Next
                GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(71)
                GLOBAL_MessageFrm.Show (1)
                txtHours.SetFocus
            ElseIf Trim(txtAttendant.text) = "" Then
                On Error Resume Next
                GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(70)
                GLOBAL_MessageFrm.Show (1)
                txtAttendant.SetFocus
            Else
                Set con = New ADODB.Connection
                
                Set rec = New ADODB.Recordset
                
                con.ConnectionString = ConnString
                con.Open
                
                If isExtend = True Then
                    'Update Status of current session
                    Set cmd = New ADODB.Command
                    cmd.ActiveConnection = con
                    cmd.CommandType = adCmdStoredProc
                    cmd.Parameters.Append cmd.CreateParameter("@POS_SessionId", adInteger, adParamInput, , Val(POS_SessionId))
                    cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , 4)
                    cmd.CommandText = "POS_Session_StatusUpdate"
                    cmd.Execute
                    POS_SessionId = 0
                End If
                
                
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                
                cmd.Parameters.Append cmd.CreateParameter("@POS_SessionId", adInteger, adParamInputOutput, , Val(POS_SessionId))
                cmd.Parameters.Append cmd.CreateParameter("@POS_SalesId", adInteger, adParamInput, , Val(POS_SalesId))
                cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInput, , Val(CustomerId))
                cmd.Parameters.Append cmd.CreateParameter("@SlotId", adInteger, adParamInput, , Val(SlotId))
                cmd.Parameters.Append cmd.CreateParameter("@RoomId", adInteger, adParamInput, , Val(RoomId))
                cmd.Parameters.Append cmd.CreateParameter("@StartDate", adDate, adParamInput, , DateFrom.value)
                cmd.Parameters.Append cmd.CreateParameter("@EndDate", adDate, adParamInput, , dateTo.value)
                cmd.Parameters.Append cmd.CreateParameter("@ChildName", adVarChar, adParamInput, 250, txtChild.text)
                cmd.Parameters.Append cmd.CreateParameter("@Attendant", adVarChar, adParamInput, 250, txtAttendant.text)
                cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
                cmd.Parameters.Append cmd.CreateParameter("@ImageUrl", adVarChar, adParamInput, 500, "") 'NOT SET
                If chkUnli.value = vbChecked Then
                    cmd.Parameters.Append cmd.CreateParameter("@IsUnlimited", adBoolean, adParamInput, , "True")
                Else
                    cmd.Parameters.Append cmd.CreateParameter("@IsUnlimited", adBoolean, adParamInput, , "False")
                End If
                
                If chkMin.value = vbChecked Then
                    cmd.Parameters.Append cmd.CreateParameter("@isMinutes", adBoolean, adParamInput, , "True")
                Else
                    cmd.Parameters.Append cmd.CreateParameter("@isMinutes", adBoolean, adParamInput, , "False")
                End If
                
                cmd.Parameters.Append cmd.CreateParameter("@Hours", adDecimal, adParamInput, , Val(txtHours.text))
                                      cmd.Parameters("@Hours").NumericScale = 2
                                      cmd.Parameters("@Hours").Precision = 18
                
                If Val(POS_SessionId) = 0 Then
                    cmd.CommandText = "POS_Session_Insert"
                    cmd.Execute
                    txtAdmissionNumber.text = cmd.Parameters("@POS_SessionId")
                    POS_SessionId = cmd.Parameters("@POS_SessionId")
                    
                    'Disable and Show Controls
                    btnCheckOut.Visible = True
                    btnItemSearch.Enabled = False
                    txtOR.Enabled = False
                    txtChild.Enabled = False
                    
                Else
                    cmd.CommandText = "POS_Session_Update"
                    cmd.Execute
                End If
                
                'UPDATE STATUS
                
                
                con.Close
                
                'PRINT SLIP
                Dim xy As Variant
                xy = MsgBox("Do you want to print a receipt?", vbQuestion + vbYesNo)
                If xy = vbYes Then
                    '**PRINT RECEIPT******
                    Dim crxApp As New CRAXDRT.Application
                    Dim crxRpt As New CRAXDRT.Report
                    Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\POS_admissionslip.rpt")
                    crxRpt.RecordSelectionFormula = "{POS_Sessions.POS_sessionId}= " & Val(POS_SessionId) & ""
                    crxRpt.DiscardSavedData
                    crxRpt.EnableParameterPrompting = False
                    crxRpt.ParameterFields(1).AddCurrentValue ""
        
                    Call ResetRptDB(crxRpt)
                    crxRpt.PrintOut False
                    '**END PRINT RECEIPT**
                End If
                
                MsgBox "Record saved.", vbInformation
                
                Unload Me
                Unload POS_PlayHouseFrm
                'DISABLE Controls
            End If
        Case 6 'PRINT
            If POS_SessionId <> 0 Then
                '**PRINT RECEIPT******
                    'Dim crxApp As New CRAXDRT.Application
                    'Dim crxRpt As New CRAXDRT.Report
                    Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\POS_admissionslip.rpt")
                    crxRpt.RecordSelectionFormula = "{POS_Sessions.POS_sessionId}= " & Val(POS_SessionId) & ""
                    crxRpt.DiscardSavedData
                    crxRpt.EnableParameterPrompting = False
                    crxRpt.ParameterFields(1).AddCurrentValue ""
        
                    Call ResetRptDB(crxRpt)
                    crxRpt.PrintOut False
                    '**END PRINT RECEIPT**
            End If
        Case 4 ' EXTEND
            Dim extend As String
            extend = InputBox("Please input POS OR #:")
            
            'CHECK IF VALID OR #
            Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            Set rec = New ADODB.Recordset
            
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "POS_Session_ORCheck"
            cmd.Parameters.Append cmd.CreateParameter("@OrNumber", adVarChar, adParamInput, 50, extend)
            Set rec = cmd.Execute
            If Not rec.EOF Then
                If rec!result = "False" Then
                    GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(68)
                    GLOBAL_MessageFrm.Show (1)
                    isExtend = False
                    Exit Sub
                Else
                    txtOR.text = extend
                    txtCustomer.text = rec!Name
                    DateFrom.value = rec!Date
                    dateTo.value = rec!Date
                    POS_SalesId = rec!POS_SalesId
                    CustomerId = rec!CustomerId
                    
                    'enable
                    'txtCustomer.Enabled = True
                    'DateFrom.Enabled = True
                    txtHours.Enabled = True
                    dateTo.Enabled = True
                    'txtChild.Enabled = True
                    txtAttendant.Enabled = True
                    isExtend = True
                    
                    MsgBox "OR # accepted.", vbInformation
                End If
            Else
                isExtend = False
            End If
            con.Close
    End Select
End Sub

Private Sub txtHours_Change()
    If IsNumeric(txtHours.text) = False Then
        txtHours.text = "0"
    Else
        dateTo.value = DateFrom.value
        If chkMin.value = vbChecked Then
            dateTo.value = DateAdd("n", Val(txtHours.text), dateTo.value)
        Else
            dateTo.value = DateAdd("h", Val(txtHours.text), dateTo.value)
        End If
    End If
    
End Sub

Private Sub txtOR_GotFocus()
    selectText txtOR
End Sub

Private Sub txtOR_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            btnItemSearch_Click
    End Select
End Sub
