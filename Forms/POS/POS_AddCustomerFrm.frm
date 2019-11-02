VERSION 5.00
Begin VB.Form POS_AddCustomerFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6150
   Icon            =   "POS_AddCustomerFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtInterestRate 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   4
      Top             =   3840
      Width           =   4095
   End
   Begin VB.TextBox txtCardnumber 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   3
      Top             =   2760
      Width           =   4095
   End
   Begin VB.TextBox txtAddress 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   1680
      MaxLength       =   50
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   5520
      Width           =   4095
   End
   Begin VB.TextBox txtMobile 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   6
      Top             =   5160
      Width           =   4095
   End
   Begin VB.TextBox txtPhone 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   5
      Top             =   4800
      Width           =   4095
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1680
      MaxLength       =   500
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1800
      Width           =   4095
   End
   Begin VB.TextBox txtCode 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1440
      Width           =   4095
   End
   Begin VB.CommandButton btnAccept 
      Caption         =   "ENTER: Accept"
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
      Left            =   2760
      Picture         =   "POS_AddCustomerFrm.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7080
      Width           =   1575
   End
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
      Left            =   4440
      Picture         =   "POS_AddCustomerFrm.frx":23E0
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Credit Info"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   345
      Left            =   360
      TabIndex        =   20
      Top             =   3360
      Width           =   1230
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Interest Rate"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   19
      Top             =   3840
      Width           =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Card #"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   18
      Top             =   2760
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Info"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   345
      Left            =   360
      TabIndex        =   17
      Top             =   2280
      Width           =   1125
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contact Info"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   345
      Left            =   360
      TabIndex        =   16
      Top             =   4320
      Width           =   1425
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Basic Info"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   345
      Left            =   360
      TabIndex        =   15
      Top             =   960
      Width           =   1125
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   14
      Top             =   5520
      Width           =   750
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   13
      Top             =   5160
      Width           =   660
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   12
      Top             =   4800
      Width           =   600
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   11
      Top             =   1800
      Width           =   555
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   360
      TabIndex        =   10
      Top             =   1440
      Width           =   480
   End
   Begin VB.Label lblCaption_Title 
      AutoSize        =   -1  'True
      Caption         =   "Add Customer"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   1800
   End
   Begin VB.Image picModuleImage 
      Height          =   480
      Left            =   120
      Picture         =   "POS_AddCustomerFrm.frx":476F
      Top             =   120
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   6255
      Left            =   120
      Top             =   720
      Width           =   5895
   End
End
Attribute VB_Name = "POS_AddCustomerFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Function isValidated() As Boolean
    isValidated = False
    'CHECK EMPTY FIELDS
'    If Trim(txtCode.Text) = "" Then
''        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(40)
''        GLOBAL_MessageFrm.Show (1)
''        txtCode.SetFocus
'    Else
    If Trim(txtName.Text) = "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(18)
        GLOBAL_MessageFrm.Show (1)
        txtName.SetFocus
    Else
        isValidated = True
    End If
End Function
Private Sub btnAccept_Click()
    If isValidated = True Then
        On Error GoTo ErrHandler
        Set con = New ADODB.Connection
        Set cmd = New ADODB.Command
        
        'SAVE MAIN Customer DETAILS
        con.ConnectionString = ConnString
        con.Open
        con.BeginTrans
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        
        cmd.Parameters.Append cmd.CreateParameter("@CustomerId", adInteger, adParamInputOutput, , 0)
        cmd.Parameters.Append cmd.CreateParameter("@CustomerCode", adVarChar, adParamInput, 50, txtCode.Text)
        cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 500, txtName.Text)
        cmd.Parameters.Append cmd.CreateParameter("@Salesrep", adVarChar, adParamInput, 50, "")
        cmd.Parameters.Append cmd.CreateParameter("@Mobile", adVarChar, adParamInput, 50, txtMobile.Text)
        cmd.Parameters.Append cmd.CreateParameter("@Phone", adVarChar, adParamInput, 50, txtPhone.Text)
        cmd.Parameters.Append cmd.CreateParameter("@Address", adVarChar, adParamInput, 500, txtAddress.Text)
        cmd.Parameters.Append cmd.CreateParameter("@CityId", adInteger, adParamInput, , 0)
        cmd.Parameters.Append cmd.CreateParameter("@CardNumber", adVarChar, adParamInput, 250, txtCardnumber.Text)
        cmd.Parameters.Append cmd.CreateParameter("@CreditLimit", adDecimal, adParamInput, , 0)
                                      cmd.Parameters("@CreditLimit").NumericScale = 2
                                      cmd.Parameters("@CreditLimit").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@CardPoints", adDecimal, adParamInput, , 0)
                                      cmd.Parameters("@CardPoints").NumericScale = 2
                                      cmd.Parameters("@CardPoints").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@PricingSchemeId", adInteger, adParamInput, , 0)
        cmd.Parameters.Append cmd.CreateParameter("@MemberId", adInteger, adParamInput, , 2)
        cmd.Parameters.Append cmd.CreateParameter("@InterestRate", adDecimal, adParamInput, , NVAL(txtInterestRate.Text))
                              cmd.Parameters("@InterestRate").NumericScale = 2
                              cmd.Parameters("@InterestRate").Precision = 18
        cmd.CommandText = "BASE_Customer_Insert"
        cmd.Execute
        'CustomerId = cmd.Parameters("@CustomerId")
      
        con.CommitTrans
        con.Close
        POS_CustomerNameFrm.txtCustomer.SelStart = 0
        POS_CustomerNameFrm.txtCustomer.SelLength = Len(POS_CustomerNameFrm.txtCustomer.Text)
        POS_CustomerNameFrm.txtCustomer_Change
        Unload Me
    End If
    Exit Sub
ErrHandler:
    con.RollbackTrans
    If IsNumeric(Err.Description) = True Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(Err.Description)
    Else
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & Err.Description
    End If
        GLOBAL_MessageFrm.Show (1)
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

