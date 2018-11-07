VERSION 5.00
Begin VB.Form POS_OtherPaymentFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   Icon            =   "POS_OtherPaymentFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnRemove 
      Caption         =   "ALT+R: Remove"
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
      Picture         =   "POS_OtherPaymentFrm.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3240
      Width           =   1815
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
      Left            =   6360
      Picture         =   "POS_OtherPaymentFrm.frx":065B
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Width           =   1575
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
      Left            =   4680
      Picture         =   "POS_OtherPaymentFrm.frx":29EA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox txtRemarks 
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
      Left            =   3120
      TabIndex        =   1
      Top             =   1440
      Width           =   4575
   End
   Begin VB.TextBox txtReferenceNumber 
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
      Left            =   3120
      TabIndex        =   0
      Top             =   960
      Width           =   4575
   End
   Begin VB.TextBox txtAmount 
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
      Left            =   3120
      TabIndex        =   2
      Text            =   "0.00"
      Top             =   2400
      Width           =   4575
   End
   Begin VB.Label lblReferenceNumber 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reference Number:"
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
      TabIndex        =   6
      Top             =   960
      Width           =   2235
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "POS_OtherPaymentFrm.frx":4DBE
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Other Payment Option"
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
      Left            =   960
      TabIndex        =   5
      Top             =   360
      Width           =   2610
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   240
      X2              =   7800
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Remarks:"
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
      TabIndex        =   4
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount:"
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
      TabIndex        =   3
      Top             =   2520
      Width           =   1005
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   3015
      Left            =   120
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "POS_OtherPaymentFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btnAccept_Click()
    If Trim(txtReferenceNumber.Text) = "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(52)
        GLOBAL_MessageFrm.Show (1)
        txtReferenceNumber.SetFocus
    ElseIf Val(Replace(txtAmount.Text, ",", "")) <= 0 Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(24)
        GLOBAL_MessageFrm.Show (1)
        txtAmount.SetFocus
    Else
        With OtherInfo
            .Amount = Val(Replace(txtAmount.Text, ",", ""))
            .ReferenceNumber = txtReferenceNumber.Text
            .Remarks = txtRemarks.Text
        End With
        POS_PayFrm.txtOthers.Text = FormatNumber(txtAmount.Text, 2, vbTrue, vbFalse)
        POS_PayFrm.ComputeChange
        Unload Me
    End If
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnRemove_Click()
    If btnRemove.Visible = False Then Exit Sub
    ClearClassData (3)
    POS_PayFrm.txtOthers.Text = "0.00"
    POS_PayFrm.ComputeChange
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            btnAccept_Click
        Case vbKeyEscape
            btnCancel_Click
        Case vbKeyR
            If btnRemove.Visible = False Then Exit Sub
            If Shift = vbAltMask Then
                btnRemove_Click
            End If
    End Select
End Sub

Private Sub Form_Load()
    With POS_OtherPaymentFrm
        .txtReferenceNumber.Text = OtherInfo.ReferenceNumber
        .txtRemarks.Text = OtherInfo.Remarks
        .txtAmount.Text = FormatNumber(OtherInfo.Amount, 2, vbTrue, vbFalse)
    End With
    
    If Val(Replace(txtAmount.Text, ",", "")) = 0 Then
        btnRemove.Visible = False
    Else
        btnRemove.Visible = True
    End If
End Sub

Private Sub txtAmount_Click()
    Set SYS_OSKFrm.txtControl = txtAmount
    SYS_OSKFrm.Caption = "Amount"
    SYS_OSKFrm.Show (1)
End Sub

Private Sub txtAmount_GotFocus()
    selectText txtAmount
End Sub



Private Sub txtReferenceNumber_Click()
    Set SYS_OskAlphaFrm.txtControl = txtReferenceNumber
    SYS_OskAlphaFrm.Caption = lblReferenceNumber.Caption
    SYS_OskAlphaFrm.Show (1)
End Sub

Private Sub txtReferenceNumber_GotFocus()
    selectText txtReferenceNumber
End Sub

Private Sub txtRemarks_Click()
    Set SYS_OskAlphaFrm.txtControl = txtRemarks
    SYS_OskAlphaFrm.Caption = "Remarks"
    SYS_OskAlphaFrm.Show (1)
End Sub

Private Sub txtRemarks_GotFocus()
    selectText txtRemarks
End Sub
