VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form POS_CheckPaymentFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   Icon            =   "POS_CheckPaymentFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
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
      Picture         =   "POS_CheckPaymentFrm.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3840
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
      Picture         =   "POS_CheckPaymentFrm.frx":065B
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3840
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
      Picture         =   "POS_CheckPaymentFrm.frx":29EA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3840
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker dtCheckDate 
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   2160
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   873
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
      Format          =   74317825
      CurrentDate     =   42301
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
      TabIndex        =   4
      Text            =   "0.00"
      Top             =   3120
      Width           =   4575
   End
   Begin VB.ComboBox cmbBank 
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
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   960
      Width           =   4575
   End
   Begin VB.TextBox txtCheckNumber 
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
      Top             =   1680
      Width           =   4575
   End
   Begin VB.Label lblAmount 
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
      TabIndex        =   8
      Top             =   3240
      Width           =   1005
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bank:"
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
      TabIndex        =   7
      Top             =   960
      Width           =   645
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Check Date:"
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
      Top             =   2280
      Width           =   1365
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      X1              =   240
      X2              =   7800
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Check Payment"
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
      Width           =   1770
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "POS_CheckPaymentFrm.frx":4DBE
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblCheckNumber 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Check Number:"
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
      TabIndex        =   0
      Top             =   1800
      Width           =   1770
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   3615
      Left            =   120
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "POS_CheckPaymentFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub Populate(ByVal data As String)
    Dim item As MSComctlLib.ListItem
    Select Case data
        Case "Bank"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Bank")
            cmbBank.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        cmbBank.AddItem rec!bankname
                        cmbBank.ItemData(cmbBank.NewIndex) = rec!BankId
                    End If
                    rec.MoveNext
                Loop
            End If
            On Error Resume Next
            cmbBank.ListIndex = 0
    End Select
End Sub

Private Sub btnAccept_Click()
    If Trim(txtCheckNumber.text) = "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(32)
        GLOBAL_MessageFrm.Show (1)
        txtCheckNumber.SetFocus
    ElseIf Val(Replace(txtAmount.text, ",", "")) <= 0 Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(24)
        GLOBAL_MessageFrm.Show (1)
        txtAmount.SetFocus
    Else
        With CheckInfo
            .Amount = Val(Replace(txtAmount.text, ",", ""))
            .BankId = cmbBank.ItemData(cmbBank.ListIndex)
            .CheckDate = dtCheckDate.value
            .CheckNumber = txtCheckNumber.text
        End With
        POS_PayFrm.txtCheck.text = FormatNumber(txtAmount.text, 2, vbTrue, vbFalse)
        POS_PayFrm.ComputeChange
        Unload Me
    End If
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnRemove_Click()
    If btnRemove.Visible = False Then Exit Sub
    ClearClassData (1)
    POS_PayFrm.txtCheck.text = "0.00"
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
    Populate "Bank"
    
    With POS_CheckPaymentFrm
        'search selected bank
        Dim X As Integer
        For X = 0 To .cmbBank.ListCount - 1
            If .cmbBank.ItemData(X) = CardInfo.BankId Then
                .cmbBank.ListIndex = X
                Exit For
            End If
        Next X
        
        .txtCheckNumber.text = CheckInfo.CheckNumber
        .dtCheckDate.value = CheckInfo.CheckDate
        .txtAmount.text = FormatNumber(CheckInfo.Amount, 2, vbTrue, vbFalse)
    End With
    
    If Val(Replace(txtAmount.text, ",", "")) = 0 Then
        btnRemove.Visible = False
    Else
        btnRemove.Visible = True
    End If
End Sub

Private Sub txtAmount_Click()
    Set SYS_OSKFrm.txtControl = txtAmount
    SYS_OSKFrm.Caption = lblAmount.Caption
    SYS_OSKFrm.Show (1)
End Sub

Private Sub txtAmount_GotFocus()
    selectText txtAmount
End Sub

Private Sub txtCheckNumber_Click()
    Set SYS_OskAlphaFrm.txtControl = txtCheckNumber
    SYS_OskAlphaFrm.Caption = lblCheckNumber.Caption
    SYS_OskAlphaFrm.Show (1)
End Sub

Private Sub txtCheckNumber_GotFocus()
    selectText txtCheckNumber
End Sub
