VERSION 5.00
Begin VB.Form POS_CardPaymentFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8055
   Icon            =   "POS_CardPaymentFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
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
      Picture         =   "POS_CardPaymentFrm.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   4680
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
      Picture         =   "POS_CardPaymentFrm.frx":065B
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4680
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
      Picture         =   "POS_CardPaymentFrm.frx":29EA
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4680
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
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
      Left            =   4200
      TabIndex        =   14
      TabStop         =   0   'False
      Text            =   "X X X X X X "
      Top             =   2640
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
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
      Left            =   3120
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "X X X X"
      Top             =   2640
      Width           =   975
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
      TabIndex        =   6
      Text            =   "0.00"
      Top             =   3840
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
      TabIndex        =   5
      Top             =   3120
      Width           =   4575
   End
   Begin VB.TextBox txtCardNumber 
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
      Left            =   6240
      MaxLength       =   4
      TabIndex        =   4
      Top             =   2640
      Width           =   1455
   End
   Begin VB.ComboBox cmbCardType 
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
      ItemData        =   "POS_CardPaymentFrm.frx":4DBE
      Left            =   3120
      List            =   "POS_CardPaymentFrm.frx":4DC0
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1440
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
   Begin VB.TextBox txtName 
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
      TabIndex        =   3
      Top             =   2160
      Width           =   4575
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Left            =   6120
      TabIndex        =   16
      Top             =   2640
      Width           =   90
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
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
      Left            =   4080
      TabIndex        =   15
      Top             =   2640
      Width           =   90
   End
   Begin VB.Label lblAmount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
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
      Top             =   3960
      Width           =   930
   End
   Begin VB.Label lblReference 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reference:"
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
      Top             =   3240
      Width           =   1230
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Card Type:"
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
      Top             =   1440
      Width           =   1230
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
      TabIndex        =   9
      Top             =   960
      Width           =   645
   End
   Begin VB.Label lblCard 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Card Number:"
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
      Top             =   2760
      Width           =   1620
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
      Caption         =   "Card Payment"
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
      TabIndex        =   7
      Top             =   360
      Width           =   1620
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "POS_CardPaymentFrm.frx":4DC2
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name on Card:"
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
      Top             =   2280
      Width           =   1710
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   4455
      Left            =   120
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "POS_CardPaymentFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Populate(ByVal data As String)
    Dim item As MSComctlLib.ListItem
    Select Case data
        Case "CardType"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("CardType")
            cmbCardType.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    cmbCardType.AddItem rec!CardType
                    cmbCardType.ItemData(cmbCardType.NewIndex) = rec!CardTypeId
                    rec.MoveNext
                Loop
            End If
            On Error Resume Next
            cmbCardType.ListIndex = 0
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
    If Trim(txtName.text) = "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(18)
        GLOBAL_MessageFrm.Show (1)
        txtName.SetFocus
    ElseIf Trim(txtCardNumber.text) = "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(51)
        GLOBAL_MessageFrm.Show (1)
        txtCardNumber.SetFocus
    ElseIf Trim(txtReferenceNumber.text) = "" Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(52)
        GLOBAL_MessageFrm.Show (1)
        txtReferenceNumber.SetFocus
    ElseIf Val(Replace(txtAmount.text, ",", "")) <= 0 Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(24)
        GLOBAL_MessageFrm.Show (1)
        txtAmount.SetFocus
    Else
        'store data in class
        With CardInfo
            .BankId = cmbBank.ItemData(cmbBank.ListIndex)
            .CardTypeId = cmbCardType.ItemData(cmbCardType.ListIndex)
            .NameOnCard = txtName.text
            .CardNumber = txtCardNumber.text
            .Reference = txtReferenceNumber.text
            .Amount = Val(Replace(txtAmount.text, ",", ""))
        End With
        POS_PayFrm.txtCard.text = FormatNumber(txtAmount.text, 2, vbTrue, vbFalse)
        POS_PayFrm.ComputeChange
        Unload Me
    End If
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnRemove_Click()
    If btnRemove.Visible = False Then Exit Sub
    ClearClassData (0)
    POS_PayFrm.txtCard.text = "0.00"
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
    Populate "CardType"
    Populate "Bank"
    
    With POS_CardPaymentFrm
        'search selected bank
        Dim X, Y As Integer
        For X = 0 To .cmbBank.ListCount - 1
            If .cmbBank.ItemData(X) = CardInfo.BankId Then
                .cmbBank.ListIndex = X
                Exit For
            End If
        Next X
        For Y = 0 To .cmbCardType.ListCount - 1
            If .cmbCardType.ItemData(Y) = CardInfo.CardTypeId Then
                .cmbCardType.ListIndex = Y
                Exit For
            End If
        Next Y
        
        .txtName.text = CardInfo.NameOnCard
        .txtCardNumber.text = CardInfo.CardNumber
        .txtReferenceNumber.text = CardInfo.Reference
        .txtAmount.text = FormatNumber(CardInfo.Amount, 2, vbTrue, vbFalse)
    End With
    
    If Val(Replace(txtAmount.text, ",", "")) = 0 Then
        btnRemove.Visible = False
    Else
        btnRemove.Visible = True
    End If
    
End Sub

Private Sub txtAmount_Change()
    If IsNumeric(txtAmount.text) = False Then
        txtAmount.text = "0.00"
        selectText txtAmount
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

Private Sub txtCardNumber_Click()
    Set SYS_OSKFrm.txtControl = txtCardNumber
    SYS_OSKFrm.Caption = lblCard.Caption
    SYS_OSKFrm.Show (1)
End Sub

Private Sub txtCardNumber_GotFocus()
    selectText txtCardNumber
End Sub

Private Sub txtname_Click()
    Set SYS_OskAlphaFrm.txtControl = txtName
    SYS_OskAlphaFrm.Caption = lblName.Caption
    SYS_OskAlphaFrm.Show (1)
End Sub

Private Sub txtName_GotFocus()
    selectText txtName
End Sub

Private Sub txtReferenceNumber_Click()
    Set SYS_OskAlphaFrm.txtControl = txtReferenceNumber
    SYS_OskAlphaFrm.Caption = lblReference.Caption
    SYS_OskAlphaFrm.Show (1)
End Sub

Private Sub txtReferenceNumber_GotFocus()
    selectText txtReferenceNumber
End Sub
