VERSION 5.00
Begin VB.Form POS_PaymentOptionsFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   Icon            =   "POS_PaymentOptionsFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCash 
      BackColor       =   &H00C0C0C0&
      Caption         =   "3: CASH"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      Picture         =   "POS_PaymentOptionsFrm.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   4335
   End
   Begin VB.CommandButton btnOthers 
      BackColor       =   &H00FF8080&
      Caption         =   "2: OTHERS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      Picture         =   "POS_PaymentOptionsFrm.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   4335
   End
   Begin VB.CommandButton btnCheck 
      BackColor       =   &H0080C0FF&
      Caption         =   "1: CHECK"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      Picture         =   "POS_PaymentOptionsFrm.frx":0C07
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   4335
   End
   Begin VB.Image picModuleImage 
      Height          =   480
      Left            =   120
      Picture         =   "POS_PaymentOptionsFrm.frx":10F7
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblCaption_Title 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Payment Options"
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
      Width           =   2160
   End
End
Attribute VB_Name = "POS_PaymentOptionsFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCheck_Click()
    POS_CheckPaymentFrm.lblAmountDue.Caption = POS_CashPayFrm.lblAmountDue.Caption
    POS_CheckPaymentFrm.Show
End Sub

Private Sub btnCash_Click()
    Unload Me
    Unload POS_PaymentOptionsFrm
End Sub

Private Sub btnNoDiscount_KeyDown(KeyCode As Integer, Shift As Integer)
    
End Sub

Private Sub btnOthers_Click()
    POS_OtherPaymentFrm.lblAmountDue.Caption = POS_CashPayFrm.lblAmountDue.Caption
    POS_OtherPaymentFrm.Show
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKey1
            btnCheck_Click
        Case vbKey2
            btnOthers_Click
        Case vbKey3
            btnCash_Click
    End Select
End Sub

