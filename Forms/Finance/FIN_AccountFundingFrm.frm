VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FIN_AccountFundingFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Account Funding"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   Icon            =   "FIN_AccountFundingFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   1000
      Begin VB.CommandButton btnFund 
         Caption         =   "Funds"
         DisabledPicture =   "FIN_AccountFundingFrm.frx":000C
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   22
         Picture         =   "FIN_AccountFundingFrm.frx":05F1
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton btnBank 
         Caption         =   "Banks"
         DisabledPicture =   "FIN_AccountFundingFrm.frx":0BF9
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   22
         Picture         =   "FIN_AccountFundingFrm.frx":1284
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Frame FRE_Main 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   7095
      Left            =   1035
      TabIndex        =   16
      Top             =   0
      Width           =   7935
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
         Height          =   375
         Left            =   6480
         TabIndex        =   15
         Top             =   6480
         Width           =   1335
      End
      Begin VB.CommandButton btnSave 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   14
         Top             =   6480
         Width           =   1335
      End
      Begin VB.Frame FRE_Fund 
         BackColor       =   &H00FFFFFF&
         Height          =   6375
         Left            =   120
         TabIndex        =   27
         Top             =   0
         Visible         =   0   'False
         Width           =   7695
         Begin VB.TextBox txtCheckAmount 
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
            Height          =   330
            Left            =   1440
            TabIndex        =   10
            Top             =   3840
            Width           =   3015
         End
         Begin VB.TextBox txtParticulars 
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
            Left            =   1440
            TabIndex        =   8
            Text            =   "Account Funding"
            Top             =   2760
            Width           =   3015
         End
         Begin VB.TextBox txtCheckNumber 
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
            Height          =   330
            Left            =   1440
            TabIndex        =   12
            Top             =   4575
            Width           =   3015
         End
         Begin VB.TextBox txtFund_Remarks 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Left            =   1440
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Top             =   5160
            Visible         =   0   'False
            Width           =   3015
         End
         Begin VB.ComboBox cmbFund 
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
            Height          =   345
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1320
            Width           =   3015
         End
         Begin VB.TextBox txtFund_Amount 
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
            Height          =   330
            Left            =   1440
            TabIndex        =   7
            Top             =   2415
            Width           =   3015
         End
         Begin MSComCtl2.DTPicker dtFund 
            Height          =   330
            Left            =   1440
            TabIndex        =   9
            Top             =   3120
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   582
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
            Format          =   69140481
            CurrentDate     =   41509
         End
         Begin MSComCtl2.DTPicker dtCheckDate 
            Height          =   330
            Left            =   1440
            TabIndex        =   11
            Top             =   4200
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   582
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
            Format          =   69140481
            CurrentDate     =   41509
         End
         Begin VB.Label lblCheckBalance 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
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
            Left            =   1440
            TabIndex        =   43
            Top             =   2070
            Width           =   3015
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Check Bal"
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
            Left            =   240
            TabIndex        =   42
            Top             =   2040
            Width           =   885
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Check Amt."
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
            Left            =   240
            TabIndex        =   41
            Top             =   3840
            Width           =   1050
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Particulars"
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
            Left            =   240
            TabIndex        =   40
            Top             =   2760
            Width           =   975
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Check #"
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
            Left            =   240
            TabIndex        =   39
            Top             =   4575
            Width           =   705
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Check Date"
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
            Left            =   240
            TabIndex        =   38
            Top             =   4200
            Width           =   1035
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Balance"
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
            Left            =   240
            TabIndex        =   37
            Top             =   1680
            Width           =   720
         End
         Begin VB.Label lblFund_Balance 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
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
            Left            =   1440
            TabIndex        =   36
            Top             =   1710
            Width           =   3015
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Funds"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   345
            Left            =   240
            TabIndex        =   33
            Top             =   240
            Width           =   705
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "You can add more cash to your funds to support your credit sales orders."
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
            TabIndex        =   32
            Top             =   600
            Width           =   6135
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
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
            Left            =   240
            TabIndex        =   31
            Top             =   3120
            Width           =   435
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
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
            Left            =   240
            TabIndex        =   30
            Top             =   5160
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
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
            Left            =   240
            TabIndex        =   29
            Top             =   2415
            Width           =   750
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fund"
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
            Left            =   240
            TabIndex        =   28
            Top             =   1320
            Width           =   465
         End
      End
      Begin VB.Frame FRE_Bank 
         BackColor       =   &H00FFFFFF&
         Height          =   6375
         Left            =   120
         TabIndex        =   19
         Top             =   0
         Width           =   7695
         Begin VB.TextBox txtBank_Remarks 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Left            =   1440
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   3600
            Width           =   3015
         End
         Begin VB.ComboBox cmbBank 
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
            Height          =   345
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   1320
            Width           =   3015
         End
         Begin VB.TextBox txtBank_Amount 
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
            Height          =   330
            Left            =   1440
            TabIndex        =   4
            Top             =   3015
            Width           =   3015
         End
         Begin VB.ComboBox cmbAccountNumber 
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
            Height          =   345
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1680
            Width           =   3015
         End
         Begin MSComCtl2.DTPicker dtBank 
            Height          =   330
            Left            =   1440
            TabIndex        =   3
            Top             =   2640
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   582
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
            Format          =   69140481
            CurrentDate     =   41509
         End
         Begin VB.Label lblBank_Balance 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
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
            Left            =   1440
            TabIndex        =   35
            Top             =   2070
            Width           =   3015
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Balance"
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
            Left            =   240
            TabIndex        =   34
            Top             =   2040
            Width           =   720
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Banks"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   345
            Left            =   240
            TabIndex        =   26
            Top             =   240
            Width           =   690
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "You can fund your bank accounts to help support your purchase orders."
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
            TabIndex        =   25
            Top             =   600
            Width           =   6135
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
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
            Left            =   240
            TabIndex        =   24
            Top             =   2640
            Width           =   435
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Particulars"
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
            Left            =   240
            TabIndex        =   23
            Top             =   3600
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account #"
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
            Left            =   240
            TabIndex        =   22
            Top             =   1680
            Width           =   900
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Amount"
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
            Left            =   240
            TabIndex        =   21
            Top             =   3015
            Width           =   750
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Bank"
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
            Left            =   240
            TabIndex        =   20
            Top             =   1320
            Width           =   450
         End
      End
   End
End
Attribute VB_Name = "FIN_AccountFundingFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub Populate(ByVal data As String)
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
        Case "Fund"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Fund")
            cmbFund.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    cmbFund.AddItem rec!FundName
                    cmbFund.ItemData(cmbFund.NewIndex) = rec!FundId
                    rec.MoveNext
                Loop
            End If
            On Error Resume Next
            cmbFund.ListIndex = 0
    End Select
End Sub
Private Sub ClearTexts()
    txtBank_Amount.text = ""
    txtBank_Remarks.text = ""
    txtFund_Amount.text = ""
    txtParticulars.text = ""
    txtCheckAmount.text = ""
    txtCheckNumber.text = ""
    txtFund_Remarks.text = ""
End Sub
Private Function isValidated() As Boolean
    If FRE_Bank.Visible = True Then
        If cmbAccountNumber.text = "" Then
            GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(21)
            GLOBAL_MessageFrm.Show (1)
            cmbAccountNumber.SetFocus
            isValidated = False
            Exit Function
        ElseIf Trim(txtBank_Amount.text) = "" Then
            GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(24)
            GLOBAL_MessageFrm.Show (1)
            txtBank_Amount.SetFocus
            isValidated = False
            Exit Function
        End If
    Else
        If Trim(txtFund_Amount.text) = "" Then
            GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(24)
            GLOBAL_MessageFrm.Show (1)
            txtFund_Amount.SetFocus
            isValidated = False
            Exit Function
        End If
    End If
    isValidated = True
End Function

Private Sub btnBank_Click()
    btnBank.Enabled = False
    btnFund.Enabled = True
    FRE_Bank.Visible = True
    FRE_Bank.Left = 120
    FRE_Fund.Visible = False
    FRE_Fund.Left = -99999
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnFund_Click()
    btnFund.Enabled = False
    btnBank.Enabled = True
    FRE_Bank.Visible = False
    FRE_Bank.Left = -99999
    FRE_Fund.Visible = True
    FRE_Fund.Left = 120
End Sub

Private Sub btnSave_Click()
    If isValidated = True Then
        Set con = New ADODB.Connection
        Set cmd = New ADODB.Command
        Set rec = New ADODB.Recordset
        
        con.ConnectionString = ConnString
        con.Open
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        If FRE_Bank.Visible = True Then
            cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , cmbAccountNumber.ItemData(cmbAccountNumber.ListIndex))
            cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtBank.value)
            cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Val(Replace(txtBank_Amount.text, ",", "")))
                                  cmd.Parameters("@Amount").Precision = 18
                                  cmd.Parameters("@Amount").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, txtBank_Remarks.text)
            cmd.CommandText = "FIN_FundBank_Insert"
            cmd.Execute
            
            'INFLOW OUTFLOW REPORT
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "FIN_InFlow_Insert"
            cmd.Parameters.Append cmd.CreateParameter("@Cash", adDecimal, adParamInput, , Val(Replace(txtBank_Amount.text, ",", "")))
                                  cmd.Parameters("@Cash").Precision = 18
                                  cmd.Parameters("@Cash").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , 0)
                                  cmd.Parameters("@CheckAmount").Precision = 18
                                  cmd.Parameters("@CheckAmount").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, Null)
            cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtBank.value)
            cmd.Parameters.Append cmd.CreateParameter("@Particulars", adVarChar, adParamInput, 250, txtBank_Remarks.text & ":" & cmbBank.text)
            cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@ExpenseId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@POS_SalesId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@SOPaymentId", adInteger, adParamInput, , Null)
            cmd.Execute
        Else
            cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , cmbFund.ItemData(cmbFund.ListIndex))
            cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtFund.value)
            cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Val(Replace(txtFund_Amount.text, ",", "")))
                                  cmd.Parameters("@Amount").Precision = 18
                                  cmd.Parameters("@Amount").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@CheckBalance", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.text, ",", "")))
                                  cmd.Parameters("@CheckBalance").Precision = 18
                                  cmd.Parameters("@CheckBalance").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, txtBank_Remarks.text)
            cmd.Parameters.Append cmd.CreateParameter("@Particulars", adVarChar, adParamInput, 250, txtParticulars.text)
           
            cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.text)
            cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
            cmd.CommandText = "FIN_Funds_Insert"
            cmd.Execute
            
            'INFLOW AND OUTFLOW
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "FIN_InFlow_Insert"
            cmd.Parameters.Append cmd.CreateParameter("@Cash", adDecimal, adParamInput, , Val(Replace(txtFund_Amount.text, ",", "")))
                                  cmd.Parameters("@Cash").Precision = 18
                                  cmd.Parameters("@Cash").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.text, ",", "")))
                                  cmd.Parameters("@CheckAmount").Precision = 18
                                  cmd.Parameters("@CheckAmount").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.text)
            cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
            cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtFund.value)
            cmd.Parameters.Append cmd.CreateParameter("@Particulars", adVarChar, adParamInput, 250, txtParticulars.text)
            cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@ExpenseId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@POS_SalesId", adInteger, adParamInput, , Null)
            cmd.Parameters.Append cmd.CreateParameter("@SOPaymentId", adInteger, adParamInput, , Null)
            cmd.Execute
            
            'CHECK REGISTRY
            If (Trim(txtCheckNumber.text) <> "" Or Val(Replace(txtCheckAmount.text, ",", "")) > 0) Then
                'Set con = New ADODB.Connection
                Set cmd = New ADODB.Command
                'con.ConnectionString = ConnString
                'con.Open
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "FIN_CheckRegistry_Insert"
                cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.text)
                cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
                cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.text, ",", "")))
                                      cmd.Parameters("@Amount").NumericScale = 2
                                      cmd.Parameters("@Amount").Precision = 18
                cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, txtParticulars.text)
                cmd.Parameters.Append cmd.CreateParameter("@isReceivable", adBoolean, adParamInput, , "True")
                cmd.Parameters.Append cmd.CreateParameter("@CheckStatusId", adInteger, adParamInput, , 1) '--open
                cmd.Parameters.Append cmd.CreateParameter("@CheckRegistryId", adInteger, adParamInputOutput, , Null)
                cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , Null)
                cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , Null)
                cmd.Parameters.Append cmd.CreateParameter("@ExpenseId", adInteger, adParamInput, , Null)
                cmd.Parameters.Append cmd.CreateParameter("@POS_SalesId", adInteger, adParamInput, , Null)
                cmd.Parameters.Append cmd.CreateParameter("@SOPaymentId", adInteger, adParamInput, , Null)
                cmd.Parameters.Append cmd.CreateParameter("@POPaymentId", adInteger, adParamInput, , Null)
                cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , Null)
                cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , 1)
                
                cmd.Execute
            End If
        End If
        con.Close
        'Unload Me
        MsgBox "Funds updated.", vbInformation, "QuickPOS"
        ClearTexts
    End If
End Sub

Private Sub cmbAccountNumber_Click()
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_BankAccount_Get"
    
    cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , cmbAccountNumber.ItemData(cmbAccountNumber.ListIndex))
    Set rec = cmd.Execute
    lblBank_Balance.Caption = "0.00"
    If Not rec.EOF Then
        lblBank_Balance.Caption = FormatNumber(rec!balance, 2, vbTrue, vbFalse)
    End If
    con.Close
End Sub

Private Sub cmbBank_Click()
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_BankAccount_Load"
    
    cmd.Parameters.Append cmd.CreateParameter("@BankId", adInteger, adParamInput, , cmbBank.ItemData(cmbBank.ListIndex))
    Set rec = cmd.Execute
    cmbAccountNumber.Clear
    lblBank_Balance.Caption = "0.00"
    If Not rec.EOF Then
        Do Until rec.EOF
            If rec!isActive = "True" Then
                cmbAccountNumber.AddItem rec!accountnumber & " - " & rec!AccountName
                cmbAccountNumber.ItemData(cmbAccountNumber.NewIndex) = rec!AccountId
            End If
            rec.MoveNext
        Loop
    End If
    con.Close
End Sub

Private Sub cmbFund_Click()
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_Fund_Get"
    
    cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , cmbFund.ItemData(cmbFund.ListIndex))
    Set rec = cmd.Execute
    lblFund_Balance.Caption = "0.00"
    lblCheckBalance.Caption = "0.00"
    If Not rec.EOF Then
        lblFund_Balance.Caption = FormatNumber(rec!balance, 2, vbTrue, vbFalse)
        lblCheckBalance.Caption = FormatNumber(rec!checkbalance, 2, vbTrue, vbFalse)
    End If
    con.Close
End Sub

Private Sub cmbType_Click()
'    If cmbType.text = "CASH" Then
'        dtCheckDate.Enabled = False
'        txtCheckNumber.Enabled = False
'    Else
'        dtCheckDate.Enabled = True
'        txtCheckNumber.Enabled = True
'    End If
End Sub

Private Sub Form_Load()
    Populate "Bank"
    Populate "Fund"
    dtFund.value = Format(Now(), "MM/DD/YY")
    dtBank.value = Format(Now(), "MM/DD/YY")
    dtCheckDate.value = Format(Now(), "MM/DD/YY")
    lblBank_Balance.Caption = "0.00"

End Sub

Private Sub txtBank_Amount_Change()
    If IsNumeric(txtBank_Amount.text) = False Then
        txtBank_Amount.text = "1.00"
    End If
End Sub

Private Sub txtFund_Amount_Change()
    If IsNumeric(txtFund_Amount.text) = False Then
        txtFund_Amount.text = "1.00"
    End If
End Sub

Private Sub txtParticulars_GotFocus()
    selectText txtParticulars
End Sub
