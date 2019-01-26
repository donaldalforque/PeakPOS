VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FIN_VendorPaymentFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Supplier Payment"
   ClientHeight    =   10125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10125
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCash 
      Caption         =   "CASH"
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
      Left            =   120
      Picture         =   "FIN_VendorPaymentFrm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5040
      Width           =   975
   End
   Begin VB.CommandButton btnCheck 
      Caption         =   "CHECK"
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
      Left            =   1200
      Picture         =   "FIN_VendorPaymentFrm.frx":060C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   975
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
      Left            =   8520
      TabIndex        =   14
      Top             =   9600
      Width           =   1575
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
      Height          =   375
      Left            =   10140
      TabIndex        =   15
      Top             =   9600
      Width           =   1575
   End
   Begin MSComctlLib.ListView lvOrders 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   6588
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "SalesOrderId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Order #"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Due Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Terms"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Balance"
         Object.Width           =   2716
      EndProperty
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   3495
      Left            =   120
      TabIndex        =   18
      Top             =   6000
      Width           =   11655
      Begin VB.Frame FRE_REMARKS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "REMARKS"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   120
         TabIndex        =   12
         Top             =   2520
         Width           =   11415
         Begin VB.TextBox txtRemarks 
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
            MultiLine       =   -1  'True
            TabIndex        =   13
            Top             =   360
            Width           =   9615
         End
         Begin VB.Label Label13 
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
            TabIndex        =   38
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame FRE_CHECK 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CHECK"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   6000
         TabIndex        =   6
         Top             =   240
         Width           =   5535
         Begin VB.ComboBox txtBank 
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
            ItemData        =   "FIN_VendorPaymentFrm.frx":0BC4
            Left            =   1560
            List            =   "FIN_VendorPaymentFrm.frx":0BC6
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1800
            Width           =   3855
         End
         Begin VB.TextBox txtCheckNumber 
            Alignment       =   1  'Right Justify
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
            Left            =   1560
            TabIndex        =   8
            Top             =   720
            Width           =   3855
         End
         Begin VB.TextBox txtCheckAmount 
            Alignment       =   1  'Right Justify
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
            Left            =   1560
            TabIndex        =   7
            Text            =   "0.00"
            Top             =   375
            Width           =   3855
         End
         Begin VB.TextBox txtTax 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H000000FF&
            Height          =   330
            Left            =   1560
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   1455
            Width           =   3855
         End
         Begin MSComCtl2.DTPicker dtCheckDate 
            Height          =   345
            Left            =   1560
            TabIndex        =   9
            Top             =   1080
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   609
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
            Format          =   148570113
            CurrentDate     =   41646
         End
         Begin VB.Label Label4 
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
            Left            =   120
            TabIndex        =   39
            Top             =   1800
            Width           =   450
         End
         Begin VB.Label Label7 
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
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Width           =   1050
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
            Left            =   120
            TabIndex        =   36
            Top             =   1080
            Width           =   1035
         End
         Begin VB.Label Label10 
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
            Left            =   120
            TabIndex        =   35
            Top             =   720
            Width           =   705
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Adjustment"
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
            Left            =   120
            TabIndex        =   34
            Top             =   1440
            Width           =   1095
         End
      End
      Begin VB.Frame FRE_CASH 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CASH"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   5775
         Begin VB.TextBox txtCash 
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
            Left            =   1680
            TabIndex        =   5
            Text            =   "0.00"
            Top             =   720
            Width           =   3855
         End
         Begin MSComCtl2.DTPicker dtDate 
            Height          =   345
            Left            =   1680
            TabIndex        =   4
            Top             =   360
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   609
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
            Format          =   148570113
            CurrentDate     =   41646
         End
         Begin VB.Label Label11 
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
            TabIndex        =   33
            Top             =   360
            Width           =   435
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cash"
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
            TabIndex        =   32
            Top             =   720
            Width           =   435
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   2175
         Left            =   -9999
         TabIndex        =   25
         Top             =   480
         Width           =   3855
         Begin VB.ComboBox cmbAccount 
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
            TabIndex        =   28
            Top             =   360
            Visible         =   0   'False
            Width           =   3855
         End
         Begin VB.CheckBox chkOnline 
            BackColor       =   &H0080C0FF&
            Caption         =   "Check Payment"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1440
            TabIndex        =   27
            Top             =   720
            Visible         =   0   'False
            Width           =   3855
         End
         Begin VB.TextBox txtAmountInWords 
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
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   26
            Top             =   1920
            Visible         =   0   'False
            Width           =   5655
         End
         Begin VB.Label lblAccount 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Account"
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
            Left            =   0
            TabIndex        =   31
            Top             =   360
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.Label lblBank 
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
            Left            =   0
            TabIndex        =   30
            Top             =   0
            Visible         =   0   'False
            Width           =   450
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "[For Check Issuance, Amount in Words]"
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
            Left            =   0
            TabIndex        =   29
            Top             =   1560
            Visible         =   0   'False
            Width           =   3615
         End
      End
      Begin VB.ComboBox cmbType 
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
         Left            =   -9999
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   960
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Method"
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
         Left            =   -9999
         TabIndex        =   20
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.Label lblPayable 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Net Payable: 0.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   7200
      TabIndex        =   24
      Top             =   5400
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "FIN_VendorPaymentFrm.frx":0BC8
      Top             =   240
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier Payment"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   345
      Left            =   720
      TabIndex        =   23
      Top             =   240
      Width           =   2025
   End
   Begin VB.Label lblSelectAll 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select All"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   120
      MouseIcon       =   "FIN_VendorPaymentFrm.frx":11E0
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   4680
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unselect All"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   1080
      MouseIcon       =   "FIN_VendorPaymentFrm.frx":1332
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Selected: 0.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   7200
      TabIndex        =   22
      Top             =   5040
      Width           =   4335
   End
   Begin VB.Label lblTotalBalance 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Balance: 0.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   7200
      TabIndex        =   21
      Top             =   4680
      Width           =   4335
   End
End
Attribute VB_Name = "FIN_VendorPaymentFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public VendorId As Long

Private Sub btnCancel_Click()
    Unload Me
End Sub
Public Sub Populate(ByVal data As String)
    Select Case data
        Case "Bank"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Bank")
            txtBank.Clear
            txtBank.AddItem ""
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        txtBank.AddItem rec!bankname
                        txtBank.ItemData(txtBank.NewIndex) = rec!BankId
                    End If
                    rec.MoveNext
                Loop
            End If
            On Error Resume Next
            txtBank.ListIndex = 0
    End Select
End Sub

Private Sub btnCash_Click()
    FRE_CASH.Enabled = True
    FRE_CHECK.Enabled = False
    txtCheckAmount.Text = "0.00"
End Sub

Private Sub btnCheck_Click()
    FRE_CASH.Enabled = False
    FRE_CHECK.Enabled = True
    txtCash.Text = "0.00"
End Sub

Private Sub btnSave_Click()
    If FRE_CHECK.Enabled = True Then
        If txtCheckNumber.Text = "" Then
            MsgBox "Check number is required.", vbCritical, "PeakPOS"
            txtCheckNumber.SetFocus
            Exit Sub
        End If
        If txtBank.Text = "" Then
            MsgBox "Bank is required.", vbCritical, "PeakPOS"
            txtBank.SetFocus
            Exit Sub
        End If
    End If
    
    If Val(Replace(Replace(lblTotal.Caption, "Total Selected:", ""), ",", "")) < Val(Replace(txtCash.Text, ",", "")) + Val(Replace(txtCheckAmount.Text, ",", "")) + Val(Replace(txtTax.Text, ",", "")) Then
        MsgBox "Cannot overpay.", vbCritical, "PeakPOS"
        Exit Sub
    End If
        
    Dim x As Variant
    x = MsgBox("Are you sure all information are correct?", vbQuestion + vbYesNo, "Verify")
    If x = vbNo Then
        Exit Sub
    End If
    Dim hasSelected As Boolean
    Dim item As MSComctlLib.ListItem
    For Each item In lvOrders.ListItems
        If item.Checked = True Then
            hasSelected = True
            Exit For
        End If
    Next
    If hasSelected = False Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(37)
        GLOBAL_MessageFrm.Show (1)
        Exit Sub
    End If
    
    Dim isOnline As Boolean
    If chkOnline.value = Checked Then isOnline = True
    If chkOnline.value = Unchecked Then isOnline = False
    
    'Save Payment
    If isOnline = True Then
        If cmbAccount.Text = "" Then
            GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(21)
            GLOBAL_MessageFrm.Show (1)
            cmbAccount.SetFocus
            Exit Sub
        End If
    End If
    
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    con.ConnectionString = ConnString
    con.Open
    con.BeginTrans
    
    'TRANSACTION ID
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_TransactionId_Insert"
    cmd.Parameters.Append cmd.CreateParameter("@TransactionId", adInteger, adParamInputOutput, , 0)
    cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, "Vendor Payment")
    cmd.Execute
    Dim TransactionId As Long
    TransactionId = cmd.Parameters("@TransactionId")
    
    'PO_Payment
    Dim Payment As Double
    Payment = Val(Replace(txtCash.Text, ",", "")) + Val(Replace(txtCheckAmount.Text, ",", "")) + Val(Replace(txtTax.Text, ",", ""))
    Dim Orders As String
    Dim PurchaseOrderId As Integer
    PurchaseOrderId = lvOrders.ListItems.item(1).SubItems(1) 'For Check Registry
    
    For Each item In lvOrders.ListItems
        If item.Checked = True Then
            If Payment <= 0 Then Exit For
            
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "PO_Payment_Insert"
            cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , item.SubItems(1))
            If Payment >= NVAL(item.SubItems(6)) Then
                If FRE_CHECK.Enabled = False Then 'IF CASH
                    cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , NVAL(item.SubItems(6)))
                                          cmd.Parameters("@Amount").NumericScale = 2
                                          cmd.Parameters("@Amount").Precision = 18
                    cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtDate.value)
                    cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , 0)
                                  cmd.Parameters("@CheckAmount").NumericScale = 2
                                  cmd.Parameters("@CheckAmount").Precision = 18
                    cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, "")
                    cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
                Else
                    cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , 0)
                                          cmd.Parameters("@Amount").NumericScale = 2
                                          cmd.Parameters("@Amount").Precision = 18
                    cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtDate.value)
                    cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , NVAL(item.SubItems(6)))
                                  cmd.Parameters("@CheckAmount").NumericScale = 2
                                  cmd.Parameters("@CheckAmount").Precision = 18
                    cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.Text)
                    cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
                End If
            Else
                If FRE_CHECK.Enabled = False Then 'IF CASH
                    cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Payment)
                                          cmd.Parameters("@Amount").NumericScale = 2
                                          cmd.Parameters("@Amount").Precision = 18
                    cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtDate.value)
                    cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , 0)
                                  cmd.Parameters("@CheckAmount").NumericScale = 2
                                  cmd.Parameters("@CheckAmount").Precision = 18
                    cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, "")
                    cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
                Else
                    cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , 0)
                                          cmd.Parameters("@Amount").NumericScale = 2
                                          cmd.Parameters("@Amount").Precision = 18
                    cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtDate.value)
                    cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Payment)
                                  cmd.Parameters("@CheckAmount").NumericScale = 2
                                  cmd.Parameters("@CheckAmount").Precision = 18
                    cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.Text)
                    cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
                End If
            End If
                                  
                                  
            cmd.Parameters.Append cmd.CreateParameter("@Tax", adDecimal, adParamInput, , Null)
                                  cmd.Parameters("@Tax").NumericScale = 2
                                  cmd.Parameters("@Tax").Precision = 18
            cmd.Parameters.Append cmd.CreateParameter("@Bank", adVarChar, adParamInput, 50, txtBank.Text)
            If isOnline = True Then
                cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , 1)
                cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , Null)
            Else
                cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , Null)
                cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , 1)
            End If
            cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, txtRemarks.Text)
            cmd.Parameters.Append cmd.CreateParameter("@POPaymentId", adInteger, adParamInputOutput, , 0)
            cmd.Parameters.Append cmd.CreateParameter("@TransactionId", adInteger, adParamInput, , TransactionId)
            cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInput, 250, item.SubItems(2))
            cmd.Parameters.Append cmd.CreateParameter("@OrderBalance", adDecimal, adParamInput, , Val(Replace(item.SubItems(6), ",", "")))
                                  cmd.Parameters("@OrderBalance").NumericScale = 2
                                  cmd.Parameters("@OrderBalance").Precision = 18
            cmd.Execute
            
            Payment = Payment - Val(Replace(item.SubItems(6), ",", ""))
            Orders = Orders & "[" & item.SubItems(2) & "]"
        End If
    Next
    
    'PAYMENT HISTORY
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "PO_PaymentHistory_Insert"
    cmd.Parameters.Append cmd.CreateParameter("@VendorId", adInteger, adParamInput, , VendorId)
    cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , dtDate.value)
    cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Val(Replace(txtCash.Text, ",", "")))
                          cmd.Parameters("@Amount").NumericScale = 2
                          cmd.Parameters("@Amount").Precision = 18
    cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.Text, ",", "")))
                          cmd.Parameters("@CheckAmount").NumericScale = 2
                          cmd.Parameters("@CheckAmount").Precision = 18
    cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 50, txtCheckNumber.Text)
    cmd.Parameters.Append cmd.CreateParameter("@Bank", adVarChar, adParamInput, 50, txtBank.Text)
    cmd.Parameters.Append cmd.CreateParameter("@Tax", adDecimal, adParamInput, , Val(Replace(txtTax.Text, ",", "")))
                          cmd.Parameters("@Tax").NumericScale = 2
                          cmd.Parameters("@Tax").Precision = 18
    cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
    cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 4000, txtRemarks.Text & " " & Orders)
    cmd.Parameters.Append cmd.CreateParameter("@TransactionId", adInteger, adParamInput, , TransactionId)
    
    cmd.Execute
    
    SavePOSAuditTrail UserId, WorkstationId, "", "New payment for supplier: " & FIN_AccountsPayable.lvSearch.SelectedItem.SubItems(2) _
    & " Amount: " & FormatNumber(txtCash.Text, 2, vbTrue, vbFalse) & " Check Amt.: " & txtCheckAmount.Text & " for orders: " & Orders, "FINANCE"
    
    'CHECK REGISTRY
    If Trim(txtCheckNumber.Text) <> "" Or Val(Replace(txtCheckAmount.Text, ",", "")) > 0 Then

        'Set con = New ADODB.Connection
        Set cmd = New ADODB.Command
        'con.ConnectionString = ConnString
        'con.Open
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "FIN_CheckRegistry_Insert"
        cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.Text)
        cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
        cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.Text, ",", "")))
                              cmd.Parameters("@Amount").NumericScale = 2
                              cmd.Parameters("@Amount").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, Orders)
        cmd.Parameters.Append cmd.CreateParameter("@isReceivable", adBoolean, adParamInput, , "False")
        cmd.Parameters.Append cmd.CreateParameter("@CheckStatusId", adInteger, adParamInput, , 1) '--open
        cmd.Parameters.Append cmd.CreateParameter("@CheckRegistryId", adInteger, adParamInputOutput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , PurchaseOrderId)
        cmd.Parameters.Append cmd.CreateParameter("@ExpenseId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@POS_SalesId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@SOPaymentId", adInteger, adParamInput, , Null)
        cmd.Parameters.Append cmd.CreateParameter("@POPaymentId", adInteger, adParamInput, , PurchaseOrderId)
        'cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , cmbAccount.ItemData(cmbAccount.ListIndex))
        
        cmd.Execute
    End If
    
    'INFLOW AND OUTFLOW
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "FIN_OutFlow_Insert"
    cmd.Parameters.Append cmd.CreateParameter("@Cash", adDecimal, adParamInput, , Val(Replace(txtCash.Text, ",", "")))
                          cmd.Parameters("@Cash").Precision = 18
                          cmd.Parameters("@Cash").NumericScale = 2
    cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.Text, ",", "")))
                          cmd.Parameters("@CheckAmount").Precision = 18
                          cmd.Parameters("@CheckAmount").NumericScale = 2
    cmd.Parameters.Append cmd.CreateParameter("@CheckNumber", adVarChar, adParamInput, 250, txtCheckNumber.Text)
    cmd.Parameters.Append cmd.CreateParameter("@CheckDate", adDate, adParamInput, , dtCheckDate.value)
    cmd.Parameters.Append cmd.CreateParameter("@Date", adDate, adParamInput, , Format(Now, "MM/DD/YY"))
    cmd.Parameters.Append cmd.CreateParameter("@Particulars", adVarChar, adParamInput, 250, txtRemarks.Text & ":" & FIN_AccountsPayable.lvSearch.SelectedItem.SubItems(2))
    cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , Null)
    cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , Null)
    cmd.Parameters.Append cmd.CreateParameter("@ExpenseId", adInteger, adParamInput, , Null)
    cmd.Parameters.Append cmd.CreateParameter("@POS_SalesId", adInteger, adParamInput, , Null)
    cmd.Parameters.Append cmd.CreateParameter("@POPaymentId", adInteger, adParamInput, , Null)
    cmd.Execute
    
    
    If chkOnline.value = Checked Then
        'UPDATE BANK ACCOUNT
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "FIN_FundBank_Deduct"
        cmd.Parameters.Append cmd.CreateParameter("@AccountId", adInteger, adParamInput, , cmbAccount.ItemData(cmbAccount.ListIndex))
        cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.Text, ",", "")))
                              cmd.Parameters("@Amount").Precision = 18
                              cmd.Parameters("@Amount").NumericScale = 2
        cmd.Execute
    Else
        'UPDATE FUND ACCOUNT
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "FIN_Funds_Deduct"
        cmd.Parameters.Append cmd.CreateParameter("@FundId", adInteger, adParamInput, , 1) 'DEFAULT
        cmd.Parameters.Append cmd.CreateParameter("@Cash", adDecimal, adParamInput, , Val(Replace(txtCash.Text, ",", "")))
                              cmd.Parameters("@Cash").Precision = 18
                              cmd.Parameters("@Cash").NumericScale = 2
        cmd.Parameters.Append cmd.CreateParameter("@CheckAmount", adDecimal, adParamInput, , Val(Replace(txtCheckAmount.Text, ",", "")))
                              cmd.Parameters("@CheckAmount").Precision = 18
                              cmd.Parameters("@CheckAmount").NumericScale = 2
        cmd.Execute
    End If
    
    con.CommitTrans
    con.Close
    
    UpdateVendorOrderDues
    
    MsgBox "Payment Successful!", vbInformation, "Success!"
    
'    'PRINTOUTS
'    Screen.MousePointer = vbHourglass
'    BASE_PrintPreviewFrm.Show
'    Dim crxApp As New CRAXDRT.Application
'    Dim crxRpt As New CRAXDRT.Report
'    Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\CashCheckVoucher.rpt")
'    crxRpt.RecordSelectionFormula = "{PO_Paymenthistory.Transactionid}= " & TransactionId & ""
'    crxRpt.DiscardSavedData
'
'    Call ResetRptDB(crxRpt)
'
'    crxRpt.ParameterFields(1).AddCurrentValue "CASH/CHECK VOUCHER"
'    crxRpt.ParameterFields(2).AddCurrentValue FIN_AccountsPayable.lvSearch.SelectedItem.SubItems(2)
'    BASE_PrintPreviewFrm.CRViewer.ReportSource = crxRpt
'    BASE_PrintPreviewFrm.CRViewer.ViewReport
'    BASE_PrintPreviewFrm.CRViewer.Zoom 1
'    Screen.MousePointer = vbDefault
'
'    'CHECK TARGET PRINT
'    Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\checkTarget.rpt")
'    'crxRpt.RecordSelectionFormula = "{PO_Paymenthistory.Transactionid}= " & TransactionId & ""
'    crxRpt.DiscardSavedData
'
'    Call ResetRptDB(crxRpt)
'
'
'    crxRpt.ParameterFields(1).AddCurrentValue dtCheckDate.value
'    crxRpt.ParameterFields(2).AddCurrentValue FIN_AccountsPayable.lvSearch.SelectedItem.SubItems(2)
'    crxRpt.ParameterFields(3).AddCurrentValue FormatNumber(Val(Replace(txtCheckAmount.text, ",", "")) + Val(Replace(txtCash.text, ",", "")), 2, vbTrue, vbFalse)
'    crxRpt.ParameterFields(4).AddCurrentValue txtAmountInWords.text
'    crxRpt.PrintOut False
    'BASE_PrintPreviewFrm.CRViewer.ReportSource = crxRpt
    
    FIN_AccountsPayable.btnSearch_Click
    Unload Me
End Sub

Private Sub chkOnline_Click()
    If chkOnline.value = Checked Then
        lblBank.Visible = True
        cmbBank.Visible = True
        lblAccount.Visible = True
        cmbAccount.Visible = True
        txtCash.Enabled = False
    Else
        txtCash.Enabled = True
        lblBank.Visible = False
        cmbBank.Visible = False
        lblAccount.Visible = False
        cmbAccount.Visible = False
    End If
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
    cmbAccount.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            If rec!isActive = "True" Then
                cmbAccount.AddItem rec!accountnumber & " - " & rec!AccountName
                cmbAccount.ItemData(cmbAccount.NewIndex) = rec!AccountId
            End If
            rec.MoveNext
        Loop
    End If
    con.Close
End Sub

Private Sub Form_Load()
    lvOrders.ColumnHeaders(1).width = lvOrders.width * 0.025
    lvOrders.ColumnHeaders(3).width = lvOrders.width * 0.1633
    lvOrders.ColumnHeaders(4).width = lvOrders.width * 0.1633
    lvOrders.ColumnHeaders(5).width = lvOrders.width * 0.1633
    lvOrders.ColumnHeaders(6).width = lvOrders.width * 0.1633
    lvOrders.ColumnHeaders(7).width = lvOrders.width * 0.295
    
    FRE_CASH.Enabled = True
    FRE_CHECK.Enabled = False
    
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandText = "PO_PurchaseOrderPayable_Get"
    cmd.CommandType = adCmdStoredProc
    cmd.Parameters.Append cmd.CreateParameter("@PurchaseOrderId", adInteger, adParamInput, , Null)
    cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInput, 50, Null)
    cmd.Parameters.Append cmd.CreateParameter("@VendorId", adInteger, adParamInput, , VendorId)
    cmd.Parameters.Append cmd.CreateParameter("@Sort", adVarChar, adParamInput, 50, "Date")
    
    Dim item As MSComctlLib.ListItem
    Set rec = cmd.Execute
    If Not rec.EOF Then
        Do Until rec.EOF
            Set item = lvOrders.ListItems.add(, , "")
                item.SubItems(1) = rec!PurchaseOrderId
                item.SubItems(2) = rec!OrderNumber
                item.SubItems(3) = Format(rec!Date, "MM/DD/YY")
                'item.SubItems(4) = Format(rec!DueDate, "MM/DD/YY")
                If Not IsNull(rec!Terms) Then item.SubItems(5) = rec!Terms
                item.SubItems(6) = FormatNumber(rec!OutStandingBalance, 2, vbTrue, vbFalse)
            rec.MoveNext
        Loop
    End If
    
    con.Close
    
    dtDate.value = Format(Now, "MM/DD/YY hh:mm:ss")
    dtCheckDate.value = Format(Now, "MM/DD/YY hh:mm:ss")
    
    Populate "Bank"
    CountTotal
End Sub
Private Sub CountTotal()
    Dim item As MSComctlLib.ListItem
    Dim Total, balance, tax As Double
    Total = 0
    For Each item In lvOrders.ListItems
        balance = balance + Val(Replace(item.SubItems(6), ",", ""))
        If item.Checked = True Then
            Total = Total + Val(Replace(item.SubItems(6), ",", ""))
        End If
    Next
    lblTotal.Caption = "Total Selected: " & FormatNumber(Total, 2, vbTrue, vbFalse)
    lblTotalBalance.Caption = "Total Balance: " & FormatNumber(balance, 2, vbTrue, vbFalse)
    'tax = (total / 1.12) * 0.01
    tax = 0
    'txtTax.text = FormatNumber(((total / 1.12) * 0.01), 2, vbTrue)
    txtTax.Text = "0.00"
    lblPayable.Caption = "Net Payable: " & FormatNumber((Total - Val(Replace(tax, ",", ""))), 2, vbTrue, vbFalse)
End Sub

Private Sub Label2_Click()
    Dim item As MSComctlLib.ListItem
    For Each item In lvOrders.ListItems
        item.Checked = False
    Next
    CountTotal
End Sub

Private Sub lblSelectAll_Click()
    Dim item As MSComctlLib.ListItem
    For Each item In lvOrders.ListItems
        item.Checked = True
    Next
    CountTotal
End Sub

Private Sub lvOrders_ItemCheck(ByVal item As MSComctlLib.ListItem)
    CountTotal
End Sub

Private Sub txtCash_Change()
    If IsNumeric(txtCash.Text) = False Then
        txtCash.Text = "0.00"
        selectText txtCash
    ElseIf Val(Replace(txtCash.Text, ",", "")) + Val(Replace(txtCheckAmount.Text, ",", "")) + Val(Replace(txtTax.Text, ",", "")) > Val(Replace(Replace(lblTotal.Caption, "Total Selected:", ""), ",", "")) Then
        txtCash.Text = FormatNumber(Val(Replace(Replace(lblTotal.Caption, "Total Selected:", ""), ",", "")) - (Val(Replace(txtCheckAmount.Text, ",", "")) + Val(Replace(txtTax.Text, ",", ""))), 2, vbTrue)
    End If
End Sub

Private Sub txtCheckAmount_Change()
    If IsNumeric(txtCheckAmount.Text) = False Then
        txtCheckAmount.Text = "0.00"
        selectText txtCheckAmount
     ElseIf Val(Replace(txtCash.Text, ",", "")) + Val(Replace(txtCheckAmount.Text, ",", "")) + Val(Replace(txtTax.Text, ",", "")) > Val(Replace(Replace(lblTotal.Caption, "Total Selected:", ""), ",", "")) Then
        txtCheckAmount.Text = FormatNumber(Val(Replace(Replace(lblTotal.Caption, "Total Selected:", ""), ",", "")) - (Val(Replace(txtCash.Text, ",", "")) + Val(Replace(txtTax.Text, ",", ""))), 2, vbTrue, vbFalse)
    End If
End Sub

Private Sub txtTax_Change()
    If IsNumeric(txtTax.Text) = False Then
        txtTax.Text = "0.00"
        selectText txtTax
     ElseIf Val(Replace(txtCash.Text, ",", "")) + Val(Replace(txtCheckAmount.Text, ",", "")) + Val(Replace(txtTax.Text, ",", "")) > Val(Replace(Replace(lblTotal.Caption, "Total Selected:", ""), ",", "")) Then
        txtTax.Text = FormatNumber(Val(Replace(Replace(lblTotal.Caption, "Total Selected:", ""), ",", "")) - (Val(Replace(txtCash.Text, ",", "")) + Val(Replace(txtCheckAmount.Text, ",", ""))), 2, vbTrue, vbFalse)
    End If
    lblPayable.Caption = "Total Payable(Taxed):" & FormatNumber(Val(Replace(Replace(lblTotal.Caption, "Total Selected:", ""), ",", "")) - Val(Replace(txtTax.Text, ",", "")), 2, vbTrue, vbFalse)
End Sub

