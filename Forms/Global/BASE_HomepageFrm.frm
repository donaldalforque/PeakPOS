VERSION 5.00
Begin VB.Form BASE_HomepageFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Homepage"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15090
   Icon            =   "BASE_HomepageFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   15090
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FRE_Body 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   8775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11295
      Begin VB.Label lblExpenses 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Manage Expenses"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   480
         MouseIcon       =   "BASE_HomepageFrm.frx":000C
         MousePointer    =   99  'Custom
         TabIndex        =   16
         Top             =   7440
         Width           =   2505
      End
      Begin VB.Image imgExpenses 
         Height          =   720
         Left            =   1080
         MouseIcon       =   "BASE_HomepageFrm.frx":015E
         MousePointer    =   99  'Custom
         Picture         =   "BASE_HomepageFrm.frx":02B0
         Top             =   6600
         Width           =   720
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Manage your Finance"
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   12.75
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   240
         TabIndex        =   15
         Top             =   6120
         Width           =   2265
      End
      Begin VB.Label lblAccountsReceivable 
         BackColor       =   &H00FFFFFF&
         Caption         =   "View Customer Credits"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2520
         MouseIcon       =   "BASE_HomepageFrm.frx":34F3
         MousePointer    =   99  'Custom
         TabIndex        =   14
         Top             =   7440
         Width           =   2505
      End
      Begin VB.Image imgAccountsReceivable 
         Height          =   720
         Left            =   3120
         MouseIcon       =   "BASE_HomepageFrm.frx":3645
         MousePointer    =   99  'Custom
         Picture         =   "BASE_HomepageFrm.frx":3797
         Top             =   6600
         Width           =   720
      End
      Begin VB.Label lblNewCustomer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Manage Customers"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6120
         MouseIcon       =   "BASE_HomepageFrm.frx":69F3
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   5040
         Width           =   2505
      End
      Begin VB.Image imgNewCustomer 
         Height          =   720
         Left            =   6600
         MouseIcon       =   "BASE_HomepageFrm.frx":6B45
         MousePointer    =   99  'Custom
         Picture         =   "BASE_HomepageFrm.frx":6C97
         Top             =   4200
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sell Products"
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   12.75
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   3360
         TabIndex        =   12
         Top             =   3600
         Width           =   1350
      End
      Begin VB.Image imgSalesOrder 
         Height          =   720
         Left            =   3960
         MouseIcon       =   "BASE_HomepageFrm.frx":A11D
         MousePointer    =   99  'Custom
         Picture         =   "BASE_HomepageFrm.frx":A26F
         Top             =   4200
         Width           =   720
      End
      Begin VB.Label lblSalesOrder 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Create Sales Orders"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3480
         MouseIcon       =   "BASE_HomepageFrm.frx":D536
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   5040
         Width           =   2505
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Accept Deliveries"
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   12.75
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   240
         TabIndex        =   10
         Top             =   3600
         Width           =   1755
      End
      Begin VB.Image imgPurchaseOrder 
         Height          =   720
         Left            =   1080
         MouseIcon       =   "BASE_HomepageFrm.frx":D688
         MousePointer    =   99  'Custom
         Picture         =   "BASE_HomepageFrm.frx":D7DA
         Top             =   4200
         Width           =   720
      End
      Begin VB.Label lblPurchaseOrder 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Issue Purchase Orders"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         MouseIcon       =   "BASE_HomepageFrm.frx":108C6
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   5040
         Width           =   2505
      End
      Begin VB.Label lblProductList 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Check your products"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6000
         MouseIcon       =   "BASE_HomepageFrm.frx":10A18
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   2400
         Width           =   2025
      End
      Begin VB.Image imgProductList 
         Height          =   720
         Left            =   6600
         MouseIcon       =   "BASE_HomepageFrm.frx":10B6A
         MousePointer    =   99  'Custom
         Picture         =   "BASE_HomepageFrm.frx":10CBC
         Top             =   1560
         Width           =   720
      End
      Begin VB.Image imgCategories 
         Height          =   720
         Left            =   3840
         MouseIcon       =   "BASE_HomepageFrm.frx":13B98
         MousePointer    =   99  'Custom
         Picture         =   "BASE_HomepageFrm.frx":13CEA
         Top             =   1560
         Width           =   720
      End
      Begin VB.Label lblCategories 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Organize product with categories "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3240
         MouseIcon       =   "BASE_HomepageFrm.frx":16F3B
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   2400
         Width           =   2025
      End
      Begin VB.Label lblNewProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Create new product"
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
         Left            =   600
         MouseIcon       =   "BASE_HomepageFrm.frx":1708D
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   2400
         Width           =   1830
      End
      Begin VB.Image imgNewProduct 
         Height          =   720
         Left            =   1080
         MouseIcon       =   "BASE_HomepageFrm.frx":171DF
         MousePointer    =   99  'Custom
         Picture         =   "BASE_HomepageFrm.frx":17331
         Top             =   1560
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Manage your Inventory"
         BeginProperty Font 
            Name            =   "Calibri Light"
            Size            =   12.75
            Charset         =   0
            Weight          =   300
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   2430
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hello, What would you like to do?"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   4410
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   11800
      MouseIcon       =   "BASE_HomepageFrm.frx":1A531
      MousePointer    =   99  'Custom
      Picture         =   "BASE_HomepageFrm.frx":1A683
      Top             =   2280
      Width           =   480
   End
   Begin VB.Label lblNotifications 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Notifications"
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
      Left            =   12480
      MouseIcon       =   "BASE_HomepageFrm.frx":1AC99
      MousePointer    =   99  'Custom
      TabIndex        =   31
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label lblSystemSettings 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "System Settings"
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
      Left            =   12480
      MouseIcon       =   "BASE_HomepageFrm.frx":1ADEB
      MousePointer    =   99  'Custom
      TabIndex        =   30
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Image imgSystemSettings 
      Height          =   480
      Left            =   11760
      MouseIcon       =   "BASE_HomepageFrm.frx":1AF3D
      MousePointer    =   99  'Custom
      Picture         =   "BASE_HomepageFrm.frx":1B08F
      Top             =   1560
      Width           =   480
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "BANK OUT"
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
      Height          =   345
      Left            =   11760
      MouseIcon       =   "BASE_HomepageFrm.frx":1B659
      MousePointer    =   99  'Custom
      TabIndex        =   29
      Top             =   7680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblBankin 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "1,230,500.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   11760
      MouseIcon       =   "BASE_HomepageFrm.frx":1B7AB
      MousePointer    =   99  'Custom
      TabIndex        =   28
      Top             =   7200
      Visible         =   0   'False
      Width           =   3195
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "BANK IN"
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
      Height          =   345
      Left            =   11760
      MouseIcon       =   "BASE_HomepageFrm.frx":1B8FD
      MousePointer    =   99  'Custom
      TabIndex        =   27
      Top             =   6840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblbankout 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "1,230,500.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   11760
      MouseIcon       =   "BASE_HomepageFrm.frx":1BA4F
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   8040
      Visible         =   0   'False
      Width           =   3195
   End
   Begin VB.Label lblCheckOut 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "1,230,500.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   11760
      MouseIcon       =   "BASE_HomepageFrm.frx":1BBA1
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   6000
      Visible         =   0   'False
      Width           =   3195
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "CHECK IN"
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
      Height          =   345
      Left            =   11760
      MouseIcon       =   "BASE_HomepageFrm.frx":1BCF3
      MousePointer    =   99  'Custom
      TabIndex        =   24
      Top             =   4800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblCheckIn 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "1,230,500.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   11760
      MouseIcon       =   "BASE_HomepageFrm.frx":1BE45
      MousePointer    =   99  'Custom
      TabIndex        =   23
      Top             =   5160
      Visible         =   0   'False
      Width           =   3195
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "CHECK OUT"
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
      Height          =   345
      Left            =   11760
      MouseIcon       =   "BASE_HomepageFrm.frx":1BF97
      MousePointer    =   99  'Custom
      TabIndex        =   22
      Top             =   5640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblCashOut 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "1,230,500.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   11760
      MouseIcon       =   "BASE_HomepageFrm.frx":1C0E9
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Top             =   4800
      Visible         =   0   'False
      Width           =   3195
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "CASH OUT"
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
      Height          =   345
      Left            =   11760
      MouseIcon       =   "BASE_HomepageFrm.frx":1C23B
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   4440
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label lblCashIn 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "1,230,500.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   11760
      MouseIcon       =   "BASE_HomepageFrm.frx":1C38D
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   3960
      Visible         =   0   'False
      Width           =   3195
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "CASH IN"
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
      Height          =   345
      Left            =   11760
      MouseIcon       =   "BASE_HomepageFrm.frx":1C4DF
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   3600
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Summary"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   465
      Left            =   11760
      TabIndex        =   17
      Top             =   3120
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Quick Launch"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   465
      Left            =   11760
      TabIndex        =   4
      Top             =   240
      Width           =   3255
   End
   Begin VB.Image imgGeneralSettings 
      Height          =   480
      Left            =   11760
      MouseIcon       =   "BASE_HomepageFrm.frx":1C631
      MousePointer    =   99  'Custom
      Picture         =   "BASE_HomepageFrm.frx":1C783
      Top             =   960
      Width           =   480
   End
   Begin VB.Label lblGeneralSettings 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "General Settings"
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
      Left            =   12480
      MouseIcon       =   "BASE_HomepageFrm.frx":1CDA7
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label FRE_RightPane 
      BackColor       =   &H00E0E0E0&
      Height          =   9015
      Left            =   11540
      TabIndex        =   2
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "BASE_HomepageFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub ComputeTransactions()
'    Set con = New ADODB.Connection
'    Set rec = New ADODB.Recordset
'    Set cmd = New ADODB.Command
'
'    con.ConnectionString = ConnString
'    con.Open
'    cmd.ActiveConnection = con
'    cmd.CommandType = adCmdText
'    cmd.CommandText = "SELECT ISNULL(SUM(CASH),0) as CASHIN FROM FIN_Inflow  " & _
'                      "WHERE DATEDIFF(day,getdate(),FIN_Inflow.Date) = 0"
'    Set rec = cmd.Execute
'
'    'CASH
'    If Not rec.EOF Then
'        lblCashIn.Caption = FormatNumber(rec!cashin, 2, vbTrue)
'    End If
'
'    cmd.CommandText = "SELECT ISNULL(SUM(CASH),0) as CASHOUT FROM FIN_Outflow  " & _
'                      "WHERE DATEDIFF(day,getdate(),FIN_Outflow.Date) = 0"
'    Set rec = cmd.Execute
'
'    If Not rec.EOF Then
'        lblCashOut.Caption = FormatNumber(rec!cashout, 2, vbTrue)
'    End If
'
'    'CHECK
'    cmd.CommandText = "SELECT ISNULL(SUM(CHECKAMOUNT),0) as checkin FROM fin_inflow  " & _
'                      "WHERE DATEDIFF(day,getdate(),fin_inflow.Date) = 0"
'    Set rec = cmd.Execute
'
'    If Not rec.EOF Then
'        lblCheckIn.Caption = FormatNumber(rec!checkin, 2, vbTrue)
'    End If
'
'    cmd.CommandText = "SELECT ISNULL(SUM(CHECKAMOUNT),0) as checkout FROM FIN_OutFlow  " & _
'                      "WHERE DATEDIFF(day,getdate(),FIN_OutFlow.Date) = 0"
'    Set rec = cmd.Execute
'
'    If Not rec.EOF Then
'        lblCheckOut.Caption = FormatNumber(rec!CHECKOUT, 2, vbTrue)
'    End If
'
'    'BANK
'    cmd.CommandText = "SELECT ISNULL((SUM(CHECKAMOUNT)+SUM(CASH)),0) as bankin FROM fin_Bankinflow  " & _
'                      "WHERE DATEDIFF(day,getdate(),fin_Bankinflow.Date) = 0"
'    Set rec = cmd.Execute
'
'    If Not rec.EOF Then
'        lblBankin.Caption = FormatNumber(rec!bankin, 2, vbTrue)
'    End If
'
'    cmd.CommandText = "SELECT ISNULL((SUM(CHECKAMOUNT)+SUM(CASH)),0) as bankout FROM FIN_bankOutFlow  " & _
'                      "WHERE DATEDIFF(day,getdate(),FIN_bankOutFlow.Date) = 0"
'    Set rec = cmd.Execute
'
'    If Not rec.EOF Then
'        lblbankout.Caption = FormatNumber(rec!Bankout, 2, vbTrue)
'    End If
'
'    con.Close
End Sub
Private Sub Form_Activate()
    ComputeTransactions
End Sub

Private Sub Form_Load()
    lblCashOut.Caption = "0.00"
    lblCashIn.Caption = "0.00"
    lblCheckIn.Caption = "0.00"
    lblCheckOut.Caption = "0.00"
    lblbankout.Caption = "0.00"
    lblBankin.Caption = "0.00"

    ComputeTransactions
    ShowNotification
    
End Sub

Private Sub Image1_Click()
    BASE_GeneralSettingsFrm.Show (1)
End Sub

Private Sub Image10_Click()
    CenterChildForm FIN_ExpensesFrm
    FIN_ExpensesFrm.Show
    FIN_ExpensesFrm.ZOrder 0
End Sub

Private Sub Image11_Click()
    CenterChildForm FIN_AccountFundingFrm
    FIN_AccountFundingFrm.Show
    FIN_AccountFundingFrm.ZOrder 0
End Sub

Private Sub Image2_Click()
    CornerChildForm INV_NewProductFrm
            INV_NewProductFrm.Show
            INV_NewProductFrm.ZOrder 0
End Sub

Private Sub Image3_Click()
    
End Sub

Private Sub Image4_Click()
    CornerChildForm INV_ProductListFrm
    INV_ProductListFrm.Show
    INV_ProductListFrm.ZOrder 0
End Sub

Private Sub Image5_Click()
    CornerChildForm PO_PurchaseOrderFrm
    PO_PurchaseOrderFrm.Show
    PO_PurchaseOrderFrm.ZOrder 0
End Sub

Private Sub imgVendor_Click()

End Sub

Private Sub FRE_RightPane_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblGeneralSettings.FontUnderline = False
    lblSystemSettings.FontUnderline = False
    lblNotifications.FontUnderline = False
End Sub

Private Sub Image6_Click()
    CornerChildForm SO_SalesOrderFrm
    SO_SalesOrderFrm.Show
    SO_SalesOrderFrm.ZOrder 0
End Sub

Private Sub Image7_Click()
    CornerChildForm SO_CustomerFrm
    SO_CustomerFrm.Show
    SO_CustomerFrm.ZOrder 0
End Sub

Private Sub Image8_Click()
    CornerChildForm FIN_AccountsReceivable
    FIN_AccountsReceivable.Show
    FIN_AccountsReceivable.ZOrder 0
End Sub

Private Sub Image9_Click()
    FIN_BanksFrm.Show (1)
End Sub

Private Sub imgAccountsPayable_Click()
   
End Sub

Private Sub imgAccountsReceivable_Click()
     CornerChildForm FIN_AccountsReceivable
    FIN_AccountsReceivable.Show
    FIN_AccountsReceivable.ZOrder 0
End Sub

Private Sub imgCategories_Click()
    CenterChildForm INV_CategoryModFrm
    INV_CategoryModFrm.Show (1)
    'INV_CategoryModFrm.ZOrder 0
End Sub

Private Sub imgExpenses_Click()
    CenterChildForm FIN_ExpensesFrm
    FIN_ExpensesFrm.Show
    FIN_ExpensesFrm.ZOrder 0
End Sub

Private Sub imgGeneralSettings_Click()
    If EditAccessRights(24) = False Then
        MsgBox ErrorCodes(74), vbCritical, "Limited Rights"
        Exit Sub
    End If
    BASE_GeneralSettingsFrm.Show (1)
End Sub

Private Sub imgNewCustomer_Click()
    CornerChildForm SO_CustomerFrm
    SO_CustomerFrm.Show
    SO_CustomerFrm.ZOrder 0
End Sub

Private Sub imgNewProduct_Click()
    CornerChildForm INV_NewProductFrm
    INV_NewProductFrm.Show
    INV_NewProductFrm.ZOrder 0
End Sub

Private Sub imgProductList_Click()
    CornerChildForm INV_ProductListFrm
    INV_ProductListFrm.Show
    INV_ProductListFrm.ZOrder 0
End Sub

Private Sub imgPurchaseOrder_Click()
    CornerChildForm PO_PurchaseOrderFrm
    PO_PurchaseOrderFrm.Show
    PO_PurchaseOrderFrm.ZOrder 0
End Sub

Private Sub imgSalesOrder_Click()
    CornerChildForm SO_SalesOrderFrm
    SO_SalesOrderFrm.Show
    SO_SalesOrderFrm.ZOrder 0
End Sub

Private Sub Label11_Click()
    CornerChildForm SO_CustomerFrm
    SO_CustomerFrm.Show
    SO_CustomerFrm.ZOrder 0
End Sub

Private Sub Label12_Click()
    
    CornerChildForm FIN_AccountsReceivable
    FIN_AccountsReceivable.Show
    FIN_AccountsReceivable.ZOrder 0
End Sub

Private Sub Label14_Click()
    FIN_BanksFrm.Show (1)
End Sub

Private Sub Label15_Click()
    CenterChildForm FIN_ExpensesFrm
    FIN_ExpensesFrm.Show
    FIN_ExpensesFrm.ZOrder 0
End Sub

Private Sub Label16_Click()
    CenterChildForm FIN_AccountFundingFrm
    FIN_AccountFundingFrm.Show
    FIN_AccountFundingFrm.ZOrder 0
End Sub

Private Sub Label23_Click()

End Sub

Private Sub Label4_Click()
    CornerChildForm INV_NewProductFrm
    INV_NewProductFrm.Show
    INV_NewProductFrm.ZOrder 0
End Sub

Private Sub Label5_Click()
    CenterChildForm INV_CategoryModFrm
    INV_CategoryModFrm.Show (1)
End Sub

Private Sub Label6_Click()
    
End Sub

Private Sub Label7_Click()
    CornerChildForm PO_PurchaseOrderFrm
    PO_PurchaseOrderFrm.Show
    PO_PurchaseOrderFrm.ZOrder 0
End Sub

Private Sub Label9_Click()
    CornerChildForm SO_SalesOrderFrm
    SO_SalesOrderFrm.Show
    SO_SalesOrderFrm.ZOrder 0
End Sub

Private Sub lblAccountsPayable_Click()
    CornerChildForm FIN_AccountsReceivable
    FIN_AccountsReceivable.Show
    FIN_AccountsReceivable.ZOrder 0
End Sub

Private Sub imgSystemSettings_Click()
    If EditAccessRights(25) = False Then
        MsgBox ErrorCodes(74), vbCritical, "Limited Rights"
        Exit Sub
    End If
    BASE_SystemSettingsFrm.Show (1)
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
End Sub

Private Sub lblAccountsReceivable_Click()
     CornerChildForm FIN_AccountsReceivable
    FIN_AccountsReceivable.Show
    FIN_AccountsReceivable.ZOrder 0
End Sub

Private Sub lblCategories_Click()
    CenterChildForm INV_CategoryModFrm
    INV_CategoryModFrm.Show (1)
    'INV_CategoryModFrm.ZOrder 0
End Sub

Private Sub lblExpenses_Click()
    CenterChildForm FIN_ExpensesFrm
    FIN_ExpensesFrm.Show
    FIN_ExpensesFrm.ZOrder 0
End Sub

Private Sub lblGeneralSettings_Click()
    If EditAccessRights(24) = False Then
        MsgBox ErrorCodes(74), vbCritical, "Limited Rights"
        Exit Sub
    End If
    BASE_GeneralSettingsFrm.Show '(1)
End Sub

Private Sub lblGeneralSettings_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If x >= 0 And x <= lblGeneralSettings.width Then
        lblGeneralSettings.FontUnderline = True
    End If
End Sub

Private Sub lblNewCustomer_Click()
    CornerChildForm SO_CustomerFrm
    SO_CustomerFrm.Show
    SO_CustomerFrm.ZOrder 0
End Sub

Private Sub lblNewProduct_Click()
    CornerChildForm INV_NewProductFrm
    INV_NewProductFrm.Show
    INV_NewProductFrm.ZOrder 0
End Sub

Private Sub lblNotifications_Click()
    ShowNotification
End Sub

Private Sub lblNotifications_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If x >= 0 And x <= lblNotifications.width Then
        lblNotifications.FontUnderline = True
    End If
End Sub

Private Sub lblProductList_Click()
    CornerChildForm INV_ProductListFrm
    INV_ProductListFrm.Show
    INV_ProductListFrm.ZOrder 0
End Sub

Private Sub lblPurchaseOrder_Click()
    CornerChildForm PO_PurchaseOrderFrm
    PO_PurchaseOrderFrm.Show
    PO_PurchaseOrderFrm.ZOrder 0
End Sub

Private Sub lblSalesOrder_Click()
    CornerChildForm SO_SalesOrderFrm
    SO_SalesOrderFrm.Show
    SO_SalesOrderFrm.ZOrder 0
End Sub

Private Sub lblSystemSettings_Click()
    If EditAccessRights(25) = False Then
        MsgBox ErrorCodes(74), vbCritical, "Limited Rights"
        Exit Sub
    End If
    BASE_SystemSettingsFrm.Show (1)
End Sub

Private Sub lblSystemSettings_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If x >= 0 And x <= lblSystemSettings.width Then
        lblSystemSettings.FontUnderline = True
    End If
End Sub
