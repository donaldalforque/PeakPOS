VERSION 5.00
Begin VB.Form POS_ItemDetailsFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6855
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCancel 
      Caption         =   "ESC: Close"
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
      Left            =   5160
      Picture         =   "POS_ItemDetailsFrm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "CODE:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   8
      Top             =   5280
      Width           =   1185
   End
   Begin VB.Label txtCode 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1320
      TabIndex        =   7
      Top             =   5280
      Width           =   5955
   End
   Begin VB.Image picMain 
      BorderStyle     =   1  'Fixed Single
      Height          =   4935
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6615
   End
   Begin VB.Label lblCost 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1320
      TabIndex        =   6
      Top             =   6720
      Width           =   5955
   End
   Begin VB.Label lblPrice 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1320
      TabIndex        =   5
      Top             =   6240
      Width           =   5955
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1320
      TabIndex        =   4
      Top             =   5760
      Width           =   5955
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   6720
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "COST:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   2
      Top             =   6720
      Width           =   1245
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PRICE:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   1
      Top             =   6240
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   0
      Top             =   5760
      Width           =   1185
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6720
      Y1              =   5160
      Y2              =   5160
   End
End
Attribute VB_Name = "POS_ItemDetailsFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ProductId As Integer

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            btnCancel_Click
    End Select
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim con As New ADODB.Connection
    con.ConnectionString = ConnString
    con.Open
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_Product_Get"
    cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , ProductId)
    Set rec = cmd.Execute
    If Not rec.EOF Then
        txtCode.Caption = rec!itemcode
        lblName.Caption = rec!Name
        lblPrice.Caption = FormatNumber(rec!unitprice, 2, vbTrue, vbFalse)
        lblCost.Caption = FormatNumber(rec!unitcost, 2, vbTrue, vbFalse)
        picMain.Picture = LoadPicture(rec!Image)
        If AccessRights(2, 1) = False Then
            lblCost.Caption = "******"
        End If
    End If
    con.Close
End Sub

