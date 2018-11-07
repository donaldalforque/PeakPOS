VERSION 5.00
Begin VB.Form INV_AuditStockOptFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4815
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSystem 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Height          =   330
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   3
      Text            =   "1"
      Top             =   1320
      Width           =   3135
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
      Left            =   3360
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "OK"
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
      Left            =   2040
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtPhysical 
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
      Left            =   1440
      MaxLength       =   50
      TabIndex        =   0
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Label lblAdjustment 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3480
      TabIndex        =   9
      Top             =   2160
      Width           =   1140
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Adjustment"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "System and Physical Quantity"
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
      Left            =   240
      TabIndex        =   7
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "This windows shows you the current system quantity and allows you to input the actual physical quantity."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   4485
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "System Qty"
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
      TabIndex        =   5
      Top             =   1320
      Width           =   1050
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   4560
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Physical Qty"
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
      TabIndex        =   4
      Top             =   1680
      Width           =   1125
   End
End
Attribute VB_Name = "INV_AuditStockOptFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function GetSystemQuantity(ByVal ProductId As Integer, ByVal LocationId As Integer) As Double
    Dim con As New ADODB.Connection
    con.ConnectionString = ConnString
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "INV_StockAuditSysQuantity_Get"
    cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , ProductId)
    cmd.Parameters.Append cmd.CreateParameter("@LocationId", adInteger, adParamInput, , LocationId)
    If Not rec.EOF Then GetSystemQuantity = rec!Quantity
    con.Close
End Function

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub txtCost_Change()

End Sub

Private Sub btnOk_Click()
    INV_AuditStockFrm.lvItems.SelectedItem.SubItems(7) = FormatNumber(NVAL(txtPhysical.text), 2, vbTrue, vbFalse)
    INV_AuditStockFrm.lvItems.SelectedItem.SubItems(8) = FormatNumber(NVAL(lblAdjustment.Caption), 2, vbTrue, vbFalse)
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            btnOk_Click
        Case vbKeyCancel
            btnCancel_Click
    End Select
End Sub

Private Sub txtPhysical_Change()
    If IsNumeric(txtPhysical.text) = False And txtPhysical.text <> "" Then
        txtPhysical.text = txtSystem.text
    Else
        'COMPUTE
        Dim adjustment As Double
        adjustment = NVAL(txtPhysical.text) - NVAL(txtSystem.text)
        lblAdjustment.Caption = FormatNumber(adjustment, 2, vbTrue, vbFalse)
    End If
End Sub
