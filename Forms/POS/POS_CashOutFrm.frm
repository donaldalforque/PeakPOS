VERSION 5.00
Begin VB.Form POS_CashOutFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5550
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnAccept 
      Caption         =   "ENTER: Pay"
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
      Left            =   2160
      Picture         =   "POS_CashOutFrm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3240
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
      Left            =   3840
      Picture         =   "POS_CashOutFrm.frx":23D4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3240
      Width           =   1575
   End
   Begin VB.TextBox txtCash 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      Text            =   "0.00"
      Top             =   1080
      Width           =   3855
   End
   Begin VB.TextBox txtRemarks 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1560
      TabIndex        =   1
      Top             =   1800
      Width           =   3855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CASH  - OUT (-)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   360
      TabIndex        =   6
      Top             =   360
      Width           =   2040
   End
   Begin VB.Label lblCash 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   360
      TabIndex        =   5
      Top             =   1200
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Notes:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   360
      TabIndex        =   4
      Top             =   1800
      Width           =   810
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H008080FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   855
      Left            =   120
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "POS_CashOutFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAccept_Click()
    Dim x As Variant
    x = MsgBox("Are you sure all information are correct?", vbQuestion + vbYesNo)
    If x = vbYes Then
        Dim CashFlowId As Long
        Dim con As New ADODB.Connection
        Set rec = New ADODB.Recordset
        Dim cmd As New ADODB.Command
        
        con.ConnectionString = ConnString
        con.Open
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "POS_CashFlow_Insert"
        cmd.Parameters.Append cmd.CreateParameter("@CashFlowId", adInteger, adParamInputOutput, , 0)
        cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , NVAL(txtCash.Text))
                              cmd.Parameters("@Amount").NumericScale = 2
                              cmd.Parameters("@Amount").Precision = 18
        cmd.Parameters.Append cmd.CreateParameter("@Remarks", adVarChar, adParamInput, 250, txtRemarks.Text)
        cmd.Parameters.Append cmd.CreateParameter("@Type", adVarChar, adParamInput, 250, "CASH-OUT")
        cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
        cmd.Parameters.Append cmd.CreateParameter("@WorkStationId", adInteger, adParamInput, , WorkstationId)
        
        cmd.Execute
        
        CashFlowId = cmd.Parameters("@CashFlowId")
        
        con.Close
        
        MsgBox "Cash-out successful!", vbInformation
        
        x = MsgBox("Do you want to print a receipt?", vbQuestion + vbYesNo)
        If x = vbYes Then
            '**PRINT RECEIPT******
            Dim crxApp As New CRAXDRT.Application
            Dim crxRpt As New CRAXDRT.Report
            Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\POS_CashFlowReceipt.rpt")
            crxRpt.DiscardSavedData
            crxRpt.EnableParameterPrompting = False
            crxRpt.ParameterFields.GetItemByName("Notice").AddCurrentValue "CASH-OUT"
            crxRpt.ParameterFields.GetItemByName("@CashFlowId").AddCurrentValue CashFlowId
            
            Call ResetRptDB(crxRpt)
            crxRpt.PrintOut False
            
            '**END PRINT RECEIPT**
        End If
        
        Unload Me
    End If
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

Private Sub Form_Load()
    selectText txtCash
End Sub

Private Sub txtCash_Change()
    If IsNumeric(txtCash.Text) = False Then
        txtCash.Text = "0"
        selectText txtCash
    End If
End Sub

