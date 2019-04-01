VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form BASE_SalesmanDetailsFrm 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Category Filter"
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
      Left            =   360
      TabIndex        =   9
      Top             =   3000
      Width           =   3975
      Begin VB.ComboBox cmbCategory 
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
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Commission Value"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   3975
      Begin VB.TextBox txtAmount 
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
         Left            =   2760
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtPercentage 
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
         Left            =   2760
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton optAmount 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Amount per sale"
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
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   2055
      End
      Begin VB.OptionButton optPercentage 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Percentage per sale"
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
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Value           =   -1  'True
         Width           =   2055
      End
   End
   Begin MSComctlLib.Toolbar tb_Standard 
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   582
      ButtonWidth     =   1349
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Accounts"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9240
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BASE_SalesmanDetailsFrm.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BASE_SalesmanDetailsFrm.frx":6862
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BASE_SalesmanDetailsFrm.frx":D0C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BASE_SalesmanDetailsFrm.frx":13926
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      X1              =   360
      X2              =   4320
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label lblSalesman 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1320
      TabIndex        =   3
      Top             =   1320
      Width           =   2925
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Salesman:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   360
      TabIndex        =   2
      Top             =   1320
      Width           =   825
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Salesman Commission"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   2565
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   3615
      Left            =   120
      Top             =   480
      Width           =   4455
   End
End
Attribute VB_Name = "BASE_SalesmanDetailsFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public SalesmanCommissionId, SalesmanId As Long
Private Sub Form_Load()
    optPercentage_Click
    Populate "Category"
    Populate "SalesmanCommission"
End Sub

Private Sub optAmount_Click()
    If optAmount.value = True Then
        txtAmount.Enabled = True
        txtAmount.BackColor = vbWhite
        txtPercentage.Enabled = False
        txtPercentage.BackColor = &HC0C0C0
    End If
End Sub

Private Sub optPercentage_Click()
    If optPercentage.value = True Then
        txtPercentage.Enabled = True
        txtPercentage.BackColor = vbWhite
        txtAmount.Enabled = False
        txtAmount.BackColor = &HC0C0C0
    End If
End Sub

Public Sub Populate(ByVal data As String)
    Dim Item As MSComctlLib.ListItem
    Select Case data
        Case "Category"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Category")
            cmbCategory.Clear
            cmbCategory.AddItem ""
            cmbCategory.ItemData(cmbCategory.NewIndex) = 0
            cmbCategory.ListIndex = 0
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        cmbCategory.AddItem rec!Category
                        cmbCategory.ItemData(cmbCategory.NewIndex) = rec!CategoryId
                    End If
                    rec.MoveNext
                Loop
            End If
        Case "SalesmanCommission"
            Dim con As New ADODB.Connection
            con.ConnectionString = ConnString
            con.Open
            Set rec = New ADODB.Recordset
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "BASE_SalesmanCommission_Get"
            cmd.Parameters.Append cmd.CreateParameter("@SalesmanId", adInteger, adParamInput, , SalesmanId)
            Set rec = cmd.Execute
            If Not rec.EOF Then
                SalesmanCommissionId = rec!SalesmanCommissionId
                If Not IsNull(rec!percentage) Then
                    optPercentage.value = True
                    optPercentage_Click
                    txtPercentage.Text = rec!percentage
                End If
                If Not IsNull(rec!Amount) Then
                    optAmount.value = True
                    optAmount_Click
                    txtAmount.Text = rec!Amount
                End If
                If Not IsNull(rec!Category) Then
                    On Error Resume Next
                    cmbCategory.Text = rec!Category
                End If
            End If
            con.Close
    End Select
End Sub

Private Sub tb_Standard_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'NEW
            
        Case 2 'Save
            Dim con As New ADODB.Connection
            Set cmd = New ADODB.Command
            Set rec = New ADODB.Recordset
            
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.Parameters.Append cmd.CreateParameter("@SalesmanCommissionId", adInteger, adParamInputOutput, , SalesmanCommissionId)
            cmd.Parameters.Append cmd.CreateParameter("@SalesmanId", adInteger, adParamInput, , SalesmanId)
            
            If optPercentage.value = True Then
                cmd.Parameters.Append cmd.CreateParameter("@Percentage", adDecimal, adParamInput, , NVAL(txtPercentage.Text))
            Else
                cmd.Parameters.Append cmd.CreateParameter("@Percentage", adDecimal, adParamInput, , Null)
            End If
            If optAmount.value = True Then
                cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , NVAL(txtAmount.Text))
            Else
                cmd.Parameters.Append cmd.CreateParameter("@Amount", adDecimal, adParamInput, , Null)
            End If
            
            cmd.Parameters.Append cmd.CreateParameter("@CategoryId", adInteger, adParamInput, , cmbCategory.ItemData(cmbCategory.ListIndex))
            cmd.Parameters("@Percentage").Precision = 18
            cmd.Parameters("@Percentage").NumericScale = 2
            cmd.Parameters("@Amount").Precision = 18
            cmd.Parameters("@Amount").NumericScale = 2
            
            cmd.CommandText = "BASE_SalesmanCommission_Update"
            cmd.Execute
            con.Close
            MsgBox "Record saved.", vbInformation
    End Select
    Exit Sub
ErrorHandler:
    If IsNumeric(Err.Description) = True Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(Val(Err.Description))
    Else
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & Err.Description
    End If
    GLOBAL_MessageFrm.Show (1)
End Sub
