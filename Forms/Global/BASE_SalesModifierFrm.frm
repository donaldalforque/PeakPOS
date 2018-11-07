VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form BASE_SalesModifierFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnModify 
      BackColor       =   &H008080FF&
      Caption         =   "PROCESS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4800
      Width           =   4455
   End
   Begin VB.CommandButton btnAdd 
      Caption         =   "ADD"
      Height          =   330
      Left            =   3120
      TabIndex        =   7
      Top             =   1080
      Width           =   1455
   End
   Begin VB.TextBox txtOrderNumber 
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
      Left            =   960
      TabIndex        =   5
      Top             =   1080
      Width           =   2055
   End
   Begin MSComCtl2.DTPicker DateTo 
      Height          =   345
      Left            =   1320
      TabIndex        =   0
      Top             =   480
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   95682561
      CurrentDate     =   41686
   End
   Begin MSComCtl2.DTPicker DateFrom 
      Height          =   345
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   95682561
      CurrentDate     =   41686
   End
   Begin MSComctlLib.ListView lvItems 
      Height          =   3135
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Order #"
         Object.Width           =   7444
      EndProperty
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order #"
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
      TabIndex        =   6
      Top             =   1080
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date From"
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
      TabIndex        =   3
      Top             =   120
      Width           =   960
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date To"
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
      TabIndex        =   2
      Top             =   480
      Width           =   705
   End
End
Attribute VB_Name = "BASE_SalesModifierFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAdd_Click()
    If Trim(txtOrderNumber.Text) = "" Then
        MsgBox "Cannot add empty value.", vbCritical
        txtOrderNumber.SetFocus
    Else
        Dim item As MSComctlLib.ListItem
        Set item = lvItems.ListItems.add(, , txtOrderNumber.Text)
        txtOrderNumber.Text = ""
        txtOrderNumber.SetFocus
    End If
End Sub

Private Sub btnModify_Click()
    Dim x As Variant
    x = MsgBox("WARNING! This process cannot be reverted. Proceed with caution. Continue?", vbExclamation + vbOKCancel)
    If x = vbOK Then
        Dim con As New ADODB.Connection
        con.ConnectionString = ConnString
        con.Open
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdText
        cmd.CommandText = "DELETE FROM SYS_SM_ProductId"
        cmd.Execute
        
        Dim item As MSComctlLib.ListItem
        For Each item In lvItems.ListItems
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "SYS_SM_GetProductId"
            cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInput, 50, item.Text)
            cmd.Execute
        Next
        
        Dim sql As String
        sql = "DELETE FROM POS_Sales_Line WHERE ProductId IN " & _
              "(SELECT ProductId FROM SYS_SM_ProductId) AND " & _
              "POS_SalesId IN (SELECT POS_SalesId FROM POS_Sales WHERE Date BETWEEN '" & DateFrom.value & " 00:00:00' AND " & _
              "'" & DateTo.value & " 23:59:59')"
              
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdText
        cmd.CommandText = sql
        cmd.Execute
        
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "SYS_SM_Check"
        cmd.Execute
        
        con.Close
        
        MsgBox "Magic complete! The tool will now close.", vbInformation
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    DateFrom.value = Format(Now, "mm/dd/yy")
    DateTo.value = Format(Now, "mm/dd/yy")
End Sub

Private Sub lvItems_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If lvItems.ListItems.Count <= 0 Then Exit Sub
        lvItems.ListItems.Remove (lvItems.SelectedItem.Index)
    End If
End Sub
