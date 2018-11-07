VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form PO_ReceiveOrderDialogFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   Icon            =   "PO_ReceiveOrderDialogFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox a 
      Height          =   285
      Left            =   -9999
      TabIndex        =   14
      Top             =   1920
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox cmbUnit 
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
      Top             =   720
      Width           =   4455
   End
   Begin VB.TextBox txtDescription 
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
      TabIndex        =   2
      Top             =   1080
      Width           =   4455
   End
   Begin VB.TextBox txtQuantity 
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
      TabIndex        =   0
      Text            =   "1.00"
      Top             =   360
      Width           =   4455
   End
   Begin VB.TextBox txtReference 
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
      TabIndex        =   6
      Top             =   2880
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ComboBox cmbLocation 
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
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   2655
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
      Left            =   4680
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton btnSave 
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
      Left            =   3240
      TabIndex        =   3
      Top             =   1800
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker dtDate 
      Height          =   345
      Left            =   1440
      TabIndex        =   7
      Top             =   3240
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      Format          =   91750401
      CurrentDate     =   41686
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit"
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
      TabIndex        =   13
      Top             =   720
      Width           =   390
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      TabIndex        =   12
      Top             =   1080
      Width           =   1065
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
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
      TabIndex        =   11
      Top             =   360
      Width           =   810
   End
   Begin VB.Label Label2 
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
      TabIndex        =   10
      Top             =   2880
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reference"
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
      TabIndex        =   9
      Top             =   2880
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Location"
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
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   1575
      Left            =   120
      Top             =   120
      Width           =   5895
   End
End
Attribute VB_Name = "PO_ReceiveOrderDialogFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Public isModify As Boolean
Dim Conversion As Double
Public Sub Populate(ByVal data As String)
    Dim item As MSComctlLib.ListItem
    Select Case data
        Case "Location"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Location")
            cmbLocation.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        cmbLocation.AddItem rec!Location
                        cmbLocation.ItemData(cmbLocation.NewIndex) = rec!LocationId
                    End If
                    rec.MoveNext
                Loop
            End If
            On Error Resume Next
            Dim x As Integer
            For x = 1 To cmbLocation.ListCount
                If cmbLocation.ItemData(x) = 2 Then 'warehouse
                    cmbLocation.ListIndex = x
                    Exit For
                End If
            Next
    End Select
End Sub
Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnSave_Click()
    Dim item As MSComctlLib.ListItem
    If isModify = False Then
        With PO_ReceiveOrderFrm
            Set item = .lvItems.ListItems.add(, , "")
                item.SubItems(1) = ""
                item.SubItems(2) = .lvItemList.SelectedItem.Text
                item.SubItems(3) = .lvItemList.SelectedItem.SubItems(1)
                item.SubItems(4) = txtDescription.Text '.lvItemList.SelectedItem.SubItems(2)
                item.SubItems(5) = FormatNumber(txtQuantity.Text, 2, vbTrue)
                item.SubItems(6) = cmbUnit.Text '.lvItemList.SelectedItem.SubItems(4)
                item.SubItems(7) = cmbLocation.Text
                item.SubItems(8) = dtDate.value
                item.SubItems(9) = txtReference.Text
                item.SubItems(10) = cmbLocation.ItemData(cmbLocation.ListIndex)
                item.SubItems(14) = .lvItemList.SelectedItem.SubItems(5)
        Unload Me
        .txtItemSearch.SetFocus
        .lvItemList.Visible = False
        End With
    Else
        With PO_ReceiveOrderFrm
            .lvItems.SelectedItem.SubItems(5) = txtQuantity.Text
            .lvItems.SelectedItem.SubItems(6) = cmbUnit.Text
            .lvItems.SelectedItem.SubItems(8) = Format(dtDate.value, "MM/DD/YY")
            .lvItems.SelectedItem.SubItems(9) = txtReference.Text
            .lvItems.SelectedItem.SubItems(4) = txtDescription.Text
            
            On Error Resume Next
            .lvItems.SelectedItem.SubItems(7) = cmbLocation.Text
            .lvItems.SelectedItem.SubItems(11) = cmbLocation.ItemData(cmbLocation.ListIndex)
            Unload Me
        End With
    End If
End Sub

Private Sub cmbUnit_Click()
    If isModify = False Then
        Conversion = GetProductConversion(PO_ReceiveOrderFrm.lvItemList.SelectedItem.Text, cmbUnit.ItemData(cmbUnit.ListIndex), "Cost", a)
        PO_ReceiveOrderFrm.lvItemList.SelectedItem.SubItems(5) = a.Text
    Else
        Conversion = GetProductConversion(PO_ReceiveOrderFrm.lvItems.SelectedItem.SubItems(2), cmbUnit.ItemData(cmbUnit.ListIndex), "Cost")
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            btnSave_Click
        Case vbKeyEscape
            btnCancel_Click
    End Select
End Sub

Private Sub Form_Load()

     'Get Uom Related
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "INV_UomConversion_Get"
    
    If isModify = True Then
        cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , PO_ReceiveOrderFrm.lvItems.SelectedItem.SubItems(2))
    Else
        cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , PO_ReceiveOrderFrm.lvItemList.SelectedItem.Text)
    End If
    
    Set rec = cmd.Execute
    'lvUom.ListItems.Clear
    cmbUnit.Clear
    Dim item As MSComctlLib.ListItem
    If Not rec.EOF Then
        Do Until rec.EOF
            cmbUnit.AddItem rec!Uom
            cmbUnit.ItemData(cmbUnit.NewIndex) = rec!UomId
            rec.MoveNext
        Loop
    End If
    con.Close
    
    'On Error Resume Next
    cmbUnit.ListIndex = 0

    Populate "Location"
    dtDate.value = Format(Now, "MM/DD/YY")
    txtReference.Text = PO_ReceiveOrderFrm.txtDeliveryNumber.Text
    cmbLocation.Text = PO_ReceiveOrderFrm.cmbLocation.Text
End Sub

Private Sub txtDescription_GotFocus()
    selectText txtDescription
End Sub

Private Sub txtQuantity_Change()
    If IsNumeric(txtQuantity.Text) = False Then
        txtQuantity.Text = "1.00"
    End If
End Sub

Private Sub txtQuantity_GotFocus()
    selectText txtQuantity
End Sub

Private Sub txtReference_GotFocus()
    selectText txtReference
End Sub
