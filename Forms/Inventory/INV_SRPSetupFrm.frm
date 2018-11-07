VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form INV_SRPSetupFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14535
   Icon            =   "INV_SRPSetupFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   14535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnRemoveSelected 
      Caption         =   "Remove Selected"
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
      TabIndex        =   15
      Top             =   8040
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pricing Details"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   10
      Top             =   840
      Width           =   14055
      Begin VB.CheckBox chkRoundOff 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Decimal Round off"
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
         Left            =   9720
         TabIndex        =   5
         Top             =   840
         Value           =   1  'Checked
         Width           =   3735
      End
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
         Height          =   345
         Left            =   6960
         TabIndex        =   4
         Text            =   "0.00"
         Top             =   840
         Width           =   2655
      End
      Begin VB.ComboBox cmbMode 
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
         ItemData        =   "INV_SRPSetupFrm.frx":000C
         Left            =   4560
         List            =   "INV_SRPSetupFrm.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   840
         Width           =   2295
      End
      Begin VB.ComboBox cmbOperator 
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
         ItemData        =   "INV_SRPSetupFrm.frx":002F
         Left            =   2040
         List            =   "INV_SRPSetupFrm.frx":003F
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   840
         Width           =   2415
      End
      Begin VB.ComboBox cmbBasePrice 
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
         ItemData        =   "INV_SRPSetupFrm.frx":006D
         Left            =   120
         List            =   "INV_SRPSetupFrm.frx":0077
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Value"
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
         Left            =   6960
         TabIndex        =   14
         Top             =   480
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mode"
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
         Left            =   4560
         TabIndex        =   13
         Top             =   480
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operator"
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
         Left            =   2040
         TabIndex        =   12
         Top             =   480
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Base Price"
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
         TabIndex        =   11
         Top             =   480
         Width           =   945
      End
   End
   Begin VB.CommandButton btnAddProduct 
      Caption         =   "Add Products"
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
      Left            =   10080
      TabIndex        =   6
      Top             =   2280
      Width           =   4215
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
      Left            =   12960
      TabIndex        =   9
      Top             =   8040
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
      Left            =   11520
      TabIndex        =   8
      Top             =   8040
      Width           =   1335
   End
   Begin MSComctlLib.ListView lvSearch 
      Height          =   5175
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   9128
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ProductId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Item Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Name"
         Object.Width           =   6526
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Category"
         Object.Width           =   2893
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Cost"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Old Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "New Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Cost"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Srp"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label lblShowMorePrice 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select/Unselect All"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   225
      Left            =   240
      MouseIcon       =   "INV_SRPSetupFrm.frx":0086
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Top             =   8040
      Width           =   1530
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Standard Retail Price (SRP) Setup"
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
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3765
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   8415
      Left            =   120
      Top             =   120
      Width           =   14295
   End
End
Attribute VB_Name = "INV_SRPSetupFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub ComputeMarkUp()
    Dim item As MSComctlLib.ListItem
        
    For Each item In lvSearch.ListItems
        If cmbBasePrice.ListIndex = 0 Then
            'base price is cost
            'check operator
            If cmbOperator.ListIndex = 0 Then 'add
                'check mode
                If cmbMode.ListIndex = 0 Then 'percent
                    item.SubItems(6) = (Val(Replace(item.SubItems(4), ",", "")) * _
                                        (Val(Replace(txtAmount.text, ",", "")) / 100)) + Val(Replace(item.SubItems(4), ",", ""))
                    
                Else 'pure integer
                    item.SubItems(6) = (Val(Replace(item.SubItems(4), ",", "")) + Val(Replace(txtAmount.text, ",", "")))
                End If
            ElseIf cmbOperator.ListIndex = 1 Then 'subtract
                'check mode
                If cmbMode.ListIndex = 0 Then 'percent
                    item.SubItems(6) = ((Val(Replace(item.SubItems(4), ",", "")) * _
                                        (Val(Replace(txtAmount.text, ",", "")) / 100)) - Val(Replace(item.SubItems(4), ",", ""))) * -1
                    
                Else 'pure integer
                    item.SubItems(6) = (Val(Replace(item.SubItems(4), ",", "")) - Val(Replace(txtAmount.text, ",", "")))
                End If
            ElseIf cmbOperator.ListIndex = 2 Then 'multiply
                'check mode
                If cmbMode.ListIndex = 0 Then 'percent
                    item.SubItems(6) = (Val(Replace(item.SubItems(4), ",", "")) * _
                                        (Val(Replace(txtAmount.text, ",", "")) / 100)) * Val(Replace(item.SubItems(4), ",", ""))
                    
                Else 'pure integer
                    item.SubItems(6) = (Val(Replace(item.SubItems(4), ",", "")) * Val(Replace(txtAmount.text, ",", "")))
                End If
            ElseIf cmbOperator.ListIndex = 3 Then 'divide
                'check mode
                If cmbMode.ListIndex = 0 Then 'percent
                    item.SubItems(6) = (Val(Replace(item.SubItems(4), ",", "")) * _
                                        (Val(Replace(txtAmount.text, ",", "")) / 100)) / Val(Replace(item.SubItems(4), ",", ""))
                    
                Else 'pure integer
                    item.SubItems(6) = (Val(Replace(item.SubItems(4), ",", "")) / Val(Replace(txtAmount.text, ",", "")))
                End If
            End If
        Else
            'base price is srp
            'check operator
            If cmbOperator.ListIndex = 0 Then 'add
                'check mode
                If cmbMode.ListIndex = 0 Then 'percent
                    item.SubItems(6) = (Val(Replace(item.SubItems(5), ",", "")) * _
                                        (Val(Replace(txtAmount.text, ",", "")) / 100)) + Val(Replace(item.SubItems(5), ",", ""))
                    
                Else 'pure integer
                    item.SubItems(6) = (Val(Replace(item.SubItems(5), ",", "")) + Val(Replace(txtAmount.text, ",", "")))
                End If
            ElseIf cmbOperator.ListIndex = 1 Then 'subtract
                'check mode
                If cmbMode.ListIndex = 0 Then 'percent
                    item.SubItems(6) = ((Val(Replace(item.SubItems(5), ",", "")) * _
                                        (Val(Replace(txtAmount.text, ",", "")) / 100)) - Val(Replace(item.SubItems(5), ",", ""))) * -1
                    
                Else 'pure integer
                    item.SubItems(6) = (Val(Replace(item.SubItems(5), ",", "")) - Val(Replace(txtAmount.text, ",", "")))
                End If
            ElseIf cmbOperator.ListIndex = 2 Then 'multiply
                'check mode
                If cmbMode.ListIndex = 0 Then 'percent
                    item.SubItems(6) = (Val(Replace(item.SubItems(5), ",", "")) * _
                                        (Val(Replace(txtAmount.text, ",", "")) / 100)) * Val(Replace(item.SubItems(5), ",", ""))
                    
                Else 'pure integer
                    item.SubItems(6) = (Val(Replace(item.SubItems(5), ",", "")) * Val(Replace(txtAmount.text, ",", "")))
                End If
            ElseIf cmbOperator.ListIndex = 3 Then 'divide
                'check mode
                If cmbMode.ListIndex = 0 Then 'percent
                    item.SubItems(6) = (Val(Replace(item.SubItems(5), ",", "")) * _
                                        (Val(Replace(txtAmount.text, ",", "")) / 100)) / Val(Replace(item.SubItems(5), ",", ""))
                    
                Else 'pure integer
                    item.SubItems(6) = (Val(Replace(item.SubItems(5), ",", "")) / Val(Replace(txtAmount.text, ",", "")))
                End If
            End If
        End If
        If chkRoundOff.value = vbChecked Then
            item.SubItems(6) = Round(Val(Replace(item.SubItems(6), ",", "")))
        End If
        
        item.SubItems(6) = FormatNumber(item.SubItems(6), 2, vbTrue, vbFalse)
    Next
End Sub

Private Sub btnAddProduct_Click()
    INV_AddProductFrm.isPriceMarkUp = True
    INV_AddProductFrm.Show (1)
End Sub

Private Sub btnCancel_Click()
   Unload Me
End Sub

Private Sub btnRemoveSelected_Click()
    Dim x As Integer
    For x = 1 To lvSearch.ListItems.Count
        If x > lvSearch.ListItems.Count Then Exit For
        If lvSearch.ListItems(x).Selected = True Then
            lvSearch.ListItems.Remove (x)
            x = x - 1
        End If
    Next
End Sub

Private Sub btnSave_Click()
    Set con = New ADODB.Connection
    Dim item As MSComctlLib.ListItem
    
    con.ConnectionString = ConnString
    con.Open
    For Each item In lvSearch.ListItems
        Set cmd = New ADODB.Command
        cmd.ActiveConnection = con
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "BASE_ProductSRP_Update"
        cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , item.text)
        cmd.Parameters.Append cmd.CreateParameter("@Price", adDecimal, adParamInput, , Val(Replace(item.SubItems(6), ",", "")))
                              cmd.Parameters("@Price").Precision = 18
                              cmd.Parameters("@Price").NumericScale = 2
        cmd.Execute
    Next
    con.Close
    MsgBox "Product prices updated.", vbInformation
End Sub

Private Sub chkRoundOff_Click()
    ComputeMarkUp
End Sub

Private Sub cmbBasePrice_Click()
    ComputeMarkUp
End Sub

Private Sub cmbMode_Click()
    ComputeMarkUp
End Sub

Private Sub cmbOperator_Click()
    ComputeMarkUp
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDelete
            btnRemoveSelected_Click
    End Select
End Sub

Private Sub Form_Load()
    lvSearch.ColumnHeaders(2).width = lvSearch.width * 0.12
    lvSearch.ColumnHeaders(3).width = lvSearch.width * 0.312
    lvSearch.ColumnHeaders(4).width = lvSearch.width * 0.175
    lvSearch.ColumnHeaders(5).width = lvSearch.width * 0.125
    lvSearch.ColumnHeaders(6).width = lvSearch.width * 0.125
    lvSearch.ColumnHeaders(7).width = lvSearch.width * 0.125
    
    cmbBasePrice.ListIndex = 0
    cmbOperator.ListIndex = 0
    cmbMode.ListIndex = 0
    
End Sub

Private Sub lblShowMorePrice_Click()
    Dim item As MSComctlLib.ListItem
    For Each item In lvSearch.ListItems
        If item.Selected = True Then
            item.Selected = False
        Else
            item.Selected = True
        End If
    Next
End Sub

Private Sub txtAmount_Change()
    If IsNumeric(txtAmount.text) = False Then
        txtAmount.text = "0.00"
    Else
        ComputeMarkUp
    End If
End Sub

Private Sub txtAmount_GotFocus()
    selectText txtAmount
End Sub
