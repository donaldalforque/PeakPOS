VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Begin VB.Form INV_ProductConversionDialogFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   Icon            =   "INV_ProductConversionDialogFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView lvItemList 
      Height          =   2655
      Left            =   1320
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   4683
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
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
      NumItems        =   4
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
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Uom"
         Object.Width           =   0
      EndProperty
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
      Left            =   5520
      TabIndex        =   3
      Top             =   3600
      Width           =   1335
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
      Left            =   6960
      TabIndex        =   5
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Product"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   8175
      Begin VB.ComboBox cmbProducts 
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
         Left            =   1200
         TabIndex        =   1
         Top             =   840
         Width           =   6615
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Selected Product Details"
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
         Left            =   1200
         TabIndex        =   9
         Top             =   1560
         Width           =   6615
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Uom:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   510
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   615
         End
         Begin VB.Label lblUom 
            BackStyle       =   0  'Transparent
            Caption         =   "Unit of Masure:"
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
            Left            =   840
            TabIndex        =   11
            Top             =   720
            Width           =   5595
         End
         Begin VB.Label lblSelectedProduct 
            BackStyle       =   0  'Transparent
            Caption         =   "Name:"
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
            Left            =   840
            TabIndex        =   10
            Top             =   360
            Width           =   5595
         End
      End
      Begin VB.ComboBox cmbCode 
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
         Left            =   1200
         TabIndex        =   0
         Top             =   480
         Width           =   6615
      End
      Begin VB.TextBox txtQuantity 
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
         Left            =   1200
         MaxLength       =   50
         TabIndex        =   2
         Text            =   "1.00"
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
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
         TabIndex        =   14
         Top             =   840
         Width           =   555
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
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
         TabIndex        =   7
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label7 
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
         TabIndex        =   6
         Top             =   1200
         Width           =   810
      End
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   3375
      Left            =   120
      Top             =   120
      Width           =   8175
   End
End
Attribute VB_Name = "INV_ProductConversionDialogFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Uom As String
Dim ProductId As Integer
'Public isModify As Boolean

Private Sub txtBasicInfo_ItemCode_Change()

End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnSave_Click()
    If ProductId = 0 Then
        MsgBox "Please select a product.", vbCritical
        cmbProducts.SetFocus
    Else
        'If isModify = False Then
            INV_ProductConversionFrm.lvSearch.SelectedItem.SubItems(5) = lvItemList.SelectedItem.SubItems(1)
            INV_ProductConversionFrm.lvSearch.SelectedItem.SubItems(6) = lvItemList.SelectedItem.SubItems(2)
            INV_ProductConversionFrm.lvSearch.SelectedItem.SubItems(7) = FormatNumber(txtQuantity.text, 2, vbTrue, vbFalse)
            INV_ProductConversionFrm.lvSearch.SelectedItem.SubItems(8) = lvItemList.SelectedItem.SubItems(3)
            INV_ProductConversionFrm.lvSearch.SelectedItem.SubItems(9) = lvItemList.SelectedItem.text
        'Else
            
        'End If
        Unload Me
    End If
End Sub

Private Sub cmbCode_Change()
    Dim item As MSComctlLib.ListItem
    For Each item In lvItemList.ListItems
        If UCase(item.SubItems(1)) Like UCase(cmbCode.text) & "*" Then
            lvItemList.Visible = True
            lvItemList.Top = 960
            item.Selected = True
            item.EnsureVisible
            'lvItemList.SetFocus
            Exit For
        Else
            lvItemList.Visible = False
        End If
    Next
End Sub

Private Sub cmbCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            If lvItemList.Visible = True Then
                lvItemList.SetFocus
            End If
    End Select
End Sub

Private Sub cmbProducts_Change()
    Dim item As MSComctlLib.ListItem
    For Each item In lvItemList.ListItems
        If UCase(item.SubItems(2)) Like UCase(cmbProducts.text) & "*" Then
            lvItemList.Visible = True
            lvItemList.Top = 1320
            item.Selected = True
            item.EnsureVisible
            'lvItemList.SetFocus
            Exit For
        Else
            lvItemList.Visible = False
        End If
    Next
End Sub

Private Sub cmbProducts_Click()
'    lblSelectedProduct.Caption = cmbProducts.text
'    Uom = cmbProducts.text
End Sub

Private Sub cmbProducts_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            If lvItemList.Visible = True Then
                lvItemList.SetFocus
            End If
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If lvItemList.Visible = True Then Exit Sub
            btnSave_Click
        Case vbKeyEscape
            btnCancel_Click
    End Select
End Sub

Private Sub Form_Load()
    Set con = New ADODB.Connection
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_Product_Get"
    cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , Null)
    Set rec = cmd.Execute
    lvItemList.ListItems.Clear
    Dim item As MSComctlLib.ListItem
    If Not rec.EOF Then
        Do Until rec.EOF
            Set item = lvItemList.ListItems.add(, , rec!ProductId)
                item.SubItems(1) = rec!itemcode
                item.SubItems(2) = rec!Name
                item.SubItems(3) = rec!Uom
            rec.MoveNext
        Loop
    End If
    con.Close
    
    
    lvItemList.ColumnHeaders(2).width = lvItemList.width * 0.22
    lvItemList.ColumnHeaders(3).width = lvItemList.width * 0.52
    lvItemList.ColumnHeaders(4).width = lvItemList.width * 0.2
End Sub

Private Sub lvItemList_DblClick()
    If lvItemList.ListItems.Count > 0 Then
        ProductId = lvItemList.SelectedItem.text
        Uom = lvItemList.SelectedItem.SubItems(3)
        lblSelectedProduct.Caption = "Name: " & lvItemList.SelectedItem.SubItems(2)
        lblUom.Caption = "Unit of Measure: " & lvItemList.SelectedItem.SubItems(3)
        lvItemList.Visible = False
        txtQuantity.SetFocus
    End If
End Sub

Private Sub lvItemList_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If lvItemList.ListItems.Count > 0 Then
                ProductId = lvItemList.SelectedItem.text
                Uom = lvItemList.SelectedItem.SubItems(3)
                lblSelectedProduct.Caption = lvItemList.SelectedItem.SubItems(2)
                lblUom.Caption = lvItemList.SelectedItem.SubItems(3)
                lvItemList.Visible = False
                txtQuantity.SetFocus
            End If
    End Select
End Sub

Private Sub txtQuantity_Change()
    If IsNumeric(txtQuantity.text) = False Then
        txtQuantity.text = "1.00"
    End If
End Sub

Private Sub txtQuantity_GotFocus()
    selectText txtQuantity
End Sub
