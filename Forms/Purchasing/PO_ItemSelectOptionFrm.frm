VERSION 5.00
Begin VB.Form PO_ItemSelectOptionFrmBAK 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   Icon            =   "PO_ItemSelectOptionFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.OptionButton optSacks 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sacks"
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
         Left            =   2760
         TabIndex        =   6
         Top             =   1920
         Width           =   1215
      End
      Begin VB.OptionButton optKilos 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kilos"
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
         Left            =   1440
         TabIndex        =   5
         Top             =   1920
         Value           =   -1  'True
         Width           =   1215
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
         Left            =   2760
         TabIndex        =   8
         Top             =   2280
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
         Left            =   1440
         TabIndex        =   7
         Top             =   2280
         Width           =   1215
      End
      Begin VB.ComboBox cmbLocation 
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
         TabIndex        =   4
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox txtPrice 
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
         TabIndex        =   3
         Text            =   "0.00"
         Top             =   840
         Width           =   2535
      End
      Begin VB.TextBox txtKilos 
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
         TabIndex        =   2
         Text            =   "50.00"
         Top             =   480
         Width           =   2535
      End
      Begin VB.TextBox txtSacks 
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
         TabIndex        =   1
         Text            =   "1.00"
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Store as"
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
         TabIndex        =   13
         Top             =   1920
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To Location"
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
         TabIndex        =   12
         Top             =   1440
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
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
         Top             =   840
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kilo(s)"
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
         TabIndex        =   10
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sack(s)"
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
         TabIndex        =   9
         Top             =   120
         Width           =   645
      End
   End
End
Attribute VB_Name = "PO_ItemSelectOptionFrmBAK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public isModify As Boolean
Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOk_Click()
    Dim isExisting As Boolean
    isExisting = False
    With PO_PurchaseOrderFrm
        Dim item As MSComctlLib.ListItem
        If isModify = True Then
            .lvItems.SelectedItem.SubItems(4) = FormatNumber(txtSacks.text, 2, vbTrue, vbFalse)
            .lvItems.SelectedItem.SubItems(5) = FormatNumber(txtKilos.text, 2, vbTrue, vbFalse)
            .lvItems.SelectedItem.SubItems(6) = FormatNumber(txtPrice.text, 2, vbTrue, vbFalse)
            .lvItems.SelectedItem.SubItems(8) = cmbLocation.ItemData(cmbLocation.ListIndex) 'location
            .CountTotal
            isModify = False
            .txtTare.text = FormatNumber((Val(Replace(txtSacks.text, ",", "")) / 2), 2, vbTrue, vbFalse)
            Unload Me
            Exit Sub
         End If
        
        For Each item In .lvItems.ListItems
            If item.SubItems(9) = .lvItemList.SelectedItem.text Then
                item.SubItems(4) = FormatNumber(Val(Replace(item.SubItems(4), ",", "")) + txtSacks.text, 2, vbTrue, vbFalse)
                item.SubItems(5) = FormatNumber(Val(Replace(item.SubItems(5), ",", "")) + txtKilos.text, 2, vbTrue, vbFalse)
                item.SubItems(6) = FormatNumber(txtPrice.text, 2, vbTrue, vbFalse)
                item.SubItems(8) = cmbLocation.ItemData(cmbLocation.ListIndex) 'location
                If optKilos.value = True Then
                    item.SubItems(10) = "Kilos"
                Else
                    item.SubItems(10) = "Sacks"
                End If
                isExisting = True
            End If
        Next
  
        If isExisting = False Then
            Set item = .lvItems.ListItems.add(, , "")
            item.SubItems(1) = ""
            item.SubItems(2) = .lvItemList.SelectedItem.SubItems(1) 'ItemCode
            item.SubItems(3) = .lvItemList.SelectedItem.SubItems(2) 'Name
            item.SubItems(4) = FormatNumber(txtSacks.text, 2, vbTrue, vbFalse)
            item.SubItems(5) = FormatNumber(txtKilos.text, 2, vbTrue, vbFalse)
            item.SubItems(6) = FormatNumber(txtPrice.text, 2, vbTrue, vbFalse)
            item.SubItems(8) = cmbLocation.ItemData(cmbLocation.ListIndex) 'location
            item.SubItems(9) = .lvItemList.SelectedItem.text
            If optKilos.value = True Then
                item.SubItems(10) = "Kilos"
            Else
                item.SubItems(10) = "Sacks"
            End If
        End If
         
        isModify = False
            
        .lvItemList.Visible = False
        .lvItems.ListItems(.lvItems.ListItems.Count).Selected = True
        .lvItems.ListItems(.lvItems.ListItems.Count).EnsureVisible
        .txtTare.text = FormatNumber((Val(Replace(txtSacks.text, ",", "")) / 2), 2, vbTrue, vbFalse)
        .CountTotal
        
        Unload Me
        .txtItemSearch.SetFocus
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            btnCancel_Click
        Case vbKeyReturn
            btnOk_Click
    End Select
End Sub

Private Sub Form_Load()
    txtSacks.text = "1.00"
    txtKilos.text = "50.00"
    txtPrice.text = "1.00"
    selectText txtSacks
    Set rec = New ADODB.Recordset
    Set rec = Global_Data("Location")
    cmbLocation.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            If rec!isActive = "True" Then
                cmbLocation.AddItem rec!Location
                cmbLocation.ItemData(cmbLocation.NewIndex) = rec!LocationId
                cmbLocation.ListIndex = 0
            End If
            rec.MoveNext
        Loop
    End If
End Sub

Private Sub txtKilos_Change()
    If IsNumeric(txtKilos.text) = False Then
        txtKilos.text = "50.00"
    End If
End Sub

Private Sub txtKilos_GotFocus()
    selectText txtKilos
End Sub

Private Sub txtPrice_Change()
    If IsNumeric(txtPrice.text) = False Then
        txtPrice.text = "1.00"
    End If
End Sub

Private Sub txtPrice_GotFocus()
    selectText txtPrice
End Sub

Private Sub txtSacks_Change()
    If IsNumeric(txtSacks.text) = False Then
        txtSacks.text = "1.00"
    End If
End Sub

Private Sub txtSacks_GotFocus()
    selectText txtSacks
End Sub
