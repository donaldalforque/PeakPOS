VERSION 5.00
Begin VB.Form INV_WarehouseWithdrawalOptFrm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   4335
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
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
         TabIndex        =   3
         Top             =   1200
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
         TabIndex        =   2
         Top             =   1200
         Width           =   1215
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
         TabIndex        =   1
         Text            =   "1"
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Available"
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
         Top             =   480
         Width           =   870
      End
      Begin VB.Label lblAvailable 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label1 
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
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   810
      End
   End
End
Attribute VB_Name = "INV_WarehouseWithdrawalOptFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public isModify, isFormSearch As Boolean

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOk_Click()
    Dim item As MSComctlLib.ListItem
    If isModify = False Then
        If isFormSearch = False Then
            With INV_WarehouseWithdrawalFrm
                'If isExisting = False Then
                    Set item = .lvItems.ListItems.add(, , "")
                    item.SubItems(1) = ""
                    item.SubItems(2) = .lvItemList.SelectedItem.SubItems(1) 'ItemCode
                    item.SubItems(3) = .lvItemList.SelectedItem.SubItems(2) 'Name
                    item.SubItems(4) = FormatNumber(txtQuantity.text, 2, vbTrue, vbFalse)
                    item.SubItems(5) = .lvItemList.SelectedItem.SubItems(3)
                    'item.SubItems(6) = FormatNumber(txtPrice.text, 2, vbTrue, vbFalse)
                    'item.SubItems(8) = cmbLocation.ItemData(cmbLocation.ListIndex) 'location
                    item.SubItems(6) = .lvItemList.SelectedItem.text
                'End If
                    
                Unload Me
                
                .lvItemList.Visible = False
                .txtItemSearch.SetFocus
                '.txtCode.SetFocus
                '.txtReferenceNumber.SetFocus
                .lvItems.ListItems(.lvItems.ListItems.Count).Selected = True
                .lvItems.ListItems(.lvItems.ListItems.Count).EnsureVisible
                .CountTotal
            End With
        Else
            With INV_WarehouseWithdrawalFrm
                'If isExisting = False Then
                    Set item = .lvItems.ListItems.add(, , "")
                    item.SubItems(1) = ""
                    item.SubItems(2) = INV_ProductSearch.lvItemList.SelectedItem.SubItems(1) 'ItemCode
                    item.SubItems(3) = INV_ProductSearch.lvItemList.SelectedItem.SubItems(2) 'Name
                    item.SubItems(4) = FormatNumber(txtQuantity.text, 2, vbTrue, vbFalse)
                    item.SubItems(5) = INV_ProductSearch.lvItemList.SelectedItem.SubItems(3)
                    'item.SubItems(6) = FormatNumber(txtPrice.text, 2, vbTrue, vbFalse)
                    'item.SubItems(8) = cmbLocation.ItemData(cmbLocation.ListIndex) 'location
                    item.SubItems(6) = INV_ProductSearch.lvItemList.SelectedItem.text
                'End If
                    
                Unload Me
                Unload INV_ProductSearch
                
                '.lvItemList.Visible = False
                .txtItemSearch.SetFocus
                '.txtCode.SetFocus
                '.txtReferenceNumber.SetFocus
                .lvItems.ListItems(.lvItems.ListItems.Count).Selected = True
                .lvItems.ListItems(.lvItems.ListItems.Count).EnsureVisible
                .CountTotal
            End With
        End If
    Else
        With INV_WarehouseWithdrawalFrm
            'If isExisting = False Then
                'Set item = .lvItems.ListItems.add(, , "")
                'item.SubItems(1) = ""
                'item.SubItems(2) = .lvItemList.SelectedItem.SubItems(1) 'ItemCode
                'item.SubItems(3) = .lvItemList.SelectedItem.SubItems(2) 'Name
                .lvItems.SelectedItem.SubItems(4) = FormatNumber(txtQuantity.text, 2, vbTrue, vbFalse)
                'item.SubItems(5) = .lvItemList.SelectedItem.SubItems(3)
                'item.SubItems(6) = FormatNumber(txtPrice.text, 2, vbTrue, vbFalse)
                'item.SubItems(8) = cmbLocation.ItemData(cmbLocation.ListIndex) 'location
                'item.SubItems(6) = .lvItemList.SelectedItem.text
            'End If
                
            Unload Me
            
            .lvItemList.Visible = False
            .txtItemSearch.SetFocus
            '.txtCode.SetFocus
            '.txtReferenceNumber.SetFocus
            '.lvItems.ListItems(.lvItems.ListItems.Count).Selected = True
            '.lvItems.ListItems(.lvItems.ListItems.Count).EnsureVisible
            .CountTotal
        End With
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            btnOk_Click
        Case vbKeyEscape
            btnCancel_Click
    End Select
End Sub

Private Sub Form_Load()
    selectText txtQuantity
    On Error Resume Next
    If isFormSearch = False Then
        If Val(Replace(INV_WarehouseWithdrawalFrm.lvItemList.SelectedItem.SubItems(6), ",", "")) <= 0 Then
            MsgBox "WARNING! Inventory below 0!", vbCritical, "WARNING!"
        End If
    Else
        If Val(Replace(INV_ProductSearch.lvItemList.SelectedItem.SubItems(6), ",", "")) <= 0 Then
            MsgBox "WARNING! Inventory below 0!", vbCritical, "WARNING!"
        End If
    End If
End Sub

Private Sub Label2_Click()

End Sub

Private Sub txtQuantity_Change()
    If IsNumeric(txtQuantity.text) = False Then
        txtQuantity.text = "1"
        selectText txtQuantity
    End If
End Sub

Private Sub txtQuantity_GotFocus()
    selectText txtQuantity
End Sub
