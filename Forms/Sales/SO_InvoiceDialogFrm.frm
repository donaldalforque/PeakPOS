VERSION 5.00
Begin VB.Form SO_InvoiceDialogFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   Icon            =   "SO_InvoiceDialogFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2055
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
         TabIndex        =   5
         Top             =   1560
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
         TabIndex        =   4
         Top             =   1560
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
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1080
         Visible         =   0   'False
         Width           =   2415
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
         Left            =   1560
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox txtQuantity 
         Alignment       =   1  'Right Justify
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
         Left            =   1560
         TabIndex        =   1
         Text            =   "1"
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From Location"
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
         TabIndex        =   8
         Top             =   1080
         Visible         =   0   'False
         Width           =   1305
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
         TabIndex        =   7
         Top             =   480
         Width           =   465
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
         TabIndex        =   6
         Top             =   120
         Width           =   810
      End
   End
End
Attribute VB_Name = "SO_InvoiceDialogFrm"
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
    Dim isExisting As Boolean
    isExisting = False
    With SO_InvoiceFrm
        Dim item As MSComctlLib.ListItem
        If isModify = True Then
            'on error resume next
            .lvItems.SelectedItem.SubItems(4) = FormatNumber(txtQuantity.text, 2, vbTrue, vbFalse)
            .lvItems.SelectedItem.SubItems(6) = FormatNumber(txtPrice.text, 2, vbTrue, vbFalse)
            '.lvItems.SelectedItem.SubItems(8) = cmbLocation.ItemData(cmbLocation.ListIndex) 'location
            .CountTotal
            isModify = False
            Unload Me
            Exit Sub
        End If
        
'        For Each item In .lvItems.ListItems
'            If item.SubItems(9) = .lvItemList.SelectedItem.text Then
'                item.SubItems(4) = FormatNumber(Val(Replace(item.SubItems(4), ",", "")) + txtQuantity.text, 2, vbTrue, vbFalse)
'                item.SubItems(6) = txtPrice.text
'                item.SubItems(8) = cmbLocation.ItemData(cmbLocation.ListIndex) 'location
'                isExisting = True
'            End If
'        Next
        Unload Me
        .CountTotal
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
    txtQuantity.text = "1.00"
    selectText txtQuantity
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
    cmbLocation.ListIndex = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    isModify = False
End Sub

Private Sub txtPrice_Change()
    If IsNumeric(txtPrice.text) = False Then txtPrice.text = "0.00"
End Sub

Private Sub txtPrice_GotFocus()
    selectText txtPrice
End Sub

Private Sub txtQuantity_Change()
    If IsNumeric(txtQuantity.text) = False Then
        txtQuantity.text = "1"
    End If
End Sub

Private Sub txtQuantity_GotFocus()
    selectText txtQuantity
End Sub

