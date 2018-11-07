VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form SO_PickOrderDialogFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   Icon            =   "SO_PickOrderDialogFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1440
      TabIndex        =   4
      Top             =   2520
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
      Left            =   2880
      TabIndex        =   3
      Top             =   2520
      Width           =   1335
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
      TabIndex        =   2
      Top             =   1080
      Width           =   2655
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
      TabIndex        =   1
      Top             =   1440
      Width           =   2655
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
      Text            =   "0.00"
      Top             =   360
      Width           =   2655
   End
   Begin MSComCtl2.DTPicker dtDate 
      Height          =   345
      Left            =   1440
      TabIndex        =   5
      Top             =   1800
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
      Format          =   52559873
      CurrentDate     =   41686
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
      TabIndex        =   9
      Top             =   1080
      Width           =   780
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
      TabIndex        =   8
      Top             =   1440
      Width           =   960
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
      TabIndex        =   7
      Top             =   1800
      Width           =   435
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
      TabIndex        =   6
      Top             =   360
      Width           =   810
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   2295
      Left            =   120
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "SO_PickOrderDialogFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public isModify As Boolean
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
        With SO_PickOrderFrm
            Set item = .lvItems.ListItems.add(, , "")
                item.SubItems(1) = ""
                item.SubItems(2) = .lvItemList.SelectedItem.text
                item.SubItems(3) = .lvItemList.SelectedItem.SubItems(1)
                item.SubItems(4) = .lvItemList.SelectedItem.SubItems(2)
                item.SubItems(5) = FormatNumber(txtQuantity.text, 2, vbTrue)
                item.SubItems(6) = .lvItemList.SelectedItem.SubItems(4)
                item.SubItems(7) = cmbLocation.text
                item.SubItems(8) = dtDate.value
                item.SubItems(9) = txtReference.text
                item.SubItems(10) = cmbLocation.ItemData(cmbLocation.ListIndex)
        Unload Me
        .txtItemSearch.SetFocus
        .lvItemList.Visible = False
        End With
    Else
        With SO_PickOrderFrm
            .lvItems.SelectedItem.SubItems(5) = txtQuantity.text
            .lvItems.SelectedItem.SubItems(8) = Format(dtDate.value, "MM/DD/YY")
            .lvItems.SelectedItem.SubItems(9) = txtReference.text
            On Error Resume Next
            .lvItems.SelectedItem.SubItems(7) = cmbLocation.text
            .lvItems.SelectedItem.SubItems(11) = cmbLocation.ItemData(cmbLocation.ListIndex)
            Unload Me
        End With
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
    Populate "Location"
    dtDate.value = Format(Now, "MM/DD/YY")
    txtReference.text = SO_PickOrderFrm.txtDeliveryNumber.text
    On Error Resume Next
    cmbLocation.text = SO_PickOrderFrm.cmbLocation.text
End Sub

Private Sub txtQuantity_Change()
    If IsNumeric(txtQuantity.text) = False Then
        txtQuantity.text = "0.00"
    End If
End Sub

Private Sub txtQuantity_GotFocus()
    selectText txtQuantity
End Sub

Private Sub txtReference_GotFocus()
    selectText txtReference
End Sub

