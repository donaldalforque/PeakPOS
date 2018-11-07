VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form POS_DiscountFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   Icon            =   "POS_DiscountFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnAccept 
      Caption         =   "ENTER: Accept"
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
      Left            =   2400
      Picture         =   "POS_DiscountFrm.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "ESC:Cancel"
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
      Left            =   4080
      Picture         =   "POS_DiscountFrm.frx":23E0
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5400
      Width           =   1575
   End
   Begin MSComctlLib.ListView lvDiscount 
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   8070
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "LocationId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Discount"
         Object.Width           =   6253
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Percent/Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Tax Exempt"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblCaption_Title 
      AutoSize        =   -1  'True
      Caption         =   "Discounts"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   1230
   End
   Begin VB.Image picModuleImage 
      Height          =   480
      Left            =   120
      Picture         =   "POS_DiscountFrm.frx":476F
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "POS_DiscountFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Populate(ByVal data As String)
    Dim item As MSComctlLib.ListItem
    Select Case data
        Case "Term"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Discount")
            lvDiscount.ListItems.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        Set item = lvDiscount.ListItems.add(, , "")
                            item.SubItems(1) = rec!DiscountId
                            item.SubItems(2) = rec!Name
                            item.SubItems(3) = rec!Percentage
                            item.SubItems(4) = "False"
                            If rec!isTaxExempt = True Then item.SubItems(4) = "True"
                            item.Checked = True
                    End If
                    rec.MoveNext
                Loop
            End If
    End Select
End Sub

Private Sub btnAccept_Click()
    If lvDiscount.SelectedItem.SubItems(1) = "1" Then 'Custom Price
        Dim Y As String
        Dim NewPrice As Double
        Y = InputBox("Please input new price", "Custom Price")
        If IsNumeric(Y) = False Then
            MsgBox "Invalid price.", vbCritical
        Else
            If NVAL(Y) >= NVAL(POS_CashierFrm.lvList.SelectedItem.SubItems(9)) Then
                POS_CashierFrm.lvList.SelectedItem.SubItems(3) = FormatNumber(Y, 2, vbTrue, vbFalse)
                
                POS_CashierFrm.lvList.SelectedItem.SubItems(4) = FormatNumber(0, 2, vbTrue, vbFalse)
                POS_CashierFrm.lvList.SelectedItem.SubItems(19) = 0
                
                POS_CashierFrm.CountTotal
                POS_CashierFrm.CountTax
                Exit Sub
            End If
            
            NewPrice = NVAL(POS_CashierFrm.lvList.SelectedItem.SubItems(3)) - NVAL(Y)
            
            POS_CashierFrm.lvList.SelectedItem.SubItems(4) = FormatNumber(NewPrice, 2, vbTrue, vbFalse)
            POS_CashierFrm.lvList.SelectedItem.SubItems(19) = (NewPrice / NVAL(POS_CashierFrm.lvList.SelectedItem.SubItems(3))) * 100
            
            POS_CashierFrm.CountTotal
            POS_CashierFrm.CountTax
            Exit Sub
        End If
    End If
    
    x = MsgBox("Do you want to apply discount for all items?", vbQuestion + vbYesNo)
    If x = vbYes Then
        'Dim item As MSComctlLib.ListItem
        For Each item In POS_CashierFrm.lvList.ListItems
            item.SubItems(20) = lvDiscount.SelectedItem.SubItems(4) 'TAX EXEMPT
            item.SubItems(19) = lvDiscount.SelectedItem.SubItems(3)
        Next
    Else
        POS_CashierFrm.lvList.SelectedItem.SubItems(20) = lvDiscount.SelectedItem.SubItems(4) 'TAX EXEMPT
        POS_CashierFrm.lvList.SelectedItem.SubItems(19) = lvDiscount.SelectedItem.SubItems(3)
    End If
    
    POS_CashierFrm.CountTotal
    POS_CashierFrm.CountTax
    Unload Me
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
    lvDiscount.ColumnHeaders(1).width = lvDiscount.width * 0
    lvDiscount.ColumnHeaders(3).width = lvDiscount.width * 0.3133
    lvDiscount.ColumnHeaders(4).width = lvDiscount.width * 0.3133
    lvDiscount.ColumnHeaders(5).width = lvDiscount.width * 0.3133
    Populate "Term"
End Sub
