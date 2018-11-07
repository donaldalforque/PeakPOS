VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form POS_PricingSchemeFrm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6390
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCancel 
      Caption         =   "ESC: Cancel"
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
      Left            =   4680
      Picture         =   "POS_PricingSchemeFrm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6600
      Width           =   1575
   End
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
      Left            =   2880
      Picture         =   "POS_PricingSchemeFrm.frx":238F
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6600
      Width           =   1695
   End
   Begin MSComctlLib.ListView lvPricingScheme 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   11245
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
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "PricingSchemeId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Pricing"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "POS_PricingSchemeFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public ProductId As String

Private Sub btnAccept_Click()
    If lvPricingScheme.ListItems.Count = 0 Then Exit Sub
    Dim item As MSComctlLib.ListItem
    With POS_CashierFrm
        .GetPrice (lvPricingScheme.SelectedItem.text)
        Unload Me
    End With
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
    lvPricingScheme.ColumnHeaders(2).width = lvPricingScheme.width * 0.92
    
    Set rec = New ADODB.Recordset
    Dim item As MSComctlLib.ListItem
    
    Set rec = Global_Data("PricingScheme")
    If Not rec.EOF Then
        Do Until rec.EOF
            Set item = lvPricingScheme.ListItems.add(, , rec!PricingSchemeId)
                item.SubItems(1) = rec!PricingScheme
            rec.MoveNext
        Loop
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ProductId = 0
End Sub



