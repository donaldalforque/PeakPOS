VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form PO_ChargesFrmBAK 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4305
   Icon            =   "PO_ChargesFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton btnSave 
         Caption         =   "Save && Close"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CommandButton btnCancel 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   1
         Top             =   4200
         Width           =   1215
      End
      Begin MSComctlLib.ListView lvCharges 
         Height          =   3375
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   5953
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FlatScrollBar   =   -1  'True
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ChargeId"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Charge"
            Object.Width           =   6253
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "PO_ChargesFrm.frx":000C
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Include tracking, tare and other charges in your purchase order."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   720
         TabIndex        =   4
         Top             =   120
         Width           =   3255
      End
   End
End
Attribute VB_Name = "PO_ChargesFrmBAK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnSave_Click()
    Dim item As MSComctlLib.ListItem
    Dim hitem As MSComctlLib.ListItem
    For Each item In lvCharges.ListItems
        If item.Checked = True Then
            Set hitem = PO_PurchaseOrderFrm.lvCharges.ListItems.add(, , "")
                hitem.text = item.SubItems(1)
                hitem.SubItems(1) = item.SubItems(2)
                hitem.SubItems(2) = "0.00"
        End If
    Next
    Unload Me
End Sub

Private Sub Form_Load()
    Dim item As MSComctlLib.ListItem
    Dim hitem As MSComctlLib.ListItem
    Set rec = New ADODB.Recordset
    Set rec = Global_Data("Charges")
    If Not rec.EOF Then
        lvCharges.ListItems.Clear
        Do Until rec.EOF
            If rec!isActive = "True" Then
                Set item = lvCharges.ListItems.add(, , "")
                        item.SubItems(1) = rec!ChargeId
                        item.SubItems(2) = rec!Charge
            End If
            rec.MoveNext
        Loop
    End If

    'Clear Chargess already selected
    For Each item In PO_PurchaseOrderFrm.lvCharges.ListItems
        For Each hitem In lvCharges.ListItems
            If item.text = hitem.SubItems(1) Then
                lvCharges.ListItems.Remove (hitem.Index)
                Exit For
            End If
        Next
    Next
End Sub
