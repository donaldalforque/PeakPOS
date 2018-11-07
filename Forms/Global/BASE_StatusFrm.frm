VERSION 5.00
Begin VB.Form BASE_StatusFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Status"
   ClientHeight    =   1455
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnOK 
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
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.ComboBox cmbStatus 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   4215
   End
End
Attribute VB_Name = "BASE_StatusFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnOK_Click()
    If PO_PurchaseReturnInvoiceFrm.Visible = True Then
        PO_PurchaseReturnInvoiceFrm.lvModules.SelectedItem.SubItems(4) = cmbStatus.text
        PO_PurchaseReturnInvoiceFrm.lvModules.SelectedItem.SubItems(6) = cmbStatus.ItemData(cmbStatus.ListIndex)
    Else
        SO_InvoiceSalesReturnFrm.lvModules.SelectedItem.SubItems(4) = cmbStatus.text
        SO_InvoiceSalesReturnFrm.lvModules.SelectedItem.SubItems(6) = cmbStatus.ItemData(cmbStatus.ListIndex)
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Populate "Status"
End Sub

Public Sub Populate(ByVal data As String)
    Dim item As MSComctlLib.ListItem
    Select Case data
        Case "Status"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Status")
            cmbStatus.Clear
            cmbStatus.AddItem ""
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!StatusId >= 16 Then
                        cmbStatus.AddItem rec!Status
                        cmbStatus.ItemData(cmbStatus.NewIndex) = rec!StatusId
                    End If
                    rec.MoveNext
                Loop
            End If
            cmbStatus.ListIndex = 0
        Case "Terms"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Terms")
            cmbTerms.Clear
            cmbTerms.AddItem ""
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        cmbTerms.AddItem rec!Terms
                        cmbTerms.ItemData(cmbTerms.NewIndex) = rec!TermId
                    End If
                    rec.MoveNext
                Loop
            End If
            cmbTerms.ListIndex = 0
        Case "Customer"
            Set rec = New ADODB.Recordset
            Set rec = Global_Data("Customer")
            cmbCustomer.Clear
            cmbCustomer.AddItem ""
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isActive = "True" Then
                        cmbCustomer.AddItem rec!Name
                        cmbCustomer.ItemData(cmbCustomer.NewIndex) = rec!CustomerId
                    End If
                    rec.MoveNext
                Loop
            End If
            cmbCustomer.ListIndex = 0
    End Select
End Sub
