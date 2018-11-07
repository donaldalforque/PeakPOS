VERSION 5.00
Begin VB.Form INV_TransferStockOptFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4335
   Icon            =   "INV_TransferStocklOptFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.ComboBox cmbUnit 
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
         Top             =   480
         Width           =   2535
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
         TabIndex        =   4
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
         TabIndex        =   3
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
      Begin VB.Label Label2 
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
         TabIndex        =   8
         Top             =   840
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
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
         Width           =   390
      End
      Begin VB.Label lblAvailable 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "0.00"
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
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
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
         TabIndex        =   5
         Top             =   120
         Width           =   810
      End
   End
End
Attribute VB_Name = "INV_TransferStockOptFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public isModify, isFormSearch As Boolean
Dim Conversion As Double

Private Sub btnCancel_Click()
    Unload Me
End Sub
Private Function CheckQuantity() As Double
    Dim quantity As Double
    Dim item As MSComctlLib.ListItem
    
    With INV_TransferStockFrm
        If .lvItems.ListItems.Count > 0 Then
            For Each item In .lvItems.ListItems
                If item.SubItems(6) = INV_TransferStockFrm.currentproductid Then
                    quantity = quantity + NVAL(item.SubItems(4))
                End If
            Next
        End If
    End With
    CheckQuantity = quantity
End Function
Private Sub btnOK_Click()
    If cmbUnit.Text = "" Then
        MsgBox "Please select a unit of measure.", vbCritical
        Exit Sub
    End If
    
    Dim item As MSComctlLib.ListItem
    Dim Available As Double
    Dim ReserveId As String
    Dim ActualQuantity As Double
    
    If isModify = False Then
        'GetInventorySettings
        With INV_TransferStockFrm
            If AllowNegativeInventory = False Then
                'CHECK AVAILABLE QUANTITY
                Available = checkAvailableQuantity(.lvItemList.SelectedItem.Text, INV_TransferStockFrm.cmbOrigin.ItemData(INV_TransferStockFrm.cmbOrigin.ListIndex))
                
                If Available < NVAL(txtQuantity.Text) * Conversion Then
                    MsgBox "Insufficient quantity. Remaining quantity for location " & .cmbOrigin.Text & ": " & FormatNumber(Available, 2, vbTrue, vbFalse), vbCritical, "Insufficient Quantity"
                    Exit Sub
                Else
                    'INSERT RESERVES
                    'ModId 4 - TransferStock
                    ReserveId = ReserveProduct(0, .lvItemList.SelectedItem.Text, Conversion * NVAL(txtQuantity.Text), UserId, WorkstationId, False, 4)
                End If
            End If
        End With
        
        With INV_TransferStockFrm
            Set item = .lvItems.ListItems.add(, , "")
            item.SubItems(1) = ""
            item.SubItems(2) = .lvItemList.SelectedItem.SubItems(1) 'ItemCode
            item.SubItems(3) = .lvItemList.SelectedItem.SubItems(2) 'Name
            item.SubItems(4) = FormatNumber(txtQuantity.Text, 2, vbTrue, vbFalse)
            item.SubItems(5) = cmbUnit.Text
            item.SubItems(6) = .lvItemList.SelectedItem.Text
            item.SubItems(8) = Conversion
            item.SubItems(9) = ReserveId
            Unload Me
            
            .cmbOrigin.Enabled = False
            'Unload INV_ProductSearch
            
            .txtItemSearch.SetFocus
            .lvItems.ListItems(.lvItems.ListItems.Count).Selected = True
            .lvItems.ListItems(.lvItems.ListItems.Count).EnsureVisible
            .lvItemList.Visible = False
            .CountTotal
        End With
    Else
        'GetInventorySettings
        If AllowNegativeInventory = False Then
            'CHECK AVAILABLE QUANTITY
            With INV_TransferStockFrm
                Available = checkAvailableQuantity(.lvItemList.SelectedItem.Text, .cmbOrigin.ItemData(.cmbOrigin.ListIndex))
                
                If Available + (NVAL(.lvItems.SelectedItem.SubItems(4)) * NVAL(.lvItems.SelectedItem.SubItems(8))) < (NVAL(txtQuantity.Text) * Conversion) Then  'less the current
                    MsgBox "Insufficient quantity. Remaining quantity for location " & .cmbOrigin.Text & ": " & FormatNumber(Available, 2, vbTrue, vbFalse), vbCritical, "Insufficient Quantity"
                    Exit Sub
                Else
                    'UDPATE RESERVES
                    'ModId 4 - transfer
                    ReserveId = ReserveProduct(.lvItems.SelectedItem.SubItems(9), .lvItems.SelectedItem.Text, (Conversion * NVAL(txtQuantity.Text)), UserId, WorkstationId, False, 4)
                End If
            End With
        End If
        
        With INV_TransferStockFrm
            .lvItems.SelectedItem.SubItems(4) = FormatNumber(txtQuantity.Text, 2, vbTrue, vbFalse)
            .lvItems.SelectedItem.SubItems(5) = cmbUnit.Text
            .lvItems.SelectedItem.SubItems(8) = Conversion
            Unload Me
            .lvItemList.Visible = False
            .txtItemSearch.SetFocus
            .CountTotal
        End With
    End If
End Sub

Private Sub cmbUnit_Click()
    If isModify = True Then
       Conversion = GetProductConversion(INV_TransferStockFrm.lvItems.SelectedItem.SubItems(6), cmbUnit.ItemData(cmbUnit.ListIndex), "")
    Else
        Conversion = GetProductConversion(INV_TransferStockFrm.lvItemList.SelectedItem.Text, cmbUnit.ItemData(cmbUnit.ListIndex), "")
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            btnOK_Click
        Case vbKeyEscape
            btnCancel_Click
    End Select
End Sub
Private Sub GetUoms()
    'Get Uom Related
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "INV_UomConversion_Get"
    
    If isModify = True Then
        cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , INV_TransferStockFrm.lvItems.SelectedItem.SubItems(6))
    Else
        cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , INV_TransferStockFrm.lvItemList.SelectedItem.Text)
    End If
    Set rec = cmd.Execute
    'lvUom.ListItems.Clear
    cmbUnit.Clear
    Dim item As MSComctlLib.ListItem
    If Not rec.EOF Then
        Do Until rec.EOF
            cmbUnit.AddItem rec!Uom
            cmbUnit.ItemData(cmbUnit.NewIndex) = rec!UomId
            rec.MoveNext
        Loop
    End If
    con.Close
End Sub
Private Sub Form_Load()
    selectText txtQuantity
    GetUoms
    
    'On Error Resume Next
    cmbUnit.ListIndex = 0
End Sub

Private Sub txtQuantity_Change()
    If IsNumeric(txtQuantity.Text) = False Then
        txtQuantity.Text = "1"
        selectText txtQuantity
    End If
End Sub

Private Sub txtQuantity_GotFocus()
    selectText txtQuantity
End Sub
