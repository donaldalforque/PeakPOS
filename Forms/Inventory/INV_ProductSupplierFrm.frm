VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form INV_ProductSupplierFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8655
   Icon            =   "INV_ProductSupplierFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkLastPurchaseDate 
      Caption         =   "Empty Date"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   4
      Top             =   6720
      Width           =   1455
   End
   Begin VB.TextBox txtLastPurchaseCost 
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
      Left            =   2160
      TabIndex        =   2
      Top             =   6360
      Width           =   2775
   End
   Begin VB.ComboBox cmbVendor 
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
      Left            =   2160
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   6000
      Width           =   6375
   End
   Begin MSComctlLib.Toolbar tb_Standard 
      Height          =   330
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   582
      ButtonWidth     =   1349
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Accounts"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8040
      Top             =   6480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_ProductSupplierFrm.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_ProductSupplierFrm.frx":686E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_ProductSupplierFrm.frx":D0D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_ProductSupplierFrm.frx":13932
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtLastPurchaseDate 
      Height          =   345
      Left            =   2160
      TabIndex        =   3
      Top             =   6720
      Width           =   2775
      _ExtentX        =   4895
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
      Format          =   148635649
      CurrentDate     =   41686
   End
   Begin MSComctlLib.ListView lvProductSuppliers 
      Height          =   4095
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   7223
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
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ProductSupplierId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "ProductId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "VendorId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Supplier"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Last Order Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Last Cost"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblAverageCost 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Average Cost: 0.00"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6360
      TabIndex        =   10
      Top             =   5400
      Width           =   1875
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Purchase Date"
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
      Top             =   6720
      Width           =   1725
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Purchase Cost"
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
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier"
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
      Top             =   6000
      Width           =   780
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "INV_ProductSupplierFrm.frx":1A194
      Top             =   600
      Width           =   480
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product Suppliers"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   345
      Left            =   840
      TabIndex        =   0
      Top             =   645
      Width           =   2040
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   5295
      Left            =   120
      Top             =   480
      Width           =   8415
   End
End
Attribute VB_Name = "INV_ProductSupplierFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Populate(ByVal data As String)
    Dim item As MSComctlLib.ListItem
    Select Case data
       Case "Vendor"
            Set rec = Global_Data("Vendor")
            cmbVendor.Clear
            cmbVendor.AddItem ""
            cmbVendor.ItemData(cmbVendor.NewIndex) = 0
            If Not rec.EOF Then
                Do Until rec.EOF
                    cmbVendor.AddItem rec!Name
                    cmbVendor.ItemData(cmbVendor.NewIndex) = rec!VendorId
                    rec.MoveNext
                Loop
            End If
        Case "ProductSuppliers"
            Dim con As New ADODB.Connection
            Set cmd = New ADODB.Command
            Set rec = New ADODB.Recordset
            
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "INV_ProductSupplier_Get"
            cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , INV_NewProductFrm.ProductId)
            Set rec = cmd.Execute
            lvProductSuppliers.ListItems.Clear
            If Not rec.EOF Then
               'Dim item As MSComctlLib.ListItem
                Do Until rec.EOF
                    Set item = lvProductSuppliers.ListItems.add(, , rec!ProductVendorId)
                        item.SubItems(1) = rec!ProductId
                        item.SubItems(2) = rec!VendorId
                        item.SubItems(3) = rec!Supplier
                        item.SubItems(4) = rec!lastpurchasedate
                        item.SubItems(5) = FormatNumber(rec!lastpurchaseCost, 2, vbTrue, vbFalse)
                    rec.MoveNext
                Loop
            End If
            con.Close
    End Select
End Sub
Private Sub chkLastPurchaseDate_Click()
    If chkLastPurchaseDate.value = Checked Then
        dtLastPurchaseDate.Enabled = False
    Else
        dtLastPurchaseDate.Enabled = True
    End If
End Sub

Public Sub CountAverageCost()
    Dim total, count As Integer
    
    For Each item In lvProductSuppliers.ListItems
        If NVAL(item.SubItems(5)) > 0 Then
            total = total + NVAL(item.SubItems(5))
            count = count + 1
        End If
    Next
    total = total / count
    lblAverageCost.Caption = "Average Cost: " & FormatNumber(total, 2, vbTrue, vbFalse)
End Sub

Private Sub Form_Load()
    lvProductSuppliers.ColumnHeaders(4).width = lvProductSuppliers.width * 0.6
    lvProductSuppliers.ColumnHeaders(5).width = lvProductSuppliers.width * 0.2
    lvProductSuppliers.ColumnHeaders(6).width = lvProductSuppliers.width * 0.16

    Initialize

    dtLastPurchaseDate.value = Format(Now, "MM/DD/YY")
    
    Populate "Vendor"
    Populate "ProductSuppliers"
    
    CountAverageCost
End Sub

Public Sub Initialize()
    txtLastPurchaseCost.Text = "0.00"
    dtLastPurchaseDate.value = Format(Now, "MM/DD/YY")
    
    On Error Resume Next
    cmbVendor.ListIndex = 0
End Sub

Private Sub tb_Standard_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'NEW
            Initialize
        Case 2 'Save
            If cmbVendor.ListIndex <= 0 Then
                MsgBox "Please select a supplier.", vbCritical
                cmbVendor.SetFocus
                Exit Sub
            End If
        
            Dim con As New ADODB.Connection
            Set cmd = New ADODB.Command
            Set rec = New ADODB.Recordset
            
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.Parameters.Append cmd.CreateParameter("@ProductVendorId", adInteger, adParamInputOutput, , 0)
            cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , INV_NewProductFrm.ProductId)
            cmd.Parameters.Append cmd.CreateParameter("@VendorId", adInteger, adParamInput, , cmbVendor.ItemData(cmbVendor.ListIndex))
            cmd.Parameters.Append cmd.CreateParameter("@LastPurchaseCost", adDecimal, adParamInput, , NVAL(txtLastPurchaseCost.Text))
                                  cmd.Parameters("@LastPurchaseCost").NumericScale = 2
                                  cmd.Parameters("@LastPurchaseCost").Precision = 18
            If chkLastPurchaseDate.value = Checked Then
                cmd.Parameters.Append cmd.CreateParameter("@LastPurchaseDate", adDate, adParamInput, , "")
            Else
                cmd.Parameters.Append cmd.CreateParameter("@LastPurchaseDate", adDate, adParamInput, , dtLastPurchaseDate.value)
            End If
            cmd.CommandText = "INV_ProductSupplier_Insert"
            Set rec = cmd.Execute
            
            'Add/update list
            Dim item As MSComctlLib.ListItem
            
            For Each item In lvProductSuppliers.ListItems
                If item.SubItems(2) = cmbVendor.ItemData(cmbVendor.ListIndex) Then
                    item.SubItems(4) = dtLastPurchaseDate.value
                    item.SubItems(5) = FormatNumber(txtLastPurchaseCost.Text, 2, vbTrue, vbFalse)
                    Exit Sub
                End If
            Next
            
            
            Set item = lvProductSuppliers.ListItems.add(, , cmd.Parameters("@ProductVendorId"))
                item.SubItems(1) = INV_NewProductFrm.ProductId
                item.SubItems(2) = cmbVendor.ItemData(cmbVendor.ListIndex)
                item.SubItems(3) = cmbVendor.Text
                If chkLastPurchaseDate.value = Unchecked Then item.SubItems(4) = dtLastPurchaseDate.value
                item.SubItems(5) = FormatNumber(txtLastPurchaseCost.Text, 2, vbTrue, vbFalse)
                
            con.Close
    End Select
End Sub
