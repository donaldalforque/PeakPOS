VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form INV_QuantityPricingFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6240
   Icon            =   "INV_QuantityPricingFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtQtyEnd 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   330
      Left            =   3240
      MaxLength       =   50
      TabIndex        =   2
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox txtQtyBegin 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   330
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   1
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox txtPrice 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Height          =   330
      Left            =   4800
      MaxLength       =   50
      TabIndex        =   3
      Top             =   6000
      Width           =   1215
   End
   Begin MSComctlLib.ListView lvPricing 
      Height          =   4095
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   5775
      _ExtentX        =   10186
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "LocationId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Qty Begin"
         Object.Width           =   6253
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Qty End"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Price"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5400
      Top             =   600
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
            Picture         =   "INV_QuantityPricingFrm.frx":000C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_QuantityPricingFrm.frx":686E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_QuantityPricingFrm.frx":D0D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "INV_QuantityPricingFrm.frx":13932
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tb_Standard 
      Height          =   330
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   582
      ButtonWidth     =   1588
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
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qty. End:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2400
      TabIndex        =   8
      Top             =   6045
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Qty. Begin:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   240
      TabIndex        =   7
      Top             =   6045
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Price:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   4320
      TabIndex        =   6
      Top             =   6045
      Width           =   435
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Product Quantity Pricing"
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
      Left            =   960
      TabIndex        =   5
      Top             =   765
      Width           =   2820
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Provide flexible product pricing to your clients based on the number of pieces they purchase."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   450
      Left            =   360
      TabIndex        =   4
      Top             =   1290
      Width           =   4815
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "INV_QuantityPricingFrm.frx":1A194
      Stretch         =   -1  'True
      Top             =   720
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   6015
      Left            =   120
      Top             =   480
      Width           =   6015
   End
End
Attribute VB_Name = "INV_QuantityPricingFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim PricingId As Long
Public Sub Initialize()
    'cmbUom.text = ""
    PricingId = 0
    txtQtyBegin.SetFocus
    txtPrice.Text = "1"
    txtQtyBegin.Text = ""
    txtQtyEnd.Text = ""
End Sub
Public Sub Populate(ByVal data As String)
    Dim Item As MSComctlLib.ListItem
    Select Case data
        Case "ConversionLoad"
            Dim con As New ADODB.Connection
            Set rec = New ADODB.Recordset
            Set cmd = New ADODB.Command
            
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "POS_ProductPricing_Get"
            cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , INV_NewProductFrm.ProductId)
            Set rec = cmd.Execute
            lvPricing.ListItems.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                     Set Item = lvPricing.ListItems.add(, , "")
                            Item.Text = rec!POS_ProductPricingId
                            Item.SubItems(1) = rec!quantitybegin
                            Item.SubItems(2) = rec!quantityend
                            Item.SubItems(3) = FormatNumber(rec!price, 2, vbTrue, vbFalse)
                            
                    rec.MoveNext
                Loop
            End If
            con.Close
    
    End Select
End Sub

Private Sub chkShow_Click()
    Dim Item As MSComctlLib.ListItem
    Set rec = New ADODB.Recordset
    Set rec = Global_Data("Tax")
    lvPricing.ListItems.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            If chkShow.value = 1 Then
                Set Item = lvPricing.ListItems.add(, , "")
                    Item.SubItems(1) = rec!PricingId
                    Item.SubItems(2) = rec!TaxName
                    Item.SubItems(3) = rec!percentage
                If rec!isActive = "True" Then Item.Checked = True
                lvPricing.ColumnHeaders(1).width = lvPricing.width * 0.06
                lvPricing.ColumnHeaders(3).width = lvPricing.width * 0.44
                lvPricing.ColumnHeaders(4).width = lvPricing.width * 0.44
            Else
                If rec!isActive = "True" Then
                    Set Item = lvPricing.ListItems.add(, , "")
                        Item.SubItems(1) = rec!PricingId
                        Item.SubItems(2) = rec!TaxName
                        Item.SubItems(3) = rec!percentage
                    If rec!isActive = "True" Then Item.Checked = True
                    lvPricing.ColumnHeaders(1).width = lvPricing.width * 0
                    lvPricing.ColumnHeaders(3).width = lvPricing.width * 0.47
                    lvPricing.ColumnHeaders(4).width = lvPricing.width * 0.47
                End If
            End If
            rec.MoveNext
        Loop
    End If
End Sub

Private Sub Form_Load()
    lvPricing.ColumnHeaders(1).width = lvPricing.width * 0
    lvPricing.ColumnHeaders(2).width = lvPricing.width * 0.323
    lvPricing.ColumnHeaders(3).width = lvPricing.width * 0.323
    lvPricing.ColumnHeaders(4).width = lvPricing.width * 0.323
    Populate "Uom"
    Populate "ConversionLoad"
End Sub


Private Sub Form_Unload(Cancel As Integer)
    PricingId = 0
End Sub

Private Sub lvPricing_ItemClick(ByVal Item As MSComctlLib.ListItem)
    PricingId = Item.Text
    On Error Resume Next
    txtQtyBegin.Text = Item.SubItems(1)
    txtQtyEnd.Text = Item.SubItems(2)
    txtPrice.Text = Item.SubItems(3)
    txtQtyBegin.SetFocus
End Sub

Private Sub tb_Standard_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo ErrorHandler:
    Select Case Button.Index
        Case 1 'NEW
            Initialize
        Case 2 'Save
            Dim Item As MSComctlLib.ListItem
            Set con = New ADODB.Connection
            Set rec = New ADODB.Recordset
            
            con.ConnectionString = ConnString
            con.Open
            
        
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.Parameters.Append cmd.CreateParameter("@POS_ProductPricingId", adInteger, adParamInputOutput, , PricingId)
            cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , INV_NewProductFrm.ProductId)
            cmd.Parameters.Append cmd.CreateParameter("@QuantityBegin", adDecimal, adParamInput, , NVAL(txtQtyBegin.Text))
                                  cmd.Parameters("@QuantityBegin").NumericScale = 2
                                  cmd.Parameters("@QuantityBegin").Precision = 18
            cmd.Parameters.Append cmd.CreateParameter("@QuantityEnd", adDecimal, adParamInput, , NVAL(txtQtyEnd.Text))
                                  cmd.Parameters("@QuantityEnd").NumericScale = 2
                                  cmd.Parameters("@QuantityEnd").Precision = 18
             cmd.Parameters.Append cmd.CreateParameter("@Price", adDecimal, adParamInput, , NVAL(txtPrice.Text))
                                  cmd.Parameters("@Price").NumericScale = 2
                                  cmd.Parameters("@Price").Precision = 18
            
            If PricingId = 0 Then
                cmd.CommandText = "POS_ProductPricing_Insert"
                cmd.Execute
                PricingId = cmd.Parameters("@POS_ProductPricingId")
                Populate "ConversionLoad"
            Else
                cmd.CommandText = "POS_ProductPricing_Update"
                cmd.Execute
                For Each Item In lvPricing.ListItems
                    If Item.Text = PricingId Then
                        Item.SubItems(1) = txtQtyBegin.Text
                        Item.SubItems(2) = txtQtyEnd.Text
                        Item.SubItems(3) = FormatNumber(Val(txtPrice.Text), 2, vbTrue, vbFalse)
                        Item.Selected = True
                        Item.EnsureVisible
                    End If
                Next
            End If
            con.Close
        Case 4 'Delete
            Dim x As Variant
            x = MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo)
            If x = vbYes Then
                If lvPricing.ListItems.count > 0 Then
                    'Dim item As MSComctlLib.ListItem
                    Set con = New ADODB.Connection
                    Set cmd = New ADODB.Command
                    
                    con.ConnectionString = ConnString
                    con.Open
                    cmd.ActiveConnection = con
                    cmd.CommandType = adCmdStoredProc
                    cmd.CommandText = "POS_ProductPricing_Delete"
                    cmd.Parameters.Append cmd.CreateParameter("@POS_ProductPricingId", adInteger, adParamInput, , lvPricing.SelectedItem.Text)
                    cmd.Execute
                    con.Close
                    
                    lvPricing.ListItems.Remove (lvPricing.SelectedItem.Index)
                    Initialize
                End If
            End If
    End Select
    Exit Sub
ErrorHandler:
    If IsNumeric(Err.Description) = True Then
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & ErrorCodes(Val(Err.Description))
    Else
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(0) & " " & Err.Description
    End If
    GLOBAL_MessageFrm.Show (1)
End Sub

Private Sub txtPrice_Change()
    If IsNumeric(txtPrice.Text) = False Then
        txtPrice.Text = 0
    End If
End Sub

Private Sub cmbUom_GotFocus()
'    selectText cmbUom
End Sub






Private Sub txtPrice_GotFocus()
    selectText txtPrice
End Sub

Private Sub txtQtyBegin_GotFocus()
    selectText txtQtyBegin
End Sub

Private Sub txtQtyEnd_GotFocus()
    selectText txtQtyEnd
End Sub
