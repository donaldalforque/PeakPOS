VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form BASE_DiscountSettingsFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAmount 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Amount Based"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1800
      TabIndex        =   8
      Top             =   5430
      Width           =   1845
   End
   Begin VB.TextBox txtPercentage 
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
      Left            =   3720
      MaxLength       =   50
      TabIndex        =   7
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CheckBox chkTaxExempt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Tax Exempt"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3720
      TabIndex        =   6
      Top             =   5430
      Width           =   1365
   End
   Begin VB.CheckBox chkShow 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Show All"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4050
      TabIndex        =   1
      Top             =   1830
      Width           =   1000
   End
   Begin VB.TextBox txtName 
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
      Left            =   240
      MaxLength       =   50
      TabIndex        =   0
      Top             =   5880
      Width           =   3375
   End
   Begin MSComctlLib.ListView lvDiscount 
      Height          =   3135
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5530
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
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
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
         Text            =   "Pct(%)/Amt."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Tax Ex."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "AmountBased"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.Toolbar tb_Standard 
      Height          =   330
      Left            =   0
      TabIndex        =   3
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
      Left            =   4680
      Top             =   -120
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
            Picture         =   "BASE_DiscounSettingsFrm.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BASE_DiscounSettingsFrm.frx":6862
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BASE_DiscounSettingsFrm.frx":D0C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "BASE_DiscounSettingsFrm.frx":13926
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "You can create discounts for different customers or company events and promotions."
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
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   4815
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "POS Discounts"
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
      TabIndex        =   4
      Top             =   720
      Width           =   1680
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "BASE_DiscounSettingsFrm.frx":1A188
      Top             =   675
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   5895
      Left            =   120
      Top             =   480
      Width           =   5055
   End
End
Attribute VB_Name = "BASE_DiscountSettingsFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DiscountId As Integer
Public Sub Initialize()
    txtName.Text = ""
    txtPercentage.Text = 0
    DiscountId = 0
    txtName.SetFocus
End Sub
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
                            If rec!isTaxExempt = True Then item.SubItems(4) = "True" Else item.SubItems(4) = "False"
                            If rec!isAmountBased = True Then item.SubItems(5) = "True" Else item.SubItems(5) = "False"
                            item.Checked = True
                    End If
                    rec.MoveNext
                Loop
            End If
    End Select
End Sub

Private Sub chkShow_Click()
    Dim item As MSComctlLib.ListItem
    Set rec = New ADODB.Recordset
    Set rec = Global_Data("Discount")
    lvDiscount.ListItems.Clear
    If Not rec.EOF Then
        Do Until rec.EOF
            If chkShow.value = 1 Then
                Set item = lvDiscount.ListItems.add(, , "")
                    item.SubItems(1) = rec!DiscountId
                    item.SubItems(2) = rec!Name
                    item.SubItems(3) = rec!Percentage
                    item.SubItems(4) = "False"
                    If rec!isTaxExempt = True Then item.SubItems(4) = "True"
                    If rec!isAmountBased = True Then item.SubItems(5) = "True"
                If rec!isActive = "True" Then item.Checked = True
                lvDiscount.ColumnHeaders(1).width = lvDiscount.width * 0.06
                lvDiscount.ColumnHeaders(3).width = lvDiscount.width * 0.2933
                lvDiscount.ColumnHeaders(4).width = lvDiscount.width * 0.2933
                lvDiscount.ColumnHeaders(5).width = lvDiscount.width * 0.2933
            Else
                If rec!isActive = "True" Then
                    Set item = lvDiscount.ListItems.add(, , "")
                        item.SubItems(1) = rec!DiscountId
                        item.SubItems(2) = rec!Name
                        item.SubItems(3) = rec!Percentage
                        item.SubItems(4) = "False"
                        If rec!isTaxExempt = True Then item.SubItems(4) = "True"
                        If rec!isAmountBased = True Then item.SubItems(5) = "True"
                    If rec!isActive = "True" Then item.Checked = True
                    lvDiscount.ColumnHeaders(1).width = lvDiscount.width * 0
                    lvDiscount.ColumnHeaders(3).width = lvDiscount.width * 0.3133
                    lvDiscount.ColumnHeaders(4).width = lvDiscount.width * 0.3133
                    lvDiscount.ColumnHeaders(5).width = lvDiscount.width * 0.3133
                End If
            End If
            rec.MoveNext
        Loop
    End If
End Sub

Private Sub Form_Load()
    lvDiscount.ColumnHeaders(1).width = lvDiscount.width * 0
    lvDiscount.ColumnHeaders(3).width = lvDiscount.width * 0.3133
    lvDiscount.ColumnHeaders(4).width = lvDiscount.width * 0.3133
    lvDiscount.ColumnHeaders(5).width = lvDiscount.width * 0.3133
    Populate "Term"
End Sub


Private Sub lvDiscount_ItemClick(ByVal item As MSComctlLib.ListItem)
    DiscountId = item.SubItems(1)
    txtName.Text = item.SubItems(2)
    txtPercentage.Text = item.SubItems(3)
    If item.SubItems(4) = "False" Then
        chkTaxExempt.value = Unchecked
    Else
        chkTaxExempt.value = Checked
    End If
    If item.SubItems(5) = "False" Then
        chkAmount.value = Unchecked
    Else
        chkAmount.value = Checked
    End If
    txtName.SetFocus
End Sub

Private Sub tb_Standard_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo ErrorHandler:
    Select Case Button.Index
        Case 1 'NEW
            Initialize
        Case 2 'Save
            Dim item As MSComctlLib.ListItem
            Set con = New ADODB.Connection
            Set rec = New ADODB.Recordset
            con.ConnectionString = ConnString
            con.Open
            
            'Check for Deactivate/Activated Lists
            For Each item In lvDiscount.ListItems
                Set cmd = New ADODB.Command
                cmd.ActiveConnection = con
                cmd.CommandType = adCmdStoredProc
                cmd.CommandText = "BASE_Discount_Update"
                cmd.Parameters.Append cmd.CreateParameter("@DiscountId", adInteger, adParamInputOutput, , item.SubItems(1))
                cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, item.SubItems(2))
                cmd.Parameters.Append cmd.CreateParameter("@Percentage", adInteger, adParamInput, , item.SubItems(3))
                cmd.Parameters.Append cmd.CreateParameter("@isTaxExempt", adBoolean, adParamInput, , item.SubItems(4))
                cmd.Parameters.Append cmd.CreateParameter("@isAmountBased", adBoolean, adParamInput, , item.SubItems(5))
                cmd.Parameters.Append cmd.CreateParameter("@isActive", adBoolean, adParamInput, , item.Checked)
                cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
                cmd.Parameters.Append cmd.CreateParameter("@WorkStationId", adInteger, adParamInput, , WorkstationId)
                cmd.Execute
            Next
            
            If Trim(txtName.Text) = "" Then
                Exit Sub
            End If
        
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.Parameters.Append cmd.CreateParameter("@DiscountId", adInteger, adParamInputOutput, , DiscountId)
            cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, txtName.Text)
            cmd.Parameters.Append cmd.CreateParameter("@Percentage", adInteger, adParamInput, , txtPercentage.Text)
            If chkTaxExempt.value = Checked Then
                cmd.Parameters.Append cmd.CreateParameter("@isTaxExempt", adBoolean, adParamInput, , True)
            Else
                cmd.Parameters.Append cmd.CreateParameter("@isTaxExempt", adBoolean, adParamInput, , False)
            End If
            If chkAmount.value = Checked Then
                cmd.Parameters.Append cmd.CreateParameter("@isAmountBased", adBoolean, adParamInput, , True)
            Else
                cmd.Parameters.Append cmd.CreateParameter("@isAmountBased", adBoolean, adParamInput, , False)
            End If
            cmd.Parameters.Append cmd.CreateParameter("@isActive", adBoolean, adParamInput, , lvDiscount.SelectedItem.Checked)
            cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
            cmd.Parameters.Append cmd.CreateParameter("@WorkStationId", adInteger, adParamInput, , WorkstationId)
            If DiscountId = 0 Then
                cmd.CommandText = "BASE_Discount_Insert"
                cmd.Execute
                DiscountId = cmd.Parameters("@DiscountId")
                Set item = lvDiscount.ListItems.add(, , "")
                    item.SubItems(1) = DiscountId
                    item.SubItems(2) = txtName.Text
                    item.SubItems(3) = txtPercentage.Text
                    If chkTaxExempt.value = Checked Then
                        item.SubItems(4) = "True"
                    Else
                        item.SubItems(4) = "False"
                    End If
                    item.Checked = True
                    item.Selected = True
                    item.EnsureVisible
            Else
                cmd.CommandText = "BASE_Discount_Update"
                cmd.Execute
                For Each item In lvDiscount.ListItems
                    If item.SubItems(1) = DiscountId Then
                        item.SubItems(2) = txtName.Text
                        item.SubItems(3) = txtPercentage.Text
                        If chkTaxExempt.value = Checked Then
                            item.SubItems(4) = "True"
                        Else
                            item.SubItems(4) = "False"
                        End If
                        If chkAmount.value = Checked Then
                            item.SubItems(5) = "True"
                        Else
                            item.SubItems(5) = "False"
                        End If
                        item.Selected = True
                        item.EnsureVisible
                    End If
                Next
            End If
            con.Close
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

Private Sub txtPercentage_Change()
    If IsNumeric(txtPercentage.Text) = False And txtPercentage.Text <> "" Then
        txtPercentage.Text = 0
    Else
        txtPercentage.Text = Int(txtPercentage.Text)
    End If
End Sub

Private Sub txtName_GotFocus()
    selectText txtName
End Sub





