VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form INV_ProductExpiryFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3840
      TabIndex        =   2
      Top             =   2640
      Width           =   1335
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Save && Close"
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
      Left            =   2280
      TabIndex        =   1
      Top             =   2640
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Product Expiry"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.TextBox txtLot 
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
         Left            =   1800
         TabIndex        =   8
         Top             =   1440
         Width           =   3015
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
         Height          =   345
         Left            =   1800
         TabIndex        =   6
         Top             =   1000
         Width           =   3015
      End
      Begin VB.TextBox txtStockInDate 
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
         Left            =   1800
         TabIndex        =   4
         Top             =   600
         Width           =   3015
      End
      Begin MSComCtl2.DTPicker dtExpiry 
         Height          =   345
         Left            =   1800
         TabIndex        =   9
         Top             =   1880
         Width           =   3015
         _ExtentX        =   5318
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
         Format          =   87031809
         CurrentDate     =   41686
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Expiry Date"
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
         TabIndex        =   10
         Top             =   1875
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lot #"
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
         Top             =   1440
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reference #"
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
         TabIndex        =   5
         Top             =   1000
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stocked-in Date"
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
         TabIndex        =   3
         Top             =   600
         Width           =   1470
      End
   End
End
Attribute VB_Name = "INV_ProductExpiryFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ProductExpiryId As Integer

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnSave_Click()
    If txtLot.text = "" Then
        x = MsgBox("Lot number is empty. Do you want to continue saving?", vbQuestion + vbYesNo)
        If x = vbYes Then
            Save
            Unload Me
        End If
    Else
        Save
        Unload Me
    End If
End Sub

Private Sub Save()
    'SAVE
    Dim con As New ADODB.Connection
    con.ConnectionString = ConnString
    con.Open
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    
    cmd.Parameters.Append cmd.CreateParameter("@ProductExpiryId", adInteger, adParamInputOutput, , ProductExpiryId)
    cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , INV_NewProductFrm.ProductId)
    cmd.Parameters.Append cmd.CreateParameter("@StockedInDate", adVarChar, adParamInput, 50, txtStockInDate.text)
    cmd.Parameters.Append cmd.CreateParameter("@ReferenceNumber", adVarChar, adParamInput, 50, txtReference.text)
    cmd.Parameters.Append cmd.CreateParameter("@LotNumber", adVarChar, adParamInput, 50, txtLot.text)
    cmd.Parameters.Append cmd.CreateParameter("@ExpiryDate", adDate, adParamInput, , dtExpiry.value)
    
    If ProductExpiryId = 0 Then
        cmd.CommandText = "INV_ProductExpiry_Insert"
        cmd.Execute
        
        'ADD TO GRID
        Dim Item As MSComctlLib.ListItem
        Set Item = INV_ProductExtraInfoFrm.lvExpiry.ListItems.add(, , cmd.Parameters("@ProductExpiryId"))
            Item.SubItems(1) = INV_NewProductFrm.ProductId
            Item.SubItems(2) = txtStockInDate.text
            Item.SubItems(3) = txtReference.text
            Item.SubItems(4) = txtLot.text
            Item.SubItems(5) = dtExpiry.value
    Else
        cmd.CommandText = "INV_ProductExpiry_Update"
        cmd.Execute
        
        'UPDATE GRID
        With INV_ProductExtraInfoFrm.lvExpiry.SelectedItem
            .SubItems(2) = txtStockInDate.text
            .SubItems(3) = txtReference.text
            .SubItems(4) = txtLot.text
            .SubItems(5) = dtExpiry.value
        End With
    End If
    con.Close
End Sub
Private Sub Form_Load()
    If ProductExpiryId <> 0 Then
        LoadExpiryData
    Else
        txtStockInDate.text = ""
        txtReference.text = ""
        txtLot.text = ""
        dtExpiry.value = Format(Now, "MM/DD/YY")
    End If
End Sub

Private Sub LoadExpiryData()
    Dim Item As MSComctlLib.ListItem
    Dim con As New ADODB.Connection
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "INV_ProductExpiry_GetById"
    cmd.Parameters.Append cmd.CreateParameter("@ProductExpiryId", adInteger, adParamInput, , ProductExpiryId)
    Set rec = cmd.Execute
    If Not rec.EOF Then
        Do Until rec.EOF
            txtStockInDate.text = rec!stockedindate
            txtReference.text = rec!Reference
            txtLot.text = rec!lotnumber
            dtExpiry.value = Format(rec!expirydate, "MM/DD/YY")
            rec.MoveNext
        Loop
    End If
    con.Close
End Sub
