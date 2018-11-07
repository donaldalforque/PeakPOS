VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form SO_PurchaseReturnStatusFrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Purchase Return Status"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Body_Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      Begin VB.Frame Frame_Body 
         BackColor       =   &H00FFFFFF&
         Height          =   3780
         Left            =   120
         TabIndex        =   32
         Top             =   2280
         Width           =   10215
         Begin VB.TextBox txtCode 
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
            Left            =   -9999
            TabIndex        =   33
            Top             =   240
            Visible         =   0   'False
            Width           =   3015
         End
         Begin MSComctlLib.ListView lvItems 
            Height          =   3375
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   5953
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
            NumItems        =   13
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "PurchaseOrderLineId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "PurchaseOrderId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Item Code"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Name"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Quantity"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Unit"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "Cost"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Text            =   "Sub-Total"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "LocationId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "ProductId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   10
               Text            =   "PrevQuantity"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   11
               Text            =   "ReserveId"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   12
               Text            =   "ActualQuantity"
               Object.Width           =   0
            EndProperty
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Code"
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
            Left            =   -9999
            TabIndex        =   35
            Top             =   240
            Visible         =   0   'False
            Width           =   480
         End
      End
      Begin VB.Frame Frame_Footer 
         BackColor       =   &H00FFFFFF&
         Height          =   2415
         Left            =   120
         TabIndex        =   21
         Top             =   6120
         Width           =   10215
         Begin VB.TextBox txtRemarks 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1410
            Left            =   1200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   23
            Top             =   240
            Width           =   5655
         End
         Begin VB.TextBox txtCash 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   -9999
            TabIndex        =   22
            Top             =   240
            Visible         =   0   'False
            Width           =   1695
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
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
            Left            =   8160
            TabIndex        =   31
            Top             =   585
            Width           =   1860
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL"
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
            Left            =   7200
            TabIndex        =   30
            Top             =   585
            Width           =   660
         End
         Begin VB.Label lblSubTotal 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   8400
            TabIndex        =   29
            Top             =   240
            Width           =   1620
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SUB-TOTAL"
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
            Left            =   7200
            TabIndex        =   28
            Top             =   240
            Width           =   1020
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
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
            TabIndex        =   27
            Top             =   240
            Width           =   795
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CASH"
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
            Left            =   -9999
            TabIndex        =   26
            Top             =   240
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "REM. BAL."
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   7200
            TabIndex        =   25
            Top             =   1125
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.Label lblRemainingBalance 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   8280
            TabIndex        =   24
            Top             =   1125
            Visible         =   0   'False
            Width           =   1740
         End
      End
      Begin VB.Frame Frame_Header2 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   1395
         Left            =   7050
         TabIndex        =   12
         Top             =   120
         Width           =   3285
         Begin VB.TextBox txtOrderNumber 
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
            Left            =   1080
            TabIndex        =   14
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox txtStatus 
            BackColor       =   &H00E0E0E0&
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
            Left            =   1080
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   960
            Width           =   2055
         End
         Begin MSComCtl2.DTPicker dtOrder 
            Height          =   330
            Left            =   1080
            TabIndex        =   15
            Top             =   600
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   89522177
            CurrentDate     =   41509
         End
         Begin MSComCtl2.DTPicker dtReceived 
            Height          =   330
            Left            =   1080
            TabIndex        =   16
            Top             =   1560
            Visible         =   0   'False
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   582
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   89522177
            CurrentDate     =   41509
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Return #"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   780
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   435
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   18
            Top             =   960
            Width           =   570
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Received"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   120
            TabIndex        =   17
            Top             =   1560
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VB.Frame Frame_Header1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Enabled         =   0   'False
         Height          =   2055
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   5655
         Begin VB.ComboBox cmbVendor 
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
            Left            =   1200
            TabIndex        =   6
            Top             =   0
            Width           =   3015
         End
         Begin VB.TextBox txtContact 
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
            TabIndex        =   5
            Top             =   720
            Width           =   3015
         End
         Begin VB.TextBox txtAddress 
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   930
            Left            =   1200
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   1080
            Width           =   3015
         End
         Begin VB.Label Label7 
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
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   780
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Phone"
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
            Left            =   0
            TabIndex        =   10
            Top             =   720
            Width           =   600
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
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
            Left            =   0
            TabIndex        =   9
            Top             =   1080
            Width           =   750
         End
         Begin VB.Label lblBalance 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
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
            Left            =   1200
            TabIndex        =   8
            Top             =   390
            Width           =   3015
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Balance"
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
            Left            =   0
            TabIndex        =   7
            Top             =   360
            Width           =   720
         End
      End
      Begin VB.CommandButton btnReceiveOrder 
         Caption         =   "Receive Order"
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
         Left            =   4200
         TabIndex        =   1
         Top             =   8040
         Visible         =   0   'False
         Width           =   2925
      End
      Begin MSComctlLib.ListView lvVendor 
         Height          =   2655
         Left            =   -99999
         TabIndex        =   2
         Top             =   930
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   4683
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
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
            Text            =   "CustomerId"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Balance"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Contact"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Address"
            Object.Width           =   0
         EndProperty
      End
   End
End
Attribute VB_Name = "SO_PurchaseReturnStatusFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PurchaseReturnId As String
Public SalesReturnId As String

Public Sub Populate(ByVal data As String)
    Select Case data
        Case "Vendor"
            Set rec = New ADODB.Recordset
            Set cmd = New ADODB.Command
            Set con = New ADODB.Connection
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "BASE_Vendor_Get"
            cmd.Parameters.Append cmd.CreateParameter("@VendorId", adInteger, adParamInput, , VendorId)
            Set rec = cmd.Execute
            If Not rec.EOF Then
                Do Until rec.EOF
                    lblBalance.Caption = FormatNumber(rec!OutStandingBalance, 2, vbTrue, vbFalse)
                    rec.MoveNext
                Loop
            End If
            con.Close
        Case "PurchaseReturnGet"
            Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            Set rec = New ADODB.Recordset
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "PO_PurchaseReturn_Get"
            cmd.Parameters.Append cmd.CreateParameter("@PurchaseReturnId", adInteger, adParamInput, , PurchaseReturnId)
            Set rec = cmd.Execute
            If Not rec.EOF Then
                'If rec!isCashAdvance = "False" Then
                    VendorId = rec!VendorId
                    'lblBalance.Caption = FormatNumber(rec!balance, 2, vbTrue, vbFalse)
                    cmbVendor.BackColor = vbWhite
                    StatusId = rec!StatusId
                    txtStatus.Text = rec!Status
                    txtOrderNumber.Text = rec!OrderNumber
                    txtAddress.Text = rec!Address
                    txtContact.Text = rec!Phone
                    dtOrder.value = Format(rec!Date, "MM/DD/YY")
                    'dtReceived.value = Format(rec!duedate, "MM/DD/YY")
                    lblSubTotal.Caption = FormatNumber(rec!subtotal, 2, vbTrue)
                    txtOrderNumber.Text = rec!OrderNumber
                    On Error Resume Next
                    'txtDays.text = rec!days
                    'txtInterestRate.text = rec!interestrate
                    'txtCash.text = FormatNumber(rec!cash, 2, vbTrue, vbFalse)
                    'On Error Resume Next
                    
'                    If rec!StatusId = 2 Then 'COMPLETED
'                        isNotCompleted (False)
'                        'picCompleted.Left = 6840
'                        'picCompleted.Visible = True
'                        ''pic_Cancelled.Left = -9999
'                        'picPaid.Left = -9999
'                        ''picCompleted.Visible = False
'                        ''pic_Cancelled.Visible = False
'                        'picPaid.Visible = False
'                        tb_Standard.Buttons(4).Caption = "Cancel"
'                        tb_Standard.Buttons(4).Image = 3
'                        'btnPaid.Visible = True
'                        btnStatus.Caption = "Reopen Order"
'                        btnStatus.Visible = False
'                        Frame_Footer.Enabled = False
'                    ElseIf rec!StatusId = 7 Then 'Cancelled
'                        isNotCompleted (False)
'                        'picCompleted.Left = -9999
'                        ''pic_Cancelled.Left = 6360
'                        'picPaid.Left = -9999
'                        'picPaid.Visible = False
'                        'picCompleted.Visible = False
'                        ''pic_Cancelled.Visible = True
'                        tb_Standard.Buttons(4).Caption = "Activate"
'                        tb_Standard.Buttons(4).Image = 6
'                        'btnPaid.Visible = False
'                        btnStatus.Caption = "Complete Order"
'                        btnStatus.Visible = False
'                    ElseIf rec!StatusId = 6 Then 'PAID
'                        isNotCompleted (False)
'                        'picPaid.Left = 6890
'                        'picCompleted.Left = -9999
'                        ''pic_Cancelled.Left = -9999
'                        'picPaid.Visible = True
'                        'picCompleted.Visible = False
'                        ''pic_Cancelled.Visible = False
'                        tb_Standard.Buttons(4).Caption = "Cancel"
'                        tb_Standard.Buttons(4).Image = 3
'                        'btnPaid.Visible = False
'                        btnStatus.Caption = "Reopen Order"
'                    Else
'                        isNotCompleted (True)
'                        'picCompleted.Left = -9999
'                        ''pic_Cancelled.Left = -9999
'                        'picPaid.Left = -9999
'                        'picCompleted.Visible = False
'                        ''pic_Cancelled.Visible = False
'                        'picPaid.Visible = False
'                        tb_Standard.Buttons(4).Caption = "Cancel"
'                        tb_Standard.Buttons(4).Image = 3
'                        'btnPaid.Visible = False
'                        btnStatus.Caption = "Complete Order"
'                        btnStatus.Visible = True
'                        Frame_Footer.Enabled = True
'                    End If
                    lblRemainingBalance.Caption = FormatNumber(rec!OutStandingBalance, 2, vbTrue, vbFalse)
                    'PREVENT ERROR ON DISPLAY FOR cmbVendor_Change Event
                    On Error Resume Next
                    cmbVendor.Text = rec!Name
                    lvVendor.Left = -9999
                    lvVendor.Visible = False
                'End If
            End If
            con.Close
            
            'GetTagged
        Case "PurchaseReturnLoad"
            Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            Set rec = New ADODB.Recordset
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "PO_PurchaseReturn_Get"
            Set rec = cmd.Execute
            Dim item As MSComctlLib.ListItem
            lvSearch.ListItems.Clear
            If Not rec.EOF Then
                Do Until rec.EOF
                    If rec!isCashAdvance = "False" Then
                        Set item = lvSearch.ListItems.add(, , rec!PurchaseReturnId)
                            item.SubItems(1) = rec!OrderNumber
                            item.SubItems(2) = rec!Status
                    End If
                    rec.MoveNext
                Loop
            End If
            con.Close
        Case "PurchaseReturnLineGet"
            Set con = New ADODB.Connection
            Set cmd = New ADODB.Command
            Set rec = New ADODB.Recordset
            con.ConnectionString = ConnString
            con.Open
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "PO_PurchaseReturnLine_Get"
            cmd.Parameters.Append cmd.CreateParameter("@PurchaseReturnId", adInteger, adParamInput, , PurchaseReturnId)
            Set rec = cmd.Execute
            lvItems.ListItems.Clear
            
            
            If Not rec.EOF Then
                Do Until rec.EOF
                    Set item = lvItems.ListItems.add(, , rec!PurchaseReturnLineId)
                        item.SubItems(1) = rec!PurchaseReturnId
                        item.SubItems(2) = rec!itemcode
                        item.SubItems(3) = rec!Name
                        item.SubItems(4) = FormatNumber(rec!quantity, 2, vbTrue)
                        item.SubItems(5) = rec!Uom
                        item.SubItems(6) = FormatNumber(rec!cost, 2, vbTrue)
                        item.SubItems(7) = FormatNumber(rec!subtotal, 2, vbTrue)
                        'item.SubItems(8) = rec!LocationId
                        item.SubItems(9) = rec!ProductId
                        item.SubItems(10) = rec!quantity
                        item.SubItems(12) = rec!ActualQuantity
                        On Error Resume Next
                        item.SubItems(11) = rec!ReserveId
                    rec.MoveNext
                Loop
            End If
            con.Close
    End Select
End Sub
Public Sub GetAssociatedPurchaseReturnId()
    Dim con As New ADODB.Connection
    con.ConnectionString = ConnString
    con.Open
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "PO_SalesReturnTag_Get"
    cmd.Parameters.Append cmd.CreateParameter("@SalesReturnId", adInteger, adParamInput, , Val(SalesReturnId))
    Set rec = cmd.Execute
    If Not rec.EOF Then
        PurchaseReturnId = rec!PurchaseReturnId
    Else
        PurchaseReturnId = 0
    End If
    con.Close
End Sub
Private Sub Form_Load()
    GetAssociatedPurchaseReturnId
    
    If Val(PurchaseReturnId) = 0 Then
        MsgBox "There are no associated purchase return for this order.", vbCritical
        Exit Sub
    End If
    
    Populate "PurchaseReturnLineGet"
    Populate "PurchaseReturnGet"
    
    lvItems.ColumnHeaders(3).width = lvItems.width * 0.14
    lvItems.ColumnHeaders(4).width = lvItems.width * 0.29
    lvItems.ColumnHeaders(5).width = lvItems.width * 0.12
    lvItems.ColumnHeaders(6).width = lvItems.width * 0.09
    lvItems.ColumnHeaders(7).width = lvItems.width * 0.13
    lvItems.ColumnHeaders(8).width = lvItems.width * 0.2
End Sub

