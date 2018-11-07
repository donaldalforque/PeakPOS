VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form POS_SalesReturnFrm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13815
   Icon            =   "POS_SalesReturnFrm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   13815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCash 
      Caption         =   "CASH"
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
      Left            =   120
      Picture         =   "POS_SalesReturnFrm.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton btnCredit 
      Caption         =   "CREDIT"
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
      Left            =   1680
      Picture         =   "POS_SalesReturnFrm.frx":062E
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8040
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker dtTimeTo 
      Height          =   375
      Left            =   10680
      TabIndex        =   25
      Top             =   2160
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   96993282
      CurrentDate     =   41686
   End
   Begin MSComCtl2.DTPicker dtTimeFrom 
      Height          =   375
      Left            =   7800
      TabIndex        =   24
      Top             =   2160
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   96993282
      CurrentDate     =   41686
   End
   Begin VB.CommandButton btnReturnSlips 
      Caption         =   "Return Slips"
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
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   8070
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancel"
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
      Left            =   12120
      Picture         =   "POS_SalesReturnFrm.frx":0C7C
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8070
      Width           =   1575
   End
   Begin VB.CommandButton btnAccept 
      Caption         =   "Return"
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
      Left            =   10440
      Picture         =   "POS_SalesReturnFrm.frx":300B
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8040
      Width           =   1575
   End
   Begin MSComctlLib.ListView lvCustomer 
      Height          =   3255
      Left            =   -9999
      TabIndex        =   20
      Top             =   2160
      Visible         =   0   'False
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   5741
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
         Object.Width           =   0
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
   Begin VB.TextBox txtComments 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   1560
      TabIndex        =   14
      Top             =   6960
      Width           =   4215
   End
   Begin VB.CommandButton btnSearchItem 
      Height          =   375
      Left            =   7215
      Picture         =   "POS_SalesReturnFrm.frx":53DF
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton btnSearchReceipt 
      Height          =   375
      Left            =   3015
      Picture         =   "POS_SalesReturnFrm.frx":5603
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox txtItem 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   1800
      Width           =   3615
   End
   Begin VB.TextBox txtReceipt 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1800
      Width           =   2655
   End
   Begin MSComctlLib.ListView lvItems 
      Height          =   4095
      Left            =   360
      TabIndex        =   8
      Top             =   2760
      Width           =   13095
      _ExtentX        =   23098
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
      NumItems        =   13
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "POSSalesLineId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "POSSalesId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "ProductId"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Disc"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Qty"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Sub-Total"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Return Qty (-)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Sales Return"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "ORNUMBER"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "TaxExempt"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Tax"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComCtl2.DTPicker DateTo 
      Height          =   375
      Left            =   10680
      TabIndex        =   17
      Top             =   1800
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   96993281
      CurrentDate     =   41686
   End
   Begin MSComCtl2.DTPicker DateFrom 
      Height          =   375
      Left            =   7800
      TabIndex        =   18
      Top             =   1800
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   96993281
      CurrentDate     =   41686
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   10320
      TabIndex        =   19
      Top             =   1800
      Width           =   225
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Filter Date:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   7800
      TabIndex        =   16
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Comments:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   360
      TabIndex        =   15
      Top             =   6960
      Width           =   2535
   End
   Begin VB.Label lblTotalSalesReturn 
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
      Height          =   330
      Left            =   11760
      TabIndex        =   13
      Top             =   7320
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   11760
      TabIndex        =   12
      Top             =   6960
      Width           =   1695
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total Sales Return:"
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
      Height          =   330
      Left            =   8520
      TabIndex        =   11
      Top             =   7320
      Width           =   3255
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Invoice Total:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   8520
      TabIndex        =   10
      Top             =   6960
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Items Purchased"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   360
      TabIndex        =   9
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Search by Item:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   3600
      TabIndex        =   5
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search by Receipt #:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   360
      TabIndex        =   3
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Use this to make a sales return done in the Point of Sale module. You can search by receipts, customers or even with date ranges."
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
      TabIndex        =   1
      Top             =   930
      Width           =   10455
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "POS Sales Return"
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
      TabIndex        =   0
      Top             =   450
      Width           =   1965
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "POS_SalesReturnFrm.frx":5827
      Top             =   405
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   7815
      Left            =   120
      Top             =   120
      Width           =   13575
   End
End
Attribute VB_Name = "POS_SalesReturnFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim orNumber As String

Private Sub btnAccept_Click()
    If EditAccessRights(15) = False Then
        MsgBox ErrorCodes(74), vbCritical, "Limited Rights"
        Exit Sub
    End If
    Dim x As Variant
    
    x = MsgBox("Are you sure all information are correct?", vbQuestion + vbYesNo, "Return Items")
    If x = vbNo Then Exit Sub
    
    Dim item As MSComctlLib.ListItem
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    
    con.ConnectionString = ConnString
    con.Open
    
    For Each item In lvItems.ListItems
        If Val(Replace(item.SubItems(8), ",", "")) > 0 Then
            Set cmd = New ADODB.Command
            cmd.ActiveConnection = con
            cmd.CommandType = adCmdStoredProc
            cmd.CommandText = "POS_SalesReturnLine_Insert"
            cmd.Parameters.Append cmd.CreateParameter("@OrderNumber", adVarChar, adParamInput, 50, item.SubItems(10))
            cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , item.SubItems(2))
            cmd.Parameters.Append cmd.CreateParameter("@POS_SalesId", adInteger, adParamInput, , item.SubItems(1))
            cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, item.SubItems(3))
            cmd.Parameters.Append cmd.CreateParameter("@Price", adDecimal, adParamInput, , Val(Replace(item.SubItems(4), ",", "")))
                                  cmd.Parameters("@Price").Precision = 18
                                  cmd.Parameters("@Price").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@LineDiscount", adDecimal, adParamInput, , Val(Replace(item.SubItems(5), ",", "")))
                                  cmd.Parameters("@LineDiscount").Precision = 18
                                  cmd.Parameters("@LineDiscount").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@QtySold", adDecimal, adParamInput, , Val(Replace(item.SubItems(6), ",", "")))
                                  cmd.Parameters("@QtySold").Precision = 18
                                  cmd.Parameters("@QtySold").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@QtyReturned", adDecimal, adParamInput, , Val(Replace(item.SubItems(8), ",", "")))
                                  cmd.Parameters("@QtyReturned").Precision = 18
                                  cmd.Parameters("@QtyReturned").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@SalesReturn", adDecimal, adParamInput, , Val(Replace(item.SubItems(9), ",", "")))
                                  cmd.Parameters("@SalesReturn").Precision = 18
                                  cmd.Parameters("@SalesReturn").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@TaxExempt", adDecimal, adParamInput, , Val(Replace(item.SubItems(11), ",", "")))
                                  cmd.Parameters("@TaxExempt").Precision = 18
                                  cmd.Parameters("@TaxExempt").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@Tax", adDecimal, adParamInput, , Val(Replace(item.SubItems(12), ",", "")))
                                  cmd.Parameters("@Tax").Precision = 18
                                  cmd.Parameters("@Tax").NumericScale = 2
            cmd.Parameters.Append cmd.CreateParameter("@POS_SalesLineId", adInteger, adParamInput, , item.Text)
            cmd.Parameters.Append cmd.CreateParameter("@Comment", adVarChar, adParamInput, 400, txtComments.Text)
            cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
            cmd.Parameters.Append cmd.CreateParameter("@WorkstationId", adInteger, adParamInput, , WorkstationId)
            cmd.Execute
        End If
    Next
    con.Close
    
    'print
    
    
    MsgBox "Sales return complete.", vbInformation, "Success"
    
    Dim Y As Variant
    Y = MsgBox("Print sales return receipt?", vbQuestion + vbYesNo)
    If Y = vbYes Then
        '**PRINT RECEIPT******
        Dim crxApp As New CRAXDRT.Application
        Dim crxRpt As New CRAXDRT.Report
        Set crxRpt = crxApp.OpenReport(App.Path & "\Reports\POS_SalesReturnSlip.rpt")
        crxRpt.RecordSelectionFormula = "{POS_SalesReturn.POS_SalesId}= " & Val(orNumber) & ""
        crxRpt.DiscardSavedData
        crxRpt.EnableParameterPrompting = False
        crxRpt.ParameterFields(1).AddCurrentValue ""
    
        Call ResetRptDB(crxRpt)
        crxRpt.PrintOut False
        '**END PRINT RECEIPT**
    End If
    
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub
Private Sub CountReturns()
    Dim item As MSComctlLib.ListItem
    Dim subtotal, Total As Double
    For Each item In lvItems.ListItems
        subtotal = (Val(Replace(item.SubItems(4), ",", "")) * Val(Replace(item.SubItems(8), ",", ""))) - Val(Replace(item.SubItems(5), ",", ""))
        Total = Total + subtotal
        item.SubItems(9) = FormatNumber(subtotal, 2, vbTrue, vbFalse)
    Next
    lblTotalSalesReturn.Caption = FormatNumber(Total, 2, vbTrue, vbFalse)
End Sub

Private Sub btnCash_Click()
    btnCash.Enabled = False
    btnCredit.Enabled = True
End Sub

Private Sub btnCredit_Click()
    btnCash.Enabled = True
    btnCredit.Enabled = False
End Sub

Private Sub btnReturnSlips_Click()
    POS_ReturnSlipsFrm.Show (1)
End Sub

Private Sub btnSearchCustomer_Click()
'    If Trim(txtCustomer.text) <> "" Then
'        Dim item As MSComctlLib.ListItem
'        Set con = New ADODB.Connection
'        Set rec = New ADODB.Recordset
'        Set cmd = New ADODB.Command
'        'Dim item As MSComctlLib.ListItem
'
'        con.ConnectionString = ConnString
'        con.Open
'        cmd.ActiveConnection = con
'        cmd.CommandType = adCmdStoredProc
'        cmd.CommandText = "BASE_Customer_Search"
'        cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, Trim(txtCustomer.text))
'        Set rec = cmd.Execute
'        If Not rec.EOF Then
'            lvCustomer.ListItems.Clear
'            lvCustomer.Left = 3600
'            lvCustomer.Visible = True
'            Do Until rec.EOF
'                If rec!isActive = "True" Then
'                    Set item = lvCustomer.ListItems.add(, , rec!CustomerId)
'                        item.SubItems(1) = rec!CustomerCode
'                        item.SubItems(2) = rec!Name
'                        item.SubItems(3) = FormatNumber(rec!balance, 2, vbTrue, vbFalse)
'                        item.SubItems(4) = rec!Phone
'                        item.SubItems(5) = rec!Address
'                End If
'                rec.MoveNext
'            Loop
'        Else
'            lvCustomer.Visible = False
'            lvCustomer.Left = -9999
'        End If
'        con.Close
'
''        If Trim(txtcustomer.text) = "" Then
''            txtcustomer.BackColor = &HC0C0FF
''        Else
''            txtcustomer.BackColor = vbWhite
''        End If
'    Else
'        lvCustomer.Visible = False
'        lvCustomer.Left = -9999
'    End If
End Sub

Private Sub btnSearchItem_Click()
    Set con = New ADODB.Connection
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "POS_SalesReturn_ItemSearch"
    cmd.Parameters.Append cmd.CreateParameter("@Description", adVarChar, adParamInput, 500, txtItem.Text)
    cmd.Parameters.Append cmd.CreateParameter("@DateFrom", adDate, adParamInput, , DateFrom.value)
    cmd.Parameters.Append cmd.CreateParameter("@DateTo", adDate, adParamInput, , DateTo.value)
    cmd.Parameters.Append cmd.CreateParameter("@TimeFrom", adVarChar, adParamInput, 50, dtTimeFrom.value)
    cmd.Parameters.Append cmd.CreateParameter("@TimeTo", adVarChar, adParamInput, 50, dtTimeTo.value)
    Set rec = cmd.Execute
    
    Dim item As MSComctlLib.ListItem
    lvItems.ListItems.Clear
    If Not rec.EOF Then
        'lblTotal.Caption = FormatNumber(rec!total, 2, vbTrue, vbFalse)
        Do Until rec.EOF
            orNumber = rec!POS_SalesId
            Dim subtot As Double
            subtot = (rec!price - rec!linediscount) * rec!quantity
            Set item = lvItems.ListItems.add(, , rec!pos_saleslineId)
                item.SubItems(1) = rec!POS_SalesId
                item.SubItems(2) = rec!ProductId
                item.SubItems(3) = rec!Name
                item.SubItems(4) = FormatNumber(rec!price, 2, vbTrue, vbFalse)
                item.SubItems(5) = FormatNumber(rec!linediscount, 2, vbTrue, vbFalse)
                item.SubItems(6) = FormatNumber(rec!quantity, 2, vbTrue, vbFalse)
                'Item.SubItems(7) = FormatNumber(subtot, 2, vbTrue, vbFalse)
                item.SubItems(7) = FormatNumber(rec!subtotal, 2, vbTrue, vbFalse)
                item.SubItems(10) = rec!pos_ordernumber 'UCase(txtReceipt.Text)
                item.SubItems(11) = rec!TaxExempt
                item.SubItems(12) = rec!tax
            rec.MoveNext
        Loop
    Else
        MsgBox "No related search found.", vbCritical, "No search found"
    End If
    con.Close
    CountReturns
End Sub

Private Sub btnSearchReceipt_Click()
    Set con = New ADODB.Connection
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "POS_SalesReturn_InvoiceSearch"
    cmd.Parameters.Append cmd.CreateParameter("@Invoice", adVarChar, adParamInput, 50, txtReceipt.Text)
    Set rec = cmd.Execute
    
    Dim item As MSComctlLib.ListItem
    lvItems.ListItems.Clear
    If Not rec.EOF Then
        'lblTotal.Caption = FormatNumber(rec!total, 2, vbTrue, vbFalse)
        Do Until rec.EOF
            orNumber = rec!POS_SalesId
            Dim subtot As Double
            subtot = (rec!price - rec!linediscount) * rec!quantity
            Set item = lvItems.ListItems.add(, , rec!pos_saleslineId)
                item.SubItems(1) = rec!POS_SalesId
                item.SubItems(2) = rec!ProductId
                item.SubItems(3) = rec!Name
                item.SubItems(4) = FormatNumber(rec!price, 2, vbTrue, vbFalse)
                item.SubItems(5) = FormatNumber(rec!linediscount, 2, vbTrue, vbFalse)
                item.SubItems(6) = FormatNumber(rec!quantity, 2, vbTrue, vbFalse)
                'Item.SubItems(7) = FormatNumber(subtot, 2, vbTrue, vbFalse)
                item.SubItems(7) = FormatNumber(rec!subtotal, 2, vbTrue, vbFalse)
                item.SubItems(10) = UCase(txtReceipt.Text)
                item.SubItems(11) = rec!TaxExempt
                item.SubItems(12) = rec!tax
            rec.MoveNext
        Loop
    Else
        MsgBox "No related search found.", vbCritical, "No search found"
    End If
    con.Close
    CountReturns
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            lvCustomer.Visible = False
            lvCustomer.Left = -9999
            txtReceipt.SetFocus
            selectText txtReceipt
    End Select
End Sub

Private Sub Form_Load()
    lvCustomer.ColumnHeaders(2).width = lvCustomer.width * 0.2
    lvCustomer.ColumnHeaders(3).width = lvCustomer.width * 0.75
    
    lvItems.ColumnHeaders(4).width = lvItems.width * 0.42
    lvItems.ColumnHeaders(5).width = lvItems.width * 0.1
    lvItems.ColumnHeaders(6).width = lvItems.width * 0.1
    lvItems.ColumnHeaders(7).width = lvItems.width * 0.1
    lvItems.ColumnHeaders(8).width = lvItems.width * 0.13
    lvItems.ColumnHeaders(9).width = lvItems.width * 0.13
    
    DateFrom.value = Format(Now, "mm/dd/yy")
    DateTo.value = Format(Now, "mm/dd/yy")
    dtTimeFrom.value = "00:00:00"
    dtTimeTo.value = "23:59:59"
    
    btnCredit_Click
End Sub

Private Sub lvItems_DblClick()
    With lvItems
        If .ListItems.Count > 0 Then
            Dim i As String
            i = InputBox("Input quantity to be returned.", "Returned Quantity", 1)
            If i = "" Then
                Exit Sub
            ElseIf IsNumeric(i) = False Then
                Exit Sub
            Else
                If Val(Replace(i, ",", "")) > Val(Replace(.SelectedItem.SubItems(6), ",", "")) Then
                    MsgBox "Quantity returned must not be greater than the purchased quantity.", vbCritical, "Invalid Quantity"
                    Exit Sub
                End If
                
                .SelectedItem.SubItems(8) = FormatNumber(i, 2, vbFalse, vbFalse)
                .SetFocus
                CountReturns
            End If
        End If
    End With
End Sub

Private Sub lvItems_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 13
        Call lvItems_DblClick
    End Select
End Sub

Private Sub txtCustomer_Change()
    btnSearchCustomer_Click
End Sub

Private Sub txtCustomer_GotFocus()
   ' selectText txtCustomer
End Sub

Private Sub txtCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    If lvCustomer.Visible = True Then
        Select Case KeyCode
            Case vbKeyDown
                lvCustomer.SetFocus
        End Select
    End If
End Sub

Private Sub txtReceipt_GotFocus()
    selectText txtReceipt
End Sub

Private Sub txtReceipt_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            btnSearchReceipt_Click
    End Select
End Sub
