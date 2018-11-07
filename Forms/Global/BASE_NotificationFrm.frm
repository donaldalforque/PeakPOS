VERSION 5.00
Begin VB.Form BASE_NotificationFrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6630
   ClientLeft      =   13245
   ClientTop       =   1065
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   4320
      Top             =   120
   End
   Begin VB.Label lblPurchase_OverdueNotification 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "- There are 0 order(s) that are already overdue."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   435
      Left            =   240
      MouseIcon       =   "BASE_NotificationFrm.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   5520
      Width           =   4335
   End
   Begin VB.Label lblPurchase_CheckDue 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "- You have 0 check(s) that are due for deposit today."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   435
      Left            =   240
      MouseIcon       =   "BASE_NotificationFrm.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   6000
      Width           =   4335
   End
   Begin VB.Label lblPurchasing_OpenSalesReturn 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "- You have 0 purchase return(s) that are pending."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   435
      Left            =   240
      MouseIcon       =   "BASE_NotificationFrm.frx":02A4
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   5040
      Width           =   4335
   End
   Begin VB.Label lblPurchasing_OpenStatus 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "- You have 0 purchase order(s) that are pending."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   435
      Left            =   240
      MouseIcon       =   "BASE_NotificationFrm.frx":03F6
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   4560
      Width           =   4335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Purchasing"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Top             =   4080
      Width           =   1050
   End
   Begin VB.Label lblSales_OpenSalesReturn 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "- You have 0 sales returns that are pending."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   435
      Left            =   240
      MouseIcon       =   "BASE_NotificationFrm.frx":0548
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   3600
      Width           =   4335
   End
   Begin VB.Label lblSales_CheckDue 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "- You have 0 check(s) that are due for deposit today."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   435
      Left            =   240
      MouseIcon       =   "BASE_NotificationFrm.frx":069A
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   3120
      Width           =   4335
   End
   Begin VB.Label lblSales_OpenStatus 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "- You have 0 order(s) that are waiting to be invoiced."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   435
      Left            =   240
      MouseIcon       =   "BASE_NotificationFrm.frx":07EC
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2640
      Width           =   4335
   End
   Begin VB.Label lblSales_OverDueNotification 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "- There are 0 order(s) that are already overdue."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   435
      Left            =   240
      MouseIcon       =   "BASE_NotificationFrm.frx":093E
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2160
      Width           =   4335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sales"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lblInventoryNotification 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "- You have 0 products that reached reorder point."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   465
      Left            =   240
      MouseIcon       =   "BASE_NotificationFrm.frx":0A90
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inventory"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "BASE_NotificationFrm.frx":0BE2
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Notifications"
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
      Left            =   720
      TabIndex        =   0
      Top             =   165
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   5775
      Left            =   120
      Top             =   720
      Width           =   4575
   End
End
Attribute VB_Name = "BASE_NotificationFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    GetNotification
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    lblInventoryNotification.FontUnderline = False
    lblPurchasing_OpenSalesReturn.FontUnderline = False
    lblPurchasing_OpenStatus.FontUnderline = False
    lblSales_CheckDue.FontUnderline = False
    lblSales_OpenSalesReturn.FontUnderline = False
    lblSales_OpenStatus.FontUnderline = False
    lblSales_OverDueNotification.FontUnderline = False
End Sub

Private Sub Label5_Click()

End Sub

Private Sub lblInventoryNotification_Click()
    If lblInventoryNotification.ForeColor = vbRed Then
        CornerChildForm INV_StockOnReorderPointFrm
        INV_StockOnReorderPointFrm.Show
        INV_StockOnReorderPointFrm.ZOrder 0
    End If
End Sub

Private Sub lblInventoryNotification_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If x >= 0 And x <= lblInventoryNotification.width Then
        lblInventoryNotification.FontUnderline = True
    End If
End Sub

Private Sub lblSystemSettings_Click()

End Sub

Private Sub lblPurchase_CheckDue_Click()
    If lblPurchase_CheckDue.ForeColor <> vbRed Then Exit Sub
    
    CornerChildForm RPT_PO_PaymentHistoryFrm
    RPT_PO_PaymentHistoryFrm.Show
    RPT_PO_PaymentHistoryFrm.ZOrder 0
    RPT_PO_PaymentHistoryFrm.btnGenerate_Click
End Sub

Private Sub lblPurchase_CheckDue_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If x >= 0 And x <= lblPurchase_CheckDue.width Then
        lblPurchase_CheckDue.FontUnderline = True
    End If
End Sub

Private Sub lblPurchase_OverdueNotification_Click()
    If lblPurchase_OverdueNotification.ForeColor <> vbRed Then Exit Sub
    
    Dim currdate As String
    currdate = "January 1," & Year(Now)
    
    CornerChildForm PO_PurchaseOrderFrm
    PO_PurchaseOrderFrm.Show
    PO_PurchaseOrderFrm.cmbSearch_Status.text = "Overdue"
    PO_PurchaseOrderFrm.DateFrom.value = currdate
    PO_PurchaseOrderFrm.btnSearch_Click
    PO_PurchaseOrderFrm.ZOrder 0
End Sub

Private Sub lblPurchase_OverdueNotification_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If x >= 0 And x <= lblPurchase_OverdueNotification.width Then
        lblPurchase_OverdueNotification.FontUnderline = True
    End If
End Sub

Private Sub lblPurchasing_OpenSalesReturn_Click()
    If lblPurchasing_OpenSalesReturn.ForeColor <> vbRed Then Exit Sub
    
    Dim currdate As String
    currdate = "January 1," & Year(Now)
    
    CornerChildForm PO_PurchaseReturnFrm
    PO_PurchaseReturnFrm.Show
    PO_PurchaseReturnFrm.cmbSearch_Status.text = "Open"
    PO_PurchaseReturnFrm.DateFrom.value = currdate
    PO_PurchaseReturnFrm.btnSearch_Click
    PO_PurchaseReturnFrm.ZOrder 0
End Sub

Private Sub lblPurchasing_OpenSalesReturn_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If x >= 0 And x <= lblPurchasing_OpenSalesReturn.width Then
        lblPurchasing_OpenSalesReturn.FontUnderline = True
    End If
End Sub

Private Sub lblPurchasing_OpenStatus_Click()
    If lblPurchasing_OpenStatus.ForeColor <> vbRed Then Exit Sub
    
    Dim currdate As String
    currdate = "January 1," & Year(Now)
    
    CornerChildForm PO_PurchaseOrderFrm
    PO_PurchaseOrderFrm.Show
    PO_PurchaseOrderFrm.cmbSearch_Status.text = "Open"
    PO_PurchaseOrderFrm.DateFrom.value = currdate
    PO_PurchaseOrderFrm.btnSearch_Click
    PO_PurchaseOrderFrm.ZOrder 0
End Sub

Private Sub lblPurchasing_OpenStatus_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If x >= 0 And x <= lblPurchasing_OpenStatus.width Then
        lblPurchasing_OpenStatus.FontUnderline = True
    End If
End Sub

Private Sub lblSales_CheckDue_Click()
    If lblSales_CheckDue.ForeColor <> vbRed Then Exit Sub
    
    CornerChildForm RPT_SO_CustomerPaymentDetailsFrm
    RPT_SO_CustomerPaymentDetailsFrm.Show
    RPT_SO_CustomerPaymentDetailsFrm.ZOrder 0
    RPT_SO_CustomerPaymentDetailsFrm.btnGenerate_Click
End Sub

Private Sub lblSales_CheckDue_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If x >= 0 And x <= lblSales_CheckDue.width Then
        lblSales_CheckDue.FontUnderline = True
    End If
End Sub

Private Sub lblSales_OpenSalesReturn_Click()
    If lblSales_OpenSalesReturn.ForeColor <> vbRed Then Exit Sub
    
    Dim currdate As String
    currdate = "January 1," & Year(Now)
    
    CornerChildForm SO_SalesReturnFrm
    SO_SalesReturnFrm.Show
    SO_SalesReturnFrm.cmbSearch_Status.text = "Open"
    SO_SalesReturnFrm.DateFrom.value = currdate
    SO_SalesReturnFrm.btnSearch_Click
    SO_SalesReturnFrm.ZOrder 0
End Sub

Private Sub lblSales_OpenSalesReturn_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If x >= 0 And x <= lblSales_OpenSalesReturn.width Then
        lblSales_OpenSalesReturn.FontUnderline = True
    End If
End Sub

Private Sub lblSales_OpenStatus_Click()
    If lblSales_OpenStatus.ForeColor <> vbRed Then Exit Sub
    Dim currdate As String
    currdate = "January 1," & Year(Now)
    
    CornerChildForm SO_SalesOrderFrm
    SO_SalesOrderFrm.Show
    SO_SalesOrderFrm.cmbSearch_Status.text = "Open"
    SO_SalesOrderFrm.DateFrom.value = currdate
    SO_SalesOrderFrm.btnSearch_Click
    SO_SalesOrderFrm.ZOrder 0
End Sub

Private Sub lblSales_OpenStatus_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If x >= 0 And x <= lblSales_OpenStatus.width Then
        lblSales_OpenStatus.FontUnderline = True
    End If
End Sub

Private Sub lblSales_OverDueNotification_Click()
    If lblSales_OverDueNotification.ForeColor <> vbRed Then Exit Sub
    
    Dim currdate As String
    currdate = "January 1," & Year(Now)
    
    CornerChildForm SO_SalesOrderFrm
    SO_SalesOrderFrm.Show
    SO_SalesOrderFrm.cmbSearch_Status.text = "Overdue"
    SO_SalesOrderFrm.DateFrom.value = currdate
    SO_SalesOrderFrm.btnSearch_Click
    SO_SalesOrderFrm.ZOrder 0
End Sub

Private Sub lblSales_OverDueNotification_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If x >= 0 And x <= lblSales_OverDueNotification.width Then
        lblSales_OverDueNotification.FontUnderline = True
    End If
End Sub
Private Sub GetNotification()
    'GET NOTIFICATIONS
    Dim Total As Long
    Total = 0
    
    'Inventory
    Total = Val(GetNotifications("Inventory"))
    If Total > 0 Then
        lblInventoryNotification.ForeColor = vbRed
    Else
        lblInventoryNotification.ForeColor = &H404040
    End If
    lblInventoryNotification.Caption = "For Order -" & FormatNumber(Total, 0, vbTrue, vbFalse) & " (products reached reorder point.)"
    
    'Sales OVERDUE STATUS
    Total = Val(GetNotifications("Overdue"))
    If Total > 0 Then
        lblSales_OverDueNotification.ForeColor = vbRed
    Else
        lblSales_OverDueNotification.ForeColor = &H404040
    End If
    lblSales_OverDueNotification.Caption = "Overdue AR - " & FormatNumber(Total, 0, vbTrue, vbFalse)
    
    'Sales OPEN STATUS
    Total = Val(GetNotifications("Open"))
    If Total > 0 Then
        lblSales_OpenStatus.ForeColor = vbRed
    Else
        lblSales_OpenStatus.ForeColor = &H404040
    End If
    lblSales_OpenStatus.Caption = "For Invoicing - " & FormatNumber(Total, 0, vbTrue, vbFalse)
    
    'Sales CHECK DUE
    Total = Val(GetNotifications("CheckDue"))
    If Total > 0 Then
        lblSales_CheckDue.ForeColor = vbRed
    Else
        lblSales_CheckDue.ForeColor = &H404040
    End If
    lblSales_CheckDue.Caption = "Check Payments for Deposit - " & FormatNumber(Total, 0, vbTrue, vbFalse)
    
    'Sales Return OPEN STATUS
    Total = Val(GetNotifications("SalesReturnOpen"))
    If Total > 0 Then
        lblSales_OpenSalesReturn.ForeColor = vbRed
    Else
        lblSales_OpenSalesReturn.ForeColor = &H404040
    End If
    lblSales_OpenSalesReturn.Caption = "Pending Sales Return - " & FormatNumber(Total, 0, vbTrue, vbFalse)
    
    'PURCHASE ORDER OPEN STATUS
    Total = Val(GetNotifications("PurchaseOrderOpen"))
    If Total > 0 Then
        lblPurchasing_OpenStatus.ForeColor = vbRed
    Else
        lblPurchasing_OpenStatus.ForeColor = &H404040
    End If
    lblPurchasing_OpenStatus.Caption = "Pending Order Delivery - " & FormatNumber(Total, 0, vbTrue, vbFalse)
    
    'Purchase Return OPEN STATUS
    Total = Val(GetNotifications("PurchaseReturnOpen"))
    If Total > 0 Then
        lblPurchasing_OpenSalesReturn.ForeColor = vbRed
    Else
        lblPurchasing_OpenSalesReturn.ForeColor = &H404040
    End If
    lblPurchasing_OpenSalesReturn.Caption = "Pending Purchase Return - " & FormatNumber(Total, 0, vbTrue, vbFalse)
    
    
    'PURCHASE CHECK DUE
    Total = Val(GetNotifications("POCheckDue"))
    If Total > 0 Then
        lblPurchase_CheckDue.ForeColor = vbRed
    Else
        lblPurchase_CheckDue.ForeColor = &H404040
    End If
    lblPurchase_CheckDue.Caption = "Check to be deposited by Supplier - " & FormatNumber(Total, 0, vbTrue, vbFalse)
    
    'PURCHASE OVERDUE STATUS
    Total = Val(GetNotifications("POOverdue"))
    If Total > 0 Then
        lblPurchase_OverdueNotification.ForeColor = vbRed
    Else
        lblPurchase_OverdueNotification.ForeColor = &H404040
    End If
    lblPurchase_OverdueNotification.Caption = "Overdue AP - " & FormatNumber(Total, 0, vbTrue, vbFalse)
End Sub
Private Sub Timer1_Timer()
    If NotificationTimer = 5 Then
        GetNotification
    Else
        NotificationTimer = NotificationTimer + 1
    End If
End Sub
