VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form GEN_PatchExtendedFRM 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   3900
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imgList_Main 
      Left            =   120
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GEN_PatchExtendedFRM.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GEN_PatchExtendedFRM.frx":065E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GEN_PatchExtendedFRM.frx":0BE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GEN_PatchExtendedFRM.frx":1286
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GEN_PatchExtendedFRM.frx":1842
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GEN_PatchExtendedFRM.frx":1F12
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GEN_PatchExtendedFRM.frx":2468
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GEN_PatchExtendedFRM.frx":2B09
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "GEN_PatchExtendedFRM.frx":313D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar_Main 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   1429
      ButtonWidth     =   1746
      ButtonHeight    =   1429
      ToolTips        =   0   'False
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "imgList_Main"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Home"
            Key             =   "Homepage"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Inventory"
            ImageIndex      =   6
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   16
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "NewProduct"
                  Text            =   "New Product"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ProductList"
                  Text            =   "Product List"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ProductCategories"
                  Text            =   "Product Categories"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "ProductConversion"
                  Text            =   "Product Conversion"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "StockCard"
                  Text            =   "Stock Card"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "StockReorderPoint"
                  Text            =   "Stock Reorder Point"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "NewStock"
                  Text            =   "New Stock"
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "AuditStock"
                  Text            =   "Audit Stock"
               EndProperty
               BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "AdjustStock"
                  Text            =   "Adjust Stock"
               EndProperty
               BeginProperty ButtonMenu12 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "TransferStock"
                  Text            =   "Transfer Stock"
               EndProperty
               BeginProperty ButtonMenu13 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu14 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "MovementHistory"
                  Text            =   "Movement History"
               EndProperty
               BeginProperty ButtonMenu15 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu16 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PriceManager"
                  Text            =   "Price Manager"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Purchasing"
            ImageIndex      =   2
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   8
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "PurchaseOrder"
                  Text            =   "Purchase Order"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PurchaseReturn"
                  Text            =   "Purchase Return"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "ProductBySupplier"
                  Text            =   "Supplier Products"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "Charges"
                  Text            =   "Charges"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "Shrinkages"
                  Text            =   "Shrinkages"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "NewVendor"
                  Text            =   "New Supplier"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "VendorList"
                  Text            =   "Supplier List"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sales"
            ImageIndex      =   4
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   6
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "SalesOrder"
                  Text            =   "Sales Order"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "SalesReturn"
                  Text            =   "Sales Return"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SalesAdjustment"
                  Text            =   "Sales Adjustment"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "NewCustomer"
                  Text            =   "New Customer"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "CustomerList"
                  Text            =   "Customer List"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Finance"
            ImageIndex      =   7
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   16
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "Banks"
                  Text            =   "Banks"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "Funds"
                  Text            =   "Funds"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Expenses"
                  Text            =   "Expenses"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "ExpenseList"
                  Text            =   "Expense List"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "AccountFunding"
                  Text            =   "Account Funding"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "BalanceForwarding"
                  Text            =   "Balance Forwarding"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "CheckRegistry"
                  Text            =   "Check Registry"
               EndProperty
               BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu12 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "AccountsReceivable"
                  Text            =   "Accounts Receivable"
               EndProperty
               BeginProperty ButtonMenu13 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "AccountsPayable"
                  Text            =   "Accounts Payable"
               EndProperty
               BeginProperty ButtonMenu14 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu15 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Object.Visible         =   0   'False
                  Key             =   "CustomerLedger"
                  Text            =   "Customer Ledger"
               EndProperty
               BeginProperty ButtonMenu16 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "PaymentHistory"
                  Text            =   "Payment History"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Reports"
            ImageIndex      =   9
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "AgingAccountsSummary"
                  Text            =   "Aging Accounts Summary"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
            Object.Width           =   1
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Settings"
            ImageIndex      =   8
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "GeneralSettings"
                  Text            =   "General Settings"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "SystemSettings"
                  Text            =   "System Settings"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "GEN_PatchExtendedFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Toolbar_Main_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "NewProduct"
            CornerChildForm INV_NewProductFrm
            INV_NewProductFrm.Show
            INV_NewProductFrm.ZOrder 0
        Case "TransferStock"
            CornerChildForm INV_TransferStockFrm
            INV_TransferStockFrm.Show
            INV_TransferStockFrm.ZOrder 0
         Case "NewStockSummary"
            CornerChildForm RPT_INV_NewStockSummaryFrm
            RPT_INV_NewStockSummaryFrm.Show
            RPT_INV_NewStockSummaryFrm.ZOrder 0
        Case "SystemSettings"
            If EditAccessRights(25) = False Then
                MsgBox ErrorCodes(74), vbCritical, "Limited Rights"
            Else
                BASE_SystemSettingsFrm.Show (1)
            End If
        Case "AuditStock"
            CornerChildForm INV_AuditStockFrm
            INV_AuditStockFrm.Show
            INV_AuditStockFrm.ZOrder 0
        Case "NewStock"
            CornerChildForm INV_NewStockFrm
            INV_NewStockFrm.Show
            INV_NewStockFrm.ZOrder 0
        Case "AdjustStock"
            CornerChildForm INV_AdjustStockFrm
            INV_AdjustStockFrm.Show
            INV_AdjustStockFrm.ZOrder 0
        Case "POSUserAuditTrail"
            CornerChildForm RPT_POS_UserAuditTrailFrm
            RPT_POS_UserAuditTrailFrm.Show
            RPT_POS_UserAuditTrailFrm.ZOrder 0
        Case "POSOrderSummary"
            CornerChildForm RPT_POS_OrderSummaryFrm
            RPT_POS_OrderSummaryFrm.Show
            RPT_POS_OrderSummaryFrm.ZOrder 0
        Case "SupplierPaymentHistory"
            CornerChildForm RPT_PO_PaymentHistoryFrm
            RPT_PO_PaymentHistoryFrm.Show
            RPT_PO_PaymentHistoryFrm.ZOrder 0
        Case "PriceManager"
            'CornerChildForm INV_PriceManagerFrm
            INV_PriceManagerFrm.Show
            INV_PriceManagerFrm.ZOrder 0
        Case "ProductConversion"
            'CornerChildForm INV_PriceManagerFrm
            INV_ProductConversionFrm.Show (1)
            INV_ProductConversionFrm.ZOrder 0
        Case "StockCard"
            CornerChildForm INV_StockCardFrm
            INV_StockCardFrm.Show
            INV_StockCardFrm.ZOrder 0
        Case "StockReorderPoint"
            CornerChildForm INV_StockOnReorderPointFrm
            INV_StockOnReorderPointFrm.Show
            INV_StockOnReorderPointFrm.ZOrder 0
        Case "PurchaseOrder"
            CornerChildForm PO_PurchaseOrderFrm
            PO_PurchaseOrderFrm.Show
            PO_PurchaseOrderFrm.ZOrder 0
        Case "PurchaseReturn"
            PO_PurchaseReturnFrm.Show
        Case "ProductBySupplier"
            CornerChildForm PO_ProductBySupplierFrm
            PO_ProductBySupplierFrm.Show
            PO_ProductBySupplierFrm.ZOrder 0
        Case "SalesOrder"
            CornerChildForm SO_SalesOrderFrm
            SO_SalesOrderFrm.Show
            SO_SalesOrderFrm.ZOrder 0
        Case "SalesReturn"
            CornerChildForm SO_SalesReturnFrm
            SO_SalesReturnFrm.Show
            SO_SalesReturnFrm.ZOrder 0
        Case "AccountsReceivable"
            CornerChildForm FIN_AccountsReceivable
            FIN_AccountsReceivable.Show
            FIN_AccountsReceivable.ZOrder 0
        Case "CashSalesByProduct"
            CornerChildForm RPT_POS_SalesDetailsFrm
            RPT_POS_SalesDetailsFrm.Show
            RPT_POS_SalesDetailsFrm.ZOrder 0
        Case "AccountsPayable"
            CornerChildForm FIN_AccountsPayable
            FIN_AccountsPayable.Show
            FIN_AccountsPayable.ZOrder 0
        Case "CashAdvanceRPT"
            CornerChildForm RPT_SO_CashAdvanceFrm
            RPT_SO_CashAdvanceFrm.Show
            RPT_SO_CashAdvanceFrm.ZOrder 0
        Case "PaymentHistory"
            CornerChildForm FIN_PaymentHistoryFrm
            FIN_PaymentHistoryFrm.Show
            FIN_PaymentHistoryFrm.ZOrder 0
        Case "Banks"
            CenterChildForm FIN_BanksFrm
            FIN_BanksFrm.Show
            FIN_BanksFrm.ZOrder 0
        Case "Funds"
            CenterChildForm FIN_FundsFrm
            FIN_FundsFrm.Show
            FIN_FundsFrm.ZOrder 0
        Case "AccountFunding"
            CenterChildForm FIN_AccountFundingFrm
            FIN_AccountFundingFrm.Show
            FIN_AccountFundingFrm.ZOrder 0
        Case "CustomerLedger"
            CornerChildForm RPT_SO_CustomerLedgerFrm
            RPT_SO_CustomerLedgerFrm.Show
            RPT_SO_CustomerLedgerFrm.ZOrder 0
        Case "SupplierLedger"
            CornerChildForm RPT_PO_SupplierLedgerFrm
            RPT_PO_SupplierLedgerFrm.Show
            RPT_PO_SupplierLedgerFrm.ZOrder 0
        Case "CollectionbyCustomer"
            CornerChildForm RPT_CollectionListbyCustomerFrm
            RPT_CollectionListbyCustomerFrm.Show
            RPT_CollectionListbyCustomerFrm.ZOrder 0
        Case "GeneralSalesTransaction"
            CornerChildForm RPT_GeneralSalesTransactionFrm
            RPT_GeneralSalesTransactionFrm.Show
            RPT_GeneralSalesTransactionFrm.ZOrder 0
        Case "CustomerInvoiceTransactions"
            CornerChildForm RPT_SO_InvoiceTransactionsFrm
            RPT_SO_InvoiceTransactionsFrm.Show
            RPT_SO_InvoiceTransactionsFrm.ZOrder 0
        Case "Locations"
            CenterChildForm INV_LocationModFrm
            INV_LocationModFrm.Show
            INV_LocationModFrm.ZOrder 0
        Case "ProductCategories"
            CenterChildForm INV_CategoryModFrm
            INV_CategoryModFrm.Show
            INV_CategoryModFrm.ZOrder 0
        Case "NewCustomer"
            CornerChildForm SO_CustomerFrm
            SO_CustomerFrm.Show
            SO_CustomerFrm.ZOrder 0
        Case "NewVendor"
            CornerChildForm PO_VendorFrm
            PO_VendorFrm.Show
            PO_VendorFrm.ZOrder 0
        Case "ProductList"
            CornerChildForm INV_ProductListFrm
            INV_ProductListFrm.Show
            INV_ProductListFrm.ZOrder 0
        Case "CustomerList"
            CornerChildForm SO_CustomerListFrm
            SO_CustomerListFrm.Show
            SO_CustomerListFrm.ZOrder 0
         Case "VendorList"
            CornerChildForm PO_VendorListFrm
            PO_VendorListFrm.Show
            PO_VendorListFrm.ZOrder 0
        Case "ExpenseList"
            CenterChildForm FIN_ExpenseListFrm
            FIN_ExpenseListFrm.Show
            FIN_ExpenseListFrm.ZOrder 0
        Case "Expenses"
            CenterChildForm FIN_ExpensesFrm
            FIN_ExpensesFrm.Show
            FIN_ExpensesFrm.ZOrder 0
        Case "InventorySummary"
            CornerChildForm RPT_INV_InventorySummaryFrm
            RPT_INV_InventorySummaryFrm.Show
            RPT_INV_InventorySummaryFrm.ZOrder 0
        Case "InventoryByLocation"
            CornerChildForm RPT_INV_InventoryByLocationFrm
            RPT_INV_InventoryByLocationFrm.Show
            RPT_INV_InventoryByLocationFrm.ZOrder 0
        Case "InventoryBySales"
            CornerChildForm RPT_INV_InventoryBySalesFrm
            RPT_INV_InventoryBySalesFrm.Show
            RPT_INV_InventoryBySalesFrm.ZOrder 0
        Case "PurchaseOrderSummary"
            CornerChildForm RPT_PO_PurchaseOrderSummary
            RPT_PO_PurchaseOrderSummary.Show
            RPT_PO_PurchaseOrderSummary.ZOrder 0
        Case "PurchaseOrderDetails"
            CornerChildForm RPT_PO_PurchaseOrderDetailsFrm
            RPT_PO_PurchaseOrderDetailsFrm.Show
            RPT_PO_PurchaseOrderDetailsFrm.ZOrder 0
        Case "PurchaseOrderPaymentDetails"
            CornerChildForm RPT_PO_PurchaseOrderPaymentDetailsFrm
            RPT_PO_PurchaseOrderPaymentDetailsFrm.Show
            RPT_PO_PurchaseOrderPaymentDetailsFrm.ZOrder 0
        Case "AccountsPayableReport"
            CornerChildForm RPT_PO_AccountsPayableFrm
            RPT_PO_AccountsPayableFrm.Show
            RPT_PO_AccountsPayableFrm.ZOrder 0
        Case "SalesOrderSummary"
            CornerChildForm RPT_SO_SalesOrderSummaryFrm
            RPT_SO_SalesOrderSummaryFrm.Show
            RPT_SO_SalesOrderSummaryFrm.ZOrder 0
        Case "SalesAdjustmentSummary"
            CornerChildForm RPT_SO_SalesAdjustmentSummaryFrm
            RPT_SO_SalesAdjustmentSummaryFrm.Show
            RPT_SO_SalesAdjustmentSummaryFrm.ZOrder 0
        Case "SalesbyProductDetails"
            CornerChildForm RPT_SO_SalesByProductDetailsFrm
            RPT_SO_SalesByProductDetailsFrm.Show
            RPT_SO_SalesByProductDetailsFrm.ZOrder 0
        Case "AccountsReceivableReport"
            CornerChildForm RPT_SO_AccountsReceivableFrm
            RPT_SO_AccountsReceivableFrm.Show
            RPT_SO_AccountsReceivableFrm.ZOrder 0
        Case "AgingAccountsSummary"
            CornerChildForm RPT_SO_AgingAccountsFrm
            RPT_SO_AgingAccountsFrm.Show
            RPT_SO_AgingAccountsFrm.ZOrder 0
        Case "CollectionSummary"
            CornerChildForm RPT_SO_CollectionByCustomerFrm
            RPT_SO_CollectionByCustomerFrm.Show
            RPT_SO_CollectionByCustomerFrm.ZOrder 0
        Case "AgingofAccounts"
            CornerChildForm RPT_SO_AgingAccountsFrm
            RPT_SO_AgingAccountsFrm.Show
            RPT_SO_AgingAccountsFrm.ZOrder 0
        Case "AccountsReceivableDetails"
            CornerChildForm RPT_SO_InvoiceTransactionsFrm
            RPT_SO_InvoiceTransactionsFrm.Show
            RPT_SO_InvoiceTransactionsFrm.ZOrder 0
        Case "CustomerListReport"
            CornerChildForm RPT_SO_CustomerListFrm
            RPT_SO_CustomerListFrm.Show
            RPT_SO_CustomerListFrm.ZOrder 0
        Case "StatementofAccount"
            CornerChildForm RPT_SO_CustomerStatementofAccountFrm
            RPT_SO_CustomerStatementofAccountFrm.Show
            RPT_SO_CustomerStatementofAccountFrm.ZOrder 0
        Case "SupplierStatement"
            CornerChildForm RPT_PO_SupplierStatementofAccountFrm
            RPT_PO_SupplierStatementofAccountFrm.Show
            RPT_PO_SupplierStatementofAccountFrm.ZOrder 0
'        Case "OrderPenalties"
'            CenterChildForm SO_OrderPenaltiesFrm
'            SO_OrderPenaltiesFrm.Show
'            SO_OrderPenaltiesFrm.ZOrder 0
        Case "SalesAdjustment"
            CornerChildForm SO_SalesAdjustmentFrm
            SO_SalesAdjustmentFrm.Show
            SO_SalesAdjustmentFrm.ZOrder 0
        Case "CustomerPaymentHistory"
            CornerChildForm RPT_SO_CustomerPaymentDetailsFrm
            RPT_SO_CustomerPaymentDetailsFrm.Show
            RPT_SO_CustomerPaymentDetailsFrm.ZOrder 0
        Case "CustomerSalesVolume"
            CornerChildForm RPT_SO_CustomerSalesVolumeFrm
            RPT_SO_CustomerSalesVolumeFrm.Show
            RPT_SO_CustomerSalesVolumeFrm.ZOrder 0
        Case "CheckRegistry"
            CornerChildForm RPT_FIN_CheckRegistryFrm
            RPT_FIN_CheckRegistryFrm.Show
            RPT_FIN_CheckRegistryFrm.ZOrder 0
        Case "ExpensesReport"
            CornerChildForm RPT_FIN_ExpensesFrm
            RPT_FIN_ExpensesFrm.Show
            RPT_FIN_ExpensesFrm.ZOrder 0
        Case "CashFlowReport"
            CornerChildForm RPT_CashInflow
            RPT_CashInflow.Show
            RPT_CashInflow.ZOrder 0
        Case "BalanceForwarding"
            CenterChildForm FIN_BalanceForwardingFrm
            FIN_BalanceForwardingFrm.Show
            FIN_BalanceForwardingFrm.ZOrder 0
        Case "GeneralSettings"
            If EditAccessRights(24) = False Then
                MsgBox ErrorCodes(74), vbCritical, "Limited Rights"
            Else
                BASE_GeneralSettingsFrm.Show '(1)
            End If
        Case "SystemSettings"
            If EditAccessRights(25) = False Then
                MsgBox ErrorCodes(74), vbCritical, "Limited Rights"
            Else
                BASE_SystemSettingsFrm.Show (1)
            End If
        Case "PointofSale"
            POS_CashierFrm.Show
        Case "CheckRegistry"
            CornerChildForm FIN_CheckRegistryFrm
            FIN_CheckRegistryFrm.Show
            FIN_CheckRegistryFrm.ZOrder 0
        Case "CashAdvance"
            CornerChildForm SO_CashAdvance
            SO_CashAdvance.Show
            SO_CashAdvance.ZOrder 0
        Case "PurchaseOrderbyProduct"
            CornerChildForm RPT_PO_PurchaseOrderByProductFrm
            RPT_PO_PurchaseOrderByProductFrm.Show
            RPT_PO_PurchaseOrderByProductFrm.ZOrder 0
         Case "POSSalesByCashier"
            CornerChildForm RPT_POS_SalesByCashierFrm
            RPT_POS_SalesByCashierFrm.Show
            RPT_POS_SalesByCashierFrm.ZOrder 0
        Case "POSSalesbyInvoice"
            CornerChildForm RPT_POS_SalesByInvoiceFrm
            RPT_POS_SalesByInvoiceFrm.Show
            RPT_POS_SalesByInvoiceFrm.ZOrder 0
         Case "POSSalesByCustomer"
            CornerChildForm RPT_POS_SalesByCustomerFrm
            RPT_POS_SalesByCustomerFrm.Show
            RPT_POS_SalesByCustomerFrm.ZOrder 0
         Case "InventoryByCategory"
            CornerChildForm RPT_INV_InventoryByCategoryFrm
            RPT_INV_InventoryByCategoryFrm.Show
            RPT_INV_InventoryByCategoryFrm.ZOrder 0
        Case "InventoryBySupplier"
            CornerChildForm RPT_INV_InventoryBySupplierFrm
            RPT_INV_InventoryBySupplierFrm.Show
            RPT_INV_InventoryBySupplierFrm.ZOrder 0
         Case "ProductPricing"
            CornerChildForm RPT_INV_InventoryProductPricingFrm
            RPT_INV_InventoryProductPricingFrm.Show
            RPT_INV_InventoryProductPricingFrm.ZOrder 0
         Case "ProductExpiry"
            CornerChildForm RPT_INV_ProductExpiry
            RPT_INV_ProductExpiry.Show
            RPT_INV_ProductExpiry.ZOrder 0
        Case "ProductListRpt"
'            CornerChildForm RPT_INV_ProductExpiry
'            RPT_INV_ProductExpiry.Show
'            RPT_INV_ProductExpiry.ZOrder 0
            MsgBox "Working on it! :)"
         Case "POSSalesSummary"
            CornerChildForm RPT_POS_SalesSummaryFrm
            RPT_POS_SalesSummaryFrm.Show
            RPT_POS_SalesSummaryFrm.ZOrder 0
         Case "POSSalesReturn"
            POS_SalesReturnFrm.Show
            POS_SalesReturnFrm.ZOrder 0
        Case "POSSalesReturnReport"
            CornerChildForm RPT_POS_SalesReturn
            RPT_POS_SalesReturn.Show
            RPT_POS_SalesReturn.ZOrder 0
          
    End Select
End Sub
