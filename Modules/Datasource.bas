Attribute VB_Name = "Datasource"
Option Explicit
Public con As ADODB.Connection
Public rec As ADODB.Recordset
Public cmd As ADODB.Command

'***************** REGION GLOBAL DATA ****************
Public Function Global_Data(ByVal data As String) As ADODB.Recordset
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.CommandType = adCmdStoredProc
    cmd.ActiveConnection = con
    
    Select Case data
        Case "Category"
            cmd.CommandText = "BASE_Category_Get"
            Set rec = cmd.Execute
        Case "Type"
            cmd.CommandText = "BASE_Type_Get"
            Set rec = cmd.Execute
        Case "WarehouseManager"
            cmd.CommandText = "BASE_WarehouseManager_Get"
            Set rec = cmd.Execute
        Case "CardType"
            cmd.CommandText = "BASE_CardType_Get"
            Set rec = cmd.Execute
        Case "Uom"
            cmd.CommandText = "BASE_Uom_Get"
            Set rec = cmd.Execute
        Case "Room"
            cmd.CommandText = "BASE_Room_Get"
            Set rec = cmd.Execute
        Case "Slot"
            cmd.CommandText = "BASE_Slot_Get"
            Set rec = cmd.Execute
        Case "UomConversion"
            cmd.CommandText = "INV_UomConversion_Get"
            Set rec = cmd.Execute
        Case "SalesRep"
            cmd.CommandText = "BASE_WarehouseManager_Get"
            Set rec = cmd.Execute
        Case "WarehouseManager"
            cmd.CommandText = "BASE_SalesRep_Get"
            Set rec = cmd.Execute
        Case "PricingScheme"
            cmd.CommandText = "BASE_PricingScheme_Get"
            Set rec = cmd.Execute
        Case "Location"
            cmd.CommandText = "BASE_Location_Get"
            Set rec = cmd.Execute
        Case "Product"
            cmd.CommandText = "BASE_Product_Get"
            Set rec = cmd.Execute
        Case "Vendor"
            cmd.CommandText = "BASE_Vendor_Get"
            Set rec = cmd.Execute
        Case "Customer"
            cmd.CommandText = "BASE_Customer_Get"
            Set rec = cmd.Execute
        Case "Vendor"
            cmd.CommandText = "BASE_Vendor_Get"
            Set rec = cmd.Execute
        Case "Terms"
            cmd.CommandText = "BASE_Terms_Get"
            Set rec = cmd.Execute
        Case "Tax"
            cmd.CommandText = "BASE_Tax_Get"
            Set rec = cmd.Execute
        Case "Discount"
            cmd.CommandText = "BASE_Discount_Get"
            Set rec = cmd.Execute
        Case "Charges"
            cmd.CommandText = "BASE_Charges_Get"
            Set rec = cmd.Execute
        Case "Shrinkage"
            cmd.CommandText = "BASE_Shrinkage_Get"
            Set rec = cmd.Execute
        Case "ShrinkageGrade"
            cmd.CommandText = "BASE_ShrinkageGrade_Get"
            Set rec = cmd.Execute
        Case "Status"
            cmd.CommandText = "GLOBAL_DocStatus_Get"
            Set rec = cmd.Execute
        Case "PaymentType"
            cmd.CommandText = "BASE_PaymentType_Get"
            Set rec = cmd.Execute
        Case "Bank"
            cmd.CommandText = "BASE_Bank_Get"
            Set rec = cmd.Execute
        Case "Fund"
            cmd.CommandText = "BASE_Fund_Get"
            Set rec = cmd.Execute
        Case "Expense"
            cmd.CommandText = "BASE_Expense_Get"
            Set rec = cmd.Execute
        Case "Company"
            cmd.CommandText = "BASE_Company_Get"
            Set rec = cmd.Execute
        Case "User"
            cmd.CommandText = "BASE_User_Get"
            Set rec = cmd.Execute
        Case "UserRoles"
            cmd.CommandText = "BASE_UserRoles_Get"
            Set rec = cmd.Execute
        Case "Documents"
            cmd.CommandText = "GLOBAL_DocNoFormat_Get"
            Set rec = cmd.Execute
        Case "POSDisplay"
            cmd.CommandText = "POS_Display_Get"
            Set rec = cmd.Execute
        Case "POSValidation"
            cmd.CommandText = "POS_UserValidation_Get"
            Set rec = cmd.Execute
        Case "UserRights"
            cmd.CommandText = "BASE_UserRights_Get"
            Set rec = cmd.Execute
        Case "CheckStatus"
            cmd.CommandText = "BASE_CheckStatus_Get"
            Set rec = cmd.Execute
        Case "RecentReceipts"
            cmd.CommandText = "POS_RecentReceipts"
            cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
            Set rec = cmd.Execute
        Case "RecentCreditReceipts"
            cmd.CommandText = "POS_RecentCreditReceipts"
            Set rec = cmd.Execute
        Case "ReturnSlips"
            cmd.CommandText = "POS_SalesReturnSlip_Get"
            Set rec = cmd.Execute
        Case "Workstation"
            cmd.CommandText = "SYS_Workstation_Get"
            Set rec = cmd.Execute
        Case "InventorySettings"
            cmd.CommandText = "INV_Settings_Get"
            Set rec = cmd.Execute
        Case "City"
            cmd.CommandText = "BASE_City_Get"
            Set rec = cmd.Execute
    End Select
    Set Global_Data = rec
    Set con = Nothing
End Function

'***************** END REGION GLOBAL DATA ************
