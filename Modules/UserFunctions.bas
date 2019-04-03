Attribute VB_Name = "UserFunctions"
Option Explicit
Global UserId, WorkstationId, gUserRoleId As Integer
Global AllowNegativeInventory As Boolean
Global ReorderPointCheckScheduled As String
Global ReorderPointCheckFrequency As Double
Global LastReorderPointDateCheck As Date
Global AllowAccess As Boolean
Global EveryLogIn As Boolean
Global EveryLogOut As Boolean
Global ExpiryDays As Double
Public CardInfo As New CardPaymentInfo
Public CheckInfo As New CheckPaymentInfo
Public LoyaltyInfo As New LoyaltyPointsInfo
Public OtherInfo As New OtherPaymentInfo
Global CurrentUser As String
Global CurrentUsers As String
Public isModify As Boolean
Public ProductSet As ADODB.Recordset
Public OrderSet As ADODB.Recordset
Public AccessRights(1 To 99, 1 To 99) As Boolean
Global PharmacyMode As String
Global OrderSlipMode As String
Global DualPharmacyMode As String
Global VoidUserId As Long
Global NotificationTimer As Integer
Global POSLogo As String
Global StatementTemplateId As Integer
Global CSVRecordset As ADODB.Recordset
Global UniversalCtr As Long
Global POS_Printer, BackOffice_Printer As String

Private Declare Function GetVolumeInformation _
    Lib "kernel32" Alias "GetVolumeInformationA" _
    (ByVal lpRootPathName As String, _
    ByVal pVolumeNameBuffer As String, _
    ByVal nVolumeNameSize As Long, _
    lpVolumeSerialNumber As Long, _
    lpMaximumComponentLength As Long, _
    lpFileSystemFlags As Long, _
    ByVal lpFileSystemNameBuffer As String, _
    ByVal nFileSystemNameSize As Long) As Long
Public Function Hostname() As String
    'Get Hostname from Text
    Open App.path & "\Resources\Hostname.txt" For Input As #1
    Input #1, Hostname
    Input #1, Hostname
    Close #1
End Function
Public Function DatabaseName() As String
    'Get Hostname from Text
    Open App.path & "\Resources\Hostname.txt" For Input As #1
    Input #1, DatabaseName '[Hostname]
    Input #1, DatabaseName '[Hostname value]
    Input #1, DatabaseName '[Databasename]
    Input #1, DatabaseName '[Database]
    Close #1
End Function

Public Function ConnString() As String
    ConnString = "Provider=SQLNCLI.1;Data Source = " & Hostname & "\PEAKSQL;User Id=sa; " & _
                 "Password=PeakPOS2015;Initial Catalog=" & DatabaseName & ";"
End Function
Public Sub ResetRptDB(ByRef crxReport As CRAXDRT.Report)
    Dim DBProviderName As String ' i.e SQLOLEDB.1;
    Dim DBDataSource As String ' i.e brandon-pc\sqlexpress
    Dim DBName As String
    Dim DBUsername As String
    Dim DBPwd As String
    Dim ConnectionString As String
    Dim crxTable As DatabaseTable
    'Dim objDataAccess As DataAccess.clsDataAccess
    Dim i As Integer
    Dim crxSection As CRAXDRT.Section
    Dim ReportObject
    Dim crxSubReportObj
    Dim crxsubreport
    Dim crxdatatable
    
    DBProviderName = "SQLNCLI.1"
    DBDataSource = Hostname & "\PEAKSQL"
    DBName = DatabaseName
    DBUsername = "sa"
    DBPwd = "PeakPOS2015"
    
    For Each crxTable In crxReport.Database.Tables
        Call crxTable.SetLogOnInfo(DBDataSource, DBName, DBUsername, DBPwd)
        Call crxTable.SetTableLocation(crxTable.Location, "", ConnString)
    Next
    
    For Each crxSection In crxReport.Sections
        For Each ReportObject In crxSection.ReportObjects
            If ReportObject.Kind = crSubreportObject Then
            
                Set crxSubReportObj = ReportObject
                Set crxsubreport = crxSubReportObj.OpenSubreport
                
                For Each crxdatatable In crxsubreport.Database.Tables
                    Call crxdatatable.SetLogOnInfo(DBDataSource, DBName, DBUsername, DBPwd)
                    Call crxdatatable.SetTableLocation(crxdatatable.Location, "", ConnectionString)
                Next
                
            End If
        Next
    Next
End Sub
Public Sub ResetRptDB_try(ByRef crxReport As CRAXDRT.Report)
    Dim DBProviderName As String ' i.e SQLOLEDB.1;
    Dim DBDataSource As String ' i.e brandon-pc\sqlexpress
    Dim DBName As String
    Dim DBUsername As String
    Dim DBPwd As String
    Dim ConnectionString As String
    
    Dim crxApp As CRAXDRT.Application
    Dim CrxRep As CRAXDRT.Report
    Dim crxDatabase As CRAXDRT.Database
    Dim crxDatabaseTables As CRAXDRT.DatabaseTables
    Dim crxDatabaseTable As CRAXDRT.DatabaseTable
    Dim crxSection
    Dim ReportObject
    Dim crxSubReportObj
    Dim crxsubreport

    
    
    DBProviderName = "SQLNCLI.1"
    DBDataSource = Hostname & "\PEAKSQL"
    DBName = DatabaseName
    DBUsername = "sa"
    DBPwd = "PeakPOS2015"
    
    Set CrxRep = crxReport
    Set crxDatabase = CrxRep.Database
    Set crxDatabaseTables = crxDatabase.Tables
    
    For Each crxDatabaseTable In crxDatabaseTables
        crxDatabaseTable.SetLogOnInfo DBDataSource, DBName, DBUsername, DBPwd
    Next crxDatabaseTable
    
    For Each crxSection In crxReport.Sections
        For Each ReportObject In crxSection.ReportObjects
            If ReportObject.Kind = crSubreportObject Then
            
                Set crxSubReportObj = ReportObject
                Set crxsubreport = crxSubReportObj.OpenSubreport
                
                For Each crxDatabaseTable In crxsubreport.Database.Tables
                    Call crxDatabaseTable.SetLogOnInfo(DBDataSource, DBName, DBUsername, DBPwd)
                    Call crxDatabaseTable.SetTableLocation(crxDatabaseTable.Location, "", ConnectionString)
                Next
                
            End If
        Next
    Next
End Sub

Public Sub selectText(ByVal Text As Control)
    Text.SelStart = 0
    Text.SelLength = Len(Text.Text)
End Sub
Public Sub CenterChildForm(ByVal Form As Form)
    Form.Left = (BASE_ContainerFrm.ScaleWidth - Form.width) / 2
    Form.Top = (BASE_ContainerFrm.ScaleHeight - Form.Height) / 2
End Sub
Public Sub CornerChildForm(ByVal Form As Form)
    On Error Resume Next
    Form.Left = 0
    Form.Top = 0
End Sub
Public Sub ShowNotification()
    On Error Resume Next
    BASE_NotificationFrm.Left = (BASE_ContainerFrm.width - BASE_NotificationFrm.width) - 600
    BASE_NotificationFrm.Top = 0
    BASE_NotificationFrm.Show
    BASE_NotificationFrm.ZOrder 0
End Sub
Public Sub StatusBarWidth(ByVal Form As Form, ByVal Statusbar As Statusbar)
    On Error Resume Next
    Dim width As Double
    width = Form.ScaleWidth
    Statusbar.Panels(1).width = width * 0.3
    Statusbar.Panels(2).width = width * 0.2
    Statusbar.Panels(3).width = width * 0.2
    Statusbar.Panels(4).width = width * 0.3
End Sub
'Public Sub DistinctList(lv As MSComctlLib.ListView)
'    Dim i As Long
'    Dim j As Long
'    With lv
'        For i = 1 To .ListItems.Count
'            For j = .ListItems.Count To (i + 1) Step -1
'                If .ListItems(j) = .ListItems(i) Then
'                    .ListItems.Remove j
'                End If
'            Next
'        Next
'    End With
'End Sub

Public Function ErrorCodes(ByVal Code As Integer) As String
    Dim Errors(100) As String
    Errors(0) = "Save failed."
    Errors(1) = "Product code is required."
    Errors(2) = "Product name is required."
    Errors(3) = "Product name is already in use."
    Errors(4) = "Probably with an inactive one."
    Errors(5) = "Category is required."
    Errors(6) = "Invalid category."
    Errors(7) = "Invalid Unit Price."
    Errors(8) = "Price must be numeric."
    Errors(9) = "Invalid Unit Cost."
    Errors(10) = "Unit of Measure is required."
    Errors(11) = "Code is already in use."
    Errors(12) = "Numeric data is required."
    Errors(13) = "Customer is required."
    Errors(14) = "Terms is required."
    Errors(15) = "Order number is already in use."
    Errors(16) = "Bank account is required."
    Errors(17) = "No valid payment found."
    Errors(18) = "Name is required."
    Errors(19) = "Name is already in use."
    Errors(20) = "Fund account is required."
    Errors(21) = "Account number is required."
    Errors(22) = "Bank is required."
    Errors(23) = "Account number is already in use."
    Errors(24) = "Amount is required."
    Errors(25) = "Amount is invalid."
    Errors(26) = "Expense is required."
    Errors(27) = "There is already a forwarded balance in this date."
    Errors(28) = "Password did not match."
    Errors(29) = "Invalid username and/or password."
    Errors(30) = "User Name is required."
    Errors(31) = "User Name is already in use."
    Errors(32) = "Check # is required."
    Errors(33) = "Insufficient quantity."
    Errors(34) = "Payment is insufficient."
    Errors(35) = "Delete failed. No item selected."
    Errors(36) = "No items selected."
    Errors(37) = "Please select accounts to pay."
    Errors(38) = "Login failed."
    Errors(39) = "Username and/or password is invalid."
    Errors(40) = "Code is required."
    Errors(41) = "Mark-up is invalid."
    Errors(42) = "Field required."
    Errors(43) = "Invalid data."
    Errors(44) = "User number must be numeric."
    Errors(45) = "Pin must be numeric."
    Errors(46) = "User cannot be deactivated."
    Errors(47) = "User number already in use."
    Errors(48) = "Name already exists."
    Errors(49) = "Password is required."
    Errors(50) = "Tax is required."
    Errors(51) = "Card number is required."
    Errors(52) = "Reference is required."
    Errors(53) = "Card number does not exist."
    Errors(54) = "Card already in use."
    Errors(55) = "Login error. Machine is not registerd in the system."
    Errors(56) = "Invalid user number."
    Errors(57) = "Invalid pin."
    Errors(58) = "Login error. Machine is not activated in the system."
    Errors(59) = "Item does not exists in the purchase order list."
    Errors(60) = "Cannot receive inventory when order is already complete, cancelled or invoiced."
    Errors(61) = "Cannot pick inventory when order is already complete."
    Errors(62) = "Cannot pick inventory when order is already invoiced."
    Errors(63) = "Cannot pick inventory when order is already paid or cancelled."
    Errors(64) = "Order is cancelled. No changes made."
    Errors(65) = "User pin not set."
    Errors(66) = "User not allowed."
    Errors(67) = "No more records to display."
    Errors(68) = "Invalid O.R. number."
    Errors(69) = "Child is required."
    Errors(70) = "Attendant is required."
    Errors(71) = "Hours must be greater than 0."
    Errors(72) = "Reorder point must be numeric."
    Errors(73) = "Reorder quantity must be numeric."
    Errors(74) = "This account is restricted in editing data in this module."
    Errors(75) = "This account is restricted in viewing details on this module/record."
    ErrorCodes = Errors(Code)
End Function

Public Function MessageCodes(ByVal Code As Integer) As String
    Dim Message(100) As String
    Message(0) = "saved."
    Message(1) = "Record/s"
    Message(2) = "deleted."
    Message(3) = "Payments"
    Message(4) = "deactivated."
    Message(5) = "activated."
    Message(6) = "New"
    Message(7) = "Record"
    MessageCodes = Message(Code)
End Function

Public Sub ClearClassData(ByVal info As Integer)
    Select Case info
        Case 0
            With CardInfo
                .Amount = 0
                .BankId = 0
                .CardNumber = ""
                .CardTypeId = 0
                .NameOnCard = ""
                .Reference = ""
            End With
        Case 1
            With CheckInfo
                .Amount = 0
                .BankId = 0
                .CheckDate = Format(Now, "MM/DD/YY")
                .CheckNumber = ""
            End With
        Case 2
            With LoyaltyInfo
                .CardNumber = ""
                .UsePoints = "0.00"
            End With
        Case 3
            With OtherInfo
                .ReferenceNumber = ""
                .Remarks = ""
                .Amount = "0.00"
            End With
    End Select
End Sub
Public Sub SavePOSAuditTrail(ByVal UserId As Integer, ByVal WorkstationId As Integer, _
                ByVal POS_SalesId As String, ByVal Activity As String, Optional Module As String = "POS")
    Dim newcon As ADODB.Connection
    Set newcon = New ADODB.Connection
    Dim newcmd As New ADODB.Command
    
    newcon.ConnectionString = ConnString
    newcon.Open
    newcmd.CommandType = adCmdStoredProc
    newcmd.ActiveConnection = newcon
    newcmd.CommandText = "POS_UserAudit_Insert"
    newcmd.Parameters.Append newcmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
    newcmd.Parameters.Append newcmd.CreateParameter("@WorkstationId", adInteger, adParamInput, , WorkstationId)
    newcmd.Parameters.Append newcmd.CreateParameter("@POS_SalesId", adInteger, adParamInput, , Val(POS_SalesId))
    newcmd.Parameters.Append newcmd.CreateParameter("@Activity", adVarChar, adParamInput, 4000, Left(Activity, 250))
    newcmd.Parameters.Append newcmd.CreateParameter("@Module", adVarChar, adParamInput, 250, Module)
    newcmd.Execute
    newcon.Close
    
    
End Sub

Public Function ProductBarcode(ByVal Barcode As String) As ADODB.Recordset
    On Error GoTo ErrMessage
    Set con = New ADODB.Connection
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_Product_Search_Barcode"
    cmd.Parameters.Append cmd.CreateParameter("@Barcode", adVarChar, adParamInput, 250, Barcode)
    
    Set rec = cmd.Execute
    'con.Close
    
    Set ProductBarcode = rec
    Set con = Nothing
    Exit Function
ErrMessage:
    MsgBox "An error occured while processing your request. " & Err.Description & " Please try again.", vbCritical
    
End Function

Public Function ProductName(ByVal Name As String) As ADODB.Recordset
    On Error GoTo ErrMessage
    Set con = New ADODB.Connection
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "POS_ItemSearch_Name"
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, Name)
    
    Set rec = cmd.Execute
    'con.Close
    
    Set ProductName = rec
    Set con = Nothing
    Exit Function
ErrMessage:
    MsgBox "An error occured while processing your request. " & Err.Description & " Please try again.", vbCritical
    
End Function

Public Sub GetInventorySettings()
    'Get Settings
    Dim con As New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "INV_Settings_Get"
    Set rec = cmd.Execute
    
    If Not rec.EOF Then
        AllowNegativeInventory = rec!AllowNegativeInventory
        ReorderPointCheckScheduled = rec!ReorderPointCheckScheduled
        ReorderPointCheckFrequency = rec!ReorderPointCheckFrequency
        LastReorderPointDateCheck = rec!LastReorderPointDateCheck
        EveryLogIn = rec!EveryLogIn
        EveryLogOut = rec!EveryLogOut
        ExpiryDays = rec!ExpiryDays
    End If
    
    con.Close
End Sub

Public Function GetNotifications(ByVal data As String) As String
    Dim con As New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "SYS_Notification_Get"
    cmd.Parameters.Append cmd.CreateParameter("@Type", adVarChar, adParamInput, 50, data)
    Set rec = cmd.Execute
    If Not rec.EOF Then
        GetNotifications = rec!Total
    End If
    con.Close
End Function

Public Function checkAvailableQuantity(ByVal ProductId As String, Optional LocationId As String = "0", Optional System As Boolean = False) As Double
    Dim chk_con As New ADODB.Connection
    Dim chkrec As New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    chk_con.ConnectionString = ConnString
    chk_con.Open
    cmd.ActiveConnection = chk_con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "INV_CheckAvailableQuantity"
    cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , Val(ProductId))
    cmd.Parameters.Append cmd.CreateParameter("@LocationId", adInteger, adParamInput, , Val(LocationId))
    cmd.Parameters.Append cmd.CreateParameter("@System", adBoolean, adParamInput, , System)
    Set chkrec = cmd.Execute
    If Not chkrec.EOF Then
        checkAvailableQuantity = chkrec!AvailableQuantity
    End If
    chk_con.Close
End Function

Public Function ReserveProduct(ByVal ReserveId As String, ByVal ProductId As String, _
    ByVal quantity As Double, ByVal UserId As Integer, ByVal WorkstationId As Integer, ByVal isPOS As Boolean, _
    ByVal ModId As Integer, Optional SalesOrderId As String = "0", Optional ByVal PurchaseReturnId As String = "0", Optional POS_SalesId As String = "0") As String
    
    Dim res_con As New ADODB.Connection
    
    Set cmd = New ADODB.Command
    
    res_con.ConnectionString = ConnString
    res_con.Open
    res_con.BeginTrans
    cmd.ActiveConnection = res_con
    cmd.CommandType = adCmdStoredProc
    
    cmd.Parameters.Append cmd.CreateParameter("@ReserveId", adInteger, adParamInputOutput, , Val(ReserveId))
    cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , Val(ProductId))
    cmd.Parameters.Append cmd.CreateParameter("@Quantity", adDecimal, adParamInput, , quantity)
                          cmd.Parameters("@Quantity").NumericScale = 2
                          cmd.Parameters("@Quantity").Precision = 18
    cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
    cmd.Parameters.Append cmd.CreateParameter("@WorkStationId", adInteger, adParamInput, , WorkstationId)
    cmd.Parameters.Append cmd.CreateParameter("@isPOS", adBoolean, adParamInput, , isPOS)
    cmd.Parameters.Append cmd.CreateParameter("@ModId", adInteger, adParamInput, , ModId)
    cmd.Parameters.Append cmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , Val(SalesOrderId))
    cmd.Parameters.Append cmd.CreateParameter("@POS_SalesId", adInteger, adParamInput, , Val(POS_SalesId))
    cmd.Parameters.Append cmd.CreateParameter("@PurchaseReturnId", adInteger, adParamInput, , Val(PurchaseReturnId))
    If Val(ReserveId) = 0 Then
        cmd.CommandText = "INV_ProductReserve_Insert"
        cmd.Execute
        ReserveProduct = cmd.Parameters("@ReserveId")
    Else
        cmd.CommandText = "INV_ProductReserve_Update"
        cmd.Execute
        ReserveProduct = cmd.Parameters("@ReserveId")
    End If
    res_con.CommitTrans
    res_con.Close
    
End Function
    
Public Sub DeleteReserves_User(ByVal UserId As Integer, ByVal isPOS As Boolean, ByVal isSalesOrder As Boolean, ByVal isPurchaseReturn As Boolean)
    Dim con As New ADODB.Connection
    con.ConnectionString = ConnString
    con.Open
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "INV_ProductReserve_DeleteByUser"
    cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
    cmd.Parameters.Append cmd.CreateParameter("@isPOS", adBoolean, adParamInput, , isPOS)
    cmd.Parameters.Append cmd.CreateParameter("@isSalesOrder", adBoolean, adParamInput, , isSalesOrder)
    cmd.Parameters.Append cmd.CreateParameter("@isPurchaseReturn", adBoolean, adParamInput, , isPurchaseReturn)
    cmd.Execute
    con.Close
End Sub
Public Sub DeleteReserves(ByVal WorkstationId As Integer, ByVal ModId As Integer)
    Dim con As New ADODB.Connection
    con.ConnectionString = ConnString
    con.Open
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "INV_ProductReserve_DeleteByWorkStation"
    cmd.Parameters.Append cmd.CreateParameter("@WorkStationId", adInteger, adParamInput, , WorkstationId)
    cmd.Parameters.Append cmd.CreateParameter("@ModId", adInteger, adParamInput, , ModId)
    cmd.Execute
    con.Close
End Sub

Public Sub DeleteReserveLine(ByVal ReserveId As String)
    Dim con As New ADODB.Connection
    Set cmd = New ADODB.Command
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "INV_ProductReserves_Delete"
    cmd.Parameters.Append cmd.CreateParameter("@ReserveId", adInteger, adParamInput, , Val(ReserveId))
    cmd.Execute
    con.Close
End Sub
Public Sub UpdateCustomerOrderDues()
    Dim con As New ADODB.Connection
    Set cmd = New ADODB.Command
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_CustomerDuesStatus_Update"
    cmd.Execute
    con.Close
End Sub
Public Sub UpdateVendorOrderDues()
    Dim con As New ADODB.Connection
    Set cmd = New ADODB.Command
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_VendorDuesStatus_Update"
    cmd.Execute
    con.Close
End Sub
Public Sub UpdateReserveQuantity(ByVal ReserveId As String, ByVal quantity As Double, ByVal ProductId As String, _
        ByVal SalesOrderId As String)
        
    Dim newcon As New ADODB.Connection
    Dim newcmd As New ADODB.Command
    newcon.ConnectionString = ConnString
    
    newcon.Open
    newcmd.ActiveConnection = newcon
    newcmd.CommandType = adCmdStoredProc
    newcmd.CommandText = "INV_ProductReserve_QuantityUpdate"
    newcmd.Parameters.Append newcmd.CreateParameter("@ReserveId", adInteger, adParamInput, , Val(ReserveId))
    newcmd.Parameters.Append newcmd.CreateParameter("@ProductId", adInteger, adParamInput, , Val(ProductId))
    newcmd.Parameters.Append newcmd.CreateParameter("@SalesOrderId", adInteger, adParamInput, , Val(SalesOrderId))
    newcmd.Parameters.Append newcmd.CreateParameter("@Quantity", adDecimal, adParamInput, , quantity)
                          newcmd.Parameters("@Quantity").NumericScale = 2
                          newcmd.Parameters("@Quantity").Precision = 18
    newcmd.Execute
    newcon.Close
End Sub

Public Function NVAL(ByVal expression As String) As Double
    NVAL = Val(Replace(expression, ",", ""))
End Function

Public Function GetRegistration() As Boolean
    Dim serial As String
    serial = GetSetting("PeakPOS", "Data", "Default")
    If "123456" <> serial Then
        GetRegistration = False
    Else
        GetRegistration = True
    End If
End Function
Public Function GetSerialNumber( _
    ByVal sDrive As String) As Long

    If Len(sDrive) Then
        If InStr(sDrive, "\\") = 1 Then
            ' Make sure we end in backslash for UNC
            If Right$(sDrive, 1) <> "\" Then
                sDrive = sDrive & "\"
            End If
        Else
            ' If not UNC, take first letter as drive
            sDrive = Left$(sDrive, 1) & ":\"
        End If
    Else
        ' Else just use current drive
        sDrive = vbNullString
    End If

    ' Grab S/N -- Most params can be NULL
    Call GetVolumeInformation( _
        sDrive, vbNullString, 0, GetSerialNumber, _
        ByVal 0&, ByVal 0&, vbNullString, 0)
End Function

Public Sub SYS_ErrorLog(ByVal UserId As Integer, ByVal WorkstationId As Integer, _
                ByVal Message As String)
    On Error Resume Next
    Dim newcon As ADODB.Connection
    Set newcon = New ADODB.Connection
    Set cmd = New ADODB.Command
    
    newcon.ConnectionString = ConnString
    newcon.Open
    cmd.CommandType = adCmdStoredProc
    cmd.ActiveConnection = newcon
    cmd.CommandText = "SYS_ErrorLog_Insert"
    cmd.Parameters.Append cmd.CreateParameter("@Message", adVarChar, adParamInput, 250, Message)
    cmd.Parameters.Append cmd.CreateParameter("@WorkstationId", adInteger, adParamInput, , WorkstationId)
    cmd.Parameters.Append cmd.CreateParameter("@UserId", adInteger, adParamInput, , UserId)
    cmd.Execute
    newcon.Close
End Sub
Public Function GetAccessRightsByModule(ByVal UserRoleId As Integer, ByVal ModuleId As Integer) As Boolean
    Dim ModuleCtr, RightsCtr As Integer
    
    Dim Item As MSComctlLib.ListItem
    Dim rrCon As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim rrRec As New ADODB.Recordset
    
    rrCon.ConnectionString = ConnString
    rrCon.Open
    
    cmd.ActiveConnection = rrCon
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_UserRoleRights_Insert"
    cmd.Parameters.Append cmd.CreateParameter("@UserRoleId", adInteger, adParamInput, , UserRoleId)
    cmd.Execute
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = rrCon
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_UserRoleRights_GetByModule"
    cmd.Parameters.Append cmd.CreateParameter("@UserRoleId", adInteger, adParamInput, , UserRoleId)
    cmd.Parameters.Append cmd.CreateParameter("@ModuleId", adInteger, adParamInput, , ModuleId)
    Set rrRec = cmd.Execute
    If Not rrRec.EOF Then
        GetAccessRightsByModule = rrRec!allowedit
    End If
    rrCon.Close
End Function


Public Sub GetAccessRights(ByVal UserRoleId As Integer)
    Dim ModuleCtr, RightsCtr As Integer
    
    Dim Item As MSComctlLib.ListItem
    Dim rrCon As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Dim rrRec As New ADODB.Recordset
    
    rrCon.ConnectionString = ConnString
    rrCon.Open
    
    cmd.ActiveConnection = rrCon
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_UserRoleRights_Insert"
    cmd.Parameters.Append cmd.CreateParameter("@UserRoleId", adInteger, adParamInput, , UserRoleId)
    cmd.Execute
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = rrCon
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_UserRoleRights_Get"
    cmd.Parameters.Append cmd.CreateParameter("@UserRoleId", adInteger, adParamInput, , UserRoleId)
    Set rrRec = cmd.Execute
    If Not rrRec.EOF Then
        ModuleCtr = 1
        Do Until rrRec.EOF
            RightsCtr = 1
            AccessRights(ModuleCtr, RightsCtr) = rrRec!allowview
            RightsCtr = RightsCtr + 1
            AccessRights(ModuleCtr, RightsCtr) = rrRec!allowedit
            ModuleCtr = ModuleCtr + 1
            rrRec.MoveNext
        Loop
    End If
    rrCon.Close
End Sub
Public Function ViewAccessRights(ByVal ModuleId As Integer) As Boolean
    Dim value As Boolean
    value = AccessRights(ModuleId, 1)
    Select Case ModuleId
        Case 1: 'New Product
                BASE_ContainerFrm.Toolbar_Main.Buttons(3).ButtonMenus(1).Visible = value
                BASE_HomepageFrm.imgNewProduct.Visible = value
                BASE_HomepageFrm.lblNewProduct.Visible = value
        Case 2: 'Product Cost
                INV_NewProductFrm.lblCostingInfo_Cost.Visible = value
                INV_NewProductFrm.txtCostingInfo_AverageCost.Visible = value
        Case 3: 'ProductList
                BASE_ContainerFrm.Toolbar_Main.Buttons(3).ButtonMenus(2).Visible = value
                BASE_HomepageFrm.imgProductList.Visible = value
                BASE_HomepageFrm.lblProductList.Visible = value
        Case 4: 'Categories
                BASE_ContainerFrm.Toolbar_Main.Buttons(3).ButtonMenus(3).Visible = value
                BASE_HomepageFrm.imgCategories.Visible = value
                BASE_HomepageFrm.lblCategories.Visible = value
        Case 5: 'Stockard
                BASE_ContainerFrm.Toolbar_Main.Buttons(3).ButtonMenus(6).Visible = value
'        Case 6: 'AdjustStock
'                BASE_ContainerFrm.Toolbar_Main.Buttons(3).ButtonMenus(9).Visible = value
        Case 7: 'Transfer Stock
                BASE_ContainerFrm.Toolbar_Main.Buttons(3).ButtonMenus(12).Visible = value
        Case 8: 'Price Manager
                BASE_ContainerFrm.Toolbar_Main.Buttons(3).ButtonMenus(16).Visible = value
        Case 9: 'Purchase Order
                BASE_ContainerFrm.Toolbar_Main.Buttons(4).ButtonMenus(1).Visible = value
                BASE_HomepageFrm.imgPurchaseOrder.Visible = value
                BASE_HomepageFrm.lblPurchaseOrder.Visible = value
        Case 10: 'Purchase Return
                BASE_ContainerFrm.Toolbar_Main.Buttons(4).ButtonMenus(2).Visible = value
        Case 11: 'New Supplier
                BASE_ContainerFrm.Toolbar_Main.Buttons(4).ButtonMenus(7).Visible = value
        Case 12: 'Supplier List
                BASE_ContainerFrm.Toolbar_Main.Buttons(4).ButtonMenus(8).Visible = value
        Case 13: 'Sales Order
                BASE_ContainerFrm.Toolbar_Main.Buttons(5).ButtonMenus(1).Visible = value
                BASE_HomepageFrm.imgSalesOrder.Visible = value
                BASE_HomepageFrm.lblSalesOrder.Visible = value
        Case 14: 'Sales Return
                BASE_ContainerFrm.Toolbar_Main.Buttons(5).ButtonMenus(2).Visible = value
        Case 15: 'Sales Adjustment
                BASE_ContainerFrm.Toolbar_Main.Buttons(5).ButtonMenus(3).Visible = value
        Case 16: 'New Customer
                BASE_ContainerFrm.Toolbar_Main.Buttons(5).ButtonMenus(5).Visible = value
                BASE_HomepageFrm.imgNewCustomer.Visible = value
                BASE_HomepageFrm.lblNewCustomer.Visible = value
        Case 17: 'Customer List
                BASE_ContainerFrm.Toolbar_Main.Buttons(5).ButtonMenus(6).Visible = value
        Case 18: 'Expenses
                BASE_ContainerFrm.Toolbar_Main.Buttons(7).ButtonMenus(4).Visible = value
                BASE_HomepageFrm.imgExpenses.Visible = value
                BASE_HomepageFrm.lblExpenses.Visible = value
        Case 19: 'Expenses List
                BASE_ContainerFrm.Toolbar_Main.Buttons(7).ButtonMenus(5).Visible = value
        Case 20: 'Accounts Receivable
                BASE_ContainerFrm.Toolbar_Main.Buttons(7).ButtonMenus(12).Visible = value
                BASE_HomepageFrm.imgAccountsReceivable.Visible = value
                BASE_HomepageFrm.lblAccountsReceivable.Visible = value
        Case 21: 'Accounts Payable
                BASE_ContainerFrm.Toolbar_Main.Buttons(7).ButtonMenus(13).Visible = value
        Case 22: 'Accounts Payable
                BASE_ContainerFrm.Toolbar_Main.Buttons(7).ButtonMenus(16).Visible = value
        Case 23: 'Reports
                BASE_ContainerFrm.Toolbar_Main.Buttons(9).Visible = value '.ButtonMenus(16).Visible = Value
        Case 24: 'General Settings
                BASE_ContainerFrm.Toolbar_Main.Buttons(11).ButtonMenus(1).Visible = value
                BASE_HomepageFrm.imgGeneralSettings.Visible = value
                BASE_HomepageFrm.lblGeneralSettings.Visible = value
        Case 25: 'System Settings
                BASE_ContainerFrm.Toolbar_Main.Buttons(11).ButtonMenus(2).Visible = value
                BASE_HomepageFrm.imgSystemSettings.Visible = value
                BASE_HomepageFrm.lblSystemSettings.Visible = value
        Case 26: 'New User
                BASE_GeneralSettingsFrm.btnUsers.Enabled = value
        Case 27: 'User Roles
                BASE_GeneralSettingsFrm.lblUserRoles.Visible = value
        Case 28: 'New Stock
                BASE_ContainerFrm.Toolbar_Main.Buttons(3).ButtonMenus(9).Visible = value
        Case 29: 'Audit Stock
                BASE_ContainerFrm.Toolbar_Main.Buttons(3).ButtonMenus(10).Visible = value
        Case 36: 'Penalty
                BASE_ContainerFrm.Toolbar_Main.Buttons(5).ButtonMenus(3).Visible = value
    End Select
    'return value for other purpose
     ViewAccessRights = value
End Function
Public Function EditAccessRights(ByVal ModuleId As Integer) As Boolean
    EditAccessRights = AccessRights(ModuleId, 2)
End Function
Public Function GetFileNameFromPath(strFullPath As String) As String
    GetFileNameFromPath = Right(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "\"))
End Function
Public Sub GetPOSSettings()
    Dim linevalue As String
    Open App.path & "\Resources\Settings.txt" For Input As #1
        Line Input #1, PharmacyMode 'Settings Title [PharmacyMode]
        Line Input #1, PharmacyMode 'Settings value
        Line Input #1, OrderSlipMode 'Settings title [OrderSlipMode]
        Line Input #1, OrderSlipMode 'settings value
        Line Input #1, DualPharmacyMode 'Settings title [DualPharmacyMode]
        Line Input #1, DualPharmacyMode 'settings value
        Line Input #1, linevalue 'null value
        Line Input #1, linevalue 'null value
        Line Input #1, linevalue 'null value
        Line Input #1, POS_Printer 'Settings title [POS Printer]
        Line Input #1, POS_Printer 'value
        Line Input #1, BackOffice_Printer 'Settings title [BackOffice Printer]
        Line Input #1, BackOffice_Printer 'value
    Close #1
    POSLogo = App.path & "\images\cashier_logo.jpg"
End Sub

Public Function CheckMachineRegistration() As Boolean
    'Check if machine is registered in the Server
    Dim ComputerName As String
    ComputerName = Environ("Computername")
    
    Dim con As New ADODB.Connection
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "POS_MachineRegistration_Check"
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, ComputerName)
    Set rec = cmd.Execute
    If Not rec.EOF Then
        If rec!isActive = "True" Then
            WorkstationId = rec!WorkstationId
            CheckMachineRegistration = True
        Else
            CheckMachineRegistration = False
            GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(55)
            GLOBAL_MessageFrm.Show (1)
        End If
    Else
        CheckMachineRegistration = False
        GLOBAL_MessageFrm.lblErrorMessage.Caption = ErrorCodes(55)
        GLOBAL_MessageFrm.Show (1)
    End If
    con.Close
End Function

Public Function GetTermDays(ByVal id As Integer)
    Dim con As New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "BASE_Terms_Get"
    cmd.Parameters.Append cmd.CreateParameter("@TermId", adInteger, adParamInput, , id)
    Set rec = cmd.Execute
    If Not rec.EOF Then
        GetTermDays = rec!DaysDue
    End If
    con.Close
End Function

Public Sub LoadImageStatus(ByVal picturebox As picturebox, ByVal Status As String)
    Status = UCase(Status)
    picturebox.Visible = True
    Select Case Status
        Case UCase("open")
            picturebox.Visible = False
        Case ""
            picturebox.Visible = False
        Case Else
            picturebox.Picture = LoadPicture(App.path & "\images\" & Status & ".jpg")
    End Select
    
End Sub

Public Function GetStatus(ByVal StatusId As Long) As String
    Dim con As New ADODB.Connection
    Set cmd = New ADODB.Command
    Set rec = New ADODB.Recordset
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "GLOBAL_DocStatus_Get"
    cmd.Parameters.Append cmd.CreateParameter("@StatusId", adInteger, adParamInput, , StatusId)
    Set rec = cmd.Execute
    If Not rec.EOF Then
        GetStatus = rec!Status
    End If
    con.Close
End Function

Public Function GetProductConversion(ByVal ProductId As String, ByVal UomId As Integer, ByVal ReturnType As String, Optional Text As TextBox = Nothing) As Double
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "INV_ProductConversion_Get"
    cmd.Parameters.Append cmd.CreateParameter("@ProductId", adInteger, adParamInput, , Val(ProductId))
    cmd.Parameters.Append cmd.CreateParameter("@UomId", adInteger, adParamInput, , UomId)
    Set rec = cmd.Execute
    If Not rec.EOF Then
       If Not rec.EOF Then
            If Not Text Is Nothing Then
                If ReturnType = "Cost" Then
                    Text.Text = FormatNumber(rec!cost, 2, vbTrue, vbFalse)
                Else
                    Text.Text = FormatNumber(rec!price, 2, vbTrue, vbFalse)
                End If
            End If
            GetProductConversion = rec!quantity
       End If
    End If
    con.Close
End Function

Public Function GetSalesSettings() As Integer
    Set con = New ADODB.Connection
    Set rec = New ADODB.Recordset
    Set cmd = New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "SO_Settings_Get"
    Set rec = cmd.Execute
    If Not rec.EOF Then
       Do Until rec.EOF
        StatementTemplateId = rec!StatementTemplateId
        rec.MoveNext
       Loop
    End If
    con.Close
End Function

Public Sub GetCSVData(ByVal filename As String, ByVal path As String, Optional ByVal importtype As String = "Product")
    Dim rst As ADODB.Recordset
    Dim cnn As ADODB.Connection
    
    'set up the connection
    Set cnn = New ADODB.Connection
    'cnn.ConnectionString = "Provider=MSDASQL.1;Extended Properties=""DBQ=" & path & ";Driver={Microsoft Text Driver (*.txt; *.csv)};DriverId=27;Extensions=csv;FIL=text;"""
    cnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;TEXT;DATABASE=" & path & ";"
    cnn.Open
    
    Set CSVRecordset = Nothing
    Set CSVRecordset = New ADODB.Recordset
    
    If importtype = "Product" Then
        CSVRecordset.Fields.Append "ITEMCODE", adVarChar, 500
        CSVRecordset.Fields.Append "NAME", adVarChar, 4000
        CSVRecordset.Fields.Append "UNIT", adVarChar, 50
        CSVRecordset.Fields.Append "BARCODE", adVarChar, 250
        CSVRecordset.Fields.Append "CATEGORY", adVarChar, 250
        CSVRecordset.Fields.Append "TAX", adVarChar, 50
        CSVRecordset.Fields.Append "SELLINGPRICE", adDecimal
                   CSVRecordset.Fields("SELLINGPRICE").Precision = 18
                   CSVRecordset.Fields("SELLINGPRICE").NumericScale = 2
        CSVRecordset.Fields.Append "COST", adDecimal
                   CSVRecordset.Fields("COST").Precision = 18
                   CSVRecordset.Fields("COST").NumericScale = 2
        CSVRecordset.Fields.Append "SUPPLIER", adVarChar, 4000
    End If
    Set CSVRecordset = cnn.Execute("SELECT * FROM " & filename & "")
    
    With CSVRecordset
        If Not .EOF Then
            Do Until .EOF
                UniversalCtr = UniversalCtr + 1
                .MoveNext
            Loop
        End If
        .MoveFirst
    End With
End Sub

Public Function CategoryImport(ByVal Name As String) As Long
    Dim con As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Set rec = New ADODB.Recordset
    
    Dim CategoryId As Long
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "SYS_Import_Category"
    
    cmd.Parameters.Append cmd.CreateParameter("@CategoryId", adInteger, adParamInputOutput, , CategoryId)
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, Name)
    
    Set rec = cmd.Execute
    CategoryImport = cmd.Parameters("@CategoryId")
    
    con.Close
End Function

Public Function SupplierImport(ByVal Name As String) As Long
    Dim con As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Set rec = New ADODB.Recordset
    
    Dim SupplierId As Long
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "SYS_Import_Supplier"
    
    cmd.Parameters.Append cmd.CreateParameter("@SupplierId", adInteger, adParamInputOutput, , SupplierId)
    cmd.Parameters.Append cmd.CreateParameter("@Name", adVarChar, adParamInput, 250, Name)
    
    Set rec = cmd.Execute
    SupplierImport = cmd.Parameters("@SupplierId")
'    MsgBox cmd.Parameters("@SupplierId")
    
    con.Close
End Function

Public Function UomImport(ByVal Name As String) As Long
    Dim con As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Set rec = New ADODB.Recordset
    
    Dim UomId As Long
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "SYS_Import_Uom"
    
    cmd.Parameters.Append cmd.CreateParameter("@UomId", adInteger, adParamInputOutput, , UomId)
    cmd.Parameters.Append cmd.CreateParameter("@UomName", adVarChar, adParamInput, 250, Name)
    
    Set rec = cmd.Execute
    UomImport = cmd.Parameters("@UomId")
    
    con.Close
End Function


Public Sub ClearDataImportLog()
    Dim con As New ADODB.Connection
    Dim cmd As New ADODB.Command
    Set rec = New ADODB.Recordset
    
    Dim UomId As Long
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "SYS_DataImportLog_Delete"
    
    Set rec = cmd.Execute
    
    con.Close
End Sub

Public Sub RemoveDuplicates(ByVal lv As ListView, ByVal ColumntoEvaluate As Integer)
    Dim item1 As MSComctlLib.ListItem
    Dim item2 As MSComctlLib.ListItem
    
    For Each item1 In lv.ListItems
        For Each item2 In lv.ListItems
            If item1.SubItems(ColumntoEvaluate) = item2.SubItems(ColumntoEvaluate) Then
                
            End If
        Next
    Next
End Sub


Public Sub UpdateCustomerIdonPOSSales()
    Dim con As New ADODB.Connection
    Dim cmd As New ADODB.Command
    
    con.ConnectionString = ConnString
    con.Open
    cmd.ActiveConnection = con
    cmd.CommandType = adCmdStoredProc
    cmd.CommandText = "POS_SalesCustomerId_Update"
    cmd.Execute
End Sub

Function DefaultPrinter(Printer As String) 'set defualt printer
    On Error Resume Next
    Dim SetDefaultPrint As New WshNetwork
    SetDefaultPrint.SetDefaultPrinter (Printer)
    Set SetDefaultPrint = Nothing
End Function

Public Function GetVersion() As String
    GetVersion = "v" & App.Major & "." & App.Minor & "." & App.Revision
End Function
