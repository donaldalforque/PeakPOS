Attribute VB_Name = "AppConfig"
Option Explicit
Sub Main()
    'If GetRegistration = True Then
        'Register Machine
        CheckMachineRegistration
    
        'Get Settings
        GetInventorySettings
        GetPOSSettings
        UpdateCustomerOrderDues
        UpdateVendorOrderDues
        GetSalesSettings
        
        DeleteReserves WorkstationId, 1 'POS
        DeleteReserves WorkstationId, 2 'SALES
        DeleteReserves WorkstationId, 3 'PR
        DeleteReserves WorkstationId, 4 'Transfer Stock
        
        'BASE_UserLoginFrm.Show
        POS_UserLoginFrm.Show
'    Else
'        MsgBox "Invalid license.", vbCritical, "PeakPOS"
'    End If
End Sub

