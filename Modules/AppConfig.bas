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
        ComputeInterest (EnableInterest)
        
        DeleteReserves WorkstationId, 1 'POS
        DeleteReserves WorkstationId, 2 'SALES
        DeleteReserves WorkstationId, 3 'PR
        DeleteReserves WorkstationId, 4 'Transfer Stock
        
        'GEN_PatchExtendedFRM.Show
        BASE_UserLoginFrm.Show
        DefaultPrinter (BackOffice_Printer)

'        POS_UserLoginFrm.Show
'        DefaultPrinter (POS_Printer)
'    Else
'        MsgBox "Invalid license.", vbCritical, "PeakPOS"
'    End If
End Sub

