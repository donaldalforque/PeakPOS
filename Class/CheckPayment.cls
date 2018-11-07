VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CheckPaymentInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private CurrentBankId As Integer
Private CurrentCheckNumber As String
Private CurrentCheckDate As Date
Private CurrentAmount As Double
Public Property Let BankId(ByVal newValue As Integer)
    CurrentBankId = newValue
End Property
Public Property Get BankId() As Integer
    BankId = CurrentBankId
End Property
Public Property Let CheckNumber(ByVal newValue As String)
    CurrentCheckNumber = newValue
End Property
Public Property Get CheckNumber() As String
    CheckNumber = CurrentCheckNumber
End Property
Public Property Let CheckDate(ByVal newValue As Date)
    CurrentCheckDate = newValue
End Property
Public Property Get CheckDate() As Date
    CheckDate = CurrentCheckDate
End Property
Public Property Let Amount(ByVal newValue As Double)
    CurrentAmount = newValue
End Property
Public Property Get Amount() As Double
    Amount = CurrentAmount
End Property

