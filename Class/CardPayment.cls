VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CardPaymentInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private CurrentBankId As Integer
Private CurrentCardTypeId As Integer
Private CurrentNameOnCard As String
Private CurrentCardNumber As String
Private CurrentReference As String
Private CurrentAmount As Double
Public Property Let BankId(ByVal newValue As Integer)
    CurrentBankId = newValue
End Property
Public Property Get BankId() As Integer
    BankId = CurrentBankId
End Property

Public Property Let CardTypeId(ByVal newValue As Integer)
    CurrentCardTypeId = newValue
End Property

Public Property Get CardTypeId() As Integer
    CardTypeId = CurrentCardTypeId
End Property

Public Property Let NameOnCard(ByVal newValue As String)
    CurrentNameOnCard = newValue
End Property

Public Property Get NameOnCard() As String
    NameOnCard = CurrentNameOnCard
End Property

Public Property Let CardNumber(ByVal newValue As String)
    CurrentCardNumber = newValue
End Property

Public Property Get CardNumber() As String
    CardNumber = CurrentCardNumber
End Property
Public Property Let Reference(ByVal newValue As String)
    CurrentReference = newValue
End Property

Public Property Get Reference() As String
    Reference = CurrentReference
End Property
Public Property Let Amount(ByVal newValue As Double)
    CurrentAmount = newValue
End Property

Public Property Get Amount() As Double
    Amount = CurrentAmount
End Property
