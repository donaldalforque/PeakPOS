VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OtherPaymentInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private CurrentReferenceNumber As String
Private CurrentRemarks As String
Private CurrentAmount As Double

Public Property Let ReferenceNumber(ByVal newValue As String)
    CurrentReferenceNumber = newValue
End Property
Public Property Get ReferenceNumber() As String
    ReferenceNumber = CurrentReferenceNumber
End Property

Public Property Let Remarks(ByVal newValue As String)
    CurrentRemarks = newValue
End Property
Public Property Get Remarks() As String
    Remarks = CurrentRemarks
End Property

Public Property Let Amount(ByVal newValue As Double)
    CurrentAmount = newValue
End Property
Public Property Get Amount() As Double
    Amount = CurrentAmount
End Property
