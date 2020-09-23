VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Dictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type CustomerUDT
    IDCustomer As Long
    CustomerName As String
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type

Public Enum eCustomers
    giIDCustomerCustomers = 0
    giNameCustomers
End Enum

Public Type OrderUDT
    IDOrder As Long
    IDCustomer As Long
    DateOrder As Date
    IsNew As Boolean
    IsDirty As Boolean
    IsDeleted As Boolean
End Type

Public Enum eOrders
    giIDOrderOrders = 0
    giIDCustomerOrders
    giDateOrderOrders
End Enum


