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
Attribute VB_Description = "Clase Diccionario"
Option Explicit

Public Enum eCustomers
    giIDCustomerCustomers = 0
    giNameCustomers
End Enum

Public Enum eOrders
    giIDOrdersOrders = 0
    giIDCustomerOrders
    giDateOrderOrders
End Enum

