VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UCController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private mcolCustomers As Collection

'This Class implements every Use Case Path
'One Controller per Use Case

'In general we need to ask for ALL the Customers for filling a ListBox or a Grid
'Then, when the User select a particular Customer, ask for the Customer info = create an instance of Customer
'in the Business layer
Public Function getCustomers(xmlCustomers As String, xmlOrders As String, Optional lErrNum As Long, Optional sErrDesc As String, Optional sErrSource As String) As Boolean
    On Error GoTo getCustomersErr
    
    Dim oCust As New CCustomer
    
    If ogAccess.Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        If oCust.readAll(lErrNum, sErrDesc, sErrSource) Then
            If oCust.getSate(xmlCustomers, xmlOrders, True, lErrNum, sErrDesc, sErrSource) Then
                getCustomers = True
                Set oCust = Nothing
            End If
        End If
        
        Call ogAccess.closeConnection(lErrNum, sErrDesc, sErrSource)
    End If
    
    Exit Function
    
getCustomersErr:
    If lErrNum = 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "UCController.getCustomers"
    End If
End Function

Public Function getCustomer(lngIDCustomer As Long, xmlCustomers As String, xmlOrders As String, Optional lErrNum As Long, Optional sErrDesc As String, Optional sErrSource As String) As Boolean
    On Error GoTo getCustomersErr
    
    Dim oCust As New CCustomer
    
    If ogAccess.Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        If oCust.read(lngIDCustomer, True, lErrNum, sErrDesc, sErrSource) Then
            If oCust.getSate(xmlCustomers, xmlOrders, False, lErrNum, sErrDesc, sErrSource) Then
                Set mcolCustomers = New Collection
                mcolCustomers.Add oCust, "CU-" & oCust.IDCustomer
                getCustomer = True
                Set oCust = Nothing
            End If
        End If
        
        Call ogAccess.closeConnection(lErrNum, sErrDesc, sErrSource)
    End If
    
    Exit Function
    
getCustomersErr:
    If lErrNum = 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "UCController.getCustomers"
    End If
End Function

Public Function getXMLHeading(strTableName As String, xmlHeading As String, Optional lErrNum As Long, Optional sErrDesc As String, Optional sErrSource As String) As Boolean
    On Error GoTo getXMLHeadingErr
    
    If ogAccess.Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        If ogAccess.getXMLHeading(strTableName, xmlHeading, lErrNum, sErrDesc, sErrSource) Then
            getXMLHeading = True
        End If
        
        Call ogAccess.closeConnection(lErrNum, sErrDesc, sErrSource)
    End If
    
    Exit Function
    
getXMLHeadingErr:
    If lErrNum = 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "UCController.getXMLHeading"
    End If
End Function

'ESTE METODO DEBE RECIBIR LA CADENA XML Y DE ACUERDO A LOS DATOS DE LOS CLIENTES PUESTOS
'EN ESA CADENA, ENTONCES INSTANCIARÁ ESOS OBJETOS Y LES MANDARÁ UN MENSAJE GRABAR CON
'ISNEW=TRUE. PARA ELLO, DEBO LEVANTAR ESOS OBJETOS DESDE LA CADENA XML QUE RECIBO.

Public Function saveCustomers(xmlCustomers As String, xmlOrders As String, Optional lErrNum As Long, Optional sErrDesc As String, Optional sErrSource As String) As Boolean
    On Error GoTo saveCustomersErr
    
    Dim customer As CCustomer
    
    loadCustomer xmlCustomers
    
    If ogAccess.Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        'You can add here the initiation of a transaction
        'ogAccess.beginTrans ...etc
        For Each customer In mcolCustomers
            customer.IsNew = True
            If Not customer.Save(lErrNum, sErrDesc, sErrSource) Then
                Exit Function
            End If
        Next customer
        
        Call ogAccess.closeConnection(lErrNum, sErrDesc, sErrSource)
    End If
    
    saveCustomers = True
    
    Exit Function
    
saveCustomersErr:
    If lErrNum = 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "UCController.saveCustomers"
    End If
End Function


Public Sub loadCustomer(xmlCustomer As String)
    Dim oDOM As New DOMDocument
    Dim node As MSXML2.IXMLDOMElement
    Dim oCust As CCustomer
    
    Set mcolCustomers = New Collection
    If oDOM.loadXML(xmlCustomer) Then
        For Each node In oDOM.documentElement.selectSingleNode("rs:data").childNodes
            Set oCust = New CCustomer
            oCust.IDCustomer = CLng(node.Attributes(giIDCustomerCustomers).Text)
            oCust.Name = node.Attributes(giNameCustomers).Text
            
            mcolCustomers.Add oCust, "CU-" & oCust.IDCustomer
        Next node
    Else
        'Wrong
    End If
End Sub

Public Function deleteCustomers(xmlCustomers As String, Optional lErrNum As Long, Optional sErrDesc As String, Optional sErrSource As String) As Boolean
    On Error GoTo deleteCustomersErr
    
    Dim customer As CCustomer
    
    loadCustomer xmlCustomers
    
    If ogAccess.Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        'You can add here the initiation of a transaction
        'ogAccess.beginTrans ...etc
        For Each customer In mcolCustomers
            customer.IsDeleted = True
            If Not customer.Save(lErrNum, sErrDesc, sErrSource) Then
                Exit Function
            End If
        Next customer
        
        Call ogAccess.closeConnection(lErrNum, sErrDesc, sErrSource)
    End If
    
    deleteCustomers = True
    
    Exit Function
    
deleteCustomersErr:
    If lErrNum = 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "UCController.deleteCustomers"
    End If
End Function

Public Function modifyCustomers(xmlCustomers As String, Optional lErrNum As Long, Optional sErrDesc As String, Optional sErrSource As String) As Boolean
    On Error GoTo modifyCustomersErr
    
    Dim customer As CCustomer
    
    loadCustomer xmlCustomers
    
    If ogAccess.Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        'You can add here the initiation of a transaction
        'ogAccess.beginTrans ...etc
        For Each customer In mcolCustomers
            customer.IsDirty = True
            If Not customer.Save(lErrNum, sErrDesc, sErrSource) Then
                Exit Function
            End If
        Next customer
        
        Call ogAccess.closeConnection(lErrNum, sErrDesc, sErrSource)
    End If
    
    modifyCustomers = True
    
    Exit Function
    
modifyCustomersErr:
    If lErrNum = 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "UCController.modifyCustomers"
    End If
End Function

Public Function saveOrders(xmlOrders As String, Optional lErrNum As Long, Optional sErrDesc As String, Optional sErrSource As String) As Boolean
    On Error GoTo saveOrdersErr
    
    Dim customer As CCustomer
    
    loadOrder xmlOrders
    
    If ogAccess.Connect(msDSN, lErrNum, sErrDesc, sErrSource) Then
        'You can add here the initiation of a transaction
        'ogAccess.beginTrans ...etc
        For Each customer In mcolCustomers
            customer.IsDirty = True
            If Not customer.Save(lErrNum, sErrDesc, sErrSource) Then
                Exit Function
            End If
        Next customer
        
        Call ogAccess.closeConnection(lErrNum, sErrDesc, sErrSource)
    End If
    
    saveOrders = True
    
    Exit Function
    
saveOrdersErr:
    If lErrNum = 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "UCController.saveCustomers"
    End If
End Function

Public Sub loadOrder(xmlOrders As String)
    Dim oDOM As New DOMDocument
    Dim node As MSXML2.IXMLDOMElement
    Dim oOrd As COrder
    Dim oCust As CCustomer
    
    If oDOM.loadXML(xmlOrders) Then
        For Each node In oDOM.documentElement.selectSingleNode("rs:data").childNodes
            'Create the instance of the Order
            Set oOrd = New COrder
            oOrd.myCustomer.IDCustomer = CLng(node.Attributes(giIDCustomerOrders).Text)
            oOrd.DateOrder = node.Attributes(giDateOrderOrders).Text
            oOrd.IDOrder = node.Attributes(giIDOrdersOrders).Text
            oOrd.IsNew = True
            
            'Search for the Customer owner of this Order
            Set oCust = mcolCustomers("CU-" & oOrd.myCustomer.IDCustomer)
            Set oOrd.myCustomer = oCust
            oCust.myOrders.Add oOrd, "O-" & oOrd.IDOrder
            oCust.IsDirty = True
        Next node
    Else
        'Wrong
    End If
End Sub


