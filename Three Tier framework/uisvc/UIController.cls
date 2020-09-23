VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UIController"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private myController As New UCController

Public rsCustomers As New Recordset
Public mcolCustomers As Collection
Public mcolOrders As Collection

Private mcolCustomersInterchange As Collection
Private mcolOrdersInterchange As Collection

Public Function getCustomers(Optional lErrNum As Long, Optional sErrDesc As String, Optional sErrSource As String) As Boolean
    On Error GoTo getCustomersErr
    
    Dim xmlCustomers As String
    Dim xmlOrders As String
    
    'Ask to the UC Controller for the data
    If myController.getCustomers(xmlCustomers, xmlOrders, lErrNum, sErrDesc, sErrSource) Then
        LoadCustomer xmlCustomers
        getCustomers = True
    End If
    
    Exit Function
    
getCustomersErr:
    If lErrNum = 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "UIController.getCustomers"
    End If
End Function

Private Function xml2Recordset(sXML As String) As Recordset
    Dim oDOM As DOMDocument
    Dim rs As Recordset
    
    Set oDOM = New DOMDocument
    Set rs = New Recordset
    
    If oDOM.loadXML(sXML) Then
        rs.Open oDOM
    Else
        MsgBox "No se pudo abrir la cadena XML"
    End If
    
    Set oDOM = Nothing
    Set xml2Recordset = rs
    Set rs = Nothing
End Function

Public Function getCustomer(ByVal lngIDCustomer As Long, Optional lErrNum As Long, Optional sErrDesc As String, Optional sErrSource As String) As Boolean
    On Error GoTo getCustomerErr
    
    Dim xmlCustomers As String
    Dim xmlOrders As String
    
    'Ask to the UC Controller for the data
    If myController.getCustomer(lngIDCustomer, xmlCustomers, xmlOrders, lErrNum, sErrDesc, sErrSource) Then
        'Fill the Customer Data UDTs
        LoadCustomer xmlCustomers
        LoadOrder xmlOrders
        getCustomer = True
    End If
    
    Exit Function
    
getCustomerErr:
    If lErrNum = 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "UIController.getCustomer"
    End If
End Function

Public Sub LoadCustomer(xmlCustomer As String)
    Dim oDOM As New DOMDocument
    Dim node As MSXML2.IXMLDOMElement
    Dim customerInfo As CustomerUDT
    
    If oDOM.loadXML(xmlCustomer) Then
        'Here you can also keep the data in memory and make a method to verify
        'if certain item exists in the collection so you don't add it.
        Set mcolCustomers = New Collection
        For Each node In oDOM.documentElement.selectSingleNode("rs:data").childNodes
            customerInfo.IDCustomer = CLng(node.Attributes(giIDCustomerCustomers).Text)
            customerInfo.CustomerName = node.Attributes(giNameCustomers).Text
            
            mcolCustomers.Add customerInfo, "CU-" & customerInfo.IDCustomer
        Next node
    Else
        'Wrong
    End If
End Sub

Public Sub LoadOrder(xmlOrders As String)
    Dim oDOM As New DOMDocument
    Dim node As MSXML2.IXMLDOMElement
    Dim orderInfo As OrderUDT
    
    If oDOM.loadXML(xmlOrders) Then
        Set mcolOrders = New Collection
        For Each node In oDOM.documentElement.selectSingleNode("rs:data").childNodes
            With orderInfo
                .DateOrder = Format(CDate(node.Attributes(giDateOrderOrders).Text), "mm-dd-yyyy hh:mm")
                .IDCustomer = CLng(node.Attributes(giIDCustomerOrders).Text)
                .IDOrder = CLng(node.Attributes(giIDOrderOrders).Text)
            End With
            
            mcolOrders.Add orderInfo, "O-" & orderInfo.IDOrder
        Next node
    Else
        'Wrong
    End If
End Sub

Public Function saveCustomers(Optional lErrNum As Long, Optional sErrDesc As String, Optional sErrSource As String) As Boolean
    On Error GoTo saveCustomersErr
    
    Dim xmlCustomers As String
    Dim xmlOrders As String
    Dim i As Integer
    Dim customerInfo As CustomerUDT
    
    'Add those Customers that need to be created (IsNew=True)
    Set mcolCustomersInterchange = New Collection
    For i = 1 To mcolCustomers.Count
        customerInfo = mcolCustomers(i)
        If customerInfo.IsNew Then
            mcolCustomersInterchange.Add customerInfo, "CU-" & customerInfo.IDCustomer
        End If
    Next i
    
    uploadCustomer xmlCustomers
    
    If myController.saveCustomers(xmlCustomers, xmlOrders, lErrNum, sErrDesc, sErrSource) Then
        saveCustomers = True
    End If
    
    Exit Function
    
saveCustomersErr:
    If lErrNum = 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "UIController.saveCustomers"
    End If
End Function

Public Function deleteCustomers(Optional lErrNum As Long, Optional sErrDesc As String, Optional sErrSource As String) As Boolean
    On Error GoTo deleteCustomersErr
    
    Dim xmlCustomers As String
    Dim i As Integer
    Dim customerInfo As CustomerUDT
    
    'Add those Customers that need to be created (IsNew=True)
    Set mcolCustomersInterchange = New Collection
    For i = 1 To mcolCustomers.Count
        customerInfo = mcolCustomers(i)
        If customerInfo.IsDeleted Then
            mcolCustomersInterchange.Add customerInfo, "CU-" & customerInfo.IDCustomer
        End If
    Next i
    
    uploadCustomer xmlCustomers
    
    If myController.deleteCustomers(xmlCustomers, lErrNum, sErrDesc, sErrSource) Then
        deleteCustomers = True
    End If
    
    Exit Function
    
deleteCustomersErr:
    If lErrNum = 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "UIController.deleteCustomers"
    End If
End Function

Public Function modifyCustomers(Optional lErrNum As Long, Optional sErrDesc As String, Optional sErrSource As String) As Boolean
    On Error GoTo modifyCustomersErr
    
    Dim xmlCustomers As String
    Dim i As Integer
    Dim customerInfo As CustomerUDT
    
    'Add those Customers that need to be created (IsNew=True)
    Set mcolCustomersInterchange = New Collection
    For i = 1 To mcolCustomers.Count
        customerInfo = mcolCustomers(i)
        If customerInfo.IsDirty Then
            mcolCustomersInterchange.Add customerInfo, "CU-" & customerInfo.IDCustomer
        End If
    Next i
    
    uploadCustomer xmlCustomers
    
    If myController.modifyCustomers(xmlCustomers, lErrNum, sErrDesc, sErrSource) Then
        modifyCustomers = True
    End If
    
    Exit Function
    
modifyCustomersErr:
    If lErrNum = 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "UIController.modifyCustomers"
    End If
End Function

Public Function saveOrders(Optional lErrNum As Long, Optional sErrDesc As String, Optional sErrSource As String) As Boolean
    On Error GoTo saveOrdersErr
    
    Dim xmlOrders As String
    Dim i As Integer
    Dim orderInfo As OrderUDT
    
    'Add those Customers that need to be created (IsNew=True)
    Set mcolOrdersInterchange = New Collection
    For i = 1 To mcolOrders.Count
        orderInfo = mcolOrders(i)
        If orderInfo.IsNew Then
            mcolOrdersInterchange.Add orderInfo, "O-" & orderInfo.IDOrder
        End If
    Next i
    
    uploadOrder xmlOrders
    
    If myController.saveOrders(xmlOrders, lErrNum, sErrDesc, sErrSource) Then
        saveOrders = True
    End If
    
    Exit Function
    
saveOrdersErr:
    If lErrNum = 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "UIController.saveOrders"
    End If
End Function

Public Sub uploadCustomer(ByRef xmlCustomer As String)
    Dim oDOM As New DOMDocument
    Dim node As MSXML2.IXMLDOMElement
    Dim customerInfo As CustomerUDT
    Dim cust As New MSXML2.DOMDocument
    Dim attr As MSXML2.IXMLDOMAttribute
    
    Dim xmlHeading As String
    Dim i As Integer
    
    Call myController.getXMLHeading("Customers", xmlHeading)
    
    oDOM.loadXML xmlHeading
    oDOM.documentElement.appendChild oDOM.createElement("rs:data")
    For i = 1 To mcolCustomersInterchange.Count
        Set cust = New MSXML2.DOMDocument
        customerInfo = mcolCustomersInterchange(i)

        cust.appendChild cust.createElement("z:row")
        Set attr = cust.createAttribute("IDCustomer")
        attr.Value = customerInfo.IDCustomer
        Set attr = cust.selectSingleNode("z:row").Attributes.setNamedItem(attr)
        
        Set attr = cust.createAttribute("Name")
        attr.Value = customerInfo.CustomerName
        Set attr = cust.selectSingleNode("z:row").Attributes.setNamedItem(attr)
        
        oDOM.documentElement.selectSingleNode("rs:data").appendChild cust.documentElement
    Next i
    
    xmlCustomer = oDOM.xml
End Sub

Public Sub uploadOrder(ByRef xmlOrders As String)
    Dim oDOM As New DOMDocument
    Dim node As MSXML2.IXMLDOMElement
    Dim cust As New MSXML2.DOMDocument
    Dim attr As MSXML2.IXMLDOMAttribute
    Dim orderInfo As OrderUDT
    
    Dim xmlHeading As String
    Dim i As Integer
    
    Call myController.getXMLHeading("Orders", xmlHeading)
    
    oDOM.loadXML xmlHeading
    oDOM.documentElement.appendChild oDOM.createElement("rs:data")
    For i = 1 To mcolOrdersInterchange.Count
        Set cust = New MSXML2.DOMDocument
        orderInfo = mcolOrdersInterchange(i)

        cust.appendChild cust.createElement("z:row")
        
'This is for every attribute of the Order
'It's important to keep the column Order that you have in the database
'if the Column Order is:
'ID
'Name
'Then you have to create the node "ID" in the fisrt place
'------------------------------------------------------------------------------------------------
        Set attr = cust.createAttribute("IDOrder")
        attr.Value = orderInfo.IDOrder
        Set attr = cust.selectSingleNode("z:row").Attributes.setNamedItem(attr)
'------------------------------------------------------------------------------------------------
        
        Set attr = cust.createAttribute("IDCustomer")
        attr.Value = orderInfo.IDCustomer
        Set attr = cust.selectSingleNode("z:row").Attributes.setNamedItem(attr)
        
        Set attr = cust.createAttribute("DateOrder")
        attr.Value = orderInfo.DateOrder
        Set attr = cust.selectSingleNode("z:row").Attributes.setNamedItem(attr)
        
        oDOM.documentElement.selectSingleNode("rs:data").appendChild cust.documentElement
    Next i
    
    xmlOrders = oDOM.xml
End Sub

