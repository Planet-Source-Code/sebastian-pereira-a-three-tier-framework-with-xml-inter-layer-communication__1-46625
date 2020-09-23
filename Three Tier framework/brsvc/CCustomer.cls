VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Clase de datos de la tabla Customers"
Option Explicit

Private SQL As String
Private sXML As String
Private mIDCustomer As Long
Private mName As String
Private mIsNew As Boolean
Private mIsDirty As Boolean
Private mIsDeleted As Boolean

Public mcolCustomers As New Collection
'Variable to reference the Orders of the Customer
'You can also build a Collection Class of the COrder Class
Public myOrders As New Collection

Property Get IDCustomer() As Long
    IDCustomer = mIDCustomer
End Property

Property Let IDCustomer(Valor As Long)
    mIDCustomer = Valor
End Property

Property Get Name() As String
    Name = mName
End Property

Property Let Name(Valor As String)
    mName = Valor
End Property

Property Get IsNew() As Boolean
    IsNew = mIsNew
End Property

Property Get IsDeleted() As Boolean
    IsDeleted = mIsDeleted
End Property

Property Get IsDirty() As Boolean
    IsDirty = mIsDirty
End Property

Property Let IsNew(Valor As Boolean)
    mIsNew = Valor
End Property

Property Let IsDeleted(Valor As Boolean)
    mIsDeleted = Valor
End Property

Property Let IsDirty(Valor As Boolean)
    mIsDirty = Valor
End Property

Public Function read(ByVal IDCustomer As Long, bExpand As Boolean, Optional lErrNum As Long, Optional sErrDesc As String, Optional sErrSource As String) As Boolean
    On Error GoTo Errread

    Dim oOrder As COrder

    SQL = "SELECT * FROM Customers WHERE IDCustomer=" & IDCustomer

    If ogAccess.retrieveData(SQL, sXML, lErrNum, sErrDesc, sErrSource) Then
        If unPack(sXML, lErrNum, sErrDesc, sErrSource) Then
            read = True
            If bExpand Then
                Set oOrder = New COrder
                Set oOrder.myCustomer = Me
                If oOrder.readByCustomer(lErrNum, sErrDesc, sErrSource) Then
                    Set myOrders = oOrder.mcolOrders
                End If
            End If
        End If
    End If

    Exit Function

Errread:
    If lErrNum = 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "CCustomer.read"
    End If
End Function

Private Function unPack(sXML As String, Optional lErrNum As Long, Optional sErrDesc As String, Optional sErrSource As String) As Boolean
    On Error GoTo ErrunPack

    Dim oDOMDocument As New MSXML2.DOMDocument
    Dim oRoot As MSXML2.IXMLDOMElement 'Estos son todos los datos
    Dim oDatos As MSXML2.IXMLDOMElement 'Este objeto contiene los datos
    Dim oFila As MSXML2.IXMLDOMElement 'Este contiene las filas
    unPack = False

    If oDOMDocument.loadXML(sXML) Then
        Set oRoot = oDOMDocument.documentElement
        Set oDatos = oRoot.childNodes(1)
        Set oFila = oDatos.childNodes.Item(0)
        
        mIDCustomer = oFila.Attributes(giIDCustomerCustomers).nodeValue
        mName = oFila.Attributes(giNameCustomers).nodeValue
        
        IsDeleted = False
        IsNew = False
        IsDirty = False

        unPack = True

        Set oRoot = Nothing
        Set oDatos = Nothing
        Set oDOMDocument = Nothing
        Set oFila = Nothing
    End If

    Exit Function

ErrunPack:
    If lErrNum = 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "CCustomer.unPack"
    End If
End Function

Private Function unPackDep(sXML As String, Optional lErrNum As Long, Optional sErrDesc As String, Optional sErrSource As String) As Boolean
    On Error GoTo ErrunPackDep

    Dim iIndice As Integer
    Dim oCustomer As CCustomer

    Dim oDOMDocument As New MSXML2.DOMDocument
    Dim oRoot As MSXML2.IXMLDOMElement 'Estos son todos los datos
    Dim oDatos As MSXML2.IXMLDOMElement 'Este objeto contiene los datos
    Dim oFila As MSXML2.IXMLDOMElement 'Este contiene las filas

    If oDOMDocument.loadXML(sXML) Then
        Set oRoot = oDOMDocument.documentElement
        Set oDatos = oRoot.childNodes(1)
        
        For iIndice = 0 To oDatos.childNodes.length - 1
            Set oFila = oDatos.childNodes.Item(iIndice)
            Set oCustomer = New CCustomer
            With oCustomer
                .IDCustomer = CLng(oFila.Attributes(giIDCustomerCustomers).nodeValue)
                .Name = oFila.Attributes(giNameCustomers).nodeValue
            End With
        
            mcolCustomers.Add oCustomer, "CU-" & oCustomer.IDCustomer
        Next iIndice
    End If
    
    unPackDep = True

    Exit Function

ErrunPackDep:
    If lErrNum = 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "CCustomer.unPackDep"
    End If
End Function

Public Function Save(Optional lErrNum As Long, Optional sErrDesc As String, Optional sErrSource As String) As Boolean
    On Error GoTo SaveErr
    
    Dim theOrder As COrder
    
    Save = False
    
    If mIsNew Then
        SQL = "INSERT INTO Customers (IDCustomer, Name) VALUES (" & IDCustomer & ", '" & Name & "'" & ")"
        mIsNew = False
        Save = True
    Else
        If mIsDirty Then
            SQL = "UPDATE Customers SET Name='" & Name & "'" & " WHERE IDCustomer=" & IDCustomer
            mIsDirty = False
            Save = True
        Else
            If mIsDeleted Then
                SQL = "DELETE FROM Customers WHERE IDCustomer=" & IDCustomer
                mIsDeleted = False
                Save = True
            End If
        End If
    End If

    If Save Then
        If ogAccess.query(SQL, lErrNum, sErrDesc, sErrSource) Then
            'Verify if this Customer has Orders to save
            For Each theOrder In myOrders
                If theOrder.IsNew Or theOrder.IsDeleted Or theOrder.IsDirty Then
                    If Not theOrder.Save(lErrNum, sErrDesc, sErrSource) Then
                        Save = False
                        Exit Function
                    End If
                End If
            Next theOrder
            Save = True
        End If
    End If

    Exit Function

SaveErr:
    If lErrNum = 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "CCustomer.Save"
    End If
End Function

Public Function readAll(Optional lErrNum As Long, Optional sErrDesc As String, Optional sErrSource As String) As Boolean
    On Error GoTo readAllErr
    
    SQL = "SELECT * FROM Customers"

    If ogAccess.retrieveData(SQL, sXML, lErrNum, sErrDesc, sErrSource) Then
        If unPackDep(sXML, lErrNum, sErrDesc, sErrSource) Then
            readAll = True
        End If
    End If

    Exit Function

readAllErr:
    If lErrNum = 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "CCustomer.readAll"
    End If
End Function

'Convert the state of the object into XML
Public Function getSate(xmlCustomers As String, xmlOrders As String, Optional bMultiple As Boolean, Optional lErrNum As Long, Optional sErrDesc As String, Optional sErrSource As String) As Boolean
    On Error GoTo getSateErr

    Dim oCust As CCustomer
    Dim cust As New MSXML2.DOMDocument
    Dim node As MSXML2.IXMLDOMElement
    Dim oDOM As New MSXML2.DOMDocument
    Dim attr As MSXML2.IXMLDOMAttribute
    
    Dim xmlHeading As String
    
    Dim xmlOrdersTemp As String
    Dim oOrder As COrder
    
    'This is the heading that we need for opening recordsets with XML
    ogAccess.getXMLHeading "Customers", xmlHeading
    
    oDOM.loadXML xmlHeading

    oDOM.documentElement.appendChild oDOM.createElement("rs:data")
    
    'Multiple tells me that I need to get the xml string of many Customers
    'When I'm asking for a particular Customer, Multiple=False
    If Not bMultiple Then
        Set cust = New MSXML2.DOMDocument

        cust.appendChild cust.createElement("z:row")
        Set attr = cust.createAttribute("IDCustomer")
        attr.Value = mIDCustomer
        Set attr = cust.selectSingleNode("z:row").Attributes.setNamedItem(attr)
        
        Set attr = cust.createAttribute("Name")
        attr.Value = mName
        Set attr = cust.selectSingleNode("z:row").Attributes.setNamedItem(attr)
        
        oDOM.documentElement.selectSingleNode("rs:data").appendChild cust.documentElement
    Else
        For Each oCust In mcolCustomers
            Set cust = New MSXML2.DOMDocument

            cust.appendChild cust.createElement("z:row")
            Set attr = cust.createAttribute("IDCustomer")
            attr.Value = oCust.IDCustomer
            Set attr = cust.selectSingleNode("z:row").Attributes.setNamedItem(attr)
            
            Set attr = cust.createAttribute("Name")
            attr.Value = oCust.Name
            Set attr = cust.selectSingleNode("z:row").Attributes.setNamedItem(attr)
            
            oDOM.documentElement.selectSingleNode("rs:data").appendChild cust.documentElement
        Next oCust
    End If

    xmlCustomers = oDOM.xml
    
    ogAccess.getXMLHeading "Orders", xmlHeading, lErrNum, sErrDesc, sErrSource
    oDOM.loadXML xmlHeading
    oDOM.documentElement.appendChild oDOM.createElement("rs:data")
    
    'Get the state of the Customer's Orders
    For Each oOrder In myOrders
        Set cust = New MSXML2.DOMDocument
        If oOrder.getSate(xmlOrders, False, lErrNum, sErrDesc, sErrSource) Then
            cust.loadXML xmlOrders
        End If
        
        oDOM.documentElement.selectSingleNode("rs:data").appendChild cust.documentElement.selectSingleNode("rs:data").selectSingleNode("z:row")
    Next oOrder
    
    xmlOrders = oDOM.xml
    
    getSate = True

    Exit Function

getSateErr:
    If lErrNum = 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "COrder.getSate"
    End If
End Function

