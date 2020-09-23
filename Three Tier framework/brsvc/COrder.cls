VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Clase de datos de la tabla Orders"
Option Explicit

Private SQL As String
Private sXML As String
Private mDateOrder As Date
Private mIDOrder As Long
'State variables
Private mIsNew As Boolean
Private mIsDirty As Boolean
Private mIsDeleted As Boolean

Public mcolOrders As New Collection

'Reference to the Order Customer
Public myCustomer As New CCustomer

Property Get DateOrder() As Date
    DateOrder = mDateOrder
End Property

Property Let DateOrder(Valor As Date)
    mDateOrder = Valor
End Property

Property Get IDOrder() As Long
    IDOrder = mIDOrder
End Property

Property Let IDOrder(Valor As Long)
    mIDOrder = Valor
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

Public Function read(ByVal IDOrder As Long, bExpand As Boolean, Optional lErrNum As Long, Optional sDescErr As String, Optional sErrSource As String) As Boolean
    On Error GoTo Errread

    SQL = "SELECT * FROM Orders WHERE IDOrder=" & IDOrder

    If ogAccess.retrieveData(SQL, sXML, lErrNum, sDescErr, sErrSource) Then
        If unPack(sXML, lErrNum, sDescErr, sErrSource) Then
            read = True
            If bExpand Then
            End If
        End If
    End If

    Exit Function

Errread:
    If lErrNum = 0 Then
        lErrNum = Err.Number
        sDescErr = Err.Description
        sErrSource = "COrder.read"
    End If
End Function

Private Function unPack(sXML As String, Optional lErrNum As Long, Optional sDescErr As String, Optional sErrSource As String) As Boolean
    On Error GoTo ErrunPack

    Dim oDOMDocument As New MSXML2.DOMDocument
    Dim oRoot As MSXML2.IXMLDOMElement 'Estos son todos los datos
    Dim oData As MSXML2.IXMLDOMElement 'Este objeto contiene los datos
    Dim oRow As MSXML2.IXMLDOMElement 'Este contiene las filas

    If oDOMDocument.loadXML(sXML) Then
        Set oRoot = oDOMDocument.documentElement
        Set oData = oRoot.childNodes(1)
        Set oRow = oData.childNodes.Item(0)
        
        mDateOrder = oRow.Attributes(giDateOrderOrders).nodeValue
        myCustomer.IDCustomer = oRow.Attributes(giIDCustomerOrders).nodeValue
        mIDOrder = oRow.Attributes(giIDOrdersOrders).nodeValue
        
        unPack = True

        Set oRoot = Nothing
        Set oData = Nothing
        Set oDOMDocument = Nothing
        Set oRow = Nothing
    End If

    Exit Function

ErrunPack:
    If lErrNum = 0 Then
        lErrNum = Err.Number
        sDescErr = Err.Description
        sErrSource = "COrder.unPack"
    End If
End Function

Public Function unPackDep(sXML As String, Optional lErrNum As Long, Optional sDescErr As String, Optional sErrSource As String) As Boolean
    On Error GoTo ErrunPackDep

    Dim iIndice As Integer
    Dim oOrder As COrder

    Dim oDOM As New DOMDocument
    Dim node As MSXML2.IXMLDOMElement

    unPackDep = False

    If oDOM.loadXML(sXML) Then
        For Each node In oDOM.documentElement.selectSingleNode("rs:data").childNodes
            Set oOrder = New COrder
            With oOrder
                .DateOrder = validateDateTime(node.Attributes(giDateOrderOrders).Text)
                .myCustomer.IDCustomer = node.Attributes(giIDCustomerOrders).Text
                .IDOrder = node.Attributes(giIDOrdersOrders).Text
            End With
        
            mcolOrders.Add oOrder, "ORD-" & oOrder.IDOrder
        Next node
    End If
    
    unPackDep = True

    Exit Function

ErrunPackDep:
    If lErrNum = 0 Then
        lErrNum = Err.Number
        sDescErr = Err.Description
        sErrSource = "COrder.unPackDep"
    End If
End Function

Public Function Save(Optional lErrNum As Long, Optional sDescErr As String, Optional sErrSource As String) As Boolean
    On Error GoTo SaveErr
    
    Save = False
    
    If mIsNew Then
        SQL = "INSERT INTO Orders (DateOrder, IDCustomer, IDOrder) VALUES (" & "#" & Format(DateOrder, "mm/dd/yyyy hh:mm") & "#, " & myCustomer.IDCustomer & ", " & IDOrder & ")"
        mIsNew = False
        Save = True
    Else
        If IsDirty Then
            SQL = "UPDATE Orders SET " & "DateOrder=#" & Format(DateOrder, "mm/dd/yyyy  hh:mm") & "#, IDCustomer=" & myCustomer.IDCustomer & ", IDOrder=" & IDOrder & " WHERE IDOrder=" & IDOrder
            IsDirty = False
            Save = True
        Else
            If IsDeleted Then
                SQL = "DELETE FROM Orders WHERE IDOrder=" & IDOrder
                IsDeleted = False
                Save = True
            End If
        End If
    End If

    If Save Then
        If ogAccess.query(SQL, lErrNum, sDescErr, sErrSource) Then
            Save = True
        End If
    End If

    Exit Function

SaveErr:
    If lErrNum = 0 Then
        lErrNum = Err.Number
        sDescErr = Err.Description
        sErrSource = "COrder.Save"
    End If
End Function

Public Function readByCustomer(Optional lErrNum As Long, Optional sDescErr As String, Optional sErrSource As String) As Boolean
    On Error GoTo Errread

    SQL = "SELECT * FROM Orders WHERE IDCustomer=" & myCustomer.IDCustomer

    If ogAccess.retrieveData(SQL, sXML, lErrNum, sDescErr, sErrSource) Then
        If unPackDep(sXML, lErrNum, sDescErr, sErrSource) Then
            readByCustomer = True
        End If
    End If

    Exit Function

Errread:
    If lErrNum = 0 Then
        lErrNum = Err.Number
        sDescErr = Err.Description
        sErrSource = "COrder.readByCustomer"
    End If
End Function

'Convert the state of the object into XML
Public Function getSate(xmlOrders As String, Optional bMultiple As Boolean, Optional lErrNum As Long, Optional sErrDesc As String, Optional sErrSource As String) As Boolean
    On Error GoTo getSateErr

    Dim oOrder As COrder
    Dim cust As New MSXML2.DOMDocument
    Dim node As MSXML2.IXMLDOMElement
    Dim oDOM As New MSXML2.DOMDocument
    Dim attr As MSXML2.IXMLDOMAttribute
    
    Dim xmlHeading As String
    
    'This is the heading that we need for opening recordsets with XML
    ogAccess.getXMLHeading "Orders", xmlHeading
    
    oDOM.loadXML xmlHeading

    oDOM.documentElement.appendChild oDOM.createElement("rs:data")
    
    'Multiple tells me that I need to get the xml string of many Customers
    'When I'm asking for a particular Customer, Multiple=False
    If Not bMultiple Then
        Set cust = New MSXML2.DOMDocument

        cust.appendChild cust.createElement("z:row")
        Set attr = cust.createAttribute("IDOrder")
        attr.Value = mIDOrder
        Set attr = cust.selectSingleNode("z:row").Attributes.setNamedItem(attr)
        
        Set attr = cust.createAttribute("IDCustomer")
        attr.Value = myCustomer.IDCustomer
        Set attr = cust.selectSingleNode("z:row").Attributes.setNamedItem(attr)
        
        Set attr = cust.createAttribute("DateOrder")
        attr.Value = mDateOrder
        Set attr = cust.selectSingleNode("z:row").Attributes.setNamedItem(attr)
        
        oDOM.documentElement.selectSingleNode("rs:data").appendChild cust.documentElement
    Else
        For Each oOrder In mcolOrders
            Set cust = New MSXML2.DOMDocument

            cust.appendChild cust.createElement("z:row")
            Set attr = cust.createAttribute("IDOrder")
            attr.Value = oOrder.IDOrder
            Set attr = cust.selectSingleNode("z:row").Attributes.setNamedItem(attr)
            
            Set attr = cust.createAttribute("IDCustomer")
            attr.Value = oOrder.myCustomer.IDCustomer
            Set attr = cust.selectSingleNode("z:row").Attributes.setNamedItem(attr)
            
            Set attr = cust.createAttribute("DateOrder")
            attr.Value = oOrder.DateOrder
            Set attr = cust.selectSingleNode("z:row").Attributes.setNamedItem(attr)
            
            oDOM.documentElement.selectSingleNode("rs:data").appendChild cust.documentElement
        Next oOrder
    End If

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


