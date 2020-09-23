VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'========================================================================
'CAccess version 2.0
'Data Access Services Library
'© Argen 2002
'========================================================================
'DATE: December 20th 2003
'TIME: 1:55 AM
'========================================================================

Private moConn As ADODB.Connection
Private mrsTemp As ADODB.Recordset

Public Function connect(ByVal sConn As String, Optional lErrNum As Long, Optional sErrDesc As String, Optional sErrSource As String) As Boolean
    On Error GoTo connectErr
    
    connect = False
    
    If moConn.State = adStateClosed Then
        moConn.Open sConn
    End If
    
    connect = True
    
    Exit Function
    
connectErr:
    If lErrNum = 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "CAccess.connect"
    End If
End Function

Public Function retrieveData(ByVal sQuery As String, ByRef sXML As String, Optional lErrNum As Long, Optional sErrDesc As String, Optional sErrSource As String) As Boolean
    On Error GoTo retrieveDataErr
    
    Dim oDOM As New MSXML2.DOMDocument
    
    retrieveData = False
    
    'Query the db in order to fill in the rs
    Set mrsTemp = moConn.Execute(sQuery)
    'Save the data of the rs into a XML DOMDocument
    mrsTemp.Save oDOM, adPersistXML
    
    'Give back the XML String
    sXML = oDOM.xml
    
    Set oDOM = Nothing
    mrsTemp.Close
    Set mrsTemp = Nothing
    
    retrieveData = True
    
    Exit Function
    
retrieveDataErr:
    If lErrNum = 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "CAccess.retrieveData"
    End If
End Function

Private Sub Class_Initialize()
    Set moConn = New ADODB.Connection
    moConn.CursorLocation = adUseServer
End Sub

Private Sub Class_Terminate()
    Set moConn = Nothing
End Sub


Public Function beginTrans(Optional lErrNum As Long, Optional sErrDesc As String, Optional sErrSource As String) As Boolean
    On Error GoTo beginTransErr
    
    beginTrans = False
    
    beginTrans = True
    
    Exit Function
    
beginTransErr:
    If lErrNum = 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "CAccess.beginTrans"
    End If
End Function

Public Function closeConnection(Optional lErrNum As Long, Optional sErrDesc As String, Optional sErrSource As String) As Boolean
    On Error GoTo closeConnectionErr
    
    closeConnection = False
    
    If moConn.State = adStateOpen Then
        moConn.Close
    End If
    
    closeConnection = True
    
    Exit Function
    
closeConnectionErr:
    If lErrNum = 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "CAccess.closeConnection"
    End If
End Function

Public Function commitTrans(Optional lErrNum As Long, Optional sErrDesc As String, Optional sErrSource As String) As Boolean
    On Error GoTo commitTransErr
    
    commitTrans = False
    
    moConn.commitTrans
    
    commitTrans = True
    
    Exit Function
    
commitTransErr:
    If lErrNum = 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "CAccess.commitTrans"
    End If
End Function

Public Function query(ByVal sQuery As String, Optional ByVal lErrNum As Long, Optional sErrDesc As String, Optional sErrSource As String) As Boolean
    On Error GoTo queryErr
    
    query = False
    
    moConn.Execute sQuery
    query = True
    Exit Function
    
queryErr:
    If lErrNum = 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "CAccess.query"
    End If
End Function

Public Function rollbackTrans(Optional lErrNum As Long, Optional sErrDesc As String, Optional sErrSource As String) As Boolean
    On Error GoTo rollbackTransErr
    
    rollbackTrans = False
    
    moConn.rollbackTrans
    
    If Not closeConnection(lErrNum, sErrDesc, sErrSource) Then
        Exit Function
    End If
    
    rollbackTrans = True
    
    Exit Function
    
rollbackTransErr:
    If Err.Number = -2147168242 Then
        lErrNum = 0
        sErrDesc = ""
        sErrSource = ""
    Else
        If lErrNum = 0 Then
            lErrNum = Err.Number
            sErrDesc = Err.Description
            sErrSource = "CAccess.rollbackTrans"
        End If
    End If
End Function

Public Function getXMLHeading(ByVal strTableName As String, ByRef xmlHeading As String, Optional lErrNum As Long, Optional sErrDesc As String, Optional sErrSource As String) As Boolean
    On Error GoTo getXMLHeadingErr
    
    Dim oDOM As New MSXML2.DOMDocument
    Dim intPlace As Variant
    Dim strXML As String
    
    'Query the db in order to fill in the rs
    Set mrsTemp = moConn.Execute("select * from " & strTableName)
    'Save the data of the rs into a XML DOMDocument
    mrsTemp.Save oDOM, adPersistXML
    
    'Parse the xml
    strXML = oDOM.xml
    intPlace = InStr(1, strXML, "</s:Schema>")
    xmlHeading = Left(strXML, intPlace + 10) & "</xml>"
    
    Set oDOM = Nothing
    mrsTemp.Close
    Set mrsTemp = Nothing
    
    getXMLHeading = True
    
    Exit Function
    
getXMLHeadingErr:
    If lErrNum = 0 Then
        lErrNum = Err.Number
        sErrDesc = Err.Description
        sErrSource = "CAccess.rollbackTrans"
    End If
End Function


