Attribute VB_Name = "brBas"
Option Explicit

'HERE YOU HAVE TO CHANGE THE DATA BASE PATH
Public Const msDSN As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\Habilitación Profesional\Investigación\XML\bd.mdb;Persist Security Info=False"

Public ogAccess As New CAccess

Public Function validateDateTime(xmlDateTime As String) As Date
    Dim Pos As Long
    Dim dDate As String
    Dim dTime As String
    
    Pos = InStr(1, xmlDateTime, "T", vbTextCompare)
    dDate = Mid(xmlDateTime, 1, Pos - 1)
    
    Pos = InStr(1, xmlDateTime, "T", vbTextCompare)
    dTime = Mid(xmlDateTime, Pos + 1, Len(xmlDateTime))
    
    validateDateTime = CDate(dDate & " " & dTime)
End Function

