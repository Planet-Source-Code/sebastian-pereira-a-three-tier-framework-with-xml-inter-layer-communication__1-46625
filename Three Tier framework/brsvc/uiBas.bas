Attribute VB_Name = "uiBas"
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const LB_FINDSTRING = &H18F

Public Enum ListBoxPosOptions
    opcItemData = 0
    opcListValue
End Enum

Public Sub Main()
    frm.Show
End Sub

Public Sub ImprimirRecordset(rs As ADODB.Recordset)
  Dim n As Integer, i As Integer
    
    With rs
        If Not (.BOF And .EOF) Then
            .MoveFirst
            For n = 0 To .RecordCount - 1
                For i = 0 To .Fields.Count - 1
                    Debug.Print .Fields(i).Name & ": " & .Fields(i).value
                Next i
                Debug.Print "============================================="
                .MoveNext
            Next n
            Debug.Print " "
            Debug.Print "**********************************************************"
            Debug.Print "CANTIDAD DE REGISTROS: " & .RecordCount
            .MoveFirst
          Else
            Debug.Print "RECORDSET VACIO"
        End If
    End With
End Sub

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

Public Function seekItemListBox(list As ListBox, ByVal value As Variant, optionSeek As ListBoxPosOptions) As Long
    Dim i As Integer
    
    If list.ListCount = 0 Then
        seekItemListBox = -1
        Exit Function
    End If
    Select Case optionSeek
        Case opcItemData
            For i = 0 To list.ListCount - 1
                If list.ItemData(i) = value Then
                    seekItemListBox = i
                    Exit Function
                Else
                    seekItemListBox = -1
                End If
                
            Next i
        Case opcListValue
            seekItemListBox = SendMessage(list.hwnd, LB_FINDSTRING, -1, ByVal CStr(value))
        Case Else
            MsgBox "I don't know that option!", vbCritical, ""
            Exit Function
    End Select
End Function

