Attribute VB_Name = "modUnicodeText"
Option Compare Database
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function MessageBoxW Lib "user32" ( _
    ByVal hwnd As LongPtr, _
    ByVal lpText As LongPtr, _
    ByVal lpCaption As LongPtr, _
    ByVal uType As Long) As Long
#Else
Private Declare Function MessageBoxW Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal lpText As Long, _
    ByVal lpCaption As Long, _
    ByVal uType As Long) As Long
#End If

' Returns a Unicode string from a list of hexadecimal code points.
' Example: U("0633 0644 0627 0645")
Public Function U(ByVal hexCodes As String) As String
    Dim parts() As String
    parts = Split(Trim$(hexCodes), " ")

    Dim i As Long
    Dim result As String
    For i = LBound(parts) To UBound(parts)
        If Len(parts(i)) > 0 Then
            result = result & ChrW$(CLng("&H" & parts(i)))
        End If
    Next i

    U = result
End Function

Public Function MsgBoxU(ByVal prompt As String, Optional ByVal buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal title As String = "") As VbMsgBoxResult
    Dim captionText As String
    If Len(title) > 0 Then
        captionText = title
    Else
        captionText = Application.Name
    End If

#If VBA7 Then
    MsgBoxU = MessageBoxW(CLngPtr(0), StrPtr(prompt), StrPtr(captionText), CLng(buttons))
#Else
    MsgBoxU = MessageBoxW(0&, StrPtr(prompt), StrPtr(captionText), CLng(buttons))
#End If
End Function
