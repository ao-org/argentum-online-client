Attribute VB_Name = "modAuth"
Public Function ByteArr2String(ByRef arr() As Byte) As String
    
    Dim str As String
    Dim i As Long
    For i = 0 To UBound(arr)
        str = str + Chr(arr(i))
    Next i
    
    ByteArr2String = str
    
End Function


'HarThaoS: Convierto el str en arr() bytes
Public Function Str2ByteArr(ByVal str As String, ByRef arr() As Byte, Optional ByVal length As Long = 0)
    Dim i As Long
    Dim asd As String
    If length = 0 Then
        ReDim arr(0 To (Len(str) - 1))
        For i = 0 To (Len(str) - 1)
            arr(i) = Asc(mid(str, i + 1, 1))
        Next i
    Else
        ReDim arr(0 To (length - 1)) As Byte
        For i = 0 To (length - 1)
            arr(i) = Asc(mid(str, i + 1, 1))
        Next i
    End If
    
End Function
