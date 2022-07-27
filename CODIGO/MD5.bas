Attribute VB_Name = "MD5"
Option Explicit

Private Declare Sub MDFile Lib "aamd532.dll" (ByVal f As String, ByVal r As String)
Private Declare Sub MDStringFix Lib "aamd532.dll" (ByVal f As String, ByVal t As Long, ByVal r As String)

Public Function MD5String(ByVal P As String) As String
    Dim r As String * 32, t As Long
    r = Space(32)
    t = Len(P)
    MDStringFix P, t, r
    MD5String = r
End Function

Public Function MD5File(ByVal f As String) As String
    Dim r As String * 32
    r = Space(32)
    MDFile f, r
    MD5File = r
End Function

Public Function hexMd52Asc(ByVal MD5 As String) As String
    Dim i As Long
    Dim L As String
    
    If Len(MD5) And &H1 Then MD5 = "0" & MD5
    
    For i = 1 To Len(MD5) \ 2
        L = mid$(MD5, (2 * i) - 1, 2)
        hexMd52Asc = hexMd52Asc & Chr$(hexHex2Dec(L))
    Next i
End Function

Public Function hexHex2Dec(ByVal hex As String) As Long
    hexHex2Dec = Val("&H" & hex)
End Function

Public Function txtOffset(ByVal Text As String, ByVal off As Integer) As String
    Dim i As Long
    Dim L As String
    
    For i = 1 To Len(Text)
        L = mid$(Text, i, 1)
        txtOffset = txtOffset & Chr$((Asc(L) + off) And &HFF)
    Next i
End Function


