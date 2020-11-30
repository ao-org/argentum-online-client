Attribute VB_Name = "ModEncrypt"

Public Function SEncriptar(ByVal Cadena As String) As String
    
    On Error GoTo SEncriptar_Err
    

    ' GSZ-AO - Encripta una cadena de texto
    Dim i As Long, RandomNum As Integer
    
    RandomNum = 99 * Rnd

    If RandomNum < 10 Then RandomNum = 10

    For i = 1 To Len(Cadena)
        Mid$(Cadena, i, 1) = Chr$(Asc(mid$(Cadena, i, 1)) + RandomNum)
    Next i

    SEncriptar = Cadena & Chr$(Asc(Left$(RandomNum, 1)) + 10) & Chr$(Asc(Right$(RandomNum, 1)) + 10)
    DoEvents

    
    Exit Function

SEncriptar_Err:
    Call RegistrarError(Err.number, Err.Description, "ModEncrypt.SEncriptar", Erl)
    Resume Next
    
End Function

' GSZAO - Encriptación basica y rapida para Strings
Public Function RndCrypt(ByVal str As String, ByVal Password As String) As String
    
    On Error GoTo RndCrypt_Err
    

    '  Made by Michael Ciurescu
    ' (CVMichael from vbforums.com)
    '  Original thread: http://www.vbforums.com/showthread.php?t=231798
    Dim SK As Long, k As Long

    Rnd -1
    Randomize Len(Password)

    For k = 1 To Len(Password)
        SK = SK + (((k Mod 256) Xor Asc(mid$(Password, k, 1))) Xor Fix(256 * Rnd))
    Next k

    Rnd -1
    Randomize SK
    
    For k = 1 To Len(str)
        Mid$(str, k, 1) = Chr(Fix(256 * Rnd) Xor Asc(mid$(str, k, 1)))
    Next k
    
    RndCrypt = str

    
    Exit Function

RndCrypt_Err:
    Call RegistrarError(Err.number, Err.Description, "ModEncrypt.RndCrypt", Erl)
    Resume Next
    
End Function

