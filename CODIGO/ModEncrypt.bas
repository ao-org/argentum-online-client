Attribute VB_Name = "ModEncrypt"
'    Argentum 20 - Game Client Program
'    Copyright (C) 2022 - Noland Studios
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'
Public Function SEncriptar(ByVal Cadena As String) As String
    
    On Error GoTo SEncriptar_Err
    
    SEncriptar = AO20CryptoSysWrapper.Encrypt("7061626C6F6D61727175657A41524731", Cadena)
    
    DoEvents

    
    Exit Function

SEncriptar_Err:

    Call MsgBox(Err.Description)
    Call RegistrarError(Err.Number, Err.Description, "ModEncrypt.SEncriptar", Erl)
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
    Call RegistrarError(Err.Number, Err.Description, "ModEncrypt.RndCrypt", Erl)
    Resume Next
    
End Function

