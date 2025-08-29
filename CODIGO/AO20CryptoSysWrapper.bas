Attribute VB_Name = "AO20CryptoSysWrapper"
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
Public Function Encrypt(ByVal hex_key As String, ByVal plain_text As String) As String
    On Error Goto Encrypt_Err
    Dim iv() As Byte
    Dim key() As Byte
    Dim plain_text_byte() As Byte
    
    Dim algstr As String
    algstr = "Aes128/CFB/nopad"
    key = cnvBytesFromHexStr(hex_key)
    iv = key
    
    ' "Now is the time for all good men to"
    
    plain_text = cnvHexStrFromString(plain_text)
    plain_text_byte = cnvBytesFromHexStr(plain_text)
    Encrypt = cnvToBase64(cipherEncryptBytes2(plain_text_byte, key, iv, algstr))
   
    Exit Function
Encrypt_Err:
    Call TraceError(Err.Number, Err.Description, "AO20CryptoSysWrapper.Encrypt", Erl)
End Function


Public Function Decrypt(ByVal hex_key As String, ByVal encrypted_text_b64 As String) As String
    On Error Goto Decrypt_Err
    Dim iv() As Byte
    Dim key() As Byte
    Dim encrypted_text_byte() As Byte
    Dim decrypted_text() As Byte
    Dim encrypted_text_hex As String
    Dim algstr As String
    algstr = "Aes128/CFB/nopad"
    key = cnvBytesFromHexStr(hex_key)
    iv = key
    
    ' "Now is the time for all good men to"
    
    encrypted_text_byte = cnvFromBase64(encrypted_text_b64)
    encrypted_text_hex = cnvToHex(encrypted_text_byte)
    encrypted_text_byte = cnvBytesFromHexStr(encrypted_text_hex)
    Decrypt = cnvStringFromHexStr(cnvToHex(cipherDecryptBytes2(encrypted_text_byte, key, iv, algstr)))
   
    Exit Function
Decrypt_Err:
    Call TraceError(Err.Number, Err.Description, "AO20CryptoSysWrapper.Decrypt", Erl)
End Function


Public Function Str2ByteArr(ByVal str As String, ByRef arr() As Byte, Optional ByVal Length As Long = 0) 'Convierto el str en arr() bytes
    On Error Goto Str2ByteArr_Err
    Dim i As Long
    Dim asd As String
    If length = 0 Then
        ReDim arr(0 To (Len(str) - 1))
        For i = 0 To (Len(str) - 1)
            arr(i) = Asc(Mid(str, i + 1, 1))
        Next i
    Else
        ReDim arr(0 To (length - 1)) As Byte
        For i = 0 To (length - 1)
            arr(i) = Asc(Mid(str, i + 1, 1))
        Next i
    End If
    
    Exit Function
Str2ByteArr_Err:
    Call TraceError(Err.Number, Err.Description, "AO20CryptoSysWrapper.Str2ByteArr", Erl)
End Function

Public Function ByteArr2String(ByRef arr() As Byte) As String
    On Error Goto ByteArr2String_Err
    
    Dim str As String
    Dim i As Long
    For i = 0 To UBound(arr)
        str = str + Chr(arr(i))
    Next i
    
    ByteArr2String = str
    
    Exit Function
ByteArr2String_Err:
    Call TraceError(Err.Number, Err.Description, "AO20CryptoSysWrapper.ByteArr2String", Erl)
End Function

Public Function hiByte(ByVal w As Integer) As Byte
    On Error Goto hiByte_Err
    Dim hi As Integer
    If w And &H8000 Then hi = &H4000
    
    hiByte = (w And &H7FFE) \ 256
    hiByte = (hiByte Or (hi \ 128))
    
    Exit Function
hiByte_Err:
    Call TraceError(Err.Number, Err.Description, "AO20CryptoSysWrapper.hiByte", Erl)
End Function

Public Function LoByte(w As Integer) As Byte
    On Error Goto LoByte_Err
 LoByte = w And &HFF
    Exit Function
LoByte_Err:
    Call TraceError(Err.Number, Err.Description, "AO20CryptoSysWrapper.LoByte", Erl)
End Function

Public Function MakeInt(ByVal LoByte As Byte, _
    On Error Goto MakeInt_Err
   ByVal hiByte As Byte) As Integer

MakeInt = ((hiByte * &H100) + LoByte)

    Exit Function
MakeInt_Err:
    Call TraceError(Err.Number, Err.Description, "AO20CryptoSysWrapper.MakeInt", Erl)
End Function

Public Function CopyBytes(ByRef src() As Byte, ByRef dst() As Byte, ByVal size As Long, Optional ByVal offset As Long = 0)
    On Error Goto CopyBytes_Err
    Dim i As Long
    
    For i = 0 To (size - 1)
        dst(i + offset) = src(i)
    Next i
    
    Exit Function
CopyBytes_Err:
    Call TraceError(Err.Number, Err.Description, "AO20CryptoSysWrapper.CopyBytes", Erl)
End Function

Public Function ByteArrayToHex(ByRef ByteArray() As Byte) As String
    On Error Goto ByteArrayToHex_Err
    Dim l As Long, strRet As String
    
    For l = LBound(ByteArray) To UBound(ByteArray)
        strRet = strRet & hex$(ByteArray(l)) & " "
    Next l
    
    'Remove last space at end.
    ByteArrayToHex = Left$(strRet, Len(strRet) - 1)
    Exit Function
ByteArrayToHex_Err:
    Call TraceError(Err.Number, Err.Description, "AO20CryptoSysWrapper.ByteArrayToHex", Erl)
End Function

Public Function ByteArrayToDecimalString(ByRef oMsg() As Byte) As String
    On Error Goto ByteArrayToDecimalString_Err
    Dim i As Long
    Dim result As String
    
      
    ' Loop through the byte array and build the string
    For i = LBound(oMsg) To UBound(oMsg)
        result = result & oMsg(i) & " "
    Next i
    
    ' Trim trailing space and return result
    ByteArrayToDecimalString = Trim(result)
    Exit Function
ByteArrayToDecimalString_Err:
    Call TraceError(Err.Number, Err.Description, "AO20CryptoSysWrapper.ByteArrayToDecimalString", Erl)
End Function


