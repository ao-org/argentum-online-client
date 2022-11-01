Attribute VB_Name = "modLocale"
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
Option Explicit

Private arrLocale_SMG() As String

Public Type Tdestino

    CityDest As Byte
    costo As Long

End Type

Public Destinos() As Tdestino

Public Function Locale_Parse_ServerMessage(ByVal bytHeader As Integer, Optional ByVal strExtra As String = vbNullString) As String

    On Error GoTo ErrorHandler
    
    Dim Fields() As String

    Dim strLocale As String

    Dim i As Long
    
    strLocale = Locale_SMG(bytHeader)

    If LenB(strExtra) = 0 Then
        Locale_Parse_ServerMessage = strLocale
        Exit Function
    End If
    
    Fields = Split(strExtra, "¬")

    ' En reversa para evitar pisar campos mayores a 10
    For i = UBound(Fields) To 0 Step -1
        strLocale = Replace(strLocale, "¬" & (i + 1), Fields(i))
    Next

ErrorHandler:
    Locale_Parse_ServerMessage = strLocale

End Function

Public Function General_Get_Line_Count(ByVal FileName As String) As Long

    '**************************************************************
    'Author: Augusto José Rando
    'Last Modify Date: 6/11/2005
    '
    '**************************************************************
    On Error GoTo ErrorHandler

    Dim N As Integer, tmpStr As String

    If LenB(FileName) Then
        N = FreeFile()
    
        Open FileName For Input As #N
    
        Do While Not EOF(N)
            General_Get_Line_Count = General_Get_Line_Count + 1
            Line Input #N, tmpStr
        Loop
    
        Close N

    End If

    Exit Function

ErrorHandler:

End Function

Public Function Integer_To_String(ByVal Var As Integer) As String
    
    On Error GoTo Integer_To_String_Err
    

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 3/12/2005
    '
    '**************************************************************
    Dim temp As String
        
    'Convertimos a hexa
    temp = hex$(Var)
    
    'Nos aseguramos tenga 4 bytes de largo
    While Len(temp) < 4

        temp = "0" & temp
    Wend
    
    'Convertimos a string
    Integer_To_String = Chr$(Val("&H" & Left$(temp, 2))) & Chr$(Val("&H" & Right$(temp, 2)))
    Exit Function

ErrorHandler:

    
    Exit Function

Integer_To_String_Err:
    Call RegistrarError(Err.number, Err.Description, "modLocale.Integer_To_String", Erl)
    Resume Next
    
End Function

Public Function String_To_Integer(ByRef str As String, ByVal Start As Integer) As Integer

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 3/12/2005
    '
    '**************************************************************
    On Error GoTo Error_Handler
    
    Dim temp_str As String
    
    'Asergurarse sea válido
    If Len(str) < Start - 1 Or Len(str) = 0 Then Exit Function
    
    'Convertimos a hexa el valor ascii del segundo byte
    temp_str = hex$(Asc(mid$(str, Start + 1, 1)))
    
    'Nos aseguramos tenga 2 bytes (los ceros a la izquierda cuentan por ser el segundo byte)
    While Len(temp_str) < 2

        temp_str = "0" & temp_str
    Wend
    
    'Convertimos a integer
    String_To_Integer = Val("&H" & hex$(Asc(mid$(str, Start, 1))) & temp_str)
            
    Exit Function
        
Error_Handler:
        
End Function

Public Function Byte_To_String(ByVal Var As Byte) As String
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 3/12/2005
    'Convierte un byte a string
    '**************************************************************
    
    On Error GoTo Byte_To_String_Err
    
    Byte_To_String = Chr$(Val("&H" & hex$(Var)))
    Exit Function

ErrorHandler:

    
    Exit Function

Byte_To_String_Err:
    Call RegistrarError(Err.number, Err.Description, "modLocale.Byte_To_String", Erl)
    Resume Next
    
End Function

Public Function String_To_Byte(ByRef str As String, ByVal Start As Integer) As Byte

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 3/12/2005
    '
    '**************************************************************
    On Error GoTo Error_Handler
    
    If Len(str) < Start Then Exit Function
    
    String_To_Byte = Asc(mid$(str, Start, 1))
    
    Exit Function
        
Error_Handler:

End Function

Public Function Long_To_String(ByVal Var As Long) As String
    
    On Error GoTo Long_To_String_Err
    

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 3/12/2005
    '
    '**************************************************************
    'No aceptamos valores que usen los 4 últimos its
    If Var > &HFFFFFFF Then GoTo ErrorHandler
    
    Dim temp As String
    
    'Vemos si el cuarto byte es cero
    If (Var And &HFF&) = 0 Then Var = Var Or &H80000001
    
    'Vemos si el tercer byte es cero
    If (Var And &HFF00&) = 0 Then Var = Var Or &H40000100
    
    'Vemos si el segundo byte es cero
    If (Var And &HFF0000) = 0 Then Var = Var Or &H20010000
    
    'Vemos si el primer byte es cero
    If Var < &H1000000 Then Var = Var Or &H10000000
    
    'Convertimos a hexa
    temp = hex$(Var)
    
    'Nos aseguramos tenga 8 bytes de largo
    While Len(temp) < 8

        temp = "0" & temp
    Wend
    
    'Convertimos a string
    Long_To_String = Chr$(Val("&H" & Left$(temp, 2))) & Chr$(Val("&H" & mid$(temp, 3, 2))) & Chr$(Val("&H" & mid$(temp, 5, 2))) & Chr$(Val("&H" & mid$(temp, 7, 2)))
    Exit Function

ErrorHandler:

    
    Exit Function

Long_To_String_Err:
    Call RegistrarError(Err.number, Err.Description, "modLocale.Long_To_String", Erl)
    Resume Next
    
End Function

Public Function String_To_Long(ByRef str As String, ByVal Start As Integer) As Long
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 3/12/2005
    '
    '**************************************************************
    
    On Error GoTo ErrorHandler
        
    If Len(str) < Start - 3 Then Exit Function
    
    Dim temp_str  As String

    Dim temp_str2 As String

    Dim temp_str3 As String
    
    'Tomamos los últimos 3 bytes y convertimos sus valroes ASCII a hexa
    temp_str = hex$(Asc(mid$(str, Start + 1, 1)))
    temp_str2 = hex$(Asc(mid$(str, Start + 2, 1)))
    temp_str3 = hex$(Asc(mid$(str, Start + 3, 1)))
    
    'Nos aseguramos todos midan 2 bytes (los ceros a la izquierda cuentan por ser bytes 2, 3 y 4)
    While Len(temp_str) < 2

        temp_str = "0" & temp_str
    Wend
    
    While Len(temp_str2) < 2

        temp_str2 = "0" & temp_str2
    Wend
    
    While Len(temp_str3) < 2

        temp_str3 = "0" & temp_str3
    Wend
    
    'Convertimos a una única cadena hexa
    String_To_Long = Val("&H" & hex$(Asc(mid$(str, Start, 1))) & temp_str & temp_str2 & temp_str3)
    
    'Si el cuarto byte era cero
    If String_To_Long And &H80000000 Then String_To_Long = String_To_Long Xor &H80000001
    
    'Si el tercer byte era cero
    If String_To_Long And &H40000000 Then String_To_Long = String_To_Long Xor &H40000100
    
    'Si el segundo byte era cero
    If String_To_Long And &H20000000 Then String_To_Long = String_To_Long Xor &H20010000
    
    'Si el primer byte era cero
    If String_To_Long And &H10000000 Then String_To_Long = String_To_Long Xor &H10000000
        
    Exit Function
        
ErrorHandler:

End Function

