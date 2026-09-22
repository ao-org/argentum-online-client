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
    CityDest As String
    costo As Long
End Type

Public Destinos() As Tdestino

Public Function Locale_Parse_ServerMessage(ByVal bytHeader As Integer, Optional ByVal strExtra As String = vbNullString) As String
    On Error GoTo Locale_Parse_ServerMessage_Err
    Dim Fields()  As String
    Dim strLocale As String
    Dim i         As Long
    If bytHeader = 2266 Then
        If Left$(strExtra, 13) = "CHAR_INVALID|" Then
            Locale_Parse_ServerMessage = Locale_Parse_InvalidCharacter(strExtra)
        Else
            Locale_Parse_ServerMessage = strExtra
        End If
        Exit Function
    End If
    strLocale = Locale_SMG(bytHeader)
    ' Manejo del caso especial del NPC
    Call HandleNpcName(bytHeader, strExtra)
    If LenB(strExtra) = 0 Then
        Locale_Parse_ServerMessage = strLocale
        Exit Function
    End If
    Fields = Split(strExtra, "¬")
    'Look At Tile when clicking an npc with or without owner
    If bytHeader = 1622 Or bytHeader = 1621 Then
        Call NpcInTileToTxtParser(Fields, bytHeader)
    End If
    'Look At Tile when clicking a player
    If bytHeader = 1105 Then
        Call UserInTileToTxtParser(Fields)
    End If
    If bytHeader = 473 Then
        Call SkillsNamesToTxtParser(Fields)
    End If
    If bytHeader = 1426 Then
        Call ClassesToTxtParser(Fields)
    End If
    If bytHeader = 1264 Then
        Call QuestsIndexToTxtParser(Fields)
    End If
    If bytHeader = 1988 Then
        Call ClassesToTxtParser(Fields)
        Call RaceToTxtParser(Fields)
    End If
    
    If bytHeader = 2122 Then
        Call ObjIndexToLocalizedName(Fields, 2)
    End If
    
    If bytHeader = 2128 Then
        Call ObjIndexToLocalizedName(Fields, 2)
    End If
    ' En reversa para evitar pisar campos mayores a 10
    For i = UBound(Fields) To 0 Step -1
        strLocale = Replace(strLocale, "¬" & (i + 1), Fields(i))
    Next
    Locale_Parse_ServerMessage = strLocale
    Exit Function
Locale_Parse_ServerMessage_Err:
    Call RegistrarError(Err.Number, Err.Description, "modLocale.Locale_Parse_ServerMessage", Erl)
    Resume Next
End Function

Private Function Locale_Parse_InvalidCharacter(ByVal payload As String) As String
    Dim sections() As String
    Dim sectionData() As String
    Dim sectionName As String
    Dim messageText As String
    Dim i As Long

    sections = Split(Mid$(payload, 14), "|")
    messageText = JsonLanguage.Item("MENSAJE_PERSONAJE_NO_VALIDO")

    For i = 0 To UBound(sections) - 1 Step 2
        sectionName = sections(i)
        sectionData = Split(sections(i + 1), ";")
        Select Case sectionName
            Case "INV"
                Call AppendInvalidCharacterObjects(messageText, sectionData, "MENSAJE_OBJETOS_INVENTARIO_FALTANTES")
            Case "BANK"
                Call AppendInvalidCharacterObjects(messageText, sectionData, "MENSAJE_OBJETOS_BANCO_FALTANTES")
            ' INV_SLOTS / BANK_SLOTS: ya quedan en el log del server, no se muestran al jugador.
        End Select
    Next i

    Locale_Parse_InvalidCharacter = messageText & vbCrLf & JsonLanguage.Item("MENSAJE_ENVIAR_TICKET_ADMIN")
End Function

Private Sub AppendInvalidCharacterObjects(ByRef messageText As String, ByRef objectData() As String, ByVal labelKey As String)
    Dim objectParts() As String
    Dim ids As String
    Dim i As Long

    For i = LBound(objectData) To UBound(objectData)
        If LenB(objectData(i)) Then
            objectParts = Split(objectData(i), ":")
            If UBound(objectParts) = 1 Then
                If LenB(ids) Then ids = ids & ", "
                ids = ids & objectParts(1)  ' solo el item_id, descartamos el slot
            End If
        End If
    Next i

    If LenB(ids) Then messageText = messageText & vbCrLf & JsonLanguage.Item(labelKey) & " " & ids
End Sub
Private Sub AppendInvalidCharacterSlots(ByRef messageText As String, ByRef slotData() As String, ByVal labelKey As String)
    Dim i As Long

    messageText = messageText & vbCrLf & JsonLanguage.Item(labelKey)
    For i = LBound(slotData) To UBound(slotData)
        If LenB(slotData(i)) Then
            messageText = messageText & vbCrLf & Replace(JsonLanguage.Item("MENSAJE_SLOT_INVALIDO_FORMATO"), "¬1", slotData(i))
        End If
    Next i
End Sub

' Manejar el nombre del NPC para los casos especiales
Private Sub HandleNpcName(ByVal bytHeader As Integer, ByRef strExtra As String)
    Dim NpcName           As String
    Dim specialNpcHeaders As Variant
    Dim i                 As Long
    ' Definir los IDs que requieren el nombre del NPC
    specialNpcHeaders = Array(1788) ' Agrega aquí otros IDs si es necesario
    ' Verificar si el ID está en la lista de IDs especiales
    For i = LBound(specialNpcHeaders) To UBound(specialNpcHeaders)
        If bytHeader = specialNpcHeaders(i) Then
            NpcName = NpcData(strExtra).Name ' Obtener el nombre del NPC
            If Len(NpcName) > 0 Then
                strExtra = NpcName
            End If
        End If
    Next
End Sub


Public Function General_Get_Line_Count(ByVal filename As String) As Long
    On Error GoTo ErrorHandler
    Dim n As Integer, tmpStr As String
    If LenB(filename) Then
        n = FreeFile()
        Open filename For Input As #n
        Do While Not EOF(n)
            General_Get_Line_Count = General_Get_Line_Count + 1
            Line Input #n, tmpStr
        Loop
        Close n
    End If
    Exit Function
ErrorHandler:
End Function

Public Function Integer_To_String(ByVal Var As Integer) As String
    On Error GoTo Integer_To_String_Err
    Dim temp As String
    'Convertimos a hexa
    temp = hex$(Var)
    'Nos aseguramos tenga 4 bytes de largo
    While Len(temp) < 4
        temp = "0" & temp
    Wend
    'Convertimos a string

    Integer_To_String = Chr$(val("&H" & Left$(temp, 2))) & Chr$(val("&H" & Right$(temp, 2)))
    Exit Function
ErrorHandler:
    Exit Function
Integer_To_String_Err:
    Call RegistrarError(Err.Number, Err.Description, "modLocale.Integer_To_String", Erl)
    Resume Next
End Function

Public Function String_To_Integer(ByRef str As String, ByVal start As Integer) As Integer
    On Error GoTo Error_Handler
    Dim temp_str As String
    'Asergurarse sea válido
    If Len(str) < start - 1 Or Len(str) = 0 Then Exit Function

    'Convertimos a hexa el valor ascii del segundo byte
    temp_str = hex$(Asc(mid$(str, start + 1, 1)))
    'Nos aseguramos tenga 2 bytes (los ceros a la izquierda cuentan por ser el segundo byte)
    While Len(temp_str) < 2
        temp_str = "0" & temp_str
    Wend
    'Convertimos a integer
    String_To_Integer = val("&H" & hex$(Asc(mid$(str, start, 1))) & temp_str)
    Exit Function
Error_Handler:
End Function

Public Function Byte_To_String(ByVal Var As Byte) As String
    'Convierte un byte a string
    On Error GoTo Byte_To_String_Err
    Byte_To_String = Chr$(val("&H" & hex$(Var)))
    Exit Function
ErrorHandler:
    Exit Function
Byte_To_String_Err:
    Call RegistrarError(Err.Number, Err.Description, "modLocale.Byte_To_String", Erl)
    Resume Next
End Function

Public Function String_To_Byte(ByRef str As String, ByVal start As Integer) As Byte
    On Error GoTo Error_Handler
    If Len(str) < start Then Exit Function
    String_To_Byte = Asc(mid$(str, start, 1))
    Exit Function
Error_Handler:
End Function

Public Function Long_To_String(ByVal Var As Long) As String
    On Error GoTo Long_To_String_Err
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
    Long_To_String = Chr$(val("&H" & Left$(temp, 2))) & Chr$(val("&H" & mid$(temp, 3, 2))) & Chr$(val("&H" & mid$(temp, 5, 2))) & Chr$(val("&H" & mid$(temp, 7, 2)))
    Exit Function
ErrorHandler:
    Exit Function
Long_To_String_Err:
    Call RegistrarError(Err.Number, Err.Description, "modLocale.Long_To_String", Erl)
    Resume Next
End Function

Public Function String_To_Long(ByRef str As String, ByVal start As Integer) As Long
    On Error GoTo ErrorHandler
    If Len(str) < start - 3 Then Exit Function
    Dim temp_str  As String
    Dim temp_str2 As String
    Dim temp_str3 As String
    'Tomamos los últimos 3 bytes y convertimos sus valroes ASCII a hexa
    temp_str = hex$(Asc(mid$(str, start + 1, 1)))
    temp_str2 = hex$(Asc(mid$(str, start + 2, 1)))
    temp_str3 = hex$(Asc(mid$(str, start + 3, 1)))
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
    String_To_Long = val("&H" & hex$(Asc(mid$(str, start, 1))) & temp_str & temp_str2 & temp_str3)
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
