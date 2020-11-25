Attribute VB_Name = "TileEngine_Map"
Option Explicit

Public Function Letter_Set(ByVal grh_index As Long, ByVal text_string As String) As Boolean
    '*****************************************************************
    'Author: Augusto José Rando
    '*****************************************************************
    letter_text = text_string
    letter_grh.GrhIndex = grh_index
    Letter_Set = True
    map_letter_fadestatus = 1

End Function

Public Function Map_Letter_Fade_Set(ByVal grh_index As Long, Optional ByVal after_grh As Long = -1) As Boolean

    '*****************************************************************
    'Author: Augusto José Rando
    '*****************************************************************
    If grh_index <= 0 Or grh_index = map_letter_grh.GrhIndex Then Exit Function
        
    If after_grh = -1 Then
    
        map_letter_grh.GrhIndex = grh_index
        map_letter_fadestatus = 1
        map_letter_a = 0
        map_letter_grh_next = 0
        
    Else
        map_letter_grh.GrhIndex = after_grh
        map_letter_fadestatus = 1
        map_letter_a = 0
        map_letter_grh_next = grh_index

    End If
    
    Map_Letter_Fade_Set = True

End Function

Public Function Map_Letter_UnSet() As Boolean
    '*****************************************************************
    'Author: Augusto José Rando
    '*****************************************************************
    map_letter_grh.GrhIndex = 0
    map_letter_fadestatus = 0
    map_letter_a = 0
    map_letter_grh_next = 0
    Map_Letter_UnSet = True

End Function

Public Function Letter_UnSet() As Boolean
    '*****************************************************************
    'Author: Augusto José Rando
    '*****************************************************************
    letter_text = vbNullString
    letter_grh.GrhIndex = 0
    Letter_UnSet = True

End Function

Public Function Map_Base_Light_Get() As Long
    '**************************************************************
    'Author: Aaron Perkins - Modified by Augusto José Rando
    'Last Modify Date: 6/12/2005
    '
    '**************************************************************
    Map_Base_Light_Get = map_base_light

End Function

Public Function Map_Base_Light_Set(ByVal base_light As Long)

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '
    '**************************************************************
    If map_base_light <> base_light Then
        map_base_light = base_light
        
    End If
    
End Function

Public Function Map_Fill(ByVal grh_index As Long, ByVal layer As Byte, Optional ByVal light_base_color As Long = -1, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single) As Boolean

    '**************************************************************
    'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
    'Last Modify Date: 1/04/2003
    '
    '**************************************************************
    Dim x As Integer

    Dim y As Integer
    
    'Base light color
    If light_base_color <> -1 Then
        If Not Map_Base_Light_Set(light_base_color) Then Exit Function

    End If
        
    Map_Fill = True

End Function

Public Function Map_FX_Group_Next_Open(ByVal x As Byte, ByVal y As Byte) As Integer

    '*****************************************************************
    'Author: Augusto José Rando
    '*****************************************************************
    On Error GoTo ErrorHandler:

    Dim loopc As Long
    
    If MapData(x, y).FxCount = 0 Then
        MapData(x, y).FxCount = 1
        ReDim MapData(x, y).FxList(1 To 1)
        Map_FX_Group_Next_Open = 1
        Exit Function

    End If
    
    loopc = 1

    Do Until MapData(x, y).FxList(loopc).FxIndex = 0

        If loopc = MapData(x, y).FxCount Then
            Map_FX_Group_Next_Open = MapData(x, y).FxCount + 1
            MapData(x, y).FxCount = Map_FX_Group_Next_Open
            ReDim Preserve MapData(x, y).FxList(1 To Map_FX_Group_Next_Open)
            Exit Function

        End If

        loopc = loopc + 1
    Loop

    Map_FX_Group_Next_Open = loopc
    Exit Function

ErrorHandler:
    MapData(x, y).FxCount = 1
    ReDim MapData(x, y).FxList(1 To 1)
    Map_FX_Group_Next_Open = 1

End Function

Public Sub Draw_Sombra(ByRef grh As grh, ByVal x As Integer, ByVal y As Integer, ByVal center As Byte, ByVal animate As Byte, Optional ByVal Alpha As Boolean, Optional ByVal map_x As Byte = 1, Optional ByVal map_y As Byte = 1, Optional ByVal angle As Single)

    On Error Resume Next

    ' If meteo_estado = 3 Or meteo_estado = 4 Then Exit Sub
    ' If UserEstado = 1 Then Exit Sub
    ' If bTecho Then Exit Sub
    Dim CurrentGrhIndex As Long

    If grh.GrhIndex = 0 Then Exit Sub
    'Por ladder
    
    CurrentGrhIndex = GrhData(grh.GrhIndex).Frames(grh.framecounter)

    If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
        x = x - Int(GrhData(CurrentGrhIndex).TileWidth * (32 \ 2)) + 32 \ 2
    End If

    If GrhData(grh.GrhIndex).TileHeight <> 1 Then
        y = y - Int(GrhData(CurrentGrhIndex).TileHeight * 32) + 32
    End If

    Call Batch_Textured_Box_Shadow(x, y, GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, GrhData(CurrentGrhIndex).FileNum, MapData(map_x, map_y).light_value)
End Sub

Sub Engine_Weather_UpdateFog()

    '*****************************************************************
    'Update the fog effects
    '*****************************************************************
    Dim TempGrh     As grh

    Dim i           As Long

    Dim x           As Long

    Dim y           As Long

    Dim cc(3)       As Long

    Dim ElapsedTime As Single

    ElapsedTime = Engine_ElapsedTime

    If WeatherFogCount = 0 Then WeatherFogCount = 13

    WeatherFogX1 = WeatherFogX1 + (ElapsedTime * (0.018 + Rnd * 0.01)) + (LastOffsetX - ParticleOffsetX)
    WeatherFogY1 = WeatherFogY1 + (ElapsedTime * (0.013 + Rnd * 0.01)) + (LastOffsetY - ParticleOffsetY)
    
    Do While WeatherFogX1 < -512
        WeatherFogX1 = WeatherFogX1 + 512
    Loop

    Do While WeatherFogY1 < -512
        WeatherFogY1 = WeatherFogY1 + 512
    Loop

    Do While WeatherFogX1 > 0
        WeatherFogX1 = WeatherFogX1 - 512
    Loop

    Do While WeatherFogY1 > 0
        WeatherFogY1 = WeatherFogY1 - 512
    Loop
    
    WeatherFogX2 = WeatherFogX2 - (ElapsedTime * (0.037 + Rnd * 0.01)) + (LastOffsetX - ParticleOffsetX)
    WeatherFogY2 = WeatherFogY2 - (ElapsedTime * (0.021 + Rnd * 0.01)) + (LastOffsetY - ParticleOffsetY)

    Do While WeatherFogX2 < -512
        WeatherFogX2 = WeatherFogX2 + 512
    Loop

    Do While WeatherFogY2 < -512
        WeatherFogY2 = WeatherFogY2 + 512
    Loop

    Do While WeatherFogX2 > 0
        WeatherFogX2 = WeatherFogX2 - 512
    Loop

    Do While WeatherFogY2 > 0
        WeatherFogY2 = WeatherFogY2 - 512
    Loop

    TempGrh.framecounter = 1
    
    'Render fog 2
    TempGrh.GrhIndex = 32014
    x = 2
    y = -1

    cc(1) = D3DColorARGB(AlphaNiebla, 255, 255, 255)
    cc(2) = D3DColorARGB(AlphaNiebla, 255, 255, 255)
    cc(3) = D3DColorARGB(AlphaNiebla, 255, 255, 255)
    cc(0) = D3DColorARGB(AlphaNiebla, 255, 255, 255)

    For i = 1 To WeatherFogCount
        Draw_Grh TempGrh, (x * 512) + WeatherFogX2, (y * 512) + WeatherFogY2, 0, 0, cc()
        x = x + 1

        If x > (1 + (ScreenWidth \ 512)) Then
            x = 0
            y = y + 1

        End If

    Next i
            
    'Render fog 1
    TempGrh.GrhIndex = 32015
    x = 0
    y = 0
    cc(1) = D3DColorARGB(AlphaNiebla, 255, 255, 255)
    cc(2) = D3DColorARGB(AlphaNiebla, 255, 255, 255)
    cc(3) = D3DColorARGB(AlphaNiebla, 255, 255, 255)
    cc(0) = D3DColorARGB(AlphaNiebla, 255, 255, 255)

    For i = 1 To WeatherFogCount
        Draw_Grh TempGrh, (x * 512) + WeatherFogX1, (y * 512) + WeatherFogY1, 0, 0, cc()
        x = x + 1

        If x > (2 + (ScreenWidth \ 512)) Then
            x = 0
            y = y + 1

        End If

    Next i

End Sub


