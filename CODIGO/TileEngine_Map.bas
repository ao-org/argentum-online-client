Attribute VB_Name = "TileEngine_Map"
Option Explicit

Sub SwitchMap(ByVal map As Integer)
    
    On Error GoTo SwitchMap_Err
    
    
    'Cargamos el mapa.
    Call Recursos.CargarMapa(map)

    map_light = global_light

    Call DibujarMiniMapa
    
    CurMap = map
    
    If Musica Then
        
        If MapDat.music_numberLow > 0 Then
        
            If Sound.MusicActual <> MapDat.music_numberLow Then
                Sound.NextMusic = MapDat.music_numberLow
                Sound.Fading = 200
            End If

        Else

            If MapDat.music_numberHi > 0 Then
                
                If Sound.MusicActual <> MapDat.music_numberHi Then
                    Sound.NextMusic = MapDat.music_numberHi
                    Sound.Fading = 100
                End If

                Call ReproducirMp3(MapDat.music_numberHi)
                
                Call Sound.Music_Load(MapDat.music_numberHi, 0, 0)
                
                Call Sound.Music_Play

            End If

        End If

    End If

    If bRain And MapDat.LLUVIA Then
        Call Graficos_Particulas.Engine_MeteoParticle_Set(Particula_Lluvia)
    
    ElseIf bNieve And MapDat.NIEVE Then
        Call Graficos_Particulas.Engine_MeteoParticle_Set(Particula_Nieve)

    End If
    
    If AmbientalActivated = 1 Then
        Call AmbientarAudio(map)
    End If

    Call NameMapa(map)

    
    Exit Sub

SwitchMap_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine_Map.SwitchMap", Erl)
    Resume Next
    
End Sub

Function HayAgua(ByVal x As Integer, ByVal y As Integer) As Boolean
    
    On Error GoTo HayAgua_Err
    

    With MapData(x, y).Graphic(1)
        HayAgua = ((.GrhIndex >= 1505 And .GrhIndex <= 1520) Or (.GrhIndex >= 24223 And .GrhIndex <= 24238) Or _
            (.GrhIndex >= 24143 And .GrhIndex <= 24158) Or (.GrhIndex >= 468 And .GrhIndex <= 483) Or _
            (.GrhIndex >= 44668 And .GrhIndex <= 44939) Or (.GrhIndex >= 24303 And .GrhIndex <= 24318))
    End With

    
    Exit Function

HayAgua_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine_Map.HayAgua", Erl)
    Resume Next
    
End Function

Function EsArbol(ByVal GrhIndex As Long) As Boolean
    
    On Error GoTo EsArbol_Err
    
    EsArbol = GrhIndex = 7000 Or GrhIndex = 7001 Or GrhIndex = 7002 Or GrhIndex = 641 Or GrhIndex = 26075 Or GrhIndex = 643 Or GrhIndex = 644 Or _
       GrhIndex = 647 Or GrhIndex = 26076 Or GrhIndex = 7222 Or GrhIndex = 7223 Or GrhIndex = 7224 Or GrhIndex = 7225 Or GrhIndex = 7226 Or _
       GrhIndex = 26077 Or GrhIndex = 26079 Or GrhIndex = 735 Or GrhIndex = 32343 Or GrhIndex = 32344 Or GrhIndex = 26080 Or GrhIndex = 26081 Or _
       GrhIndex = 32345 Or GrhIndex = 32346 Or GrhIndex = 32347 Or GrhIndex = 32348 Or GrhIndex = 32349 Or GrhIndex = 32350 Or GrhIndex = 32351 Or _
       GrhIndex = 32352 Or GrhIndex = 14961 Or GrhIndex = 14950 Or GrhIndex = 14951 Or GrhIndex = 14952 Or GrhIndex = 14953 Or GrhIndex = 14954 Or _
       GrhIndex = 14955 Or GrhIndex = 14956 Or GrhIndex = 14957 Or GrhIndex = 14958 Or GrhIndex = 14959 Or GrhIndex = 14962 Or GrhIndex = 14963 Or _
       GrhIndex = 14964 Or GrhIndex = 14967 Or GrhIndex = 14968 Or GrhIndex = 14969 Or GrhIndex = 14970 Or GrhIndex = 14971 Or GrhIndex = 14972 Or _
       GrhIndex = 14973 Or GrhIndex = 14974 Or GrhIndex = 14975 Or GrhIndex = 14976 Or GrhIndex = 14978 Or GrhIndex = 14980 Or GrhIndex = 14982 Or _
       GrhIndex = 14983 Or GrhIndex = 14984 Or GrhIndex = 14985 Or GrhIndex = 14987 Or GrhIndex = 14988 Or GrhIndex = 26078 Or GrhIndex = 26192

    
    Exit Function

EsArbol_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine_Map.EsArbol", Erl)
    Resume Next
    
End Function

Public Function EsObjetoFijo(ByVal x As Integer, ByVal y As Integer) As Boolean
    
    On Error GoTo EsObjetoFijo_Err
    
    Dim OBJIndex As Integer
    OBJIndex = MapData(x, y).OBJInfo.OBJIndex
    
    Dim ObjType As eObjType
    ObjType = ObjData(OBJIndex).ObjType
    
    EsObjetoFijo = ObjType = eObjType.otForos Or ObjType = eObjType.otCarteles Or ObjType = eObjType.otArboles Or ObjType = eObjType.otYacimiento Or ObjType = eObjType.OtDecoraciones

    
    Exit Function

EsObjetoFijo_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine_Map.EsObjetoFijo", Erl)
    Resume Next
    
End Function

Public Function Letter_Set(ByVal grh_index As Long, ByVal text_string As String) As Boolean
    '*****************************************************************
    'Author: Augusto José Rando
    '*****************************************************************
    
    On Error GoTo Letter_Set_Err
    
    letter_text = text_string
    letter_grh.GrhIndex = grh_index
    Letter_Set = True
    map_letter_fadestatus = 1

    
    Exit Function

Letter_Set_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine_Map.Letter_Set", Erl)
    Resume Next
    
End Function

Public Function Map_Letter_Fade_Set(ByVal grh_index As Long, Optional ByVal after_grh As Long = -1) As Boolean
    
    On Error GoTo Map_Letter_Fade_Set_Err
    

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

    
    Exit Function

Map_Letter_Fade_Set_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine_Map.Map_Letter_Fade_Set", Erl)
    Resume Next
    
End Function

Public Function Map_Letter_UnSet() As Boolean
    '*****************************************************************
    'Author: Augusto José Rando
    '*****************************************************************
    
    On Error GoTo Map_Letter_UnSet_Err
    
    map_letter_grh.GrhIndex = 0
    map_letter_fadestatus = 0
    map_letter_a = 0
    map_letter_grh_next = 0
    Map_Letter_UnSet = True

    
    Exit Function

Map_Letter_UnSet_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine_Map.Map_Letter_UnSet", Erl)
    Resume Next
    
End Function

Public Function Letter_UnSet() As Boolean
    '*****************************************************************
    'Author: Augusto José Rando
    '*****************************************************************
    
    On Error GoTo Letter_UnSet_Err
    
    letter_text = vbNullString
    letter_grh.GrhIndex = 0
    Letter_UnSet = True

    
    Exit Function

Letter_UnSet_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine_Map.Letter_UnSet", Erl)
    Resume Next
    
End Function

Public Sub SetGlobalLight(ByVal base_light As Long)
    
    On Error GoTo SetGlobalLight_Err
    
    Call Long_2_RGBA(global_light, base_light)
    global_light.A = 255
    light_transition = 1#
    
    Exit Sub

SetGlobalLight_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine_Map.SetGlobalLight", Erl)
    Resume Next
    
End Sub

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

Public Sub Draw_Sombra(ByRef grh As grh, ByVal x As Integer, ByVal y As Integer, ByVal center As Byte, ByVal animate As Byte, Optional ByVal Alpha As Boolean, Optional ByVal map_x As Byte = 1, Optional ByVal map_y As Byte = 1, Optional ByVal Angle As Single)
    
    On Error GoTo Draw_Sombra_Err
    

    

    If grh.GrhIndex = 0 Or grh.GrhIndex > MaxGrh Then Exit Sub
    
    Dim CurrentFrame As Integer
    CurrentFrame = 1

    If animate Then
        If grh.Started > 0 Then
            Dim ElapsedFrames As Long
            ElapsedFrames = Fix((FrameTime - grh.Started) / grh.speed)

            If grh.Loops = INFINITE_LOOPS Or ElapsedFrames < GrhData(grh.GrhIndex).NumFrames * (grh.Loops + 1) Then
                CurrentFrame = ElapsedFrames Mod GrhData(grh.GrhIndex).NumFrames + 1

            Else
                grh.Started = 0
            End If

        End If

    End If
    
    Dim CurrentGrhIndex As Long
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(grh.GrhIndex).Frames(CurrentFrame)

    If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
        x = x - Int(GrhData(CurrentGrhIndex).TileWidth * (32 \ 2)) + 32 \ 2
    End If

    If GrhData(grh.GrhIndex).TileHeight <> 1 Then
        y = y - Int(GrhData(CurrentGrhIndex).TileHeight * 32) + 32
    End If

    Call Batch_Textured_Box_Shadow(x, y, GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, GrhData(CurrentGrhIndex).FileNum, MapData(map_x, map_y).light_value)
    
    Exit Sub

Draw_Sombra_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine_Map.Draw_Sombra", Erl)
    Resume Next
    
End Sub

Sub Engine_Weather_UpdateFog()
    
    On Error GoTo Engine_Weather_UpdateFog_Err
    

    '*****************************************************************
    'Update the fog effects
    '*****************************************************************
    Dim TempGrh     As grh

    Dim i           As Long

    Dim x           As Long

    Dim y           As Long

    Dim cc(3)       As RGBA

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

    Call InitGrh(TempGrh, 32014)

    x = 2
    y = -1

    Call RGBAList(cc, 255, 255, 255, AlphaNiebla)

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

    For i = 1 To WeatherFogCount
        Draw_Grh TempGrh, (x * 512) + WeatherFogX1, (y * 512) + WeatherFogY1, 0, 0, cc()
        x = x + 1

        If x > (2 + (ScreenWidth \ 512)) Then
            x = 0
            y = y + 1

        End If

    Next i

    
    Exit Sub

Engine_Weather_UpdateFog_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine_Map.Engine_Weather_UpdateFog", Erl)
    Resume Next
    
End Sub

Sub MapUpdateGlobalLight()
    
    On Error GoTo MapUpdateGlobalLight_Err
    

    Dim x As Integer, y As Integer
    
    ' Reseteamos toda la luz del mapa
    For y = YMinMapSize To YMaxMapSize
        For x = XMinMapSize To XMaxMapSize
            With MapData(x, y)
            
                .light_value(0) = global_light
                .light_value(1) = global_light
                .light_value(2) = global_light
                .light_value(3) = global_light
                
            End With
        Next x
    Next y
    
    Call LucesRedondas.LightRenderAll
    Call LucesCuadradas.Light_Render_All
    
    
    Exit Sub

MapUpdateGlobalLight_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine_Map.MapUpdateGlobalLight", Erl)
    Resume Next
    
End Sub
