Attribute VB_Name = "TileEngine_Map"
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

Sub SwitchMap(ByVal map As Integer)
    
    On Error GoTo SwitchMap_Err
    
    
    'Cargamos el mapa.
    Call Recursos.CargarMapa(map)

    map_light = global_light

    Call DibujarMiniMapa
    map_letter_a = 0
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
    
    If MapDat.Seguro = 1 Then
        frmMain.Coord.ForeColor = RGB(0, 170, 0)
    Else
         If MostrarTutorial And tutorial_index <= 0 And isLogged Then
            If tutorial(e_tutorialIndex.TUTORIAL_ZONA_INSEGURA).Activo = 1 Then
                tutorial_index = e_tutorialIndex.TUTORIAL_ZONA_INSEGURA
                'TUTORIAL MAPA INSEGURO
                Call mostrarCartel(tutorial(tutorial_index).titulo, tutorial(tutorial_index).textos(1), tutorial(tutorial_index).grh, _
                -1, &H164B8A, , , False, 100, 479, 100, 535, 640, 530, 64, 64)
            End If
        End If
        frmMain.Coord.ForeColor = RGB(170, 0, 0)
    End If

    
    Exit Sub

SwitchMap_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine_Map.SwitchMap", Erl)
    Resume Next
    
End Sub

Function HayAgua(ByVal x As Integer, ByVal y As Integer) As Boolean
    
    On Error GoTo HayAgua_Err
    

    With MapData(x, y).Graphic(1)
            HayAgua = (.GrhIndex >= 1505 And .GrhIndex <= 1520) Or _
                        (.GrhIndex >= 124 And .GrhIndex <= 139) Or _
                        (.GrhIndex >= 24223 And .GrhIndex <= 24238) Or _
                        (.GrhIndex >= 24303 And .GrhIndex <= 24318) Or _
                        (.GrhIndex >= 468 And .GrhIndex <= 483) Or _
                        (.GrhIndex >= 44668 And .GrhIndex <= 44683) Or _
                        (.GrhIndex >= 24143 And .GrhIndex <= 24158) Or _
                        (.GrhIndex >= 12628 And .GrhIndex <= 12643) Or _
                        (.GrhIndex >= 2948 And .GrhIndex <= 2963)
    End With

    
    Exit Function

HayAgua_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine_Map.HayAgua", Erl)
    Resume Next
    
End Function

Function HayLava(ByVal x As Integer, ByVal y As Integer) As Boolean
    
    On Error GoTo HayLava_Err
    

    With MapData(x, y).Graphic(1)
        HayLava = .GrhIndex >= 57400 And .GrhIndex <= 57415
    End With

    
    Exit Function

HayLava_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine_Map.HayLava", Erl)
    Resume Next
    
End Function

Function EsArbol(ByVal GrhIndex As Long) As Boolean
    
    On Error GoTo EsArbol_Err
    
    EsArbol = GrhIndex = 304 Or GrhIndex = 305 Or GrhIndex = 641 Or GrhIndex = 643 Or GrhIndex = 644 Or GrhIndex = 647 Or GrhIndex = 735 Or GrhIndex = 1121 Or GrhIndex = 1126 Or GrhIndex = 2931 Or _
              GrhIndex = 12309 Or GrhIndex = 12310 Or GrhIndex = 16833 Or GrhIndex = 16834 Or GrhIndex = 7020 Or GrhIndex = 11903 Or GrhIndex = 11904 Or _
              GrhIndex = 11905 Or GrhIndex = 11906 Or GrhIndex = 12160 Or GrhIndex = 15698 Or GrhIndex = 14504 Or GrhIndex = 15697 Or _
              (GrhIndex >= 12581 And GrhIndex <= 12586) Or (GrhIndex >= 12164 And GrhIndex <= 12179) Or _
              (GrhIndex >= 26075 And GrhIndex <= 26081) Or GrhIndex = 26192 Or (GrhIndex >= 32142 And GrhIndex <= 32162) Or (GrhIndex >= 32343 And GrhIndex <= 32352) Or _
              (GrhIndex >= 55626 And GrhIndex <= 55640) Or GrhIndex = 55642 Or _
              (GrhIndex >= 50985 And GrhIndex <= 50991) Or (GrhIndex >= 2547 And GrhIndex <= 2549) Or (GrhIndex >= 6597 And GrhIndex <= 6598) Or (GrhIndex >= 15108 And GrhIndex <= 15110) Or GrhIndex = 11904 Or GrhIndex = 11905 Or GrhIndex = 11906 Or GrhIndex = 12160 Or _
              GrhIndex = 7220 Or GrhIndex = 50990 Or GrhIndex = 6597 Or GrhIndex = 6598 Or GrhIndex = 2548 Or GrhIndex = 2549 Or _
              GrhIndex = 463 Or GrhIndex = 1880 Or GrhIndex = 1878 Or GrhIndex = 9513 Or GrhIndex = 9514 Or GrhIndex = 9515 Or GrhIndex = 9518 Or GrhIndex = 9519 Or GrhIndex = 9520 Or GrhIndex = 9529 Or _
              GrhIndex = 55633 Or GrhIndex = 55627 Or GrhIndex = 15510 Or GrhIndex = 14775 Or GrhIndex = 14687
              
    
    Exit Function

EsArbol_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine_Map.EsArbol", Erl)
    Resume Next
    
End Function

Function AgregarSombra(ByVal GrhIndex As Long) As Boolean
    
    On Error GoTo AgregarSombra_Err
    
    AgregarSombra = GrhIndex = 5624 Or GrhIndex = 5625 Or GrhIndex = 5626 Or GrhIndex = 5627 Or GrhIndex = 51716

    
    Exit Function

AgregarSombra_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine_Map.AgregarSombra", Erl)
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
            ElapsedFrames = Fix(0.5 * (FrameTime - grh.Started) / grh.speed)

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
    
    Exit Sub

MapUpdateGlobalLight_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine_Map.MapUpdateGlobalLight", Erl)
    Resume Next
    
End Sub

Sub MapUpdateGlobalLightRender()
    
    On Error GoTo MapUpdateGlobalLight_Err
    

    Dim x As Integer, y As Integer
    Dim MinX As Long, MinY As Long, MaxX As Long, MaxY As Long
    MinX = 1
    MinY = 1
    MaxX = 100
    MaxY = 100
    
    ' Reseteamos toda la luz del mapa
    For y = MinY To MaxY
        For x = MinX To MaxX
            With MapData(x, y)
            
                .light_value(0) = global_light
                .light_value(1) = global_light
                .light_value(2) = global_light
                .light_value(3) = global_light
                
            End With
        Next x
    Next y
    
    Call LucesRedondas.LightRenderAll(MinX, MinY, MaxX, MaxY) '(MinX, MinY, MaxX, MaxY)
    Call LucesCuadradas.Light_Render_All(MinX, MinY, MaxX, MaxY)  '(MinX, MinY, MaxX, MaxY)
        
    Exit Sub

MapUpdateGlobalLight_Err:
   ' Call RegistrarError(Err.Number, Err.Description, "TileEngine_Map.MapUpdateGlobalLightRender", Erl)
    Resume Next
    
End Sub
