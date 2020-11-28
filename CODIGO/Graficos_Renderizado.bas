Attribute VB_Name = "Graficos_Renderizado"
Option Explicit

'Letter showing on screen
Private letter_text           As String
Private letter_grh            As grh
Private map_letter_grh        As grh
Private map_letter_grh_next   As Long
Private map_letter_a          As Single
Private map_letter_fadestatus As Byte

Public Sub render()

    '*****************************************************
    '****** Coded by Menduz (lord.yo.wo@gmail.com) *******
    '*****************************************************
    Rem On Error GoTo ErrorHandler:
    Dim temp_array(3) As Long
    
    If Map_light_base = -1 And Not EfectoEnproceso Then
        Meteo_Engine.Meteo_Logic
    ElseIf UserEstado = 1 Then
        Meteo_Engine.Meteo_Logic
       
    End If
    
    Call Engine_BeginScene
    
    Call ShowNextFrame

    frmmain.fps.Caption = "FPS: " & fps
    frmmain.ms.Caption = PingRender & "ms"
       
    If frmmain.Contadores.Enabled Then

        Dim PosY As Integer
       
        Dim PosX As Integer

        If FullScreen Then
            PosY = 90
            PosX = 10
            
            temp_array(0) = RGB(0, 255, 0)
            temp_array(1) = temp_array(0)
            temp_array(2) = temp_array(0)
            temp_array(3) = temp_array(0)
            Engine_Draw_Box 665, 480, 37, 15, D3DColorARGB(150, 100, 100, 100)
            
            Engine_Text_Render Val(UserAtributos(eAtributos.Fuerza)), 665, 480, temp_array, 1, True, 10, 160
            temp_array(0) = RGB(255, 255, 0)
            temp_array(1) = temp_array(0)
            temp_array(2) = temp_array(0)
            temp_array(3) = temp_array(0)
            Engine_Text_Render Val(UserAtributos(eAtributos.Agilidad)), 685, 480, temp_array, 1, True, 0, 160
        Else
            PosY = -10
            PosX = 5

        End If

        If DrogaCounter > 0 Then
            temp_array(0) = D3DColorXRGB(0, 153, 0)
            temp_array(1) = temp_array(0)
            temp_array(2) = temp_array(0)
            temp_array(3) = temp_array(0)
            
            PosY = PosY + 15
            Engine_Text_Render "Potenciado: " & CLng(DrogaCounter) & "s", PosX, PosY, temp_array, 1, True, 0, 160

        End If
        
        If OxigenoCounter > 0 Then

            Dim HR                  As Integer

            Dim ms                  As Integer

            Dim SS                  As Integer

            Dim secs                As Integer

            Dim TextoOxigenoCounter As String
        
            temp_array(0) = RGB(50, 100, 255)
            temp_array(1) = temp_array(0)
            temp_array(2) = temp_array(0)
            temp_array(3) = temp_array(0)
            secs = OxigenoCounter
            HR = secs \ 3600
            ms = (secs Mod 3600) \ 60
            SS = (secs Mod 3600) Mod 60

            If SS > 9 Then
                TextoOxigenoCounter = ms & ":" & SS
            Else
                TextoOxigenoCounter = ms & ":0" & SS

            End If
            
            PosY = PosY + 15

            If ms < 1 Then
                frmmain.oxigenolbl = SS
                frmmain.oxigenolbl.ForeColor = vbRed
            Else
                frmmain.oxigenolbl = ms
                frmmain.oxigenolbl.ForeColor = vbWhite

            End If

            Engine_Text_Render "Oxigeno: " & TextoOxigenoCounter, PosX, PosY, temp_array, 1, True, 0, 128

        End If

    End If
    
    Call Engine_EndScene(Render_Main_Rect)
    
    lFrameLimiter = (GetTickCount() And &H7FFFFFFF)
    FramesPerSecCounter = FramesPerSecCounter + 1
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed

    Exit Sub

End Sub

Sub ShowNextFrame()

    'Call RenderSounds
    Static OffsetCounterX As Single

    Static OffsetCounterY As Single
     
    If UserMoving Then

        '****** Move screen Left and Right if needed ******
        If AddtoUserPos.x <> 0 Then
            OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.x * timerTicksPerFrame * charlist(UserCharIndex).Speeding

            If Abs(OffsetCounterX) >= Abs(32 * AddtoUserPos.x) Then
                OffsetCounterX = 0
                AddtoUserPos.x = 0
                UserMoving = False

            End If

        End If
            
        '****** Move screen Up and Down if needed ******
        If AddtoUserPos.y <> 0 Then
            OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.y * timerTicksPerFrame * charlist(UserCharIndex).Speeding

            If Abs(OffsetCounterY) >= Abs(32 * AddtoUserPos.y) Then
                OffsetCounterY = 0
                AddtoUserPos.y = 0
                UserMoving = False

            End If

        End If

    End If

    If UserCiego Then
        Call ConvertCPtoTP(MouseX, MouseY, TX, TY)
        Call RenderScreenCiego(UserPos.x - AddtoUserPos.x, UserPos.y - AddtoUserPos.y, OffsetCounterX, OffsetCounterY)
    Else
        'Reparacion de pj
        Call ConvertCPtoTP(MouseX, MouseY, TX, TY)
        Call RenderScreen(UserPos.x - AddtoUserPos.x, UserPos.y - AddtoUserPos.y, OffsetCounterX, OffsetCounterY)
                
    End If

End Sub

Public Sub Grh_Render_Advance(ByRef grh As grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByVal Height As Integer, ByVal Width As Integer, ByRef rgb_list() As Long, Optional ByVal h_center As Boolean, Optional ByVal v_center As Boolean, Optional ByVal alpha_blend As Boolean = False)

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
    'Last Modify Date: 11/19/2003
    'Similar to Grh_Render, but let´s you resize the Grh
    '**************************************************************
    Dim tile_width  As Integer

    Dim tile_height As Integer

    Dim grh_index   As Long
    
    'Animation
    If grh.Started Then
        grh.framecounter = grh.framecounter + (timerTicksPerFrame * grh.speed)

        If grh.framecounter > GrhData(grh.GrhIndex).NumFrames Then
            'If Grh.noloop Then
            '    Grh.FrameCounter = GrhData(Grh.GrhIndex).NumFrames
            'Else
            grh.framecounter = 1

            'End If
        End If

    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    If grh.framecounter = 0 Then grh.framecounter = 1
    grh_index = GrhData(grh.GrhIndex).Frames(grh.framecounter)
    
    'Center Grh over X, Y pos
    If GrhData(grh.GrhIndex).TileWidth <> 1 Then
        screen_x = screen_x - Int(GrhData(grh.GrhIndex).TileWidth * (32 \ 2)) + 32 \ 2

    End If
    
    If GrhData(grh.GrhIndex).TileHeight <> 1 Then
        screen_y = screen_y - Int(GrhData(grh.GrhIndex).TileHeight * 32) + 32

    End If
    
    'Draw it to device
    'Device_Box_Textured_Render_Advance grh_index, screen_x, screen_y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, rgb_list, GrhData(grh_index).sX, GrhData(grh_index).sY, Width, Height, alpha_blend, grh.angle
    Device_Textured_Render screen_x, screen_y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, GrhData(grh_index).sX, GrhData(grh_index).sY, GrhData(grh_index).FileNum, rgb_list(), alpha_blend, grh.angle

End Sub

Public Sub Grh_Render(ByRef grh As grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByRef rgb_list() As Long, Optional ByVal h_centered As Boolean = True, Optional ByVal v_centered As Boolean = True, Optional ByVal alpha_blend As Boolean = False)

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 2/28/2003
    'Modified by Juan Martín Sotuyo Dodero
    'Added centering
    '**************************************************************
    Dim tile_width  As Integer

    Dim tile_height As Integer

    Dim grh_index   As Long
    
    If grh.GrhIndex = 0 Then Exit Sub
        
    'Animation
    If grh.Started Then
        grh.framecounter = grh.framecounter + (timerTicksPerFrame * grh.speed)

        If grh.framecounter > GrhData(grh.GrhIndex).NumFrames Then
            'If Grh.noloop Then
            '    Grh.FrameCounter = GrhData(Grh.GrhIndex).NumFrames
            'Else
            grh.framecounter = 1

            'End If
        End If

    End If

    ' particle_group_list(particle_group_index).frame_counter = particle_group_list(particle_group_index).frame_counter + timer_ticks_per_frame
    ' If particle_group_list(particle_group_index).frame_counter > particle_group_list(particle_group_index).frame_speed Then
    '     particle_group_list(particle_group_index).frame_counter = 0
    '      no_move = False
    '  Else
    '     no_move = True
    '  End If

    'Figure out what frame to draw (always 1 if not animated)
    If grh.framecounter = 0 Then grh.framecounter = 1
    ' If Not Grh_Check(Grh.grhindex) Then Exit Sub
    grh_index = GrhData(grh.GrhIndex).Frames(grh.framecounter)

    If grh_index <= 0 Then Exit Sub
    If GrhData(grh_index).FileNum = 0 Then Exit Sub
        
    'Modified by Augusto José Rando
    'Simplier function - according to basic ORE engine
    If h_centered Then
        If GrhData(grh.GrhIndex).TileWidth <> 1 Then
            screen_x = screen_x - Int(GrhData(grh.GrhIndex).TileWidth * (32 \ 2)) + 32 \ 2

        End If

    End If
    
    If v_centered Then
        If GrhData(grh.GrhIndex).TileHeight <> 1 Then
            screen_y = screen_y - Int(GrhData(grh.GrhIndex).TileHeight * 32) + 32

        End If

    End If
    
    'Draw it to device
    Device_Box_Textured_Render grh_index, screen_x, screen_y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, rgb_list(), GrhData(grh_index).sX, GrhData(grh_index).sY, alpha_blend, grh.angle

End Sub

Sub RenderScreenCiego(ByVal tilex As Integer, ByVal tiley As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 8/14/2007
    'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
    'Renders everything to the viewport
    '**************************************************************

    Dim y                As Integer     'Keeps track of where on map we are

    Dim x                As Integer     'Keeps track of where on map we are

    Dim screenminY       As Integer  'Start Y pos on current screen

    Dim screenmaxY       As Integer  'End Y pos on current screen

    Dim screenminX       As Integer  'Start X pos on current screen

    Dim screenmaxX       As Integer  'End X pos on current screen

    Dim minY             As Integer  'Start Y pos on current map

    Dim MaxY             As Integer  'End Y pos on current map

    Dim minX             As Integer  'Start X pos on current map

    Dim MaxX             As Integer  'End X pos on current map

    Dim ScreenX          As Integer  'Keeps track of where to place tile on screen

    Dim ScreenY          As Integer  'Keeps track of where to place tile on screen

    Dim minXOffset       As Integer

    Dim minYOffset       As Integer

    Dim PixelOffsetXTemp As Integer 'For centering grhs

    Dim PixelOffsetYTemp As Integer 'For centering grhs

    Dim CurrentGrhIndex  As Integer

    Dim OffX             As Integer

    Dim Offy             As Integer

    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    minY = screenminY - TileBufferSize
    MaxY = screenmaxY + TileBufferSize
    minX = screenminX - TileBufferSize
    MaxX = screenmaxX + TileBufferSize
    
    'Make sure mins and maxs are allways in map bounds
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize

    End If
    
    If MaxY > YMaxMapSize Then MaxY = YMaxMapSize
    
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize

    End If
    
    If MaxX > XMaxMapSize Then MaxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1

    End If
    
    If screenmaxY < YMaxMapSize Then screenmaxY = screenmaxY + 1
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1

    End If
    
    If screenmaxX < XMaxMapSize Then screenmaxX = screenmaxX + 1
    
    If screenminY < 1 Then screenminY = 1
    If screenminX < 1 Then screenminX = 1
    If screenmaxY > 100 Then screenmaxY = 100
    If screenmaxX > 100 Then screenmaxX = 100
    
    Dim PicClimaRGB(0 To 3) As Long

    Dim Climapic            As grh
   
    ColorCiego(0) = D3DColorARGB(255, 15, 15, 15)
    ColorCiego(1) = ColorCiego(0)
    ColorCiego(2) = ColorCiego(0)
    ColorCiego(3) = ColorCiego(0)
    'If minY < 1 Then minY = 1
    'If minX < 1 Then minX = 1
    ' If maxY > 100 Then maxY = 100
    ' If maxX > 100 Then maxX = 100
    
    'Draw floor layer
    For y = screenminY To screenmaxY
        For x = screenminX To screenmaxX
            'Layer 1 **********************************
            Call Draw_Grh(MapData(x, y).Graphic(1), (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, 0, 1, ColorCiego, , x, y)
            '******************************************
            ScreenX = ScreenX + 1
        Next x

        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - x + screenminX
        ScreenY = ScreenY + 1
    Next y
    
    If HayLayer2 Then
        ScreenY = minYOffset - TileBufferSize

        For y = minY To MaxY
            ScreenX = minXOffset - TileBufferSize

            For x = minX To MaxX

                With MapData(x, y)

                    '***********************************************
                    If MapData(x, y).Graphic(2).GrhIndex <> 0 Then
                        Call Draw_Grh(MapData(x, y).Graphic(2), (ScreenX * 32 + PixelOffsetX), (ScreenY * 32 + PixelOffsetY), 1, 1, ColorCiego, , x, y)

                    End If
              
                End With

                ScreenX = ScreenX + 1
            Next x

            ScreenY = ScreenY + 1
        Next y

    End If
    
    ScreenY = minYOffset - TileBufferSize

    For y = minY To MaxY
        ScreenX = minXOffset - TileBufferSize

        For x = minX To MaxX
            PixelOffsetXTemp = ScreenX * 32 + PixelOffsetX
            PixelOffsetYTemp = ScreenY * 32 + PixelOffsetY
            
            With MapData(x, y)
                '******************************************

                'Object Layer **********************************
                If MapData(x, y).ObjGrh.GrhIndex <> 0 Then
                    Call Draw_Grh(MapData(x, y).ObjGrh, (ScreenX * 32 + PixelOffsetX), (ScreenY * 32 + PixelOffsetY), 1, 1, ColorCiego, , x, y)

                End If
                
                'Char layer ************************************
                'clones
            
                If MapData(x, y).charindex = UserCharIndex Then
                    If x <> UserPos.x Then
                        MapData(x, y).charindex = 0

                    End If
                    
                End If
                
                If .charindex <> 0 Then
                    If charlist(.charindex).AlphaPJ = 255 And charlist(.charindex).active = 1 Then
                        Call Char_RenderCiego(.charindex, PixelOffsetXTemp, PixelOffsetYTemp, x, y)

                    End If

                End If
                
                'Layer 3 *****************************************
                If .Graphic(3).GrhIndex <> 0 Then
                
                    Call Draw_Grh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, ColorCiego, False, x, y)
                            
                End If

                '************************************************

            End With
            
            ScreenX = ScreenX + 1
        Next x

        ScreenY = ScreenY + 1
    Next y

    ScreenY = minYOffset - 5

    ScreenY = minYOffset - TileBufferSize

    For y = minY To MaxY
        ScreenX = minXOffset - TileBufferSize

        For x = minX To MaxX

            With MapData(x, y)
                '***********************************************
                
                If .particle_Index = 184 Then
                    If meteo_estado = 3 Or meteo_estado = 4 Then
                        If .particle_group > 0 Then
                            Call Particle_Group_Render(.particle_group, ScreenX * 32 + PixelOffsetX + 15, ScreenY * 32 + PixelOffsetY + 15)

                        End If

                    End If

                End If

                If .particle_Index <> 184 Then
                    If .particle_group > 0 Then
                        Call Particle_Group_Render(.particle_group, ScreenX * 32 + PixelOffsetX + 15, ScreenY * 32 + PixelOffsetY + 15)

                    End If

                End If
                
            End With

            ScreenX = ScreenX + 1
        Next x

        ScreenY = ScreenY + 1
    Next y
 
    'Draw blocked tiles and grid
 
    If HayLayer4 Then

        Dim rgb_list(0 To 3)  As Long

        Dim rgb_list2(0 To 3) As Long

        rgb_list2(0) = D3DColorARGB(255, ColorAmbiente.r, ColorAmbiente.g, ColorAmbiente.b)
        rgb_list2(1) = D3DColorARGB(255, ColorAmbiente.r, ColorAmbiente.g, ColorAmbiente.b)
        rgb_list2(2) = D3DColorARGB(255, ColorAmbiente.r, ColorAmbiente.g, ColorAmbiente.b)
        rgb_list2(3) = D3DColorARGB(255, ColorAmbiente.r, ColorAmbiente.g, ColorAmbiente.b)
    
        ScreenY = minYOffset - TileBufferSize

        For y = minY To MaxY
            ScreenX = minXOffset - TileBufferSize

            For x = minX To MaxX
        
                If MapData(x, y).Graphic(4).GrhIndex Then
        
                    'Layer 4 **********************************
                    If bTecho Then

                        If MapData(UserPos.x, UserPos.y).Trigger = MapData(x, y).Trigger Then
                    
                            If MapData(x, y).GrhBlend <= 20 Then MapData(x, y).GrhBlend = 20
                            MapData(x, y).GrhBlend = MapData(x, y).GrhBlend - (timerTicksPerFrame * 12)
                    
                            rgb_list(0) = D3DColorARGB(CInt(MapData(x, y).GrhBlend), ColorAmbiente.r, ColorAmbiente.g, ColorAmbiente.b)
                            rgb_list(1) = rgb_list(0)
                            rgb_list(2) = rgb_list(0)
                            rgb_list(3) = rgb_list(0)
                        
                            Call Draw_Grh(MapData(x, y).Graphic(4), ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, 1, 1, ColorCiego, , x, y)
                        Else
                            Call Draw_Grh(MapData(x, y).Graphic(4), ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, 1, 1, ColorCiego, , x, y)

                        End If

                    Else
                 
                        MapData(x, y).GrhBlend = MapData(x, y).GrhBlend + (timerTicksPerFrame * 12)

                        If MapData(x, y).GrhBlend >= 255 Then MapData(x, y).GrhBlend = 255

                        rgb_list(0) = D3DColorARGB(CInt(MapData(x, y).GrhBlend), ColorAmbiente.r, ColorAmbiente.g, ColorAmbiente.b)
                        rgb_list(1) = rgb_list(0)
                        rgb_list(2) = rgb_list(0)
                        rgb_list(3) = rgb_list(0)
                        
                        Call Draw_Grh(MapData(x, y).Graphic(4), ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, 1, 1, ColorCiego, , x, y)
          
                    End If

                End If
 
                '**********************************
                ScreenX = ScreenX + 1
            Next x

            ScreenY = ScreenY + 1
        Next y
        
    End If

    'If MostrarTrofeo Then

    '    Dim TrofeoRGB(0 To 3) As Long
    '    Dim Trofeo As grh
    '    Trofeo.FrameCounter = 1
    '    Trofeo.grhindex = 32018
    '    Trofeo.Started = 1
    '    TrofeoRGB(0) = D3DColorARGB(100, 255, 0, 0)
    '    TrofeoRGB(1) = D3DColorARGB(100, 255, 0, 0)
    '    TrofeoRGB(2) = D3DColorARGB(100, 0, 0, 255)
    '    TrofeoRGB(3) = D3DColorARGB(100, 0, 0, 255)
    '  Engine_Draw_Box CInt(clicX), CInt(clicY), 190, 180, D3DColorARGB(180, 100, 100, 100)
    '        Grh_Render Trofeo, 690, 50, TrofeoRGB, True, True, True
    ' Call Draw_Grh(Trofeo, 690, 50, 1, 0, TrofeoRGB, False, 0, 0, 0)
    'End If

    If Pregunta Then
        'PreguntaScreen = "¿Esta seguro que asen es gay? ¿Que se lo come a fede?"
        Engine_Draw_Box 283, 180, 170, 80, D3DColorARGB(200, 219, 116, 3)
        Engine_Draw_Box 288, 185, 160, 70, D3DColorARGB(200, 51, 27, 3)

        Dim preguntaGrh As grh

        preguntaGrh.framecounter = 1
        preguntaGrh.GrhIndex = 32120
        preguntaGrh.Started = 1
        rgb_list(0) = D3DColorARGB(255, 255, 255, 255)
        rgb_list(1) = rgb_list(0)
        rgb_list(2) = rgb_list(0)
        rgb_list(3) = rgb_list(0)
        Engine_Text_Render PreguntaScreen, 290, 190, rgb_list, 1, True
        Call Draw_Grh(preguntaGrh, 392, 223, 1, 0, rgb_list, False, 0, 0, 0)

    End If

    If bRain Then
        If MapDat.LLUVIA Then
            'Screen positions were hardcoded by now
            ScreenX = 250
            ScreenY = 0
            Call Particle_Group_Render(meteo_particle, ScreenX, ScreenY)
            LastOffsetX = ParticleOffsetX
            LastOffsetY = ParticleOffsetY

        End If

    End If

    If AlphaNiebla Then
        If MapDat.niebla Then
            Engine_Weather_UpdateFog

        End If

    End If

    If bNieve Then
        If MapDat.NIEVE Then
            If Engine_Meteo_Particle_Get <> 0 Then
                'Screen positions were hardcoded by now
                ScreenX = 250
                ScreenY = 0
                Call Particle_Group_Render(meteo_particle, ScreenX, ScreenY)

            End If

        End If

    End If

    'Pelota
    'If DibujarPelota Then

    'If Pelota.Fps = 100 Then DibujarPelota = False: Exit Sub
    '   Pelota.X = Pelota.X + Pelota.DireccionX
    '   Pelota.Y = Pelota.Y + Pelota.DireccionY
    '  Pelota.Fps = Pelota.Fps + 1
    '     Call Particle_Group_Render(spell_particle, Pelota.X, Pelota.Y)
    'End If
    'Pelota

    'If CaminandoMacro Then
    'Call Particle_Group_Render(spell_particle, CaminarX, CaminarY)
    'End If

    If cartel Then

        Dim cartelito(0 To 3) As Long

        Dim TempGrh           As grh

        TempGrh.framecounter = 1
        TempGrh.GrhIndex = GrhCartel
        cartelito(0) = D3DColorARGB(200, 255, 255, 255)
        cartelito(1) = rgb_list(0)
        cartelito(2) = rgb_list(0)
        cartelito(3) = rgb_list(0)
        Call Draw_Grh(TempGrh, CInt(clicX), CInt(clicY), 1, 0, cartelito, False, 0, 0, 0)
        Engine_Text_Render Leyenda, CInt(clicX - 100), CInt(clicY - 130), cartelito, 1, False

    End If

End Sub

Sub RenderScreen(ByVal tilex As Integer, ByVal tiley As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 8/14/2007
    'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
    'Renders everything to the viewport
    On Error Resume Next

    '**************************************************************
    Dim y                As Integer     'Keeps track of where on map we are

    Dim x                As Integer     'Keeps track of where on map we are

    Dim screenminY       As Integer  'Start Y pos on current screen

    Dim screenmaxY       As Integer  'End Y pos on current screen

    Dim screenminX       As Integer  'Start X pos on current screen

    Dim screenmaxX       As Integer  'End X pos on current screen

    Dim minY             As Integer  'Start Y pos on current map

    Dim MaxY             As Integer  'End Y pos on current map

    Dim minX             As Integer  'Start X pos on current map

    Dim MaxX             As Integer  'End X pos on current map

    Dim ScreenX          As Integer  'Keeps track of where to place tile on screen

    Dim ScreenY          As Integer  'Keeps track of where to place tile on screen

    Dim minXOffset       As Integer

    Dim minYOffset       As Integer

    Dim PixelOffsetXTemp As Integer 'For centering grhs

    Dim PixelOffsetYTemp As Integer 'For centering grhs

    Dim CurrentGrhIndex  As Integer

    Dim OffX             As Integer

    Dim Offy             As Integer

    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    minY = screenminY - TileBufferSize
    MaxY = screenmaxY + TileBufferSize
    minX = screenminX - TileBufferSize
    MaxX = screenmaxX + TileBufferSize
    
    'Make sure mins and maxs are allways in map bounds
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize

    End If
    
    If MaxY > YMaxMapSize Then MaxY = YMaxMapSize
    
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize

    End If
    
    If MaxX > XMaxMapSize Then MaxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1

    End If
    
    If screenmaxY < YMaxMapSize Then screenmaxY = screenmaxY + 1
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1

    End If
    
    If screenmaxX < XMaxMapSize Then screenmaxX = screenmaxX + 1
    
    If screenminY < 1 Then screenminY = 1
    If screenminX < 1 Then screenminX = 1
    If screenmaxY > 100 Then screenmaxY = 100
    If screenmaxX > 100 Then screenmaxX = 100
    
    Dim PicClimaRGB(0 To 3) As Long

    Dim Climapic            As grh
   
    'If minY < 1 Then minY = 1
    'If minX < 1 Then minX = 1
    ' If maxY > 100 Then maxY = 100
    ' If maxX > 100 Then maxX = 100
    'estoy renderizando 20 de Y y deberian ser 18
    'estoy renderizando 24 de x y deberian ser 18

    screenmaxY = screenmaxY ' 1 tile menos dibujo, vamos a ver que onda

    'Draw floor layer
    For y = screenminY To screenmaxY
        For x = screenminX To screenmaxX
     
            'Layer 1 **********************************
            Call Draw_Grh(MapData(x, y).Graphic(1), (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, 0, 1, MapData(x, y).light_value, , x, y)
            '******************************************
            ScreenX = ScreenX + 1
        Next x

        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - x + screenminX
        ScreenY = ScreenY + 1
        
    Next y
    
    If HayLayer2 Then
        ScreenY = minYOffset - TileBufferSize

        For y = minY To MaxY ' el -8 lo agrego
            ScreenX = minXOffset - TileBufferSize

            For x = minX To MaxX  ' -7 lo agrego

                With MapData(x, y)

                    '***********************************************
                    If MapData(x, y).Graphic(2).GrhIndex <> 0 Then
                        Call Draw_Grh(MapData(x, y).Graphic(2), (ScreenX * 32 + PixelOffsetX), (ScreenY * 32 + PixelOffsetY), 1, 1, MapData(x, y).light_value(), , x, y)

                    End If
              
                End With

                ScreenX = ScreenX + 1
            Next x

            ScreenY = ScreenY + 1
        Next y

    End If
    
    ScreenY = minYOffset - TileBufferSize

    For y = minY To MaxY
        ScreenX = minXOffset - TileBufferSize

        For x = minX To MaxX
            PixelOffsetXTemp = ScreenX * 32 + PixelOffsetX
            PixelOffsetYTemp = ScreenY * 32 + PixelOffsetY
            
            With MapData(x, y)
                '******************************************

                'Object Layer **********************************
                If MapData(x, y).ObjGrh.GrhIndex <> 0 Then
                    Call Draw_Grh(MapData(x, y).ObjGrh, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(x, y).light_value(), , x, y)

                End If
                
                'Char layer ************************************
                'evitamos reenderizar un clon del usuario
                If MapData(x, y).charindex = UserCharIndex Then
                    If x <> UserPos.x Then
                        MapData(x, y).charindex = 0

                    End If
                    
                End If
                
                If .charindex <> 0 Then
                    If charlist(.charindex).active = 1 Then
                        Call Char_Render(.charindex, PixelOffsetXTemp, PixelOffsetYTemp, x, y)

                    End If

                End If

                If .CharFantasma.Activo Then

                    Dim ColorFantasma(3) As Long
                    
                    If MapData(x, y).CharFantasma.AlphaB >= 3 Then
                        MapData(x, y).CharFantasma.AlphaB = MapData(x, y).CharFantasma.AlphaB - (timerTicksPerFrame * 6.7)
                        ColorFantasma(0) = D3DColorARGB(CInt(MapData(x, y).CharFantasma.AlphaB), ColorAmbiente.r, ColorAmbiente.g, ColorAmbiente.b)
                        ColorFantasma(1) = ColorFantasma(0)
                        ColorFantasma(2) = ColorFantasma(0)
                        ColorFantasma(3) = ColorFantasma(0)

                        If .CharFantasma.Heading = 1 Or .CharFantasma.Heading = 2 Then
                            Call Draw_Grh(.CharFantasma.Escudo, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, ColorFantasma, False, x, y)
                            Call Draw_Grh(.CharFantasma.Body, PixelOffsetXTemp + 1, PixelOffsetYTemp, 1, 1, ColorFantasma, False, x, y)
                            Call Draw_Grh(.CharFantasma.Head, PixelOffsetXTemp + .CharFantasma.OffX, PixelOffsetYTemp + .CharFantasma.Offy, 1, 1, ColorFantasma, False, x, y)
                            Call Draw_Grh(.CharFantasma.Casco, PixelOffsetXTemp + .CharFantasma.OffX, PixelOffsetYTemp + .CharFantasma.Offy, 1, 1, ColorFantasma, False, x, y)
                            Call Draw_Grh(.CharFantasma.Arma, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, ColorFantasma, False, x, y)
                        Else
                        
                            Call Draw_Grh(.CharFantasma.Body, PixelOffsetXTemp + 1, PixelOffsetYTemp, 1, 1, ColorFantasma, False, x, y)
                            Call Draw_Grh(.CharFantasma.Head, PixelOffsetXTemp + .CharFantasma.OffX, PixelOffsetYTemp + .CharFantasma.Offy, 1, 1, ColorFantasma, False, x, y)
                            Call Draw_Grh(.CharFantasma.Escudo, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, ColorFantasma, False, x, y)
                            Call Draw_Grh(.CharFantasma.Casco, PixelOffsetXTemp + .CharFantasma.OffX, PixelOffsetYTemp + .CharFantasma.Offy, 1, 1, ColorFantasma, False, x, y)
                            Call Draw_Grh(.CharFantasma.Arma, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, ColorFantasma, False, x, y)

                        End If

                    Else
                        .CharFantasma.Activo = False

                    End If

                End If

                '*************************************************
                If EsArbol(.Graphic(3).GrhIndex) Then
                    Call Draw_Sombra(.Graphic(3), PixelOffsetXTemp + 40, PixelOffsetYTemp, 1, 1, False, x, y)

                End If
                
                'Layer 3 *****************************************
                If .Graphic(3).GrhIndex <> 0 Then
                    Call Draw_Grh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(x, y).light_value, False, x, y)

                End If

                '************************************************
                
            End With
            
            ScreenX = ScreenX + 1
        Next x

        ScreenY = ScreenY + 1
    Next y

    ScreenY = minYOffset - 5

    ScreenY = minYOffset - TileBufferSize

    For y = minY To MaxY
        ScreenX = minXOffset - TileBufferSize

        For x = minX To MaxX

            With MapData(x, y)

                '***********************************************
                If .particle_group > 0 Then
                    Call Particle_Group_Render(.particle_group, ScreenX * 32 + PixelOffsetX + 15, ScreenY * 32 + PixelOffsetY + 15)

                End If

            End With

            ScreenX = ScreenX + 1
        Next x

        ScreenY = ScreenY + 1
    Next y
 
    'Draw blocked tiles and grid
 
    If HayLayer4 Then

        Dim rgb_list(0 To 3)  As Long

        Dim rgb_list2(0 To 3) As Long

        rgb_list2(0) = D3DColorARGB(255, ColorAmbiente.r, ColorAmbiente.g, ColorAmbiente.b)
        rgb_list2(1) = D3DColorARGB(255, ColorAmbiente.r, ColorAmbiente.g, ColorAmbiente.b)
        rgb_list2(2) = D3DColorARGB(255, ColorAmbiente.r, ColorAmbiente.g, ColorAmbiente.b)
        rgb_list2(3) = D3DColorARGB(255, ColorAmbiente.r, ColorAmbiente.g, ColorAmbiente.b)
    
        ScreenY = minYOffset - TileBufferSize

        For y = minY To MaxY
            ScreenX = minXOffset - TileBufferSize

            For x = minX To MaxX
        
                If MapData(x, y).Graphic(4).GrhIndex Then
                            
                    Dim r, g, b As Byte

                    b = (map_base_light And 16711680) / 65536
                    g = (map_base_light And 65280) / 256
                    r = map_base_light And 255

                    'Layer 4 **********************************
                    If bTecho Then
                        If MapData(UserPos.x, UserPos.y).Trigger = MapData(x, y).Trigger Then
                    
                            If MapData(x, y).GrhBlend <= 20 Then MapData(x, y).GrhBlend = 20
                            MapData(x, y).GrhBlend = MapData(x, y).GrhBlend - (timerTicksPerFrame * 12)
                    
                            rgb_list(0) = D3DColorARGB(CInt(MapData(x, y).GrhBlend), b, g, r)
                            rgb_list(1) = rgb_list(0)
                            rgb_list(2) = rgb_list(0)
                            rgb_list(3) = rgb_list(0)
                        
                            Call Draw_Grh(MapData(x, y).Graphic(4), ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, 1, 1, rgb_list(), , x, y)
                        Else
                            Call Draw_Grh(MapData(x, y).Graphic(4), ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, 1, 1, rgb_list2(), , x, y)

                        End If

                    Else
                 
                        MapData(x, y).GrhBlend = MapData(x, y).GrhBlend + (timerTicksPerFrame * 12)

                        If MapData(x, y).GrhBlend >= 255 Then MapData(x, y).GrhBlend = 255

                        rgb_list(0) = D3DColorARGB(CInt(MapData(x, y).GrhBlend), b, g, r)
                        rgb_list(1) = rgb_list(0)
                        rgb_list(2) = rgb_list(0)
                        rgb_list(3) = rgb_list(0)
                        
                        Call Draw_Grh(MapData(x, y).Graphic(4), ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, 1, 1, rgb_list(), , x, y)
          
                    End If

                End If

                ScreenX = ScreenX + 1
            Next x

            ScreenY = ScreenY + 1
        Next y
        
    End If
        
    ScreenY = minYOffset - TileBufferSize

    For y = minY To MaxY
        ScreenX = minXOffset - TileBufferSize

        For x = minX To MaxX
            PixelOffsetXTemp = ScreenX * 32 + PixelOffsetX
            PixelOffsetYTemp = ScreenY * 32 + PixelOffsetY

            With MapData(x, y)
                
                If MapData(x, y).charindex <> 0 Then
                    If charlist(MapData(x, y).charindex).active = 1 Then
                        Call Char_TextRender(MapData(x, y).charindex, PixelOffsetXTemp, PixelOffsetYTemp, x, y)

                    End If

                End If

                modRenderValue.Draw x, y, PixelOffsetXTemp + 16, PixelOffsetYTemp, timerTicksPerFrame
                
                Dim i         As Byte

                Dim colorz(3) As Long

                If .FxCount > 0 Then

                    For i = 1 To .FxCount

                        If .FxList(i).FxIndex > 0 And .FxList(i).Started <> 0 Then
                            colorz(0) = D3DColorARGB(220, 255, 255, 255)
                            colorz(1) = D3DColorARGB(220, 255, 255, 255)
                            colorz(2) = D3DColorARGB(220, 255, 255, 255)
                            colorz(3) = D3DColorARGB(220, 255, 255, 255)

                            If FxData(.FxList(i).FxIndex).IsPNG = 1 Then
                                Call Draw_GrhFX(.FxList(i), PixelOffsetXTemp + FxData(.FxList(i).FxIndex).OffsetX, PixelOffsetYTemp + FxData(.FxList(i).FxIndex).OffsetY + 20, 1, 1, colorz, False)
                                ' Call Draw_GrhFX(.FxList(i), PixelOffsetXTemp + FxData(.FxList(i).OffsetX, PixelOffsetYTemp + FxData(.FxList(i)).Offsety + 20, 1, 1, colorz, False)
                            Else
                                Call Draw_GrhFX(.FxList(i), PixelOffsetXTemp + FxData(.FxList(i).FxIndex).OffsetX, PixelOffsetYTemp + FxData(.FxList(i).FxIndex).OffsetY + 20, 1, 1, colorz, True)

                            End If

                        End If

                        If .FxList(i).Started = 0 Then
                            .FxList(i).FxIndex = 0

                        End If

                    Next i

                    If .FxList(.FxCount).Started = 0 Then
                        .FxCount = .FxCount - 1

                    End If

                End If

            End With
            
            ScreenX = ScreenX + 1
        Next x

        ScreenY = ScreenY + 1
    Next y

    ScreenY = minYOffset - 5
    
    If bRain Then
        If MapDat.LLUVIA Then
            'Screen positions were hardcoded by now
            ScreenX = 250
            ScreenY = 0
            Call Particle_Group_Render(meteo_particle, ScreenX, ScreenY)
            LastOffsetX = ParticleOffsetX
            LastOffsetY = ParticleOffsetY

        End If

    End If

    If AlphaNiebla Then
        If MapDat.niebla Then
            Engine_Weather_UpdateFog

        End If

    End If

    If bNieve Then
        If MapDat.NIEVE Then
        
            If Graficos_Particulas.Engine_Meteo_Particle_Get <> 0 Then
                'Screen positions were hardcoded by now
                ScreenX = 250
                ScreenY = 0
                Call Particle_Group_Render(meteo_particle, ScreenX, ScreenY)

            End If

        End If

    End If

    Dim macroPic(0 To 3) As Long

    Dim TempGrh          As grh

    If Pregunta Then
        Engine_Draw_Box 283, 180, 170, 80, D3DColorARGB(200, 150, 20, 3)
        Engine_Draw_Box 288, 185, 160, 70, D3DColorARGB(200, 25, 25, 23)

        Dim preguntaGrh As grh

        preguntaGrh.framecounter = 1
        preguntaGrh.GrhIndex = 32120
        preguntaGrh.Started = 1
        macroPic(0) = D3DColorARGB(255, 255, 255, 255)
        macroPic(1) = macroPic(0)
        macroPic(2) = macroPic(0)
        macroPic(3) = macroPic(0)
        Engine_Text_Render PreguntaScreen, 290, 190, macroPic, 1, True
        Call Draw_Grh(preguntaGrh, 392, 223, 1, 0, macroPic, False, 0, 0, 0)

    End If

    Effect_Render_All

    If cartel Then

        Dim cartelito(0 To 3) As Long

        TempGrh.framecounter = 1
        TempGrh.GrhIndex = GrhCartel
        cartelito(0) = D3DColorARGB(200, 255, 255, 255)
        cartelito(1) = rgb_list(0)
        cartelito(2) = rgb_list(0)
        cartelito(3) = rgb_list(0)
        Call Draw_Grh(TempGrh, CInt(clicX), CInt(clicY), 1, 0, cartelito, False, 0, 0, 0)
        Engine_Text_Render Leyenda, CInt(clicX - 100), CInt(clicY - 130), cartelito, 1, False

    End If

    Dim temp_array(0 To 3) As Long

    If map_letter_fadestatus > 0 Then
        If map_letter_fadestatus = 1 Then
            map_letter_a = map_letter_a + (timerTicksPerFrame * 3.5)

            If map_letter_a >= 255 Then
                map_letter_a = 255
                map_letter_fadestatus = 2

            End If

        Else
            map_letter_a = map_letter_a - (timerTicksPerFrame * 3.5)

            If map_letter_a <= 0 Then
                map_letter_fadestatus = 0
                map_letter_a = 0
                 
                If map_letter_grh_next > 0 Then
                    map_letter_grh.GrhIndex = map_letter_grh_next
                    map_letter_fadestatus = 1
                    map_letter_grh_next = 0

                End If
                
            End If

        End If

    End If
    
    If Len(letter_text) Then
        temp_array(0) = D3DColorARGB(CInt(map_letter_a), 179, 95, 0)
        temp_array(1) = D3DColorARGB(CInt(map_letter_a), 179, 95, 0)
        temp_array(2) = D3DColorARGB(CInt(map_letter_a), 179, 95, 0)
        temp_array(3) = D3DColorARGB(CInt(map_letter_a), 179, 95, 0)
        Grh_Render letter_grh, 250, 300, temp_array()
        Engine_Text_RenderGrande letter_text, 360 - Engine_Text_Width(letter_text, False, 4) / 2, 1, temp_array, 5, False, , CInt(map_letter_a)

    End If

    If FullScreen Then
        RenderConsola

    End If

End Sub

Private Sub Char_Render(ByVal charindex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer, ByVal x As Byte, ByVal y As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 12/03/04
    'Draw char's to screen without offcentering them
    '***************************************************
    Dim moved                As Boolean

    Dim Pos                  As Integer

    Dim line                 As String

    Dim color(0 To 3)        As Long

    Dim colorCorazon(0 To 3) As Long

    Dim i                    As Long

    Dim OffsetYname          As Byte

    Dim OffsetYClan          As Byte
    
    Dim OffArma              As Byte
    
    With charlist(charindex)

        If .Heading = 0 Then Exit Sub
    
        If .Moving Then

            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame * .Speeding
                
                'Start animations
                'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).speed > 0 Then .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                .MovArmaEscudo = False
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0

                End If

            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame * .Speeding
                
                'Start animations
                If .Body.Walk(.Heading).speed > 0 Then .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                .MovArmaEscudo = False
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0

                End If

            End If

        End If
        
        'If done moving stop animation
        If Not moved Then
            'Stop animations
            .Body.Walk(.Heading).Started = 0
            .Body.Walk(.Heading).framecounter = 1
            
            If Not .MovArmaEscudo Then
                .Arma.WeaponWalk(.Heading).Started = 0
                .Arma.WeaponWalk(.Heading).framecounter = 1

                .Escudo.ShieldWalk(.Heading).Started = 0
                .Escudo.ShieldWalk(.Heading).framecounter = 1

            End If
            
            .Moving = False

        End If

        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
 
        If .EsNpc Then
            If Len(.nombre) > 0 Then
                If Abs(TX - .Pos.x) < 1 And (Abs(TY - .Pos.y)) < 1 Then

                    Dim colornpcs(3) As Long

                    colornpcs(0) = D3DColorXRGB(0, 129, 195)
                    colornpcs(1) = colornpcs(0)
                    colornpcs(2) = colornpcs(0)
                    colornpcs(3) = colornpcs(0)
                    Pos = InStr(.nombre, "<")

                    If Pos = 0 Then Pos = Len(.nombre) + 2
                    'Nick
                    line = Left$(.nombre, Pos - 2)
                    Engine_Text_Render line, PixelOffsetX + 16 - CInt(Engine_Text_Width(line, True) / 2), PixelOffsetY + 30 - Engine_Text_Height(line, True), colornpcs, 1, True
                        
                    'Clan
                    line = mid$(.nombre, Pos)
                    Engine_Text_Render line, PixelOffsetX + 16 - CInt(Engine_Text_Width(line, True) / 2), PixelOffsetY + 45 - Engine_Text_Height(line, True), colornpcs, 1, True

                End If
                    
                If .simbolo <> 0 Then
                    ' Dim simbolo As grh
                    ' simbolo.framecounter = 1
                    ' simbolo.GrhIndex = 5259 + .simbolo
                    'Call Draw_Grh(TempGrh, PixelOffsetX + 20, PixelOffsetY - 45, 1, 0, colorz, False, 0, 0, 0)
                    Call Draw_GrhIndex(5259 + .simbolo, PixelOffsetX + 6, PixelOffsetY + .Body.HeadOffset.y - 10)
                                
                    ' Debug.Print .simbolo
                End If

            End If

        End If

        colornpcs(0) = D3DColorXRGB(255, 255, 255)
        'line = "me gusta el vino, me quiero casar con tu hermana. pero no se si vos qu e"
        ' Engine_Text_Render line, PixelOffsetX + 16 - Engine_Text_Width(line, True), PixelOffsetY + 30 - Engine_Text_Height(line, True), colornpcs, 1, True
        
        If .Body.Walk(.Heading).GrhIndex Then

            If Not .invisible Then

                Dim colorz(3) As Long

                'Draw Body

                colorz(0) = MapData(x, y).light_value(0)
                colorz(1) = MapData(x, y).light_value(1)
                colorz(2) = MapData(x, y).light_value(2)
                colorz(3) = MapData(x, y).light_value(3)
                
                If .EsEnano Then OffArma = 7
                                
                If .Body_Aura <> "" Then Call Renderizar_Aura(.Body_Aura, PixelOffsetX, PixelOffsetY + OffArma, x, y, charindex)
                If .Arma_Aura <> "" Then Call Renderizar_Aura(.Arma_Aura, PixelOffsetX, PixelOffsetY + OffArma, x, y, charindex)
                If .Otra_Aura <> "" Then Call Renderizar_Aura(.Otra_Aura, PixelOffsetX, PixelOffsetY + OffArma, x, y, charindex)
                If .Escudo_Aura <> "" Then Call Renderizar_Aura(.Escudo_Aura, PixelOffsetX, PixelOffsetY + OffArma, x, y, charindex)
                                
                Select Case .Heading

                    Case EAST

                        If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y)
                                                                    
                        If .iBody < 488 Then
                            Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX + 1, PixelOffsetY, 1, 1, colorz, False, x, y, 0)
                        Else
                            Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y, 0)

                        End If
                                         
                        If .Head.Head(.Heading).GrhIndex Then Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, colorz, False, x, y)
                                         
                        If .Casco.Head(.Heading).GrhIndex Then Call Draw_Grh(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, colorz, False, x, y)
                                             
                        If .Arma.WeaponWalk(.Heading).GrhIndex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY + OffArma, 1, 1, colorz, False, x, y)
                                     
                    Case NORTH

                        If .Arma.WeaponWalk(.Heading).GrhIndex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY + OffArma, 1, 1, colorz, False, x, y)
                                             
                        If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y)
                                             
                        If .iBody < 488 Then
                            Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX + 1, PixelOffsetY, 1, 1, colorz, False, x, y, 0)
                        Else
                            Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y, 0)

                        End If
                                         
                        If .Head.Head(.Heading).GrhIndex Then Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, colorz, False, x, y)
                                         
                        If .Casco.Head(.Heading).GrhIndex Then Call Draw_Grh(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, colorz, False, x, y)
                                     
                    Case WEST

                        If .Arma.WeaponWalk(.Heading).GrhIndex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y)
                                             
                        If .iBody < 488 Then
                            Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX + 1, PixelOffsetY, 1, 1, colorz, False, x, y, 0)
                        Else
                            Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y, 0)

                        End If
                                         
                        If .Head.Head(.Heading).GrhIndex Then Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, colorz, False, x, y)
                                         
                        If .Casco.Head(.Heading).GrhIndex Then Call Draw_Grh(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, colorz, False, x, y)
                                         
                        If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY + OffArma, 1, 1, colorz, False, x, y)

                    Case south
                                         
                        If .iBody < 488 Then
                            Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX + 1, PixelOffsetY, 1, 1, colorz, False, x, y, 0)
                        Else
                            Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y, 0)

                        End If
                                         
                        If .Head.Head(.Heading).GrhIndex Then Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, colorz, False, x, y)
                                         
                        If .Casco.Head(.Heading).GrhIndex Then Call Draw_Grh(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, colorz, False, x, y)
                                         
                        If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY + OffArma, 1, 1, colorz, False, x, y)
                                             
                        If .Arma.WeaponWalk(.Heading).GrhIndex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY + OffArma, 1, 1, colorz, False, x, y)

                End Select

                'Draw name over head
                '  If .transformado = False Then
                If Nombres Then
                    If Len(.nombre) > 0 And Not .EsNpc Then
                        Pos = InStr(.nombre, "<")

                        If Pos = 0 Then Pos = Len(.nombre) + 2
                        If .priv = 0 Then
                                
                            Select Case .status

                                Case 0
                                    color(0) = RGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b)
                                    color(1) = RGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b)
                                    color(2) = RGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b)
                                    color(3) = RGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b)
                                    colorCorazon(0) = D3DColorXRGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b)
                                    colorCorazon(1) = D3DColorXRGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b)
                                    colorCorazon(2) = D3DColorXRGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b)
                                    colorCorazon(3) = D3DColorXRGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b)

                                Case 1
                                    color(0) = RGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b)
                                    color(1) = RGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b)
                                    color(2) = RGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b)
                                    color(3) = RGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b)
                                    colorCorazon(0) = D3DColorXRGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b)
                                    colorCorazon(1) = D3DColorXRGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b)
                                    colorCorazon(2) = D3DColorXRGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b)
                                    colorCorazon(3) = D3DColorXRGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b)

                                Case 2
                                    color(0) = RGB(ColoresPJ(6).r, ColoresPJ(6).g, ColoresPJ(6).b)
                                    color(1) = RGB(ColoresPJ(6).r, ColoresPJ(6).g, ColoresPJ(6).b)
                                    color(2) = RGB(ColoresPJ(6).r, ColoresPJ(6).g, ColoresPJ(6).b)
                                    color(3) = RGB(ColoresPJ(6).r, ColoresPJ(6).g, ColoresPJ(6).b)
                                    colorCorazon(0) = D3DColorXRGB(ColoresPJ(6).r, ColoresPJ(6).g, ColoresPJ(6).b)
                                    colorCorazon(1) = D3DColorXRGB(ColoresPJ(6).r, ColoresPJ(6).g, ColoresPJ(6).b)
                                    colorCorazon(2) = D3DColorXRGB(ColoresPJ(6).r, ColoresPJ(6).g, ColoresPJ(6).b)
                                    colorCorazon(3) = D3DColorXRGB(ColoresPJ(6).r, ColoresPJ(6).g, ColoresPJ(6).b)

                                Case 3
                                    color(0) = RGB(ColoresPJ(7).r, ColoresPJ(7).g, ColoresPJ(7).b)
                                    color(1) = RGB(ColoresPJ(7).r, ColoresPJ(7).g, ColoresPJ(7).b)
                                    color(2) = RGB(ColoresPJ(7).r, ColoresPJ(7).g, ColoresPJ(7).b)
                                    color(3) = RGB(ColoresPJ(7).r, ColoresPJ(7).g, ColoresPJ(7).b)
                                    colorCorazon(0) = D3DColorXRGB(ColoresPJ(7).r, ColoresPJ(7).g, ColoresPJ(7).b)
                                    colorCorazon(1) = D3DColorXRGB(ColoresPJ(7).r, ColoresPJ(7).g, ColoresPJ(7).b)
                                    colorCorazon(2) = D3DColorXRGB(ColoresPJ(7).r, ColoresPJ(7).g, ColoresPJ(7).b)
                                    colorCorazon(3) = D3DColorXRGB(ColoresPJ(7).r, ColoresPJ(7).g, ColoresPJ(7).b)

                            End Select
                                    
                        Else
                            color(0) = RGB(ColoresPJ(.priv).r, ColoresPJ(.priv).g, ColoresPJ(.priv).b)
                            color(1) = RGB(ColoresPJ(.priv).r, ColoresPJ(.priv).g, ColoresPJ(.priv).b)
                            color(2) = RGB(ColoresPJ(.priv).r, ColoresPJ(.priv).g, ColoresPJ(.priv).b)
                            color(3) = RGB(ColoresPJ(.priv).r, ColoresPJ(.priv).g, ColoresPJ(.priv).b)
                            colorCorazon(0) = D3DColorXRGB(ColoresPJ(.priv).r, ColoresPJ(.priv).g, ColoresPJ(.priv).b)
                            colorCorazon(1) = D3DColorXRGB(ColoresPJ(.priv).r, ColoresPJ(.priv).g, ColoresPJ(.priv).b)
                            colorCorazon(2) = D3DColorXRGB(ColoresPJ(.priv).r, ColoresPJ(.priv).g, ColoresPJ(.priv).b)
                            colorCorazon(3) = D3DColorXRGB(ColoresPJ(.priv).r, ColoresPJ(.priv).g, ColoresPJ(.priv).b)
                                    
                        End If
                                            
                        If .group_index > 0 Then
                            If charlist(charindex).group_index = charlist(UserCharIndex).group_index Then
                                color(0) = D3DColorXRGB(255, 255, 255)
                                color(1) = D3DColorXRGB(255, 255, 255)
                                color(2) = D3DColorXRGB(255, 255, 255)
                                color(3) = D3DColorXRGB(255, 255, 255)
                                colorCorazon(0) = D3DColorXRGB(255, 255, 0)
                                colorCorazon(1) = D3DColorXRGB(0, 255, 255)
                                colorCorazon(2) = D3DColorXRGB(0, 255, 0)
                                colorCorazon(3) = D3DColorXRGB(0, 255, 255)

                            End If

                        End If

                        If FullScreen And charindex = UserCharIndex And UserEstado = 0 Then
                            OffsetYname = 16
                            OffsetYClan = 14
                            line = Left$(.nombre, Pos - 2)
                            Grh_Render Marco, PixelOffsetX, PixelOffsetY + 5, colorz, True, True, False
                            Draw_FilledBox PixelOffsetX + 3, PixelOffsetY + 31, (((UserMinHp + 1 / 100) / (UserMaxHp + 1 / 100))) * 26, 4, D3DColorARGB(255, 200, 0, 0), D3DColorARGB(0, 200, 200, 200)
                            Grh_Render Marco, PixelOffsetX, PixelOffsetY + 14, colorz, True, True, False
                            Draw_FilledBox PixelOffsetX + 3, PixelOffsetY + 40, (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100))) * 26, 4, D3DColorARGB(255, 0, 0, 255), D3DColorARGB(0, 200, 200, 200)

                        End If
                            
                        If .clan_index > 0 Then
                            If .clan_index = charlist(UserCharIndex).clan_index And charindex <> UserCharIndex And .MUERTO = 0 Then
                                If .clan_nivel = 5 Then
                                    OffsetYname = 8
                                    OffsetYClan = 6
                                    Grh_Render Marco, PixelOffsetX, PixelOffsetY + 5, colorz, True, True, False
                                    Draw_FilledBox PixelOffsetX + 3, PixelOffsetY + 31, (((.UserMinHp + 1 / 100) / (.UserMaxHp + 1 / 100))) * 26, 4, D3DColorARGB(255, 200, 0, 0), D3DColorARGB(0, 200, 200, 200)

                                End If

                            End If

                        End If
  
                        'Nick
                        line = Left$(.nombre, Pos - 2)
                        Engine_Text_Render line, PixelOffsetX + 15 - CInt(Engine_Text_Width(line, True) / 2), PixelOffsetY + 30 + OffsetYname - Engine_Text_Height(line, True), color, 1, True
                        
                        'Clan
                        Select Case .priv

                            Case 1
                                line = "<Game Design>"

                            Case 2
                                line = "<Game Master>"

                            Case 3, 4
                                line = "<Administrador>"

                            Case Else
                                line = mid$(.nombre, Pos)

                        End Select
                            
                        Engine_Text_Render line, PixelOffsetX + 15 - CInt(Engine_Text_Width(line, True) / 2), PixelOffsetY + 45 + OffsetYClan - Engine_Text_Height(line, True), color, 1, True

                        If .Donador = 1 Then
                            line = Left$(.nombre, Pos - 2)
                            Grh_Render Estrella, PixelOffsetX + 7 + CInt(Engine_Text_Width(line, 1) / 2), PixelOffsetY + 10 + OffsetYname, colorCorazon, True, True, False

                        End If

                    End If

                End If

                ' End If
            Else
            
                Dim mostrarlo As Boolean
                         
                If .priv < charlist(UserCharIndex).priv Then
                    mostrarlo = True

                End If

                If .group_index > 0 Then
                    If charlist(charindex).group_index = charlist(UserCharIndex).group_index Then
                        mostrarlo = True

                    End If

                End If

                If .clan_index > 0 Then
                    If .clan_index = charlist(UserCharIndex).clan_index Then
                        If .clan_nivel >= 3 Then
                            mostrarlo = True

                        End If

                    End If

                End If
                    
                If charindex = UserCharIndex Or mostrarlo = True Then
                    colorz(0) = D3DColorARGB(100, 255, 255, 255)
                    colorz(1) = D3DColorARGB(100, 255, 255, 255)
                    colorz(2) = D3DColorARGB(100, 255, 255, 255)
                    colorz(3) = D3DColorARGB(100, 255, 255, 255)

                    If .Body.Walk(.Heading).GrhIndex Then Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y)

                    If .Head.Head(.Heading).GrhIndex Then Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, colorz, False, x, y)

                    If .Casco.Head(.Heading).GrhIndex Then Call Draw_Grh(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, colorz, False, x, y)

                    If .Arma.WeaponWalk(.Heading).GrhIndex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y)

                    If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y)
                                
                    Pos = InStr(.nombre, "<")

                    If Pos = 0 Then Pos = Len(.nombre) + 2

                    color(0) = D3DColorXRGB(255, 255, 255)
                    color(1) = color(0)
                    color(2) = color(0)
                    color(3) = color(0)
                    colorCorazon(0) = D3DColorXRGB(120, 100, 200)
                    colorCorazon(1) = colorCorazon(0)
                    colorCorazon(2) = colorCorazon(0)
                    colorCorazon(3) = colorCorazon(0)
                                
                    If FullScreen And charindex = UserCharIndex And UserEstado = 0 Then
                        OffsetYname = 16
                        OffsetYClan = 14
                        line = Left$(.nombre, Pos - 2)
                        Grh_Render Marco, PixelOffsetX, PixelOffsetY + 5, color, True, True, False
                        Draw_FilledBox PixelOffsetX + 3, PixelOffsetY + 31, (((UserMinHp + 1 / 100) / (UserMaxHp + 1 / 100))) * 26, 4, D3DColorARGB(255, 200, 0, 0), D3DColorARGB(0, 200, 200, 200)
                        Grh_Render Marco, PixelOffsetX, PixelOffsetY + 14, color, True, True, False
                        Draw_FilledBox PixelOffsetX + 3, PixelOffsetY + 40, (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100))) * 26, 4, D3DColorARGB(255, 0, 0, 255), D3DColorARGB(0, 200, 200, 200)

                    End If
                                
                    color(0) = D3DColorXRGB(200, 100, 100)
                    color(1) = color(0)
                    color(2) = color(0)
                    color(3) = color(0)

                    line = Left$(.nombre, Pos - 2)
                    Engine_Text_Render line, PixelOffsetX + 15 - CInt(Engine_Text_Width(line, True) / 2), PixelOffsetY + 30 + OffsetYname - Engine_Text_Height(line, True), color, 1, True
                        
                    'Clan
                    Select Case .priv

                        Case 1
                            line = "<Game Design>"

                        Case 2
                            line = "<Game Master>"

                        Case 3, 4
                            line = "<Administrador>"

                        Case Else
                            line = mid$(.nombre, Pos)

                    End Select
                            
                    Engine_Text_Render line, PixelOffsetX + 15 - CInt(Engine_Text_Width(line, True) / 2), PixelOffsetY + 45 + OffsetYClan - Engine_Text_Height(line, True), color, 1, True

                    If .Donador = 1 Then
                        line = Left$(.nombre, Pos - 2)
                        Grh_Render Estrella, PixelOffsetX + 7 + CInt(Engine_Text_Width(line, 1) / 2), PixelOffsetY + 10 + OffsetYname, colorCorazon, True, True, False

                    End If

                Else

                    If .TimerI <= 0 Then .TimerIAct = True
                    If .TimerIAct = False Then
                        .TimerI = .TimerI - (timerTicksPerFrame * 1)
                    Else
                        .TimerI = .TimerI + (timerTicksPerFrame * 0.3)

                        If .TimerI >= 40 Then .TimerIAct = False

                    End If

                    colorz(0) = D3DColorARGB(.TimerI, 255, 255, 255)
                    colorz(1) = D3DColorARGB(.TimerI, 255, 255, 255)
                    colorz(2) = D3DColorARGB(.TimerI, 255, 255, 255)
                    colorz(3) = D3DColorARGB(.TimerI, 255, 255, 255)

                    If .Body.Walk(.Heading).GrhIndex Then Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y)

                    If .Head.Head(.Heading).GrhIndex Then Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, colorz, False, x, y)

                    If .Casco.Head(.Heading).GrhIndex Then Call Draw_Grh(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, colorz, False, x, y)

                    If .Arma.WeaponWalk(.Heading).GrhIndex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y)

                    If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y)

                End If

            End If
        
            If .particle_count > 0 Then

                For i = 1 To .particle_count

                    If .particle_group(i) > 0 Then
                        Particle_Group_Render .particle_group(i), PixelOffsetX + .Body.HeadOffset.x + (32 / 2), PixelOffsetY

                    End If

                Next i

            End If
    
            'Barra de tiempo
            If .BarTime < .MaxBarTime Then
                Draw_FilledBox PixelOffsetX - 17, PixelOffsetY - 40, 70, 7, D3DColorARGB(100, 0, 0, 0), D3DColorARGB(100, 0, 0, 0)
                Draw_FilledBox PixelOffsetX - 17, PixelOffsetY - 40, (((.BarTime / 100) / (.MaxBarTime / 100))) * 69, 7, D3DColorARGB(100, 200, 0, 0), D3DColorARGB(1, 200, 200, 200)
                .BarTime = .BarTime + (4 * timerTicksPerFrame * Sgn(1))
                                 
                If .BarTime >= .MaxBarTime Then
                    If charindex = UserCharIndex Then
                        Call CompletarAccionBarra(.BarAccion)

                    End If

                    charlist(charindex).BarTime = 0
                    charlist(charindex).BarAccion = 99
                    charlist(charindex).MaxBarTime = 0

                End If

            End If
                            
            If .Escribiendo = True And Not .invisible Then

                Dim TempGrh As grh

                TempGrh.framecounter = 1
                TempGrh.GrhIndex = 32017
                colorz(0) = D3DColorARGB(200, 255, 255, 255)
                colorz(1) = D3DColorARGB(200, 255, 255, 255)
                colorz(2) = D3DColorARGB(200, 255, 255, 255)
                colorz(3) = D3DColorARGB(200, 255, 255, 255)
                Call Draw_Grh(TempGrh, PixelOffsetX + 20, PixelOffsetY - 45, 1, 0, colorz, False, 0, 0, 0)

            End If
                             
            If .FxCount > 0 Then

                For i = 1 To .FxCount

                    If .FxList(i).FxIndex > 0 And .FxList(i).Started <> 0 Then
                        colorz(0) = D3DColorARGB(220, 255, 255, 255)
                        colorz(1) = D3DColorARGB(220, 255, 255, 255)
                        colorz(2) = D3DColorARGB(220, 255, 255, 255)
                        colorz(3) = D3DColorARGB(220, 255, 255, 255)

                        If FxData(.FxList(i).FxIndex).IsPNG = 1 Then
                            Call Draw_GrhFX(.FxList(i), PixelOffsetX + FxData(.FxList(i).FxIndex).OffsetX, PixelOffsetY + FxData(.FxList(i).FxIndex).OffsetY + 20, 1, 1, colorz, False, , , , charindex)
                        Else
                            Call Draw_GrhFX(.FxList(i), PixelOffsetX + FxData(.FxList(i).FxIndex).OffsetX, PixelOffsetY + FxData(.FxList(i).FxIndex).OffsetY + 20, 1, 1, colorz, True, , , , charindex)

                        End If

                    End If

                    If .FxList(i).Started = 0 Then
                        .FxList(i).FxIndex = 0

                    End If

                Next i

                If .FxList(.FxCount).Started = 0 Then
                    .FxCount = .FxCount - 1

                End If

            End If
            
            ' Meditación
            If .FxIndex <> 0 And .fX.Started <> 0 Then
                colorz(0) = D3DColorARGB(180, 255, 255, 255)
                colorz(1) = D3DColorARGB(180, 255, 255, 255)
                colorz(2) = D3DColorARGB(180, 255, 255, 255)
                colorz(3) = D3DColorARGB(180, 255, 255, 255)

                Call Draw_GrhFX(.fX, PixelOffsetX + FxData(.FxIndex).OffsetX, PixelOffsetY + FxData(.FxIndex).OffsetY + 4, 1, 1, colorz, False, , , , charindex)
           
            End If

        End If

    End With

End Sub

Private Sub Char_RenderCiego(ByVal charindex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer, ByVal x As Byte, ByVal y As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 12/03/04
    'Draw char's to screen without offcentering them
    '***************************************************
    Dim moved         As Boolean

    Dim Pos           As Integer

    Dim line          As String

    Dim color(0 To 3) As Long

    Dim i             As Long
    
    With charlist(charindex)

        If .Heading = 0 Then Exit Sub
    
        If .Moving Then

            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame * .Speeding
                
                'Start animations
                'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).speed > 0 Then .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0

                End If

            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame * .Speeding
                
                'Start animations
                If .Body.Walk(.Heading).speed > 0 Then .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0

                End If

            End If

        End If
        
        'If done moving stop animation
        If Not moved Then
            'Stop animations
            .Body.Walk(.Heading).Started = 0
            .Body.Walk(.Heading).framecounter = 1
            
            .Escudo.ShieldWalk(.Heading).Started = 0
            .Escudo.ShieldWalk(.Heading).framecounter = 1
            
            .Moving = False

        End If

        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
        
        Dim ColorCiego(0 To 3) As Long

        ColorCiego(0) = D3DColorARGB(255, 30, 30, 30)
        ColorCiego(1) = ColorCiego(0)
        ColorCiego(2) = ColorCiego(0)
        ColorCiego(3) = ColorCiego(0)
        
        If .Body.Walk(.Heading).GrhIndex Then
        
            If Not .invisible Then
 
                Dim colorz(3) As Long

                'Draw Body
                If .MUERTO = True Then
                    If .TimerM = 0 Then .TimerAct = True
                    If .TimerAct = False Then
                        .TimerM = .TimerM - 2
                    Else
                        .TimerM = .TimerM + 2

                        If .TimerM = 254 Then .TimerAct = False

                    End If
                    
                    colorz(0) = ColorCiego(0)
                    colorz(1) = ColorCiego(0)
                    colorz(2) = ColorCiego(0)
                    colorz(3) = ColorCiego(0)
                    
                Else
                    colorz(0) = ColorCiego(0)
                    colorz(1) = ColorCiego(0)
                    colorz(2) = ColorCiego(0)
                    colorz(3) = ColorCiego(0)

                End If
                        
                If .Body_Aura <> "" Then Call Renderizar_AuraCiego(.Body_Aura, PixelOffsetX, PixelOffsetY, x, y)
                If .Arma_Aura <> "" Then Call Renderizar_AuraCiego(.Arma_Aura, PixelOffsetX, PixelOffsetY, x, y)
                If .Otra_Aura <> "" Then Call Renderizar_AuraCiego(.Otra_Aura, PixelOffsetX, PixelOffsetY, x, y)
                If .Escudo_Aura <> "" Then Call Renderizar_AuraCiego(.Escudo_Aura, PixelOffsetX, PixelOffsetY, x, y)

                If .Heading = EAST Or .Heading = NORTH Then
                    If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y)

                    If .Body.Walk(.Heading).GrhIndex Then
                        Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y, 0)

                    End If

                Else

                    If .Heading = WEST Then
                        If .Arma.WeaponWalk(.Heading).GrhIndex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y)

                    End If

                    If .Body.Walk(.Heading).GrhIndex Then
                        If .iBody < 488 Then
                            Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX + 1, PixelOffsetY, 1, 1, colorz, False, x, y, 0)
                        Else
                            Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y, 0)

                        End If

                    End If
                            
                    If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y)

                End If
                            
                If .Head.Head(.Heading).GrhIndex Then

                    Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, colorz, False, x, y)
                    'Else
                    ' Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X - 1, PixelOffsetY + .Body.HeadOffset.Y, 1, 0, colorz, False, X, Y)
                                
                    ' End If
                End If

                If .Casco.Head(.Heading).GrhIndex Then Call Draw_Grh(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, colorz, False, x, y)
                            
                If .Heading <> WEST Then
                    If .Arma.WeaponWalk(.Heading).GrhIndex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y)

                End If

                'Draw name over head
                '  If .transformado = False Then
                    
                ' End If
            Else

                If charindex = UserCharIndex Or charlist(UserCharIndex).priv > 0 And .priv >= 0 Then
                    colorz(0) = ColorCiego(0)
                    colorz(1) = ColorCiego(0)
                    colorz(2) = ColorCiego(0)
                    colorz(3) = ColorCiego(0)

                    If .Body.Walk(.Heading).GrhIndex Then Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y)

                    If .Head.Head(.Heading).GrhIndex Then Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, colorz, False, x, y)

                    If .Casco.Head(.Heading).GrhIndex Then Call Draw_Grh(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, colorz, False, x, y)

                    If .Arma.WeaponWalk(.Heading).GrhIndex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y)

                    If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y)
          
                End If

            End If
        
            If .particle_count > 0 Then

                For i = 1 To .particle_count

                    If .particle_group(i) > 0 Then
            
                        Particle_Group_Render .particle_group(i), PixelOffsetX + .Body.HeadOffset.x + (32 / 2), PixelOffsetY

                    End If

                Next i

            End If
    
            'Barra de tiempo
            If .BarTime < .MaxBarTime Then
                Draw_FilledBox PixelOffsetX - 17, PixelOffsetY - 40, 70, 7, D3DColorARGB(100, 0, 0, 0), D3DColorARGB(100, 0, 0, 0)
                Draw_FilledBox PixelOffsetX - 17, PixelOffsetY - 40, (((.BarTime / 100) / (.MaxBarTime / 100))) * 69, 7, D3DColorARGB(100, 200, 0, 0), D3DColorARGB(1, 200, 200, 200)
                .BarTime = .BarTime + (4 * timerTicksPerFrame * Sgn(1))
                                 
                '  Engine_Text_Render "time: " & .BarTime, 50, 50, color, 1, True
                If .BarTime >= .MaxBarTime And charindex = UserCharIndex Then
                    Call CompletarAccionBarra(.BarAccion)
                                
                End If

            End If
                            
            If .Escribiendo = True Then
                            
                Dim cartelito(0 To 3) As Long

                Dim rgb_list(3)       As Long

                Dim TempGrh           As grh

                TempGrh.framecounter = 1
                TempGrh.GrhIndex = 32017
                cartelito(0) = D3DColorARGB(200, 255, 255, 255)
                cartelito(1) = rgb_list(0)
                cartelito(2) = rgb_list(0)
                cartelito(3) = rgb_list(0)
    
                Call Draw_Grh(TempGrh, PixelOffsetX + 20, PixelOffsetY - 45, 1, 0, cartelito, False, 0, 0, 0)
                            
            End If
          
            'Draw FX

            If .FxIndex <> 0 And .fX.Started <> 0 Then
                colorz(0) = D3DColorARGB(220, 255, 255, 255)
                colorz(1) = D3DColorARGB(220, 255, 255, 255)
                colorz(2) = D3DColorARGB(220, 255, 255, 255)
                colorz(3) = D3DColorARGB(220, 255, 255, 255)

                If FxData(.FxIndex).IsPNG = 1 Then
                    Call Draw_GrhFX(.fX, PixelOffsetX + FxData(.FxIndex).OffsetX, PixelOffsetY + FxData(.FxIndex).OffsetY + 20, 1, 1, colorz, False, , , , charindex)
                Else
                    Call Draw_GrhFX(.fX, PixelOffsetX + FxData(.FxIndex).OffsetX, PixelOffsetY + FxData(.FxIndex).OffsetY + 20, 1, 1, colorz, True, , , , charindex)

                End If
                    
                If .fX.Started = 0 Then .FxIndex = 0
           
            End If
        
        End If

    End With

End Sub

Private Sub Char_TextRender(ByVal charindex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer, ByVal x As Byte, ByVal y As Byte)

    Dim moved         As Boolean

    Dim Pos           As Integer

    Dim line          As String

    Dim color(0 To 3) As Long

    Dim i             As Long
    
    Dim screen_x      As Integer

    Dim screen_y      As Integer
    
    ' screen_x = Convert_Tile_To_View_X(PixelOffsetX) + MoveOffsetX
    ' screen_y = Convert_Tile_To_View_Y(PixelOffsetY) +

    With charlist(charindex)

        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
        
        'screen_x = Convert_Tile_To_View_X(PixelOffsetX) + MoveOffsetX

        '*** Start Dialogs ***
        If charlist(MapData(x, y).charindex).dialog <> "" Then

            'Figure out screen position
            Dim temp_array(3) As Long

            Dim PixelY        As Integer

            PixelY = PixelOffsetY
            temp_array(0) = charlist(MapData(x, y).charindex).dialog_color
            temp_array(1) = temp_array(0)
            temp_array(2) = temp_array(0)
            temp_array(3) = temp_array(0)

            If charlist(MapData(x, y).charindex).dialog_scroll Then
                charlist(MapData(x, y).charindex).dialog_offset_counter_y = charlist(MapData(x, y).charindex).dialog_offset_counter_y + (scroll_dialog_pixels_per_frame * timerTicksPerFrame * Sgn(-1))

                If Sgn(charlist(MapData(x, y).charindex).dialog_offset_counter_y) = -1 Then
                    charlist(MapData(x, y).charindex).dialog_offset_counter_y = 0
                    charlist(MapData(x, y).charindex).dialog_scroll = False

                End If

                Engine_Text_Render charlist(MapData(x, y).charindex).dialog, PixelOffsetX + 14 - CInt(Engine_Text_Width(charlist(MapData(x, y).charindex).dialog, True) / 2), PixelY + charlist(MapData(x, y).charindex).Body.HeadOffset.y - Engine_Text_Height(charlist(MapData(x, y).charindex).dialog, True) + charlist(MapData(x, y).charindex).dialog_offset_counter_y, temp_array, 1, True, MapData(x, y).charindex
            Else
                Engine_Text_Render charlist(MapData(x, y).charindex).dialog, PixelOffsetX + 14 - CInt(Engine_Text_Width(charlist(MapData(x, y).charindex).dialog, True) / 2), PixelY + charlist(MapData(x, y).charindex).Body.HeadOffset.y - Engine_Text_Height(charlist(MapData(x, y).charindex).dialog, True), temp_array, 1, True, MapData(x, y).charindex

            End If

        End If
        
        If charlist(MapData(x, y).charindex).dialogEfec <> "" Then

            charlist(MapData(x, y).charindex).SubeEfecto = charlist(MapData(x, y).charindex).SubeEfecto - timerTicksPerFrame
            charlist(MapData(x, y).charindex).dialog_Efect_color.a = charlist(MapData(x, y).charindex).dialog_Efect_color.a - (timerTicksPerFrame * 8.2)

            If charlist(MapData(x, y).charindex).dialog_Efect_color.a < 0 Then
                charlist(MapData(x, y).charindex).SubeEfecto = 0
                charlist(MapData(x, y).charindex).dialogEfec = ""
            Else
                temp_array(0) = D3DColorARGB(.dialog_Efect_color.a, .dialog_Efect_color.r, .dialog_Efect_color.g, .dialog_Efect_color.b)
                temp_array(1) = temp_array(0)
                temp_array(2) = temp_array(0)
                temp_array(3) = temp_array(0)
        
                Engine_Text_Render_Efect MapData(x, y).charindex, .dialogEfec, PixelOffsetX + 14 - Engine_Text_Width(.dialogEfec, True) / 2, PixelOffsetY - 100 + .Body.HeadOffset.y - Engine_Text_Height(.dialogEfec, True) + .SubeEfecto, temp_array, 1, True, max(CDbl(charlist(MapData(x, y).charindex).dialog_Efect_color.a), 0)

            End If

        End If
            
        ' If charlist(MapData(X, Y).charindex).dialogExp <> "" Then
    
        '  charlist(MapData(X, Y).charindex).SubeExp = charlist(MapData(X, Y).charindex).SubeExp + (5 * timerTicksPerFrame * Sgn(-1))
        ' If charlist(MapData(X, Y).charindex).SubeExp <= 5 Then
        '   charlist(MapData(X, Y).charindex).SubeExp = 0
        '  charlist(MapData(X, Y).charindex).dialogExp = ""
        'End If
                    
        'temp_array(0) = D3DColorARGB(charlist(MapData(X, Y).charindex).SubeExp, 42, 169, 222)
        ' temp_array(1) = temp_array(0)
        ' temp_array(2) = temp_array(0)
        ' temp_array(3) = temp_array(0)
        'Engine_Text_Render_Exp MapData(X, Y).charindex, .dialogExp, PixelOffsetX + 14 - Engine_Text_Width(.dialogExp, True) / 2, PixelOffsetY + 14 + .Body.HeadOffset.Y - Engine_Text_Height(.dialogExp, True), temp_array, 1, True
        ' End If
            
        'If charlist(MapData(X, Y).charindex).dialogOro <> "" Then

        '  charlist(MapData(X, Y).charindex).SubeOro = charlist(MapData(X, Y).charindex).SubeOro + (5 * timerTicksPerFrame * Sgn(-1))
                
        'If charlist(MapData(X, Y).charindex).SubeOro <= 5 Then
        '    charlist(MapData(X, Y).charindex).SubeOro = 0
        '    charlist(MapData(X, Y).charindex).dialogOro = ""
        'End If
                
        ' temp_array(0) = D3DColorARGB(charlist(MapData(X, Y).charindex).SubeOro, 255, 255, 115)
        ' temp_array(1) = temp_array(0)
        ' temp_array(2) = temp_array(0)
        ' temp_array(3) = temp_array(0)
        ' Engine_Text_Render_Exp MapData(X, Y).charindex, .dialogOro, PixelOffsetX + 14 - Engine_Text_Width(.dialogOro, True) / 2, PixelOffsetY + 1 + .Body.HeadOffset.Y - Engine_Text_Height(.dialogOro, True), temp_array, 1, True
                
        '  End If
        '*** End Dialogs ***
    End With

End Sub

Public Sub DrawMainInventory()

    ' Sólo dibujamos cuando es necesario
    If Not frmmain.Inventario.NeedsRedraw Then Exit Sub

    Dim InvRect As RECT

    InvRect.Left = 0
    InvRect.Top = 0
    InvRect.Right = frmmain.picInv.ScaleWidth
    InvRect.bottom = frmmain.picInv.ScaleHeight

    ' Comenzamos la escena
    Call Engine_BeginScene

    ' Dibujamos el fondo del inventario principal
    'Call Draw_GrhIndex(6, 0, 0)

    ' Dibujamos items
    Call frmmain.Inventario.DrawInventory
    
    ' Dibujamos item arrastrado
    Call frmmain.Inventario.DrawDraggedItem

    ' Presentamos la escena
    Call Engine_EndScene(InvRect, frmmain.picInv.hwnd)

End Sub

Public Sub DrawInterfaceComerciar()

    ' Sólo dibujamos cuando es necesario
    If Not frmComerciar.InvComNpc.NeedsRedraw And Not frmComerciar.InvComUsu.NeedsRedraw Then Exit Sub

    Dim InvRect As RECT

    InvRect.Left = 0
    InvRect.Top = 0
    InvRect.Right = frmComerciar.interface.ScaleWidth
    InvRect.bottom = frmComerciar.interface.ScaleHeight

    ' Comenzamos la escena
    Call Engine_BeginScene

    ' Dibujamos el fondo del inventario de comercio
    Call Draw_GrhIndex(837, 0, 0)

    ' Dibujamos items del NPC
    Call frmComerciar.InvComNpc.DrawInventory
    
    ' Dibujamos items del usuario
    Call frmComerciar.InvComUsu.DrawInventory

    ' Dibujamos "ambos" items arrastrados (aunque sólo puede estar uno activo a la vez)
    Call frmComerciar.InvComNpc.DrawDraggedItem
    Call frmComerciar.InvComUsu.DrawDraggedItem
    
    ' Me fijo qué inventario está seleccionado
    Dim CurrentInventory As clsGrapchicalInventory
    
    Dim cantidad         As Integer

    If frmComerciar.InvComNpc.SelectedItem > 0 Then
        Set CurrentInventory = frmComerciar.InvComNpc
        ' Al comprar, calculamos el valor según la cantidad exacta que ingresó
        cantidad = Val(frmComerciar.cantidad.Text)
    ElseIf frmComerciar.InvComUsu.SelectedItem > 0 Then
        Set CurrentInventory = frmComerciar.InvComUsu
        ' Al vender, calculamos el valor según el min(cantidad_ingresada, cantidad_items)
        cantidad = min(Val(frmComerciar.cantidad.Text), CurrentInventory.Amount(CurrentInventory.SelectedItem))

    End If
    
    ' Si hay alguno seleccionado
    If Not CurrentInventory Is Nothing Then
        ' Dibujo el item seleccionado
        'Call Draw_GrhColor(CurrentInventory.GrhIndex(CurrentInventory.SelectedItem), 282, 251, COLOR_WHITE)
    
        ' Muestro info del item
        Dim str As String

        str = " (No usa: "
        
        Select Case CurrentInventory.PuedeUsar(CurrentInventory.SelectedItem)

            Case 1
                str = str & "Genero)"

            Case 2
                str = str & "Clase)"

            Case 3
                str = str & "Facción)"

            Case 4
                str = str & "Skill)"

            Case 5
                str = str & "Raza)"

            Case 6
                str = str & "Nivel)"

            Case 0
                str = " (Usable)"

        End Select
                           
        frmComerciar.lblnombre = CurrentInventory.ItemName(CurrentInventory.SelectedItem) & str
        frmComerciar.lbldesc = CurrentInventory.GetInfo(CurrentInventory.OBJIndex(CurrentInventory.SelectedItem))
        frmComerciar.lblCosto = PonerPuntos(CLng(CurrentInventory.Valor(CurrentInventory.SelectedItem) * cantidad))
        
        Set CurrentInventory = Nothing

    End If

    ' Presentamos la escena
    Call Engine_EndScene(InvRect, frmComerciar.interface.hwnd)

End Sub

Public Sub DrawInterfaceBoveda()

    ' Sólo dibujamos cuando es necesario
    If Not frmBancoObj.InvBoveda.NeedsRedraw And Not frmBancoObj.InvBankUsu.NeedsRedraw Then Exit Sub

    Dim InvRect As RECT

    InvRect.Left = 0
    InvRect.Top = 0
    InvRect.Right = frmBancoObj.interface.ScaleWidth
    InvRect.bottom = frmBancoObj.interface.ScaleHeight

    ' Comenzamos la escena
    Call Engine_BeginScene

    ' Dibujamos el fondo de la bóveda
    Call Draw_GrhIndex(838, 0, 0)

    ' Dibujamos items de la bóveda
    Call frmBancoObj.InvBoveda.DrawInventory
    
    ' Dibujamos items del usuario
    Call frmBancoObj.InvBankUsu.DrawInventory

    ' Dibujamos "ambos" items arrastrados (aunque sólo puede estar uno activo a la vez)
    Call frmBancoObj.InvBoveda.DrawDraggedItem
    Call frmBancoObj.InvBankUsu.DrawDraggedItem
    
    ' Me fijo qué inventario está seleccionado
    Dim CurrentInventory As clsGrapchicalInventory
    
    If frmBancoObj.InvBoveda.SelectedItem > 0 Then
        Set CurrentInventory = frmBancoObj.InvBoveda
    ElseIf frmBancoObj.InvBankUsu.SelectedItem > 0 Then
        Set CurrentInventory = frmBancoObj.InvBankUsu

    End If
    
    ' Si hay alguno seleccionado
    If Not CurrentInventory Is Nothing Then

        ' Muestro info del item
        Dim str As String

        str = " (No usa: "
        
        Select Case CurrentInventory.PuedeUsar(CurrentInventory.SelectedItem)

            Case 1
                str = str & "Genero)"

            Case 2
                str = str & "Clase)"

            Case 3
                str = str & "Facción)"

            Case 4
                str = str & "Skill)"

            Case 5
                str = str & "Raza)"

            Case 6
                str = str & "Nivel)"

            Case 0
                str = " (Usable)"

        End Select
        
        frmBancoObj.lblnombre.Caption = CurrentInventory.ItemName(CurrentInventory.SelectedItem) & str
        frmBancoObj.lbldesc.Caption = CurrentInventory.GetInfo(CurrentInventory.OBJIndex(CurrentInventory.SelectedItem))
        
        Set CurrentInventory = Nothing

    End If

    ' Presentamos la escena
    Call Engine_EndScene(InvRect, frmBancoObj.interface.hwnd)

End Sub

Public Sub DrawMapaMundo()

    On Error Resume Next

    Static re          As RECT

    Static rgb_list(3) As Long

    re.Left = 0
    re.Top = 0
    re.bottom = 89
    re.Right = 177
    
    frmMapaGrande.PlayerView.Height = 89
    frmMapaGrande.PlayerView.Width = 177
    frmMapaGrande.PlayerView.ScaleHeight = 89
    frmMapaGrande.PlayerView.ScaleWidth = 177
    
    Call Engine_BeginScene
        
    Dim color(0 To 3) As Long

    color(0) = D3DColorARGB(255, 255, 255, 255)
    color(1) = color(0)
    color(2) = color(0)
    color(3) = color(0)
        
    Dim i    As Byte

    Dim x    As Integer

    Dim y    As Integer
    
    Dim Head As grh

    Head = HeadData(NpcData(frmMapaGrande.ListView1.SelectedItem.SubItems(2)).Head).Head(3)
    
    Dim grh As grh

    grh = BodyData(NpcData(frmMapaGrande.ListView1.SelectedItem.SubItems(2)).Body).Walk(3)
    
    Dim tmp           As String

    Dim temp_array(3) As Long 'Si le queres dar color a la letra pasa este parametro dsp xD

    Engine_Draw_Box x, y, 177, 89, D3DColorARGB(255, 7, 7, 7) 'Fondo del inventario
    
    x = frmMapaGrande.PlayerView.ScaleWidth / 2 - GrhData(grh.GrhIndex).pixelWidth / 2
    y = frmMapaGrande.PlayerView.ScaleHeight / 2 - GrhData(grh.GrhIndex).pixelHeight / 2
    Call Draw_Grh(grh, x, y, 0, 0, color, False, 0, 0, 0)

    x = frmMapaGrande.PlayerView.ScaleWidth / 2 - GrhData(Head.GrhIndex).pixelWidth / 2
    y = frmMapaGrande.PlayerView.ScaleHeight / 2 - GrhData(Head.GrhIndex).pixelHeight + 8 + BodyData(NpcData(frmMapaGrande.ListView1.SelectedItem.SubItems(2)).Body).HeadOffset.y / 2
    Call Draw_Grh(Head, x, y, 0, 0, color, False, 0, 0, 0)
    
    Call Engine_EndScene(re, frmMapaGrande.PlayerView.hwnd)

End Sub

Private Sub Renderizar_Aura(ByVal aura_index As String, ByVal x As Integer, ByVal y As Integer, ByVal map_x As Byte, ByVal map_y As Byte, Optional ByVal userindex As Long = 0)

    Dim rgb_list(0 To 3) As Long

    Dim i                As Byte

    Dim Index            As Long

    Dim color            As Long

    Dim aura_grh         As grh

    Dim TRANS            As Integer

    Dim giro             As Single

    Dim lado             As Byte

    Index = Val(ReadField(1, aura_index, Asc(":")))
    color = Val(ReadField(2, aura_index, Asc(":")))
    giro = Val(ReadField(3, aura_index, Asc(":")))
    lado = Val(ReadField(4, aura_index, Asc(":")))

    'Debug.Print charlist(userindex).AuraAngle
    If giro > 0 And userindex > 0 Then
        'If lado = 0 Then
        charlist(userindex).AuraAngle = charlist(userindex).AuraAngle + (timerTicksPerFrame * giro)
        'Else
        'charlist(userindex).AuraAngle = charlist(userindex).AuraAngle - (timerTicksPerFrame * giro)
        ' End If
    
        If charlist(userindex).AuraAngle >= 360 Then charlist(userindex).AuraAngle = 0

    End If

    'If charlist(userindex).AuraAngle <> 0 Then
    'Debug.Print charlist(userindex).AuraAngle
    'End If
    Dim r As Integer

    Dim g As Integer

    Dim b As Integer

    r = &HFF& And color
    g = (&HFF00& And color) \ 256
    b = (&HFF0000 And color) \ 65536
    TRANS = 255

    rgb_list(0) = D3DColorARGB(TRANS, b, g, r)
    rgb_list(1) = D3DColorARGB(TRANS, b, g, r)
    rgb_list(2) = D3DColorARGB(TRANS, b, g, r)
    rgb_list(3) = D3DColorARGB(TRANS, b, g, r)

    'Convertimos el Aura en un GRH
    Call InitGrh(aura_grh, Index)

    'Y por ultimo renderizamos esta capa con Draw_Grh
    If giro > 0 And userindex > 0 Then
        Call Draw_Grh(aura_grh, x, y + 30, 1, 0, rgb_list(), True, map_x, map_y, charlist(userindex).AuraAngle)
    Else
        Call Draw_Grh(aura_grh, x, y + 30, 1, 0, rgb_list(), True, map_x, map_y, 0)

    End If
    
End Sub

Private Sub Renderizar_AuraCiego(ByVal aura_index As String, ByVal x As Integer, ByVal y As Integer, ByVal map_x As Byte, ByVal map_y As Byte)

    Dim rgb_list(0 To 3) As Long

    Dim i                As Byte

    Dim Index            As Long

    Dim color            As Long

    Dim aura_grh         As grh

    Dim TRANS            As Integer

    Index = Val(ReadField(1, aura_index, Asc(":")))
    color = Val(ReadField(2, aura_index, Asc(":")))
    TRANS = 1 'Val(ReadField(4, aura_index, Asc(":")))

    Dim r As Integer

    Dim g As Integer

    Dim b As Integer

    r = &HFF& And color
    g = (&HFF00& And color) \ 256
    b = (&HFF0000 And color) \ 65536

    Dim ColorCiego(0 To 3) As Long

    ColorCiego(0) = D3DColorARGB(255, 30, 30, 30)
    ColorCiego(1) = ColorCiego(0)
    ColorCiego(2) = ColorCiego(0)
    ColorCiego(3) = ColorCiego(0)

    rgb_list(0) = ColorCiego(0)
    rgb_list(1) = ColorCiego(0)
    rgb_list(2) = ColorCiego(0)
    rgb_list(3) = ColorCiego(0)

    'Convertimos el Aura en un GRH
    Call InitGrh(aura_grh, Index)
    'Y por ultimo renderizamos esta capa con Draw_Grh
    Call Draw_Grh(aura_grh, x, y + 30, 1, 0, rgb_list(), True, map_x, map_y)
    
End Sub

Public Sub RenderConnect(ByVal tilex As Integer, ByVal tiley As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)

    Call Engine_BeginScene

    Dim y                As Integer     'Keeps track of where on map we are

    Dim x                As Integer     'Keeps track of where on map we are

    Dim screenminY       As Integer  'Start Y pos on current screen

    Dim screenmaxY       As Integer  'End Y pos on current screen

    Dim screenminX       As Integer  'Start X pos on current screen

    Dim screenmaxX       As Integer  'End X pos on current screen

    Dim minY             As Integer  'Start Y pos on current map

    Dim MaxY             As Integer  'End Y pos on current map

    Dim minX             As Integer  'Start X pos on current map

    Dim MaxX             As Integer  'End X pos on current map

    Dim ScreenX          As Integer  'Keeps track of where to place tile on screen

    Dim ScreenY          As Integer  'Keeps track of where to place tile on screen

    Dim minXOffset       As Integer

    Dim minYOffset       As Integer

    Dim PixelOffsetXTemp As Integer 'For centering grhs

    Dim PixelOffsetYTemp As Integer 'For centering grhs

    Dim CurrentGrhIndex  As Integer

    Dim OffX             As Integer

    Dim Offy             As Integer

    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    minY = screenminY - TileBufferSize
    MaxY = screenmaxY + TileBufferSize
    minX = screenminX - TileBufferSize
    MaxX = screenmaxX + TileBufferSize
    
    'Make sure mins and maxs are allways in map bounds
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize

    End If
    
    If MaxY > YMaxMapSize Then MaxY = YMaxMapSize
    
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize

    End If
    
    If MaxX > XMaxMapSize Then MaxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1

    End If
    
    If screenmaxY < YMaxMapSize Then screenmaxY = screenmaxY + 1
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1

    End If
    
    If screenmaxX < XMaxMapSize Then screenmaxX = screenmaxX + 1
    
    If screenminY < 1 Then screenminY = 1
    If screenminX < 1 Then screenminX = 1
    If screenmaxY > 100 Then screenmaxY = 100
    If screenmaxX > 100 Then screenmaxX = 100
    
    screenmaxY = screenmaxY + 9
    screenmaxX = screenmaxY + 9
  
    'Draw floor layer
    For y = screenminY To screenmaxY
        For x = screenminX To screenmaxX
            'Layer 1 **********************************
            Call Draw_Grh(MapData(x, y).Graphic(1), (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, 0, 1, MapData(x, y).light_value, , x, y)
            '******************************************
            ScreenX = ScreenX + 1
        Next x

        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - x + screenminX
        ScreenY = ScreenY + 1
    Next y
    
    If HayLayer2 Then
        ScreenY = minYOffset - TileBufferSize

        For y = minY To MaxY
            ScreenX = minXOffset - TileBufferSize

            For x = minX To MaxX + 2

                With MapData(x, y)

                    '***********************************************
                    If MapData(x, y).Graphic(2).GrhIndex <> 0 Then
                        Call Draw_Grh(MapData(x, y).Graphic(2), (ScreenX * 32 + PixelOffsetX), (ScreenY * 32 + PixelOffsetY), 1, 1, MapData(x, y).light_value(), , x, y)

                    End If
              
                End With

                ScreenX = ScreenX + 1
            Next x

            ScreenY = ScreenY + 1
        Next y

    End If
    
    ScreenY = minYOffset - TileBufferSize

    For y = minY To MaxY
        ScreenX = minXOffset - TileBufferSize

        For x = minX To MaxX
            PixelOffsetXTemp = ScreenX * 32 + PixelOffsetX
            PixelOffsetYTemp = ScreenY * 32 + PixelOffsetY
            
            With MapData(x, y)
                '******************************************

                'Object Layer **********************************
                If MapData(x, y).ObjGrh.GrhIndex <> 0 Then
                    Call Draw_Grh(MapData(x, y).ObjGrh, (ScreenX * 32 + PixelOffsetX), (ScreenY * 32 + PixelOffsetY), 1, 1, MapData(x, y).light_value(), , x, y)

                End If
             
                'Layer 3 *****************************************
                If .Graphic(3).GrhIndex <> 0 Then
                    Call Draw_Grh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(x, y).light_value, False, x, y)

                End If

                '************************************************

            End With
            
            ScreenX = ScreenX + 1
        Next x

        ScreenY = ScreenY + 1
    Next y

    ScreenY = minYOffset - 5
    
    Dim cc(3)   As Long

    Dim TempGrh As grh

    'nubes negras
    TempGrh.framecounter = 1
    TempGrh.GrhIndex = 1170
    cc(0) = D3DColorARGB(180, 255, 255, 255)
    cc(1) = cc(0)
    cc(2) = cc(0)
    cc(3) = cc(0)
    ' Draw_Grh TempGrh, 494, 735, 1, 1, cc(), False
    'nubes negras

    ScreenY = minYOffset - TileBufferSize

    For y = minY To MaxY
        ScreenX = minXOffset - TileBufferSize

        For x = minX To MaxX

            With MapData(x, y)

                '***********************************************
                If .particle_group > 0 Then
                    Call Particle_Group_Render(.particle_group, ScreenX * 32 + PixelOffsetX + 15, ScreenY * 32 + PixelOffsetY + 15)

                End If
          
            End With

            ScreenX = ScreenX + 1
        Next x

        ScreenY = ScreenY + 1
    Next y
 
    'Draw blocked tiles and grid
 
    If HayLayer4 Then

        Dim rgb_list(0 To 3) As Long
    
        ScreenY = minYOffset - TileBufferSize

        For y = minY To MaxY
            ScreenX = minXOffset - TileBufferSize

            For x = minX To MaxX
        
                If MapData(x, y).Graphic(4).GrhIndex Then

                    rgb_list(0) = D3DColorARGB(255, ColorAmbiente.r, ColorAmbiente.g, ColorAmbiente.b)
                    rgb_list(1) = rgb_list(0)
                    rgb_list(2) = rgb_list(0)
                    rgb_list(3) = rgb_list(0)
                        
                    Call Draw_Grh(MapData(x, y).Graphic(4), ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, 1, 1, rgb_list(), , x, y)
          
                End If
 
                '**********************************
                ScreenX = ScreenX + 1
            Next x

            ScreenY = ScreenY + 1
        Next y
        
    End If
        
    ScreenY = minYOffset - TileBufferSize

    For y = minY To MaxY
        ScreenX = minXOffset - TileBufferSize

        For x = minX To MaxX
            PixelOffsetXTemp = ScreenX * 32 + PixelOffsetX
            PixelOffsetYTemp = ScreenY * 32 + PixelOffsetY

            With MapData(x, y)
                
                If MapData(x, y).charindex <> 0 Then
                    If charlist(MapData(x, y).charindex).active = 1 Then
                        Call Char_TextRender(MapData(x, y).charindex, PixelOffsetXTemp, PixelOffsetYTemp, x, y)

                    End If

                End If

            End With
            
            ScreenX = ScreenX + 1
        Next x

        ScreenY = ScreenY + 1
    Next y

    ScreenY = minYOffset - 5
        
    Dim DefaultColor(3) As Long

    Dim color           As Long

    intro = 1

    If intro = 1 Then

        DefaultColor(0) = D3DColorXRGB(255, 255, 255)
        DefaultColor(1) = DefaultColor(0)
        DefaultColor(2) = DefaultColor(0)
        DefaultColor(3) = DefaultColor(0)
        '    Call Renderizar_Aura("35457:&HFF8000:0:0", 400 + 15, 310, 0, 0)
        Draw_Grh BodyData(640).Walk(3), 470 + 15, 366, 1, 0, DefaultColor()
        Draw_Grh HeadData(602).Head(3), 470 + 15, 327 + 2, 1, 0, DefaultColor()
            
        Draw_Grh CascoAnimData(48).Head(3), 470 + 15, 327, 1, 0, DefaultColor()
        Draw_Grh WeaponAnimData(82).WeaponWalk(3), 470 + 15, 366, 1, 0, DefaultColor()
            
        Engine_Text_Render_LetraChica "v" & App.Major & "." & App.Minor & " Build: " & App.Revision, 870, 750, DefaultColor, 4, False

        Dim ItemName As String

        'itemname = "abcdfghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789¡!¿TEST?#$100%&/\()=-@^[]<>*+.,:; pálmas séso te píso sólo púto ý LÁL LÉ"
            
        ' itemname = "pálmas séso te píso sólo púto ý lÁ Élefante PÍSÓS PÚTO ÑOño"
        Engine_Text_Render_LetraChica ItemName, 100, 730, DefaultColor, 4, False

        If ClickEnAsistente < 30 Then
            Call Particle_Group_Render(spell_particle, 500, 365)

        End If

    End If
 
    ScreenX = 250
    ScreenY = 0
    'Call Particle_Group_Render(meteo_particle, ScreenX, ScreenY)

    LastOffsetX = ParticleOffsetX
    LastOffsetY = ParticleOffsetY
    'Engine_Weather_UpdateFog
    
    TextEfectAsistente = TextEfectAsistente + (15 * timerTicksPerFrame * Sgn(-1))

    If TextEfectAsistente <= 1 Then
        TextEfectAsistente = 0

    End If

    Engine_Text_Render TextAsistente, 510 - Engine_Text_Width(TextAsistente, True, 1) / 2, 320 - Engine_Text_Height(TextAsistente, True) + TextEfectAsistente, textcolorAsistente, 1, True, , 200

    '
    ' Engine_Draw_Box 975, 5, 15, 15, D3DColorARGB(100, 70, 0, 0)
    'Engine_Text_Render UserCuenta, 490 - Engine_Text_Width(UserCuenta, False, 3) / 2, 38 - Engine_Text_Height(UserCuenta, False, 3), DefaultColor, 3, False
    ' Engine_Text_Render "X", 977, 5, DefaultColor, 1, False
    
    '   Engine_Draw_Box 955, 5, 15, 15, D3DColorARGB(100, 70, 0, 0)
    'Engine_Text_Render UserCuenta, 490 - Engine_Text_Width(UserCuenta, False, 3) / 2, 38 - Engine_Text_Height(UserCuenta, False, 3), DefaultColor, 3, False
    ' Engine_Text_Render "_", 957, 3, DefaultColor, 1, False

    'Logo viejo
    TempGrh.framecounter = 1
    TempGrh.GrhIndex = 1171

    cc(0) = D3DColorARGB(255, 255, 255, 255)
    cc(1) = cc(0)
    cc(2) = cc(0)
    cc(3) = cc(0)

    ' Draw_Grh TempGrh, 494, 200, 1, 1, cc(), False
    'Logo viejo

    'Logo viejo
    
    TempGrh.framecounter = 1
    TempGrh.GrhIndex = 1172

    cc(0) = D3DColorARGB(220, 255, 255, 255)
    cc(1) = cc(0)
    cc(2) = cc(0)
    cc(3) = cc(0)

    Draw_Grh TempGrh, 494, 275, 1, 1, cc(), False

    'Logo nuevo
    'Marco
    TempGrh.framecounter = 1
    TempGrh.GrhIndex = 1169

    cc(0) = D3DColorARGB(255, 255, 255, 255)
    cc(1) = cc(0)
    cc(2) = cc(0)
    cc(3) = cc(0)

    Draw_Grh TempGrh, 0, 0, 0, 0, cc(), False
    
    'Marco

    #If DEBUGGING = 1 Then
        ' Botones debug
        Engine_Text_Render "Debug:", 50, 300, DefaultColor
    
        ' Crear cuenta a manopla
        Engine_Draw_Box 40, 330, 155, 35, D3DColorARGB(150, 0, 0, 0)
        Engine_Text_Render "Crear cuenta en cliente", 50, 340, DefaultColor
    #End If

    'TempGrh.framecounter = 1
    'TempGrh.GrhIndex = 32016

    ' cc(0) = D3DColorARGB(255, 255, 255, 255)
    ' cc(1) = cc(0)
    ' cc(2) = cc(0)
    ' cc(3) = cc(0)

    ' Draw_Grh TempGrh, 480, 100, 1, 1, cc(), False
    Call Engine_EndScene(Render_Connect_Rect, frmConnect.render.hwnd)
    
    lFrameLimiter = (GetTickCount() And &H7FFFFFFF)
    'FramesPerSecCounter = FramesPerSecCounter + 1
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed

    Exit Sub

End Sub

Public Sub RenderCrearPJ(ByVal tilex As Integer, ByVal tiley As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)

    Call Engine_BeginScene

    Dim y                As Integer     'Keeps track of where on map we are

    Dim x                As Integer     'Keeps track of where on map we are

    Dim screenminY       As Integer  'Start Y pos on current screen

    Dim screenmaxY       As Integer  'End Y pos on current screen

    Dim screenminX       As Integer  'Start X pos on current screen

    Dim screenmaxX       As Integer  'End X pos on current screen

    Dim minY             As Integer  'Start Y pos on current map

    Dim MaxY             As Integer  'End Y pos on current map

    Dim minX             As Integer  'Start X pos on current map

    Dim MaxX             As Integer  'End X pos on current map

    Dim ScreenX          As Integer  'Keeps track of where to place tile on screen

    Dim ScreenY          As Integer  'Keeps track of where to place tile on screen

    Dim minXOffset       As Integer

    Dim minYOffset       As Integer

    Dim PixelOffsetXTemp As Integer 'For centering grhs

    Dim PixelOffsetYTemp As Integer 'For centering grhs

    Dim CurrentGrhIndex  As Integer

    Dim OffX             As Integer

    Dim Offy             As Integer

    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    minY = screenminY - TileBufferSize
    MaxY = screenmaxY + TileBufferSize
    minX = screenminX - TileBufferSize
    MaxX = screenmaxX + TileBufferSize
    
    'Make sure mins and maxs are allways in map bounds
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize

    End If
    
    If MaxY > YMaxMapSize Then MaxY = YMaxMapSize
    
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize

    End If
    
    If MaxX > XMaxMapSize Then MaxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1

    End If
    
    If screenmaxY < YMaxMapSize Then screenmaxY = screenmaxY + 1
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1

    End If
    
    If screenmaxX < XMaxMapSize Then screenmaxX = screenmaxX + 1
    
    If screenminY < 1 Then screenminY = 1
    If screenminX < 1 Then screenminX = 1
    If screenmaxY > 100 Then screenmaxY = 100
    If screenmaxX > 100 Then screenmaxX = 100
    screenmaxY = screenmaxY + 8
    screenmaxX = screenmaxY + 8

    'Draw floor layer
    For y = screenminY To screenmaxY
        For x = screenminX To screenmaxX
            'Layer 1 **********************************
            Call Draw_Grh(MapData(x, y).Graphic(1), (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, 0, 1, MapData(x, y).light_value, , x, y)
            '******************************************
            ScreenX = ScreenX + 1
        Next x

        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - x + screenminX
        ScreenY = ScreenY + 1
    Next y
    
    If HayLayer2 Then
        ScreenY = minYOffset - TileBufferSize

        For y = minY To MaxY
            ScreenX = minXOffset - TileBufferSize

            For x = minX To MaxX

                With MapData(x, y)

                    '***********************************************
                    If MapData(x, y).Graphic(2).GrhIndex <> 0 Then
                        Call Draw_Grh(MapData(x, y).Graphic(2), (ScreenX * 32 + PixelOffsetX), (ScreenY * 32 + PixelOffsetY), 1, 1, MapData(x, y).light_value(), , x, y)

                    End If
              
                End With

                ScreenX = ScreenX + 1
            Next x

            ScreenY = ScreenY + 1
        Next y

    End If
    
    ScreenY = minYOffset - TileBufferSize

    For y = minY To MaxY
        ScreenX = minXOffset - TileBufferSize

        For x = minX To MaxX
            PixelOffsetXTemp = ScreenX * 32 + PixelOffsetX
            PixelOffsetYTemp = ScreenY * 32 + PixelOffsetY
            
            With MapData(x, y)
                '******************************************

                'Object Layer **********************************
                If MapData(x, y).ObjGrh.GrhIndex <> 0 Then
                    Call Draw_Grh(MapData(x, y).ObjGrh, (ScreenX * 32 + PixelOffsetX), (ScreenY * 32 + PixelOffsetY), 1, 1, MapData(x, y).light_value(), , x, y)

                End If
             
                'Layer 3 *****************************************
                If .Graphic(3).GrhIndex <> 0 Then
                    Call Draw_Grh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(x, y).light_value, False, x, y)

                End If

                '************************************************

            End With
            
            ScreenX = ScreenX + 1
        Next x

        ScreenY = ScreenY + 1
    Next y

    ScreenY = minYOffset - 5

    ScreenY = minYOffset - TileBufferSize

    For y = minY To MaxY
        ScreenX = minXOffset - TileBufferSize

        For x = minX To MaxX

            With MapData(x, y)

                '***********************************************
                If .particle_group > 0 Then
                    Call Particle_Group_Render(.particle_group, ScreenX * 32 + PixelOffsetX + 15, ScreenY * 32 + PixelOffsetY + 15)

                End If
          
            End With

            ScreenX = ScreenX + 1
        Next x

        ScreenY = ScreenY + 1
    Next y
 
    'Draw blocked tiles and grid
 
    If HayLayer4 Then

        Dim rgb_list(0 To 3) As Long
    
        ScreenY = minYOffset - TileBufferSize

        For y = minY To MaxY
            ScreenX = minXOffset - TileBufferSize

            For x = minX To MaxX
        
                If MapData(x, y).Graphic(4).GrhIndex Then

                    rgb_list(0) = D3DColorARGB(255, ColorAmbiente.r, ColorAmbiente.g, ColorAmbiente.b)
                    rgb_list(1) = rgb_list(0)
                    rgb_list(2) = rgb_list(0)
                    rgb_list(3) = rgb_list(0)
                        
                    Call Draw_Grh(MapData(x, y).Graphic(4), ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, 1, 1, rgb_list(), , x, y)
          
                End If
 
                '**********************************
                ScreenX = ScreenX + 1
            Next x

            ScreenY = ScreenY + 1
        Next y
        
    End If

    Engine_Weather_UpdateFog

    RenderUICrearPJ

    Dim cc(3)   As Long

    Dim TempGrh As grh

    'Logo viejo
    TempGrh.framecounter = 1
    TempGrh.GrhIndex = 1171

    cc(0) = D3DColorARGB(255, 255, 255, 255)
    cc(1) = cc(0)
    cc(2) = cc(0)
    cc(3) = cc(0)

    Draw_Grh TempGrh, 494, 190, 1, 1, cc(), False
    'Logo viejo
    
    'Marco
    TempGrh.framecounter = 1
    TempGrh.GrhIndex = 1169

    cc(0) = D3DColorARGB(255, 255, 255, 255)
    cc(1) = cc(0)
    cc(2) = cc(0)
    cc(3) = cc(0)

    Draw_Grh TempGrh, 0, 0, 0, 0, cc(), False

    Call Engine_EndScene(Render_Connect_Rect, frmConnect.render.hwnd)

    lFrameLimiter = (GetTickCount() And &H7FFFFFFF)
    FramesPerSecCounter = FramesPerSecCounter + 1
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed

    'RenderPjsCuenta

End Sub

Public Sub rendercuenta(ByVal tilex As Integer, ByVal tiley As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)

    Call Engine_BeginScene

    lFrameLimiter = (GetTickCount() And &H7FFFFFFF)
    FramesPerSecCounter = FramesPerSecCounter + 1
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed

    RenderPjsCuenta
    
    Call Particle_Group_Render(ParticleLluviaDorada, 400, 0)

    Call Engine_EndScene(Render_Connect_Rect, frmConnect.render.hwnd)
    
    Exit Sub

End Sub

Public Sub RenderUICrearPJ()

    Dim TempGrh         As grh
    
    Dim DefaultColor(3) As Long
    
    TempGrh.framecounter = 1
    TempGrh.GrhIndex = 727
    DefaultColor(0) = D3DColorXRGB(255, 255, 255)
    DefaultColor(1) = DefaultColor(0)
    DefaultColor(2) = DefaultColor(0)
    DefaultColor(3) = DefaultColor(0)
    
    Draw_Grh TempGrh, 475, 545, 1, 1, DefaultColor(), False

    DefaultColor(0) = D3DColorXRGB(200, 200, 200)
    DefaultColor(1) = DefaultColor(0)
    DefaultColor(2) = DefaultColor(0)
    DefaultColor(3) = DefaultColor(0)
    
    'Engine_Text_Render "Nombre del personaje", 230 + -Engine_Text_Width("Nombre del personaje", False) / 2, 110 + 40 - Engine_Text_Height("Nombre del personaje", False), DefaultColor, 3, True
    'Engine_Text_Render "Creacion de personajes", 210, 120, DefaultColor, 3, False
    
    'Basico
    'Engine_Draw_Box 217, 183, 550, 386, D3DColorARGB(20, 219, 116, 3)
    
    'Engine_Draw_Box 250, 190, 490, 356, D3DColorARGB(50, 128, 128, 128)
    'Engine_Draw_Box 250, 190, 490, 356, D3DColorARGB(100, 0, 0, 0)
    
    'Engine_Draw_Box 220, 186, 550, 380, D3DColorARGB(80, 20, 27, 3)
    'Engine_Text_Render UserCuenta, 490 - Engine_Text_Width(UserCuenta, False, 3) / 2, 38 - Engine_Text_Height(UserCuenta, False, 3), DefaultColor, 3, False
    Engine_Text_Render "Creacion de Personaje", 280, 125, DefaultColor, 5, False

    'Engine_Draw_Box 400, 215, 180, 250, D3DColorARGB(200, 100, 100, 100)
    
    DefaultColor(0) = D3DColorXRGB(255, 255, 255)
    DefaultColor(1) = DefaultColor(0)
    DefaultColor(2) = DefaultColor(0)
    DefaultColor(3) = DefaultColor(0)
    
    Engine_Text_Render_LetraChica "Nombre ", 470, 198, DefaultColor, 6, False
    Engine_Text_Render_LetraChica "Clase ", 477, 240, DefaultColor, 6, False
    
    '
    
    Engine_Draw_Box 450, 255, 95, 21, D3DColorARGB(100, 1, 1, 1)
    
    Engine_Text_Render "<", 435, 258, DefaultColor, 1, False
        
    Engine_Text_Render ">", 548, 258, DefaultColor, 1, False
    'Engine_Text_Render ">", 403, 412, DefaultColor, 1, True
    
    DefaultColor(0) = D3DColorXRGB(200, 200, 200)
    DefaultColor(1) = DefaultColor(0)
    DefaultColor(2) = DefaultColor(0)
    DefaultColor(3) = DefaultColor(0)
    
    Engine_Text_Render frmCrearPersonaje.lstProfesion.List(frmCrearPersonaje.lstProfesion.ListIndex), 498 - Engine_Text_Width(frmCrearPersonaje.lstProfesion.List(frmCrearPersonaje.lstProfesion.ListIndex), True, 1) / 2, 258, DefaultColor, 1, True
    
    DefaultColor(0) = D3DColorXRGB(255, 255, 255)
    DefaultColor(1) = DefaultColor(0)
    DefaultColor(2) = DefaultColor(0)
    DefaultColor(3) = DefaultColor(0)
    
    Engine_Text_Render_LetraChica "Raza ", 481, 285, DefaultColor, 6, False
    Engine_Draw_Box 450, 302, 95, 21, D3DColorARGB(100, 1, 1, 1)
    
    DefaultColor(0) = D3DColorXRGB(200, 200, 200)
    DefaultColor(1) = DefaultColor(0)
    DefaultColor(2) = DefaultColor(0)
    DefaultColor(3) = DefaultColor(0)

    'Engine_Text_Render "Humano", 470 - Engine_Text_Height("Humano", False), 304, DefaultColor, 1, False
    Engine_Text_Render frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.ListIndex), 495 - Engine_Text_Width(frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.ListIndex), True, 1) / 2, 305, DefaultColor, 1, True
    
    Engine_Text_Render "<", 435, 305, DefaultColor, 1, False
    Engine_Text_Render ">", 548, 305, DefaultColor, 1, False
    
    DefaultColor(0) = D3DColorXRGB(255, 255, 255)
    DefaultColor(1) = DefaultColor(0)
    DefaultColor(2) = DefaultColor(0)
    DefaultColor(3) = DefaultColor(0)
    
    Engine_Text_Render_LetraChica "Genero ", 475, 330, DefaultColor, 6, False
    Engine_Draw_Box 450, 346, 95, 21, D3DColorARGB(100, 1, 1, 1)
        
    DefaultColor(0) = D3DColorXRGB(200, 200, 200)
    DefaultColor(1) = DefaultColor(0)
    DefaultColor(2) = DefaultColor(0)
    DefaultColor(3) = DefaultColor(0)
    
    Engine_Text_Render frmCrearPersonaje.lstGenero.List(frmCrearPersonaje.lstGenero.ListIndex), 495 - Engine_Text_Width(frmCrearPersonaje.lstGenero.List(frmCrearPersonaje.lstGenero.ListIndex), True, 1) / 2, 349, DefaultColor, 1, True
    
    Engine_Text_Render "<", 435, 350, DefaultColor, 1, False
    Engine_Text_Render ">", 548, 350, DefaultColor, 1, False
    
    DefaultColor(0) = D3DColorXRGB(200, 200, 50)
    DefaultColor(1) = DefaultColor(0)
    DefaultColor(2) = DefaultColor(0)
    DefaultColor(3) = DefaultColor(0)
    
    'Engine_Text_Render RazaRecomendada, 489 - Engine_Text_Width(RazaRecomendada, False, 1) / 2, 278, DefaultColor, 1, False
    
    If Len(RazaRecomendada) > 0 Then
        Engine_Text_Render "Raza sugerida:", 570, 290, DefaultColor, 4, False
        Engine_Text_Render RazaRecomendada, 570, 300, DefaultColor, 4, False

    End If
    
    '     DefaultColor(0) = D3DColorXRGB(255, 50, 50)
    '  DefaultColor(1) = DefaultColor(0)
    '  DefaultColor(2) = DefaultColor(0)
    '  DefaultColor(3) = DefaultColor(0)
    
    '' Engine_Text_Render "¡Atención! ", 240, 250, DefaultColor, 1, False
    '     DefaultColor(0) = D3DColorXRGB(255, 255, 255)
    '  DefaultColor(1) = DefaultColor(0)
    ' DefaultColor(2) = DefaultColor(0)
    ' DefaultColor(3) = DefaultColor(0)
    ' Engine_Text_Render "Se cuidadoso al momento de distribuir tus atributos. De estos dependen aspectos basicos como la vida o maná de tu personaje. ", 190, 270, DefaultColor, 1, True
    
    DefaultColor(0) = D3DColorXRGB(255, 255, 255)
    DefaultColor(1) = DefaultColor(0)
    DefaultColor(2) = DefaultColor(0)
    DefaultColor(3) = DefaultColor(0)
        
    Dim Offy As Long
     
    Offy = 2

    Dim OffX As Long
     
    OffX = 350
    
    'Atributos
    Engine_Text_Render_LetraChica "Atributos ", 240 + OffX, 385 + Offy, DefaultColor, 6, True
    Engine_Draw_Box 175 + OffX, 405 + Offy, 185, 120, D3DColorARGB(80, 0, 0, 0)
    '  Engine_Draw_Box 610, 405, 220, 180, D3DColorARGB(120, 100, 100, 100)
    
    Engine_Text_Render_LetraChica "Fuerza ", 185 + OffX, 410 + Offy, DefaultColor, 1, True
    ' Engine_Text_Render "<", 260, 410, DefaultColor, 1, True
    ' Engine_Text_Render ">", 310, 410, DefaultColor, 1, True
    Engine_Draw_Box 280 + OffX, 409 + Offy, 20, 20, D3DColorARGB(100, 1, 1, 1)
    Engine_Text_Render_LetraChica frmCrearPersonaje.lbFuerza.Caption, 282 + OffX, 413 + Offy, DefaultColor, 1, True ' Atributo fuerza
    'Engine_Text_Render "+", 335, 410, DefaultColor, 1, True
    Engine_Draw_Box 317 + OffX, 409 + Offy, 25, 20, D3DColorARGB(100, 1, 1, 1)
    Engine_Text_Render_LetraChica frmCrearPersonaje.modfuerza.Caption, 320 + OffX, 413 + Offy, DefaultColor, 1, True ' Bonificacion fuerza
    
    Engine_Text_Render "Agilidad ", 185 + OffX, 440 + Offy, DefaultColor, 1, True
    ' Engine_Text_Render "<", 260, 440, DefaultColor, 1, True
    ' Engine_Text_Render ">", 310, 440, DefaultColor, 1, True
    Engine_Draw_Box 280 + OffX, 440 + Offy, 20, 20, D3DColorARGB(100, 1, 1, 1)
    Engine_Text_Render frmCrearPersonaje.lbAgilidad.Caption, 282 + OffX, 443 + Offy, DefaultColor, 1, True ' Atributo Agilidad
    ' Engine_Text_Render "+", 335, 440, DefaultColor, 1, True
    Engine_Draw_Box 317 + OffX, 440 + Offy, 25, 20, D3DColorARGB(100, 1, 1, 1)
    Engine_Text_Render frmCrearPersonaje.modAgilidad.Caption, 320 + OffX, 443 + Offy, DefaultColor, 1, True ' Bonificacion Agilidad
    
    Engine_Text_Render "Inteligencia ", 185 + OffX, 470 + Offy, DefaultColor, 1, True
    'Engine_Text_Render "<", 260, 470, DefaultColor, 1, True
    'Engine_Text_Render ">", 310, 470, DefaultColor, 1, True
    Engine_Draw_Box 280 + OffX, 470 + Offy, 20, 20, D3DColorARGB(100, 1, 1, 1)
    Engine_Text_Render frmCrearPersonaje.lbInteligencia.Caption, 282 + OffX, 473 + Offy, DefaultColor, 1, True ' Atributo Inteligencia
    'Engine_Text_Render "+", 335, 470, DefaultColor, 1, True
    Engine_Draw_Box 317 + OffX, 470 + Offy, 25, 20, D3DColorARGB(100, 1, 1, 1)
    Engine_Text_Render frmCrearPersonaje.modInteligencia.Caption, 320 + OffX, 473 + Offy, DefaultColor, 1, True ' Bonificacion Inteligencia
    
    Engine_Text_Render "Constitución ", 185 + OffX, 500 + Offy, DefaultColor, , True
    'Engine_Text_Render "<", 260, 500, DefaultColor, 1, True
    ' Engine_Text_Render ">", 310, 500, DefaultColor, 1, True
    Engine_Draw_Box 280 + OffX, 500 + Offy, 20, 20, D3DColorARGB(100, 1, 1, 1)
    Engine_Text_Render frmCrearPersonaje.lbConstitucion.Caption, 283 + OffX, 503 + Offy, DefaultColor, 1, True ' Atributo Constitución
    '
    ' Engine_Text_Render "+", 335, 500, DefaultColor, 1, True
    Engine_Draw_Box 317 + OffX, 500 + Offy, 25, 20, D3DColorARGB(100, 1, 1, 1)
    Engine_Text_Render frmCrearPersonaje.modConstitucion.Caption, 320 + OffX, 503 + Offy, DefaultColor, 1, True ' Bonificacion Constitución
    
    Engine_Text_Render "Constitución ", 185 + OffX, 500 + Offy, DefaultColor, , True
    'Engine_Text_Render "<", 260, 500, DefaultColor, 1, True
    ' Engine_Text_Render ">", 310, 500, DefaultColor, 1, True
    Engine_Draw_Box 280 + OffX, 530 + Offy, 20, 20, D3DColorARGB(100, 1, 1, 1)
    Engine_Text_Render frmCrearPersonaje.lbCarisma.Caption, 283 + OffX, 533 + Offy, DefaultColor, 1, True ' Atributo Carisma
    '
    ' Engine_Text_Render "+", 335, 500, DefaultColor, 1, True
    Engine_Draw_Box 317 + OffX, 530 + Offy, 25, 20, D3DColorARGB(100, 1, 1, 1)
    Engine_Text_Render frmCrearPersonaje.modCarisma.Caption, 320 + OffX, 533 + Offy, DefaultColor, 1, True ' Bonificacion Carisma
      
    '
    'Engine_Draw_Box 290, 528, 20, 20, D3DColorARGB(120, 1, 150, 150)
    'Engine_Text_Render "Puntos disponibles", 175, 530, DefaultColor, 1, True '
    'Engine_Text_Render frmCrearPersonaje.lbLagaRulzz.Caption, 291, 530, DefaultColor, 1, True '
    'Cabeza
    'Engine_Draw_Box 425, 415, 140, 100, D3DColorARGB(120, 100, 100, 100)

    ' Engine_Text_Render "Selecciona el rostro que más te agrade.", 662, 260, DefaultColor, 1, True

    OffX = -345
    Offy = -100
     
    Engine_Draw_Box 280, 407, 185, 120, D3DColorARGB(80, 0, 0, 0)
     
    Engine_Text_Render_LetraChica "Aspecto", 345, 385, DefaultColor, 6, False
    
    ' Engine_Draw_Box 345, 502, 12, 12, D3DColorARGB(120, 100, 0, 0)
    
    'Engine_Text_Render_LetraChica "Equipado", 360, 502, DefaultColor, 4, False
     
    ' CPHeading = 3
     
    DefaultColor(0) = D3DColorXRGB(255, 255, 255)
    DefaultColor(1) = DefaultColor(0)
    DefaultColor(2) = DefaultColor(0)
    DefaultColor(3) = DefaultColor(0)

    If CPHead <> 0 And CPArma <> 0 Then
         
        Engine_Text_Render_LetraChica "Cabeza", 350, 410, DefaultColor, 1, False
        Engine_Text_Render "<", 335, 412, DefaultColor, 1, True
        Engine_Text_Render ">", 403, 412, DefaultColor, 1, True
        
        Engine_Text_Render ">", 423, 428, DefaultColor, 3, True
        Engine_Text_Render "<", 293, 428, DefaultColor, 3, True
    
        'If CPEquipado Then
        '    Engine_Draw_Box 347, 512, 12, 12, D3DColorARGB(100, 255, 1, 1)
        '    Engine_Text_Render_LetraChica "Equipado", 360, 512, DefaultColor, 4, False
        '    Engine_Text_Render_LetraChica "x", 348, 512, DefaultColor, 6, False
        'Else
        '    Engine_Draw_Box 347, 512, 12, 12, D3DColorARGB(100, 255, 1, 1)
        '    Engine_Text_Render_LetraChica "Equipado", 360, 512, DefaultColor, 4, False
        'End If
    
        Dim Raza As Byte

        If frmCrearPersonaje.lstRaza.ListIndex < 0 Then
            frmCrearPersonaje.lstRaza.ListIndex = 0

        End If

        Raza = frmCrearPersonaje.lstRaza.ListIndex

        Dim enanooff As Byte

        If Raza = 0 Or Raza = 1 Or Raza = 2 Or Raza = 5 Then
            enanooff = 0
    
        Else
            enanooff = 10

        End If
    
        If CPEquipado Then
            Call Renderizar_Aura(CPAura, 686 + 15 + OffX, 360 - Offy + enanooff, 0, 0)

        End If
            
        If enanooff > 0 Then
            Draw_Grh BodyData(CPBodyE).Walk(CPHeading), 685 + 15 + OffX, 366 - Offy, 1, 0, DefaultColor()
        Else
            Draw_Grh BodyData(CPBody).Walk(CPHeading), 685 + 15 + OffX, 366 - Offy, 1, 0, DefaultColor()

        End If
            
        Draw_Grh HeadData(CPHead).Head(CPHeading), 685 + 15 + OffX, 366 - Offy + BodyData(CPBody).HeadOffset.y + enanooff, 1, 0, DefaultColor()
            
        'If CPEquipado Then
        'Draw_Grh CascoAnimData(CPGorro).Head(CPHeading), 700 + OffX, 366 - Offy + BodyData(CPBody).HeadOffset.y + enanooff, 1, 0, DefaultColor()
        'Draw_Grh WeaponAnimData(CPArma).WeaponWalk(CPHeading), 685 + 15 + OffX, 365 - Offy + enanooff, 1, 0, DefaultColor()
        'Call Renderizar_Aura(CPAura, 686 + 15 + offx, 360 - offy, 0, 0)
        'End If
            
        DefaultColor(0) = D3DColorXRGB(0, 128, 190)
        DefaultColor(1) = DefaultColor(0)
        DefaultColor(2) = DefaultColor(0)
        DefaultColor(3) = DefaultColor(0)
        Engine_Text_Render CPName, 372 - Engine_Text_Width(CPName, True) / 2, 495, DefaultColor, 1, True
    Else
        Engine_Text_Render "X", 355, 428, DefaultColor, 3, True

    End If
    
    'DefaultColor(0) = D3DColorXRGB(255, 255, 255)
    'DefaultColor(1) = DefaultColor(0)
    'DefaultColor(2) = DefaultColor(0)
    'DefaultColor(3) = DefaultColor(0)

    'Boton Atras
    'Engine_Draw_Box 147, 628, 100, 40, D3DColorARGB(80, 0, 0, 0)
    'Engine_Text_Render "< Volver", 170, 640, DefaultColor, 1, True
    
    'Boton Crear
    'If StopCreandoCuenta Then
    '    Engine_Draw_Box 730, 630, 100, 40, D3DColorARGB(120, 100, 180, 100)
    '    Engine_Text_Render "Creando...", 750, 640, DefaultColor, 1, True
    'Else
    '    Engine_Draw_Box 730, 630, 100, 40, D3DColorARGB(80, 0, 0, 0)
    '    Engine_Text_Render "Crear PJ >", 750, 640, DefaultColor, 1, True
    'End If
       
    'Engine_Text_Render "DADO", 670, 390, DefaultColor()
    Draw_GrhIndex 1123, 665, 385

End Sub

Public Sub RenderPjsCuenta()

    ' Renderiza el menu para seleccionar las clases
        
    Dim i               As Long

    Dim x               As Integer

    Dim y               As Integer

    Dim notY            As Integer

    Dim DefaultColor(3) As Long

    Dim color           As Long

    Dim Texto           As String

    Texto = CuentaEmail

    'Render fondo
    Draw_GrhIndex 1170, 0, 0
    
    Dim temp_array(3) As Long 'Si le queres dar color a la letra pasa este parametro dsp xD

    DefaultColor(0) = D3DColorXRGB(255, 255, 255)
    DefaultColor(1) = DefaultColor(0)
    DefaultColor(2) = DefaultColor(0)
    DefaultColor(3) = DefaultColor(0)

    Dim sumax As Long

    sumax = 84
            
    For i = 1 To 10
            
        If (i > 5) Then
            x = ((i * 132) - (5 * 132))
            y = 440
        Else
            x = (i * 132)
            y = 283

        End If

        x = x + sumax

        temp_array(0) = Pjs(i).LetraColor
        temp_array(1) = Pjs(i).LetraColor
        temp_array(2) = Pjs(i).LetraColor
        temp_array(3) = Pjs(i).LetraColor
        
        'Offset de la cabeza / enanos.
        ' If (Pjs(i).Clase <> eClass.Warrior) Then
        ' notY = 5
        ' Else
        Rem   notY = -5
        ' End If
        
        'Si tiene cuerpo dibuja
        If (Pjs(i).Body <> 0) Then
        
            If PJSeleccionado = i Then
                Call Particle_Group_Render(Select_part, x + 32, y + 5)

            End If

            If (Pjs(i).Body <> 0) Then
                  
                'Else
                'Engine_Draw_Box X - 40, Y - 40, 145, 150, D3DColorARGB(20, 28, 18, 9)
                Draw_Grh BodyData(Pjs(i).Body).Walk(3), x + 15, y + 10, 1, 1, DefaultColor()

            End If

            If (Pjs(i).Head <> 0) Then
                'If Not nohead Then
                Draw_Grh HeadData(Pjs(i).Head).Head(3), x + 15, y - notY + BodyData(Pjs(i).Body).HeadOffset.y + 10, 1, 0, DefaultColor()

                ' End If
            End If
            
            If (Pjs(i).Casco <> 0) Then
                'If Not nohead Then
                Draw_Grh CascoAnimData(Pjs(i).Casco).Head(3), x + 15, y - notY + BodyData(Pjs(i).Body).HeadOffset.y + 10, 1, 0, DefaultColor()

                ' End If
            End If
            
            If (Pjs(i).Escudo <> 0) Then
                'If Not nohead Then
                Draw_Grh ShieldAnimData(Pjs(i).Escudo).ShieldWalk(3), x + 14, y - notY + 10, 1, 0, DefaultColor()

                ' End If
            End If
                        
            If (Pjs(i).Arma <> 0) Then
                'If Not nohead Then
                Draw_Grh WeaponAnimData(Pjs(i).Arma).WeaponWalk(3), x + 14, y - notY + 10, 1, 0, DefaultColor()

                ' End If
            End If
            
            Dim colorCorazon(0 To 4) As Long

            Dim b                    As Long

            Dim g                    As Long

            Dim r                    As Long

            colorCorazon(0) = temp_array(1)
            colorCorazon(1) = temp_array(1)
            colorCorazon(2) = temp_array(1)
            colorCorazon(3) = temp_array(1)
            
            'Convert LONG to RGB:
            ' b = temp_array(1) \ 65536
            ' g = (temp_array(1) - b * 65536) \ 256
            'r = temp_array(1) - b * 65536 - g * 256
                
            '' r = (temp_array(1) And 16711680) / 65536
            ' g = (temp_array(1) And 65280) / 256
            ' b = temp_array(1) And 255
                
            colorCorazon(0) = D3DColorXRGB(r, g, b)
            colorCorazon(1) = colorCorazon(0)
            colorCorazon(2) = colorCorazon(0)
            colorCorazon(3) = colorCorazon(0)
        
            If CuentaDonador = 1 Then
                Grh_Render Estrella, x + 17 + 6 + Engine_Text_Width(Pjs(i).nombre, 1) / 2, y + 19, temp_array(), True, True, False

            End If

            Engine_Text_Render Pjs(i).nombre, x + 30 - Engine_Text_Width(Pjs(i).nombre, True) / 2, y + 56 - Engine_Text_Height(Pjs(i).nombre, True), temp_array(), 1, True
            
            If PJSeleccionado = i Then
            
                Dim Offy As Byte

                Offy = 0
            
                Engine_Text_Render Pjs(i).nombre, 511 - Engine_Text_Width(Pjs(i).nombre, True) / 2, 565 - Engine_Text_Height(Pjs(i).nombre, True), temp_array(), 1, True
                
                If Pjs(i).ClanName <> "<>" Then
                    Engine_Text_Render Pjs(i).ClanName, 511 - Engine_Text_Width(Pjs(i).ClanName, True) / 2, 565 + 15 - Engine_Text_Height(Pjs(i).ClanName, True), temp_array(), 1, True
                    Offy = 15
                Else
                
                    Offy = 0

                End If

                Engine_Text_Render "Clase: " & ListaClases(Pjs(i).Clase), 511 - Engine_Text_Width("Clase:" & ListaClases(Pjs(i).Clase), True) / 2, Offy + 585 - Engine_Text_Height("Clase:" & ListaClases(Pjs(i).Clase), True), DefaultColor, 1, True
                
                Engine_Text_Render "Nivel: " & Pjs(i).nivel, 511 - Engine_Text_Width("Nivel:" & Pjs(i).nivel, True) / 2, Offy + 600 - Engine_Text_Height("Nivel:" & Pjs(i).nivel, True), DefaultColor, 1, True
                Engine_Text_Render CStr(Pjs(i).NameMapa), 5111 - Engine_Text_Width(CStr(Pjs(i).NameMapa), True) / 2, Offy + 615 - Engine_Text_Height(CStr(Pjs(i).NameMapa), True), DefaultColor, 1, True

            End If
            
        End If

    Next i

End Sub

Sub RenderConsola()

    Dim i As Byte
 
    If OffSetConsola > 0 Then OffSetConsola = OffSetConsola - 1
    If OffSetConsola = 0 Then UltimaLineavisible = True
 
    For i = 1 To MaxLineas - 1
 
        Text_Render font_list(1), Con(i).T, ComienzoY + (i * 15) + OffSetConsola - 20, 10, frmmain.renderer.Width, frmmain.renderer.Height, ARGB(Con(i).r, Con(i).g, Con(i).b, i * (255 / MaxLineas)), DT_TOP Or DT_LEFT, False
        
    Next i
 
    If UltimaLineavisible = True Then Text_Render font_list(1), Con(i).T, ComienzoY + (MaxLineas * 15) + OffSetConsola - 20, 10, frmmain.renderer.Width, frmmain.renderer.Height, ARGB(Con(MaxLineas).r, Con(MaxLineas).g, Con(i).b, 255), DT_TOP Or DT_LEFT, False
 
End Sub


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

