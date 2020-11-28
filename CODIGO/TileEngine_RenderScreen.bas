Attribute VB_Name = "TileEngine_RenderScreen"
Option Explicit

'Letter showing on screen
Public letter_text           As String
Public letter_grh            As grh
Public map_letter_grh        As grh
Public map_letter_grh_next   As Long
Public map_letter_a          As Single
Public map_letter_fadestatus As Byte

Sub RenderScreen(ByVal center_x As Integer, ByVal center_y As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer, ByVal HalfTileWidth As Integer, ByVal HalfTileHeight As Integer)

    '**************************************************************
    ' Author: Aaron Perkins
    ' Last Modify Date: 23/11/2020
    ' Modified by: Juan Martín Sotuyo Dodero (Maraxus)
    ' Last modified by: Alexis Caraballo (WyroX)
    ' Renders everything to the viewport
    '**************************************************************
    
    On Error Resume Next
    
    Dim y                   As Integer      ' Keeps track of where on map we are
    Dim x                   As Integer      ' Keeps track of where on map we are

    Dim MinX                As Integer
    Dim MaxX                As Integer

    Dim MinY                As Integer
    Dim MaxY                As Integer
    
    Dim MinBufferedX        As Integer
    Dim MaxBufferedX        As Integer
    
    Dim MinBufferedY        As Integer
    Dim MaxBufferedY        As Integer
    
    Dim StartX              As Integer
    Dim StartY              As Integer
    
    Dim StartBufferedX      As Integer
    Dim StartBufferedY      As Integer

    Dim ScreenX             As Integer      ' Keeps track of where to place tile on screen
    Dim ScreenY             As Integer      ' Keeps track of where to place tile on screen

    Dim TempColor(3)        As Long         ' Temporarily store a Long type color into a list

    ' Tiles that are in range
    MinX = center_x - HalfTileWidth
    MaxX = center_x + HalfTileWidth
    MinY = center_y - HalfTileHeight
    MaxY = center_y + HalfTileHeight

    ' Buffer tiles (for layer 2, chars, big objects, etc.)
    MinBufferedX = MinX - TileBufferSizeX
    MaxBufferedX = MaxX + TileBufferSizeX
    MinBufferedY = MinY
    MaxBufferedY = MaxY + TileBufferSizeY

    ' Screen start (with movement offset)
    StartX = PixelOffsetX - MinX * TilePixelWidth
    StartY = PixelOffsetY - MinY * TilePixelHeight

    ' Screen start with tiles buffered (for layer 2, chars, big objects, etc.)
    StartBufferedX = TileBufferPixelOffsetX + PixelOffsetX
    StartBufferedY = PixelOffsetY

    ' Add 1 tile to the left if going left, else add it to the right
    If PixelOffsetX > 0 Then
        MinX = MinX - 1
    Else
        MaxX = MaxX + 1
    End If
    
    If PixelOffsetY > 0 Then
        MinY = MinY - 1
        MinBufferedY = MinBufferedY - 1
        StartBufferedY = StartBufferedY - TilePixelHeight
    
    Else
        MaxY = MaxY + 1
    End If
    
    ' Map border checks
    If MinX < XMinMapSize Then
        StartBufferedX = PixelOffsetX - MinX * TilePixelWidth
        MaxX = MaxX - MinX
        MaxBufferedX = MaxBufferedX - MinX
        MinX = XMinMapSize
        MinBufferedX = XMinMapSize
    
    ElseIf MinBufferedX < XMinMapSize Then
        StartBufferedX = StartBufferedX - (MinBufferedX - XMinMapSize) * TilePixelWidth
        MinBufferedX = XMinMapSize
    
    ElseIf MaxX > XMaxMapSize Then
        MaxX = XMaxMapSize
        MaxBufferedX = XMaxMapSize
        
    ElseIf MaxBufferedX > XMaxMapSize Then
        MaxBufferedX = XMaxMapSize
    End If
    
    If MinY < YMinMapSize Then
        StartBufferedY = PixelOffsetY - MinY * TilePixelHeight
        MaxY = MaxY - MinY
        MaxBufferedY = MaxBufferedY - MinY
        MinY = YMinMapSize
        MinBufferedY = YMinMapSize
    
    ElseIf MinBufferedY < YMinMapSize Then
        StartBufferedY = StartBufferedY - MinBufferedY * TilePixelHeight
        MinBufferedY = YMinMapSize
    
    ElseIf MaxY > YMaxMapSize Then
        MaxY = YMaxMapSize
        MaxBufferedY = YMaxMapSize
        
    ElseIf MaxBufferedY > YMaxMapSize Then
        MaxBufferedY = YMaxMapSize
    End If
    
    Call SpriteBatch.BeginPrecalculated(StartX, StartY)

    ' *********************************
    ' Layer 1 loop
    For y = MinY To MaxY
        For x = MinX To MaxX
            
            With MapData(x, y)

                ' Layer 1 *********************************
                Call Draw_Grh_Precalculated(.Graphic(1), .light_value, (.Blocked And FLAG_AGUA) <> 0)
                '******************************************
          
            End With

        Next x
    Next y
    
    Call SpriteBatch.EndPrecalculated

    ' *********************************
    ' Layer 2 & small objects loop
    Call DirectDevice.SetRenderState(D3DRS_ALPHATESTENABLE, True) ' Para no pisar los reflejos
    
    ScreenY = StartBufferedY

    For y = MinBufferedY To MaxBufferedY
        ScreenX = StartBufferedX

        For x = MinBufferedX To MaxBufferedX

            With MapData(x, y)

                ' Layer 2 *********************************
                If .Graphic(2).GrhIndex <> 0 Then
                    Call Draw_Grh(.Graphic(2), ScreenX, ScreenY, 1, 1, .light_value, , x, y)
                End If
                '******************************************
                
                ' Objects *********************************
                If .ObjGrh.GrhIndex <> 0 Then
                    Select Case ObjData(.OBJInfo.OBJIndex).ObjType
                    
                        Case eObjType.otArboles, eObjType.otPuertas, eObjType.otTeleport, eObjType.otCarteles, eObjType.OtPozos, eObjType.otYacimiento, eObjType.OtCorreo

                        Case Else
                            ' Objetos en el suelo (items, decorativos, etc)
                            Call Draw_Grh(.ObjGrh, ScreenX, ScreenY, 1, 1, .light_value, , x, y)
                    
                    End Select
                End If
                '******************************************

            End With

            ScreenX = ScreenX + TilePixelWidth
        Next x

        ScreenY = ScreenY + TilePixelHeight
    Next y
    
    Call DirectDevice.SetRenderState(D3DRS_ALPHATESTENABLE, False)
    
    ' *********************************
    '  Layer 3 & chars
    ScreenY = StartBufferedY

    For y = MinBufferedY To MaxBufferedY
        ScreenX = StartBufferedX

        For x = MinBufferedX To MaxBufferedX
            
            With MapData(x, y)
                ' Chars ***********************************
                If .charindex = UserCharIndex Then 'evitamos reenderizar un clon del usuario
                    If x <> UserPos.x Or y <> UserPos.y Then
                        .charindex = 0
                    End If
                End If
                
                If .CharFantasma.Activo Then
                    
                    If .CharFantasma.AlphaB > 0 Then
                    
                        .CharFantasma.AlphaB = .CharFantasma.AlphaB - (timerTicksPerFrame * 30)
                        
                        'Redondeamos a 0 para prevenir errores
                        If .CharFantasma.AlphaB < 0 Then .CharFantasma.AlphaB = 0
                        
                        'Seteamos el color
                        Call Long_To_RGBList(TempColor, D3DColorARGB(CInt(.CharFantasma.AlphaB), ColorAmbiente.r, ColorAmbiente.g, ColorAmbiente.b))

                        If .CharFantasma.Heading = 1 Or .CharFantasma.Heading = 2 Then
                            Call Draw_Grh(.CharFantasma.Escudo, ScreenX, ScreenY, 1, 1, TempColor(), False, x, y)
                            Call Draw_Grh(.CharFantasma.Body, ScreenX + 1, ScreenY, 1, 1, TempColor(), False, x, y)
                            Call Draw_Grh(.CharFantasma.Head, ScreenX + .CharFantasma.OffX, ScreenY + .CharFantasma.Offy, 1, 1, TempColor(), False, x, y)
                            Call Draw_Grh(.CharFantasma.Casco, ScreenX + .CharFantasma.OffX, ScreenY + .CharFantasma.Offy, 1, 1, TempColor(), False, x, y)
                            Call Draw_Grh(.CharFantasma.Arma, ScreenX, ScreenY, 1, 1, TempColor(), False, x, y)
                        Else
                            Call Draw_Grh(.CharFantasma.Body, ScreenX + 1, ScreenY, 1, 1, TempColor(), False, x, y)
                            Call Draw_Grh(.CharFantasma.Head, ScreenX + .CharFantasma.OffX, ScreenY + .CharFantasma.Offy, 1, 1, TempColor(), False, x, y)
                            Call Draw_Grh(.CharFantasma.Escudo, ScreenX, ScreenY, 1, 1, TempColor(), False, x, y)
                            Call Draw_Grh(.CharFantasma.Casco, ScreenX + .CharFantasma.OffX, ScreenY + .CharFantasma.Offy, 1, 1, TempColor(), False, x, y)
                            Call Draw_Grh(.CharFantasma.Arma, ScreenX, ScreenY, 1, 1, TempColor(), False, x, y)
                        End If

                    Else
                        .CharFantasma.Activo = False

                    End If

                End If
                
                If .charindex <> 0 Then
                    If charlist(.charindex).active = 1 Then
                        Call Char_Render(.charindex, ScreenX, ScreenY, x, y)
                    End If
                End If
                '******************************************
                
            End With
            
            ScreenX = ScreenX + TilePixelWidth
        Next x
        
        ' Recorremos de nuevo esta fila para dibujar objetos grandes y capa 3 encima de chars
        ScreenX = StartBufferedX

        For x = MinBufferedX To MaxBufferedX

            With MapData(x, y)
                ' Objects *********************************
                If .ObjGrh.GrhIndex <> 0 Then
                    Select Case ObjData(.OBJInfo.OBJIndex).ObjType
                    
                        Case eObjType.otPuertas, eObjType.otTeleport, eObjType.otCarteles, eObjType.OtPozos, eObjType.otYacimiento, eObjType.OtCorreo
                            ' Objetos grandes (menos árboles)
                            Call Draw_Grh(.ObjGrh, ScreenX, ScreenY, 1, 1, .light_value, , x, y)
                    
                    End Select
                End If
                '******************************************
    
                'Layer 3 **********************************
                If .Graphic(3).GrhIndex <> 0 Then
                
                    If (.Blocked And FLAG_ARBOL) <> 0 Then
                    
                        Call Draw_Sombra(.Graphic(3), ScreenX, ScreenY, 1, 1, False, x, y)
                        
                        If Abs(UserPos.x - x) < 3 And (Abs(UserPos.y - y)) < 5 And (Abs(UserPos.y) < y) Then
    
                            Call Long_To_RGBList(TempColor(), D3DColorARGB(200, ColorAmbiente.r, ColorAmbiente.g, ColorAmbiente.b))

                            Call Draw_Grh(.Graphic(3), ScreenX, ScreenY, 1, 1, TempColor, False, x, y)
                            
                        Else

                            Call Draw_Grh(.Graphic(3), ScreenX, ScreenY, 1, 1, .light_value, False, x, y)
                            
                        End If
                        
                    Else
                    
                        Call Draw_Grh(.Graphic(3), ScreenX, ScreenY, 1, 1, .light_value, False, x, y)
                        
                    End If
    
                End If
                '******************************************
            End With
            
            ScreenX = ScreenX + TilePixelWidth
        Next x

        ScreenY = ScreenY + TilePixelHeight
    Next y
    
    
    ' *********************************
    ' Particles loop
    ScreenY = StartBufferedY

    For y = MinBufferedY To MaxBufferedY
        ScreenX = StartBufferedX

        For x = MinBufferedX To MaxBufferedX
            
            With MapData(x, y)
                ' Particles *******************************
                If .particle_group > 0 Then
                    Call Particle_Group_Render(.particle_group, ScreenX + 16, ScreenY + 16)
                End If
                '******************************************
            End With
            
            ScreenX = ScreenX + TilePixelWidth
        Next x

        ScreenY = ScreenY + TilePixelHeight
    Next y
 
    ' *********************************
    ' Layer 4 loop
    If HayLayer4 Then

        ' Actualizo techos
        Dim Trigger As Integer
        For Trigger = LBound(RoofsLight) To UBound(RoofsLight)
        
            With RoofsLight(Trigger)

                ' Si estoy bajo este techo
                If Trigger = MapData(UserPos.x, UserPos.y).Trigger Then
                
                    If .Alpha > 0 Then
                        ' Reduzco el alpha
                        .Alpha = .Alpha - timerTicksPerFrame * 12
                        If .Alpha < 0 Then .Alpha = 0
                    End If
    
                ElseIf .Alpha < 255 Then
                
                    ' Aumento el alpha
                    .Alpha = .Alpha + timerTicksPerFrame * 12
                    If .Alpha > 255 Then .Alpha = 255
                
                End If
                
                ' Guardo el color nuevo
                .Color = D3DColorARGB(.Alpha, ColorAmbiente.r, ColorAmbiente.g, ColorAmbiente.b)

            End With
            
        Next

        ScreenY = StartBufferedY

        For y = MinBufferedY To MaxBufferedY
            ScreenX = StartBufferedX
    
            For x = MinBufferedX To MaxBufferedX
            
                With MapData(x, y)
                    ' Layer 4 - roofs *******************************
                    If .Graphic(4).GrhIndex Then

                        If .Trigger >= PRIMER_TRIGGER_TECHO Then
                            
                            Call Long_To_RGBList(TempColor(), RoofsLight(.Trigger).Color)
                            Call Draw_Grh(.Graphic(4), ScreenX, ScreenY, 1, 1, TempColor(), , x, y)
                            
                        Else
                            
                            Call Long_To_RGBList(TempColor(), map_base_light)
                            Call Draw_Grh(.Graphic(4), ScreenX, ScreenY, 1, 1, TempColor(), , x, y)
    
                        End If
    
                    End If
                    '******************************************
                End With

                ScreenX = ScreenX + TilePixelWidth
            Next x

            ScreenY = ScreenY + TilePixelHeight
        Next y
        
    End If
    
    
    ' *********************************
    ' FXs, dialogs, rendered values loop
    ScreenY = StartBufferedY

    For y = MinBufferedY To MaxBufferedY
        ScreenX = StartBufferedX

        For x = MinBufferedX To MaxBufferedX
            
            With MapData(x, y)
                ' Dialogs *******************************
                If MapData(x, y).charindex <> 0 Then
                
                    If charlist(.charindex).active = 1 Then
                    
                        Call Char_TextRender(.charindex, ScreenX, ScreenY, x, y)
                    
                    End If
                    
                End If
                '******************************************

                ' Render text value *******************************
                Call modRenderValue.Draw(x, y, ScreenX + 16, ScreenY, timerTicksPerFrame)
                '******************************************

                ' FXs *******************************
                If .FxCount > 0 Then
                    Dim i As Long
                    For i = 1 To .FxCount

                        If .FxList(i).FxIndex > 0 And .FxList(i).Started <> 0 Then
    
                            Call Long_To_RGBList(TempColor(), D3DColorARGB(220, 255, 255, 255))

                            If FxData(.FxList(i).FxIndex).IsPNG = 1 Then
                            
                                Call Draw_GrhFX(.FxList(i), ScreenX + FxData(.FxList(i).FxIndex).OffsetX, ScreenY + FxData(.FxList(i).FxIndex).OffsetY + 20, 1, 1, TempColor(), False)

                            Else
                                Call Draw_GrhFX(.FxList(i), ScreenX + FxData(.FxList(i).FxIndex).OffsetX, ScreenY + FxData(.FxList(i).FxIndex).OffsetY + 20, 1, 1, TempColor(), True)

                            End If

                        End If

                        If .FxList(i).Started = 0 Then .FxList(i).FxIndex = 0

                    Next i

                    If .FxList(.FxCount).Started = 0 Then .FxCount = .FxCount - 1

                End If
                '******************************************
                End With
            
            ScreenX = ScreenX + TilePixelWidth
        Next x

        ScreenY = ScreenY + TilePixelHeight
    Next y

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
    
        If MapDat.niebla Then Call Engine_Weather_UpdateFog

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

    If Pregunta Then
        
        Call Engine_Draw_Box(283, 180, 170, 80, D3DColorARGB(200, 150, 20, 3))
        Call Engine_Draw_Box(288, 185, 160, 70, D3DColorARGB(200, 25, 25, 23))

        Dim preguntaGrh As grh
            preguntaGrh.framecounter = 1
            preguntaGrh.GrhIndex = 32120
            preguntaGrh.Started = 1
        
        Call Long_To_RGBList(TempColor(), D3DColorARGB(255, 255, 255, 255))
        
        Call Engine_Text_Render(PreguntaScreen, 290, 190, TempColor(), 1, True)
        
        Call Draw_Grh(preguntaGrh, 392, 223, 1, 0, TempColor(), False, 0, 0, 0)

    End If

    Call Effect_Render_All

    If cartel Then

        Call Long_To_RGBList(TempColor(), D3DColorARGB(200, 255, 255, 255))
        
        Dim TempGrh  As grh
            TempGrh.framecounter = 1
            TempGrh.GrhIndex = GrhCartel
        
        Call Draw_Grh(TempGrh, CInt(clicX), CInt(clicY), 1, 0, TempColor(), False, 0, 0, 0)
        
        Call Engine_Text_Render(Leyenda, CInt(clicX - 100), CInt(clicY - 130), TempColor(), 1, False)

    End If

    Call RenderScreen_NombreMapa

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

    Dim MinY             As Integer  'Start Y pos on current map
    Dim MaxY             As Integer  'End Y pos on current map

    Dim MinX             As Integer  'Start X pos on current map
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
    
    Dim target_color(3)  As Long    ' Temporarily store a Long type color into a list
    Dim ColorCiego(3)    As Long

    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    MinY = screenminY
    MaxY = screenmaxY + TileBufferSizeY
    MinX = screenminX - TileBufferSizeX
    MaxX = screenmaxX + TileBufferSizeX
    
    'Make sure mins and maxs are allways in map bounds
    If MinY < XMinMapSize Then
        minYOffset = YMinMapSize - MinY
        MinY = YMinMapSize

    End If
    
    If MaxY > YMaxMapSize Then MaxY = YMaxMapSize
    
    If MinX < XMinMapSize Then
        minXOffset = XMinMapSize - MinX
        MinX = XMinMapSize

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
    
    Call Long_To_RGBList(ColorCiego(), D3DColorARGB(255, 15, 15, 15))
    
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

    ScreenY = minYOffset - TileBufferSizeX

    For y = MinY To MaxY
        ScreenX = minXOffset - TileBufferSizeY

        For x = MinX To MaxX

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
    
    ScreenY = minYOffset - TileBufferSizeY

    For y = MinY To MaxY
        ScreenX = minXOffset - TileBufferSizeX

        For x = MinX To MaxX
            PixelOffsetXTemp = ScreenX * 32 + PixelOffsetX
            PixelOffsetYTemp = ScreenY * 32 + PixelOffsetY
            
            With MapData(x, y)

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

    ScreenY = minYOffset - TileBufferSizeY

    For y = MinY To MaxY
        ScreenX = minXOffset - TileBufferSizeY

        For x = MinX To MaxX

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
        ScreenY = minYOffset - TileBufferSizeY

        For y = MinY To MaxY
            ScreenX = minXOffset - TileBufferSizeY

            For x = MinX To MaxX
        
                If MapData(x, y).Graphic(4).GrhIndex Then
        
                    'Layer 4 **********************************
                    If bTecho Then

                        If MapData(UserPos.x, UserPos.y).Trigger = MapData(x, y).Trigger Then
                    
                            If MapData(x, y).GrhBlend <= 20 Then MapData(x, y).GrhBlend = 20
                            
                            MapData(x, y).GrhBlend = MapData(x, y).GrhBlend - (timerTicksPerFrame * 12)
                        
                            Call Draw_Grh(MapData(x, y).Graphic(4), ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, 1, 1, ColorCiego, , x, y)
                        
                        Else
                            Call Draw_Grh(MapData(x, y).Graphic(4), ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, 1, 1, ColorCiego, , x, y)

                        End If

                    Else
                 
                        MapData(x, y).GrhBlend = MapData(x, y).GrhBlend + (timerTicksPerFrame * 12)

                        If MapData(x, y).GrhBlend >= 255 Then MapData(x, y).GrhBlend = 255
                        
                        Call Draw_Grh(MapData(x, y).Graphic(4), ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, 1, 1, ColorCiego, , x, y)
          
                    End If

                End If
 
                '**********************************
                ScreenX = ScreenX + 1
            Next x

            ScreenY = ScreenY + 1
        Next y
        
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

    ElseIf AlphaNiebla Then
    
        If MapDat.niebla Then Call Engine_Weather_UpdateFog

    ElseIf bNieve Then
    
        If MapDat.NIEVE Then
        
            If Engine_Meteo_Particle_Get <> 0 Then
            
                'Screen positions were hardcoded by now
                ScreenX = 250
                ScreenY = 0
                
                Call Particle_Group_Render(meteo_particle, ScreenX, ScreenY)

            End If

        End If

    End If
    
    If Pregunta Then
    
        'PreguntaScreen = "¿Esta seguro que asen es gay? ¿Que se lo come a fede?"
        Call Engine_Draw_Box(283, 180, 170, 80, D3DColorARGB(200, 219, 116, 3))
        Call Engine_Draw_Box(288, 185, 160, 70, D3DColorARGB(200, 51, 27, 3))

        Dim preguntaGrh As grh
            preguntaGrh.framecounter = 1
            preguntaGrh.GrhIndex = 32120
            preguntaGrh.Started = 1

        Call Long_To_RGBList(target_color(), D3DColorARGB(255, 255, 255, 255))
        
        Call Engine_Text_Render(PreguntaScreen, 290, 190, target_color(), 1, True)
        
        Call Draw_Grh(preguntaGrh, 392, 223, 1, 0, target_color(), False, 0, 0, 0)

    ElseIf cartel Then
    
        Dim TempGrh As grh
            TempGrh.framecounter = 1
            TempGrh.GrhIndex = GrhCartel
        
        Call Long_To_RGBList(target_color(), D3DColorARGB(200, 255, 255, 255))
        
        Call Draw_Grh(TempGrh, CInt(clicX), CInt(clicY), 1, 0, target_color(), False, 0, 0, 0)
        
        Call Engine_Text_Render(Leyenda, CInt(clicX - 100), CInt(clicY - 130), target_color(), 1, False)

    End If

End Sub

Private Sub RenderScreen_NombreMapa()
    
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
        
        Dim Color(3) As Long
        Call Long_To_RGBList(Color(), D3DColorARGB(CInt(map_letter_a), 179, 95, 0))
        
        Call Grh_Render(letter_grh, 250, 300, Color())
        
        Call Engine_Text_RenderGrande(letter_text, 360 - Engine_Text_Width(letter_text, False, 4) / 2, 1, Color(), 5, False, , CInt(map_letter_a))

    End If

End Sub




