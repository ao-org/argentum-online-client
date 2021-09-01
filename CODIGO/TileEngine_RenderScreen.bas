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
    
    On Error GoTo RenderScreen_Err
    

    '**************************************************************
    ' Author: Aaron Perkins
    ' Last Modify Date: 23/11/2020
    ' Modified by: Juan Martín Sotuyo Dodero (Maraxus)
    ' Last modified by: Alexis Caraballo (WyroX)
    ' Renders everything to the viewport
    '**************************************************************
    
    
    
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
    
    Dim DeltaTime                   As Long

    Dim TempColor(3)        As RGBA

    ' Tiles that are in range
    MinX = center_x - HalfTileWidth
    MaxX = center_x + HalfTileWidth
    MinY = center_y - HalfTileHeight
    MaxY = center_y + HalfTileHeight

    ' Buffer tiles (for layer 2, chars, big objects, etc.)
    MinBufferedX = MinX - TileBufferSizeX
    MaxBufferedX = MaxX + TileBufferSizeX
    MinBufferedY = MinY - 1
    MaxBufferedY = MaxY + TileBufferSizeY

    ' Screen start (with movement offset)
    StartX = PixelOffsetX - MinX * TilePixelWidth
    StartY = PixelOffsetY - MinY * TilePixelHeight

    ' Screen start with tiles buffered (for layer 2, chars, big objects, etc.)
    StartBufferedX = TileBufferPixelOffsetX + PixelOffsetX
    StartBufferedY = PixelOffsetY - TilePixelHeight

    ' Add 1 tile to the left if going left, else add it to the right
    If PixelOffsetX > 0 Then
        MinX = MinX - 1
    Else
        MaxX = MaxX + 1
    End If
    
    If PixelOffsetY > 0 Then
        MinY = MinY - 1
    
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
        StartBufferedY = StartBufferedY - (MinBufferedY - 1) * TilePixelHeight
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
                Call Draw_Grh_Precalculated(.Graphic(1), .light_value, (.Blocked And FLAG_AGUA) <> 0, (.Blocked And FLAG_LAVA) <> 0, x, y, MinX, MaxX, MinY, MaxY)
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
            
            End With

            ScreenX = ScreenX + TilePixelWidth
        Next x

        ScreenY = ScreenY + TilePixelHeight
    Next y
    
    
     ScreenY = StartBufferedY

    For y = MinBufferedY To MaxBufferedY
        ScreenX = StartBufferedX

        For x = MinBufferedX To MaxBufferedX

            With MapData(x, y)
                
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
                        
                        Call Copy_RGBAList_WithAlpha(TempColor, .light_value, .CharFantasma.AlphaB)
                        
                        'Seteamos el color
                        If .CharFantasma.Heading = 1 Or .CharFantasma.Heading = 2 Then
                            Call Draw_Grh(.CharFantasma.Escudo, ScreenX, ScreenY, 1, 1, TempColor(), False, x, y)
                            Call Draw_Grh(.CharFantasma.Body, ScreenX, ScreenY, 1, 1, TempColor(), False, x, y)
                            Call Draw_Grh(.CharFantasma.Head, ScreenX + .CharFantasma.OffX, ScreenY + .CharFantasma.Offy, 1, 1, TempColor(), False, x, y)
                            Call Draw_Grh(.CharFantasma.Casco, ScreenX + .CharFantasma.OffX, ScreenY + .CharFantasma.Offy, 1, 1, TempColor(), False, x, y)
                            Call Draw_Grh(.CharFantasma.Arma, ScreenX, ScreenY, 1, 1, TempColor(), False, x, y)
                        Else
                            Call Draw_Grh(.CharFantasma.Body, ScreenX, ScreenY, 1, 1, TempColor(), False, x, y)
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
                        
                        Case eObjType.otArboles
                            Call Draw_Sombra(.Graphic(3), ScreenX, ScreenY, 1, 1, False, x, y)

                            ' Debajo del arbol
                            If Abs(UserPos.x - x) < 3 And (Abs(UserPos.y - y)) < 8 And (Abs(UserPos.y) < y) Then
    
                                If .ArbolAlphaTimer <= 0 Then
                                    .ArbolAlphaTimer = LastMove
                                End If
    
                                DeltaTime = FrameTime - .ArbolAlphaTimer
    
                                Call Copy_RGBAList_WithAlpha(TempColor, .light_value, IIf(DeltaTime > ARBOL_ALPHA_TIME, ARBOL_MIN_ALPHA, 255 - DeltaTime / ARBOL_ALPHA_TIME * (255 - ARBOL_MIN_ALPHA)))
                                Call Draw_Grh(.ObjGrh, ScreenX, ScreenY, 1, 1, TempColor, False, x, y)
    
                            Else    ' Lejos del arbol
                                If .ArbolAlphaTimer = 0 Then
                                    Call Draw_Grh(.ObjGrh, ScreenX, ScreenY, 1, 1, .light_value, False, x, y)
    
                                Else
                                    If .ArbolAlphaTimer > 0 Then
                                        .ArbolAlphaTimer = -LastMove
                                    End If
    
                                    DeltaTime = FrameTime + .ArbolAlphaTimer
    
                                    If DeltaTime > ARBOL_ALPHA_TIME Then
                                        .ArbolAlphaTimer = 0
                                        Call Draw_Grh(.ObjGrh, ScreenX, ScreenY, 1, 1, .light_value, False, x, y)
                                    Else
                                        Call Copy_RGBAList_WithAlpha(TempColor, .light_value, ARBOL_MIN_ALPHA + DeltaTime * (255 - ARBOL_MIN_ALPHA) / ARBOL_ALPHA_TIME)
                                        Call Draw_Grh(.ObjGrh, ScreenX, ScreenY, 1, 1, TempColor, False, x, y)
                                    End If
                                End If
    
                            End If
                        
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

                        ' Debajo del arbol
                        If Abs(UserPos.x - x) <= 3 And (Abs(UserPos.y - y)) < 8 And (Abs(UserPos.y) < y) Then

                            If .ArbolAlphaTimer <= 0 Then
                                .ArbolAlphaTimer = LastMove
                            End If

                            DeltaTime = FrameTime - .ArbolAlphaTimer

                            Call Copy_RGBAList_WithAlpha(TempColor, .light_value, IIf(DeltaTime > ARBOL_ALPHA_TIME, ARBOL_MIN_ALPHA, 255 - DeltaTime / ARBOL_ALPHA_TIME * (255 - ARBOL_MIN_ALPHA)))
                            Call Draw_Grh(.Graphic(3), ScreenX, ScreenY, 1, 1, TempColor, False, x, y)

                        Else    ' Lejos del arbol
                            If .ArbolAlphaTimer = 0 Then
                                Call Draw_Grh(.Graphic(3), ScreenX, ScreenY, 1, 1, .light_value, False, x, y)

                            Else
                                If .ArbolAlphaTimer > 0 Then
                                    .ArbolAlphaTimer = -LastMove
                                End If

                                DeltaTime = FrameTime + .ArbolAlphaTimer

                                If DeltaTime > ARBOL_ALPHA_TIME Then
                                    .ArbolAlphaTimer = 0
                                    Call Draw_Grh(.Graphic(3), ScreenX, ScreenY, 1, 1, .light_value, False, x, y)
                                Else
                                    Call Copy_RGBAList_WithAlpha(TempColor, .light_value, ARBOL_MIN_ALPHA + DeltaTime * (255 - ARBOL_MIN_ALPHA) / ARBOL_ALPHA_TIME)
                                    Call Draw_Grh(.Graphic(3), ScreenX, ScreenY, 1, 1, TempColor, False, x, y)
                                End If
                            End If

                        End If

                    Else
                        If AgregarSombra(.Graphic(3).GrhIndex) Then
                            Call Draw_Sombra(.Graphic(3), ScreenX, ScreenY, 1, 1, False, x, y)
                        End If

                        Call Draw_Grh(.Graphic(3), ScreenX, ScreenY, 1, 1, .light_value, False, x, y)

                    End If

                End If
                '******************************************
            End With
            
            ScreenX = ScreenX + TilePixelWidth
        Next x

        ScreenY = ScreenY + TilePixelHeight
    Next y

    If InfoItemsEnRender And tX And tY Then
        With MapData(tX, tY)
            If .OBJInfo.OBJIndex Then
                If Not ObjData(.OBJInfo.OBJIndex).Agarrable Then
                    Dim Text As String, Amount As String
                    If .OBJInfo.Amount > 1000 Then
                        Amount = Round(.OBJInfo.Amount * 0.001, 1) & "K"
                    Else
                        Amount = .OBJInfo.Amount
                    End If
                    Text = ObjData(.OBJInfo.OBJIndex).Name & " (" & Amount & ")"
                    Call Engine_Text_Render(Text, MouseX + 15, MouseY, COLOR_WHITE, , , , 160)
                End If
            End If
        End With
    End If

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
        Dim Trigger As eTrigger
        For Trigger = LBound(RoofsLight) To UBound(RoofsLight)

            ' Si estoy bajo este techo
            If Trigger = MapData(UserPos.x, UserPos.y).Trigger Then
            
                If RoofsLight(Trigger) > 0 Then
                    ' Reduzco el alpha
                    RoofsLight(Trigger) = RoofsLight(Trigger) - timerTicksPerFrame * 48
                    If RoofsLight(Trigger) < 0 Then RoofsLight(Trigger) = 0
                End If

            ElseIf RoofsLight(Trigger) < 255 Then
            
                ' Aumento el alpha
                RoofsLight(Trigger) = RoofsLight(Trigger) + timerTicksPerFrame * 48
                If RoofsLight(Trigger) > 255 Then RoofsLight(Trigger) = 255
            
            End If
            
        Next

        ScreenY = StartBufferedY

        For y = MinBufferedY To MaxBufferedY
            ScreenX = StartBufferedX
    
            For x = MinBufferedX To MaxBufferedX
            
                With MapData(x, y)
                    ' Layer 4 - roofs *******************************
                    If .Graphic(4).GrhIndex Then

                        Trigger = NearRoof(x, y)

                        If Trigger Then
                            Call Copy_RGBAList_WithAlpha(TempColor, .light_value, RoofsLight(Trigger))

                            Call Draw_Grh(.Graphic(4), ScreenX, ScreenY, 1, 1, TempColor, , x, y)
                            
                        Else

                            Call Draw_Grh(.Graphic(4), ScreenX, ScreenY, 1, 1, .light_value, , x, y)
    
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
    ' FXs and dialogs loop
    ScreenY = StartBufferedY

    For y = MinBufferedY To MaxBufferedY
        ScreenX = StartBufferedX

        For x = MinBufferedX To MaxBufferedX
            
            With MapData(x, y)

                Dim i As Long

                ' Dialogs *******************************
                If MapData(x, y).charindex <> 0 Then
                
                    If charlist(.charindex).active = 1 Then
                    
                        Call Char_TextRender(.charindex, ScreenX, ScreenY, x, y)
                    
                    End If
                    
                End If
                '******************************************

                ' Render text value *******************************
                If UBound(.DialogEffects) > 0 Then
                    For i = 1 To UBound(.DialogEffects)
                        With .DialogEffects(i)
                            If LenB(.Text) <> 0 Then
                                Dim DialogTime As Long
                                DialogTime = FrameTime - .Start
            
                                If DialogTime > 1300 Then
                                    .Text = vbNullString
                                Else
                                    If DialogTime > 900 Then
                                        Call RGBAList(TempColor, .Color.r, .Color.G, .Color.B, .Color.A * (1300 - DialogTime) * 0.0025)
                                    Else
                                        Call RGBAList(TempColor, .Color.r, .Color.G, .Color.B, .Color.A)
                                    End If
                            
                                    Engine_Text_Render_Efect 0, .Text, ScreenX + 16 - Int(Engine_Text_Width(.Text, False) * 0.5) + .offset.x, ScreenY - Engine_Text_Height(.Text, False) + .offset.y - DialogTime * 0.025, TempColor, 1, False
                    
                                End If
                            End If
                        End With
                    Next
                End If
                '******************************************

                ' FXs *******************************
                If .FxCount > 0 Then
                    For i = 1 To .FxCount

                        If .FxList(i).FxIndex > 0 And .FxList(i).Started <> 0 Then
    
                            Call RGBAList(TempColor, 255, 255, 255, 220)

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
            
            Call Particle_Group_Render(MeteoParticle, ScreenX, ScreenY)
            
            LastOffsetX = ParticleOffsetX
            LastOffsetY = ParticleOffsetY

        End If

    End If

    If AlphaNiebla Then
    
        If MapDat.niebla Then Call Engine_Weather_UpdateFog

    End If

    If bNieve Then
    
        If MapDat.NIEVE Then
        
            If Graficos_Particulas.Engine_MeteoParticle_Get <> 0 Then
            
                'Screen positions were hardcoded by now
                ScreenX = 250
                ScreenY = 0
                
                Call Particle_Group_Render(MeteoParticle, ScreenX, ScreenY)

            End If

        End If

    End If
    
    Call Effect_Render_All
    
    If InvasionActual Then
        
        Call Engine_Draw_Box(190, 550, 356, 36, RGBA_From_Comp(0, 0, 0, 200))
        Call Engine_Draw_Box(193, 553, 3.5 * InvasionPorcentajeVida, 30, RGBA_From_Comp(20, 196, 255, 200))
        
        Call Engine_Draw_Box(340, 586, 54, 9, RGBA_From_Comp(0, 0, 0, 200))
        Call Engine_Draw_Box(342, 588, 0.5 * InvasionPorcentajeTiempo, 5, RGBA_From_Comp(220, 200, 0, 200))
        
    End If

    If Pregunta Then
        
        Call Engine_Draw_Box(283, 170, 170, 100, RGBA_From_Comp(150, 20, 3, 200))
        Call Engine_Draw_Box(288, 175, 160, 90, RGBA_From_Comp(25, 25, 23, 200))

        Dim preguntaGrh As grh
        Call InitGrh(preguntaGrh, 32120)
        
        Call Engine_Text_Render(PreguntaScreen, 290, 180, COLOR_WHITE, 1, True)
        
        Call Draw_Grh(preguntaGrh, 392, 233, 1, 0, COLOR_WHITE, False, 0, 0, 0)

    End If
    
    If pescandoPezEspecial Then
        Dim barraPezEspecial As grh
        Call InitGrh(barraPezEspecial, 60666)
        Call Draw_Grh(barraPezEspecial, 350, 570, 1, 0, COLOR_WHITE, False, 0, 0, 0)
        Dim cursorPezEspecial As grh
        Call InitGrh(cursorPezEspecial, 60667)
        
        cursorOffsetX = Clamp(cursorOffsetX + pescaDirection * timerElapsedTime * 0.01, -100, 100)
        
        Call Draw_Grh(cursorPezEspecial, 360 + cursorOffsetX, 562, 1, 0, COLOR_WHITE, False, 0, 0, 0)
    End If

    If cartel Then

        Call RGBAList(TempColor, 255, 255, 255, 220)
        
        Dim TempGrh  As grh
        Call InitGrh(TempGrh, GrhCartel)
        
        Call Draw_Grh(TempGrh, CInt(clicX), CInt(clicY), 1, 0, TempColor(), False, 0, 0, 0)
        
        Call Engine_Text_Render(Leyenda, CInt(clicX - 100), CInt(clicY - 130), TempColor(), 1, False)

    End If

    Call RenderScreen_NombreMapa

    
    Exit Sub

RenderScreen_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine_RenderScreen.RenderScreen", Erl)
    Resume Next
    
End Sub

Private Sub RenderScreen_NombreMapa()
    
    On Error GoTo RenderScreen_NombreMapa_Err
    
    
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
        
        Dim Color(3) As RGBA
        Call RGBAList(Color(), 179, 95, 0, map_letter_a)
        
        Call Grh_Render(letter_grh, 250, 300, Color())
        
        Call Engine_Text_RenderGrande(letter_text, 360 - Engine_Text_Width(letter_text, False, 4) / 2, 1, Color(), 5, False, , CInt(map_letter_a))

    End If

    
    Exit Sub

RenderScreen_NombreMapa_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine_RenderScreen.RenderScreen_NombreMapa", Erl)
    Resume Next
    
End Sub




