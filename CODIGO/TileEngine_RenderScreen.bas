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

    Dim TempColor(3)        As RGBA

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
                Call Draw_Grh_Precalculated(.Graphic(1), .light_value, (.Blocked And FLAG_AGUA) <> 0, x, y, MinX, MaxX, MinY, MaxY)
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
                        
                            Call Draw_Grh(.Graphic(3), ScreenX, ScreenY, 1, 1, COLOR_WHITE, False, x, y)
                            
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

            ' Si estoy bajo este techo
            If Trigger = MapData(UserPos.x, UserPos.y).Trigger Then
            
                If RoofsLight(Trigger) > 0 Then
                    ' Reduzco el alpha
                    RoofsLight(Trigger) = RoofsLight(Trigger) - timerTicksPerFrame * 12
                    If RoofsLight(Trigger) < 0 Then RoofsLight(Trigger) = 0
                End If

            ElseIf RoofsLight(Trigger) < 255 Then
            
                ' Aumento el alpha
                RoofsLight(Trigger) = RoofsLight(Trigger) + timerTicksPerFrame * 12
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

                        If .Trigger >= PRIMER_TRIGGER_TECHO Then

                            Call SetRGBA(TempColor(0), .light_value(0).R, .light_value(0).G, .light_value(0).B, RoofsLight(.Trigger))
                            Call SetRGBA(TempColor(1), .light_value(1).R, .light_value(1).G, .light_value(1).B, RoofsLight(.Trigger))
                            Call SetRGBA(TempColor(2), .light_value(2).R, .light_value(2).G, .light_value(2).B, RoofsLight(.Trigger))
                            Call SetRGBA(TempColor(3), .light_value(3).R, .light_value(3).G, .light_value(3).B, RoofsLight(.Trigger))

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

    If Pregunta Then
        
        Call Engine_Draw_Box(283, 180, 170, 80, RGBA_From_Comp(150, 20, 3, 200))
        Call Engine_Draw_Box(288, 185, 160, 70, RGBA_From_Comp(25, 25, 23, 200))

        Dim preguntaGrh As grh
        Call InitGrh(preguntaGrh, 32120)
        
        Call Engine_Text_Render(PreguntaScreen, 290, 190, COLOR_WHITE, 1, True)
        
        Call Draw_Grh(preguntaGrh, 392, 223, 1, 0, COLOR_WHITE, False, 0, 0, 0)

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
    Call RegistrarError(Err.number, Err.Description, "TileEngine_RenderScreen.RenderScreen", Erl)
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
    Call RegistrarError(Err.number, Err.Description, "TileEngine_RenderScreen.RenderScreen_NombreMapa", Erl)
    Resume Next
    
End Sub




