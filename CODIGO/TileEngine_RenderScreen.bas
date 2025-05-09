Attribute VB_Name = "TileEngine_RenderScreen"
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

'Letter showing on screen
Public letter_text           As String
Public letter_grh            As grh
Public map_letter_grh        As grh
Public map_letter_grh_next   As Long
Public map_letter_a          As Single
Public map_letter_fadestatus As Byte
Public gameplay_render_offset As Vector2
Public Const hotkey_render_posX = 200
Public Const hotkey_render_posY = 40
Public Const hotkey_arrow_posx = 200 + 36 * 5 - 5
Public Const hotkey_arrow_posy = 10

Sub RenderScreen(ByVal center_x As Integer, ByVal center_y As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer, ByVal HalfTileWidth As Integer, ByVal HalfTileHeight As Integer)
    
    On Error GoTo RenderScreen_Err

    ' Renders everything to the viewport

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

    Dim screenX             As Integer      ' Keeps track of where to place tile on screen
    Dim screenY             As Integer      ' Keeps track of where to place tile on screen
    
    Dim DeltaTime                   As Long

    Dim TempColor(3)        As RGBA
    Dim ColorBarraPesca(3)  As RGBA

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
    StartX = PixelOffsetX - MinX * TilePixelWidth + gameplay_render_offset.x
    StartY = PixelOffsetY - MinY * TilePixelHeight + gameplay_render_offset.y

    ' Screen start with tiles buffered (for layer 2, chars, big objects, etc.)
    StartBufferedX = TileBufferPixelOffsetX + PixelOffsetX + gameplay_render_offset.x
    StartBufferedY = PixelOffsetY - TilePixelHeight + gameplay_render_offset.y

    ' Add 1 tile to the left if going left, else add it to the right
    If PixelOffsetX > 0 Then
        MinX = MinX - 1
    Else
        MaxX = MaxX + 10
    End If
    
    If PixelOffsetY > 0 Then
        MinY = MinY - 1
    
    Else
        MaxY = MaxY + 5
    End If
    
    If MapData(UserPos.x, UserPos.y).CharIndex = 0 And UserCharIndex > 0 Then
        UserPos.x = charlist(UserCharIndex).Pos.x
        UserPos.y = charlist(UserCharIndex).Pos.y
        MapData(UserPos.x, UserPos.y).CharIndex = UserCharIndex
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
    
    If UpdateLights Then
        Call RestaurarLuz
        Call MapUpdateGlobalLightRender
        UpdateLights = False
    End If
    
    Call SpriteBatch.BeginPrecalculated(StartX, StartY)

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

    ' Layer 2 & small objects loop
    Call DirectDevice.SetRenderState(D3DRS_ALPHATESTENABLE, True) ' Para no pisar los reflejos
    
    screenY = StartBufferedY

    For y = MinBufferedY To MaxBufferedY
        screenX = StartBufferedX

        For x = MinBufferedX To MaxBufferedX

            With MapData(x, y)
                
                ' Layer 2 *********************************
                If .Graphic(2).GrhIndex <> 0 Then
                    Call Draw_Grh(.Graphic(2), screenX, screenY, 1, 1, .light_value, , x, y)
                End If
            
            End With

            screenX = screenX + TilePixelWidth
        Next x

        screenY = screenY + TilePixelHeight
    Next y
    
 
    
    
    Dim grhSpellArea As grh
    grhSpellArea.GrhIndex = 20058
    
    Dim temp_color(3) As RGBA
    
    Call SetRGBA(temp_color(0), 255, 20, 25, 255)
    Call SetRGBA(temp_color(1), 0, 255, 25, 255)
    Call SetRGBA(temp_color(2), 55, 255, 55, 255)
    Call SetRGBA(temp_color(3), 145, 70, 70, 255)
    
   ' Call SetRGBA(MapData(15, 15).light_value(0), 255, 20, 20)
    'size 96x96 - mitad = 48
    If casteaArea And MouseX > 0 And MouseY > 0 And GetGameplayForm.MousePointer = 2 Then
        Call Draw_Grh(grhSpellArea, MouseX - 48, MouseY - 48, 0, 1, temp_color, True, , , 70)
    End If
    
     screenY = StartBufferedY

    For y = MinBufferedY To MaxBufferedY
        screenX = StartBufferedX

        For x = MinBufferedX To MaxBufferedX

            With MapData(x, y)
                If .Trap.GrhIndex > 0 Then
                    Call RGBAList(temp_color, 255, 255, 255, 100)
                    Call Draw_Grh(.Trap, screenX, screenY, 1, 1, temp_color, False)
                End If
                ' Objects *********************************
                If .ObjGrh.GrhIndex <> 0 Then
                    Select Case ObjData(.OBJInfo.ObjIndex).ObjType
                        Case eObjType.otArboles, eObjType.otPuertas, eObjType.otTeleport, eObjType.otCarteles, eObjType.OtPozos, eObjType.otYacimiento, eObjType.OtCorreo, eObjType.otFragua, eObjType.OtDecoraciones, eObjType.otFishingPool
                            Call Draw_Grh(.ObjGrh, screenX, screenY, 1, 1, .light_value)

                        Case Else
                            ' Objetos en el suelo (items, decorativos, etc)
                            
                             If ((.Blocked And FLAG_AGUA) <> 0) And .Graphic(2).GrhIndex = 0 Then
                             
                                object_angle = (object_angle + (timerElapsedTime * 0.002))
                                
                                .light_value(1).a = 85
                                .light_value(3).a = 85
                                
                                Call Draw_Grh_ItemInWater(.ObjGrh, screenX, screenY, False, False, .light_value, False, , , (object_angle + x * 45 + y * 90))
                                
                                .light_value(1).a = 255
                                .light_value(3).a = 255
                                .light_value(0).a = 255
                                .light_value(2).a = 255
                            Else
                                Call Draw_Grh(.ObjGrh, screenX, screenY, 1, 1, .light_value)
                            End If
                    End Select
                End If

            End With

            screenX = screenX + TilePixelWidth
        Next x

        screenY = screenY + TilePixelHeight
    Next y
    
    Call DirectDevice.SetRenderState(D3DRS_ALPHATESTENABLE, False)

    '  Layer 3 & chars
    screenY = StartBufferedY

    For y = MinBufferedY To MaxBufferedY
        screenX = StartBufferedX

        For x = MinBufferedX To MaxBufferedX
            
            With MapData(x, y)
                ' Chars ***********************************
                If .CharIndex = UserCharIndex Then 'evitamos reenderizar un clon del usuario
                    If x <> UserPos.x Or y <> UserPos.y Then
                        .CharIndex = 0
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
                            Call Draw_Grh(.CharFantasma.Escudo, screenX, screenY, 1, 1, TempColor(), False, x, y)
                            Call Draw_Grh(.CharFantasma.Body, screenX, screenY, 1, 1, TempColor(), False, x, y)
                            Call Draw_Grh(.CharFantasma.Head, screenX + .CharFantasma.OffX, screenY + .CharFantasma.Offy, 1, 1, TempColor(), False, x, y)
                            Call Draw_Grh(.CharFantasma.Casco, screenX + .CharFantasma.OffX, screenY + .CharFantasma.Offy, 1, 1, TempColor(), False, x, y)
                            Call Draw_Grh(.CharFantasma.Arma, screenX, screenY, 1, 1, TempColor(), False, x, y)
                        Else
                            Call Draw_Grh(.CharFantasma.Body, screenX, screenY, 1, 1, TempColor(), False, x, y)
                            Call Draw_Grh(.CharFantasma.Head, screenX + .CharFantasma.OffX, screenY + .CharFantasma.Offy, 1, 1, TempColor(), False, x, y)
                            Call Draw_Grh(.CharFantasma.Escudo, screenX, screenY, 1, 1, TempColor(), False, x, y)
                            Call Draw_Grh(.CharFantasma.Casco, screenX + .CharFantasma.OffX, screenY + .CharFantasma.Offy, 1, 1, TempColor(), False, x, y)
                            Call Draw_Grh(.CharFantasma.Arma, screenX, screenY, 1, 1, TempColor(), False, x, y)
                        End If

                    Else
                        .CharFantasma.Activo = False

                    End If

                End If
                
                If .CharIndex <> 0 Then
                    If charlist(.CharIndex).active = 1 Then
                        If mascota.visible And .CharIndex = UserCharIndex Then
                          '  Call Mascota_Render(.charindex, PixelOffsetX, PixelOffsetY)
                        End If
                        Call Char_Render(.CharIndex, screenX, screenY, x, y)
                    End If
                End If

                
            End With
            
            screenX = screenX + TilePixelWidth
        Next x
        ' Recorremos de nuevo esta fila para dibujar objetos grandes y capa 3 encima de chars
        screenX = StartBufferedX

        For x = MinBufferedX To MaxBufferedX

            With MapData(x, y)
                ' Objects *********************************
                If .ObjGrh.GrhIndex <> 0 Then
                           
                    Select Case ObjData(.OBJInfo.ObjIndex).ObjType
                         
                        Case eObjType.otArboles
                          
                            Call Draw_Sombra(.ObjGrh, screenX, screenY, 1, 1, False, x, y)

                            ' Debajo del arbol
                            If Abs(UserPos.x - x) < 3 And (Abs(UserPos.y - y)) < 8 And (Abs(UserPos.y) < y) Then
    
                                If .ArbolAlphaTimer <= 0 Then
                                    .ArbolAlphaTimer = lastMove
                                End If
    
                                DeltaTime = FrameTime - .ArbolAlphaTimer
    
                                Call Copy_RGBAList_WithAlpha(TempColor, .light_value, IIf(DeltaTime > ARBOL_ALPHA_TIME, ARBOL_MIN_ALPHA, 255 - DeltaTime / ARBOL_ALPHA_TIME * (255 - ARBOL_MIN_ALPHA)))
                                Call Draw_Grh(.ObjGrh, screenX, screenY, 1, 1, TempColor, False, x, y)
    
                            Else    ' Lejos del arbol
                                If .ArbolAlphaTimer = 0 Then
                                    Call Draw_Grh(.ObjGrh, screenX, screenY, 1, 1, .light_value, False, x, y)
    
                                Else
                                    If .ArbolAlphaTimer > 0 Then
                                        .ArbolAlphaTimer = -lastMove
                                    End If
    
                                    DeltaTime = FrameTime + .ArbolAlphaTimer
    
                                    If DeltaTime > ARBOL_ALPHA_TIME Then
                                        .ArbolAlphaTimer = 0
                                        Call Draw_Grh(.ObjGrh, screenX, screenY, 1, 1, .light_value, False, x, y)
                                    Else
                                        Call Copy_RGBAList_WithAlpha(TempColor, .light_value, ARBOL_MIN_ALPHA + DeltaTime * (255 - ARBOL_MIN_ALPHA) / ARBOL_ALPHA_TIME)
                                        Call Draw_Grh(.ObjGrh, screenX, screenY, 1, 1, TempColor, False, x, y)
                                    End If
                                End If
    
                            End If
                        
                        Case eObjType.otPuertas, eObjType.otTeleport, eObjType.otCarteles, eObjType.OtPozos, eObjType.otYacimiento, eObjType.OtCorreo, eObjType.otYunque, eObjType.otFragua, eObjType.OtDecoraciones
                            ' Objetos grandes (menos árboles)
                            Call Draw_Grh(.ObjGrh, screenX, screenY, 1, 1, .light_value, False, x, y)
                            
                        'Case Else
                        '    Call Draw_Grh(.ObjGrh, ScreenX, ScreenY, 1, 1, .light_value, False, x, y)
                    
                    End Select
                End If

                
                'Layer 3 **********************************
                If .Graphic(3).GrhIndex <> 0 Then

                    If (.Blocked And FLAG_ARBOL) <> 0 Then
                        
                        
                       ' Call Draw_Sombra(.Graphic(3), ScreenX, ScreenY, 1, 1, False, x, y)

                        ' Debajo del arbol
                        If Abs(UserPos.x - x) <= 3 And (Abs(UserPos.y - y)) < 12 And (Abs(UserPos.y) < y) Then

                            If .ArbolAlphaTimer <= 0 Then
                                .ArbolAlphaTimer = lastMove
                            End If

                            DeltaTime = FrameTime - .ArbolAlphaTimer

                            Call Copy_RGBAList_WithAlpha(TempColor, .light_value, IIf(DeltaTime > ARBOL_ALPHA_TIME, ARBOL_MIN_ALPHA, 255 - DeltaTime / ARBOL_ALPHA_TIME * (255 - ARBOL_MIN_ALPHA)))
                            Call Draw_Grh(.Graphic(3), screenX, screenY, 1, 1, TempColor, False, x, y)

                        Else    ' Lejos del arbol
                            If .ArbolAlphaTimer = 0 Then
                                Call Draw_Grh(.Graphic(3), screenX, screenY, 1, 1, .light_value, False, x, y)

                            Else
                                If .ArbolAlphaTimer > 0 Then
                                    .ArbolAlphaTimer = -lastMove
                                End If

                                DeltaTime = FrameTime + .ArbolAlphaTimer

                                If DeltaTime > ARBOL_ALPHA_TIME Then
                                    .ArbolAlphaTimer = 0
                                    Call Draw_Grh(.Graphic(3), screenX, screenY, 1, 1, .light_value, False, x, y)
                                Else
                                    Call Copy_RGBAList_WithAlpha(TempColor, .light_value, ARBOL_MIN_ALPHA + DeltaTime * (255 - ARBOL_MIN_ALPHA) / ARBOL_ALPHA_TIME)
                                    Call Draw_Grh(.Graphic(3), screenX, screenY, 1, 1, TempColor, False, x, y)
                                End If
                            End If

                        End If

                    Else
                        If AgregarSombra(.Graphic(3).GrhIndex) Then
                            Call Draw_Sombra(.Graphic(3), screenX, screenY, 1, 1, False, x, y)
                        End If

                        Call Draw_Grh(.Graphic(3), screenX, screenY, 1, 1, .light_value, False, x, y)

                    End If

                End If

            End With
            
            screenX = screenX + TilePixelWidth
        Next x
        screenY = screenY + TilePixelHeight
    Next y
    
    If InfoItemsEnRender And tX And tY Then
        With MapData(tX, tY)
            If .OBJInfo.ObjIndex Then
                If Not ObjData(.OBJInfo.ObjIndex).Agarrable Then
                    Dim Text As String, Amount As String
                    If .OBJInfo.Amount > 1000 Then
                        Amount = Round(.OBJInfo.Amount * 0.001, 1) & "K"
                    Else
                        Amount = .OBJInfo.Amount
                    End If
                    Text = ObjData(.OBJInfo.ObjIndex).Name & " (" & Amount & ")"
                    Call Engine_Text_Render(Text, MouseX + 15 + gameplay_render_offset.x, MouseY + gameplay_render_offset.y, COLOR_WHITE, , , , 160)
                End If
            End If
        End With
    End If

    ' Particles loop
    screenY = StartBufferedY

    For y = MinBufferedY To MaxBufferedY
        screenX = StartBufferedX

        For x = MinBufferedX To MaxBufferedX
            
            With MapData(x, y)
                ' Particles *******************************
                If .particle_group > 0 Then
                    Call Particle_Group_Render(.particle_group, screenX + 16, screenY + 16)
                End If
                '******************************************
            End With
            
            screenX = screenX + TilePixelWidth
        Next x

        screenY = screenY + TilePixelHeight
    Next y
    
    'draw projectiles
    Dim transform As Vector2
    Dim complete As Boolean
    Dim Index As Integer
    Index = 1
    Do While Index <= ActiveProjectile.CurrentIndex
        complete = UpdateProjectile(AllProjectile(ActiveProjectile.IndexInfo(Index)))
        Call WorldToScreen(AllProjectile(ActiveProjectile.IndexInfo(Index)).CurrentPos, transform, StartBufferedX, StartBufferedY, MinBufferedX, MinBufferedY)
        Call RenderProjectile(AllProjectile(ActiveProjectile.IndexInfo(Index)), transform, temp_color)
        If complete Then
            ReleaseProjectile (Index)
        Else
            Index = Index + 1
        End If
    Loop

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

        screenY = StartBufferedY

        For y = MinBufferedY To MaxBufferedY
            screenX = StartBufferedX
    
            For x = MinBufferedX To MaxBufferedX
            
                With MapData(x, y)
                    ' Layer 4 - roofs *******************************
                    If .Graphic(4).GrhIndex Then

                        Trigger = NearRoof(x, y)

                        If Trigger Then
                            Call Copy_RGBAList_WithAlpha(TempColor, .light_value, RoofsLight(Trigger))
                            Call Draw_Grh(.Graphic(4), screenX, screenY, 1, 1, TempColor, , x, y)
                        Else
                            Call Draw_Grh(.Graphic(4), screenX, screenY, 1, 1, .light_value, , x, y)
                        End If
    
                    End If
                    '******************************************
                End With

                screenX = screenX + TilePixelWidth
            Next x

            screenY = screenY + TilePixelHeight
        Next y
        
    End If
    
    
    
    
    If TieneAntorcha Then
    
        Dim randX As Double, randY As Double
        
        If GetTickCount - (10 * Rnd + 50) >= DeltaAntorcha Then
            randX = RandomNumber(-8, 0)
            randY = RandomNumber(-8, 0)
            
            DeltaAntorcha = GetTickCount
        End If
    Call Draw_GrhIndex(63333, randX, randY)
    
    End If
    
       
    If mascota.dialog <> "" And mascota.visible Then
        Call Engine_Text_Render(mascota.dialog, mascota.PosX + 14 - CInt(Engine_Text_Width(mascota.dialog, True) / 2) + 150, mascota.PosY - Engine_Text_Height(mascota.dialog, True) - 25 + 150, mascota_text_color(), 1, True, , mascota.Color(0).a)
    End If

    ' FXs and dialogs loop
    screenY = StartBufferedY

    For y = MinBufferedY To MaxBufferedY
        screenX = StartBufferedX

        For x = MinBufferedX To MaxBufferedX
            
            With MapData(x, y)


                ' Dialogs *******************************
                If MapData(x, y).CharIndex <> 0 Then
                
                    If charlist(.CharIndex).active = 1 Then
                    
                        Call Char_TextRender(.CharIndex, screenX, screenY, x, y)
                    
                    End If
                    
                End If
                '******************************************

                ' Render text value *******************************
                Dim i As Long
                If UBound(.DialogEffects) > 0 Then
                    For i = 1 To UBound(.DialogEffects)
                        With .DialogEffects(i)
                            If LenB(.Text) <> 0 Then
                                Dim DialogTime As Long
                                DialogTime = FrameTime - .Start
            
                                If DialogTime > .Duration Then
                                    .Text = vbNullString
                                Else
                                    If DialogTime > 900 Then
                                        Call RGBAList(TempColor, .Color.r, .Color.G, .Color.b, .Color.a * (1300 - DialogTime) * 0.0025)
                                    Else
                                        Call RGBAList(TempColor, .Color.r, .Color.G, .Color.b, .Color.a)
                                    End If
                                    If .Animated Then
                                        Engine_Text_Render_Efect 0, .Text, screenX + 16 - Int(Engine_Text_Width(.Text, False) * 0.5) + .offset.x, screenY - Engine_Text_Height(.Text, False) + .offset.y - DialogTime * 0.025, TempColor, 1, False
                                    Else
                                        Engine_Text_Render_Efect 0, .Text, screenX + 16 - Int(Engine_Text_Width(.Text, False) * 0.5) + .offset.x, screenY - Engine_Text_Height(.Text, False) + .offset.y, TempColor, 1, False
                                    End If
                                End If
                            End If
                        End With
                    Next
                End If
                '******************************************

                ' FXs *******************************
                If .FxCount > 0 Then
                    For i = 1 To .FxCount

                        If .FxList(i).FxIndex > 0 And .FxList(i).started <> 0 Then
    
                            Call RGBAList(TempColor, 255, 255, 255, 220)

                            If FxData(.FxList(i).FxIndex).IsPNG = 1 Then
                            
                                Call Draw_GrhFX(.FxList(i), screenX + FxData(.FxList(i).FxIndex).OffsetX, screenY + FxData(.FxList(i).FxIndex).OffsetY + 20, 1, 1, TempColor(), False)

                            Else
                                Call Draw_GrhFX(.FxList(i), screenX + FxData(.FxList(i).FxIndex).OffsetX, screenY + FxData(.FxList(i).FxIndex).OffsetY + 20, 1, 1, TempColor(), True)

                            End If

                        End If

                        If .FxList(i).started = 0 Then .FxList(i).FxIndex = 0

                    Next i

                    If .FxList(.FxCount).started = 0 Then .FxCount = .FxCount - 1

                End If
                '******************************************
                End With
            
            screenX = screenX + TilePixelWidth
        Next x

        screenY = screenY + TilePixelHeight
    Next y

    If MeteoParticle >= LBound(particle_group_list) And MeteoParticle <= UBound(particle_group_list) Then
        If particle_group_list(MeteoParticle).active Then
            If MapDat.LLUVIA Then
                'Screen positions were hardcoded by now
                screenX = 250
                screenY = 0
                
                Call Particle_Group_Render(MeteoParticle, screenX, screenY)
                
                LastOffsetX = ParticleOffsetX
                LastOffsetY = ParticleOffsetY
        
            End If
        
            If MapDat.NIEVE Then
            
                If Graficos_Particulas.Engine_MeteoParticle_Get <> 0 Then
                
                    'Screen positions were hardcoded by now
                    screenX = 250 + gameplay_render_offset.x
                    screenY = 0 + gameplay_render_offset.y
                    
                    Call Particle_Group_Render(MeteoParticle, screenX, screenY)
        
                End If
        
            End If
        Else
            MeteoParticle = 0
        End If
    End If

    If AlphaNiebla Then
    
        If MapDat.niebla Then Call Engine_Weather_UpdateFog

    End If
    
    Call Effect_Render_All
    
    If IsSet(FeatureToggles, eEnableHotkeys) And g_game_state.State = e_state_gameplay_screen Then
        Dim Color(3) As RGBA
        Call RGBAList(Color, 255, 255, 255, 200)
        Dim ArrowPos As Vector2
        ArrowPos.x = hotkey_arrow_posx
        ArrowPos.y = frmMain.renderer.Height - hotkey_arrow_posy
        If HideHotkeys Then
            Call DrawSingleGrh(HideArrowGrh, ArrowPos, 1, 270, Color)
        Else
            For i = 0 To 9
                Call DrawHotkey(i, i * 36 + hotkey_render_posX, frmMain.renderer.Height - hotkey_render_posY)
            Next
            Call DrawSingleGrh(HideArrowGrh, ArrowPos, 1, 90, Color)
            If gDragState.active Then
                Call Draw_GrhColor(gDragState.grh, gDragState.PosX - 16 - frmMain.renderer.Left, gDragState.PosY - frmMain.renderer.Top - 16, Color)
            End If
        End If
    End If
    
    Call renderCooldowns(710 + gameplay_render_offset.x, 25 + gameplay_render_offset.y)
    
    If InvasionActual Then
        
        Call Engine_Draw_Box(190 + gameplay_render_offset.x, 550 + gameplay_render_offset.y, 356, 36, RGBA_From_Comp(0, 0, 0, 200))
        Call Engine_Draw_Box(193 + gameplay_render_offset.x, 553 + gameplay_render_offset.y, 3.5 * InvasionPorcentajeVida, 30, RGBA_From_Comp(20, 196, 255, 200))
        
        Call Engine_Draw_Box(340 + gameplay_render_offset.x, 586 + gameplay_render_offset.y, 54, 9, RGBA_From_Comp(0, 0, 0, 200))
        Call Engine_Draw_Box(342 + gameplay_render_offset.x, 588 + gameplay_render_offset.y, 0.5 * InvasionPorcentajeTiempo, 5, RGBA_From_Comp(220, 200, 0, 200))
        
    End If

    If Pregunta Then
        
        Call Engine_Draw_Box(283 + gameplay_render_offset.x, 170 + gameplay_render_offset.y, 190, 100, RGBA_From_Comp(150, 20, 3, 200))
        Call Engine_Draw_Box(288 + gameplay_render_offset.x, 175 + gameplay_render_offset.y, 180, 90, RGBA_From_Comp(25, 25, 23, 200))

        Dim preguntaGrh As grh
        Call InitGrh(preguntaGrh, 32120)
        
        Call Engine_Text_Render(PreguntaScreen, 290 + gameplay_render_offset.x, 180 + gameplay_render_offset.y, COLOR_WHITE, 1, True)
        
        Call Draw_Grh(preguntaGrh, 416 + gameplay_render_offset.x, 233 + gameplay_render_offset.y, 1, 0, COLOR_WHITE, False, 0, 0, 0)

    End If

    If cartel Then

        Call RGBAList(TempColor, 255, 255, 255, 220)
        
        Dim TempGrh  As grh
        Call InitGrh(TempGrh, GrhCartel)
        
        Call Draw_Grh(TempGrh, CInt(clicX), CInt(clicY), 1, 0, TempColor(), False, 0, 0, 0)
        
        Call Engine_Text_Render(Leyenda, CInt(clicX - 100), CInt(clicY - 130), TempColor(), 1, False)

    End If

    Call RenderScreen_NombreMapa
    If PescandoEspecial Then
        Call RGBAList(ColorBarraPesca, 255, 255, 255)
        Dim grh As grh
        grh.GrhIndex = GRH_BARRA_PESCA
        Call Draw_Grh(grh, 239 + gameplay_render_offset.x, 550 + gameplay_render_offset.y, 0, 0, ColorBarraPesca())
        grh.GrhIndex = GRH_CURSOR_PESCA
        Call Draw_Grh(grh, 271 + PosicionBarra + gameplay_render_offset.x, 558 + gameplay_render_offset.y, 0, 0, ColorBarraPesca())
        frmDebug.add_text_tracebox PescandoEspecial
        For i = 1 To MAX_INTENTOS
            If intentosPesca(i) = 1 Then
                grh.GrhIndex = GRH_CIRCULO_VERDE
                Call Draw_Grh(grh, 394 + (i * 10) + gameplay_render_offset.x, 573 + gameplay_render_offset.y, 0, 0, ColorBarraPesca())
            ElseIf intentosPesca(i) = 2 Then
                grh.GrhIndex = GRH_CIRCULO_ROJO
                Call Draw_Grh(grh, 394 + (i * 10) + gameplay_render_offset.x, 573 + gameplay_render_offset.y, 0, 0, ColorBarraPesca())
            End If
        Next i
                
        If PosicionBarra <= 0 Then
            DireccionBarra = 1
            PuedeIntentar = True
        ElseIf PosicionBarra > 199 Then
            DireccionBarra = -1
            PuedeIntentar = True
        End If
        If PosicionBarra < 0 Then
            PosicionBarra = 0
        ElseIf PosicionBarra > 199 Then
            PosicionBarra = 199
        End If
        '90 - 111 es incluido (saca el pecesito)
        PosicionBarra = PosicionBarra + (DireccionBarra * VelocidadBarra * timerElapsedTime * 0.2)
        
        
        If (GetTickCount() - startTimePezEspecial) >= 20000 Then
            PescandoEspecial = False
            Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_PEZ_ROMPIO_LINEA_PESCA"), 255, 0, 0, 1, 0)
            Call WriteRomperCania
        End If
        
    End If
    
    If cartel_visible Then Call RenderScreen_Cartel
    Exit Sub

RenderScreen_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine_RenderScreen.RenderScreen", Erl)
    Resume Next
    
End Sub

Private Sub WorldToScreen(ByRef world As Vector2, ByRef screen As Vector2, ByVal screenX As Integer, ByVal screenY As Integer, ByVal tilesOffsetX As Integer, ByVal tilesOffsetY As Integer)
    screen.y = world.y - tilesOffsetY * TilePixelHeight + screenY
    screen.x = world.x - tilesOffsetX * TilePixelWidth + screenX
End Sub

Private Sub RenderProjectile(ByRef projetileInstance As Projectile, ByRef screenPos As Vector2, ByRef rgba_list() As RGBA)
On Error GoTo RenderProjectile_Err
    Call RGBAList(rgba_list, 255, 255, 255, 255)
    Call DrawSingleGrh(projetileInstance.GrhIndex, screenPos, 0, projetileInstance.Rotation, rgba_list)
    Exit Sub
RenderProjectile_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine_RenderScreen.RenderProjectile_Err", Erl)
End Sub

Function UpdateProjectile(ByRef Projectile As Projectile) As Boolean
On Error GoTo UpdateProjectile_Err
    Dim direction As Vector2
    direction = VSubs(Projectile.TargetPos, Projectile.CurrentPos)
    If VecLength(direction) < Projectile.speed * timerElapsedTime Then
        UpdateProjectile = True
        Exit Function
    End If
    Call Normalize(direction)
    direction = VMul(direction, Projectile.speed * timerElapsedTime)
    Projectile.CurrentPos = VAdd(Projectile.CurrentPos, direction)
    Projectile.Rotation = FixAngle(Projectile.Rotation + Projectile.RotationSpeed * timerElapsedTime)
    UpdateProjectile = False
    Exit Function
UpdateProjectile_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine_RenderScreen.UpdateProjectile_Err", Erl)
End Function

Public Sub InitializeProjectile(ByRef Projectile As Projectile, ByVal StartX As Byte, ByVal StartY As Byte, ByVal endX As Byte, ByVal endY As Byte, ByVal projectileType As Integer)
On Error GoTo InitializeProjectile_Err
    With ProjectileData(projectileType)
        Dim Index As Integer
        If AvailableProjectile.CurrentIndex > 0 Then
            Index = AvailableProjectile.IndexInfo(AvailableProjectile.CurrentIndex)
            AvailableProjectile.CurrentIndex = AvailableProjectile.CurrentIndex - 1
        Else
            'increase projectile active/ inactive/ instance arrays size
        End If
        AllProjectile(Index).CurrentPos.x = StartX * TilePixelWidth
        AllProjectile(Index).CurrentPos.y = StartY * TilePixelHeight
        AllProjectile(Index).TargetPos.x = endX * TilePixelWidth
        AllProjectile(Index).TargetPos.y = endY * TilePixelHeight
        AllProjectile(Index).speed = .speed
        AllProjectile(Index).RotationSpeed = .RotationSpeed
        AllProjectile(Index).GrhIndex = .grh
        If endX > StartX And .RigthGrh > 0 Then
            AllProjectile(Index).GrhIndex = .RigthGrh
            AllProjectile(Index).RotationSpeed = .RotationSpeed * -1
        End If

        AllProjectile(Index).Rotation = RadToDeg(GetAngle(StartX, endY, endX, StartY))
        AllProjectile(Index).Rotation = FixAngle(AllProjectile(Index).Rotation + .OffsetRotation)
        ActiveProjectile.CurrentIndex = ActiveProjectile.CurrentIndex + 1
        ActiveProjectile.IndexInfo(ActiveProjectile.CurrentIndex) = Index
    End With
    Exit Sub
InitializeProjectile_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine_RenderScreen.InitializeProjectile_Err", Erl)
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
        
        Call Grh_Render(letter_grh, 250 + gameplay_render_offset.x, 300 + gameplay_render_offset.y, Color())
        
        Call Engine_Text_RenderGrande(letter_text, 360 - Engine_Text_Width(letter_text, False, 4) / 2 + gameplay_render_offset.x, 1 + gameplay_render_offset.y, Color(), 5, False, , CInt(map_letter_a))

    End If

    
    Exit Sub

RenderScreen_NombreMapa_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine_RenderScreen.RenderScreen_NombreMapa", Erl)
    Resume Next
    
End Sub

Private Sub DrawHotkey(ByVal HkIndex As Integer, ByVal PosX As Integer, ByVal PosY As Integer)
    Call Draw_GrhIndex(GRH_INVENTORYSLOT, PosX, PosY)
    If HotkeyList(HkIndex).Index > 0 Then
        If HotkeyList(HkIndex).Type = e_HotkeyType.Item Then
            Call Draw_GrhIndex(ObjData(HotkeyList(HkIndex).Index).GrhIndex, PosX, PosY)
        ElseIf HotkeyList(HkIndex).Type = e_HotkeyType.Spell Then
            Call Draw_GrhIndex(HechizoData(HotkeyList(HkIndex).Index).IconoIndex, PosX, PosY)
        End If
    End If
    Call Engine_Text_Render(HkIndex + 1, PosX + 12, PosY, COLOR_WHITE, 1, True)
End Sub

