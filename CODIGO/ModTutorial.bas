Attribute VB_Name = "ModTutorial"
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

Private cartel_title As String
Private cartel_message As String
Private cartel_icon As Long
Private cartel_duration As Long

Private cartel_fade As Single
Private cartel_fadestatus As Byte

'COLORES
Private cartel_title_color(3) As RGBA
Private cartel_message_color(3) As RGBA
Private cartel_background_color(3) As RGBA
Private cartel_icono_color(3) As RGBA
Private cartel_continue_color(3) As RGBA

'Color texto mascota
Public mascota_text_color(3) As RGBA

'GRHS
Private cartel_background_grh As grh
Private cartel_icon_grh As grh
Private Const GRH_CARTEL_FONDO As Long = 22728

'TAMAÑOS Y POSICIONES
Private cartel_title_pos_x As Long
Private cartel_title_pos_y As Long
Private cartel_message_pos_x As Long
Private cartel_message_pos_y As Long
Private cartel_grh_pos_x  As Long
Private cartel_grh_pos_y As Long
Private grh_width  As Long
Private grh_height  As Long
Private cartel_npc As Boolean

Private text_length As Integer
Private text_duration As Long
Private text_duration_total As Long
Private typing As Boolean
Private Const TYPING_SOUND = 230
Private sonido_activado As Boolean
Public tutorial_index As Integer
Public cartel_visible As Boolean
Private cartel_index As Byte
Private text_message_render As String
Public Enum e_tutorialIndex
    TUTORIAL_Muerto = 1
    TUTORIAL_ZONA_INSEGURA = 2
    TUTORIAL_NUEVO_USER = 3
    TUTORIAL_SkillPoints = 4
End Enum

Private enabled As Boolean

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Public Sub initMascotaTutorial()
    mascota.posX = 50
    mascota.posY = 50
    
    Call InitGrh(mascota.Body(1), 275, 1)
    Call InitGrh(mascota.Body(2), 275, 1)
    Call InitGrh(mascota.Body(3), 275, 1)
    Call InitGrh(mascota.Body(4), 275, 1)
    Call InitGrh(mascota.Body(5), 275, 1)
    Call InitGrh(mascota.Body(6), 275, 1)
    Call InitGrh(mascota.Body(7), 275, 1)
    Call InitGrh(mascota.Body(8), 275, 1)
End Sub
Private Function RGBA2Lng(r As Byte, G As Byte, B As Byte, A As Byte) As Long
    RGBA2Lng = r + G * 256 + B * 65536 + A * 16777216
End Function
Private Function Lng2RGBA(MyLong As Long) As Byte()
    Dim MyBytes(3) As Byte
    CopyMemory MyBytes(0), MyLong, 4
    Lng2RGBA = MyBytes
End Function
Public Sub nextCartel()
    If tutorial_index = 0 Then Exit Sub
    
    cartel_index = cartel_index + 1
    If cartel_index < UBound(tutorial(tutorial_index).textos()) Then
        text_duration = Len(cartel_message) * 20
        text_duration_total = text_duration
        If text_duration_total = 0 Then text_duration_total = 1
        sonido_activado = True
        Call Sound.Sound_Stop(TYPING_SOUND)
        Call Sound.Sound_Play(TYPING_SOUND)
        cartel_message = tutorial(tutorial_index).textos(cartel_index + 1)
    Else
        Call toggleTutorialActivo(tutorial_index)
        Call cerrarCartel
    End If
    
End Sub
Public Sub cerrarCartel()
    tutorial_index = 0
    cartel_index = 0
    cartel_duration = 0
    If mascota.visible Then mascota.visible = False
        Call Sound.Sound_Stop(TYPING_SOUND)
End Sub
Public Sub resetearCartel()
    tutorial_index = 0
    cartel_index = 0
    cartel_visible = False
End Sub
'Duration = 0 calcula solo con largo de texto, Duration = -1 infinito
Public Sub mostrarCartel(ByVal title As String, ByVal message As String, Optional ByVal icon As Long = 0, Optional ByVal duration As Long = 0, Optional ByVal titleColor As Long = -1, Optional ByVal messageColor As Long = -1, Optional ByVal backgroundColor As Long = -1, Optional ByVal esNpc As Boolean = False, Optional ByVal CartelTitlePosX = 0, Optional ByVal CartelTitlePosY As Long = 0, Optional ByVal CartelMessagePosX As Long = 0, Optional ByVal CartelMessagePosY As Long = 0, Optional ByVal cartelGrhPosX As Long = 0, Optional ByVal cartelGrhPosY As Long = 0, Optional ByVal grhWidth As Long = 0, Optional ByVal grhHeight As Long = 0)
    
    Dim titleColor_byte() As Byte
    Dim messageColor_byte() As Byte
    Dim backgroundColor_byte() As Byte
    
    If titleColor > -1 Then
        titleColor_byte = Lng2RGBA(titleColor)
        Call RGBAList(cartel_title_color(), titleColor_byte(0), titleColor_byte(1), titleColor_byte(2))
    Else
        Call RGBAList(cartel_title_color(), 255, 255, 255)
    End If
    
    If messageColor > -1 Then
        messageColor_byte = Lng2RGBA(messageColor)
        Call RGBAList(cartel_message_color(), messageColor_byte(0), messageColor_byte(1), messageColor_byte(2))
    Else
        Call RGBAList(cartel_message_color(), 255, 255, 255)
    End If
    
    If backgroundColor > -1 Then
        messageColor_byte = Lng2RGBA(backgroundColor)
        Call RGBAList(cartel_background_color(), backgroundColor_byte(0), backgroundColor_byte(1), backgroundColor_byte(2))
    Else
        Call RGBAList(cartel_background_color(), 255, 255, 255)
    End If
    
    'Inicializo GRH de fondo
    Call InitGrh(cartel_background_grh, GRH_CARTEL_FONDO)
    
    'Inicializo GRG de ícono
    Call InitGrh(cartel_icon_grh, icon)
    
    cartel_title = title
    cartel_message = message
    cartel_icon = icon
    cartel_duration = duration
    cartel_fadestatus = 1
    cartel_fade = 1
    cartel_visible = True
    
    cartel_title_pos_x = CartelTitlePosX
    cartel_title_pos_y = CartelTitlePosY
    cartel_message_pos_x = CartelMessagePosX
    cartel_message_pos_y = CartelMessagePosY
    
    cartel_grh_pos_x = cartelGrhPosX
    cartel_grh_pos_y = cartelGrhPosY
    grh_width = grhWidth
    grh_height = grhHeight
    text_message_render = cartel_message
    If Not esNpc Then
        text_length = Len(cartel_message)
        text_duration = Len(cartel_message) * 16
        text_duration_total = text_duration
        Call Sound.Sound_Stop(TYPING_SOUND)
        Call Sound.Sound_Play(TYPING_SOUND)
        If text_duration_total = 0 Then text_duration_total = 1
        sonido_activado = True
    End If
    cartel_fadestatus = 1
    cartel_fade = 1
    cartel_npc = esNpc
End Sub
Public Sub RenderScreen_Cartel()
 On Error GoTo RenderScreen_Cartel_Err
    
    
    If cartel_fadestatus > 0 Then
        
        If cartel_fadestatus = 2 And cartel_duration > 0 Then
            cartel_duration = cartel_duration - (timerTicksPerFrame * 40)
            If cartel_duration < 0 Then cartel_duration = 0
        End If
        
        If cartel_fadestatus = 1 Then
        
            cartel_fade = cartel_fade + (timerTicksPerFrame * 40)

            If cartel_fade >= 255 Then
                cartel_fade = 255
                cartel_fadestatus = 2
            End If

        ElseIf cartel_fadestatus = 2 And cartel_duration = 0 Then
            cartel_fade = cartel_fade - (timerTicksPerFrame * 40)

            If cartel_fade <= 0 Then
                cartel_fadestatus = 0
                cartel_fade = 0
                cartel_message = ""
            End If
        End If
    End If
    
    cartel_visible = (cartel_fade > 0)

    
    If Not cartel_npc Then
        Dim charCount As Integer
        charCount = (text_duration * text_length) / text_duration_total
        If charCount = 0 And sonido_activado Then
            Call Sound.Sound_Stop(TYPING_SOUND)
            sonido_activado = False
        End If
        text_message_render = Left(cartel_message, text_length - charCount)
        text_duration = text_duration - (timerTicksPerFrame * 40)
    End If
    
    If cartel_visible = False Then Exit Sub
        
    'Renderizo cartel
    Call RGBAList(cartel_background_color(), cartel_background_color(0).r, cartel_background_color(0).G, cartel_background_color(0).B, cartel_fade)
    'Call Grh_Render(cartel_background_grh, 350, 556, cartel_background_color())
    If Not cartel_npc Then
        Call Grh_Render_Advance(cartel_background_grh, 350, 615, 70, 644, cartel_background_color())
    End If
    
    If UserCharIndex > 0 Then
        'Renderizo titulo
        Call Engine_Text_Render_Cartel(cartel_title, cartel_title_pos_x, cartel_title_pos_y, cartel_title_color(), 5, False, , cartel_fade)
        'Renderizo texto
        Call Engine_Text_Render_Cartel(text_message_render, cartel_message_pos_x, cartel_message_pos_y, cartel_message_color(), 1, False, , cartel_fade)
                
        'Renderizo texto
        If cartel_duration = -1 Then
            Call RGBAList(cartel_continue_color(), 203, 156, 156, 255)
            If language = e_language.English Then
                Call Engine_Text_Render_Cartel("Click left to continue...", 516, 572, cartel_continue_color(), 1, False, , cartel_fade)
            Else
                Call Engine_Text_Render_Cartel("Click para continuar...", 539, 572, cartel_continue_color(), 1, False, , cartel_fade)
            End If
        End If
                
        'Renderizo ícono
        Call RGBAList(cartel_icono_color(), 255, 255, 255, cartel_fade)
        Call Grh_Render_Advance(cartel_icon_grh, cartel_grh_pos_x, cartel_grh_pos_y, grh_height, grh_width, cartel_icono_color())
    End If
    
    Exit Sub

RenderScreen_Cartel_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine_RenderScreen.RenderScreen_Cartel", Erl)
    Resume Next
End Sub
Private Function changeAlphaColor(Color() As RGBA, ByVal Alpha As Byte)
    Color(0).A = Alpha
    Color(1).A = Alpha
    Color(2).A = Alpha
    Color(3).A = Alpha
End Function

Public Sub toggleTutorialActivo(ByVal tutorial_index As Byte)
    Dim file As String
    With tutorial(tutorial_index)
        If .Activo = 0 Then
            .Activo = 1
        Else
            .Activo = 0
        End If
        Call SaveSetting("TUTORIAL" & tutorial_index, "Activo", .Activo)
    End With
End Sub
Public Sub cargarTutoriales()
    Dim CantidadTutoriales As Long
    Dim i As Long, j As Long
    
    CantidadTutoriales = GetSetting("INITTUTORIAL", "Cantidad")
    MostrarTutorial = GetSetting("INITTUTORIAL", "MostrarTutorial")
    If CantidadTutoriales <= 0 Then Exit Sub
    
    ReDim tutorial(1 To CantidadTutoriales)
    
    For i = 1 To CantidadTutoriales
        
        tutorial(i).grh = Val(GetSetting("TUTORIAL" & i, "Grh"))
        tutorial(i).Activo = Val(GetSetting("TUTORIAL" & i, "Activo"))
        tutorial(i).titulo = GetSetting("TUTORIAL" & i, IIf(language = e_language.English, "en_titulo", "titulo"))
        Dim CantidadTextos As Long
        CantidadTextos = Val(GetSetting("TUTORIAL" & i, "Cantidad"))
        ReDim tutorial(i).textos(1 To CantidadTextos)
        
        If CantidadTextos > 0 Then
            For j = 1 To CantidadTextos
                tutorial(i).textos(j) = GetSetting("TUTORIAL" & i, IIf(language = e_language.English, "en_texto" & j, "texto" & j))
            Next j
        End If
    Next i
End Sub


Public Sub Engine_Text_Render_Cartel(Texto As String, ByVal x As Integer, ByVal y As Integer, ByRef text_color() As RGBA, Optional ByVal font_index As Integer = 1, Optional multi_line As Boolean = False, Optional charindex As Integer = 0, Optional ByVal Alpha As Byte = 255)
    
    On Error GoTo Engine_Text_Render_Cartel_Err

    Dim A, B, c, d, e, f As Integer

    Dim graf          As grh

    Dim temp_array(3) As RGBA

    If charindex = 0 Then
        A = 255
    Else
        A = Clamp(charlist(charindex).AlphaText, 0, 255)
    End If

    If Alpha <> 255 Then
        A = Alpha
    End If
    
    Call RGBAList(temp_array, text_color(0).r, text_color(0).G, text_color(0).B, A)

    Dim i              As Long

    Dim removedDialogs As Long

    For i = 0 To dialogCount - 1

        'Decrease index to prevent jumping over a dialog
        'Crappy VB will cache the limit of the For loop, so even if it changed, it won't matter
        With dialogs(i - removedDialogs)

            If FrameTime - .startTime >= .lifeTime Then
                Call Char_Dialog_Remove(.charindex, charindex)
                             
                If A <= 0 Then
                    removedDialogs = removedDialogs + 1

                End If

            Else
            
            End If

        End With

    Next i

    Dim Sombra(3) As RGBA 'Sombra
    Call RGBAList(Sombra, text_color(0).r / 6, text_color(0).G / 6, text_color(0).B / 6, 0.8 * A)

    If (Len(Texto) = 0) Then Exit Sub

    d = 0

    If multi_line = False Then
        e = 0
        f = 0

        For A = 1 To Len(Texto)
            B = Asc(mid(Texto, A, 1))
            graf.GrhIndex = Fuentes(font_index).Caracteres(B)

            If B = 32 Or B = 13 Then
                If e >= 80 Then 'reemplazar por lo que os plazca
                    f = f + 1
                    e = 0
                    d = 0
                Else

                    If B = 32 Then d = d + 4

                End If

            Else

                If graf.GrhIndex > 12 Then

                    'mega sombra O-matica
                    graf.GrhIndex = Fuentes(font_index).Caracteres(B)

                    If font_index <> 3 Then
                        Call Draw_GrhFont(graf.GrhIndex, (x + d) + 1, y + 1 + f * 14, Sombra())

                    End If

                    Call Draw_GrhFont(graf.GrhIndex, (x + d), y + f * 14, temp_array())
                
                    ' graf.grhindex = Fuentes(font_index).Caracteres(b)
                    ' Grh_Render graf, (X + d), Y + f * 14, temp_array, False, False, False '14 es el height de esta fuente dsp lo hacemos dinamico
                    d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth

                End If

            End If

            e = e + 1
        Next A

    Else
        e = 0
        f = 0

        For A = 1 To Len(Texto)
            B = Asc(mid(Texto, A, 1))
            graf.GrhIndex = Fuentes(font_index).Caracteres(B)

            If B = 32 Or B = 13 Then
                If e >= 20 Then 'reemplazar por lo que os plazca
                    f = f + 1
                    e = 0
                    d = 0
                Else

                    If B = 32 Then d = d + 4

                End If

            Else

                If graf.GrhIndex > 12 Then

                    'mega sombra O-matica
                    graf.GrhIndex = Fuentes(font_index).Caracteres(B)
                    Call Draw_GrhFont(graf.GrhIndex, (x + d) + 1, y + 1 + f * 14, Sombra())
                    Call Draw_GrhFont(graf.GrhIndex, (x + d), y + f * 14, temp_array())
                
                    ' graf.grhindex = Fuentes(font_index).Caracteres(b)
                    'Grh_Render graf, (x + d), y + f * 14, temp_array, False, False, False '14 es el height de esta fuente dsp lo hacemos dinamico
                    If font_index = 4 Then
                        d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth - 1
                    Else
                        d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth

                    End If

                End If

            End If

            e = e + 1
        Next A

    End If

    
    Exit Sub

Engine_Text_Render_Cartel_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Textos.Engine_Text_Render_Cartel", Erl)
    Resume Next
    
End Sub

Public Sub checkTutorial()
    
    If charlist(UserCharIndex).Pos.x > 10 And charlist(UserCharIndex).Pos.y > 10 And charlist(UserCharIndex).Pos.x < 80 And charlist(UserCharIndex).Pos.y < 80 Then
        
        If Not mascota.visible Then
            Call RGBAList(mascota_text_color, 211, 153, 94, 255)
            mascota.dialog = "Bienvenido, soy tu guia de entrenamiento en las tierras de Argentum 20, estaré siguiendo todos tus movimientos para que te conviertas en un enorme guerrero."
            Call InitGrh(mascota.fx, 4841, FrameTime, 0)
            mascota.fx.speed = mascota.fx.speed / 2
            Call RGBAList(mascota.color, 255, 255, 255, 0)
            mascota.visible = True
        End If
        'If MostrarTutorial And tutorial_index <= 0 Then
        '    If tutorial(e_tutorialIndex.TUTORIAL_NUEVO_USER).Activo = 1 Then
        '        tutorial_index = e_tutorialIndex.TUTORIAL_NUEVO_USER
        '        mascota.visible = True
        '        Call mostrarCartel(tutorial(tutorial_index).titulo, tutorial(tutorial_index).textos(1), 275, -1, &H164B8A, , , False, 100, 479, 100, 535, 640, 490, 64, 64)
        '    End If
        'End If
    Else
        mascota.visible = False
        'mascota.dialog = ""
    End If
    'If charlist(UserCharIndex).Pos.x >= 27 And charlist(UserCharIndex).Pos.y = 14 And charlist(UserCharIndex).Pos.x <= 30 And charlist(UserCharIndex).Pos.y = 14 And Not enabled Then
    '    Call RGBAList(mascota_text_color, 211, 153, 94, 255)
   '     mascota.dialog = "Estás saliendo a una zona insegura, ten en cuenta que podrás ser atacado por otros."
   '     enabled = True
   ' Else
  '      enabled = False
        'mascota.dialog = ""
  '  End If
End Sub
