Attribute VB_Name = "ModTutorial"
Option Explicit

Public cartel_title As String
Public cartel_message As String
Public cartel_icon As Long
Public cartel_duration As Long

Public cartel_fade As Single
Public cartel_fadestatus As Byte
Public cartel_visible As Boolean

Private cartel_title_color(3) As RGBA
Private cartel_message_color(3) As RGBA
Private cartel_background_color(3) As RGBA
Private cartel_icono_color(3) As RGBA
Private cartel_continue_color(3) As RGBA

Private cartel_background_grh As grh
Private cartel_icon_grh As grh

Public tutorial_texto_actual As Byte


'TAMAÑOS Y POSICIONES

Private cartel_title_pos_x As Long
Private cartel_title_pos_y As Long
Private cartel_message_pos_x As Long
Private cartel_message_pos_y As Long
Private cartel_npc As Boolean

Public grh_width  As Long
Public grh_height  As Long
Public cartel_grh_pos_x  As Long
Public cartel_grh_pos_y As Long

Public tutorial_index As Integer

Public Const GRH_CARTEL_FONDO As Long = 22728

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)

Private Function RGBA2Lng(r As Byte, G As Byte, B As Byte, A As Byte) As Long
    RGBA2Lng = r + G * 256 + B * 65536 + A * 16777216
End Function
Private Function Lng2RGBA(MyLong As Long) As Byte()
    Dim MyBytes(3) As Byte
    CopyMemory MyBytes(0), MyLong, 4
    Lng2RGBA = MyBytes
End Function

'Duration = 0 calcula solo con largo de texto, Duration = -1 infinito
Public Sub mostrarCartel(ByVal title As String, ByVal message As String, Optional ByVal icon As Long = 0, Optional ByVal duration As Long = 0, Optional ByVal titleColor As Long = -1, Optional ByVal messageColor As Long = -1, Optional ByVal backgroundColor As Long = -1, Optional ByVal esNpc As Boolean = False, Optional ByVal CartelTitlePosX = 0, Optional ByVal CartelTitlePosY As Long = 0, Optional ByVal CartelMessagePosX As Long = 0, Optional ByVal CartelMessagePosY As Long = 0)
    
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
    
    cartel_npc = esNpc
End Sub
Public Sub RenderScreen_Cartel()
 On Error GoTo RenderScreen_Cartel_Err
    
    
    If cartel_fadestatus > 0 Then
        
        If cartel_fadestatus = 2 And cartel_duration > 0 Then
            Debug.Print cartel_duration
            cartel_duration = cartel_duration - IIf(cartel_duration = -1, 0, 1)
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
        Call Engine_Text_Render_Cartel(cartel_message, cartel_message_pos_x, cartel_message_pos_y, cartel_message_color(), 1, False, , cartel_fade)
                
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

Public Sub toggleTutorialActivo(ByVal NumeroTutorial As Byte)
    Dim file As String
    With tutorial(NumeroTutorial)
        If .Activo = 0 Then
            .Activo = 1
        Else
            .Activo = 0
        End If
        
        file = App.Path & "\..\Recursos\OUTPUT\Configuracion.ini"
        Call WriteVar(file, "TUTORIAL" & NumeroTutorial, "Activo", .Activo)
    End With
End Sub
Public Sub cargarTutoriales()
    
    
    Dim FileName As String
    Dim CantidadTutoriales As Long
    Dim file As clsIniManager
    Dim i As Long, j As Long
    FileName = App.Path & "\..\Recursos\OUTPUT\Configuracion.ini"
    Set file = New clsIniManager
    Call file.Initialize(FileName)
    
    
    CantidadTutoriales = file.GetValue("INITTUTORIAL", "Cantidad")
    MostrarTutorial = file.GetValue("INITTUTORIAL", "MostrarTutorial")
    If CantidadTutoriales <= 0 Then Exit Sub
    
    ReDim tutorial(1 To CantidadTutoriales)
    
    For i = 1 To CantidadTutoriales
        
        tutorial(i).grh = Val(file.GetValue("TUTORIAL" & i, "Grh"))
        tutorial(i).Activo = Val(file.GetValue("TUTORIAL" & i, "Activo"))
        tutorial(i).titulo = file.GetValue("TUTORIAL" & i, IIf(language = e_language.English, "en_titulo", "titulo"))
        Dim CantidadTextos As Long
        CantidadTextos = Val(file.GetValue("TUTORIAL" & i, "Cantidad"))
        ReDim tutorial(i).textos(1 To CantidadTextos)
        
        If CantidadTextos > 0 Then
            For j = 1 To CantidadTextos
                tutorial(i).textos(j) = file.GetValue("TUTORIAL" & i, IIf(language = e_language.English, "en_texto" & j, "texto" & j))
            Next j
        End If
    Next i
    
    Set file = Nothing
    
End Sub

Public Sub RenderizarTutoriales()
    If UserEstado = 1 And tutorial(1).Activo = 1 And tutorial(1).Mostrando Then
        If tutorial_texto_actual <= UBound(tutorial(1).textos) Then
            Dim ColorTextoCartel(3) As RGBA
            Call RGBAList(ColorTextoCartel(), 213, 250, 255, 255)
            Dim ColorTitulo(3) As RGBA
            Call RGBAList(ColorTitulo(), 191, 0, 0, 255)
            Dim ColorFondo(3) As RGBA
            Call RGBAList(ColorFondo(), 202, 28, 2, 255)
            Call RenderScreen_Cartel
        Else
            tutorial_texto_actual = 0
            tutorial(1).Mostrando = False
            Call toggleTutorialActivo(1)
        End If
    End If
End Sub

Public Function MostrandoTutorial() As Long
    Dim i As Long
    
    For i = 1 To UBound(tutorial)
        If tutorial(i).Mostrando Then
            MostrandoTutorial = i
            Exit Function
        End If
    Next i
End Function

Public Sub ResetearCartel()
    Dim i As Long
    
    For i = 1 To UBound(tutorial)
        tutorial(i).Mostrando = False
    Next i
    
    cartel_fade = 0
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
