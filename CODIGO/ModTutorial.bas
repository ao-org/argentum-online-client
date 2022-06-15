Attribute VB_Name = "ModTutorial"
Option Explicit

Public cartel_fade As Single
Public cartel_fadestatus As Byte
Public cartel_icon As Long
Public cartel_text As String
Public cartel_title As String
Public cartel_duration As Long
Public cartel_visible As Boolean

Public tutorial_texto_actual As Byte

Public Const GRH_CARTEL_FONDO As Long = 22728

Public Sub RenderScreen_Cartel(ByVal Text As String, ColorTexto() As RGBA, ByVal icon_index As Long, ColorFondo() As RGBA, ColorTitulo() As RGBA, Optional ByVal title As String = "Argentum 20", Optional ByVal infinito As Boolean = False)
 On Error GoTo RenderScreen_Cartel_Err
    
    
    If cartel_fadestatus > 0 Then
        
        If cartel_fadestatus = 2 And cartel_duration > 0 Then
            cartel_duration = cartel_duration - IIf(infinito, 0, 1)
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
                cartel_text = ""
            End If

        End If

    End If
    
    cartel_visible = (cartel_fade > 0)
    
    'Renderizo cartel
    Dim background_grh As grh
    Call InitGrh(background_grh, GRH_CARTEL_FONDO)
    Call Grh_Render(background_grh, 350, 560, ColorFondo())
    
    If UserCharIndex > 0 Then
        'Renderizo titulo
        Call Engine_Text_Render_Cartel(title, 155, 444, ColorTitulo(), 5, False, , cartel_fade)
        'Renderizo texto
        Call Engine_Text_Render_Cartel(Text, 160, 500, ColorTexto(), 1, False, , cartel_fade)
        
        
        'Renderizo texto
        If infinito Then
            If language = e_language.English Then
                Call Engine_Text_Render_Cartel("Press left click to continue...", 476, 566, ColorTexto(), 1, False, , cartel_fade)
            Else
                Call Engine_Text_Render_Cartel("Click para continuar...", 499, 566, ColorTexto(), 1, False, , cartel_fade)
            End If
            Call Engine_Text_Render_Cartel(Text, 160, 500, ColorTexto(), 1, False, , cartel_fade)
        End If
        
        'Renderizo ícono
        Dim grh_icon As grh
        Dim ColorIcono(3) As RGBA
        Call RGBAList(ColorIcono(), 255, 255, 255, cartel_fade)
        Call InitGrh(grh_icon, icon_index)
        Call Grh_Render_Advance(grh_icon, 30, 430, 150, 150, ColorIcono())
    End If
    Exit Sub

RenderScreen_Cartel_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine_RenderScreen.RenderScreen_Cartel", Erl)
    Resume Next
End Sub
Public Sub toggleTutorialActivo(ByVal NumeroTutorial As Byte)
    Dim file As String
    With tutorial(NumeroTutorial)
        If .Activo = 0 Then
            .Activo = 1
        Else
            .Activo = 0
        End If
        
        file = App.Path & "\..\Recursos\INIT\tutoriales.ini"
            Call WriteVar(file, "TUTORIAL" & NumeroTutorial, "Activo", .Activo)
    End With
End Sub
Public Sub cargarTutoriales()
    
    
    Dim FileName As String
    Dim CantidadTutoriales As Long
    Dim file As clsIniManager
    Dim i As Long, j As Long
    
    FileName = App.Path & "\..\Recursos\INIT\tutoriales.ini"
    Set file = New clsIniManager
    Call file.Initialize(FileName)
    
    
    CantidadTutoriales = file.GetValue("INIT", "Cantidad")
    MostrarTutorial = file.GetValue("INIT", "MostrarTutorial")
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
            Call RenderScreen_Cartel(tutorial(1).textos(tutorial_texto_actual), ColorTextoCartel(), tutorial(1).grh, ColorFondo(), ColorTitulo(), tutorial(1).titulo, True)
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
    cartel_fade = 0
    cartel_fadestatus = 0
    cartel_icon = 0
    cartel_text = ""
    cartel_title = ""
    cartel_duration = 0
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
