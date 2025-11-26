Attribute VB_Name = "TextRenderer"
Option Explicit

Public Function Engine_Text_Height(Texto As String, Optional multi As Boolean = False, Optional font As Byte = 1) As Integer
    On Error GoTo Engine_Text_Height_Err
    Dim A As Integer, B As Integer, d As Integer, e As Integer, f As Integer
    Dim graf As Grh
    If multi = False Then
        Engine_Text_Height = 0
    Else
        e = 0
        f = 0
        If font = 1 Then
            For A = 1 To Len(Texto)
                B = Asc(mid(Texto, A, 1))
                graf.GrhIndex = Fuentes(1).Caracteres(B)
                If B = 32 Or B = 13 Then
                    If e >= 20 Then 'reemplazar por lo que os plazca
                        f = f + 1
                        e = 0
                        d = 0
                    Else
                        If B = 32 Then
                            d = d + 4
                        End If
                    End If
                    'Else
                    'If graf.GrhIndex > 12 Then
                End If
                e = e + 1
            Next A
        Else
            For A = 1 To Len(Texto)
                B = Asc(mid(Texto, A, 1))
                graf.GrhIndex = Fuentes(font).Caracteres(B)
                If B = 32 Or B = 13 Then
                    If e >= 20 Then 'reemplazar por lo que os plazca
                        f = f + 1
                        e = 0
                        d = 0
                    Else
                        If B = 32 Then
                            d = d + 4
                        End If
                    End If
                    'Else
                    'If graf.GrhIndex > 12 Then
                End If
                e = e + 1
            Next A
        End If
        Engine_Text_Height = f * 14
    End If
    Exit Function
Engine_Text_Height_Err:
    Call RegistrarError(Err.Number, Err.Description, "TextRenderer.Engine_Text_Height", Erl)
    Resume Next
End Function

Public Sub RenderText(ByVal Texto As String, _
                      ByVal x As Integer, _
                      ByVal y As Integer, _
                      ByRef text_color() As RGBA, _
                      Optional ByVal font_index As Integer = 1, _
                      Optional multi_line As Boolean = False, _
                      Optional charindex As Integer = 0, _
                      Optional ByVal alpha As Byte = 255)
    On Error GoTo RenderText_Err
    If (Len(Texto) = 0) Then Exit Sub
    Dim A As Integer, B As Integer, d As Integer, e As Integer, f As Integer
    Dim graf As Grh
    If charindex = 0 Then
        A = 255
    Else
        A = charlist(charindex).AlphaText
    End If
    If alpha <> 255 Then
        A = alpha
    End If
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
            End If
        End With
    Next i
    d = 0
    If multi_line = False Then
        e = 0
        f = 0
        For A = 1 To Len(Texto)
            B = Asc(mid(Texto, A, 1))
            graf.GrhIndex = Fuentes(font_index).Caracteres(B)
            If B = 32 Or B = 13 Then
                If e >= 30 Then 'reemplazar por lo que os plazca
                    f = f + 1
                    e = 0
                    d = 0
                Else
                    If B = 32 Then d = d + 2
                End If
            Else
                If graf.GrhIndex > 12 Then
                    graf.GrhIndex = Fuentes(font_index).Caracteres(B)
                    Call Draw_GrhFont(graf.GrhIndex, (x + d) + 1, y + 1 + f * 14, text_color())
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
                If e >= 33 Then 'reemplazar por lo que os plazca
                    f = f + 1
                    e = 0
                    d = 0
                Else
                    If B = 32 Then d = d + 2
                End If
            Else
                If graf.GrhIndex > 12 Then
                    graf.GrhIndex = Fuentes(font_index).Caracteres(B)
                    Call Draw_GrhFont(graf.GrhIndex, (x + d), y + f * 14, text_color())
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
RenderText_Err:
    Call RegistrarError(Err.Number, Err.Description, "TextRenderer.RenderText", Erl)
    Resume Next
End Sub

Public Sub Engine_Text_Render(Texto As String, _
                              ByVal x As Integer, _
                              ByVal y As Integer, _
                              ByRef text_color() As RGBA, _
                              Optional ByVal font_index As Integer = 1, _
                              Optional multi_line As Boolean = False, _
                              Optional charindex As Integer = 0, _
                              Optional ByVal alpha As Byte = 255)
    On Error GoTo Engine_Text_Render_Err
    Dim A As Integer, B As Integer, d As Integer, e As Integer, f As Integer
    Dim graf          As Grh
    Dim temp_array(3) As RGBA
    If charindex = 0 Then
        A = 255
    Else
        A = Clamp(charlist(charindex).AlphaText, 0, 255)
    End If
    If alpha <> 255 Then
        A = alpha
    End If
    Call RGBAList(temp_array, text_color(0).R, text_color(0).G, text_color(0).B, A)
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
    Call RGBAList(Sombra, text_color(0).R / 6, text_color(0).G / 6, text_color(0).B / 6, 0.8 * A)
    If (Len(Texto) = 0) Then Exit Sub
    d = 0
    If multi_line = False Then
        e = 0
        f = 0
        For A = 1 To Len(Texto)
            B = Asc(mid(Texto, A, 1))
            graf.GrhIndex = Fuentes(font_index).Caracteres(B)
            If B = 32 Or B = 13 Then
                If e >= 32 Then 'reemplazar por lo que os plazca
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
Engine_Text_Render_Err:
    Call RegistrarError(Err.Number, Err.Description, "TextRenderer.Engine_Text_Render", Erl)
    Resume Next
End Sub

Public Sub simple_text_render(Texto As String, _
                              ByVal x As Integer, _
                              ByVal y As Integer, _
                              ByRef text_color() As RGBA, _
                              Optional ByVal font_index As Integer = 1, _
                              Optional multi_line As Boolean = False, _
                              Optional charindex As Integer = 0, _
                              Optional ByVal alpha As Byte = 255)
    On Error GoTo Engine_Text_Render_Err
    Dim A As Integer, B As Integer, d As Integer, e As Integer, f As Integer
    Dim graf          As Grh
    Dim temp_array(3) As RGBA
    If charindex = 0 Then
        A = 255
    Else
        A = Clamp(charlist(charindex).AlphaText, 0, 255)
    End If
    If alpha <> 255 Then
        A = alpha
    End If
    Call RGBAList(temp_array, text_color(0).R, text_color(0).G, text_color(0).B, A)
    Dim i         As Long
    Dim Sombra(3) As RGBA 'Sombra
    Call RGBAList(Sombra, text_color(0).R / 6, text_color(0).G / 6, text_color(0).B / 6, 0.8 * A)
    If (Len(Texto) = 0) Then Exit Sub
    d = 0
    f = 0
    For A = 1 To Len(Texto)
        B = Asc(mid(Texto, A, 1))
        graf.GrhIndex = Fuentes(font_index).Caracteres(B)
        If graf.GrhIndex > 12 Then
            'mega sombra O-matica
            graf.GrhIndex = Fuentes(font_index).Caracteres(B)
            Call Draw_GrhFont(graf.GrhIndex, (x + d) + 1, y + 1 + f * 14, Sombra())
            Call Draw_GrhFont(graf.GrhIndex, (x + d), y + f * 14, temp_array())
            ' graf.grhindex = Fuentes(font_index).Caracteres(b)
            If font_index = 4 Then
                d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth - 1
            Else
                d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth
            End If
        End If
    Next A
    Exit Sub
Engine_Text_Render_Err:
    Call RegistrarError(Err.Number, Err.Description, "TextRenderer.Engine_Text_Render", Erl)
    Resume Next
End Sub

Public Sub Engine_Text_Render_No_Ladder(Texto As String, _
                                        ByVal x As Integer, _
                                        ByVal y As Integer, _
                                        ByRef text_color() As RGBA, _
                                        ByVal status As Byte, _
                                        Optional ByVal font_index As Integer = 1, _
                                        Optional multi_line As Boolean = False, _
                                        Optional charindex As Integer = 0, _
                                        Optional ByVal alpha As Byte = 255)
    On Error GoTo Engine_Text_Render_Err
    Dim A         As Integer, B As Integer, c As Integer, d As Integer
    Dim graf      As Grh
    Dim color1(3) As RGBA
    Dim color2(3) As RGBA
    If charindex = 0 Then
        A = 255
    Else
        A = Clamp(charlist(charindex).AlphaText, 0, 255)
    End If
    If alpha <> 255 Then
        A = alpha
    End If
    Select Case status
        Case 0 'criminal
            Call RGBAList(color1, 225, 0, 0, A)
            Call RGBAList(color2, 255, 255, 255, A)
        Case 1 'ciuda
            Call RGBAList(color1, 0, 128, 255, A)
            Call RGBAList(color2, 255, 255, 255, A)
        Case 2 'legión oscura
            Call RGBAList(color1, 155, 0, 0, A)
            Call RGBAList(color2, 255, 255, 255, A)
        Case 3 'armada real
            Call RGBAList(color1, 0, 175, 255, A)
            Call RGBAList(color2, 255, 255, 255, A)
        Case 4 'Legión
            Call RGBAList(color1, 155, 0, 0, A)
            Call RGBAList(color2, 255, 255, 255, A)
        Case 5 'Consejo
            Call RGBAList(color1, 22, 239, 253, A)
            Call RGBAList(color2, 255, 255, 255, A)
        Case 7 'aviso solicitud
            Call RGBAList(color2, 255, 255, 0, A)
        Case 8 'aviso desconectado
            Call RGBAList(color2, 255, 0, 0, A)
        Case 9 'aviso conectado
            Call RGBAList(color2, 10, 182, 70, A)
        Case 10 'lider
            Call RGBAList(color1, 222, 194, 112, A)
            Call RGBAList(color2, 255, 255, 255, A)
    End Select
    'Call RGBAList(color2, 255, 255, 255, A)
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
    Call RGBAList(Sombra, text_color(0).R / 6, text_color(0).G / 6, text_color(0).B / 6, 0.8 * A)
    If (Len(Texto) = 0) Then Exit Sub
    Dim row As Integer, charPos As Integer
    d = 0
    row = 0
    charPos = 0
    Dim separador As Boolean
    For A = 1 To Len(Texto)
        B = Asc(mid(Texto, A, 1))
        graf.GrhIndex = Fuentes(font_index).Caracteres(B)
        If B = 1 Then separador = Not separador
        If graf.GrhIndex > 12 Then
            'mega sombra O-matica
            graf.GrhIndex = Fuentes(font_index).Caracteres(B)
            If font_index <> 3 Then
                Call Draw_GrhFont(graf.GrhIndex, (x + d) + 1, y + 1 + 10, Sombra())
            End If
            If status >= 0 And status <= 5 Or status = 10 Then
                If separador Then
                    Call Draw_GrhFont(graf.GrhIndex, (x + d), y + 10, color1)
                Else
                    Call Draw_GrhFont(graf.GrhIndex, (x + d), y + 10, color2)
                End If
            Else
                Call Draw_GrhFont(graf.GrhIndex, (x + d), y + 10, color2)
            End If
            d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth
        End If
        charPos = charPos + 1
    Next A
    Exit Sub
Engine_Text_Render_Err:
    Call RegistrarError(Err.Number, Err.Description, "TextRenderer.Engine_Text_Render", Erl)
    Resume Next
End Sub

Public Sub Engine_Text_RenderGrande(Texto As String, _
                                    x As Integer, _
                                    y As Integer, _
                                    ByRef text_color() As RGBA, _
                                    Optional ByVal font_index As Integer = 1, _
                                    Optional multi_line As Boolean = False, _
                                    Optional charindex As Integer = 0, _
                                    Optional ByVal alpha As Byte = 255)
    On Error GoTo Engine_Text_RenderGrande_Err
    Dim A As Integer, B As Integer, d As Integer, e As Integer, f As Integer
    Dim graf          As Grh
    Dim temp_array(3) As RGBA
    If charindex = 0 Then
        A = 255
    Else
        A = charlist(charindex).AlphaText
    End If
    If alpha <> 255 Then
        A = alpha
    End If
    Call RGBAList(temp_array, text_color(0).R, text_color(0).G, text_color(0).B, A)
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
    Call RGBAList(Sombra, text_color(0).R / 6, text_color(0).G / 6, text_color(0).B / 6, 0.8 * alpha)
    If (Len(Texto) = 0) Then Exit Sub
    d = 0
    If multi_line = False Then
        e = 0
        f = 0
        For A = 1 To Len(Texto)
            B = Asc(mid(Texto, A, 1))
            graf.GrhIndex = Fuentes(font_index).Caracteres(B)
            If B = 32 Or B = 13 Then
                If e >= 35 Then 'reemplazar por lo que os plazca
                    f = f + 1
                    e = 0
                    d = 0
                Else
                    If B = 32 Then d = d + 12
                End If
            Else
                If graf.GrhIndex > 12 Then
                    'mega sombra O-matica
                    graf.GrhIndex = Fuentes(font_index).Caracteres(B)
                    If font_index <> 3 Then
                        Call Draw_GrhFont(graf.GrhIndex, (x + d), y + f * 14, Sombra())
                    End If
                    Call Draw_GrhFont(graf.GrhIndex, (x + d) + 1, y + 1 + f * 14, temp_array())
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
                If e >= 10 Then 'reemplazar por lo que os plazca
                    f = f + 3
                    e = 0
                    d = 0
                Else
                    If B = 32 Then d = d + 12
                End If
            Else
                If graf.GrhIndex > 12 Then
                    'mega sombra O-matica
                    graf.GrhIndex = Fuentes(font_index).Caracteres(B)
                    'Call Draw_GrhColor(graf.GrhIndex, (x + d) + 1, y + 1 + f * 14, Sombra())
                    Call Draw_GrhFont(graf.GrhIndex, (x + d), y + f * 14, temp_array())
                    ' graf.grhindex = Fuentes(font_index).Caracteres(b)
                    'Grh_Render graf, (x + d), y + f * 14, temp_array, False, False, False '14 es el height de esta fuente dsp lo hacemos dinamico
                    If font_index = 4 Then
                        d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth
                    Else
                        d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth
                    End If
                End If
            End If
            e = e + 1
        Next A
    End If
    Exit Sub
Engine_Text_RenderGrande_Err:
    Call RegistrarError(Err.Number, Err.Description, "TextRenderer.Engine_Text_RenderGrande", Erl)
    Resume Next
End Sub

Public Sub Engine_Text_Render2(Texto As String, _
                               x As Integer, _
                               y As Integer, _
                               ByRef text_color As RGBA, _
                               Optional ByVal font_index As Integer = 1, _
                               Optional multi_line As Boolean = False, _
                               Optional charindex As Long = 0, _
                               Optional ByVal alpha As Boolean = False)
    On Error GoTo Engine_Text_Render2_Err
    Dim A As Integer, B As Integer, d As Integer, e As Integer, f As Integer
    Dim graf          As Grh
    Dim temp_array(3) As RGBA
    Call RGBAList(temp_array, text_color.R, text_color.G, text_color.B, text_color.A)
    Dim Sombra(3) As RGBA 'Sombra
    Call RGBAList(Sombra, text_color.R / 6, text_color.G / 6, text_color.B / 6, 0.8 * text_color.A)
    If (Len(Texto) = 0) Then Exit Sub
    d = 0
    If multi_line = False Then
        e = 0
        f = 0
        For A = 1 To Len(Texto)
            B = Asc(mid(Texto, A, 1))
            graf.GrhIndex = Fuentes(font_index).Caracteres(B)
            If B = 32 Or B = 13 Then
                If e >= 35 Then 'reemplazar por lo que os plazca
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
                    If font_index <> 3 Then
                        d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth
                    Else
                        d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth
                    End If
                End If
            End If
            e = e + 1
        Next A
    End If
    Exit Sub
Engine_Text_Render2_Err:
    Call RegistrarError(Err.Number, Err.Description, "TextRenderer.Engine_Text_Render2", Erl)
    Resume Next
End Sub

Public Sub Engine_Text_Render_Efect(charindex As Integer, _
                                    Texto As String, _
                                    x As Integer, _
                                    y As Integer, _
                                    ByRef text_color() As RGBA, _
                                    Optional ByVal font_index As Integer = 1, _
                                    Optional multi_line As Boolean = False)
    On Error GoTo Engine_Text_Render_Efect_Err
    Dim A As Integer, B As Integer, d As Integer, e As Integer, f As Integer
    Dim graf As Grh
    If (Len(Texto) = 0) Then Exit Sub
    d = 0
    e = 0
    f = 0
    Dim Sombra(3) As RGBA 'Sombra
    Call RGBAList(Sombra, text_color(0).R / 6, text_color(0).G / 6, text_color(0).B / 6, 0.8 * text_color(0).A)
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
                Call Draw_GrhFont(graf.GrhIndex, (x + d), y + f * 14, text_color())
                ' graf.grhindex = Fuentes(font_index).Caracteres(b)
                'Grh_Render graf, (x + d), y + f * 14, temp_array, False, False, False '14 es el height de esta fuente dsp lo hacemos dinamico
                d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth
            End If
        End If
        e = e + 1
    Next A
    Exit Sub
Engine_Text_Render_Efect_Err:
    Call RegistrarError(Err.Number, Err.Description, "TextRenderer.Engine_Text_Render_Efect", Erl)
    Resume Next
End Sub

Public Function Engine_Text_Width(Texto As String, Optional multi As Boolean = False, Optional Fon As Byte = 1) As Integer
    On Error GoTo Engine_Text_Width_Err
    Dim A As Integer, B As Integer, d As Integer, e As Integer, f As Integer
    Dim graf As Grh
    Select Case Fon
        Case 1
            If multi = False Then
                For A = 1 To Len(Texto)
                    B = Asc(mid(Texto, A, 1))
                    graf.GrhIndex = Fuentes(1).Caracteres(B)
                    If graf.GrhIndex = 0 Then graf.GrhIndex = 1
                    If B <> 32 Then
                        Engine_Text_Width = Engine_Text_Width + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1
                    Else
                        Engine_Text_Width = Engine_Text_Width + 4
                    End If
                Next A
            Else
                e = 0
                f = 0
                For A = 1 To Len(Texto)
                    B = Asc(mid(Texto, A, 1))
                    graf.GrhIndex = Fuentes(1).Caracteres(B)
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
                            d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1
                            If d > Engine_Text_Width Then Engine_Text_Width = d
                        End If
                    End If
                    e = e + 1
                Next A
            End If
        Case 4
            If multi = False Then
                For A = 1 To Len(Texto)
                    B = Asc(mid(Texto, A, 1))
                    graf.GrhIndex = Fuentes(Fon).Caracteres(B)
                    If graf.GrhIndex = 0 Then graf.GrhIndex = 1
                    If B <> 20 Then
                        Engine_Text_Width = Engine_Text_Width + GrhData(GrhData(graf.GrhIndex + 1).Frames(1)).pixelWidth + 10
                    Else
                        Engine_Text_Width = Engine_Text_Width - 15
                    End If
                Next A
            Else
                e = 0
                f = 0
                For A = 1 To Len(Texto)
                    B = Asc(mid(Texto, A, 1))
                    graf.GrhIndex = Fuentes(Fon).Caracteres(B)
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
                            d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1
                            If d > Engine_Text_Width Then Engine_Text_Width = d
                        End If
                    End If
                    e = e + 1
                Next A
            End If
        Case Else
            If multi = False Then
                For A = 1 To Len(Texto)
                    B = Asc(mid(Texto, A, 1))
                    graf.GrhIndex = Fuentes(Fon).Caracteres(B)
                    If graf.GrhIndex = 0 Then graf.GrhIndex = 1
                    If B <> 32 Then
                        Engine_Text_Width = Engine_Text_Width + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1
                    Else
                        Engine_Text_Width = Engine_Text_Width + 4
                    End If
                Next A
            Else
                e = 0
                f = 0
                For A = 1 To Len(Texto)
                    B = Asc(mid(Texto, A, 1))
                    graf.GrhIndex = Fuentes(Fon).Caracteres(B)
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
                            d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1
                            If d > Engine_Text_Width Then Engine_Text_Width = d
                        End If
                    End If
                    e = e + 1
                Next A
            End If
    End Select
    Exit Function
Engine_Text_Width_Err:
    Call RegistrarError(Err.Number, Err.Description, "TextRenderer.Engine_Text_Width", Erl)
    Resume Next
End Function

Public Function Engine_Text_WidthCentrado(Texto As String, Optional multi As Boolean = False, Optional Fon As Byte = 1) As Integer
    On Error GoTo Engine_Text_WidthCentrado_Err
    Dim A As Integer, B As Integer, d As Integer, e As Integer, f As Integer
    Dim graf As Grh
    Select Case Fon
        Case 1
            '
            If multi = False Then
                For A = 1 To Len(Texto)
                    B = Asc(mid(Texto, A, 1))
                    graf.GrhIndex = Fuentes(1).Caracteres(B)
                    If graf.GrhIndex = 0 Then graf.GrhIndex = 1
                    If B <> 32 Then
                        Engine_Text_WidthCentrado = Engine_Text_WidthCentrado + GrhData(GrhData(graf.GrhIndex + 1).Frames(1)).pixelWidth '+ 1
                    Else
                        Engine_Text_WidthCentrado = Engine_Text_WidthCentrado + 4
                    End If
                Next A
            Else
                e = 0
                f = 0
                For A = 1 To Len(Texto)
                    B = Asc(mid(Texto, A, 1))
                    graf.GrhIndex = Fuentes(1).Caracteres(B)
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
                            d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1
                            If d > Engine_Text_WidthCentrado Then Engine_Text_WidthCentrado = d
                        End If
                    End If
                    e = e + 1
                Next A
            End If
        Case 4
            If multi = False Then
                For A = 1 To Len(Texto)
                    B = Asc(mid(Texto, A, 1))
                    graf.GrhIndex = Fuentes(Fon).Caracteres(B)
                    If graf.GrhIndex = 0 Then graf.GrhIndex = 1
                    If B <> 20 Then
                        Engine_Text_WidthCentrado = Engine_Text_WidthCentrado + GrhData(GrhData(graf.GrhIndex + 1).Frames(1)).pixelWidth + 10
                    Else
                        Engine_Text_WidthCentrado = Engine_Text_WidthCentrado - 15
                    End If
                Next A
            Else
                e = 0
                f = 0
                For A = 1 To Len(Texto)
                    B = Asc(mid(Texto, A, 1))
                    graf.GrhIndex = Fuentes(Fon).Caracteres(B)
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
                            d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1
                            If d > Engine_Text_WidthCentrado Then Engine_Text_WidthCentrado = d
                        End If
                    End If
                    e = e + 1
                Next A
            End If
        Case Else
            If multi = False Then
                For A = 1 To Len(Texto)
                    B = Asc(mid(Texto, A, 1))
                    graf.GrhIndex = Fuentes(Fon).Caracteres(B)
                    If graf.GrhIndex = 0 Then graf.GrhIndex = 1
                    If B <> 32 Then
                        Engine_Text_WidthCentrado = Engine_Text_WidthCentrado + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1
                    Else
                        Engine_Text_WidthCentrado = Engine_Text_WidthCentrado + 4
                    End If
                Next A
            Else
                e = 0
                f = 0
                For A = 1 To Len(Texto)
                    B = Asc(mid(Texto, A, 1))
                    graf.GrhIndex = Fuentes(Fon).Caracteres(B)
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
                            d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1
                            If d > Engine_Text_WidthCentrado Then Engine_Text_WidthCentrado = d
                        End If
                    End If
                    e = e + 1
                Next A
            End If
    End Select
    Exit Function
Engine_Text_WidthCentrado_Err:
    Call RegistrarError(Err.Number, Err.Description, "TextRenderer.Engine_Text_WidthCentrado", Erl)
    Resume Next
End Function

Public Sub Text_Render(ByVal font As D3DXFont, _
                       text As String, _
                       ByVal Top As Long, _
                       ByVal Left As Long, _
                       ByVal Width As Long, _
                       ByVal Height As Long, _
                       ByVal color As Long, _
                       ByVal format As Long, _
                       Optional ByVal Shadow As Boolean = False)
    On Error GoTo Text_Render_Err
    Dim TextRect   As Rect
    Dim ShadowRect As Rect
    TextRect.Top = Top
    TextRect.Left = Left
    TextRect.Bottom = Top + Height
    TextRect.Right = Left + Width
    If Shadow Then
        ShadowRect.Top = Top - 1
        ShadowRect.Left = Left - 2
        ShadowRect.Bottom = (Top + Height) - 1
        ShadowRect.Right = (Left + Width) - 2
        DirectD3D8.drawText font, &HFF000000, text, ShadowRect, format
    End If
    DirectD3D8.drawText font, color, text, TextRect, format
    Exit Sub
Text_Render_Err:
    Call RegistrarError(Err.Number, Err.Description, "TextRenderer.Text_Render", Erl)
    Resume Next
End Sub

Public Sub Text_Render_ext(text As String, _
                           ByVal Top As Long, _
                           ByVal Left As Long, _
                           ByVal Width As Long, _
                           ByVal Height As Long, _
                           ByVal color As Long, _
                           Optional ByVal Shadow As Boolean = False, _
                           Optional ByVal center As Boolean = False, _
                           Optional ByVal font As Long = 0)
    On Error GoTo Text_Render_ext_Err
    If center = True Then
        Call Text_Render(font_list(font), text, Top, Left, Width, Height, color, DT_VCENTER & DT_CENTER, Shadow)
    Else
        Call Text_Render(font_list(font), text, Top, Left, Width, Height, color, DT_TOP Or DT_LEFT, Shadow)
    End If
    Exit Sub
Text_Render_ext_Err:
    Call RegistrarError(Err.Number, Err.Description, "TextRenderer.Text_Render_ext", Erl)
    Resume Next
End Sub

Private Sub Font_Make(ByVal font_index As Long, ByVal Style As String, ByVal bold As Boolean, ByVal italic As Boolean, ByVal size As Long)
    On Error GoTo Font_Make_Err
    If font_index > font_last Then
        font_last = font_index
        ReDim Preserve font_list(1 To font_last)
    End If
    font_count = font_count + 1
    Dim font_desc As IFont
    Dim fnt       As New StdFont
    fnt.Name = Style
    fnt.size = size
    fnt.bold = bold
    fnt.italic = italic
    Set font_desc = fnt
    Set font_list(font_index) = DirectD3D8.CreateFont(DirectDevice, font_desc.hFont)
    Exit Sub
Font_Make_Err:
    Call RegistrarError(Err.Number, Err.Description, "TextRenderer.Font_Make", Erl)
    Resume Next
End Sub

Public Function Font_Create(ByVal Style As String, ByVal size As Long, ByVal bold As Boolean, ByVal italic As Boolean) As Long
    On Error GoTo ErrorHandler:
    Font_Create = Font_Next_Open
    Font_Make Font_Create, Style, bold, italic, size
ErrorHandler:
    Font_Create = 0
End Function

Public Function Font_Next_Open() As Long
    On Error GoTo Font_Next_Open_Err
    Font_Next_Open = font_last + 1
    Exit Function
Font_Next_Open_Err:
    Call RegistrarError(Err.Number, Err.Description, "TextRenderer.Font_Next_Open", Erl)
    Resume Next
End Function

Public Function Font_Check(ByVal font_index As Long) As Boolean
    On Error GoTo Font_Check_Err
    If font_index > 0 And font_index <= font_last Then
        Font_Check = True
    End If
    Exit Function
Font_Check_Err:
    Call RegistrarError(Err.Number, Err.Description, "TextRenderer.Font_Check", Erl)
    Resume Next
End Function

Public Function Prepare_Multiline_Text(text As String, ByVal MaxWidth As Integer, Optional ByVal FontIndex As Integer = 1) As String()
    On Error GoTo Handler
    Dim Lines() As String
    If LenB(text) = 0 Then
        ReDim Lines(0)
        Prepare_Multiline_Text = Lines
        Exit Function
    End If
    Dim LetterIndex As Long, CurLetter As Integer, LastBreak As Long, CanBreak As Long, CurWidth As Integer, CurLine As Integer, CanBreakWidth As Integer
    With Fuentes(FontIndex)
        LastBreak = 1
        For LetterIndex = 1 To Len(text)
            CurLetter = Asc(mid$(text, LetterIndex, 1))
            If CurLetter = vbKeyReturn Then
                ReDim Preserve Lines(CurLine)
                If LetterIndex - LastBreak > 0 Then
                    Lines(CurLine) = mid$(text, LastBreak, LetterIndex - LastBreak)
                End If
                LastBreak = LetterIndex + 2
                CanBreak = LastBreak
                CurLine = CurLine + 1
                CurWidth = 0
            Else
                If .Caracteres(CurLetter) <> 0 Then CurWidth = CurWidth + GrhData(.Caracteres(CurLetter)).pixelWidth
                If CurLetter = vbKeySpace Or CurLetter = vbKeyTab Then
                    CanBreak = LetterIndex
                    CanBreakWidth = CurWidth
                End If
                If CurWidth > MaxWidth And MaxWidth > 0 Then
                    ReDim Preserve Lines(CurLine)
                    If CanBreak - LastBreak > 0 Then
                        Lines(CurLine) = mid$(text, LastBreak, CanBreak - LastBreak)
                        CurWidth = CurWidth - CanBreakWidth
                        LastBreak = CanBreak + 1
                    Else
                        Lines(CurLine) = mid$(text, LastBreak, LetterIndex - LastBreak)
                        CurWidth = GrhData(.Caracteres(CurLetter)).pixelWidth
                        LastBreak = LetterIndex
                    End If
                    CanBreak = LastBreak
                    CurLine = CurLine + 1
                End If
            End If
        Next
        If LetterIndex - LastBreak > 0 Then
            ReDim Preserve Lines(CurLine)
            Lines(CurLine) = mid$(text, LastBreak, LetterIndex - LastBreak)
        End If
    End With
    Prepare_Multiline_Text = Lines
    Exit Function
Handler:
    Call RegistrarError(Err.Number, Err.Description, "TextRenderer.Prepare_Multiline_Text", Erl)
    ReDim Lines(0)
    Prepare_Multiline_Text = Lines
End Function

Public Function Text_Width(text As String, Optional ByVal FontIndex As Byte = 1) As Integer
    On Error GoTo Handler
    Dim LetterIndex As Long, CurLetter As Integer
    With Fuentes(FontIndex)
        For LetterIndex = 1 To Len(text)
            CurLetter = Asc(mid$(text, LetterIndex, 1))
            Text_Width = Text_Width + GrhData(.Caracteres(CurLetter)).pixelWidth
        Next
    End With
    Exit Function
Handler:
    Call RegistrarError(Err.Number, Err.Description, "TextRenderer.Text_Width", Erl)
End Function
