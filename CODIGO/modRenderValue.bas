Attribute VB_Name = "modRenderValue"
Option Explicit

' GS-Zone AO
' Basad en el Sistema de Daño aportado por maTih.- adaptado by ^[GS]^
' Fuente: http://www.gs-zone.org/dano_tds_style_en_mapa_tl6O.html
 
Const RENDER_TIME As Integer = 300

Enum RVType

    ePuñal = 1                'Apuñalo.
    eNormal = 2               'Golpe común.
    eMagic = 3                'Hechizo. ' GSZAO
    eGold = 4                 'Oro ' GSZAO
    eTrabajo = 5
    eExp = 6

End Enum
 
Private RVNormalFont As New StdFont
 
Type RVList

    RenderVal      As Double  'Cantidad.
    ColorRGB       As Long     'Color.
    RenderType     As RVType   'Tipo, se usa para saber si es apu o no.
    'RenderFont     As New StdFont  'Efecto del apu.
    TimeRendered   As Single  'Tiempo transcurrido.
    Downloading    As Single     'Contador para la posicion Y.
    Activated      As Boolean  'Si está activado..

End Type
 
Sub Create(ByVal x As Byte, ByVal y As Byte, ByVal ColorRGB As Long, ByVal rValue As Double, ByVal eMode As Byte)
     
    ' @ Agrega un nuevo valor.
     
    With MapData(x, y).RenderValue
         
        .Activated = True
        .ColorRGB = ColorRGB
        .RenderType = eMode
        .RenderVal = rValue
        .TimeRendered = 255
        .Downloading = 0
         
    End With
 
End Sub
 
Sub Draw(ByVal x As Byte, ByVal y As Byte, ByVal PixelX As Integer, ByVal PixelY As Integer, ByVal TicksPerFrame As Single)
 
    ' @ Dibuja un valor
    
    Dim Text As String, Width As Integer
     
    With MapData(x, y).RenderValue
         
        If (Not .Activated) Or (Not .RenderVal <> 0) Then Exit Sub
        If .TimeRendered < RENDER_TIME Then
            
            'Sumo el contador del tiempo.
            .TimeRendered = .TimeRendered - (9 * TicksPerFrame * Sgn(1))
                
            If (.TimeRendered / 2) > 0 Then
                .Downloading = (.TimeRendered / 6)

            End If
                
            .ColorRGB = ModifyColor(.TimeRendered, .RenderType)
            Call ColorToDX8(.ColorRGB)
                
            Select Case .RenderType

                Case eGold
                    Text = "+" & PonerPuntos(CLng(.RenderVal)) & " ORO"

                Case eExp
                    Text = "+" & PonerPuntos(CLng(.RenderVal)) & " EXP"

                Case eTrabajo
                    Text = "+" & PonerPuntos(CLng(.RenderVal))

                Case Else
                    Text = "-" & PonerPuntos(CLng(.RenderVal))

            End Select
                
            Width = Engine_Text_Width(Text)
                
            'Dibujo ; D
            Engine_Text_Render2 Text, (PixelX - Width \ 2), (PixelY - 48) + .Downloading, .ColorRGB, , , Int(.TimeRendered) ' .RenderFont,
               
            'Si llego al tiempo lo limpio
            If .TimeRendered <= 0 Then
                Call Clear(x, y)

            End If
                
        End If
           
    End With
 
End Sub
 
Private Sub Clear(ByVal x As Byte, ByVal y As Byte)
 
    ' @ Limpia todo.
     
    With MapData(x, y).RenderValue
        .Activated = False
        .ColorRGB = 0
        .RenderVal = 0
        .TimeRendered = 0

    End With
 
End Sub

Public Function ColorToDX8(ByVal Long_Color As Long) As Long

    ' DX8 engine
    Dim temp_color As String

    Dim red        As Integer, blue As Integer, green As Integer
    
    temp_color = Hex$(Long_Color)

    If Len(temp_color) < 6 Then
        'Give is 6 digits for easy RGB conversion.
        temp_color = String(6 - Len(temp_color), "0") + temp_color

    End If
    
    red = CLng("&H" + mid$(temp_color, 1, 2))
    green = CLng("&H" + mid$(temp_color, 3, 2))
    blue = CLng("&H" + mid$(temp_color, 5, 2))
    
    ColorToDX8 = D3DColorXRGB(red, green, blue)

End Function

Private Function ModifyColor(ByVal TimeNowRendered As Integer, ByVal RenderType As RVType) As Long
 
    ' @ Se usa para los "efectos" en el tiempo.
    
    ' 512 ---- 255
    ' 120 ---- x = 255 * 120 / 512
    
    Dim TimeX2 As Integer

    TimeX2 = TimeNowRendered ' * 2

    If TimeX2 < 0 Then TimeX2 = 0
    
    Select Case RenderType

        Case RVType.ePuñal
            ModifyColor = ARGB(0, 0, 0, TimeX2)

        Case RVType.eNormal
            ModifyColor = ARGB(255, 0, 0, TimeX2)

        Case RVType.eMagic
            ModifyColor = ARGB(0, 0, 0, TimeX2)

        Case RVType.eGold
            ModifyColor = ARGB(204, 193, 115, TimeX2)

            ' ModifyColor = ARGB(0, 0, 0, TimeX2)
        Case RVType.eExp
            ModifyColor = ARGB(0, 169, 255, TimeX2)

        Case RVType.eTrabajo
            ModifyColor = ARGB(255, 255, 255, TimeX2)

    End Select
 
End Function

