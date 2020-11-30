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
    ColorRGB       As RGBA     'Color.
    RenderType     As RVType   'Tipo, se usa para saber si es apu o no.
    'RenderFont     As New StdFont  'Efecto del apu.
    TimeRendered   As Single  'Tiempo transcurrido.
    Downloading    As Single     'Contador para la posicion Y.
    Activated      As Boolean  'Si está activado..

End Type
 
Sub Create(ByVal x As Byte, ByVal y As Byte, ColorRGB As RGBA, ByVal rValue As Double, ByVal eMode As Byte)
    
    On Error GoTo Create_Err
    
     
    ' @ Agrega un nuevo valor.
     
    With MapData(x, y).RenderValue
         
        .Activated = True
        .ColorRGB = ColorRGB
        .RenderType = eMode
        .RenderVal = rValue
        .TimeRendered = 255
        .Downloading = 0
         
    End With
 
    
    Exit Sub

Create_Err:
    Call RegistrarError(Err.number, Err.Description, "modRenderValue.Create", Erl)
    Resume Next
    
End Sub
 
Sub Draw(ByVal x As Byte, ByVal y As Byte, ByVal PixelX As Integer, ByVal PixelY As Integer, ByVal TicksPerFrame As Single)
    
    On Error GoTo Draw_Err
    
 
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
                
            Call ModifyColor(.ColorRGB, .TimeRendered, .RenderType)
                
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
 
    
    Exit Sub

Draw_Err:
    Call RegistrarError(Err.number, Err.Description, "modRenderValue.Draw", Erl)
    Resume Next
    
End Sub
 
Private Sub Clear(ByVal x As Byte, ByVal y As Byte)
    
    On Error GoTo Clear_Err
    
 
    ' @ Limpia todo.
     
    With MapData(x, y).RenderValue
        .Activated = False
        .ColorRGB = COLOR_EMPTY
        .RenderVal = 0
        .TimeRendered = 0

    End With
 
    
    Exit Sub

Clear_Err:
    Call RegistrarError(Err.number, Err.Description, "modRenderValue.Clear", Erl)
    Resume Next
    
End Sub

Private Sub ModifyColor(Color As RGBA, ByVal TimeNowRendered As Integer, ByVal RenderType As RVType)
    
    On Error GoTo ModifyColor_Err
    
 
    ' @ Se usa para los "efectos" en el tiempo.
    
    ' 512 ---- 255
    ' 120 ---- x = 255 * 120 / 512
    
    Dim TimeX2 As Integer

    TimeX2 = TimeNowRendered ' * 2

    If TimeX2 < 0 Then TimeX2 = 0
    
    Select Case RenderType

        Case RVType.ePuñal
            Call SetRGBA(Color, 0, 0, 0, TimeX2)

        Case RVType.eNormal
            Call SetRGBA(Color, 255, 0, 0, TimeX2)

        Case RVType.eMagic
            Call SetRGBA(Color, 0, 0, 0, TimeX2)

        Case RVType.eGold
            Call SetRGBA(Color, 204, 193, 115, TimeX2)

        Case RVType.eExp
            Call SetRGBA(Color, 0, 169, 255, TimeX2)

        Case RVType.eTrabajo
            Call SetRGBA(Color, 255, 255, 255, TimeX2)

    End Select
 
    
    Exit Sub

ModifyColor_Err:
    Call RegistrarError(Err.number, Err.Description, "modRenderValue.ModifyColor", Erl)
    Resume Next
    
End Sub

