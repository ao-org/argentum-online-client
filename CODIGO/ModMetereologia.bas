Attribute VB_Name = "ModMetereologia"
Option Explicit

Public Const LIGHT_TRANSITION_DURATION = 5000

Public Const STEP_LIGHT_TRANSITION = 1 / LIGHT_TRANSITION_DURATION

'Status
Public Const Normal        As Byte = 0
Public Const NUBLADO       As Byte = 1
Public Const LLUVIA        As Byte = 2
Public Const NIEVE         As Byte = 3
Public Const TORMENTA      As Byte = 4

Public DayColors()         As RGBA
Public DeathColor          As RGBA
Public BlindColor          As RGBA
Public DungeonColor          As RGBA
Public TimeIndex           As Integer

Public NightIndex          As Integer
Public MorningIndex        As Integer

Public MeteoParticle        As Integer

Public Sub IniciarMeteorologia()
    
    On Error GoTo IniciarMeteorologia_Err
    
    ReDim DayColors(0 To 23) As RGBA
    ' 0hs
    Call SetRGBA(DayColors(23), 120, 120, 120)
    ' 1hs
    Call SetRGBA(DayColors(0), 120, 120, 120)
    ' 2hs
    Call SetRGBA(DayColors(1), 120, 120, 120)
    NightIndex = 0
    ' 3hs
    Call SetRGBA(DayColors(2), 120, 120, 120)
    ' 4hs
    Call SetRGBA(DayColors(3), 120, 120, 120)
    ' 5hs
    Call SetRGBA(DayColors(4), 138, 138, 138)
    MorningIndex = 3
    ' 6hs
    Call SetRGBA(DayColors(5), 156, 156, 145)
    ' 7hs
    Call SetRGBA(DayColors(6), 170, 170, 155)
    ' 8hs
    Call SetRGBA(DayColors(7), 185, 185, 185)
    ' 9hs
    Call SetRGBA(DayColors(8), 200, 200, 200)
    ' 10hs
    Call SetRGBA(DayColors(9), 220, 220, 220)
    ' 11hs
    Call SetRGBA(DayColors(10), 235, 235, 235)
    ' 12hs
    Call SetRGBA(DayColors(11), 245, 245, 245)
    ' 13hs
    Call SetRGBA(DayColors(12), 255, 255, 255)
    NightIndex = 0
    ' 14hs
    Call SetRGBA(DayColors(13), 255, 255, 255)
    ' 15hs
    Call SetRGBA(DayColors(14), 255, 255, 255)
    ' 16hs
    Call SetRGBA(DayColors(15), 245, 245, 245)
    MorningIndex = 3
    ' 17hs
    Call SetRGBA(DayColors(16), 230, 230, 230)
    ' 18hs
    Call SetRGBA(DayColors(17), 220, 220, 220)
    ' 19hs
    Call SetRGBA(DayColors(18), 200, 200, 180)
    ' 20hs
    Call SetRGBA(DayColors(19), 180, 160, 160)
    ' 21hs
    Call SetRGBA(DayColors(20), 160, 160, 160)
    ' 21hs
    Call SetRGBA(DayColors(21), 140, 140, 140)
    ' 23hs
    Call SetRGBA(DayColors(22), 120, 120, 140)
        
    ' Muerto
    Call SetRGBA(DeathColor, 120, 120, 120)
    
    ' Ciego
    Call SetRGBA(BlindColor, 4, 4, 4)
    
    ' Dungeon
    Call SetRGBA(DungeonColor, 130, 130, 130)
    
    TimeIndex = -1

    
    Exit Sub

IniciarMeteorologia_Err:
    Call RegistrarError(Err.number, Err.Description, "ModMetereologia.IniciarMeteorologia", Erl)
    Resume Next
    
End Sub

Public Sub RevisarHoraMundo(Optional ByVal Instantaneo As Boolean = False)
    
    On Error GoTo RevisarHoraMundo_Err

    Dim Elapsed As Single
    Elapsed = (FrameTime - HoraMundo) / DuracionDia
    Elapsed = (Elapsed - Fix(Elapsed)) * 24

    Dim HoraActual As Integer
    HoraActual = Fix(Elapsed)

    Dim CurrentIndex As Integer
    CurrentIndex = HoraActual \ 2
    If CurrentIndex <> TimeIndex Then
        TimeIndex = CurrentIndex
        If MapDat.base_light = 0 Then
            If Instantaneo Then
                global_light = DayColors(HoraActual)
            Else
                Call ActualizarLuz(DayColors(HoraActual))
            End If
            
            If TimeIndex = NightIndex Then
                Call Sound.Sound_Play(FXSound.Lobo_Sound, False, 0, 0)
    
            ElseIf TimeIndex = MorningIndex Then
                Call Sound.Sound_Play(FXSound.Gallo_Sound, False, 0, 0)
    
            End If
        End If
    End If
    
    Dim Minutos As Integer
    Dim Factor As Double
    
    Minutos = (Elapsed - HoraActual) * 60
    
    Factor = CDbl(Minutos) / CDbl(60)
    
    Dim HoraAnterior As Integer
    
    HoraAnterior = HoraActual - 1
    
    Call LerpRGB(global_light, DayColors((24 + HoraAnterior) Mod 24), DayColors(HoraActual), Factor)
    
    UpdateLights = True
    
    
    
    frmMain.lblhora = Right$("00" & HoraActual, 2) & ":" & Right$("00" & Minutos, 2)
    
    Exit Sub

RevisarHoraMundo_Err:
    Call RegistrarError(Err.number, Err.Description, "ModMetereologia.RevisarHoraMundo", Erl)
    Resume Next
    
End Sub

Public Sub ActualizarLuz(Color As RGBA)
    
    On Error GoTo ActualizarLuz_Err
    
    last_light = global_light
    next_light = Color
    light_transition = 0#
    
    Exit Sub

ActualizarLuz_Err:
    Call RegistrarError(Err.number, Err.Description, "ModMetereologia.ActualizarLuz", Erl)
    Resume Next
    
End Sub

Public Sub RestaurarLuz()
    
    On Error GoTo RestaurarLuz_Err
    
    If UserEstado = 1 Then
        global_light = DeathColor
        
    ElseIf UserCiego Then
        global_light = BlindColor
        
    ElseIf MapDat.zone = "DUNGEON" Then
        global_light = DungeonColor
        
    ElseIf TimeIndex >= 0 Then
       ' Dim Elapsed As Single
       '     Elapsed = (FrameTime - HoraMundo) / DuracionDia
       '     Elapsed = (Elapsed - Fix(Elapsed)) * 24
       '
       '     Dim HoraActual As Integer
       '     HoraActual = Fix(Elapsed)
       ' global_light = DayColors(HoraActual)
        
    Else
        global_light = COLOR_WHITE(0)
    End If
    light_transition = 1#
    
    Exit Sub

RestaurarLuz_Err:
    Call RegistrarError(Err.number, Err.Description, "ModMetereologia.RestaurarLuz", Erl)
    Resume Next
    
End Sub

Public Function EsNoche() As Boolean
    
    On Error GoTo EsNoche_Err
    
    EsNoche = (TimeIndex >= NightIndex And TimeIndex < MorningIndex)
    
    Exit Function

EsNoche_Err:
    Call RegistrarError(Err.number, Err.Description, "ModMetereologia.EsNoche", Erl)
    Resume Next
    
End Function
