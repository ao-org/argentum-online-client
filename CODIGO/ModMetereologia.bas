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
Public TimeIndex           As Integer

Public NightIndex          As Integer
Public MorningIndex        As Integer

Public MeteoParticle        As Integer

Public Sub IniciarMeteorologia()
    
    On Error GoTo IniciarMeteorologia_Err
    
    ReDim DayColors(0 To 24) As RGBA

    ' 00:00 - 02:00
    Call SetRGBA(DayColors(0), 70, 70, 70)
    NightIndex = 0
    ' 02:00 - 04:00
    Call SetRGBA(DayColors(1), 60, 60, 60)
    ' 04:00 - 06:00
    Call SetRGBA(DayColors(2), 80, 80, 80)
    ' 06:00 - 08:00
    'Call SetRGBA(DayColors(3), 20, 20, 20)
    Call SetRGBA(DayColors(3), 100, 100, 100)
    MorningIndex = 3
    ' 08:00 - 10:00
    Call SetRGBA(DayColors(4), 100, 100, 100)
    ' 10:00 - 12:00
    Call SetRGBA(DayColors(5), 130, 125, 125)
    ' 12:00 - 14:00
    Call SetRGBA(DayColors(6), 150, 150, 150)
    ' 14:00 - 16:00
    Call SetRGBA(DayColors(7), 170, 170, 170)
    ' 16:00 - 18:00
    Call SetRGBA(DayColors(8), 180, 170, 170)
    ' 18:00 - 20:00
    Call SetRGBA(DayColors(9), 190, 180, 190)
    ' 20:00 - 22:00
    Call SetRGBA(DayColors(10), 200, 210, 200)
    ' 22:00 - 00:00
    Call SetRGBA(DayColors(11), 220, 220, 220)
    
    Call SetRGBA(DayColors(12), 255, 255, 255)
    NightIndex = 0
    ' 02:00 - 04:00
    Call SetRGBA(DayColors(13), 255, 255, 255)
    ' 04:00 - 06:00
    Call SetRGBA(DayColors(14), 240, 240, 240)
    ' 06:00 - 08:00
    Call SetRGBA(DayColors(15), 240, 235, 240)
    MorningIndex = 3
    ' 08:00 - 10:00
    Call SetRGBA(DayColors(16), 230, 230, 230)
    ' 10:00 - 12:00
    Call SetRGBA(DayColors(17), 210, 200, 210)
    ' 12:00 - 14:00
    Call SetRGBA(DayColors(18), 220, 190, 200)
    ' 14:00 - 16:00
    Call SetRGBA(DayColors(19), 190, 170, 170)
    ' 16:00 - 18:00
    Call SetRGBA(DayColors(20), 130, 130, 170)
    ' 18:00 - 20:00
    Call SetRGBA(DayColors(21), 130, 130, 170)
    ' 20:00 - 22:00
    Call SetRGBA(DayColors(22), 120, 120, 140)
    ' 22:00 - 00:00
    Call SetRGBA(DayColors(23), 110, 110, 110)

    Call SetRGBA(DayColors(24), 90, 90, 90)
        
    ' Muerto
    Call SetRGBA(DeathColor, 120, 120, 120)
    
    ' Ciego
    Call SetRGBA(BlindColor, 4, 4, 4)
    
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
