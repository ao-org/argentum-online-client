Attribute VB_Name = "ModMetereologia"
Option Explicit

Public Const LIGHT_TRANSITION_DURATION = 5000

Public Const STEP_LIGHT_TRANSITION = 1 / LIGHT_TRANSITION_DURATION

'Status
Private Const Normal        As Byte = 0
Private Const NUBLADO       As Byte = 1
Private Const LLUVIA        As Byte = 2
Private Const NIEVE         As Byte = 3
Private Const TORMENTA      As Byte = 4

Private DayColors()         As RGBA
Private DeathColor          As RGBA
Private BlindColor          As RGBA
Private TimeIndex           As Integer

Private NightIndex          As Integer
Private MorningIndex        As Integer

Public MeteoParticle        As Integer

Public Sub IniciarMeteorologia()
    
    On Error GoTo IniciarMeteorologia_Err
    
    ReDim DayColors(11)

    ' 00:00 - 02:00
    Call SetRGBA(DayColors(0), 130, 130, 130)
    NightIndex = 0
    ' 02:00 - 04:00
    Call SetRGBA(DayColors(1), 130, 130, 160)
    ' 04:00 - 06:00
    Call SetRGBA(DayColors(2), 150, 150, 180)
    ' 06:00 - 08:00
    Call SetRGBA(DayColors(3), 200, 200, 200)
    MorningIndex = 3
    ' 08:00 - 10:00
    Call SetRGBA(DayColors(4), 230, 200, 200)
    ' 10:00 - 12:00
    Call SetRGBA(DayColors(5), 255, 230, 220)
    ' 12:00 - 14:00
    Call SetRGBA(DayColors(6), 255, 240, 220)
    ' 14:00 - 16:00
    Call SetRGBA(DayColors(7), 255, 250, 200)
    ' 16:00 - 18:00
    Call SetRGBA(DayColors(8), 255, 200, 200)
    ' 18:00 - 20:00
    Call SetRGBA(DayColors(9), 255, 180, 160)
    ' 20:00 - 22:00
    Call SetRGBA(DayColors(10), 180, 150, 140)
    ' 22:00 - 00:00
    Call SetRGBA(DayColors(11), 150, 140, 130)

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
                global_light = DayColors(TimeIndex)
            Else
                Call ActualizarLuz(DayColors(TimeIndex))
            End If
            
            If TimeIndex = NightIndex Then
                Call Sound.Sound_Play(FXSound.Lobo_Sound, False, 0, 0)
    
            ElseIf TimeIndex = MorningIndex Then
                Call Sound.Sound_Play(FXSound.Gallo_Sound, False, 0, 0)
    
            End If
        End If
    End If
    
    Dim Minutos As Integer
    Minutos = (Elapsed - HoraActual) * 60
    
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
        global_light = DayColors(TimeIndex)
        
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
