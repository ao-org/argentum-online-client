Attribute VB_Name = "ModPelota"
'MODULO DEDICADO A ROCHI! :D MI PELOTA PREFERIDA!
'11-01-2010
'RevolucionAo 1.0
'Pablo Mercavides


Option Explicit
Const CantidadDeFps = 15
Public DibujarPelota As Boolean
Public Type TPelota
    X As Integer
    Y As Integer
    DireccionX As Integer
    DireccionY As Integer
    fps As Integer
End Type

Public Type TCosa
    X As Integer
    Y As Integer
End Type
Public Cosa1 As TCosa
Public Cosa2 As TCosa
Public Pelota As TPelota

Public Sub CalcularTodo()
   ' Pelota.X = Cosa1.X
    'Pelota.Y = Cosa1.Y
    Pelota.DireccionX = Abs(Cosa1.X - Cosa2.X) / CantidadDeFps
    Pelota.DireccionY = Abs(Cosa1.Y - Cosa2.Y) / CantidadDeFps
    If Cosa1.X > Cosa2.X Then
        Pelota.DireccionX = -Pelota.DireccionX
    ElseIf Cosa1.X < Cosa2.X Then
        Pelota.DireccionX = Pelota.DireccionX
    End If
    If Cosa1.Y > Cosa2.Y Then
        Pelota.DireccionY = -Pelota.DireccionY
    ElseIf Cosa1.X < Cosa2.X Then
        Pelota.DireccionY = Pelota.DireccionY
    End If
    
    Pelota.fps = 0
End Sub

