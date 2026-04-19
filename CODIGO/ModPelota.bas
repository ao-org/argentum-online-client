Attribute VB_Name = "ModPelota"
'MODULO DEDICADO A ROCHI! :D MI PELOTA PREFERIDA!
'11-01-2010
'RevolucionAo 1.0
'Pablo Mercavides


Option Explicit
Const CantidadDeFps = 15
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


