Attribute VB_Name = "ModAreas"
'    Argentum 20 - Game Client Program
'    Copyright (C) 2022 - Noland Studios
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'
Option Explicit
'LAS GUARDAMOS PARA PROCESAR LOS MPs y sabes si borrar personajes
Public MinLimiteX      As Integer
Public MaxLimiteX      As Integer
Public MinLimiteY      As Integer
Public MaxLimiteY      As Integer
Private Const AREA_DIM As Byte = 12
#If PYMMO = 1 Then
    ' Keep the client-side culling window aligned with the server area
    ' window so AreaChanged does not delete entities the server still sends.
    Private Const AREA_RADIUS_X As Byte = 4
    Private Const AREA_RADIUS_Y As Byte = 3
#Else
    Private Const AREA_RADIUS_X As Byte = 1
    Private Const AREA_RADIUS_Y As Byte = 1
#End If
Private Const AREA_TILE_WIDTH As Integer = AREA_DIM * (AREA_RADIUS_X * 2 + 1)
Private Const AREA_TILE_HEIGHT As Integer = AREA_DIM * (AREA_RADIUS_Y * 2 + 1)

Public Sub CambioDeArea(ByVal x As Byte, ByVal y As Byte)
    On Error GoTo CambioDeArea_Err
    Dim loopX As Long, loopY As Long
    MinLimiteX = (x \ AREA_DIM - AREA_RADIUS_X) * AREA_DIM
    MaxLimiteX = MinLimiteX + AREA_TILE_WIDTH - 1
    MinLimiteY = (y \ AREA_DIM - AREA_RADIUS_Y) * AREA_DIM
    MaxLimiteY = MinLimiteY + AREA_TILE_HEIGHT - 1
    For loopX = 1 To 100
        For loopY = 1 To 100
            If (loopY < MinLimiteY) Or (loopY > MaxLimiteY) Or (loopX < MinLimiteX) Or (loopX > MaxLimiteX) Then
                'Erase NPCs
                With MapData(loopX, loopY)
                    If .charindex > 0 Then
                        If .charindex <> UserCharIndex Then
                            Call EraseChar(.charindex)
                        End If
                    End If
                    'Erase OBJs
                    If Not EsObjetoFijo(loopX, loopY) Then
                        .ObjGrh.GrhIndex = 0
                        .OBJInfo.ObjIndex = 0
                    End If
                End With
            End If
        Next loopY
    Next loopX
    Call RefreshAllChars
    Exit Sub
CambioDeArea_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModAreas.CambioDeArea", Erl)
    Resume Next
End Sub
