VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLucesRedondas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private Type tLight

    RGBcolor As RGBA
    active As Boolean
    map_x As Long
    map_y As Long
    range As Long

End Type
     
Private light_list() As tLight

Private NumLights    As Integer

Friend Function Create_Light_To_Map(ByVal map_x As Byte, ByVal map_y As Byte, Color As RGBA, Optional range As Byte = 3)
    
    On Error GoTo Create_Light_To_Map_Err
    
    NumLights = NumLights + 1
       
    ReDim Preserve light_list(1 To NumLights) As tLight
    
    'Le damos color
    light_list(NumLights).RGBcolor = Color
       
    'Alpha (Si borras esto RE KB!!)
    light_list(NumLights).RGBcolor.A = 255
       
    'Un rango determinado
    light_list(NumLights).range = range
       
    'La activamos para que se pueda renderizar
    light_list(NumLights).active = True
       
    'Pos en el mapa
    light_list(NumLights).map_x = map_x
    light_list(NumLights).map_y = map_y
    
    UpdateLights = True
    
    Exit Function

Create_Light_To_Map_Err:
    Call RegistrarError(Err.number, Err.Description, "clsLucesRedondas.Create_Light_To_Map", Erl)
    Resume Next
    
End Function

Public Function Delete_Light_To_Map(ByVal x As Byte, ByVal y As Byte)
    
    On Error GoTo Delete_Light_To_Map_Err
    
       
    Dim i As Long
       
    For i = 1 To NumLights

        If light_list(i).map_x = x And light_list(i).map_y = y Then
            Delete_Light_To_Index i
            Exit Function

        End If

    Next i
     
    
    Exit Function

Delete_Light_To_Map_Err:
    Call RegistrarError(Err.number, Err.Description, "clsLucesRedondas.Delete_Light_To_Map", Erl)
    Resume Next
    
End Function

Public Function Delete_Light_To_Index(ByVal light_index As Byte, Optional RenderFinal As Boolean = True)
    
    On Error GoTo Delete_Light_To_Index_Err
    
       
    Dim min_x As Integer

    Dim min_y As Integer

    Dim max_x As Integer

    Dim max_y As Integer

    Dim Ya    As Integer

    Dim Xa    As Integer
     
    light_list(light_index).active = False
    NumLights = NumLights - 1
     
    min_x = light_list(light_index).map_x - light_list(light_index).range
    max_x = light_list(light_index).map_x + light_list(light_index).range
    min_y = light_list(light_index).map_y - light_list(light_index).range
    max_y = light_list(light_index).map_y + light_list(light_index).range
    
    UpdateLights = True
    
    For Ya = min_y To max_y
        For Xa = min_x To max_x

            If InMapBounds(Xa, Ya) Then
                MapData(Xa, Ya).light_value(0) = COLOR_EMPTY
                MapData(Xa, Ya).light_value(1) = COLOR_EMPTY
                MapData(Xa, Ya).light_value(2) = COLOR_EMPTY
                MapData(Xa, Ya).light_value(3) = COLOR_EMPTY
            End If

        Next Xa
    Next Ya
    
    Exit Function

Delete_Light_To_Index_Err:
    Call RegistrarError(Err.number, Err.Description, "clsLucesRedondas.Delete_Light_To_Index", Erl)
    Resume Next
    
End Function

Private Function LightCalculate(ByVal Radio2 As Long, ByVal LightX As Integer, ByVal LightY As Integer, ByVal XCoord As Integer, ByVal YCoord As Integer, TileLight As RGBA, LightColor As RGBA) As RGBA
    
    On Error GoTo LightCalculate_Err
    
    Dim XDist        As Single
    Dim YDist        As Single
    Dim VertexDist2  As Single

    XDist = LightX - XCoord
    YDist = LightY - YCoord

    VertexDist2 = XDist * XDist + YDist * YDist

    If VertexDist2 <= Radio2 Then
        Call LerpRGBA(LightCalculate, LightColor, TileLight, VertexDist2 / Radio2)
    Else
        LightCalculate = TileLight
    End If

    
    Exit Function

LightCalculate_Err:
    Call RegistrarError(Err.number, Err.Description, "clsLucesRedondas.LightCalculate", Erl)
    Resume Next
    
End Function
     
Private Sub LightRender(ByVal light_index As Integer)
    
    On Error GoTo LightRender_Err
    
     
    If light_index = 0 Then Exit Sub
    If light_list(light_index).active = False Then Exit Sub
       
    Dim min_x        As Long

    Dim min_y        As Long

    Dim max_x        As Long

    Dim max_y        As Long

    Dim Color        As Long

    Dim Ya           As Long

    Dim Xa           As Long
       
    Dim TileLight    As RGBA

    Dim LightColor   As RGBA
       
    Dim XCoord       As Integer

    Dim YCoord       As Integer
    
    Dim LightX       As Integer
    
    Dim LightY       As Integer
    
    Dim radius       As Long
    
    Dim range        As Integer

    LightColor = light_list(light_index).RGBcolor
    
    LightX = light_list(light_index).map_x * 32 + 16
    LightY = light_list(light_index).map_y * 32 + 16

    range = light_list(light_index).range
    
    radius = range * 32 + 16
    radius = radius * radius
           
    min_x = light_list(light_index).map_x - range
    max_x = light_list(light_index).map_x + range
    min_y = light_list(light_index).map_y - range
    max_y = light_list(light_index).map_y + range

    
    
    For Ya = min_y To max_y
        For Xa = min_x To max_x

            If InMapBounds(Xa, Ya) Then
                XCoord = Xa * 32
                YCoord = Ya * 32 + 32
                MapData(Xa, Ya).light_value(0) = LightCalculate(radius, LightX, LightY, XCoord, YCoord, MapData(Xa, Ya).light_value(0), LightColor)
     
                XCoord = Xa * 32
                YCoord = Ya * 32
                MapData(Xa, Ya).light_value(1) = LightCalculate(radius, LightX, LightY, XCoord, YCoord, MapData(Xa, Ya).light_value(1), LightColor)
                           
                XCoord = Xa * 32 + 32
                YCoord = Ya * 32 + 32
                MapData(Xa, Ya).light_value(2) = LightCalculate(radius, LightX, LightY, XCoord, YCoord, MapData(Xa, Ya).light_value(2), LightColor)
       
                XCoord = Xa * 32 + 32
                YCoord = Ya * 32
                MapData(Xa, Ya).light_value(3) = LightCalculate(radius, LightX, LightY, XCoord, YCoord, MapData(Xa, Ya).light_value(3), LightColor)
                   
            End If

        Next Xa
    Next Ya

    
    Exit Sub

LightRender_Err:
    Call RegistrarError(Err.number, Err.Description, "clsLucesRedondas.LightRender", Erl)
    Resume Next
    
End Sub

Public Sub LightRenderAll(ByVal MinX As Long, ByVal MinY As Long, ByVal MaxX As Long, ByVal MaxY As Long)
    
    On Error GoTo LightRenderAll_Err
    

    Dim i As Long
       
    For i = 1 To NumLights
        If light_list(i).map_x + light_list(i).range > MinX And light_list(i).map_x - light_list(i).range < MaxX And light_list(i).map_y + light_list(i).range > MinY And light_list(i).map_y - light_list(i).range < MaxY Then
            LightRender i
        End If
    Next i

    
    Exit Sub

LightRenderAll_Err:
    Call RegistrarError(Err.number, Err.Description, "clsLucesRedondas.LightRenderAll", Erl)
    Resume Next
    
End Sub

Public Function Delete_All_LigthRound(Optional RenderFinal As Boolean = True)
    
    On Error GoTo Delete_All_LigthRound_Err
    

    Dim i As Long
       
    For i = 1 To NumLights
        Delete_Light_To_Index i, RenderFinal
    Next i
    
    Exit Function

Delete_All_LigthRound_Err:
    Call RegistrarError(Err.number, Err.Description, "clsLucesRedondas.Delete_All_LigthRound", Erl)
    Resume Next
    
End Function
