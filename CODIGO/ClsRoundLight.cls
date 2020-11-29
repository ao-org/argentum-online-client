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
       
    LucesRedondas.LightRenderAll

End Function

Public Function Delete_Light_To_Map(ByVal x As Byte, ByVal y As Byte)
       
    Dim i As Long
       
    For i = 1 To NumLights

        If light_list(i).map_x = x And light_list(i).map_y = y Then
            Delete_Light_To_Index i
            Exit Function

        End If

    Next i
     
End Function

Public Function Delete_Light_To_Index(ByVal light_index As Byte, Optional RenderFinal As Boolean = True)
       
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

    If RenderFinal Then
        LucesRedondas.LightRenderAll

    End If
       
End Function

Private Function LightCalculate(ByVal Radio2 As Long, ByVal LightX As Integer, ByVal LightY As Integer, ByVal XCoord As Integer, ByVal YCoord As Integer, TileLight As RGBA, LightColor As RGBA, AmbientColor As RGBA) As RGBA
    Dim XDist        As Single
    Dim YDist        As Single
    Dim VertexDist2  As Single

    XDist = LightX - XCoord
    YDist = LightY - YCoord

    VertexDist2 = XDist * XDist + YDist * YDist

    If VertexDist2 <= Radio2 Then
        Call LerpRGBA(LightCalculate, LightColor, TileLight, Sqr(VertexDist2 / Radio2))
    Else
        LightCalculate = TileLight
    End If

End Function
     
Private Sub LightRender(ByVal light_index As Integer)
     
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

    Dim AmbientColor As RGBA

    Dim LightColor   As RGBA
       
    Dim XCoord       As Integer

    Dim YCoord       As Integer
    
    Dim LightX       As Integer
    
    Dim LightY       As Integer
    
    Dim Radius       As Long

    AmbientColor = map_light

    LightColor = light_list(light_index).RGBcolor
    
    LightX = light_list(light_index).map_x * 32 + 16
    LightY = light_list(light_index).map_y * 32 + 16
    
    Radius = light_list(light_index).range * 32
    Radius = Radius * Radius
           
    min_x = light_list(light_index).map_x - light_list(light_index).range
    max_x = light_list(light_index).map_x + light_list(light_index).range
    min_y = light_list(light_index).map_y - light_list(light_index).range
    max_y = light_list(light_index).map_y + light_list(light_index).range

    For Ya = min_y To max_y
        For Xa = min_x To max_x

            If InMapBounds(Xa, Ya) Then
                XCoord = Xa * 32
                YCoord = Ya * 32 + 32
                MapData(Xa, Ya).light_value(0) = LightCalculate(Radius, LightX, LightY, XCoord, YCoord, MapData(Xa, Ya).light_value(0), LightColor, AmbientColor)
     
                XCoord = Xa * 32
                YCoord = Ya * 32
                MapData(Xa, Ya).light_value(1) = LightCalculate(Radius, LightX, LightY, XCoord, YCoord, MapData(Xa, Ya).light_value(1), LightColor, AmbientColor)
                           
                XCoord = Xa * 32 + 32
                YCoord = Ya * 32 + 32
                MapData(Xa, Ya).light_value(2) = LightCalculate(Radius, LightX, LightY, XCoord, YCoord, MapData(Xa, Ya).light_value(2), LightColor, AmbientColor)
       
                XCoord = Xa * 32 + 32
                YCoord = Ya * 32
                MapData(Xa, Ya).light_value(3) = LightCalculate(Radius, LightX, LightY, XCoord, YCoord, MapData(Xa, Ya).light_value(3), LightColor, AmbientColor)
                   
            End If

        Next Xa
    Next Ya

End Sub

Public Sub LightRenderAll()

    Dim i As Long
       
    For i = 1 To NumLights
        LightRender i
                
    Next i

End Sub

Public Function Delete_All_LigthRound(Optional RenderFinal As Boolean = True)

    Dim i As Long
       
    For i = 1 To NumLights
        Delete_Light_To_Index i, RenderFinal
                
    Next i

End Function
