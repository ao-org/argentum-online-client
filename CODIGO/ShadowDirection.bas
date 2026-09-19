Attribute VB_Name = "ShadowDirection"
Option Explicit

Public Function Shadow_DirectionForTile(ByVal map_x As Long, ByVal map_y As Long) As Single
    Shadow_DirectionForTile = LucesCuadradas.Light_ShadowDirectionX(map_x, map_y)
    If Shadow_DirectionForTile = 32767! Then Shadow_DirectionForTile = LucesRedondas.Light_ShadowDirectionX(map_x, map_y)
    If Shadow_DirectionForTile = 32767! Then Shadow_DirectionForTile = Shadow_SunDirectionX(Shadow_CurrentHour())
End Function

Public Function Shadow_DepthForTile(ByVal map_x As Long, ByVal map_y As Long) As Single
    Shadow_DepthForTile = LucesCuadradas.Light_ShadowDepth(map_x, map_y)
    If Shadow_DepthForTile = 32767! Then Shadow_DepthForTile = LucesRedondas.Light_ShadowDepth(map_x, map_y)
    If Shadow_DepthForTile = 32767! Then Shadow_DepthForTile = 16
End Function