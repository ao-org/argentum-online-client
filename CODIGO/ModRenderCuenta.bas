Attribute VB_Name = "ModRenderCuenta"
'RevolucionAo 1.0
'Pablo Mercavides
Option Explicit

Public Sub Engine_Convert_List(rgb_list() As Long, long_color As Long)

    ' / Author: Dunkansdk
    ' / Note: Convierte en array's los D3DColorArgb

    rgb_list(0) = long_color
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)
    
End Sub

Public Sub Engine_Draw_Box(ByVal x As Integer, ByVal y As Integer, ByVal Width As Integer, ByVal Height As Integer, Color As Long)

    Call Engine_Long_To_RGB_List(temp_rgb(), Color)

    Call SpriteBatch.SetTexture(Nothing)
    Call SpriteBatch.Draw(x, y, Width, Height, temp_rgb())

End Sub

Public Sub Engine_Draw_Box_Border(ByVal x As Integer, ByVal y As Integer, ByVal Width As Integer, ByVal Height As Integer, Color As Long, ColorLine As Long)

    Call Engine_Draw_Box(x, y, Width, Height, Color)

    Call Engine_Long_To_RGB_List(temp_rgb(), ColorLine)

    Call Engine_Draw_Box(x, y, Width, 1, ColorLine)
    Call Engine_Draw_Box(x, y + Height, Width, 1, ColorLine)
    Call Engine_Draw_Box(x, y, 1, Height, ColorLine)
    Call Engine_Draw_Box(x + Width, y, 1, Height, ColorLine)

End Sub

Public Sub Engine_Long_To_RGB_List(rgb_list() As Long, long_color As Long)
    '***************************************************
    'Author: Ezequiel Juarez (Standelf)
    'Last Modification: 16/05/10
    'Blisse-AO | Set a Long Color to a RGB List
    '***************************************************
    rgb_list(0) = long_color
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)

End Sub
