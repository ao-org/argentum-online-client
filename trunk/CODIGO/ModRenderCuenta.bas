Attribute VB_Name = "ModRenderCuenta"
'RevolucionAo 1.0
'Pablo Mercavides
Option Explicit

Public Sub Engine_Convert_List(rgb_list() As Long, Long_Color As Long)

    ' / Author: Dunkansdk
    ' / Note: Convierte en array's los D3DColorArgb

    rgb_list(0) = Long_Color
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)
    
End Sub

Public Sub Engine_Draw_Box(ByVal x As Integer, ByVal y As Integer, ByVal Width As Integer, ByVal Height As Integer, color As Long)

    ' / Author: Ezequiel Juárez (Standelf)
    ' / Note: Extract to Blisse AO, modified by Dunkansdk

    Dim b_Rect           As RECT

    Dim b_Color(0 To 3)  As Long

    Dim b_Vertex(0 To 3) As TLVERTEX
    
    With b_Rect
        .bottom = y + Height
        .Left = x
        .Right = x + Width
        .Top = y

    End With

    Engine_Convert_List b_Color(), color

    Geometry_Create_Box b_Vertex(), b_Rect, b_Rect, b_Color(), 0, 0
    
    D3DDevice.SetTexture 0, Nothing
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, b_Vertex(0), Len(b_Vertex(0))

End Sub

