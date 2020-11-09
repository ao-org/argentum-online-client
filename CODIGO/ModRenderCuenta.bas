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

Public Sub Geometry_Create_Box(ByRef verts() As TLVERTEX, ByRef dest As RECT, ByRef src As RECT, ByRef rgb_list() As Long, Optional ByRef Textures_Width As Long, Optional ByRef Textures_Height As Long)

    ' / Author: Dunkansdk

    ' * v0      * v1
    ' |        /|
    ' |      /  |
    ' |    /    |
    ' |  /      |
    ' |/        |
    ' * v2      * v3

    Dim x_Cor As Single

    Dim y_Cor As Single
    
    ' * - - - - - - - Vertice 0 -
    x_Cor = dest.Left
    y_Cor = dest.bottom
    
    '0 - Bottom left vertex
    If Textures_Width Or Textures_Height Then
        verts(0) = CreateVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.Left / Textures_Width + 0.001, (src.bottom) / Textures_Height)
    Else
        verts(0) = CreateVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 0)

    End If

    ' * - - - - - - - Vertice 0 -
    
    ' * - - - - - - - Vertice 1 -
    x_Cor = dest.Left
    y_Cor = dest.Top
       
    '1 - Top left vertex
    If Textures_Width Or Textures_Height Then
        verts(1) = CreateVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, src.Left / Textures_Width + 0.001, src.Top / Textures_Height + 0.001)
    Else
        verts(1) = CreateVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 0, 1)

    End If

    ' * - - - - - - - Vertice 1 -

    ' * - - - - - - - Vertice 2 -
    x_Cor = dest.Right
    y_Cor = dest.bottom
    
    '2 - Bottom right vertex
    If Textures_Width Or Textures_Height Then
        verts(2) = CreateVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, (src.Right) / Textures_Width, (src.bottom) / Textures_Height)
    Else
        verts(2) = CreateVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 1, 0)

    End If

    ' * - - - - - - - Vertice 2 -
    
    ' * - - - - - - - Vertice 3 -
    x_Cor = dest.Right
    y_Cor = dest.Top
    
    '3 - Top right vertex
    If Textures_Width Or Textures_Height Then
        verts(3) = CreateVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.Right) / Textures_Width, src.Top / Textures_Height + 0.001)
    Else
        verts(3) = CreateVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 1, 1)

    End If

    ' * - - - - - - - Vertice 3 -

End Sub

Public Function CreateVertex(ByVal x As Single, ByVal y As Single, ByVal Z As Single, ByVal rhw As Single, ByVal color As Long, ByVal Specular As Long, tu As Single, ByVal tv As Single) As TLVERTEX

    ' / Author: Aaron Perkins
    ' / Last Modify Date: 10/07/2002

    CreateVertex.x = x
    CreateVertex.y = y
    CreateVertex.Z = Z
    CreateVertex.rhw = rhw
    CreateVertex.color = color
    CreateVertex.Specular = Specular
    CreateVertex.tu = tu
    CreateVertex.tv = tv
    
End Function

