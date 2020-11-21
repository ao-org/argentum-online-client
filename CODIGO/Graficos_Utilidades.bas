Attribute VB_Name = "Graficos_Utilidades"
Option Explicit

'For composed texture
Public ComposedTexture As Direct3DTexture8
Public ComposedTextureSurface As Direct3DSurface8
Public pBackbuffer As Direct3DSurface8
Public ComposedTextureWidth As Integer
Public ComposedTextureHeight As Integer
Public ComposedTextureCenterX As Integer
Public ProjectionComposedTexture As D3DMATRIX

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)

Function MakeVector(ByVal x As Single, ByVal y As Single, ByVal Z As Single) As D3DVECTOR
    '*****************************************************
    '****** Coded by Menduz (lord.yo.wo@gmail.com) *******
    '*****************************************************
    MakeVector.x = x
    MakeVector.y = y
    MakeVector.Z = Z

End Function

Private Function CreateVertex(x As Single, y As Single, Z As Single, Color As Long, tu As Single, tv As Single) As TYPE_VERTEX
    
    CreateVertex.x = x
    CreateVertex.y = y
    CreateVertex.Z = Z
    CreateVertex.Color = Color
    CreateVertex.tx = tu
    CreateVertex.ty = tv

End Function


Private Function Geometry_Create_Vertex(ByVal x As Single, ByVal y As Single, ByVal Z As Single, ByVal Color As Long, tu As Single, ByVal tv As Single) As TYPE_VERTEX
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '**************************************************************
    Geometry_Create_Vertex.x = x
    Geometry_Create_Vertex.y = y
    Geometry_Create_Vertex.Z = Z
    Geometry_Create_Vertex.Color = Color
    Geometry_Create_Vertex.tx = tu
    Geometry_Create_Vertex.ty = tv

End Function

Public Sub Geometry_Create_Box(ByRef verts() As TYPE_VERTEX, ByRef dest As RECT, ByRef src As RECT, ByRef rgb_list() As Long, Optional ByVal Textures_Width As Long, Optional ByVal Textures_Height As Long, Optional ByVal angle As Single)

    '**************************************************************
    'Author: Aaron Perkins
    'Modified by Juan Martín Sotuyo Dodero
    'Last Modify Date: 11/17/2002
    '
    ' * v1      * v3
    ' |\        |
    ' |  \      |
    ' |    \    |
    ' |      \  |
    ' |        \|
    ' * v0      * v2
    '**************************************************************
    Dim x_center    As Single

    Dim y_center    As Single

    Dim radius      As Single

    Dim x_Cor       As Single

    Dim y_Cor       As Single

    Dim left_point  As Single

    Dim right_point As Single

    Dim temp        As Single
    
    If angle > 0 Then
        'Center coordinates on screen of the square
        x_center = dest.Left + (dest.Right - dest.Left) / 2
        y_center = dest.Top + (dest.bottom - dest.Top) / 2
        
        'Calculate radius
        radius = Sqr((dest.Right - x_center) ^ 2 + (dest.bottom - y_center) ^ 2)
        
        'Calculate left and right points
        temp = (dest.Right - x_center) / radius
        right_point = Atn(temp / Sqr(-temp * temp + 1))
        left_point = PI - right_point
    End If
    
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-left_point - angle) * radius
        y_Cor = y_center - Sin(-left_point - angle) * radius
    End If
    
    '0 - Bottom left vertex
    If Textures_Width And Textures_Height Then
        verts(0) = Geometry_Create_Vertex(x_Cor, y_Cor, 0, rgb_list(0), src.Left / Textures_Width, src.bottom / Textures_Height)
    Else
        verts(0) = Geometry_Create_Vertex(x_Cor, y_Cor, 0, rgb_list(0), 0, 0)
    End If

    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(left_point - angle) * radius
        y_Cor = y_center - Sin(left_point - angle) * radius
    End If
    
    '1 - Top left vertex
    If Textures_Width And Textures_Height Then
        verts(1) = Geometry_Create_Vertex(x_Cor, y_Cor, 0, rgb_list(1), src.Left / Textures_Width, src.Top / Textures_Height)
    Else
        verts(1) = Geometry_Create_Vertex(x_Cor, y_Cor, 0, rgb_list(1), 0, 1)
    End If

    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-right_point - angle) * radius
        y_Cor = y_center - Sin(-right_point - angle) * radius
    End If
    
    '2 - Bottom right vertex
    If Textures_Width And Textures_Height Then
        verts(2) = Geometry_Create_Vertex(x_Cor, y_Cor, 0, rgb_list(2), src.Right / Textures_Width, src.bottom / Textures_Height)
    Else
        verts(2) = Geometry_Create_Vertex(x_Cor, y_Cor, 0, rgb_list(2), 1, 0)
    End If

    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(right_point - angle) * radius
        y_Cor = y_center - Sin(right_point - angle) * radius
    End If
    
    '3 - Top right vertex
    If Textures_Width And Textures_Height Then
        verts(3) = Geometry_Create_Vertex(x_Cor, y_Cor, 0, rgb_list(3), src.Right / Textures_Width, src.Top / Textures_Height)
    Else
        verts(3) = Geometry_Create_Vertex(x_Cor, y_Cor, 0, rgb_list(3), 1, 1)
    End If

End Sub

Public Function BinarySearch(ByVal charindex As Integer) As Integer

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 07/28/07
    'Returns the index of the dialog in the list, or the negation
    'of the position were it should be if not found (for binary insertion)
    '**************************************************************
    Dim min As Long

    Dim max As Long

    Dim mid As Long
    
    min = 0
    max = dialogCount - 1
    
    Do While min <= max
        mid = (min + max) \ 2
        
        If dialogs(mid).charindex < charindex Then
            min = mid + 1
        ElseIf dialogs(mid).charindex > charindex Then
            max = mid - 1
        Else
            'We found it
            BinarySearch = mid
            Exit Function

        End If

    Loop
    
    'Not found, return the negation of the position where it should be
    '(all higher values are to the right of the list and lower values are to the left)
    BinarySearch = Not min

End Function


Public Sub InitComposedTexture()

    ComposedTextureWidth = 256
    ComposedTextureHeight = 256
    
    ComposedTextureCenterX = ComposedTextureWidth \ 2
    
    Set ComposedTexture = DirectD3D8.CreateTexture(DirectDevice, ComposedTextureWidth, ComposedTextureHeight, 0, D3DUSAGE_RENDERTARGET, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT)
    Set pBackbuffer = DirectDevice.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO)
    Set ComposedTextureSurface = ComposedTexture.GetSurfaceLevel(0)
    
    Call D3DXMatrixOrthoOffCenterLH(ProjectionComposedTexture, 0, ComposedTextureWidth, ComposedTextureHeight, 0, -1#, 1#)

End Sub

Public Sub BeginComposedTexture()

    Call SpriteBatch.Flush
    Call DirectDevice.EndScene

    ' Render to texture
    DirectDevice.SetRenderTarget ComposedTextureSurface, Nothing, 0
    
    ' Change viewport
    Call DirectDevice.SetTransform(D3DTS_PROJECTION, ProjectionComposedTexture)

    Call Engine_BeginScene
    'Call DirectDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, &HFF00FFFF, 1#, 0)
End Sub

Public Sub EndComposedTexture()

    Call SpriteBatch.Flush
    Call DirectDevice.EndScene

    ' Render to backbuffer
    DirectDevice.SetRenderTarget pBackbuffer, Nothing, ByVal 0
    
    ' Change viewport
    Call DirectDevice.SetTransform(D3DTS_PROJECTION, Projection)

    ' Render to BackBuffer
    Call DirectDevice.BeginScene
    Call SpriteBatch.Begin

End Sub

Public Sub PresentComposedTexture(ByVal x As Integer, ByVal y As Integer, ByRef Color_List() As Long, Optional ByVal angle As Single = 0, Optional ByVal Shadow As Boolean = False)

    Static src_rect            As RECT
    Static dest_rect           As RECT
    Static vertices(3)         As TYPE_VERTEX
    Static d3dTextures         As D3D8Textures
    Static light_value(0 To 3) As Long
    Static tmpColor            As D3DCOLORVALUE
    
    light_value(0) = Color_List(0)
    light_value(1) = Color_List(1)
    light_value(2) = Color_List(2)
    light_value(3) = Color_List(3)
 
    If (light_value(0) = 0) Then light_value(0) = map_base_light
    If (light_value(1) = 0) Then light_value(1) = map_base_light
    If (light_value(2) = 0) Then light_value(2) = map_base_light
    If (light_value(3) = 0) Then light_value(3) = map_base_light
        
    'Set up the source rectangle
    With src_rect
        .Right = ComposedTextureWidth
        .bottom = ComposedTextureHeight
    End With
                
    'Set up the destination rectangle
    With dest_rect
        .bottom = y + 32
        .Left = x + 16 - ComposedTextureWidth \ 2
        .Right = .Left + ComposedTextureHeight
        .Top = y + 32 - ComposedTextureHeight
    End With
    
    If Shadow Then
        Call ARGBtoD3DCOLORVALUE(light_value(0), tmpColor)

        Dim IntensidadSombra As Single
        IntensidadSombra = (0.2126 * tmpColor.r + 0.7152 * tmpColor.g + 0.0722 * tmpColor.b) ^ 2 / 65025
        
        Dim ColorShadow(3) As Long
        Call Engine_Long_To_RGB_List(ColorShadow(), D3DColorARGB(IntensidadSombra * 60, 0, 0, 0))
        
        Geometry_Create_Box vertices(), dest_rect, src_rect, ColorShadow(), ComposedTextureWidth, ComposedTextureHeight, angle
    Else
    
        'Set up the vertices(3) vertices
        Geometry_Create_Box vertices(), dest_rect, src_rect, light_value(), ComposedTextureWidth, ComposedTextureHeight, angle
    End If

    If Shadow Then
        vertices(1).x = vertices(1).x + (dest_rect.Right - dest_rect.Left) * 0.5
        vertices(1).y = vertices(1).y - (dest_rect.bottom - dest_rect.Top) * 0.5
       
        vertices(3).x = vertices(3).x + (dest_rect.Right - dest_rect.Left)
        vertices(3).y = vertices(3).y - (dest_rect.Right - dest_rect.Left) * 0.5
    End If

    With SpriteBatch

        Call .SetTexture(ComposedTexture)
        
        Call .SetAlpha(False)
        
        Call .DrawVertices(vertices)
            
    End With
 
End Sub

Public Function ARGBtoD3DCOLORVALUE(ByVal ARGB As Long, ByRef Color As D3DCOLORVALUE)
    Dim dest(3) As Byte
    CopyMemory dest(0), ARGB, 4
    Color.a = dest(3)
    Color.r = dest(2)
    Color.g = dest(1)
    Color.b = dest(0)
End Function
