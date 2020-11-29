Attribute VB_Name = "Graficos_Utilidades"
Option Explicit

'For composed texture
Public ComposedTexture As Direct3DTexture8
Public ComposedTextureSurface As Direct3DSurface8
Public ComposedZBufferSurface As Direct3DSurface8
Public pBackbuffer As Direct3DSurface8
Public pZbuffer As Direct3DSurface8
Public ComposedTextureWidth As Integer
Public ComposedTextureHeight As Integer
Public ComposedTextureCenterX As Integer
Public ProjectionComposedTexture As D3DMATRIX

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)

Function MakeVector(ByVal X As Single, ByVal Y As Single, ByVal z As Single) As D3DVECTOR
    '*****************************************************
    '****** Coded by Menduz (lord.yo.wo@gmail.com) *******
    '*****************************************************
    MakeVector.X = X
    MakeVector.Y = Y
    MakeVector.z = z

End Function

Private Function CreateVertex(X As Single, Y As Single, z As Single, Color As RGBA, tu As Single, tv As Single) As TYPE_VERTEX
    
    CreateVertex.X = X
    CreateVertex.Y = Y
    CreateVertex.z = z
    CreateVertex.Color = Color
    CreateVertex.TX = tu
    CreateVertex.TY = tv

End Function


Private Function Geometry_Create_Vertex(ByVal X As Single, ByVal Y As Single, ByVal z As Single, Color As RGBA, tu As Single, ByVal tv As Single) As TYPE_VERTEX
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '**************************************************************
    Geometry_Create_Vertex.X = X
    Geometry_Create_Vertex.Y = Y
    Geometry_Create_Vertex.z = z
    Geometry_Create_Vertex.Color = Color
    Geometry_Create_Vertex.TX = tu
    Geometry_Create_Vertex.TY = tv

End Function

Public Sub Geometry_Create_Box(ByRef verts() As TYPE_VERTEX, ByRef Dest As RECT, ByRef Src As RECT, ByRef rgb_list() As RGBA, Optional ByVal Textures_Width As Long, Optional ByVal Textures_Height As Long, Optional ByVal Angle As Single)

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
    
    If Angle > 0 Then
        'Center coordinates on screen of the square
        x_center = Dest.Left + (Dest.Right - Dest.Left) / 2
        y_center = Dest.Top + (Dest.bottom - Dest.Top) / 2
        
        'Calculate radius
        radius = Sqr((Dest.Right - x_center) ^ 2 + (Dest.bottom - y_center) ^ 2)
        
        'Calculate left and right points
        temp = (Dest.Right - x_center) / radius
        right_point = Atn(temp / Sqr(-temp * temp + 1))
        left_point = PI - right_point

    End If
    
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        x_Cor = Dest.Left
        y_Cor = Dest.bottom
    Else
        x_Cor = x_center + Cos(-left_point - Angle) * radius
        y_Cor = y_center - Sin(-left_point - Angle) * radius

    End If
    
    '0 - Bottom left vertex
    If Textures_Width And Textures_Height Then
        verts(0) = Geometry_Create_Vertex(x_Cor, y_Cor, 0, rgb_list(0), Src.Left / Textures_Width, Src.bottom / Textures_Height)
    Else
        verts(0) = Geometry_Create_Vertex(x_Cor, y_Cor, 0, rgb_list(0), 0, 0)
    End If

    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        x_Cor = Dest.Left
        y_Cor = Dest.Top
    Else
        x_Cor = x_center + Cos(left_point - Angle) * radius
        y_Cor = y_center - Sin(left_point - Angle) * radius

    End If
    
    '1 - Top left vertex
    If Textures_Width And Textures_Height Then
        verts(1) = Geometry_Create_Vertex(x_Cor, y_Cor, 0, rgb_list(1), Src.Left / Textures_Width, Src.Top / Textures_Height)
    Else
        verts(1) = Geometry_Create_Vertex(x_Cor, y_Cor, 0, rgb_list(1), 0, 1)
    End If

    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        x_Cor = Dest.Right
        y_Cor = Dest.bottom
    Else
        x_Cor = x_center + Cos(-right_point - Angle) * radius
        y_Cor = y_center - Sin(-right_point - Angle) * radius

    End If
    
    '2 - Bottom right vertex
    If Textures_Width And Textures_Height Then
        verts(2) = Geometry_Create_Vertex(x_Cor, y_Cor, 0, rgb_list(2), Src.Right / Textures_Width, Src.bottom / Textures_Height)
    Else
        verts(2) = Geometry_Create_Vertex(x_Cor, y_Cor, 0, rgb_list(2), 1, 0)
    End If

    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        x_Cor = Dest.Right
        y_Cor = Dest.Top
    Else
        x_Cor = x_center + Cos(right_point - Angle) * radius
        y_Cor = y_center - Sin(right_point - Angle) * radius

    End If
    
    '3 - Top right vertex
    If Textures_Width And Textures_Height Then
        verts(3) = Geometry_Create_Vertex(x_Cor, y_Cor, 0, rgb_list(3), Src.Right / Textures_Width, Src.Top / Textures_Height)
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

    ComposedTextureWidth = 128
    ComposedTextureHeight = 128

    ComposedTextureCenterX = ComposedTextureWidth \ 2

    Set ComposedTexture = DirectD3D8.CreateTexture(DirectDevice, ComposedTextureWidth, ComposedTextureHeight, 0, D3DUSAGE_RENDERTARGET, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT)
    Set ComposedTextureSurface = ComposedTexture.GetSurfaceLevel(0)
    
    Dim ComposedZBuffer As Direct3DTexture8
    Set ComposedZBuffer = DirectD3D8.CreateTexture(DirectDevice, ComposedTextureWidth, ComposedTextureHeight, 0, D3DUSAGE_DEPTHSTENCIL, D3DFMT_D24S8, D3DPOOL_DEFAULT)
    
    Set ComposedZBufferSurface = ComposedZBuffer.GetSurfaceLevel(0)

    Set pBackbuffer = DirectDevice.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO)
    Set pZbuffer = DirectDevice.GetDepthStencilSurface()

    Call D3DXMatrixOrthoOffCenterLH(ProjectionComposedTexture, 0, ComposedTextureWidth, ComposedTextureHeight, 0, -1#, 1#)

End Sub

Public Sub BeginComposedTexture()

    Call SpriteBatch.Flush
    Call DirectDevice.EndScene

    ' Render to texture
    DirectDevice.SetRenderTarget ComposedTextureSurface, ComposedZBufferSurface, 0

    ' Change viewport
    Call DirectDevice.SetTransform(D3DTS_PROJECTION, ProjectionComposedTexture)

    Call Engine_BeginScene

End Sub

Public Sub EndComposedTexture()

    Call SpriteBatch.Flush
    Call DirectDevice.EndScene

    ' Render to backbuffer
    DirectDevice.SetRenderTarget pBackbuffer, pZbuffer, 0
    
    ' Change viewport
    Call DirectDevice.SetTransform(D3DTS_PROJECTION, Projection)

    ' Render to BackBuffer
    Call DirectDevice.BeginScene
    Call SpriteBatch.Begin

End Sub

Public Sub PresentComposedTexture(ByVal X As Integer, ByVal Y As Integer, ByRef light_value() As RGBA, Optional ByVal Angle As Single = 0, Optional ByVal Shadow As Boolean = False, Optional ByVal Reflection As Boolean = False)

    Static src_rect            As RECT
    Static dest_rect           As RECT
    Static vertices(3)         As TYPE_VERTEX
    Static d3dTextures         As D3D8Textures
    
    X = X - ComposedTextureWidth \ 2 + 16
    Y = Y - ComposedTextureHeight + 32

    With SpriteBatch

        Call .SetTexture(ComposedTexture)

        Call .SetAlpha(False)
    
        If Shadow Then
            Call .DrawShadow(X, Y, ComposedTextureWidth, ComposedTextureHeight, light_value)
            
        ElseIf Reflection Then
            Call .DrawReflection(X, Y, ComposedTextureWidth, ComposedTextureHeight, light_value)
                    
        Else
            Call .Draw(X, Y, ComposedTextureWidth, ComposedTextureHeight, light_value, , , , , Angle)
        End If

    End With
 
End Sub

Public Function EaseBreathing(ByVal T As Single) As Single
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************

    If T < 0.25 Then
        Dim c1 As Single, c3 As Single
        c1 = 1.70158
        c3 = c1 + 1
        
        T = T * 4
        EaseBreathing = 1 + c3 * (T - 1) ^ 3 + c1 * (T - 1) ^ 2
    
    ElseIf T < 0.5 Then
        EaseBreathing = 2 - T * 4

    End If

End Function
