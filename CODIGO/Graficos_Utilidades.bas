Attribute VB_Name = "Graficos_Utilidades"
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

Function MakeVector(ByVal x As Single, ByVal y As Single, ByVal z As Single) As D3DVECTOR
    '*****************************************************
    '****** Coded by Menduz (lord.yo.wo@gmail.com) *******
    '*****************************************************
    
    On Error GoTo MakeVector_Err
    
    MakeVector.x = x
    MakeVector.y = y
    MakeVector.z = z

    
    Exit Function

MakeVector_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Utilidades.MakeVector", Erl)
    Resume Next
    
End Function

Private Function CreateVertex(x As Single, y As Single, z As Single, Color As RGBA, tu As Single, tv As Single) As TYPE_VERTEX
    
    On Error GoTo CreateVertex_Err
    
    
    CreateVertex.x = x
    CreateVertex.y = y
    CreateVertex.z = z
    CreateVertex.Color = Color
    CreateVertex.tX = tu
    CreateVertex.tY = tv

    
    Exit Function

CreateVertex_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Utilidades.CreateVertex", Erl)
    Resume Next
    
End Function


Private Function Geometry_Create_Vertex(ByVal x As Single, ByVal y As Single, ByVal z As Single, Color As RGBA, tu As Single, ByVal tv As Single) As TYPE_VERTEX
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '**************************************************************
    
    On Error GoTo Geometry_Create_Vertex_Err
    
    Geometry_Create_Vertex.x = x
    Geometry_Create_Vertex.y = y
    Geometry_Create_Vertex.z = z
    Geometry_Create_Vertex.Color = Color
    Geometry_Create_Vertex.tX = tu
    Geometry_Create_Vertex.tY = tv

    
    Exit Function

Geometry_Create_Vertex_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Utilidades.Geometry_Create_Vertex", Erl)
    Resume Next
    
End Function

Public Sub Geometry_Create_Box(ByRef verts() As TYPE_VERTEX, ByRef Dest As RECT, ByRef Src As RECT, ByRef rgb_list() As RGBA, Optional ByVal Textures_Width As Long, Optional ByVal Textures_Height As Long, Optional ByVal Angle As Single)
    
    On Error GoTo Geometry_Create_Box_Err
    

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
        y_center = Dest.Top + (Dest.Bottom - Dest.Top) / 2
        
        'Calculate radius
        radius = Sqr((Dest.Right - x_center) ^ 2 + (Dest.Bottom - y_center) ^ 2)
        
        'Calculate left and right points
        temp = (Dest.Right - x_center) / radius
        right_point = Atn(temp / Sqr(-temp * temp + 1))
        left_point = PI - right_point

    End If
    
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If Angle = 0 Then
        x_Cor = Dest.Left
        y_Cor = Dest.Bottom
    Else
        x_Cor = x_center + Cos(-left_point - Angle) * radius
        y_Cor = y_center - Sin(-left_point - Angle) * radius

    End If
    
    '0 - Bottom left vertex
    If Textures_Width And Textures_Height Then
        verts(0) = Geometry_Create_Vertex(x_Cor, y_Cor, 0, rgb_list(0), Src.Left / Textures_Width, Src.Bottom / Textures_Height)
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
        y_Cor = Dest.Bottom
    Else
        x_Cor = x_center + Cos(-right_point - Angle) * radius
        y_Cor = y_center - Sin(-right_point - Angle) * radius

    End If
    
    '2 - Bottom right vertex
    If Textures_Width And Textures_Height Then
        verts(2) = Geometry_Create_Vertex(x_Cor, y_Cor, 0, rgb_list(2), Src.Right / Textures_Width, Src.Bottom / Textures_Height)
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

    
    Exit Sub

Geometry_Create_Box_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Utilidades.Geometry_Create_Box", Erl)
    Resume Next
    
End Sub

Public Function BinarySearch(ByVal charindex As Integer) As Integer
    
    On Error GoTo BinarySearch_Err
    

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

    
    Exit Function

BinarySearch_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Utilidades.BinarySearch", Erl)
    Resume Next
    
End Function

Public Sub InitComposedTexture()
    
    On Error GoTo InitComposedTexture_Err
    

100 ComposedTextureWidth = 256
102 ComposedTextureHeight = 256

104 ComposedTextureCenterX = ComposedTextureWidth \ 2
106 Set ComposedTexture = DirectD3D8.CreateTexture(DirectDevice, ComposedTextureWidth, ComposedTextureHeight, 0, D3DUSAGE_RENDERTARGET, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT)
108 Set ComposedTextureSurface = ComposedTexture.GetSurfaceLevel(0)
110 Dim ComposedZBuffer As Direct3DTexture8
112 Set ComposedZBuffer = DirectD3D8.CreateTexture(DirectDevice, ComposedTextureWidth, ComposedTextureHeight, 0, D3DUSAGE_DEPTHSTENCIL, D3DFMT_D24S8, D3DPOOL_DEFAULT)
114 Set ComposedZBufferSurface = ComposedZBuffer.GetSurfaceLevel(0)
116 Set pBackbuffer = DirectDevice.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO)
118 Set pZbuffer = DirectDevice.GetDepthStencilSurface()
120 Call D3DXMatrixOrthoOffCenterLH(ProjectionComposedTexture, 0, ComposedTextureWidth, ComposedTextureHeight, 0, -1#, 1#)

    Exit Sub

InitComposedTexture_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Utilidades.InitComposedTexture", Erl)
    Resume Next
    
End Sub

Public Sub BeginComposedTexture()
    
    On Error GoTo BeginComposedTexture_Err
    

    Call SpriteBatch.Flush
    Call DirectDevice.EndScene

    ' Render to texture
    DirectDevice.SetRenderTarget ComposedTextureSurface, ComposedZBufferSurface, 0

    ' Change viewport
    Call DirectDevice.SetTransform(D3DTS_PROJECTION, ProjectionComposedTexture)

    Call Engine_BeginScene

    
    Exit Sub

BeginComposedTexture_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Utilidades.BeginComposedTexture", Erl)
    Resume Next
    
End Sub

Public Sub EndComposedTexture()
    
    On Error GoTo EndComposedTexture_Err
    

    Call SpriteBatch.Flush
    Call DirectDevice.EndScene

    ' Render to backbuffer
    DirectDevice.SetRenderTarget pBackbuffer, pZbuffer, 0
    
    ' Change viewport
    Call DirectDevice.SetTransform(D3DTS_PROJECTION, Projection)

    ' Render to BackBuffer
    Call DirectDevice.BeginScene
    Call SpriteBatch.Begin

    
    Exit Sub

EndComposedTexture_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Utilidades.EndComposedTexture", Erl)
    Resume Next
    
End Sub

Public Sub PresentComposedTexture(ByVal x As Integer, ByVal y As Integer, ByRef light_value() As RGBA, Optional ByVal Angle As Single = 0, Optional ByVal Shadow As Boolean = False, Optional ByVal Reflection As Boolean = False)
    
    On Error GoTo PresentComposedTexture_Err
    

    Static src_rect            As RECT
    Static dest_rect           As RECT
    Static vertices(3)         As TYPE_VERTEX
    Static d3dTextures         As D3D8Textures
    
    x = x - ComposedTextureWidth \ 2 + 16
    y = y - ComposedTextureHeight + 32

    With SpriteBatch

        Call .SetTexture(ComposedTexture)

        Call .SetAlpha(False)
    
        If Shadow Then
            Call .DrawShadow(x, y, ComposedTextureWidth, ComposedTextureHeight, light_value)
            
        ElseIf Reflection Then
            Call .DrawReflection(x, y, ComposedTextureWidth, ComposedTextureHeight, light_value)
                    
        Else
            Call .Draw(x, y, ComposedTextureWidth, ComposedTextureHeight, light_value, , , , , Angle)
        End If

    End With
 
    
    Exit Sub

PresentComposedTexture_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Utilidades.PresentComposedTexture", Erl)
    Resume Next
    
End Sub

Public Function EaseBreathing(ByVal t As Single) As Single
    '***************************************************
    'Author: Alexis Caraballo (WyroX)
    '***************************************************
    
    On Error GoTo EaseBreathing_Err
    

    If t < 0.25 Then
        Dim c1 As Single, c3 As Single
        c1 = 1.70158
        c3 = c1 + 1
        
        t = t * 4
        EaseBreathing = 1 + c3 * (t - 1) ^ 3 + c1 * (t - 1) ^ 2
    
    ElseIf t < 0.5 Then
        EaseBreathing = 2 - t * 4

    End If

    
    Exit Function

EaseBreathing_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Utilidades.EaseBreathing", Erl)
    Resume Next
    
End Function
