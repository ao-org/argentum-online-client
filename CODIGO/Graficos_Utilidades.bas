Attribute VB_Name = "Graficos_Utilidades"
Option Explicit

Function MakeVector(ByVal x As Single, ByVal y As Single, ByVal Z As Single) As D3DVECTOR
    '*****************************************************
    '****** Coded by Menduz (lord.yo.wo@gmail.com) *******
    '*****************************************************
    MakeVector.x = x
    MakeVector.y = y
    MakeVector.Z = Z

End Function

Private Function CreateTLVertex(x As Single, y As Single, Z As Single, rhw As Single, Color As Long, Specular As Long, tu As Single, tv As Single) As TLVERTEX
    '*****************************************************
    '****** Coded by Menduz (lord.yo.wo@gmail.com) *******
    '*****************************************************
    CreateTLVertex.x = x
    CreateTLVertex.y = y
    CreateTLVertex.Z = Z
    CreateTLVertex.rhw = rhw
    CreateTLVertex.Color = Color
    CreateTLVertex.Specular = Specular
    CreateTLVertex.tu = tu
    CreateTLVertex.tv = tv

End Function


Private Function Geometry_Create_TLVertex(ByVal x As Single, ByVal y As Single, ByVal Z As Single, ByVal rhw As Single, ByVal Color As Long, ByVal Specular As Long, tu As Single, ByVal tv As Single) As TLVERTEX
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '**************************************************************
    Geometry_Create_TLVertex.x = x
    Geometry_Create_TLVertex.y = y
    Geometry_Create_TLVertex.Z = Z
    Geometry_Create_TLVertex.rhw = rhw
    Geometry_Create_TLVertex.Color = Color
    Geometry_Create_TLVertex.Specular = Specular
    Geometry_Create_TLVertex.tu = tu
    Geometry_Create_TLVertex.tv = tv

End Function

Private Function Geometry_Create_TLVertex2(x As Single, y As Single, Z As Single, rhw As Single, Color As Long, Specular As Long, tu1 As Single, tv1 As Single, tu2 As Single, tv2 As Single) As TLVERTEX2
    'mz
    Geometry_Create_TLVertex2.x = x
    Geometry_Create_TLVertex2.y = y
    Geometry_Create_TLVertex2.Z = Z
    Geometry_Create_TLVertex2.rhw = rhw
    Geometry_Create_TLVertex2.Color = Color
    Geometry_Create_TLVertex2.Specular = Specular
    Geometry_Create_TLVertex2.tu1 = tu1
    Geometry_Create_TLVertex2.tv1 = tv1
    Geometry_Create_TLVertex2.tu2 = tu2
    Geometry_Create_TLVertex2.tv2 = tv2

End Function

Public Sub Geometry_Create_Box(ByRef verts() As TLVERTEX, ByRef dest As RECT, ByRef src As RECT, ByRef rgb_list() As Long, Optional ByRef Textures_Width As Long, Optional ByRef Textures_Height As Long, Optional ByVal angle As Single)

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
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.Left / Textures_Width, src.bottom / Textures_Height)
    Else
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 0)

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
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, src.Left / Textures_Width, src.Top / Textures_Height)
    Else
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 0, 1)

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
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, src.Right / Textures_Width, src.bottom / Textures_Height)
    Else
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 1, 0)

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
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, src.Right / Textures_Width, src.Top / Textures_Height)
    Else
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 1, 1)

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


