Attribute VB_Name = "Graficos_Particulas"
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

Public Type Particle

    PartCountLive As Integer
    destruir  As Boolean
    friction As Single
    x As Single
    y As Single
    vector_x As Single
    vector_y As Single
    Angle As Single
    grh As grh
    alive_counter As Long
    x1 As Integer
    x2 As Integer
    y1 As Integer
    y2 As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Integer
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    rgb_list(0 To 3) As Long
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer

End Type

'*******************************************************
' PARTICULAS
'*******************************************************
Public Type Stream

    Name As String
    NumOfParticles As Long
    NumGrhs As Long
    id As Long
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
    Angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    friction As Long
    spin As Byte
    spin_speedL As Single
    spin_speedH As Single
    AlphaBlend As Byte
    gravity As Byte
    grav_strength As Long
    bounce_strength As Long
    XMove As Byte
    YMove As Byte
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    grh_list() As Long
    colortint(0 To 3) As RGB
    
    speed As Single
    life_counter As Long
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer

End Type

Public Type particle_group

    PartCountLive As Integer
    active As Boolean
    destruir As Boolean
    Creando As Integer
    Creada As Boolean
    id As Long
    map_x As Integer
    map_y As Integer
    char_index As Long

    frame_counter As Single
    frame_speed As Single
    
    stream_type As Integer

    particle_stream() As Particle
    particle_count As Long
    
    grh_index_list() As Long
    grh_index_count As Long
    
    alpha_blend As Boolean
    
    alive_counter As Long
    never_die As Boolean
    
    x1 As Integer
    x2 As Integer
    y1 As Integer
    y2 As Integer
    Angle As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Long
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    rgb_list(3) As RGBA
    
    
    'Added by Juan Martín Sotuyo Dodero
    speed As Single
    life_counter As Long
    
    'Added by David Justus
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
    
    noBorrar As Boolean

End Type

'Particle system
Public particle_group_list()  As particle_group
Public particle_group_count   As Long
Public particle_group_last As Long

'Loaded Particle groups list
Public StreamData()           As Stream
Public ParticulasTotales      As Integer
'*******************************************************
' FIN - PARTICULAS
'*******************************************************

Public Function Particle_Group_Create(ByVal map_x As Integer, ByVal map_y As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As RGBA, _
   Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
   Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
   Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
   Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal Angle As Integer, _
   Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
   Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
   Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
   Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
   Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
   Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
   Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
   Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
   Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
   Optional grh_resizex As Integer, Optional grh_resizey As Integer, Optional ByVal noBorrar As Boolean)
    
    On Error GoTo Particle_Group_Create_Err
    

    '**************************************************************
    'Author: Aaron Perkins
    'Modified by: Ryan Cain (Onezero)
    'Last Modify Date: 5/14/2003
    'Returns the particle_group_index if successful, else 0
    'Modified by Juan Martín Sotuyo Dodero
    'Modified by Augusto José Rando
    '**************************************************************
    If (map_x <> -1) And (map_y <> -1) Then
        If Map_Particle_Group_Get(map_x, map_y) = 0 Then
            Particle_Group_Create = Particle_Group_Next_Open
            Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, Angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
            particle_group_list(Particle_Group_Create).noBorrar = noBorrar

        End If

    Else
        Particle_Group_Create = Particle_Group_Next_Open
      
        Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, Angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
        particle_group_list(Particle_Group_Create).noBorrar = noBorrar

    End If

    
    Exit Function

Particle_Group_Create_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Particulas.Particle_Group_Create", Erl)
    Resume Next
    
End Function

Public Function Particle_Group_Remove(ByVal particle_group_index As Long) As Boolean
    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 1/04/2003
    '
    '*****************************************************************
    'Make sure it's a legal index
    
    On Error GoTo Particle_Group_Remove_Err
    

    If Particle_Group_Check(particle_group_index) Then
        particle_group_list(particle_group_index).never_die = False
        particle_group_list(particle_group_index).alive_counter = 0
        
        particle_group_list(particle_group_index).destruir = True
    
        Rem Particle_Group_Destroy particle_group_index
        Particle_Group_Remove = True

    End If

    
    Exit Function

Particle_Group_Remove_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Particulas.Particle_Group_Remove", Erl)
    Resume Next
    
End Function

Public Function Particle_Group_Remove_All() As Boolean
    
    On Error GoTo Particle_Group_Remove_All_Err
    

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 1/04/2003
    '
    '*****************************************************************
    Dim Index As Long
    
    For Index = 1 To particle_group_last

        'Make sure it's a legal index
        If Particle_Group_Check(Index) Then
            If Not particle_group_list(Index).noBorrar Then
                Particle_Group_Destroy Index
            End If

        End If

    Next Index
    
    Particle_Group_Remove_All = True

    
    Exit Function

Particle_Group_Remove_All_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Particulas.Particle_Group_Remove_All", Erl)
    Resume Next
    
End Function

Public Function Particle_Group_Find(ByVal id As Long) As Long

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 1/04/2003
    'Find the index related to the handle
    '*****************************************************************
    On Error GoTo ErrorHandler:

    Dim loopc As Long
    
    loopc = 1

    Do Until particle_group_list(loopc).id = id

        If loopc = particle_group_last Then
            Particle_Group_Find = 0
            Exit Function

        End If

        loopc = loopc + 1
    Loop
    
    Particle_Group_Find = loopc
    Exit Function
ErrorHandler:
    Particle_Group_Find = 0

End Function

Public Function Particle_Group_Edit(ByVal id As Long) As Long
    On Error GoTo ErrorHandler:

    particle_group_list(id).particle_count = CantPartLLuvia
        
    Exit Function
ErrorHandler:

End Function

Public Sub Particle_Group_Destroy(ByVal particle_group_index As Long)
    
    On Error GoTo Particle_Group_Destroy_Err
    

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '
    '**************************************************************
    Dim temp  As particle_group

    Dim i     As Integer

    Dim ii    As Integer

    Dim B     As Integer

    Dim antes As Integer
    
    If particle_group_list(particle_group_index).map_x > 0 And particle_group_list(particle_group_index).map_y > 0 Then
        MapData(particle_group_list(particle_group_index).map_x, particle_group_list(particle_group_index).map_y).particle_group = 0
    
    ElseIf particle_group_list(particle_group_index).char_index Then

        If Char_Check(particle_group_list(particle_group_index).char_index) Then
       
            For i = 1 To charlist(particle_group_list(particle_group_index).char_index).particle_count

                If charlist(particle_group_list(particle_group_index).char_index).particle_group(i) = particle_group_index Then
                    antes = charlist(particle_group_list(particle_group_index).char_index).particle_count
                    charlist(particle_group_list(particle_group_index).char_index).particle_count = charlist(particle_group_list(particle_group_index).char_index).particle_count - 1
                    charlist(particle_group_list(particle_group_index).char_index).particle_group(i) = 0
                   
                    ii = i

                    For B = ii To antes - 1
                        charlist(particle_group_list(particle_group_index).char_index).particle_group(B) = charlist(particle_group_list(particle_group_index).char_index).particle_group(B + 1)
                        ' charlist(particle_group_list(particle_group_index).char_index).particle_group(b + 1) = 0
                    Next B

                    Rem       ReDim Preserve charlist(particle_group_list(particle_group_index).char_index).particle_group(1 To charlist(particle_group_list(particle_group_index).char_index).particle_count)
                    Rem Else
                    Rem ReDim charlist(particle_group_list(particle_group_index).char_index).particle_group(0)
                    '  End If
                          
                    Exit For
                    
                End If
                
            Next i

        End If
    
    End If
    
    particle_group_list(particle_group_index) = temp
    
    'Update array size
    If particle_group_index = particle_group_last Then

        Do Until particle_group_list(particle_group_last).active
            particle_group_last = particle_group_last - 1

            If particle_group_last = 0 Then
                particle_group_count = 0
                Exit Sub

            End If

        Loop
        ReDim Preserve particle_group_list(1 To particle_group_last)
        
    End If

    particle_group_count = particle_group_count - 1
    
    
    Exit Sub

Particle_Group_Destroy_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Particulas.Particle_Group_Destroy", Erl)
    Resume Next
    
End Sub

Public Sub Particle_Group_Make(ByVal particle_group_index As Long, ByVal map_x As Integer, ByVal map_y As Integer, _
   ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As RGBA, _
   Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
   Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
   Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal Angle As Integer, _
   Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
   Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
   Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
   Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
   Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
   Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
   Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
   Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
   Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
   Optional grh_resizex As Integer, Optional grh_resizey As Integer)
    
    On Error GoTo Particle_Group_Make_Err
    
                                
    '*****************************************************************
    'Author: Aaron Perkins
    'Modified by: Ryan Cain (Onezero)
    'Last Modify Date: 5/15/2003
    'Makes a new particle effect
    'Modified by Juan Martín Sotuyo Dodero
    '*****************************************************************
    'Update array size
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)

    End If

    particle_group_count = particle_group_count + 1
    
    'Make active
    particle_group_list(particle_group_index).active = True
    
    'Map pos
    If (map_x <> -1) And (map_y <> -1) Then
        particle_group_list(particle_group_index).map_x = map_x
        particle_group_list(particle_group_index).map_y = map_y

    End If
    
    'Grh list
    ReDim particle_group_list(particle_group_index).grh_index_list(1 To UBound(grh_index_list))
    particle_group_list(particle_group_index).grh_index_list() = grh_index_list()
    particle_group_list(particle_group_index).grh_index_count = UBound(grh_index_list)
    
    'Sets alive vars
    If alive_counter = -1 Then
        particle_group_list(particle_group_index).alive_counter = -1
        particle_group_list(particle_group_index).never_die = True
    Else
        particle_group_list(particle_group_index).alive_counter = alive_counter
        particle_group_list(particle_group_index).never_die = False

    End If
    
    'alpha blending
    particle_group_list(particle_group_index).alpha_blend = alpha_blend

    'stream type
    particle_group_list(particle_group_index).stream_type = stream_type
    
    'speed
    particle_group_list(particle_group_index).frame_speed = frame_speed
    
    particle_group_list(particle_group_index).x1 = x1
    particle_group_list(particle_group_index).y1 = y1
    particle_group_list(particle_group_index).x2 = x2
    particle_group_list(particle_group_index).y2 = y2
    particle_group_list(particle_group_index).Angle = Angle
    particle_group_list(particle_group_index).vecx1 = vecx1
    particle_group_list(particle_group_index).vecx2 = vecx2
    particle_group_list(particle_group_index).vecy1 = vecy1
    particle_group_list(particle_group_index).vecy2 = vecy2
    particle_group_list(particle_group_index).life1 = life1
    particle_group_list(particle_group_index).life2 = life2
    particle_group_list(particle_group_index).fric = fric
    particle_group_list(particle_group_index).spin = spin
    particle_group_list(particle_group_index).spin_speedL = spin_speedL
    particle_group_list(particle_group_index).spin_speedH = spin_speedH
    particle_group_list(particle_group_index).gravity = gravity
    particle_group_list(particle_group_index).grav_strength = grav_strength
    particle_group_list(particle_group_index).bounce_strength = bounce_strength
    particle_group_list(particle_group_index).XMove = XMove
    particle_group_list(particle_group_index).YMove = YMove
    particle_group_list(particle_group_index).move_x1 = move_x1
    particle_group_list(particle_group_index).move_x2 = move_x2
    particle_group_list(particle_group_index).move_y1 = move_y1
    particle_group_list(particle_group_index).move_y2 = move_y2
    
    particle_group_list(particle_group_index).rgb_list(0) = rgb_list(0)
    particle_group_list(particle_group_index).rgb_list(1) = rgb_list(1)
    particle_group_list(particle_group_index).rgb_list(2) = rgb_list(2)
    particle_group_list(particle_group_index).rgb_list(3) = rgb_list(3)
    
    particle_group_list(particle_group_index).grh_resize = grh_resize
    particle_group_list(particle_group_index).grh_resizex = grh_resizex
    particle_group_list(particle_group_index).grh_resizey = grh_resizey
    
    'handle
    particle_group_list(particle_group_index).id = id
    particle_group_list(particle_group_index).Creando = True
    
    'create particle stream

    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)
    
    'Escribo La particula en el mapdata(x,y).particle_group :P
    If (map_x <> -1) And (map_y <> -1) Then
        MapData(map_x, map_y).particle_group = particle_group_index

    End If
    
    
    Exit Sub

Particle_Group_Make_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Particulas.Particle_Group_Make", Erl)
    Resume Next
    
End Sub

Public Sub Char_Particle_Group_Make(ByVal particle_group_index As Long, ByVal char_index As Integer, ByVal particle_char_index As Integer, _
   ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As RGBA, _
   Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
   Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
   Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal Angle As Integer, _
   Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
   Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
   Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
   Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
   Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
   Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
   Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
   Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
   Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
   Optional grh_resizex As Integer, Optional grh_resizey As Integer)
    
    On Error GoTo Char_Particle_Group_Make_Err
    
                                
    '*****************************************************************
    'Author: Aaron Perkins
    'Modified by: Ryan Cain (Onezero)
    'Last Modify Date: 5/15/2003
    'Makes a new particle effect
    'Modified by Juan Martín Sotuyo Dodero
    '*****************************************************************
    'Update array size
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)

    End If

    particle_group_count = particle_group_count + 1
    
    'Make active
    particle_group_list(particle_group_index).active = True
    
    'Char index
    particle_group_list(particle_group_index).char_index = char_index
    
    'Grh list
    ReDim particle_group_list(particle_group_index).grh_index_list(1 To UBound(grh_index_list))
    particle_group_list(particle_group_index).grh_index_list() = grh_index_list()
    particle_group_list(particle_group_index).grh_index_count = UBound(grh_index_list)
    
    'Sets alive vars
    If alive_counter = -1 Then

        particle_group_list(particle_group_index).alive_counter = -1
        particle_group_list(particle_group_index).never_die = True
    Else
        '  Debug.Print alive_counter
        particle_group_list(particle_group_index).alive_counter = alive_counter
        particle_group_list(particle_group_index).never_die = False

    End If
    
    'alpha blending
    particle_group_list(particle_group_index).alpha_blend = alpha_blend
    
    'stream type
    particle_group_list(particle_group_index).stream_type = stream_type
    
    'speed
    particle_group_list(particle_group_index).frame_speed = frame_speed
    
    particle_group_list(particle_group_index).x1 = x1
    particle_group_list(particle_group_index).y1 = y1
    particle_group_list(particle_group_index).x2 = x2
    particle_group_list(particle_group_index).y2 = y2
    particle_group_list(particle_group_index).Angle = Angle
    particle_group_list(particle_group_index).vecx1 = vecx1
    particle_group_list(particle_group_index).vecx2 = vecx2
    particle_group_list(particle_group_index).vecy1 = vecy1
    particle_group_list(particle_group_index).vecy2 = vecy2
    particle_group_list(particle_group_index).life1 = life1
    particle_group_list(particle_group_index).life2 = life2
    particle_group_list(particle_group_index).fric = fric
    particle_group_list(particle_group_index).spin = spin
    particle_group_list(particle_group_index).spin_speedL = spin_speedL
    particle_group_list(particle_group_index).spin_speedH = spin_speedH
    particle_group_list(particle_group_index).gravity = gravity
    particle_group_list(particle_group_index).grav_strength = grav_strength
    particle_group_list(particle_group_index).bounce_strength = bounce_strength
    particle_group_list(particle_group_index).XMove = XMove
    particle_group_list(particle_group_index).YMove = YMove
    particle_group_list(particle_group_index).move_x1 = move_x1
    particle_group_list(particle_group_index).move_x2 = move_x2
    particle_group_list(particle_group_index).move_y1 = move_y1
    particle_group_list(particle_group_index).move_y2 = move_y2
    
    particle_group_list(particle_group_index).rgb_list(0) = rgb_list(0)
    particle_group_list(particle_group_index).rgb_list(1) = rgb_list(1)
    particle_group_list(particle_group_index).rgb_list(2) = rgb_list(2)
    particle_group_list(particle_group_index).rgb_list(3) = rgb_list(3)
    
    particle_group_list(particle_group_index).grh_resize = grh_resize
    particle_group_list(particle_group_index).grh_resizex = grh_resizex
    particle_group_list(particle_group_index).grh_resizey = grh_resizey
    
    'handle
    particle_group_list(particle_group_index).id = id
    
    'create particle stream
    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)
    
    'plot particle group on char
    
    charlist(char_index).particle_group(particle_char_index) = particle_group_index

    'MsgBox particle_group_list(particle_group_index).stream_type
    
    Exit Sub

Char_Particle_Group_Make_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Particulas.Char_Particle_Group_Make", Erl)
    Resume Next
    
End Sub

Public Sub Particle_Incrementar(ByVal id As Integer)
    
    On Error GoTo Particle_Incrementar_Err
    

    If particle_group_list(id).Creando < particle_group_list(id).particle_count Then
        particle_group_list(id).Creando = particle_group_list(id).Creando + 1
    Else
        particle_group_list(id).Creada = True

    End If

    If particle_group_list(id).char_index > 0 Then
        charlist(particle_group_list(id).char_index).particle_count = particle_group_list(id).Creando

    End If

    
    Exit Sub

Particle_Incrementar_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Particulas.Particle_Incrementar", Erl)
    Resume Next
    
End Sub

Public Sub Particle_Group_Render(ByVal particle_group_index As Long, ByVal screen_x As Integer, ByVal screen_y As Integer)
    
    On Error GoTo Particle_Group_Render_Err
    

    '*****************************************************************
    'Author: Aaron Perkins
    'Modified by: Ryan Cain (Onezero)
    'Modified by: Juan Martín Sotuyo Dodero
    'Last Modify Date: 5/15/2003
    'Renders a particle stream at a paticular screen point
    '*****************************************************************
    

    Dim loopc            As Long

    Dim temp_rgb(0 To 3) As RGBA

    Dim no_move          As Boolean
    
    'Set colors
    'Exit Sub
    particle_group_index = min(particle_group_index, UBound(particle_group_list))
    temp_rgb(0) = particle_group_list(particle_group_index).rgb_list(0)
    temp_rgb(1) = particle_group_list(particle_group_index).rgb_list(1)
    temp_rgb(2) = particle_group_list(particle_group_index).rgb_list(2)
    temp_rgb(3) = particle_group_list(particle_group_index).rgb_list(3)

    If particle_group_list(particle_group_index).alive_counter Then
        'See if it is time to move a particle
        particle_group_list(particle_group_index).frame_counter = particle_group_list(particle_group_index).frame_counter + timerTicksPerFrame

        If particle_group_list(particle_group_index).frame_counter > particle_group_list(particle_group_index).frame_speed Then
            particle_group_list(particle_group_index).frame_counter = 0
            no_move = False
                    
        Else
            no_move = True
                    
        End If
  
        Dim cantidad As Long

        cantidad = particle_group_list(particle_group_index).particle_count
 
        For loopc = 1 To cantidad
            
            Particle_Render particle_group_list(particle_group_index).particle_stream(loopc), _
               screen_x, screen_y, _
               particle_group_list(particle_group_index).grh_index_list(Round(RandomNumber(1, particle_group_list(particle_group_index).grh_index_count), 0)), _
               temp_rgb(), _
               particle_group_list(particle_group_index).alpha_blend, no_move, _
               particle_group_list(particle_group_index).x1, particle_group_list(particle_group_index).y1, particle_group_list(particle_group_index).Angle, _
               particle_group_list(particle_group_index).vecx1, particle_group_list(particle_group_index).vecx2, _
               particle_group_list(particle_group_index).vecy1, particle_group_list(particle_group_index).vecy2, _
               particle_group_list(particle_group_index).life1, particle_group_list(particle_group_index).life2, _
               particle_group_list(particle_group_index).fric, particle_group_list(particle_group_index).spin_speedL, _
               particle_group_list(particle_group_index).gravity, particle_group_list(particle_group_index).grav_strength, _
               particle_group_list(particle_group_index).bounce_strength, particle_group_list(particle_group_index).x2, _
               particle_group_list(particle_group_index).y2, particle_group_list(particle_group_index).XMove, _
               particle_group_list(particle_group_index).move_x1, particle_group_list(particle_group_index).move_x2, _
               particle_group_list(particle_group_index).move_y1, particle_group_list(particle_group_index).move_y2, _
               particle_group_list(particle_group_index).YMove, particle_group_list(particle_group_index).spin_speedH, _
               particle_group_list(particle_group_index).spin, particle_group_list(particle_group_index).grh_resize, particle_group_list(particle_group_index).grh_resizex, particle_group_list(particle_group_index).grh_resizey, particle_group_index, particle_group_list(particle_group_index).destruir
        Next loopc

        '
       
        'Render particle

        If no_move = False Then
        
            'Update the group alive counter
            If particle_group_list(particle_group_index).never_die = False Then
                particle_group_list(particle_group_index).alive_counter = particle_group_list(particle_group_index).alive_counter - 1

            End If

        End If
  
        'If it's dead destroy it
    Else
      
        'Revisar si se saca esto. Ladder

        particle_group_list(particle_group_index).destruir = True
            
        If particle_group_list(particle_group_index).PartCountLive <= 2 Then
                  
            Particle_Group_Destroy particle_group_index
            Exit Sub

        End If

        particle_group_list(particle_group_index).frame_counter = particle_group_list(particle_group_index).frame_counter + timerTicksPerFrame

        If particle_group_list(particle_group_index).frame_counter > particle_group_list(particle_group_index).frame_speed Then
            particle_group_list(particle_group_index).frame_counter = 0
            no_move = False
        Else
            no_move = True

        End If

        'If it's still alive render all the particles inside
        For loopc = 1 To particle_group_list(particle_group_index).particle_count
        
            'Render particle
            Particle_Render particle_group_list(particle_group_index).particle_stream(loopc), _
               screen_x, screen_y, _
               particle_group_list(particle_group_index).grh_index_list(Round(RandomNumber(1, particle_group_list(particle_group_index).grh_index_count), 0)), _
               temp_rgb(), _
               particle_group_list(particle_group_index).alpha_blend, no_move, _
               particle_group_list(particle_group_index).x1, particle_group_list(particle_group_index).y1, particle_group_list(particle_group_index).Angle, _
               particle_group_list(particle_group_index).vecx1, particle_group_list(particle_group_index).vecx2, _
               particle_group_list(particle_group_index).vecy1, particle_group_list(particle_group_index).vecy2, _
               particle_group_list(particle_group_index).life1, particle_group_list(particle_group_index).life2, _
               particle_group_list(particle_group_index).fric, particle_group_list(particle_group_index).spin_speedL, _
               particle_group_list(particle_group_index).gravity, particle_group_list(particle_group_index).grav_strength, _
               particle_group_list(particle_group_index).bounce_strength, particle_group_list(particle_group_index).x2, _
               particle_group_list(particle_group_index).y2, particle_group_list(particle_group_index).XMove, _
               particle_group_list(particle_group_index).move_x1, particle_group_list(particle_group_index).move_x2, _
               particle_group_list(particle_group_index).move_y1, particle_group_list(particle_group_index).move_y2, _
               particle_group_list(particle_group_index).YMove, particle_group_list(particle_group_index).spin_speedH, _
               particle_group_list(particle_group_index).spin, particle_group_list(particle_group_index).grh_resize, particle_group_list(particle_group_index).grh_resizex, particle_group_list(particle_group_index).grh_resizey, particle_group_index, particle_group_list(particle_group_index).destruir
        Next loopc

        particle_group_list(particle_group_index).destruir = True

    End If

    
    Exit Sub

Particle_Group_Render_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Particulas.Particle_Group_Render", Erl)
    Resume Next
    
End Sub

Public Sub Particle_Render(ByRef temp_particle As Particle, ByVal screen_x As Integer, ByVal screen_y As Integer, _
   ByVal grh_index As Long, ByRef rgb_list() As RGBA, _
   Optional ByVal alpha_blend As Boolean, Optional ByVal no_move As Boolean, _
   Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal Angle As Integer, _
   Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
   Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
   Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
   Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
   Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
   Optional ByVal bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
   Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
   Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
   Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
   Optional grh_resizex As Integer, Optional grh_resizey As Integer, Optional particle_group_index As Long, Optional destruir As Boolean)
    '**************************************************************
    'Author: Aaron Perkins
    'Modified by: Ryan Cain (Onezero)
    'Modified by: Juan Martín Sotuyo Dodero
    'Last Modify Date: 5/15/2003
    '**************************************************************
    
    On Error GoTo Particle_Render_Err
    

    If no_move = False Then

        If temp_particle.alive_counter = 0 And Not destruir Then
            'Start new particle
            InitGrh temp_particle.grh, grh_index
            temp_particle.grh.Alpha = 255
            temp_particle.x = RandomNumber(x1, x2) - (32 / 2)
            temp_particle.y = RandomNumber(y1, y2) - (32 / 2)
            temp_particle.vector_x = RandomNumber(vecx1, vecx2)
            temp_particle.vector_y = RandomNumber(vecy1, vecy2)
            temp_particle.Angle = Angle
            temp_particle.alive_counter = RandomNumber(life1, life2)
            temp_particle.friction = fric
            particle_group_list(particle_group_index).PartCountLive = particle_group_list(particle_group_index).PartCountLive + 1

        Else
            'Continue old particle
            'Do gravity
            
            If temp_particle.alive_counter = 0 And destruir Then
                temp_particle.grh.GrhIndex = 0
                
            End If

            If gravity = True Then
                temp_particle.vector_y = temp_particle.vector_y + grav_strength

                If temp_particle.y > 0 Then
                    'bounce
                    temp_particle.vector_y = bounce_strength

                End If

            End If

            'Do rotation
            If spin = True Then temp_particle.grh.Angle = temp_particle.grh.Angle + RandomNumber(spin_speedL, spin_speedH) / 5
            Do While temp_particle.grh.Angle >= 360
                temp_particle.grh.Angle = temp_particle.grh.Angle - 360
            Loop
            
            If XMove = True Then temp_particle.vector_x = RandomNumber(move_x1, move_x2)
            If YMove = True Then temp_particle.vector_y = RandomNumber(move_y1, move_y2)

        End If
        
        'Add in vector
        temp_particle.x = temp_particle.x + (temp_particle.vector_x \ temp_particle.friction)
        temp_particle.y = temp_particle.y + (temp_particle.vector_y \ temp_particle.friction)
    
        'decrement counter
        temp_particle.alive_counter = temp_particle.alive_counter - 1

        If temp_particle.alive_counter = 0 Then
            particle_group_list(particle_group_index).PartCountLive = particle_group_list(particle_group_index).PartCountLive - 1
        End If
                    
    End If
    
    temp_particle.grh_resize = grh_resize
    temp_particle.grh_resizex = grh_resizex
    temp_particle.grh_resizey = grh_resizey
    'Draw it
    If screen_x = -1000 Then Exit Sub
    
    'Particulas Grises si esta muerto Ladder
    If UserEstado = 1 Then
        Call RGBAList(rgb_list, 100, 100, 100, 100)

    ElseIf UserCiego = True Then
        Call RGBAList(rgb_list, 5, 5, 5, 5)

    ElseIf Not alpha_blend Then
        Call SetRGBA(rgb_list(0), rgb_list(0).r, rgb_list(0).G, rgb_list(0).B, temp_particle.grh.Alpha)
        Call SetRGBA(rgb_list(1), rgb_list(1).r, rgb_list(1).G, rgb_list(1).B, temp_particle.grh.Alpha)
        Call SetRGBA(rgb_list(2), rgb_list(2).r, rgb_list(2).G, rgb_list(2).B, temp_particle.grh.Alpha)
        Call SetRGBA(rgb_list(3), rgb_list(3).r, rgb_list(3).G, rgb_list(3).B, temp_particle.grh.Alpha)
    End If
    
    If grh_resize = True Then
    
        If temp_particle.grh.GrhIndex Then
    
            Grh_Render_Advance temp_particle.grh, temp_particle.x + screen_x, temp_particle.y + screen_y, grh_resizex, grh_resizey, rgb_list(), True, True, alpha_blend
            
            Exit Sub

        End If

    End If

    If temp_particle.grh.GrhIndex Then

        Grh_Render temp_particle.grh, temp_particle.x + screen_x, temp_particle.y + screen_y, rgb_list(), True, True, alpha_blend

    End If
    
    
    Exit Sub

Particle_Render_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Particulas.Particle_Render", Erl)
    Resume Next
    
End Sub

Public Function Particle_Type_Get(ByVal particle_Index As Long) As Long
    
    On Error GoTo Particle_Type_Get_Err
    

    '*****************************************************************
    'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
    'Last Modify Date: 8/27/2003
    'Returns the stream type of a particle stream
    '*****************************************************************
    If Particle_Group_Check(particle_Index) Then
        Particle_Type_Get = particle_group_list(particle_Index).stream_type

    End If

    
    Exit Function

Particle_Type_Get_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Particulas.Particle_Type_Get", Erl)
    Resume Next
    
End Function

Public Function Engine_MeteoParticle_Get() As Long
    '*****************************************************************
    'Author: Augusto José Rando
    'Last Modify Date: 6/11/2002
    '*****************************************************************
    
    On Error GoTo Engine_MeteoParticle_Get_Err
    
    Engine_MeteoParticle_Get = MeteoParticle

    
    Exit Function

Engine_MeteoParticle_Get_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Particulas.Engine_MeteoParticle_Get", Erl)
    Resume Next
    
End Function

Public Function Map_Particle_Group_Get(ByVal map_x As Integer, ByVal map_y As Integer) As Long
    
    On Error GoTo Map_Particle_Group_Get_Err
    

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 2/20/2003
    'Checks to see if a tile position has a particle_group_index and return it
    '*****************************************************************
    If InMapBounds(map_x, map_y) Then
        Map_Particle_Group_Get = MapData(map_x, map_y).particle_group
    Else
        Map_Particle_Group_Get = 0

    End If

    
    Exit Function

Map_Particle_Group_Get_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Particulas.Map_Particle_Group_Get", Erl)
    Resume Next
    
End Function

Public Sub Engine_MeteoParticle_Set(ByVal meteo_part As Long)
    '*****************************************************************
    'Author: Augusto José Rando
    'Last Modify Date: 6/11/2002
    '*****************************************************************
    
    On Error GoTo Engine_MeteoParticle_Set_Err
    
    
    If (meteo_part = -1) And (MeteoParticle <> 0) Then
        Call Particle_Group_Remove(MeteoParticle)
        
    ElseIf (meteo_part <> -1) Then

        If MeteoParticle <> 0 Then Call Particle_Group_Remove(MeteoParticle)
        MeteoParticle = General_Particle_Create(meteo_part, -1, -1)
        MeteoIndex = particle_group_last

        Dim i As Integer
        For i = 1 To 500
            Call Particle_Group_Render(MeteoIndex, -1000, -1000)
        Next i

    End If
    
    
    Exit Sub

Engine_MeteoParticle_Set_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Particulas.Engine_MeteoParticle_Set", Erl)
    Resume Next
    
End Sub

Public Sub Engine_spell_Particle_Set(ByVal spell_part As Long)
    '*****************************************************************
    'Author: Augusto José Rando
    'Last Modify Date: 6/11/2002
    '*****************************************************************
    
    On Error GoTo Engine_spell_Particle_Set_Err
    
    
    If (spell_part = -1) And (spell_particle <> 0) Then
        Call Particle_Group_Remove(spell_particle)
    ElseIf (spell_part <> -1) Then

        If spell_particle <> 0 Then Call Particle_Group_Remove(spell_particle)
        spell_particle = General_Particle_Create(spell_part, -1, -1)

    End If
    
    
    Exit Sub

Engine_spell_Particle_Set_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Particulas.Engine_spell_Particle_Set", Erl)
    Resume Next
    
End Sub

Public Sub Engine_Select_Particle_Set(ByVal Select_particle As Long)
    '*****************************************************************
    'Author: Augusto José Rando
    'Last Modify Date: 6/11/2002
    '*****************************************************************
    
    On Error GoTo Engine_Select_Particle_Set_Err
    
    
    If (Select_particle = -1) And (Select_part <> 0) Then
        Call Particle_Group_Remove(Select_part)
    ElseIf (Select_part <> -1) Then

        If Select_part <> 0 Then Call Particle_Group_Remove(Select_part)
        Select_part = General_Particle_Create(Select_particle, -1, -1, , True)

    End If
    
    
    Exit Sub

Engine_Select_Particle_Set_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Particulas.Engine_Select_Particle_Set", Erl)
    Resume Next
    
End Sub

Public Function Particle_Group_Check(ByVal particle_group_index As Long) As Boolean
    
    On Error GoTo Particle_Group_Check_Err
    

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 1/04/2003
    '
    '**************************************************************
    'check index
    If particle_group_index > 0 And particle_group_index <= particle_group_last Then
        If particle_group_list(particle_group_index).active Then
            Particle_Group_Check = True

        End If

    End If

    
    Exit Function

Particle_Group_Check_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Particulas.Particle_Group_Check", Erl)
    Resume Next
    
End Function

Public Function Particle_Group_Next_Open() As Long

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '
    '*****************************************************************
    On Error GoTo ErrorHandler:

    Dim loopc As Long
    
    If particle_group_last = 0 Then
        Particle_Group_Next_Open = 1
        Exit Function

    End If
    
    loopc = 1

    Do Until particle_group_list(loopc).active = False

        If loopc = particle_group_last Then
            Particle_Group_Next_Open = particle_group_last + 1
            Exit Function

        End If

        loopc = loopc + 1
    Loop
    
    Particle_Group_Next_Open = loopc

    Exit Function

ErrorHandler:
    Particle_Group_Next_Open = 1

End Function

Public Function General_Char_Particle_Create(ByVal ParticulaInd As Long, ByVal char_index As Integer, Optional ByVal particle_life As Long = 0, Optional ByVal grh As Long = 0) As Long
    
    On Error GoTo General_Char_Particle_Create_Err

    If ParticulaInd = 0 Then Exit Function
    
    If grh > 0 Then
        Dim i As Byte
        StreamData(ParticulaInd).grh_list(1) = grh
        For i = 0 To 3
            StreamData(ParticulaInd).colortint(i).r = 255
            StreamData(ParticulaInd).colortint(i).G = 255
            StreamData(ParticulaInd).colortint(i).B = 255
        Next i
    End If
    
    Dim rgb_list(0 To 3) As RGBA
    
    Call SetRGBA(rgb_list(0), StreamData(ParticulaInd).colortint(0).r, StreamData(ParticulaInd).colortint(0).G, StreamData(ParticulaInd).colortint(0).r)
    Call SetRGBA(rgb_list(1), StreamData(ParticulaInd).colortint(1).B, StreamData(ParticulaInd).colortint(1).G, StreamData(ParticulaInd).colortint(1).R)
    Call SetRGBA(rgb_list(2), StreamData(ParticulaInd).colortint(2).B, StreamData(ParticulaInd).colortint(2).G, StreamData(ParticulaInd).colortint(2).R)
    Call SetRGBA(rgb_list(3), StreamData(ParticulaInd).colortint(3).B, StreamData(ParticulaInd).colortint(3).G, StreamData(ParticulaInd).colortint(3).R)

    General_Char_Particle_Create = Char_Particle_Group_Create(char_index, StreamData(ParticulaInd).grh_list, rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
       grh = 0 And StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).speed, , StreamData(ParticulaInd).x1, StreamData(ParticulaInd).y1, StreamData(ParticulaInd).Angle, _
       StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
       StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
       StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).x2, _
       StreamData(ParticulaInd).y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
       StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin, StreamData(ParticulaInd).grh_resize, StreamData(ParticulaInd).grh_resizex, StreamData(ParticulaInd).grh_resizey)

    
    Exit Function

General_Char_Particle_Create_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Particulas.General_Char_Particle_Create", Erl)
    Resume Next
    
End Function

Public Function General_Particle_Create(ByVal ParticulaInd As Long, ByVal x As Integer, ByVal y As Integer, Optional ByVal particle_life As Long = 0, Optional ByVal noBorrar As Boolean) As Long
    
    On Error GoTo General_Particle_Create_Err
    

    If ParticulaInd <= 0 Or ParticulaInd > UBound(StreamData) Then Exit Function

    Dim rgb_list(0 To 3) As RGBA

    Call SetRGBA(rgb_list(0), StreamData(ParticulaInd).colortint(0).B, StreamData(ParticulaInd).colortint(0).G, StreamData(ParticulaInd).colortint(0).R)
    Call SetRGBA(rgb_list(1), StreamData(ParticulaInd).colortint(1).B, StreamData(ParticulaInd).colortint(1).G, StreamData(ParticulaInd).colortint(1).R)
    Call SetRGBA(rgb_list(2), StreamData(ParticulaInd).colortint(2).B, StreamData(ParticulaInd).colortint(2).G, StreamData(ParticulaInd).colortint(2).R)
    Call SetRGBA(rgb_list(3), StreamData(ParticulaInd).colortint(3).B, StreamData(ParticulaInd).colortint(3).G, StreamData(ParticulaInd).colortint(3).R)

    General_Particle_Create = Graficos_Particulas.Particle_Group_Create(x, y, StreamData(ParticulaInd).grh_list, rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
       StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).speed, , StreamData(ParticulaInd).x1, StreamData(ParticulaInd).y1, StreamData(ParticulaInd).Angle, _
       StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
       StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
       StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).x2, _
       StreamData(ParticulaInd).y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
       StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin, StreamData(ParticulaInd).grh_resize, StreamData(ParticulaInd).grh_resizex, StreamData(ParticulaInd).grh_resizey, noBorrar)

    
    Exit Function

General_Particle_Create_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Particulas.General_Particle_Create", Erl)
    Resume Next
    
End Function

Public Function Char_Particle_Group_Create(ByVal char_index As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As RGBA, _
   Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
   Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
   Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
   Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal Angle As Integer, _
   Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
   Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
   Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
   Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
   Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
   Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
   Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
   Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
   Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
   Optional grh_resizex As Integer, Optional grh_resizey As Integer)
    '**************************************************************
    'Author: Augusto José Rando
    '**************************************************************
    
    On Error GoTo Char_Particle_Group_Create_Err
    

    Dim char_part_free_index As Integer
    
    Rem If charlist(char_index).Particula = stream_type Then Exit Function
    'If Char_Particle_Group_Find(char_index, stream_type) Then Exit Function ' hay que ver si dejar o sacar esto...
    If Not Char_Check(char_index) Then Exit Function
    char_part_free_index = Char_Particle_Group_Next_Open(char_index)
    
    If char_part_free_index > 0 Then
        Char_Particle_Group_Create = Particle_Group_Next_Open
        Char_Particle_Group_Make Char_Particle_Group_Create, char_index, char_part_free_index, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, Angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey

    End If

    
    Exit Function

Char_Particle_Group_Create_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Particulas.Char_Particle_Group_Create", Erl)
    Resume Next
    
End Function


Public Function Char_Particle_Group_Remove(ByVal char_index As Integer, ByVal stream_type As Long)
    '**************************************************************
    'Author: Augusto José Rando
    '**************************************************************
    
    On Error GoTo Char_Particle_Group_Remove_Err
    

    Dim char_part_index As Integer

    If Char_Check(char_index) Then
        char_part_index = Char_Particle_Group_Find(char_index, stream_type)

        If char_part_index = -1 Then Exit Function
        If char_part_index = 0 Then Exit Function
        
        'Call Particle_Group_Remove(char_part_index)
        Rem  particle_group_list(char_part_index).alive_counter = 20
        particle_group_list(char_part_index).alive_counter = 0
        particle_group_list(char_part_index).never_die = False
        particle_group_list(char_part_index).destruir = True
     
        'Ladder
    End If

    
    Exit Function

Char_Particle_Group_Remove_Err:
    Call RegistrarError(Err.Number, Err.Description, "Graficos_Particulas.Char_Particle_Group_Remove", Erl)
    Resume Next
    
End Function

Public Function Char_Particle_Group_Remove_All(ByVal char_index As Integer)
    
    On Error GoTo Char_Particle_Group_Remove_All_Err
    

    '**************************************************************
    'Author: Augusto José Rando
    '**************************************************************
    Dim i As Integer
    
    If Char_Check(char_index) Then

        For i = 1 To charlist(char_index).particle_count

            If charlist(char_index).particle_group(i) <> 0 Then Call Particle_Group_Remove(charlist(char_index).particle_group(i))
        Next i

    End If
    
    
    Exit Function

Char_Particle_Group_Remove_All_Err:
    Call RegistrarError(Err.number, Err.Description, "Graficos_Particulas.Char_Particle_Group_Remove_All", Erl)
    Resume Next
    
End Function

Private Function Char_Particle_Group_Find(ByVal char_index As Integer, ByVal stream_type As Long) As Integer

    '*****************************************************************
    'Author: Augusto José Rando
    'Modified: returns slot or -1
    '*****************************************************************
    On Error GoTo ErrorHandler:

    Dim i As Integer

    For i = 1 To charlist(char_index).particle_count

        If particle_group_list(charlist(char_index).particle_group(i)).stream_type = stream_type Then
            If particle_group_list(charlist(char_index).particle_group(i)).destruir = False Then
                Char_Particle_Group_Find = charlist(char_index).particle_group(i)
                Exit Function

            End If

        End If

    Next i

    Char_Particle_Group_Find = -1
ErrorHandler:

End Function

Public Function Char_Particle_Group_Next_Open(ByVal char_index As Integer) As Integer

    '*****************************************************************
    'Author: Augusto José Rando
    '*****************************************************************
    On Error GoTo ErrorHandler:

    Dim loopc As Long
    
    If charlist(char_index).particle_count = 0 Then
        charlist(char_index).particle_count = 1
        ReDim charlist(char_index).particle_group(1 To 1)
        Char_Particle_Group_Next_Open = 1
        Exit Function

    End If
    
    loopc = 1

    Do Until charlist(char_index).particle_group(loopc) = 0

        If loopc = charlist(char_index).particle_count Then
            Char_Particle_Group_Next_Open = charlist(char_index).particle_count + 1
            charlist(char_index).particle_count = Char_Particle_Group_Next_Open
            ReDim Preserve charlist(char_index).particle_group(1 To Char_Particle_Group_Next_Open)
            Exit Function

        End If

        loopc = loopc + 1
    Loop
    
    Char_Particle_Group_Next_Open = loopc

    Exit Function

ErrorHandler:
    charlist(char_index).particle_count = 1
    ReDim charlist(char_index).particle_group(1 To 1)
    Char_Particle_Group_Next_Open = 1

End Function


