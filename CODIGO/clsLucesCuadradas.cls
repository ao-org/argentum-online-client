VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLucesCuadradas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'*******************************************************
' LUCES
'*******************************************************
Private Type Light

    active As Boolean 'Do we ignore this light?
    id As Long
    map_x As Integer 'Coordinates
    map_y As Integer
    Color As RGBA 'Start Color
    range As Byte

End Type

'Light list
Private light_list()  As Light
Private light_count   As Long
Private light_last    As Long
'*******************************************************
' FIN - LUCES
'*******************************************************

Public Function Light_Remove(ByVal light_index As Long) As Boolean

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 1/04/2003
    '
    '*****************************************************************
    'Make sure it's a legal index
    If Light_Check(light_index) Then
        Light_Destroy light_index
        Light_Remove = True

    End If

End Function

Friend Function Light_Color_Value_Get(ByVal light_index As Long, ByRef color_value As RGBA) As Boolean

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 2/28/2003
    '
    '*****************************************************************
    'Make sure it's a legal index
    If Light_Check(light_index) Then
        color_value = light_list(light_index).Color
        Light_Color_Value_Get = True

    End If

End Function

Friend Function Light_Create(ByVal map_x As Integer, ByVal map_y As Integer, ByRef color_value As RGBA, Optional ByVal range As Byte = 1, Optional ByVal id As Long) As Long

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Returns the light_index if successful, else 0
    'Edited by Juan Martín Sotuyo Dodero
    '**************************************************************
    If InMapBounds(map_x, map_y) Then

        'Make sure there is no light in the given map pos
        If Map_Light_Get(map_x, map_y) <> 0 Then
            Light_Create = 0
            Exit Function

        End If

        Light_Create = Light_Next_Open
        Light_Make Light_Create, map_x, map_y, color_value, range, id

    End If

End Function

Public Function Map_Light_Get(ByVal map_x As Integer, ByVal map_y As Integer) As Long

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 2/20/2003
    'Checks to see if a tile position has a light_index and return it
    '*****************************************************************
    On Error GoTo ErrorHandler:

    Dim loopc As Long
    
    'We start from the back, to get the last light to be placed on the tile first
    If light_last = 0 Then Exit Function
    
    loopc = light_last

    Do Until light_list(loopc).map_x = map_x And light_list(loopc).map_y = map_y

        If loopc = 0 Then
            Map_Light_Get = 0
            Exit Function

        End If

        loopc = loopc - 1

        If loopc = 0 Then Exit Function
    Loop
    
    Map_Light_Get = loopc
    Exit Function
ErrorHandler:
    Map_Light_Get = 0

End Function

Public Function Light_Move(ByVal light_index As Long, ByVal map_x As Integer, ByVal map_y As Integer) As Boolean

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Returns true if successful, else false
    '**************************************************************
    'Make sure it's a legal CharIndex
    If Light_Check(light_index) Then

        'Make sure it's a legal move
        If InMapBounds(map_x, map_y) Then
        
            'Move it
            Light_Erase light_index
            light_list(light_index).map_x = map_x
            light_list(light_index).map_y = map_y
    
            Light_Move = True
            
        End If

    End If

End Function

Public Function Light_Move_By_Head(ByVal light_index As Long, ByVal Heading As Byte) As Boolean

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 15/05/2002
    'Returns true if successful, else false
    '**************************************************************
    Dim map_x As Integer

    Dim map_y As Integer

    Dim nX    As Integer

    Dim nY    As Integer

    Dim addy  As Byte

    Dim addx  As Byte

    'Check for valid heading
    If Heading < 1 Or Heading > 8 Then
        Light_Move_By_Head = False
        Exit Function

    End If

    'Make sure it's a legal CharIndex
    If Light_Check(light_index) Then
    
        map_x = light_list(light_index).map_x
        map_y = light_list(light_index).map_y

        Select Case Heading

            Case E_Heading.NORTH
                addy = -1
        
            Case E_Heading.EAST
                addx = 1
        
            Case E_Heading.south
                addy = 1
            
            Case E_Heading.WEST
                addx = -1

        End Select
        
        nX = map_x + addx
        nY = map_y + addy
        
        'Make sure it's a legal move
        If InMapBounds(nX, nY) Then
        
            'Move it
            Light_Erase light_index

            light_list(light_index).map_x = nX
            light_list(light_index).map_y = nY
    
            Light_Move_By_Head = True
            
        End If

    End If

End Function

Private Sub Light_Make(ByVal light_index As Long, ByVal map_x As Integer, ByVal map_y As Integer, rgb_value As RGBA, ByVal range As Long, Optional ByVal id As Long)

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '
    '*****************************************************************
    'Update array size
    If light_index > light_last Then
        light_last = light_index
        ReDim Preserve light_list(1 To light_last)

    End If

    light_count = light_count + 1
    
    'Make active
    light_list(light_index).active = True

    light_list(light_index).map_x = map_x
    light_list(light_index).map_y = map_y

    light_list(light_index).Color = rgb_value
    light_list(light_index).range = range
    light_list(light_index).id = id

End Sub

Private Function Light_Check(ByVal light_index As Long) As Boolean

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 1/04/2003
    '
    '**************************************************************
    'check light_index
    If light_index > 0 And light_index <= light_last Then
        If light_list(light_index).active Then
            Light_Check = True

        End If

    End If

End Function

Public Sub Light_Render_All()

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '
    '**************************************************************
    Dim loop_counter As Long
            
    For loop_counter = 1 To light_count
        
        If light_list(loop_counter).active Then
            Light_Render loop_counter

        End If
    
    Next loop_counter

End Sub

Private Sub Light_Render(ByVal light_index As Long)

    'menduz
    Dim min_x As Integer

    Dim min_y As Integer

    Dim max_x As Integer

    Dim max_y As Integer

    Dim X     As Integer

    Dim Y     As Integer

    Dim ia    As Single

    Dim i     As Integer

    Dim Color As RGBA
    
    'Set up light borders
    min_x = light_list(light_index).map_x - light_list(light_index).range
    min_y = light_list(light_index).map_y - light_list(light_index).range
    max_x = light_list(light_index).map_x + light_list(light_index).range
    max_y = light_list(light_index).map_y + light_list(light_index).range
    
    'Set color
    Color = light_list(light_index).Color
    Color.A = 255
    
    MapData(light_list(light_index).map_x, light_list(light_index).map_y).light_value(0) = Color
    MapData(light_list(light_index).map_x, light_list(light_index).map_y).light_value(1) = Color
    MapData(light_list(light_index).map_x, light_list(light_index).map_y).light_value(2) = Color
    MapData(light_list(light_index).map_x, light_list(light_index).map_y).light_value(3) = Color
                
    'For ia = 0 To 6.5 Step 0.05
    '    For i = 1 To light_list(light_index).range / 2
    '        x = (Sin(ia) * i) + light_list(light_index).map_x
    '        y = (Cos(ia) * i) + light_list(light_index).map_y
    '        'Debug.Print x; y
    '        If InMapBounds(x, y) Then
    '            If i = (light_list(light_index).range / 2) Then
    '                If MapData(x - 1, y - 1).light_value(2) <> color Then MapData(x, y).light_value(0) = 0
    '                If MapData(x + 1, y + 1).light_value(0) <> color Then MapData(x, y).light_value(2) = 0
    '                If MapData(x + 1, y - 1).light_value(1) <> color Then MapData(x, y).light_value(1) = 0
    '                If MapData(x - 1, y + 1).light_value(3) <> color Then MapData(x, y).light_value(3) = 0
    '                If MapData(x - 1, y).light_value(0) = color Then MapData(x, y).light_value(3) = color
    '                If MapData(x + 1, y).light_value(1) = color And MapData(x + 1, y).light_value(2) = color Then
    '                    MapData(x, y).light_value(0) = &HFFFF0FFF
    '                    MapData(x, y).light_value(3) = &HFFFF0FFF
    '                End If
    '                If MapData(x, y - 1).light_value(2) = color Then MapData(x, y).light_value(3) = color
    '                If MapData(x, y + 1).light_value(3) = color Then MapData(x, y).light_value(1) = color''

    '            Else
    '                MapData(x, y).light_value(0) = &HFFFFFFFF
    '                MapData(x, y).light_value(1) = &HFFFFFFFF
    '                MapData(x, y).light_value(2) = &HFFFFFFFF
    '                MapData(x, y).light_value(3) = &HFFFFFFFF
    '            End If
    '        End If
    '    Next i
    'Next ia
    'Arrange corners
    'NE
    If InMapBounds(min_x, min_y) Then
        MapData(min_x, min_y).light_value(2) = Color

    End If

    'NW
    If InMapBounds(max_x, min_y) Then
        MapData(max_x, min_y).light_value(0) = Color

    End If

    'SW
    If InMapBounds(max_x, max_y) Then
        MapData(max_x, max_y).light_value(1) = Color

    End If

    'SE
    If InMapBounds(min_x, max_y) Then
        MapData(min_x, max_y).light_value(3) = Color

    End If
    
    'Arrange borders
    'Upper border
    For X = min_x + 1 To max_x - 1

        If InMapBounds(X, min_y) Then
            MapData(X, min_y).light_value(0) = Color
            MapData(X, min_y).light_value(2) = Color

        End If

    Next X
    
    'Lower border
    For X = min_x + 1 To max_x - 1

        If InMapBounds(X, max_y) Then
            MapData(X, max_y).light_value(1) = Color
            MapData(X, max_y).light_value(3) = Color

        End If

    Next X
    
    'Left border
    For Y = min_y + 1 To max_y - 1

        If InMapBounds(min_x, Y) Then
            MapData(min_x, Y).light_value(2) = Color
            MapData(min_x, Y).light_value(3) = Color

        End If

    Next Y
    
    'Right border
    For Y = min_y + 1 To max_y - 1

        If InMapBounds(max_x, Y) Then
            MapData(max_x, Y).light_value(0) = Color
            MapData(max_x, Y).light_value(1) = Color

        End If

    Next Y
    
    'Set the inner part of the light
    For X = min_x + 1 To max_x - 1
        For Y = min_y + 1 To max_y - 1

            If InMapBounds(X, Y) Then
                MapData(X, Y).light_value(0) = Color
                MapData(X, Y).light_value(1) = Color
                MapData(X, Y).light_value(2) = Color
                MapData(X, Y).light_value(3) = Color

            End If

        Next Y
    Next X

End Sub

Private Function Light_Next_Open() As Long

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '
    '*****************************************************************
    On Error GoTo ErrorHandler:

    Dim loopc As Long
    
    If light_last = 0 Then
        Light_Next_Open = 1
        Exit Function

    End If
    
    loopc = 1

    Do Until light_list(loopc).active = False

        If loopc = light_last Then
            Light_Next_Open = light_last + 1
            Exit Function

        End If

        loopc = loopc + 1
    Loop
    
    Light_Next_Open = loopc
    Exit Function
ErrorHandler:
    Light_Next_Open = 1

End Function

Public Function Light_Find(ByVal id As Long) As Long

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 1/04/2003
    'Find the index related to the handle
    '*****************************************************************
    On Error GoTo ErrorHandler:

    Dim loopc As Long
    
    loopc = 1

    Do Until light_list(loopc).id = id

        If loopc = light_last Then
            Light_Find = 0
            Exit Function

        End If

        loopc = loopc + 1
    Loop
    
    Light_Find = loopc
    Exit Function
ErrorHandler:
    Light_Find = 0

End Function

Public Function Light_Remove_All() As Boolean

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 1/04/2003
    '
    '*****************************************************************
    Dim Index As Long
    
    For Index = 1 To light_last

        'Make sure it's a legal index
        If Light_Check(Index) Then
            Light_Destroy Index

        End If

    Next Index
    
    Light_Remove_All = True

End Function

Private Sub Light_Destroy(ByVal light_index As Long)

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '
    '**************************************************************
    Dim temp As Light
    
    Light_Erase light_index
    
    light_list(light_index) = temp
    
    'Update array size
    If light_index = light_last Then

        Do Until light_list(light_last).active
            light_last = light_last - 1

            If light_last = 0 Then
                light_count = 0
                Exit Sub

            End If

        Loop
        ReDim Preserve light_list(1 To light_last)

    End If

    light_count = light_count - 1

End Sub

Private Sub Light_Erase(ByVal light_index As Long)

    '***************************************'
    'Author: Juan Martín Sotuyo Dodero
    'Last modified: 3/31/2003
    'Correctly erases a light
    '***************************************'
    Dim min_x As Integer

    Dim min_y As Integer

    Dim max_x As Integer

    Dim max_y As Integer

    Dim X     As Integer

    Dim Y     As Integer
    
    'Set up light borders
    min_x = light_list(light_index).map_x - light_list(light_index).range
    min_y = light_list(light_index).map_y - light_list(light_index).range
    max_x = light_list(light_index).map_x + light_list(light_index).range
    max_y = light_list(light_index).map_y + light_list(light_index).range
    
    'Arrange corners
    'NE
    If InMapBounds(min_x, min_y) Then
        MapData(min_x, min_y).light_value(2) = COLOR_EMPTY
    End If

    'NW
    If InMapBounds(max_x, min_y) Then
        MapData(max_x, min_y).light_value(0) = COLOR_EMPTY
    End If

    'SW
    If InMapBounds(max_x, max_y) Then
        MapData(max_x, max_y).light_value(1) = COLOR_EMPTY
    End If

    'SE
    If InMapBounds(min_x, max_y) Then
        MapData(min_x, max_y).light_value(3) = COLOR_EMPTY
    End If

    'Arrange borders
    'Upper border
    For X = min_x + 1 To max_x - 1

        If InMapBounds(X, min_y) Then
            MapData(X, min_y).light_value(2) = COLOR_EMPTY
        End If

    Next X
    
    'Lower border
    For X = min_x + 1 To max_x - 1

        If InMapBounds(X, max_y) Then
            MapData(X, max_y).light_value(1) = COLOR_EMPTY
            MapData(X, max_y).light_value(3) = COLOR_EMPTY
        End If

    Next X
    
    'Left border
    For Y = min_y + 1 To max_y - 1

        If InMapBounds(min_x, Y) Then
            MapData(min_x, Y).light_value(2) = COLOR_EMPTY
            MapData(min_x, Y).light_value(3) = COLOR_EMPTY
        End If

    Next Y
    
    'Right border
    For Y = min_y + 1 To max_y - 1

        If InMapBounds(max_x, Y) Then
            MapData(max_x, Y).light_value(0) = COLOR_EMPTY
            MapData(max_x, Y).light_value(1) = COLOR_EMPTY
        End If

    Next Y
    
    'Set the inner part of the light
    For X = min_x + 1 To max_x - 1
        For Y = min_y + 1 To max_y - 1

            If InMapBounds(X, Y) Then
                MapData(X, Y).light_value(0) = COLOR_EMPTY
                MapData(X, Y).light_value(1) = COLOR_EMPTY
                MapData(X, Y).light_value(2) = COLOR_EMPTY
                MapData(X, Y).light_value(3) = COLOR_EMPTY
            End If

        Next Y
    Next X

End Sub

