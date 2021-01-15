Attribute VB_Name = "TileEngine"
'MENDUZ DX8 VERSION www.noicoder.com
'RevolucionAo 1.0
'Pablo Mercavides

Option Explicit

Public PreguntaScreen        As String

Public Pregunta              As Boolean

Public PreguntaLocal         As Boolean

Public PreguntaNUM           As Byte

'Map sizes in tiles
Public Const XMaxMapSize     As Byte = 100

Public Const XMinMapSize     As Byte = 1

Public Const YMaxMapSize     As Byte = 100

Public Const YMinMapSize     As Byte = 1

Private Const GrhFogata      As Integer = 1521

' Transparencia de techos
Public RoofsLight()          As Single

''
'Sets a Grh animation to loop indefinitely.
Public Const INFINITE_LOOPS As Integer = -1

Public MaxGrh                As Long

'Encabezado bmp
Type BITMAPFILEHEADER

    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long

End Type

'Info del encabezado del bmp
Type BITMAPINFOHEADER

    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long

End Type

'Posicion en un mapa
Public Type Position

    x As Long
    y As Long

End Type

'Posicion en el Mundo
Public Type WorldPos

    map As Integer
    x As Integer
    y As Integer

End Type

'Contiene info acerca de donde se puede encontrar un grh tamaño y animacion
Public Type GrhData

    sX As Integer
    sY As Integer
    
    FileNum As Long
    
    pixelWidth As Integer
    pixelHeight As Integer
    
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames() As Long 'gs-long
    speed As Single
    active As Boolean
    MiniMap_color As Long
    
    ' Precalculated
    Tx1 As Single
    Tx2 As Single
    Ty1 As Single
    Ty2 As Single

End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type grh

    GrhIndex As Long
    speed As Single
    Started As Long
    Loops As Integer
    Angle As Single
    AnimacionContador As Single
    CantAnim As Long
    Alpha As Byte
    FxIndex As Integer
    
    ' Precalculated
    x As Single
    y As Single

End Type

'Lista de cuerpos
Public Type BodyData

    Walk(E_Heading.NORTH To E_Heading.WEST) As grh
    HeadOffset As Position

End Type

'Lista de cabezas
Public Type HeadData

    Head(E_Heading.NORTH To E_Heading.WEST) As grh

End Type

'Lista de las animaciones de las armas
Type WeaponAnimData

    WeaponWalk(E_Heading.NORTH To E_Heading.WEST) As grh

End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData

    ShieldWalk(E_Heading.NORTH To E_Heading.WEST) As grh

End Type

' Dialog effect
Public Type DialogEffect
    Text As String
    Sube As Single
    Color As RGBA
End Type

'Apariencia del personaje
Public Type Char
    Navegando As Boolean

    UserMinHp As Long
    UserMaxHp As Long
    
    EsEnano As Boolean
    active As Byte
    Heading As E_Heading
    Pos As Position
    
    NowPosX As Integer
    NowPosY As Integer
    
    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    MovArmaEscudo As Boolean
    
    fX As grh
    FxIndex As Integer
    BarTime As Single
    Escribiendo As Boolean
    MaxBarTime As Integer
    BarAccion As Byte
    Particula As Byte
    
    ParticulaTime As Long
    
    Speeding As Single
    EsNpc As Boolean
    EsMascota As Boolean
    
    Donador As Byte
    appear As Byte
    simbolo As Byte
    Idle As Boolean

    Head_Aura As String
    Body_Aura As String
    Arma_Aura As String
    Escudo_Aura As String
    DM_Aura As String
    RM_Aura As String
    Otra_Aura As String

    AuraAngle As Single
    
    FxCount As Integer
    FxList() As grh
        
    particle_count As Integer
    CreandoCant As Integer
    particle_group() As Integer

    TimerM As Byte
    TimerAct As Boolean
    
    TimerI As Single
    TimerIAct As Boolean

    status As Byte
    
    nombre As String
    clan As String
    
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    
    Moving As Boolean
    MoveOffsetX As Single
    MoveOffsetY As Single
    
    Pie As Boolean
    LastStep As Long
    Muerto As Boolean
    Invisible As Boolean
    TimeCreated As Long
    priv As Byte
    
    dialog As String
    dialog_offset_counter_y As Single
    dialog_scroll As Boolean
    AlphaText As Single
    AlphaPJ As Single
    dialog_color As Long
    dialog_life As Byte
    dialog_font_index As Integer
    
    dialogEffects() As DialogEffect
    
    dialogExp As String
    SubeExp As Single
    dialog_Exp_color As RGB
    
    dialogOro As String
    SubeOro As Single
    dialog_Oro_color As RGB
    
    group_index As Integer
    
    clan_index As Integer
    clan_nivel As Byte

End Type

'Info de un objeto
Public Type Obj
    OBJIndex As Integer
    Amount As Integer
End Type

'Tipo de las celdas del mapa
Public Type Light
    Rango As Integer
    Color As RGBA
End Type

Public Type Fantasma

    Activo As Boolean
    Body As grh
    Head As grh
    Arma As grh
    Casco As grh
    Escudo As grh
    Body_Aura As String
    AlphaB As Single
    OffX As Integer
    Offy As Integer
    Heading As Byte

End Type

Public Type MapBlock

    fX As grh
    FxIndex As Byte
    
    FxCount As Integer
    FxList() As grh
    
    Graphic(1 To 4) As grh
    charindex As Integer
    ObjGrh As grh
    GrhBlend As Single
    light_value(3) As RGBA
    
    luz As Light
    particle_group As Integer
    particle_Index As Integer
    
    RenderValue As RVList
    
    NpcIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Integer
    
    Trigger As Integer
    CharFantasma As Fantasma
    ArbolAlphaTimer As Long

End Type

'Info de cada mapa
Public Type MapInfo
    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
    Llueve As Byte
    Nieba As Byte
End Type

Public IniPath                 As String
Public MapPath                 As String

'Bordes del mapa
Public MinXBorder              As Byte
Public MaxXBorder              As Byte
Public MinYBorder              As Byte
Public MaxYBorder              As Byte

'Status del user
Public CurMap                  As Integer 'Mapa actual

Public userindex               As Integer

Public UserMoving              As Byte
Public UserBody                As Integer
Public UserHead                As Integer
Public UserPos                 As Position 'Posicion

Public AddtoUserPos            As Position 'Si se mueve

Public UserCharIndex           As Integer

Public EngineRun               As Boolean

Public fps                     As Long
Public FramesPerSecCounter     As Long
Private fpsLastCheck           As Long

'Tamaño del la vista en Tiles
Public WindowTileWidth        As Integer
Public WindowTileHeight       As Integer
Public HalfWindowTileWidth    As Integer
Public HalfWindowTileHeight   As Integer

'Tamaño del connect
Public HalfConnectTileWidth   As Integer
Public HalfConnectTileHeight  As Integer

'Offset del desde 0,0 del main view
Public MainViewTop            As Integer
Public MainViewLeft           As Integer

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamaño muy grande puede
'volver el engine muy lento
Public TileBufferSizeX        As Integer
Public TileBufferSizeY        As Integer
Public TileBufferPixelOffsetX As Integer
Public TileBufferPixelOffsetY As Integer

'Tamaño de los tiles en pixels
Public Const TilePixelHeight   As Integer = 32
Public Const TilePixelWidth    As Integer = 32

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX   As Single
Public ScrollPixelsPerFrameY   As Single

Public timerElapsedTime           As Single
Public timerTicksPerFrame         As Single
Public engineBaseSpeed            As Single

Public NumBodies               As Integer
Public Numheads                As Integer
Public NumFxs                  As Integer
Public NumChars                As Integer
Public LastChar                As Long
Public NumWeaponAnims          As Integer
Public NumShieldAnims          As Integer

Public MainDestRect           As RECT
Public MainViewRect           As RECT
Public BackBufferRect         As RECT

Public MainViewWidth          As Integer
Public MainViewHeight         As Integer

Public MouseTileX             As Byte
Public MouseTileY             As Byte

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public GrhData()               As GrhData 'Guarda todos los grh
Public BodyData()              As BodyData
Public HeadData()              As HeadData
Public FxData()                As tIndiceFx
Public WeaponAnimData()        As WeaponAnimData
Public ShieldAnimData()        As ShieldAnimData
Public CascoAnimData()         As HeadData
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Mapa?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public MapData()               As MapBlock ' Mapa
Public MapInfo                 As MapInfo ' Info acerca del mapa en uso
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public bRain                   As Boolean 'está raineando?
Public bNieve                  As Boolean 'está nevando?
Public bNiebla                 As Boolean 'Hay niebla?
Public bTecho                  As Boolean 'hay techo?
Public LastMove                As Long ' Tiempo de último paso

Public brstTick                As Long
Private iFrameIndex            As Byte  'Frame actual de la LL
Private llTick                 As Long  'Contador

Public charlist(1 To 10000)    As Char


' Used by GetTextExtentPoint32
Private Type size
    cx As Long
    cy As Long
End Type

'[CODE 001]:MatuX
Public Enum PlayLoop
    plNone = 0
    plLluviain = 1
    plLluviaout = 2
End Enum

'[END]'
'
'       [END]
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'Added by Juan Martín Sotuyo Dodero
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
'Added by Barrin

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

'Text width computation. Needed to center text.
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As size) As Long

Public keysMovementPressedQueue As clsArrayList

Public Sub Init_TileEngine()
    
    On Error GoTo Init_TileEngine_Err
    
    
    'Esto es para el movimiento suave de pjs, para que el pj termine de hacer el movimiento antes de empezar otro
    Set keysMovementPressedQueue = New clsArrayList
    Call keysMovementPressedQueue.Initialize(1, 4)
    
    HalfWindowTileHeight = (frmMain.renderer.ScaleHeight / 32) \ 2
    HalfWindowTileWidth = (frmMain.renderer.ScaleWidth / 32) \ 2
    
    HalfConnectTileHeight = (frmConnect.render.ScaleHeight / 32) \ 2
    HalfConnectTileWidth = (frmConnect.render.ScaleWidth / 32) \ 2
    
    TileBufferSizeX = 7
    TileBufferSizeY = 12
    TileBufferPixelOffsetX = -TileBufferSizeX * TilePixelWidth
    TileBufferPixelOffsetY = -TileBufferSizeY * TilePixelHeight

    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    UserPos.x = 50
    UserPos.y = 50
    
    MinXBorder = XMinMapSize + (frmMain.renderer.ScaleWidth / 64)
    MaxXBorder = XMaxMapSize - (frmMain.renderer.ScaleWidth / 64)
    MinYBorder = YMinMapSize + (frmMain.renderer.ScaleHeight / 64)
    MaxYBorder = YMaxMapSize - (frmMain.renderer.ScaleHeight / 64)
    MinYBorder = MinYBorder

    
    Exit Sub

Init_TileEngine_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine.Init_TileEngine", Erl)
    Resume Next
    
End Sub

Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tX As Byte, ByRef tY As Byte)
    '******************************************
    'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
    '******************************************
    
    On Error GoTo ConvertCPtoTP_Err
    
    
    If viewPortX < 0 Or viewPortX > frmMain.renderer.ScaleWidth Then Exit Sub
    If viewPortY < 0 Or viewPortY > frmMain.renderer.ScaleHeight Then Exit Sub

    tX = UserPos.x + viewPortX \ 32 - frmMain.renderer.ScaleWidth \ 64
    tY = UserPos.y + viewPortY \ 32 - frmMain.renderer.ScaleHeight \ 64

    
    Exit Sub

ConvertCPtoTP_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine.ConvertCPtoTP", Erl)
    Resume Next
    
End Sub

Public Sub InitGrh(ByRef grh As grh, ByVal GrhIndex As Long, Optional ByVal Started As Long = -1, Optional ByVal Loops As Integer = INFINITE_LOOPS)
    '*****************************************************************
    'Sets up a grh. MUST be done before rendering
    '*****************************************************************
    
    On Error GoTo InitGrh_Err

    If GrhIndex = 0 Or GrhIndex > MaxGrh Then Exit Sub
    
     grh.GrhIndex = GrhIndex

    If GrhData(GrhIndex).NumFrames > 1 Then
        If Started >= 0 Then
            grh.Started = Started
        Else
            grh.Started = FrameTime
        End If
        
        grh.Loops = Loops
        grh.speed = GrhData(GrhIndex).speed / GrhData(GrhIndex).NumFrames
    Else
        grh.Started = 0
    End If

    'Precalculate texture coordinates
    With GrhData(grh.GrhIndex)
        If .Tx2 = 0 And .FileNum > 0 Then
            Dim Texture As Direct3DTexture8

            Dim TextureWidth As Long, TextureHeight As Long
            Set Texture = SurfaceDB.GetTexture(.FileNum, TextureWidth, TextureHeight)
        
            .Tx1 = .sX / TextureWidth
            .Tx2 = (.sX + .pixelWidth) / TextureWidth
            .Ty1 = .sY / TextureHeight
            .Ty2 = (.sY + .pixelHeight) / TextureHeight
        End If
    End With

    
    Exit Sub

InitGrh_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine.InitGrh", Erl)
    Resume Next
    
End Sub

Public Sub DoFogataFx()
    
    On Error GoTo DoFogataFx_Err
    

    Dim location As Position
    
    If bFogata Then
        bFogata = HayFogata(location)

        If Not bFogata Then
            ' Call Audio.StopWave(FogataBufferIndex)
            FogataBufferIndex = 0

        End If

    Else
        bFogata = HayFogata(location)

        ' If bFogata And FogataBufferIndex = 0 Then FogataBufferIndex = Audio.PlayWave("fuego.wav", location.x, location.y, LoopStyle.Enabled)
    End If

    
    Exit Sub

DoFogataFx_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine.DoFogataFx", Erl)
    Resume Next
    
End Sub

Private Function EstaPCarea(ByVal charindex As Integer) As Boolean
    
    On Error GoTo EstaPCarea_Err
    

    With charlist(charindex).Pos
        EstaPCarea = .x > UserPos.x - MinXBorder And .x < UserPos.x + MinXBorder And .y > UserPos.y - MinYBorder And .y < UserPos.y + MinYBorder

    End With

    
    Exit Function

EstaPCarea_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine.EstaPCarea", Erl)
    Resume Next
    
End Function

Sub DoPasosFx(ByVal charindex As Integer)
    
    On Error GoTo DoPasosFx_Err
    

    Static TerrenoDePaso As TipoPaso

    Static FileNum       As Integer

    If Not UserNavegando Then

        With charlist(charindex)

            If Not .Muerto And EstaPCarea(charindex) And .priv <= charlist(UserCharIndex).priv Then
                If .Speeding > 1.3 Then
                   
                    Call Sound.Sound_Play(Pasos(CONST_CABALLO).wav(1), , Sound.Calculate_Volume(.Pos.x, .Pos.y), Sound.Calculate_Pan(.Pos.x, .Pos.y))
                    Exit Sub

                End If
           
                .Pie = Not .Pie

                If .Pie Then
                    FileNum = GrhData(MapData(.Pos.x, .Pos.y).Graphic(1).GrhIndex).FileNum
                    TerrenoDePaso = GetTerrenoDePaso(FileNum)
                    Call Sound.Sound_Play(Pasos(TerrenoDePaso).wav(1), , Sound.Calculate_Volume(.Pos.x, .Pos.y), Sound.Calculate_Pan(.Pos.x, .Pos.y))
                    'Call Audio.PlayWave(SND_PASOS3, .Pos.X, .Pos.Y)
                Else
                    Call Sound.Sound_Play(Pasos(TerrenoDePaso).wav(2), , Sound.Calculate_Volume(.Pos.x, .Pos.y), Sound.Calculate_Pan(.Pos.x, .Pos.y))

                End If

            End If

        End With

    Else

        If FxNavega Then
            Call Sound.Sound_Play(SND_NAVEGANDO)

            '  Call Audio.PlayWave(SND_NAVEGANDO, charlist(charindex).Pos.x, charlist(charindex).Pos.y)
        End If

    End If

    
    Exit Sub

DoPasosFx_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine.DoPasosFx", Erl)
    Resume Next
    
End Sub

Private Function GetTerrenoDePaso(ByVal TerrainFileNum As Integer) As TipoPaso
    
    On Error GoTo GetTerrenoDePaso_Err
    

    If (TerrainFileNum >= 6000 And TerrainFileNum <= 6004) Or (TerrainFileNum >= 550 And TerrainFileNum <= 552) Or (TerrainFileNum >= 6018 And TerrainFileNum <= 6020) Then
        GetTerrenoDePaso = CONST_BOSQUE
        Exit Function
    ElseIf (TerrainFileNum >= 7501 And TerrainFileNum <= 7507) Or (TerrainFileNum = 7500 Or TerrainFileNum = 7508 Or TerrainFileNum = 1533 Or TerrainFileNum = 2508) Then
        GetTerrenoDePaso = CONST_DUNGEON
        Exit Function
    ElseIf (TerrainFileNum >= 5000 And TerrainFileNum <= 5004) Then
        GetTerrenoDePaso = CONST_NIEVE
        Exit Function
    ElseIf (TerrainFileNum >= 6018 And TerrainFileNum <= 6021) Or (TerrainFileNum = 186 Or TerrainFileNum = 8007) Then
        GetTerrenoDePaso = CONST_DESIERTO
        Exit Function
    Else
        GetTerrenoDePaso = CONST_PISO

    End If

    
    Exit Function

GetTerrenoDePaso_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine.GetTerrenoDePaso", Erl)
    Resume Next
    
End Function

Sub MoveScreen(ByVal nHeading As E_Heading)
    
    On Error GoTo MoveScreen_Err
    

    '******************************************
    'Starts the screen moving in a direction
    '******************************************
    Dim x  As Integer

    Dim y  As Integer

    Dim tX As Integer

    Dim tY As Integer
    
    'Figure out which way to move
    Select Case nHeading

        Case E_Heading.NORTH
            y = -1
        
        Case E_Heading.EAST
            x = 1
        
        Case E_Heading.south
            y = 1
        
        Case E_Heading.WEST
            x = -1

    End Select
    
    'Fill temp pos
    tX = UserPos.x + x
    tY = UserPos.y + y
    
    'Check to see if its out of bounds
    If tX < MinXBorder Or tX > MaxXBorder Or tY < MinYBorder Or tY > MaxYBorder Then
        Exit Sub
    Else
        'Start moving... MainLoop does the rest
        AddtoUserPos.x = x
        UserPos.x = tX
        AddtoUserPos.y = y
        UserPos.y = tY
        UserMoving = 1
        
        bTecho = HayTecho(UserPos.x, UserPos.y)
        
        LastMove = FrameTime

    End If

    
    Exit Sub

MoveScreen_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine.MoveScreen", Erl)
    Resume Next
    
End Sub

Public Function HayTecho(ByVal x As Integer, ByVal y As Integer) As Boolean
    
    On Error GoTo HayTecho_Err
    
    
    Select Case MapData(x, y).Trigger
        
        Case 1, 2, 4, 6
            HayTecho = True
                
        Case Is > PRIMER_TRIGGER_TECHO
            HayTecho = True
                
        Case Else
            HayTecho = False
        
    End Select
    
    
    Exit Function

HayTecho_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine.HayTecho", Erl)
    Resume Next
    
End Function

Public Function HayFogata(ByRef location As Position) As Boolean
    
    On Error GoTo HayFogata_Err
    

    Dim j As Long

    Dim k As Long
    
    For j = UserPos.x - 13 To UserPos.x + 13
        For k = UserPos.y - 15 To UserPos.y + 15

            If InMapBounds(j, k) Then
                If MapData(j, k).ObjGrh.GrhIndex = GrhFogata Then
                    location.x = j
                    location.y = k
                    
                    HayFogata = True
                    Exit Function

                End If

            End If

        Next k
    Next j

    
    Exit Function

HayFogata_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine.HayFogata", Erl)
    Resume Next
    
End Function

Public Function HayWavAmbiental(ByRef location As Position) As Boolean
    
    On Error GoTo HayWavAmbiental_Err
    

    Dim j As Long

    Dim k As Long
    
    For j = UserPos.x - 13 To UserPos.x + 13
        For k = UserPos.y - 15 To UserPos.y + 15

            If InMapBounds(j, k) Then
                If MapData(j, k).Trigger = 150 Then
                    location.x = j
                    location.y = k
                    
                    '  HayFogata = True
                    '    Exit Function
                End If

            End If

        Next k
    Next j

    
    Exit Function

HayWavAmbiental_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine.HayWavAmbiental", Erl)
    Resume Next
    
End Function

Function NextOpenChar() As Integer
    
    On Error GoTo NextOpenChar_Err
    

    '*****************************************************************
    'Finds next open char slot in CharList
    '*****************************************************************
    Dim loopc As Long

    Dim Dale  As Boolean
    
    loopc = 1

    Do While charlist(loopc).active And Dale
        loopc = loopc + 1
        Dale = (loopc <= UBound(charlist))
    Loop
    
    NextOpenChar = loopc

    
    Exit Function

NextOpenChar_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine.NextOpenChar", Erl)
    Resume Next
    
End Function

''
' Loads grh data using the new file format.
'
' @return   True if the load was successfull, False otherwise.

Function LegalPos(ByVal x As Integer, ByVal y As Integer, ByVal Heading As E_Heading) As Boolean
    
    On Error GoTo LegalPos_Err
    

    '*****************************************************************
    'Checks to see if a tile position is legal
    '*****************************************************************
    'Limites del mapa
    If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
        Exit Function
    End If
    
    '¿Hay un personaje?
    If MapData(x, y).charindex > 0 Then
        With charlist(MapData(x, y).charindex)

            If Not (.Muerto Or (.Invisible And .priv > charlist(UserCharIndex).priv)) Then
                Exit Function
            End If

        End With
    End If
    
    'Tile Bloqueado?
    If (MapData(x, y).Blocked And 2 ^ (Heading - 1)) <> 0 Then
        Exit Function
    End If

    'If Not UserNadando And MapData(x, y).Trigger = 8 Then
    ' If Not UserAvisado Then
    '  Call AddtoRichTextBox(frmMain.RecTxt, "El terreno es rocoso y tu barca podria romperse, solo puedes nadar.", 65, 190, 156, False, False, False)
    ' UserAvisado = True
    ' End If
    'Exit Function

    'Else
    ' If UserNadando And MapData(x, y).Trigger <> 8 Then
    ' Exit Function
    ' End If
    ' LegalPos = True
    ' Exit Function
    '  End If
    
    If UserMontado And MapData(x, y).Trigger > 9 Then
        Exit Function

    End If

    '
    If UserNadando And MapData(x, y).Trigger = 8 Then
        LegalPos = True
        Exit Function

    End If
   
    If UserNavegando <> ((MapData(x, y).Blocked And FLAG_AGUA) <> 0 And (MapData(x, y).Blocked And FLAG_COSTA) = 0) Then
        Exit Function

    End If
    
    If UserNavegando And MapData(x, y).Trigger = 8 And Not UserNadando And Not UserEstado = 1 Then
        If Not UserAvisadoBarca Then
            Call AddtoRichTextBox(frmMain.RecTxt, "¡Atención! El terreno es rocoso y tu barca podria romperse, solo puedes nadar.", 255, 255, 255, True, False, False)
            UserAvisadoBarca = True

        End If

        Exit Function

    End If
    
    'If UserNadando <> HayAgua(x, y) Then
    '    Exit Function
    'End If
    
    UserAvisadoBarca = False
    LegalPos = True
    UserAvisado = False

    
    Exit Function

LegalPos_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine.LegalPos", Erl)
    Resume Next
    
End Function

Function InMapBounds(ByVal x As Integer, ByVal y As Integer) As Boolean
    
    On Error GoTo InMapBounds_Err
    

    '*****************************************************************
    'Checks to see if a tile position is in the maps bounds
    '*****************************************************************
    If x < XMinMapSize Or x > XMaxMapSize Or y < YMinMapSize Or y > YMaxMapSize Then
        Exit Function

    End If
    
    InMapBounds = True

    
    Exit Function

InMapBounds_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine.InMapBounds", Erl)
    Resume Next
    
End Function

Function GetBitmapDimensions(ByVal BmpFile As String, ByRef bmWidth As Long, ByRef bmHeight As Long)
    
    On Error GoTo GetBitmapDimensions_Err
    

    '*****************************************************************
    'Gets the dimensions of a bmp
    '*****************************************************************
    Dim BMHeader    As BITMAPFILEHEADER

    Dim BINFOHeader As BITMAPINFOHEADER
    
    Open BmpFile For Binary Access Read As #1
    
    Get #1, , BMHeader
    Get #1, , BINFOHeader
    
    Close #1
    
    bmWidth = BINFOHeader.biWidth
    bmHeight = BINFOHeader.biHeight

    
    Exit Function

GetBitmapDimensions_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine.GetBitmapDimensions", Erl)
    Resume Next
    
End Function

Public Sub Grh_Render_To_Hdc(ByRef pic As PictureBox, ByVal GrhIndex As Long, ByVal screen_x As Integer, ByVal screen_y As Integer, Optional ByVal Alpha As Integer = False, Optional ByVal ClearColor As Long = &O0)
    
    On Error GoTo Grh_Render_To_Hdc_Err
    

    If GrhIndex = 0 Then Exit Sub

    Static Picture As RECT

    With Picture
        .Left = 0
        .Top = 0

        .Bottom = pic.ScaleHeight
        .Right = pic.ScaleWidth

    End With

    Call DirectDevice.BeginScene
    Call DirectDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, ClearColor, 1#, 0)
    
    Device_Box_Textured_Render GrhIndex, screen_x, screen_y, GrhData(GrhIndex).pixelWidth, GrhData(GrhIndex).pixelHeight, COLOR_WHITE, GrhData(GrhIndex).sX, GrhData(GrhIndex).sY, Alpha, 0

    Call DirectDevice.EndScene
    Call DirectDevice.Present(Picture, ByVal 0, pic.hWnd, ByVal 0)
    
    
    Exit Sub

Grh_Render_To_Hdc_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine.Grh_Render_To_Hdc", Erl)
    Resume Next
    
End Sub

Public Sub Grh_Render_To_HdcSinBorrar(ByRef pic As PictureBox, ByVal GrhIndex As Long, ByVal screen_x As Integer, ByVal screen_y As Integer, Optional ByVal Alpha As Integer = False)
    
    On Error GoTo Grh_Render_To_HdcSinBorrar_Err
    

    If GrhIndex = 0 Then Exit Sub

    Static Picture As RECT

    With Picture
        .Left = 0
        .Top = 0

        .Bottom = pic.ScaleHeight
        .Right = pic.ScaleWidth

    End With

    Call DirectDevice.BeginScene
    
    Device_Box_Textured_Render GrhIndex, screen_x, screen_y, GrhData(GrhIndex).pixelWidth, GrhData(GrhIndex).pixelHeight, COLOR_WHITE, GrhData(GrhIndex).sX, GrhData(GrhIndex).sY, Alpha, 0
                           
    Call DirectDevice.EndScene
    Call DirectDevice.Present(Picture, ByVal 0, pic.hWnd, ByVal 0)
    
    
    Exit Sub

Grh_Render_To_HdcSinBorrar_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine.Grh_Render_To_HdcSinBorrar", Erl)
    Resume Next
    
End Sub


Public Function RenderSounds()
    
    On Error GoTo RenderSounds_Err
    

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 3/30/2008
    'Actualiza todos los sonidos del mapa.
    '**************************************************************
    If bRain Then
        If MapDat.LLUVIA Then
        
            If bTecho Then
                If frmMain.IsPlaying <> PlayLoop.plLluviain Then
                    '  If RainBufferIndex Then _
                    '   Call Audio.StopWave(RainBufferIndex)
                    ' RainBufferIndex = Audio.PlayWave("lluviain.wav", 0, 0, LoopStyle.Enabled)
                    frmMain.IsPlaying = PlayLoop.plLluviain

                End If

            Else

                If frmMain.IsPlaying <> PlayLoop.plLluviaout Then
                
                    ' If RainBufferIndex Then _
                    '   Call Audio.StopWave(RainBufferIndex)
                    '  RainBufferIndex = Audio.PlayWave("lluviaout.wav", 0, 0, LoopStyle.Enabled)
                    frmMain.IsPlaying = PlayLoop.plLluviaout

                End If

            End If

        End If

    End If
    
    DoFogataFx

    
    Exit Function

RenderSounds_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine.RenderSounds", Erl)
    Resume Next
    
End Function

Function HayUserAbajo(ByVal x As Integer, ByVal y As Integer, ByVal GrhIndex As Long) As Boolean
    
    On Error GoTo HayUserAbajo_Err
    

    If GrhIndex > 0 Then
        HayUserAbajo = charlist(UserCharIndex).Pos.x >= x - (GrhData(GrhIndex).TileWidth \ 2) And charlist(UserCharIndex).Pos.x <= x + (GrhData(GrhIndex).TileWidth \ 2) And charlist(UserCharIndex).Pos.y >= y - (GrhData(GrhIndex).TileHeight - 1) And charlist(UserCharIndex).Pos.y <= y

    End If

    
    Exit Function

HayUserAbajo_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine.HayUserAbajo", Erl)
    Resume Next
    
End Function

Private Function GetElapsedTime() As Single
    
    On Error GoTo GetElapsedTime_Err
    

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Gets the time that past since the last call
    '**************************************************************
    Dim Start_Time    As Currency

    Static end_time   As Currency

    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq

    End If
    
    'Get current time
    Call QueryPerformanceCounter(Start_Time)
    
    'Calculate elapsed time
    GetElapsedTime = (Start_Time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)

    
    Exit Function

GetElapsedTime_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine.GetElapsedTime", Erl)
    Resume Next
    
End Function

Private Sub Grh_Create_Mask(ByRef hdcsrc As Long, ByRef MaskDC As Long, ByVal src_x As Integer, ByVal src_y As Integer, ByVal src_width As Integer, ByVal src_height As Integer)
    
    On Error GoTo Grh_Create_Mask_Err
    

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 8/30/2004
    'Creates a Mask hDC, and sets the source hDC to work for trans bliting.
    '**************************************************************
    Dim x          As Integer

    Dim y          As Integer

    Dim TransColor As Long

    Dim ColorKey   As String
    
    'ColorKey = hex(COLOR_KEY)
    
    'Check if it has an alpha component
    'If Len(ColorKey) > 6 Then
    'get rid of alpha
    '    ColorKey = "&H" & Right$(ColorKey, 6)
    'End If
    'piluex prueba
    'TransColor = Val(ColorKey)
    ColorKey = "0"
    TransColor = &H0

    'Make it a mask (set background to black and foreground to white)
    'And set the sprite's background white
    For y = src_y To src_height + src_y
        For x = src_x To src_width + src_x

            If GetPixel(hdcsrc, x, y) = TransColor Then
                SetPixel MaskDC, x, y, vbWhite
                SetPixel hdcsrc, x, y, vbBlack
            Else
                SetPixel MaskDC, x, y, vbBlack

            End If

        Next x
    Next y

    
    Exit Sub

Grh_Create_Mask_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine.Grh_Create_Mask", Erl)
    Resume Next
    
End Sub

Public Function Convert_Tile_To_View_X(ByVal x As Integer) As Integer
    '**************************************************************
    'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
    'Last Modify Date: 10/07/2002
    'Convert tile position into position in view area
    '**************************************************************
    'If engine_windowed Then
    
    On Error GoTo Convert_Tile_To_View_X_Err
    
    Convert_Tile_To_View_X = ((x - 1) * 32)

    ' Else
    '  Convert_Tile_To_View_X = view_screen_left + ((x - 1) * base_tile_size)
    '  End If
    
    Exit Function

Convert_Tile_To_View_X_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine.Convert_Tile_To_View_X", Erl)
    Resume Next
    
End Function

Public Function Convert_Tile_To_View_Y(ByVal y As Integer) As Integer
    '**************************************************************
    'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
    'Last Modify Date: 10/07/2002
    'Convert tile position into position in view area
    '**************************************************************
    ' If engine_windowed Then
    
    On Error GoTo Convert_Tile_To_View_Y_Err
    
    Convert_Tile_To_View_Y = ((y - 1) * 32)

    'Else
    '   Convert_Tile_To_View_Y = view_screen_top + ((y - 1) * base_tile_size)
    'End If
    
    Exit Function

Convert_Tile_To_View_Y_Err:
    Call RegistrarError(Err.number, Err.Description, "TileEngine.Convert_Tile_To_View_Y", Erl)
    Resume Next
    
End Function


