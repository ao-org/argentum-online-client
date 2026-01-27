Attribute VB_Name = "TileEngine"
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
Public TileEngineHelper                        As New clsTileEngineHelper
Public PreguntaScreen                          As String
Public Pregunta                                As Boolean
Public PreguntaLocal                           As Boolean
Public PreguntaNUM                             As Byte
Public Const DwarvenExoesqueletonBody          As Integer = 4993
'Map sizes in tiles
Public Const XMaxMapSize                       As Byte = 100
Public Const XMinMapSize                       As Byte = 1
Public Const YMaxMapSize                       As Byte = 100
Public Const YMinMapSize                       As Byte = 1
Private Const GrhFogata                        As Integer = 1521
' Transparencia de techos
Public RoofsLight()                            As Single
''
'Sets a Grh animation to loop indefinitely.
Public Const INFINITE_LOOPS                    As Integer = -1
Public MaxGrh                                  As Long
'PescaEspecial VARS
Public Const MAX_INTENTOS                      As Byte = 5
Public intentosPesca(1 To MAX_INTENTOS)        As Byte
Public PuedeIntentar                           As Boolean
Public PescandoEspecial                        As Boolean
Public MostrarTutorial                         As Boolean
Public Const BarWidth                          As Long = 199
Public PosicionBarra                           As Single
Public ContadorIntentosPescaEspecial_Acertados As Long
Public ContadorIntentosPescaEspecial_Fallados  As Long
Public startTimePezEspecial                    As Long
Public LastTimePezEspecial                     As Long
Public DireccionBarra                          As Single
Public Const VelocidadBarra                    As Single = 1
Public Const GRH_BARRA_PESCA                   As Long = 60666
Public Const GRH_CURSOR_PESCA                  As Long = 60667
Public Const GRH_CIRCULO_VERDE                 As Long = 38367
Public Const GRH_CIRCULO_ROJO                  As Long = 38366
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

'Contiene info acerca de donde se puede encontrar un grh tamaï¿½o y animacion
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
    ' Precalculated
    Tx1 As Single
    Tx2 As Single
    Ty1 As Single
    Ty2 As Single
End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh
    GrhIndex As Long
    speed As Single
    started As Long
    Loops As Integer
    angle As Single
    AnimacionContador As Single
    CantAnim As Long
    alpha As Byte
    FxIndex As Integer
    ' Precalculated
    x As Single
    y As Single
End Type

'Lista de cuerpos
Public Type BodyData
    Walk(E_Heading.NORTH To E_Heading.WEST) As Grh
    HeadOffset As Position
    BodyOffset As Position
    ShadowOffset As Position
    BodyIndex As Integer
    IdleBody As Integer
    AnimateOnIdle As Integer
End Type

'Lista de cabezas
Public Type HeadData
    Head(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

' Dialog effect
Public Type DialogEffect
    text As String
    start As Long
    color As RGBA
    offset As Position
    duration As Integer
    Animated As Boolean
End Type

Public Enum eTipoUsuario
    User = 0
    aventurero
    heroe
    Legend
End Enum

Public Type tMascota
    PosX As Double
    PosY As Double
    delta As Double
    Body(1 To 8) As Grh
    Heading As Long
    last_time As Double
    dialog As String
    dialog_life As Single
    Fx As Grh
    color(3) As RGBA
    visible As Boolean
End Type

Public mascota As tMascota

Public Enum e_CharValue
    eDontBlockTile = 1
End Enum

'Apariencia del personaje
Public Type Char
    Navegando As Boolean
    UserMinHp As Long
    UserMaxHp As Long
    Shield As Long
    UserMinMAN As Long
    UserMaxMAN As Long
    EsEnano As Boolean
    active As Byte
    Heading As E_Heading
    Pos As Position
    NowPosX As Integer
    NowPosY As Integer
    IHead As Integer
    iBody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Cart As BodyData
    Backpack As BodyData
    tmpBackPack As Integer
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    MovArmaEscudo As Boolean
    HasCart As Boolean
    HasBackpack As Boolean
    AnimatingBody As Integer
    ActiveAnimation As tAnimationPlaybackState
    BarTime As Single
    MaxBarTime As Integer
    BarAccion As Byte
    Particula As Integer
    ParticulaTime As Long
    Speeding As Single
    EsNpc As Boolean
    EsMascota As Boolean
    bodyOnWater As Integer
    appear As Byte
    simbolo As Byte
    Idle As Boolean
    Meditating As Boolean
    Head_Aura As String
    Body_Aura As String
    Arma_Aura As String
    Escudo_Aura As String
    DM_Aura As String
    RM_Aura As String
    Otra_Aura As String
    AuraAngle As Single
    FxCount As Integer
    FxList() As Grh
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
    DialogEffects() As DialogEffect
    group_index As Integer
    clan_index As Integer
    clan_nivel As Byte
    tipoUsuario As eTipoUsuario
    Team As Byte
    banderaIndex As Byte
    AnimAtaque1 As Integer
    TranslationActive As Boolean
    TranslationStartTime As Long
    TranslationTime As Long
    DontBlockTile As Boolean
End Type

'Info de un objeto
Public Type Obj
    ObjIndex As Integer
    Amount As Integer
    ElementalTags As Long
End Type

'Tipo de las celdas del mapa
Public Type Light
    Rango As Integer
    color As RGBA
End Type

Public Type Fantasma
    Activo As Boolean
    Body As Grh
    Head As Grh
    Arma As Grh
    Casco As Grh
    Escudo As Grh
    Body_Aura As String
    AlphaB As Single
    OffX As Integer
    Offy As Integer
    Heading As Byte
End Type

Public Type MapZone
    OcultarNombre As Boolean
    NumMapa As Integer
    Musica As Integer
    x1 As Byte
    x2 As Byte
    y1 As Byte
    y2 As Byte
End Type

Public Type MapBlock
    Fx As Grh
    FxIndex As Byte
    FxCount As Integer
    FxList() As Grh
    Graphic(1 To 4) As Grh
    charindex As Integer
    ObjGrh As Grh
    GrhBlend As Single
    light_value(3) As RGBA
    luz As Light
    particle_group As Integer
    particle_Index As Integer
    DialogEffects() As DialogEffect
    NpcIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Integer
    Trigger As Integer
    CharFantasma As Fantasma
    ArbolAlphaTimer As Long
    zone As MapZone
    Trap As Grh
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

Public IniPath                As String
Public MapPath                As String
'Bordes del mapa
Public MinXBorder             As Byte
Public MaxXBorder             As Byte
Public MinYBorder             As Byte
Public MaxYBorder             As Byte
'Status del user
Public CurMap                 As Integer 'Mapa actual
Public userIndex              As Integer
Public UserMoving             As Boolean
Public CharindexSeguido       As Integer
Public UserBody               As Integer
Public UserHead               As Integer
Public UserPos                As Position 'Posicion
Public AddtoUserPos           As Position 'Si se mueve
Public UserCharIndex          As Integer
Public EngineRun              As Boolean
Public fps                    As Long
Public FramesPerSecCounter    As Long
Private fpsLastCheck          As Long
'Tamaï¿½o del la vista en Tiles
Public WindowTileWidth        As Integer
Public WindowTileHeight       As Integer
Public HalfWindowTileWidth    As Integer
Public HalfWindowTileHeight   As Integer
'Tamaï¿½o del connect
Public HalfConnectTileWidth   As Integer
Public HalfConnectTileHeight  As Integer
'Offset del desde 0,0 del main view
Public MainViewTop            As Integer
Public MainViewLeft           As Integer
'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamaï¿½o muy grande puede
'volver el engine muy lento
Public TileBufferSizeX        As Integer
Public TileBufferSizeY        As Integer
Public TileBufferPixelOffsetX As Integer
Public TileBufferPixelOffsetY As Integer
'Tamaï¿½o de los tiles en pixels
Public Const TilePixelHeight  As Integer = 32
Public Const TilePixelWidth   As Integer = 32
'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX  As Single
Public ScrollPixelsPerFrameY  As Single
Public timerElapsedTime       As Single
Public timerTicksPerFrame     As Single
Public engineBaseSpeed        As Single
Public UpdateLights           As Boolean
Public NumBodies              As Integer
Public Numheads               As Integer
Public NumFxs                 As Integer
Public NumChars               As Integer
Public LastChar               As Long
Public NumWeaponAnims         As Integer
Public NumShieldAnims         As Integer
Public MainDestRect           As Rect
Public MainViewRect           As Rect
Public BackBufferRect         As Rect
Public MainViewWidth          As Integer
Public MainViewHeight         As Integer
Public MouseTileX             As Byte
Public MouseTileY             As Byte
'ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½Graficosï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?
Public GrhData()              As GrhData 'Guarda todos los grh
Public BodyData()             As BodyData
Public HeadData()             As HeadData
Public FxData()               As tIndiceFx
Public FxToAnimationMap()     As Integer
Public ComposedFxData()       As tComposedAnimation
Public WeaponAnimData()       As WeaponAnimData
Public ShieldAnimData()       As ShieldAnimData
Public CascoAnimData()        As HeadData
'ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?
'ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½Mapa?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?
Public MapData()              As MapBlock ' Mapa
Public MapInfo                As MapInfo ' Info acerca del mapa en uso
Public Zonas()                As MapZone
'ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?
Public bRain                  As Boolean 'estï¿½ raineando?
Public bNieve                 As Boolean 'estï¿½ nevando?
Public bNiebla                As Boolean 'Hay niebla?
Public bTecho                 As Boolean 'hay techo?
Public lastMove               As Long ' Tiempo de ï¿½ltimo paso
Public brstTick               As Long
Private iFrameIndex           As Byte  'Frame actual de la LL
Private llTick                As Long  'Contador
Public charlist(1 To 10000)   As Char

' Used by GetTextExtentPoint32
Private Type size
    cx As Long
    cy As Long
End Type

'ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?ï¿½?
Private Declare Function BitBlt _
                Lib "gdi32" (ByVal hDestDC As Long, _
                             ByVal x As Long, _
                             ByVal y As Long, _
                             ByVal nWidth As Long, _
                             ByVal nHeight As Long, _
                             ByVal hSrcDC As Long, _
                             ByVal xSrc As Long, _
                             ByVal ySrc As Long, _
                             ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'Added by Juan Martï¿½n Sotuyo Dodero
Private Declare Function StretchBlt _
                Lib "gdi32" (ByVal hDestDC As Long, _
                             ByVal x As Long, _
                             ByVal y As Long, _
                             ByVal nWidth As Long, _
                             ByVal nHeight As Long, _
                             ByVal hSrcDC As Long, _
                             ByVal xSrc As Long, _
                             ByVal ySrc As Long, _
                             ByVal nSrcWidth As Long, _
                             ByVal nSrcHeight As Long, _
                             ByVal dwRop As Long) As Long
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
    TileBufferSizeX = 14
    TileBufferSizeY = 18
    TileBufferPixelOffsetX = -TileBufferSizeX * TilePixelWidth
    TileBufferPixelOffsetY = -TileBufferSizeY * TilePixelHeight
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    MinXBorder = XMinMapSize + (frmMain.renderer.ScaleWidth \ 64)
    MaxXBorder = XMaxMapSize - (frmMain.renderer.ScaleWidth \ 64)
    MinYBorder = YMinMapSize + (frmMain.renderer.ScaleHeight \ 64)
    MaxYBorder = YMaxMapSize - (frmMain.renderer.ScaleHeight \ 64)
    MinYBorder = MinYBorder
    Exit Sub
Init_TileEngine_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine.Init_TileEngine", Erl)
    Resume Next
End Sub

Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef tX As Byte, ByRef tY As Byte)
    On Error GoTo ConvertCPtoTP_Err
    Dim ltx As Long: Dim lty As Long
    If viewPortX < 0 Or viewPortX > frmMain.renderer.ScaleWidth Then Exit Sub
    If viewPortY < 0 Or viewPortY > frmMain.renderer.ScaleHeight Then Exit Sub
    ltx = UserPos.x + viewPortX \ 32 - frmMain.renderer.ScaleWidth \ 64
    lty = UserPos.y + viewPortY \ 32 - frmMain.renderer.ScaleHeight \ 64
    tX = max(0, ltx)
    tY = max(0, lty)
    Exit Sub
ConvertCPtoTP_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine.ConvertCPtoTP", Erl)
    Resume Next
End Sub

Public Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Long, Optional ByVal started As Long = -1, Optional ByVal Loops As Integer = INFINITE_LOOPS)
    On Error GoTo InitGrh_Err
    If GrhIndex = 0 Or GrhIndex > MaxGrh Then Exit Sub
    Grh.GrhIndex = GrhIndex
    If GrhData(GrhIndex).NumFrames > 1 Then
        If started >= 0 Then
            Grh.started = started
        Else
            Grh.started = FrameTime
        End If
        Grh.Loops = Loops
        Grh.speed = GrhData(GrhIndex).speed / GrhData(GrhIndex).NumFrames
        If Grh.speed = 0 Then
            Grh.speed = 1
        End If
    Else
        Grh.started = 0
        Grh.speed = 1
    End If
    'Precalculate texture coordinates
    With GrhData(Grh.GrhIndex)
        If .Tx2 = 0 And .FileNum > 0 Then
            Dim Texture      As Direct3DTexture8
            Dim TextureWidth As Long, TextureHeight As Long
            Set Texture = SurfaceDB.GetTexture(.FileNum, TextureWidth, TextureHeight)
            Debug.Assert Not Texture Is Nothing
            Debug.Assert TextureWidth > 0 And TextureHeight > 0
            .Tx1 = (.sX + 0.25) / TextureWidth
            .Tx2 = (.sX + .pixelWidth) / TextureWidth
            .Ty1 = (.sY + 0.25) / TextureHeight
            .Ty2 = (.sY + .pixelHeight) / TextureHeight
        End If
    End With
    Exit Sub
InitGrh_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine.InitGrh", Erl)
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
    Call RegistrarError(Err.Number, Err.Description, "TileEngine.DoFogataFx", Erl)
    Resume Next
End Sub

Private Function EstaPCarea(ByVal charindex As Integer) As Boolean
    On Error GoTo EstaPCarea_Err
    With charlist(charindex).Pos
        EstaPCarea = .x > UserPos.x - MinXBorder And .x < UserPos.x + MinXBorder And .y > UserPos.y - MinYBorder And .y < UserPos.y + MinYBorder
    End With
    Exit Function
EstaPCarea_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine.EstaPCarea", Erl)
    Resume Next
End Function

Sub DoPasosFx(ByVal charindex As Integer)
    On Error GoTo DoPasosFx_Err
    With charlist(charindex)
        ' Mascots don't make noise
        If .EsMascota Then Exit Sub
        If EstaPCarea(charindex) And Not .Muerto And .priv <= charlist(UserCharIndex).priv And charlist(UserCharIndex).Muerto = False Then
            .Pie = Not .Pie
            Dim StepIndex     As Byte: StepIndex = IIf(.Pie, 1, 2)
            Dim TerrenoDePaso As Byte
            If .Navegando Then
                TerrenoDePaso = CONST_AGUA
            ElseIf MapData(.Pos.x, .Pos.y).Graphic(1).GrhIndex > 0 Then
                Dim FileNum As Long: FileNum = GrhData(MapData(.Pos.x, .Pos.y).Graphic(1).GrhIndex).FileNum
                TerrenoDePaso = GetTerrenoDePaso(FileNum, MapData(.Pos.x, .Pos.y).Graphic(2).GrhIndex)
                If .Speeding > 1.2 And TerrenoDePaso = CONST_PISO Then
                    TerrenoDePaso = CONST_CABALLO
                End If
            Else
                TerrenoDePaso = CONST_PISO
            End If
            Dim steps_vol As Long
            Dim steps_pan As Long
            steps_vol = ao20audio.ComputeVolumeAtPos(eFxSteps, .Pos)
            steps_pan = ao20audio.ComputeCharFxPan(.Pos)
            Dim SndIndex As Integer: SndIndex = Pasos(TerrenoDePaso).wav(StepIndex)
            Dim SndLabel As String: SndLabel = "pasos" & charindex & "_" & StepIndex
            Call ao20audio.StopWav(SndIndex, SndLabel)
            Call ao20audio.PlayFx(Pasos(TerrenoDePaso).wav(StepIndex), False, steps_vol, steps_pan, "pasos" & charindex & "_" & StepIndex, eFxSteps)
        End If
    End With
    Exit Sub
DoPasosFx_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine.DoPasosFx", Erl)
    Resume Next
End Sub

Sub DoPasosInvi(ByVal Grh As Long, ByVal Grh2 As Long, ByVal distancia As Byte, ByVal balance As Integer, ByVal step As Boolean)
    On Error GoTo DoPasosInvi_Err
    Static TerrenoDePaso As TipoPaso
    Dim FileNum          As Long
    If Grh > 0 Then
        FileNum = GrhData(Grh).FileNum
        TerrenoDePaso = GetTerrenoDePaso(FileNum, Grh2)
        Dim steps_vol2 As Long
        Dim steps_pan2 As Long
        steps_vol2 = ao20audio.ComputeVolumeByDistance(eFxSteps, distancia)
        steps_pan2 = ao20audio.ComputeCharFxPanByDistance(distancia, balance)
        Call ao20audio.PlayFx(Pasos(TerrenoDePaso).wav(IIf(step, 1, 2)), False, steps_vol2, steps_pan2, , eFxSteps)
    End If
    Exit Sub
DoPasosInvi_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine.DoPasosInvi", Erl)
    Resume Next
End Sub

Public Function GetTerrenoDePaso(ByVal TerrainFileNum As Integer, ByVal Layer2Grh As Long) As TipoPaso
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
    ElseIf TerrainFileNum = 20 And Layer2Grh = 0 Then
        GetTerrenoDePaso = CONST_AGUA
        Exit Function
    Else
        GetTerrenoDePaso = CONST_PISO
    End If
    Exit Function
GetTerrenoDePaso_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine.GetTerrenoDePaso", Erl)
    Resume Next
End Function

Sub MoveScreen(ByVal nHeading As E_Heading)
    On Error GoTo MoveScreen_Err
    'Starts the screen moving in a direction
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
        UserMoving = True
        UpdatePlayerRoof
        lastMove = FrameTime
    End If
    Exit Sub
MoveScreen_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine.MoveScreen", Erl)
    Resume Next
End Sub

Public Function NearRoof(ByVal x As Integer, ByVal y As Integer) As eTrigger
    On Error GoTo NearRoof_Err
    Dim lX As Integer, lY As Integer
    For lY = y - 1 To y + 1
        For lX = x - 1 To x + 1
            If lX >= XMinMapSize And lX <= XMaxMapSize Then
                If lY >= YMinMapSize And lY <= YMaxMapSize Then
                    If HayTecho(lX, lY) Then
                        NearRoof = MapData(lX, lY).Trigger
                        Exit Function
                    End If
                End If
            End If
        Next
    Next
    Exit Function
NearRoof_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine.NearRoof", Erl)
    Resume Next
End Function

Public Function HayTecho(ByVal x As Integer, ByVal y As Integer) As Boolean
    On Error GoTo HayTecho_Err
    With MapData(x, y)
        HayTecho = .Trigger >= PRIMER_TRIGGER_TECHO Or .Trigger = eTrigger.BAJOTECHO Or .Trigger = eTrigger.ZONASEGURA Or .Trigger = eTrigger.NADOBAJOTECHO
    End With
    Exit Function
HayTecho_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine.HayTecho", Erl)
    Resume Next
End Function

Public Function HayFogata(ByRef location As Position) As Boolean
    On Error GoTo HayFogata_Err
    Dim J As Long
    Dim k As Long
    For J = UserPos.x - 13 To UserPos.x + 13
        For k = UserPos.y - 15 To UserPos.y + 15
            If InMapBounds(J, k) Then
                If MapData(J, k).ObjGrh.GrhIndex = GrhFogata Then
                    location.x = J
                    location.y = k
                    HayFogata = True
                    Exit Function
                End If
            End If
        Next k
    Next J
    Exit Function
HayFogata_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine.HayFogata", Erl)
    Resume Next
End Function

Public Function HayWavAmbiental(ByRef location As Position) As Boolean
    On Error GoTo HayWavAmbiental_Err
    Dim J As Long
    Dim k As Long
    For J = UserPos.x - 13 To UserPos.x + 13
        For k = UserPos.y - 15 To UserPos.y + 15
            If InMapBounds(J, k) Then
                If MapData(J, k).Trigger = 150 Then
                    location.x = J
                    location.y = k
                    '  HayFogata = True
                    '    Exit Function
                End If
            End If
        Next k
    Next J
    Exit Function
HayWavAmbiental_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine.HayWavAmbiental", Erl)
    Resume Next
End Function

Function NextOpenChar() As Integer
    On Error GoTo NextOpenChar_Err
    'Finds next open char slot in CharList
    Dim loopC As Long
    Dim Dale  As Boolean
    loopC = 1
    Do While charlist(loopC).active And Dale
        loopC = loopC + 1
        Dale = (loopC <= UBound(charlist))
    Loop
    NextOpenChar = loopC
    Exit Function
NextOpenChar_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine.NextOpenChar", Erl)
    Resume Next
End Function

''
' Loads grh data using the new file format.
'
' @return   True if the load was successfull, False otherwise.
Function LegalPos(ByVal x As Integer, ByVal y As Integer, ByVal Heading As E_Heading) As Boolean
    On Error GoTo LegalPos_Err
    'Checks to see if a tile position is legal
    'Limites del mapa
    If x < MinXBorder Or x > MaxXBorder Or y < MinYBorder Or y > MaxYBorder Then
        Exit Function
    End If
    'ï¿½Hay un personaje?
    If MapData(x, y).charindex > 0 Then
        With charlist(MapData(x, y).charindex)
            If Not (.Muerto Or (.Invisible And .priv > charlist(UserCharIndex).priv) Or .DontBlockTile) Then
                Exit Function
            End If
        End With
    End If
    If (MapData(x, y).Blocked And 2 ^ (Heading - 1)) <> 0 Then Exit Function
    If UserMontado And MapData(x, y).Trigger > 9 Then
        Exit Function
    End If
    If MapData(x, y).Trigger = WORKERONLY Then
        If Not UserStats.Clase = Trabajador Then Exit Function
    End If
    If UserNadando And MapData(x, y).Trigger = DETALLEAGUA Then
        LegalPos = True
        Exit Function
    End If
    If UserNadando And MapData(x, y).Trigger = NADOBAJOTECHO Then
        LegalPos = True
        Exit Function
    End If
    '0 <>
    If UserNavegando <> ((MapData(x, y).Blocked And FLAG_AGUA) <> 0 And (MapData(x, y).Blocked And FLAG_COSTA) = 0) And MapData(x, y).Trigger <> eTrigger.VALIDOPUENTE Then
        Exit Function
    End If
    If UserNadando And Not (MapData(x, y).Trigger = eTrigger.DETALLEAGUA Or MapData(x, y).Trigger = eTrigger.NADOCOMBINADO Or MapData(x, y).Trigger = eTrigger.VALIDONADO Or _
            MapData(x, y).Trigger = eTrigger.NADOBAJOTECHO) Then
        LegalPos = False
        Exit Function
    End If
    If UserNavegando And MapData(x, y).Trigger = 8 And Not UserNadando And Not UserStats.estado = 1 Then
        If Not UserAvisadoBarca Then
            Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_AGUA_POCO_PROFUNDA_BARCA_PODRIA_ROMPERSE"), 255, 255, 255, True, False, False)
            UserAvisadoBarca = True
        End If
        Exit Function
    End If
    If UserNavegando And MapData(x, y).Trigger = 11 And Not UserNadando And Not UserStats.estado = 1 Then
        If Not UserAvisadoBarca Then
            Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_TERRENO_ROCOSO_BARCA_PODRIA_ROMPERSE"), 255, 255, 255, True, False, False)
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
    Call RegistrarError(Err.Number, Err.Description, "TileEngine.LegalPos", Erl)
    Resume Next
End Function

Function InMapBounds(ByVal x As Integer, ByVal y As Integer) As Boolean
    On Error GoTo InMapBounds_Err
    'Checks to see if a tile position is in the maps bounds
    If x < XMinMapSize Or x > XMaxMapSize Or y < YMinMapSize Or y > YMaxMapSize Then
        Exit Function
    End If
    InMapBounds = True
    Exit Function
InMapBounds_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine.InMapBounds", Erl)
    Resume Next
End Function

Function GetBitmapDimensions(ByVal BmpFile As String, ByRef bmWidth As Long, ByRef bmHeight As Long)
    On Error GoTo GetBitmapDimensions_Err
    'Gets the dimensions of a bmp
    Dim BMHeader    As BITMAPFILEHEADER
    Dim BINFOHeader As BITMAPINFOHEADER
    Dim fh          As Integer
    fh = FreeFile
    Open BmpFile For Binary Access Read As #fh
    Get #fh, , BMHeader
    Get #fh, , BINFOHeader
    Close #fh
    bmWidth = BINFOHeader.biWidth
    bmHeight = BINFOHeader.biHeight
    Exit Function
GetBitmapDimensions_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine.GetBitmapDimensions", Erl)
    Resume Next
End Function

Public Sub Grh_Render_To_Hdc(ByRef pic As PictureBox, _
                             ByVal GrhIndex As Long, _
                             ByVal screen_x As Integer, _
                             ByVal screen_y As Integer, _
                             Optional ByVal alpha As Integer = False, _
                             Optional ByVal ClearColor As Long = &O0)
    On Error GoTo Grh_Render_To_Hdc_Err
    If GrhIndex = 0 Then Exit Sub
    Static Picture As Rect
    With Picture
        .Left = 0
        .Top = 0
        .Bottom = pic.ScaleHeight
        .Right = pic.ScaleWidth
    End With
    Call DirectDevice.BeginScene
    Call DirectDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, ClearColor, 1#, 0)
    Device_Box_Textured_Render GrhIndex, screen_x, screen_y, GrhData(GrhIndex).pixelWidth, GrhData(GrhIndex).pixelHeight, COLOR_WHITE, GrhData(GrhIndex).sX, GrhData( _
            GrhIndex).sY, alpha, 0
    Call DirectDevice.EndScene
    Call DirectDevice.Present(Picture, ByVal 0, pic.hWnd, ByVal 0)
    Exit Sub
Grh_Render_To_Hdc_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine.Grh_Render_To_Hdc", Erl)
    Resume Next
End Sub

Public Sub Grh_Render_To_HdcSinBorrar(ByRef pic As PictureBox, _
                                      ByVal GrhIndex As Long, _
                                      ByVal screen_x As Integer, _
                                      ByVal screen_y As Integer, _
                                      Optional ByVal alpha As Integer = False)
    On Error GoTo Grh_Render_To_HdcSinBorrar_Err
    If GrhIndex = 0 Then Exit Sub
    Static Picture As Rect
    With Picture
        .Left = 0
        .Top = 0
        .Bottom = pic.ScaleHeight
        .Right = pic.ScaleWidth
    End With
    Call DirectDevice.BeginScene
    Device_Box_Textured_Render GrhIndex, screen_x, screen_y, GrhData(GrhIndex).pixelWidth, GrhData(GrhIndex).pixelHeight, COLOR_WHITE, GrhData(GrhIndex).sX, GrhData( _
            GrhIndex).sY, alpha, 0
    Call DirectDevice.EndScene
    Call DirectDevice.Present(Picture, ByVal 0, pic.hWnd, ByVal 0)
    Exit Sub
Grh_Render_To_HdcSinBorrar_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine.Grh_Render_To_HdcSinBorrar", Erl)
    Resume Next
End Sub

Function HayUserAbajo(ByVal x As Integer, ByVal y As Integer, ByVal GrhIndex As Long) As Boolean
    On Error GoTo HayUserAbajo_Err
    If GrhIndex > 0 Then
        HayUserAbajo = charlist(UserCharIndex).Pos.x >= x - (GrhData(GrhIndex).TileWidth \ 2) And charlist(UserCharIndex).Pos.x <= x + (GrhData(GrhIndex).TileWidth \ 2) And _
                charlist(UserCharIndex).Pos.y >= y - (GrhData(GrhIndex).TileHeight - 1) And charlist(UserCharIndex).Pos.y <= y
    End If
    Exit Function
HayUserAbajo_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine.HayUserAbajo", Erl)
    Resume Next
End Function

Public Function GetElapsedTime() As Single
    On Error GoTo GetElapsedTime_Err
    'Gets the time that past since the last call
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
    Call RegistrarError(Err.Number, Err.Description, "TileEngine.GetElapsedTime", Erl)
    Resume Next
End Function

Private Sub Grh_Create_Mask(ByRef hdcsrc As Long, ByRef MaskDC As Long, ByVal src_x As Integer, ByVal src_y As Integer, ByVal src_width As Integer, ByVal src_height As Integer)
    On Error GoTo Grh_Create_Mask_Err
    'Creates a Mask hDC, and sets the source hDC to work for trans bliting.
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
    Call RegistrarError(Err.Number, Err.Description, "TileEngine.Grh_Create_Mask", Erl)
    Resume Next
End Sub

Public Function Convert_Tile_To_View_X(ByVal x As Integer) As Integer
    'Convert tile position into position in view area
    'If engine_windowed Then
    On Error GoTo Convert_Tile_To_View_X_Err
    Convert_Tile_To_View_X = ((x - 1) * 32)
    ' Else
    '  Convert_Tile_To_View_X = view_screen_left + ((x - 1) * base_tile_size)
    '  End If
    Exit Function
Convert_Tile_To_View_X_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine.Convert_Tile_To_View_X", Erl)
    Resume Next
End Function

Public Function Convert_Tile_To_View_Y(ByVal y As Integer) As Integer
    'Convert tile position into position in view area
    ' If engine_windowed Then
    On Error GoTo Convert_Tile_To_View_Y_Err
    Convert_Tile_To_View_Y = ((y - 1) * 32)
    'Else
    '   Convert_Tile_To_View_Y = view_screen_top + ((y - 1) * base_tile_size)
    'End If
    Exit Function
Convert_Tile_To_View_Y_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine.Convert_Tile_To_View_Y", Erl)
    Resume Next
End Function

Public Function GetTerrainHeight(x As Byte, y As Byte) As Integer
    With MapData(x, y)
        Select Case .Graphic(2).GrhIndex
            Case 12682
                GetTerrainHeight = 5
            Case 12683
                GetTerrainHeight = 10
            Case 12684
                GetTerrainHeight = 14
            Case 12685
                GetTerrainHeight = 14
            Case 12686
                GetTerrainHeight = 14
            Case 12687
                GetTerrainHeight = 14
            Case 12688
                GetTerrainHeight = 10
            Case 12689
                GetTerrainHeight = 5
            Case 12692
                GetTerrainHeight = 5
            Case 12693
                GetTerrainHeight = 10
            Case 12694
                GetTerrainHeight = 14
            Case 12695
                GetTerrainHeight = 14
            Case 12696
                GetTerrainHeight = 14
            Case 12697
                GetTerrainHeight = 14
            Case 12698
                GetTerrainHeight = 10
            Case 12699
                GetTerrainHeight = 5
            Case Else
                GetTerrainHeight = 0
        End Select
    End With
End Function

Public Sub ConvertToMinimapPosition(ByRef x As Single, ByRef y As Single, ByVal MarkerWidth As Single, ByVal MarkerHeight As Single)
    'this hard to understand, minimaps only shows playable area, with the borders cut, so instead of showing every tile on the map we have
    '100x100 pixels for ~78x82 tiles
    x = (x - HalfWindowTileWidth - 2) * (100 / (100 - 2 * HalfWindowTileWidth - 4)) - MarkerWidth \ 2 - 1
    y = (y - HalfWindowTileHeight - 1) * (100 / (100 - 2 * HalfWindowTileHeight - 2)) - MarkerHeight \ 2 - 1
End Sub

Public Sub UpdatePlayerRoof()
    Dim WasUnderRoof As Boolean: WasUnderRoof = bTecho
    bTecho = HayTecho(UserPos.x, UserPos.y)
    If WasUnderRoof = bTecho Then Exit Sub
    If bRain And MapDat.LLUVIA = 1 Then
        Call ao20audio.PlayWeatherAudio(IIf(bTecho, SND_RAIN_IN_LOOP, SND_RAIN_OUT_LOOP))
    ElseIf bNieve And MapDat.NIEVE = 1 Then
        Call ao20audio.PlayWeatherAudio(IIf(bTecho, SND_NIEVEIN, SND_NIEVEOUT))
    End If
End Sub

Private Function SyncGrhPhase(ByRef Grh As Grh, ByVal newGrhIndex As Long) As Long
    Dim oldNum As Long, Elapsed As Long, phase As Long
    If Grh.started <= 0 Then SyncGrhPhase = FrameTime: Exit Function
    oldNum = GrhData(Grh.GrhIndex).NumFrames
    If oldNum <= 0 Then SyncGrhPhase = FrameTime: Exit Function
    Elapsed = Fix((FrameTime - Grh.started) / Grh.speed)
    phase = Elapsed Mod oldNum
    SyncGrhPhase = FrameTime - (phase * Grh.speed)
End Function

Public Sub InitGrhPreserve(ByRef Grh As Grh, ByVal GrhIndex As Long, Optional ByVal Loops As Integer = INFINITE_LOOPS)
    Dim keepStarted As Long
    If Grh.GrhIndex <> 0 And Grh.started > 0 Then
        keepStarted = SyncGrhPhase(Grh, GrhIndex)
    Else
        keepStarted = -1
    End If
    Call InitGrh(Grh, GrhIndex, keepStarted, Loops)
End Sub
