VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDX8Engine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'MOTOR GRÁFICO ESCRITO(mayormente) POR MENDUZ@NOICODER.COM
Option Explicit

Option Base 0

Public FrameNum               As Long

'Letter showing on screen
Private letter_text           As String

Private letter_grh            As grh

Private map_letter_grh        As grh

Private map_letter_grh_next   As Long

Private map_letter_a          As Single

Private map_letter_fadestatus As Byte

''
' Maximum number of dialogs that can exist.
Private Const MAX_DIALOGS     As Byte = 100

''
' Maximum length of any dialog line without having to split it.
Private Const MAX_LENGTH      As Byte = 18

''
' Number of milliseconds to add to the lifetime per dialog character
Private Const MS_PER_CHAR     As Byte = 100

''
' Number of extra milliseconds to add to the lifetime of a new dialog
Private Const MS_ADD_EXTRA    As Integer = 5000

''
' The dialog structure
'
' @param    textLine    Array of lines of a formated chat.
' @param    x           X coord for rendering.
' @param    y           Y coord for rendering.
' @param    startTime   The time (in ms) at which the dialog was created.
' @param    lifeTime    Time (in ms) this dialog should last.
' @param    charIndex   The charIndex that created this dialog.
' @param    color       The color to be used when rendering the text.
' @param    renderable  Set to True if the chat should be rendered this frame, False otherwise
'                           (used to skip dialogs from people outside render area).
Private Type dialog

    textLine()  As String
    x           As Integer
    y           As Integer
    startTime   As Long
    lifeTime    As Long
    charindex   As Integer
    color       As Long
    renderable  As Boolean
    Sube As Byte

End Type

Dim scroll_dialog_pixels_per_frame As Single

''
' Array if dialogs, sorted by the charIndex.
Private dialogs(MAX_DIALOGS - 1)   As dialog

''
' The number of dialogs being used at the moment.
Private dialogCount                As Byte

Private Type Fuente

    Tamanio As Integer
    Caracteres(0 To 255) As Long 'indice de cada letra

End Type

Private Fuentes(1 To 6)    As Fuente

Public WeatherFogX1        As Single

Public WeatherFogY1        As Single

Public WeatherFogX2        As Single

Public WeatherFogY2        As Single

Public WeatherDoFog        As Byte

Public WeatherFogCount     As Byte

Public ParticleOffsetX     As Long

Public ParticleOffsetY     As Long

Public LastOffsetX         As Integer

Public LastOffsetY         As Integer

Public EndTime             As Long

Private Const ScreenWidth  As Long = 538

Private Const ScreenHeight As Long = 376

Public bRunning            As Boolean

Private Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

Private Const FVF2 = D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX2

Dim font_count      As Long

Dim font_last       As Long

Private font_list() As D3DXFont

Public Enum FontAlignment

    fa_center = DT_CENTER
    fa_top = DT_TOP
    fa_left = DT_LEFT
    fa_topleft = DT_TOP Or DT_LEFT
    fa_bottomleft = DT_BOTTOM Or DT_LEFT
    fa_bottom = DT_BOTTOM
    fa_right = DT_RIGHT
    fa_bottomright = DT_BOTTOM Or DT_RIGHT
    fa_topright = DT_TOP Or DT_RIGHT

End Enum

Dim texture      As Direct3DTexture8

Dim TransTexture As Direct3DTexture8

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Long

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public fps                     As Long

Private FramesPerSecCounter    As Long

Private timerElapsedTime       As Single

Private timerTicksPerFrame     As Double

Public engineBaseSpeed         As Single

Private lFrameTimer            As Long

Private lFrameLimiter          As Long

Private ScrollPixelsPerFrameX  As Single

Private ScrollPixelsPerFrameY  As Single

Private TileBufferPixelOffsetX As Integer

Private TileBufferPixelOffsetY As Integer

Private MainViewTop            As Integer

Private MainViewLeft           As Integer

Private MainDestRect           As RECT

Private MainViewRect           As RECT

Private BackBufferRect         As RECT

Private MainViewWidth          As Integer

Private MainViewHeight         As Integer

Private MouseTileX             As Byte

Private MouseTileY             As Byte

Private iFrameIndex            As Byte  'Frame actual de la LL

Private llTick                 As Long  'Contador

Private WindowTileWidth        As Integer

Private WindowTileHeight       As Integer

Private HalfWindowTileWidth    As Integer

Private HalfWindowTileHeight   As Integer

Private Const GrhFogata        As Integer = 1521

Private Type Light

    active As Boolean 'Do we ignore this light?
    id As Long
    map_x As Integer 'Coordinates
    map_y As Integer
    color As Long 'Start colour
    range As Byte
    red As Byte
    green As Byte
    blue As Byte

End Type

'Light list
Dim light_list()  As Light

Dim light_count   As Long

Dim light_last    As Long

Private Estrella  As grh

Private Marco     As grh

Private BarraMana As grh

Private BarraVida As grh

Private BarraGris As grh

Private Type Particle

    PartCountLive As Integer
    destruir  As Boolean
    friction As Single
    x As Single
    y As Single
    vector_x As Single
    vector_y As Single
    angle As Single
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

'Modified by: Ryan Cain (Onezero)
'Last modify date: 5/14/2003
Private Type particle_group

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
    
    stream_type As Byte

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
    angle As Integer
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
    rgb_list(0 To 3) As Long
    
    'Added by Juan Martín Sotuyo Dodero
    speed As Single
    life_counter As Long
    
    'Added by David Justus
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer

End Type

'Particle system
Dim particle_group_list()  As particle_group

Dim particle_group_count   As Long

Public particle_group_last As Long

'Loaded Particle groups list
Dim StreamData()           As particle_group

Dim TotalStreams           As Long

'Dim fxs_list() As New clsFXS
'Dim fxs_count As Integer
'Dim fxs_last As Integer

Dim valoreBlur             As Long

Dim dimeTex                As Long

Dim tex                    As Direct3DTexture8

Dim D3DbackBuffer          As Direct3DSurface8

Dim zTarget                As Direct3DSurface8

Dim stencil                As Direct3DSurface8

Dim superTex               As Direct3DSurface8

Dim blur                   As Boolean

Public blur_factor         As Byte

Dim bump_map_texture       As Direct3DTexture8

Dim bump_map_texture_ex    As Direct3DTexture8

Dim bump_map_supported     As Boolean

Dim bump_map_powa          As Boolean

'Sets a Grh animation to loop indefinitely.

#Const HARDCODED = False 'True ' == MÁS FPS ^^

Private Function GetElapsedTime() As Single

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

End Function

Public Sub Text_Render(ByVal font As D3DXFont, Text As String, ByVal Top As Long, ByVal Left As Long, ByVal Width As Long, ByVal Height As Long, ByVal color As Long, ByVal format As Long, Optional ByVal shadow As Boolean = False)

    '*****************************************************
    '****** Coded by Menduz (lord.yo.wo@gmail.com) *******
    '*****************************************************
    Dim TextRect   As RECT

    Dim ShadowRect As RECT
    
    TextRect.Top = Top
    TextRect.Left = Left
    TextRect.bottom = Top + Height
    TextRect.Right = Left + Width
    
    If shadow Then
        ShadowRect.Top = Top - 1
        ShadowRect.Left = Left - 2
        ShadowRect.bottom = (Top + Height) - 1
        ShadowRect.Right = (Left + Width) - 2
        D3DX.DrawText font, &HFF000000, Text, ShadowRect, format

    End If
    
    D3DX.DrawText font, color, Text, TextRect, format

End Sub

Public Sub Text_Render_ext(Text As String, ByVal Top As Long, ByVal Left As Long, ByVal Width As Long, ByVal Height As Long, ByVal color As Long, Optional ByVal shadow As Boolean = False, Optional ByVal center As Boolean = False, Optional ByVal font As Long = 0)

    If center = True Then
        Call Text_Render(font_list(font), Text, Top, Left, Width, Height, color, DT_VCENTER & DT_CENTER, shadow)
    Else
        Call Text_Render(font_list(font), Text, Top, Left, Width, Height, color, DT_TOP Or DT_LEFT, shadow)

    End If

End Sub

Private Sub Font_Make(ByVal font_index As Long, ByVal Style As String, ByVal bold As Boolean, ByVal italic As Boolean, ByVal size As Long)

    If font_index > font_last Then
        font_last = font_index
        ReDim Preserve font_list(1 To font_last)

    End If

    font_count = font_count + 1
    
    Dim font_desc As IFont

    Dim fnt       As New StdFont

    fnt.name = Style
    fnt.size = size
    fnt.bold = bold
    fnt.italic = italic
    
    Set font_desc = fnt
    Set font_list(font_index) = D3DX.CreateFont(D3DDevice, font_desc.hFont)

End Sub

Public Function Font_Create(ByVal Style As String, ByVal size As Long, ByVal bold As Boolean, ByVal italic As Boolean) As Long

    On Error GoTo ErrorHandler:

    Font_Create = Font_Next_Open
    Font_Make Font_Create, Style, bold, italic, size
ErrorHandler:
    Font_Create = 0

End Function

Private Function Font_Next_Open() As Long
    Font_Next_Open = font_last + 1

End Function

Private Function Font_Check(ByVal font_index As Long) As Boolean

    '*****************************************************
    '****** Coded by Menduz (lord.yo.wo@gmail.com) *******
    '*****************************************************
    If font_index > 0 And font_index <= font_last Then
        Font_Check = True

    End If

End Function

Function MakeVector(ByVal x As Single, ByVal y As Single, ByVal Z As Single) As D3DVECTOR
    '*****************************************************
    '****** Coded by Menduz (lord.yo.wo@gmail.com) *******
    '*****************************************************
    MakeVector.x = x
    MakeVector.y = y
    MakeVector.Z = Z

End Function

Public Sub Engine_ReInit()

    '*****************************************************
    '****** Coded by Menduz (lord.yo.wo@gmail.com) *******
    '*****************************************************
    On Error GoTo errhandler:

    Dim DispMode    As D3DDISPLAYMODE

    Dim DispModeBK  As D3DDISPLAYMODE

    Dim D3DWindow   As D3DPRESENT_PARAMETERS

    Dim ColorKeyVal As Long
    
    Set SurfaceDB = New clsTexManager
    
    Set dX = New DirectX8
    Set D3D = dX.Direct3DCreate()
    Set D3DX = New D3DX8
    
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispModeBK
    
    Call Engine_Font_Initialize

    With D3DWindow
        .Windowed = True

        'VSync_FPS = True
        If VSync_FPS = True Then
            .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
        Else
            .SwapEffect = D3DSWAPEFFECT_DISCARD

        End If
        
        .BackBufferFormat = DispMode.format
        'set color depth
        .BackBufferWidth = 1024 'frmMain.renderer.ScaleWidth
        .BackBufferHeight = 768 'frmMain.renderer.ScaleHeight
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
        .hDeviceWindow = frmmain.renderer.hwnd

    End With

    DispMode.format = D3DFMT_X8R8G8B8

    If D3D.CheckDeviceFormat(0, D3DDEVTYPE_HAL, DispMode.format, 0, D3DRTYPE_TEXTURE, D3DFMT_A8R8G8B8) = D3D_OK Then

        Dim Caps8 As D3DCAPS8

        D3D.GetDeviceCaps 0, D3DDEVTYPE_HAL, Caps8

        If (Caps8.TextureOpCaps And D3DTEXOPCAPS_DOTPRODUCT3) = D3DTEXOPCAPS_DOTPRODUCT3 Then
            bump_map_supported = True
        Else
            bump_map_supported = False
            DispMode.format = DispModeBK.format

        End If

    Else
        bump_map_supported = False
        DispMode.format = DispModeBK.format

    End If

    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmmain.renderer.hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
                                                            
    'frmMain.renderer.Width = 545
    'frmMain.renderer.Height = 415
    'frmMain.renderer.ScaleMode = 3
    'frmMain.renderer.ScaleWidth = 545
    'frmMain.renderer.ScaleHeight = 415
    
    HalfWindowTileHeight = (frmmain.renderer.ScaleHeight / 32) \ 2
    HalfWindowTileWidth = (frmmain.renderer.ScaleWidth / 32) \ 2
    
    TileBufferSize = 8
    TileBufferPixelOffsetX = (TileBufferSize - 1) * 32
    TileBufferPixelOffsetY = (TileBufferSize - 1) * 32
    
    D3DDevice.SetVertexShader FVF
    
    '//Transformed and lit vertices dont need lighting
    '   so we disable it...
    D3DDevice.SetRenderState D3DRS_LIGHTING, False
    
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
    
    Call SurfaceDB.Init(D3DX, D3DDevice, General_Get_Free_Ram_Bytes)
    
    MinXBorder = XMinMapSize + (frmmain.renderer.ScaleWidth / 64)
    MaxXBorder = XMaxMapSize - (frmmain.renderer.ScaleWidth / 64)
    MinYBorder = YMinMapSize + (frmmain.renderer.ScaleHeight / 64)
    MaxYBorder = YMaxMapSize - (frmmain.renderer.ScaleHeight / 64)
    MinYBorder = MinYBorder

    With Render_Connect_Rect
        .Top = 0
        .Left = 0
        .Right = frmConnect.render.ScaleWidth
        .bottom = frmConnect.render.ScaleHeight

    End With
    
    With Render_Main_Rect
        .Top = 0
        .Left = 0
        .Right = frmmain.renderer.ScaleWidth
        .bottom = frmmain.renderer.ScaleHeight

    End With
    
    D3DDevice.SetRenderState D3DRS_POINTSIZE, Engine_FToDW(2)
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
    D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    
    Font_Create "Tahoma", 8, True, 0
    Font_Create "Verdana", 8, False, 0
    Font_Create "Verdana", 11, True, False
    
    bRunning = True
    Exit Sub
errhandler:
    MsgBox ("Se ha producido un error desconocido al inicializar el engine. En juego se cerrara." & " Error: " & Err.number & " - " & Err.Description)
    Debug.Print "Error Number Returned: " & Err.number
    End

    bRunning = False

End Sub

Public Sub Engine_Init()

    '*****************************************************
    '****** Coded by Menduz (lord.yo.wo@gmail.com) *******
    '*****************************************************
    On Error GoTo errhandler:

    Dim DispMode    As D3DDISPLAYMODE

    Dim DispModeBK  As D3DDISPLAYMODE

    Dim D3DWindow   As D3DPRESENT_PARAMETERS

    Dim ColorKeyVal As Long
    
    Set SurfaceDB = New clsTexManager
    
    Set dX = New DirectX8
    Set D3D = dX.Direct3DCreate()
    Set D3DX = New D3DX8
    
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispModeBK
    
    Call Engine_Font_Initialize

    With D3DWindow
        .Windowed = True

        'VSync_FPS = True
        If VSync_FPS = True Then
            .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
        Else
            .SwapEffect = D3DSWAPEFFECT_DISCARD

        End If
        
        .BackBufferFormat = DispMode.format
        'set color depth
        .BackBufferWidth = 1024 'frmMain.renderer.ScaleWidth
        .BackBufferHeight = 768 'frmMain.renderer.ScaleHeight
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
        .hDeviceWindow = frmmain.renderer.hwnd

    End With

    DispMode.format = D3DFMT_X8R8G8B8

    If D3D.CheckDeviceFormat(0, D3DDEVTYPE_HAL, DispMode.format, 0, D3DRTYPE_TEXTURE, D3DFMT_A8R8G8B8) = D3D_OK Then

        Dim Caps8 As D3DCAPS8

        D3D.GetDeviceCaps 0, D3DDEVTYPE_HAL, Caps8

        If (Caps8.TextureOpCaps And D3DTEXOPCAPS_DOTPRODUCT3) = D3DTEXOPCAPS_DOTPRODUCT3 Then
            bump_map_supported = True
        Else
            bump_map_supported = False
            DispMode.format = DispModeBK.format

        End If

    Else
        bump_map_supported = False
        DispMode.format = DispModeBK.format

    End If

    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmmain.renderer.hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
                                                            
    'frmMain.renderer.Width = 545
    'frmMain.renderer.Height = 415
    'frmMain.renderer.ScaleMode = 3
    'frmMain.renderer.ScaleWidth = 545
    'frmMain.renderer.ScaleHeight = 415
    
    HalfWindowTileHeight = (frmmain.renderer.ScaleHeight / 32) \ 2
    HalfWindowTileWidth = (frmmain.renderer.ScaleWidth / 32) \ 2
    
    TileBufferSize = 8
    TileBufferPixelOffsetX = (TileBufferSize - 1) * 32
    TileBufferPixelOffsetY = (TileBufferSize - 1) * 32
    
    D3DDevice.SetVertexShader FVF
    
    '//Transformed and lit vertices dont need lighting
    '   so we disable it...
    D3DDevice.SetRenderState D3DRS_LIGHTING, False
    
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
    
    Call SurfaceDB.Init(D3DX, D3DDevice, General_Get_Free_Ram_Bytes)

    engineBaseSpeed = 0.018
    
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    'Set FPS value to 60 for startup
    fps = 60
    FramesPerSecCounter = 60
    scroll_dialog_pixels_per_frame = 4
    
    ScrollPixelsPerFrameX = 8.5
    ScrollPixelsPerFrameY = 8.5
    
    UserPos.x = 50
    UserPos.y = 50
    
    Estrella.framecounter = 1
    Estrella.GrhIndex = 35764
    Estrella.Started = 1
    
    Marco.framecounter = 1
    Marco.GrhIndex = 839
    Marco.Started = 1
    
    BarraMana.framecounter = 1
    BarraMana.GrhIndex = 840
    BarraMana.Started = 1
    
    BarraVida.framecounter = 1
    BarraVida.GrhIndex = 841
    BarraVida.Started = 1
    
    BarraGris.framecounter = 1
    BarraGris.GrhIndex = 842
    BarraGris.Started = 1
    
    MinXBorder = XMinMapSize + (frmmain.renderer.ScaleWidth / 64)
    MaxXBorder = XMaxMapSize - (frmmain.renderer.ScaleWidth / 64)
    MinYBorder = YMinMapSize + (frmmain.renderer.ScaleHeight / 64)
    MaxYBorder = YMaxMapSize - (frmmain.renderer.ScaleHeight / 64)
    MinYBorder = MinYBorder

    With Render_Connect_Rect
        .Top = 0
        .Left = 0
        .Right = frmConnect.render.ScaleWidth
        .bottom = frmConnect.render.ScaleHeight

    End With
    
    With Render_Main_Rect
        .Top = 0
        .Left = 0
        .Right = frmmain.renderer.ScaleWidth
        .bottom = frmmain.renderer.ScaleHeight

    End With
    
    D3DDevice.SetRenderState D3DRS_POINTSIZE, Engine_FToDW(2)
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
    D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    
    Font_Create "Tahoma", 8, True, 0
    Font_Create "Verdana", 8, False, 0
    Font_Create "Verdana", 11, True, False
    
    ' Colores comunes
    COLOR_WHITE(0) = D3DColorXRGB(255, 255, 255)
    COLOR_WHITE(1) = D3DColorXRGB(255, 255, 255)
    COLOR_WHITE(2) = D3DColorXRGB(255, 255, 255)
    COLOR_WHITE(3) = D3DColorXRGB(255, 255, 255)
    
    bRunning = True
    Exit Sub
errhandler:
    MsgBox ("Se ha producido un error desconocido al inicializar el engine. En juego se cerrara." & " Error Number Returned: " & Err.number)
    Debug.Print "Error Number Returned: " & Err.number
    End

    bRunning = False

End Sub

Public Sub Engine_Deinit()
    Erase MapData
    Erase charlist
    Set D3DDevice = Nothing
    Set D3D = Nothing
    Set dX = Nothing

End Sub

Private Function CreateTLVertex(x As Single, y As Single, Z As Single, rhw As Single, color As Long, Specular As Long, tu As Single, tv As Single) As TLVERTEX
    '*****************************************************
    '****** Coded by Menduz (lord.yo.wo@gmail.com) *******
    '*****************************************************
    CreateTLVertex.x = x
    CreateTLVertex.y = y
    CreateTLVertex.Z = Z
    CreateTLVertex.rhw = rhw
    CreateTLVertex.color = color
    CreateTLVertex.Specular = Specular
    CreateTLVertex.tu = tu
    CreateTLVertex.tv = tv

End Function

Public Sub Engine_ActFPS()

    If (GetTickCount() And &H7FFFFFFF) - lFrameTimer > 1000 Then
        fps = FramesPerSecCounter
        FramesPerSecCounter = 0
        lFrameTimer = GetTickCount

    End If

End Sub

Function InMapBounds(ByVal x As Integer, ByVal y As Integer) As Boolean

    '*****************************************************************
    'Checks to see if a tile position is in the maps bounds
    '*****************************************************************
    If x < XMinMapSize Or x > XMaxMapSize Or y < YMinMapSize Or y > YMaxMapSize Then
        Exit Function

    End If
    
    InMapBounds = True

End Function

Public Sub Draw_GrhIndex(ByVal grh_index As Long, ByVal x As Integer, ByVal y As Integer)

    If grh_index <= 0 Then Exit Sub

    Dim rgb_list(3) As Long
    
    rgb_list(0) = D3DColorXRGB(255, 255, 255)
    rgb_list(1) = D3DColorXRGB(255, 255, 255)
    rgb_list(2) = D3DColorXRGB(255, 255, 255)
    rgb_list(3) = D3DColorXRGB(255, 255, 255)
    
    Device_Box_Textured_Render grh_index, x, y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, rgb_list, GrhData(grh_index).sX, GrhData(grh_index).sY

End Sub

Public Sub Draw_GrhColor(ByVal grh_index As Long, ByVal x As Integer, ByVal y As Integer, ByRef text_color() As Long)

    If grh_index <= 0 Then Exit Sub
    
    Device_Box_Textured_Render grh_index, x, y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, text_color, GrhData(grh_index).sX, GrhData(grh_index).sY

End Sub

Public Sub Draw_GrhIndexColor(ByVal grh_index As Long, ByVal x As Integer, ByVal y As Integer)

    If grh_index <= 0 Then Exit Sub

    Dim rgb_list(3) As Long
    
    rgb_list(0) = D3DColorXRGB(255, 255, 255)
    rgb_list(1) = rgb_list(0)
    rgb_list(2) = rgb_list(0)
    rgb_list(3) = rgb_list(0)
    
    Device_Box_Textured_Render grh_index, x, y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, rgb_list, GrhData(grh_index).sX, GrhData(grh_index).sY, True

End Sub

Private Sub Draw_Grh(ByRef grh As grh, ByVal x As Integer, ByVal y As Integer, ByVal center As Byte, ByVal animate As Byte, ByRef rgb_list() As Long, Optional ByVal Alpha As Boolean, Optional ByVal map_x As Byte = 1, Optional ByVal map_y As Byte = 1, Optional ByVal angle As Single)

    On Error Resume Next

    Dim CurrentGrhIndex As Long

    If grh.GrhIndex = 0 Then Exit Sub
    If animate Then
        If grh.Started = 1 Then
            grh.framecounter = grh.framecounter + (timerElapsedTime * GrhData(grh.GrhIndex).NumFrames / grh.speed) * 0.5

            If grh.framecounter > GrhData(grh.GrhIndex).NumFrames Then
                grh.framecounter = (grh.framecounter Mod GrhData(grh.GrhIndex).NumFrames) + 1
                
                If grh.Loops <> -1 Then
                    If grh.Loops > 0 Then
                        grh.Loops = grh.Loops - 1
                    Else
                        grh.Started = 0

                        Rem Exit Sub 'Agregado por Ladder 08/09/2014
                    End If

                End If

            End If

        End If

    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(grh.GrhIndex).Frames(grh.framecounter)

    'Center Grh over X,Y pos
    If center Then
        If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
            x = x - Int(GrhData(CurrentGrhIndex).TileWidth * (32 \ 2)) + 32 \ 2

        End If

        If GrhData(grh.GrhIndex).TileHeight <> 1 Then
            y = y - Int(GrhData(CurrentGrhIndex).TileHeight * 32) + 32

        End If

    End If

    Device_Box_Textured_Render CurrentGrhIndex, x, y, GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, rgb_list, GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, Alpha, angle
    'exits:

End Sub

Private Sub Draw_GrhFX(ByRef grh As grh, ByVal x As Integer, ByVal y As Integer, ByVal center As Byte, ByVal animate As Byte, ByRef rgb_list() As Long, Optional ByVal Alpha As Boolean, Optional ByVal map_x As Byte = 1, Optional ByVal map_y As Byte = 1, Optional ByVal angle As Single, Optional ByVal charindex As Integer)

    On Error Resume Next

    Dim cantidaddeframes As Long

    Dim CurrentGrhIndex  As Long

    If grh.GrhIndex = 0 Then Exit Sub
    If animate Then
        If grh.Started = 1 Then
            grh.framecounter = grh.framecounter + (timerElapsedTime * GrhData(grh.GrhIndex).NumFrames / grh.speed)
            
            If grh.AnimacionContador > 0 Then
                grh.AnimacionContador = grh.AnimacionContador - (timerElapsedTime * GrhData(grh.GrhIndex).NumFrames / grh.speed)

            End If
            
            If grh.framecounter > GrhData(grh.GrhIndex).NumFrames Then
            
                grh.framecounter = (grh.framecounter Mod GrhData(grh.GrhIndex).NumFrames) + 1

                If grh.Loops <> -1 Then
                    If grh.Loops > 0 Then
                        
                        grh.Loops = grh.Loops - 1
                    Else
                        grh.Started = 0

                        Exit Sub 'Agregado por Ladder 08/09/2014

                    End If

                End If

            End If

        End If

    End If
    
    'If grh.Loops > 3 Then
    Dim colorz(0 To 3) As Long
    
    '  If grh.AnimacionContador < 30 Then
    
    '    Dim alphablen As Byte
    
    '  grh.Alpha = grh.Alpha - 1
        
    '  colorz(0) = D3DColorARGB(grh.Alpha, 255, 255, 255)
    '   colorz(1) = D3DColorARGB(grh.Alpha, 255, 255, 255)
    '   colorz(2) = D3DColorARGB(grh.Alpha, 255, 255, 255)
    '   colorz(3) = D3DColorARGB(grh.Alpha, 255, 255, 255)
    
    '   rgb_list(0) = colorz(0)
    '   rgb_list(1) = colorz(0)
    '   rgb_list(2) = colorz(0)
    '   rgb_list(3) = colorz(0)
    ' End If
    If grh.AnimacionContador < grh.CantAnim * 0.1 Then
            
        grh.Alpha = grh.Alpha - 1
            
        colorz(0) = D3DColorARGB(grh.Alpha, 255, 255, 255)
        colorz(1) = D3DColorARGB(grh.Alpha, 255, 255, 255)
        colorz(2) = D3DColorARGB(grh.Alpha, 255, 255, 255)
        colorz(3) = D3DColorARGB(grh.Alpha, 255, 255, 255)
        
        rgb_list(0) = colorz(0)
        rgb_list(1) = colorz(0)
        rgb_list(2) = colorz(0)
        rgb_list(3) = colorz(0)

        If grh.Alpha = 0 And charindex > 0 Then
            charlist(charindex).fX.Started = 0
            Exit Sub

        End If

    End If
    
    If grh.AnimacionContador > grh.CantAnim * 0.6 Then
        If grh.Alpha < 220 Then
            grh.Alpha = grh.Alpha + 1

        End If
        
        colorz(0) = D3DColorARGB(grh.Alpha, 255, 255, 255)
        colorz(1) = D3DColorARGB(grh.Alpha, 255, 255, 255)
        colorz(2) = D3DColorARGB(grh.Alpha, 255, 255, 255)
        colorz(3) = D3DColorARGB(grh.Alpha, 255, 255, 255)
    
        rgb_list(0) = colorz(0)
        rgb_list(1) = colorz(0)
        rgb_list(2) = colorz(0)
        rgb_list(3) = colorz(0)

    End If
    
    ' End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(grh.GrhIndex).Frames(grh.framecounter)

    'Center Grh over X,Y pos
    If center Then
        If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
            x = x - Int(GrhData(CurrentGrhIndex).TileWidth * (32 \ 2)) + 32 \ 2

        End If

        If GrhData(grh.GrhIndex).TileHeight <> 1 Then
            y = y - Int(GrhData(CurrentGrhIndex).TileHeight * 32) + 32

        End If

    End If

    Device_Box_Textured_Render CurrentGrhIndex, x, y, GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, rgb_list, GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, Alpha, angle
    'exits:

End Sub

Private Sub Draw_GrhSinLuz(ByRef grh As grh, ByVal x As Integer, ByVal y As Integer, ByVal center As Byte, ByVal animate As Byte, Optional ByVal Alpha As Boolean, Optional ByVal map_x As Byte = 1, Optional ByVal map_y As Byte = 1, Optional ByVal angle As Single)

    Dim CurrentGrhIndex As Long

    If grh.GrhIndex = 0 Then Exit Sub
    If animate Then
        If grh.Started = 1 Then
            grh.framecounter = grh.framecounter + (timerElapsedTime * GrhData(grh.GrhIndex).NumFrames / grh.speed)

            If grh.framecounter > GrhData(grh.GrhIndex).NumFrames Then
                grh.framecounter = (grh.framecounter Mod GrhData(grh.GrhIndex).NumFrames) + 1
                
                If grh.Loops <> -1 Then
                    If grh.Loops > 0 Then
                        grh.Loops = grh.Loops - 1
                    Else
                        grh.Started = 0

                    End If

                End If

            End If

        End If

    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(grh.GrhIndex).Frames(grh.framecounter)

    'Center Grh over X,Y pos
    If center Then
        If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
            x = x - Int(GrhData(CurrentGrhIndex).TileWidth * (32 \ 2)) + 32 \ 2

        End If

        If GrhData(grh.GrhIndex).TileHeight <> 1 Then
            y = y - Int(GrhData(CurrentGrhIndex).TileHeight * 32) + 32

        End If

    End If
    
    Static light_value(0 To 3) As Long

    light_value(0) = map_base_light
    light_value(1) = light_value(0)
    light_value(2) = light_value(0)
    light_value(3) = light_value(0)

    Device_Box_Textured_Render CurrentGrhIndex, x, y, GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, light_value, GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, Alpha, angle
    'exits:

End Sub

Public Function Map_Base_Light_Get() As Long
    '**************************************************************
    'Author: Aaron Perkins - Modified by Augusto José Rando
    'Last Modify Date: 6/12/2005
    '
    '**************************************************************
    Map_Base_Light_Get = map_base_light

End Function

Public Function Map_Base_Light_Set(ByVal base_light As Long)

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '
    '**************************************************************
    If map_base_light <> base_light Then
        map_base_light = base_light
        
    End If
    
End Function

Public Function Map_Fill(ByVal grh_index As Long, ByVal layer As Byte, Optional ByVal light_base_color As Long = -1, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single) As Boolean

    '**************************************************************
    'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
    'Last Modify Date: 1/04/2003
    '
    '**************************************************************
    Dim x As Integer

    Dim y As Integer
    
    'Base light color
    If light_base_color <> -1 Then
        If Not Map_Base_Light_Set(light_base_color) Then Exit Function

    End If
        
    Map_Fill = True

End Function

Public Sub render()

    '*****************************************************
    '****** Coded by Menduz (lord.yo.wo@gmail.com) *******
    '*****************************************************
    Rem On Error GoTo ErrorHandler:
    Dim temp_array(3) As Long

    On Error GoTo ErrorHandler:
    
    If Map_light_base = -1 And Not EfectoEnproceso Then
        Meteo_Engine.Meteo_Logic
    ElseIf UserEstado = 1 Then
        Meteo_Engine.Meteo_Logic
       
    End If

    D3DDevice.BeginScene
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 1#, 1#, 0
    ShowNextFrame

    frmmain.fps.Caption = "FPS: " & fps
    frmmain.ms.Caption = PingRender & "ms"
       
    If frmmain.Contadores.Enabled Then

        Dim PosY As Integer
       
        Dim PosX As Integer

        If FullScreen Then
            PosY = 90
            PosX = 10
            
            temp_array(0) = RGB(0, 255, 0)
            temp_array(1) = temp_array(0)
            temp_array(2) = temp_array(0)
            temp_array(3) = temp_array(0)
            Engine_Draw_Box 665, 480, 37, 15, D3DColorARGB(150, 100, 100, 100)
            
            Engine_Text_Render Val(UserAtributos(eAtributos.Fuerza)), 665, 480, temp_array, 1, True, 10, 160
            temp_array(0) = RGB(255, 255, 0)
            temp_array(1) = temp_array(0)
            temp_array(2) = temp_array(0)
            temp_array(3) = temp_array(0)
            Engine_Text_Render Val(UserAtributos(eAtributos.Agilidad)), 685, 480, temp_array, 1, True, 0, 160
        Else
            PosY = -10
            PosX = 5

        End If

        If DrogaCounter > 0 Then
            temp_array(0) = D3DColorXRGB(0, 153, 0)
            temp_array(1) = temp_array(0)
            temp_array(2) = temp_array(0)
            temp_array(3) = temp_array(0)
            
            PosY = PosY + 15
            Engine_Text_Render "Potenciado: " & CLng(DrogaCounter) & "s", PosX, PosY, temp_array, 1, True, 0, 160

        End If
        
        If OxigenoCounter > 0 Then

            Dim HR                  As Integer

            Dim ms                  As Integer

            Dim SS                  As Integer

            Dim secs                As Integer

            Dim TextoOxigenoCounter As String
        
            temp_array(0) = RGB(50, 100, 255)
            temp_array(1) = temp_array(0)
            temp_array(2) = temp_array(0)
            temp_array(3) = temp_array(0)
            secs = OxigenoCounter
            HR = secs \ 3600
            ms = (secs Mod 3600) \ 60
            SS = (secs Mod 3600) Mod 60

            If SS > 9 Then
                TextoOxigenoCounter = ms & ":" & SS
            Else
                TextoOxigenoCounter = ms & ":0" & SS

            End If
            
            PosY = PosY + 15

            If ms < 1 Then
                frmmain.oxigenolbl = SS
                frmmain.oxigenolbl.ForeColor = vbRed
            Else
                frmmain.oxigenolbl = ms
                frmmain.oxigenolbl.ForeColor = vbWhite

            End If

            Engine_Text_Render "Oxigeno: " & TextoOxigenoCounter, PosX, PosY, temp_array, 1, True, 0, 128

        End If

    End If

    D3DDevice.EndScene
    D3DDevice.Present Render_Main_Rect, ByVal 0, 0, ByVal 0

    lFrameLimiter = (GetTickCount() And &H7FFFFFFF)
    FramesPerSecCounter = FramesPerSecCounter + 1
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed

    Exit Sub
ErrorHandler:

    If D3DDevice.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        Call engine.Engine_ReInit
        prgRun = True
        pausa = False
        QueRender = 0

    End If
 
End Sub

Public Sub Char_Dialog_Set(ByVal char_index As Integer, ByVal char_dialog As String, ByVal char_dialog_color As Long, ByVal char_dialog_life As Byte, ByVal Sube As Byte, Optional ByVal font_index As Integer = 1)
    
    If Char_Check(char_index) Then
        charlist(char_index).dialog = char_dialog
        charlist(char_index).dialog_color = char_dialog_color
        charlist(char_index).dialog_life = char_dialog_life
        charlist(char_index).dialog_font_index = font_index
        charlist(char_index).dialog_scroll = True
        charlist(char_index).dialog_offset_counter_y = -(IIf(BodyData(charlist(char_index).iBody).HeadOffset.y = 0, -32, BodyData(charlist(char_index).iBody).HeadOffset.y) / 2)
        charlist(char_index).AlphaText = 255

    End If

    Dim slot As Integer

    Dim i    As Long
    
    slot = BinarySearch(char_index)
    
    If slot < 0 Then
        If dialogCount = MAX_DIALOGS Then Exit Sub  'Out of space! Should never happen....
        
        'We need to add it. Get insertion index and move list backwards.
        slot = Not slot
        
        For i = dialogCount To slot + 1 Step -1
            dialogs(i) = dialogs(i - 1)
        Next i
        
        dialogCount = dialogCount + 1

    End If
    
    If char_dialog_life = 250 Then

        With dialogs(slot)
            .startTime = (GetTickCount() And &H7FFFFFFF)
            .lifeTime = MS_ADD_EXTRA + (MS_PER_CHAR * Len(char_dialog))
            .charindex = char_index

        End With

    Else

        With dialogs(slot)
            .startTime = (GetTickCount() And &H7FFFFFFF)
            .lifeTime = (MS_PER_CHAR * Len(char_dialog))
            .charindex = char_index

        End With

    End If
    
End Sub

Private Function BinarySearch(ByVal charindex As Integer) As Integer

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

Public Sub Char_Dialog_Remove(ByVal char_index As Integer, ByVal Index As Integer)

    If char_index = 0 Then Exit Sub

    If charlist(char_index).AlphaText > 0 Then
        charlist(char_index).AlphaText = charlist(char_index).AlphaText - (scroll_dialog_pixels_per_frame * timerTicksPerFrame)
        Exit Sub

    End If

    Dim slot As Integer

    Dim i    As Long
    
    slot = BinarySearch(char_index)
    
    If slot < 0 Then Exit Sub
    
    For i = slot To MAX_DIALOGS - 2
        dialogs(i) = dialogs(i + 1)
    Next i
    
    dialogCount = dialogCount - 1
    
    If Char_Check(char_index) Then
        charlist(char_index).dialog = ""
        charlist(char_index).dialog_color = 0
        charlist(char_index).dialog_life = 0

    End If

End Sub

Sub ShowNextFrame()

    'Call RenderSounds
    Static OffsetCounterX As Single

    Static OffsetCounterY As Single
     
    If UserMoving Then

        '****** Move screen Left and Right if needed ******
        If AddtoUserPos.x <> 0 Then
            OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.x * timerTicksPerFrame * charlist(UserCharIndex).Speeding

            If Abs(OffsetCounterX) >= Abs(32 * AddtoUserPos.x) Then
                OffsetCounterX = 0
                AddtoUserPos.x = 0
                UserMoving = False

            End If

        End If
            
        '****** Move screen Up and Down if needed ******
        If AddtoUserPos.y <> 0 Then
            OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.y * timerTicksPerFrame * charlist(UserCharIndex).Speeding

            If Abs(OffsetCounterY) >= Abs(32 * AddtoUserPos.y) Then
                OffsetCounterY = 0
                AddtoUserPos.y = 0
                UserMoving = False

            End If

        End If

    End If

    If UserCiego Then
        Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
        Call RenderScreenCiego(UserPos.x - AddtoUserPos.x, UserPos.y - AddtoUserPos.y, OffsetCounterX, OffsetCounterY)
    Else
        'Reparacion de pj
        Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
        Call RenderScreen(UserPos.x - AddtoUserPos.x, UserPos.y - AddtoUserPos.y, OffsetCounterX, OffsetCounterY)
                
    End If

End Sub

Sub RenderScreenCiego(ByVal tilex As Integer, ByVal tiley As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 8/14/2007
    'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
    'Renders everything to the viewport
    '**************************************************************

    Dim y                As Integer     'Keeps track of where on map we are

    Dim x                As Integer     'Keeps track of where on map we are

    Dim screenminY       As Integer  'Start Y pos on current screen

    Dim screenmaxY       As Integer  'End Y pos on current screen

    Dim screenminX       As Integer  'Start X pos on current screen

    Dim screenmaxX       As Integer  'End X pos on current screen

    Dim minY             As Integer  'Start Y pos on current map

    Dim MaxY             As Integer  'End Y pos on current map

    Dim minX             As Integer  'Start X pos on current map

    Dim MaxX             As Integer  'End X pos on current map

    Dim ScreenX          As Integer  'Keeps track of where to place tile on screen

    Dim ScreenY          As Integer  'Keeps track of where to place tile on screen

    Dim minXOffset       As Integer

    Dim minYOffset       As Integer

    Dim PixelOffsetXTemp As Integer 'For centering grhs

    Dim PixelOffsetYTemp As Integer 'For centering grhs

    Dim CurrentGrhIndex  As Integer

    Dim OffX             As Integer

    Dim Offy             As Integer

    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    minY = screenminY - TileBufferSize
    MaxY = screenmaxY + TileBufferSize
    minX = screenminX - TileBufferSize
    MaxX = screenmaxX + TileBufferSize
    
    'Make sure mins and maxs are allways in map bounds
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize

    End If
    
    If MaxY > YMaxMapSize Then MaxY = YMaxMapSize
    
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize

    End If
    
    If MaxX > XMaxMapSize Then MaxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1

    End If
    
    If screenmaxY < YMaxMapSize Then screenmaxY = screenmaxY + 1
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1

    End If
    
    If screenmaxX < XMaxMapSize Then screenmaxX = screenmaxX + 1
    
    If screenminY < 1 Then screenminY = 1
    If screenminX < 1 Then screenminX = 1
    If screenmaxY > 100 Then screenmaxY = 100
    If screenmaxX > 100 Then screenmaxX = 100
    
    Dim PicClimaRGB(0 To 3) As Long

    Dim Climapic            As grh
   
    ColorCiego(0) = D3DColorARGB(255, 15, 15, 15)
    ColorCiego(1) = ColorCiego(0)
    ColorCiego(2) = ColorCiego(0)
    ColorCiego(3) = ColorCiego(0)
    'If minY < 1 Then minY = 1
    'If minX < 1 Then minX = 1
    ' If maxY > 100 Then maxY = 100
    ' If maxX > 100 Then maxX = 100
    
    'Draw floor layer
    For y = screenminY To screenmaxY
        For x = screenminX To screenmaxX
            'Layer 1 **********************************
            Call Draw_Grh(MapData(x, y).Graphic(1), (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, 0, 1, ColorCiego, , x, y)
            '******************************************
            ScreenX = ScreenX + 1
        Next x

        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - x + screenminX
        ScreenY = ScreenY + 1
    Next y
    
    If HayLayer2 Then
        ScreenY = minYOffset - TileBufferSize

        For y = minY To MaxY
            ScreenX = minXOffset - TileBufferSize

            For x = minX To MaxX

                With MapData(x, y)

                    '***********************************************
                    If MapData(x, y).Graphic(2).GrhIndex <> 0 Then
                        Call Draw_Grh(MapData(x, y).Graphic(2), (ScreenX * 32 + PixelOffsetX), (ScreenY * 32 + PixelOffsetY), 1, 1, ColorCiego, , x, y)

                    End If
              
                End With

                ScreenX = ScreenX + 1
            Next x

            ScreenY = ScreenY + 1
        Next y

    End If
    
    ScreenY = minYOffset - TileBufferSize

    For y = minY To MaxY
        ScreenX = minXOffset - TileBufferSize

        For x = minX To MaxX
            PixelOffsetXTemp = ScreenX * 32 + PixelOffsetX
            PixelOffsetYTemp = ScreenY * 32 + PixelOffsetY
            
            With MapData(x, y)
                '******************************************

                'Object Layer **********************************
                If MapData(x, y).ObjGrh.GrhIndex <> 0 Then
                    Call Draw_Grh(MapData(x, y).ObjGrh, (ScreenX * 32 + PixelOffsetX), (ScreenY * 32 + PixelOffsetY), 1, 1, ColorCiego, , x, y)

                End If
                
                'Char layer ************************************
                'clones
            
                If MapData(x, y).charindex = UserCharIndex Then
                    If x <> UserPos.x Then
                        MapData(x, y).charindex = 0

                    End If
                    
                End If
                
                If .charindex <> 0 Then
                    If charlist(.charindex).AlphaPJ = 255 And charlist(.charindex).active = 1 Then
                        Call Char_RenderCiego(.charindex, PixelOffsetXTemp, PixelOffsetYTemp, x, y)

                    End If

                End If
                
                'Layer 3 *****************************************
                If .Graphic(3).GrhIndex <> 0 Then
                
                    Call Draw_Grh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, ColorCiego, False, x, y)
                            
                End If

                '************************************************

            End With
            
            ScreenX = ScreenX + 1
        Next x

        ScreenY = ScreenY + 1
    Next y

    ScreenY = minYOffset - 5

    ScreenY = minYOffset - TileBufferSize

    For y = minY To MaxY
        ScreenX = minXOffset - TileBufferSize

        For x = minX To MaxX

            With MapData(x, y)
                '***********************************************
                
                If .particle_Index = 184 Then
                    If meteo_estado = 3 Or meteo_estado = 4 Then
                        If .particle_group > 0 Then
                            Call Particle_Group_Render(.particle_group, ScreenX * 32 + PixelOffsetX + 15, ScreenY * 32 + PixelOffsetY + 15)

                        End If

                    End If

                End If

                If .particle_Index <> 184 Then
                    If .particle_group > 0 Then
                        Call Particle_Group_Render(.particle_group, ScreenX * 32 + PixelOffsetX + 15, ScreenY * 32 + PixelOffsetY + 15)

                    End If

                End If
                
            End With

            ScreenX = ScreenX + 1
        Next x

        ScreenY = ScreenY + 1
    Next y
 
    'Draw blocked tiles and grid
 
    If HayLayer4 Then

        Dim rgb_list(0 To 3)  As Long

        Dim rgb_list2(0 To 3) As Long

        rgb_list2(0) = D3DColorARGB(255, ColorAmbiente.r, ColorAmbiente.g, ColorAmbiente.b)
        rgb_list2(1) = D3DColorARGB(255, ColorAmbiente.r, ColorAmbiente.g, ColorAmbiente.b)
        rgb_list2(2) = D3DColorARGB(255, ColorAmbiente.r, ColorAmbiente.g, ColorAmbiente.b)
        rgb_list2(3) = D3DColorARGB(255, ColorAmbiente.r, ColorAmbiente.g, ColorAmbiente.b)
    
        ScreenY = minYOffset - TileBufferSize

        For y = minY To MaxY
            ScreenX = minXOffset - TileBufferSize

            For x = minX To MaxX
        
                If MapData(x, y).Graphic(4).GrhIndex Then
        
                    'Layer 4 **********************************
                    If bTecho Then

                        If MapData(UserPos.x, UserPos.y).Trigger = MapData(x, y).Trigger Then
                    
                            If MapData(x, y).GrhBlend <= 20 Then MapData(x, y).GrhBlend = 20
                            MapData(x, y).GrhBlend = MapData(x, y).GrhBlend - (timerTicksPerFrame * 12)
                    
                            rgb_list(0) = D3DColorARGB(CInt(MapData(x, y).GrhBlend), ColorAmbiente.r, ColorAmbiente.g, ColorAmbiente.b)
                            rgb_list(1) = rgb_list(0)
                            rgb_list(2) = rgb_list(0)
                            rgb_list(3) = rgb_list(0)
                        
                            Call Draw_Grh(MapData(x, y).Graphic(4), ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, 1, 1, ColorCiego, , x, y)
                        Else
                            Call Draw_Grh(MapData(x, y).Graphic(4), ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, 1, 1, ColorCiego, , x, y)

                        End If

                    Else
                 
                        MapData(x, y).GrhBlend = MapData(x, y).GrhBlend + (timerTicksPerFrame * 12)

                        If MapData(x, y).GrhBlend >= 255 Then MapData(x, y).GrhBlend = 255

                        rgb_list(0) = D3DColorARGB(CInt(MapData(x, y).GrhBlend), ColorAmbiente.r, ColorAmbiente.g, ColorAmbiente.b)
                        rgb_list(1) = rgb_list(0)
                        rgb_list(2) = rgb_list(0)
                        rgb_list(3) = rgb_list(0)
                        
                        Call Draw_Grh(MapData(x, y).Graphic(4), ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, 1, 1, ColorCiego, , x, y)
          
                    End If

                End If
 
                '**********************************
                ScreenX = ScreenX + 1
            Next x

            ScreenY = ScreenY + 1
        Next y
        
    End If

    'If MostrarTrofeo Then

    '    Dim TrofeoRGB(0 To 3) As Long
    '    Dim Trofeo As grh
    '    Trofeo.FrameCounter = 1
    '    Trofeo.grhindex = 32018
    '    Trofeo.Started = 1
    '    TrofeoRGB(0) = D3DColorARGB(100, 255, 0, 0)
    '    TrofeoRGB(1) = D3DColorARGB(100, 255, 0, 0)
    '    TrofeoRGB(2) = D3DColorARGB(100, 0, 0, 255)
    '    TrofeoRGB(3) = D3DColorARGB(100, 0, 0, 255)
    '  Engine_Draw_Box CInt(clicX), CInt(clicY), 190, 180, D3DColorARGB(180, 100, 100, 100)
    '        Grh_Render Trofeo, 690, 50, TrofeoRGB, True, True, True
    ' Call Draw_Grh(Trofeo, 690, 50, 1, 0, TrofeoRGB, False, 0, 0, 0)
    'End If

    If Pregunta Then
        'PreguntaScreen = "¿Esta seguro que asen es gay? ¿Que se lo come a fede?"
        Engine_Draw_Box 283, 180, 170, 80, D3DColorARGB(200, 219, 116, 3)
        Engine_Draw_Box 288, 185, 160, 70, D3DColorARGB(200, 51, 27, 3)

        Dim preguntaGrh As grh

        preguntaGrh.framecounter = 1
        preguntaGrh.GrhIndex = 32120
        preguntaGrh.Started = 1
        rgb_list(0) = D3DColorARGB(255, 255, 255, 255)
        rgb_list(1) = rgb_list(0)
        rgb_list(2) = rgb_list(0)
        rgb_list(3) = rgb_list(0)
        Engine_Text_Render PreguntaScreen, 290, 190, rgb_list, 1, True
        Call Draw_Grh(preguntaGrh, 392, 223, 1, 0, rgb_list, False, 0, 0, 0)

    End If

    If bRain Then
        If MapDat.LLUVIA Then
            'Screen positions were hardcoded by now
            ScreenX = 250
            ScreenY = 0
            Call Particle_Group_Render(meteo_particle, ScreenX, ScreenY)
            LastOffsetX = ParticleOffsetX
            LastOffsetY = ParticleOffsetY

        End If

    End If

    If AlphaNiebla Then
        If MapDat.niebla Then
            Engine_Weather_UpdateFog

        End If

    End If

    If bNieve Then
        If MapDat.NIEVE Then
            If engine.Engine_Meteo_Particle_Get <> 0 Then
                'Screen positions were hardcoded by now
                ScreenX = 250
                ScreenY = 0
                Call Particle_Group_Render(meteo_particle, ScreenX, ScreenY)

            End If

        End If

    End If

    'Pelota
    'If DibujarPelota Then

    'If Pelota.Fps = 100 Then DibujarPelota = False: Exit Sub
    '   Pelota.X = Pelota.X + Pelota.DireccionX
    '   Pelota.Y = Pelota.Y + Pelota.DireccionY
    '  Pelota.Fps = Pelota.Fps + 1
    '     Call Particle_Group_Render(spell_particle, Pelota.X, Pelota.Y)
    'End If
    'Pelota

    'If CaminandoMacro Then
    'Call Particle_Group_Render(spell_particle, CaminarX, CaminarY)
    'End If

    If cartel Then

        Dim cartelito(0 To 3) As Long

        Dim TempGrh           As grh

        TempGrh.framecounter = 1
        TempGrh.GrhIndex = GrhCartel
        cartelito(0) = D3DColorARGB(200, 255, 255, 255)
        cartelito(1) = rgb_list(0)
        cartelito(2) = rgb_list(0)
        cartelito(3) = rgb_list(0)
        Call Draw_Grh(TempGrh, CInt(clicX), CInt(clicY), 1, 0, cartelito, False, 0, 0, 0)
        Engine_Text_Render Leyenda, CInt(clicX - 100), CInt(clicY - 130), cartelito, 1, False

    End If

End Sub

Sub RenderScreen(ByVal tilex As Integer, ByVal tiley As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 8/14/2007
    'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
    'Renders everything to the viewport
    On Error Resume Next

    '**************************************************************
    Dim y                As Integer     'Keeps track of where on map we are

    Dim x                As Integer     'Keeps track of where on map we are

    Dim screenminY       As Integer  'Start Y pos on current screen

    Dim screenmaxY       As Integer  'End Y pos on current screen

    Dim screenminX       As Integer  'Start X pos on current screen

    Dim screenmaxX       As Integer  'End X pos on current screen

    Dim minY             As Integer  'Start Y pos on current map

    Dim MaxY             As Integer  'End Y pos on current map

    Dim minX             As Integer  'Start X pos on current map

    Dim MaxX             As Integer  'End X pos on current map

    Dim ScreenX          As Integer  'Keeps track of where to place tile on screen

    Dim ScreenY          As Integer  'Keeps track of where to place tile on screen

    Dim minXOffset       As Integer

    Dim minYOffset       As Integer

    Dim PixelOffsetXTemp As Integer 'For centering grhs

    Dim PixelOffsetYTemp As Integer 'For centering grhs

    Dim CurrentGrhIndex  As Integer

    Dim OffX             As Integer

    Dim Offy             As Integer

    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    minY = screenminY - TileBufferSize
    MaxY = screenmaxY + TileBufferSize
    minX = screenminX - TileBufferSize
    MaxX = screenmaxX + TileBufferSize
    
    'Make sure mins and maxs are allways in map bounds
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize

    End If
    
    If MaxY > YMaxMapSize Then MaxY = YMaxMapSize
    
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize

    End If
    
    If MaxX > XMaxMapSize Then MaxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1

    End If
    
    If screenmaxY < YMaxMapSize Then screenmaxY = screenmaxY + 1
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1

    End If
    
    If screenmaxX < XMaxMapSize Then screenmaxX = screenmaxX + 1
    
    If screenminY < 1 Then screenminY = 1
    If screenminX < 1 Then screenminX = 1
    If screenmaxY > 100 Then screenmaxY = 100
    If screenmaxX > 100 Then screenmaxX = 100
    
    Dim PicClimaRGB(0 To 3) As Long

    Dim Climapic            As grh
   
    'If minY < 1 Then minY = 1
    'If minX < 1 Then minX = 1
    ' If maxY > 100 Then maxY = 100
    ' If maxX > 100 Then maxX = 100
    'estoy renderizando 20 de Y y deberian ser 18
    'estoy renderizando 24 de x y deberian ser 18

    screenmaxY = screenmaxY ' 1 tile menos dibujo, vamos a ver que onda

    'Draw floor layer
    For y = screenminY To screenmaxY
        For x = screenminX To screenmaxX
     
            'Layer 1 **********************************
            Call Draw_Grh(MapData(x, y).Graphic(1), (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, 0, 1, MapData(x, y).light_value, , x, y)
            '******************************************
            ScreenX = ScreenX + 1
        Next x

        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - x + screenminX
        ScreenY = ScreenY + 1
        
    Next y
    
    If HayLayer2 Then
        ScreenY = minYOffset - TileBufferSize

        For y = minY To MaxY ' el -8 lo agrego
            ScreenX = minXOffset - TileBufferSize

            For x = minX To MaxX  ' -7 lo agrego

                With MapData(x, y)

                    '***********************************************
                    If MapData(x, y).Graphic(2).GrhIndex <> 0 Then
                        Call Draw_Grh(MapData(x, y).Graphic(2), (ScreenX * 32 + PixelOffsetX), (ScreenY * 32 + PixelOffsetY), 1, 1, MapData(x, y).light_value(), , x, y)

                    End If
              
                End With

                ScreenX = ScreenX + 1
            Next x

            ScreenY = ScreenY + 1
        Next y

    End If
    
    ScreenY = minYOffset - TileBufferSize

    For y = minY To MaxY
        ScreenX = minXOffset - TileBufferSize

        For x = minX To MaxX
            PixelOffsetXTemp = ScreenX * 32 + PixelOffsetX
            PixelOffsetYTemp = ScreenY * 32 + PixelOffsetY
            
            With MapData(x, y)
                '******************************************

                'Object Layer **********************************
                If MapData(x, y).ObjGrh.GrhIndex <> 0 Then
                    Call Draw_Grh(MapData(x, y).ObjGrh, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(x, y).light_value(), , x, y)

                End If
                
                'Char layer ************************************
                'evitamos reenderizar un clon del usuario
                If MapData(x, y).charindex = UserCharIndex Then
                    If x <> UserPos.x Then
                        MapData(x, y).charindex = 0

                    End If
                    
                End If
                
                If .charindex <> 0 Then
                    If charlist(.charindex).active = 1 Then
                        Call Char_Render(.charindex, PixelOffsetXTemp, PixelOffsetYTemp, x, y)

                    End If

                End If

                If .CharFantasma.Activo Then

                    Dim ColorFantasma(3) As Long
                    
                    If MapData(x, y).CharFantasma.AlphaB >= 3 Then
                        MapData(x, y).CharFantasma.AlphaB = MapData(x, y).CharFantasma.AlphaB - (timerTicksPerFrame * 6.7)
                        ColorFantasma(0) = D3DColorARGB(CInt(MapData(x, y).CharFantasma.AlphaB), ColorAmbiente.r, ColorAmbiente.g, ColorAmbiente.b)
                        ColorFantasma(1) = ColorFantasma(0)
                        ColorFantasma(2) = ColorFantasma(0)
                        ColorFantasma(3) = ColorFantasma(0)

                        If .CharFantasma.Heading = 1 Or .CharFantasma.Heading = 2 Then
                            Call Draw_Grh(.CharFantasma.Escudo, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, ColorFantasma, False, x, y)
                            Call Draw_Grh(.CharFantasma.Body, PixelOffsetXTemp + 1, PixelOffsetYTemp, 1, 1, ColorFantasma, False, x, y)
                            Call Draw_Grh(.CharFantasma.Head, PixelOffsetXTemp + .CharFantasma.OffX, PixelOffsetYTemp + .CharFantasma.Offy, 1, 1, ColorFantasma, False, x, y)
                            Call Draw_Grh(.CharFantasma.Casco, PixelOffsetXTemp + .CharFantasma.OffX, PixelOffsetYTemp + .CharFantasma.Offy, 1, 1, ColorFantasma, False, x, y)
                            Call Draw_Grh(.CharFantasma.Arma, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, ColorFantasma, False, x, y)
                        Else
                        
                            Call Draw_Grh(.CharFantasma.Body, PixelOffsetXTemp + 1, PixelOffsetYTemp, 1, 1, ColorFantasma, False, x, y)
                            Call Draw_Grh(.CharFantasma.Head, PixelOffsetXTemp + .CharFantasma.OffX, PixelOffsetYTemp + .CharFantasma.Offy, 1, 1, ColorFantasma, False, x, y)
                            Call Draw_Grh(.CharFantasma.Escudo, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, ColorFantasma, False, x, y)
                            Call Draw_Grh(.CharFantasma.Casco, PixelOffsetXTemp + .CharFantasma.OffX, PixelOffsetYTemp + .CharFantasma.Offy, 1, 1, ColorFantasma, False, x, y)
                            Call Draw_Grh(.CharFantasma.Arma, PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, ColorFantasma, False, x, y)

                        End If

                    Else
                        .CharFantasma.Activo = False

                    End If

                End If

                '*************************************************
                If EsArbol(.Graphic(3).GrhIndex) Then
                    Call Draw_Sombra(.Graphic(3), PixelOffsetXTemp + 40, PixelOffsetYTemp, 1, 1, False, x, y)

                End If
                
                'Layer 3 *****************************************
                If .Graphic(3).GrhIndex <> 0 Then
                    Call Draw_Grh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(x, y).light_value, False, x, y)

                End If

                '************************************************
                
            End With
            
            ScreenX = ScreenX + 1
        Next x

        ScreenY = ScreenY + 1
    Next y

    ScreenY = minYOffset - 5

    ScreenY = minYOffset - TileBufferSize

    For y = minY To MaxY
        ScreenX = minXOffset - TileBufferSize

        For x = minX To MaxX

            With MapData(x, y)

                '***********************************************
                If .particle_group > 0 Then
                    Call Particle_Group_Render(.particle_group, ScreenX * 32 + PixelOffsetX + 15, ScreenY * 32 + PixelOffsetY + 15)

                End If

            End With

            ScreenX = ScreenX + 1
        Next x

        ScreenY = ScreenY + 1
    Next y
 
    'Draw blocked tiles and grid
 
    If HayLayer4 Then

        Dim rgb_list(0 To 3)  As Long

        Dim rgb_list2(0 To 3) As Long

        rgb_list2(0) = D3DColorARGB(255, ColorAmbiente.r, ColorAmbiente.g, ColorAmbiente.b)
        rgb_list2(1) = D3DColorARGB(255, ColorAmbiente.r, ColorAmbiente.g, ColorAmbiente.b)
        rgb_list2(2) = D3DColorARGB(255, ColorAmbiente.r, ColorAmbiente.g, ColorAmbiente.b)
        rgb_list2(3) = D3DColorARGB(255, ColorAmbiente.r, ColorAmbiente.g, ColorAmbiente.b)
    
        ScreenY = minYOffset - TileBufferSize

        For y = minY To MaxY
            ScreenX = minXOffset - TileBufferSize

            For x = minX To MaxX
        
                If MapData(x, y).Graphic(4).GrhIndex Then
                            
                    Dim r, g, b As Byte

                    b = (map_base_light And 16711680) / 65536
                    g = (map_base_light And 65280) / 256
                    r = map_base_light And 255

                    'Layer 4 **********************************
                    If bTecho Then
                        If MapData(UserPos.x, UserPos.y).Trigger = MapData(x, y).Trigger Then
                    
                            If MapData(x, y).GrhBlend <= 20 Then MapData(x, y).GrhBlend = 20
                            MapData(x, y).GrhBlend = MapData(x, y).GrhBlend - (timerTicksPerFrame * 12)
                    
                            rgb_list(0) = D3DColorARGB(CInt(MapData(x, y).GrhBlend), b, g, r)
                            rgb_list(1) = rgb_list(0)
                            rgb_list(2) = rgb_list(0)
                            rgb_list(3) = rgb_list(0)
                        
                            Call Draw_Grh(MapData(x, y).Graphic(4), ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, 1, 1, rgb_list(), , x, y)
                        Else
                            Call Draw_Grh(MapData(x, y).Graphic(4), ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, 1, 1, rgb_list2(), , x, y)

                        End If

                    Else
                 
                        MapData(x, y).GrhBlend = MapData(x, y).GrhBlend + (timerTicksPerFrame * 12)

                        If MapData(x, y).GrhBlend >= 255 Then MapData(x, y).GrhBlend = 255

                        rgb_list(0) = D3DColorARGB(CInt(MapData(x, y).GrhBlend), b, g, r)
                        rgb_list(1) = rgb_list(0)
                        rgb_list(2) = rgb_list(0)
                        rgb_list(3) = rgb_list(0)
                        
                        Call Draw_Grh(MapData(x, y).Graphic(4), ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, 1, 1, rgb_list(), , x, y)
          
                    End If

                End If

                ScreenX = ScreenX + 1
            Next x

            ScreenY = ScreenY + 1
        Next y
        
    End If
        
    ScreenY = minYOffset - TileBufferSize

    For y = minY To MaxY
        ScreenX = minXOffset - TileBufferSize

        For x = minX To MaxX
            PixelOffsetXTemp = ScreenX * 32 + PixelOffsetX
            PixelOffsetYTemp = ScreenY * 32 + PixelOffsetY

            With MapData(x, y)
                
                If MapData(x, y).charindex <> 0 Then
                    If charlist(MapData(x, y).charindex).active = 1 Then
                        Call Char_TextRender(MapData(x, y).charindex, PixelOffsetXTemp, PixelOffsetYTemp, x, y)

                    End If

                End If

                modRenderValue.Draw x, y, PixelOffsetXTemp + 16, PixelOffsetYTemp, timerTicksPerFrame
                
                Dim i         As Byte

                Dim colorz(3) As Long

                If .FxCount > 0 Then

                    For i = 1 To .FxCount

                        If .FxList(i).FxIndex > 0 And .FxList(i).Started <> 0 Then
                            colorz(0) = D3DColorARGB(220, 255, 255, 255)
                            colorz(1) = D3DColorARGB(220, 255, 255, 255)
                            colorz(2) = D3DColorARGB(220, 255, 255, 255)
                            colorz(3) = D3DColorARGB(220, 255, 255, 255)

                            If FxData(.FxList(i).FxIndex).IsPNG = 1 Then
                                Call Draw_GrhFX(.FxList(i), PixelOffsetXTemp + FxData(.FxList(i).FxIndex).OffsetX, PixelOffsetYTemp + FxData(.FxList(i).FxIndex).OffsetY + 20, 1, 1, colorz, False)
                                ' Call Draw_GrhFX(.FxList(i), PixelOffsetXTemp + FxData(.FxList(i).OffsetX, PixelOffsetYTemp + FxData(.FxList(i)).Offsety + 20, 1, 1, colorz, False)
                            Else
                                Call Draw_GrhFX(.FxList(i), PixelOffsetXTemp + FxData(.FxList(i).FxIndex).OffsetX, PixelOffsetYTemp + FxData(.FxList(i).FxIndex).OffsetY + 20, 1, 1, colorz, True)

                            End If

                        End If

                        If .FxList(i).Started = 0 Then
                            .FxList(i).FxIndex = 0

                        End If

                    Next i

                    If .FxList(.FxCount).Started = 0 Then
                        .FxCount = .FxCount - 1

                    End If

                End If

            End With
            
            ScreenX = ScreenX + 1
        Next x

        ScreenY = ScreenY + 1
    Next y

    ScreenY = minYOffset - 5
    
    If bRain Then
        If MapDat.LLUVIA Then
            'Screen positions were hardcoded by now
            ScreenX = 250
            ScreenY = 0
            Call Particle_Group_Render(meteo_particle, ScreenX, ScreenY)
            LastOffsetX = ParticleOffsetX
            LastOffsetY = ParticleOffsetY

        End If

    End If

    If AlphaNiebla Then
        If MapDat.niebla Then
            Engine_Weather_UpdateFog

        End If

    End If

    If bNieve Then
        If MapDat.NIEVE Then
            If engine.Engine_Meteo_Particle_Get <> 0 Then
                'Screen positions were hardcoded by now
                ScreenX = 250
                ScreenY = 0
                Call Particle_Group_Render(meteo_particle, ScreenX, ScreenY)

            End If

        End If

    End If

    Dim macroPic(0 To 3) As Long

    Dim TempGrh          As grh

    If Pregunta Then
        Engine_Draw_Box 283, 180, 170, 80, D3DColorARGB(200, 150, 20, 3)
        Engine_Draw_Box 288, 185, 160, 70, D3DColorARGB(200, 25, 25, 23)

        Dim preguntaGrh As grh

        preguntaGrh.framecounter = 1
        preguntaGrh.GrhIndex = 32120
        preguntaGrh.Started = 1
        macroPic(0) = D3DColorARGB(255, 255, 255, 255)
        macroPic(1) = macroPic(0)
        macroPic(2) = macroPic(0)
        macroPic(3) = macroPic(0)
        Engine_Text_Render PreguntaScreen, 290, 190, macroPic, 1, True
        Call Draw_Grh(preguntaGrh, 392, 223, 1, 0, macroPic, False, 0, 0, 0)

    End If

    Effect_Render_All

    If cartel Then

        Dim cartelito(0 To 3) As Long

        TempGrh.framecounter = 1
        TempGrh.GrhIndex = GrhCartel
        cartelito(0) = D3DColorARGB(200, 255, 255, 255)
        cartelito(1) = rgb_list(0)
        cartelito(2) = rgb_list(0)
        cartelito(3) = rgb_list(0)
        Call Draw_Grh(TempGrh, CInt(clicX), CInt(clicY), 1, 0, cartelito, False, 0, 0, 0)
        Engine_Text_Render Leyenda, CInt(clicX - 100), CInt(clicY - 130), cartelito, 1, False

    End If

    Dim temp_array(0 To 3) As Long

    If map_letter_fadestatus > 0 Then
        If map_letter_fadestatus = 1 Then
            map_letter_a = map_letter_a + (timerTicksPerFrame * 3.5)

            If map_letter_a >= 255 Then
                map_letter_a = 255
                map_letter_fadestatus = 2

            End If

        Else
            map_letter_a = map_letter_a - (timerTicksPerFrame * 3.5)

            If map_letter_a <= 0 Then
                map_letter_fadestatus = 0
                map_letter_a = 0
                 
                If map_letter_grh_next > 0 Then
                    map_letter_grh.GrhIndex = map_letter_grh_next
                    map_letter_fadestatus = 1
                    map_letter_grh_next = 0

                End If
                
            End If

        End If

    End If
    
    If Len(letter_text) Then
        temp_array(0) = D3DColorARGB(CInt(map_letter_a), 179, 95, 0)
        temp_array(1) = D3DColorARGB(CInt(map_letter_a), 179, 95, 0)
        temp_array(2) = D3DColorARGB(CInt(map_letter_a), 179, 95, 0)
        temp_array(3) = D3DColorARGB(CInt(map_letter_a), 179, 95, 0)
        Grh_Render letter_grh, 250, 300, temp_array()
        Engine_Text_RenderGrande letter_text, 360 - Engine_Text_Width(letter_text, False, 4) / 2, 1, temp_array, 5, False, , CInt(map_letter_a)

    End If

    If FullScreen Then
        RenderConsola

    End If

End Sub

Private Function Convert_Tile_To_View_X(ByVal x As Integer) As Integer
    '**************************************************************
    'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
    'Last Modify Date: 10/07/2002
    'Convert tile position into position in view area
    '**************************************************************
    'If engine_windowed Then
    Convert_Tile_To_View_X = ((x - 1) * 32)

    ' Else
    '  Convert_Tile_To_View_X = view_screen_left + ((x - 1) * base_tile_size)
    '  End If
End Function

Private Function Convert_Tile_To_View_Y(ByVal y As Integer) As Integer
    '**************************************************************
    'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
    'Last Modify Date: 10/07/2002
    'Convert tile position into position in view area
    '**************************************************************
    ' If engine_windowed Then
    Convert_Tile_To_View_Y = ((y - 1) * 32)

    'Else
    '   Convert_Tile_To_View_Y = view_screen_top + ((y - 1) * base_tile_size)
    'End If
End Function

Private Function Geometry_Create_TLVertex(ByVal x As Single, ByVal y As Single, ByVal Z As Single, ByVal rhw As Single, ByVal color As Long, ByVal Specular As Long, tu As Single, ByVal tv As Single) As TLVERTEX
    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '**************************************************************
    Geometry_Create_TLVertex.x = x
    Geometry_Create_TLVertex.y = y
    Geometry_Create_TLVertex.Z = Z
    Geometry_Create_TLVertex.rhw = rhw
    Geometry_Create_TLVertex.color = color
    Geometry_Create_TLVertex.Specular = Specular
    Geometry_Create_TLVertex.tu = tu
    Geometry_Create_TLVertex.tv = tv

End Function

Private Function Geometry_Create_TLVertex2(x As Single, y As Single, Z As Single, rhw As Single, color As Long, Specular As Long, tu1 As Single, tv1 As Single, tu2 As Single, tv2 As Single) As TLVERTEX2
    'mz
    Geometry_Create_TLVertex2.x = x
    Geometry_Create_TLVertex2.y = y
    Geometry_Create_TLVertex2.Z = Z
    Geometry_Create_TLVertex2.rhw = rhw
    Geometry_Create_TLVertex2.color = color
    Geometry_Create_TLVertex2.Specular = Specular
    Geometry_Create_TLVertex2.tu1 = tu1
    Geometry_Create_TLVertex2.tv1 = tv1
    Geometry_Create_TLVertex2.tu2 = tu2
    Geometry_Create_TLVertex2.tv2 = tv2

End Function

Private Sub Geometry_Create_Box(ByRef verts() As TLVERTEX, ByRef dest As RECT, ByRef src As RECT, ByRef rgb_list() As Long, Optional ByRef Textures_Width As Long, Optional ByRef Textures_Height As Long, Optional ByVal angle As Single)

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
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.Left / Textures_Width, (src.bottom + 1) / Textures_Height)
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
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, (src.Right + 1) / Textures_Width, (src.bottom + 1) / Textures_Height)
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
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.Right + 1) / Textures_Width, src.Top / Textures_Height)
    Else
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 1, 1)

    End If

End Sub

Private Sub Device_Box_Textured_Render_Advance(ByVal GrhIndex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, ByVal src_height As Integer, ByRef rgb_list() As Long, ByVal src_x As Integer, ByVal src_y As Integer, ByVal dest_width As Integer, Optional ByVal dest_height As Integer, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 5/15/2003
    'Copies the Textures allowing resizing
    'Modified by Juan Martín Sotuyo Dodero
    '**************************************************************
    Static src_rect            As RECT

    Static dest_rect           As RECT

    Static temp_verts(3)       As TLVERTEX

    Static d3dTextures         As D3D8Textures

    Static light_value(0 To 3) As Long
    
    If GrhIndex = 0 Then Exit Sub
    
    Set d3dTextures.texture = SurfaceDB.GetTexture(GrhData(GrhIndex).FileNum, d3dTextures.texwidth, d3dTextures.texheight)
    
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)
        
    'If Not char_current_blind Then
    '    If (light_value(0) = 0) Then light_value(0) = 0
    '    If (light_value(1) = 0) Then light_value(1) = 0
    '    If (light_value(2) = 0) Then light_value(2) = 0
    '    If (light_value(3) = 0) Then light_value(3) = 0
    'Else
    '    light_value(0) = &HFFFFFFFF 'blind_color
    '    light_value(1) = &HFFFFFFFF 'blind_color
    '    light_value(2) = &HFFFFFFFF 'blind_color
    '    light_value(3) = &HFFFFFFFF 'blind_color
    'End If
    
    'Set up the source rectangle
    With src_rect
        .bottom = src_y + src_height
        .Left = src_x
        .Right = src_x + src_width
        .Top = src_y

    End With
        
    'Set up the destination rectangle
    With dest_rect
        .bottom = dest_y + dest_height
        .Left = dest_x
        .Right = dest_x + dest_width
        .Top = dest_y

    End With
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), d3dTextures.texwidth, d3dTextures.texheight, angle
        
    'Set Textures
    D3DDevice.SetTexture 0, d3dTextures.texture
    
    If alpha_blend Then
        'Set Rendering for alphablending
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE

    End If
    
    'Draw the triangles that make up our square Textures
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

    End If

End Sub

Public Sub Device_Box_Textured_Render(ByVal GrhIndex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, ByVal src_height As Integer, ByRef rgb_list() As Long, ByVal src_x As Integer, ByVal src_y As Integer, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 2/12/2004
    'Just copies the Textures
    '**************************************************************
    Static src_rect            As RECT

    Static dest_rect           As RECT

    Static temp_verts(3)       As TLVERTEX

    Static d3dTextures         As D3D8Textures

    Static light_value(0 To 3) As Long
    
    If GrhIndex = 0 Then Exit Sub
    
    Set d3dTextures.texture = SurfaceDB.GetTexture(GrhData(GrhIndex).FileNum, d3dTextures.texwidth, d3dTextures.texheight)
    
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)
 
    If (light_value(0) = 0) Then light_value(0) = map_base_light
    If (light_value(1) = 0) Then light_value(1) = map_base_light
    If (light_value(2) = 0) Then light_value(2) = map_base_light
    If (light_value(3) = 0) Then light_value(3) = map_base_light
        
    'Set up the source rectangle
    With src_rect
        .bottom = src_y + src_height
        .Left = src_x
        .Right = src_x + src_width
        .Top = src_y

    End With
                
    'Set up the destination rectangle
    With dest_rect
        .bottom = dest_y + src_height
        .Left = dest_x
        .Right = dest_x + src_width
        .Top = dest_y

    End With
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), d3dTextures.texwidth, d3dTextures.texheight, angle
     
    'Set Textures
    D3DDevice.SetTexture 0, d3dTextures.texture
    
    If alpha_blend Then
        'Set Rendering for alphablending
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE

    End If
    
    'Draw the triangles that make up our square Textures
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

    End If
    
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    D3DDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
 
    'D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_DIFFUSE
    'D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1

End Sub

Public Sub Engine_MoveScreen(ByVal nHeading As E_Heading)

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
        
        bTecho = IIf(MapData(UserPos.x, UserPos.y).Trigger = 1 Or MapData(UserPos.x, UserPos.y).Trigger = 2 Or MapData(UserPos.x, UserPos.y).Trigger = 6 Or MapData(UserPos.x, UserPos.y).Trigger > 9 Or MapData(UserPos.x, UserPos.y).Trigger = 4, True, False)

    End If

End Sub

Private Sub Char_TextRender(ByVal charindex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer, ByVal x As Byte, ByVal y As Byte)

    Dim moved         As Boolean

    Dim Pos           As Integer

    Dim line          As String

    Dim color(0 To 3) As Long

    Dim i             As Long
    
    Dim screen_x      As Integer

    Dim screen_y      As Integer
    
    ' screen_x = Convert_Tile_To_View_X(PixelOffsetX) + MoveOffsetX
    ' screen_y = Convert_Tile_To_View_Y(PixelOffsetY) +

    With charlist(charindex)

        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
        
        'screen_x = Convert_Tile_To_View_X(PixelOffsetX) + MoveOffsetX

        '*** Start Dialogs ***
        If charlist(MapData(x, y).charindex).dialog <> "" Then

            'Figure out screen position
            Dim temp_array(3) As Long

            Dim PixelY        As Integer

            PixelY = PixelOffsetY
            temp_array(0) = charlist(MapData(x, y).charindex).dialog_color
            temp_array(1) = temp_array(0)
            temp_array(2) = temp_array(0)
            temp_array(3) = temp_array(0)

            If charlist(MapData(x, y).charindex).dialog_scroll Then
                charlist(MapData(x, y).charindex).dialog_offset_counter_y = charlist(MapData(x, y).charindex).dialog_offset_counter_y + (scroll_dialog_pixels_per_frame * timerTicksPerFrame * Sgn(-1))

                If Sgn(charlist(MapData(x, y).charindex).dialog_offset_counter_y) = -1 Then
                    charlist(MapData(x, y).charindex).dialog_offset_counter_y = 0
                    charlist(MapData(x, y).charindex).dialog_scroll = False

                End If

                Engine_Text_Render charlist(MapData(x, y).charindex).dialog, PixelOffsetX + 14 - CInt(Engine_Text_Width(charlist(MapData(x, y).charindex).dialog, True) / 2), PixelY + charlist(MapData(x, y).charindex).Body.HeadOffset.y - Engine_Text_Height(charlist(MapData(x, y).charindex).dialog, True) + charlist(MapData(x, y).charindex).dialog_offset_counter_y, temp_array, 1, True, MapData(x, y).charindex
            Else
                Engine_Text_Render charlist(MapData(x, y).charindex).dialog, PixelOffsetX + 14 - CInt(Engine_Text_Width(charlist(MapData(x, y).charindex).dialog, True) / 2), PixelY + charlist(MapData(x, y).charindex).Body.HeadOffset.y - Engine_Text_Height(charlist(MapData(x, y).charindex).dialog, True), temp_array, 1, True, MapData(x, y).charindex

            End If

        End If
        
        If charlist(MapData(x, y).charindex).dialogEfec <> "" Then

            charlist(MapData(x, y).charindex).SubeEfecto = charlist(MapData(x, y).charindex).SubeEfecto - timerTicksPerFrame
            charlist(MapData(x, y).charindex).dialog_Efect_color.a = charlist(MapData(x, y).charindex).dialog_Efect_color.a - (timerTicksPerFrame * 8.2)

            If charlist(MapData(x, y).charindex).dialog_Efect_color.a < 0 Then
                charlist(MapData(x, y).charindex).SubeEfecto = 0
                charlist(MapData(x, y).charindex).dialogEfec = ""
            Else
                temp_array(0) = D3DColorARGB(.dialog_Efect_color.a, .dialog_Efect_color.r, .dialog_Efect_color.g, .dialog_Efect_color.b)
                temp_array(1) = temp_array(0)
                temp_array(2) = temp_array(0)
                temp_array(3) = temp_array(0)
        
                Engine_Text_Render_Efect MapData(x, y).charindex, .dialogEfec, PixelOffsetX + 14 - Engine_Text_Width(.dialogEfec, True) / 2, PixelOffsetY - 100 + .Body.HeadOffset.y - Engine_Text_Height(.dialogEfec, True) + .SubeEfecto, temp_array, 1, True, max(CDbl(charlist(MapData(x, y).charindex).dialog_Efect_color.a), 0)

            End If

        End If
            
        ' If charlist(MapData(X, Y).charindex).dialogExp <> "" Then
    
        '  charlist(MapData(X, Y).charindex).SubeExp = charlist(MapData(X, Y).charindex).SubeExp + (5 * timerTicksPerFrame * Sgn(-1))
        ' If charlist(MapData(X, Y).charindex).SubeExp <= 5 Then
        '   charlist(MapData(X, Y).charindex).SubeExp = 0
        '  charlist(MapData(X, Y).charindex).dialogExp = ""
        'End If
                    
        'temp_array(0) = D3DColorARGB(charlist(MapData(X, Y).charindex).SubeExp, 42, 169, 222)
        ' temp_array(1) = temp_array(0)
        ' temp_array(2) = temp_array(0)
        ' temp_array(3) = temp_array(0)
        'Engine_Text_Render_Exp MapData(X, Y).charindex, .dialogExp, PixelOffsetX + 14 - Engine_Text_Width(.dialogExp, True) / 2, PixelOffsetY + 14 + .Body.HeadOffset.Y - Engine_Text_Height(.dialogExp, True), temp_array, 1, True
        ' End If
            
        'If charlist(MapData(X, Y).charindex).dialogOro <> "" Then

        '  charlist(MapData(X, Y).charindex).SubeOro = charlist(MapData(X, Y).charindex).SubeOro + (5 * timerTicksPerFrame * Sgn(-1))
                
        'If charlist(MapData(X, Y).charindex).SubeOro <= 5 Then
        '    charlist(MapData(X, Y).charindex).SubeOro = 0
        '    charlist(MapData(X, Y).charindex).dialogOro = ""
        'End If
                
        ' temp_array(0) = D3DColorARGB(charlist(MapData(X, Y).charindex).SubeOro, 255, 255, 115)
        ' temp_array(1) = temp_array(0)
        ' temp_array(2) = temp_array(0)
        ' temp_array(3) = temp_array(0)
        ' Engine_Text_Render_Exp MapData(X, Y).charindex, .dialogOro, PixelOffsetX + 14 - Engine_Text_Width(.dialogOro, True) / 2, PixelOffsetY + 1 + .Body.HeadOffset.Y - Engine_Text_Height(.dialogOro, True), temp_array, 1, True
                
        '  End If
        '*** End Dialogs ***
    End With

End Sub

Private Sub Char_Render(ByVal charindex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer, ByVal x As Byte, ByVal y As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 12/03/04
    'Draw char's to screen without offcentering them
    '***************************************************
    Dim moved                As Boolean

    Dim Pos                  As Integer

    Dim line                 As String

    Dim color(0 To 3)        As Long

    Dim colorCorazon(0 To 3) As Long

    Dim i                    As Long

    Dim OffsetYname          As Byte

    Dim OffsetYClan          As Byte
    
    Dim OffArma              As Byte
    
    With charlist(charindex)

        If .Heading = 0 Then Exit Sub
    
        If .Moving Then

            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame * .Speeding
                
                'Start animations
                'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).speed > 0 Then .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                .MovArmaEscudo = False
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0

                End If

            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame * .Speeding
                
                'Start animations
                If .Body.Walk(.Heading).speed > 0 Then .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                .MovArmaEscudo = False
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0

                End If

            End If

        End If
        
        'If done moving stop animation
        If Not moved Then
            'Stop animations
            .Body.Walk(.Heading).Started = 0
            .Body.Walk(.Heading).framecounter = 1
            
            If Not .MovArmaEscudo Then
                .Arma.WeaponWalk(.Heading).Started = 0
                .Arma.WeaponWalk(.Heading).framecounter = 1

                .Escudo.ShieldWalk(.Heading).Started = 0
                .Escudo.ShieldWalk(.Heading).framecounter = 1

            End If
            
            .Moving = False

        End If

        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
 
        If .EsNpc Then
            If Len(.nombre) > 0 Then
                If Abs(tX - .Pos.x) < 1 And (Abs(tY - .Pos.y)) < 1 Then

                    Dim colornpcs(3) As Long

                    colornpcs(0) = D3DColorXRGB(0, 129, 195)
                    colornpcs(1) = colornpcs(0)
                    colornpcs(2) = colornpcs(0)
                    colornpcs(3) = colornpcs(0)
                    Pos = InStr(.nombre, "<")

                    If Pos = 0 Then Pos = Len(.nombre) + 2
                    'Nick
                    line = Left$(.nombre, Pos - 2)
                    Engine_Text_Render line, PixelOffsetX + 16 - CInt(Engine_Text_Width(line, True) / 2), PixelOffsetY + 30 - Engine_Text_Height(line, True), colornpcs, 1, True
                        
                    'Clan
                    line = mid$(.nombre, Pos)
                    Engine_Text_Render line, PixelOffsetX + 16 - CInt(Engine_Text_Width(line, True) / 2), PixelOffsetY + 45 - Engine_Text_Height(line, True), colornpcs, 1, True

                End If
                    
                If .simbolo <> 0 Then
                    ' Dim simbolo As grh
                    ' simbolo.framecounter = 1
                    ' simbolo.GrhIndex = 5259 + .simbolo
                    'Call Draw_Grh(TempGrh, PixelOffsetX + 20, PixelOffsetY - 45, 1, 0, colorz, False, 0, 0, 0)
                    Call engine.Draw_GrhIndex(5259 + .simbolo, PixelOffsetX + 6, PixelOffsetY + .Body.HeadOffset.y - 10)
                                
                    ' Debug.Print .simbolo
                End If

            End If

        End If

        colornpcs(0) = D3DColorXRGB(255, 255, 255)
        'line = "me gusta el vino, me quiero casar con tu hermana. pero no se si vos qu e"
        ' Engine_Text_Render line, PixelOffsetX + 16 - Engine_Text_Width(line, True), PixelOffsetY + 30 - Engine_Text_Height(line, True), colornpcs, 1, True
        
        If .Body.Walk(.Heading).GrhIndex Then

            If Not .invisible Then

                Dim colorz(3) As Long

                'Draw Body

                colorz(0) = MapData(x, y).light_value(0)
                colorz(1) = MapData(x, y).light_value(1)
                colorz(2) = MapData(x, y).light_value(2)
                colorz(3) = MapData(x, y).light_value(3)
                
                If .EsEnano Then OffArma = 7
                                
                If .Body_Aura <> "" Then Call Renderizar_Aura(.Body_Aura, PixelOffsetX, PixelOffsetY + OffArma, x, y, charindex)
                If .Arma_Aura <> "" Then Call Renderizar_Aura(.Arma_Aura, PixelOffsetX, PixelOffsetY + OffArma, x, y, charindex)
                If .Otra_Aura <> "" Then Call Renderizar_Aura(.Otra_Aura, PixelOffsetX, PixelOffsetY + OffArma, x, y, charindex)
                If .Escudo_Aura <> "" Then Call Renderizar_Aura(.Escudo_Aura, PixelOffsetX, PixelOffsetY + OffArma, x, y, charindex)
                                
                Select Case .Heading

                    Case EAST

                        If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y)
                                                                    
                        If .iBody < 488 Then
                            Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX + 1, PixelOffsetY, 1, 1, colorz, False, x, y, 0)
                        Else
                            Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y, 0)

                        End If
                                         
                        If .Head.Head(.Heading).GrhIndex Then Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, colorz, False, x, y)
                                         
                        If .Casco.Head(.Heading).GrhIndex Then Call Draw_Grh(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, colorz, False, x, y)
                                             
                        If .Arma.WeaponWalk(.Heading).GrhIndex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY + OffArma, 1, 1, colorz, False, x, y)
                                     
                    Case NORTH

                        If .Arma.WeaponWalk(.Heading).GrhIndex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY + OffArma, 1, 1, colorz, False, x, y)
                                             
                        If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y)
                                             
                        If .iBody < 488 Then
                            Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX + 1, PixelOffsetY, 1, 1, colorz, False, x, y, 0)
                        Else
                            Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y, 0)

                        End If
                                         
                        If .Head.Head(.Heading).GrhIndex Then Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, colorz, False, x, y)
                                         
                        If .Casco.Head(.Heading).GrhIndex Then Call Draw_Grh(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, colorz, False, x, y)
                                     
                    Case WEST

                        If .Arma.WeaponWalk(.Heading).GrhIndex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y)
                                             
                        If .iBody < 488 Then
                            Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX + 1, PixelOffsetY, 1, 1, colorz, False, x, y, 0)
                        Else
                            Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y, 0)

                        End If
                                         
                        If .Head.Head(.Heading).GrhIndex Then Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, colorz, False, x, y)
                                         
                        If .Casco.Head(.Heading).GrhIndex Then Call Draw_Grh(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, colorz, False, x, y)
                                         
                        If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY + OffArma, 1, 1, colorz, False, x, y)

                    Case south
                                         
                        If .iBody < 488 Then
                            Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX + 1, PixelOffsetY, 1, 1, colorz, False, x, y, 0)
                        Else
                            Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y, 0)

                        End If
                                         
                        If .Head.Head(.Heading).GrhIndex Then Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, colorz, False, x, y)
                                         
                        If .Casco.Head(.Heading).GrhIndex Then Call Draw_Grh(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, colorz, False, x, y)
                                         
                        If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY + OffArma, 1, 1, colorz, False, x, y)
                                             
                        If .Arma.WeaponWalk(.Heading).GrhIndex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY + OffArma, 1, 1, colorz, False, x, y)

                End Select

                'Draw name over head
                '  If .transformado = False Then
                If Nombres Then
                    If Len(.nombre) > 0 And Not .EsNpc Then
                        Pos = InStr(.nombre, "<")

                        If Pos = 0 Then Pos = Len(.nombre) + 2
                        If .priv = 0 Then
                                
                            Select Case .status

                                Case 0
                                    color(0) = RGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b)
                                    color(1) = RGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b)
                                    color(2) = RGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b)
                                    color(3) = RGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b)
                                    colorCorazon(0) = D3DColorXRGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b)
                                    colorCorazon(1) = D3DColorXRGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b)
                                    colorCorazon(2) = D3DColorXRGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b)
                                    colorCorazon(3) = D3DColorXRGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b)

                                Case 1
                                    color(0) = RGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b)
                                    color(1) = RGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b)
                                    color(2) = RGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b)
                                    color(3) = RGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b)
                                    colorCorazon(0) = D3DColorXRGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b)
                                    colorCorazon(1) = D3DColorXRGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b)
                                    colorCorazon(2) = D3DColorXRGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b)
                                    colorCorazon(3) = D3DColorXRGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b)

                                Case 2
                                    color(0) = RGB(ColoresPJ(6).r, ColoresPJ(6).g, ColoresPJ(6).b)
                                    color(1) = RGB(ColoresPJ(6).r, ColoresPJ(6).g, ColoresPJ(6).b)
                                    color(2) = RGB(ColoresPJ(6).r, ColoresPJ(6).g, ColoresPJ(6).b)
                                    color(3) = RGB(ColoresPJ(6).r, ColoresPJ(6).g, ColoresPJ(6).b)
                                    colorCorazon(0) = D3DColorXRGB(ColoresPJ(6).r, ColoresPJ(6).g, ColoresPJ(6).b)
                                    colorCorazon(1) = D3DColorXRGB(ColoresPJ(6).r, ColoresPJ(6).g, ColoresPJ(6).b)
                                    colorCorazon(2) = D3DColorXRGB(ColoresPJ(6).r, ColoresPJ(6).g, ColoresPJ(6).b)
                                    colorCorazon(3) = D3DColorXRGB(ColoresPJ(6).r, ColoresPJ(6).g, ColoresPJ(6).b)

                                Case 3
                                    color(0) = RGB(ColoresPJ(7).r, ColoresPJ(7).g, ColoresPJ(7).b)
                                    color(1) = RGB(ColoresPJ(7).r, ColoresPJ(7).g, ColoresPJ(7).b)
                                    color(2) = RGB(ColoresPJ(7).r, ColoresPJ(7).g, ColoresPJ(7).b)
                                    color(3) = RGB(ColoresPJ(7).r, ColoresPJ(7).g, ColoresPJ(7).b)
                                    colorCorazon(0) = D3DColorXRGB(ColoresPJ(7).r, ColoresPJ(7).g, ColoresPJ(7).b)
                                    colorCorazon(1) = D3DColorXRGB(ColoresPJ(7).r, ColoresPJ(7).g, ColoresPJ(7).b)
                                    colorCorazon(2) = D3DColorXRGB(ColoresPJ(7).r, ColoresPJ(7).g, ColoresPJ(7).b)
                                    colorCorazon(3) = D3DColorXRGB(ColoresPJ(7).r, ColoresPJ(7).g, ColoresPJ(7).b)

                            End Select
                                    
                        Else
                            color(0) = RGB(ColoresPJ(.priv).r, ColoresPJ(.priv).g, ColoresPJ(.priv).b)
                            color(1) = RGB(ColoresPJ(.priv).r, ColoresPJ(.priv).g, ColoresPJ(.priv).b)
                            color(2) = RGB(ColoresPJ(.priv).r, ColoresPJ(.priv).g, ColoresPJ(.priv).b)
                            color(3) = RGB(ColoresPJ(.priv).r, ColoresPJ(.priv).g, ColoresPJ(.priv).b)
                            colorCorazon(0) = D3DColorXRGB(ColoresPJ(.priv).r, ColoresPJ(.priv).g, ColoresPJ(.priv).b)
                            colorCorazon(1) = D3DColorXRGB(ColoresPJ(.priv).r, ColoresPJ(.priv).g, ColoresPJ(.priv).b)
                            colorCorazon(2) = D3DColorXRGB(ColoresPJ(.priv).r, ColoresPJ(.priv).g, ColoresPJ(.priv).b)
                            colorCorazon(3) = D3DColorXRGB(ColoresPJ(.priv).r, ColoresPJ(.priv).g, ColoresPJ(.priv).b)
                                    
                        End If
                                            
                        If .group_index > 0 Then
                            If charlist(charindex).group_index = charlist(UserCharIndex).group_index Then
                                color(0) = D3DColorXRGB(255, 255, 255)
                                color(1) = D3DColorXRGB(255, 255, 255)
                                color(2) = D3DColorXRGB(255, 255, 255)
                                color(3) = D3DColorXRGB(255, 255, 255)
                                colorCorazon(0) = D3DColorXRGB(255, 255, 0)
                                colorCorazon(1) = D3DColorXRGB(0, 255, 255)
                                colorCorazon(2) = D3DColorXRGB(0, 255, 0)
                                colorCorazon(3) = D3DColorXRGB(0, 255, 255)

                            End If

                        End If

                        If FullScreen And charindex = UserCharIndex And UserEstado = 0 Then
                            OffsetYname = 16
                            OffsetYClan = 14
                            line = Left$(.nombre, Pos - 2)
                            Grh_Render Marco, PixelOffsetX, PixelOffsetY + 5, colorz, True, True, False
                            Draw_FilledBox PixelOffsetX + 3, PixelOffsetY + 31, (((UserMinHp + 1 / 100) / (UserMaxHp + 1 / 100))) * 26, 4, D3DColorARGB(255, 200, 0, 0), D3DColorARGB(0, 200, 200, 200)
                            Grh_Render Marco, PixelOffsetX, PixelOffsetY + 14, colorz, True, True, False
                            Draw_FilledBox PixelOffsetX + 3, PixelOffsetY + 40, (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100))) * 26, 4, D3DColorARGB(255, 0, 0, 255), D3DColorARGB(0, 200, 200, 200)

                        End If
                            
                        If .clan_index > 0 Then
                            If .clan_index = charlist(UserCharIndex).clan_index And charindex <> UserCharIndex And .MUERTO = 0 Then
                                If .clan_nivel = 5 Then
                                    OffsetYname = 8
                                    OffsetYClan = 6
                                    Grh_Render Marco, PixelOffsetX, PixelOffsetY + 5, colorz, True, True, False
                                    Draw_FilledBox PixelOffsetX + 3, PixelOffsetY + 31, (((.UserMinHp + 1 / 100) / (.UserMaxHp + 1 / 100))) * 26, 4, D3DColorARGB(255, 200, 0, 0), D3DColorARGB(0, 200, 200, 200)

                                End If

                            End If

                        End If
  
                        'Nick
                        line = Left$(.nombre, Pos - 2)
                        Engine_Text_Render line, PixelOffsetX + 15 - CInt(Engine_Text_Width(line, True) / 2), PixelOffsetY + 30 + OffsetYname - Engine_Text_Height(line, True), color, 1, True
                        
                        'Clan
                        Select Case .priv

                            Case 1
                                line = "<Game Design>"

                            Case 2
                                line = "<Game Master>"

                            Case 3, 4
                                line = "<Administrador>"

                            Case Else
                                line = mid$(.nombre, Pos)

                        End Select
                            
                        Engine_Text_Render line, PixelOffsetX + 15 - CInt(Engine_Text_Width(line, True) / 2), PixelOffsetY + 45 + OffsetYClan - Engine_Text_Height(line, True), color, 1, True

                        If .Donador = 1 Then
                            line = Left$(.nombre, Pos - 2)
                            Grh_Render Estrella, PixelOffsetX + 7 + CInt(Engine_Text_Width(line, 1) / 2), PixelOffsetY + 10 + OffsetYname, colorCorazon, True, True, False

                        End If

                    End If

                End If

                ' End If
            Else
            
                Dim mostrarlo As Boolean
                         
                If .priv < charlist(UserCharIndex).priv Then
                    mostrarlo = True

                End If

                If .group_index > 0 Then
                    If charlist(charindex).group_index = charlist(UserCharIndex).group_index Then
                        mostrarlo = True

                    End If

                End If

                If .clan_index > 0 Then
                    If .clan_index = charlist(UserCharIndex).clan_index Then
                        If .clan_nivel >= 3 Then
                            mostrarlo = True

                        End If

                    End If

                End If
                    
                If charindex = UserCharIndex Or mostrarlo = True Then
                    colorz(0) = D3DColorARGB(100, 255, 255, 255)
                    colorz(1) = D3DColorARGB(100, 255, 255, 255)
                    colorz(2) = D3DColorARGB(100, 255, 255, 255)
                    colorz(3) = D3DColorARGB(100, 255, 255, 255)

                    If .Body.Walk(.Heading).GrhIndex Then Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y)

                    If .Head.Head(.Heading).GrhIndex Then Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, colorz, False, x, y)

                    If .Casco.Head(.Heading).GrhIndex Then Call Draw_Grh(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, colorz, False, x, y)

                    If .Arma.WeaponWalk(.Heading).GrhIndex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y)

                    If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y)
                                
                    Pos = InStr(.nombre, "<")

                    If Pos = 0 Then Pos = Len(.nombre) + 2

                    color(0) = D3DColorXRGB(255, 255, 255)
                    color(1) = color(0)
                    color(2) = color(0)
                    color(3) = color(0)
                    colorCorazon(0) = D3DColorXRGB(120, 100, 200)
                    colorCorazon(1) = colorCorazon(0)
                    colorCorazon(2) = colorCorazon(0)
                    colorCorazon(3) = colorCorazon(0)
                                
                    If FullScreen And charindex = UserCharIndex And UserEstado = 0 Then
                        OffsetYname = 16
                        OffsetYClan = 14
                        line = Left$(.nombre, Pos - 2)
                        Grh_Render Marco, PixelOffsetX, PixelOffsetY + 5, color, True, True, False
                        Draw_FilledBox PixelOffsetX + 3, PixelOffsetY + 31, (((UserMinHp + 1 / 100) / (UserMaxHp + 1 / 100))) * 26, 4, D3DColorARGB(255, 200, 0, 0), D3DColorARGB(0, 200, 200, 200)
                        Grh_Render Marco, PixelOffsetX, PixelOffsetY + 14, color, True, True, False
                        Draw_FilledBox PixelOffsetX + 3, PixelOffsetY + 40, (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100))) * 26, 4, D3DColorARGB(255, 0, 0, 255), D3DColorARGB(0, 200, 200, 200)

                    End If
                                
                    color(0) = D3DColorXRGB(200, 100, 100)
                    color(1) = color(0)
                    color(2) = color(0)
                    color(3) = color(0)

                    line = Left$(.nombre, Pos - 2)
                    Engine_Text_Render line, PixelOffsetX + 15 - CInt(Engine_Text_Width(line, True) / 2), PixelOffsetY + 30 + OffsetYname - Engine_Text_Height(line, True), color, 1, True
                        
                    'Clan
                    Select Case .priv

                        Case 1
                            line = "<Game Design>"

                        Case 2
                            line = "<Game Master>"

                        Case 3, 4
                            line = "<Administrador>"

                        Case Else
                            line = mid$(.nombre, Pos)

                    End Select
                            
                    Engine_Text_Render line, PixelOffsetX + 15 - CInt(Engine_Text_Width(line, True) / 2), PixelOffsetY + 45 + OffsetYClan - Engine_Text_Height(line, True), color, 1, True

                    If .Donador = 1 Then
                        line = Left$(.nombre, Pos - 2)
                        Grh_Render Estrella, PixelOffsetX + 7 + CInt(Engine_Text_Width(line, 1) / 2), PixelOffsetY + 10 + OffsetYname, colorCorazon, True, True, False

                    End If

                Else

                    If .TimerI <= 0 Then .TimerIAct = True
                    If .TimerIAct = False Then
                        .TimerI = .TimerI - (timerTicksPerFrame * 1)
                    Else
                        .TimerI = .TimerI + (timerTicksPerFrame * 0.3)

                        If .TimerI >= 40 Then .TimerIAct = False

                    End If

                    colorz(0) = D3DColorARGB(.TimerI, 255, 255, 255)
                    colorz(1) = D3DColorARGB(.TimerI, 255, 255, 255)
                    colorz(2) = D3DColorARGB(.TimerI, 255, 255, 255)
                    colorz(3) = D3DColorARGB(.TimerI, 255, 255, 255)

                    If .Body.Walk(.Heading).GrhIndex Then Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y)

                    If .Head.Head(.Heading).GrhIndex Then Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, colorz, False, x, y)

                    If .Casco.Head(.Heading).GrhIndex Then Call Draw_Grh(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, colorz, False, x, y)

                    If .Arma.WeaponWalk(.Heading).GrhIndex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y)

                    If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y)

                End If

            End If
        
            If .particle_count > 0 Then

                For i = 1 To .particle_count

                    If .particle_group(i) > 0 Then
                        Particle_Group_Render .particle_group(i), PixelOffsetX + .Body.HeadOffset.x + (32 / 2), PixelOffsetY

                    End If

                Next i

            End If
    
            'Barra de tiempo
            If .BarTime < .MaxBarTime Then
                Draw_FilledBox PixelOffsetX - 17, PixelOffsetY - 40, 70, 7, D3DColorARGB(100, 0, 0, 0), D3DColorARGB(100, 0, 0, 0)
                Draw_FilledBox PixelOffsetX - 17, PixelOffsetY - 40, (((.BarTime / 100) / (.MaxBarTime / 100))) * 69, 7, D3DColorARGB(100, 200, 0, 0), D3DColorARGB(1, 200, 200, 200)
                .BarTime = .BarTime + (4 * timerTicksPerFrame * Sgn(1))
                                 
                If .BarTime >= .MaxBarTime Then
                    If charindex = UserCharIndex Then
                        Call CompletarAccionBarra(.BarAccion)

                    End If

                    charlist(charindex).BarTime = 0
                    charlist(charindex).BarAccion = 99
                    charlist(charindex).MaxBarTime = 0

                End If

            End If
                            
            If .Escribiendo = True And Not .invisible Then

                Dim TempGrh As grh

                TempGrh.framecounter = 1
                TempGrh.GrhIndex = 32017
                colorz(0) = D3DColorARGB(200, 255, 255, 255)
                colorz(1) = D3DColorARGB(200, 255, 255, 255)
                colorz(2) = D3DColorARGB(200, 255, 255, 255)
                colorz(3) = D3DColorARGB(200, 255, 255, 255)
                Call Draw_Grh(TempGrh, PixelOffsetX + 20, PixelOffsetY - 45, 1, 0, colorz, False, 0, 0, 0)

            End If
                             
            If .FxCount > 0 Then

                For i = 1 To .FxCount

                    If .FxList(i).FxIndex > 0 And .FxList(i).Started <> 0 Then
                        colorz(0) = D3DColorARGB(220, 255, 255, 255)
                        colorz(1) = D3DColorARGB(220, 255, 255, 255)
                        colorz(2) = D3DColorARGB(220, 255, 255, 255)
                        colorz(3) = D3DColorARGB(220, 255, 255, 255)

                        If FxData(.FxList(i).FxIndex).IsPNG = 1 Then
                            Call Draw_GrhFX(.FxList(i), PixelOffsetX + FxData(.FxList(i).FxIndex).OffsetX, PixelOffsetY + FxData(.FxList(i).FxIndex).OffsetY + 20, 1, 1, colorz, False, , , , charindex)
                        Else
                            Call Draw_GrhFX(.FxList(i), PixelOffsetX + FxData(.FxList(i).FxIndex).OffsetX, PixelOffsetY + FxData(.FxList(i).FxIndex).OffsetY + 20, 1, 1, colorz, True, , , , charindex)

                        End If

                    End If

                    If .FxList(i).Started = 0 Then
                        .FxList(i).FxIndex = 0

                    End If

                Next i

                If .FxList(.FxCount).Started = 0 Then
                    .FxCount = .FxCount - 1

                End If

            End If
            
            ' Meditación
            If .FxIndex <> 0 And .fX.Started <> 0 Then
                colorz(0) = D3DColorARGB(180, 255, 255, 255)
                colorz(1) = D3DColorARGB(180, 255, 255, 255)
                colorz(2) = D3DColorARGB(180, 255, 255, 255)
                colorz(3) = D3DColorARGB(180, 255, 255, 255)

                Call Draw_GrhFX(.fX, PixelOffsetX + FxData(.FxIndex).OffsetX, PixelOffsetY + FxData(.FxIndex).OffsetY + 4, 1, 1, colorz, False, , , , charindex)
           
            End If

        End If

    End With

End Sub

Private Sub Char_RenderCiego(ByVal charindex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer, ByVal x As Byte, ByVal y As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 12/03/04
    'Draw char's to screen without offcentering them
    '***************************************************
    Dim moved         As Boolean

    Dim Pos           As Integer

    Dim line          As String

    Dim color(0 To 3) As Long

    Dim i             As Long
    
    With charlist(charindex)

        If .Heading = 0 Then Exit Sub
    
        If .Moving Then

            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame * .Speeding
                
                'Start animations
                'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).speed > 0 Then .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0

                End If

            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame * .Speeding
                
                'Start animations
                If .Body.Walk(.Heading).speed > 0 Then .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0

                End If

            End If

        End If
        
        'If done moving stop animation
        If Not moved Then
            'Stop animations
            .Body.Walk(.Heading).Started = 0
            .Body.Walk(.Heading).framecounter = 1
            
            .Escudo.ShieldWalk(.Heading).Started = 0
            .Escudo.ShieldWalk(.Heading).framecounter = 1
            
            .Moving = False

        End If

        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
        
        Dim ColorCiego(0 To 3) As Long

        ColorCiego(0) = D3DColorARGB(255, 30, 30, 30)
        ColorCiego(1) = ColorCiego(0)
        ColorCiego(2) = ColorCiego(0)
        ColorCiego(3) = ColorCiego(0)
        
        If .Body.Walk(.Heading).GrhIndex Then
        
            If Not .invisible Then
 
                Dim colorz(3) As Long

                'Draw Body
                If .MUERTO = True Then
                    If .TimerM = 0 Then .TimerAct = True
                    If .TimerAct = False Then
                        .TimerM = .TimerM - 2
                    Else
                        .TimerM = .TimerM + 2

                        If .TimerM = 254 Then .TimerAct = False

                    End If
                    
                    colorz(0) = ColorCiego(0)
                    colorz(1) = ColorCiego(0)
                    colorz(2) = ColorCiego(0)
                    colorz(3) = ColorCiego(0)
                    
                Else
                    colorz(0) = ColorCiego(0)
                    colorz(1) = ColorCiego(0)
                    colorz(2) = ColorCiego(0)
                    colorz(3) = ColorCiego(0)

                End If
                        
                If .Body_Aura <> "" Then Call Renderizar_AuraCiego(.Body_Aura, PixelOffsetX, PixelOffsetY, x, y)
                If .Arma_Aura <> "" Then Call Renderizar_AuraCiego(.Arma_Aura, PixelOffsetX, PixelOffsetY, x, y)
                If .Otra_Aura <> "" Then Call Renderizar_AuraCiego(.Otra_Aura, PixelOffsetX, PixelOffsetY, x, y)
                If .Escudo_Aura <> "" Then Call Renderizar_AuraCiego(.Escudo_Aura, PixelOffsetX, PixelOffsetY, x, y)

                If .Heading = EAST Or .Heading = NORTH Then
                    If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y)

                    If .Body.Walk(.Heading).GrhIndex Then
                        Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y, 0)

                    End If

                Else

                    If .Heading = WEST Then
                        If .Arma.WeaponWalk(.Heading).GrhIndex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y)

                    End If

                    If .Body.Walk(.Heading).GrhIndex Then
                        If .iBody < 488 Then
                            Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX + 1, PixelOffsetY, 1, 1, colorz, False, x, y, 0)
                        Else
                            Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y, 0)

                        End If

                    End If
                            
                    If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y)

                End If
                            
                If .Head.Head(.Heading).GrhIndex Then

                    Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, colorz, False, x, y)
                    'Else
                    ' Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X - 1, PixelOffsetY + .Body.HeadOffset.Y, 1, 0, colorz, False, X, Y)
                                
                    ' End If
                End If

                If .Casco.Head(.Heading).GrhIndex Then Call Draw_Grh(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, colorz, False, x, y)
                            
                If .Heading <> WEST Then
                    If .Arma.WeaponWalk(.Heading).GrhIndex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y)

                End If

                'Draw name over head
                '  If .transformado = False Then
                    
                ' End If
            Else

                If charindex = UserCharIndex Or charlist(UserCharIndex).priv > 0 And .priv >= 0 Then
                    colorz(0) = ColorCiego(0)
                    colorz(1) = ColorCiego(0)
                    colorz(2) = ColorCiego(0)
                    colorz(3) = ColorCiego(0)

                    If .Body.Walk(.Heading).GrhIndex Then Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y)

                    If .Head.Head(.Heading).GrhIndex Then Call Draw_Grh(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, colorz, False, x, y)

                    If .Casco.Head(.Heading).GrhIndex Then Call Draw_Grh(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y, 1, 0, colorz, False, x, y)

                    If .Arma.WeaponWalk(.Heading).GrhIndex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y)

                    If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, colorz, False, x, y)
          
                End If

            End If
        
            If .particle_count > 0 Then

                For i = 1 To .particle_count

                    If .particle_group(i) > 0 Then
            
                        Particle_Group_Render .particle_group(i), PixelOffsetX + .Body.HeadOffset.x + (32 / 2), PixelOffsetY

                    End If

                Next i

            End If
    
            'Barra de tiempo
            If .BarTime < .MaxBarTime Then
                Draw_FilledBox PixelOffsetX - 17, PixelOffsetY - 40, 70, 7, D3DColorARGB(100, 0, 0, 0), D3DColorARGB(100, 0, 0, 0)
                Draw_FilledBox PixelOffsetX - 17, PixelOffsetY - 40, (((.BarTime / 100) / (.MaxBarTime / 100))) * 69, 7, D3DColorARGB(100, 200, 0, 0), D3DColorARGB(1, 200, 200, 200)
                .BarTime = .BarTime + (4 * timerTicksPerFrame * Sgn(1))
                                 
                '  Engine_Text_Render "time: " & .BarTime, 50, 50, color, 1, True
                If .BarTime >= .MaxBarTime And charindex = UserCharIndex Then
                    Call CompletarAccionBarra(.BarAccion)
                                
                End If

            End If
                            
            If .Escribiendo = True Then
                            
                Dim cartelito(0 To 3) As Long

                Dim rgb_list(3)       As Long

                Dim TempGrh           As grh

                TempGrh.framecounter = 1
                TempGrh.GrhIndex = 32017
                cartelito(0) = D3DColorARGB(200, 255, 255, 255)
                cartelito(1) = rgb_list(0)
                cartelito(2) = rgb_list(0)
                cartelito(3) = rgb_list(0)
    
                Call Draw_Grh(TempGrh, PixelOffsetX + 20, PixelOffsetY - 45, 1, 0, cartelito, False, 0, 0, 0)
                            
            End If
          
            'Draw FX

            If .FxIndex <> 0 And .fX.Started <> 0 Then
                colorz(0) = D3DColorARGB(220, 255, 255, 255)
                colorz(1) = D3DColorARGB(220, 255, 255, 255)
                colorz(2) = D3DColorARGB(220, 255, 255, 255)
                colorz(3) = D3DColorARGB(220, 255, 255, 255)

                If FxData(.FxIndex).IsPNG = 1 Then
                    Call Draw_GrhFX(.fX, PixelOffsetX + FxData(.FxIndex).OffsetX, PixelOffsetY + FxData(.FxIndex).OffsetY + 20, 1, 1, colorz, False, , , , charindex)
                Else
                    Call Draw_GrhFX(.fX, PixelOffsetX + FxData(.FxIndex).OffsetX, PixelOffsetY + FxData(.FxIndex).OffsetY + 20, 1, 1, colorz, True, , , , charindex)

                End If
                    
                If .fX.Started = 0 Then .FxIndex = 0
           
            End If
        
        End If

    End With

End Sub

Public Sub Start()

    DoEvents

    Do While prgRun

        Call FlushBuffer

        If frmmain.WindowState <> vbMinimized Then
            
            Select Case QueRender

                Case 0
                    engine.render
                
                    Check_Keys
                    Moviendose = False
                    engine.DrawMainInventory

                    If HayFormularioAbierto Then
                        If frmComerciar.Visible Then
                            engine.DrawInterfaceComerciar
                    
                        ElseIf frmBancoObj.Visible Then
                            engine.DrawInterfaceBoveda

                        End If

                        If frmMapaGrande.Visible Then
                            DrawMapaMundo

                        End If

                    End If

                Case 1
                    engine.RenderConnect 48, 49, 0, 0

                    If Not frmConnect.Visible Then
                        frmConnect.Show
                        FrmLogear.Show , frmConnect
                        FrmLogear.Top = FrmLogear.Top + 3500

                    End If

                Case 2
                    engine.rendercuenta 42, 43, 0, 0

                Case 3
                    engine.RenderCrearPJ 13, 67, 0, 0

            End Select

            Sound.Sound_Render
        Else
            Sleep 60&
            Call frmmain.Inventario.ReDraw

        End If

        DoEvents
        Rem Limitar FPS
        '   If LimitarFps Then
        '   While (GetTickCount - lFrameLimiter) < FramesPerSecCounter
        '     Sleep 1
        '  Wend
        '   While GetTickCount - lFrameLimiter < 55
        '     Sleep 5
        '   Wend
        '  End If
    Loop '

    EngineRun = False

    Call CloseClient

End Sub

Public Sub SetCharacterFx(ByVal charindex As Integer, ByVal fX As Integer, ByVal Loops As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 12/03/04
    'Sets an FX to the character.
    '***************************************************
    Dim indice As Byte

    With charlist(charindex)
        indice = engine.Char_FX_Group_Next_Open(charindex)
        .FxList(indice).FxIndex = fX
        Call InitGrh(.FxList(indice), FxData(fX).Animacion)
        .FxList(indice).Loops = Loops

    End With

End Sub

Public Sub SetMapFx(ByVal x As Byte, ByVal y As Byte, ByVal fX As Integer, ByVal Loops As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 12/03/04
    'Sets an FX to the character.
    '***************************************************

    On Error Resume Next

    Dim indice As Byte

    With MapData(x, y)
    
        indice = engine.Map_FX_Group_Next_Open(x, y)
    
        .FxList(indice).FxIndex = fX
        Call InitGrh(.FxList(indice), FxData(fX).Animacion)
        .FxList(indice).Loops = Loops
    
    End With

End Sub

Public Sub Char_Move_by_Head(ByVal charindex As Integer, ByVal nHeading As E_Heading)
    '*****************************************************************
    'Starts the movement of a character in nHeading direction
    '*****************************************************************

    If nHeading = 0 Then
        Debug.Print "Heading: " & nHeading

    End If

    On Error Resume Next

    Dim addx As Integer

    Dim addy As Integer

    Dim x    As Integer

    Dim y    As Integer

    Dim nX   As Integer

    Dim nY   As Integer
    
    With charlist(charindex)
        x = .Pos.x
        y = .Pos.y
        
        'Figure out which way to move
        Select Case nHeading

            Case E_Heading.NORTH
                addy = -1
        
            Case E_Heading.EAST
                addx = 1
        
            Case E_Heading.south
                addy = 1
            
            Case E_Heading.WEST
                addx = -1

        End Select
        
        nX = x + addx
        nY = y + addy
        
        MapData(nX, nY).charindex = charindex
        .Pos.x = nX
        .Pos.y = nY
        MapData(x, y).charindex = 0
        
        .MoveOffsetX = -1 * (32 * addx)
        .MoveOffsetY = -1 * (32 * addy)
        
        .Moving = 1
        
        'Attached to ladder ;)
        If MapData(nX, nY).ObjGrh.GrhIndex = 26940 Then
            .Heading = E_Heading.NORTH
        Else
            .Heading = nHeading

        End If
        
        .scrollDirectionX = addx
        .scrollDirectionY = addy

    End With
    
    If UserEstado <> 1 Then Call DoPasosFx(charindex)
    
    'areas viejos
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        Call EraseChar(charindex)

    End If
    
End Sub

Public Sub Char_Move_by_Pos(ByVal charindex As Integer, ByVal nX As Integer, ByVal nY As Integer)

    On Error Resume Next

    Dim x        As Integer

    Dim y        As Integer

    Dim addx     As Integer

    Dim addy     As Integer

    Dim nHeading As E_Heading
    
    With charlist(charindex)
        x = .Pos.x
        y = .Pos.y
        
        MapData(x, y).charindex = 0
        
        addx = nX - x
        addy = nY - y
        
        If Sgn(addx) = 1 Then
            nHeading = E_Heading.EAST

        End If
        
        If Sgn(addx) = -1 Then
            nHeading = E_Heading.WEST

        End If
        
        If Sgn(addy) = -1 Then
            nHeading = E_Heading.NORTH

        End If
        
        If Sgn(addy) = 1 Then
            nHeading = E_Heading.south

        End If
        
        MapData(nX, nY).charindex = charindex
        
        .Pos.x = nX
        .Pos.y = nY
        
        .MoveOffsetX = -1 * (TilePixelWidth * addx)
        .MoveOffsetY = -1 * (TilePixelHeight * addy)
        
        .Moving = 1
        
        If MapData(nX, nY).ObjGrh.GrhIndex = 26940 Then
            .Heading = E_Heading.NORTH
        Else
            .Heading = nHeading

        End If
        
        .scrollDirectionX = Sgn(addx)
        .scrollDirectionY = Sgn(addy)

    End With
    
    If Not EstaPCarea(charindex) Then

        'Call Dialogos.RemoveDialog(CharIndex)
        'Call Hits.RemoveHit(CharIndex)
    End If

    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        Call EraseChar(charindex)

    End If

End Sub

Private Function EstaPCarea(ByVal charindex As Integer) As Boolean

    With charlist(charindex).Pos
        EstaPCarea = .x > UserPos.x - MinXBorder And .x < UserPos.x + MinXBorder And .y > UserPos.y - MinYBorder And .y < UserPos.y + MinYBorder

    End With

End Function

Public Function EstaEnArea(ByVal x As Integer, ByVal y As Integer) As Boolean
    EstaEnArea = x > UserPos.x - MinXBorder And x < UserPos.x + MinXBorder And y > UserPos.y - MinYBorder And y < UserPos.y + MinYBorder

End Function

Private Function Engine_FToDW(f As Single) As Long

    ' single > long
    Dim buf As D3DXBuffer

    Set buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData buf, 0, 4, 1, f
    D3DX.BufferGetData buf, 0, 4, 1, Engine_FToDW

End Function

Private Function VectorToRGBA(Vec As D3DVECTOR, fHeight As Single) As Long

    Dim r As Integer, g As Integer, b As Integer, a As Integer

    r = 127 * Vec.x + 128
    g = 127 * Vec.y + 128
    b = 127 * Vec.Z + 128
    a = 255 * fHeight
    VectorToRGBA = D3DColorARGB(a, r, g, b)

End Function

Public Sub DrawMainInventory()

    ' Sólo dibujamos cuando es necesario
    If Not frmmain.Inventario.NeedsRedraw Then Exit Sub

    Dim InvRect As RECT

    InvRect.Left = 0
    InvRect.Top = 0
    InvRect.Right = frmmain.picInv.ScaleWidth
    InvRect.bottom = frmmain.picInv.ScaleHeight

    ' Comenzamos la escena
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0
    D3DDevice.BeginScene

    ' Dibujamos el fondo del inventario principal
    'Call engine.Draw_GrhIndex(6, 0, 0)

    ' Dibujamos items
    Call frmmain.Inventario.DrawInventory
    
    ' Dibujamos item arrastrado
    Call frmmain.Inventario.DrawDraggedItem

    ' Presentamos la escena
    D3DDevice.EndScene
    D3DDevice.Present InvRect, ByVal 0, frmmain.picInv.hwnd, ByVal 0

End Sub

Public Sub DrawInterfaceComerciar()

    ' Sólo dibujamos cuando es necesario
    If Not frmComerciar.InvComNpc.NeedsRedraw And Not frmComerciar.InvComUsu.NeedsRedraw Then Exit Sub

    Dim InvRect As RECT

    InvRect.Left = 0
    InvRect.Top = 0
    InvRect.Right = frmComerciar.interface.ScaleWidth
    InvRect.bottom = frmComerciar.interface.ScaleHeight

    ' Comenzamos la escena
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0
    D3DDevice.BeginScene

    ' Dibujamos el fondo del inventario de comercio
    Call engine.Draw_GrhIndex(837, 0, 0)

    ' Dibujamos items del NPC
    Call frmComerciar.InvComNpc.DrawInventory
    
    ' Dibujamos items del usuario
    Call frmComerciar.InvComUsu.DrawInventory

    ' Dibujamos "ambos" items arrastrados (aunque sólo puede estar uno activo a la vez)
    Call frmComerciar.InvComNpc.DrawDraggedItem
    Call frmComerciar.InvComUsu.DrawDraggedItem
    
    ' Me fijo qué inventario está seleccionado
    Dim CurrentInventory As clsGrapchicalInventory
    
    Dim cantidad         As Integer

    If frmComerciar.InvComNpc.SelectedItem > 0 Then
        Set CurrentInventory = frmComerciar.InvComNpc
        ' Al comprar, calculamos el valor según la cantidad exacta que ingresó
        cantidad = Val(frmComerciar.cantidad.Text)
    ElseIf frmComerciar.InvComUsu.SelectedItem > 0 Then
        Set CurrentInventory = frmComerciar.InvComUsu
        ' Al vender, calculamos el valor según el min(cantidad_ingresada, cantidad_items)
        cantidad = min(Val(frmComerciar.cantidad.Text), CurrentInventory.Amount(CurrentInventory.SelectedItem))

    End If
    
    ' Si hay alguno seleccionado
    If Not CurrentInventory Is Nothing Then
        ' Dibujo el item seleccionado
        'Call engine.Draw_GrhColor(CurrentInventory.GrhIndex(CurrentInventory.SelectedItem), 282, 251, COLOR_WHITE)
    
        ' Muestro info del item
        Dim str As String

        str = " (No usa: "
        
        Select Case CurrentInventory.PuedeUsar(CurrentInventory.SelectedItem)

            Case 1
                str = str & "Genero)"

            Case 2
                str = str & "Clase)"

            Case 3
                str = str & "Facción)"

            Case 4
                str = str & "Skill)"

            Case 5
                str = str & "Raza)"

            Case 6
                str = str & "Nivel)"

            Case 0
                str = " (Usable)"

        End Select
                           
        frmComerciar.lblnombre = CurrentInventory.ItemName(CurrentInventory.SelectedItem) & str
        frmComerciar.lbldesc = CurrentInventory.GetInfo(CurrentInventory.OBJIndex(CurrentInventory.SelectedItem))
        frmComerciar.lblCosto = PonerPuntos(CLng(CurrentInventory.Valor(CurrentInventory.SelectedItem) * cantidad))
        
        Set CurrentInventory = Nothing

    End If

    ' Presentamos la escena
    D3DDevice.EndScene
    D3DDevice.Present InvRect, ByVal 0, frmComerciar.interface.hwnd, ByVal 0

End Sub

Public Sub DrawInterfaceBoveda()

    ' Sólo dibujamos cuando es necesario
    If Not frmBancoObj.InvBoveda.NeedsRedraw And Not frmBancoObj.InvBankUsu.NeedsRedraw Then Exit Sub

    Dim InvRect As RECT

    InvRect.Left = 0
    InvRect.Top = 0
    InvRect.Right = frmBancoObj.interface.ScaleWidth
    InvRect.bottom = frmBancoObj.interface.ScaleHeight

    ' Comenzamos la escena
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0
    D3DDevice.BeginScene

    ' Dibujamos el fondo de la bóveda
    Call engine.Draw_GrhIndex(838, 0, 0)

    ' Dibujamos items de la bóveda
    Call frmBancoObj.InvBoveda.DrawInventory
    
    ' Dibujamos items del usuario
    Call frmBancoObj.InvBankUsu.DrawInventory

    ' Dibujamos "ambos" items arrastrados (aunque sólo puede estar uno activo a la vez)
    Call frmBancoObj.InvBoveda.DrawDraggedItem
    Call frmBancoObj.InvBankUsu.DrawDraggedItem
    
    ' Me fijo qué inventario está seleccionado
    Dim CurrentInventory As clsGrapchicalInventory
    
    If frmBancoObj.InvBoveda.SelectedItem > 0 Then
        Set CurrentInventory = frmBancoObj.InvBoveda
    ElseIf frmBancoObj.InvBankUsu.SelectedItem > 0 Then
        Set CurrentInventory = frmBancoObj.InvBankUsu

    End If
    
    ' Si hay alguno seleccionado
    If Not CurrentInventory Is Nothing Then

        ' Muestro info del item
        Dim str As String

        str = " (No usa: "
        
        Select Case CurrentInventory.PuedeUsar(CurrentInventory.SelectedItem)

            Case 1
                str = str & "Genero)"

            Case 2
                str = str & "Clase)"

            Case 3
                str = str & "Facción)"

            Case 4
                str = str & "Skill)"

            Case 5
                str = str & "Raza)"

            Case 6
                str = str & "Nivel)"

            Case 0
                str = " (Usable)"

        End Select
        
        frmBancoObj.lblnombre.Caption = CurrentInventory.ItemName(CurrentInventory.SelectedItem) & str
        frmBancoObj.lbldesc.Caption = CurrentInventory.GetInfo(CurrentInventory.OBJIndex(CurrentInventory.SelectedItem))
        
        Set CurrentInventory = Nothing

    End If

    ' Presentamos la escena
    D3DDevice.EndScene
    D3DDevice.Present InvRect, ByVal 0, frmBancoObj.interface.hwnd, ByVal 0

End Sub

Public Sub DrawMapaMundo()

    On Error Resume Next

    Static re          As RECT

    Static rgb_list(3) As Long

    re.Left = 0
    re.Top = 0
    re.bottom = 89
    re.Right = 177
    
    frmMapaGrande.PlayerView.Height = 89
    frmMapaGrande.PlayerView.Width = 177
    frmMapaGrande.PlayerView.ScaleHeight = 89
    frmMapaGrande.PlayerView.ScaleWidth = 177
    
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0
    D3DDevice.BeginScene
        
    Dim color(0 To 3) As Long

    color(0) = D3DColorARGB(255, 255, 255, 255)
    color(1) = color(0)
    color(2) = color(0)
    color(3) = color(0)
        
    Dim i    As Byte

    Dim x    As Integer

    Dim y    As Integer
    
    Dim Head As grh

    Head = HeadData(NpcData(frmMapaGrande.ListView1.SelectedItem.SubItems(2)).Head).Head(3)
    
    Dim grh As grh

    grh = BodyData(NpcData(frmMapaGrande.ListView1.SelectedItem.SubItems(2)).Body).Walk(3)
    
    Dim tmp           As String

    Dim temp_array(3) As Long 'Si le queres dar color a la letra pasa este parametro dsp xD

    Engine_Draw_Box x, y, 177, 89, D3DColorARGB(255, 7, 7, 7) 'Fondo del inventario
    
    x = frmMapaGrande.PlayerView.ScaleWidth / 2 - GrhData(grh.GrhIndex).pixelWidth / 2
    y = frmMapaGrande.PlayerView.ScaleHeight / 2 - GrhData(grh.GrhIndex).pixelHeight / 2
    Call Draw_Grh(grh, x, y, 0, 0, color, False, 0, 0, 0)

    x = frmMapaGrande.PlayerView.ScaleWidth / 2 - GrhData(Head.GrhIndex).pixelWidth / 2
    y = frmMapaGrande.PlayerView.ScaleHeight / 2 - GrhData(Head.GrhIndex).pixelHeight + 8 + BodyData(NpcData(frmMapaGrande.ListView1.SelectedItem.SubItems(2)).Body).HeadOffset.y / 2
    Call Draw_Grh(Head, x, y, 0, 0, color, False, 0, 0, 0)
    
    D3DDevice.EndScene
    
    D3DDevice.Present re, ByVal 0, frmMapaGrande.PlayerView.hwnd, ByVal 0

End Sub

Public Sub Draw_FilledBox(ByVal x As Integer, ByVal y As Integer, ByVal Width As Integer, ByVal Height As Integer, color As Long, outlinecolor As Long)

    Static box_rect     As RECT

    Static Outline      As RECT

    Static rgb_list(3)  As Long

    Static rgb_list2(3) As Long

    Static Vertex(3)    As TLVERTEX

    Static Vertex2(3)   As TLVERTEX
    
    rgb_list(0) = color
    rgb_list(1) = color
    rgb_list(2) = color
    rgb_list(3) = color
    
    rgb_list2(0) = outlinecolor
    rgb_list2(1) = outlinecolor
    rgb_list2(2) = outlinecolor
    rgb_list2(3) = outlinecolor
    
    With box_rect
        .bottom = y + Height
        .Left = x
        .Right = x + Width
        .Top = y

    End With
    
    With Outline
        .bottom = y + Height + 1
        .Left = x - 1
        .Right = x + Width + 1
        .Top = y - 1

    End With
    
    Geometry_Create_Box Vertex2(), Outline, Outline, rgb_list2(), 0, 0
    Geometry_Create_Box Vertex(), box_rect, box_rect, rgb_list(), 0, 0
    
    D3DDevice.SetTexture 0, Nothing
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex2(0), Len(Vertex2(0))
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex(0), Len(Vertex(0))

End Sub

Private Function Particle_Group_Next_Open() As Long

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

Private Function Particle_Group_Check(ByVal particle_group_index As Long) As Boolean

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

End Function

Public Function Particle_Group_Create(ByVal map_x As Integer, ByVal map_y As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
   Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
   Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
   Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
   Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
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

            Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey

        End If

    Else
        Particle_Group_Create = Particle_Group_Next_Open
      
        Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey

    End If

End Function

Public Function Particle_Group_Remove(ByVal particle_group_index As Long) As Boolean
    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 1/04/2003
    '
    '*****************************************************************
    'Make sure it's a legal index

    If Particle_Group_Check(particle_group_index) Then
        particle_group_list(particle_group_index).never_die = False
        particle_group_list(particle_group_index).alive_counter = 0
        
        particle_group_list(particle_group_index).destruir = True
    
        Rem Particle_Group_Destroy particle_group_index
        Particle_Group_Remove = True

    End If

End Function

Public Function Particle_Group_Remove_All() As Boolean

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 1/04/2003
    '
    '*****************************************************************
    Dim Index As Long
    
    For Index = 1 To particle_group_last

        'Make sure it's a legal index
        If Particle_Group_Check(Index) Then
            Particle_Group_Destroy Index

        End If

    Next Index
    
    Particle_Group_Remove_All = True

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

    '*****************************************************************
    'Author: Pablo Mercavides
    'Last Modify Date: 1/04/2003
    'Find the index related to the handle
    '*****************************************************************
    On Error GoTo ErrorHandler:

    particle_group_list(id).particle_count = CantPartLLuvia
        
    Exit Function
ErrorHandler:

End Function

Private Sub Particle_Group_Destroy(ByVal particle_group_index As Long)

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    '
    '**************************************************************
    Dim temp  As particle_group

    Dim i     As Integer

    Dim ii    As Integer

    Dim b     As Integer

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

                    For b = ii To antes - 1
                        charlist(particle_group_list(particle_group_index).char_index).particle_group(b) = charlist(particle_group_list(particle_group_index).char_index).particle_group(b + 1)
                        ' charlist(particle_group_list(particle_group_index).char_index).particle_group(b + 1) = 0
                    Next b

                    Rem       ReDim Preserve charlist(particle_group_list(particle_group_index).char_index).particle_group(1 To charlist(particle_group_list(particle_group_index).char_index).particle_count)
                    Rem Else
                    Rem ReDim charlist(particle_group_list(particle_group_index).char_index).particle_group(0)
                    '  End If
                          
                    Exit For
                    
                End If
                
            Next i

        End If

    ElseIf particle_group_index = meteo_particle Then
        meteo_particle = 0

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
    
End Sub

Private Sub Particle_Group_Make(ByVal particle_group_index As Long, ByVal map_x As Integer, ByVal map_y As Integer, _
   ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
   Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
   Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
   Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
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
    particle_group_list(particle_group_index).angle = angle
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
    
End Sub

Private Sub Char_Particle_Group_Make(ByVal particle_group_index As Long, ByVal char_index As Integer, ByVal particle_char_index As Integer, _
   ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
   Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
   Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
   Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
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
    particle_group_list(particle_group_index).angle = angle
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
End Sub

Private Sub Particle_Incrementar(ByVal id As Integer)

    If particle_group_list(id).Creando < particle_group_list(id).particle_count Then
        particle_group_list(id).Creando = particle_group_list(id).Creando + 1
    Else
        particle_group_list(id).Creada = True

    End If

    If particle_group_list(id).char_index > 0 Then
        charlist(particle_group_list(id).char_index).particle_count = particle_group_list(id).Creando

    End If

End Sub

Private Sub Particle_Group_Render(ByVal particle_group_index As Long, ByVal screen_x As Integer, ByVal screen_y As Integer)

    '*****************************************************************
    'Author: Aaron Perkins
    'Modified by: Ryan Cain (Onezero)
    'Modified by: Juan Martín Sotuyo Dodero
    'Last Modify Date: 5/15/2003
    'Renders a particle stream at a paticular screen point
    '*****************************************************************
    On Error Resume Next

    Dim loopc            As Long

    Dim temp_rgb(0 To 3) As Long

    Dim no_move          As Boolean
    
    'Set colors
  
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
               particle_group_list(particle_group_index).x1, particle_group_list(particle_group_index).y1, particle_group_list(particle_group_index).angle, _
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
               particle_group_list(particle_group_index).x1, particle_group_list(particle_group_index).y1, particle_group_list(particle_group_index).angle, _
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

End Sub

Private Sub Particle_Render(ByRef temp_particle As Particle, ByVal screen_x As Integer, ByVal screen_y As Integer, _
   ByVal grh_index As Long, ByRef rgb_list() As Long, _
   Optional ByVal alpha_blend As Boolean, Optional ByVal no_move As Boolean, _
   Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
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

    If no_move = False Then

        If temp_particle.alive_counter = 0 And Not destruir Then
            'Start new particle
            InitGrh temp_particle.grh, grh_index, alpha_blend
            
            temp_particle.x = RandomNumber(x1, x2) - (32 / 2)
            temp_particle.y = RandomNumber(y1, y2) - (32 / 2)
            temp_particle.vector_x = RandomNumber(vecx1, vecx2)
            temp_particle.vector_y = RandomNumber(vecy1, vecy2)
            temp_particle.angle = angle
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
            If spin = True Then temp_particle.grh.angle = temp_particle.grh.angle + (RandomNumber(spin_speedL, spin_speedH) / 100)
            If temp_particle.angle >= 360 Then
                temp_particle.angle = 0

            End If
            
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
    
    'Particulas Grises si esta muerto Ladder
    If UserEstado = 1 Then
        rgb_list(0) = D3DColorARGB(100, 100, 100, 100)
        rgb_list(1) = D3DColorARGB(100, 100, 100, 100)
        rgb_list(2) = D3DColorARGB(100, 100, 100, 100)
        rgb_list(3) = D3DColorARGB(100, 100, 100, 100)

    End If
    
    If Not alpha_blend Then

        Dim r, g, b As Byte

        r = (rgb_list(0) And 16711680) / 65536
        g = (rgb_list(0) And 65280) / 256
        b = rgb_list(0) And 255
             
        rgb_list(0) = D3DColorARGB(temp_particle.angle, r, g, b)
        
        r = (rgb_list(1) And 16711680) / 65536
        g = (rgb_list(1) And 65280) / 256
        b = rgb_list(1) And 255
        
        rgb_list(1) = D3DColorARGB(temp_particle.angle, r, g, b)
        
        r = (rgb_list(2) And 16711680) / 65536
        g = (rgb_list(2) And 65280) / 256
        b = rgb_list(2) And 255
        rgb_list(2) = D3DColorARGB(temp_particle.angle, r, g, b)
        
        r = (rgb_list(3) And 16711680) / 65536
        g = (rgb_list(3) And 65280) / 256
        b = rgb_list(3) And 255
        rgb_list(3) = D3DColorARGB(temp_particle.angle, r, g, b)

    End If
    
    If UserCiego = True Then
        rgb_list(0) = D3DColorARGB(5, 5, 5, 5)
        rgb_list(1) = rgb_list(0)
        rgb_list(2) = rgb_list(0)
        rgb_list(3) = rgb_list(0)

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
    
End Sub

Public Function Particle_Type_Get(ByVal particle_Index As Long) As Long

    '*****************************************************************
    'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
    'Last Modify Date: 8/27/2003
    'Returns the stream type of a particle stream
    '*****************************************************************
    If Particle_Group_Check(particle_Index) Then
        Particle_Type_Get = particle_group_list(particle_Index).stream_type

    End If

End Function

Public Function Engine_Meteo_Particle_Get() As Long
    '*****************************************************************
    'Author: Augusto José Rando
    'Last Modify Date: 6/11/2002
    '*****************************************************************
    Engine_Meteo_Particle_Get = meteo_particle

End Function

Public Function Engine_Meteo_Estado_Get() As Byte
    '*****************************************************************
    'Author: Augusto José Rando
    'Last Modify Date: 6/11/2002
    '*****************************************************************
    Engine_Meteo_Estado_Get = meteo_estado

End Function

Public Function Map_Particle_Group_Get(ByVal map_x As Integer, ByVal map_y As Integer) As Long

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

End Function

Public Sub Engine_Meteo_Particle_Set(ByVal meteo_part As Long)
    '*****************************************************************
    'Author: Augusto José Rando
    'Last Modify Date: 6/11/2002
    '*****************************************************************
    
    If (meteo_part = -1) And (meteo_particle <> 0) Then
        Call Particle_Group_Remove(meteo_particle)
        
    ElseIf (meteo_part <> -1) Then

        If meteo_particle <> 0 Then Call Particle_Group_Remove(meteo_particle)
        meteo_particle = General_Particle_Create(meteo_part, -1, -1)
        MeteoIndex = particle_group_last

    End If
    
End Sub

Public Sub Engine_spell_Particle_Set(ByVal spell_part As Long)
    '*****************************************************************
    'Author: Augusto José Rando
    'Last Modify Date: 6/11/2002
    '*****************************************************************
    
    If (spell_part = -1) And (spell_particle <> 0) Then
        Call Particle_Group_Remove(spell_particle)
    ElseIf (spell_part <> -1) Then

        If spell_particle <> 0 Then Call Particle_Group_Remove(spell_particle)
        spell_particle = General_Particle_Create(spell_part, -1, -1)

    End If
    
End Sub

Public Sub Engine_Select_Particle_Set(ByVal Select_particle As Long)
    '*****************************************************************
    'Author: Augusto José Rando
    'Last Modify Date: 6/11/2002
    '*****************************************************************
    
    If (Select_particle = -1) And (Select_part <> 0) Then
        Call Particle_Group_Remove(Select_part)
    ElseIf (Select_part <> -1) Then

        If Select_part <> 0 Then Call Particle_Group_Remove(Select_part)
        Select_part = General_Particle_Create(Select_particle, -1, -1)

    End If
    
End Sub

Private Sub Grh_Render_Advance(ByRef grh As grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByVal Height As Integer, ByVal Width As Integer, ByRef rgb_list() As Long, Optional ByVal h_center As Boolean, Optional ByVal v_center As Boolean, Optional ByVal alpha_blend As Boolean = False)

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
    'Last Modify Date: 11/19/2003
    'Similar to Grh_Render, but let´s you resize the Grh
    '**************************************************************
    Dim tile_width  As Integer

    Dim tile_height As Integer

    Dim grh_index   As Long
    
    'Animation
    If grh.Started Then
        grh.framecounter = grh.framecounter + (timerTicksPerFrame * grh.speed)

        If grh.framecounter > GrhData(grh.GrhIndex).NumFrames Then
            'If Grh.noloop Then
            '    Grh.FrameCounter = GrhData(Grh.GrhIndex).NumFrames
            'Else
            grh.framecounter = 1

            'End If
        End If

    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    If grh.framecounter = 0 Then grh.framecounter = 1
    grh_index = GrhData(grh.GrhIndex).Frames(grh.framecounter)
    
    'Center Grh over X, Y pos
    If GrhData(grh.GrhIndex).TileWidth <> 1 Then
        screen_x = screen_x - Int(GrhData(grh.GrhIndex).TileWidth * (32 \ 2)) + 32 \ 2

    End If
    
    If GrhData(grh.GrhIndex).TileHeight <> 1 Then
        screen_y = screen_y - Int(GrhData(grh.GrhIndex).TileHeight * 32) + 32

    End If
    
    'Draw it to device
    Device_Box_Textured_Render_Advance grh_index, screen_x, screen_y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, rgb_list, GrhData(grh_index).sX, GrhData(grh_index).sY, Width, Height, alpha_blend, grh.angle

End Sub

Private Sub Grh_Render(ByRef grh As grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByRef rgb_list() As Long, Optional ByVal h_centered As Boolean = True, Optional ByVal v_centered As Boolean = True, Optional ByVal alpha_blend As Boolean = False)

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 2/28/2003
    'Modified by Juan Martín Sotuyo Dodero
    'Added centering
    '**************************************************************
    Dim tile_width  As Integer

    Dim tile_height As Integer

    Dim grh_index   As Long
    
    If grh.GrhIndex = 0 Then Exit Sub
        
    'Animation
    If grh.Started Then
        grh.framecounter = grh.framecounter + (timerTicksPerFrame * grh.speed)

        If grh.framecounter > GrhData(grh.GrhIndex).NumFrames Then
            'If Grh.noloop Then
            '    Grh.FrameCounter = GrhData(Grh.GrhIndex).NumFrames
            'Else
            grh.framecounter = 1

            'End If
        End If

    End If

    ' particle_group_list(particle_group_index).frame_counter = particle_group_list(particle_group_index).frame_counter + timer_ticks_per_frame
    ' If particle_group_list(particle_group_index).frame_counter > particle_group_list(particle_group_index).frame_speed Then
    '     particle_group_list(particle_group_index).frame_counter = 0
    '      no_move = False
    '  Else
    '     no_move = True
    '  End If

    'Figure out what frame to draw (always 1 if not animated)
    If grh.framecounter = 0 Then grh.framecounter = 1
    ' If Not Grh_Check(Grh.grhindex) Then Exit Sub
    grh_index = GrhData(grh.GrhIndex).Frames(grh.framecounter)

    If grh_index <= 0 Then Exit Sub
    If GrhData(grh_index).FileNum = 0 Then Exit Sub
        
    'Modified by Augusto José Rando
    'Simplier function - according to basic ORE engine
    If h_centered Then
        If GrhData(grh.GrhIndex).TileWidth <> 1 Then
            screen_x = screen_x - Int(GrhData(grh.GrhIndex).TileWidth * (32 \ 2)) + 32 \ 2

        End If

    End If
    
    If v_centered Then
        If GrhData(grh.GrhIndex).TileHeight <> 1 Then
            screen_y = screen_y - Int(GrhData(grh.GrhIndex).TileHeight * 32) + 32

        End If

    End If
    
    'Draw it to device
    Device_Box_Textured_Render grh_index, screen_x, screen_y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, rgb_list(), GrhData(grh_index).sX, GrhData(grh_index).sY, alpha_blend, grh.angle

End Sub

Private Function Grh_Check(ByVal grh_index As Long) As Boolean

    '**************************************************************
    'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
    'Last Modify Date: 1/04/2003
    '
    '**************************************************************
    'check grh_index
    If grh_index > 0 And grh_index <= MaxGrh Then
        Grh_Check = GrhData(grh_index).NumFrames

    End If

End Function

Public Function Char_Particle_Group_Create(ByVal char_index As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
   Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
   Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
   Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
   Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
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

    Dim char_part_free_index As Integer
    
    Rem If charlist(char_index).Particula = stream_type Then Exit Function
    'If Char_Particle_Group_Find(char_index, stream_type) Then Exit Function ' hay que ver si dejar o sacar esto...
    If Not Char_Check(char_index) Then Exit Function
    char_part_free_index = Char_Particle_Group_Next_Open(char_index)
    
    If char_part_free_index > 0 Then
        Char_Particle_Group_Create = Particle_Group_Next_Open
        Char_Particle_Group_Make Char_Particle_Group_Create, char_index, char_part_free_index, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, id, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey

    End If

End Function

Private Function Char_Check(ByVal char_index As Integer) As Boolean

    '**************************************************************
    'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
    'Last Modify Date: 1/04/2003
    '
    '**************************************************************
    'check char_index
    If char_index > 0 And char_index <= LastChar Then
        Char_Check = (charlist(char_index).Heading > 0)

    End If
    
End Function

Public Function Char_Particle_Group_Remove(ByVal char_index As Integer, ByVal stream_type As Long)
    '**************************************************************
    'Author: Augusto José Rando
    '**************************************************************

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

End Function

Public Function Char_Particle_Group_Remove_All(ByVal char_index As Integer)

    '**************************************************************
    'Author: Augusto José Rando
    '**************************************************************
    Dim i As Integer
    
    If Char_Check(char_index) Then

        For i = 1 To charlist(char_index).particle_count

            If charlist(char_index).particle_group(i) <> 0 Then Call Particle_Group_Remove(charlist(char_index).particle_group(i))
        Next i

    End If
    
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

Private Function Char_Particle_Group_Next_Open(ByVal char_index As Integer) As Integer

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

Public Function Char_FX_Group_Next_Open(ByVal char_index As Integer) As Integer

    '*****************************************************************
    'Author: Augusto José Rando
    '*****************************************************************
    On Error GoTo ErrorHandler:

    Dim loopc As Long
    
    If charlist(char_index).FxCount = 0 Then
        charlist(char_index).FxCount = 1
        ReDim charlist(char_index).FxList(1 To 1)
        Char_FX_Group_Next_Open = 1
        Exit Function

    End If
    
    loopc = 1

    Do Until charlist(char_index).FxList(loopc).FxIndex = 0

        If loopc = charlist(char_index).FxCount Then
            Char_FX_Group_Next_Open = charlist(char_index).FxCount + 1
            charlist(char_index).FxCount = Char_FX_Group_Next_Open
            ReDim Preserve charlist(char_index).FxList(1 To Char_FX_Group_Next_Open)
            Exit Function

        End If

        loopc = loopc + 1
    Loop

    Char_FX_Group_Next_Open = loopc
    Exit Function

ErrorHandler:
    charlist(char_index).FxCount = 1
    ReDim charlist(char_index).FxList(1 To 1)
    Char_FX_Group_Next_Open = 1

End Function

Public Function Map_FX_Group_Next_Open(ByVal x As Byte, ByVal y As Byte) As Integer

    '*****************************************************************
    'Author: Augusto José Rando
    '*****************************************************************
    On Error GoTo ErrorHandler:

    Dim loopc As Long
    
    If MapData(x, y).FxCount = 0 Then
        MapData(x, y).FxCount = 1
        ReDim MapData(x, y).FxList(1 To 1)
        Map_FX_Group_Next_Open = 1
        Exit Function

    End If
    
    loopc = 1

    Do Until MapData(x, y).FxList(loopc).FxIndex = 0

        If loopc = MapData(x, y).FxCount Then
            Map_FX_Group_Next_Open = MapData(x, y).FxCount + 1
            MapData(x, y).FxCount = Map_FX_Group_Next_Open
            ReDim Preserve MapData(x, y).FxList(1 To Map_FX_Group_Next_Open)
            Exit Function

        End If

        loopc = loopc + 1
    Loop

    Map_FX_Group_Next_Open = loopc
    Exit Function

ErrorHandler:
    MapData(x, y).FxCount = 1
    ReDim MapData(x, y).FxList(1 To 1)
    Map_FX_Group_Next_Open = 1

End Function

Private Sub Draw_Sombra(ByRef grh As grh, ByVal x As Integer, ByVal y As Integer, ByVal center As Byte, ByVal animate As Byte, Optional ByVal Alpha As Boolean, Optional ByVal map_x As Byte = 1, Optional ByVal map_y As Byte = 1, Optional ByVal angle As Single)

    On Error Resume Next

    ' If meteo_estado = 3 Or meteo_estado = 4 Then Exit Sub
    ' If UserEstado = 1 Then Exit Sub
    ' If bTecho Then Exit Sub
    Dim CurrentGrhIndex As Long

    If grh.GrhIndex = 0 Then Exit Sub
    'Por ladder
    
    CurrentGrhIndex = GrhData(grh.GrhIndex).Frames(grh.framecounter)

    If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
        x = x - Int(GrhData(CurrentGrhIndex).TileWidth * (32 \ 2)) + 32 \ 2

    End If

    If GrhData(grh.GrhIndex).TileHeight <> 1 Then
        y = y - Int(GrhData(CurrentGrhIndex).TileHeight * 32) + 32

    End If

    Dim SRGB(3) As Long

    SRGB(0) = D3DColorARGB(50, 0, 0, 0)
    SRGB(1) = D3DColorARGB(50, 0, 0, 0)
    SRGB(2) = D3DColorARGB(50, 0, 0, 0)
    SRGB(3) = D3DColorARGB(50, 0, 0, 0)
    'Por ladder
    Device_Box_Textured_Render CurrentGrhIndex, x + 5, y + 5, GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, SRGB(), GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, False, 0.4

    'Por ladder
End Sub

Sub Engine_Weather_UpdateFog()

    '*****************************************************************
    'Update the fog effects
    '*****************************************************************
    Dim TempGrh     As grh

    Dim i           As Long

    Dim x           As Long

    Dim y           As Long

    Dim cc(3)       As Long

    Dim ElapsedTime As Single

    ElapsedTime = Engine_ElapsedTime

    If WeatherFogCount = 0 Then WeatherFogCount = 13

    WeatherFogX1 = WeatherFogX1 + (ElapsedTime * (0.018 + Rnd * 0.01)) + (LastOffsetX - ParticleOffsetX)
    WeatherFogY1 = WeatherFogY1 + (ElapsedTime * (0.013 + Rnd * 0.01)) + (LastOffsetY - ParticleOffsetY)
    
    Do While WeatherFogX1 < -512
        WeatherFogX1 = WeatherFogX1 + 512
    Loop

    Do While WeatherFogY1 < -512
        WeatherFogY1 = WeatherFogY1 + 512
    Loop

    Do While WeatherFogX1 > 0
        WeatherFogX1 = WeatherFogX1 - 512
    Loop

    Do While WeatherFogY1 > 0
        WeatherFogY1 = WeatherFogY1 - 512
    Loop
    
    WeatherFogX2 = WeatherFogX2 - (ElapsedTime * (0.037 + Rnd * 0.01)) + (LastOffsetX - ParticleOffsetX)
    WeatherFogY2 = WeatherFogY2 - (ElapsedTime * (0.021 + Rnd * 0.01)) + (LastOffsetY - ParticleOffsetY)

    Do While WeatherFogX2 < -512
        WeatherFogX2 = WeatherFogX2 + 512
    Loop

    Do While WeatherFogY2 < -512
        WeatherFogY2 = WeatherFogY2 + 512
    Loop

    Do While WeatherFogX2 > 0
        WeatherFogX2 = WeatherFogX2 - 512
    Loop

    Do While WeatherFogY2 > 0
        WeatherFogY2 = WeatherFogY2 - 512
    Loop

    TempGrh.framecounter = 1
    
    'Render fog 2
    TempGrh.GrhIndex = 32014
    x = 2
    y = -1

    cc(1) = D3DColorARGB(AlphaNiebla, 255, 255, 255)
    cc(2) = D3DColorARGB(AlphaNiebla, 255, 255, 255)
    cc(3) = D3DColorARGB(AlphaNiebla, 255, 255, 255)
    cc(0) = D3DColorARGB(AlphaNiebla, 255, 255, 255)

    For i = 1 To WeatherFogCount
        Draw_Grh TempGrh, (x * 512) + WeatherFogX2, (y * 512) + WeatherFogY2, 0, 0, cc()
        x = x + 1

        If x > (1 + (ScreenWidth \ 512)) Then
            x = 0
            y = y + 1

        End If

    Next i
            
    'Render fog 1
    TempGrh.GrhIndex = 32015
    x = 0
    y = 0
    cc(1) = D3DColorARGB(AlphaNiebla, 255, 255, 255)
    cc(2) = D3DColorARGB(AlphaNiebla, 255, 255, 255)
    cc(3) = D3DColorARGB(AlphaNiebla, 255, 255, 255)
    cc(0) = D3DColorARGB(AlphaNiebla, 255, 255, 255)

    For i = 1 To WeatherFogCount
        Draw_Grh TempGrh, (x * 512) + WeatherFogX1, (y * 512) + WeatherFogY1, 0, 0, cc()
        x = x + 1

        If x > (2 + (ScreenWidth \ 512)) Then
            x = 0
            y = y + 1

        End If

    Next i

End Sub

Function Engine_PixelPosX(ByVal x As Integer) As Integer
    '*****************************************************************
    'Converts a tile position to a screen position
    '*****************************************************************
    Engine_PixelPosX = (x - 1) * 32

End Function

Function Engine_PixelPosY(ByVal y As Integer) As Integer
    '*****************************************************************
    'Converts a tile position to a screen position
    '*****************************************************************
    Engine_PixelPosY = (y - 1) * 32

End Function

Private Function Engine_ElapsedTime() As Long

    '**************************************************************
    'Gets the time that past since the last call
    '**************************************************************
    Dim Start_Time As Long

    Start_Time = (GetTickCount() And &H7FFFFFFF)
    Engine_ElapsedTime = Start_Time - EndTime

    If Engine_ElapsedTime > 1000 Then Engine_ElapsedTime = 1000
    EndTime = Start_Time

End Function

Private Sub Renderizar_Aura(ByVal aura_index As String, ByVal x As Integer, ByVal y As Integer, ByVal map_x As Byte, ByVal map_y As Byte, Optional ByVal userindex As Long = 0)

    Dim rgb_list(0 To 3) As Long

    Dim i                As Byte

    Dim Index            As Long

    Dim color            As Long

    Dim aura_grh         As grh

    Dim TRANS            As Integer

    Dim giro             As Single

    Dim lado             As Byte

    Index = Val(ReadField(1, aura_index, Asc(":")))
    color = Val(ReadField(2, aura_index, Asc(":")))
    giro = Val(ReadField(3, aura_index, Asc(":")))
    lado = Val(ReadField(4, aura_index, Asc(":")))

    'Debug.Print charlist(userindex).AuraAngle
    If giro > 0 And userindex > 0 Then
        'If lado = 0 Then
        charlist(userindex).AuraAngle = charlist(userindex).AuraAngle + (timerTicksPerFrame * giro)
        'Else
        'charlist(userindex).AuraAngle = charlist(userindex).AuraAngle - (timerTicksPerFrame * giro)
        ' End If
    
        If charlist(userindex).AuraAngle >= 360 Then charlist(userindex).AuraAngle = 0

    End If

    'If charlist(userindex).AuraAngle <> 0 Then
    'Debug.Print charlist(userindex).AuraAngle
    'End If
    Dim r As Integer

    Dim g As Integer

    Dim b As Integer

    r = &HFF& And color
    g = (&HFF00& And color) \ 256
    b = (&HFF0000 And color) \ 65536
    TRANS = 255

    rgb_list(0) = D3DColorARGB(TRANS, b, g, r)
    rgb_list(1) = D3DColorARGB(TRANS, b, g, r)
    rgb_list(2) = D3DColorARGB(TRANS, b, g, r)
    rgb_list(3) = D3DColorARGB(TRANS, b, g, r)

    'Convertimos el Aura en un GRH
    Call InitGrh(aura_grh, Index)

    'Y por ultimo renderizamos esta capa con Draw_Grh
    If giro > 0 And userindex > 0 Then
        Call Draw_Grh(aura_grh, x, y + 30, 1, 0, rgb_list(), True, map_x, map_y, charlist(userindex).AuraAngle)
    Else
        Call Draw_Grh(aura_grh, x, y + 30, 1, 0, rgb_list(), True, map_x, map_y, 0)

    End If
    
End Sub

Private Sub Renderizar_AuraCiego(ByVal aura_index As String, ByVal x As Integer, ByVal y As Integer, ByVal map_x As Byte, ByVal map_y As Byte)

    Dim rgb_list(0 To 3) As Long

    Dim i                As Byte

    Dim Index            As Long

    Dim color            As Long

    Dim aura_grh         As grh

    Dim TRANS            As Integer

    Index = Val(ReadField(1, aura_index, Asc(":")))
    color = Val(ReadField(2, aura_index, Asc(":")))
    TRANS = 1 'Val(ReadField(4, aura_index, Asc(":")))

    Dim r As Integer

    Dim g As Integer

    Dim b As Integer

    r = &HFF& And color
    g = (&HFF00& And color) \ 256
    b = (&HFF0000 And color) \ 65536

    Dim ColorCiego(0 To 3) As Long

    ColorCiego(0) = D3DColorARGB(255, 30, 30, 30)
    ColorCiego(1) = ColorCiego(0)
    ColorCiego(2) = ColorCiego(0)
    ColorCiego(3) = ColorCiego(0)

    rgb_list(0) = ColorCiego(0)
    rgb_list(1) = ColorCiego(0)
    rgb_list(2) = ColorCiego(0)
    rgb_list(3) = ColorCiego(0)

    'Convertimos el Aura en un GRH
    Call InitGrh(aura_grh, Index)
    'Y por ultimo renderizamos esta capa con Draw_Grh
    Call Draw_Grh(aura_grh, x, y + 30, 1, 0, rgb_list(), True, map_x, map_y)
    
End Sub

Private Sub Engine_Font_Initialize()

    Dim a As Integer

    Fuentes(1).Tamanio = 9
    Fuentes(1).Caracteres(48) = 21452
    Fuentes(1).Caracteres(49) = 21453
    Fuentes(1).Caracteres(50) = 21454
    Fuentes(1).Caracteres(51) = 21455
    Fuentes(1).Caracteres(52) = 21456
    Fuentes(1).Caracteres(53) = 21457
    Fuentes(1).Caracteres(54) = 21458
    Fuentes(1).Caracteres(55) = 21459
    Fuentes(1).Caracteres(56) = 21460
    Fuentes(1).Caracteres(57) = 21461

    For a = 0 To 25
        Fuentes(1).Caracteres(a + 97) = 21400 + a
    Next a

    For a = 0 To 25
        Fuentes(1).Caracteres(a + 65) = 21426 + a
    Next a

    Fuentes(1).Caracteres(33) = 21462
    Fuentes(1).Caracteres(161) = 21463
    Fuentes(1).Caracteres(34) = 21464
    Fuentes(1).Caracteres(36) = 21465
    Fuentes(1).Caracteres(191) = 21466
    Fuentes(1).Caracteres(35) = 21467
    Fuentes(1).Caracteres(36) = 21468
    Fuentes(1).Caracteres(37) = 21469
    Fuentes(1).Caracteres(38) = 21470
    Fuentes(1).Caracteres(47) = 21471
    Fuentes(1).Caracteres(92) = 21472
    Fuentes(1).Caracteres(40) = 21473
    Fuentes(1).Caracteres(41) = 21474
    Fuentes(1).Caracteres(61) = 21475
    Fuentes(1).Caracteres(39) = 21476
    Fuentes(1).Caracteres(123) = 21477
    Fuentes(1).Caracteres(125) = 21478
    Fuentes(1).Caracteres(95) = 21479
    Fuentes(1).Caracteres(45) = 21480
    Fuentes(1).Caracteres(63) = 21465
    Fuentes(1).Caracteres(64) = 21481
    Fuentes(1).Caracteres(94) = 21482
    Fuentes(1).Caracteres(91) = 21483
    Fuentes(1).Caracteres(93) = 21484
    Fuentes(1).Caracteres(60) = 21485
    Fuentes(1).Caracteres(62) = 21486
    Fuentes(1).Caracteres(42) = 21487
    Fuentes(1).Caracteres(43) = 21488
    Fuentes(1).Caracteres(46) = 21489
    Fuentes(1).Caracteres(44) = 21490
    Fuentes(1).Caracteres(58) = 21491
    Fuentes(1).Caracteres(59) = 21492
    Fuentes(1).Caracteres(124) = 21493
    Fuentes(1).Caracteres(252) = 21800
    Fuentes(1).Caracteres(220) = 21801
    Fuentes(1).Caracteres(225) = 21802
    Fuentes(1).Caracteres(233) = 21803
    Fuentes(1).Caracteres(237) = 21804
    Fuentes(1).Caracteres(243) = 21805
    Fuentes(1).Caracteres(250) = 21806
    Fuentes(1).Caracteres(253) = 21807
    Fuentes(1).Caracteres(193) = 21808
    Fuentes(1).Caracteres(201) = 21809
    Fuentes(1).Caracteres(205) = 21810
    Fuentes(1).Caracteres(211) = 21811
    Fuentes(1).Caracteres(218) = 21812
    Fuentes(1).Caracteres(221) = 21813
    Fuentes(1).Caracteres(224) = 21814
    Fuentes(1).Caracteres(232) = 21815
    Fuentes(1).Caracteres(236) = 21816
    Fuentes(1).Caracteres(242) = 21817
    Fuentes(1).Caracteres(249) = 21818
    Fuentes(1).Caracteres(192) = 21819
    Fuentes(1).Caracteres(200) = 21820
    Fuentes(1).Caracteres(204) = 21821
    Fuentes(1).Caracteres(210) = 21822
    Fuentes(1).Caracteres(217) = 21823
    Fuentes(1).Caracteres(241) = 21824
    Fuentes(1).Caracteres(209) = 21825
    Fuentes(1).Caracteres(196) = 25238
    Fuentes(1).Caracteres(194) = 25239
    Fuentes(1).Caracteres(203) = 25240
    Fuentes(1).Caracteres(207) = 25241
    Fuentes(1).Caracteres(214) = 25242
    Fuentes(1).Caracteres(212) = 25243

    Fuentes(2).Tamanio = 9
    Fuentes(2).Caracteres(97) = 21936
    Fuentes(2).Caracteres(108) = 21937
    Fuentes(2).Caracteres(115) = 21938
    Fuentes(2).Caracteres(70) = 21939
    Fuentes(2).Caracteres(48) = 21940
    Fuentes(2).Caracteres(49) = 21941
    Fuentes(2).Caracteres(50) = 21942
    Fuentes(2).Caracteres(51) = 21943
    Fuentes(2).Caracteres(52) = 21944
    Fuentes(2).Caracteres(53) = 21945
    Fuentes(2).Caracteres(54) = 21946
    Fuentes(2).Caracteres(55) = 21947
    Fuentes(2).Caracteres(56) = 21948
    Fuentes(2).Caracteres(57) = 21949
    Fuentes(2).Caracteres(33) = 21950
    Fuentes(2).Caracteres(161) = 21951
    Fuentes(2).Caracteres(42) = 21952

    Fuentes(3).Tamanio = 40
    Fuentes(3).Caracteres(48) = 20428 '0
    Fuentes(3).Caracteres(49) = 20429 '1
    Fuentes(3).Caracteres(50) = 20430 '2
    Fuentes(3).Caracteres(51) = 20431 '3
    Fuentes(3).Caracteres(52) = 20432 '4
    Fuentes(3).Caracteres(53) = 20433 '5
    Fuentes(3).Caracteres(54) = 20434 '6
    Fuentes(3).Caracteres(55) = 20435 '7
    Fuentes(3).Caracteres(56) = 20436 '8
    Fuentes(3).Caracteres(57) = 20437 '9

    For a = 0 To 25
        Fuentes(3).Caracteres(a + 97) = 20477 + a 'Desde la a hasta la z (sin ñ)
    Next a

    For a = 0 To 25
        Fuentes(3).Caracteres(a + 65) = 20445 + a 'Desde la A hasta la Z (sin Ñ)

    Next a

    Fuentes(3).Caracteres(33) = 20413 '!
    Fuentes(3).Caracteres(161) = 20541 '¡
    Fuentes(3).Caracteres(34) = 20414 '"
    Fuentes(3).Caracteres(191) = 8488 '¿
    Fuentes(3).Caracteres(35) = 8332 '#
    Fuentes(3).Caracteres(36) = 20416    '$
    Fuentes(3).Caracteres(37) = 20417 '%
    Fuentes(3).Caracteres(38) = 20418 '&
    Fuentes(3).Caracteres(47) = 20427 '/
    Fuentes(3).Caracteres(92) = 8389 '\
    Fuentes(3).Caracteres(40) = 20420 '(
    Fuentes(3).Caracteres(41) = 20421 ')
    Fuentes(3).Caracteres(61) = 20441 '=
    Fuentes(3).Caracteres(39) = 24930 ''
    Fuentes(3).Caracteres(123) = 24932 ' '
    Fuentes(3).Caracteres(125) = 24931 '}
    Fuentes(3).Caracteres(95) = 20475  '_
    Fuentes(3).Caracteres(45) = 20425 '-
    Fuentes(3).Caracteres(63) = 20443 ' ?
    Fuentes(3).Caracteres(64) = 20444 '@
    Fuentes(3).Caracteres(94) = 20516 '^
    Fuentes(3).Caracteres(91) = 8388 '[
    Fuentes(3).Caracteres(93) = 8390 ']
    Fuentes(3).Caracteres(60) = 20440 '<
    Fuentes(3).Caracteres(62) = 20442 '>
    Fuentes(3).Caracteres(42) = 20422 '*
    Fuentes(3).Caracteres(43) = 20423 '+
    Fuentes(3).Caracteres(46) = 20426 '.
    Fuentes(3).Caracteres(44) = 20510 ',
    Fuentes(3).Caracteres(58) = 8355 ':
    Fuentes(3).Caracteres(59) = 8356 ';
    Fuentes(3).Caracteres(124) = 20504 '|
    Fuentes(3).Caracteres(252) = 24948 '    ü
    Fuentes(3).Caracteres(220) = 24949 'Ü
    Fuentes(3).Caracteres(225) = 8490 'á
    Fuentes(3).Caracteres(233) = 8498 'é
    Fuentes(3).Caracteres(237) = 8502 'í
    Fuentes(3).Caracteres(243) = 8508 'ó
    Fuentes(3).Caracteres(250) = 8515 'ú
    Fuentes(3).Caracteres(253) = 24955 'ý
    Fuentes(3).Caracteres(193) = 8490 'Á
    Fuentes(3).Caracteres(201) = 8498 'É
    Fuentes(3).Caracteres(205) = 8502 'Í
    Fuentes(3).Caracteres(211) = 8508 'Ó
    Fuentes(3).Caracteres(218) = 8515 'Ú
    Fuentes(3).Caracteres(221) = 24961 'Ý
    Fuentes(3).Caracteres(224) = 24962 'à
    Fuentes(3).Caracteres(232) = 24963 'è
    Fuentes(3).Caracteres(236) = 24964 'ì
    Fuentes(3).Caracteres(242) = 24965 'ò
    Fuentes(3).Caracteres(249) = 24966 'ù
    Fuentes(3).Caracteres(192) = 24967 'ü
    Fuentes(3).Caracteres(200) = 24968 '
    Fuentes(3).Caracteres(204) = 24969 '
    Fuentes(3).Caracteres(210) = 24970 '
    Fuentes(3).Caracteres(217) = 24971 '
    Fuentes(3).Caracteres(241) = 8506 'ñ
    Fuentes(3).Caracteres(209) = 24872 '
    Fuentes(3).Caracteres(196) = 24874 '
    Fuentes(3).Caracteres(194) = 24875 '
    Fuentes(3).Caracteres(203) = 24876 '
    Fuentes(3).Caracteres(207) = 24877 '
    Fuentes(3).Caracteres(214) = 24878 '
    Fuentes(3).Caracteres(212) = 24879 '

    Fuentes(3).Caracteres(172) = 20552 '¬
    Fuentes(3).Caracteres(186) = 20556 'º

    Fuentes(4).Tamanio = 3
    Fuentes(4).Caracteres(48) = 13852
    Fuentes(4).Caracteres(49) = 13853
    Fuentes(4).Caracteres(50) = 13854
    Fuentes(4).Caracteres(51) = 13855
    Fuentes(4).Caracteres(52) = 13856
    Fuentes(4).Caracteres(53) = 13857
    Fuentes(4).Caracteres(54) = 13858
    Fuentes(4).Caracteres(55) = 13859
    Fuentes(4).Caracteres(56) = 13860
    Fuentes(4).Caracteres(57) = 13861

    For a = 0 To 25
        Fuentes(4).Caracteres(a + 97) = 13800 + a
    Next a

    For a = 0 To 25
        Fuentes(4).Caracteres(a + 65) = 13826 + a
    Next a

    Fuentes(4).Caracteres(33) = 13862
    Fuentes(4).Caracteres(161) = 13863
    Fuentes(4).Caracteres(34) = 13864
    Fuentes(4).Caracteres(36) = 13865
    Fuentes(4).Caracteres(191) = 13866
    Fuentes(4).Caracteres(35) = 13867
    Fuentes(4).Caracteres(36) = 13868
    Fuentes(4).Caracteres(37) = 13869
    Fuentes(4).Caracteres(38) = 13870
    Fuentes(4).Caracteres(47) = 13871
    Fuentes(4).Caracteres(92) = 13872
    Fuentes(4).Caracteres(40) = 13873
    Fuentes(4).Caracteres(41) = 13874
    Fuentes(4).Caracteres(61) = 13875
    Fuentes(4).Caracteres(39) = 13876
    Fuentes(4).Caracteres(123) = 13877
    Fuentes(4).Caracteres(125) = 13878
    Fuentes(4).Caracteres(95) = 13879
    Fuentes(4).Caracteres(45) = 13880
    Fuentes(4).Caracteres(63) = 13865
    Fuentes(4).Caracteres(64) = 13881
    Fuentes(4).Caracteres(94) = 13882
    Fuentes(4).Caracteres(91) = 13883
    Fuentes(4).Caracteres(93) = 13884
    Fuentes(4).Caracteres(60) = 13885
    Fuentes(4).Caracteres(62) = 13886
    Fuentes(4).Caracteres(42) = 13887
    Fuentes(4).Caracteres(43) = 13888
    Fuentes(4).Caracteres(46) = 13889
    Fuentes(4).Caracteres(44) = 13890
    Fuentes(4).Caracteres(58) = 13891
    Fuentes(4).Caracteres(59) = 13892
    Fuentes(4).Caracteres(124) = 13893

    Fuentes(4).Caracteres(252) = 24948 '    ü
    Fuentes(4).Caracteres(220) = 24949 'Ü
    Fuentes(3).Caracteres(225) = 8490 'á
    Fuentes(3).Caracteres(233) = 8498 'é
    Fuentes(3).Caracteres(237) = 8502 'í
    Fuentes(3).Caracteres(243) = 8508 'ó
    Fuentes(3).Caracteres(250) = 8515 'ú
    Fuentes(3).Caracteres(253) = 24955 'ý
    Fuentes(3).Caracteres(193) = 8490 'Á
    Fuentes(3).Caracteres(201) = 8498 'É
    Fuentes(3).Caracteres(205) = 8502 'Í
    Fuentes(3).Caracteres(211) = 8508 'Ó
    Fuentes(3).Caracteres(218) = 8515 'Ú
    Fuentes(3).Caracteres(221) = 24961 'Ý
    Fuentes(3).Caracteres(224) = 24962 'à
    Fuentes(3).Caracteres(232) = 24963 'è
    Fuentes(3).Caracteres(236) = 24964 'ì
    Fuentes(3).Caracteres(242) = 24965 'ò
    Fuentes(3).Caracteres(249) = 24966 'ù
    Fuentes(3).Caracteres(192) = 24967 'ü
    Fuentes(3).Caracteres(200) = 24968 '
    Fuentes(3).Caracteres(204) = 24969 '
    Fuentes(3).Caracteres(210) = 24970 '
    Fuentes(3).Caracteres(217) = 24971 '
    Fuentes(3).Caracteres(241) = 8506 'ñ
    Fuentes(3).Caracteres(209) = 24872 '
    Fuentes(3).Caracteres(196) = 24874 '
    Fuentes(3).Caracteres(194) = 24875 '
    Fuentes(3).Caracteres(203) = 24876 '
    Fuentes(3).Caracteres(207) = 24877 '
    Fuentes(3).Caracteres(214) = 24878 '
    Fuentes(3).Caracteres(212) = 24879 '

    Fuentes(3).Caracteres(172) = 20552 '¬
    Fuentes(3).Caracteres(186) = 20556 'º

    Fuentes(1).Caracteres(196) = 25238
    Fuentes(1).Caracteres(194) = 25239
    Fuentes(1).Caracteres(203) = 25240
    Fuentes(1).Caracteres(207) = 25241
    Fuentes(1).Caracteres(214) = 25242
    Fuentes(1).Caracteres(212) = 25243

    Fuentes(5).Tamanio = 50
    Fuentes(5).Caracteres(48) = 30127
    Fuentes(5).Caracteres(49) = 30128
    Fuentes(5).Caracteres(50) = 30129
    Fuentes(5).Caracteres(51) = 30130
    Fuentes(5).Caracteres(52) = 30131
    Fuentes(5).Caracteres(53) = 30132
    Fuentes(5).Caracteres(54) = 30133
    Fuentes(5).Caracteres(55) = 30134
    Fuentes(5).Caracteres(56) = 30135
    Fuentes(5).Caracteres(57) = 30136

    For a = 0 To 25
        Fuentes(5).Caracteres(a + 97) = 30176 + a
    Next a

    For a = 0 To 25
        Fuentes(5).Caracteres(a + 65) = 30144 + a
    Next a

    Fuentes(5).Caracteres(33) = 30112 '!
    Fuentes(5).Caracteres(161) = 20541 '¡
    Fuentes(5).Caracteres(34) = 30113 '"
    Fuentes(5).Caracteres(191) = 8488 '¿
    Fuentes(5).Caracteres(35) = 8332 '#
    Fuentes(5).Caracteres(36) = 20416    '$
    Fuentes(5).Caracteres(37) = 20417 '%
    Fuentes(5).Caracteres(38) = 20418 '&
    Fuentes(5).Caracteres(47) = 20427 '/
    Fuentes(5).Caracteres(92) = 8389 '\
    Fuentes(5).Caracteres(40) = 30119 '(
    Fuentes(5).Caracteres(41) = 30120 ')
    Fuentes(5).Caracteres(61) = 30140 '=
    Fuentes(5).Caracteres(39) = 24930 ''
    Fuentes(5).Caracteres(123) = 24932 ' '
    Fuentes(5).Caracteres(125) = 24931 '}
    Fuentes(5).Caracteres(95) = 20475  '_
    Fuentes(5).Caracteres(45) = 20425 '-
    Fuentes(5).Caracteres(63) = 20443 ' ?
    Fuentes(5).Caracteres(64) = 20444 '@
    Fuentes(5).Caracteres(94) = 20516 '^
    Fuentes(5).Caracteres(91) = 8388 '[
    Fuentes(5).Caracteres(93) = 8390 ']
    Fuentes(5).Caracteres(60) = 30139 '<
    Fuentes(5).Caracteres(62) = 30141 '>
    Fuentes(5).Caracteres(42) = 20422 '*
    Fuentes(5).Caracteres(43) = 20423 '+
    Fuentes(5).Caracteres(46) = 20426 '.
    Fuentes(5).Caracteres(44) = 20510 ',
    Fuentes(5).Caracteres(58) = 8355 ':
    Fuentes(5).Caracteres(59) = 8356 ';
    Fuentes(5).Caracteres(124) = 20504 '|
    Fuentes(5).Caracteres(252) = 24948 '    ü
    Fuentes(5).Caracteres(220) = 24949 'Ü
    Fuentes(5).Caracteres(225) = 30304 'á
    Fuentes(5).Caracteres(233) = 30312 'é
    Fuentes(5).Caracteres(237) = 30316 'í
    Fuentes(5).Caracteres(243) = 30322 'ó
    Fuentes(5).Caracteres(250) = 30329 'ú
    Fuentes(5).Caracteres(253) = 24955 'ý
    Fuentes(5).Caracteres(193) = 30272 'Á
    Fuentes(5).Caracteres(201) = 30280 'É
    Fuentes(5).Caracteres(205) = 8502 'Í
    Fuentes(5).Caracteres(211) = 30290 'Ó
    Fuentes(5).Caracteres(218) = 8515 'Ú
    Fuentes(5).Caracteres(221) = 24961 'Ý
    Fuentes(5).Caracteres(224) = 24962 'à
    Fuentes(5).Caracteres(232) = 24963 'è
    Fuentes(5).Caracteres(236) = 24964 'ì
    Fuentes(5).Caracteres(242) = 24965 'ò
    Fuentes(5).Caracteres(249) = 24966 'ù
    Fuentes(5).Caracteres(192) = 24967 'ü
    Fuentes(5).Caracteres(200) = 24968 '
    Fuentes(5).Caracteres(204) = 24969 '
    Fuentes(5).Caracteres(210) = 24970 '
    Fuentes(5).Caracteres(217) = 24971 '
    Fuentes(5).Caracteres(241) = 30288 'ñ
    Fuentes(5).Caracteres(209) = 24872 '
    Fuentes(5).Caracteres(196) = 24874 '
    Fuentes(5).Caracteres(194) = 30305 'â
    Fuentes(5).Caracteres(203) = 24876 '
    Fuentes(5).Caracteres(207) = 24877 '
    Fuentes(5).Caracteres(214) = 24878 '
    Fuentes(5).Caracteres(212) = 24879 '

    Fuentes(5).Caracteres(172) = 20552 '¬
    Fuentes(5).Caracteres(186) = 20556 'º

    Fuentes(6).Tamanio = 50
    Fuentes(6).Caracteres(48) = 45866
    Fuentes(6).Caracteres(49) = 45867
    Fuentes(6).Caracteres(50) = 45868
    Fuentes(6).Caracteres(51) = 45869
    Fuentes(6).Caracteres(52) = 45870
    Fuentes(6).Caracteres(53) = 45871
    Fuentes(6).Caracteres(54) = 45872
    Fuentes(6).Caracteres(55) = 45873
    Fuentes(6).Caracteres(56) = 45874
    Fuentes(6).Caracteres(57) = 45875

    For a = 0 To 25
        Fuentes(6).Caracteres(a + 97) = 45915 + a
    Next a

    For a = 0 To 25
        Fuentes(6).Caracteres(a + 65) = 45883 + a
    Next a

    Fuentes(6).Caracteres(33) = 13862
    Fuentes(6).Caracteres(161) = 13863
    Fuentes(6).Caracteres(34) = 13864
    Fuentes(6).Caracteres(36) = 13865
    Fuentes(6).Caracteres(191) = 13866
    Fuentes(6).Caracteres(35) = 13867
    Fuentes(6).Caracteres(36) = 13868
    Fuentes(6).Caracteres(37) = 13869
    Fuentes(6).Caracteres(38) = 13870
    Fuentes(6).Caracteres(47) = 13871
    Fuentes(6).Caracteres(92) = 13872
    Fuentes(6).Caracteres(40) = 13873
    Fuentes(6).Caracteres(41) = 13874
    Fuentes(6).Caracteres(61) = 13875
    Fuentes(6).Caracteres(39) = 13876
    Fuentes(6).Caracteres(123) = 13877
    Fuentes(6).Caracteres(125) = 13878
    Fuentes(6).Caracteres(95) = 13879
    Fuentes(6).Caracteres(45) = 13880
    Fuentes(6).Caracteres(63) = 13865
    Fuentes(6).Caracteres(64) = 13881
    Fuentes(6).Caracteres(94) = 13882
    Fuentes(6).Caracteres(91) = 13883
    Fuentes(6).Caracteres(93) = 13884
    Fuentes(6).Caracteres(60) = 13885
    Fuentes(6).Caracteres(62) = 13886
    Fuentes(6).Caracteres(42) = 13887
    Fuentes(6).Caracteres(43) = 13888
    Fuentes(6).Caracteres(46) = 13889
    Fuentes(6).Caracteres(44) = 13890
    Fuentes(6).Caracteres(58) = 13891
    Fuentes(6).Caracteres(59) = 13892
    Fuentes(6).Caracteres(124) = 13893

    Fuentes(6).Caracteres(252) = 18200
    Fuentes(6).Caracteres(220) = 18201
    Fuentes(6).Caracteres(225) = 18202
    Fuentes(6).Caracteres(233) = 18203
    Fuentes(6).Caracteres(237) = 18204
    Fuentes(6).Caracteres(243) = 18205
    Fuentes(6).Caracteres(250) = 18206
    Fuentes(6).Caracteres(253) = 18207
    Fuentes(6).Caracteres(193) = 18208
    Fuentes(6).Caracteres(201) = 18209
    Fuentes(6).Caracteres(205) = 18210
    Fuentes(6).Caracteres(211) = 18211
    Fuentes(6).Caracteres(218) = 18212
    Fuentes(6).Caracteres(221) = 18213
    Fuentes(6).Caracteres(224) = 18214
    Fuentes(6).Caracteres(232) = 18215
    Fuentes(6).Caracteres(236) = 18216
    Fuentes(6).Caracteres(242) = 18217
    Fuentes(6).Caracteres(249) = 18218
    Fuentes(6).Caracteres(192) = 18219
    Fuentes(6).Caracteres(200) = 18220
    Fuentes(6).Caracteres(204) = 18221
    Fuentes(6).Caracteres(210) = 18222
    Fuentes(6).Caracteres(217) = 18223
    Fuentes(6).Caracteres(241) = 18224
    Fuentes(6).Caracteres(209) = 18225

End Sub

Function Engine_Text_Height(Texto As String, Optional multi As Boolean = False, Optional font As Byte = 1) As Integer

    Dim a, b, c, d, e, f As Integer

    Dim graf As grh
  
    If multi = False Then
        Engine_Text_Height = 0
    Else
        e = 0
        f = 0

        If font = 1 Then

            For a = 1 To Len(Texto)
                b = Asc(mid(Texto, a, 1))
                graf.GrhIndex = Fuentes(1).Caracteres(b)

                If b = 32 Or b = 13 Then
                    If e >= 20 Then 'reemplazar por lo que os plazca
                        f = f + 1
                        e = 0
                        d = 0
                    Else

                        If b = 32 Then
                            d = d + 4

                        End If

                    End If

                    'Else
                    'If graf.GrhIndex > 12 Then
                End If

                e = e + 1
            Next a

        Else
    
            For a = 1 To Len(Texto)
                b = Asc(mid(Texto, a, 1))
                graf.GrhIndex = Fuentes(font).Caracteres(b)

                If b = 32 Or b = 13 Then
                    If e >= 20 Then 'reemplazar por lo que os plazca
                        f = f + 1
                        e = 0
                        d = 0
                    Else

                        If b = 32 Then
                            d = d + 4

                        End If

                    End If

                    'Else
                    'If graf.GrhIndex > 12 Then
                End If

                e = e + 1
            Next a
  
        End If

        Engine_Text_Height = f * 14
  
    End If

End Function

Sub Engine_Text_Render_LetraGrande(Texto As String, x As Integer, y As Integer, ByRef text_color() As Long, Optional ByVal font_index As Integer = 1, Optional multi_line As Boolean = False, Optional charindex As Integer = 0, Optional ByVal Alpha As Byte = 255)

    On Error Resume Next

    Dim a, b, c, d, e, f As Integer

    Dim graf          As grh

    Dim temp_array(3) As Long 'Si le queres dar color a la letra pasa este parametro dsp xD

    temp_array(0) = text_color(0)

    If charindex = 0 Then
        a = 255
    Else
        a = charlist(charindex).AlphaText

    End If

    If Alpha <> 255 Then
        a = Alpha

    End If

    Dim r, g As Byte

    r = (temp_array(0) And 16711680) / 65536
    g = (temp_array(0) And 65280) / 256
    b = temp_array(0) And 255
             
    temp_array(0) = D3DColorARGB(a, r, g, b)

    temp_array(1) = temp_array(0)
    temp_array(2) = temp_array(0)
    temp_array(3) = temp_array(0)

    Dim i              As Long

    Dim removedDialogs As Long

    For i = 0 To dialogCount - 1

        'Decrease index to prevent jumping over a dialog
        'Crappy VB will cache the limit of the For loop, so even if it changed, it won't matter
        With dialogs(i - removedDialogs)

            If ((GetTickCount() And &H7FFFFFFF) - .startTime) >= .lifeTime Then
                Call Char_Dialog_Remove(.charindex, charindex)
                             
                If charlist(charindex).AlphaText = 0 Then
                    removedDialogs = removedDialogs + 1

                End If

            Else
            
            End If

        End With

    Next i

    Dim Sombra(3) As Long 'Sombra

    Sombra(0) = D3DColorARGB(a, r / 6, g / 6, b / 6)
    Sombra(1) = Sombra(0)
    Sombra(2) = Sombra(0)
    Sombra(3) = Sombra(0)

    If (Len(Texto) = 0) Then Exit Sub

    d = 0

    If multi_line = False Then
        e = 0
        f = 0

        For a = 1 To Len(Texto)
            b = Asc(mid(Texto, a, 1))
            graf.GrhIndex = Fuentes(font_index).Caracteres(b)

            If b = 32 Or b = 13 Then
                If e >= 35 Then 'reemplazar por lo que os plazca
                    f = f + 1
                    e = 0
                    d = 0
                Else

                    If b = 32 Then d = d + 30

                End If

            Else

                If graf.GrhIndex > 12 Then

                    'mega sombra O-matica
                    graf.GrhIndex = Fuentes(font_index).Caracteres(b)

                    If font_index <> 3 Then

                        'Call Draw_GrhColor(graf.GrhIndex, (x + d), y + f * 14, Sombra())
                    End If

                    Call Draw_GrhColor(graf.GrhIndex, (x + d) + 1, y + 1 + f * 14, temp_array())
                
                    ' graf.grhindex = Fuentes(font_index).Caracteres(b)
                    ' Grh_Render graf, (X + d), Y + f * 14, temp_array, False, False, False '14 es el height de esta fuente dsp lo hacemos dinamico
                    d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth - 70

                End If

            End If

            e = e + 1
        Next a

    Else
        e = 0
        f = 0

        For a = 1 To Len(Texto)
            b = Asc(mid(Texto, a, 1))
            graf.GrhIndex = Fuentes(font_index).Caracteres(b)

            If b = 32 Or b = 13 Then
                If e >= 33 Then 'reemplazar por lo que os plazca
                    f = f + 1
                    e = 0
                    d = 0
                Else

                    If b = 32 Then d = d + 2

                End If

            Else

                If graf.GrhIndex > 12 Then

                    'mega sombra O-matica
                    graf.GrhIndex = Fuentes(font_index).Caracteres(b)
                    ' Call Draw_GrhColor(graf.GrhIndex, (x + d) + 1, y + 1 + f * 14, Sombra())
                    Call Draw_GrhColor(graf.GrhIndex, (x + d), y + f * 14, temp_array())
                
                    ' graf.grhindex = Fuentes(font_index).Caracteres(b)
                    'Grh_Render graf, (x + d), y + f * 14, temp_array, False, False, False '14 es el height de esta fuente dsp lo hacemos dinamico
                    If font_index = 5 Then
                        d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth - 50
                    Else
                        d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth

                    End If

                End If

            End If

            e = e + 1
        Next a

    End If

End Sub

Sub Engine_Text_Render_LetraChica(Texto As String, x As Integer, y As Integer, ByRef text_color() As Long, Optional ByVal font_index As Integer = 1, Optional multi_line As Boolean = False, Optional charindex As Integer = 0, Optional ByVal Alpha As Byte = 255)

    On Error Resume Next

    Dim a, b, c, d, e, f As Integer

    Dim graf          As grh

    Dim temp_array(3) As Long 'Si le queres dar color a la letra pasa este parametro dsp xD

    temp_array(0) = text_color(0)

    If charindex = 0 Then
        a = 255
    Else
        a = charlist(charindex).AlphaText

    End If

    If Alpha <> 255 Then
        a = Alpha

    End If

    Dim r, g As Byte

    b = (temp_array(0) And 16711680) / 65536
    g = (temp_array(0) And 65280) / 256
    r = temp_array(0) And 255
             
    temp_array(0) = D3DColorARGB(a, r, g, b)

    temp_array(1) = temp_array(0)
    temp_array(2) = temp_array(0)
    temp_array(3) = temp_array(0)

    Dim i              As Long

    Dim removedDialogs As Long

    For i = 0 To dialogCount - 1

        'Decrease index to prevent jumping over a dialog
        'Crappy VB will cache the limit of the For loop, so even if it changed, it won't matter
        With dialogs(i - removedDialogs)

            If ((GetTickCount() And &H7FFFFFFF) - .startTime) >= .lifeTime Then
                Call Char_Dialog_Remove(.charindex, charindex)
                             
                If a <= 0 Then
                    removedDialogs = removedDialogs + 1

                End If

            Else
            
            End If

        End With

    Next i

    Dim Sombra(3) As Long 'Sombra

    Sombra(0) = D3DColorARGB(a, r / 6, g / 6, b / 6)
    Sombra(1) = Sombra(0)
    Sombra(2) = Sombra(0)
    Sombra(3) = Sombra(0)

    If (Len(Texto) = 0) Then Exit Sub

    d = 0

    If multi_line = False Then
        e = 0
        f = 0

        For a = 1 To Len(Texto)
            b = Asc(mid(Texto, a, 1))
            graf.GrhIndex = Fuentes(font_index).Caracteres(b)

            If b = 32 Or b = 13 Then
                If e >= 30 Then 'reemplazar por lo que os plazca
                    f = f + 1
                    e = 0
                    d = 0
                Else

                    If b = 32 Then d = d + 2

                End If

            Else

                If graf.GrhIndex > 12 Then

                    'mega sombra O-matica
                    graf.GrhIndex = Fuentes(font_index).Caracteres(b)

                    If font_index <> 3 Then

                        'Call Draw_GrhColor(graf.GrhIndex, (x + d), y + f * 14, Sombra())
                    End If

                    'Call Draw_GrhColor(graf.GrhIndex, (x + d) + 1, y + 1 + f * 14, temp_array())
                
                    Call InitGrh(graf, graf.GrhIndex)
                    'Call Draw_Grh(graf, (x + d) + 1, y + 1 + f * 14, 0, 0, Sombra(), True, 0, 0, 0)
                    Call Draw_Grh(graf, (x + d) + 1, y + 1 + f * 14, 0, 0, temp_array(), True, 0, 0, 0)
                
                    ' graf.grhindex = Fuentes(font_index).Caracteres(b)
                    ' Grh_Render graf, (X + d), Y + f * 14, temp_array, False, False, False '14 es el height de esta fuente dsp lo hacemos dinamico
                    d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth

                End If

            End If

            e = e + 1
        Next a

    Else
        e = 0
        f = 0

        For a = 1 To Len(Texto)
            b = Asc(mid(Texto, a, 1))
            graf.GrhIndex = Fuentes(font_index).Caracteres(b)

            If b = 32 Or b = 13 Then
                If e >= 33 Then 'reemplazar por lo que os plazca
                    f = f + 1
                    e = 0
                    d = 0
                Else

                    If b = 32 Then d = d + 2

                End If

            Else

                If graf.GrhIndex > 12 Then

                    'mega sombra O-matica
                    graf.GrhIndex = Fuentes(font_index).Caracteres(b)
                    ' Call Draw_GrhColor(graf.GrhIndex, (x + d) + 1, y + 1 + f * 14, Sombra())
                    Call Draw_GrhColor(graf.GrhIndex, (x + d), y + f * 14, temp_array())
                
                    ' graf.grhindex = Fuentes(font_index).Caracteres(b)
                    'Grh_Render graf, (x + d), y + f * 14, temp_array, False, False, False '14 es el height de esta fuente dsp lo hacemos dinamico
                    If font_index = 4 Then
                        d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth - 1
                    Else
                        d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth

                    End If

                End If

            End If

            e = e + 1
        Next a

    End If

End Sub

Sub Engine_Text_Render(Texto As String, x As Integer, y As Integer, ByRef text_color() As Long, Optional ByVal font_index As Integer = 1, Optional multi_line As Boolean = False, Optional charindex As Integer = 0, Optional ByVal Alpha As Byte = 255)

    On Error Resume Next

    Dim a, b, c, d, e, f As Integer

    Dim graf          As grh

    Dim temp_array(3) As Long 'Si le queres dar color a la letra pasa este parametro dsp xD

    temp_array(0) = text_color(0)

    If charindex = 0 Then
        a = 255
    Else
        a = charlist(charindex).AlphaText

    End If

    If Alpha <> 255 Then
        a = Alpha

    End If

    Dim r, g As Byte

    b = (temp_array(0) And 16711680) / 65536
    g = (temp_array(0) And 65280) / 256
    r = temp_array(0) And 255
             
    temp_array(0) = D3DColorARGB(a, r, g, b)

    temp_array(1) = temp_array(0)
    temp_array(2) = temp_array(0)
    temp_array(3) = temp_array(0)

    Dim i              As Long

    Dim removedDialogs As Long

    For i = 0 To dialogCount - 1

        'Decrease index to prevent jumping over a dialog
        'Crappy VB will cache the limit of the For loop, so even if it changed, it won't matter
        With dialogs(i - removedDialogs)

            If ((GetTickCount() And &H7FFFFFFF) - .startTime) >= .lifeTime Then
                Call Char_Dialog_Remove(.charindex, charindex)
                             
                If a <= 0 Then
                    removedDialogs = removedDialogs + 1

                End If

            Else
            
            End If

        End With

    Next i

    Dim Sombra(3) As Long 'Sombra

    Sombra(0) = D3DColorARGB(a, r / 4, g / 4, b / 4)
    Sombra(1) = Sombra(0)
    Sombra(2) = Sombra(0)
    Sombra(3) = Sombra(0)

    If (Len(Texto) = 0) Then Exit Sub

    d = 0

    If multi_line = False Then
        e = 0
        f = 0

        For a = 1 To Len(Texto)
            b = Asc(mid(Texto, a, 1))
            graf.GrhIndex = Fuentes(font_index).Caracteres(b)

            If b = 32 Or b = 13 Then
                If e >= 35 Then 'reemplazar por lo que os plazca
                    f = f + 1
                    e = 0
                    d = 0
                Else

                    If b = 32 Then d = d + 4

                End If

            Else

                If graf.GrhIndex > 12 Then

                    'mega sombra O-matica
                    graf.GrhIndex = Fuentes(font_index).Caracteres(b)

                    If font_index <> 3 Then
                        Call Draw_GrhColor(graf.GrhIndex, (x + d), y + f * 14, Sombra())

                    End If

                    Call Draw_GrhColor(graf.GrhIndex, (x + d) + 1, y + 1 + f * 14, temp_array())
                
                    ' graf.grhindex = Fuentes(font_index).Caracteres(b)
                    ' Grh_Render graf, (X + d), Y + f * 14, temp_array, False, False, False '14 es el height de esta fuente dsp lo hacemos dinamico
                    d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth

                End If

            End If

            e = e + 1
        Next a

    Else
        e = 0
        f = 0

        For a = 1 To Len(Texto)
            b = Asc(mid(Texto, a, 1))
            graf.GrhIndex = Fuentes(font_index).Caracteres(b)

            If b = 32 Or b = 13 Then
                If e >= 20 Then 'reemplazar por lo que os plazca
                    f = f + 1
                    e = 0
                    d = 0
                Else

                    If b = 32 Then d = d + 4

                End If

            Else

                If graf.GrhIndex > 12 Then

                    'mega sombra O-matica
                    graf.GrhIndex = Fuentes(font_index).Caracteres(b)
                    Call Draw_GrhColor(graf.GrhIndex, (x + d) + 1, y + 1 + f * 14, Sombra())
                    Call Draw_GrhColor(graf.GrhIndex, (x + d), y + f * 14, temp_array())
                
                    ' graf.grhindex = Fuentes(font_index).Caracteres(b)
                    'Grh_Render graf, (x + d), y + f * 14, temp_array, False, False, False '14 es el height de esta fuente dsp lo hacemos dinamico
                    If font_index = 4 Then
                        d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth - 1
                    Else
                        d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth

                    End If

                End If

            End If

            e = e + 1
        Next a

    End If

End Sub

Sub Engine_Text_RenderGrande(Texto As String, x As Integer, y As Integer, ByRef text_color() As Long, Optional ByVal font_index As Integer = 1, Optional multi_line As Boolean = False, Optional charindex As Integer = 0, Optional ByVal Alpha As Byte = 255)

    On Error Resume Next

    Dim a, b, c, d, e, f As Integer

    Dim graf          As grh

    Dim temp_array(3) As Long 'Si le queres dar color a la letra pasa este parametro dsp xD

    temp_array(0) = text_color(0)

    If charindex = 0 Then
        a = 255
    Else
        a = charlist(charindex).AlphaText

    End If

    If Alpha <> 255 Then
        a = Alpha

    End If

    Dim r, g As Byte

    r = (temp_array(0) And 16711680) / 65536
    g = (temp_array(0) And 65280) / 256
    b = temp_array(0) And 255
             
    temp_array(0) = D3DColorARGB(a, r, g, b)

    temp_array(1) = temp_array(0)
    temp_array(2) = temp_array(0)
    temp_array(3) = temp_array(0)

    Dim i              As Long

    Dim removedDialogs As Long

    For i = 0 To dialogCount - 1

        'Decrease index to prevent jumping over a dialog
        'Crappy VB will cache the limit of the For loop, so even if it changed, it won't matter
        With dialogs(i - removedDialogs)

            If ((GetTickCount() And &H7FFFFFFF) - .startTime) >= .lifeTime Then
                Call Char_Dialog_Remove(.charindex, charindex)
                             
                If a <= 0 Then
                    removedDialogs = removedDialogs + 1

                End If

            Else
            
            End If

        End With

    Next i

    Dim Sombra(3) As Long 'Sombra

    Sombra(0) = D3DColorARGB(a, r / 6, g / 6, b / 6)
    Sombra(1) = Sombra(0)
    Sombra(2) = Sombra(0)
    Sombra(3) = Sombra(0)

    If (Len(Texto) = 0) Then Exit Sub

    d = 0

    If multi_line = False Then
        e = 0
        f = 0

        For a = 1 To Len(Texto)
            b = Asc(mid(Texto, a, 1))
            graf.GrhIndex = Fuentes(font_index).Caracteres(b)

            If b = 32 Or b = 13 Then
                If e >= 35 Then 'reemplazar por lo que os plazca
                    f = f + 1
                    e = 0
                    d = 0
                Else

                    If b = 32 Then d = d + 12

                End If

            Else

                If graf.GrhIndex > 12 Then

                    'mega sombra O-matica
                    graf.GrhIndex = Fuentes(font_index).Caracteres(b)

                    If font_index <> 3 Then
                        Call Draw_GrhColor(graf.GrhIndex, (x + d), y + f * 14, Sombra())

                    End If

                    Call Draw_GrhColor(graf.GrhIndex, (x + d) + 1, y + 1 + f * 14, temp_array())
                
                    ' graf.grhindex = Fuentes(font_index).Caracteres(b)
                    ' Grh_Render graf, (X + d), Y + f * 14, temp_array, False, False, False '14 es el height de esta fuente dsp lo hacemos dinamico
                    d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth

                End If

            End If

            e = e + 1
        Next a

    Else
        e = 0
        f = 0

        For a = 1 To Len(Texto)
            b = Asc(mid(Texto, a, 1))
            graf.GrhIndex = Fuentes(font_index).Caracteres(b)

            If b = 32 Or b = 13 Then
                If e >= 10 Then 'reemplazar por lo que os plazca
                    f = f + 3
                    e = 0
                    d = 0
                Else

                    If b = 32 Then d = d + 12

                End If

            Else

                If graf.GrhIndex > 12 Then

                    'mega sombra O-matica
                    graf.GrhIndex = Fuentes(font_index).Caracteres(b)
                    'Call Draw_GrhColor(graf.GrhIndex, (x + d) + 1, y + 1 + f * 14, Sombra())
                    Call Draw_GrhColor(graf.GrhIndex, (x + d), y + f * 14, temp_array())
                
                    ' graf.grhindex = Fuentes(font_index).Caracteres(b)
                    'Grh_Render graf, (x + d), y + f * 14, temp_array, False, False, False '14 es el height de esta fuente dsp lo hacemos dinamico
                    If font_index = 4 Then
                        d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth
                    Else
                        d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth

                    End If

                End If

            End If

            e = e + 1
        Next a

    End If

End Sub

Sub Engine_Text_Render2(Texto As String, x As Integer, y As Integer, ByRef text_color As Long, Optional ByVal font_index As Integer = 1, Optional multi_line As Boolean = False, Optional charindex As Long = 0)

    On Error Resume Next

    Dim a, b, c, d, e, f As Integer

    Dim graf          As grh

    Dim temp_array(3) As Long 'Si le queres dar color a la letra pasa este parametro dsp xD

    temp_array(0) = text_color

    Dim r, g As Byte

    r = (temp_array(0) And 16711680) / 65536
    g = (temp_array(0) And 65280) / 256
    b = temp_array(0) And 255
             
    temp_array(0) = text_color

    temp_array(1) = temp_array(0)
    temp_array(2) = temp_array(0)
    temp_array(3) = temp_array(0)

    Dim Sombra(3) As Long 'Sombra

    Sombra(0) = D3DColorARGB(charindex, r / 6, g / 6, b / 6)
    Sombra(1) = Sombra(0)
    Sombra(2) = Sombra(0)
    Sombra(3) = Sombra(0)

    If (Len(Texto) = 0) Then Exit Sub

    d = 0

    If multi_line = False Then
        e = 0
        f = 0

        For a = 1 To Len(Texto)
            b = Asc(mid(Texto, a, 1))
            graf.GrhIndex = Fuentes(font_index).Caracteres(b)

            If b = 32 Or b = 13 Then
                If e >= 35 Then 'reemplazar por lo que os plazca
                    f = f + 1
                    e = 0
                    d = 0
                Else

                    If b = 32 Then d = d + 4

                End If

            Else

                If graf.GrhIndex > 12 Then

                    'mega sombra O-matica
                    graf.GrhIndex = Fuentes(font_index).Caracteres(b)

                    If font_index <> 3 Then
                        Call Draw_GrhColor(graf.GrhIndex, (x + d) + 1, y + 1 + f * 14, Sombra())

                    End If

                    Call Draw_GrhColor(graf.GrhIndex, (x + d), y + f * 14, temp_array())
                
                    ' graf.grhindex = Fuentes(font_index).Caracteres(b)
                    ' Grh_Render graf, (X + d), Y + f * 14, temp_array, False, False, False '14 es el height de esta fuente dsp lo hacemos dinamico
                    d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth

                End If

            End If

            e = e + 1
        Next a

    Else
        e = 0
        f = 0

        For a = 1 To Len(Texto)
            b = Asc(mid(Texto, a, 1))
            graf.GrhIndex = Fuentes(font_index).Caracteres(b)

            If b = 32 Or b = 13 Then
                If e >= 20 Then 'reemplazar por lo que os plazca
                    f = f + 1
                    e = 0
                    d = 0
                Else

                    If b = 32 Then d = d + 4

                End If

            Else

                If graf.GrhIndex > 12 Then

                    'mega sombra O-matica
                    graf.GrhIndex = Fuentes(font_index).Caracteres(b)
                    Call Draw_GrhColor(graf.GrhIndex, (x + d) + 1, y + 1 + f * 14, Sombra())
                    Call Draw_GrhColor(graf.GrhIndex, (x + d), y + f * 14, temp_array())
                
                    ' graf.grhindex = Fuentes(font_index).Caracteres(b)
                    'Grh_Render graf, (x + d), y + f * 14, temp_array, False, False, False '14 es el height de esta fuente dsp lo hacemos dinamico
                    If font_index <> 3 Then
                        d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth
                    Else
                        d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth

                    End If

                End If

            End If

            e = e + 1
        Next a

    End If

End Sub

Sub Engine_Text_Render_Efect(charindex As Integer, Texto As String, x As Integer, y As Integer, ByRef text_color() As Long, Optional ByVal font_index As Integer = 1, Optional multi_line As Boolean = False, Optional ByVal Alpha As Byte = 255)

    Dim a, b, c, d, e, f As Integer

    Dim graf As grh

    If (Len(Texto) = 0) Then Exit Sub

    d = 0
    e = 0
    f = 0

    Dim r, g As Byte

    r = (text_color(0) And 16711680) / 65536
    g = (text_color(0) And 65280) / 256
    b = text_color(0) And 255

    Dim Sombra(3) As Long 'Sombra

    Sombra(0) = D3DColorARGB(Alpha, r / 6, g / 6, b / 6)
    Sombra(1) = Sombra(0)
    Sombra(2) = Sombra(0)
    Sombra(3) = Sombra(0)

    For a = 1 To Len(Texto)
        b = Asc(mid(Texto, a, 1))
        graf.GrhIndex = Fuentes(font_index).Caracteres(b)

        If b = 32 Or b = 13 Then
            If e >= 20 Then 'reemplazar por lo que os plazca
                f = f + 1
                e = 0
                d = 0
            Else

                If b = 32 Then d = d + 4

            End If

        Else

            If graf.GrhIndex > 12 Then

                'mega sombra O-matica
                graf.GrhIndex = Fuentes(font_index).Caracteres(b)
                
                Call Draw_GrhColor(graf.GrhIndex, (x + d) + 1, y + 1 + f * 14, Sombra())
      
                Call Draw_GrhColor(graf.GrhIndex, (x + d), y + f * 14, text_color())
                
                ' graf.grhindex = Fuentes(font_index).Caracteres(b)
                'Grh_Render graf, (x + d), y + f * 14, temp_array, False, False, False '14 es el height de esta fuente dsp lo hacemos dinamico
                d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth

            End If

        End If

        e = e + 1
    Next a

End Sub

Sub Engine_Text_Render_Exp(charindex As Integer, Texto As String, x As Integer, y As Integer, ByRef text_color() As Long, Optional ByVal font_index As Integer = 1, Optional multi_line As Boolean = False)

    Dim a, b, c, d, e, f As Integer

    Dim graf As grh

    Dim r, g As Byte

    r = (text_color(0) And 16711680) / 65536
    g = (text_color(0) And 65280) / 256
    b = text_color(0) And 255

    Dim Sombra(3) As Long 'Sombra

    Sombra(0) = D3DColorARGB(200, r / 6, g / 6, b / 6)
    Sombra(1) = Sombra(0)
    Sombra(2) = Sombra(0)
    Sombra(3) = Sombra(0)

    If (Len(Texto) = 0) Then Exit Sub

    d = 0
    e = 0
    f = 0

    For a = 1 To Len(Texto)
        b = Asc(mid(Texto, a, 1))
        graf.GrhIndex = Fuentes(font_index).Caracteres(b)

        If b = 32 Or b = 13 Then
            If e >= 20 Then 'reemplazar por lo que os plazca
                f = f + 1
                e = 0
                d = 0
            Else

                If b = 32 Then d = d + 4

            End If

        Else

            If graf.GrhIndex > 12 Then

                'mega sombra O-matica
                graf.GrhIndex = Fuentes(font_index).Caracteres(b)
      
                'Call Draw_GrhColor(graf.GrhIndex, (X + d) + 1, Y + 1 + f * 14, text_color())
                ' Call Draw_GrhColor(graf.GrhIndex, (X + d), Y + f * 14, text_color())
                
                ' graf.grhindex = Fuentes(font_index).Caracteres(b)
                'Grh_Render graf, (x + d), y + f * 14, temp_array, False, False, False '14 es el height de esta fuente dsp lo hacemos dinamico
                d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth

            End If

        End If

        e = e + 1
    Next a

End Sub

Function Engine_Text_Width(Texto As String, Optional multi As Boolean = False, Optional Fon As Byte = 1) As Integer

    Dim a, b, d, e, f As Integer

    Dim graf As grh

    Select Case Fon

        Case 1

            If multi = False Then

                For a = 1 To Len(Texto)
                    b = Asc(mid(Texto, a, 1))
                    graf.GrhIndex = Fuentes(1).Caracteres(b)

                    If graf.GrhIndex = 0 Then graf.GrhIndex = 1
                    If b <> 32 Then
                        Engine_Text_Width = Engine_Text_Width + GrhData(GrhData(graf.GrhIndex + 1).Frames(1)).pixelWidth '+ 1
                    Else
                        Engine_Text_Width = Engine_Text_Width + 4

                    End If

                Next a

            Else
                e = 0
                f = 0

                For a = 1 To Len(Texto)
                    b = Asc(mid(Texto, a, 1))
                    graf.GrhIndex = Fuentes(1).Caracteres(b)

                    If b = 32 Or b = 13 Then
                        If e >= 20 Then 'reemplazar por lo que os plazca
                            f = f + 1
                            e = 0
                            d = 0
                        Else

                            If b = 32 Then d = d + 4

                        End If

                    Else

                        If graf.GrhIndex > 12 Then
                            d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1

                            If d > Engine_Text_Width Then Engine_Text_Width = d

                        End If

                    End If

                    e = e + 1
                Next a

            End If

        Case 4

            If multi = False Then

                For a = 1 To Len(Texto)
                    b = Asc(mid(Texto, a, 1))
                    graf.GrhIndex = Fuentes(Fon).Caracteres(b)

                    If graf.GrhIndex = 0 Then graf.GrhIndex = 1
                    If b <> 20 Then
                        Engine_Text_Width = Engine_Text_Width + GrhData(GrhData(graf.GrhIndex + 1).Frames(1)).pixelWidth + 10
                    Else
                        Engine_Text_Width = Engine_Text_Width - 15

                    End If

                Next a

            Else
                e = 0
                f = 0

                For a = 1 To Len(Texto)
                    b = Asc(mid(Texto, a, 1))
                    graf.GrhIndex = Fuentes(Fon).Caracteres(b)

                    If b = 32 Or b = 13 Then
                        If e >= 20 Then 'reemplazar por lo que os plazca
                            f = f + 1
                            e = 0
                            d = 0
                        Else

                            If b = 32 Then d = d + 4

                        End If

                    Else

                        If graf.GrhIndex > 12 Then
                            d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1

                            If d > Engine_Text_Width Then Engine_Text_Width = d

                        End If

                    End If

                    e = e + 1
                Next a

            End If

    End Select

End Function

Function Engine_Text_WidthCentrado(Texto As String, Optional multi As Boolean = False, Optional Fon As Byte = 1) As Integer

    Dim a, b, d, e, f As Integer

    Dim graf As grh

    Select Case Fon

        Case 1
            '

            If multi = False Then

                For a = 1 To Len(Texto)
                    b = Asc(mid(Texto, a, 1))
                    graf.GrhIndex = Fuentes(1).Caracteres(b)

                    If graf.GrhIndex = 0 Then graf.GrhIndex = 1
                    If b <> 32 Then
                        Engine_Text_WidthCentrado = Engine_Text_WidthCentrado + GrhData(GrhData(graf.GrhIndex + 1).Frames(1)).pixelWidth '+ 1
                    Else
                        Engine_Text_WidthCentrado = Engine_Text_WidthCentrado + 4

                    End If

                Next a

            Else
                e = 0
                f = 0

                For a = 1 To Len(Texto)
                    b = Asc(mid(Texto, a, 1))
                    graf.GrhIndex = Fuentes(1).Caracteres(b)

                    If b = 32 Or b = 13 Then
                        If e >= 20 Then 'reemplazar por lo que os plazca
                            f = f + 1
                            e = 0
                            d = 0
                        Else

                            If b = 32 Then d = d + 4

                        End If

                    Else

                        If graf.GrhIndex > 12 Then
                            d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1

                            If d > Engine_Text_WidthCentrado Then Engine_Text_WidthCentrado = d

                        End If

                    End If

                    e = e + 1
                Next a

            End If

        Case 4

            If multi = False Then

                For a = 1 To Len(Texto)
                    b = Asc(mid(Texto, a, 1))
                    graf.GrhIndex = Fuentes(Fon).Caracteres(b)

                    If graf.GrhIndex = 0 Then graf.GrhIndex = 1
                    If b <> 20 Then
                        Engine_Text_WidthCentrado = Engine_Text_WidthCentrado + GrhData(GrhData(graf.GrhIndex + 1).Frames(1)).pixelWidth + 10
                    Else
                        Engine_Text_WidthCentrado = Engine_Text_WidthCentrado - 15

                    End If

                Next a

            Else
                e = 0
                f = 0

                For a = 1 To Len(Texto)
                    b = Asc(mid(Texto, a, 1))
                    graf.GrhIndex = Fuentes(Fon).Caracteres(b)

                    If b = 32 Or b = 13 Then
                        If e >= 20 Then 'reemplazar por lo que os plazca
                            f = f + 1
                            e = 0
                            d = 0
                        Else

                            If b = 32 Then d = d + 4

                        End If

                    Else

                        If graf.GrhIndex > 12 Then
                            d = d + GrhData(GrhData(graf.GrhIndex).Frames(1)).pixelWidth '+ 1

                            If d > Engine_Text_WidthCentrado Then Engine_Text_WidthCentrado = d

                        End If

                    End If

                    e = e + 1
                Next a

            End If

    End Select

End Function

Public Sub RenderConnect(ByVal tilex As Integer, ByVal tiley As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)

    On Error GoTo ErrorHandler:

    D3DDevice.BeginScene

    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 1#, 1#, 0

    Dim y                As Integer     'Keeps track of where on map we are

    Dim x                As Integer     'Keeps track of where on map we are

    Dim screenminY       As Integer  'Start Y pos on current screen

    Dim screenmaxY       As Integer  'End Y pos on current screen

    Dim screenminX       As Integer  'Start X pos on current screen

    Dim screenmaxX       As Integer  'End X pos on current screen

    Dim minY             As Integer  'Start Y pos on current map

    Dim MaxY             As Integer  'End Y pos on current map

    Dim minX             As Integer  'Start X pos on current map

    Dim MaxX             As Integer  'End X pos on current map

    Dim ScreenX          As Integer  'Keeps track of where to place tile on screen

    Dim ScreenY          As Integer  'Keeps track of where to place tile on screen

    Dim minXOffset       As Integer

    Dim minYOffset       As Integer

    Dim PixelOffsetXTemp As Integer 'For centering grhs

    Dim PixelOffsetYTemp As Integer 'For centering grhs

    Dim CurrentGrhIndex  As Integer

    Dim OffX             As Integer

    Dim Offy             As Integer

    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    minY = screenminY - TileBufferSize
    MaxY = screenmaxY + TileBufferSize
    minX = screenminX - TileBufferSize
    MaxX = screenmaxX + TileBufferSize
    
    'Make sure mins and maxs are allways in map bounds
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize

    End If
    
    If MaxY > YMaxMapSize Then MaxY = YMaxMapSize
    
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize

    End If
    
    If MaxX > XMaxMapSize Then MaxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1

    End If
    
    If screenmaxY < YMaxMapSize Then screenmaxY = screenmaxY + 1
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1

    End If
    
    If screenmaxX < XMaxMapSize Then screenmaxX = screenmaxX + 1
    
    If screenminY < 1 Then screenminY = 1
    If screenminX < 1 Then screenminX = 1
    If screenmaxY > 100 Then screenmaxY = 100
    If screenmaxX > 100 Then screenmaxX = 100
    
    screenmaxY = screenmaxY + 9
    screenmaxX = screenmaxY + 9
  
    'Draw floor layer
    For y = screenminY To screenmaxY
        For x = screenminX To screenmaxX
            'Layer 1 **********************************
            Call Draw_Grh(MapData(x, y).Graphic(1), (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, 0, 1, MapData(x, y).light_value, , x, y)
            '******************************************
            ScreenX = ScreenX + 1
        Next x

        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - x + screenminX
        ScreenY = ScreenY + 1
    Next y
    
    If HayLayer2 Then
        ScreenY = minYOffset - TileBufferSize

        For y = minY To MaxY
            ScreenX = minXOffset - TileBufferSize

            For x = minX To MaxX + 2

                With MapData(x, y)

                    '***********************************************
                    If MapData(x, y).Graphic(2).GrhIndex <> 0 Then
                        Call Draw_Grh(MapData(x, y).Graphic(2), (ScreenX * 32 + PixelOffsetX), (ScreenY * 32 + PixelOffsetY), 1, 1, MapData(x, y).light_value(), , x, y)

                    End If
              
                End With

                ScreenX = ScreenX + 1
            Next x

            ScreenY = ScreenY + 1
        Next y

    End If
    
    ScreenY = minYOffset - TileBufferSize

    For y = minY To MaxY
        ScreenX = minXOffset - TileBufferSize

        For x = minX To MaxX
            PixelOffsetXTemp = ScreenX * 32 + PixelOffsetX
            PixelOffsetYTemp = ScreenY * 32 + PixelOffsetY
            
            With MapData(x, y)
                '******************************************

                'Object Layer **********************************
                If MapData(x, y).ObjGrh.GrhIndex <> 0 Then
                    Call Draw_Grh(MapData(x, y).ObjGrh, (ScreenX * 32 + PixelOffsetX), (ScreenY * 32 + PixelOffsetY), 1, 1, MapData(x, y).light_value(), , x, y)

                End If
             
                'Layer 3 *****************************************
                If .Graphic(3).GrhIndex <> 0 Then
                    Call Draw_Grh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(x, y).light_value, False, x, y)

                End If

                '************************************************

            End With
            
            ScreenX = ScreenX + 1
        Next x

        ScreenY = ScreenY + 1
    Next y

    ScreenY = minYOffset - 5
    
    Dim cc(3)   As Long

    Dim TempGrh As grh

    'nubes negras
    TempGrh.framecounter = 1
    TempGrh.GrhIndex = 1170
    cc(0) = D3DColorARGB(180, 255, 255, 255)
    cc(1) = cc(0)
    cc(2) = cc(0)
    cc(3) = cc(0)
    ' Draw_Grh TempGrh, 494, 735, 1, 1, cc(), False
    'nubes negras

    ScreenY = minYOffset - TileBufferSize

    For y = minY To MaxY
        ScreenX = minXOffset - TileBufferSize

        For x = minX To MaxX

            With MapData(x, y)

                '***********************************************
                If .particle_group > 0 Then
                    Call Particle_Group_Render(.particle_group, ScreenX * 32 + PixelOffsetX + 15, ScreenY * 32 + PixelOffsetY + 15)

                End If
          
            End With

            ScreenX = ScreenX + 1
        Next x

        ScreenY = ScreenY + 1
    Next y
 
    'Draw blocked tiles and grid
 
    If HayLayer4 Then

        Dim rgb_list(0 To 3) As Long
    
        ScreenY = minYOffset - TileBufferSize

        For y = minY To MaxY
            ScreenX = minXOffset - TileBufferSize

            For x = minX To MaxX
        
                If MapData(x, y).Graphic(4).GrhIndex Then

                    rgb_list(0) = D3DColorARGB(255, ColorAmbiente.r, ColorAmbiente.g, ColorAmbiente.b)
                    rgb_list(1) = rgb_list(0)
                    rgb_list(2) = rgb_list(0)
                    rgb_list(3) = rgb_list(0)
                        
                    Call Draw_Grh(MapData(x, y).Graphic(4), ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, 1, 1, rgb_list(), , x, y)
          
                End If
 
                '**********************************
                ScreenX = ScreenX + 1
            Next x

            ScreenY = ScreenY + 1
        Next y
        
    End If
        
    ScreenY = minYOffset - TileBufferSize

    For y = minY To MaxY
        ScreenX = minXOffset - TileBufferSize

        For x = minX To MaxX
            PixelOffsetXTemp = ScreenX * 32 + PixelOffsetX
            PixelOffsetYTemp = ScreenY * 32 + PixelOffsetY

            With MapData(x, y)
                
                If MapData(x, y).charindex <> 0 Then
                    If charlist(MapData(x, y).charindex).active = 1 Then
                        Call Char_TextRender(MapData(x, y).charindex, PixelOffsetXTemp, PixelOffsetYTemp, x, y)

                    End If

                End If

            End With
            
            ScreenX = ScreenX + 1
        Next x

        ScreenY = ScreenY + 1
    Next y

    ScreenY = minYOffset - 5
        
    Dim DefaultColor(3) As Long

    Dim color           As Long

    intro = 1

    If intro = 1 Then

        DefaultColor(0) = D3DColorXRGB(255, 255, 255)
        DefaultColor(1) = DefaultColor(0)
        DefaultColor(2) = DefaultColor(0)
        DefaultColor(3) = DefaultColor(0)
        '    Call Renderizar_Aura("35457:&HFF8000:0:0", 400 + 15, 310, 0, 0)
        Draw_Grh BodyData(640).Walk(3), 470 + 15, 366, 1, 0, DefaultColor()
        Draw_Grh HeadData(602).Head(3), 470 + 15, 327 + 2, 1, 0, DefaultColor()
            
        Draw_Grh CascoAnimData(48).Head(3), 470 + 15, 327, 1, 0, DefaultColor()
        Draw_Grh WeaponAnimData(82).WeaponWalk(3), 470 + 15, 366, 1, 0, DefaultColor()
            
        Engine_Text_Render_LetraChica "v" & App.Major & "." & App.Minor & " Build: " & App.Revision, 870, 750, DefaultColor, 4, False

        Dim ItemName As String

        'itemname = "abcdfghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789¡!¿TEST?#$100%&/\()=-@^[]<>*+.,:; pálmas séso te píso sólo púto ý LÁL LÉ"
            
        ' itemname = "pálmas séso te píso sólo púto ý lÁ Élefante PÍSÓS PÚTO ÑOño"
        engine.Engine_Text_Render_LetraChica ItemName, 100, 730, DefaultColor, 4, False

        If ClickEnAsistente < 30 Then
            Call Particle_Group_Render(spell_particle, 500, 365)

        End If

    End If
 
    ScreenX = 250
    ScreenY = 0
    'Call Particle_Group_Render(meteo_particle, ScreenX, ScreenY)

    LastOffsetX = ParticleOffsetX
    LastOffsetY = ParticleOffsetY
    'Engine_Weather_UpdateFog
    
    TextEfectAsistente = TextEfectAsistente + (15 * timerTicksPerFrame * Sgn(-1))

    If TextEfectAsistente <= 1 Then
        TextEfectAsistente = 0

    End If

    Engine_Text_Render TextAsistente, 510 - Engine_Text_Width(TextAsistente, True, 1) / 2, 320 - Engine_Text_Height(TextAsistente, True) + TextEfectAsistente, textcolorAsistente, 1, True, , 200

    '
    ' Engine_Draw_Box 975, 5, 15, 15, D3DColorARGB(100, 70, 0, 0)
    'Engine_Text_Render UserCuenta, 490 - Engine_Text_Width(UserCuenta, False, 3) / 2, 38 - Engine_Text_Height(UserCuenta, False, 3), DefaultColor, 3, False
    ' Engine_Text_Render "X", 977, 5, DefaultColor, 1, False
    
    '   Engine_Draw_Box 955, 5, 15, 15, D3DColorARGB(100, 70, 0, 0)
    'Engine_Text_Render UserCuenta, 490 - Engine_Text_Width(UserCuenta, False, 3) / 2, 38 - Engine_Text_Height(UserCuenta, False, 3), DefaultColor, 3, False
    ' Engine_Text_Render "_", 957, 3, DefaultColor, 1, False

    'Logo viejo
    TempGrh.framecounter = 1
    TempGrh.GrhIndex = 1171

    cc(0) = D3DColorARGB(255, 255, 255, 255)
    cc(1) = cc(0)
    cc(2) = cc(0)
    cc(3) = cc(0)

    ' Draw_Grh TempGrh, 494, 200, 1, 1, cc(), False
    'Logo viejo

    'Logo viejo
    
    TempGrh.framecounter = 1
    TempGrh.GrhIndex = 1172

    cc(0) = D3DColorARGB(220, 255, 255, 255)
    cc(1) = cc(0)
    cc(2) = cc(0)
    cc(3) = cc(0)

    Draw_Grh TempGrh, 494, 275, 1, 1, cc(), False

    'Logo nuevo
    'Marco
    TempGrh.framecounter = 1
    TempGrh.GrhIndex = 1169

    cc(0) = D3DColorARGB(255, 255, 255, 255)
    cc(1) = cc(0)
    cc(2) = cc(0)
    cc(3) = cc(0)

    Draw_Grh TempGrh, 0, 0, 0, 0, cc(), False
    
    'Marco

    #If DEBUGGING = 1 Then
        ' Botones debug
        Engine_Text_Render "Debug:", 50, 300, DefaultColor
    
        ' Crear cuenta a manopla
        Engine_Draw_Box 40, 330, 155, 35, D3DColorARGB(150, 0, 0, 0)
        Engine_Text_Render "Crear cuenta en cliente", 50, 340, DefaultColor
    #End If

    'TempGrh.framecounter = 1
    'TempGrh.GrhIndex = 32016

    ' cc(0) = D3DColorARGB(255, 255, 255, 255)
    ' cc(1) = cc(0)
    ' cc(2) = cc(0)
    ' cc(3) = cc(0)

    ' Draw_Grh TempGrh, 480, 100, 1, 1, cc(), False
    D3DDevice.EndScene
    D3DDevice.Present Render_Connect_Rect, ByVal 0, frmConnect.render.hwnd, ByVal 0
    
    lFrameLimiter = (GetTickCount() And &H7FFFFFFF)
    'FramesPerSecCounter = FramesPerSecCounter + 1
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed

    Exit Sub

ErrorHandler:

    If D3DDevice.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        Call engine.Engine_ReInit
        prgRun = True
        pausa = False

    End If

End Sub

Public Sub RenderCrearPJ(ByVal tilex As Integer, ByVal tiley As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)

    On Error GoTo ErrorHandler:

    D3DDevice.BeginScene

    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 1#, 1#, 0

    Dim y                As Integer     'Keeps track of where on map we are

    Dim x                As Integer     'Keeps track of where on map we are

    Dim screenminY       As Integer  'Start Y pos on current screen

    Dim screenmaxY       As Integer  'End Y pos on current screen

    Dim screenminX       As Integer  'Start X pos on current screen

    Dim screenmaxX       As Integer  'End X pos on current screen

    Dim minY             As Integer  'Start Y pos on current map

    Dim MaxY             As Integer  'End Y pos on current map

    Dim minX             As Integer  'Start X pos on current map

    Dim MaxX             As Integer  'End X pos on current map

    Dim ScreenX          As Integer  'Keeps track of where to place tile on screen

    Dim ScreenY          As Integer  'Keeps track of where to place tile on screen

    Dim minXOffset       As Integer

    Dim minYOffset       As Integer

    Dim PixelOffsetXTemp As Integer 'For centering grhs

    Dim PixelOffsetYTemp As Integer 'For centering grhs

    Dim CurrentGrhIndex  As Integer

    Dim OffX             As Integer

    Dim Offy             As Integer

    'Figure out Ends and Starts of screen
    screenminY = tiley - HalfWindowTileHeight
    screenmaxY = tiley + HalfWindowTileHeight
    screenminX = tilex - HalfWindowTileWidth
    screenmaxX = tilex + HalfWindowTileWidth
    
    minY = screenminY - TileBufferSize
    MaxY = screenmaxY + TileBufferSize
    minX = screenminX - TileBufferSize
    MaxX = screenmaxX + TileBufferSize
    
    'Make sure mins and maxs are allways in map bounds
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize

    End If
    
    If MaxY > YMaxMapSize Then MaxY = YMaxMapSize
    
    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize

    End If
    
    If MaxX > XMaxMapSize Then MaxX = XMaxMapSize
    
    'If we can, we render around the view area to make it smoother
    If screenminY > YMinMapSize Then
        screenminY = screenminY - 1
    Else
        screenminY = 1
        ScreenY = 1

    End If
    
    If screenmaxY < YMaxMapSize Then screenmaxY = screenmaxY + 1
    
    If screenminX > XMinMapSize Then
        screenminX = screenminX - 1
    Else
        screenminX = 1
        ScreenX = 1

    End If
    
    If screenmaxX < XMaxMapSize Then screenmaxX = screenmaxX + 1
    
    If screenminY < 1 Then screenminY = 1
    If screenminX < 1 Then screenminX = 1
    If screenmaxY > 100 Then screenmaxY = 100
    If screenmaxX > 100 Then screenmaxX = 100
    screenmaxY = screenmaxY + 8
    screenmaxX = screenmaxY + 8

    'Draw floor layer
    For y = screenminY To screenmaxY
        For x = screenminX To screenmaxX
            'Layer 1 **********************************
            Call Draw_Grh(MapData(x, y).Graphic(1), (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, 0, 1, MapData(x, y).light_value, , x, y)
            '******************************************
            ScreenX = ScreenX + 1
        Next x

        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - x + screenminX
        ScreenY = ScreenY + 1
    Next y
    
    If HayLayer2 Then
        ScreenY = minYOffset - TileBufferSize

        For y = minY To MaxY
            ScreenX = minXOffset - TileBufferSize

            For x = minX To MaxX

                With MapData(x, y)

                    '***********************************************
                    If MapData(x, y).Graphic(2).GrhIndex <> 0 Then
                        Call Draw_Grh(MapData(x, y).Graphic(2), (ScreenX * 32 + PixelOffsetX), (ScreenY * 32 + PixelOffsetY), 1, 1, MapData(x, y).light_value(), , x, y)

                    End If
              
                End With

                ScreenX = ScreenX + 1
            Next x

            ScreenY = ScreenY + 1
        Next y

    End If
    
    ScreenY = minYOffset - TileBufferSize

    For y = minY To MaxY
        ScreenX = minXOffset - TileBufferSize

        For x = minX To MaxX
            PixelOffsetXTemp = ScreenX * 32 + PixelOffsetX
            PixelOffsetYTemp = ScreenY * 32 + PixelOffsetY
            
            With MapData(x, y)
                '******************************************

                'Object Layer **********************************
                If MapData(x, y).ObjGrh.GrhIndex <> 0 Then
                    Call Draw_Grh(MapData(x, y).ObjGrh, (ScreenX * 32 + PixelOffsetX), (ScreenY * 32 + PixelOffsetY), 1, 1, MapData(x, y).light_value(), , x, y)

                End If
             
                'Layer 3 *****************************************
                If .Graphic(3).GrhIndex <> 0 Then
                    Call Draw_Grh(.Graphic(3), PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(x, y).light_value, False, x, y)

                End If

                '************************************************

            End With
            
            ScreenX = ScreenX + 1
        Next x

        ScreenY = ScreenY + 1
    Next y

    ScreenY = minYOffset - 5

    ScreenY = minYOffset - TileBufferSize

    For y = minY To MaxY
        ScreenX = minXOffset - TileBufferSize

        For x = minX To MaxX

            With MapData(x, y)

                '***********************************************
                If .particle_group > 0 Then
                    Call Particle_Group_Render(.particle_group, ScreenX * 32 + PixelOffsetX + 15, ScreenY * 32 + PixelOffsetY + 15)

                End If
          
            End With

            ScreenX = ScreenX + 1
        Next x

        ScreenY = ScreenY + 1
    Next y
 
    'Draw blocked tiles and grid
 
    If HayLayer4 Then

        Dim rgb_list(0 To 3) As Long
    
        ScreenY = minYOffset - TileBufferSize

        For y = minY To MaxY
            ScreenX = minXOffset - TileBufferSize

            For x = minX To MaxX
        
                If MapData(x, y).Graphic(4).GrhIndex Then

                    rgb_list(0) = D3DColorARGB(255, ColorAmbiente.r, ColorAmbiente.g, ColorAmbiente.b)
                    rgb_list(1) = rgb_list(0)
                    rgb_list(2) = rgb_list(0)
                    rgb_list(3) = rgb_list(0)
                        
                    Call Draw_Grh(MapData(x, y).Graphic(4), ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, 1, 1, rgb_list(), , x, y)
          
                End If
 
                '**********************************
                ScreenX = ScreenX + 1
            Next x

            ScreenY = ScreenY + 1
        Next y
        
    End If

    Engine_Weather_UpdateFog

    engine.RenderUICrearPJ

    Dim cc(3)   As Long

    Dim TempGrh As grh

    'Logo viejo
    TempGrh.framecounter = 1
    TempGrh.GrhIndex = 1171

    cc(0) = D3DColorARGB(255, 255, 255, 255)
    cc(1) = cc(0)
    cc(2) = cc(0)
    cc(3) = cc(0)

    Draw_Grh TempGrh, 494, 190, 1, 1, cc(), False
    'Logo viejo
    
    'Marco
    TempGrh.framecounter = 1
    TempGrh.GrhIndex = 1169

    cc(0) = D3DColorARGB(255, 255, 255, 255)
    cc(1) = cc(0)
    cc(2) = cc(0)
    cc(3) = cc(0)

    Draw_Grh TempGrh, 0, 0, 0, 0, cc(), False
    
    D3DDevice.EndScene
    ' D3DDevice.Present Render_CrearPj_Rect, ByVal 0, frmCrearPersonaje.render.hwnd, ByVal 0
    D3DDevice.Present Render_Connect_Rect, ByVal 0, frmConnect.render.hwnd, ByVal 0

    lFrameLimiter = (GetTickCount() And &H7FFFFFFF)
    FramesPerSecCounter = FramesPerSecCounter + 1
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed

    'engine.RenderPjsCuenta

ErrorHandler:

    If D3DDevice.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        Call engine.Engine_ReInit
        prgRun = True
        pausa = False

    End If

End Sub

Public Sub rendercuenta(ByVal tilex As Integer, ByVal tiley As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)

    On Error GoTo ErrorHandler:

    D3DDevice.BeginScene

    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 1#, 1#, 0

    lFrameLimiter = (GetTickCount() And &H7FFFFFFF)
    FramesPerSecCounter = FramesPerSecCounter + 1
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed

    engine.RenderPjsCuenta
    
    Call Particle_Group_Render(ParticleLluviaDorada, 400, 0)

    D3DDevice.EndScene
    D3DDevice.Present Render_Connect_Rect, ByVal 0, frmConnect.render.hwnd, ByVal 0
    Exit Sub
ErrorHandler:

    If D3DDevice.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        Call engine.Engine_ReInit
        prgRun = True
        pausa = False

    End If

End Sub

Public Sub RenderUICrearPJ()

    Dim TempGrh         As grh
    
    Dim DefaultColor(3) As Long
    
    TempGrh.framecounter = 1
    TempGrh.GrhIndex = 727
    DefaultColor(0) = D3DColorXRGB(255, 255, 255)
    DefaultColor(1) = DefaultColor(0)
    DefaultColor(2) = DefaultColor(0)
    DefaultColor(3) = DefaultColor(0)
    
    Draw_Grh TempGrh, 475, 545, 1, 1, DefaultColor(), False

    DefaultColor(0) = D3DColorXRGB(200, 200, 200)
    DefaultColor(1) = DefaultColor(0)
    DefaultColor(2) = DefaultColor(0)
    DefaultColor(3) = DefaultColor(0)
    
    'Engine_Text_Render "Nombre del personaje", 230 + -Engine_Text_Width("Nombre del personaje", False) / 2, 110 + 40 - Engine_Text_Height("Nombre del personaje", False), DefaultColor, 3, True
    'Engine_Text_Render "Creacion de personajes", 210, 120, DefaultColor, 3, False
    
    'Basico
    'Engine_Draw_Box 217, 183, 550, 386, D3DColorARGB(20, 219, 116, 3)
    
    'Engine_Draw_Box 250, 190, 490, 356, D3DColorARGB(50, 128, 128, 128)
    'Engine_Draw_Box 250, 190, 490, 356, D3DColorARGB(100, 0, 0, 0)
    
    'Engine_Draw_Box 220, 186, 550, 380, D3DColorARGB(80, 20, 27, 3)
    'Engine_Text_Render UserCuenta, 490 - Engine_Text_Width(UserCuenta, False, 3) / 2, 38 - Engine_Text_Height(UserCuenta, False, 3), DefaultColor, 3, False
    Engine_Text_Render "Creacion de Personaje", 280, 125, DefaultColor, 5, False

    'Engine_Draw_Box 400, 215, 180, 250, D3DColorARGB(200, 100, 100, 100)
    
    DefaultColor(0) = D3DColorXRGB(255, 255, 255)
    DefaultColor(1) = DefaultColor(0)
    DefaultColor(2) = DefaultColor(0)
    DefaultColor(3) = DefaultColor(0)
    
    Engine_Text_Render_LetraChica "Nombre ", 470, 198, DefaultColor, 6, False
    Engine_Text_Render_LetraChica "Clase ", 477, 240, DefaultColor, 6, False
    
    '
    
    Engine_Draw_Box 450, 255, 95, 21, D3DColorARGB(100, 1, 1, 1)
    
    Engine_Text_Render "<", 435, 258, DefaultColor, 1, False
        
    Engine_Text_Render ">", 548, 258, DefaultColor, 1, False
    'Engine_Text_Render ">", 403, 412, DefaultColor, 1, True
    
    DefaultColor(0) = D3DColorXRGB(200, 200, 200)
    DefaultColor(1) = DefaultColor(0)
    DefaultColor(2) = DefaultColor(0)
    DefaultColor(3) = DefaultColor(0)
    
    Engine_Text_Render frmCrearPersonaje.lstProfesion.List(frmCrearPersonaje.lstProfesion.ListIndex), 498 - Engine_Text_Width(frmCrearPersonaje.lstProfesion.List(frmCrearPersonaje.lstProfesion.ListIndex), True, 1) / 2, 258, DefaultColor, 1, True
    
    DefaultColor(0) = D3DColorXRGB(255, 255, 255)
    DefaultColor(1) = DefaultColor(0)
    DefaultColor(2) = DefaultColor(0)
    DefaultColor(3) = DefaultColor(0)
    
    Engine_Text_Render_LetraChica "Raza ", 481, 285, DefaultColor, 6, False
    Engine_Draw_Box 450, 302, 95, 21, D3DColorARGB(100, 1, 1, 1)
    
    DefaultColor(0) = D3DColorXRGB(200, 200, 200)
    DefaultColor(1) = DefaultColor(0)
    DefaultColor(2) = DefaultColor(0)
    DefaultColor(3) = DefaultColor(0)

    'Engine_Text_Render "Humano", 470 - Engine_Text_Height("Humano", False), 304, DefaultColor, 1, False
    Engine_Text_Render frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.ListIndex), 495 - Engine_Text_Width(frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.ListIndex), True, 1) / 2, 305, DefaultColor, 1, True
    
    Engine_Text_Render "<", 435, 305, DefaultColor, 1, False
    Engine_Text_Render ">", 548, 305, DefaultColor, 1, False
    
    DefaultColor(0) = D3DColorXRGB(255, 255, 255)
    DefaultColor(1) = DefaultColor(0)
    DefaultColor(2) = DefaultColor(0)
    DefaultColor(3) = DefaultColor(0)
    
    Engine_Text_Render_LetraChica "Genero ", 475, 330, DefaultColor, 6, False
    Engine_Draw_Box 450, 346, 95, 21, D3DColorARGB(100, 1, 1, 1)
        
    DefaultColor(0) = D3DColorXRGB(200, 200, 200)
    DefaultColor(1) = DefaultColor(0)
    DefaultColor(2) = DefaultColor(0)
    DefaultColor(3) = DefaultColor(0)
    
    Engine_Text_Render frmCrearPersonaje.lstGenero.List(frmCrearPersonaje.lstGenero.ListIndex), 495 - Engine_Text_Width(frmCrearPersonaje.lstGenero.List(frmCrearPersonaje.lstGenero.ListIndex), True, 1) / 2, 349, DefaultColor, 1, True
    
    Engine_Text_Render "<", 435, 350, DefaultColor, 1, False
    Engine_Text_Render ">", 548, 350, DefaultColor, 1, False
    
    DefaultColor(0) = D3DColorXRGB(200, 200, 50)
    DefaultColor(1) = DefaultColor(0)
    DefaultColor(2) = DefaultColor(0)
    DefaultColor(3) = DefaultColor(0)
    
    'Engine_Text_Render RazaRecomendada, 489 - Engine_Text_Width(RazaRecomendada, False, 1) / 2, 278, DefaultColor, 1, False
    
    If Len(RazaRecomendada) > 0 Then
        Engine_Text_Render "Raza sugerida:", 570, 290, DefaultColor, 4, False
        Engine_Text_Render RazaRecomendada, 570, 300, DefaultColor, 4, False

    End If
    
    '     DefaultColor(0) = D3DColorXRGB(255, 50, 50)
    '  DefaultColor(1) = DefaultColor(0)
    '  DefaultColor(2) = DefaultColor(0)
    '  DefaultColor(3) = DefaultColor(0)
    
    '' Engine_Text_Render "¡Atención! ", 240, 250, DefaultColor, 1, False
    '     DefaultColor(0) = D3DColorXRGB(255, 255, 255)
    '  DefaultColor(1) = DefaultColor(0)
    ' DefaultColor(2) = DefaultColor(0)
    ' DefaultColor(3) = DefaultColor(0)
    ' Engine_Text_Render "Se cuidadoso al momento de distribuir tus atributos. De estos dependen aspectos basicos como la vida o maná de tu personaje. ", 190, 270, DefaultColor, 1, True
    
    DefaultColor(0) = D3DColorXRGB(255, 255, 255)
    DefaultColor(1) = DefaultColor(0)
    DefaultColor(2) = DefaultColor(0)
    DefaultColor(3) = DefaultColor(0)
        
    Dim Offy As Long
     
    Offy = 2

    Dim OffX As Long
     
    OffX = 350
    
    'Atributos
    Engine_Text_Render_LetraChica "Atributos ", 240 + OffX, 385 + Offy, DefaultColor, 6, True
    Engine_Draw_Box 175 + OffX, 405 + Offy, 185, 120, D3DColorARGB(80, 0, 0, 0)
    '  Engine_Draw_Box 610, 405, 220, 180, D3DColorARGB(120, 100, 100, 100)
    
    Engine_Text_Render_LetraChica "Fuerza ", 185 + OffX, 410 + Offy, DefaultColor, 1, True
    ' Engine_Text_Render "<", 260, 410, DefaultColor, 1, True
    ' Engine_Text_Render ">", 310, 410, DefaultColor, 1, True
    Engine_Draw_Box 280 + OffX, 409 + Offy, 20, 20, D3DColorARGB(100, 1, 1, 1)
    Engine_Text_Render_LetraChica frmCrearPersonaje.lbFuerza.Caption, 282 + OffX, 413 + Offy, DefaultColor, 1, True ' Atributo fuerza
    'Engine_Text_Render "+", 335, 410, DefaultColor, 1, True
    Engine_Draw_Box 317 + OffX, 409 + Offy, 25, 20, D3DColorARGB(100, 1, 1, 1)
    Engine_Text_Render_LetraChica frmCrearPersonaje.modfuerza.Caption, 320 + OffX, 413 + Offy, DefaultColor, 1, True ' Bonificacion fuerza
    
    Engine_Text_Render "Agilidad ", 185 + OffX, 440 + Offy, DefaultColor, 1, True
    ' Engine_Text_Render "<", 260, 440, DefaultColor, 1, True
    ' Engine_Text_Render ">", 310, 440, DefaultColor, 1, True
    Engine_Draw_Box 280 + OffX, 440 + Offy, 20, 20, D3DColorARGB(100, 1, 1, 1)
    Engine_Text_Render frmCrearPersonaje.lbAgilidad.Caption, 282 + OffX, 443 + Offy, DefaultColor, 1, True ' Atributo Agilidad
    ' Engine_Text_Render "+", 335, 440, DefaultColor, 1, True
    Engine_Draw_Box 317 + OffX, 440 + Offy, 25, 20, D3DColorARGB(100, 1, 1, 1)
    Engine_Text_Render frmCrearPersonaje.modAgilidad.Caption, 320 + OffX, 443 + Offy, DefaultColor, 1, True ' Bonificacion Agilidad
    
    Engine_Text_Render "Inteligencia ", 185 + OffX, 470 + Offy, DefaultColor, 1, True
    'Engine_Text_Render "<", 260, 470, DefaultColor, 1, True
    'Engine_Text_Render ">", 310, 470, DefaultColor, 1, True
    Engine_Draw_Box 280 + OffX, 470 + Offy, 20, 20, D3DColorARGB(100, 1, 1, 1)
    Engine_Text_Render frmCrearPersonaje.lbInteligencia.Caption, 282 + OffX, 473 + Offy, DefaultColor, 1, True ' Atributo Inteligencia
    'Engine_Text_Render "+", 335, 470, DefaultColor, 1, True
    Engine_Draw_Box 317 + OffX, 470 + Offy, 25, 20, D3DColorARGB(100, 1, 1, 1)
    Engine_Text_Render frmCrearPersonaje.modInteligencia.Caption, 320 + OffX, 473 + Offy, DefaultColor, 1, True ' Bonificacion Inteligencia
    
    Engine_Text_Render "Constitución ", 185 + OffX, 500 + Offy, DefaultColor, , True
    'Engine_Text_Render "<", 260, 500, DefaultColor, 1, True
    ' Engine_Text_Render ">", 310, 500, DefaultColor, 1, True
    Engine_Draw_Box 280 + OffX, 500 + Offy, 20, 20, D3DColorARGB(100, 1, 1, 1)
    Engine_Text_Render frmCrearPersonaje.lbConstitucion.Caption, 283 + OffX, 503 + Offy, DefaultColor, 1, True ' Atributo Carisma
    '
    ' Engine_Text_Render "+", 335, 500, DefaultColor, 1, True
    Engine_Draw_Box 317 + OffX, 500 + Offy, 25, 20, D3DColorARGB(100, 1, 1, 1)
    Engine_Text_Render frmCrearPersonaje.modConstitucion.Caption, 320 + OffX, 503 + Offy, DefaultColor, 1, True ' Bonificacion Carisma
      
    '
    'Engine_Draw_Box 290, 528, 20, 20, D3DColorARGB(120, 1, 150, 150)
    'Engine_Text_Render "Puntos disponibles", 175, 530, DefaultColor, 1, True '
    'Engine_Text_Render frmCrearPersonaje.lbLagaRulzz.Caption, 291, 530, DefaultColor, 1, True '
    'Cabeza
    'Engine_Draw_Box 425, 415, 140, 100, D3DColorARGB(120, 100, 100, 100)

    ' Engine_Text_Render "Selecciona el rostro que más te agrade.", 662, 260, DefaultColor, 1, True

    OffX = -345
    Offy = -100
     
    Engine_Draw_Box 280, 407, 185, 120, D3DColorARGB(80, 0, 0, 0)
     
    Engine_Text_Render_LetraChica "Aspecto", 345, 385, DefaultColor, 6, False
    
    ' Engine_Draw_Box 345, 502, 12, 12, D3DColorARGB(120, 100, 0, 0)
    
    'Engine_Text_Render_LetraChica "Equipado", 360, 502, DefaultColor, 4, False
     
    ' CPHeading = 3
     
    DefaultColor(0) = D3DColorXRGB(255, 255, 255)
    DefaultColor(1) = DefaultColor(0)
    DefaultColor(2) = DefaultColor(0)
    DefaultColor(3) = DefaultColor(0)

    If CPHead <> 0 And CPArma <> 0 Then
         
        Engine_Text_Render_LetraChica "Cabeza", 350, 410, DefaultColor, 1, False
        Engine_Text_Render "<", 335, 412, DefaultColor, 1, True
        Engine_Text_Render ">", 403, 412, DefaultColor, 1, True
        
        Engine_Text_Render ">", 423, 428, DefaultColor, 3, True
        Engine_Text_Render "<", 293, 428, DefaultColor, 3, True
    
        'If CPEquipado Then
        '    Engine_Draw_Box 347, 512, 12, 12, D3DColorARGB(100, 255, 1, 1)
        '    Engine_Text_Render_LetraChica "Equipado", 360, 512, DefaultColor, 4, False
        '    Engine_Text_Render_LetraChica "x", 348, 512, DefaultColor, 6, False
        'Else
        '    Engine_Draw_Box 347, 512, 12, 12, D3DColorARGB(100, 255, 1, 1)
        '    Engine_Text_Render_LetraChica "Equipado", 360, 512, DefaultColor, 4, False
        'End If
    
        Dim Raza As Byte

        If frmCrearPersonaje.lstRaza.ListIndex < 0 Then
            frmCrearPersonaje.lstRaza.ListIndex = 0

        End If

        Raza = frmCrearPersonaje.lstRaza.ListIndex

        Dim enanooff As Byte

        If Raza = 0 Or Raza = 1 Or Raza = 2 Or Raza = 5 Then
            enanooff = 0
    
        Else
            enanooff = 10

        End If
    
        If CPEquipado Then
            Call Renderizar_Aura(CPAura, 686 + 15 + OffX, 360 - Offy + enanooff, 0, 0)

        End If
            
        If enanooff > 0 Then
            Draw_Grh BodyData(CPBodyE).Walk(CPHeading), 685 + 15 + OffX, 366 - Offy, 1, 0, DefaultColor()
        Else
            Draw_Grh BodyData(CPBody).Walk(CPHeading), 685 + 15 + OffX, 366 - Offy, 1, 0, DefaultColor()

        End If
            
        Draw_Grh HeadData(CPHead).Head(CPHeading), 685 + 15 + OffX, 366 - Offy + BodyData(CPBody).HeadOffset.y + enanooff, 1, 0, DefaultColor()
            
        'If CPEquipado Then
        'Draw_Grh CascoAnimData(CPGorro).Head(CPHeading), 700 + OffX, 366 - Offy + BodyData(CPBody).HeadOffset.y + enanooff, 1, 0, DefaultColor()
        'Draw_Grh WeaponAnimData(CPArma).WeaponWalk(CPHeading), 685 + 15 + OffX, 365 - Offy + enanooff, 1, 0, DefaultColor()
        'Call Renderizar_Aura(CPAura, 686 + 15 + offx, 360 - offy, 0, 0)
        'End If
            
        DefaultColor(0) = D3DColorXRGB(0, 128, 190)
        DefaultColor(1) = DefaultColor(0)
        DefaultColor(2) = DefaultColor(0)
        DefaultColor(3) = DefaultColor(0)
        Engine_Text_Render CPName, 372 - Engine_Text_Width(CPName, True) / 2, 495, DefaultColor, 1, True
    Else
        Engine_Text_Render "X", 355, 428, DefaultColor, 3, True

    End If
    
    'DefaultColor(0) = D3DColorXRGB(255, 255, 255)
    'DefaultColor(1) = DefaultColor(0)
    'DefaultColor(2) = DefaultColor(0)
    'DefaultColor(3) = DefaultColor(0)

    'Boton Atras
    'Engine_Draw_Box 147, 628, 100, 40, D3DColorARGB(80, 0, 0, 0)
    'Engine_Text_Render "< Volver", 170, 640, DefaultColor, 1, True
    
    'Boton Crear
    'If StopCreandoCuenta Then
    '    Engine_Draw_Box 730, 630, 100, 40, D3DColorARGB(120, 100, 180, 100)
    '    Engine_Text_Render "Creando...", 750, 640, DefaultColor, 1, True
    'Else
    '    Engine_Draw_Box 730, 630, 100, 40, D3DColorARGB(80, 0, 0, 0)
    '    Engine_Text_Render "Crear PJ >", 750, 640, DefaultColor, 1, True
    'End If
       
    'Engine_Text_Render "DADO", 670, 390, DefaultColor()
    Draw_GrhIndex 1123, 665, 385

End Sub

Public Sub RenderPjsCuenta()

    ' Renderiza el menu para seleccionar las clases
        
    Dim i               As Long

    Dim x               As Integer

    Dim y               As Integer

    Dim notY            As Integer

    Dim DefaultColor(3) As Long

    Dim color           As Long

    Dim Texto           As String

    Texto = CuentaEmail

    'Render fondo
    Draw_GrhIndex 1170, 0, 0
    
    Dim temp_array(3) As Long 'Si le queres dar color a la letra pasa este parametro dsp xD

    DefaultColor(0) = D3DColorXRGB(255, 255, 255)
    DefaultColor(1) = DefaultColor(0)
    DefaultColor(2) = DefaultColor(0)
    DefaultColor(3) = DefaultColor(0)

    Dim sumax As Long

    sumax = 84
            
    For i = 1 To 10
            
        If (i > 5) Then
            x = ((i * 132) - (5 * 132))
            y = 440
        Else
            x = (i * 132)
            y = 283

        End If

        x = x + sumax

        temp_array(0) = Pjs(i).LetraColor
        temp_array(1) = Pjs(i).LetraColor
        temp_array(2) = Pjs(i).LetraColor
        temp_array(3) = Pjs(i).LetraColor
        
        'Offset de la cabeza / enanos.
        ' If (Pjs(i).Clase <> eClass.Warrior) Then
        ' notY = 5
        ' Else
        Rem   notY = -5
        ' End If
        
        'Si tiene cuerpo dibuja
        If (Pjs(i).Body <> 0) Then
        
            If PJSeleccionado = i Then
                Call Particle_Group_Render(Select_part, x + 32, y + 5)

            End If

            If (Pjs(i).Body <> 0) Then
                  
                'Else
                'Engine_Draw_Box X - 40, Y - 40, 145, 150, D3DColorARGB(20, 28, 18, 9)
                Draw_Grh BodyData(Pjs(i).Body).Walk(3), x + 15, y + 10, 1, 1, DefaultColor()

            End If

            If (Pjs(i).Head <> 0) Then
                'If Not nohead Then
                Draw_Grh HeadData(Pjs(i).Head).Head(3), x + 15, y - notY + BodyData(Pjs(i).Body).HeadOffset.y + 10, 1, 0, DefaultColor()

                ' End If
            End If
            
            If (Pjs(i).Casco <> 0) Then
                'If Not nohead Then
                Draw_Grh CascoAnimData(Pjs(i).Casco).Head(3), x + 15, y - notY + BodyData(Pjs(i).Body).HeadOffset.y + 10, 1, 0, DefaultColor()

                ' End If
            End If
            
            If (Pjs(i).Escudo <> 0) Then
                'If Not nohead Then
                Draw_Grh ShieldAnimData(Pjs(i).Escudo).ShieldWalk(3), x + 14, y - notY + 10, 1, 0, DefaultColor()

                ' End If
            End If
                        
            If (Pjs(i).Arma <> 0) Then
                'If Not nohead Then
                Draw_Grh WeaponAnimData(Pjs(i).Arma).WeaponWalk(3), x + 14, y - notY + 10, 1, 0, DefaultColor()

                ' End If
            End If
            
            Dim colorCorazon(0 To 4) As Long

            Dim b                    As Long

            Dim g                    As Long

            Dim r                    As Long

            colorCorazon(0) = temp_array(1)
            colorCorazon(1) = temp_array(1)
            colorCorazon(2) = temp_array(1)
            colorCorazon(3) = temp_array(1)
            
            'Convert LONG to RGB:
            ' b = temp_array(1) \ 65536
            ' g = (temp_array(1) - b * 65536) \ 256
            'r = temp_array(1) - b * 65536 - g * 256
                
            '' r = (temp_array(1) And 16711680) / 65536
            ' g = (temp_array(1) And 65280) / 256
            ' b = temp_array(1) And 255
                
            colorCorazon(0) = D3DColorXRGB(r, g, b)
            colorCorazon(1) = colorCorazon(0)
            colorCorazon(2) = colorCorazon(0)
            colorCorazon(3) = colorCorazon(0)
        
            If CuentaDonador = 1 Then
                Grh_Render Estrella, x + 17 + 6 + Engine_Text_Width(Pjs(i).nombre, 1) / 2, y + 19, temp_array(), True, True, False

            End If

            Engine_Text_Render Pjs(i).nombre, x + 30 - Engine_Text_Width(Pjs(i).nombre, True) / 2, y + 56 - Engine_Text_Height(Pjs(i).nombre, True), temp_array(), 1, True
            
            If PJSeleccionado = i Then
            
                Dim Offy As Byte

                Offy = 0
            
                Engine_Text_Render Pjs(i).nombre, 511 - Engine_Text_Width(Pjs(i).nombre, True) / 2, 565 - Engine_Text_Height(Pjs(i).nombre, True), temp_array(), 1, True
                
                If Pjs(i).ClanName <> "<>" Then
                    Engine_Text_Render Pjs(i).ClanName, 511 - Engine_Text_Width(Pjs(i).ClanName, True) / 2, 565 + 15 - Engine_Text_Height(Pjs(i).ClanName, True), temp_array(), 1, True
                    Offy = 15
                Else
                
                    Offy = 0

                End If

                Engine_Text_Render "Clase: " & ListaClases(Pjs(i).Clase), 511 - Engine_Text_Width("Clase:" & ListaClases(Pjs(i).Clase), True) / 2, Offy + 585 - Engine_Text_Height("Clase:" & ListaClases(Pjs(i).Clase), True), DefaultColor, 1, True
                
                Engine_Text_Render "Nivel: " & Pjs(i).nivel, 511 - Engine_Text_Width("Nivel:" & Pjs(i).nivel, True) / 2, Offy + 600 - Engine_Text_Height("Nivel:" & Pjs(i).nivel, True), DefaultColor, 1, True
                Engine_Text_Render CStr(Pjs(i).NameMapa), 5111 - Engine_Text_Width(CStr(Pjs(i).NameMapa), True) / 2, Offy + 615 - Engine_Text_Height(CStr(Pjs(i).NameMapa), True), DefaultColor, 1, True

            End If
            
        End If

    Next i

End Sub

Sub EfectoEnPantalla(ByVal color As Long, ByVal time As Long)
    frmmain.Efecto.Interval = time
    frmmain.Efecto.Enabled = True
    EfectoEnproceso = True
    Call engine.Map_Base_Light_Set(color)

End Sub

Public Sub SetBarFx(ByVal charindex As Integer, ByVal BarTime As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 12/03/04
    'Sets an FX to the character.
    '***************************************************

    With charlist(charindex)
        .BarTime = BarTime

    End With

End Sub

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

Public Function Light_Color_Value_Get(ByVal light_index As Long, ByRef color_value As Long) As Boolean

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 2/28/2003
    '
    '*****************************************************************
    'Make sure it's a legal index
    If Light_Check(light_index) Then
        color_value = light_list(light_index).color
        Light_Color_Value_Get = True

    End If

End Function

Public Function Light_Create(ByVal map_x As Integer, ByVal map_y As Integer, Optional ByVal color_value As Long = &HFFFFFFFF, Optional ByVal range As Byte = 1, Optional ByVal id As Long) As Long

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

            Case NORTH
                addy = -1
        
            Case EAST
                addx = 1
        
            Case south
                addy = 1
            
            Case WEST
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

Private Sub Light_Make(ByVal light_index As Long, ByVal map_x As Integer, ByVal map_y As Integer, ByVal rgb_value As Long, ByVal range As Long, Optional ByVal id As Long)

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
    
    If rgb_value = 0 Then
        rgb_value = map_base_light

    End If

    Dim r, g, b As Byte

    r = rgb_value And 255
    g = (rgb_value \ 256) And 255
    b = (rgb_value \ 65536) And 255
    light_list(light_index).map_x = map_x
    light_list(light_index).map_y = map_y
    
    Dim rgb_list As Long

    rgb_list = D3DColorARGB(255, b, g, r)

    light_list(light_index).color = rgb_list
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

    Dim x     As Integer

    Dim y     As Integer

    Dim ia    As Single

    Dim i     As Integer

    Dim color As Long
    
    'Set up light borders
    min_x = light_list(light_index).map_x - light_list(light_index).range
    min_y = light_list(light_index).map_y - light_list(light_index).range
    max_x = light_list(light_index).map_x + light_list(light_index).range
    max_y = light_list(light_index).map_y + light_list(light_index).range
    
    'Set color
    color = light_list(light_index).color
    
    MapData(light_list(light_index).map_x, light_list(light_index).map_y).light_value(0) = color
    MapData(light_list(light_index).map_x, light_list(light_index).map_y).light_value(1) = color
    MapData(light_list(light_index).map_x, light_list(light_index).map_y).light_value(2) = color
    MapData(light_list(light_index).map_x, light_list(light_index).map_y).light_value(3) = color
                
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
        MapData(min_x, min_y).light_value(2) = color

    End If

    'NW
    If InMapBounds(max_x, min_y) Then
        MapData(max_x, min_y).light_value(0) = color

    End If

    'SW
    If InMapBounds(max_x, max_y) Then
        MapData(max_x, max_y).light_value(1) = color

    End If

    'SE
    If InMapBounds(min_x, max_y) Then
        MapData(min_x, max_y).light_value(3) = color

    End If
    
    'Arrange borders
    'Upper border
    For x = min_x + 1 To max_x - 1

        If InMapBounds(x, min_y) Then
            MapData(x, min_y).light_value(0) = color
            MapData(x, min_y).light_value(2) = color

        End If

    Next x
    
    'Lower border
    For x = min_x + 1 To max_x - 1

        If InMapBounds(x, max_y) Then
            MapData(x, max_y).light_value(1) = color
            MapData(x, max_y).light_value(3) = color

        End If

    Next x
    
    'Left border
    For y = min_y + 1 To max_y - 1

        If InMapBounds(min_x, y) Then
            MapData(min_x, y).light_value(2) = color
            MapData(min_x, y).light_value(3) = color

        End If

    Next y
    
    'Right border
    For y = min_y + 1 To max_y - 1

        If InMapBounds(max_x, y) Then
            MapData(max_x, y).light_value(0) = color
            MapData(max_x, y).light_value(1) = color

        End If

    Next y
    
    'Set the inner part of the light
    For x = min_x + 1 To max_x - 1
        For y = min_y + 1 To max_y - 1

            If InMapBounds(x, y) Then
                MapData(x, y).light_value(0) = color
                MapData(x, y).light_value(1) = color
                MapData(x, y).light_value(2) = color
                MapData(x, y).light_value(3) = color

            End If

        Next y
    Next x

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

    Dim x     As Integer

    Dim y     As Integer
    
    'Set up light borders
    min_x = light_list(light_index).map_x - light_list(light_index).range
    min_y = light_list(light_index).map_y - light_list(light_index).range
    max_x = light_list(light_index).map_x + light_list(light_index).range
    max_y = light_list(light_index).map_y + light_list(light_index).range
    
    'Arrange corners
    'NE
    If InMapBounds(min_x, min_y) Then
        MapData(min_x, min_y).light_value(2) = 0

    End If

    'NW
    If InMapBounds(max_x, min_y) Then
        MapData(max_x, min_y).light_value(0) = 0

    End If

    'SW
    If InMapBounds(max_x, max_y) Then
        MapData(max_x, max_y).light_value(1) = 0

    End If

    'SE
    If InMapBounds(min_x, max_y) Then
        MapData(min_x, max_y).light_value(3) = 0

    End If
    
    'Arrange borders
    'Upper border
    For x = min_x + 1 To max_x - 1

        If InMapBounds(x, min_y) Then
            MapData(x, min_y).light_value(0) = 0
            MapData(x, min_y).light_value(2) = 0

        End If

    Next x
    
    'Lower border
    For x = min_x + 1 To max_x - 1

        If InMapBounds(x, max_y) Then
            MapData(x, max_y).light_value(1) = 0
            MapData(x, max_y).light_value(3) = 0

        End If

    Next x
    
    'Left border
    For y = min_y + 1 To max_y - 1

        If InMapBounds(min_x, y) Then
            MapData(min_x, y).light_value(2) = 0
            MapData(min_x, y).light_value(3) = 0

        End If

    Next y
    
    'Right border
    For y = min_y + 1 To max_y - 1

        If InMapBounds(max_x, y) Then
            MapData(max_x, y).light_value(0) = 0
            MapData(max_x, y).light_value(1) = 0

        End If

    Next y
    
    'Set the inner part of the light
    For x = min_x + 1 To max_x - 1
        For y = min_y + 1 To max_y - 1

            If InMapBounds(x, y) Then
                MapData(x, y).light_value(0) = 0
                MapData(x, y).light_value(1) = 0
                MapData(x, y).light_value(2) = 0
                MapData(x, y).light_value(3) = 0

            End If

        Next y
    Next x

End Sub

Public Function Engine_Get_2_Points_Angle(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Double
    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 18/10/2012
    '**************************************************************

    Engine_Get_2_Points_Angle = Engine_Get_X_Y_Angle((x2 - x1), (y2 - y1))
   
End Function

Public Function Engine_Get_X_Y_Angle(ByVal x As Double, ByVal y As Double) As Double
    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 18/10/2012
    '**************************************************************

    Dim dblres As Double
 
    dblres = 0
   
    If (y <> 0) Then
        dblres = Engine_Convert_Radians_To_Degrees(Atn(x / y))

        If (x <= 0 And y < 0) Then
            dblres = dblres + 180
        ElseIf (x > 0 And y < 0) Then
            dblres = dblres + 180
        ElseIf (x < 0 And y > 0) Then
            dblres = dblres + 360

        End If

    Else

        If (x > 0) Then
            dblres = 90
        ElseIf (x < 0) Then
            dblres = 270

        End If

    End If
   
    Engine_Get_X_Y_Angle = dblres
   
End Function

Public Function Engine_Convert_Radians_To_Degrees(ByVal s_radians As Double) As Integer
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 8/25/2004
    'Converts a radian to degrees
    '**************************************************************

    Engine_Convert_Radians_To_Degrees = (s_radians * 180) / 3.14159265358979
 
End Function

' programado por maTih.-
 
Public Sub Initialize()
    '
    ' @ Inicializa el array de efectos.
     
    ReDim Effect(1 To 255) As Effect_Type
    
    ' Inicializo inventarios
    Set frmmain.Inventario = New clsGrapchicalInventory
    Set frmComerciar.InvComUsu = New clsGrapchicalInventory
    Set frmComerciar.InvComNpc = New clsGrapchicalInventory
    Set frmBancoObj.InvBankUsu = New clsGrapchicalInventory
    Set frmBancoObj.InvBoveda = New clsGrapchicalInventory

    Call frmmain.Inventario.Initialize(frmmain.picInv, MAX_INVENTORY_SLOTS, , , 0, 0, 3, 3, True, 9)
    Call frmComerciar.InvComUsu.Initialize(frmComerciar.interface, MAX_INVENTORY_SLOTS, 210, 0, 252, 0, 3, 3, True) 'Nuev
    Call frmComerciar.InvComNpc.Initialize(frmComerciar.interface, MAX_INVENTORY_SLOTS, 210, , 1, 0, 3, 3) 'Nuevo
   
    Call frmBancoObj.InvBankUsu.Initialize(frmBancoObj.interface, MAX_INVENTORY_SLOTS, 210, 0, 252, 0, 3, 3, True)
    Call frmBancoObj.InvBoveda.Initialize(frmBancoObj.interface, MAX_BANCOINVENTORY_SLOTS, 210, 0, 0, 0, 3, 3)
 
End Sub

Public Sub Terminate_Index(ByVal effect_Index As Integer)
 
    '
    ' @ Destruye un indice del array
 
    Dim clear_Index As Effect_Type
 
    'Si es un slot válido
    If (effect_Index <> 0) And (effect_Index <= UBound(Effect())) Then
        Effect(effect_Index) = clear_Index

    End If
 
End Sub
 
Public Function Effect_Begin(ByVal Fx_Index As Integer, ByVal Bind_Speed As Single, ByVal x As Single, ByVal y As Single, Optional ByVal explosion_FX_Index As Integer = -1, Optional ByVal explosion_FX_Loops As Integer = -1, Optional ByVal receptor As Integer = 1, Optional ByVal Emisor As Integer = 1, Optional ByVal wav As Integer = 1, Optional ByVal fX As Integer = -1) As Integer
 
    '
    ' @ Inicia un nuevo efecto y devuelve el index.
 
    Effect_Begin = GetFreeIndex()
 
    ' Debug.Print "fx =" & fX
    'Si hay efecto
    If (Effect_Begin <> 0) Then
    
        With Effect(Effect_Begin)
            .Now_X = CInt(x) - 16
            .Now_Y = CInt(y) - 20
         
            .Fx_Index = Fx_Index
         
            .ViajeSpeed = Bind_Speed
            .ViajeChar = Emisor
            .DestinoChar = receptor
         
            .wav = wav
         
            'Explosión?
            If (explosion_FX_Index <> 0) And (fX = 0) Then
                .End_Effect = explosion_FX_Index
                .End_Loops = explosion_FX_Loops
                .FxEnd_Effect = 0
                
            End If
         
            If (fX = 1) Then
                .End_Effect = 0
                .FxEnd_Effect = explosion_FX_Index
                .End_Loops = explosion_FX_Loops
            
            End If
         
            .Slot_Used = True
         
        End With
    
    End If
 
End Function

Public Function Effect_BeginXY(ByVal Fx_Index As Integer, ByVal Bind_Speed As Single, ByVal x As Single, ByVal y As Single, ByVal DestinoX As Byte, ByVal Destinoy As Byte, Optional ByVal explosion_FX_Index As Integer = -1, Optional ByVal explosion_FX_Loops As Integer = -1, Optional ByVal Emisor As Integer = 1, Optional ByVal wav As Integer = 1, Optional ByVal fX As Integer = 0) As Integer
    '
    ' @ Inicia un nuevo efecto y devuelve el index.
 
    ' Debug.Print "fx =" & fX
    Effect_BeginXY = GetFreeIndex()
 
    'Si hay efecto
    If (Effect_BeginXY <> 0) Then
    
        With Effect(Effect_BeginXY)
            .Now_X = CInt(x) - 16
            .Now_Y = CInt(y) - 20
         
            .Fx_Index = Fx_Index
         
            .ViajeSpeed = Bind_Speed
            .ViajeChar = Emisor
            .DestinoChar = 0
            .DestX = DestinoX
            .DesyY = Destinoy
         
            .wav = wav
         
            'Explosión?
            If (explosion_FX_Index <> 0) And (fX = 0) Then
                .End_Effect = explosion_FX_Index
                .End_Loops = explosion_FX_Loops
                .FxEnd_Effect = 0

            End If
         
            If (fX = 1) Then
                .End_Effect = 0
                .FxEnd_Effect = explosion_FX_Index
                .End_Loops = explosion_FX_Loops

            End If
         
            .Slot_Used = True
         
        End With
    
    End If
 
End Function
 
Public Sub Effect_Render_All()
 
    '
    ' @ Dibuja todos los efectos
 
    Dim i As Long

    For i = 1 To UBound(Effect())

        With Effect(i)
         
            If .Slot_Used Then
    
                Effect_Render_Slot CInt(i)

            End If
         
        End With

    Next i
 
End Sub
 
Public Sub Effect_Render_Slot(ByVal effect_Index As Integer)
 
    '
    ' @ Renderiza un efecto.

    Dim colornpcs(3) As Long

    colornpcs(0) = D3DColorXRGB(255, 255, 255)
    colornpcs(1) = colornpcs(0)
    colornpcs(2) = colornpcs(0)
    colornpcs(3) = colornpcs(0)
    
    With Effect(effect_Index)
 
        Dim target_Angle As Single
     
        .Now_Moved = (GetTickCount() And &H7FFFFFFF)
     
        'Controla el intervalo de vuelo
        If (.Last_Move + 10) < .Now_Moved Then
            .Last_Move = (GetTickCount() And &H7FFFFFFF)
        
            'Si tiene char de destino.
            If (.DestinoChar <> 0) Then
     
                'Actualiza la pos de destino.
                '.Viaje_X = charlist(.ViajeChar).NowPosX
                '.Viaje_Y = charlist(.ViajeChar).NowPosY
            
                .Viaje_X = Get_Pixelx_Of_Char(.DestinoChar) - 0
                .Viaje_Y = Get_PixelY_Of_Char(.DestinoChar) - 32

            Else
                .Viaje_X = Get_Pixelx_Of_XY(.DestX) - 0
                .Viaje_Y = Get_PixelY_Of_XY(.DesyY) - 32

            End If
        
        End If
      
        'Actualiza el ángulo.
        target_Angle = Engine_GetAngle(.Now_X, .Now_Y, CInt(.Viaje_X), CInt(.Viaje_Y))
    
        'Actualiza la posición del efecto.
        .Now_X = (.Now_X + Sin(target_Angle * DegreeToRadian) * .ViajeSpeed)
        .Now_Y = (.Now_Y - Cos(target_Angle * DegreeToRadian) * .ViajeSpeed)

        'Si hay posición dibuja.
        If (.Now_X <> 0) And (.Now_Y <> 0) Then
            ' Call DDrawTransGrhtoSurface(.FX_Grh, .Now_X, .Now_Y, 1, 1)

            Call Particle_Group_Render(spell_particle, .Now_X, .Now_Y)
        
            'Check si terminó.
            ' If (.FX_Grh.Started = 0) Then .Fx_Index = 0: .Slot_Used = False
        
            If Abs(CInt(.Viaje_X) - CInt(.Now_X)) < 5 Then
                .Now_X = .Viaje_X

            End If

            If Abs(CInt(.Viaje_Y) - CInt(.Now_Y)) < 5 Then
        
                .Now_Y = .Viaje_Y

            End If
        
            If (.Now_X = .Viaje_X) And (.Now_Y = .Viaje_Y) Then
       
                'Inicializa la explosión : p
                If (.End_Effect <> 0) And .DestinoChar <> 0 Then
                    If .DestinoChar <> 0 Then
                        Call General_Char_Particle_Create(.End_Effect, .DestinoChar, .End_Loops)
                        Call Sound.Sound_Play(.wav, , Sound.Calculate_Volume(charlist(.DestinoChar).Pos.x, charlist(.DestinoChar).Pos.y), Sound.Calculate_Pan(charlist(.DestinoChar).Pos.x, charlist(.DestinoChar).Pos.y))
                        .Slot_Used = False
                        Exit Sub

                    End If

                End If
            
                If (.End_Effect <> 0) And .DestinoChar = 0 Then
                    MapData(.DestX, .DesyY).particle_group = 0
                    General_Particle_Create .End_Effect, .DestX, .DesyY, .End_Loops
                    Call Sound.Sound_Play(.wav, , Sound.Calculate_Volume(.DestX, .DesyY), Sound.Calculate_Pan(.DestX, .DesyY))
                    .Slot_Used = False
                    Exit Sub

                End If
            
                If (.FxEnd_Effect > 0) And .DestinoChar <> 0 Then
                    Call Sound.Sound_Play(.wav, , Sound.Calculate_Volume(charlist(.DestinoChar).Pos.x, charlist(.DestinoChar).Pos.y), Sound.Calculate_Pan(charlist(.DestinoChar).Pos.x, charlist(.DestinoChar).Pos.y))
                    Call SetCharacterFx(.DestinoChar, .FxEnd_Effect, .End_Loops)
                    .Slot_Used = False
                    Exit Sub

                End If
            
                If (.FxEnd_Effect > 0) And (.DestinoChar = 0) Then
                    Call Sound.Sound_Play(.wav, , Sound.Calculate_Volume(.DestX, .DesyY), Sound.Calculate_Pan(.DestX, .DesyY))
              
                    Call engine.SetMapFx(.DestX, .DesyY, .FxEnd_Effect, 0)
                    .Slot_Used = False
                    Exit Sub

                End If
          
            End If
        
        End If

    End With
 
End Sub
 
Public Function Engine_GetAngle(ByVal CenterX As Integer, ByVal CenterY As Integer, ByVal TargetX As Integer, ByVal TargetY As Integer) As Single

    '************************************************************
    'Gets the angle between two points in a 2d plane
    'More info: [url=http://www.vbgore.com/GameClient.TileEn]http://www.vbgore.com/GameClient.TileEn[/url] ... e_GetAngle" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
    '************************************************************
    Dim SideA As Single

    Dim SideC As Single
 
    On Error GoTo ErrOut
 
    'Check for horizontal lines (90 or 270 degrees)
    If CenterY = TargetY Then
 
        'Check for going right (90 degrees)
        If CenterX < TargetX Then
            Engine_GetAngle = 90
 
            'Check for going left (270 degrees)
        Else
            Engine_GetAngle = 270

        End If
 
        'Exit the function
        Exit Function
 
    End If
 
    'Check for horizontal lines (360 or 180 degrees)
    If CenterX = TargetX Then
 
        'Check for going up (360 degrees)
        If CenterY > TargetY Then
            Engine_GetAngle = 360
 
            'Check for going down (180 degrees)
        Else
            Engine_GetAngle = 180

        End If
 
        'Exit the function
        Exit Function
 
    End If
 
    'Calculate Side C
    SideC = Sqr(Abs(TargetX - CenterX) ^ 2 + Abs(TargetY - CenterY) ^ 2)
 
    'Side B = CenterY
 
    'Calculate Side A
    SideA = Sqr(Abs(TargetX - CenterX) ^ 2 + TargetY ^ 2)
 
    'Calculate the angle
    Engine_GetAngle = (SideA ^ 2 - CenterY ^ 2 - SideC ^ 2) / (CenterY * SideC * -2)
    Engine_GetAngle = (Atn(-Engine_GetAngle / Sqr(-Engine_GetAngle * Engine_GetAngle + 1)) + 1.5708) * 57.29583
 
    'If the angle is >180, subtract from 360
    If TargetX < CenterX Then Engine_GetAngle = 360 - Engine_GetAngle
 
    'Exit function
 
    Exit Function
 
    'Check for error
ErrOut:
 
    'Return a 0 saying there was an error
    Engine_GetAngle = 0
 
    Exit Function
 
End Function
 
Public Function GetFreeIndex() As Integer
 
    '
    ' @ Devuelve un índice para un nuevo FX.
 
    Dim i As Long
 
    For i = 1 To UBound(Effect())

        'No está usado.
        If Not Effect(i).Slot_Used Then
            GetFreeIndex = CInt(i)
            Exit Function

        End If

    Next i
 
    GetFreeIndex = NO_INDEX
 
End Function

Public Function Get_PixelY_Of_Char(ByVal char_index As Integer) As Integer

    '*****************************************************************
    'Author: Pablo Mercavides
    '*****************************************************************
    'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        Get_PixelY_Of_Char = (charlist(char_index).Pos.y - 2 - UserPos.y) * 32 + frmmain.renderer.ScaleWidth / 2
        Get_PixelY_Of_Char = Get_PixelY_Of_Char - 16

    End If

End Function

Public Function Get_Pixelx_Of_Char(ByVal char_index As Integer) As Integer

    '*****************************************************************
    'Author: Pablo Mercavides
    '*****************************************************************
    'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        Get_Pixelx_Of_Char = (charlist(char_index).Pos.x - UserPos.x) * 32 + frmmain.renderer.ScaleWidth / 2
        Get_Pixelx_Of_Char = Get_Pixelx_Of_Char

    End If

End Function

Public Function Get_Pixelx_Of_XY(ByVal x As Byte) As Integer
    '*****************************************************************
    'Author: Pablo Mercavides
    '*****************************************************************
    'Make sure it's a legal char_index
    Get_Pixelx_Of_XY = (x - UserPos.x) * 32 + frmmain.renderer.ScaleWidth / 2
    Get_Pixelx_Of_XY = Get_Pixelx_Of_XY

End Function

Public Function Get_PixelY_Of_XY(ByVal y As Byte) As Integer
    '*****************************************************************
    'Author: Pablo Mercavides
    '*****************************************************************
    'Make sure it's a legal char_index
    Get_PixelY_Of_XY = (y - 2 - UserPos.y) * 32 + frmmain.renderer.ScaleWidth / 2
    Get_PixelY_Of_XY = Get_PixelY_Of_XY - 16

End Function

Public Function Letter_Set(ByVal grh_index As Long, ByVal text_string As String) As Boolean
    '*****************************************************************
    'Author: Augusto José Rando
    '*****************************************************************
    letter_text = text_string
    letter_grh.GrhIndex = grh_index
    Letter_Set = True
    map_letter_fadestatus = 1

End Function

Public Function Map_Letter_Fade_Set(ByVal grh_index As Long, Optional ByVal after_grh As Long = -1) As Boolean

    '*****************************************************************
    'Author: Augusto José Rando
    '*****************************************************************
    If grh_index <= 0 Or grh_index = map_letter_grh.GrhIndex Then Exit Function
        
    If after_grh = -1 Then
        map_letter_grh.GrhIndex = grh_index
        map_letter_fadestatus = 1
        map_letter_a = 0
        map_letter_grh_next = 0
    Else
        map_letter_grh.GrhIndex = after_grh
        map_letter_fadestatus = 1
        map_letter_a = 0
        map_letter_grh_next = grh_index

    End If
    
    Map_Letter_Fade_Set = True

End Function

Public Function Map_Letter_UnSet() As Boolean
    '*****************************************************************
    'Author: Augusto José Rando
    '*****************************************************************
    map_letter_grh.GrhIndex = 0
    map_letter_fadestatus = 0
    map_letter_a = 0
    map_letter_grh_next = 0
    Map_Letter_UnSet = True

End Function

Public Function Letter_UnSet() As Boolean
    '*****************************************************************
    'Author: Augusto José Rando
    '*****************************************************************
    letter_text = vbNullString
    letter_grh.GrhIndex = 0
    Letter_UnSet = True

End Function

Sub RenderConsola()

    Dim i As Byte
 
    If OffSetConsola > 0 Then OffSetConsola = OffSetConsola - 1
    If OffSetConsola = 0 Then UltimaLineavisible = True
 
    For i = 1 To MaxLineas - 1
 
        Text_Render font_list(1), Con(i).T, ComienzoY + (i * 15) + OffSetConsola - 20, 10, frmmain.renderer.Width, frmmain.renderer.Height, ARGB(Con(i).r, Con(i).g, Con(i).b, i * (255 / MaxLineas)), DT_TOP Or DT_LEFT, False
        
    Next i
 
    If UltimaLineavisible = True Then Text_Render font_list(1), Con(i).T, ComienzoY + (MaxLineas * 15) + OffSetConsola - 20, 10, frmmain.renderer.Width, frmmain.renderer.Height, ARGB(Con(MaxLineas).r, Con(MaxLineas).g, Con(i).b, 255), DT_TOP Or DT_LEFT, False
 
End Sub
