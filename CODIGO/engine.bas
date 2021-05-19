Attribute VB_Name = "engine"
Option Explicit


Public FrameNum               As Long

'Depentientes del motor grafico
Public Dialogos                 As clsDialogs
Public LucesRedondas            As clsLucesRedondas
Public LucesCuadradas           As clsLucesCuadradas
Private Estrella                As grh
Private Marco                   As grh
Private BarraMana               As grh
Private BarraVida               As grh
Private BarraGris               As grh

''
' Maximum number of dialogs that can exist.
Public Const MAX_DIALOGS     As Byte = 100

''
' Maximum length of any dialog line without having to split it.
Public Const MAX_LENGTH      As Byte = 18

''
' Number of milliseconds to add to the lifetime per dialog character
Public Const MS_PER_CHAR     As Byte = 100

''
' Number of extra milliseconds to add to the lifetime of a new dialog
Public Const MS_ADD_EXTRA    As Integer = 5000
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
    Color       As Long
    renderable  As Boolean
    Sube As Byte

End Type

Public scroll_dialog_pixels_per_frame As Single

''
' Array if dialogs, sorted by the charIndex.
Public dialogs(MAX_DIALOGS - 1)   As dialog

''
' The number of dialogs being used at the moment.
Public dialogCount                As Byte



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

Public Const ScreenWidth   As Long = 538
Public Const ScreenHeight  As Long = 376

Public temp_rgb(3)         As RGBA

Public bRunning            As Boolean

Dim Texture      As Direct3DTexture8
Dim TransTexture As Direct3DTexture8

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public fps                     As Long
Private FramesPerSecCounter    As Long
Private lFrameTimer            As Long
Public FrameTime               As Long

Public FadeInAlpha             As Integer

Private ScrollPixelsPerFrameX  As Single
Private ScrollPixelsPerFrameY  As Single

Private TileBufferPixelOffsetX As Integer
Private TileBufferPixelOffsetY As Integer
Private TimeLast As Long

Private Const GrhFogata        As Integer = 1521

' Colores estaticos

Public FLASH(3)        As RGBA
Public COLOR_EMPTY              As RGBA
Public COLOR_WHITE(3)           As RGBA
Public r As Byte
Public G As Byte
Public B As Byte
Public textcolorAsistente(3)    As RGBA

'Sets a Grh animation to loop indefinitely.

#Const HARDCODED = False 'True ' == MÁS FPS ^^

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
    If timer_freq = 0 Then Call QueryPerformanceFrequency(timer_freq)
    
    'Get current time
    Call QueryPerformanceCounter(Start_Time)
    
    'Calculate elapsed time
    GetElapsedTime = (Start_Time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)

    
    Exit Function

GetElapsedTime_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.GetElapsedTime", Erl)
    Resume Next
    
End Function

Private Function Init_DirectDevice(ByVal ModoAceleracion As CONST_D3DCREATEFLAGS) As Boolean
On Error GoTo ErrorHandler:
    
    ' Retrieve the information about your current display adapter.
    Call DirectD3D.GetAdapterDisplayMode(D3DADAPTER_DEFAULT, DispMode)
    
    With D3DWindow
    
        .Windowed = True

        If VSyncActivado Then
            .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
        Else
            .SwapEffect = D3DSWAPEFFECT_DISCARD
        End If
        
        .BackBufferFormat = DispMode.format
        
        'set color depth
        .BackBufferWidth = 1024
        .BackBufferHeight = 768
        
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D24S8
        
        .hDeviceWindow = frmMain.renderer.hwnd
        
    End With
    
    If Not DirectDevice Is Nothing Then Set DirectDevice = Nothing
    
    Set DirectDevice = DirectD3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, D3DWindow.hDeviceWindow, ModoAceleracion, D3DWindow)
    
    Init_DirectDevice = True
    
    Exit Function
    
ErrorHandler:
    
    Set DirectDevice = Nothing
    
    Init_DirectDevice = False

End Function

Private Sub Engine_InitExtras()
    
    On Error GoTo Engine_InitExtras_Err
    

    Call InitGrh(Estrella, 35764)
    Call InitGrh(Marco, 839)
    Call InitGrh(BarraMana, 840)
    Call InitGrh(BarraVida, 841)
    Call InitGrh(BarraGris, 842)
    
    With Render_Connect_Rect
        .Top = 0
        .Left = 0
        .Right = frmConnect.render.ScaleWidth
        .Bottom = frmConnect.render.ScaleHeight
    End With
    
    With Render_Main_Rect
        .Top = 0
        .Left = 0
        .Right = frmMain.renderer.ScaleWidth
        .Bottom = frmMain.renderer.ScaleHeight
    End With
    
    Call Engine_InitColors
    
    ' Sistemas dependientes de el motor grafico
    Set LucesRedondas = New clsLucesRedondas
    Set LucesCuadradas = New clsLucesCuadradas
    Set Dialogos = New clsDialogs
    
    Call IniciarMeteorologia
    
    ' Fuentes graficas.
    Call Engine_Font_Initialize
    'Call Font_Create("Tahoma", 8, True, 0)
    'Call Font_Create("Verdana", 8, False, 0)
    'Call Font_Create("Verdana", 11, True, False)
        
    ' Inicializar textura compuesta
    Call InitComposedTexture
    
    
    Exit Sub

Engine_InitExtras_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Engine_InitExtras", Erl)
    Resume Next
    
End Sub

Public Sub Engine_InitColors()
    
    On Error GoTo Engine_InitColors_Err
    

    ' Colores comunes
    Call Long_2_RGBAList(COLOR_WHITE, -1)

    Call RGBAList(textcolorAsistente, 0, 200, 0)

    
    Exit Sub

Engine_InitColors_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Engine_InitColors", Erl)
    Resume Next
    
End Sub

Public Sub Engine_Init()

On Error GoTo errhandler:
    
    ' Initialize all DirectX objects.
    Set DirectX = New DirectX8
    Set DirectD3D = DirectX.Direct3DCreate()
    Set DirectD3D8 = New D3DX8

    Select Case ModoAceleracion
    
        Case "Auto"
            If Not Init_DirectDevice(D3DCREATE_HARDWARE_VERTEXPROCESSING) Then
                If Not Init_DirectDevice(D3DCREATE_MIXED_VERTEXPROCESSING) Then
                    If Not Init_DirectDevice(D3DCREATE_SOFTWARE_VERTEXPROCESSING) Then
                        
                        GoTo errhandler
                        
                    End If
                End If
            End If
                
        Case "Hardware"
            If Init_DirectDevice(D3DCREATE_HARDWARE_VERTEXPROCESSING) = False Then GoTo errhandler
            Debug.Print "Modo de Renderizado: HARDWARE"
            
        Case "Mixed"
            If Init_DirectDevice(D3DCREATE_MIXED_VERTEXPROCESSING) = False Then GoTo errhandler
            Debug.Print "Modo de Renderizado: MIXED"
        
        Case Else
            If Init_DirectDevice(D3DCREATE_SOFTWARE_VERTEXPROCESSING) = False Then GoTo errhandler
            Debug.Print "Modo de Renderizado: SOFTWARE"
    
    End Select
    
    'Seteamos la matriz de proyeccion.
    Call D3DXMatrixOrthoOffCenterLH(Projection, 0, 1024, 768, 0, -1#, 1#)
    Call D3DXMatrixIdentity(IdentityMatrix)
    Call DirectDevice.SetTransform(D3DTS_PROJECTION, Projection)
    Call DirectDevice.SetTransform(D3DTS_VIEW, IdentityMatrix)

    'Set the render states
    With DirectDevice
        Call .SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_MODULATE)
        
        Call .SetVertexShader(D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)
        
        Call .SetRenderState(D3DRS_LIGHTING, False)
        Call .SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
        Call .SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
        Call .SetRenderState(D3DRS_ALPHABLENDENABLE, True)
        Call .SetRenderState(D3DRS_ZENABLE, True)
        Call .SetRenderState(D3DRS_FILLMODE, D3DFILL_SOLID)
        Call .SetRenderState(D3DRS_CULLMODE, D3DCULL_NONE)
        Call .SetRenderState(D3DRS_ALPHAFUNC, D3DCMP_GREATER)
        Call .SetRenderState(D3DRS_ALPHAREF, 0)
    End With
    
    ' Carga de texturas
    Set SurfaceDB = New clsTexManager
    Call SurfaceDB.Init(DirectD3D8, DirectDevice, General_Get_Free_Ram_Bytes)
    
    'Sprite batching.
    Set SpriteBatch = New clsBatch
    Call SpriteBatch.Initialize(2000)

    ' Configuracion del motor
    engineBaseSpeed = 0.018
    
    'Set FPS value to 60 for startup
    fps = 60
    FramesPerSecCounter = 60
    scroll_dialog_pixels_per_frame = 4
    
    ScrollPixelsPerFrameX = 8.5
    ScrollPixelsPerFrameY = 8.5
    
    Call Engine_InitExtras

    bRunning = True
    
    Exit Sub
    
errhandler:
    
    Call MsgBox("Ha ocurrido un error al iniciar el motor grafico." & vbNewLine & _
                "Asegúrate de tener los drivers gráficos actualizados y la librería DX8VB.dll registrada correctamente.", vbCritical, "Argentum20")
    
    Debug.Print "Error Number Returned: " & Err.Number

    End

End Sub

Public Sub Engine_BeginScene(Optional ByVal Color As Long = 0)
    
    On Error GoTo Engine_BeginScene_Err
    
    
    Call DirectDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, Color, 1, 0)
    
    Call DirectDevice.BeginScene
    
    Call SpriteBatch.Begin

    
    Exit Sub

Engine_BeginScene_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Engine_BeginScene", Erl)
    Resume Next
    
End Sub

Public Sub Engine_EndScene(ByRef DestRect As RECT, Optional ByVal hwnd As Long = 0)

On Error GoTo ErrorHandler:
    
    Call SpriteBatch.Flush
    
    Call DirectDevice.EndScene
        
    Call DirectDevice.Present(DestRect, ByVal 0, hwnd, ByVal 0)
    
    Exit Sub
    
ErrorHandler:

    If DirectDevice.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        
        Call Engine_Init
        
        prgRun = True
        pausa = False
        QueRender = 0

    End If
        
End Sub

Public Sub Engine_Deinit()
    
    On Error GoTo Engine_Deinit_Err
    
    
    Erase MapData
    Erase charlist
    
    Set DirectDevice = Nothing
    Set DirectD3D = Nothing
    Set DirectX = Nothing
    Set SpriteBatch = Nothing

    
    Exit Sub

Engine_Deinit_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Engine_Deinit", Erl)
    Resume Next
    
End Sub

Public Sub Engine_ActFPS()
    
    On Error GoTo Engine_ActFPS_Err
    

    If FrameTime - lFrameTimer >= 1000 Then
        fps = FramesPerSecCounter
        FramesPerSecCounter = 0
        lFrameTimer = FrameTime
    End If

    
    Exit Sub

Engine_ActFPS_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Engine_ActFPS", Erl)
    Resume Next
    
End Sub

Public Sub Draw_GrhIndex(ByVal grh_index As Long, ByVal x As Integer, ByVal y As Integer)
    
    On Error GoTo Draw_GrhIndex_Err
    

    If grh_index <= 0 Then Exit Sub

    Call Batch_Textured_Box(x, y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, GrhData(grh_index).sX, GrhData(grh_index).sY, GrhData(grh_index).FileNum, COLOR_WHITE)

    
    Exit Sub

Draw_GrhIndex_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Draw_GrhIndex", Erl)
    Resume Next
    
End Sub

Public Sub Draw_GrhColor(ByVal grh_index As Long, ByVal x As Integer, ByVal y As Integer, ByRef text_color() As RGBA)
    
    On Error GoTo Draw_GrhColor_Err
    

    If grh_index <= 0 Then Exit Sub
    
    'Device_Box_Textured_Render grh_index, x, y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, text_color, GrhData(grh_index).sX, GrhData(grh_index).sY
    Call Batch_Textured_Box(x, y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, GrhData(grh_index).sX, GrhData(grh_index).sY, GrhData(grh_index).FileNum, text_color)

    
    Exit Sub

Draw_GrhColor_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Draw_GrhColor", Erl)
    Resume Next
    
End Sub

Public Sub Draw_GrhFont(ByVal grh_index As Long, ByVal x As Integer, ByVal y As Integer, ByRef text_color() As RGBA)
    
    On Error GoTo Draw_GrhFont_Err
    

    If grh_index <= 0 Then Exit Sub

    Call Batch_Textured_Box_Advance(x, y, GrhData(grh_index).pixelWidth + 1, GrhData(grh_index).pixelHeight + 1, GrhData(grh_index).sX, GrhData(grh_index).sY, GrhData(grh_index).FileNum, GrhData(grh_index).pixelWidth + 1, GrhData(grh_index).pixelHeight + 1, text_color)

    
    Exit Sub

Draw_GrhFont_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Draw_GrhFont", Erl)
    Resume Next
    
End Sub

Public Sub Draw_GrhIndexColor(ByVal grh_index As Long, ByVal x As Integer, ByVal y As Integer)
    
    On Error GoTo Draw_GrhIndexColor_Err
    

    If grh_index <= 0 Then Exit Sub

    Call Batch_Textured_Box(x, y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, GrhData(grh_index).sX, GrhData(grh_index).sY, GrhData(grh_index).FileNum, COLOR_WHITE, True)

    
    Exit Sub

Draw_GrhIndexColor_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Draw_GrhIndexColor", Erl)
    Resume Next
    
End Sub

Public Sub Draw_Grh(ByRef grh As grh, ByVal x As Integer, ByVal y As Integer, ByVal center As Byte, ByVal animate As Byte, ByRef rgb_list() As RGBA, Optional ByVal Alpha As Boolean = False, Optional ByVal map_x As Byte = 1, Optional ByVal map_y As Byte = 1, Optional ByVal Angle As Single)
    
    On Error GoTo Draw_Grh_Err

    If grh.GrhIndex = 0 Or grh.GrhIndex > MaxGrh Then Exit Sub
    
    Dim CurrentFrame As Integer
    CurrentFrame = 1

    If animate Then
        If grh.Started > 0 Then
            Dim ElapsedFrames As Long
            ElapsedFrames = Fix(0.5 * (FrameTime - grh.Started) / grh.speed)

            If grh.Loops = INFINITE_LOOPS Or ElapsedFrames < GrhData(grh.GrhIndex).NumFrames * (grh.Loops + 1) Then
                CurrentFrame = ElapsedFrames Mod GrhData(grh.GrhIndex).NumFrames + 1

            Else
                grh.Started = 0
            End If

        End If

    End If
    
    Dim CurrentGrhIndex As Long
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(grh.GrhIndex).Frames(CurrentFrame)

    'Center Grh over X,Y pos
    If center Then
        If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
            x = x - Int(GrhData(CurrentGrhIndex).TileWidth * (TilePixelWidth \ 2)) + TilePixelWidth \ 2
        End If

        If GrhData(grh.GrhIndex).TileHeight <> 1 Then
            y = y - Int(GrhData(CurrentGrhIndex).TileHeight * TilePixelHeight) + TilePixelHeight
        End If
    End If

    With GrhData(CurrentGrhIndex)

        If .Tx2 = 0 And .FileNum > 0 Then
            Dim Texture As Direct3DTexture8

            Dim TextureWidth As Long, TextureHeight As Long
            Set Texture = SurfaceDB.GetTexture(.FileNum, TextureWidth, TextureHeight)
        
            .Tx1 = .sX / TextureWidth
            .Tx2 = (.sX + .pixelWidth) / TextureWidth
            .Ty1 = .sY / TextureHeight
            .Ty2 = (.sY + .pixelHeight) / TextureHeight
        End If
        
        Call Batch_Textured_Box_Pre(x, y, .pixelWidth, .pixelHeight, .Tx1, .Ty1, .Tx2, .Ty2, .FileNum, rgb_list, Alpha, Angle)
    
    End With
    
    Exit Sub

Draw_Grh_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Draw_Grh", Erl)
    Resume Next
    
End Sub

Public Sub Draw_Grh_Breathing(ByRef grh As grh, ByVal x As Integer, ByVal y As Integer, ByVal center As Byte, ByVal animate As Byte, ByRef rgb_list() As RGBA, ByVal ease As Single, Optional ByVal Alpha As Boolean = False)
    
    On Error GoTo Draw_Grh_Breathing_Err

    If grh.GrhIndex = 0 Or grh.GrhIndex > MaxGrh Then Exit Sub
    
    Dim CurrentFrame As Integer
    CurrentFrame = 1

    If animate Then
        If grh.Started > 0 Then
            Dim ElapsedFrames As Long
            ElapsedFrames = Fix(0.5 * (FrameTime - grh.Started) / grh.speed)

            If grh.Loops = INFINITE_LOOPS Or ElapsedFrames < GrhData(grh.GrhIndex).NumFrames * (grh.Loops + 1) Then
                CurrentFrame = ElapsedFrames Mod GrhData(grh.GrhIndex).NumFrames + 1

            Else
                grh.Started = 0
            End If

        End If

    End If
    
    Dim CurrentGrhIndex As Long
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(grh.GrhIndex).Frames(CurrentFrame)

    'Center Grh over X,Y pos
    If center Then
        If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
            x = x - Int(GrhData(CurrentGrhIndex).TileWidth * (TilePixelWidth \ 2)) + TilePixelWidth \ 2
        End If

        If GrhData(grh.GrhIndex).TileHeight <> 1 Then
            y = y - Int(GrhData(CurrentGrhIndex).TileHeight * TilePixelHeight) + TilePixelHeight
        End If
    End If

    With GrhData(CurrentGrhIndex)

        Dim Texture As Direct3DTexture8

        Dim TextureWidth As Long, TextureHeight As Long
        Set Texture = SurfaceDB.GetTexture(.FileNum, TextureWidth, TextureHeight)

        Call SpriteBatch.SetTexture(Texture)

        Call SpriteBatch.SetAlpha(Alpha)
        
        If .Tx2 = 0 And .FileNum > 0 Then
            .Tx1 = .sX / TextureWidth
            .Tx2 = (.sX + .pixelWidth) / TextureWidth
            .Ty1 = .sY / TextureHeight
            .Ty2 = (.sY + .pixelHeight) / TextureHeight
        End If

        Call SpriteBatch.DrawBreathing(x, y, .pixelWidth, .pixelHeight, ease, rgb_list, .Tx1, .Ty1, .Tx2, .Ty2)

    End With
    
    Exit Sub

Draw_Grh_Breathing_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Draw_Grh_Breathing", Erl)
    Resume Next
    
End Sub

Sub Draw_GrhFX(ByRef grh As grh, ByVal x As Integer, ByVal y As Integer, ByVal center As Byte, ByVal animate As Byte, ByRef rgb_list() As RGBA, Optional ByVal Alpha As Boolean, Optional ByVal map_x As Byte = 1, Optional ByVal map_y As Byte = 1, Optional ByVal Angle As Single, Optional ByVal charindex As Integer)
    
    On Error GoTo Draw_GrhFX_Err
    

    

    If grh.GrhIndex = 0 Or grh.GrhIndex > MaxGrh Then Exit Sub
    
    Dim CurrentFrame As Integer
    CurrentFrame = 1

    If animate Then
        If grh.Started > 0 Then
            Dim ElapsedFrames As Long
            ElapsedFrames = Fix((FrameTime - grh.Started) / grh.speed)
            
            If grh.AnimacionContador > 0 Then
                grh.AnimacionContador = grh.AnimacionContador - ElapsedFrames
            End If

            If grh.Loops = INFINITE_LOOPS Or ElapsedFrames < GrhData(grh.GrhIndex).NumFrames * (grh.Loops + 1) Then
                CurrentFrame = ElapsedFrames Mod GrhData(grh.GrhIndex).NumFrames + 1

            Else
                grh.Started = 0
            End If

        End If

    End If
    
    Dim CurrentGrhIndex As Long
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(grh.GrhIndex).Frames(CurrentFrame)

    If grh.AnimacionContador < grh.CantAnim * 0.1 Then
            
        grh.Alpha = grh.Alpha - 1

        Call RGBAList(rgb_list, 255, 255, 255, grh.Alpha)

        If grh.Alpha = 0 And charindex > 0 Then
            charlist(charindex).fX.Started = 0
            Exit Sub

        End If

    End If
    
    If grh.AnimacionContador > grh.CantAnim * 0.6 Then
        If grh.Alpha < 220 Then
            grh.Alpha = grh.Alpha + 1
        End If
        
        Call RGBAList(rgb_list, 255, 255, 255, grh.Alpha)

    End If

    'Center Grh over X,Y pos
    If center Then
        If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
            x = x - Int(GrhData(CurrentGrhIndex).TileWidth * (TilePixelWidth \ 2)) + TilePixelWidth \ 2
        End If

        If GrhData(grh.GrhIndex).TileHeight <> 1 Then
            y = y - Int(GrhData(CurrentGrhIndex).TileHeight * TilePixelHeight) + TilePixelHeight
        End If
    End If

    With GrhData(CurrentGrhIndex)

        Dim Texture As Direct3DTexture8

        Dim TextureWidth As Long, TextureHeight As Long
        Set Texture = SurfaceDB.GetTexture(.FileNum, TextureWidth, TextureHeight)

        Call SpriteBatch.SetTexture(Texture)

        Call SpriteBatch.SetAlpha(Alpha)
        
        If .Tx2 = 0 And .FileNum > 0 Then
            .Tx1 = .sX / TextureWidth
            .Tx2 = (.sX + .pixelWidth) / TextureWidth
            .Ty1 = .sY / TextureHeight
            .Ty2 = (.sY + .pixelHeight) / TextureHeight
        End If

        Call SpriteBatch.Draw(x, y, .pixelWidth, .pixelHeight, rgb_list, .Tx1, .Ty1, .Tx2, .Ty2, Angle)

    End With

    
    Exit Sub

Draw_GrhFX_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Draw_GrhFX", Erl)
    Resume Next
    
End Sub

Private Sub Draw_GrhSinLuz(ByRef grh As grh, ByVal x As Integer, ByVal y As Integer, ByVal center As Byte, ByVal animate As Byte, Optional ByVal Alpha As Boolean, Optional ByVal map_x As Byte = 1, Optional ByVal map_y As Byte = 1, Optional ByVal Angle As Single)
    
    On Error GoTo Draw_GrhSinLuz_Err
    

    

    If grh.GrhIndex = 0 Or grh.GrhIndex > MaxGrh Then Exit Sub
    
    Dim CurrentFrame As Integer
    CurrentFrame = 1

    If animate Then
        If grh.Started > 0 Then
            Dim ElapsedFrames As Long
            ElapsedFrames = Fix((FrameTime - grh.Started) / grh.speed)

            If grh.Loops = INFINITE_LOOPS Or ElapsedFrames < GrhData(grh.GrhIndex).NumFrames * (grh.Loops + 1) Then
                CurrentFrame = ElapsedFrames Mod GrhData(grh.GrhIndex).NumFrames + 1

            Else
                grh.Started = 0
            End If

        End If

    End If
    
    Dim CurrentGrhIndex As Long
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(grh.GrhIndex).Frames(CurrentFrame)

    'Center Grh over X,Y pos
    If center Then
        If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
            x = x - Int(GrhData(CurrentGrhIndex).TileWidth * (TilePixelWidth \ 2)) + TilePixelWidth \ 2

        End If

        If GrhData(grh.GrhIndex).TileHeight <> 1 Then
            y = y - Int(GrhData(CurrentGrhIndex).TileHeight * TilePixelHeight) + TilePixelHeight

        End If

    End If
    
    Static light_value(3) As RGBA

    light_value(0) = global_light
    light_value(1) = light_value(0)
    light_value(2) = light_value(0)
    light_value(3) = light_value(0)

    'Device_Box_Textured_Render CurrentGrhIndex, x, y, GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, light_value, GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, Alpha, angle
    Call Batch_Textured_Box(x, y, GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, GrhData(CurrentGrhIndex).FileNum, light_value, Alpha, Angle)

    
    Exit Sub

Draw_GrhSinLuz_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Draw_GrhSinLuz", Erl)
    Resume Next
    
End Sub

Public Sub render()
    
    On Error GoTo render_Err
    

    '*****************************************************
    '****** Coded by Menduz (lord.yo.wo@gmail.com) *******
    '*****************************************************
    Rem On Error GoTo ErrorHandler:
    Dim temp_array(3) As RGBA

    Call Engine_BeginScene
    
    Call ShowNextFrame

    frmMain.ms.Caption = PingRender & "ms"
       
    If frmMain.Contadores.Enabled Then

        Dim PosY As Integer
       
        Dim PosX As Integer

        PosY = -10
        PosX = 5

        If DrogaCounter > 0 Then
            Call RGBAList(temp_array, 0, 153, 0)
            
            PosY = PosY + 15
            Engine_Text_Render "Potenciado: " & CLng(DrogaCounter) & "s", PosX, PosY, temp_array, 1, True, 0, 160
            
        End If
        
        If OxigenoCounter > 0 Then

            Dim HR                  As Integer

            Dim ms                  As Integer

            Dim SS                  As Integer

            Dim secs                As Integer

            Dim TextoOxigenoCounter As String
        
            Call RGBAList(temp_array, 50, 100, 255)

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
                frmMain.oxigenolbl = SS
                frmMain.oxigenolbl.ForeColor = vbRed
            Else
                frmMain.oxigenolbl = ms
                frmMain.oxigenolbl.ForeColor = vbWhite

            End If

            Engine_Text_Render "Oxigeno: " & TextoOxigenoCounter, PosX, PosY, temp_array, 1, True, 0, 128

        End If

    End If
    
    If FadeInAlpha > 0 Then
        Call Engine_Draw_Box(0, 0, frmMain.renderer.ScaleWidth, frmMain.renderer.ScaleHeight, RGBA_From_Comp(0, 0, 0, FadeInAlpha))
        FadeInAlpha = FadeInAlpha - 1
    End If
    
    Call Engine_EndScene(Render_Main_Rect)
    
    FrameTime = (timeGetTime() And &H7FFFFFFF)
    FramesPerSecCounter = FramesPerSecCounter + 1
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
    
    Engine_ActFPS

    Exit Sub

render_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.render", Erl)
    Resume Next
    
End Sub

Sub ShowNextFrame()
    
    On Error GoTo ShowNextFrame_Err
    

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

    Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
    
    Call RenderScreen(UserPos.x - AddtoUserPos.x, UserPos.y - AddtoUserPos.y, OffsetCounterX, OffsetCounterY, HalfWindowTileWidth, HalfWindowTileHeight)

    
    Exit Sub

ShowNextFrame_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.ShowNextFrame", Erl)
    Resume Next
    
End Sub

Private Sub Device_Box_Textured_Render_Advance(ByVal GrhIndex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, ByVal src_height As Integer, ByRef rgb_list() As RGBA, ByVal src_x As Integer, ByVal src_y As Integer, ByVal dest_width As Integer, Optional ByVal dest_height As Integer, Optional ByVal alpha_blend As Boolean, Optional ByVal Angle As Single)
    
    On Error GoTo Device_Box_Textured_Render_Advance_Err
    

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 5/15/2003
    'Copies the Textures allowing resizing
    'Modified by Juan Martín Sotuyo Dodero
    '**************************************************************
    Static src_rect            As RECT

    Static dest_rect           As RECT

    Static temp_verts(3)       As TYPE_VERTEX

    Static d3dTextures         As D3D8Textures

    Static light_value(0 To 3) As RGBA
    
    If GrhIndex = 0 Then Exit Sub
    
    Set d3dTextures.Texture = SurfaceDB.GetTexture(GrhData(GrhIndex).FileNum, d3dTextures.texwidth, d3dTextures.texheight)
    
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
        .Bottom = src_y + src_height
        .Left = src_x
        .Right = src_x + src_width
        .Top = src_y

    End With
        
    'Set up the destination rectangle
    With dest_rect
        .Bottom = dest_y + dest_height
        .Left = dest_x
        .Right = dest_x + dest_width
        .Top = dest_y

    End With
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value, d3dTextures.texwidth, d3dTextures.texheight, Angle
        
    'Set Textures
    DirectDevice.SetTexture 0, d3dTextures.Texture
    
    If alpha_blend Then
        'Set Rendering for alphablending
        DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE

    End If
    
    'Draw the triangles that make up our square Textures
    DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

    End If

    
    Exit Sub

Device_Box_Textured_Render_Advance_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Device_Box_Textured_Render_Advance", Erl)
    Resume Next
    
End Sub

Public Sub Batch_Textured_Box(ByVal x As Long, ByVal y As Long, _
                                ByVal Width As Integer, ByVal Height As Integer, _
                                ByVal sX As Integer, ByVal sY As Integer, _
                                ByVal tex As Long, _
                                ByRef Color() As RGBA, _
                                Optional ByVal Alpha As Boolean = False, _
                                Optional ByVal Angle As Single = 0, _
                                Optional ByVal ScaleX As Single = 1!, _
                                Optional ByVal ScaleY As Single = 1!)
    
    On Error GoTo Batch_Textured_Box_Err
    

    Dim Texture As Direct3DTexture8
        
    Dim TextureWidth As Long, TextureHeight As Long
    Set Texture = SurfaceDB.GetTexture(tex, TextureWidth, TextureHeight)
    
    With SpriteBatch

        Call .SetTexture(Texture)
            
        Call .SetAlpha(Alpha)

        If TextureWidth <> 0 And TextureHeight <> 0 Then
            Call .Draw(x, y, Width * ScaleX, Height * ScaleY, Color, (sX + 0.25) / TextureWidth, (sY + 0.25) / TextureHeight, (sX + Width) / TextureWidth, (sY + Height) / TextureHeight, Angle)
        Else
            Call .Draw(x, y, TextureWidth * ScaleX, TextureHeight * ScaleY, Color, , , , , Angle)
        End If
            
    End With
        
    
    Exit Sub

Batch_Textured_Box_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Batch_Textured_Box", Erl)
    Resume Next
    
End Sub

Public Sub Batch_Textured_Box_Advance(ByVal x As Long, ByVal y As Long, _
                                ByVal Width As Integer, ByVal Height As Integer, _
                                ByVal sX As Integer, ByVal sY As Integer, _
                                ByVal tex As Long, _
                                ByVal dw As Integer, ByVal dH As Integer, _
                                ByRef Color() As RGBA, _
                                Optional ByVal Alpha As Boolean = False, _
                                Optional ByVal Angle As Single = 0, _
                                Optional ByVal ScaleX As Single = 1!, _
                                Optional ByVal ScaleY As Single = 1!, _
                                Optional ByVal z As Long = 1)
    
    On Error GoTo Batch_Textured_Box_Advance_Err
    

    Dim Texture As Direct3DTexture8
        
    Dim TextureWidth As Long, TextureHeight As Long
    Set Texture = SurfaceDB.GetTexture(tex, TextureWidth, TextureHeight)
    
    With SpriteBatch

        Call .SetTexture(Texture)
            
        Call .SetAlpha(Alpha)
        
        If TextureWidth <> 0 And TextureHeight <> 0 Then
            Call .Draw(x, y, dw * ScaleX, dH * ScaleY, Color, (sX + 0.25) / TextureWidth, (sY + 0.25) / TextureHeight, (sX + Width) / TextureWidth, (sY + Height) / TextureHeight, Angle)
        Else
            Call .Draw(x, y, TextureWidth * ScaleX, TextureHeight * ScaleY, Color, , , , , Angle)
        End If
            
    End With
        
    
    Exit Sub

Batch_Textured_Box_Advance_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Batch_Textured_Box_Advance", Erl)
    Resume Next
    
End Sub

Public Sub Batch_Textured_Box_Pre(ByVal x As Long, ByVal y As Long, _
                                ByVal Width As Integer, ByVal Height As Integer, _
                                ByVal sX As Single, ByVal sY As Single, _
                                ByVal sW As Single, ByVal sH As Single, _
                                ByVal tex As Long, _
                                ByRef Color() As RGBA, _
                                Optional ByVal Alpha As Boolean = False, _
                                Optional ByVal Angle As Single = 0, _
                                Optional ByVal ScaleX As Single = 1!, _
                                Optional ByVal ScaleY As Single = 1!)
    
    On Error GoTo Batch_Textured_Box_Pre_Err
    

    Dim Texture As Direct3DTexture8
        
    Dim TextureWidth As Long, TextureHeight As Long
    Set Texture = SurfaceDB.GetTexture(tex, TextureWidth, TextureHeight)
    
    With SpriteBatch

        Call .SetTexture(Texture)

        Call .SetAlpha(Alpha)

        Call .Draw(x, y, Width * ScaleX, Height * ScaleY, Color, sX, sY, sW, sH, Angle)

    End With
        
    
    Exit Sub

Batch_Textured_Box_Pre_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Batch_Textured_Box_Pre", Erl)
    Resume Next
    
End Sub

Public Sub Batch_Textured_Box_Shadow(ByVal x As Long, ByVal y As Long, _
                                ByVal Width As Integer, ByVal Height As Integer, _
                                ByVal sX As Integer, ByVal sY As Integer, _
                                ByVal tex As Long, _
                                ByRef Color() As RGBA)
    
    On Error GoTo Batch_Textured_Box_Shadow_Err
    

    Dim Texture As Direct3DTexture8
        
    Dim TextureWidth As Long, TextureHeight As Long
    Set Texture = SurfaceDB.GetTexture(tex, TextureWidth, TextureHeight)
    
    With SpriteBatch

        Call .SetTexture(Texture)
            
        Call .SetAlpha(False)
        
        If TextureWidth <> 0 And TextureHeight <> 0 Then
            Call .DrawShadow(x, y, Width, Height, Color, (sX + 0.25) / TextureWidth, (sY + 0.25) / TextureHeight, (sX + Width) / TextureWidth, (sY + Height) / TextureHeight)
        Else
            Call .DrawShadow(x, y, TextureWidth, TextureHeight, Color)
        End If
            
    End With
        
    
    Exit Sub

Batch_Textured_Box_Shadow_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Batch_Textured_Box_Shadow", Erl)
    Resume Next
    
End Sub

Public Sub Device_Box_Textured_Render(ByVal GrhIndex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, ByVal src_height As Integer, ByRef Color() As RGBA, ByVal src_x As Integer, ByVal src_y As Integer, Optional ByVal alpha_blend As Boolean, Optional ByVal Angle As Single)
    
    On Error GoTo Device_Box_Textured_Render_Err
    

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 2/12/2004
    'Just copies the Textures
    '**************************************************************
    Static src_rect            As RECT

    Static dest_rect           As RECT

    Static temp_verts(3)       As TYPE_VERTEX

    Static d3dTextures         As D3D8Textures
    
    If GrhIndex = 0 Then Exit Sub
    
    Set d3dTextures.Texture = SurfaceDB.GetTexture(GrhData(GrhIndex).FileNum, d3dTextures.texwidth, d3dTextures.texheight)

    'Set up the source rectangle
    With src_rect
        .Bottom = src_y + src_height
        .Left = src_x
        .Right = src_x + src_width
        .Top = src_y
    End With
                
    'Set up the destination rectangle
    With dest_rect
        .Bottom = dest_y + src_height
        .Left = dest_x
        .Right = dest_x + src_width
        .Top = dest_y
    End With
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, Color(), d3dTextures.texwidth, d3dTextures.texheight, Angle
     
    'Set Textures
    DirectDevice.SetTexture 0, d3dTextures.Texture
    
    If alpha_blend Then
        'Set Rendering for alphablending
        DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE

    End If
    
    'Draw the triangles that make up our square Textures
    DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

    End If
    
    DirectDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    DirectDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
 
    'DirectDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_DIFFUSE
    'DirectDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1

    
    Exit Sub

Device_Box_Textured_Render_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Device_Box_Textured_Render", Erl)
    Resume Next
    
End Sub

Sub Char_TextRender(ByVal charindex As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer, ByVal x As Byte, ByVal y As Byte)
    
    On Error GoTo Char_TextRender_Err
    

    Dim moved         As Boolean

    Dim Pos           As Integer

    Dim line          As String

    Dim Color(0 To 3) As Long

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
        If .dialog <> "" Then

            'Figure out screen position
            Dim temp_array(3) As RGBA

            Dim PixelY        As Integer

            PixelY = PixelOffsetY
            
            Call Long_2_RGBAList(temp_array, .dialog_color)

            If .dialog_scroll Then
                .dialog_offset_counter_y = .dialog_offset_counter_y + (scroll_dialog_pixels_per_frame * timerTicksPerFrame * Sgn(-1))

                If Sgn(.dialog_offset_counter_y) = -1 Then
                    .dialog_offset_counter_y = 0
                    .dialog_scroll = False

                End If

                Engine_Text_Render .dialog, PixelOffsetX + 14 - CInt(Engine_Text_Width(.dialog, True) / 2), PixelY + .Body.HeadOffset.y - Engine_Text_Height(.dialog, True) + .dialog_offset_counter_y, temp_array, 1, True, MapData(x, y).charindex
            Else
                Engine_Text_Render .dialog, PixelOffsetX + 14 - CInt(Engine_Text_Width(.dialog, True) / 2), PixelY + .Body.HeadOffset.y - Engine_Text_Height(.dialog, True), temp_array, 1, True, MapData(x, y).charindex

            End If

        End If
        
        If UBound(.DialogEffects) > 0 Then

            For i = 1 To UBound(.DialogEffects)
            
                With .DialogEffects(i)
                
                    If LenB(.Text) <> 0 Then
                        Dim DeltaTime As Long
                        DeltaTime = FrameTime - .Start
    
                        If DeltaTime > 1300 Then
                            .Text = vbNullString
                        Else
                            If DeltaTime > 900 Then
                                Call RGBAList(temp_array, .Color.r, .Color.G, .Color.B, .Color.A * (1300 - DeltaTime) * 0.0025)
                            Else
                                Call RGBAList(temp_array, .Color.r, .Color.G, .Color.B, .Color.A)
                            End If
                    
                            Engine_Text_Render_Efect charindex, .Text, PixelOffsetX + 14 - Int(Engine_Text_Width(.Text, True) * 0.5), PixelOffsetY + charlist(charindex).Body.HeadOffset.y - Engine_Text_Height(.Text, True) - DeltaTime * 0.025, temp_array, 1, True
            
                        End If
                    End If
                    
                End With
                
            Next

        End If
            
        '*** End Dialogs ***
    End With

    
    Exit Sub

Char_TextRender_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Char_TextRender", Erl)
    Resume Next
    
End Sub

Sub Char_Render(ByVal charindex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer, ByVal x As Byte, ByVal y As Byte)
    
    On Error GoTo Char_Render_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 12/03/04
    'Draw char's to screen without offcentering them
    '***************************************************
    Dim Pos                 As Integer

    Dim line                As String

    Dim Color(3)            As RGBA
    
    Dim NameColor(3)        As RGBA

    Dim colorCorazon(3)     As RGBA

    Dim i                   As Long

    Dim OffsetYname         As Byte

    Dim OffsetYClan         As Byte
    
    Dim TextureX            As Integer

    Dim TextureY            As Integer
    
    Dim OffArma             As Single

    Dim OffAuras            As Integer

    Dim OffHead             As Single
    
    Dim MostrarNombre       As Boolean
    
    
    Dim TempGrh As grh
    
    With charlist(charindex)

        If .Heading = 0 Then Exit Sub
    
        If .Moving Then

            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame * .Speeding

                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0
                End If

            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame * .Speeding

                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0
                End If

            End If
            
            If .scrollDirectionX = 0 And .scrollDirectionY = 0 Then
                .Moving = False
            End If

        ElseIf Not .Idle Then
            
            If .Muerto Then
                If charindex <> UserCharIndex Then
                    ' Si no somos nosotros, esperamos un intervalo
                    ' antes de poner la animación idle para evitar saltos
                    If FrameTime - .LastStep > TIME_CASPER_IDLE Then
                        .Body = BodyData(CASPER_BODY_IDLE)
                        .Body.Walk(.Heading).Started = FrameTime
                        .Idle = True
                    End If
                    
                Else
                    .Body = BodyData(CASPER_BODY_IDLE)
                    .Body.Walk(.Heading).Started = FrameTime
                    .Idle = True
                End If

            Else
                'Stop animations
                If .Navegando = False Then
                    .Body.Walk(.Heading).Started = 0
                    If Not .MovArmaEscudo Then
                        .Arma.WeaponWalk(.Heading).Started = 0
                        .Escudo.ShieldWalk(.Heading).Started = 0
                    End If
                End If
            End If
        End If

        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
        
        Dim ease As Single
        If MostrarRespiracion Then
            ease = EaseBreathing((((FrameTime - .TimeCreated) * 0.25) Mod 1000) * 0.001)
        Else
            ease = 0
        End If

        If .Body.Walk(.Heading).GrhIndex Then

            If UserCiego Then
                MostrarNombre = False
                
            ElseIf .Invisible Then

                If IsCharVisible(charindex) Then
                    Call RGBAList(Color, 255, 255, 255, 100)
                    
                    If .priv = 0 Then
                        
                        Select Case .status
                            ' Criminal
                            Case 0
                                Call RGBAList(NameColor, ColoresPJ(50).r, ColoresPJ(50).G, ColoresPJ(50).B)
                                Call RGBAList(colorCorazon, ColoresPJ(50).r, ColoresPJ(50).G, ColoresPJ(50).B)
                            
                            ' Ciudadano
                            Case 1
                                Call RGBAList(NameColor, ColoresPJ(49).r, ColoresPJ(49).G, ColoresPJ(49).B)
                                Call RGBAList(colorCorazon, ColoresPJ(49).r, ColoresPJ(49).G, ColoresPJ(49).B)
                            
                            ' Caos
                            Case 2
                                Call RGBAList(NameColor, ColoresPJ(6).r, ColoresPJ(6).G, ColoresPJ(6).B)
                                Call RGBAList(colorCorazon, ColoresPJ(6).r, ColoresPJ(6).G, ColoresPJ(6).B)
    
                            ' Armada
                            Case 3
                                Call RGBAList(NameColor, ColoresPJ(8).r, ColoresPJ(8).G, ColoresPJ(8).B)
                                Call RGBAList(colorCorazon, ColoresPJ(8).r, ColoresPJ(8).G, ColoresPJ(8).B)
    
                        End Select
                                
                    Else
                        Call RGBAList(NameColor, ColoresPJ(.priv).r, ColoresPJ(.priv).G, ColoresPJ(.priv).B)
                        Call RGBAList(colorCorazon, ColoresPJ(.priv).r, ColoresPJ(.priv).G, ColoresPJ(.priv).B)
                        
                    End If
                                
                    MostrarNombre = True
                        
                Else
                    Call RGBAList(Color, 0, 0, 0, 0)
                    MostrarNombre = False
                End If

            Else
                If .Muerto Then
                    Call Copy_RGBAList_WithAlpha(Color, MapData(x, y).light_value, 150)
                Else
                    Call Copy_RGBAList(Color, MapData(x, y).light_value)
                End If

                If .EsNpc Then
                    If Abs(tX - .Pos.x) < 1 And tY - .Pos.y < 1 And .Pos.y - tY < 2 Then
                        MostrarNombre = True
                        Call RGBAList(NameColor, 210, 105, 30)
                        Call InitGrh(TempGrh, 839)
                        
                        If .UserMinHp > 0 Then
                            Dim TempColor(3) As RGBA
                            Call RGBAList(TempColor, 255, 255, 255, 200)
                            Call Draw_Grh(TempGrh, PixelOffsetX + 1, PixelOffsetY + 10, 1, 0, TempColor, False, 0, 0, 0)
                           
                            Engine_Draw_Box PixelOffsetX + 4, PixelOffsetY + 36, (((.UserMinHp + 1 / 100) / (.UserMaxHp + 1 / 100))) * 26, 4, RGBA_From_Comp(255, 0, 0, 255)
                        End If
                    End If

                    If .simbolo <> 0 Then
                        Call Draw_GrhIndex(5257 + .simbolo, PixelOffsetX + 6, PixelOffsetY + .Body.HeadOffset.y - 12 - 10 * Sin((FrameTime Mod 31415) * 0.002) ^ 2)
                    End If
                    
                Else
                    MostrarNombre = True
                    
                    If .priv = 0 Then
                        
                        Select Case .status
                            ' Criminal
                            Case 0
                                Call RGBAList(NameColor, ColoresPJ(50).r, ColoresPJ(50).G, ColoresPJ(50).B)
                                Call RGBAList(colorCorazon, ColoresPJ(50).r, ColoresPJ(50).G, ColoresPJ(50).B)
                            
                            ' Ciudadano
                            Case 1
                                Call RGBAList(NameColor, ColoresPJ(49).r, ColoresPJ(49).G, ColoresPJ(49).B)
                                Call RGBAList(colorCorazon, ColoresPJ(49).r, ColoresPJ(49).G, ColoresPJ(49).B)
                            
                            ' Caos
                            Case 2
                                Call RGBAList(NameColor, ColoresPJ(6).r, ColoresPJ(6).G, ColoresPJ(6).B)
                                Call RGBAList(colorCorazon, ColoresPJ(6).r, ColoresPJ(6).G, ColoresPJ(6).B)
    
                            ' Armada
                            Case 3
                                Call RGBAList(NameColor, ColoresPJ(8).r, ColoresPJ(8).G, ColoresPJ(8).B)
                                Call RGBAList(colorCorazon, ColoresPJ(8).r, ColoresPJ(8).G, ColoresPJ(8).B)
    
                        End Select
                                
                    Else
                        Call RGBAList(NameColor, ColoresPJ(.priv).r, ColoresPJ(.priv).G, ColoresPJ(.priv).B)
                        Call RGBAList(colorCorazon, ColoresPJ(.priv).r, ColoresPJ(.priv).G, ColoresPJ(.priv).B)
                        
                    End If
                                        
                    If .group_index > 0 Then
                        If charlist(charindex).group_index = charlist(UserCharIndex).group_index Then
                            Call Copy_RGBAList(Color, COLOR_WHITE)
                            Call SetRGBA(colorCorazon(0), 255, 255, 0)
                            Call SetRGBA(colorCorazon(1), 0, 255, 255)
                            Call SetRGBA(colorCorazon(2), 0, 255, 0)
                            Call SetRGBA(colorCorazon(3), 0, 255, 255)
                        End If
                    End If
                        
                    If .clan_index > 0 Then
                        If .clan_index = charlist(UserCharIndex).clan_index And charindex <> UserCharIndex And .Muerto = 0 Then
                            If .clan_nivel = 5 Then
                                OffsetYname = 8
                                OffsetYClan = 6
                                Grh_Render Marco, PixelOffsetX, PixelOffsetY + 5, Color, True, True, False
                                Engine_Draw_Box_Border PixelOffsetX + 3, PixelOffsetY + 31, (((.UserMinHp + 1 / 100) / (.UserMaxHp + 1 / 100))) * 26, 4, RGBA_From_Comp(255, 200, 0, 0), RGBA_From_Comp(0, 200, 200, 200)
                            End If
                        End If
                    End If
                End If
            End If
            
            ' Si tiene cabeza, componemos la textura
            If .Head.Head(.Heading).GrhIndex Then
            
                If .EsEnano Then
                    OffArma = 7
                    OffAuras = 7
                End If

                OffArma = OffArma - ease * 4
                OffHead = .Body.HeadOffset.y - ease * 2
    
                BeginComposedTexture
                
                TextureX = ComposedTextureCenterX - 16
                TextureY = ComposedTextureHeight - 32
                                
                Select Case .Heading

                    Case E_Heading.EAST
    
                        If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), TextureX, TextureY + OffArma, 1, 1, COLOR_WHITE, False, x, y)
                                                                    
                        Call Draw_Grh_Breathing(.Body.Walk(.Heading), TextureX, TextureY, 1, 1, COLOR_WHITE, ease)
                                         
                        Call Draw_Grh(.Head.Head(.Heading), TextureX + .Body.HeadOffset.x, TextureY + OffHead, 1, 0, COLOR_WHITE, False, x, y)
                                         
                        If .Casco.Head(.Heading).GrhIndex Then Call Draw_Grh(.Casco.Head(.Heading), TextureX + .Body.HeadOffset.x, TextureY + OffHead, 1, 0, COLOR_WHITE, False, x, y)
                                             
                        If .Arma.WeaponWalk(.Heading).GrhIndex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), TextureX, TextureY + OffArma, 1, 1, COLOR_WHITE, False, x, y)
                                     
                    Case E_Heading.NORTH
    
                        If .Arma.WeaponWalk(.Heading).GrhIndex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), TextureX, TextureY + OffArma, 1, 1, COLOR_WHITE, False, x, y)
                                             
                        If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), TextureX, TextureY + OffArma, 1, 1, COLOR_WHITE, False, x, y)
                                             
                        Call Draw_Grh_Breathing(.Body.Walk(.Heading), TextureX, TextureY, 1, 1, COLOR_WHITE, ease)
                                         
                        Call Draw_Grh(.Head.Head(.Heading), TextureX + .Body.HeadOffset.x, TextureY + OffHead, 1, 0, COLOR_WHITE, False, x, y)
                                         
                        If .Casco.Head(.Heading).GrhIndex Then Call Draw_Grh(.Casco.Head(.Heading), TextureX + .Body.HeadOffset.x, TextureY + OffHead, 1, 0, COLOR_WHITE, False, x, y)
                                     
                    Case E_Heading.WEST
    
                        If .Arma.WeaponWalk(.Heading).GrhIndex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), TextureX, TextureY + OffArma, 1, 1, COLOR_WHITE, False, x, y)
                                             
                        Call Draw_Grh_Breathing(.Body.Walk(.Heading), TextureX, TextureY, 1, 1, COLOR_WHITE, ease)
                                         
                        Call Draw_Grh(.Head.Head(.Heading), TextureX + .Body.HeadOffset.x, TextureY + OffHead, 1, 0, COLOR_WHITE, False, x, y)
                                         
                        If .Casco.Head(.Heading).GrhIndex Then Call Draw_Grh(.Casco.Head(.Heading), TextureX + .Body.HeadOffset.x, TextureY + OffHead, 1, 0, COLOR_WHITE, False, x, y)
                                         
                        If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), TextureX, TextureY + OffArma, 1, 1, COLOR_WHITE, False, x, y)
    
                    Case E_Heading.south
                                         
                        Call Draw_Grh_Breathing(.Body.Walk(.Heading), TextureX, TextureY, 1, 1, COLOR_WHITE, ease)
                                         
                        Call Draw_Grh(.Head.Head(.Heading), TextureX + .Body.HeadOffset.x, TextureY + OffHead, 1, 0, COLOR_WHITE, False, x, y)
                                         
                        If .Casco.Head(.Heading).GrhIndex Then Call Draw_Grh(.Casco.Head(.Heading), TextureX + .Body.HeadOffset.x, TextureY + OffHead, 1, 0, COLOR_WHITE, False, x, y)
                                         
                        If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), TextureX, TextureY + OffArma, 1, 1, COLOR_WHITE, False, x, y)
                                             
                        If .Arma.WeaponWalk(.Heading).GrhIndex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), TextureX, TextureY + OffArma, 1, 1, COLOR_WHITE, False, x, y)
    
                End Select

                EndComposedTexture

                If Not .Invisible Then
                    ' Reflejo
                    PresentComposedTexture PixelOffsetX, PixelOffsetY, Color, 0, , True

                    ' Sombra
                    PresentComposedTexture PixelOffsetX, PixelOffsetY, Color, 0, True

                    If LenB(.Body_Aura) <> 0 And .Body_Aura <> "0" Then Call Renderizar_Aura(.Body_Aura, PixelOffsetX, PixelOffsetY + OffArma, x, y, charindex)
                    If LenB(.Arma_Aura) <> 0 And .Arma_Aura <> "0" Then Call Renderizar_Aura(.Arma_Aura, PixelOffsetX, PixelOffsetY + OffArma, x, y, charindex)
                    If LenB(.Otra_Aura) <> 0 And .Otra_Aura <> "0" Then Call Renderizar_Aura(.Otra_Aura, PixelOffsetX, PixelOffsetY + OffArma, x, y, charindex)
                    If LenB(.Escudo_Aura) <> 0 And .Escudo_Aura <> "0" Then Call Renderizar_Aura(.Escudo_Aura, PixelOffsetX, PixelOffsetY + OffArma, x, y, charindex)
                    If LenB(.DM_Aura) <> 0 And .DM_Aura <> "0" Then Call Renderizar_Aura(.DM_Aura, PixelOffsetX, PixelOffsetY + OffArma, x, y, charindex)
                    If LenB(.RM_Aura) <> 0 And .RM_Aura <> "0" Then Call Renderizar_Aura(.RM_Aura, PixelOffsetX, PixelOffsetY + OffArma, x, y, charindex)
                End If

                ' Char
                PresentComposedTexture PixelOffsetX, PixelOffsetY, Color, False
                
            ' Si no, solo dibujamos body
            Else
                Call Draw_Sombra(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, False, x, y)
                Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, Color, False, x, y)
            End If
    
            'Draw name over head
            If Nombres And Len(.nombre) > 0 And MostrarNombre Then
                
                Pos = InStr(.nombre, "<")
                
                If Pos = 0 Then Pos = InStr(.nombre, "[")

                If Pos = 0 Then Pos = Len(.nombre) + 2

                'Nick
                line = Left$(.nombre, Pos - 2)
                Engine_Text_Render line, PixelOffsetX + 15 - CInt(Engine_Text_Width(line, True) / 2), PixelOffsetY + 30 + OffsetYname - Engine_Text_Height(line, True), NameColor, 1, False, 0, IIf(.Invisible, 160, 255)

                
                'Clan
                If .priv > 1 And .priv < &H40 Then
                    line = "<Game Master>"
                Else
                    line = .clan
                End If
                    
                Engine_Text_Render line, PixelOffsetX + 15 - CInt(Engine_Text_Width(line, True) / 2), PixelOffsetY + 45 + OffsetYClan - Engine_Text_Height(line, True), NameColor, 1, False, 0, IIf(.Invisible, 160, 255)

                If .Donador = 1 Then
                    Grh_Render Estrella, PixelOffsetX + 7 + CInt(Engine_Text_Width(.nombre, 1) / 2), PixelOffsetY + 10 + OffsetYname, colorCorazon, True, True, False
                End If
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
            Call InitGrh(TempGrh, 839)
            Call RGBAList(Color, 255, 255, 255, 200)

            Call Draw_Grh(TempGrh, PixelOffsetX + 1, PixelOffsetY - 55, 1, 0, Color, False, 0, 0, 0)
            
            Engine_Draw_Box_Border PixelOffsetX + 5, PixelOffsetY - 29, (((.BarTime / 100) / (.MaxBarTime / 100))) * 24, 3, RGBA_From_Comp(0, 128, 128, 255), RGBA_From_Comp(0, 0, 0, 255)

            .BarTime = .BarTime + (timerTicksPerFrame * 4)
                             
            If .BarTime >= .MaxBarTime Then
                charlist(charindex).BarTime = 0
                charlist(charindex).BarAccion = 99
                charlist(charindex).MaxBarTime = 0
            End If

        End If
                        
        If .Escribiendo = True And Not .Invisible Then

            
            Call InitGrh(TempGrh, 32017)
            Call RGBAList(Color, 255, 255, 255, 200)

            Call Draw_Grh(TempGrh, PixelOffsetX + 20, PixelOffsetY - 45, 1, 0, Color, False, 0, 0, 0)

        End If

        ' Meditación
        If .FxIndex <> 0 And .fX.Started <> 0 Then
            Call RGBAList(Color, 255, 255, 255, 180)

            Call Draw_GrhFX(.fX, PixelOffsetX + FxData(.FxIndex).OffsetX, PixelOffsetY + FxData(.FxIndex).OffsetY + 4, 1, 1, Color, False, , , , charindex)
       
        End If

        If .FxCount > 0 Then

            For i = 1 To .FxCount

                If .FxList(i).FxIndex > 0 And .FxList(i).Started <> 0 Then
                    Call RGBAList(Color, 255, 255, 255, 220)

                    If FxData(.FxList(i).FxIndex).IsPNG = 1 Then
                        Call Draw_GrhFX(.FxList(i), PixelOffsetX + FxData(.FxList(i).FxIndex).OffsetX, PixelOffsetY + FxData(.FxList(i).FxIndex).OffsetY + 20, 1, 1, Color, False, , , , charindex)
                    Else
                        Call Draw_GrhFX(.FxList(i), PixelOffsetX + FxData(.FxList(i).FxIndex).OffsetX, PixelOffsetY + FxData(.FxList(i).FxIndex).OffsetY + 20, 1, 1, Color, True, , , , charindex)
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

    
    Exit Sub

Char_Render_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Char_Render", Erl)
    Resume Next
    
End Sub

Public Function IsCharVisible(ByVal charindex As Integer) As Boolean

    With charlist(charindex)
    
        If charindex = UserCharIndex Then
            IsCharVisible = True
            Exit Function
        End If
        
        If charlist(UserCharIndex).priv > 0 And .priv <= charlist(UserCharIndex).priv Then
            IsCharVisible = True
            Exit Function
        End If
        
        If .clan_index > 0 Then
            If .clan_index = charlist(UserCharIndex).clan_index Then
                If .clan_nivel >= 3 Then
                    IsCharVisible = True
                    Exit Function
                End If
            End If
        End If

    End With

End Function

Public Sub Start()
    
    On Error GoTo Start_Err
    

    DoEvents

    Do While prgRun

        Call FlushBuffer

        If frmMain.WindowState <> vbMinimized Then
            Select Case QueRender

                Case 0
                    render
                
                    Check_Keys
                    Moviendose = False
                    DrawMainInventory

                    If frmComerciar.Visible Then
                        DrawInterfaceComerciar

                    ElseIf frmBancoObj.Visible Then
                        DrawInterfaceBoveda
                    End If
                    
                    If frmBancoCuenta.Visible Then
                        DrawInterfaceBovedaCuenta
                    End If
                    
                    If frmMapaGrande.Visible Then
                        DrawMapaMundo
                    End If
                    
                    If FrmKeyInv.Visible Then
                        DrawInterfaceKeys
                    End If
                    
                    If frmComerciarUsu.Visible Then
                        DrawInventoryComercio
                        DrawInventoryUserComercio
                        DrawInventoryOtherComercio
                    End If

                Case 1
                    If Not frmConnect.Visible Then
                        frmConnect.Show
                        FrmLogear.Show , frmConnect
                        FrmLogear.Top = FrmLogear.Top + 3500
                    End If
                    
                    RenderConnect 57, 45, 0, 0

                Case 2
                    rendercuenta 42, 43, 0, 0

                Case 3
                    RenderCrearPJ 76, 82, 0, 0

            End Select

            Sound.Sound_Render
        Else
            Sleep 60&
            Call frmMain.Inventario.ReDraw
        End If

        DoEvents

    Loop

    EngineRun = False

    Call CloseClient

    
    Exit Sub

Start_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Start", Erl)
    Resume Next
    
End Sub

Public Sub SetMapFx(ByVal x As Byte, ByVal y As Byte, ByVal fX As Integer, ByVal Loops As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 12/03/04
    'Sets an FX to the character.
    '***************************************************
    
    On Error GoTo SetMapFx_Err
    

    

    Dim indice As Byte

    With MapData(x, y)
    
        indice = Map_FX_Group_Next_Open(x, y)
    
        .FxList(indice).FxIndex = fX
        Call InitGrh(.FxList(indice), FxData(fX).Animacion)
        .FxList(indice).Loops = Loops
    
    End With

    
    Exit Sub

SetMapFx_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.SetMapFx", Erl)
    Resume Next
    
End Sub

Private Function Engine_FToDW(f As Single) As Long
    
    On Error GoTo Engine_FToDW_Err
    

    ' single > long
    Dim Buf As D3DXBuffer

    Set Buf = DirectD3D8.CreateBuffer(4)
    DirectD3D8.BufferSetData Buf, 0, 4, 1, f
    DirectD3D8.BufferGetData Buf, 0, 4, 1, Engine_FToDW

    
    Exit Function

Engine_FToDW_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Engine_FToDW", Erl)
    Resume Next
    
End Function

Private Function VectorToRGBA(Vec As D3DVECTOR, fHeight As Single) As Long
    
    On Error GoTo VectorToRGBA_Err
    

    Dim r As Integer, G As Integer, B As Integer, A As Integer

    r = 127 * Vec.x + 128
    G = 127 * Vec.y + 128
    B = 127 * Vec.z + 128
    A = 255 * fHeight
    VectorToRGBA = D3DColorARGB(A, r, G, B)

    
    Exit Function

VectorToRGBA_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.VectorToRGBA", Erl)
    Resume Next
    
End Function

Public Sub DrawMainInventory()
    
    On Error GoTo DrawMainInventory_Err
    

    ' Sólo dibujamos cuando es necesario
    If Not frmMain.Inventario.NeedsRedraw Then Exit Sub

    Dim InvRect As RECT

    InvRect.Left = 0
    InvRect.Top = 0
    InvRect.Right = frmMain.picInv.ScaleWidth
    InvRect.Bottom = frmMain.picInv.ScaleHeight

    ' Comenzamos la escena
    Call Engine_BeginScene

    ' Dibujamos el fondo del inventario principal
    'Call Draw_GrhIndex(6, 0, 0)

    ' Dibujamos items
    Call frmMain.Inventario.DrawInventory
    
    ' Dibujamos item arrastrado
    Call frmMain.Inventario.DrawDraggedItem

    ' Presentamos la escena
    Call Engine_EndScene(InvRect, frmMain.picInv.hwnd)

    
    Exit Sub

DrawMainInventory_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.DrawMainInventory", Erl)
    Resume Next
    
End Sub

Public Sub DrawInterfaceComerciar()
    
    On Error GoTo DrawInterfaceComerciar_Err
    

    ' Sólo dibujamos cuando es necesario
    If Not frmComerciar.InvComNpc.NeedsRedraw And Not frmComerciar.InvComUsu.NeedsRedraw Then Exit Sub

    Dim InvRect As RECT

    InvRect.Left = 0
    InvRect.Top = 0
    InvRect.Right = frmComerciar.interface.ScaleWidth
    InvRect.Bottom = frmComerciar.interface.ScaleHeight

    ' Comenzamos la escena
    Call Engine_BeginScene

    ' Dibujamos el fondo del inventario de comercio
    Call Draw_GrhIndex(837, 0, 0)

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
        'Call Draw_GrhColor(CurrentInventory.GrhIndex(CurrentInventory.SelectedItem), 282, 251, COLOR_WHITE)
    
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
        frmComerciar.lblcosto = PonerPuntos(Fix(CurrentInventory.Valor(CurrentInventory.SelectedItem) * cantidad))
        
        Set CurrentInventory = Nothing

    End If

    ' Presentamos la escena
    Call Engine_EndScene(InvRect, frmComerciar.interface.hwnd)

    
    Exit Sub

DrawInterfaceComerciar_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.DrawInterfaceComerciar", Erl)
    Resume Next
    
End Sub
Public Sub DrawInterfaceBovedaCuenta()
    
    On Error GoTo DrawInterfaceBoveda_Err
    

    ' Sólo dibujamos cuando es necesario
    If Not frmBancoCuenta.InvBovedaCuenta.NeedsRedraw And Not frmBancoCuenta.InvBankUsuCuenta.NeedsRedraw Then Exit Sub

    Dim InvRect As RECT

    InvRect.Left = 0
    InvRect.Top = 0
    InvRect.Right = frmBancoCuenta.interface.ScaleWidth
    InvRect.Bottom = frmBancoCuenta.interface.ScaleHeight

    ' Comenzamos la escena
    Call Engine_BeginScene

    ' Dibujamos el fondo de la bóveda
    Call Draw_GrhIndex(838, 0, 0)

    ' Dibujamos items de la bóveda
    Call frmBancoCuenta.InvBovedaCuenta.DrawInventory
    
    ' Dibujamos items del usuario
    Call frmBancoCuenta.InvBankUsuCuenta.DrawInventory

    ' Dibujamos "ambos" items arrastrados (aunque sólo puede estar uno activo a la vez)
    Call frmBancoCuenta.InvBovedaCuenta.DrawDraggedItem
    Call frmBancoCuenta.InvBankUsuCuenta.DrawDraggedItem
    
    ' Me fijo qué inventario está seleccionado
    Dim CurrentInventory As clsGrapchicalInventory
    
    If frmBancoCuenta.InvBovedaCuenta.SelectedItem > 0 Then
        Set CurrentInventory = frmBancoCuenta.InvBovedaCuenta
    ElseIf frmBancoCuenta.InvBankUsuCuenta.SelectedItem > 0 Then
        Set CurrentInventory = frmBancoCuenta.InvBankUsuCuenta

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
        
        frmBancoCuenta.lblnombre.Caption = CurrentInventory.ItemName(CurrentInventory.SelectedItem) & str
        frmBancoCuenta.lbldesc.Caption = CurrentInventory.GetInfo(CurrentInventory.OBJIndex(CurrentInventory.SelectedItem))
        
        Set CurrentInventory = Nothing

    End If

    ' Presentamos la escena
    Call Engine_EndScene(InvRect, frmBancoCuenta.interface.hwnd)
    Call Engine_EndScene(InvRect, frmBancoCuenta.interface.hwnd)

    
    Exit Sub

DrawInterfaceBoveda_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.DrawInterfaceBoveda", Erl)
    Resume Next
    
End Sub
Public Sub DrawInterfaceBoveda()
    
    On Error GoTo DrawInterfaceBoveda_Err
    

    ' Sólo dibujamos cuando es necesario
    If Not frmBancoObj.InvBoveda.NeedsRedraw And Not frmBancoObj.InvBankUsu.NeedsRedraw Then Exit Sub

    Dim InvRect As RECT

    InvRect.Left = 0
    InvRect.Top = 0
    InvRect.Right = frmBancoObj.interface.ScaleWidth
    InvRect.Bottom = frmBancoObj.interface.ScaleHeight

    ' Comenzamos la escena
    Call Engine_BeginScene

    ' Dibujamos el fondo de la bóveda
    Call Draw_GrhIndex(838, 0, 0)

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
    Call Engine_EndScene(InvRect, frmBancoObj.interface.hwnd)

    
    Exit Sub

DrawInterfaceBoveda_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.DrawInterfaceBoveda", Erl)
    Resume Next
    
End Sub
Public Sub DrawInterfaceKeys()
    
    On Error GoTo DrawInterfaceKeys_Err
    

    ' Sólo dibujamos cuando es necesario
    If Not FrmKeyInv.InvKeys.NeedsRedraw Then Exit Sub

    Dim InvRect As RECT

    InvRect.Left = 0
    InvRect.Top = 0
    InvRect.Right = FrmKeyInv.interface.ScaleWidth
    InvRect.Bottom = FrmKeyInv.interface.ScaleHeight

    ' Comenzamos la escena
    Call Engine_BeginScene

    ' Dibujamos el fondo de la bóveda
    'Call Draw_GrhIndex(838, 0, 0)
    
    ' Dibujamos llaves
    Call FrmKeyInv.InvKeys.DrawInventory

    ' Presentamos la escena
    Call Engine_EndScene(InvRect, FrmKeyInv.interface.hwnd)

    
    Exit Sub

DrawInterfaceKeys_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.DrawInterfaceKeys", Erl)
    Resume Next
    
End Sub

Public Sub DrawInventoryComercio()
    
    On Error GoTo DrawInventorysComercio_Err
    

    ' Sólo dibujamos cuando es necesario
    If Not frmComerciarUsu.InvUser.NeedsRedraw Then Exit Sub

    Dim InvRect As RECT

    InvRect.Left = 0
    InvRect.Top = 0
    InvRect.Right = frmComerciarUsu.picInv.ScaleWidth
    InvRect.Bottom = frmComerciarUsu.picInv.ScaleHeight

    ' Comenzamos la escena
    Call Engine_BeginScene
    
    ' Dibujamos llaves
    Call frmComerciarUsu.InvUser.DrawInventory
    
    ' Presentamos la escena
    Call Engine_EndScene(InvRect, frmComerciarUsu.picInv.hwnd)

    
    Exit Sub

DrawInventorysComercio_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.DrawInventorysComercio", Erl)
    Resume Next
    
End Sub


Public Sub DrawInventoryUserComercio()
    
    On Error GoTo DrawInventoryUserComercio_Err
    

    ' Sólo dibujamos cuando es necesario
    If Not frmComerciarUsu.InvUserSell.NeedsRedraw Then Exit Sub

    Dim InvRect As RECT

    InvRect.Left = 0
    InvRect.Top = 0
    InvRect.Right = frmComerciarUsu.picInvUserSell.ScaleWidth
    InvRect.Bottom = frmComerciarUsu.picInvUserSell.ScaleHeight

    ' Comenzamos la escena
    Call Engine_BeginScene
    
    ' Dibujamos llaves
    Call frmComerciarUsu.InvUserSell.DrawInventory
    
    ' Presentamos la escena
    Call Engine_EndScene(InvRect, frmComerciarUsu.picInvUserSell.hwnd)

    
    Exit Sub

DrawInventoryUserComercio_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.DrawInventoryUserComercio", Erl)
    Resume Next
    
End Sub


Public Sub DrawInventoryOtherComercio()
    
    On Error GoTo DrawInventoryOtherComercio_Err
    

    ' Sólo dibujamos cuando es necesario
    If Not frmComerciarUsu.InvOtherSell.NeedsRedraw Then Exit Sub

    Dim InvRect As RECT

    InvRect.Left = 0
    InvRect.Top = 0
    InvRect.Right = frmComerciarUsu.picInvOtherSell.ScaleWidth
    InvRect.Bottom = frmComerciarUsu.picInvOtherSell.ScaleHeight

    ' Comenzamos la escena
    Call Engine_BeginScene
    
    ' Dibujamos llaves
    Call frmComerciarUsu.InvOtherSell.DrawInventory
    
    ' Presentamos la escena
    Call Engine_EndScene(InvRect, frmComerciarUsu.picInvOtherSell.hwnd)

    
    Exit Sub

DrawInventoryOtherComercio_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.DrawInventoryOtherComercio", Erl)
    Resume Next
    
End Sub

Public Sub DrawMapaMundo()
    
    On Error GoTo DrawMapaMundo_Err
    

    

    Static re As RECT
    re.Left = 0
    re.Top = 0
    re.Bottom = 89
    re.Right = 177
    
    frmMapaGrande.PlayerView.Height = 89
    frmMapaGrande.PlayerView.Width = 177
    frmMapaGrande.PlayerView.ScaleHeight = 89
    frmMapaGrande.PlayerView.ScaleWidth = 177
    
    If frmMapaGrande.ListView1.ListItems.count <= 0 Then Exit Sub
    
    Call Engine_BeginScene
        
    Dim i    As Byte

    Dim x    As Integer

    Dim y    As Integer
    
    Dim Head As grh, grh As grh
    Dim HeadID As Integer, BodyID As Integer
    HeadID = NpcData(frmMapaGrande.ListView1.SelectedItem.SubItems(2)).Head
    BodyID = NpcData(frmMapaGrande.ListView1.SelectedItem.SubItems(2)).Body
    
    Dim tmp           As String

    Dim temp_array As RGBA
    Call SetRGBA(temp_array, 7, 7, 7)

    Engine_Draw_Box x, y, 177, 89, temp_array 'Fondo del inventario
    
    If BodyID > 0 Then
        grh = BodyData(BodyID).Walk(3)
        x = (frmMapaGrande.PlayerView.ScaleWidth - GrhData(grh.GrhIndex).pixelWidth) \ 2
        y = (frmMapaGrande.PlayerView.ScaleHeight - GrhData(grh.GrhIndex).pixelHeight) \ 2
        Call Draw_Grh(grh, x, y, 0, 1, COLOR_WHITE, False, 0, 0, 0)
    End If

    If HeadID > 0 Then
        Head = HeadData(HeadID).Head(3)
        x = (frmMapaGrande.PlayerView.ScaleWidth - GrhData(Head.GrhIndex).pixelWidth) \ 2
        y = (frmMapaGrande.PlayerView.ScaleHeight - GrhData(Head.GrhIndex).pixelHeight) \ 2 + 8 + BodyData(NpcData(frmMapaGrande.ListView1.SelectedItem.SubItems(2)).Body).HeadOffset.y
        Call Draw_Grh(Head, x, y, 0, 1, COLOR_WHITE, False, 0, 0, 0)
    End If
    
    Call Engine_EndScene(re, frmMapaGrande.PlayerView.hwnd)

    
    Exit Sub

DrawMapaMundo_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.DrawMapaMundo", Erl)
    Resume Next
    
End Sub

Public Sub Grh_Render_Advance(ByRef grh As grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByVal Height As Integer, ByVal Width As Integer, ByRef rgb_list() As RGBA, Optional ByVal h_center As Boolean, Optional ByVal v_center As Boolean, Optional ByVal alpha_blend As Boolean = False)
    
    On Error GoTo Grh_Render_Advance_Err
    

    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
    'Last Modify Date: 11/19/2003
    'Similar to Grh_Render, but let´s you resize the Grh
    '**************************************************************
    Dim tile_width  As Integer

    Dim tile_height As Integer

    Dim grh_index   As Long
    
    If grh.GrhIndex = 0 Or grh.GrhIndex > MaxGrh Then Exit Sub
    
    Dim CurrentFrame As Integer
    CurrentFrame = 1

    If grh.Started > 0 Then
        Dim ElapsedFrames As Long
        ElapsedFrames = Fix((FrameTime - grh.Started) / grh.speed)

        If grh.Loops = INFINITE_LOOPS Or ElapsedFrames < GrhData(grh.GrhIndex).NumFrames * (grh.Loops + 1) Then
            CurrentFrame = ElapsedFrames Mod GrhData(grh.GrhIndex).NumFrames + 1

        Else
            grh.Started = 0
        End If

    End If

    'Figure out what frame to draw (always 1 if not animated)
    grh_index = GrhData(grh.GrhIndex).Frames(CurrentFrame)
    
    'Center Grh over X, Y pos
    If GrhData(grh_index).TileWidth <> 1 Then
        screen_x = screen_x - Int(GrhData(grh_index).TileWidth * (32 \ 2)) + 32 \ 2

    End If
    
    If GrhData(grh_index).TileHeight <> 1 Then
        screen_y = screen_y - Int(GrhData(grh_index).TileHeight * 32) + 32

    End If
    
    'Draw it to device
    'Device_Box_Textured_Render_Advance grh_index, screen_x, screen_y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, rgb_list, GrhData(grh_index).sX, GrhData(grh_index).sY, Width, Height, alpha_blend, grh.angle
    Call Batch_Textured_Box_Advance(screen_x, screen_y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, GrhData(grh_index).sX, GrhData(grh_index).sY, GrhData(grh_index).FileNum, Width, Height, rgb_list, alpha_blend, grh.Angle)

    
    Exit Sub

Grh_Render_Advance_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Grh_Render_Advance", Erl)
    Resume Next
    
End Sub

Public Sub Grh_Render(ByRef grh As grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByRef rgb_list() As RGBA, Optional ByVal h_centered As Boolean = True, Optional ByVal v_centered As Boolean = True, Optional ByVal alpha_blend As Boolean = False)
    
    On Error GoTo Grh_Render_Err
    

    '**************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 2/28/2003
    'Modified by Juan Martín Sotuyo Dodero
    'Added centering
    '**************************************************************
    Dim tile_width  As Integer

    Dim tile_height As Integer

    Dim grh_index   As Long
    
    If grh.GrhIndex = 0 Or grh.GrhIndex > MaxGrh Then Exit Sub
    
    Dim CurrentFrame As Integer
    CurrentFrame = 1

    If grh.Started > 0 Then
        Dim ElapsedFrames As Long
        ElapsedFrames = Fix((FrameTime - grh.Started) / grh.speed)

        If grh.Loops = INFINITE_LOOPS Or ElapsedFrames < GrhData(grh.GrhIndex).NumFrames * (grh.Loops + 1) Then
            CurrentFrame = ElapsedFrames Mod GrhData(grh.GrhIndex).NumFrames + 1

        Else
            grh.Started = 0
        End If

    End If

    'Figure out what frame to draw (always 1 if not animated)
    grh_index = GrhData(grh.GrhIndex).Frames(CurrentFrame)
    
    'Center Grh over X, Y pos
    If GrhData(grh_index).TileWidth <> 1 Then
        screen_x = screen_x - Int(GrhData(grh_index).TileWidth * (32 \ 2)) + 32 \ 2

    End If
    
    If GrhData(grh_index).TileHeight <> 1 Then
        screen_y = screen_y - Int(GrhData(grh_index).TileHeight * 32) + 32

    End If
    
    'Draw it to device
    'Device_Box_Textured_Render grh_index, screen_x, screen_y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, rgb_list(), GrhData(grh_index).sX, GrhData(grh_index).sY, alpha_blend, grh.angle
    Call Batch_Textured_Box(screen_x, screen_y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, GrhData(grh_index).sX, GrhData(grh_index).sY, GrhData(grh_index).FileNum, rgb_list, alpha_blend, grh.Angle)

    
    Exit Sub

Grh_Render_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Grh_Render", Erl)
    Resume Next
    
End Sub

Private Function Grh_Check(ByVal grh_index As Long) As Boolean
    
    On Error GoTo Grh_Check_Err
    

    '**************************************************************
    'Author: Aaron Perkins - Modified by Juan Martín Sotuyo Dodero
    'Last Modify Date: 1/04/2003
    '
    '**************************************************************
    'check grh_index
    If grh_index > 0 And grh_index <= MaxGrh Then
        Grh_Check = GrhData(grh_index).NumFrames

    End If

    
    Exit Function

Grh_Check_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Grh_Check", Erl)
    Resume Next
    
End Function

Function Engine_PixelPosX(ByVal x As Integer) As Integer
    '*****************************************************************
    'Converts a tile position to a screen position
    '*****************************************************************
    
    On Error GoTo Engine_PixelPosX_Err
    
    Engine_PixelPosX = (x - 1) * 32

    
    Exit Function

Engine_PixelPosX_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Engine_PixelPosX", Erl)
    Resume Next
    
End Function

Function Engine_PixelPosY(ByVal y As Integer) As Integer
    '*****************************************************************
    'Converts a tile position to a screen position
    '*****************************************************************
    
    On Error GoTo Engine_PixelPosY_Err
    
    Engine_PixelPosY = (y - 1) * 32

    
    Exit Function

Engine_PixelPosY_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Engine_PixelPosY", Erl)
    Resume Next
    
End Function

Function Engine_ElapsedTime() As Long
    
    On Error GoTo Engine_ElapsedTime_Err
    

    '**************************************************************
    'Gets the time that past since the last call
    '**************************************************************
    Dim Start_Time As Long

    Start_Time = FrameTime
    Engine_ElapsedTime = Start_Time - EndTime

    If Engine_ElapsedTime > 1000 Then Engine_ElapsedTime = 1000
    EndTime = Start_Time

    
    Exit Function

Engine_ElapsedTime_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Engine_ElapsedTime", Erl)
    Resume Next
    
End Function

Private Sub Renderizar_Aura(ByVal aura_index As String, ByVal x As Integer, ByVal y As Integer, ByVal map_x As Byte, ByVal map_y As Byte, Optional ByVal userIndex As Long = 0)
    
    On Error GoTo Renderizar_Aura_Err
    

    Dim rgb_list(3)      As RGBA

    Dim i                As Byte

    Dim Index            As Long

    Dim Color            As Long

    Dim aura_grh         As grh

    Dim giro             As Single

    Dim lado             As Byte

    Index = Val(ReadField(1, aura_index, Asc(":")))
    Color = Val(ReadField(2, aura_index, Asc(":")))
    giro = Val(ReadField(3, aura_index, Asc(":")))
    lado = Val(ReadField(4, aura_index, Asc(":")))

    'Debug.Print charlist(userindex).AuraAngle
    If giro > 0 And userIndex > 0 Then
        'If lado = 0 Then
        charlist(userIndex).AuraAngle = charlist(userIndex).AuraAngle + (timerTicksPerFrame * giro)
        'Else
        'charlist(userindex).AuraAngle = charlist(userindex).AuraAngle - (timerTicksPerFrame * giro)
        ' End If
    
        If charlist(userIndex).AuraAngle >= 360 Then charlist(userIndex).AuraAngle = 0

    End If

    Call Long_2_RGBAList(rgb_list, Color)

    'Convertimos el Aura en un GRH
    Call InitGrh(aura_grh, Index)

    'Y por ultimo renderizamos esta capa con Draw_Grh
    If giro > 0 And userIndex > 0 Then
        Call Draw_Grh(aura_grh, x, y + 30, 1, 0, rgb_list(), True, map_x, map_y, charlist(userIndex).AuraAngle)
    Else
        Call Draw_Grh(aura_grh, x, y + 30, 1, 0, rgb_list(), True, map_x, map_y, 0)

    End If
    
    
    Exit Sub

Renderizar_Aura_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Renderizar_Aura", Erl)
    Resume Next
    
End Sub

Public Sub RenderConnect(ByVal tilex As Integer, ByVal tiley As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
    
    On Error GoTo RenderConnect_Err
    

    Call Engine_BeginScene

     Select Case UserMap
        Case 1 ' ulla 45-43
            tilex = 45
            tiley = 43
        Case 34 ' nix 22-75
            tilex = 22
            tiley = 75
        Case 59 ' bander 49-43
            tilex = 49
            tiley = 43
        Case 151 ' Arghal 38-41
            tilex = 38
            tiley = 41
        Case 62 ' Lindos 63-40
            tilex = 63
            tiley = 40
        Case 195 ' Arkhein 64-32
            tilex = 64
            tiley = 32
        Case 112 ' Esperanza 50-45
            tilex = 50
            tiley = 45
        Case 354 ' Polo 78-66
            tilex = 78
            tiley = 66
    End Select
    
    
    Call RenderScreen(tilex, tiley, PixelOffsetX, PixelOffsetY, HalfConnectTileWidth, HalfConnectTileHeight)
        
    Dim DefaultColor(3) As Long

    Dim Color           As Long
    Dim ColorGM(3) As RGBA
    ColorGM(0) = RGBA_From_Comp(248, 107, 3)
    ColorGM(1) = ColorGM(0)
    ColorGM(2) = ColorGM(0)
    ColorGM(3) = ColorGM(0)
    intro = 1

    If intro = 1 Then
        Draw_Grh BodyData(773).Walk(3), 470 + 15, 366, 1, 0, COLOR_WHITE
        Draw_Grh HeadData(118).Head(3), 470 + 15, 327 + 2, 1, 0, COLOR_WHITE
            
        Draw_Grh CascoAnimData(13).Head(3), 470 + 15, 327, 1, 0, COLOR_WHITE
        Draw_Grh WeaponAnimData(6).WeaponWalk(3), 470 + 15, 366, 1, 0, COLOR_WHITE
        Engine_Text_Render "Gulfas Morgolock", 449, 400, ColorGM, 1
        Engine_Text_Render "<Creador del Mundo>", 438, 415, ColorGM, 1

        Engine_Text_Render_LetraChica "v" & App.Major & "." & App.Minor & " Build: " & App.Revision, 870, 740, COLOR_WHITE, 4, False

        Dim ItemName As String

        'itemname = "abcdfghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789¡!¿TEST?#$100%&/\()=-@^[]<>*+.,:; pálmas séso te píso sólo púto ý LÁL LÉ"
            
        ' itemname = "pálmas séso te píso sólo púto ý lÁ Élefante PÍSÓS PÚTO ÑOño"
        Engine_Text_Render_LetraChica ItemName, 100, 730, COLOR_WHITE, 4, False

        If ClickEnAsistente < 30 Then
          '  Call Particle_Group_Render(spell_particle, 500, 365)
        End If

    End If

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
    Dim TempGrh As grh, cc(3) As RGBA
    Call InitGrh(TempGrh, 1171)
    
    Call InitGrh(TempGrh, 1172)

    Call RGBAList(cc, 255, 255, 255, 220)

    Draw_Grh TempGrh, (frmConnect.ScaleWidth - GrhData(TempGrh.GrhIndex).pixelWidth) \ 2 + 6, 20, 0, 1, cc(), False

    'Logo nuevo
    'Marco
    Call InitGrh(TempGrh, 1169)

    Draw_Grh TempGrh, 0, 0, 0, 0, COLOR_WHITE, False

    #If DEBUGGING = 1 Then
        Engine_Text_Render "CLIENTE DEBUG", (frmConnect.ScaleWidth - Engine_Text_Width("CLIENTE DEBUG")) \ 2, 30, COLOR_WHITE
    #End If

    If FadeInAlpha > 0 Then
        Call Engine_Draw_Box(0, 0, frmConnect.ScaleWidth, frmConnect.ScaleHeight, RGBA_From_Comp(0, 0, 0, FadeInAlpha))
        FadeInAlpha = FadeInAlpha - 1
    End If

    ' Draw_Grh TempGrh, 480, 100, 1, 1, cc(), False
    Call Engine_EndScene(Render_Connect_Rect, frmConnect.render.hwnd)
    
    FrameTime = (timeGetTime() And &H7FFFFFFF)
    'FramesPerSecCounter = FramesPerSecCounter + 1
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
    
    Exit Sub

RenderConnect_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.RenderConnect", Erl)
    Resume Next
    
End Sub

Public Sub RenderCrearPJ(ByVal tilex As Integer, ByVal tiley As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
    
    On Error GoTo RenderCrearPJ_Err
    

    Call Engine_BeginScene

    Call RenderScreen(tilex, tiley, PixelOffsetX, PixelOffsetY, HalfConnectTileWidth, HalfConnectTileHeight)

    RenderUICrearPJ

    Dim TempGrh As grh
    Call InitGrh(TempGrh, 1171)

    Draw_Grh TempGrh, 494, 190, 1, 1, COLOR_WHITE, False
    'Logo viejo
    
    'Marco
    Call InitGrh(TempGrh, 1169)

    Draw_Grh TempGrh, 0, 0, 0, 0, COLOR_WHITE, False

    Call Engine_EndScene(Render_Connect_Rect, frmConnect.render.hwnd)

    FrameTime = (timeGetTime() And &H7FFFFFFF)
    FramesPerSecCounter = FramesPerSecCounter + 1
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed

    'RenderPjsCuenta

    
    Exit Sub

RenderCrearPJ_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.RenderCrearPJ", Erl)
    Resume Next
    
End Sub

Public Sub rendercuenta(ByVal tilex As Integer, ByVal tiley As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
    
    On Error GoTo rendercuenta_Err
    

    Call Engine_BeginScene

    FrameTime = (timeGetTime() And &H7FFFFFFF)
    FramesPerSecCounter = FramesPerSecCounter + 1
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed

    RenderPjsCuenta
    
    'Call Particle_Group_Render(ParticleLluviaDorada, 400, 0)

    Call Engine_EndScene(Render_Connect_Rect, frmConnect.render.hwnd)
    
    Exit Sub

    
    Exit Sub

rendercuenta_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.rendercuenta", Erl)
    Resume Next
    
End Sub

Public Sub RenderUICrearPJ()
    
    On Error GoTo RenderUICrearPJ_Err
    

    Dim TempGrh As grh
    
    Dim ColorGray(3) As RGBA
    Call RGBAList(ColorGray, 200, 200, 200)

    Call InitGrh(TempGrh, 727)
    
    Draw_Grh TempGrh, 475, 545, 1, 1, COLOR_WHITE, False

    Engine_Text_Render "Creacion de Personaje", 280, 125, ColorGray, 5, False

    Dim OffsetX As Integer
    Dim OffsetY As Integer

    Engine_Text_Render_LetraChica "Nombre ", 460, 205, COLOR_WHITE, 6, False

    OffsetX = 240
    OffsetY = 15
    Engine_Text_Render_LetraChica "Clase ", 345 + OffsetX, 240 + OffsetY, COLOR_WHITE, 6, False

    Engine_Draw_Box 317 + OffsetX, 260 + OffsetY, 95, 21, RGBA_From_Comp(1, 1, 1, 100)
    Engine_Text_Render "<", 300 + OffsetX, 260 + OffsetY, COLOR_WHITE, 1, False
        
    Engine_Text_Render ">", 418 + OffsetX, 261 + OffsetY, COLOR_WHITE, 1, False

    Engine_Text_Render frmCrearPersonaje.lstProfesion.List(frmCrearPersonaje.lstProfesion.ListIndex), 365 + OffsetX - Engine_Text_Width(frmCrearPersonaje.lstProfesion.List(frmCrearPersonaje.lstProfesion.ListIndex), True, 1) / 2, 262 + OffsetY, ColorGray, 1, True

    Engine_Text_Render_LetraChica "Raza ", 347 + OffsetX, 290 + OffsetY, COLOR_WHITE, 6, False
    Engine_Draw_Box 317 + OffsetX, 305 + OffsetY, 95, 21, RGBA_From_Comp(1, 1, 1, 100)

    'Engine_Text_Render "Humano", 470 - Engine_Text_Height("Humano", False), 304, DefaultColor, 1, False
    Engine_Text_Render frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.ListIndex), 360 + OffsetX - Engine_Text_Width(frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.ListIndex), True, 1) / 2, 308 + OffsetY, ColorGray, 1, True
    
    
    
    Engine_Text_Render "<", 300 + OffsetX, 305 + OffsetY, ColorGray, 1, False
    Engine_Text_Render ">", 418 + OffsetX, 305 + OffsetY, ColorGray, 1, False

    OffsetX = 5
    OffsetY = 5

    Engine_Text_Render_LetraChica "Genero ", 340 + OffsetX, 255, COLOR_WHITE, 6, False
    
    
    Engine_Draw_Box 317 + OffsetX, 275, 95, 21, RGBA_From_Comp(1, 1, 1, 100)

    Engine_Text_Render frmCrearPersonaje.lstGenero.List(frmCrearPersonaje.lstGenero.ListIndex), 360 + OffsetX - Engine_Text_Width(frmCrearPersonaje.lstGenero.List(frmCrearPersonaje.lstGenero.ListIndex), True, 1) / 2, 277, ColorGray, 1, True
    
    Engine_Text_Render "<", 300 + OffsetX, 275, ColorGray, 1, False
    
    Engine_Text_Render ">", 418 + OffsetX, 275, ColorGray, 1, False
    
    'NACIMIENTO
    

    OffsetY = 30
    Engine_Text_Render_LetraChica "Hogar ", 340 + OffsetX, 305, ColorGray, 6, False
    Engine_Draw_Box 317 + OffsetX, 320, 95, 21, RGBA_From_Comp(1, 1, 1, 100)
    
    Engine_Text_Render frmCrearPersonaje.lstHogar.List(frmCrearPersonaje.lstHogar.ListIndex), 360 + OffsetX - Engine_Text_Width(frmCrearPersonaje.lstHogar.List(frmCrearPersonaje.lstHogar.ListIndex), True, 1) / 2, 322, ColorGray, 1, True

    Engine_Text_Render "<", 300 + OffsetX, 320, ColorGray, 1, False
    
    Engine_Text_Render ">", 418 + OffsetX, 320, ColorGray, 1, False

    
    'NACIMIENTO
     Dim Offy As Long
     Offy = -38

     Dim OffX As Long
     OffX = 340
    
    'Atributos
    Engine_Text_Render_LetraChica "Atributos ", 235 + OffX, 385 + Offy, COLOR_WHITE, 6, True
    Engine_Draw_Box 175 + OffX, 405 + Offy, 185, 150, RGBA_From_Comp(0, 0, 0, 80)
  '  Engine_Draw_Box 610, 405, 220, 180, D3DColorARGB(120, 100, 100, 100)
    
    Engine_Text_Render_LetraChica "Fuerza ", 185 + OffX, 410 + Offy, COLOR_WHITE, 1, True
   ' Engine_Text_Render "<", 260, 410, DefaultColor, 1, True
   ' Engine_Text_Render ">", 310, 410, DefaultColor, 1, True
    Engine_Draw_Box 280 + OffX, 409 + Offy, 20, 20, RGBA_From_Comp(1, 1, 1, 100)
    Engine_Text_Render_LetraChica frmCrearPersonaje.lbFuerza.Caption, 282 + OffX, 413 + Offy, COLOR_WHITE, 1, True ' Atributo fuerza
    'Engine_Text_Render "+", 335, 410, DefaultColor, 1, True
    Engine_Draw_Box 317 + OffX, 409 + Offy, 25, 20, RGBA_From_Comp(1, 1, 1, 100)
    Engine_Text_Render_LetraChica frmCrearPersonaje.modfuerza.Caption, 320 + OffX, 413 + Offy, COLOR_WHITE, 1, True ' Bonificacion fuerza
    
    
    Engine_Text_Render "Agilidad ", 185 + OffX, 440 + Offy, COLOR_WHITE, 1, True
   ' Engine_Text_Render "<", 260, 440, DefaultColor, 1, True
   ' Engine_Text_Render ">", 310, 440, DefaultColor, 1, True
    Engine_Draw_Box 280 + OffX, 440 + Offy, 20, 20, RGBA_From_Comp(1, 1, 1, 100)
    Engine_Text_Render frmCrearPersonaje.lbAgilidad.Caption, 282 + OffX, 443 + Offy, COLOR_WHITE, 1, True ' Atributo Agilidad
   ' Engine_Text_Render "+", 335, 440, DefaultColor, 1, True
    Engine_Draw_Box 317 + OffX, 440 + Offy, 25, 20, RGBA_From_Comp(1, 1, 1, 100)
    Engine_Text_Render frmCrearPersonaje.modAgilidad.Caption, 320 + OffX, 443 + Offy, COLOR_WHITE, 1, True ' Bonificacion Agilidad
    
    
    
    Engine_Text_Render "Inteligencia ", 185 + OffX, 470 + Offy, COLOR_WHITE, 1, True
    'Engine_Text_Render "<", 260, 470, DefaultColor, 1, True
    'Engine_Text_Render ">", 310, 470, DefaultColor, 1, True
    Engine_Draw_Box 280 + OffX, 470 + Offy, 20, 20, RGBA_From_Comp(1, 1, 1, 100)
    Engine_Text_Render frmCrearPersonaje.lbInteligencia.Caption, 282 + OffX, 473 + Offy, COLOR_WHITE, 1, True ' Atributo Inteligencia
    'Engine_Text_Render "+", 335, 470, DefaultColor, 1, True
    Engine_Draw_Box 317 + OffX, 470 + Offy, 25, 20, RGBA_From_Comp(1, 1, 1, 100)
    Engine_Text_Render frmCrearPersonaje.modInteligencia.Caption, 320 + OffX, 473 + Offy, COLOR_WHITE, 1, True ' Bonificacion Inteligencia
    
    
    Engine_Text_Render "Constitución ", 185 + OffX, 500 + Offy, COLOR_WHITE, , True
    'Engine_Text_Render "<", 260, 500, DefaultColor, 1, True
   ' Engine_Text_Render ">", 310, 500, DefaultColor, 1, True
    Engine_Draw_Box 280 + OffX, 500 + Offy, 20, 20, RGBA_From_Comp(1, 1, 1, 100)
    Engine_Text_Render frmCrearPersonaje.lbConstitucion.Caption, 283 + OffX, 503 + Offy, COLOR_WHITE, 1, True ' Atributo Constitución
    '
   ' Engine_Text_Render "+", 335, 500, DefaultColor, 1, True
    Engine_Draw_Box 317 + OffX, 500 + Offy, 25, 20, RGBA_From_Comp(1, 1, 1, 100)
    Engine_Text_Render frmCrearPersonaje.modConstitucion.Caption, 320 + OffX, 503 + Offy, COLOR_WHITE, 1, True ' Bonificacion Constitución
    
    
    
        Engine_Text_Render "Carisma ", 185 + OffX, 530 + Offy, COLOR_WHITE, , True
    'Engine_Text_Render "<", 260, 500, DefaultColor, 1, True
   ' Engine_Text_Render ">", 310, 500, DefaultColor, 1, True
    Engine_Draw_Box 280 + OffX, 530 + Offy, 20, 20, RGBA_From_Comp(1, 1, 1, 100)
    Engine_Text_Render frmCrearPersonaje.lbCarisma.Caption, 283 + OffX, 533 + Offy, COLOR_WHITE, 1, True ' Atributo Carisma
    '
   ' Engine_Text_Render "+", 335, 500, DefaultColor, 1, True
    Engine_Draw_Box 317 + OffX, 530 + Offy, 25, 20, RGBA_From_Comp(1, 1, 1, 100)
    Engine_Text_Render frmCrearPersonaje.modCarisma.Caption, 320 + OffX, 533 + Offy, COLOR_WHITE, 1, True ' Bonificacion Carisma
    
    
      
    '
    'Engine_Draw_Box 290, 528, 20, 20, D3DColorARGB(120, 1, 150, 150)
    'Engine_Text_Render "Puntos disponibles", 175, 530, DefaultColor, 1, True '
    'Engine_Text_Render frmCrearPersonaje.lbLagaRulzz.Caption, 291, 530, DefaultColor, 1, True '
    'Cabeza
    'Engine_Draw_Box 425, 415, 140, 100, D3DColorARGB(120, 100, 100, 100)

   ' Engine_Text_Render "Selecciona el rostro que más te agrade.", 662, 260, DefaultColor, 1, True
    
    
    
    
    

     OffX = -340
     Offy = -100




    Dim OffAspectoX As Integer
    Dim OffAspectoY As Integer
    
    
    OffAspectoX = -5
    OffAspectoY = -40
     
     
         
     Engine_Draw_Box 280 + OffAspectoX, 407 + OffAspectoY, 185, 148, RGBA_From_Comp(0, 0, 0, 80)
     
     
     Engine_Text_Render_LetraChica "Aspecto", 345 + OffAspectoX, 385 + OffAspectoY, COLOR_WHITE, 6, False
     
     
     
    
   ' Engine_Draw_Box 345, 502, 12, 12, D3DColorARGB(120, 100, 0, 0)
    
    'Engine_Text_Render_LetraChica "Equipado", 360, 502, DefaultColor, 4, False
    
    

    
    

     
    ' CPHeading = 3

    If CPHead <> 0 And CPArma <> 0 Then
    
         
    Engine_Text_Render_LetraChica "Cabeza", 350 + OffAspectoX, 410 + OffAspectoY, COLOR_WHITE, 1, False
    Engine_Text_Render "<", 335 + OffAspectoX, 412 + OffAspectoY, COLOR_WHITE, 1, True
    Engine_Text_Render ">", 403 + OffAspectoX, 412 + OffAspectoY, COLOR_WHITE, 1, True
        
    Engine_Text_Render ">", 423 + OffAspectoX, 448 + OffAspectoY, COLOR_WHITE, 3, True
    Engine_Text_Render "<", 293 + OffAspectoX, 448 + OffAspectoY, COLOR_WHITE, 3, True
        
    
    
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

            
            If enanooff > 0 Then
                Draw_Grh BodyData(CPBodyE).Walk(CPHeading), 689 + OffX, 346 - Offy, 1, 0, COLOR_WHITE
            Else
                Draw_Grh BodyData(CPBody).Walk(CPHeading), 689 + OffX, 346 - Offy, 1, 0, COLOR_WHITE
            End If
            
            Draw_Grh HeadData(CPHead).Head(CPHeading), 689 + OffX, 346 - Offy + BodyData(CPBody).HeadOffset.y + enanooff, 1, 0, COLOR_WHITE
            
            'If CPEquipado Then
            'Draw_Grh CascoAnimData(CPGorro).Head(CPHeading), 700 + OffX, 366 - Offy + BodyData(CPBody).HeadOffset.y + enanooff, 1, 0, DefaultColor()
            'Draw_Grh WeaponAnimData(CPArma).WeaponWalk(CPHeading), 685 + 15 + OffX, 365 - Offy + enanooff, 1, 0, DefaultColor()
            'Call Renderizar_Aura(CPAura, 686 + 15 + offx, 360 - offy, 0, 0)
            'End If
            
            Dim Color(3) As RGBA
            
            Color(0) = RGBA_From_Comp(0, 128, 190)
            Color(1) = Color(0)
            Color(2) = Color(0)
            Color(3) = Color(0)
            Engine_Text_Render CPName, 365 - Engine_Text_Width(CPName, True) / 2, 478, Color, 1, True
    Else
        Engine_Text_Render "X", 360, 428, COLOR_WHITE, 3, True
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
    Draw_GrhIndex 1123, 655, 345

    
    Exit Sub

RenderUICrearPJ_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.RenderUICrearPJ", Erl)
    Resume Next
    
End Sub

Public Sub RenderPjsCuenta()
    
    On Error GoTo RenderPjsCuenta_Err
    

    ' Renderiza el menu para seleccionar las clases
        
    Dim i               As Long

    Dim x               As Integer

    Dim y               As Integer

    Dim notY            As Integer

    Dim Color           As RGBA

    Dim Texto           As String

    Texto = CuentaEmail

    'Render fondo
    
   
   
   'Draw_GrhIndex 1170, 0, 0
    
    Dim temp_array(3) As RGBA

    Dim sumax As Long

    sumax = 84
    
    Dim TempColor(3) As RGBA
    Dim grh As grh
    
    'Dibujo la escena debajo del mapa
    Call RenderScreen(RenderCuenta_PosX, RenderCuenta_PosY, 0, 0, HalfConnectTileWidth, HalfConnectTileHeight)
    
    If LastPJSeleccionado <> PJSeleccionado Then
        If AlphaRenderCuenta < MAX_ALPHA_RENDER_CUENTA Then
            AlphaRenderCuenta = AlphaRenderCuenta + 1
        Else
            LastPJSeleccionado = PJSeleccionado

            If PJSeleccionado <> 0 Then
                Call SwitchMap(Pjs(PJSeleccionado).Mapa)
                RenderCuenta_PosX = Pjs(PJSeleccionado).PosX
                RenderCuenta_PosY = Pjs(PJSeleccionado).PosY
            End If
        End If
    ElseIf PJSeleccionado <> 0 And AlphaRenderCuenta > 0 Then
        If Pjs(PJSeleccionado).Mapa <> 0 Then
            AlphaRenderCuenta = AlphaRenderCuenta - 1
        End If
    End If

    Call RGBAList(TempColor, 255, 255, 255, 170 + AlphaRenderCuenta)
    
    Call InitGrh(grh, 4531)
                        
    Call Draw_Grh(grh, 0, 0, 0, 0, TempColor, False, 0, 0, 0)

    'Dibujamos frente 3839
    Draw_GrhIndex 3839, 0, 0
      
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
                Draw_Grh BodyData(Pjs(i).Body).Walk(3), x + 15, y + 10, 1, 1, COLOR_WHITE

            End If

            If (Pjs(i).Head <> 0) Then
                'If Not nohead Then
                Draw_Grh HeadData(Pjs(i).Head).Head(3), x + 15, y - notY + BodyData(Pjs(i).Body).HeadOffset.y + 10, 1, 0, COLOR_WHITE

                ' End If
            End If
            
            If (Pjs(i).Casco <> 0) Then
                'If Not nohead Then
                Draw_Grh CascoAnimData(Pjs(i).Casco).Head(3), x + 15, y - notY + BodyData(Pjs(i).Body).HeadOffset.y + 10, 1, 0, COLOR_WHITE

                ' End If
            End If
            
            If (Pjs(i).Escudo <> 0) Then
                'If Not nohead Then
                Draw_Grh ShieldAnimData(Pjs(i).Escudo).ShieldWalk(3), x + 14, y - notY + 10, 1, 0, COLOR_WHITE

                ' End If
            End If
                        
            If (Pjs(i).Arma <> 0) Then
                'If Not nohead Then
                Draw_Grh WeaponAnimData(Pjs(i).Arma).WeaponWalk(3), x + 14, y - notY + 10, 1, 0, COLOR_WHITE

                ' End If
            End If
        
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

                Engine_Text_Render "Clase: " & ListaClases(Pjs(i).Clase), 511 - Engine_Text_Width("Clase:" & ListaClases(Pjs(i).Clase), True) / 2, Offy + 585 - Engine_Text_Height("Clase:" & ListaClases(Pjs(i).Clase), True), COLOR_WHITE, 1, True
                
                Engine_Text_Render "Nivel: " & Pjs(i).nivel, 511 - Engine_Text_Width("Nivel:" & Pjs(i).nivel, True) / 2, Offy + 600 - Engine_Text_Height("Nivel:" & Pjs(i).nivel, True), COLOR_WHITE, 1, True
                Engine_Text_Render CStr(Pjs(i).NameMapa), 511 - Engine_Text_Width(CStr(Pjs(i).NameMapa), True) / 2, Offy + 615 - Engine_Text_Height(CStr(Pjs(i).NameMapa), True), COLOR_WHITE, 1, True

            End If
            
        End If

    Next i

    Exit Sub

RenderPjsCuenta_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.RenderPjsCuenta", Erl)
    Resume Next
    
End Sub

Sub EfectoEnPantalla(ByVal Color As Long, ByVal Time As Long)
    
    On Error GoTo EfectoEnPantalla_Err
    
    frmMain.Efecto.Interval = Time
    frmMain.Efecto.Enabled = True
    EfectoEnproceso = True
    Call SetGlobalLight(Color)

    
    Exit Sub

EfectoEnPantalla_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.EfectoEnPantalla", Erl)
    Resume Next
    
End Sub

Public Sub SetBarFx(ByVal charindex As Integer, ByVal BarTime As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 12/03/04
    'Sets an FX to the character.
    '***************************************************
    
    On Error GoTo SetBarFx_Err
    

    With charlist(charindex)
        .BarTime = BarTime

    End With

    
    Exit Sub

SetBarFx_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.SetBarFx", Erl)
    Resume Next
    
End Sub

Public Function Engine_Get_2_Points_Angle(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Double
    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 18/10/2012
    '**************************************************************
    
    On Error GoTo Engine_Get_2_Points_Angle_Err
    

    Engine_Get_2_Points_Angle = Engine_Get_X_Y_Angle((x2 - x1), (y2 - y1))
   
    
    Exit Function

Engine_Get_2_Points_Angle_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Engine_Get_2_Points_Angle", Erl)
    Resume Next
    
End Function

Public Function Engine_Get_X_Y_Angle(ByVal x As Double, ByVal y As Double) As Double
    '**************************************************************
    'Author: Unknown
    'Last Modify Date: 18/10/2012
    '**************************************************************
    
    On Error GoTo Engine_Get_X_Y_Angle_Err
    

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
   
    
    Exit Function

Engine_Get_X_Y_Angle_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Engine_Get_X_Y_Angle", Erl)
    Resume Next
    
End Function

Public Function Engine_Convert_Radians_To_Degrees(ByVal s_radians As Double) As Integer
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 8/25/2004
    'Converts a radian to degrees
    '**************************************************************
    
    On Error GoTo Engine_Convert_Radians_To_Degrees_Err
    

    Engine_Convert_Radians_To_Degrees = (s_radians * 180) / 3.14159265358979
 
    
    Exit Function

Engine_Convert_Radians_To_Degrees_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Engine_Convert_Radians_To_Degrees", Erl)
    Resume Next
    
End Function

' programado por maTih.-
 
Public Sub InitializeInventory()
    '
    ' @ Inicializa el array de efectos.
    
    On Error GoTo Initialize_Err
    
     
    ReDim Effect(1 To 255) As Effect_Type
    
    ' Inicializo inventarios
    Set frmMain.Inventario = New clsGrapchicalInventory
    Set frmComerciar.InvComUsu = New clsGrapchicalInventory
    Set frmComerciar.InvComNpc = New clsGrapchicalInventory
    Set frmBancoObj.InvBankUsu = New clsGrapchicalInventory
    Set frmBancoObj.InvBoveda = New clsGrapchicalInventory
    Set frmComerciarUsu.InvUser = New clsGrapchicalInventory
    Set frmComerciarUsu.InvUserSell = New clsGrapchicalInventory
    Set frmComerciarUsu.InvOtherSell = New clsGrapchicalInventory
    
    Set frmBancoCuenta.InvBankUsuCuenta = New clsGrapchicalInventory
    Set frmBancoCuenta.InvBovedaCuenta = New clsGrapchicalInventory
    
    Set FrmKeyInv.InvKeys = New clsGrapchicalInventory
    
    Call frmMain.Inventario.Initialize(frmMain.picInv, MAX_INVENTORY_SLOTS, , , 0, 0, 3, 3, True, 9)
    Call frmComerciar.InvComUsu.Initialize(frmComerciar.interface, MAX_INVENTORY_SLOTS, 210, 0, 252, 0, 3, 3, True)
    Call frmComerciar.InvComNpc.Initialize(frmComerciar.interface, MAX_INVENTORY_SLOTS, 210, , 1, 0, 3, 3)
    Call frmComerciarUsu.InvUser.Initialize(frmComerciarUsu.picInv, MAX_INVENTORY_SLOTS, , , 0, 0, 3, 3, True)
    Call frmComerciarUsu.InvUserSell.Initialize(frmComerciarUsu.picInvUserSell, 6, , , 0, 0, 3, 3, True)
    Call frmComerciarUsu.InvOtherSell.Initialize(frmComerciarUsu.picInvOtherSell, 6, , , 0, 0, 3, 3, True)
   
    Call frmBancoObj.InvBankUsu.Initialize(frmBancoObj.interface, MAX_INVENTORY_SLOTS, 210, 0, 252, 0, 3, 3, True)
    Call frmBancoObj.InvBoveda.Initialize(frmBancoObj.interface, MAX_BANCOINVENTORY_SLOTS, 210, 0, 0, 0, 3, 3)
    
    Call frmBancoCuenta.InvBankUsuCuenta.Initialize(frmBancoCuenta.interface, MAX_INVENTORY_SLOTS, 210, 0, 252, 0, 3, 3, True)
    Call frmBancoCuenta.InvBovedaCuenta.Initialize(frmBancoCuenta.interface, MAX_BANCOINVENTORY_SLOTS, 210, 0, 0, 0, 3, 3)
    
    
    'Ladder
    Call FrmKeyInv.InvKeys.Initialize(FrmKeyInv.interface, MAX_KEYS, , , 0, 0, 3, 3, True) 'Inventario de llaves
    FrmKeyInv.InvKeys.MostrarCantidades = False
 
    
    Exit Sub

Initialize_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Initialize", Erl)
    Resume Next
    
End Sub

Public Sub Terminate_Index(ByVal effect_Index As Integer)
    
    On Error GoTo Terminate_Index_Err
    
 
    '
    ' @ Destruye un indice del array
 
    Dim clear_Index As Effect_Type
 
    'Si es un slot válido
    If (effect_Index <> 0) And (effect_Index <= UBound(Effect())) Then
        Effect(effect_Index) = clear_Index

    End If
 
    
    Exit Sub

Terminate_Index_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Terminate_Index", Erl)
    Resume Next
    
End Sub
 
Public Function Effect_Begin(ByVal Fx_Index As Integer, ByVal Bind_Speed As Single, ByVal x As Single, ByVal y As Single, Optional ByVal explosion_FX_Index As Integer = -1, Optional ByVal explosion_FX_Loops As Integer = -1, Optional ByVal receptor As Integer = 1, Optional ByVal Emisor As Integer = 1, Optional ByVal wav As Integer = 1, Optional ByVal fX As Integer = -1) As Integer
    
    On Error GoTo Effect_Begin_Err
    
 
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
 
    
    Exit Function

Effect_Begin_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Effect_Begin", Erl)
    Resume Next
    
End Function

Public Function Effect_BeginXY(ByVal Fx_Index As Integer, ByVal Bind_Speed As Single, ByVal x As Single, ByVal y As Single, ByVal DestinoX As Byte, ByVal Destinoy As Byte, Optional ByVal explosion_FX_Index As Integer = -1, Optional ByVal explosion_FX_Loops As Integer = -1, Optional ByVal Emisor As Integer = 1, Optional ByVal wav As Integer = 1, Optional ByVal fX As Integer = 0) As Integer
    '
    ' @ Inicia un nuevo efecto y devuelve el index.
    
    On Error GoTo Effect_BeginXY_Err
    
 
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
 
    
    Exit Function

Effect_BeginXY_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Effect_BeginXY", Erl)
    Resume Next
    
End Function
 
Public Sub Effect_Render_All()
    
    On Error GoTo Effect_Render_All_Err
    
 
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
 
    
    Exit Sub

Effect_Render_All_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Effect_Render_All", Erl)
    Resume Next
    
End Sub
 
Public Sub Effect_Render_Slot(ByVal effect_Index As Integer)
    
    On Error GoTo Effect_Render_Slot_Err
    
 
    '
    ' @ Renderiza un efecto.

    Dim colornpcs(3) As Long

    colornpcs(0) = D3DColorXRGB(255, 255, 255)
    colornpcs(1) = colornpcs(0)
    colornpcs(2) = colornpcs(0)
    colornpcs(3) = colornpcs(0)
    
    With Effect(effect_Index)
 
        Dim target_Angle As Single
     
        .Now_Moved = (timeGetTime() And &H7FFFFFFF)
     
        'Controla el intervalo de vuelo
        If (.Last_Move + 10) < .Now_Moved Then
            .Last_Move = (timeGetTime() And &H7FFFFFFF)
        
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
              
                    Call SetMapFx(.DestX, .DesyY, .FxEnd_Effect, 0)
                    .Slot_Used = False
                    Exit Sub

                End If
          
            End If
        
        End If

    End With
 
    
    Exit Sub

Effect_Render_Slot_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Effect_Render_Slot", Erl)
    Resume Next
    
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
    
    On Error GoTo GetFreeIndex_Err
    
 
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
 
    
    Exit Function

GetFreeIndex_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.GetFreeIndex", Erl)
    Resume Next
    
End Function

Public Sub Draw_Grh_Precalculated(ByRef grh As grh, ByRef rgb_list() As RGBA, ByVal EsAgua As Boolean, ByVal EsLava As Boolean, ByVal MapX As Integer, ByVal MapY As Integer, ByVal MinX As Integer, ByVal MaxX As Integer, ByVal MinY As Integer, ByVal MaxY As Integer)
    
    On Error GoTo Draw_Grh_Precalculated_Err
    

    

    If grh.GrhIndex = 0 Or grh.GrhIndex > MaxGrh Then Exit Sub
    
    Dim CurrentFrame As Integer
    CurrentFrame = 1

    If grh.Started > 0 Then
        Dim ElapsedFrames As Long
        ElapsedFrames = Fix((FrameTime - grh.Started) / grh.speed)

        If grh.Loops = INFINITE_LOOPS Or ElapsedFrames < GrhData(grh.GrhIndex).NumFrames * (grh.Loops + 1) Then
            CurrentFrame = ElapsedFrames Mod GrhData(grh.GrhIndex).NumFrames + 1

        Else
            grh.Started = 0
        End If

    End If
    
    Dim CurrentGrhIndex As Long
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(grh.GrhIndex).Frames(CurrentFrame)

    Dim Texture As Direct3DTexture8

    Dim TextureWidth As Long, TextureHeight As Long
    Set Texture = SurfaceDB.GetTexture(GrhData(CurrentGrhIndex).FileNum, TextureWidth, TextureHeight)
    
    With GrhData(CurrentGrhIndex)

        Call SpriteBatch.SetTexture(Texture)
        
        If .Tx2 = 0 And .FileNum > 0 Then
            .Tx1 = .sX / TextureWidth
            .Tx2 = (.sX + .pixelWidth) / TextureWidth
            .Ty1 = .sY / TextureHeight
            .Ty2 = (.sY + .pixelHeight) / TextureHeight
        End If
        
        Dim Top As Byte, Right As Byte, Bottom As Byte, Left As Byte
    
        If EsAgua Then
            If MapY > MinY Then Top = (MapData(MapX, MapY - 1).Blocked And FLAG_AGUA) * INV_FLAG_AGUA
            If MapX < MaxX Then Right = (MapData(MapX + 1, MapY).Blocked And FLAG_AGUA) * INV_FLAG_AGUA
            If MapY < MaxY Then Bottom = (MapData(MapX, MapY + 1).Blocked And FLAG_AGUA) * INV_FLAG_AGUA
            If MapX > MinX Then Left = (MapData(MapX - 1, MapY).Blocked And FLAG_AGUA) * INV_FLAG_AGUA
        
            Call SpriteBatch.DrawWater(grh.x, grh.y, TilePixelWidth, TilePixelHeight, rgb_list, .Tx1, .Ty1, .Tx2, .Ty2, MapX, MapY, Top, Right, Bottom, Left)
        
        ElseIf EsLava Then
            If MapY > MinY Then Top = (MapData(MapX, MapY - 1).Blocked And FLAG_LAVA) * INV_FLAG_LAVA
            If MapX < MaxX Then Right = (MapData(MapX + 1, MapY).Blocked And FLAG_LAVA) * INV_FLAG_LAVA
            If MapY < MaxY Then Bottom = (MapData(MapX, MapY + 1).Blocked And FLAG_LAVA) * INV_FLAG_LAVA
            If MapX > MinX Then Left = (MapData(MapX - 1, MapY).Blocked And FLAG_LAVA) * INV_FLAG_LAVA
        
            Call SpriteBatch.DrawLava(grh.x, grh.y, TilePixelWidth, TilePixelHeight, rgb_list, .Tx1, .Ty1, .Tx2, .Ty2, MapX, MapY, Top, Right, Bottom, Left)
        
        Else
            Call SpriteBatch.Draw(grh.x, grh.y, TilePixelWidth, TilePixelHeight, rgb_list, .Tx1, .Ty1, .Tx2, .Ty2)
        End If
    
    End With

    
    Exit Sub

Draw_Grh_Precalculated_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Draw_Grh_Precalculated", Erl)
    Resume Next
    
End Sub

Public Sub Engine_Draw_Box(ByVal x As Integer, ByVal y As Integer, ByVal Width As Integer, ByVal Height As Integer, Color As RGBA)
    
    On Error GoTo Engine_Draw_Box_Err
    

    Call RGBAList(temp_rgb, Color.r, Color.G, Color.B, Color.A)

    Call SpriteBatch.SetTexture(Nothing)
    Call SpriteBatch.SetAlpha(False)
    Call SpriteBatch.Draw(x, y, Width, Height, temp_rgb())

    
    Exit Sub

Engine_Draw_Box_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Engine_Draw_Box", Erl)
    Resume Next
    
End Sub

Public Sub Engine_Draw_Box_Border(ByVal x As Integer, ByVal y As Integer, ByVal Width As Integer, ByVal Height As Integer, Color As RGBA, ColorLine As RGBA)
    
    On Error GoTo Engine_Draw_Box_Border_Err
    

    Call Engine_Draw_Box(x, y, Width, Height, Color)

    Call Engine_Draw_Box(x, y, Width, 1, ColorLine)
    Call Engine_Draw_Box(x, y + Height, Width, 1, ColorLine)
    Call Engine_Draw_Box(x, y, 1, Height, ColorLine)
    Call Engine_Draw_Box(x + Width, y, 1, Height, ColorLine)

    
    Exit Sub

Engine_Draw_Box_Border_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Engine_Draw_Box_Border", Erl)
    Resume Next
    
End Sub

Public Sub DibujarBody(PicBox As PictureBox, ByVal MyBody As Integer, Optional ByVal Heading As Byte = 3)
    
    On Error GoTo DibujarBody_Err

    Dim grh As grh

    grh = BodyData(NpcData(MyBody).Body).Walk(3)

    Dim x    As Long

    Dim y    As Long

    Dim grhH As grh

    grhH = HeadData(NpcData(MyBody).Head).Head(3)

    x = (PicBox.ScaleWidth - GrhData(grh.GrhIndex).pixelWidth) / 2
    y = max((PicBox.ScaleHeight - GrhData(grh.GrhIndex).pixelHeight) / 2, BodyData(NpcData(MyBody).Body).HeadOffset.y)
     Call Grh_Render_To_Hdc(PicBox, GrhData(grh.GrhIndex).Frames(1), x, y, False, RGB(11, 11, 11))
    

    If NpcData(MyBody).Head <> 0 Then
        x = (PicBox.ScaleWidth - GrhData(grhH.GrhIndex).pixelWidth) / 2
        y = y + 8 + BodyData(NpcData(MyBody).Body).HeadOffset.y
        PicBox.BackColor = RGB(11, 11, 11)
        Call Grh_Render_To_HdcSinBorrar(PicBox, GrhData(grhH.GrhIndex).Frames(1), x, y, False)
    End If

    
    Exit Sub

DibujarBody_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmQuestInfo.DibujarBody", Erl)
    Resume Next
    
End Sub
