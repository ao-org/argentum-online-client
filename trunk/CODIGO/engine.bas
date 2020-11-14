Attribute VB_Name = "engine"
Option Explicit


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
    color       As Long
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

Public Const ScreenWidth  As Long = 538

Public Const ScreenHeight As Long = 376

Public bRunning            As Boolean

Private Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

Private Const FVF2 = D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX2


Dim texture      As Direct3DTexture8

Dim TransTexture As Direct3DTexture8

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Long

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public fps                     As Long

Private FramesPerSecCounter    As Long

Private lFrameTimer            As Long

Private lFrameLimiter          As Long

Private ScrollPixelsPerFrameX  As Single

Private ScrollPixelsPerFrameY  As Single

Private TileBufferPixelOffsetX As Integer

Private TileBufferPixelOffsetY As Integer

Private Const GrhFogata        As Integer = 1521

Private Estrella  As grh

Private Marco     As grh

Private BarraMana As grh

Private BarraVida As grh

Private BarraGris As grh

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

Private Function Init_DirectDevice(ByVal ModoAceleracion As CONST_D3DCREATEFLAGS) As Boolean
On Error GoTo ErrorHandler:

    Dim DispMode    As D3DDISPLAYMODE
    Dim D3DWindow   As D3DPRESENT_PARAMETERS
    
    Dim VSync As String: VSync = CByte(GetVar(App.Path & "\..\Recursos\OUTPUT\Configuracion.ini", "VIDEO", "VSync"))
    
    Set dX = New DirectX8
    Set D3D = dX.Direct3DCreate()
    Set D3DX = New D3DX8
    
    Call D3D.GetAdapterDisplayMode(D3DADAPTER_DEFAULT, DispMode)
    
    With D3DWindow
    
        .Windowed = True

        If VSync Then
            .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
        Else
            .SwapEffect = D3DSWAPEFFECT_DISCARD
        End If
        
        .BackBufferFormat = D3DFMT_X8R8G8B8
        
        'set color depth
        .BackBufferWidth = 1024
        .BackBufferHeight = 768
        
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
        .hDeviceWindow = frmmain.renderer.hwnd
        
    End With
    
    If Not D3DDevice Is Nothing Then Set D3DDevice = Nothing
    
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, D3DWindow.hDeviceWindow, ModoAceleracion, D3DWindow)
    
    Init_DirectDevice = True
    
    Exit Function
    
ErrorHandler:
    
    Set D3DDevice = Nothing
    
    Init_DirectDevice = False

End Function

Private Sub Engine_InitExtras()
    
    Call Engine_Font_Initialize
    
    Set LucesRedondas = New clsLucesRedondas
    Set LucesCuadradas = New clsLucesCuadradas
    Set Meteo_Engine = New clsMeteorologic
    
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
    
    Call Font_Create("Tahoma", 8, True, 0)
    Call Font_Create("Verdana", 8, False, 0)
    Call Font_Create("Verdana", 11, True, False)
    
    ' Colores comunes
    COLOR_WHITE(0) = D3DColorXRGB(255, 255, 255)
    COLOR_WHITE(1) = D3DColorXRGB(255, 255, 255)
    COLOR_WHITE(2) = D3DColorXRGB(255, 255, 255)
    COLOR_WHITE(3) = D3DColorXRGB(255, 255, 255)
    
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
    
End Sub

Public Sub Engine_Init()

    '*****************************************************
    '****** Coded by Menduz (lord.yo.wo@gmail.com) *******
    '*****************************************************
    On Error GoTo errhandler:
    
    Dim Modo As String: Modo = GetVar(App.Path & "\..\Recursos\OUTPUT\Configuracion.ini", "VIDEO", "Aceleracion")
    
    Select Case Modo
    
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
    
    With D3DDevice
    
        .SetVertexShader FVF
    
        '//Transformed and lit vertices dont need lighting
        '   so we disable it...
        .SetRenderState D3DRS_LIGHTING, False
        
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        
        .SetRenderState D3DRS_POINTSIZE, Engine_FToDW(2)
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    
    End With
    
    ' Carga de texturas
    Set SurfaceDB = New clsTexManager
    Call SurfaceDB.Init(D3DX, D3DDevice, General_Get_Free_Ram_Bytes)

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
    
    Debug.Print "Error Number Returned: " & Err.number

    End

End Sub

Public Sub Engine_BeginScene()
    
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 1#, 1#, 0)
    Call D3DDevice.BeginScene

End Sub

Public Sub Engine_EndScene(ByRef DestRect As RECT, Optional ByVal hwnd As Long = 0)

On Error GoTo ErrorHandler:
    
    Call D3DDevice.EndScene
    Call D3DDevice.Present(DestRect, ByVal 0, hwnd, ByVal 0)
    
    Exit Sub
    
ErrorHandler:

    If D3DDevice.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        
        Call Engine_Init
        
        prgRun = True
        pausa = False
        QueRender = 0

    End If
        
End Sub

Public Sub Engine_Deinit()
    
    Erase MapData
    Erase charlist
    
    Set D3DDevice = Nothing
    Set D3D = Nothing
    Set dX = Nothing

End Sub

Public Sub Engine_ActFPS()

    If (GetTickCount() And &H7FFFFFFF) - lFrameTimer > 1000 Then
        fps = FramesPerSecCounter
        FramesPerSecCounter = 0
        lFrameTimer = GetTickCount

    End If

End Sub

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

Public Sub Draw_Grh(ByRef grh As grh, ByVal x As Integer, ByVal y As Integer, ByVal center As Byte, ByVal animate As Byte, ByRef rgb_list() As Long, Optional ByVal Alpha As Boolean, Optional ByVal map_x As Byte = 1, Optional ByVal map_y As Byte = 1, Optional ByVal angle As Single)

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

Public Sub render()

    '*****************************************************
    '****** Coded by Menduz (lord.yo.wo@gmail.com) *******
    '*****************************************************
    Rem On Error GoTo ErrorHandler:
    Dim temp_array(3) As Long
    
    If Map_light_base = -1 And Not EfectoEnproceso Then
        Meteo_Engine.Meteo_Logic
    ElseIf UserEstado = 1 Then
        Meteo_Engine.Meteo_Logic
       
    End If
    
    Call Engine_BeginScene
    
    Call ShowNextFrame

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
    
    Call Engine_EndScene(Render_Main_Rect)
    
    lFrameLimiter = (GetTickCount() And &H7FFFFFFF)
    FramesPerSecCounter = FramesPerSecCounter + 1
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed

    Exit Sub

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
            If Engine_Meteo_Particle_Get <> 0 Then
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
                    
                    If MapData(x, y).CharFantasma.AlphaB > 0 Then
                        MapData(x, y).CharFantasma.AlphaB = MapData(x, y).CharFantasma.AlphaB - (timerTicksPerFrame * 30)
                        If MapData(x, y).CharFantasma.AlphaB < 0 Then MapData(x, y).CharFantasma.AlphaB = 0

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
        
            If Graficos_Particulas.Engine_Meteo_Particle_Get <> 0 Then
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
                    Call Draw_GrhIndex(5259 + .simbolo, PixelOffsetX + 6, PixelOffsetY + .Body.HeadOffset.y - 10)
                                
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
                    render
                
                    Check_Keys
                    Moviendose = False
                    DrawMainInventory

                    If HayFormularioAbierto Then
                        If frmComerciar.Visible Then
                            DrawInterfaceComerciar
                    
                        ElseIf frmBancoObj.Visible Then
                            DrawInterfaceBoveda

                        End If

                        If frmMapaGrande.Visible Then
                            DrawMapaMundo

                        End If
                        
                        If FrmKeyInv.Visible Then
                            DrawInterfaceKeys
                        End If

                    End If

                Case 1
                    RenderConnect 48, 49, 0, 0

                    If Not frmConnect.Visible Then
                        frmConnect.Show
                        FrmLogear.Show , frmConnect
                        FrmLogear.Top = FrmLogear.Top + 3500

                    End If

                Case 2
                    rendercuenta 42, 43, 0, 0

                Case 3
                    RenderCrearPJ 13, 67, 0, 0

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

Public Sub SetMapFx(ByVal x As Byte, ByVal y As Byte, ByVal fX As Integer, ByVal Loops As Integer)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modify Date: 12/03/04
    'Sets an FX to the character.
    '***************************************************

    On Error Resume Next

    Dim indice As Byte

    With MapData(x, y)
    
        indice = Map_FX_Group_Next_Open(x, y)
    
        .FxList(indice).FxIndex = fX
        Call InitGrh(.FxList(indice), FxData(fX).Animacion)
        .FxList(indice).Loops = Loops
    
    End With

End Sub

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
    Call Engine_BeginScene

    ' Dibujamos el fondo del inventario principal
    'Call Draw_GrhIndex(6, 0, 0)

    ' Dibujamos items
    Call frmmain.Inventario.DrawInventory
    
    ' Dibujamos item arrastrado
    Call frmmain.Inventario.DrawDraggedItem

    ' Presentamos la escena
    Call Engine_EndScene(InvRect, frmmain.picInv.hwnd)

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
        frmComerciar.lblCosto = PonerPuntos(CLng(CurrentInventory.Valor(CurrentInventory.SelectedItem) * cantidad))
        
        Set CurrentInventory = Nothing

    End If

    ' Presentamos la escena
    Call Engine_EndScene(InvRect, frmComerciar.interface.hwnd)

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

End Sub
Public Sub DrawInterfaceKeys()

    ' Sólo dibujamos cuando es necesario
    'If Not FrmKeyInv.InvKeys.NeedsRedraw Then Exit Sub

    Dim InvRect As RECT

    InvRect.Left = 0
    InvRect.Top = 0
    InvRect.Right = FrmKeyInv.interface.ScaleWidth
    InvRect.bottom = FrmKeyInv.interface.ScaleHeight

    ' Comenzamos la escena
    Call Engine_BeginScene

    ' Dibujamos el fondo de la bóveda
    'Call Draw_GrhIndex(838, 0, 0)
    
    
    Engine_Draw_Box 0, 0, 32, 32, D3DColorARGB(200, 255, 7, 7) 'pongo un color de fondo para chequear que dibuje
    Engine_Draw_Box 32, 0, 32, 32, D3DColorARGB(200, 0, 255, 7) 'pongo un color de fondo para chequear que dibuje
    Engine_Draw_Box 64, 0, 32, 32, D3DColorARGB(200, 0, 0, 255) 'pongo un color de fondo para chequear que dibuje
    Engine_Draw_Box 96, 0, 32, 32, D3DColorARGB(200, 255, 255, 0) 'pongo un color de fondo para chequear que dibuje
    Engine_Draw_Box 128, 0, 32, 32, D3DColorARGB(200, 0, 255, 255) 'pongo un color de fondo para chequear que dibuje
    
    
    
    Engine_Draw_Box 0, 32, 32, 32, D3DColorARGB(200, 0, 255, 255) 'pongo un color de fondo para chequear que dibuje
    Engine_Draw_Box 32, 32, 32, 32, D3DColorARGB(200, 255, 255, 7) 'pongo un color de fondo para chequear que dibuje
    Engine_Draw_Box 64, 32, 32, 32, D3DColorARGB(200, 255, 0, 0) 'pongo un color de fondo para chequear que dibuje
    Engine_Draw_Box 96, 32, 32, 32, D3DColorARGB(200, 0, 0, 250) 'pongo un color de fondo para chequear que dibuje
    Engine_Draw_Box 128, 32, 32, 32, D3DColorARGB(200, 0, 255, 0) 'pongo un color de fondo para chequear que dibuje
    
    
    ' Dibujamos items de la bóveda
    Call FrmKeyInv.InvKeys.DrawInventory
    
    ' Dibujamos items del usuario


    ' Dibujamos "ambos" items arrastrados (aunque sólo puede estar uno activo a la vez)
    Call FrmKeyInv.InvKeys.DrawDraggedItem
    
    ' Me fijo qué inventario está seleccionado
    Dim CurrentInventory As clsGrapchicalInventory
    
    If FrmKeyInv.InvKeys.SelectedItem > 0 Then
        Set CurrentInventory = FrmKeyInv.InvKeys
    ElseIf frmBancoObj.InvBankUsu.SelectedItem > 0 Then
        Set CurrentInventory = frmBancoObj.InvBankUsu

    End If
    

    ' Presentamos la escena
    Call Engine_EndScene(InvRect, FrmKeyInv.interface.hwnd)

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
    
    Call Engine_BeginScene
        
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
    
    Call Engine_EndScene(re, frmMapaGrande.PlayerView.hwnd)

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

Public Sub Grh_Render_Advance(ByRef grh As grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByVal Height As Integer, ByVal Width As Integer, ByRef rgb_list() As Long, Optional ByVal h_center As Boolean, Optional ByVal v_center As Boolean, Optional ByVal alpha_blend As Boolean = False)

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

Public Sub Grh_Render(ByRef grh As grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByRef rgb_list() As Long, Optional ByVal h_centered As Boolean = True, Optional ByVal v_centered As Boolean = True, Optional ByVal alpha_blend As Boolean = False)

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

Function Engine_ElapsedTime() As Long

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

Public Sub RenderConnect(ByVal tilex As Integer, ByVal tiley As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)

    Call Engine_BeginScene

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
        Engine_Text_Render_LetraChica ItemName, 100, 730, DefaultColor, 4, False

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
    Call Engine_EndScene(Render_Connect_Rect, frmConnect.render.hwnd)
    
    lFrameLimiter = (GetTickCount() And &H7FFFFFFF)
    'FramesPerSecCounter = FramesPerSecCounter + 1
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed

    Exit Sub

End Sub

Public Sub RenderCrearPJ(ByVal tilex As Integer, ByVal tiley As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)

    Call Engine_BeginScene

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

    RenderUICrearPJ

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

    Call Engine_EndScene(Render_Connect_Rect, frmConnect.render.hwnd)

    lFrameLimiter = (GetTickCount() And &H7FFFFFFF)
    FramesPerSecCounter = FramesPerSecCounter + 1
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed

    'RenderPjsCuenta

End Sub

Public Sub rendercuenta(ByVal tilex As Integer, ByVal tiley As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)

    Call Engine_BeginScene

    lFrameLimiter = (GetTickCount() And &H7FFFFFFF)
    FramesPerSecCounter = FramesPerSecCounter + 1
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed

    RenderPjsCuenta
    
    Call Particle_Group_Render(ParticleLluviaDorada, 400, 0)

    Call Engine_EndScene(Render_Connect_Rect, frmConnect.render.hwnd)
    
    Exit Sub

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
    Engine_Text_Render_LetraChica "Atributos", 240 + OffX, 385 + Offy, DefaultColor, 6, True
    Engine_Draw_Box 175 + OffX, 405 + Offy, 185, 120, D3DColorARGB(80, 0, 0, 0)
    '  Engine_Draw_Box 610, 405, 220, 180, D3DColorARGB(120, 100, 100, 100)
    
    Engine_Text_Render_LetraChica "Fuerza", 185 + OffX, 410 + Offy, DefaultColor, 1, True
    ' Engine_Text_Render "<", 260, 410, DefaultColor, 1, True
    ' Engine_Text_Render ">", 310, 410, DefaultColor, 1, True
    Engine_Draw_Box 280 + OffX, 407 + Offy, 20, 20, D3DColorARGB(100, 1, 1, 1)
    Engine_Text_Render frmCrearPersonaje.lbFuerza.Caption, 282 + OffX, 410 + Offy, DefaultColor, 1, True ' Atributo fuerza
    'Engine_Text_Render "+", 335, 410, DefaultColor, 1, True
    Engine_Draw_Box 317 + OffX, 407 + Offy, 25, 20, D3DColorARGB(100, 1, 1, 1)
    Engine_Text_Render frmCrearPersonaje.modfuerza.Caption, 320 + OffX, 410 + Offy, DefaultColor, 1, True ' Bonificacion fuerza
    
    Engine_Text_Render "Agilidad", 185 + OffX, 430 + Offy, DefaultColor, 1, True
    ' Engine_Text_Render "<", 260, 440, DefaultColor, 1, True
    ' Engine_Text_Render ">", 310, 440, DefaultColor, 1, True
    Engine_Draw_Box 280 + OffX, 427 + Offy, 20, 20, D3DColorARGB(100, 1, 1, 1)
    Engine_Text_Render frmCrearPersonaje.lbAgilidad.Caption, 282 + OffX, 430 + Offy, DefaultColor, 1, True ' Atributo Agilidad
    ' Engine_Text_Render "+", 335, 440, DefaultColor, 1, True
    Engine_Draw_Box 317 + OffX, 427 + Offy, 25, 20, D3DColorARGB(100, 1, 1, 1)
    Engine_Text_Render frmCrearPersonaje.modAgilidad.Caption, 320 + OffX, 430 + Offy, DefaultColor, 1, True ' Bonificacion Agilidad
    
    Engine_Text_Render "Inteligencia", 185 + OffX, 450 + Offy, DefaultColor, 1, True
    'Engine_Text_Render "<", 260, 470, DefaultColor, 1, True
    'Engine_Text_Render ">", 310, 470, DefaultColor, 1, True
    Engine_Draw_Box 280 + OffX, 447 + Offy, 20, 20, D3DColorARGB(100, 1, 1, 1)
    Engine_Text_Render frmCrearPersonaje.lbInteligencia.Caption, 282 + OffX, 450 + Offy, DefaultColor, 1, True ' Atributo Inteligencia
    'Engine_Text_Render "+", 335, 470, DefaultColor, 1, True
    Engine_Draw_Box 317 + OffX, 447 + Offy, 25, 20, D3DColorARGB(100, 1, 1, 1)
    Engine_Text_Render frmCrearPersonaje.modInteligencia.Caption, 320 + OffX, 450 + Offy, DefaultColor, 1, True ' Bonificacion Inteligencia
    
    Engine_Text_Render "Constitución", 185 + OffX, 470 + Offy, DefaultColor, , True
    'Engine_Text_Render "<", 260, 500, DefaultColor, 1, True
    ' Engine_Text_Render ">", 310, 500, DefaultColor, 1, True
    Engine_Draw_Box 280 + OffX, 467 + Offy, 20, 20, D3DColorARGB(100, 1, 1, 1)
    Engine_Text_Render frmCrearPersonaje.lbConstitucion.Caption, 283 + OffX, 470 + Offy, DefaultColor, 1, True ' Atributo Constitución
    '
    ' Engine_Text_Render "+", 335, 500, DefaultColor, 1, True
    Engine_Draw_Box 317 + OffX, 467 + Offy, 25, 20, D3DColorARGB(100, 1, 1, 1)
    Engine_Text_Render frmCrearPersonaje.modConstitucion.Caption, 320 + OffX, 470 + Offy, DefaultColor, 1, True ' Bonificacion Constitución
    
    Engine_Text_Render "Carisma", 185 + OffX, 490 + Offy, DefaultColor, , True
    'Engine_Text_Render "<", 260, 500, DefaultColor, 1, True
    ' Engine_Text_Render ">", 310, 500, DefaultColor, 1, True
    Engine_Draw_Box 280 + OffX, 487 + Offy, 20, 20, D3DColorARGB(100, 1, 1, 1)
    Engine_Text_Render frmCrearPersonaje.lbCarisma.Caption, 283 + OffX, 490 + Offy, DefaultColor, 1, True ' Atributo Carisma
    '
    ' Engine_Text_Render "+", 335, 500, DefaultColor, 1, True
    Engine_Draw_Box 317 + OffX, 487 + Offy, 25, 20, D3DColorARGB(100, 1, 1, 1)
    Engine_Text_Render frmCrearPersonaje.modCarisma.Caption, 320 + OffX, 490 + Offy, DefaultColor, 1, True ' Bonificacion Carisma
      
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
        Dim Genero As Byte

        If frmCrearPersonaje.lstRaza.ListIndex < 0 Then
            frmCrearPersonaje.lstRaza.ListIndex = 0

        End If

        Raza = frmCrearPersonaje.lstRaza.ListIndex
        
        
        
        If frmCrearPersonaje.lstGenero.ListIndex < 0 Then
            frmCrearPersonaje.lstGenero.ListIndex = 0

        End If

        Genero = frmCrearPersonaje.lstGenero.ListIndex

        Dim enanooff As Byte

        If Raza = 0 Or Raza = 1 Or Raza = 2 Or Raza = 5 Then
            enanooff = 0
    
        Else
            enanooff = 10

        End If
    
            
        If enanooff > 0 Then
            If Genero = 0 Then
                Draw_Grh BodyData(52).Walk(CPHeading), 685 + 15 + OffX, 366 - Offy, 1, 0, DefaultColor()
            Else
                Draw_Grh BodyData(52).Walk(CPHeading), 685 + 15 + OffX, 366 - Offy, 1, 0, DefaultColor()
            End If
        Else
            If Genero = 0 Then
                Draw_Grh BodyData(1).Walk(CPHeading), 685 + 15 + OffX, 366 - Offy, 1, 0, DefaultColor()
            Else
                 Draw_Grh BodyData(80).Walk(CPHeading), 685 + 15 + OffX, 366 - Offy, 1, 0, DefaultColor()
            End If

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
    Call Map_Base_Light_Set(color)

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

    Set FrmKeyInv.InvKeys = New clsGrapchicalInventory
    
    Call frmmain.Inventario.Initialize(frmmain.picInv, MAX_INVENTORY_SLOTS, , , 0, 0, 3, 3, True, 9)
    Call frmComerciar.InvComUsu.Initialize(frmComerciar.interface, MAX_INVENTORY_SLOTS, 210, 0, 252, 0, 3, 3, True)
    Call frmComerciar.InvComNpc.Initialize(frmComerciar.interface, MAX_INVENTORY_SLOTS, 210, , 1, 0, 3, 3)
   
    Call frmBancoObj.InvBankUsu.Initialize(frmBancoObj.interface, MAX_INVENTORY_SLOTS, 210, 0, 252, 0, 3, 3, True)
    Call frmBancoObj.InvBoveda.Initialize(frmBancoObj.interface, MAX_BANCOINVENTORY_SLOTS, 210, 0, 0, 0, 3, 3)
    
    
    'Ladder
    Call FrmKeyInv.InvKeys.Initialize(FrmKeyInv.interface, 10, 0, 0, 0, 0, 0, 0) 'Inventario de llaves
 
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
              
                    Call SetMapFx(.DestX, .DesyY, .FxEnd_Effect, 0)
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

Public Sub Draw_Grh_Picture(ByVal grh As Long, ByVal pic As PictureBox, ByVal x As Integer, ByVal y As Integer, ByVal Alpha As Boolean, ByVal angle As Single, Optional ByVal ModSizeX2 As Byte = 0, Optional ByVal color As Long = -1)
    '**************************************************************
    'Author: Mannakia
    'Last Modify Date: 14/05/2009
    'Modificado hoy(?) agregue funcion de agrandar y achicar para ladder :P
    '**************************************************************

    Static Piture As RECT

    With Piture
        .Left = 0
        .Top = 0
        
        If ModSizeX2 = 1 Then
            .bottom = pic.ScaleHeight / 2
            .Right = pic.ScaleWidth / 2
        ElseIf ModSizeX2 = 2 Then
            .bottom = pic.ScaleHeight * 2
            .Right = pic.ScaleWidth * 2
        Else
            .bottom = pic.ScaleHeight
            .Right = pic.ScaleWidth

        End If
        
    End With

    Dim s(3) As Long

    s(0) = color
    s(1) = color
    s(2) = color
    s(3) = color

    Call Engine_BeginScene
    
        Device_Box_Textured_Render grh, x, y, GrhData(grh).pixelWidth, GrhData(grh).pixelHeight, s, GrhData(grh).sX, GrhData(grh).sY, Alpha, angle

    Call Engine_EndScene(Piture, pic.hwnd)

End Sub


