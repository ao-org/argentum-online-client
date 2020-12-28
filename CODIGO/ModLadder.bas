Attribute VB_Name = "ModLadder"
Option Explicit

Public StopCreandoCuenta    As Boolean

Public Const DegreeToRadian As Single = 0.01745329251994 'Pi / 180
Public Const RadianToDegree As Single = 57.2958279087977 '180 / Pi

'Nueva seguridad
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal length As Long)
Private Declare Function GetAdaptersInfo Lib "iphlpapi" (lpAdapterInfo As Any, lpSize As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'get mac adress

Public Type Tclan
    nombre As String
    Alineacion As Byte
    indice As Integer
End Type

Public ListaClanes      As Boolean
Public ClanesList()     As Tclan

Public MacAdress        As String
Public HDserial         As Long

Public intro            As Byte

Public InviCounter      As Integer
Public ScrollExpCounter As Long
Public ScrollOroCounter As Long
Public OxigenoCounter   As Long
Public DrogaCounter     As Integer

Type Effect_Type

    FX_Grh     As grh      '< FxGrh.
    Fx_Index   As Integer  '< Indice del fx.
    ViajeChar  As Integer  '< CharIndex al que viaja.
    DestinoChar As Integer
    Viaje_X    As Integer   '< X hacia donde se dirije.
    End_Effect As Integer  '< Particula De la explosión.
    FxEnd_Effect As Integer  '< Particula De la explosión.
    End_Loops  As Integer  '< Loops del fx de la explosión.
    Viaje_Y    As Integer   '< Y hacia donde se dirije.
    ViajeSpeed As Single   '< Velocidad de viaje.
    Now_Moved  As Long     '< Tiempo del movimiento actual.
    Last_Move  As Long     '< Tiempo del último movimiento.
    Now_X      As Integer  '< Posición X actual
    Now_Y      As Integer  '< Posición Y actual
    Slot_Used  As Boolean  '< Si está usandose este slot.
    wav        As Integer
    DestX As Byte
    DesyY As Byte

End Type
 
Public Const NO_INDEX = -1         '< Índice no válido.
 
Public Effect()     As Effect_Type

'Destruccion de items
Public DestItemSlot As Byte
Public DestItemCant As Integer

Public Enum FXSound
    Lobo_Sound = 124
    Gallo_Sound = 137
    Dropeo_Sound = 132
    Casamiento_sound = 161
    BARCA_SOUND = 202
    MP_SOUND = 150
End Enum

Public Const MAX_CORREOS_SLOTS = 60

Public LastIndex2                        As Integer

Public CorreoMsj(1 To MAX_CORREOS_SLOTS) As CorreoMsj

Public ItemLista(1 To 10)                As Obj
Public ItemCount                         As Byte

Public Type CorreoMsj
    Remitente As String
    mensaje As String
    ItemCount As Byte
    ItemArray As String
    Leido As Byte
    Fecha As String
End Type

Public TieneFamiliar As Long

Public PetPercExp    As Long

Public HayLayer4     As Boolean

'Logros
Public NPcLogros     As TLogros
Public UserLogros    As TLogros
Public LevelLogros   As TLogros
Public MostrarTrofeo As Boolean

Type TLogros
    nombre As String
    desc As String
    cant As Long
    TipoRecompensa As Byte
    ObjRecompensa As String
    OroRecompensa As Long
    ExpRecompensa As Long
    HechizoRecompensa As Byte
    NpcsMatados As Integer
    NivelUser As Byte
    UserMatados As Integer
    Finalizada As Boolean
End Type

Public CantPartLLuvia     As Integer
Public MeteoIndex         As Integer

'Servidores
Public ChequeandoServidor As Byte
Public CantServer         As Byte

'Dropeo
Public CantdPaquetes      As Long
Public PingRender         As Integer
Public InBytes            As Long
Public OutBytes           As Long

Public NumOBJs            As Integer
Public NumNpcs            As Integer
Public NumHechizos        As Integer
Public NumLocaleMsg       As Integer
Public NumQuest           As Integer
Public NumSug             As Integer
Public Sugerencia()       As String




Public Type tQuestNpc

    NpcIndex As Integer
    Amount As Integer

End Type
 
Public Type tUserQuest

    NPCsKilled() As Integer
    QuestIndex As Integer

End Type

Public QuestList() As tQuest

Public Type tQuest

    nombre As String
    desc As String
    NextQuest As String
    DescFinal As String
    RequiredLevel As Byte
    
    RequiredQuest As Byte
    
    RequiredOBJs As Byte
    RequiredOBJ() As Obj
    
    RequiredNPCs As Byte
    RequiredNPC() As tQuestNpc
    
    RewardGLD As Long
    RewardEXP As Long
    
    RewardOBJs As Byte
    RewardOBJ() As Obj
    Repetible As Byte

End Type


Public PosMap()           As Integer

Public ObjData()          As ObjDatas
Public NpcData()          As NpcDatas

Public Locale_SMG()       As String



'Sistema de mapa del mundo
Public TotalWorlds As Byte

Public Type WorldMap
    MapIndice() As Integer
    Ancho As Integer
    Alto As Integer
End Type

Public Mundo() As WorldMap


Public PosREAL          As Integer
Public Dungeon          As Boolean
Public idmap            As Integer

Public WorldActual As Byte


'Sistema de mapa del mundo


Public HechizoData()      As HechizoDatas

Public NameMaps(1 To 1000) As NameMapas

Public ShowMacros         As Byte

Public OcultarMacro       As Boolean

Public ModoCaminata       As Boolean

Public MacrosBloqeados    As Boolean

Public Type ObjDatas

    GrhIndex As Long ' Indice del grafico que representa el obj
    Name As String
    MinDef As Integer
    MaxDef As Integer
    MinHit As Integer
    MaxHit As Integer
    ObjType As Byte
    Texto As String
    info As String
    CreaGRH As String
    CreaLuz As String
    CreaParticulaPiso As Integer
    proyectil As Byte
    Raices As Integer
    Madera As Integer
    PielLobo As Integer
    PielOsoPardo As Integer
    PielOsoPolar As Integer
    LingH As Integer
    LingP As Integer
    LingO As Integer
    Destruye As Byte
    SkHerreria As Byte
    SkPociones As Byte
    Sksastreria As Byte
    Valor As Long

End Type

Public Type NpcDatas

    Name As String
    desc As String
    Body As Integer
    Hp As Long
    exp As Long
    oro As Long
    MinHit As Integer
    MaxHit As Integer
    Head As Integer
    NumQuiza As Byte
    QuizaDropea() As Integer
    ExpClan As Long
    
End Type

Public Type HechizoDatas

    nombre As String ' Indice del grafico que representa el obj
    desc As String
    PalabrasMagicas As String
    HechizeroMsg As String
    TargetMsg As String
    PropioMsg As String
    ManaRequerido As Integer
    MinSkill As Byte
    StaRequerido As Integer
    IconoIndex As Long

End Type

Public Type NameMapas

    Name As String ' Indice del grafico que representa el obj
    desc As String

End Type

Public Type SvMsg

    nombre As String ' Indice del grafico que representa el obj

End Type

Public Enum Accion_Barra

    Runa = 1
    Resucitar = 2
    Intermundia = 3
    BattleModo = 4
    GoToPareja = 5
    CancelarAccion = 99

End Enum

Public UserMacro As Macro

Type Macro

    Activado As Boolean
    Intervalo As Integer
    TIPO As Byte
    cantidad As Integer
    Index As Integer
    tX As Byte
    tY As Byte
    Skill As Byte

End Type

Public MouseS As Long

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
 
Private Const SPI_SETMOUSESPEED = 113

Private Const SPI_GETMOUSESPEED = 112

Public Declare Function SwapMouseButton Lib "user32" (ByVal bSwap As Long) As Long

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Type COPYDATASTRUCT

    dwData As Long
    cbData As Long
    lpData As Long

End Type

Public Const WM_COPYDATA = &H4A

Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long

Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long

Private hBuffersTimer As Long

'Compresion
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Const MAX_LENGTH = 512

Private Const CONST_INTERVALO_ANIM        As Long = 150

Private Const CONST_INTERVALO_TIRAR       As Long = 1200

Private Const CONST_INTERVALO_Conectar    As Long = 100

Private Const CONST_INTERVALO_LLAMADACLAN As Long = 5000

Private Const CONST_INTERVALO_COMBO       As Long = 450

Private Const CONST_INTERVALO_HEADING     As Long = 120

Private Const CONST_INTERVALO_CLICK       As Long = 200

Public Intervalos                         As tIntervalos

Public Type tIntervalos

    Anim As Long
    Ataque As Long
    Uso As Long
    Trabajo As Long
    Hechizo As Long
    tirar As Long
    Conectar As Long
    Subir As Long
    Presentacion As Long
    ComboGolpeMagia As Long
    ComboMagiaGolpe As Long
    Heading As Long
    Click As Long
    LLamadaClan As Long
    HechizoMacro As Long
    UsarDespuesDeAtacar As Long

End Type

Public LogeoAlgunaVez               As Boolean

Public Pjs(1 To 10)                 As UserCuentaPJS

Public RecordarCuenta               As Boolean

Public CuentaRecordada              As CuentasGuardadas

Public CantidadDePersonajesEnCuenta As Byte

Type UserCuentaPJS

    nombre As String
    nivel As Byte
    Mapa As Integer
    Body As Integer
    Head As Integer
    Criminal As Byte
    Clase As Byte
    NameMapa As String
    LetraColor As RGBA
    Arma As Integer
    Escudo As Integer
    Casco As Integer
    ClanName As String
    priv As Byte

End Type

Type CuentasGuardadas

    nombre As String
    Password As String

End Type
 
Public Type POINTAPI

    x As Long
    y As Long

End Type
 
Private Type SAFEARRAYBOUND

    cElements As Long
    lLbound As Long

End Type
 
Private Type SAFEARRAY2D

    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND

End Type
 
Private Type BITMAP

    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long

End Type
 
Const HTCAPTION = 2

Const WM_NCLBUTTONDOWN = &HA1

Const RGN_OR = 2
 
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long

Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long

Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long
 
Public lRegion             As Long

Public Render_Connect_Rect As RECT

Public Render_Main_Rect    As RECT

Public Const SWP_NOMOVE = 2

Public Const SWP_NOSIZE = 1

Public Const FLAGSz = SWP_NOMOVE Or SWP_NOSIZE

Public Const HWND_TOPMOST = -1

Public Const HWND_NOTOPMOST = -2

Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function CreateIconFromResourceEx Lib "user32.dll" (ByRef presbits As Any, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal flags As Long) As Long

Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (lpPictDesc As Any, riid As Any, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long

Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long

Public Declare Function DrawIconEx Lib "user32.dll" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
  
Private Declare Function SendMessageLongRef Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Long) As Long
                           
Private m_ASC As Long

Sub inputbox_Password(El_Form As Form, Caracter As String)
    
    On Error GoTo inputbox_Password_Err
    
      
    m_ASC = Asc(Caracter)
      
    Call SetTimer(El_Form.hWnd, &H5000&, 100, AddressOf TimerProc)
  
    
    Exit Sub

inputbox_Password_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.inputbox_Password", Erl)
    Resume Next
    
End Sub
  
Private Sub TimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    
    On Error GoTo TimerProc_Err
    
           
    Dim Handle_InputBox As Long
      
    'Captura el handle del textBox del InputBox
    Handle_InputBox = FindWindowEx(FindWindow("#32770", App.title), 0, "Edit", "")
                  
    'Le establece el PasswordChar
    Call SendMessageLongRef(Handle_InputBox, &HCC&, m_ASC, 0)
    'Finaliza el Timer
    Call KillTimer(hWnd, idEvent)
  
    
    Exit Sub

TimerProc_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.TimerProc", Erl)
    Resume Next
    
End Sub

Public Function LoadPNGtoICO(pngData() As Byte) As IPicture
    
    On Error GoTo LoadPNGtoICO_Err
    
    
    Dim hIcon              As Long

    Dim lpPictDesc(0 To 3) As Long, aGUID(0 To 3) As Long
    
    hIcon = CreateIconFromResourceEx(pngData(0), UBound(pngData) + 1&, 1&, &H30000, 0&, 0&, 0&)

    If hIcon Then
        lpPictDesc(0) = 16&
        lpPictDesc(1) = vbPicTypeIcon
        lpPictDesc(2) = hIcon
        ' IPicture GUID {7BF80980-BF32-101A-8BBB-00AA00300CAB}
        aGUID(0) = &H7BF80980
        aGUID(1) = &H101ABF32
        aGUID(2) = &HAA00BB8B
        aGUID(3) = &HAB0C3000

        ' create stdPicture
        If OleCreatePictureIndirect(lpPictDesc(0), aGUID(0), True, LoadPNGtoICO) Then
            DestroyIcon hIcon

        End If
        
    End If
    
    
    Exit Function

LoadPNGtoICO_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.LoadPNGtoICO", Erl)
    Resume Next
    
End Function

Public Function SetTopMostWindow(hWnd As Long, Topmost As Boolean) As Long
    
    On Error GoTo SetTopMostWindow_Err
    

    If Topmost = True Then 'Make the window topmost
        SetTopMostWindow = SetWindowPos(hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGSz)
    Else
        SetTopMostWindow = SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGSz)
        SetTopMostWindow = False

    End If

    
    Exit Function

SetTopMostWindow_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.SetTopMostWindow", Erl)
    Resume Next
    
End Function

Public Sub LogError(desc As String)

    On Error GoTo errhandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.Path & "\errores.log" For Append Shared As #nfile
    Print #nfile, Date & "-" & Time & ":" & desc
    Close #nfile

    Exit Sub

errhandler:

End Sub

Sub IniciarCrearPj()
    
    On Error GoTo IniciarCrearPj_Err
    
    StopCreandoCuenta = False

    frmCrearPersonaje.lbFuerza.Caption = 18
    frmCrearPersonaje.lbAgilidad.Caption = 18
    frmCrearPersonaje.lbInteligencia.Caption = 18
    frmCrearPersonaje.lbConstitucion.Caption = 18
    frmCrearPersonaje.lbCarisma.Caption = 18
    frmCrearPersonaje.lbLagaRulzz.Caption = 0

    Dim i As Integer

    frmCrearPersonaje.lstRaza.Clear

    For i = LBound(ListaRazas()) To UBound(ListaRazas())
        frmCrearPersonaje.lstRaza.AddItem ListaRazas(i)
    Next i
    
    
    frmCrearPersonaje.lstHogar.Clear
    For i = LBound(ListaCiudades()) To UBound(ListaCiudades())
        frmCrearPersonaje.lstHogar.AddItem (ListaCiudades(i))
    Next i

    frmCrearPersonaje.lstProfesion.Clear

    For i = LBound(ListaClases()) To UBound(ListaClases())
        frmCrearPersonaje.lstProfesion.AddItem ListaClases(i)
    
    Next i

    
    Exit Sub

IniciarCrearPj_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.IniciarCrearPj", Erl)
    Resume Next
    
End Sub

Sub General_Set_Connect()
    
    On Error GoTo General_Set_Connect_Err
    

    AlphaNiebla = 75
    EntradaY = 10
    EntradaX = 10

    UserMap = 1
    Call SwitchMap(UserMap)

    If QueRender <> 1 Then
        frmConnect.Show
        FrmLogear.Show , frmConnect
        FrmLogear.Top = FrmLogear.Top + 3500
    End If
            
    intro = 1
    frmMain.Picture = LoadInterface("ventanaprincipal.bmp")
    frmMain.panel.Picture = LoadInterface("centroinventario.bmp")
    frmMain.ExpBar.Picture = LoadInterface("barraexperiencia.bmp")
    frmMain.COMIDAsp.Picture = LoadInterface("barradehambre.bmp")
    frmMain.AGUAsp.Picture = LoadInterface("barradesed.bmp")
    frmMain.MANShp.Picture = LoadInterface("barrademana.bmp")
    frmMain.STAShp.Picture = LoadInterface("barradeenergia.bmp")
    frmMain.Hpshp.Picture = LoadInterface("barradevida.bmp")
            
    Sound.Sound_Play CStr(SND_LLUVIAIN), True, 0, 0
    AlphaNiebla = 10
    
    Call Graficos_Particulas.Engine_spell_Particle_Set(41)

    If intro = 1 Then
        Call Graficos_Particulas.Engine_MeteoParticle_Set(207)

    End If

    Call Graficos_Particulas.Engine_Select_Particle_Set(203)
    
    ParticleLluviaDorada = Graficos_Particulas.General_Particle_Create(208, -1, -1)

    Sound.Music_Load 1, Sound.VolumenActualMusicMax
    Sound.Music_Play
    'Sound.Fading = 600

    mFadingMusicMod = 0
    CurMp3 = 1

    QueRender = 1
    frmConnect.relampago.Enabled = True
    'Sound.Sound_Play 650, False, 0, 0

    'frmConnect.Timer1.Enabled = True

    frmConnect.relampago.Enabled = True
    'Sound.Sound_Play 404, False, 0, 0   LADDER REVISAR SAQUE TRUENO
    ClickEnAsistente = 0

    If CuentaRecordada.nombre <> "" Then
        Call TextoAlAsistente("¡Bienvenido de nuevo! ¡Disfruta tu viaje por Argentum20!") ' hay que poner 20 aniversario
    Else
        Call TextoAlAsistente("¡Bienvenido a Argentum20! ¿Ya tenes tu cuenta? Logea! sino, toca sobre Cuenta para crearte una.") ' hay que poner 20 aniversario

    End If

    
    Exit Sub

General_Set_Connect_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.General_Set_Connect", Erl)
    Resume Next
    
End Sub
 
Public Sub InitializeSurfaceCapture(frm As Form)
    
    On Error GoTo InitializeSurfaceCapture_Err
    
    lRegion = CreateRectRgn(0, 0, 0, 0)
    frm.Visible = False

    
    Exit Sub

InitializeSurfaceCapture_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.InitializeSurfaceCapture", Erl)
    Resume Next
    
End Sub

Public Sub ReleaseSurfaceCapture(frm As Form)
    
    On Error GoTo ReleaseSurfaceCapture_Err
    
    ApplySurfaceTo frm
    frm.Visible = True
    Call DeleteObject(lRegion)

    
    Exit Sub

ReleaseSurfaceCapture_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.ReleaseSurfaceCapture", Erl)
    Resume Next
    
End Sub
 
Public Sub ApplySurfaceTo(frm As Form)
    
    On Error GoTo ApplySurfaceTo_Err
    
    Call SetWindowRgn(frm.hWnd, lRegion, True)

    
    Exit Sub

ApplySurfaceTo_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.ApplySurfaceTo", Erl)
    Resume Next
    
End Sub
 
' Create a polygonal region - has to be more than 2 pts (or 4 input values)
Public Sub CreateSurfacefromPoints(ParamArray XY())
    
    On Error GoTo CreateSurfacefromPoints_Err
    

    Dim lRegionTemp As Long

    Dim XY2()       As POINTAPI

    Dim nIndex      As Integer

    Dim nTemp       As Integer

    Dim nSize       As Integer

    nSize = CInt(UBound(XY) / 2) - 1
    ReDim XY2(nSize + 2)
    nIndex = 0

    For nTemp = 0 To nSize
        XY2(nTemp).x = XY(nIndex)
        nIndex = nIndex + 1
        XY2(nTemp).y = XY(nIndex)
        nIndex = nIndex + 1
    Next nTemp

    lRegionTemp = CreatePolygonRgn(XY2(0), (UBound(XY2) + 1), 2)
    Call CombineRgn(lRegion, lRegion, lRegionTemp, RGN_OR)
    Call DeleteObject(lRegionTemp)

    
    Exit Sub

CreateSurfacefromPoints_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.CreateSurfacefromPoints", Erl)
    Resume Next
    
End Sub
 
' Create a ciruclar/elliptical region
Public Sub CreateSurfacefromEllipse(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer)
    
    On Error GoTo CreateSurfacefromEllipse_Err
    

    Dim lRegionTemp As Long

    lRegionTemp = CreateEllipticRgn(x1, y1, x2, y2)
    Call CombineRgn(lRegion, lRegion, lRegionTemp, RGN_OR)
    Call DeleteObject(lRegionTemp)

    
    Exit Sub

CreateSurfacefromEllipse_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.CreateSurfacefromEllipse", Erl)
    Resume Next
    
End Sub
 
' Create a rectangular region
Public Sub CreateSurfacefromRect(x1 As Integer, y1 As Integer, x2 As Integer, y2 As Integer)
    
    On Error GoTo CreateSurfacefromRect_Err
    

    Dim lRegionTemp As Long

    lRegionTemp = CreateRectRgn(x1, y1, x2, y2)
    Call CombineRgn(lRegion, lRegion, lRegionTemp, RGN_OR)
    Call DeleteObject(lRegionTemp)

    
    Exit Sub

CreateSurfacefromRect_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.CreateSurfacefromRect", Erl)
    Resume Next
    
End Sub
 
' My best creation (more like tweak) yet! Super fast routines qown j00!
Public Sub CreateSurfacefromMask(Obj As Object, Optional lBackColor As Long)
    
    On Error GoTo CreateSurfacefromMask_Err
    

    ' Insight: Down with getpixel!!
    Dim lReturn  As Long

    Dim lRgnTmp  As Long

    Dim lSkinRgn As Long

    Dim lStart   As Long

    Dim lRow     As Long

    Dim lCol     As Long

    Dim glHeight As Integer

    Dim glWidth  As Integer

    Dim pict()   As Byte

    Dim pict2()  As Byte

    Dim sa       As SAFEARRAY2D

    Dim bmp      As BITMAP

    GetObjectAPI Obj.Picture, Len(bmp), bmp

    ' Load the bmp into a safearray ptr
    With sa
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = bmp.bmHeight
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = bmp.bmWidthBytes
        .pvData = bmp.bmBits

    End With

    ' Unfortunately this only supports 256 color bmps (damn high bit graphics!!)
    If bmp.bmBitsPixel <> 8 Then
        CreateSurfacefromMask_GetPixel Obj
        Exit Sub

    End If

    CopyMemory ByVal VarPtrArray(pict), VarPtr(sa), 4
    ' Get the dimensions for future reference
    glHeight = UBound(pict, 2)
    glWidth = UBound(pict, 1)
    ' Create an identity array to flip the damn inversed regions
    ReDim pict2(glWidth, glHeight)

    ' Flip em!
    Dim nTempX As Integer

    Dim nTempY As Integer

    For nTempX = glWidth To 0 Step -1
        For nTempY = glHeight To 0 Step -1
            pict2(nTempX, nTempY) = pict(nTempX, glHeight - nTempY)
        Next nTempY
    Next nTempX

    ' Clear the original array
    CopyMemory ByVal VarPtrArray(pict), 0&, 4
    ' Let's make our regions!
    lSkinRgn = CreateRectRgn(0, 0, 0, 0)

    With Obj

        If lBackColor < 1 Then lBackColor = pict2(0, 0)

        For lRow = 0 To glHeight
            lCol = 0

            Do While lCol < glWidth
                Do While lCol < glWidth

                    If pict2(lCol, lRow) = lBackColor Then
                        lCol = lCol + 1
                    Else
                        Exit Do

                    End If

                Loop

                If lCol < glWidth Then
                    lStart = lCol

                    Do While lCol < glWidth

                        If pict2(lCol, lRow) <> lBackColor Then
                            lCol = lCol + 1
                        Else
                            Exit Do

                        End If

                    Loop

                    If lCol > glWidth Then lCol = glWidth
                    lRgnTmp = CreateRectRgn(lStart, lRow, lCol, (lRow + 1))
                    lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
                    Call DeleteObject(lRgnTmp)

                End If

            Loop
        Next

    End With

    ' Clear the identity array
    CopyMemory ByVal VarPtrArray(pict2), 0&, 4
    ' Return the f****** fast generated region!
    lReturn = CombineRgn(lRegion, lRegion, lSkinRgn, RGN_OR)

    
    Exit Sub

CreateSurfacefromMask_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.CreateSurfacefromMask", Erl)
    Resume Next
    
End Sub
 
' XCopied from The Scarms! Felt like my obligation to leave this code intact w/o
' any changes to variables, etc (cept for the sub's name). Thanks d00d!
Public Sub CreateSurfacefromMask_GetPixel(Obj As Object, Optional lBackColor As Long)
    
    On Error GoTo CreateSurfacefromMask_GetPixel_Err
    

    Dim lReturn  As Long

    Dim lRgnTmp  As Long

    Dim lSkinRgn As Long

    Dim lStart   As Long

    Dim lRow     As Long

    Dim lCol     As Long

    Dim glHeight As Integer

    Dim glWidth  As Integer

    lSkinRgn = CreateRectRgn(0, 0, 0, 0)

    With Obj
        glHeight = .Height / Screen.TwipsPerPixelY
        glWidth = .Width / Screen.TwipsPerPixelX

        If lBackColor < 1 Then lBackColor = GetPixel(.hdc, 0, 0)

        For lRow = 0 To glHeight - 1
            lCol = 0

            Do While lCol < glWidth
                Do While lCol < glWidth And GetPixel(.hdc, lCol, lRow) = lBackColor
                    lCol = lCol + 1
                Loop

                If lCol < glWidth Then
                    lStart = lCol

                    Do While lCol < glWidth And GetPixel(.hdc, lCol, lRow) <> lBackColor
                        lCol = lCol + 1
                    Loop

                    If lCol > glWidth Then lCol = glWidth
                    lRgnTmp = CreateRectRgn(lStart, lRow, lCol, lRow + 1)
                    lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
                    Call DeleteObject(lRgnTmp)

                End If

            Loop
        Next

    End With

    lReturn = CombineRgn(lRegion, lRegion, lSkinRgn, RGN_OR)

    
    Exit Sub

CreateSurfacefromMask_GetPixel_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.CreateSurfacefromMask_GetPixel", Erl)
    Resume Next
    
End Sub

Public Sub General_Var_Write(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)
    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Writes a var to a text file
    '*****************************************************************
    
    On Error GoTo General_Var_Write_Err
    
    writeprivateprofilestring Main, Var, Value, File

    
    Exit Sub

General_Var_Write_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.General_Var_Write", Erl)
    Resume Next
    
End Sub

Public Sub MensajeAdvertencia(ByVal mensaje As String)
    
    On Error GoTo MensajeAdvertencia_Err
    
    Call MsgBox(mensaje, vbInformation + vbOKOnly, "Advertencia")

    
    Exit Sub

MensajeAdvertencia_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.MensajeAdvertencia", Erl)
    Resume Next
    
End Sub

Public Sub ReproducirMp3(ByVal mp3 As Byte)
    
    On Error GoTo ReproducirMp3_Err
    

    If mp3 <> CurMp3 Then
        If mp3 <> 0 Then
            NextMP3 = mp3
            mFadingMusicMod = 0

            ' frmMain.TimerMusica.Enabled = True
        End If

    End If

    
    Exit Sub

ReproducirMp3_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.ReproducirMp3", Erl)
    Resume Next
    
End Sub

Public Sub ForzarMp3(ByVal mp3 As Byte)
    
    On Error GoTo ForzarMp3_Err
    

    If mp3 = 0 Then Exit Sub

    mFadingMusicMod = 0
    CurMp3 = mp3

    
    Exit Sub

ForzarMp3_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.ForzarMp3", Erl)
    Resume Next
    
End Sub

Public Sub CargarCuentasGuardadas()
    
    On Error GoTo CargarCuentasGuardadas_Err
    

    Dim Arch As String

    Arch = App.Path & "\..\Recursos\OUTPUT\" & "Configuracion.ini"
    CuentaRecordada.nombre = GetVar(Arch, "CUENTA", "Nombre")
    CuentaRecordada.Password = UnEncryptStr(GetVar(Arch, "CUENTA", "Password"), 9256)
    FrmLogear.Image4.Tag = "0"
 
    If CuentaRecordada.nombre <> "" Then
        FrmLogear.NameTxt = CuentaRecordada.nombre
        FrmLogear.PasswordTxt = CuentaRecordada.Password
        FrmLogear.Image4.Picture = LoadInterface("check-amarillo.bmp")
        FrmLogear.Image4.Tag = "1"
        'FrmLogear.Check1.value = 1
         
        FrmLogear.PasswordTxt.TabIndex = 0
        
        FrmLogear.PasswordTxt.SelStart = Len(FrmLogear.PasswordTxt)

        'FrmLogear.lstServers.TabIndex = 1
        'FrmLogear.cmdConnect.TabIndex = 2
    End If

    Rem FrmLogear.PasswordTxt = CuentaRecordada(1).Password
    
    Exit Sub

CargarCuentasGuardadas_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.CargarCuentasGuardadas", Erl)
    Resume Next
    
End Sub

Public Sub GrabarNuevaCuenta(ByVal Name As String, ByVal Password As String)
    
    On Error GoTo GrabarNuevaCuenta_Err
    

    Dim Arch As String

    Arch = App.Path & "\..\Recursos\OUTPUT\" & "Configuracion.ini"
    Call WriteVar(Arch, "CUENTA", "Nombre", Name)
    Call WriteVar(Arch, "CUENTA", "Password", EncryptStr(Password, 9256))
    Call CargarCuentasGuardadas

    
    Exit Sub

GrabarNuevaCuenta_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.GrabarNuevaCuenta", Erl)
    Resume Next
    
End Sub

Public Sub ResetearCuentas()
    
    On Error GoTo ResetearCuentas_Err
    

    Dim Arch As String

    Arch = App.Path & "\..\Recursos\OUTPUT\Configuracion.ini"
    Call WriteVar(Arch, "CUENTA", "Nombre", "")
    Call WriteVar(Arch, "CUENTA", "Password", "")
    Call CargarCuentasGuardadas

    
    Exit Sub

ResetearCuentas_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.ResetearCuentas", Erl)
    Resume Next
    
End Sub

'*****************************************************************
'modTimer - ImperiumAO - v1.3.0
'
'Windows API timer functions and handles.
'
'*****************************************************************
'Respective portions copyrighted by contributors listed below.
'
'This library is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation version 2.1 of
'the License
'
'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
'Lesser General Public License for more details.
'
'You should have received a copy of the GNU Lesser General Public
'License along with this library; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'*****************************************************************

'*****************************************************************
'Augusto José Rando (barrin@imperiumao.com.ar)
'   - First Relase
'*****************************************************************

Public Function IntervaloPermiteClick(Optional ByVal Actualizar As Boolean = True) As Boolean
    
    On Error GoTo IntervaloPermiteClick_Err
    

    If FrameTime - Intervalos.Click >= CONST_INTERVALO_CLICK Then
        If Actualizar Then
            Intervalos.Click = FrameTime

        End If

        IntervaloPermiteClick = True
        'Call AddtoRichTextBox(frmMain.RecTxt, "Golpe - Magia OK.", 255, 0, 0, True, False, False)
    Else
        IntervaloPermiteClick = False

        'Call AddtoRichTextBox(frmMain.RecTxt, "Golpe - Magia NO.", 255, 0, 0, True, False, False)
    End If

    
    Exit Function

IntervaloPermiteClick_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.IntervaloPermiteClick", Erl)
    Resume Next
    
End Function

Public Function IntervaloPermiteHeading(Optional ByVal Actualizar As Boolean = True) As Boolean
    
    On Error GoTo IntervaloPermiteHeading_Err
    

    If FrameTime - Intervalos.Heading >= CONST_INTERVALO_HEADING Then
        If Actualizar Then
            Intervalos.Heading = FrameTime

        End If

        IntervaloPermiteHeading = True
        'Call AddtoRichTextBox(frmMain.RecTxt, "Golpe - Magia OK.", 255, 0, 0, True, False, False)
    Else
        IntervaloPermiteHeading = False

        'Call AddtoRichTextBox(frmMain.RecTxt, "Golpe - Magia NO.", 255, 0, 0, True, False, False)
    End If

    
    Exit Function

IntervaloPermiteHeading_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.IntervaloPermiteHeading", Erl)
    Resume Next
    
End Function

Public Function IntervaloPermiteLLamadaClan() As Boolean
    
    On Error GoTo IntervaloPermiteLLamadaClan_Err
    

    If FrameTime - Intervalos.LLamadaClan >= CONST_INTERVALO_LLAMADACLAN Then
        
        '  Call AddtoRichTextBox(frmMain.RecTxt, "Usar OK.", 255, 0, 0, True, False, False)
        Intervalos.LLamadaClan = FrameTime
        IntervaloPermiteLLamadaClan = True
    Else
        IntervaloPermiteLLamadaClan = False

        ' Call AddtoRichTextBox(frmMain.RecTxt, "Debes aguardar unos instantes para volver a llamar a tu clan.", 255, 0, 0, True, False, False)
    End If

    
    Exit Function

IntervaloPermiteLLamadaClan_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.IntervaloPermiteLLamadaClan", Erl)
    Resume Next
    
End Function

Public Function IntervaloPermiteAnim() As Boolean
    
    On Error GoTo IntervaloPermiteAnim_Err
    

    If FrameTime - Intervalos.Anim >= CONST_INTERVALO_ANIM Then
        
        '  Call AddtoRichTextBox(frmMain.RecTxt, "Usar OK.", 255, 0, 0, True, False, False)
        Intervalos.Anim = FrameTime
        IntervaloPermiteAnim = True
    Else
        IntervaloPermiteAnim = False

    End If

    
    Exit Function

IntervaloPermiteAnim_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.IntervaloPermiteAnim", Erl)
    Resume Next
    
End Function

Public Function IntervaloPermiteConectar() As Boolean
    
    On Error GoTo IntervaloPermiteConectar_Err
    

    If FrameTime - Intervalos.Conectar >= CONST_INTERVALO_Conectar Then
        
        ' Call AddtoRichTextBox(frmMain.RecTxt, "Usar OK.", 255, 0, 0, True, False, False)
        Intervalos.Conectar = FrameTime
        IntervaloPermiteConectar = True
    Else
        IntervaloPermiteConectar = False

    End If

    
    Exit Function

IntervaloPermiteConectar_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.IntervaloPermiteConectar", Erl)
    Resume Next
    
End Function

Sub CargarOpciones()

    On Error GoTo ErrorHandler
    
    If FileExist(App.Path & "\..\Recursos\OUTPUT\Configuracion.ini", vbArchive) Then
        Call LoadImpAoInit
    Else
        Call MsgBox("¡No se puede cargar el archivo de opciones! La reinstalacion del juego podria solucionar el problema.", vbCritical, "Error al cargar")
        End
    End If
    
    Dim ConfigFile As clsIniManager
    Set ConfigFile = New clsIniManager
    Call ConfigFile.Initialize(App.Path & "\..\Recursos\OUTPUT\" & "Configuracion.ini")
    
    'Musica y Sonido
    Musica = ConfigFile.GetValue("AUDIO", "Musica")
    Sonido = ConfigFile.GetValue("AUDIO", "Sonido")
    fX = ConfigFile.GetValue("AUDIO", "Fx")
    AmbientalActivated = ConfigFile.GetValue("AUDIO", "AmbientalActivated")
    InvertirSonido = ConfigFile.GetValue("AUDIO", "InvertirSonido")
    
    'Musica y Sonido - Volumen
    VolMusicFadding = VolMusic
    VolMusic = Val(ConfigFile.GetValue("AUDIO", "VolMusic"))
    VolFX = Val(ConfigFile.GetValue("AUDIO", "VolFX"))
    VolAmbient = Val(ConfigFile.GetValue("AUDIO", "VolAmbient"))
    
    'Video
    PantallaCompleta = ConfigFile.GetValue("VIDEO", "PantallaCompleta")
    CursoresGraficos = IIf(RunningInVB, 0, ConfigFile.GetValue("VIDEO", "CursoresGraficos"))
    UtilizarPreCarga = ConfigFile.GetValue("VIDEO", "UtilizarPreCarga")
    
    FxNavega = ConfigFile.GetValue("OPCIONES", "FxNavega")
    OcultarMacrosAlCastear = ConfigFile.GetValue("OPCIONES", "OcultarMacrosAlCastear")
    MostrarIconosMeteorologicos = ConfigFile.GetValue("OPCIONES", "MostrarIconosMeteorologicos")
    CopiarDialogoAConsola = ConfigFile.GetValue("OPCIONES", "CopiarDialogoAConsola")
    PermitirMoverse = ConfigFile.GetValue("OPCIONES", "PermitirMoverse")
    MoverVentana = ConfigFile.GetValue("OPCIONES", "MoverVentana")
    FPSFLAG = ConfigFile.GetValue("OPCIONES", "FPSFLAG")
    AlphaMacro = ConfigFile.GetValue("OPCIONES", "AlphaMacro")
    MostrarEscribiendo = Val(ConfigFile.GetValue("OPCIONES", "MostrarEscribiendo"))
    
    'Init
    ServerIndex = Val(ConfigFile.GetValue("INIT", "ServerIndex"))

    SensibilidadMouse = ConfigFile.GetValue("OPCIONES", "SensibilidadMouse")
    
    Set ConfigFile = Nothing
    
    If SensibilidadMouse = 0 Then: SensibilidadMouse = 10

    SensibilidadMouseOriginal = General_Get_Mouse_Speed
    
    Call General_Set_Mouse_Speed(SensibilidadMouse)
    
    Exit Sub
    
ErrorHandler:
    
    Set ConfigFile = Nothing
    
    Call MsgBox("Ha ocurrido un error al cargar la configuración del juego.", vbCritical, "Configuración del Juego")
    
    End
    
End Sub

Sub GuardarOpciones()
    
    On Error GoTo GuardarOpciones_Err
    

    Dim Arch As String: Arch = App.Path & "\..\Recursos\OUTPUT\" & "Configuracion.ini"
    
    Call WriteVar(Arch, "INIT", "ServerIndex", Int(ServerIndex))
    
    Call WriteVar(Arch, "AUDIO", "Musica", Musica)
    Call WriteVar(Arch, "AUDIO", "Fx", fX)
    Call WriteVar(Arch, "AUDIO", "VolMusic", VolMusic)
    Call WriteVar(Arch, "AUDIO", "Volfx", VolFX)
    Call WriteVar(Arch, "AUDIO", "VolAmbient", VolAmbient)
    Call WriteVar(Arch, "AUDIO", "AmbientalActivated", AmbientalActivated)
    
    'Call WriteVar(Arch, "VIDEO", "CursoresGraficos", CursoresGraficos)
    
    Call WriteVar(Arch, "OPCIONES", "MoverVentana", MoverVentana)
    Call WriteVar(Arch, "OPCIONES", "PermitirMoverse", PermitirMoverse)
    Call WriteVar(Arch, "OPCIONES", "CopiarDialogoAConsola", CopiarDialogoAConsola)
    Call WriteVar(Arch, "OPCIONES", "InvertirSonido", InvertirSonido)
    Call WriteVar(Arch, "OPCIONES", "FPSFLAG", FPSFLAG)
    Call WriteVar(Arch, "OPCIONES", "AlphaMacro", AlphaMacro)
    Call WriteVar(Arch, "OPCIONES", "FxNavega", FxNavega)
    
    Call WriteVar(Arch, "OPCIONES", "MostrarEscribiendo", MostrarEscribiendo)
    

    Call WriteVar(Arch, "OPCIONES", "OcultarMacrosAlCastear", OcultarMacrosAlCastear)
    
    Call WriteVar(Arch, "OPCIONES", "SensibilidadMouse", SensibilidadMouse)

    
    Exit Sub

GuardarOpciones_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.GuardarOpciones", Erl)
    Resume Next
    
End Sub

Public Sub WriteChatOverHeadInConsole(ByVal charindex As Integer, ByVal ChatText As String, ByVal red As Byte, ByVal green As Byte, ByVal blue As Byte)
    
    On Error GoTo WriteChatOverHeadInConsole_Err
    

    Dim NameRed   As Byte

    Dim NameGreen As Byte

    Dim NameBlue  As Byte
    
    If red = 20 And green = 226 And blue = 157 Then
        Exit Sub

    End If
    
    With charlist(charindex)
        'Todo: Hacer que los colores se usen de Colores.dat
        'Haciendo uso de ColoresPj ya que el mismo en algun momento lo hace para DX
        Select Case .priv

            Case 0

                If .status = 0 Then
                    NameRed = 128
                    NameGreen = 128
                    NameBlue = 128
                ElseIf .status = 1 Then
                    NameRed = 0
                    NameGreen = 128
                    NameBlue = 190
                ElseIf .status = 2 Then
                    NameRed = 179
                    NameGreen = 0
                    NameBlue = 4
                ElseIf .status = 3 Then
                    NameRed = 31
                    NameGreen = 139
                    NameBlue = 139

                End If

            Case 1, 2

                NameRed = 2
                NameGreen = 161
                NameBlue = 38

            Case 3, 4
                NameRed = 217
                NameGreen = 164
                NameBlue = 32
            
        End Select

        Dim Pos As Integer

        Pos = InStr(.nombre, "<")
            
        If Pos = 0 Then Pos = LenB(.nombre) + 2
        
        Dim Name As String

        Name = Left$(.nombre, Pos - 2)
       
        'Si el npc tiene nombre lo escribimos en la consola
        ChatText = Trim$(ChatText)

        If LenB(.nombre) <> 0 And LenB(ChatText) > 0 Then
            Call AddtoRichTextBox2(frmMain.RecTxt, "[" & Name & "] ", NameRed, NameGreen, NameBlue, True, False, True, rtfLeft)
            Call AddtoRichTextBox2(frmMain.RecTxt, ChatText, red, green, blue, False, False, False, rtfLeft)

        End If

        Dim i As Byte

    End With
    
    
    Exit Sub

WriteChatOverHeadInConsole_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.WriteChatOverHeadInConsole", Erl)
    Resume Next
    
End Sub

Public Sub CopiarDialogoToConsola(ByVal NickName As String, Dialogo As String, Color As Long)
    
    On Error GoTo CopiarDialogoToConsola_Err
    

    If NickName = "" Then Exit Sub
    If Right$(Dialogo, 1) = " " Or Left(Dialogo, 1) = " " Then
        Dialogo = Trim(Dialogo)

    End If

    Dim Pos  As Long

    Dim Nick As String

    Pos = InStr(NickName, "<")

    If Pos = 0 Then Pos = Len(NickName) + 2
    'Nick
    Nick = Left$(NickName, Pos - 2)

    Select Case Color

        Case 255255255 ' Blanco comun
            Call AddtoRichTextBox(frmMain.RecTxt, Nick & "> " & Dialogo, 255, 255, 255, False, True, False)

        Case 25513015 'Gritar GMS!
            Call AddtoRichTextBox(frmMain.RecTxt, Nick & "> " & Dialogo, 225, 225, 0, False, True, False)

        Case 25500 ' Gritar!
            Call AddtoRichTextBox(frmMain.RecTxt, Nick & "> " & Dialogo, 255, 0, 0, False, True, False)

        Case 2000 'GM
            Call AddtoRichTextBox(frmMain.RecTxt, Nick & "> " & Dialogo, 0, 200, , False, True, False)

        Case -14117888 ' Global
            Call AddtoRichTextBox(frmMain.RecTxt, Nick & "> " & Dialogo, 0, 201, 197, False, True, False)

        Case 192192192 'Gris
            Call AddtoRichTextBox(frmMain.RecTxt, Nick & "> " & Dialogo, 164, 164, 164, False, True, False)

        Case 15722620 'Privado
            Call AddtoRichTextBox(frmMain.RecTxt, Nick & "> " & Dialogo, 157, 226, 20, False, True, False)

    End Select

    
    Exit Sub

CopiarDialogoToConsola_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.CopiarDialogoToConsola", Erl)
    Resume Next
    
End Sub

Public Function PonerPuntos(Numero As Long) As String
    
    On Error GoTo PonerPuntos_Err
    

    Dim i     As Integer

    Dim Cifra As String
 
    Cifra = str(Numero)
    Cifra = Right$(Cifra, Len(Cifra) - 1)

    For i = 0 To 4

        If Len(Cifra) - 3 * i >= 3 Then
            If mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) <> "" Then
                PonerPuntos = mid$(Cifra, Len(Cifra) - (2 + 3 * i), 3) & "." & PonerPuntos

            End If

        Else

            If Len(Cifra) - 3 * i > 0 Then
                PonerPuntos = Left$(Cifra, Len(Cifra) - 3 * i) & "." & PonerPuntos

            End If

            Exit For

        End If

    Next
 
    PonerPuntos = Left$(PonerPuntos, Len(PonerPuntos) - 1)
 
    
    Exit Function

PonerPuntos_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.PonerPuntos", Erl)
    Resume Next
    
End Function

Sub AmbientarAudio(ByVal UserMap As Long)
    
    On Error GoTo AmbientarAudio_Err
    

    

    Dim wav As Integer

    If EsNoche Then
   
        wav = ReadField(1, Val(MapDat.ambient), Asc("-"))

        If Sound.AmbienteActual <> wav Then
            Sound.LastAmbienteActual = wav
        End If
         
        Sound.Ambient_Play

        If wav = 0 Then
            Sound.Ambient_Stop
        End If

        '  AmbientalesBufferIndex = Audio.PlayWave(Wav & ".wav", , , LoopStyle.Enabled)
    Else
   
        wav = ReadField(2, Val(MapDat.extra1), Asc("-"))

        If wav = 0 Then Exit Sub
        If Sound.AmbienteActual <> wav Then
            Sound.LastAmbienteActual = wav

        End If

        If wav = 0 Then
            Sound.Ambient_Stop

        End If

        '  AmbientalesBufferIndex = Audio.PlayWave(Wav & ".wav", , , LoopStyle.Enabled)
    End If

    Sound.Ambient_Volume_Set VolAmbient
    'Debug.Print VolAmbient

    
    Exit Sub

AmbientarAudio_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.AmbientarAudio", Erl)
    Resume Next
    
End Sub

Public Function General_Var_Get(ByVal File As String, ByVal Main As String, ByVal Var As String) As String
    
    On Error GoTo General_Var_Get_Err
    

    '*****************************************************************
    'Author: Aaron Perkins
    'Last Modify Date: 10/07/2002
    'Get a var to from a text file
    '*****************************************************************
    Dim L        As Long

    Dim Char     As String

    Dim sSpaces  As String 'Input that the program will retrieve

    Dim szReturn As String 'Default value if the string is not found
    
    szReturn = ""
    
    sSpaces = Space$(5000)
    
    getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), File
    
    General_Var_Get = RTrim$(sSpaces)
    General_Var_Get = Left$(General_Var_Get, Len(General_Var_Get) - 1)

    
    Exit Function

General_Var_Get_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.General_Var_Get", Erl)
    Resume Next
    
End Function

Public Sub DibujarMiniMapa()
    
    On Error GoTo DibujarMiniMapa_Err
    

    Dim map_x   As Long, map_y As Long

    Dim termine As Boolean

    frmMain.MiniMap.BackColor = vbBlack

    For map_y = 1 To 100
        For map_x = 1 To 100

            If MapData(map_x, map_y).Graphic(1).GrhIndex > 0 Then
                SetPixel frmMain.MiniMap.hdc, map_x - 1, map_y - 1, GrhData(MapData(map_x, map_y).Graphic(1).GrhIndex).MiniMap_color

            End If

            If MapData(map_x, map_y).Graphic(2).GrhIndex > 0 Then
                SetPixel frmMain.MiniMap.hdc, map_x - 1, map_y - 1, GrhData(MapData(map_x, map_y).Graphic(2).GrhIndex).MiniMap_color

            End If

            If MapData(map_x, map_y).Graphic(4).GrhIndex > 0 Then
                SetPixel frmMain.MiniMap.hdc, map_x - 1, map_y - 1, GrhData(MapData(map_x, map_y).Graphic(4).GrhIndex).MiniMap_color

            End If
            
        Next map_x
    Next map_y
     
    frmMain.MiniMap.Refresh

    
    Exit Sub

DibujarMiniMapa_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.DibujarMiniMapa", Erl)
    Resume Next
    
End Sub

Rem Encripta una cadena de caracteres.
Rem S = Cadena a encriptar
Rem P = Password
Function EncryptStr(ByVal s As String, ByVal p As String) As String
    
    On Error GoTo EncryptStr_Err
    

    Dim i  As Integer, r As String

    Dim c1 As Integer, C2 As Integer

    r = ""

    If Len(p) > 0 Then

        For i = 1 To Len(s)
            c1 = Asc(mid(s, i, 1))

            If i > Len(p) Then
                C2 = Asc(mid(p, i Mod Len(p) + 1, 1))
            Else
                C2 = Asc(mid(p, i, 1))

            End If

            c1 = c1 + C2 + 64

            If c1 > 255 Then c1 = c1 - 256
            r = r + Chr(c1)
        Next i

    Else
        r = s

    End If

    EncryptStr = r

    
    Exit Function

EncryptStr_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.EncryptStr", Erl)
    Resume Next
    
End Function

Rem Desencripta una cadena de caracteres.
Rem S = Cadena a desencriptar
Rem P = Password
Function UnEncryptStr(ByVal s As String, ByVal p As String) As String
    
    On Error GoTo UnEncryptStr_Err
    

    Dim i  As Integer, r As String

    Dim c1 As Integer, C2 As Integer

    r = ""

    If Len(p) > 0 Then

        For i = 1 To Len(s)
            c1 = Asc(mid(s, i, 1))

            If i > Len(p) Then
                C2 = Asc(mid(p, i Mod Len(p) + 1, 1))
            Else
                C2 = Asc(mid(p, i, 1))

            End If

            c1 = c1 - C2 - 64

            If Sgn(c1) = -1 Then c1 = 256 + c1
            r = r + Chr(c1)
        Next i

    Else
        r = s

    End If

    UnEncryptStr = r

    
    Exit Function

UnEncryptStr_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.UnEncryptStr", Erl)
    Resume Next
    
End Function

Public Function Input_Key_Get(ByVal key_code As Byte) As Boolean
    '**************************************************************
    'Author: Aaron Perkins - Juan Martín Sotuyo Dodero
    'Modified by Augusto José Rando
    'Now we use DirectInput Keyboard
    'Last Modify Date: 10/07/2002
    '
    '**************************************************************
    'Input_Key_Get = (key_state.Key(key_code) > 0)
    
    On Error GoTo Input_Key_Get_Err
    
    Input_Key_Get = (GetKeyState(key_code) < 0)

    
    Exit Function

Input_Key_Get_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.Input_Key_Get", Erl)
    Resume Next
    
End Function

Public Function Input_Click_Get(ByVal Botton As Byte) As Boolean
    '**************************************************************
    'Author: Pablo Mercavides
    'Modified by Augusto José Rando
    'Now we use DirectInput Keyboard
    'Last Modify Date: 10/07/2002
    '
    '**************************************************************
    'Input_Key_Get = (key_state.Key(key_code) > 0)
    
    On Error GoTo Input_Click_Get_Err
    
    Input_Click_Get = (GetAsyncKeyState(Botton) < 0)

    
    Exit Function

Input_Click_Get_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.Input_Click_Get", Erl)
    Resume Next
    
End Function

Public Function General_Get_Temp_Dir() As String
    
    On Error GoTo General_Get_Temp_Dir_Err
    

    '**************************************************************
    'Author: Augusto José Rando
    'Last Modify Date: 6/11/2005
    'Gets windows temporary directory
    '**************************************************************
    Dim s As String

    Dim c As Long

    s = Space$(MAX_LENGTH)
    c = GetTempPath(MAX_LENGTH, s)

    If c > 0 Then
        If c > Len(s) Then
            s = Space$(c + 1)
            c = GetTempPath(MAX_LENGTH, s)

        End If

    End If

    General_Get_Temp_Dir = IIf(c > 0, Left$(s, c), "")

    
    Exit Function

General_Get_Temp_Dir_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.General_Get_Temp_Dir", Erl)
    Resume Next
    
End Function

Public Function General_Get_Mouse_Speed() As Long
    '**************************************************************
    'Author: Augusto José Rando
    'Last Modify Date: 6/11/2005
    '
    '**************************************************************
    
    On Error GoTo General_Get_Mouse_Speed_Err
    
 
    SystemParametersInfo SPI_GETMOUSESPEED, 0, General_Get_Mouse_Speed, 0
 
    
    Exit Function

General_Get_Mouse_Speed_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.General_Get_Mouse_Speed", Erl)
    Resume Next
    
End Function
 
Public Sub General_Set_Mouse_Speed(ByVal lngSpeed As Long)
    '**************************************************************
    'Author: Augusto José Rando
    'Last Modify Date: 6/11/2005
    '
    '**************************************************************
    
    On Error GoTo General_Set_Mouse_Speed_Err
    
 
    SystemParametersInfo SPI_SETMOUSESPEED, 0, ByVal lngSpeed, 0
 
    
    Exit Sub

General_Set_Mouse_Speed_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.General_Set_Mouse_Speed", Erl)
    Resume Next
    
End Sub

Public Sub ResetearUserMacro()
    '**************************************************************
    'Author: Augusto José Rando
    'Last Modify Date: 6/11/2005
    '
    '**************************************************************
    
    On Error GoTo ResetearUserMacro_Err
    
    Call WriteFlagTrabajar
    frmMain.MacroLadder.Enabled = False
    UserMacro.Activado = False
    UserMacro.cantidad = 0
    UserMacro.Index = 0
    UserMacro.Intervalo = 0
    UserMacro.TIPO = 0
    UserMacro.tX = 0
    UserMacro.tY = 0
    UserMacro.Skill = 0

    If UsingSkill <> 0 Then
        UsingSkill = 0
        Call FormParser.Parse_Form(frmMain)

    End If

    AddtoRichTextBox frmMain.RecTxt, "Has dejado de trabajar.", 223, 51, 2, 1, 0

    
    Exit Sub

ResetearUserMacro_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.ResetearUserMacro", Erl)
    Resume Next
    
End Sub

Public Sub CargarLst()
    
    On Error GoTo CargarLst_Err
    

    Dim i As Integer

    FrmLogear.lstServers.Clear

    For i = 1 To CantServer
        FrmLogear.lstServers.AddItem ServersLst(i).desc
    Next i
    
#If DEBUGGING = 1 Then
    FrmLogear.lstServers.ListIndex = Val(ServerIndex)
#Else
    FrmLogear.lstServers.ListIndex = 1
#End If

    
    Exit Sub

CargarLst_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.CargarLst", Erl)
    Resume Next
    
End Sub

Public Sub CrearFantasma(ByVal charindex As Integer)
    
    On Error GoTo CrearFantasma_Err
    

    If charlist(charindex).Body.Walk(charlist(charindex).Heading).GrhIndex = 0 Then Exit Sub

    MapData(charlist(charindex).Pos.x, charlist(charindex).Pos.y).CharFantasma.Body.GrhIndex = charlist(charindex).Body.Walk(charlist(charindex).Heading).GrhIndex
    MapData(charlist(charindex).Pos.x, charlist(charindex).Pos.y).CharFantasma.Head.GrhIndex = charlist(charindex).Head.Head(charlist(charindex).Heading).GrhIndex
    MapData(charlist(charindex).Pos.x, charlist(charindex).Pos.y).CharFantasma.Arma.GrhIndex = charlist(charindex).Arma.WeaponWalk(charlist(charindex).Heading).GrhIndex
    MapData(charlist(charindex).Pos.x, charlist(charindex).Pos.y).CharFantasma.Casco.GrhIndex = charlist(charindex).Casco.Head(charlist(charindex).Heading).GrhIndex
    MapData(charlist(charindex).Pos.x, charlist(charindex).Pos.y).CharFantasma.Escudo.GrhIndex = charlist(charindex).Escudo.ShieldWalk(charlist(charindex).Heading).GrhIndex
    MapData(charlist(charindex).Pos.x, charlist(charindex).Pos.y).CharFantasma.Body_Aura = charlist(charindex).Body_Aura
    MapData(charlist(charindex).Pos.x, charlist(charindex).Pos.y).CharFantasma.AlphaB = 255
    MapData(charlist(charindex).Pos.x, charlist(charindex).Pos.y).CharFantasma.Activo = True
    MapData(charlist(charindex).Pos.x, charlist(charindex).Pos.y).CharFantasma.OffX = charlist(charindex).Body.HeadOffset.x
    MapData(charlist(charindex).Pos.x, charlist(charindex).Pos.y).CharFantasma.Offy = charlist(charindex).Body.HeadOffset.y
    MapData(charlist(charindex).Pos.x, charlist(charindex).Pos.y).CharFantasma.Heading = charlist(charindex).Heading

    
    Exit Sub

CrearFantasma_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.CrearFantasma", Erl)
    Resume Next
    
End Sub

Public Sub CompletarAccionBarra(ByVal BarAccion As Byte)
    
    On Error GoTo CompletarAccionBarra_Err
    

    If BarAccion = Accion_Barra.CancelarAccion Then Exit Sub

    Call WriteCompletarAccion(BarAccion)

    
    Exit Sub

CompletarAccionBarra_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.CompletarAccionBarra", Erl)
    Resume Next
    
End Sub

Public Sub ComprobarEstado()
    
    On Error GoTo ComprobarEstado_Err
    

    Call InitServersList(RawServersList)

    Call CargarLst

    
    Exit Sub

ComprobarEstado_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.ComprobarEstado", Erl)
    Resume Next
    
End Sub

Public Function General_Distance_Get(ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer) As Integer
    
    On Error GoTo General_Distance_Get_Err
    
    General_Distance_Get = Abs(x1 - x2) + Abs(y1 - y2)

    
    Exit Function

General_Distance_Get_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.General_Distance_Get", Erl)
    Resume Next
    
End Function

Public Sub EndGame(Optional ByVal Closed_ByUser As Boolean = False, Optional ByVal Init_Launcher As Boolean = False)
    
    On Error GoTo EndGame_Err
    

    

    Sound.Engine_DeInitialize
    Sound.Music_Stop
    Set Sound = Nothing

    prgRun = False

    '0. Cerramos el socket
    If frmMain.Socket1.State <> sckClosed Then frmMain.Socket1.Disconnect

    '2. Eliminamos objetos DX
    Call Client_UnInitialize_DirectX_Objects

    '6. Cerramos los forms y nos vamos
    Call UnloadAllForms

    '7. Adiós MuteX - Restauramos MouseSpeed

    End

    
    Exit Sub

EndGame_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.EndGame", Erl)
    Resume Next
    
End Sub

Public Sub Client_UnInitialize_DirectX_Objects()
    
    On Error GoTo Client_UnInitialize_DirectX_Objects_Err
    

    

    '1. Cerramos el engine de sonido y borramos buffers
    Sound.Engine_DeInitialize
    Set Sound = Nothing

    '2. Cerramos el engine gráfico y borramos textures

    
    Exit Sub

Client_UnInitialize_DirectX_Objects_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.Client_UnInitialize_DirectX_Objects", Erl)
    Resume Next
    
End Sub

Public Sub TextoAlAsistente(ByVal Texto As String)
    
    On Error GoTo TextoAlAsistente_Err
    
    TextEfectAsistente = 35
    TextAsistente = Texto

    
    Exit Sub

TextoAlAsistente_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.TextoAlAsistente", Erl)
    Resume Next
    
End Sub

Public Function GetTimeFormated(Mins As Integer) As String
    
    On Error GoTo GetTimeFormated_Err
    

    Dim Horita    As Byte

    Dim Minutitos As Byte

    Dim A         As String

    Horita = Fix(Mins / 60)
    Minutitos = Mins - 60 * Horita

    If Minutitos < 10 Then

        If Horita < 10 Then
            GetTimeFormated = "0" & Horita & ":0" & Minutitos
        Else
            GetTimeFormated = Horita & ":0" & Minutitos

        End If
    
    Else

        If Horita < 10 Then
            GetTimeFormated = "0" & Horita & ":" & Minutitos
        Else
            GetTimeFormated = Horita & ":" & Minutitos

        End If
    
    End If

    
    Exit Function

GetTimeFormated_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.GetTimeFormated", Erl)
    Resume Next
    
End Function

Public Function GetHora(Mins As Integer) As String
    
    On Error GoTo GetHora_Err
    

    Dim Horita As Byte

    Horita = Fix(Mins / 60)

    GetHora = Horita

    
    Exit Function

GetHora_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.GetHora", Erl)
    Resume Next
    
End Function

Public Sub PreloadGraphics()
    
    On Error GoTo PreloadGraphics_Err
    

    Dim PreloadFile   As String

    Dim strPreload    As String

    Dim NumPreload    As Integer
    
    Dim i             As Integer

    Dim j             As Integer
    
    Dim MinVal        As Integer

    Dim MaxVal        As Integer

    Dim Priority      As Byte
    
    Dim TotalPreloads As Integer
    
    #If Compresion = 1 Then

        If Not Extract_File(Scripts, App.Path & "\..\Recursos\OUTPUT\", "preload.ind", Windows_Temp_Dir, False) Then
            Err.Description = "No se ha logrado extraer el archivo de recurso."
            GoTo ErrorHandler

        End If
    
        PreloadFile = Windows_Temp_Dir & "Preload.ind"
    #Else
        PreloadFile = App.Path & "\..\Recursos\init\Preload.ind"
    #End If
    
    TotalPreloads = Val(General_Var_Get(PreloadFile, "GRAPHICS", "TotalPreloads"))

    If TotalPreloads = 0 Then TotalPreloads = 1
    
    NumPreload = Val(General_Var_Get(PreloadFile, "GRAPHICS", "NumGraphics"))
    
    For i = 1 To NumPreload
        strPreload = General_Var_Get(PreloadFile, "GRAPHICS", str(i))
        MinVal = Val(General_Field_Read(1, strPreload, "-"))
        MaxVal = Val(General_Field_Read(2, strPreload, "-"))
        Priority = Val(General_Field_Read(3, strPreload, "-"))
        
        For j = MinVal To MaxVal

            Static d3dTextures As D3D8Textures

            Set d3dTextures.Texture = SurfaceDB.GetTexture(j, d3dTextures.texwidth, d3dTextures.texheight)
            'Call SurfaceDB.GetTexture(j, 1024, 1024)
            DoEvents
        Next j
 
    Next i
    
    #If Compresion = 1 Then
        Delete_File Windows_Temp_Dir & "Preload.ind"
    #End If
    
    Exit Sub
    
ErrorHandler:
    '  If General_File_Exists(Windows_Temp_Dir & "Preload.ind", vbNormal) Then Delete_File Windows_Temp_Dir & "Preload.ind"

    
    Exit Sub

PreloadGraphics_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.PreloadGraphics", Erl)
    Resume Next
    
End Sub

Public Sub CalcularPosicionMAPA()
    
    On Error GoTo CalcularPosicionMAPA_Err
    
    frmMapaGrande.lblMapInfo(0) = MapDat.map_name & "(" & UserMap & ")"

    If NameMaps(UserMap).desc <> "" Then
        frmMapaGrande.Label1.Caption = NameMaps(UserMap).desc
    Else
        frmMapaGrande.Label1.Caption = "Sin información relevante."

    End If

    Dim i       As Integer
    Dim j       As Byte

    Dim Encontre As Boolean
    
    
    For j = 1 To TotalWorlds
        For i = 1 To Mundo(j).Ancho * Mundo(j).Alto
    
            If Mundo(j).MapIndice(i) = UserMap Then
                idmap = i
                Encontre = True
                frmMapaGrande.picMap.Picture = LoadInterface("mapa" & j & ".bmp")
                frmMapaGrande.Image2.Picture = Nothing
                WorldActual = j
                frmMapaGrande.ComMundo.ListIndex = j - 1
                Exit For
            End If
        Next i
        
        If Encontre Then
            Exit For
        End If
    Next j
    
    If Encontre = False Then
        If frmMapaGrande.Visible = False Then
            frmMapaGrande.picMap.Picture = LoadInterface("mapa1.bmp")
            frmMapaGrande.Image2.Picture = Nothing
        End If

    End If
    
    Call CargarDatosMapa(UserMap)

    Dim x As Long

    Dim y As Long

    x = (idmap - 1) Mod 16
    y = Int((idmap - 1) / 16)

    frmMapaGrande.lblAllies.Top = y * 27
    frmMapaGrande.lblAllies.Left = x * 27

    frmMapaGrande.Shape1.Top = y * 27 + (UserPos.y / 4.5)
    frmMapaGrande.Shape1.Left = x * 27 + (UserPos.x / 4.5)

    
    Exit Sub

CalcularPosicionMAPA_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.CalcularPosicionMAPA", Erl)
    Resume Next
    
End Sub

Public Function GetDriveSerialNumber(Optional ByVal DriveLetter As String) As Long
    
    On Error GoTo GetDriveSerialNumber_Err
    

    '***************************************************
    'Author: Nahuel Casas (Zagen)
    'Last Modify Date: 07/12/2009
    ' 07/12/2009: Zagen - Convertì las funciones, en formulas mas fàciles de modificar.
    '***************************************************
    

    Dim fso As Object, Drv As Object, DriveSerial As Long
         
    'Creamos el objeto FileSystemObject.
    Set fso = CreateObject("Scripting.FileSystemObject")
         
    'Asignamos el driver principal.
    If DriveLetter <> "" Then
        Set Drv = fso.GetDrive(DriveLetter)
    Else
        Set Drv = fso.GetDrive(fso.GetDriveName(App.Path))

    End If
     
    With Drv

        If .IsReady Then
            DriveSerial = Abs(.SerialNumber)
        Else    '"Si el driver no està como para empezar ..."
            DriveSerial = -1

        End If

    End With
         
    'Borramos y limpiamos.
    Set Drv = Nothing
    Set fso = Nothing
    'Seteamos :)
    GetDriveSerialNumber = DriveSerial
         
    
    Exit Function

GetDriveSerialNumber_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.GetDriveSerialNumber", Erl)
    Resume Next
    
End Function

Public Function GetMacAddress() As String
    
    On Error GoTo GetMacAddress_Err
    

    Const OFFSET_LENGTH As Long = 400

    Dim lSize           As Long

    Dim baBuffer()      As Byte

    Dim lIdx            As Long

    Dim sRetVal         As String
    
    Call GetAdaptersInfo(ByVal 0, lSize)

    If lSize <> 0 Then
        ReDim baBuffer(0 To lSize - 1) As Byte
        Call GetAdaptersInfo(baBuffer(0), lSize)
        Call CopyMemory(lSize, baBuffer(OFFSET_LENGTH), 4)

        For lIdx = OFFSET_LENGTH + 4 To OFFSET_LENGTH + 4 + lSize - 1
            sRetVal = IIf(LenB(sRetVal) <> 0, sRetVal & ":", vbNullString) & Right$("0" & Hex$(baBuffer(lIdx)), 2)
        Next

    End If

    GetMacAddress = sRetVal

    
    Exit Function

GetMacAddress_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.GetMacAddress", Erl)
    Resume Next
    
End Function
Public Function ObtenerIdMapaDeLlamadaDeClan(ByVal Mapa As Integer) As Integer
    
    On Error GoTo ObtenerIdMapaDeLlamadaDeClan_Err
    

    Dim i        As Integer

    Dim Encontre As Boolean

    For i = 1 To WordMapaNum

        If WordMapa(i) = Mapa Then
            ObtenerIdMapaDeLlamadaDeClan = i
            frmMapaGrande.llamadadeclan.Tag = 0
            Exit Function
            Encontre = True

            PosREAL = 1
            Exit For

        End If

    Next i
    
    If Encontre = False Then
    
        For i = 1 To DungeonDataNum

            If DungeonData(i) = Mapa Then
                frmMapaGrande.llamadadeclan.Tag = 1
                ObtenerIdMapaDeLlamadaDeClan = i
                Exit Function
                Encontre = True
                PosREAL = 0
                Dungeon = True
                Exit For

            End If

        Next i

    End If

    ObtenerIdMapaDeLlamadaDeClan = 0

    
    Exit Function

ObtenerIdMapaDeLlamadaDeClan_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.ObtenerIdMapaDeLlamadaDeClan", Erl)
    Resume Next
    
End Function

Public Sub Auto_Drag(ByVal hWnd As Long)
    
    On Error GoTo Auto_Drag_Err
    
    Call ReleaseCapture
    Call SendMessage(hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&)

    
    Exit Sub

Auto_Drag_Err:
    Call RegistrarError(Err.number, Err.Description, "ModLadder.Auto_Drag", Erl)
    Resume Next
    
End Sub
