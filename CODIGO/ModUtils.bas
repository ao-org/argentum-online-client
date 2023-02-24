Attribute VB_Name = "ModUtils"
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

Public StopCreandoCuenta    As Boolean

Public Const DegreeToRadian As Single = 0.01745329251994 'Pi / 180
Public Const RadianToDegree As Single = 57.2958279087977 '180 / Pi

'Nueva seguridad
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal Length As Long)
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'get mac adress

Public Type Tclan
    nombre As String
    Alineacion As Byte
    indice As Integer
End Type

Public ListaClanes      As Boolean
Public ClanesList()     As Tclan

Public CheckMD5         As String

Public intro            As Byte

Public InviCounter      As Integer
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

Public HayLayer4     As Boolean

Public CantPartLLuvia     As Integer
Public MeteoIndex         As Integer

'Dropeo
Public PingRender         As Integer

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
Public ObjShop()          As ObjDatas
Public NpcData()          As NpcDatas
Public ProjectileData() As t_Projectile
Public GProjectile As Projectile

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
    MaderaElfica As Integer
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
    Agarrable As Boolean
    Llave As Integer
    ObjNum As Long
    Cooldown As Long
    CDType As Integer
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
    PuedeInvocar As Byte
    
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

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Public Type COPYDATASTRUCT

    dwData As Long
    cbData As Long
    lpData As Long

End Type

Public Const WM_COPYDATA = &H4A

Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long

Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

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

Public Pjs(1 To MAX_PERSONAJES_EN_CUENTA)       As UserCuentaPJS

Public RecordarCuenta               As Boolean

Public CuentaRecordada              As CuentasGuardadas

Public CantidadDePersonajesEnCuenta As Byte

Type UserCuentaPJS

    nombre As String
    nivel As Byte
    Mapa As Integer
    posX As Integer
    posY As Integer
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

Public Type Projectile
    GrhIndex As Long
    CurrentPos As Vector2
    speed As Single
    RotationSpeed As Single
    TargetPos As Vector2
    Rotation As Single
End Type

Public Type t_IndexHeap
    CurrentIndex As Integer
    IndexInfo() As Integer
End Type



Public Const InitialProjectileSize As Integer = 45
Public AllProjectile(InitialProjectileSize) As Projectile
Public AvailableProjectile As t_IndexHeap
Public ActiveProjectile As t_IndexHeap

Const HTCAPTION = 2

Const WM_NCLBUTTONDOWN = &HA1

Const RGN_OR = 2
 
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long

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

Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Declare Function CreateIconFromResourceEx Lib "user32.dll" (ByRef presbits As Any, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal flags As Long) As Long

Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (lpPictDesc As Any, riid As Any, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long

Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long

Public Declare Function DrawIconEx Lib "user32.dll" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
  
Private Declare Function SendMessageLongRef Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Long) As Long
                           
Private m_ASC As Long

Public Sub InitilializeProjectiles()
    Dim i As Integer
    ReDim AvailableProjectile.IndexInfo(InitialProjectileSize)
    ReDim ActiveProjectile.IndexInfo(InitialProjectileSize)
    ActiveProjectile.CurrentIndex = 0
    For i = 1 To InitialProjectileSize
        AvailableProjectile.IndexInfo(i) = InitialProjectileSize - i + 1
    Next i
    AvailableProjectile.CurrentIndex = InitialProjectileSize
End Sub

Public Sub ReleaseProjectile(Index As Integer)
    AvailableProjectile.CurrentIndex = AvailableProjectile.CurrentIndex + 1
    AvailableProjectile.IndexInfo(AvailableProjectile.CurrentIndex) = ActiveProjectile.IndexInfo(Index)
    ActiveProjectile.IndexInfo(Index) = ActiveProjectile.IndexInfo(ActiveProjectile.CurrentIndex - 1)
    ActiveProjectile.CurrentIndex = ActiveProjectile.CurrentIndex - 1
End Sub

Sub inputbox_Password(El_Form As Form, Caracter As String)
    
    On Error GoTo inputbox_Password_Err
    
      
    m_ASC = Asc(Caracter)
      
    Call SetTimer(El_Form.hwnd, &H5000&, 100, AddressOf TimerProc)
  
    
    Exit Sub

inputbox_Password_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.inputbox_Password", Erl)
    Resume Next
    
End Sub
  
Private Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    
    On Error GoTo TimerProc_Err
    
           
    Dim Handle_InputBox As Long
      
    'Captura el handle del textBox del InputBox
    Handle_InputBox = FindWindowEx(FindWindow("#32770", App.title), 0, "Edit", "")
                  
    'Le establece el PasswordChar
    Call SendMessageLongRef(Handle_InputBox, &HCC&, m_ASC, 0)
    'Finaliza el Timer
    Call KillTimer(hwnd, idEvent)
  
    
    Exit Sub

TimerProc_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.TimerProc", Erl)
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
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.LoadPNGtoICO", Erl)
    Resume Next
    
End Function

Public Function SetTopMostWindow(hwnd As Long, Topmost As Boolean) As Long
    
    On Error GoTo SetTopMostWindow_Err
    

    If Topmost = True Then 'Make the window topmost
        SetTopMostWindow = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGSz)
    Else
        SetTopMostWindow = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGSz)
        SetTopMostWindow = False

    End If

    
    Exit Function

SetTopMostWindow_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.SetTopMostWindow", Erl)
    Resume Next
    
End Function

Public Sub LogError(desc As String)

    On Error GoTo ErrHandler

    Dim nfile As Integer

    nfile = FreeFile ' obtenemos un canal
    Open App.path & "\errores.log" For Append Shared As #nfile
    Print #nfile, Date & "-" & Time & ":" & desc
    Close #nfile

    Exit Sub

ErrHandler:

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
     frmCrearPersonaje.lstRaza.ListIndex = 0
    
    frmCrearPersonaje.lstHogar.Clear
    For i = LBound(ListaCiudades()) To UBound(ListaCiudades())
        frmCrearPersonaje.lstHogar.AddItem (ListaCiudades(i))
    Next i
     frmCrearPersonaje.lstHogar.ListIndex = 0

    frmCrearPersonaje.lstProfesion.Clear

    For i = LBound(ListaClases()) To UBound(ListaClases())
        frmCrearPersonaje.lstProfesion.AddItem ListaClases(i)
    
    Next i
    frmCrearPersonaje.lstProfesion.ListIndex = 1
    
    
        
    MiCabeza = Val(frmCrearPersonaje.Cabeza.List(1))
    Call DibujarCPJ(MiCabeza, 3)
    CPHead = MiCabeza
    
    Exit Sub

IniciarCrearPj_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.IniciarCrearPj", Erl)
    Resume Next
    
End Sub

Sub General_Set_Connect()
    
    On Error GoTo General_Set_Connect_Err
    

    AlphaNiebla = 75
    EntradaY = 10
    EntradaX = 10
    
    UserMap = randomMap()
    Call SwitchMap(UserMap)

    If g_game_state.state() <> e_state_connect_screen Then
        frmConnect.Show
        FrmLogear.Show , frmConnect
    End If
            
    intro = 1
    frmMain.Picture = LoadInterface("ventanaprincipal.bmp")
    frmMain.panelInf.Picture = LoadInterface("ventanaprincipal_stats.bmp")
    frmMain.panel.Picture = LoadInterface("centroinventario.bmp")
    frmMain.EXPBAR.Picture = LoadInterface("barraexperiencia.bmp")
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
    mFadingMusicMod = 0
    CurMp3 = 1
    
     g_game_state.state = e_state_connect_screen
    ClickEnAsistente = 0
    If CuentaRecordada.nombre <> "" Then
        Call TextoAlAsistente("¡Bienvenido de nuevo! ¡Disfruta tu viaje por Argentum20!") ' hay que poner 20 aniversario
    Else
        Call TextoAlAsistente("¡Bienvenido a Argentum20! ¿Ya tenes tu cuenta? Logea! sino, toca sobre Cuenta para crearte una.") ' hay que poner 20 aniversario

    End If
    
    engine.FadeInAlpha = 255
    
    Exit Sub

General_Set_Connect_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.General_Set_Connect", Erl)
    Resume Next
    
End Sub
 
Public Sub InitializeSurfaceCapture(Frm As Form)
    
    On Error GoTo InitializeSurfaceCapture_Err
    
    lRegion = CreateRectRgn(0, 0, 0, 0)
    Frm.visible = False

    
    Exit Sub

InitializeSurfaceCapture_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.InitializeSurfaceCapture", Erl)
    Resume Next
    
End Sub

Public Sub ReleaseSurfaceCapture(Frm As Form)
    
    On Error GoTo ReleaseSurfaceCapture_Err
    
    ApplySurfaceTo Frm
    Frm.visible = True
    Call DeleteObject(lRegion)

    
    Exit Sub

ReleaseSurfaceCapture_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.ReleaseSurfaceCapture", Erl)
    Resume Next
    
End Sub
 
Public Sub ApplySurfaceTo(Frm As Form)
    
    On Error GoTo ApplySurfaceTo_Err
    
    Call SetWindowRgn(Frm.hwnd, lRegion, True)

    
    Exit Sub

ApplySurfaceTo_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.ApplySurfaceTo", Erl)
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
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.CreateSurfacefromPoints", Erl)
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
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.CreateSurfacefromEllipse", Erl)
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
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.CreateSurfacefromRect", Erl)
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
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.CreateSurfacefromMask", Erl)
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
        glHeight = .Height / screen.TwipsPerPixelY
        glWidth = .Width / screen.TwipsPerPixelX

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
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.CreateSurfacefromMask_GetPixel", Erl)
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
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.General_Var_Write", Erl)
    Resume Next
    
End Sub

Public Sub MensajeAdvertencia(ByVal mensaje As String)
    
    On Error GoTo MensajeAdvertencia_Err
    
    Call MsgBox(mensaje, vbInformation + vbOKOnly, "Advertencia")

    
    Exit Sub

MensajeAdvertencia_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.MensajeAdvertencia", Erl)
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
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.ReproducirMp3", Erl)
    Resume Next
    
End Sub

Public Sub ForzarMp3(ByVal mp3 As Byte)
    
    On Error GoTo ForzarMp3_Err
    

    If mp3 = 0 Then Exit Sub

    mFadingMusicMod = 0
    CurMp3 = mp3

    
    Exit Sub

ForzarMp3_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.ForzarMp3", Erl)
    Resume Next
    
End Sub

Public Sub CargarCuentasGuardadas()

    Dim Arch As String
        Arch = App.path & "\..\Recursos\OUTPUT\Cuenta.ini"
    
    CuentaRecordada.nombre = GetVar(Arch, "CUENTA", "Nombre")
    CuentaRecordada.Password = UnEncryptStr(GetVar(Arch, "CUENTA", "Password"), 9256)
    
    FrmLogear.chkRecordar.Tag = "0"
 
    If LenB(CuentaRecordada.nombre) <> 0 Then
        FrmLogear.NameTxt = CuentaRecordada.nombre
        FrmLogear.PasswordTxt = CuentaRecordada.Password
        FrmLogear.chkRecordar.Picture = LoadInterface("check-amarillo.bmp")
        FrmLogear.chkRecordar.Tag = "1"
        FrmLogear.PasswordTxt.TabIndex = 0
        FrmLogear.PasswordTxt.SelStart = Len(FrmLogear.PasswordTxt)
    End If

End Sub

Public Sub GuardarCuenta(ByVal Name As String, ByVal Password As String)

    Dim Archivo As String
        Archivo = App.path & "\..\Recursos\OUTPUT\Cuenta.ini"
    
    ' Si el parametro Password no es vbNullString, encriptamos el string
    If LenB(Password) Then Password = EncryptStr(Password, 9256)
    
    Call WriteVar(Archivo, "CUENTA", "Nombre", Name)
    Call WriteVar(Archivo, "CUENTA", "Password", Password)
    
    Call CargarCuentasGuardadas

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
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.IntervaloPermiteClick", Erl)
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
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.IntervaloPermiteHeading", Erl)
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
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.IntervaloPermiteLLamadaClan", Erl)
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
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.IntervaloPermiteAnim", Erl)
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
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.IntervaloPermiteConectar", Erl)
    Resume Next
    
End Function
Sub initPacketControl()
    Dim i As Long, j As Long
    
    'ReDim packetControl(1 To CANT_PACKETS_CONTROL) As t_packetControl
    
    For i = 1 To CANT_PACKETS_CONTROL
        With packetControl(i)
            .last_count = 0
           ' .cant_iterations = 0
           For j = 1 To 10
                .iterations(j) = 0
            Next j
        End With
    Next i
    
End Sub
Sub load_game_settings()

    On Error GoTo ErrorHandler
    Set DialogosClanes = New clsGuildDlg
    
    If InitializeSettings() Then
        Call LoadImpAoInit
    Else
        Call MsgBox("¡No se puede cargar el archivo de opciones! La reinstalacion del juego podria solucionar el problema.", vbCritical, "Error al cargar")
        End
    End If
    'Musica y Sonido
    Musica = GetSetting("AUDIO", "Musica")
    Sonido = GetSetting("AUDIO", "Sonido")
    Fx = GetSetting("AUDIO", "Fx")
    AmbientalActivated = GetSetting("AUDIO", "AmbientalActivated")
    InvertirSonido = GetSetting("AUDIO", "InvertirSonido")
    
    'Musica y Sonido - Volumen
    VolMusicFadding = VolMusic
    VolMusic = Val(GetSetting("AUDIO", "VolMusic"))
    VolFX = Val(GetSetting("AUDIO", "VolFX"))
    VolAmbient = Val(GetSetting("AUDIO", "VolAmbient"))
    
    'Video
    PantallaCompleta = GetSetting("VIDEO", "PantallaCompleta")
    CursoresGraficos = IIf(RunningInVB, 0, GetSetting("VIDEO", "CursoresGraficos"))
    UtilizarPreCarga = GetSetting("VIDEO", "UtilizarPreCarga")
    InfoItemsEnRender = Val(GetSetting("VIDEO", "InfoItemsEnRender"))
    ModoAceleracion = GetSetting("VIDEO", "Aceleracion")
    
    Dim Value As String
    Value = GetSetting("VIDEO", "MostrarRespiracion")
    MostrarRespiracion = IIf(LenB(Value) > 0, Val(Value), True)

    FxNavega = GetSetting("OPCIONES", "FxNavega")
    MostrarIconosMeteorologicos = GetSetting("OPCIONES", "MostrarIconosMeteorologicos")
    CopiarDialogoAConsola = GetSetting("OPCIONES", "CopiarDialogoAConsola")
    PermitirMoverse = GetSetting("OPCIONES", "PermitirMoverse")
    ScrollArrastrar = Val(GetSetting("OPCIONES", "ScrollArrastrar"))
    LastScroll = Val(GetSetting("OPCIONES", "LastScroll"))
    
    MoverVentana = GetSetting("OPCIONES", "MoverVentana")
    FPSFLAG = GetSetting("OPCIONES", "FPSFLAG")
    AlphaMacro = GetSetting("OPCIONES", "AlphaMacro")
    ModoHechizos = Val(GetSetting("OPCIONES", "ModoHechizos"))
    DialogosClanes.Activo = Val(GetSetting("OPCIONES", "DialogosClanes"))
    NumerosCompletosInventario = Val(GetSetting("OPCIONES", "NumerosCompletosInventario"))
    
    'Init
    #If PYMMO = 0 Or DEBUGGING = 1 Then
        ServerIndex = GetSetting("INIT", "ServerIndex")
    #End If

    SensibilidadMouse = GetSetting("OPCIONES", "SensibilidadMouse")
    If SensibilidadMouse = 0 Then: SensibilidadMouse = 10
    SensibilidadMouseOriginal = General_Get_Mouse_Speed
    Call General_Set_Mouse_Speed(SensibilidadMouse)
    
    'Dialogos clanes
    Exit Sub
    
ErrorHandler:
    Call MsgBox("Ha ocurrido un error al cargar la configuración del juego.", vbCritical, "Configuración del Juego")
    End
End Sub

Sub GuardarOpciones()
    On Error GoTo GuardarOpciones_Err
    
    #If PYMMO = 0 Or DEBUGGING = 1 Then
    Call SaveSetting("INIT", "ServerIndex", IPdelServidor & ":" & PuertoDelServidor)
    #End If
    Call SaveSetting("AUDIO", "Musica", Musica)
    Call SaveSetting("AUDIO", "Fx", Fx)
    Call SaveSetting("AUDIO", "VolMusic", VolMusic)
    Call SaveSetting("AUDIO", "Volfx", VolFX)
    Call SaveSetting("AUDIO", "VolAmbient", VolAmbient)
    Call SaveSetting("AUDIO", "AmbientalActivated", AmbientalActivated)
    
    Call SaveSetting("OPCIONES", "MoverVentana", MoverVentana)
    Call SaveSetting("OPCIONES", "PermitirMoverse", PermitirMoverse)
    Call SaveSetting("OPCIONES", "ScrollArrastrar", ScrollArrastrar)
    
    Call SaveSetting("OPCIONES", "CopiarDialogoAConsola", CopiarDialogoAConsola)
    Call SaveSetting("OPCIONES", "FPSFLAG", FPSFLAG)
    Call SaveSetting("OPCIONES", "AlphaMacro", AlphaMacro)
    Call SaveSetting("OPCIONES", "ModoHechizos", ModoHechizos)
    Call SaveSetting("OPCIONES", "FxNavega", FxNavega)
    
    Call SaveSetting("OPCIONES", "NumerosCompletosInventario", NumerosCompletosInventario)

    Call SaveSetting("VIDEO", "MostrarRespiracion", IIf(MostrarRespiracion, 1, 0))
    Call SaveSetting("VIDEO", "PantallaCompleta", IIf(PantallaCompleta, 1, 0))
    Call SaveSetting("VIDEO", "InfoItemsEnRender", IIf(InfoItemsEnRender, 1, 0))
    Call SaveSetting("VIDEO", "Aceleracion", ModoAceleracion)

    Call SaveSetting("OPCIONES", "SensibilidadMouse", SensibilidadMouse)
    Call SaveSetting("OPCIONES", "DialogosClanes", IIf(DialogosClanes.Activo, 1, 0))

    
    Exit Sub

GuardarOpciones_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.GuardarOpciones", Erl)
    Resume Next
    
End Sub

Public Sub WriteConsoleUserChat(ByVal Text As String, ByVal userName As String, ByVal red As Byte, ByVal green As Byte, ByVal blue As Byte, ByVal userStatus As Integer, ByVal Privileges As Integer)
    Dim NameRed   As Byte
    Dim NameGreen As Byte
    Dim NameBlue  As Byte
    If Privileges > 0 Then
        NameRed = ColoresPJ(Privileges).r
        NameGreen = ColoresPJ(Privileges).G
        NameBlue = ColoresPJ(Privileges).b
    Else
        Select Case userStatus
            Case 0: ' Criminal
                NameRed = ColoresPJ(23).r
                NameGreen = ColoresPJ(23).G
                NameBlue = ColoresPJ(23).b
            Case 1: ' Ciudadano
                NameRed = ColoresPJ(20).r
                NameGreen = ColoresPJ(20).G
                NameBlue = ColoresPJ(20).b
            Case 2: ' Caos
                NameRed = ColoresPJ(24).r
                NameGreen = ColoresPJ(24).G
                NameBlue = ColoresPJ(24).b
            Case 3: ' Armada
                NameRed = ColoresPJ(21).r
                NameGreen = ColoresPJ(21).G
                NameBlue = ColoresPJ(21).b
            Case 4: ' Conciclio
                NameRed = ColoresPJ(25).r
                NameGreen = ColoresPJ(25).G
                NameBlue = ColoresPJ(25).b
            Case 5: ' Consejo
                NameRed = ColoresPJ(22).r
                NameGreen = ColoresPJ(22).G
                NameBlue = ColoresPJ(22).b
        End Select
    End If
    Dim Pos As Integer
    Pos = InStr(userName, "<")
    If Pos = 0 Then Pos = LenB(userName) + 2
    Dim name As String
    name = Left$(userName, Pos - 2)

    Text = Trim$(Text)
    If LenB(userName) <> 0 And LenB(Text) > 0 Then
        Call AddtoRichTextBox2(frmMain.RecTxt, "[" & name & "] ", NameRed, NameGreen, NameBlue, True, False, True, rtfLeft)
        Call AddtoRichTextBox2(frmMain.RecTxt, Text, red, green, blue, False, False, False, rtfLeft)
    End If
    Exit Sub
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
        Call WriteConsoleUserChat(ChatText, .nombre, red, green, blue, .status, .priv)
    End With
    Exit Sub
WriteChatOverHeadInConsole_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.WriteChatOverHeadInConsole", Erl)
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
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.PonerPuntos", Erl)
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
   
        wav = ReadField(2, Val(MapDat.ambient), Asc("-"))

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
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.AmbientarAudio", Erl)
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
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.General_Var_Get", Erl)
    Resume Next
    
End Function

Public Sub DibujarMiniMapa()
    
    On Error GoTo DibujarMiniMapa_Err

    frmMain.MiniMap.Picture = LoadMinimap(UserMap)
    'Pintamos los NPCs en Minimapa:
    If ListNPCMapData(UserMap, 1).NPCNumber > 0 Then
        Dim i As Long
        For i = 1 To MAX_QUESTNPCS_VISIBLE
            Dim posX As Long
            Dim posY As Long
            
            posX = (ListNPCMapData(UserMap, i).Position.x - HalfWindowTileWidth - 2) * (100 / (100 - 2 * HalfWindowTileWidth - 4)) - 2
            posY = (ListNPCMapData(UserMap, i).Position.y - HalfWindowTileHeight - 1) * (100 / (100 - 2 * HalfWindowTileHeight - 2)) - 1
            
            
            Dim color As Long
            
            Select Case ListNPCMapData(UserMap, i).state
                Case 1
                    color = RGB(0, 198, 254)
                Case 2
                    color = RGB(255, 201, 14)
                    Case Else
                    color = RGB(255, 201, 14)
            End Select
            
            
            
            Call SetPixel(frmMain.MiniMap.hdc, posX + 1, posY, color)
            Call SetPixel(frmMain.MiniMap.hdc, posX, posY + 1, color)
            Call SetPixel(frmMain.MiniMap.hdc, posX + 1, posY + 1, color)
            Call SetPixel(frmMain.MiniMap.hdc, posX, posY, color)
            
            Call SetPixel(frmMain.MiniMap.hdc, posX, posY - 1, &H808080)
            Call SetPixel(frmMain.MiniMap.hdc, posX + 1, posY - 1, &H808080)
            Call SetPixel(frmMain.MiniMap.hdc, posX + 2, posY, &H808080)
            Call SetPixel(frmMain.MiniMap.hdc, posX + 2, posY + 1, &H808080)
            Call SetPixel(frmMain.MiniMap.hdc, posX + 1, posY + 2, &H808080)
            Call SetPixel(frmMain.MiniMap.hdc, posX, posY + 2, &H808080)
            Call SetPixel(frmMain.MiniMap.hdc, posX - 1, posY + 1, &H808080)
            Call SetPixel(frmMain.MiniMap.hdc, posX - 1, posY, &H808080)

        Next i
        
        frmMain.MiniMap.Refresh
    End If
    Exit Sub

DibujarMiniMapa_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.DibujarMiniMapa", Erl)
    
End Sub

Rem Encripta una cadena de caracteres.
Rem S = Cadena a encriptar
Rem P = Password
Function EncryptStr(ByVal s As String, ByVal P As String) As String
    
    On Error GoTo EncryptStr_Err
    

    Dim i  As Integer, r As String

    Dim c1 As Integer, C2 As Integer

    r = ""

    If Len(P) > 0 Then

        For i = 1 To Len(s)
            c1 = Asc(mid(s, i, 1))

            If i > Len(P) Then
                C2 = Asc(mid(P, i Mod Len(P) + 1, 1))
            Else
                C2 = Asc(mid(P, i, 1))

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
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.EncryptStr", Erl)
    Resume Next
    
End Function

Rem Desencripta una cadena de caracteres.
Rem S = Cadena a desencriptar
Rem P = Password
Function UnEncryptStr(ByVal s As String, ByVal P As String) As String
    
    On Error GoTo UnEncryptStr_Err
    

    Dim i  As Integer, r As String

    Dim c1 As Integer, C2 As Integer

    r = ""

    If Len(P) > 0 Then

        For i = 1 To Len(s)
            c1 = Asc(mid(s, i, 1))

            If i > Len(P) Then
                C2 = Asc(mid(P, i Mod Len(P) + 1, 1))
            Else
                C2 = Asc(mid(P, i, 1))

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
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.UnEncryptStr", Erl)
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
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.Input_Key_Get", Erl)
    Resume Next
    
End Function

Public Function Input_Click_Get(ByVal Botton As Byte) As Boolean
 On Error GoTo Input_Click_Get_Err
    
    Input_Click_Get = (GetAsyncKeyState(Botton) < 0)

    
    Exit Function

Input_Click_Get_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.Input_Click_Get", Erl)
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
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.General_Get_Temp_Dir", Erl)
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
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.General_Get_Mouse_Speed", Erl)
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
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.General_Set_Mouse_Speed", Erl)
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
    frmMain.MacroLadder.enabled = False
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
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.ResetearUserMacro", Erl)
    Resume Next
    
End Sub

Public Sub CargarLst()
    
    On Error GoTo CargarLst_Err
    
        
    #If PYMMO = 0 Or DEBUGGING = 1 Then
        Dim server() As String
        server = Split(ServerIndex, ":")
        FrmLogear.txtIp.Text = server(0)
        FrmLogear.txtPort.Text = server(1)
    #Else
        FrmLogear.txtIp.Text = "45.235.99.71"
        FrmLogear.txtPort.Text = "6501"
    #End If
        
    Exit Sub

CargarLst_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.CargarLst", Erl)
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
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.CrearFantasma", Erl)
    Resume Next
    
End Sub

Public Sub CompletarAccionBarra(ByVal BarAccion As Byte)
    
    On Error GoTo CompletarAccionBarra_Err
    

    If BarAccion = Accion_Barra.CancelarAccion Then Exit Sub

    Call WriteCompletarAccion(BarAccion)

    
    Exit Sub

CompletarAccionBarra_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.CompletarAccionBarra", Erl)
    Resume Next
    
End Sub

Public Sub ComprobarEstado()
    
    On Error GoTo ComprobarEstado_Err
    

    Call CargarLst

    
    Exit Sub

ComprobarEstado_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.ComprobarEstado", Erl)
    Resume Next
    
End Sub

Public Function General_Distance_Get(ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer) As Integer
    
    On Error GoTo General_Distance_Get_Err
    
    General_Distance_Get = Abs(x1 - x2) + Abs(y1 - y2)

    
    Exit Function

General_Distance_Get_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.General_Distance_Get", Erl)
    Resume Next
    
End Function

Public Sub EndGame(Optional ByVal Closed_ByUser As Boolean = False, Optional ByVal Init_Launcher As Boolean = False)
    
    On Error GoTo EndGame_Err
    

    

    Sound.Engine_DeInitialize
    Sound.Music_Stop
    Set Sound = Nothing

    prgRun = False

    '0. Cerramos el socket
    Call modNetwork.Disconnect

    '2. Eliminamos objetos DX
    Call Client_UnInitialize_DirectX_Objects

    '6. Cerramos los forms y nos vamos
    Call UnloadAllForms

    '7. Adiós MuteX - Restauramos MouseSpeed

    End

    
    Exit Sub

EndGame_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.EndGame", Erl)
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
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.Client_UnInitialize_DirectX_Objects", Erl)
    Resume Next
    
End Sub

Public Sub TextoAlAsistente(ByVal Texto As String)
    
    On Error GoTo TextoAlAsistente_Err
    
    TextEfectAsistente = 35
    TextAsistente = Texto

    
    Exit Sub

TextoAlAsistente_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.TextoAlAsistente", Erl)
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
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.GetTimeFormated", Erl)
    Resume Next
    
End Function

Public Function GetHora(Mins As Integer) As String
    
    On Error GoTo GetHora_Err
    

    Dim Horita As Byte

    Horita = Fix(Mins / 60)

    GetHora = Horita

    
    Exit Function

GetHora_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.GetHora", Erl)
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

        If Not Extract_File(Scripts, App.path & "\..\Recursos\OUTPUT\", "preload.ind", Windows_Temp_Dir, ResourcesPassword, False) Then
            Err.Description = "No se ha logrado extraer el archivo de recurso."
            GoTo ErrorHandler

        End If
    
        PreloadFile = Windows_Temp_Dir & "Preload.ind"
    #Else
        PreloadFile = App.path & "\..\Recursos\init\Preload.ind"
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
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.PreloadGraphics", Erl)
    Resume Next
    
End Sub

Public Function ObtenerIdMapaDeLlamadaDeClan(ByVal Mapa As Integer) As Integer
    
    On Error GoTo ObtenerIdMapaDeLlamadaDeClan_Err
    

    Dim i        As Integer
    Dim j       As Byte

    Dim Encontre As Boolean

    For j = 1 To TotalWorlds
        For i = 1 To Mundo(j).Ancho * Mundo(j).Alto
            If Mundo(j).MapIndice(i) = Mapa Then
                ObtenerIdMapaDeLlamadaDeClan = i
                frmMapaGrande.llamadadeclan.Tag = 0
                Exit Function
                Exit For
            End If
    
        Next i
    Next j

    ObtenerIdMapaDeLlamadaDeClan = 0

    
    Exit Function

ObtenerIdMapaDeLlamadaDeClan_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.ObtenerIdMapaDeLlamadaDeClan", Erl)
    Resume Next
    
End Function

Public Sub Auto_Drag(ByVal hwnd As Long)
    
    On Error GoTo Auto_Drag_Err
    
    Call ReleaseCapture
    Call SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&)

    
    Exit Sub

Auto_Drag_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.Auto_Drag", Erl)
    Resume Next
    
End Sub

Public Function IsArrayInitialized(ByRef arr() As e_ActiveEffect) As Boolean
  Dim rv As Long
  On Error Resume Next
  rv = UBound(arr)
  IsArrayInitialized = (Err.Number = 0) And rv >= 0
End Function
