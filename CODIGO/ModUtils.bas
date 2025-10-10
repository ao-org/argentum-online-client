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
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'get mac adress
Public Type Tclan
    Alineacion As Byte
    indice As Integer
    nombre As String
End Type

Public ListaClanes  As Boolean
Public ClanesList() As Tclan
Public CheckMD5     As String
Public intro        As Byte
Public InviCounter  As Integer
Public DrogaCounter As Integer
Type Effect_Type
    FX_Grh     As Grh      '< FxGrh.
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

Public Const NO_INDEX = -1         '< índice no válido.
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

Public HayLayer4      As Boolean
Public CantPartLLuvia As Integer
Public MeteoIndex     As Integer
'Dropeo
Public PingRender     As Integer
Public NumOBJs        As Integer
Public NumNpcs        As Integer
Public NumHechizos    As Integer
Public NumLocaleMsg   As Integer
Public NumQuest       As Integer
Public NumSug         As Integer
Public Sugerencia()   As String

Public Type tQuestNpc
    NpcIndex As Integer
    Amount As Integer
End Type

Public Type tUserQuest
    NPCsKilled() As Integer
    QuestIndex As Integer
End Type

Public QuestList() As tQuest

Public Type t_QuestSkill
    SkillType As eSkill
    RequiredValue As Byte
End Type

Public Type tQuest
    nombre As String
    desc As String
    NextQuest As String
    DescFinal As String
    RequiredLevel As Integer
    RequiredClass As Integer
    RequiredQuest As Integer
    LimitLevel As Byte
    RequiredOBJ() As Obj
    RequiredNPC() As tQuestNpc
    RequiredSpellList() As Integer
    RequiredSkill As t_QuestSkill
    RewardGLD As Long
    RewardEXP As Long
    RewardOBJ() As Obj
    RewardSkillCount As Integer
    RewardSkill() As Integer
    Repetible As Byte
End Type

Public PosMap()         As Integer
Public ObjData()        As ObjDatas
Public ObjShop()        As ObjDatas
Public NpcData()        As NpcDatas
Public ProjectileData() As t_Projectile
Public GProjectile      As Projectile
Public Locale_SMG()     As String
'Sistema de mapa del mundo
Public TotalWorlds      As Byte

Public Type WorldMap
    MapIndice() As Integer
    Ancho As Integer
    Alto As Integer
End Type

Public Mundo()             As WorldMap
Public PosREAL             As Integer
Public Dungeon             As Boolean
Public idmap               As Integer
Public WorldActual         As Byte
'Sistema de mapa del mundo
Public HechizoData()       As HechizoDatas
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
    en_texto As String
    info As String
    CreaGRH As String
    CreaLuz As String
    CreaParticulaPiso As Integer
    proyectil As Byte
    Amunition As Byte
    Hechizo As Integer
    Raices As Integer
    Cuchara As Integer
    Botella As Integer
    Mortero As Integer
    FrascoAlq As Integer
    FrascoElixir As Integer
    Dosificador As Integer
    Orquidea As Integer
    Carmesi As Integer
    HongoDeLuz As Integer
    Esporas As Integer
    Tuna As Integer
    Cala As Integer
    ColaDeZorro As Integer
    FlorOceano As Integer
    FlorRoja As Integer
    Hierva As Integer
    HojasDeRin As Integer
    HojasRojas As Integer
    SemillasPros As Integer
    Pimiento As Integer
    Madera As Integer
    MaderaElfica As Integer
    PielLobo As Integer
    PielOsoPardo As Integer
    PielOsoPolar As Integer
    PielLoboNegro As Integer
    PielTigre As Integer
    PielTigreBengala As Integer
    LingH As Integer
    LingP As Integer
    LingO As Integer
    Coal As Integer
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
    Blodium As Integer
    FireEssence As Integer
    WaterEssence As Integer
    EarthEssence As Integer
    WindEssence As Integer
    ElementalTags As Long
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
    NoMapInfo As Byte
    ElementalTags As Long
End Type

Public Type HechizoDatas
    nombre As String
    desc As String
    PalabrasMagicas As String
    HechizeroMsg As String
    TargetMsg As String
    PropioMsg As String
    ManaRequerido As Integer
    MinSkill As Byte
    StaRequerido As Integer
    IconoIndex As Long
    Cooldown As Long
    IsBindable As Boolean
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

Public Enum e_EquipmentStyle
    Modern = 0
    Classic = 1
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
Private Declare Function SystemParametersInfo _
                Lib "user32" _
                Alias "SystemParametersInfoA" (ByVal uAction As Long, _
                                               ByVal uParam As Long, _
                                               ByRef lpvParam As Any, _
                                               ByVal fuWinIni As Long) As Long
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

Public Pjs(1 To MAX_PERSONAJES_EN_CUENTA) As UserCuentaPJS
Public RecordarCuenta                     As Boolean
Public CuentaRecordada                    As CuentasGuardadas
Public CantidadDePersonajesEnCuenta       As Byte
Type UserCuentaPJS
    nombre As String
    Nivel As Byte
    Mapa As Integer
    PosX As Integer
    PosY As Integer
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
    Backpack As Integer
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

Public Const InitialProjectileSize          As Integer = 45
Public AllProjectile(InitialProjectileSize) As Projectile
Public AvailableProjectile                  As t_IndexHeap
Public ActiveProjectile                     As t_IndexHeap
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
Public lRegion              As Long
Public Render_Connect_Rect  As Rect
Public Render_Main_Rect     As Rect
Public GameplayDrawAreaRect As Rect
Public RenderCullingRect    As Rect
Public Const StartRenderX = 10
Public Const StartRenderY = 152
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const FLAGSz = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Declare Function SetWindowPos _
        Lib "user32" (ByVal hWnd As Long, _
                      ByVal hWndInsertAfter As Long, _
                      ByVal x As Long, _
                      ByVal y As Long, _
                      ByVal cx As Long, _
                      ByVal cy As Long, _
                      ByVal wFlags As Long) As Long
Private Declare Function CreateIconFromResourceEx _
                Lib "user32.dll" (ByRef presbits As Any, _
                                  ByVal dwResSize As Long, _
                                  ByVal fIcon As Long, _
                                  ByVal dwVer As Long, _
                                  ByVal cxDesired As Long, _
                                  ByVal cyDesired As Long, _
                                  ByVal flags As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (lpPictDesc As Any, riid As Any, ByVal fPictureOwnsHandle As Long, ipic As IPicture) As Long
Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Public Declare Function DrawIconEx _
               Lib "user32.dll" (ByVal hdc As Long, _
                                 ByVal xLeft As Long, _
                                 ByVal yTop As Long, _
                                 ByVal hIcon As Long, _
                                 ByVal cxWidth As Long, _
                                 ByVal cyWidth As Long, _
                                 ByVal istepIfAniCur As Long, _
                                 ByVal hbrFlickerFreeDraw As Long, _
                                 ByVal diFlags As Long) As Long
Private Declare Function SendMessageLongRef Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Long) As Long
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
    ActiveProjectile.IndexInfo(Index) = ActiveProjectile.IndexInfo(ActiveProjectile.CurrentIndex)
    ActiveProjectile.CurrentIndex = ActiveProjectile.CurrentIndex - 1
End Sub

Sub inputbox_Password(El_Form As Form, Caracter As String)
    On Error GoTo inputbox_Password_Err
    m_ASC = Asc(Caracter)
    Call SetTimer(El_Form.hWnd, &H5000&, 100, AddressOf TimerProc)
    Exit Sub
inputbox_Password_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.inputbox_Password", Erl)
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
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.SetTopMostWindow", Erl)
    Resume Next
End Function

Public Sub LogError(desc As String)
    On Error GoTo errhandler
    frmDebug.add_text_tracebox "ERROR: " & desc
    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.path & "\Logs\errores.log" For Append Shared As #nfile
    Print #nfile, Date & "-" & Time & ":" & desc
    frmDebug.add_text_tracebox Date & "-" & Time & ":" & desc
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
    frmCrearPersonaje.lstRaza.ListIndex = 0
    frmCrearPersonaje.lstHogar.Clear
    For i = LBound(ListaCiudades()) To UBound(ListaCiudades())
        frmCrearPersonaje.lstHogar.AddItem ListaCiudades(i)
    Next i
    frmCrearPersonaje.lstHogar.ListIndex = 0
    frmCrearPersonaje.lstProfesion.Clear
    For i = LBound(ListaClases()) To UBound(ListaClases())
        frmCrearPersonaje.lstProfesion.AddItem ListaClases(i)
    Next i
    frmCrearPersonaje.lstProfesion.ListIndex = 0
    MiCabeza = val(frmCrearPersonaje.Cabeza.List(1))
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
    If g_game_state.State() <> e_state_connect_screen Then
        Call ShowLogin
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
    frmMain.shieldBar.Picture = LoadInterface("shield-bar.bmp", False)
    AlphaNiebla = 10
    Call Graficos_Particulas.Engine_spell_Particle_Set(41)
    If intro = 1 Then
        Call Graficos_Particulas.Engine_MeteoParticle_Set(207)
    End If
    Call Graficos_Particulas.Engine_Select_Particle_Set(203)
    ParticleLluviaDorada = Graficos_Particulas.General_Particle_Create(208, -1, -1)
    Call ao20audio.PlayMP3("31.mp3", True)
    mFadingMusicMod = 0
    CurMp3 = 1
    Call GoToLogIn
    ClickEnAsistente = 0
    If CuentaRecordada.nombre <> "" Then
        Call TextoAlAsistente(JsonLanguage.Item("LOGIN_SCREEN_WELCOME_MESSAGE"), False, True)
    Else
        Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_BIENVENIDO"), False, True)
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
    Call SetWindowRgn(Frm.hWnd, lRegion, True)
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

Public Sub General_Var_Write(ByVal File As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
    'Writes a var to a text file
    On Error GoTo General_Var_Write_Err
    writeprivateprofilestring Main, Var, value, File
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

'modTimer - ImperiumAO - v1.3.0
'
'Windows API timer functions and handles.
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
'Augusto José Rando (barrin@imperiumao.com.ar)
'   - First Relase
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
    Dim i As Long, J As Long
    For i = LBound(packetControl) To UBound(packetControl)
        With packetControl(i)
            .last_count = 0
            For J = 1 To 10
                .iterations(J) = 0
            Next J
        End With
    Next i
End Sub

Public Sub WriteConsoleUserChat(ByVal text As String, _
                                ByVal userName As String, _
                                ByVal red As Byte, _
                                ByVal green As Byte, _
                                ByVal blue As Byte, _
                                ByVal userStatus As Integer, _
                                ByVal Privileges As Integer)
    Dim NameRed   As Byte
    Dim NameGreen As Byte
    Dim NameBlue  As Byte
    If Privileges > 0 Then
        NameRed = ColoresPJ(Privileges).R
        NameGreen = ColoresPJ(Privileges).G
        NameBlue = ColoresPJ(Privileges).B
    Else
        Select Case userStatus
            Case 0: ' Criminal
                NameRed = ColoresPJ(23).R
                NameGreen = ColoresPJ(23).G
                NameBlue = ColoresPJ(23).B
            Case 1: ' Ciudadano
                NameRed = ColoresPJ(20).R
                NameGreen = ColoresPJ(20).G
                NameBlue = ColoresPJ(20).B
            Case 2: ' Caos
                NameRed = ColoresPJ(24).R
                NameGreen = ColoresPJ(24).G
                NameBlue = ColoresPJ(24).B
            Case 3: ' Armada
                NameRed = ColoresPJ(21).R
                NameGreen = ColoresPJ(21).G
                NameBlue = ColoresPJ(21).B
            Case 4: ' Conciclio
                NameRed = ColoresPJ(25).R
                NameGreen = ColoresPJ(25).G
                NameBlue = ColoresPJ(25).B
            Case 5: ' Consejo
                NameRed = ColoresPJ(22).R
                NameGreen = ColoresPJ(22).G
                NameBlue = ColoresPJ(22).B
        End Select
    End If
    Dim Pos As Integer
    Pos = InStr(userName, "<")
    If Pos = 0 Then Pos = LenB(userName) + 2
    Dim Name As String
    Name = Left$(userName, Pos - 2)
    text = Trim$(text)
    If LenB(userName) <> 0 And LenB(text) > 0 Then
        Call AddtoRichTextBox2(frmMain.RecTxt, "[" & Name & "] ", NameRed, NameGreen, NameBlue, True, False, True, rtfLeft)
        Call AddtoRichTextBox2(frmMain.RecTxt, text, red, green, blue, False, False, False, rtfLeft)
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

Public Function General_Var_Get(ByVal File As String, ByVal Main As String, ByVal Var As String) As String
    On Error GoTo General_Var_Get_Err
    'Get a var to from a text file
    Dim l        As Long
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
    frmMain.MiniMap.Picture = LoadMinimap(ResourceMap)
    'Pintamos los NPCs en Minimapa:
    If ListNPCMapData(ResourceMap).NpcCount > 0 Then
        Dim i As Long
        For i = 1 To MAX_QUESTNPCS_VISIBLE
            Dim PosX As Long
            Dim PosY As Long
            PosX = ListNPCMapData(ResourceMap).NpcList(i).Position.x
            PosY = ListNPCMapData(ResourceMap).NpcList(i).Position.y
            Dim color As Long
            Select Case ListNPCMapData(ResourceMap).NpcList(i).State
                Case 1
                    color = RGB(0, 198, 254)
                Case 2
                    color = RGB(255, 201, 14)
                Case Else
                    color = RGB(255, 201, 14)
            End Select
            Call SetPixel(frmMain.MiniMap.hdc, PosX + 1, PosY, color)
            Call SetPixel(frmMain.MiniMap.hdc, PosX, PosY + 1, color)
            Call SetPixel(frmMain.MiniMap.hdc, PosX + 1, PosY + 1, color)
            Call SetPixel(frmMain.MiniMap.hdc, PosX, PosY, color)
            Call SetPixel(frmMain.MiniMap.hdc, PosX, PosY - 1, &H808080)
            Call SetPixel(frmMain.MiniMap.hdc, PosX + 1, PosY - 1, &H808080)
            Call SetPixel(frmMain.MiniMap.hdc, PosX + 2, PosY, &H808080)
            Call SetPixel(frmMain.MiniMap.hdc, PosX + 2, PosY + 1, &H808080)
            Call SetPixel(frmMain.MiniMap.hdc, PosX + 1, PosY + 2, &H808080)
            Call SetPixel(frmMain.MiniMap.hdc, PosX, PosY + 2, &H808080)
            Call SetPixel(frmMain.MiniMap.hdc, PosX - 1, PosY + 1, &H808080)
            Call SetPixel(frmMain.MiniMap.hdc, PosX - 1, PosY, &H808080)
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
Function EncryptStr(ByVal s As String, ByVal p As String) As String
    On Error GoTo EncryptStr_Err
    Dim i  As Integer, R As String
    Dim c1 As Integer, C2 As Integer
    R = ""
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
            R = R + Chr(c1)
        Next i
    Else
        R = s
    End If
    EncryptStr = R
    Exit Function
EncryptStr_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.EncryptStr", Erl)
    Resume Next
End Function

Rem Desencripta una cadena de caracteres.
Rem S = Cadena a desencriptar
Rem P = Password
Function UnEncryptStr(ByVal s As String, ByVal p As String) As String
    On Error GoTo UnEncryptStr_Err
    Dim i  As Integer, R As String
    Dim c1 As Integer, C2 As Integer
    R = ""
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
            R = R + Chr(c1)
        Next i
    Else
        R = s
    End If
    UnEncryptStr = R
    Exit Function
UnEncryptStr_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.UnEncryptStr", Erl)
    Resume Next
End Function

Public Function Input_Key_Get(ByVal key_code As Byte) As Boolean
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
    'Gets windows temporary directory
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
    On Error GoTo General_Get_Mouse_Speed_Err
    SystemParametersInfo SPI_GETMOUSESPEED, 0, General_Get_Mouse_Speed, 0
    Exit Function
General_Get_Mouse_Speed_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.General_Get_Mouse_Speed", Erl)
    Resume Next
End Function
 
Public Sub General_Set_Mouse_Speed(ByVal lngSpeed As Long)
    On Error GoTo General_Set_Mouse_Speed_Err
    SystemParametersInfo SPI_SETMOUSESPEED, 0, ByVal lngSpeed, 0
    Exit Sub
General_Set_Mouse_Speed_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.General_Set_Mouse_Speed", Erl)
    Resume Next
End Sub

Public Sub ResetearUserMacro()
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
    AddtoRichTextBox frmMain.RecTxt, JsonLanguage.Item("MENSAJE_DEJAS_DE_TRABAJAR"), 223, 51, 2, 1, 0
    Exit Sub
ResetearUserMacro_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.ResetearUserMacro", Erl)
    Resume Next
End Sub

Public Sub CargarLst()
    On Error GoTo CargarLst_Err
    #If PYMMO = 0 Or DEBUGGING = 1 Then
        Dim server() As String
        If Len(ServerIndex) > 0 Then
            server = Split(ServerIndex, ":")
            FrmLogear.txtIp.text = server(0)
            FrmLogear.txtPort.text = server(1)
        End If
    #Else
        FrmLogear.txtIp.text = "45.235.98.188"
        FrmLogear.txtPort.text = "6501"
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
    prgRun = False
    Call modNetwork.Disconnect
    Call Client_UnInitialize_DirectX_Objects
    Call UnloadAllForms
    End
    Exit Sub
EndGame_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.EndGame", Erl)
    Resume Next
End Sub

Public Sub Client_UnInitialize_DirectX_Objects()
    On Error GoTo Client_UnInitialize_DirectX_Objects_Err
    Set ao20audio.AudioEngine = Nothing
    #If DIRECT_PLAY = 1 Then
        Call modDplayClient.shutdown_direct_play
    #End If
    Exit Sub
Client_UnInitialize_DirectX_Objects_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.Client_UnInitialize_DirectX_Objects", Erl)
    Resume Next
End Sub

Public Sub TextoAlAsistente(ByVal Texto As String, ByVal IsLoading As Boolean, ByVal ForceAssistant As Boolean)
    On Error GoTo TextoAlAsistente_Err
    frmDebug.add_text_tracebox Texto
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

Public Function ObtenerIdMapaDeLlamadaDeClan(ByVal Mapa As Integer) As Integer
    On Error GoTo ObtenerIdMapaDeLlamadaDeClan_Err
    Dim i        As Integer
    Dim J        As Byte
    Dim Encontre As Boolean
    For J = 1 To TotalWorlds
        For i = 1 To Mundo(J).Ancho * Mundo(J).Alto
            If Mundo(J).MapIndice(i) = Mapa Then
                ObtenerIdMapaDeLlamadaDeClan = i
                frmMapaGrande.llamadadeclan.Tag = 0
                Exit Function
                Exit For
            End If
        Next i
    Next J
    ObtenerIdMapaDeLlamadaDeClan = 0
    Exit Function
ObtenerIdMapaDeLlamadaDeClan_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.ObtenerIdMapaDeLlamadaDeClan", Erl)
    Resume Next
End Function

Public Sub Auto_Drag(ByVal hWnd As Long)
    On Error GoTo Auto_Drag_Err
    Call ReleaseCapture
    Call SendMessage(hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0&)
    Exit Sub
Auto_Drag_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.Auto_Drag", Erl)
    Resume Next
End Sub

Public Function IsArrayInitialized(ByRef arr() As t_ActiveEffect) As Boolean
    Dim rv As Long
    On Error Resume Next
    rv = UBound(arr)
    IsArrayInitialized = (Err.Number = 0) And rv >= 0
End Function

Public Function ElementalTagsToTxtParser(ByVal ElementalTags As Long) As String
    Dim tmpString As String
    If ElementalTags = e_ElementalTags.Normal Then
        tmpString = tmpString + "[" & JsonLanguage.Item("MENSAJE_ELEMENTO_NORMAL") & "]"
        ElementalTagsToTxtParser = tmpString
        Exit Function
    End If
    If IsSet(ElementalTags, e_ElementalTags.Fire) Then
        tmpString = tmpString + "[" & JsonLanguage.Item("MENSAJE_ELEMENTO_FUEGO") & "]"
    End If
    If IsSet(ElementalTags, e_ElementalTags.Water) Then
        tmpString = tmpString + "[" & JsonLanguage.Item("MENSAJE_ELEMENTO_AGUA") & "]"
    End If
    If IsSet(ElementalTags, e_ElementalTags.Earth) Then
        tmpString = tmpString + "[" & JsonLanguage.Item("MENSAJE_ELEMENTO_TIERRA") & "]"
    End If
    If IsSet(ElementalTags, e_ElementalTags.Wind) Then
        tmpString = tmpString + "[" & JsonLanguage.Item("MENSAJE_ELEMENTO_AIRE") & "]"
    End If
    ElementalTagsToTxtParser = tmpString
End Function

Public Function UserInTileToTxtParser(ByRef Fields() As String)
    On Error GoTo UserInTileToTxtParser_Err
    Dim targetName          As String
    Dim targetDescription   As String
    Dim guildName           As String
    Dim Spouse              As String
    Dim CharClass           As String
    Dim CharRace            As String
    Dim level               As String
    Dim Elo                 As String
    Dim StatusMask          As Long
    Dim StatusMask2         As Long
    Dim SplitServerFields() As String
    SplitServerFields = Split(Fields(0), "-")
    Dim i As Byte
    i = LBound(SplitServerFields)
    targetName = SplitServerFields(i)
    i = i + 1
    targetDescription = SplitServerFields(i)
    i = i + 1
    guildName = SplitServerFields(i)
    i = i + 1
    Spouse = SplitServerFields(i)
    i = i + 1
    CharClass = IIf(SplitServerFields(i) <> "", SplitServerFields(i), "0")
    i = i + 1
    CharRace = IIf(SplitServerFields(i) <> "", SplitServerFields(i), "0")
    i = i + 1
    level = SplitServerFields(i)
    i = i + 1
    Elo = SplitServerFields(i)
    Select Case CByte(CharClass)
        Case e_Class.Mage
            CharClass = JsonLanguage.Item("MENSAJE_CLASE_MAGO")
        Case e_Class.Cleric
            CharClass = JsonLanguage.Item("MENSAJE_CLASE_CLERIGO")
        Case e_Class.Warrior
            CharClass = JsonLanguage.Item("MENSAJE_CLASE_GUERRERO")
        Case e_Class.Assasin
            CharClass = JsonLanguage.Item("MENSAJE_CLASE_ASESINO")
        Case e_Class.Bard
            CharClass = JsonLanguage.Item("MENSAJE_CLASE_BARDO")
        Case e_Class.Druid
            CharClass = JsonLanguage.Item("MENSAJE_CLASE_DRUIDA")
        Case e_Class.paladin
            CharClass = JsonLanguage.Item("MENSAJE_CLASE_PALADIN")
        Case e_Class.Hunter
            CharClass = JsonLanguage.Item("MENSAJE_CLASE_CAZADOR")
        Case e_Class.Worker
            CharClass = JsonLanguage.Item("MENSAJE_CLASE_TRABAJADOR")
        Case e_Class.Pirat
            CharClass = JsonLanguage.Item("MENSAJE_CLASE_PIRATA")
        Case e_Class.Thief
            CharClass = JsonLanguage.Item("MENSAJE_CLASE_LADRON")
        Case e_Class.Bandit
            CharClass = JsonLanguage.Item("MENSAJE_CLASE_BANDIDO")
        Case Else
            CharClass = ""
    End Select
    Select Case CByte(CharRace)
        Case e_Race.Human
            CharRace = JsonLanguage.Item("MENSAJE_RAZA_HUMANO")
        Case e_Race.Elf
            CharRace = JsonLanguage.Item("MENSAJE_RAZA_ELFO")
        Case e_Race.DrowElf
            CharRace = JsonLanguage.Item("MENSAJE_RAZA_ELFO_OSCURO")
        Case e_Race.Gnome
            CharRace = JsonLanguage.Item("MENSAJE_RAZA_GNOMO")
        Case e_Race.Dwarf
            CharRace = JsonLanguage.Item("MENSAJE_RAZA_ENANO")
        Case e_Race.Orc
            CharRace = JsonLanguage.Item("MENSAJE_RAZA_ORCO")
        Case Else
            CharRace = ""
    End Select
    StatusMask = CLng(Fields(1))
    StatusMask2 = CLng(Fields(2))
    Dim StatusString        As String
    Dim FactionStatusString As String
    If IsSet(StatusMask, e_InfoTxts.Newbie) Then
        StatusString = StatusString & "<" & JsonLanguage.Item("MENSAJE_ESTADO_NOVATO") & ">" & " "
    End If
    StatusString = StatusString & "<"
    If IsSet(StatusMask, e_InfoTxts.Poisoned) Then
        StatusString = StatusString & JsonLanguage.Item("MENSAJE_ESTADO_ENVENENADO") & "|"
    End If
    If IsSet(StatusMask, e_InfoTxts.Blind) Then
        StatusString = StatusString & JsonLanguage.Item("MENSAJE_ESTADO_CEGADO") & "|"
    End If
    If IsSet(StatusMask, e_InfoTxts.Paralized) Then
        StatusString = StatusString & JsonLanguage.Item("MENSAJE_ESTADO_PARALIZADO") & "|"
    End If
    If IsSet(StatusMask, e_InfoTxts.Inmovilized) Then
        StatusString = StatusString & JsonLanguage.Item("MENSAJE_ESTADO_INMOVILIZADO") & "|"
    End If
    If IsSet(StatusMask, e_InfoTxts.Working) Then
        StatusString = StatusString & JsonLanguage.Item("MENSAJE_ESTADO_TRABAJANDO") & "|"
    End If
    If IsSet(StatusMask, e_InfoTxts.Invisible) Then
        StatusString = StatusString & JsonLanguage.Item("MENSAJE_ESTADO_INVISIBLE") & "|"
    End If
    If IsSet(StatusMask, e_InfoTxts.Hidden) Then
        StatusString = StatusString & JsonLanguage.Item("MENSAJE_ESTADO_OCULTO") & "|"
    End If
    If IsSet(StatusMask, e_InfoTxts.Stupid) Then
        StatusString = StatusString & JsonLanguage.Item("MENSAJE_ESTADO_TORPE") & "|"
    End If
    If IsSet(StatusMask, e_InfoTxts.Cursed) Then
        StatusString = StatusString & JsonLanguage.Item("MENSAJE_ESTADO_MALDITO") & "|"
    End If
    If IsSet(StatusMask, e_InfoTxts.Silenced) Then
        StatusString = StatusString & JsonLanguage.Item("MENSAJE_ESTADO_SILENCIADO") & "|"
    End If
    If IsSet(StatusMask, e_InfoTxts.Trading) Then
        StatusString = StatusString & JsonLanguage.Item("MENSAJE_ESTADO_COMERCIANDO") & "|"
    End If
    If IsSet(StatusMask, e_InfoTxts.Resting) Then
        StatusString = StatusString & JsonLanguage.Item("MENSAJE_ESTADO_DESCANSANDO") & "|"
    End If
    If IsSet(StatusMask, e_InfoTxts.Focusing) Then
        StatusString = StatusString & JsonLanguage.Item("MENSAJE_ESTADO_CONCENTRANDOSE") & "|"
    End If
    If IsSet(StatusMask, e_InfoTxts.Incinerated) Then
        StatusString = StatusString & JsonLanguage.Item("MENSAJE_ESTADO_INCINERADO") & "|"
    End If
    If IsSet(StatusMask, e_InfoTxts.Dead) Then
        StatusString = StatusString & JsonLanguage.Item("MENSAJE_ESTADO_MUERTO")
    End If
    If IsSet(StatusMask, e_InfoTxts.AlmostDead) Then
        StatusString = StatusString & JsonLanguage.Item("MENSAJE_ESTADO_CASI_MUERTO")
    End If
    If IsSet(StatusMask, e_InfoTxts.SeriouslyWounded) Then
        StatusString = StatusString & JsonLanguage.Item("MENSAJE_ESTADO_GRAVEMENTE_HERIDO")
    End If
    If IsSet(StatusMask, e_InfoTxts.Wounded) Then
        StatusString = StatusString & JsonLanguage.Item("MENSAJE_ESTADO_HERIDO")
    End If
    If IsSet(StatusMask, e_InfoTxts.LightlyWounded) Then
        StatusString = StatusString & JsonLanguage.Item("MENSAJE_ESTADO_LEVEMENTE_HERIDO")
    End If
    If IsSet(StatusMask, e_InfoTxts.Intact) Then
        StatusString = StatusString & JsonLanguage.Item("MENSAJE_ESTADO_INTACTO")
    End If
    StatusString = StatusString & ">"
    If IsSet(StatusMask, e_InfoTxts.Counselor) Then
        StatusString = StatusString & "<" & JsonLanguage.Item("MENSAJE_ESTADO_CONSEJERO") & ">" & " "
    End If
    If IsSet(StatusMask, e_InfoTxts.DemiGod) Then
        StatusString = StatusString & "<" & JsonLanguage.Item("MENSAJE_ESTADO_DEMIGOD") & ">" & " "
    End If
    If IsSet(StatusMask, e_InfoTxts.God) Then
        StatusString = StatusString & "<" & JsonLanguage.Item("MENSAJE_ESTADO_DIOS") & ">" & " "
    End If
    If IsSet(StatusMask, e_InfoTxts.Admin) Then
        StatusString = StatusString & "<" & JsonLanguage.Item("MENSAJE_ESTADO_ADMIN") & ">" & " "
    End If
    If IsSet(StatusMask, e_InfoTxts.RoleMaster) Then
        StatusString = StatusString & "<" & JsonLanguage.Item("MENSAJE_ESTADO_ROLEMASTER") & ">" & " "
    End If
    If IsSet(StatusMask2, e_InfoTxts2.ChaoticCouncil) Then
        FactionStatusString = FactionStatusString & "<" & JsonLanguage.Item("MENSAJE_ESTADO_CONSEJO_CAOS") & ">" & " "
    End If
    If IsSet(StatusMask2, e_InfoTxts2.Chaotic) Then
        FactionStatusString = FactionStatusString & "<" & JsonLanguage.Item("MENSAJE_ESTADO_CAOS") & ">" & " "
    End If
    If IsSet(StatusMask2, e_InfoTxts2.ChaosFirstHierarchy) Then
        FactionStatusString = FactionStatusString & "<" & JsonLanguage.Item("MENSAJE_ESTADO_CAOS_PRIMERA_JERARQUIA") & ">"
    End If
    If IsSet(StatusMask2, e_InfoTxts2.ChaosSecondHierarchy) Then
        FactionStatusString = FactionStatusString & "<" & JsonLanguage.Item("MENSAJE_ESTADO_CAOS_SEGUNDA_JERARQUIA") & ">"
    End If
    If IsSet(StatusMask2, e_InfoTxts2.ChaosThirdHierarchy) Then
        FactionStatusString = FactionStatusString & "<" & JsonLanguage.Item("MENSAJE_ESTADO_CAOS_TERCERA_JERARQUIA") & ">"
    End If
    If IsSet(StatusMask2, e_InfoTxts2.ChaosFourthHierarchy) Then
        FactionStatusString = FactionStatusString & "<" & JsonLanguage.Item("MENSAJE_ESTADO_CAOS_CUARTA_JERARQUIA") & ">"
    End If
    If IsSet(StatusMask2, e_InfoTxts2.ChaosFifthHierarchy) Then
        FactionStatusString = FactionStatusString & "<" & JsonLanguage.Item("MENSAJE_ESTADO_CAOS_QUINTA_JERARQUIA") & ">"
    End If
    If IsSet(StatusMask2, e_InfoTxts2.Criminal) Then
        FactionStatusString = FactionStatusString & "<" & JsonLanguage.Item("MENSAJE_ESTADO_CRIMINAL") & ">"
    End If
    If IsSet(StatusMask2, e_InfoTxts2.RoyalCouncil) Then
        FactionStatusString = FactionStatusString & "<" & JsonLanguage.Item("MENSAJE_ESTADO_CONSEJO_REAL") & ">" & " "
    End If
    If IsSet(StatusMask2, e_InfoTxts2.Army) Then
        FactionStatusString = FactionStatusString & "<" & JsonLanguage.Item("MENSAJE_ESTADO_ARMADA") & ">" & " "
    End If
    If IsSet(StatusMask2, e_InfoTxts2.ArmyFirstHierarchy) Then
        FactionStatusString = FactionStatusString & "<" & JsonLanguage.Item("MENSAJE_ESTADO_ARMADA_PRIMERA_JERARQUIA") & ">"
    End If
    If IsSet(StatusMask2, e_InfoTxts2.ArmySecondHierarchy) Then
        FactionStatusString = FactionStatusString & "<" & JsonLanguage.Item("MENSAJE_ESTADO_ARMADA_SEGUNDA_JERARQUIA") & ">"
    End If
    If IsSet(StatusMask2, e_InfoTxts2.ArmyThirdHierarchy) Then
        FactionStatusString = FactionStatusString & "<" & JsonLanguage.Item("MENSAJE_ESTADO_ARMADA_TERCERA_JERARQUIA") & ">"
    End If
    If IsSet(StatusMask2, e_InfoTxts2.ArmyFourthHierarchy) Then
        FactionStatusString = FactionStatusString & "<" & JsonLanguage.Item("MENSAJE_ESTADO_ARMADA_CUARTA_JERARQUIA") & ">"
    End If
    If IsSet(StatusMask2, e_InfoTxts2.ArmyFifthHierarchy) Then
        FactionStatusString = FactionStatusString & "<" & JsonLanguage.Item("MENSAJE_ESTADO_ARMADA_QUINTA_JERARQUIA") & ">"
    End If
    If IsSet(StatusMask2, e_InfoTxts2.Citizen) Then
        FactionStatusString = FactionStatusString & "<" & JsonLanguage.Item("MENSAJE_ESTADO_CIUDADANO") & ">"
    End If
    Fields(0) = targetName & " "
    If targetDescription <> "" Then
        Fields(0) = Fields(0) & "<" & targetDescription & ">" & " "
    End If
    If guildName <> "" Then
        Fields(0) = Fields(0) & "<" & guildName & ">" & " "
    End If
    If Spouse <> "" Then
        Fields(0) = Fields(0) & "<" & JsonLanguage.Item("MENSAJE_ESTADO_CASADO") & " " & Spouse & ">" & " "
    End If
    If CharClass <> "" Then
        Fields(0) = Fields(0) & "<" & CharClass & "|"
    End If
    If CharRace <> "" Then
        Fields(0) = Fields(0) & CharRace & ">" & " "
    End If
    If level <> "" Then
        Fields(0) = Fields(0) & "<" & JsonLanguage.Item("MENSAJE_NIVEL") & ":" & level & "> "
    End If
    If Elo <> "" Then
        Fields(0) = Fields(0) & "Elo:" & Elo & " "
    End If
    If StatusString <> "" Then
        StatusString = Replace(StatusString, "<>", "")
        StatusString = Replace(StatusString, " ", "")
        Fields(1) = StatusString
    Else
        Fields(1) = ""
    End If
    If FactionStatusString <> "" Then
        StatusString = Replace(StatusString, "<>", "")
        StatusString = Replace(StatusString, " ", "")
        Fields(2) = FactionStatusString
    End If
    Exit Function
UserInTileToTxtParser_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.UserInTileToTxtParser", Erl)
    Resume Next
End Function

Public Function NpcInTileToTxtParser(ByRef Fields() As String, ByVal bytHeader As Integer)
    On Error GoTo NpcInTileToTxtParser_Err
    Dim SplitNpcStatus()     As String
    Dim NpcName              As String
    Dim NpcElementalTags     As Long
    Dim NpcStatuses          As String
    Dim NpcStatusMask        As Long
    Dim extraInfo            As String
    Dim NpcFightingWith      As String
    Dim NpcFightingWithTimer As String
    Dim NpcHpInfo            As String
    Dim NpcIndex             As String
    Dim ParalisisTime        As String
    Dim InmovilizedTime      As String
    NpcName = Fields(0)
    NpcElementalTags = CLng(Fields(1))
    NpcStatuses = Fields(2)
    SplitNpcStatus = Split(NpcStatuses, "-")
    NpcHpInfo = SplitNpcStatus(0)
    NpcStatusMask = CLng(SplitNpcStatus(5))
    NpcIndex = SplitNpcStatus(4)
    If NpcIndex <> "" Then
        extraInfo = extraInfo & " NpcIndex: " & NpcIndex
    End If
    If NpcStatusMask > 0 Then
        If IsSet(NpcStatusMask, e_NpcInfoMask.AlmostDead) Then
            extraInfo = extraInfo & "[" & JsonLanguage.Item("MENSAJE_ESTADO_CASIMUERTO") & "]"
        End If
        If IsSet(NpcStatusMask, e_NpcInfoMask.SeriouslyWounded) Then
            extraInfo = extraInfo & "[" & JsonLanguage.Item("MENSAJE_ESTADO_GRAVEMENTE_HERIDO") & "]"
        End If
        If IsSet(NpcStatusMask, e_NpcInfoMask.Wounded) Then
            extraInfo = extraInfo & "[" & JsonLanguage.Item("MENSAJE_ESTADO_HERIDO") & "]"
        End If
        If IsSet(NpcStatusMask, e_NpcInfoMask.LightlyWounded) Then
            extraInfo = extraInfo & "[" & JsonLanguage.Item("MENSAJE_ESTADO_LEVEMENTE_HERIDO") & "]"
        End If
        If IsSet(NpcStatusMask, e_NpcInfoMask.Intact) Then
            extraInfo = extraInfo & "[" & JsonLanguage.Item("MENSAJE_ESTADO_INTACTO") & "]"
        End If
        If IsSet(NpcStatusMask, e_NpcInfoMask.Paralized) Then
            extraInfo = extraInfo & "[" & JsonLanguage.Item("MENSAJE_ESTADO_PARALIZADO") & "]"
            ParalisisTime = SplitNpcStatus(1)
            If ParalisisTime <> "" Then
                extraInfo = extraInfo & "(" & ParalisisTime & "s)" & "]"
            End If
        End If
        If IsSet(NpcStatusMask, e_NpcInfoMask.Inmovilized) Then
            extraInfo = extraInfo & "[" & JsonLanguage.Item("MENSAJE_ESTADO_INMOVILIZADO") & "]"
            InmovilizedTime = SplitNpcStatus(2)
            If InmovilizedTime <> "" Then
                extraInfo = extraInfo & "(" & InmovilizedTime & "s)" & "]"
            End If
        End If
        If IsSet(NpcStatusMask, e_NpcInfoMask.Fighting) Then
            extraInfo = extraInfo & "[" & JsonLanguage.Item("MENSAJE_ESTADO_PELEANDO")
            NpcFightingWith = Split(SplitNpcStatus(3), "|")(0)
            NpcFightingWithTimer = Split(SplitNpcStatus(3), "|")(1)
            extraInfo = extraInfo & NpcFightingWith & "(" & NpcFightingWithTimer & "s)" & "]"
        End If
    End If
    Fields(1) = ElementalTagsToTxtParser(NpcElementalTags)
    If NpcHpInfo <> "" Then
        Fields(2) = "<" & NpcHpInfo & ">" & extraInfo
    Else
        Fields(2) = extraInfo
    End If
    If bytHeader = 1621 Then
        Fields(3) = "[" & JsonLanguage.Item("MENSAJE_ESTADO_MASCOTA") & " " & Fields(3) & "]"
    End If
    Exit Function
NpcInTileToTxtParser_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.NpcInTileToTxtParser", Erl)
    Resume Next
End Function

Public Function SkillsNamesToTxtParser(ByRef Fields() As String)
On Error GoTo SkillsNamesToTxtParser_Err

    Dim skillID As Integer
    skillID = CInt(Fields(0))
    
    Select Case skillID
        Case 1:  Fields(0) = CStr(JsonLanguage.Item("MENSAJE_SKILL_MAGIA"))
        Case 2:  Fields(0) = CStr(JsonLanguage.Item("MENSAJE_SKILL_ROBAR"))
        Case 3:  Fields(0) = CStr(JsonLanguage.Item("MENSAJE_SKILL_TACTICAS"))
        Case 4:  Fields(0) = CStr(JsonLanguage.Item("MENSAJE_SKILL_ARMAS"))
        Case 5:  Fields(0) = CStr(JsonLanguage.Item("MENSAJE_SKILL_MEDITAR"))
        Case 6:  Fields(0) = CStr(JsonLanguage.Item("MENSAJE_SKILL_APUÑALAR"))
        Case 7:  Fields(0) = CStr(JsonLanguage.Item("MENSAJE_SKILL_OCULTARSE"))
        Case 8:  Fields(0) = CStr(JsonLanguage.Item("MENSAJE_SKILL_SUPERVIVENCIA"))
        Case 9:  Fields(0) = CStr(JsonLanguage.Item("MENSAJE_SKILL_COMERCIAR"))
        Case 10: Fields(0) = CStr(JsonLanguage.Item("MENSAJE_SKILL_DEFENSA"))
        Case 11: Fields(0) = CStr(JsonLanguage.Item("MENSAJE_SKILL_LIDERAZGO"))
        Case 12: Fields(0) = CStr(JsonLanguage.Item("MENSAJE_SKILL_PROYECTILES"))
        Case 13: Fields(0) = CStr(JsonLanguage.Item("MENSAJE_SKILL_WRESTLING"))
        Case 14: Fields(0) = CStr(JsonLanguage.Item("MENSAJE_SKILL_NAVEGACION"))
        Case 15: Fields(0) = CStr(JsonLanguage.Item("MENSAJE_SKILL_EQUITACION"))
        Case 16: Fields(0) = CStr(JsonLanguage.Item("MENSAJE_SKILL_RESISTENCIA"))
        Case 17: Fields(0) = CStr(JsonLanguage.Item("MENSAJE_SKILL_TALAR"))
        Case 18: Fields(0) = CStr(JsonLanguage.Item("MENSAJE_SKILL_PESCAR"))
        Case 19: Fields(0) = CStr(JsonLanguage.Item("MENSAJE_SKILL_MINERIA"))
        Case 20: Fields(0) = CStr(JsonLanguage.Item("MENSAJE_SKILL_HERRERIA"))
        Case 21: Fields(0) = CStr(JsonLanguage.Item("MENSAJE_SKILL_CARPINTERIA"))
        Case 22: Fields(0) = CStr(JsonLanguage.Item("MENSAJE_SKILL_ALQUIMIA"))
        Case 23: Fields(0) = CStr(JsonLanguage.Item("MENSAJE_SKILL_SASTRERIA"))
        Case 24: Fields(0) = CStr(JsonLanguage.Item("MENSAJE_SKILL_DOMAR"))
    End Select
        
    Exit Function
SkillsNamesToTxtParser_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.SkillsNamesToTxtParser", Erl)
    Resume Next
End Function
