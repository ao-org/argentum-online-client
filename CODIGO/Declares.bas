Attribute VB_Name = "Mod_Declaraciones"
'RevolucionAo 1.0
'Pablo Mercavides
Option Explicit

Public ViajarInterface                  As Byte

Public FormParser                       As clsCursor

Public EsGM                             As Boolean

Public HayFormularioAbierto             As Boolean

Public HayLLamadaDeclan                 As Boolean

Public LLamadaDeclanMapa                As Byte

Public LLamadaDeclanX                   As Byte

Public LLamadaDeclanY                   As Byte

Public SugerenciaAMostrar               As Byte

Public OffSetConsola                    As Byte

Public Const ComienzoY                  As Integer = 0

Public UltimaLineavisible               As Boolean

Public Const MaxLineas                  As Byte = 6

Public UserInvUnlocked                  As Byte

'Slots de Inventarios Generales
Public Const GRH_SLOT_INVENTARIO_NEGRO  As Long = 26095

Public Const GRH_SLOT_INVENTARIO_ROJO   As Long = 26096

'Slots de Inventario Principal
Public Const GRH_INVENTORYSLOT          As Long = 47743

Public Const GRH_INVENTORYSLOT_EXTRA    As Long = 47742

Public Const GRH_INVENTORYSLOT_LOCKED   As Long = 1122

Public Const GRH_INVENTORYSLOT_SELECTED As Long = 32873

' Cantidad de "slots" en el inventario basico
Public Const MAX_NORMAL_INVENTORY_SLOTS As Byte = 24

' Cantidad de "slots" en el inventario con slots desbloqueados
Public Const MAX_INVENTORY_SLOTS        As Byte = 42
 
Type TConsola

    T As String
    '   Color As Long
    r As Byte
    g As Byte
    b As Byte

End Type
 
Public Con(1 To MaxLineas) As TConsola

Type TRanking

    nombre As String
    puntos As Long

End Type

Public LRanking(1 To 10) As TRanking

'Creacion de PJ 17/8/20
Public RazaRecomendada   As String

Public CPBody            As Long

Public CPBodyE           As Long

Public CPArma            As Long

Public CPGorro           As Long

Public CPEscudo          As Long

'Public CPArma As Long
Public CPAura            As String

Public CPHead            As Long

Public CPHeading         As Long

Public CPEquipado        As Boolean

Public CPName            As String

Public Enum TipoPaso

    CONST_BOSQUE = 1
    CONST_NIEVE = 2
    CONST_CABALLO = 3
    CONST_DUNGEON = 4
    CONST_PISO = 5
    CONST_DESIERTO = 6
    CONST_PESADO = 7

End Enum

Public Type tPaso

    CantPasos As Byte
    wav() As Integer

End Type

Public Const NUM_PASOS       As Byte = 6

Public Pasos()               As tPaso

Public PosXMacro             As Integer

Public PosYMacro             As Long

Public MacrosHorizontal      As Boolean

Public MacroPos              As Byte

Public UserWeaponEqpSlot     As Byte

Public UserArmourEqpSlot     As Byte

Public UserHelmEqpSlot       As Byte

Public UserShieldEqpSlot     As Byte

Public TextAsistente         As String

Public TextEfectAsistente    As Single

Public ClickEnAsistente      As Long

Public PJSeleccionado        As Byte

Public AlphaNiebla           As Byte

Public MaxAlphaNiebla        As Byte
      
Public ExpMult               As Integer

Public OroMult               As Integer

Public DireccionDeCaminata   As String

Public CaminandoMacro        As Boolean

Public CaminarX              As Integer

Public CaminarY              As Integer

Public OpcionSeleccionada    As Byte

Public clicX                 As Long

Public clicY                 As Long

Public FxLoops               As Long

'¿Estamos haciendo efecto fade?
Public mFadingStatus         As Byte

Public mFadingMusicMod       As Long

Public NextMP3               As Byte

Public Enum E_SISTEMA_MUSICA

    CONST_DESHABILITADA = 0
    CONST_MIDI = 1
    CONST_MP3 = 2

End Enum

Public Music                       As E_SISTEMA_MUSICA

Public PermitirMoverse             As Byte

Public MoverVentana                As Byte

Public CursoresGraficos            As Byte

Public UtilizarPreCarga            As Byte

Public SensibilidadMouse           As Byte

Public SensibilidadMouseOriginal   As Byte

Public CopiarDialogoAConsola       As Byte

Public Musica                      As Byte

Public fX                          As Byte

Public AmbientalActivated          As Byte

Public InvertirSonido              As Byte

Public VolMusic                    As Long

Public VolFX                       As Long

Public VolAmbient                  As Long

Public FxNavega                    As Byte

Public ChatCombate                 As Byte

Public ChatGlobal                  As Byte

Public PantallaCompleta            As Byte

Public Sonido                      As Byte

Public OcultarMacrosAlCastear      As Byte

Public MostrarIconosMeteorologicos As Byte

Public OpcionMenu                  As Byte

Public EntradaX                    As Byte

Public EntradaY                    As Byte

Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Public MouseX                 As Long

Public MouseY                 As Long

Public TX                     As Byte

Public TY                     As Byte

Public UltimaTextura          As Long

Public QueRender              As Byte

Public ParticleLluviaDorada   As Long

'Compresion
Public Windows_Temp_Dir       As String

'Declaraciones Ladder
Public spell_particle         As Long
Public Select_part            As Long
Public EfectoEnproceso        As Boolean
Public ColorAmbiente          As D3DCOLORVALUE
Public VSync_FPS              As Boolean
Public MostrarOnline          As Boolean
Public usersOnline            As Integer
Public meteo_particle         As Long
Public meteo_estado           As Byte
Public map_base_light         As Long
Public Map_light_base         As Long
Public Map_light_baseBackup   As Long
Public Const Particula_Lluvia As Long = 58
Public Const Particula_Nieve  As Long = 57
Public VolMusicFadding        As Integer
Public RawServersList         As String

Public Type tServerInfo

    IP As String
    puerto As Integer
    desc As String
    estado As Boolean

End Type

Public ServersLst()   As tServerInfo

Public EngineStats    As Boolean

Public CuentaDonador  As Byte

Public DeleteUser     As String

Public CuentaPassword As String

Public CuentaEmail    As String

Public NamePj(1 To 8) As String

Public ValidacionCode As String

'Objetos públicos
Public CurMp3         As Byte

Public Const Mp3_Dir = "\..\Recursos\Mp3\"

'Opciones Clasicas


'RGB Type
Public Type RGB

    r As Long
    g As Long
    b As Long

End Type

Public Type ARGB

    a As Single
    r As Long
    g As Long
    b As Long

End Type

Public ObjFile    As String

Public StreamFile As String

Public NumAuras   As Byte

Public InvOroComUsu(2)         As New clsGrapchicalInventory ' Inventarios de oro (ambos usuarios)

Public InvOfferComUsu(1)       As New clsGrapchicalInventory ' Inventarios de ofertas (ambos usuarios)

Public CustomKeys              As New clsCustomKeys

Public Sound                   As New clsSoundEngine

Public incomingData            As New clsByteQueue
Public outgoingData            As New clsByteQueue

Public Audio_MP3_Load          As Boolean

Public Audio_MP3_Play          As Boolean

''
'The main timer of the game.
Public MainTimer               As New clsTimer

'Sonidos
Public Const SND_CAM           As String = "cam.wav"

Public Const SND_EXCLAMACION   As Integer = 451

Public Const SND_CLICK         As Integer = 190

Public Const SND_NAVEGANDO     As Integer = 50

Public Const SND_OVER          As Integer = 0

Public Const SND_DICE          As Integer = 188

Public Const SND_FUEGO         As Integer = 116

Public Const SND_LLUVIAIN      As Integer = 191

Public Const SND_LLUVIAOUT     As Integer = 194

Public Const SND_NIEVEIN       As Integer = 191

Public Const SND_NIEVEOUT      As Integer = 194

Public Const SND_RESUCITAR     As Integer = 104

Public Const SND_CURAR         As Integer = 101

Public TargetXMacro            As Byte

Public TargetYMacro            As Byte

' Head index of the casper. Used to know if a char is killed

' Constantes de intervalo
Public Const INT_MACRO_HECHIS  As Integer = 500

Public Const INT_MACRO_TRABAJO As Integer = 1200

Public IntervaloGolpe          As Long

Public IntervaloArco           As Long

Public IntervaloMagia          As Long

Public IntervaloTrabajo        As Long

Public IntervaloCaminar        As Long

Public IntervaloTirar          As Long

Public IntervaloUsarU          As Long

Public IntervaloUsarClic       As Long

Public IntervaloGolpeMagia     As Long

Public IntervaloMagiaGolpe     As Long

Public IntervaloGolpeUsar      As Long

Public Const INT_SENTRPU       As Integer = 2000

Public MacroBltIndex           As Integer

Public Const CASPER_HEAD       As Integer = 621

Public Const NUMATRIBUTES      As Byte = 5

'Musica
Public Const MIdi_Inicio       As Byte = 6

Public Const Mp3_Inicio        As Byte = 1

Public MActivated              As Boolean
 
''
'States wether sound is currently activated or not
Public sActivated              As Boolean

Public Type tColor

    r As Byte
    g As Byte
    b As Byte

End Type

Public ColoresPJ(0 To 50)     As tColor

Public CurServer              As Integer

Public CreandoClan            As Boolean

Public ClanName               As String

Public Site                   As String

Public UserCiego              As Boolean

Public UserEstupido           As Boolean

Public NoRes                  As Boolean 'no cambiar la resolucion

Public Launcher               As Boolean '¿Habrio desde el Launcher?

Public AmbientalesBufferIndex As Long

Public RainBufferIndex        As Long

Public FogataBufferIndex      As Long

Public Const bCabeza = 1

Public Const bPiernaIzquierda = 2

Public Const bPiernaDerecha = 3

Public Const bBrazoDerecho = 4

Public Const bBrazoIzquierdo = 5

Public Const bTorso = 6

'Timers de GetTickCount
Public Const tAt = 2000

Public Const tUs = 600

Public Const PrimerBodyBarco = 84

Public Const UltimoBodyBarco = 87

Public NumEscudosAnims As Integer

Type tHerreria

    LHierro As Integer
    LPlata As Integer
    LOro As Integer
    Index As Integer

End Type

Type tSasteria

    PielLobo As Integer
    PielOsoPardo As Integer
    PielOsoPolar As Integer
    Index As Integer

End Type

Public ArmasHerrero(0 To 100)     As tHerreria

Public DefensasHerrero(0 To 100)  As tHerreria

Public ArmadurasHerrero(0 To 100) As tHerreria

Public CascosHerrero(0 To 100)    As tHerreria

Public EscudosHerrero(0 To 100)   As tHerreria

Public ObjCarpintero(0 To 100)    As Integer

Public ObjAlquimista(0 To 100)    As Integer

Public ObjSastre(0 To 100)        As tSasteria

Public SastreRopas(0 To 100)      As tSasteria

Public SastreGorros(0 To 100)     As tSasteria

Public UsaLanzar                  As Boolean

Public UsaMacro                   As Boolean

Public CnTd                       As Byte

Public Type ObjDonador

    precio As Integer
    Index As Integer

End Type

Public ObjDonador(0 To 100)           As ObjDonador

Public Const MAX_BANCOINVENTORY_SLOTS As Byte = 42

Public Const MAX_KEYS As Byte = 10

Public Const LoopAdEternum            As Integer = 999

'Direcciones
Public Enum E_Heading

    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4

End Enum

Public Enum eBlock

    NORTH = &H1
    EAST = &H2
    SOUTH = &H4
    WEST = &H8
    ALL_SIDES = &HF
    GM = &H10

End Enum

'Objetos
Public Const MAX_INVENTORY_OBJS      As Integer = 10000

Public Const MAX_NPC_INVENTORY_SLOTS As Byte = 36

Public Const MAXHECHI                As Byte = 150

Public Const MAXSKILLPOINTS          As Byte = 100

Public Const FLAGORO                 As Integer = 200

Public Const FLAG_AGUA               As Byte = &H20

Public Const FLAG_ARBOL              As Byte = &H40

Public Const PRIMER_TRIGGER_TECHO    As Byte = 9

Public Const FOgata                  As Integer = 1521

Public Enum eClass

    Mage = 1    'Mago
    Cleric      'Clérigo
    Warrior     'Guerrero
    Assasin     'Asesino
    Bard        'Bardo
    Druid       'Druida
    paladin     'Paladín
    Hunter      'Cazador
    Trabajador  'Trabajador

End Enum

Public Enum eCiudad

    cUllathorpe = 1
    cNix
    cBanderbill
    cLindos
    cArghal
    cHillidan

End Enum

Enum eRaza

    Humano = 1
    Elfo
    ElfoOscuro
    Gnomo
    Enano
    Orco

End Enum

Public Enum eSkill

    magia = 1
    Robar = 2
    Tacticas = 3
    Armas = 4
    Meditar = 5
    Apuñalar = 6
    Ocultarse = 7
    Supervivencia = 8
    Comerciar = 9
    Defensa = 10
    Liderazgo = 11
    Proyectiles = 12
    Wrestling = 13
    Navegacion = 14
    equitacion = 15
    Resistencia = 16
    
    Talar = 17
    Pescar = 18
    Mineria = 19
    Herreria = 20
    Carpinteria = 21
    Alquimia = 22
    Sastreria = 23
    Domar = 24
    
    Grupo = 90
    MarcaDeClan = 91
    MarcaDeGM = 92

End Enum

Public Enum eAtributos

    Fuerza = 1
    Agilidad = 2
    Inteligencia = 3
    Constitucion = 4
    Carisma = 5

End Enum

Enum eGenero

    Hombre = 1
    Mujer

End Enum

Public Enum PlayerType

    User = &H1
    Consejero = &H2
    SemiDios = &H4
    Dios = &H8
    Admin = &H10
    RoleMaster = &H20
    ChaosCouncil = &H40
    RoyalCouncil = &H80

End Enum

Public Enum eObjType

    otUseOnce = 1
    otWeapon = 2
    otArmadura = 3
    otArboles = 4
    otGuita = 5
    otPuertas = 6
    otContenedores = 7
    otCarteles = 8
    otLlaves = 9
    otForos = 10
    otPociones = 11
    otBebidas = 13
    otLeña = 14
    otFogata = 15
    otESCUDO = 16
    otCASCO = 17
    OtHerramientas = 18
    otTeleport = 19
    OtDecoraciones = 20
    otmagicos = 21
    otYacimiento = 22
    otMinerales = 23
    otPergaminos = 24
    otInstrumentos = 26
    otYunque = 27
    otFragua = 28
    otAnillos = 30
    otBarcos = 31
    otFlechas = 32
    otBotellaVacia = 33
    otBotellaLlena = 34
    otManchas = 35          'No se usa
    otpasajes = 36
    otmapa = 38
    OtPozos = 40
    otMonturas = 44
    otRunas = 45
    otNudillos = 46
    OtCorreo = 47
    OtCofre = 48
    OtDonador = 50
    otCualquiera = 1000

End Enum

Public Const FundirMetal                           As Integer = 88

'
' Mensajes
'
' MENSAJE_*  --> Mensajes de texto que se muestran en el cuadro de texto
'

Public Const MENSAJE_CRIATURA_FALLA_GOLPE          As String = "La criatura fallo el golpe."

Public Const MENSAJE_CRIATURA_MATADO               As String = "La criatura te ha matado."

Public Const MENSAJE_RECHAZO_ATAQUE_ESCUDO         As String = "Has rechazado el ataque con el escudo."

Public Const MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO As String = "El usuario rechazo el ataque con su escudo."

Public Const MENSAJE_FALLADO_GOLPE                 As String = "Has fallado el golpe."

Public Const MENSAJE_SEGURO_ACTIVADO               As String = "Seguro Activado."

Public Const MENSAJE_SEGURO_DESACTIVADO            As String = "Seguro Desactivado."

Public Const MENSAJE_USAR_MEDITANDO                As String = "¡Estás meditando! Debes dejar de meditar para usar objetos."

Public Const MENSAJE_SEGURO_PARTY_ON               As String = "Ahora nadie te podra invitar a un grupo."

Public Const MENSAJE_SEGURO_PARTY_OFF              As String = "Ahora podras recibir solicitudes a grupos."

Public Const MENSAJE_GOLPE_CABEZA                  As String = "La criatura te ha pegado en la cabeza por "

Public Const MENSAJE_GOLPE_BRAZO_IZQ               As String = "La criatura te ha pegado el brazo izquierdo por "

Public Const MENSAJE_GOLPE_BRAZO_DER               As String = "La criatura te ha pegado el brazo derecho por "

Public Const MENSAJE_GOLPE_PIERNA_IZQ              As String = "La criatura te ha pegado la pierna izquierda por "

Public Const MENSAJE_GOLPE_PIERNA_DER              As String = "La criatura te ha pegado la pierna derecha por "

Public Const MENSAJE_GOLPE_TORSO                   As String = "La criatura te ha pegado en el torso por "

' MENSAJE_[12]: Aparecen antes y despues del valor de los mensajes anteriores (MENSAJE_GOLPE_*)
Public Const MENSAJE_1                             As String = "¡¡"

Public Const MENSAJE_2                             As String = "."

Public Const MENSAJE_GOLPE_CRIATURA_1              As String = "Le has pegado a la criatura por "

Public Const MENSAJE_ATAQUE_FALLO                  As String = " te ataco y fallo."

Public Const MENSAJE_RECIVE_IMPACTO_CABEZA         As String = " te ha pegado en la cabeza por "

Public Const MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ      As String = " te ha pegado el brazo izquierdo por "

Public Const MENSAJE_RECIVE_IMPACTO_BRAZO_DER      As String = " te ha pegado el brazo derecho por "

Public Const MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ     As String = " te ha pegado la pierna izquierda por "

Public Const MENSAJE_RECIVE_IMPACTO_PIERNA_DER     As String = " te ha pegado la pierna derecha por "

Public Const MENSAJE_RECIVE_IMPACTO_TORSO          As String = " te ha pegado en el torso por "

Public Const MENSAJE_PRODUCE_IMPACTO_1             As String = "Le has pegado a "

Public Const MENSAJE_PRODUCE_IMPACTO_CABEZA        As String = " en la cabeza por "

Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ     As String = " en el brazo izquierdo por "

Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_DER     As String = " en el brazo derecho por "

Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ    As String = " en la pierna izquierda por "

Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_DER    As String = " en la pierna derecha por "

Public Const MENSAJE_PRODUCE_IMPACTO_TORSO         As String = " en el torso por "

Public Const MENSAJE_TRABAJO_MAGIA                 As String = "Haz click sobre el objetivo..."

Public Const MENSAJE_TRABAJO_PESCA                 As String = "Haz click sobre el sitio donde quieres pescar..."

Public Const MENSAJE_TRABAJO_ROBAR                 As String = "Haz click sobre la victima..."

Public Const MENSAJE_TRABAJO_TALAR                 As String = "Haz click sobre el árbol..."

Public Const MENSAJE_TRABAJO_MINERIA               As String = "Haz click sobre el yacimiento..."

Public Const MENSAJE_TRABAJO_FUNDIRMETAL           As String = "Haz click sobre la fragua..."

Public Const MENSAJE_TRABAJO_PROYECTILES           As String = "Haz click sobre la victima..."

Public Const MENSAJE_NENE                          As String = "Cantidad de NPCs: "

'Inventario
Type Inventory

    OBJIndex As Integer
    Name As String
    GrhIndex As Long
    '[Alejo]: tipo de datos ahora es Long
    Amount As Long
    '[/Alejo]
    Equipped As Byte
    Valor As Single
    ObjType As Integer
    Def As Integer
    MaxHit As Integer
    MinHit As Integer
    PuedeUsar As Byte

End Type

Type MakeObj

    GrhIndex As Long ' Indice del grafico que representa el obj
    Name As String
    MinDef As Integer
    MaxDef As Integer
    MinHit As Integer
    MaxHit As Integer
    ObjType As Byte

End Type

Type NpCinV

    OBJIndex As Integer
    Name As String
    GrhIndex As Long
    Amount As Integer
    Valor As Single
    PuedeUsar As Byte
    ObjType As Integer
    Def As Integer
    MaxHit As Integer
    MinHit As Integer
    C1 As String
    C2 As String
    C3 As String
    C4 As String
    C5 As String
    C6 As String
    C7 As String

End Type

Type tEstadisticasUsu

    CiudadanosMatados As Long
    CriminalesMatados As Long
    UsuariosMatados As Long
    Alineacion As Byte
    NpcsMatados As Long
    Clase As String
    PenaCarcel As Long
    VecesQueMoriste As Long
    Genero As String
    Raza As String
    Donador As Byte
    CreditoDonador As Long
    DiasRestantes As Integer
    BattlePuntos As Long

End Type

Public Nombres                                  As Boolean

'User status vars
Global OtroInventario(1 To MAX_INVENTORY_SLOTS) As Inventory

Public UserInventario(1 To MAX_INVENTORY_SLOTS) As Inventory

Public UserHechizos(1 To MAXHECHI)              As Integer

Public UserHechizosInterval(1 To MAXHECHI)      As Integer

Public UserMeditar                              As Boolean

Public UserName                                 As String

Public UserPassword                             As String

Public UserMaxHp                                As Integer

Public UserMinHp                                As Integer

Public UserMaxMAN                               As Integer

Public UserMinMAN                               As Integer

Public UserMaxSTA                               As Integer

Public UserMinSTA                               As Integer

Public UserMaxAGU                               As Byte

Public UserMinAGU                               As Byte

Public UserMaxHAM                               As Byte

Public UserMinHAM                               As Byte

Public UserGLD                                  As Long

Public UserLvl                                  As Integer

Public UserPort                                 As Integer

Public UserServerIP                             As String

Public UserEstado                               As Byte '0 = Vivo & 1 = Muerto

Public UserPasarNivel                           As Long

Public UserExp                                  As Long

Public UserEstadisticas                         As tEstadisticasUsu

Public UserDescansar                            As Boolean

Public Moviendose                               As Boolean

Public FPSFLAG                                  As Byte

Public VSync                                    As Byte

Public AlphaMacro                               As Byte

Public pausa                                    As Boolean

Public UserParalizado                           As Boolean

Public UserInmovilizado                         As Boolean

Public UserNavegando                            As Boolean

Public UserMontado                              As Boolean

Public UserNadando                              As Boolean

Public UserAvisado                              As Boolean

Public UserAvisadoBarca                         As Boolean

Public UserSaliendo                             As Boolean

'<-------------------------NUEVO-------------------------->
Public Comerciando                              As Boolean
'<-------------------------NUEVO-------------------------->

Public UserClase                                As eClass

Public UserSexo                                 As eGenero

Public UserRaza                                 As eRaza

'Declaraciones LADDER!
Public SendingType                              As Byte

Public sndPrivateTo                             As String

Public Const NUMSKILLS                          As Byte = 24

Public Const NUMATRIBUTOS                       As Byte = 5

Public Const NUMCLASES                          As Byte = 9

Public Const NUMRAZAS                           As Byte = 5

Public Const NUMCIUDADES                        As Byte = 6

Type tModRaza

    Fuerza As Integer
    Agilidad As Integer
    Inteligencia As Integer
    Constitucion As Integer
    Carisma As Integer

End Type

Public ModRaza(1 To NUMRAZAS)            As tModRaza

Public ListaCiudades(1 To NUMCIUDADES)   As String

Public UserSkills(1 To NUMSKILLS)        As Byte

Public SkillsNames(1 To NUMSKILLS)       As String

Public SkillsDesc(1 To NUMSKILLS)        As String

Public UserAtributos(1 To NUMATRIBUTOS)  As Byte

Public AtributosNames(1 To NUMATRIBUTOS) As String

Public ListaRazas(1 To NUMRAZAS)         As String

Public ListaClases(1 To NUMCLASES)       As String

Public SkillPoints                       As Integer

Public Alocados                          As Integer

Public flags()                           As Integer

Public Oscuridad                         As Integer

Public logged                            As Boolean

Public UsingSkill                        As Integer

Public pingTime                          As Long

Public Enum E_MODO

    Normal = 1
    CrearNuevoPj = 2
    Dados = 3
    CreandoCuenta = 4
    ActivandoCuenta = 5
    IngresandoConCuenta = 6
    ReValidandoCuenta = 7
    BorrandoPJ = 8
    RecuperandoConstraseña = 9
    BorrandoCuenta = 10

End Enum

Public EstadoLogin As E_MODO

Public Enum eClanType

    ct_RoyalArmy
    ct_Evil
    ct_Neutral
    ct_GM
    ct_Legal
    ct_Criminal

End Enum

Public Enum eEditOptions

    eo_Gold = 1
    eo_Experience
    eo_Body
    eo_Head
    eo_CiticensKilled
    eo_CriminalsKilled
    eo_Level
    eo_Class
    eo_Skills
    eo_SkillPointsLeft
    eo_Sex
    eo_Raza

End Enum

''
' TRIGGERS
'
' @param NADA nada
' @param BAJOTECHO bajo techo
' @param trigger_2 ???
' @param POSINVALIDA los npcs no pueden pisar tiles con este trigger
' @param ZONASEGURA no se puede robar o pelear desde este trigger
' @param ANTIPIQUETE
' @param ZONAPELEA al pelear en este trigger no se caen las cosas y no cambia el estado de ciuda o crimi
'
Public Enum eTrigger

    NADA = 0
    BAJOTECHO = 1
    trigger_2 = 2
    POSINVALIDA = 3
    ZONASEGURA = 4
    ANTIPIQUETE = 5
    ZONAPELEA = 6
    CURA = 7

End Enum

'Server stuff
Public RequestPosTimer   As Integer 'Used in main loop

Public stxtbuffer        As String 'Holds temp raw data from server

Public stxtbuffercmsg    As String 'Holds temp raw data from server

Public SendNewChar       As Boolean 'Used during login

Public Connected         As Boolean 'True when connected to server

Public DownloadingMap    As Boolean 'Currently downloading a map from server

Public UserMap           As Integer

'Control
Public prgRun            As Boolean 'When true the program ends

Public IPdelServidor     As String

Public PuertoDelServidor As String

'
'********** FUNCIONES API ***********
'

Public Declare Function GetTickCount Lib "kernel32" () As Long

'para escribir y leer variables
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long

Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'Teclado
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Para ejecutar el Internet Explorer para el manual
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Lista de cabezas
Public Type tIndiceCabeza

    Head(1 To 4) As Long

End Type

Public Type tIndiceCuerpo

    Body(1 To 4) As Long
    HeadOffsetX As Integer
    HeadOffsetY As Integer

End Type

Public Type tIndiceFx

    Animacion As Long
    OffsetX As Integer
    OffsetY As Integer
    IsPNG As Integer

End Type

#If AntiExternos Then

    Public Security As New clsSecurity
#End If

