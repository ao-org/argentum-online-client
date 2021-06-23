Attribute VB_Name = "Protocol"
'**************************************************************
' Protocol.bas - Handles all incoming / outgoing messages for client-server communications.
' Uses a binary protocol designed by myself.
'
' Designed and implemented by Juan Martín Sotuyo Dodero (Maraxus)
' (juansotuyo@gmail.com)
'**************************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'**************************************************************************

''
'Handles all incoming / outgoing packets for client - server communications
'The binary prtocol here used was designed by Juan Martín Sotuyo Dodero.
'This is the first time it's used in Alkon, though the second time it's coded.
'This implementation has several enhacements from the first design.
'
' @file     Protocol.bas
' @author   Juan Martín Sotuyo Dodero (Maraxus) juansotuyo@gmail.com
' @version  1.0.0
' @date     20060517

Option Explicit

''
' TODO : /BANIP y /UNBANIP ya no trabajan con nicks. Esto lo puede mentir en forma local el cliente con un paquete a NickToIp

''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Private Const SEPARATOR As String * 1 = vbNullChar

Private LastPacket      As Byte

Private IterationsHID   As Integer

Private Const MAX_ITERATIONS_HID = 200

Private Enum ServerPacketID

    logged                  ' LOGGED  0
    RemoveDialogs           ' QTDL
    RemoveCharDialog        ' QDL
    NavigateToggle          ' NAVEG
    EquiteToggle
    Disconnect              ' FINOK
    CommerceEnd             ' FINCOMOK
    BankEnd                 ' FINBANOK
    CommerceInit            ' INITCOM
    BankInit                ' INITBANCO
    UserCommerceInit        ' INITCOMUSU   10
    UserCommerceEnd         ' FINCOMUSUOK
    ShowBlacksmithForm      ' SFH
    ShowCarpenterForm       ' SFC
    NPCKillUser             ' 6
    BlockedWithShieldUser   ' 7
    BlockedWithShieldOther  ' 8
    CharSwing               ' U1
    SafeModeOn              ' SEGON
    SafeModeOff             ' SEGOFF 20
    PartySafeOn
    PartySafeOff
    CantUseWhileMeditating  ' M!
    UpdateSta               ' ASS
    UpdateMana              ' ASM
    UpdateHP                ' ASH
    UpdateGold              ' ASG
    UpdateExp               ' ASE 30
    ChangeMap               ' CM
    PosUpdate               ' PU
    NPCHitUser              ' N2
    UserHitNPC              ' U2
    UserAttackedSwing       ' U3
    UserHittedByUser        ' N4
    UserHittedUser          ' N5
    ChatOverHead            ' ||
    ConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
    GuildChat               ' |+   40
    ShowMessageBox          ' !!
    MostrarCuenta
    UserIndexInServer       ' IU
    UserCharIndexInServer   ' IP
    CharacterCreate         ' CC
    CharacterRemove         ' BP
    CharacterMove           ' MP, +, * and _ '
    ForceCharMove
    CharacterChange         ' CP
    ObjectCreate            ' HO
    fxpiso
    ObjectDelete            ' BO  50
    BlockPosition           ' BQ
    PlayMIDI                ' TM
    PlayWave                ' TW
    guildList               ' GL
    AreaChanged             ' CA
    PauseToggle             ' BKW
    RainToggle              ' LLU
    CreateFX                ' CFX
    UpdateUserStats         ' EST
    WorkRequestTarget       ' T01 60
    ChangeInventorySlot     ' CSI
    InventoryUnlockSlots
    ChangeBankSlot          ' SBO
    ChangeSpellSlot         ' SHS
    Atributes               ' ATR
    BlacksmithWeapons       ' LAH
    BlacksmithArmors        ' LAR
    CarpenterObjects        ' OBR
    RestOK                  ' DOK
    ErrorMsg                ' ERR
    Blind                   ' CEGU 70
    Dumb                    ' DUMB
    ShowSignal              ' MCAR
    ChangeNPCInventorySlot  ' NPCI
    UpdateHungerAndThirst   ' EHYS
    MiniStats               ' MEST
    LevelUp                 ' SUNI
    AddForumMsg             ' FMSG
    ShowForumForm           ' MFOR
    SetInvisible            ' NOVER 80
    DiceRoll                ' DADOS
    MeditateToggle          ' MEDOK
    BlindNoMore             ' NSEGUE
    DumbNoMore              ' NESTUP
    SendSkills              ' SKILLS
    TrainerCreatureList     ' LSTCRI
    guildNews               ' GUILDNE
    OfferDetails            ' PEACEDE & ALLIEDE
    AlianceProposalsList    ' ALLIEPR
    PeaceProposalsList      ' PEACEPR 90
    CharacterInfo           ' CHRINFO
    GuildLeaderInfo         ' LEADERI
    GuildDetails            ' CLANDET
    ShowGuildFundationForm  ' SHOWFUN
    ParalizeOK              ' PARADOK
    ShowUserRequest         ' PETICIO
    ChangeUserTradeSlot     ' COMUSUINV
    SendNight               ' NOC
    Pong
    UpdateTagAndStatus
    FYA
    CerrarleCliente
    Contadores
    
    'GM messages
    SpawnList               ' SPL
    ShowSOSForm             ' MSOS
    ShowMOTDEditionForm     ' ZMOTD
    ShowGMPanelForm         ' ABPANEL
    UserNameList            ' LISTUSU
    PersonajesDeCuenta
    UserOnline '110
    ParticleFX
    ParticleFXToFloor
    ParticleFXWithDestino
    ParticleFXWithDestinoXY
    hora
    Light
    AuraToChar
    SpeedToChar
    LightToFloor
    NieveToggle
    NieblaToggle
    Goliath
    TextOverChar
    TextOverTile
    TextCharDrop
    FlashScreen
    AlquimistaObj
    ShowAlquimiaForm
    familiar
    SastreObj
    ShowSastreForm ' 126
    VelocidadToggle
    MacroTrabajoToggle
    RefreshAllInventorySlot
    BindKeys
    ShowFrmLogear
    ShowFrmMapa
    InmovilizadoOK
    BarFx
    SetEscribiendo
    Logros
    TrofeoToggleOn
    TrofeoToggleOff
    LocaleMsg
    ListaCorreo
    ShowPregunta
    DatosGrupo
    ubicacion
    CorreoPicOn
    DonadorObj
    ArmaMov
    EscudoMov
    ActShop
    ViajarForm
    oxigeno
    NadarToggle
    ShowFundarClanForm
    CharUpdateHP
    CharUpdateMAN
    Ranking
    PosLLamadaDeClan
    QuestDetails
    QuestListSend
    NpcQuestListSend
    UpdateNPCSimbolo
    ClanSeguro
    Intervals
    UpdateUserKey
    UpdateRM
    UpdateDM
    RequestScreenShot
    ShowScreenShot
    ScreenShotData
    Tolerancia0
    SeguroResu
    Stopped
    InvasionInfo
    CommerceRecieveChatMessage
    DoAnimation
    OpenCrafting
    CraftingItem
    CraftingCatalyst
    CraftingResult
    ForceUpdate
    
    [PacketCount]
End Enum

Public Enum ClientPacketID

    LoginExistingChar       'OLOGIN
    LoginNewChar            'NLOGIN
    ThrowDice               'TirarDados
    Talk                    ';
    Yell                    '-
    Whisper                 '\
    Walk
    RequestPositionUpdate   'RPU
    Attack                  'AT
    PickUp                  'AG
    SafeToggle              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
    PartySafeToggle
    RequestGuildLeaderInfo  'GLINFO
    RequestAtributes        'ATR
    RequestSkills           'ESKI
    RequestMiniStats        'FEST
    CommerceEnd             'FINCOM
    UserCommerceEnd         'FINCOMUSU
    BankEnd                 'FINBAN
    UserCommerceOk          'COMUSUOK
    UserCommerceReject      'COMUSUNO
    Drop                    'TI
    CastSpell               'LH
    LeftClick               'LC
    DoubleClick             'RC
    Work                    'UK
    UseSpellMacro           'UMH
    UseItem                 'USA
    CraftBlacksmith         'CNS
    CraftCarpenter          'CNC
    WorkLeftClick           'WLC
    CreateNewGuild          'CIG
    SpellInfo               'INFS
    EquipItem               'EQUI
    ChangeHeading           'CHEA
    ModifySkills            'SKSE
    Train                   'ENTR
    CommerceBuy             'COMP
    BankExtractItem         'RETI
    CommerceSell            'VEND
    BankDeposit             'DEPO
    ForumPost               'DEMSG
    MoveSpell               'DESPHE
    ClanCodexUpdate         'DESCOD
    UserCommerceOffer       'OFRECER
    GuildAcceptPeace        'ACEPPEAT
    GuildRejectAlliance     'RECPALIA
    GuildRejectPeace        'RECPPEAT
    GuildAcceptAlliance     'ACEPALIA
    GuildOfferPeace         'PEACEOFF
    GuildOfferAlliance      'ALLIEOFF
    GuildAllianceDetails    'ALLIEDET
    GuildPeaceDetails       'PEACEDET
    GuildRequestJoinerInfo  'ENVCOMEN
    GuildAlliancePropList   'ENVALPRO
    GuildPeacePropList      'ENVPROPP
    GuildDeclareWar         'DECGUERR
    GuildNewWebsite         'NEWWEBSI
    GuildAcceptNewMember    'ACEPTARI
    GuildRejectNewMember    'RECHAZAR
    GuildKickMember         'ECHARCLA
    GuildUpdateNews         'ACTGNEWS
    GuildMemberInfo         '1HRINFO<
    GuildOpenElections      'ABREELEC
    GuildRequestMembership  'SOLICITUD
    GuildRequestDetails     'CLANDETAILS
    Online                  '/ONLINE
    Quit                    '/SALIR
    GuildLeave              '/SALIRCLAN
    RequestAccountState     '/BALANCE
    PetStand                '/QUIETO
    PetFollow               '/ACOMPAÑAR
    PetLeave                '/LIBERAR
    GrupoMsg                '/GrupoMsg
    TrainList               '/ENTRENAR
    Rest                    '/DESCANSAR
    Meditate                '/MEDITAR
    Resucitate              '/RESUCITAR
    Heal                    '/CURAR
    Help                    '/AYUDA
    RequestStats            '/EST
    CommerceStart           '/COMERCIAR
    BankStart               '/BOVEDA
    Enlist                  '/ENLISTAR
    Information             '/INFORMACION
    Reward                  '/RECOMPENSA
    RequestMOTD             '/MOTD
    UpTime                  '/UPTIME
    Inquiry                 '/ENCUESTA ( with no params )
    GuildMessage            '/CMSG
    CentinelReport          '/CENTINELA
    GuildOnline             '/ONLINECLAN
    CouncilMessage          '/BMSG
    RoleMasterRequest       '/ROL
    GMRequest               '/GM
    ChangeDescription       '/DESC
    GuildVote               '/VOTO
    punishments             '/PENAS
    ChangePassword          '/CONTRASEÑA
    Gamble                  '/APOSTAR
    InquiryVote             '/ENCUESTA ( with parameters )
    LeaveFaction            '/RETIRAR ( with no arguments )
    BankExtractGold         '/RETIRAR ( with arguments )
    BankDepositGold         '/DEPOSITAR
    Denounce                '/DENUNCIAR
    Ping                    '/PING
    
    'GM messages
    GMMessage               '/GMSG
    showName                '/SHOWNAME
    OnlineRoyalArmy         '/ONLINEREAL
    OnlineChaosLegion       '/ONLINECAOS
    GoNearby                '/IRCERCA
    comment                 '/REM
    serverTime              '/HORA
    Where                   '/DONDE
    CreaturesInMap          '/NENE
    WarpMeToTarget          '/TELEPLOC
    WarpChar                '/TELEP
    Silence                 '/SILENCIAR
    SOSShowList             '/SHOW SOS
    SOSRemove               'SOSDONE
    GoToChar                '/IRA
    Invisible               '/INVISIBLE
    GMPanel                 '/PANELGM
    RequestUserList         'LISTUSU
    Working                 '/TRABAJANDO
    Hiding                  '/OCULTANDO
    Jail                    '/CARCEL
    KillNPC                 '/RMATA
    WarnUser                '/ADVERTENCIA
    EditChar                '/MOD
    RequestCharInfo         '/INFO
    RequestCharStats        '/STAT
    RequestCharGold         '/BAL
    RequestCharInventory    '/INV
    RequestCharBank         '/BOV
    RequestCharSkills       '/SKILLS
    ReviveChar              '/REVIVIR
    OnlineGM                '/ONLINEGM
    OnlineMap               '/ONLINEMAP
    Forgive                 '/PERDON
    Kick                    '/ECHAR
    Execute                 '/EJECUTAR
    BanChar                 '/BAN
    UnbanChar               '/UNBAN
    NPCFollow               '/SEGUIR
    SummonChar              '/SUM
    SpawnListRequest        '/CC
    SpawnCreature           'SPA
    ResetNPCInventory       '/RESETINV
    CleanWorld              '/LIMPIAR
    ServerMessage           '/RMSG
    NickToIP                '/NICK2IP
    IPToNick                '/IP2NICK
    GuildOnlineMembers      '/ONCLAN
    TeleportCreate          '/CT
    TeleportDestroy         '/DT
    RainToggle              '/LLUVIA
    SetCharDescription      '/SETDESC
    ForceMIDIToMap          '/FORCEMIDIMAP
    ForceWAVEToMap          '/FORCEWAVMAP
    RoyalArmyMessage        '/REALMSG
    ChaosLegionMessage      '/CAOSMSG
    CitizenMessage          '/CIUMSG
    CriminalMessage         '/CRIMSG
    TalkAsNPC               '/TALKAS
    DestroyAllItemsInArea   '/MASSDEST
    AcceptRoyalCouncilMember '/ACEPTCONSE
    AcceptChaosCouncilMember '/ACEPTCONSECAOS
    ItemsInTheFloor         '/PISO
    MakeDumb                '/ESTUPIDO
    MakeDumbNoMore          '/NOESTUPIDO
    DumpIPTables            '/DUMPSECURITY
    CouncilKick             '/KICKCONSE
    SetTrigger              '/TRIGGER
    AskTrigger              '/TRIGGER with no arguments
    BannedIPList            '/BANIPLIST
    BannedIPReload          '/BANIPRELOAD
    GuildMemberList         '/MIEMBROSCLAN
    GuildBan                '/BANCLAN
    banip                   '/BANIP
    UnBanIp                 '/UNBANIP
    CreateItem              '/CI
    DestroyItems            '/DEST
    ChaosLegionKick         '/NOCAOS
    RoyalArmyKick           '/NOREAL
    ForceMIDIAll            '/FORCEMIDI
    ForceWAVEAll            '/FORCEWAV
    RemovePunishment        '/BORRARPENA
    TileBlockedToggle       '/BLOQ
    KillNPCNoRespawn        '/MATA
    KillAllNearbyNPCs       '/MASSKILL
    LastIP                  '/LASTIP
    ChangeMOTD              '/MOTDCAMBIA
    SetMOTD                 'ZMOTD
    SystemMessage           '/SMSG
    CreateNPC               '/ACC
    CreateNPCWithRespawn    '/RACC
    ImperialArmour          '/AI1 - 4
    ChaosArmour             '/AC1 - 4
    NavigateToggle          '/NAVE
    ServerOpenToUsersToggle '/HABILITAR
    Participar              '/APAGAR
    TurnCriminal            '/CONDEN
    ResetFactions           '/RAJAR
    RemoveCharFromGuild     '/RAJARCLAN
    RequestCharMail         '/LASTEMAIL
    AlterPassword           '/APASS
    AlterMail               '/AEMAIL
    AlterName               '/ANAME
    DoBackUp                '/DOBACKUP
    ShowGuildMessages       '/SHOWCMSG
    SaveMap                 '/GUARDAMAPA
    ChangeMapInfoPK         '/MODMAPINFO PK
    ChangeMapInfoBackup     '/MODMAPINFO BACKUP
    ChangeMapInfoRestricted '/MODMAPINFO RESTRINGIR
    ChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
    ChangeMapInfoNoInvi     '/MODMAPINFO INVISINEFECTO
    ChangeMapInfoNoResu     '/MODMAPINFO RESUSINEFECTO
    ChangeMapInfoLand       '/MODMAPINFO TERRENO
    ChangeMapInfoZone       '/MODMAPINFO ZONA
    SaveChars               '/GRABAR
    CleanSOS                '/BORRAR SOS
    ShowServerForm          '/SHOW INT
    night                   '/NOCHE
    KickAllChars            '/ECHARTODOSPJS
    RequestTCPStats         '/TCPESSTATS
    ReloadNPCs              '/RELOADNPCS
    ReloadServerIni         '/RELOADSINI
    ReloadSpells            '/RELOADHECHIZOS
    ReloadObjects           '/RELOADOBJ
    Restart                 '/REINICIAR
    ResetAutoUpdate         '/AUTOUPDATE
    ChatColor               '/CHATCOLOR
    Ignored                 '/IGNORADO
    CheckSlot               '/SLOT
    
    'Nuevas Ladder
    GlobalMessage           '/CONSOLA
    GlobalOnOff
    IngresarConCuenta
    BorrarPJ
    Desbuggear
    DarLlaveAUsuario
    SacarLlave
    VerLlaves
    UseKey
    Day
    SetTime
    DonateGold              '/DONAR
    Promedio                '/PROMEDIO
    GiveItem                '/DAR
    OfertaInicial
    OfertaDeSubasta
    QuestionGM
    CuentaRegresiva
    PossUser
    Duel
    AcceptDuel
    CancelDuel
    QuitDuel
    NieveToggle
    NieblaToggle
    TransFerGold
    Moveitem
    Genio                 '/GENIO
    Casarse
    CraftAlquimista
    RequestFamiliar
    FlagTrabajar
    CraftSastre
    MensajeUser
    TraerBoveda
    CompletarAccion
    Escribiendo
    TraerRecompensas
    ReclamarRecompensa
    Correo
    SendCorreo
    RetirarItemCorreo
    BorrarCorreo
    InvitarGrupo
    ResponderPregunta
    RequestGrupo
    AbandonarGrupo
    HecharDeGrupo
    MacroPossent
    SubastaInfo
    BanCuenta
    UnbanCuenta
    BanSerial
    unBanSerial
    CerrarCliente
    EventoInfo
    CrearEvento
    BanTemporal
    Traershop
    ComprarItem
    SCROLLINFO
    CancelarExit
    EnviarCodigo
    CrearTorneo
    ComenzarTorneo
    CancelarTorneo
    BusquedaTesoro
    CompletarViaje
    BovedaMoveItem
    QuieroFundarClan
    llamadadeclan
    MarcaDeClanPack
    MarcaDeGMPack
    TraerRanking
    Pareja
    Quest
    QuestAccept
    QuestListRequest
    QuestDetailsRequest
    QuestAbandon
    SeguroClan
    CreatePretorianClan     '/CREARPRETORIANOS
    RemovePretorianClan     '/ELIMINARPRETORIANOS
    Home                    '/HOGAR
    Consulta                '/CONSULTA
    RequestScreenShot       '/SS
    SendScreenShot
    Tolerancia0
    GetMapInfo
    FinEvento
    SeguroResu
    CuentaExtractItem
    CuentaDeposit
    CreateEvent
    CommerceSendChatMessage
    LogMacroClickHechizo
    AddItemCrafting
    RemoveItemCrafting
    AddCatalyst
    RemoveCatalyst
    CraftItem
    CloseCrafting
    MoveCraftItem
    PetLeaveAll
    
    [PacketCount]
End Enum

' Rezniaq: Sacamos alv la busqueda lineal que hacia el Select Case de la funcion HandleIncomingData.
Private PacketList(0 To ServerPacketID.[PacketCount] - 1) As Long
Private Declare Sub CallHandle Lib "ao20.dll" (ByVal address As Long, ByVal userIndex As Integer)

Public Sub InitializePacketList()

    PacketList(ServerPacketID.logged) = GetAddress(AddressOf HandleLogged)
    PacketList(ServerPacketID.RemoveDialogs) = GetAddress(AddressOf HandleRemoveDialogs)
    PacketList(ServerPacketID.RemoveCharDialog) = GetAddress(AddressOf HandleRemoveCharDialog)
    PacketList(ServerPacketID.NavigateToggle) = GetAddress(AddressOf HandleNavigateToggle)
    PacketList(ServerPacketID.EquiteToggle) = GetAddress(AddressOf HandleEquiteToggle)
    PacketList(ServerPacketID.Disconnect) = GetAddress(AddressOf HandleDisconnect)
    PacketList(ServerPacketID.CommerceEnd) = GetAddress(AddressOf HandleCommerceEnd)
    PacketList(ServerPacketID.BankEnd) = GetAddress(AddressOf HandleBankEnd)
    PacketList(ServerPacketID.CommerceInit) = GetAddress(AddressOf HandleCommerceInit)
    PacketList(ServerPacketID.BankInit) = GetAddress(AddressOf HandleBankInit)
    PacketList(ServerPacketID.UserCommerceInit) = GetAddress(AddressOf HandleUserCommerceInit)
    PacketList(ServerPacketID.UserCommerceEnd) = GetAddress(AddressOf HandleUserCommerceEnd)
    PacketList(ServerPacketID.ShowBlacksmithForm) = GetAddress(AddressOf HandleShowBlacksmithForm)
    PacketList(ServerPacketID.ShowCarpenterForm) = GetAddress(AddressOf HandleShowCarpenterForm)
    PacketList(ServerPacketID.NPCKillUser) = GetAddress(AddressOf HandleNPCKillUser)
    PacketList(ServerPacketID.BlockedWithShieldUser) = GetAddress(AddressOf HandleBlockedWithShieldUser)
    PacketList(ServerPacketID.BlockedWithShieldOther) = GetAddress(AddressOf HandleBlockedWithShieldOther)
    PacketList(ServerPacketID.CharSwing) = GetAddress(AddressOf HandleCharSwing)
    PacketList(ServerPacketID.SafeModeOn) = GetAddress(AddressOf HandleSafeModeOn)
    PacketList(ServerPacketID.SafeModeOff) = GetAddress(AddressOf HandleSafeModeOff)
    PacketList(ServerPacketID.PartySafeOn) = GetAddress(AddressOf HandlePartySafeOn)
    PacketList(ServerPacketID.PartySafeOff) = GetAddress(AddressOf HandlePartySafeOff)
    PacketList(ServerPacketID.CantUseWhileMeditating) = GetAddress(AddressOf HandleCantUseWhileMeditating)
    PacketList(ServerPacketID.UpdateSta) = GetAddress(AddressOf HandleUpdateSta)
    PacketList(ServerPacketID.UpdateMana) = GetAddress(AddressOf HandleUpdateMana)
    PacketList(ServerPacketID.UpdateHP) = GetAddress(AddressOf HandleUpdateHP)
    PacketList(ServerPacketID.UpdateGold) = GetAddress(AddressOf HandleUpdateGold)
    PacketList(ServerPacketID.UpdateExp) = GetAddress(AddressOf HandleUpdateExp)
    PacketList(ServerPacketID.ChangeMap) = GetAddress(AddressOf HandleChangeMap)
    PacketList(ServerPacketID.PosUpdate) = GetAddress(AddressOf HandlePosUpdate)
    PacketList(ServerPacketID.NPCHitUser) = GetAddress(AddressOf HandleNPCHitUser)
    PacketList(ServerPacketID.UserHitNPC) = GetAddress(AddressOf HandleUserHitNPC)
    PacketList(ServerPacketID.UserAttackedSwing) = GetAddress(AddressOf HandleUserAttackedSwing)
    PacketList(ServerPacketID.UserHittedByUser) = GetAddress(AddressOf HandleUserHittedByUser)
    PacketList(ServerPacketID.UserHittedUser) = GetAddress(AddressOf HandleUserHittedUser)
    PacketList(ServerPacketID.ChatOverHead) = GetAddress(AddressOf HandleChatOverHead)
    PacketList(ServerPacketID.ConsoleMsg) = GetAddress(AddressOf HandleConsoleMessage)
    PacketList(ServerPacketID.GuildChat) = GetAddress(AddressOf HandleGuildChat)
    PacketList(ServerPacketID.ShowMessageBox) = GetAddress(AddressOf HandleShowMessageBox)
    PacketList(ServerPacketID.MostrarCuenta) = GetAddress(AddressOf HandleMostrarCuenta)
    PacketList(ServerPacketID.UserIndexInServer) = GetAddress(AddressOf HandleUserIndexInServer)
    PacketList(ServerPacketID.UserCharIndexInServer) = GetAddress(AddressOf HandleUserCharIndexInServer)
    PacketList(ServerPacketID.CharacterCreate) = GetAddress(AddressOf HandleCharacterCreate)
    PacketList(ServerPacketID.CharacterRemove) = GetAddress(AddressOf HandleCharacterRemove)
    PacketList(ServerPacketID.CharacterMove) = GetAddress(AddressOf HandleCharacterMove)
    PacketList(ServerPacketID.ForceCharMove) = GetAddress(AddressOf HandleForceCharMove)
    PacketList(ServerPacketID.CharacterChange) = GetAddress(AddressOf HandleCharacterChange)
    PacketList(ServerPacketID.ObjectCreate) = GetAddress(AddressOf HandleObjectCreate)
    PacketList(ServerPacketID.fxpiso) = GetAddress(AddressOf HandleFxPiso)
    PacketList(ServerPacketID.ObjectDelete) = GetAddress(AddressOf HandleObjectDelete)
    PacketList(ServerPacketID.BlockPosition) = GetAddress(AddressOf HandleBlockPosition)
    PacketList(ServerPacketID.PlayMIDI) = GetAddress(AddressOf HandlePlayMIDI)
    PacketList(ServerPacketID.PlayWave) = GetAddress(AddressOf HandlePlayWave)
    PacketList(ServerPacketID.guildList) = GetAddress(AddressOf HandleGuildList)
    PacketList(ServerPacketID.AreaChanged) = GetAddress(AddressOf HandleAreaChanged)
    PacketList(ServerPacketID.PauseToggle) = GetAddress(AddressOf HandlePauseToggle)
    PacketList(ServerPacketID.RainToggle) = GetAddress(AddressOf HandleRainToggle)
    PacketList(ServerPacketID.CreateFX) = GetAddress(AddressOf HandleCreateFX)
    PacketList(ServerPacketID.UpdateUserStats) = GetAddress(AddressOf HandleUpdateUserStats)
    PacketList(ServerPacketID.WorkRequestTarget) = GetAddress(AddressOf HandleWorkRequestTarget)
    PacketList(ServerPacketID.ChangeInventorySlot) = GetAddress(AddressOf HandleChangeInventorySlot)
    PacketList(ServerPacketID.InventoryUnlockSlots) = GetAddress(AddressOf HandleInventoryUnlockSlots)
    PacketList(ServerPacketID.ChangeBankSlot) = GetAddress(AddressOf HandleChangeBankSlot)
    PacketList(ServerPacketID.ChangeSpellSlot) = GetAddress(AddressOf HandleChangeSpellSlot)
    PacketList(ServerPacketID.Atributes) = GetAddress(AddressOf HandleAtributes)
    PacketList(ServerPacketID.BlacksmithWeapons) = GetAddress(AddressOf HandleBlacksmithWeapons)
    PacketList(ServerPacketID.BlacksmithArmors) = GetAddress(AddressOf HandleBlacksmithArmors)
    PacketList(ServerPacketID.CarpenterObjects) = GetAddress(AddressOf HandleCarpenterObjects)
    PacketList(ServerPacketID.RestOK) = GetAddress(AddressOf HandleRestOK)
    PacketList(ServerPacketID.ErrorMsg) = GetAddress(AddressOf HandleErrorMessage)
    PacketList(ServerPacketID.Blind) = GetAddress(AddressOf HandleBlind)
    PacketList(ServerPacketID.Dumb) = GetAddress(AddressOf HandleDumb)
    PacketList(ServerPacketID.ShowSignal) = GetAddress(AddressOf HandleShowSignal)
    PacketList(ServerPacketID.ChangeNPCInventorySlot) = GetAddress(AddressOf HandleChangeNPCInventorySlot)
    PacketList(ServerPacketID.UpdateHungerAndThirst) = GetAddress(AddressOf HandleUpdateHungerAndThirst)
    PacketList(ServerPacketID.MiniStats) = GetAddress(AddressOf HandleMiniStats)
    PacketList(ServerPacketID.LevelUp) = GetAddress(AddressOf HandleLevelUp)
    PacketList(ServerPacketID.AddForumMsg) = GetAddress(AddressOf HandleAddForumMessage)
    PacketList(ServerPacketID.ShowForumForm) = GetAddress(AddressOf HandleShowForumForm)
    PacketList(ServerPacketID.SetInvisible) = GetAddress(AddressOf HandleSetInvisible)
    PacketList(ServerPacketID.DiceRoll) = GetAddress(AddressOf HandleDiceRoll)
    PacketList(ServerPacketID.MeditateToggle) = GetAddress(AddressOf HandleMeditateToggle)
    PacketList(ServerPacketID.BlindNoMore) = GetAddress(AddressOf HandleBlindNoMore)
    PacketList(ServerPacketID.DumbNoMore) = GetAddress(AddressOf HandleDumbNoMore)
    PacketList(ServerPacketID.SendSkills) = GetAddress(AddressOf HandleSendSkills)
    PacketList(ServerPacketID.TrainerCreatureList) = GetAddress(AddressOf HandleTrainerCreatureList)
    PacketList(ServerPacketID.guildNews) = GetAddress(AddressOf HandleGuildNews)
    PacketList(ServerPacketID.OfferDetails) = GetAddress(AddressOf HandleOfferDetails)
    PacketList(ServerPacketID.AlianceProposalsList) = GetAddress(AddressOf HandleAlianceProposalsList)
    PacketList(ServerPacketID.PeaceProposalsList) = GetAddress(AddressOf HandlePeaceProposalsList)
    PacketList(ServerPacketID.CharacterInfo) = GetAddress(AddressOf HandleCharacterInfo)
    PacketList(ServerPacketID.GuildLeaderInfo) = GetAddress(AddressOf HandleGuildLeaderInfo)
    PacketList(ServerPacketID.GuildDetails) = GetAddress(AddressOf HandleGuildDetails)
    PacketList(ServerPacketID.ShowGuildFundationForm) = GetAddress(AddressOf HandleShowGuildFundationForm)
    PacketList(ServerPacketID.ParalizeOK) = GetAddress(AddressOf HandleParalizeOK)
    PacketList(ServerPacketID.ShowUserRequest) = GetAddress(AddressOf HandleShowUserRequest)
    PacketList(ServerPacketID.ChangeUserTradeSlot) = GetAddress(AddressOf HandleChangeUserTradeSlot)
    'PacketList(ServerPacketID.SendNight) = GetAddress(AddressOf HandleSendNight)
    PacketList(ServerPacketID.Pong) = GetAddress(AddressOf HandlePong)
    PacketList(ServerPacketID.UpdateTagAndStatus) = GetAddress(AddressOf HandleUpdateTagAndStatus)
    PacketList(ServerPacketID.FYA) = GetAddress(AddressOf HandleFYA)
    PacketList(ServerPacketID.CerrarleCliente) = GetAddress(AddressOf HandleCerrarleCliente)
    PacketList(ServerPacketID.Contadores) = GetAddress(AddressOf HandleContadores)
    
    ' GM Messages
    PacketList(ServerPacketID.SpawnList) = GetAddress(AddressOf HandleSpawnList)
    PacketList(ServerPacketID.ShowSOSForm) = GetAddress(AddressOf HandleShowSOSForm)
    PacketList(ServerPacketID.ShowMOTDEditionForm) = GetAddress(AddressOf HandleShowMOTDEditionForm)
    PacketList(ServerPacketID.ShowGMPanelForm) = GetAddress(AddressOf HandleShowGMPanelForm)
    PacketList(ServerPacketID.UserNameList) = GetAddress(AddressOf HandleUserNameList)
    PacketList(ServerPacketID.PersonajesDeCuenta) = GetAddress(AddressOf HandlePersonajesDeCuenta)
    PacketList(ServerPacketID.UserOnline) = GetAddress(AddressOf HandleUserOnline)
    PacketList(ServerPacketID.ParticleFX) = GetAddress(AddressOf HandleParticleFX)
    PacketList(ServerPacketID.ParticleFXToFloor) = GetAddress(AddressOf HandleParticleFXToFloor)
    PacketList(ServerPacketID.ParticleFXWithDestino) = GetAddress(AddressOf HandleParticleFXWithDestino)
    PacketList(ServerPacketID.ParticleFXWithDestinoXY) = GetAddress(AddressOf HandleParticleFXWithDestinoXY)
    PacketList(ServerPacketID.hora) = GetAddress(AddressOf HandleHora)
    PacketList(ServerPacketID.Light) = GetAddress(AddressOf HandleLight)
    PacketList(ServerPacketID.AuraToChar) = GetAddress(AddressOf HandleAuraToChar)
    PacketList(ServerPacketID.SpeedToChar) = GetAddress(AddressOf HandleSpeedToChar)
    PacketList(ServerPacketID.LightToFloor) = GetAddress(AddressOf HandleLightToFloor)
    PacketList(ServerPacketID.NieveToggle) = GetAddress(AddressOf HandleNieveToggle)
    PacketList(ServerPacketID.NieblaToggle) = GetAddress(AddressOf HandleNieblaToggle)
    PacketList(ServerPacketID.Goliath) = GetAddress(AddressOf HandleGoliathInit)
    PacketList(ServerPacketID.TextOverChar) = GetAddress(AddressOf HandleTextOverChar)
    PacketList(ServerPacketID.TextOverTile) = GetAddress(AddressOf HandleTextOverTile)
    PacketList(ServerPacketID.TextCharDrop) = GetAddress(AddressOf HandleTextCharDrop)
    PacketList(ServerPacketID.FlashScreen) = GetAddress(AddressOf HandleFlashScreen)
    PacketList(ServerPacketID.AlquimistaObj) = GetAddress(AddressOf HandleAlquimiaObjects)
    PacketList(ServerPacketID.ShowAlquimiaForm) = GetAddress(AddressOf HandleShowAlquimiaForm)
    PacketList(ServerPacketID.familiar) = GetAddress(AddressOf HandleFamiliar)
    PacketList(ServerPacketID.SastreObj) = GetAddress(AddressOf HandleSastreObjects)
    PacketList(ServerPacketID.ShowSastreForm) = GetAddress(AddressOf HandleShowSastreForm)
    PacketList(ServerPacketID.VelocidadToggle) = GetAddress(AddressOf HandleVelocidadToggle)
    PacketList(ServerPacketID.MacroTrabajoToggle) = GetAddress(AddressOf HandleMacroTrabajoToggle)
    PacketList(ServerPacketID.RefreshAllInventorySlot) = GetAddress(AddressOf HandleRefreshAllInventorySlot)
    PacketList(ServerPacketID.BindKeys) = GetAddress(AddressOf HandleBindKeys)
    PacketList(ServerPacketID.ShowFrmLogear) = GetAddress(AddressOf HandleShowFrmLogear)
    PacketList(ServerPacketID.ShowFrmMapa) = GetAddress(AddressOf HandleShowFrmMapa)
    PacketList(ServerPacketID.InmovilizadoOK) = GetAddress(AddressOf HandleInmovilizadoOK)
    PacketList(ServerPacketID.BarFx) = GetAddress(AddressOf HandleBarFx)
    PacketList(ServerPacketID.SetEscribiendo) = GetAddress(AddressOf HandleSetEscribiendo)
    PacketList(ServerPacketID.Logros) = GetAddress(AddressOf HandleLogros)
    PacketList(ServerPacketID.TrofeoToggleOn) = GetAddress(AddressOf HandleTrofeoToggleOn)
    PacketList(ServerPacketID.TrofeoToggleOff) = GetAddress(AddressOf HandleTrofeoToggleOff)
    PacketList(ServerPacketID.LocaleMsg) = GetAddress(AddressOf HandleLocaleMsg)
    PacketList(ServerPacketID.ListaCorreo) = GetAddress(AddressOf HandleListaCorreo)
    PacketList(ServerPacketID.ShowPregunta) = GetAddress(AddressOf HandleShowPregunta)
    PacketList(ServerPacketID.DatosGrupo) = GetAddress(AddressOf HandleDatosGrupo)
    PacketList(ServerPacketID.ubicacion) = GetAddress(AddressOf HandleUbicacion)
    PacketList(ServerPacketID.CorreoPicOn) = GetAddress(AddressOf HandleCorreoPicOn)
    PacketList(ServerPacketID.DonadorObj) = GetAddress(AddressOf HandleDonadorObjects)
    PacketList(ServerPacketID.ArmaMov) = GetAddress(AddressOf HandleArmaMov)
    PacketList(ServerPacketID.EscudoMov) = GetAddress(AddressOf HandleEscudoMov)
    PacketList(ServerPacketID.ActShop) = GetAddress(AddressOf HandleActShop)
    PacketList(ServerPacketID.ViajarForm) = GetAddress(AddressOf HandleViajarForm)
    PacketList(ServerPacketID.oxigeno) = GetAddress(AddressOf HandleOxigeno)
    PacketList(ServerPacketID.NadarToggle) = GetAddress(AddressOf HandleNadarToggle)
    PacketList(ServerPacketID.ShowFundarClanForm) = GetAddress(AddressOf HandleShowFundarClanForm)
    PacketList(ServerPacketID.CharUpdateHP) = GetAddress(AddressOf HandleCharUpdateHP)
    PacketList(ServerPacketID.CharUpdateMAN) = GetAddress(AddressOf HandleCharUpdateMAN)
    PacketList(ServerPacketID.Ranking) = GetAddress(AddressOf HandleRanking)
    PacketList(ServerPacketID.PosLLamadaDeClan) = GetAddress(AddressOf HandlePosLLamadaDeClan)
    PacketList(ServerPacketID.QuestDetails) = GetAddress(AddressOf HandleQuestDetails)
    PacketList(ServerPacketID.QuestListSend) = GetAddress(AddressOf HandleQuestListSend)
    PacketList(ServerPacketID.NpcQuestListSend) = GetAddress(AddressOf HandleNpcQuestListSend)
    PacketList(ServerPacketID.UpdateNPCSimbolo) = GetAddress(AddressOf HandleUpdateNPCSimbolo)
    PacketList(ServerPacketID.ClanSeguro) = GetAddress(AddressOf HandleClanSeguro)
    PacketList(ServerPacketID.Intervals) = GetAddress(AddressOf HandleIntervals)
    PacketList(ServerPacketID.UpdateUserKey) = GetAddress(AddressOf HandleUpdateUserKey)
    PacketList(ServerPacketID.UpdateRM) = GetAddress(AddressOf HandleUpdateRM)
    PacketList(ServerPacketID.UpdateDM) = GetAddress(AddressOf HandleUpdateDM)
    PacketList(ServerPacketID.RequestScreenShot) = GetAddress(AddressOf HandleRequestScreenShot)
    PacketList(ServerPacketID.ShowScreenShot) = GetAddress(AddressOf HandleShowScreenShot)
    PacketList(ServerPacketID.ScreenShotData) = GetAddress(AddressOf HandleScreenShotData)
    PacketList(ServerPacketID.Tolerancia0) = GetAddress(AddressOf HandleTolerancia0)
    PacketList(ServerPacketID.SeguroResu) = GetAddress(AddressOf HandleSeguroResu)
    PacketList(ServerPacketID.Stopped) = GetAddress(AddressOf HandleStopped)
    PacketList(ServerPacketID.InvasionInfo) = GetAddress(AddressOf HandleInvasionInfo)
    PacketList(ServerPacketID.CommerceRecieveChatMessage) = GetAddress(AddressOf HandleCommerceRecieveChatMessage)
    PacketList(ServerPacketID.DoAnimation) = GetAddress(AddressOf HandleDoAnimation)
    PacketList(ServerPacketID.OpenCrafting) = GetAddress(AddressOf HandleOpenCrafting)
    PacketList(ServerPacketID.CraftingItem) = GetAddress(AddressOf HandleCraftingItem)
    PacketList(ServerPacketID.CraftingCatalyst) = GetAddress(AddressOf HandleCraftingCatalyst)
    PacketList(ServerPacketID.CraftingResult) = GetAddress(AddressOf HandleCraftingResult)
    PacketList(ServerPacketID.ForceUpdate) = GetAddress(AddressOf HandleForceUpdate)

End Sub

Private Sub ParsePacket(ByVal packetIndex As Long)

    If packetIndex > UBound(PacketList()) Then
        Debug.Print "Paquete inexistente: " & packetIndex
        Exit Sub
    End If

    If PacketList(packetIndex) = 0 Then
        Debug.Print "Paquete inexistente: " & packetIndex
        Exit Sub
    End If

    'llamamos al sub mediante su dirección en memoria
    Call CallHandle(PacketList(packetIndex), 0)
 
End Sub

'Devuelve el argumento que se le pasó (sirve para usar AddressOf en variables)
Private Function GetAddress(ByVal address As Long) As Long
 
    GetAddress = address
 
End Function

''
' Handles incoming data.

Public Function HandleIncomingData() As Boolean
    
    ' WyroX: No remover
    On Error Resume Next

    Dim PacketID As Long
    
    If Not incomingData.CheckLength Then
        HandleIncomingData = False
        Exit Function

    End If
    
    If Not incomingData.ValidCRC Then
        HandleIncomingData = False
        Exit Function

    End If

    PacketID = CLng(incomingData.ReadID())

    InBytes = InBytes + incomingData.Length

    Call ParsePacket(PacketID)
    
    With incomingData
    
        Call .ReadNewPacket
    
        If (Not .BufferOver Or .Length > 0) And .errNumber = 0 Then    'Done with this packet, move on to next one
            Err.Clear
            HandleIncomingData = True
        
        ElseIf .errNumber <> 0 And .errNumber <> .NotEnoughDataErrCode Then
            Call RegistrarError(Err.Number, Err.Description & ". PacketID: " & PacketID, "Protocol.HandleIncomingData", Erl)
            Err.Clear
            HandleIncomingData = False
        
        Else
            Err.Clear
            HandleIncomingData = False
        End If
        
        .errNumber = 0
    
    End With
    
    
End Function

''
' Handles the Logged message.

Private Sub HandleLogged()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleLogged_Err
 
    ' Variable initialization
    UserCiego = False
    EngineRun = True
    UserDescansar = False
    Nombres = True
    Pregunta = False

    frmMain.stabar.Visible = True
    
    frmMain.panelInf.Picture = LoadInterface("ventanaprincipal_stats.bmp")
    frmMain.HpBar.Visible = True

    If UserMaxMAN <> 0 Then
        frmMain.manabar.Visible = True

    End If

    frmMain.hambar.Visible = True
    frmMain.AGUbar.Visible = True
    frmMain.Hpshp.Visible = (UserMinHp > 0)
    frmMain.MANShp.Visible = (UserMinMAN > 0)
    frmMain.STAShp.Visible = (UserMinSTA > 0)
    frmMain.AGUAsp.Visible = (UserMinAGU > 0)
    frmMain.COMIDAsp.Visible = (UserMinHAM > 0)
    frmMain.GldLbl.Visible = True
    ' frmMain.Label6.Visible = True
    frmMain.Fuerzalbl.Visible = True
    frmMain.AgilidadLbl.Visible = True
    frmMain.oxigenolbl.Visible = True
    QueRender = 0
    
    frmMain.ImgSegParty = LoadInterface("boton-seguro-party-on.bmp")
    frmMain.ImgSegClan = LoadInterface("boton-seguro-clan-on.bmp")
    frmMain.ImgSegResu = LoadInterface("boton-fantasma-on.bmp")
    SeguroParty = True
    SeguroClanX = True
    SeguroResuX = True
    
    'Set connected state
    
    Call SetConnected
    
    'Show tip
    
    Exit Sub

HandleLogged_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleLogged", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the RemoveDialogs message.

Private Sub HandleRemoveDialogs()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleRemoveDialogs_Err

    Call Dialogos.RemoveAllDialogs
    
    Exit Sub

HandleRemoveDialogs_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleRemoveDialogs", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the RemoveCharDialog message.

Private Sub HandleRemoveCharDialog()
    
    On Error GoTo HandleRemoveCharDialog_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    Call Dialogos.RemoveDialog(incomingData.ReadInteger())
    
    Exit Sub

HandleRemoveCharDialog_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleRemoveCharDialog", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the NavigateToggle message.

Private Sub HandleNavigateToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleNavigateToggle_Err

    UserNavegando = Not UserNavegando
    
    Exit Sub

HandleNavigateToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleNavigateToggle", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleNadarToggle()
    
    On Error GoTo HandleNadarToggle_Err

    UserNadando = incomingData.ReadBoolean()
    
    Exit Sub

HandleNadarToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleNadarToggle", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleEquiteToggle()
 
    On Error GoTo HandleEquiteToggle_Err
    
    UserMontado = Not UserMontado

    Exit Sub

HandleEquiteToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleEquiteToggle", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleVelocidadToggle()
    
    On Error GoTo HandleVelocidadToggle_Err

    If UserCharIndex = 0 Then Exit Sub
    
    charlist(UserCharIndex).Speeding = incomingData.ReadSingle()
    
    Call MainTimer.SetInterval(TimersIndex.Walk, IntervaloCaminar / charlist(UserCharIndex).Speeding)
    
    Exit Sub

HandleVelocidadToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleVelocidadToggle", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleMacroTrabajoToggle()
    'Activa o Desactiva el macro de trabajo  06/07/2014 Ladder
    
    On Error GoTo HandleMacroTrabajoToggle_Err

    Dim activar As Boolean
    activar = incomingData.ReadBoolean()

    If activar = False Then
    
        Call ResetearUserMacro
        
    Else
    
        Call AddtoRichTextBox(frmMain.RecTxt, "Has comenzado a trabajar...", 2, 223, 51, 1, 0)
        
        frmMain.MacroLadder.Interval = IntervaloTrabajoConstruir
        frmMain.MacroLadder.Enabled = True
        
        UserMacro.Intervalo = IntervaloTrabajoConstruir
        UserMacro.Activado = True
        UserMacro.cantidad = 999
        UserMacro.TIPO = 6
        
        TargetXMacro = tX
        TargetYMacro = tY

    End If
    
    Exit Sub

HandleMacroTrabajoToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleMacroTrabajoToggle", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the Disconnect message.

Private Sub HandleDisconnect()
    
    On Error GoTo HandleDisconnect_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim i As Long

    Call ResetearUserMacro

    'Close connection
    frmMain.MainSocket.Close
    
    'Hide main form
    'FrmCuenta.Visible = True
    
    frmConnect.Visible = True
    QueRender = 2

    Call Graficos_Particulas.Particle_Group_Remove_All
    Call Graficos_Particulas.Engine_Select_Particle_Set(203)
    
    ParticleLluviaDorada = General_Particle_Create(208, -1, -1)

    frmMain.hlst.Visible = False
    frmMain.Timerping.Enabled = False
    
    frmMain.UpdateLight.Enabled = False
    frmMain.UpdateDaytime.Enabled = False
    
    frmMain.Visible = False
    
    OpcionMenu = 0

    frmMain.picInv.Visible = True
    frmMain.hlst.Visible = False

    frmMain.cmdlanzar.Visible = False
    'frmMain.lblrefuerzolanzar.Visible = False
    frmMain.cmdMoverHechi(0).Visible = False
    frmMain.cmdMoverHechi(1).Visible = False
    
    QuePestañaInferior = 0
    frmMain.stabar.Visible = True
    frmMain.HpBar.Visible = True
    frmMain.manabar.Visible = True
    frmMain.hambar.Visible = True
    frmMain.AGUbar.Visible = True
    frmMain.Hpshp.Visible = True
    frmMain.MANShp.Visible = True
    frmMain.STAShp.Visible = True
    frmMain.AGUAsp.Visible = True
    frmMain.COMIDAsp.Visible = True
    frmMain.GldLbl.Visible = True
    ' Label6.Visible = True
    frmMain.Fuerzalbl.Visible = True
    frmMain.AgilidadLbl.Visible = True
    frmMain.oxigenolbl.Visible = True
    frmMain.TiendaBoton.Visible = False
    frmMain.rankingBoton.Visible = False
    frmMain.manualboton.Visible = False
    frmMain.QuestBoton.Visible = False
    frmMain.ImgHogar.Visible = False
    frmMain.lblWeapon.Visible = True
    frmMain.lblShielder.Visible = True
    frmMain.lblHelm.Visible = True
    frmMain.lblArmor.Visible = True
    frmMain.lblResis.Visible = True
    frmMain.lbldm.Visible = True
    frmMain.imgBugReport.Visible = False
    frmMain.panelinferior(0).Picture = Nothing
    frmMain.panelinferior(1).Picture = Nothing
    frmMain.mapMundo.Visible = False
    frmMain.Image5.Visible = False
    frmMain.clanimg.Visible = False
    frmMain.cmdLlavero.Visible = False
    frmMain.QuestBoton.Visible = False
    frmMain.ImgSeg.Visible = False
    frmMain.ImgSegParty.Visible = False
    frmMain.ImgSegClan.Visible = False
    frmMain.ImgSegResu.Visible = False
    
    'Stop audio
    If Sonido Then
        Sound.Sound_Stop_All
        Sound.Ambient_Stop

    End If

    Call CleanDialogs
    
    'frmMain.IsPlaying = PlayLoop.plNone
    
    'Show connection form

    LogeoAlgunaVez = True
    UserMap = 1
    
    EntradaY = 1
    EntradaX = 1

    Call SwitchMap(UserMap)
    
    frmMain.personaje(1).Visible = False
    frmMain.personaje(2).Visible = False
    frmMain.personaje(3).Visible = False
    frmMain.personaje(4).Visible = False
    frmMain.personaje(5).Visible = False
    
    UserClase = 0
    UserSexo = 0
    UserRaza = 0
    MiCabeza = 0
    UserHogar = 0

    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i

    For i = 1 To UserInvUnlocked
        frmMain.imgInvLock(i - 1).Picture = Nothing
    Next i

    For i = 1 To MAX_INVENTORY_SLOTS
        Call frmMain.Inventario.ClearSlot(i)
        Call frmBancoObj.InvBankUsu.ClearSlot(i)
        Call frmComerciar.InvComNpc.ClearSlot(i)
        Call frmComerciar.InvComUsu.ClearSlot(i)
        Call frmBancoCuenta.InvBankUsuCuenta.ClearSlot(i)
        Call frmComerciarUsu.InvUser.ClearSlot(i)
        Call frmCrafteo.InvCraftUser.ClearSlot(i)
    Next i

    For i = 1 To MAX_BANCOINVENTORY_SLOTS
        Call frmBancoObj.InvBoveda.ClearSlot(i)
    Next i

    For i = 1 To MAX_KEYS
        Call FrmKeyInv.InvKeys.ClearSlot(i)
    Next i

    For i = 1 To MAX_SLOTS_CRAFTEO
        Call frmCrafteo.InvCraftItems.ClearSlot(i)
    Next i

    Call frmCrafteo.InvCraftCatalyst.ClearSlot(1)
    
    UserInvUnlocked = 0

    Alocados = 0

    'Reset global vars
    UserParalizado = False
    UserSaliendo = False
    UserStopped = False
    UserInmovilizado = False
    pausa = False
    UserMeditar = False
    UserDescansar = False
    UserNavegando = False
    UserMontado = False
    UserNadando = False
    bRain = False
    AlphaNiebla = 30
    frmMain.TimerNiebla.Enabled = False
    bNiebla = False
    MostrarTrofeo = False
    bNieve = False
    bFogata = False
    SkillPoints = 0
    UserEstado = 0
    
    InviCounter = 0
    ScrollExpCounter = 0
    ScrollOroCounter = 0
    DrogaCounter = 0
    OxigenoCounter = 0
     
    frmMain.Contadores.Enabled = False
    
    InvasionActual = 0
    frmMain.Evento.Enabled = False
     
    'Delete all kind of dialogs
    
    'Reset some char variables...
    For i = 1 To LastChar + 1
        charlist(i).Invisible = False
        charlist(i).Arma_Aura = ""
        charlist(i).Body_Aura = ""
        charlist(i).Escudo_Aura = ""
        charlist(i).DM_Aura = ""
        charlist(i).RM_Aura = ""
        charlist(i).Otra_Aura = ""
        charlist(i).Head_Aura = ""
        charlist(i).Speeding = 0
        charlist(i).AuraAngle = 0
    Next i

    For i = 1 To LastChar + 1
        charlist(i).dialog = ""
    Next i
        
    'Unload all forms except frmMain and frmConnect
    Dim Frm As Form
    
    For Each Frm In Forms

        If Frm.Name <> frmMain.Name And Frm.Name <> frmConnect.Name And Frm.Name <> frmMensaje.Name Then
            Unload Frm

        End If

    Next
    
    Exit Sub

HandleDisconnect_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleDisconnect", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the CommerceEnd message.

Private Sub HandleCommerceEnd()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo HandleCommerceEnd_Err

    'Reset vars
    Comerciando = False
    
    'Hide form
    ' Unload frmComerciar
    
    Exit Sub

HandleCommerceEnd_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCommerceEnd", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the BankEnd message.

Private Sub HandleBankEnd()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo HandleBankEnd_Err

    'Unload frmBancoObj
    Comerciando = False
    
    Exit Sub

HandleBankEnd_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBankEnd", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the CommerceInit message.

Private Sub HandleCommerceInit()
    
    On Error GoTo HandleCommerceInit_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim i       As Long

    Dim NpcName As String

    NpcName = incomingData.ReadASCIIString()

    'Fill our inventory list
    For i = 1 To MAX_INVENTORY_SLOTS

        With frmMain.Inventario
            Call frmComerciar.InvComUsu.SetItem(i, .OBJIndex(i), .Amount(i), .Equipped(i), .GrhIndex(i), .ObjType(i), .MaxHit(i), .MinHit(i), .Def(i), .Valor(i), .ItemName(i), .PuedeUsar(i))

        End With

    Next i

    'Set state and show form
    Comerciando = True
    'Call Inventario.Initialize(frmComerciar.PicInvUser)
    frmComerciar.Show , frmMain
    
    Exit Sub

HandleCommerceInit_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCommerceInit", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the BankInit message.

Private Sub HandleBankInit()
    
    On Error GoTo HandleBankInit_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim i As Long

    'Fill our inventory list
    For i = 1 To MAX_INVENTORY_SLOTS

        With frmMain.Inventario
            Call frmBancoObj.InvBankUsu.SetItem(i, .OBJIndex(i), .Amount(i), .Equipped(i), .GrhIndex(i), .ObjType(i), .MaxHit(i), .MinHit(i), .Def(i), .Valor(i), .ItemName(i), .PuedeUsar(i))

        End With

    Next i

    'Set state and show form
    Comerciando = True
    'Call Inventario.Initialize(frmBancoObj.PicInvUser)
    frmBancoObj.Show , frmMain
    frmBancoObj.lblcosto = PonerPuntos(UserGLD)
    
    Exit Sub

HandleBankInit_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBankInit", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleGoliathInit()
    
    On Error GoTo HandleGoliathInit_Err

    '***************************************************
    '
    '***************************************************

    Dim UserBoveOro As Long

    Dim UserInvBove As Byte
    
    UserBoveOro = incomingData.ReadLong()
    UserInvBove = incomingData.ReadByte()
    Call frmGoliath.ParseBancoInfo(UserBoveOro, UserInvBove)
    
    Exit Sub

HandleGoliathInit_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGoliathInit", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleShowFrmLogear()
    
    On Error GoTo HandleShowFrmLogear_Err

    '***************************************************
    '
    '***************************************************
    FrmLogear.Show , frmConnect
    FrmLogear.Top = FrmLogear.Top + 4000
    
    Exit Sub

HandleShowFrmLogear_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowFrmLogear", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleShowFrmMapa()
    
    On Error GoTo HandleShowFrmMapa_Err

    '***************************************************
    '
    '***************************************************
    ExpMult = incomingData.ReadInteger()
    OroMult = incomingData.ReadInteger()
    
    Call frmMapaGrande.CalcularPosicionMAPA

    frmMapaGrande.Picture = LoadInterface("ventanamapa.bmp")
    frmMapaGrande.Show , frmMain
    
    Exit Sub

HandleShowFrmMapa_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowFrmMapa", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the UserCommerceInit message.

Private Sub HandleUserCommerceInit()
    
    On Error GoTo HandleUserCommerceInit_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim i As Long
    
    'Clears lists if necessary
    
    'Fill inventory list
    With frmMain.Inventario

        For i = 1 To MAX_INVENTORY_SLOTS
            frmComerciarUsu.InvUser.SetItem i, .OBJIndex(i), .Amount(i), .Equipped(i), .GrhIndex(i), .ObjType(i), 0, 0, 0, 0, .ItemName(i), 0
        Next i

    End With
        
    frmComerciarUsu.lblMyGold.Caption = frmMain.GldLbl.Caption
    
    Dim j As Byte

    For j = 1 To 6
        Call frmComerciarUsu.InvOtherSell.SetItem(j, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0)
        Call frmComerciarUsu.InvUserSell.SetItem(j, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0)
    Next j
    
    'Set state and show form
    Comerciando = True
    
    frmComerciarUsu.Show , frmMain
    
    Exit Sub

HandleUserCommerceInit_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUserCommerceInit", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the UserCommerceEnd message.

Private Sub HandleUserCommerceEnd()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo HandleUserCommerceEnd_Err
    
    'Destroy the form and reset the state
    Unload frmComerciarUsu
    Comerciando = False
    
    Exit Sub

HandleUserCommerceEnd_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUserCommerceEnd", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the ShowBlacksmithForm message.

Private Sub HandleShowBlacksmithForm()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo HandleShowBlacksmithForm_Err
    
    If frmMain.macrotrabajo.Enabled And (MacroBltIndex > 0) Then
    
        Call WriteCraftBlacksmith(MacroBltIndex)
        
    Else
    
        frmHerrero.lstArmas.Clear

        Dim i As Byte

        For i = 0 To UBound(CascosHerrero())

            If CascosHerrero(i).Index = 0 Then Exit For
            Call frmHerrero.lstArmas.AddItem(ObjData(CascosHerrero(i).Index).Name)
        Next i

        frmHerrero.Command3.Picture = LoadInterface("boton-casco-over.bmp")
    
        COLOR_AZUL = RGB(0, 0, 0)
        Call Establecer_Borde(frmHerrero.lstArmas, frmHerrero, COLOR_AZUL, 0, 0)
        Call Establecer_Borde(frmHerrero.List1, frmHerrero, COLOR_AZUL, 0, 0)
        Call Establecer_Borde(frmHerrero.List2, frmHerrero, COLOR_AZUL, 0, 0)
        frmHerrero.Show , frmMain

    End If
    
    Exit Sub

HandleShowBlacksmithForm_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowBlacksmithForm", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the ShowCarpenterForm message.

Private Sub HandleShowCarpenterForm()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    On Error GoTo HandleShowCarpenterForm_Err
        
    If frmMain.macrotrabajo.Enabled And (MacroBltIndex > 0) Then
    
        Call WriteCraftCarpenter(MacroBltIndex)
        
    Else
         
        COLOR_AZUL = RGB(0, 0, 0)
    
        ' establece el borde al listbox
        Call Establecer_Borde(frmCarp.lstArmas, frmCarp, COLOR_AZUL, 0, 0)
        Call Establecer_Borde(frmCarp.List1, frmCarp, COLOR_AZUL, 0, 0)
        Call Establecer_Borde(frmCarp.List2, frmCarp, COLOR_AZUL, 0, 0)
        frmCarp.Show , frmMain

    End If
    
    Exit Sub

HandleShowCarpenterForm_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowCarpenterForm", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleShowAlquimiaForm()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo HandleShowAlquimiaForm_Err
    
    If frmMain.macrotrabajo.Enabled And (MacroBltIndex > 0) Then
    
        Call WriteCraftAlquimista(MacroBltIndex)
        
    Else
    
        frmAlqui.Picture = LoadInterface("alquimia.bmp")
    
        COLOR_AZUL = RGB(0, 0, 0)
        
        ' establece el borde al listbox
        Call Establecer_Borde(frmAlqui.lstArmas, frmAlqui, COLOR_AZUL, 1, 1)
        Call Establecer_Borde(frmAlqui.List1, frmAlqui, COLOR_AZUL, 1, 1)
        Call Establecer_Borde(frmAlqui.List2, frmAlqui, COLOR_AZUL, 1, 1)

        frmAlqui.Show , frmMain

    End If
    
    Exit Sub

HandleShowAlquimiaForm_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowAlquimiaForm", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleShowSastreForm()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo HandleShowSastreForm_Err
        
    If frmMain.macrotrabajo.Enabled And (MacroBltIndex > 0) Then
    
        Call WriteCraftSastre(MacroBltIndex)
        
    Else
    
        COLOR_AZUL = RGB(0, 0, 0)

        ' establece el borde al listbox
        Call Establecer_Borde(FrmSastre.lstArmas, FrmSastre, COLOR_AZUL, 1, 1)
        Call Establecer_Borde(FrmSastre.List1, FrmSastre, COLOR_AZUL, 1, 1)
        Call Establecer_Borde(FrmSastre.List2, FrmSastre, COLOR_AZUL, 1, 1)
        FrmSastre.Picture = LoadInterface("sastreria.bmp")

        Dim i As Byte

        FrmSastre.lstArmas.Clear

        For i = 1 To UBound(SastreRopas())

            If SastreRopas(i).Index = 0 Then Exit For
            FrmSastre.lstArmas.AddItem (ObjData(SastreRopas(i).Index).Name)
        Next i
    
        FrmSastre.Command1.Picture = LoadInterface("sastreria_vestimentahover.bmp")
        FrmSastre.Show , frmMain

    End If
    
    Exit Sub

HandleShowSastreForm_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowSastreForm", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the NPCKillUser message.

Private Sub HandleNPCKillUser()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo HandleNPCKillUser_Err
        
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_MATADO, 255, 0, 0, True, False, False)
    
    Exit Sub

HandleNPCKillUser_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleNPCKillUser", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the BlockedWithShieldUser message.

Private Sub HandleBlockedWithShieldUser()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    On Error GoTo HandleBlockedWithShieldUser_Err
        
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)
    
    Exit Sub

HandleBlockedWithShieldUser_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBlockedWithShieldUser", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the BlockedWithShieldOther message.

Private Sub HandleBlockedWithShieldOther()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    On Error GoTo HandleBlockedWithShieldOther_Err
        
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)
    
    Exit Sub

HandleBlockedWithShieldOther_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBlockedWithShieldOther", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the UserSwing message.

Private Sub HandleCharSwing()
    
    On Error GoTo HandleCharSwing_Err
    
    Dim charindex As Integer

    charindex = incomingData.ReadInteger
    
    Dim ShowFX As Boolean

    ShowFX = incomingData.ReadBoolean
    
    Dim ShowText As Boolean

    ShowText = incomingData.ReadBoolean
        
    With charlist(charindex)

        If ShowText Then
            Call SetCharacterDialogFx(charindex, IIf(charindex = UserCharIndex, "Fallas", "Falló"), RGBA_From_Comp(255, 0, 0))

        End If
        
        Call Sound.Sound_Play(2, False, Sound.Calculate_Volume(.pos.x, .pos.y), Sound.Calculate_Pan(.pos.x, .pos.y)) ' Swing
        
        If ShowFX Then Call SetCharacterFx(charindex, 90, 0)

    End With
    
    Exit Sub

HandleCharSwing_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCharSwing", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the SafeModeOn message.

Private Sub HandleSafeModeOn()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    On Error GoTo HandleSafeModeOn_Err
        
    Call frmMain.DibujarSeguro
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_ACTIVADO, 65, 190, 156, False, False, False)
    
    Exit Sub

HandleSafeModeOn_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSafeModeOn", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the SafeModeOff message.

Private Sub HandleSafeModeOff()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo HandleSafeModeOff_Err
    
    Call frmMain.DesDibujarSeguro
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_DESACTIVADO, 65, 190, 156, False, False, False)
    
    Exit Sub

HandleSafeModeOff_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSafeModeOff", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the ResuscitationSafeOff message.

Private Sub HandlePartySafeOff()
    '***************************************************
    'Author: Rapsodius
    'Creation date: 10/10/07
    '***************************************************
    
    On Error GoTo HandlePartySafeOff_Err
    
    Call frmMain.ControlSeguroParty(False)
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_PARTY_OFF, 250, 250, 0, False, True, False)
    
    Exit Sub

HandlePartySafeOff_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePartySafeOff", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleClanSeguro()
    
    On Error GoTo HandleClanSeguro_Err

    '***************************************************
    'Author: Rapsodius
    'Creation date: 10/10/07
    '***************************************************
    Dim Seguro As Boolean
    
    'Get data and update form
    Seguro = incomingData.ReadBoolean()
    
    If SeguroClanX Then
    
        Call AddtoRichTextBox(frmMain.RecTxt, "Seguro de clan desactivado.", 65, 190, 156, False, False, False)
        frmMain.ImgSegClan = LoadInterface("boton-seguro-clan-off.bmp")
        SeguroClanX = False
        
    Else
        Call AddtoRichTextBox(frmMain.RecTxt, "Seguro de clan activado.", 65, 190, 156, False, False, False)
        frmMain.ImgSegClan = LoadInterface("boton-seguro-clan-on.bmp")
        SeguroClanX = True

    End If
    
    Exit Sub

HandleClanSeguro_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleClanSeguro", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleIntervals()
    
    On Error GoTo HandleIntervals_Err

    IntervaloArco = incomingData.ReadLong()
    IntervaloCaminar = incomingData.ReadLong()
    IntervaloGolpe = incomingData.ReadLong()
    IntervaloGolpeMagia = incomingData.ReadLong()
    IntervaloMagia = incomingData.ReadLong()
    IntervaloMagiaGolpe = incomingData.ReadLong()
    IntervaloGolpeUsar = incomingData.ReadLong()
    IntervaloTrabajoExtraer = incomingData.ReadLong()
    IntervaloTrabajoConstruir = incomingData.ReadLong()
    IntervaloUsarU = incomingData.ReadLong()
    IntervaloUsarClic = incomingData.ReadLong()
    IntervaloTirar = incomingData.ReadLong()
    
    'Set the intervals of timers
    Call MainTimer.SetInterval(TimersIndex.Attack, IntervaloGolpe)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithU, IntervaloUsarU)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithDblClick, IntervaloUsarClic)
    Call MainTimer.SetInterval(TimersIndex.SendRPU, INT_SENTRPU)
    Call MainTimer.SetInterval(TimersIndex.CastSpell, IntervaloMagia)
    Call MainTimer.SetInterval(TimersIndex.Arrows, IntervaloArco)
    Call MainTimer.SetInterval(TimersIndex.CastAttack, IntervaloMagiaGolpe)
    Call MainTimer.SetInterval(TimersIndex.AttackSpell, IntervaloGolpeMagia)
    Call MainTimer.SetInterval(TimersIndex.AttackUse, IntervaloGolpeUsar)
    Call MainTimer.SetInterval(TimersIndex.Drop, IntervaloTirar)
    Call MainTimer.SetInterval(TimersIndex.Walk, IntervaloCaminar)

    'Init timers
    Call MainTimer.Start(TimersIndex.Attack)
    Call MainTimer.Start(TimersIndex.UseItemWithU)
    Call MainTimer.Start(TimersIndex.UseItemWithDblClick)
    Call MainTimer.Start(TimersIndex.SendRPU)
    Call MainTimer.Start(TimersIndex.CastSpell)
    Call MainTimer.Start(TimersIndex.Arrows)
    Call MainTimer.Start(TimersIndex.CastAttack)
    Call MainTimer.Start(TimersIndex.AttackSpell)
    Call MainTimer.Start(TimersIndex.AttackUse)
    Call MainTimer.Start(TimersIndex.Drop)
    Call MainTimer.Start(TimersIndex.Walk)
    
    Exit Sub

HandleIntervals_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleIntervals", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleUpdateUserKey()
    
    On Error GoTo HandleUpdateUserKey_Err
 
    Dim Slot As Integer, Llave As Integer
    
    Slot = incomingData.ReadInteger
    Llave = incomingData.ReadInteger

    Call FrmKeyInv.InvKeys.SetItem(Slot, Llave, 1, 0, ObjData(Llave).GrhIndex, eObjType.otLlaves, 0, 0, 0, 0, ObjData(Llave).Name, 0)
    
    Exit Sub

HandleUpdateUserKey_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateUserKey", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleUpdateDM()
    
    On Error GoTo HandleUpdateDM_Err
 
    Dim Value As Integer

    Value = incomingData.ReadInteger

    frmMain.lbldm = "+" & Value & "%"
    
    Exit Sub

HandleUpdateDM_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateDM", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleUpdateRM()
    
    On Error GoTo HandleUpdateRM_Err
 
    Dim Value As Integer

    Value = incomingData.ReadInteger

    frmMain.lblResis = "+" & Value
    
    Exit Sub

HandleUpdateRM_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateRM", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

' Handles the ResuscitationSafeOn message.
Private Sub HandlePartySafeOn()

    '***************************************************
    'Author: Rapsodius
    'Creation date: 10/10/07
    '***************************************************
    On Error GoTo HandlePartySafeOn_Err

    Call frmMain.ControlSeguroParty(True)
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_PARTY_ON, 250, 250, 0, False, True, False)
    
    Exit Sub

HandlePartySafeOn_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePartySafeOn", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleCorreoPicOn()

    '***************************************************
    'Author: Rapsodius
    'Creation date: 10/10/07
    '***************************************************
    On Error GoTo HandleCorreoPicOn_Err

    frmMain.PicCorreo.Visible = True

    Exit Sub

HandleCorreoPicOn_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCorreoPicOn", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the CantUseWhileMeditating message.

Private Sub HandleCantUseWhileMeditating()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo HandleCantUseWhileMeditating_Err

    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USAR_MEDITANDO, 255, 0, 0, False, False, False)
    
    Exit Sub

HandleCantUseWhileMeditating_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCantUseWhileMeditating", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the UpdateSta message.

Private Sub HandleUpdateSta()
    
    On Error GoTo HandleUpdateSta_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    'Get data and update form
    UserMinSTA = incomingData.ReadInteger()
    frmMain.STAShp.Width = UserMinSTA / UserMaxSTA * 89
    frmMain.stabar.Caption = UserMinSTA & " / " & UserMaxSTA

    If QuePestañaInferior = 0 Then
        frmMain.STAShp.Visible = (UserMinSTA > 0)

    End If
    
    Exit Sub

HandleUpdateSta_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateSta", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the UpdateMana message.

Private Sub HandleUpdateMana()
    
    On Error GoTo HandleUpdateMana_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    Dim OldMana As Integer
    OldMana = UserMinMAN
    
    'Get data and update form
    UserMinMAN = incomingData.ReadInteger()
    
    If UserMeditar And UserMinMAN - OldMana > 0 Then

        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("Has ganado " & UserMinMAN - OldMana & " de maná.", .red, .green, .blue, .bold, .italic)

        End With

    End If
    
    If UserMaxMAN > 0 Then
        frmMain.MANShp.Width = UserMinMAN / UserMaxMAN * 216
        frmMain.manabar.Caption = UserMinMAN & " / " & UserMaxMAN

        If QuePestañaInferior = 0 Then
            frmMain.MANShp.Visible = (UserMinMAN > 0)
            frmMain.manabar.Visible = True

        End If

    Else
        frmMain.MANShp.Width = 0
        frmMain.manabar.Visible = False
        frmMain.MANShp.Visible = False

    End If
    
    Exit Sub

HandleUpdateMana_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateMana", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the UpdateHP message.

Private Sub HandleUpdateHP()
    
    On Error GoTo HandleUpdateHP_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    Dim NuevoValor As Long
    NuevoValor = incomingData.ReadInteger()
    
    ' Si perdió vida, mostramos los stats en el frmMain
    If NuevoValor < UserMinHp Then
        Call frmMain.ShowStats

    End If
    
    'Get data and update form
    UserMinHp = NuevoValor
    frmMain.Hpshp.Width = UserMinHp / UserMaxHp * 216
    frmMain.HpBar.Caption = UserMinHp & " / " & UserMaxHp
    
    If QuePestañaInferior = 0 Then
        frmMain.Hpshp.Visible = (UserMinHp > 0)

    End If
    
    'Velocidad de la musica
    
    'Is the user alive??
    If UserMinHp = 0 Then
        UserEstado = 1
    Else
        UserEstado = 0

    End If
    
    Exit Sub

HandleUpdateHP_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateHP", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the UpdateGold message.

Private Sub HandleUpdateGold()
    
    On Error GoTo HandleUpdateGold_Err

    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 08/14/07
    'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
    '- 08/14/07: Added GldLbl color variation depending on User Gold and Level
    '***************************************************

    'Get data and update form
    UserGLD = incomingData.ReadLong()
    
    frmMain.GldLbl.Caption = PonerPuntos(UserGLD)
    
    Exit Sub

HandleUpdateGold_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateGold", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the UpdateExp message.

Private Sub HandleUpdateExp()
    
    On Error GoTo HandleUpdateExp_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    'Get data and update form
    UserExp = incomingData.ReadLong()

    If UserPasarNivel > 0 Then
        frmMain.EXPBAR.Width = UserExp / UserPasarNivel * 235
        frmMain.lblPorcLvl.Caption = Round(UserExp * (100 / UserPasarNivel), 2) & "%"
        frmMain.exp.Caption = PonerPuntos(UserExp) & "/" & PonerPuntos(UserPasarNivel)
        
    Else
        frmMain.EXPBAR.Width = 235
        frmMain.lblPorcLvl.Caption = "¡Nivel máximo!"
        frmMain.exp.Caption = "¡Nivel máximo!"

    End If
    
    Exit Sub

HandleUpdateExp_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateExp", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the ChangeMap message.

Private Sub HandleChangeMap()
    
    On Error GoTo HandleChangeMap_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    UserMap = incomingData.ReadInteger()
    
    'TODO: Once on-the-fly editor is implemented check for map version before loading....
    'For now we just drop it
    Call incomingData.ReadInteger
    
    If bRain Then
    
        If Not MapDat.LLUVIA Then
            frmMain.IsPlaying = PlayLoop.plNone

        End If

    End If

    If frmComerciar.Visible Then Unload frmComerciar
    If frmBancoObj.Visible Then Unload frmBancoObj
    If FrmShop.Visible Then Unload FrmShop
    If frmEstadisticas.Visible Then Unload frmEstadisticas
    If frmHerrero.Visible Then Unload frmHerrero
    If FrmSastre.Visible Then Unload FrmSastre
    If frmAlqui.Visible Then Unload frmAlqui
    If frmCarp.Visible Then Unload frmCarp
    If FrmGrupo.Visible Then Unload FrmGrupo
    If FrmCorreo.Visible Then Unload FrmCorreo
    If frmGoliath.Visible Then Unload frmGoliath
    If FrmViajes.Visible Then Unload FrmViajes
    If frmCantidad.Visible Then Unload frmCantidad
    If FrmRanking.Visible Then Unload FrmRanking

    Call SwitchMap(UserMap)
    
    Exit Sub

HandleChangeMap_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChangeMap", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the PosUpdate message.

Private Sub HandlePosUpdate()
    
    On Error GoTo HandlePosUpdate_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    'Remove char from old position
    If MapData(UserPos.x, UserPos.y).charindex = UserCharIndex Then
        MapData(UserPos.x, UserPos.y).charindex = 0

    End If
    
    'Set new pos
    UserPos.x = incomingData.ReadByte()
    UserPos.y = incomingData.ReadByte()

    'Set char
    MapData(UserPos.x, UserPos.y).charindex = UserCharIndex
    charlist(UserCharIndex).pos = UserPos
        
    'Are we under a roof?
    bTecho = HayTecho(UserPos.x, UserPos.y)
                
    'Update pos label and minimap
    frmMain.Coord.Caption = UserMap & "-" & UserPos.x & "-" & UserPos.y
    Call frmMain.SetMinimapPosition(0, UserPos.x, UserPos.y)

    Call RefreshAllChars
    
    Exit Sub

HandlePosUpdate_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePosUpdate", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the NPCHitUser message.

Private Sub HandleNPCHitUser()
    
    On Error GoTo HandleNPCHitUser_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim Lugar As Byte, DañoStr As String
    
    Lugar = incomingData.ReadByte()

    DañoStr = PonerPuntos(incomingData.ReadInteger)

    Select Case Lugar

        Case bCabeza
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CABEZA & DañoStr, 255, 0, 0, True, False, False)

        Case bBrazoIzquierdo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_IZQ & DañoStr, 255, 0, 0, True, False, False)

        Case bBrazoDerecho
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_DER & DañoStr, 255, 0, 0, True, False, False)

        Case bPiernaIzquierda
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_IZQ & DañoStr, 255, 0, 0, True, False, False)

        Case bPiernaDerecha
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_DER & DañoStr, 255, 0, 0, True, False, False)

        Case bTorso
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_TORSO & DañoStr, 255, 0, 0, True, False, False)

    End Select
    
    Exit Sub

HandleNPCHitUser_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleNPCHitUser", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the UserHitNPC message.

Private Sub HandleUserHitNPC()
    
    On Error GoTo HandleUserHitNPC_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CRIATURA_1 & PonerPuntos(incomingData.ReadLong()) & MENSAJE_2, 255, 0, 0, True, False, False)
    
    Exit Sub

HandleUserHitNPC_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUserHitNPC", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the UserAttackedSwing message.

Private Sub HandleUserAttackedSwing()
    
    On Error GoTo HandleUserAttackedSwing_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & charlist(incomingData.ReadInteger()).nombre & MENSAJE_ATAQUE_FALLO, 255, 0, 0, True, False, False)
    
    Exit Sub

HandleUserAttackedSwing_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUserAttackedSwing", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the UserHittingByUser message.

Private Sub HandleUserHittedByUser()
    
    On Error GoTo HandleUserHittedByUser_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim attacker As String
    Dim intt     As Integer
    
    intt = incomingData.ReadInteger()
    
    Dim pos As String

    pos = InStr(charlist(intt).nombre, "<")
    
    If pos = 0 Then pos = Len(charlist(intt).nombre) + 2
    
    attacker = Left$(charlist(intt).nombre, pos - 2)
    
    Dim Lugar As Byte
    Lugar = incomingData.ReadByte
    
    Dim DañoStr As String
    DañoStr = PonerPuntos(incomingData.ReadInteger())
    
    Select Case Lugar

        Case bCabeza
            Call AddtoRichTextBox(frmMain.RecTxt, attacker & MENSAJE_RECIVE_IMPACTO_CABEZA & DañoStr & MENSAJE_2, 255, 0, 0, True, False, False)

        Case bBrazoIzquierdo
            Call AddtoRichTextBox(frmMain.RecTxt, attacker & MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ & DañoStr & MENSAJE_2, 255, 0, 0, True, False, False)

        Case bBrazoDerecho
            Call AddtoRichTextBox(frmMain.RecTxt, attacker & MENSAJE_RECIVE_IMPACTO_BRAZO_DER & DañoStr & MENSAJE_2, 255, 0, 0, True, False, False)

        Case bPiernaIzquierda
            Call AddtoRichTextBox(frmMain.RecTxt, attacker & MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ & DañoStr & MENSAJE_2, 255, 0, 0, True, False, False)

        Case bPiernaDerecha
            Call AddtoRichTextBox(frmMain.RecTxt, attacker & MENSAJE_RECIVE_IMPACTO_PIERNA_DER & DañoStr & MENSAJE_2, 255, 0, 0, True, False, False)

        Case bTorso
            Call AddtoRichTextBox(frmMain.RecTxt, attacker & MENSAJE_RECIVE_IMPACTO_TORSO & DañoStr & MENSAJE_2, 255, 0, 0, True, False, False)

    End Select
    
    Exit Sub

HandleUserHittedByUser_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUserHittedByUser", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the UserHittedUser message.

Private Sub HandleUserHittedUser()
    
    On Error GoTo HandleUserHittedUser_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim victim As String
    
    Dim intt   As Integer
    
    intt = incomingData.ReadInteger()
    'attacker = charlist().Nombre
    
    Dim pos As String

    pos = InStr(charlist(intt).nombre, "<")
    
    If pos = 0 Then pos = Len(charlist(intt).nombre) + 2
    
    victim = Left$(charlist(intt).nombre, pos - 2)
    
    Dim Lugar As Byte
    Lugar = incomingData.ReadByte()
    
    Dim DañoStr As String
    DañoStr = PonerPuntos(incomingData.ReadInteger())
    
    Select Case Lugar

        Case bCabeza
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_CABEZA & DañoStr & MENSAJE_2, 255, 0, 0, True, False, False)

        Case bBrazoIzquierdo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ & DañoStr & MENSAJE_2, 255, 0, 0, True, False, False)

        Case bBrazoDerecho
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_BRAZO_DER & DañoStr & MENSAJE_2, 255, 0, 0, True, False, False)

        Case bPiernaIzquierda
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ & DañoStr & MENSAJE_2, 255, 0, 0, True, False, False)

        Case bPiernaDerecha
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_PIERNA_DER & DañoStr & MENSAJE_2, 255, 0, 0, True, False, False)

        Case bTorso
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_TORSO & DañoStr & MENSAJE_2, 255, 0, 0, True, False, False)

    End Select
    
    Exit Sub

HandleUserHittedUser_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUserHittedUser", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the ChatOverHead message.

Private Sub HandleChatOverHead()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    On Error GoTo ErrHandler

    Dim chat       As String

    Dim charindex  As Integer

    Dim r          As Byte

    Dim G          As Byte

    Dim B          As Byte

    Dim colortexto As Long

    Dim QueEs      As String

    chat = incomingData.ReadASCIIString()
    charindex = incomingData.ReadInteger()
    
    r = incomingData.ReadByte()
    G = incomingData.ReadByte()
    B = incomingData.ReadByte()
    
    colortexto = vbColor_2_Long(incomingData.ReadLong())

    'Optimizacion de protocolo por Ladder
    QueEs = ReadField(1, chat, Asc("*"))
    
    Dim copiar As Boolean

    copiar = False
    
    Dim duracion As Integer

    duracion = 250
    
    Select Case QueEs

        Case "NPCDESC"
            chat = NpcData(ReadField(2, chat, Asc("*"))).desc
            copiar = True

        Case "PMAG"
            chat = HechizoData(ReadField(2, chat, Asc("*"))).PalabrasMagicas
            copiar = True
            duracion = 20
            
        Case "QUESTFIN"
            chat = QuestList(ReadField(2, chat, Asc("*"))).DescFinal
            copiar = True
            duracion = 20
            
        Case "QUESTNEXT"
            chat = QuestList(ReadField(2, chat, Asc("*"))).NextQuest
            copiar = True
            duracion = 20
            
            If LenB(chat) = 0 Then
                chat = "Ya has completado esa misión para mí."

            End If
        
    End Select
            
    'Only add the chat if the character exists (a CharacterRemove may have been sent to the PC / NPC area before the buffer was flushed)
    If charlist(charindex).active Then

        Call Char_Dialog_Set(charindex, chat, colortexto, duracion, 30)

    End If
    
    If charlist(charindex).EsNpc = False Then
         
        If CopiarDialogoAConsola = 1 And Not copiar Then
    
            Call WriteChatOverHeadInConsole(charindex, chat, r, G, B)

        End If

    End If

    Exit Sub
    
ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChatOverHead", Erl)
    Call incomingData.SafeClearPacket

End Sub

Private Sub HandleTextOverChar()

    On Error GoTo ErrHandler
    
    Dim chat      As String

    Dim charindex As Integer

    Dim Color     As Long
    
    chat = incomingData.ReadASCIIString()
    charindex = incomingData.ReadInteger()
    
    Color = incomingData.ReadLong()
    
    Call SetCharacterDialogFx(charindex, chat, RGBA_From_vbColor(Color))

    Exit Sub
    
ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleTextOverChar", Erl)
    Call incomingData.SafeClearPacket

End Sub

Private Sub HandleTextOverTile()

    On Error GoTo ErrHandler
    
    Dim Text  As String

    Dim x     As Integer, y As Integer

    Dim Color As Long
    
    Text = incomingData.ReadASCIIString()
    x = incomingData.ReadInteger()
    y = incomingData.ReadInteger()
    Color = incomingData.ReadLong()
    
    If InMapBounds(x, y) Then
    
        With MapData(x, y)
            Dim Index As Integer
            
            If UBound(.DialogEffects) = 0 Then
                ReDim .DialogEffects(1 To 1)
                
                Index = 1
            Else

                For Index = 1 To UBound(.DialogEffects)

                    If .DialogEffects(Index).Text = vbNullString Then
                        Exit For

                    End If

                Next
                
                If Index > UBound(.DialogEffects) Then
                    ReDim .DialogEffects(1 To UBound(.DialogEffects) + 1)

                End If

            End If
            
            With .DialogEffects(Index)
            
                .Color = RGBA_From_vbColor(Color)
                .Start = FrameTime
                .Text = Text
                .offset.x = 0
                .offset.y = 0
            
            End With

        End With
        
    End If

    Exit Sub
    
ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleTextOverTile", Erl)
    Call incomingData.SafeClearPacket

End Sub

Private Sub HandleTextCharDrop()

    On Error GoTo ErrHandler
    
    Dim Text      As String

    Dim charindex As Integer

    Dim Color     As Long
    
    Text = incomingData.ReadASCIIString()
    charindex = incomingData.ReadInteger()
    Color = incomingData.ReadLong()
    
    If charindex = 0 Then Exit Sub

    Dim x As Integer, y As Integer, OffsetX As Integer, OffsetY As Integer
    
    With charlist(charindex)
        x = .pos.x
        y = .pos.y
        
        OffsetX = .MoveOffsetX + .Body.HeadOffset.x
        OffsetY = .MoveOffsetY + .Body.HeadOffset.y

    End With
    
    If InMapBounds(x, y) Then
    
        With MapData(x, y)
            Dim Index As Integer
            
            If UBound(.DialogEffects) = 0 Then
                ReDim .DialogEffects(1 To 1)
                
                Index = 1
            Else

                For Index = 1 To UBound(.DialogEffects)

                    If .DialogEffects(Index).Text = vbNullString Then
                        Exit For

                    End If

                Next
                
                If Index > UBound(.DialogEffects) Then
                    ReDim .DialogEffects(1 To UBound(.DialogEffects) + 1)

                End If

            End If
            
            With .DialogEffects(Index)
            
                .Color = RGBA_From_vbColor(Color)
                .Start = FrameTime
                .Text = Text
                .offset.x = OffsetX
                .offset.y = OffsetY
            
            End With

        End With
        
    End If

    Exit Sub
    
ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleTextCharDrop", Erl)
    Call incomingData.SafeClearPacket

End Sub

''
' Handles the ConsoleMessage message.

Private Sub HandleConsoleMessage()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim chat      As String
    Dim FontIndex As Integer
    Dim str       As String
    Dim r         As Byte
    Dim G         As Byte
    Dim B         As Byte
    Dim QueEs     As String
    Dim NpcName   As String
    Dim objname   As String
    Dim Hechizo   As Byte
    Dim UserName  As String
    Dim Valor     As String

    chat = incomingData.ReadASCIIString()
    FontIndex = incomingData.ReadByte()
    
    If ChatGlobal = 0 And FontIndex = FontTypeNames.FONTTYPE_GLOBAL Then Exit Sub

    QueEs = ReadField(1, chat, Asc("*"))

    Select Case QueEs

        Case "NPCNAME"
            NpcName = NpcData(ReadField(2, chat, Asc("*"))).Name
            chat = NpcName & ReadField(3, chat, Asc("*"))

        Case "O" 'OBJETO
            objname = ObjData(ReadField(2, chat, Asc("*"))).Name
            chat = objname & ReadField(3, chat, Asc("*"))

        Case "HECINF"
            Hechizo = ReadField(2, chat, Asc("*"))
            chat = "------------< Información del hechizo >------------" & vbCrLf & "Nombre: " & HechizoData(Hechizo).nombre & vbCrLf & "Descripción: " & HechizoData(Hechizo).desc & vbCrLf & "Skill requerido: " & HechizoData(Hechizo).MinSkill & " de magia." & vbCrLf & "Mana necesario: " & HechizoData(Hechizo).ManaRequerido & " puntos." & vbCrLf & "Stamina necesaria: " & HechizoData(Hechizo).StaRequerido & " puntos."

        Case "ProMSG"
            Hechizo = ReadField(2, chat, Asc("*"))
            chat = HechizoData(Hechizo).PropioMsg

        Case "HecMSG"
            Hechizo = ReadField(2, chat, Asc("*"))
            chat = HechizoData(Hechizo).HechizeroMsg & " la criatura."

        Case "HecMSGU"
            Hechizo = ReadField(2, chat, Asc("*"))
            UserName = ReadField(3, chat, Asc("*"))
            chat = HechizoData(Hechizo).HechizeroMsg & " " & UserName & "."
                
        Case "HecMSGA"
            Hechizo = ReadField(2, chat, Asc("*"))
            UserName = ReadField(3, chat, Asc("*"))
            chat = UserName & " " & HechizoData(Hechizo).TargetMsg
                
        Case "EXP"
            Valor = ReadField(2, chat, Asc("*"))
            'chat = "Has ganado " & valor & " puntos de experiencia."
        
        Case "ID"

            Dim ID    As Integer
            Dim extra As String

            ID = ReadField(2, chat, Asc("*"))
            extra = ReadField(3, chat, Asc("*"))
                
            chat = Locale_Parse_ServerMessage(ID, extra)
           
    End Select
    
    If InStr(1, chat, "~") Then
        str = ReadField(2, chat, 126)

        If Val(str) > 255 Then
            r = 255
        Else
            r = Val(str)

        End If
            
        str = ReadField(3, chat, 126)

        If Val(str) > 255 Then
            G = 255
        Else
            G = Val(str)

        End If
            
        str = ReadField(4, chat, 126)

        If Val(str) > 255 Then
            B = 255
        Else
            B = Val(str)

        End If
            
        Call AddtoRichTextBox(frmMain.RecTxt, Left$(chat, InStr(1, chat, "~") - 1), r, G, B, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
    
    Else

        With FontTypes(FontIndex)
            Call AddtoRichTextBox(frmMain.RecTxt, chat, .red, .green, .blue, .bold, .italic)

        End With

    End If
    
    Exit Sub
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleConsoleMessage", Erl)
    Call incomingData.SafeClearPacket

End Sub

Private Sub HandleLocaleMsg()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim chat      As String

    Dim FontIndex As Integer

    Dim str       As String

    Dim r         As Byte

    Dim G         As Byte

    Dim B         As Byte

    Dim QueEs     As String

    Dim NpcName   As String

    Dim objname   As String

    Dim Hechizo   As Byte

    Dim UserName  As String

    Dim Valor     As String

    Dim ID        As Integer

    ID = incomingData.ReadInteger()
    chat = incomingData.ReadASCIIString()
    FontIndex = incomingData.ReadByte()

    chat = Locale_Parse_ServerMessage(ID, chat)
    
    If InStr(1, chat, "~") Then
        str = ReadField(2, chat, 126)

        If Val(str) > 255 Then
            r = 255
        Else
            r = Val(str)

        End If
            
        str = ReadField(3, chat, 126)

        If Val(str) > 255 Then
            G = 255
        Else
            G = Val(str)

        End If
            
        str = ReadField(4, chat, 126)

        If Val(str) > 255 Then
            B = 255
        Else
            B = Val(str)

        End If
            
        Call AddtoRichTextBox(frmMain.RecTxt, Left$(chat, InStr(1, chat, "~") - 1), r, G, B, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
    Else

        With FontTypes(FontIndex)
            Call AddtoRichTextBox(frmMain.RecTxt, chat, .red, .green, .blue, .bold, .italic)

        End With

    End If
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleLocaleMsg", Erl)
    Call incomingData.SafeClearPacket

End Sub

''
' Handles the GuildChat message.

Private Sub HandleGuildChat()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/07/08 (NicoNZ)
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim chat As String

    Dim status As Byte

    Dim str  As String

    Dim r    As Byte

    Dim G    As Byte

    Dim B    As Byte

    Dim tmp  As Integer

    Dim Cont As Integer
    status = incomingData.ReadByte()
    chat = incomingData.ReadASCIIString()
    
    If Not DialogosClanes.Activo Then
        If InStr(1, chat, "~") Then
            str = ReadField(2, chat, 126)
    
            If Val(str) > 255 Then
                r = 255
            Else
                r = Val(str)
    
            End If
                
            str = ReadField(3, chat, 126)
    
            If Val(str) > 255 Then
                G = 255
            Else
                G = Val(str)
    
            End If
                
            str = ReadField(4, chat, 126)
    
            If Val(str) > 255 Then
                B = 255
            Else
                B = Val(str)
            End If
                
            Call AddtoRichTextBox(frmMain.RecTxt, Left$(chat, InStr(1, chat, "~") - 1), r, G, B, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
        Else
            With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
                Call AddtoRichTextBox(frmMain.RecTxt, chat, .red, .green, .blue, .bold, .italic)
            End With
        End If
    Else
        Call DialogosClanes.PushBackText(ReadField(1, chat, 126), status)
    End If
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildChat", Erl)
    Call incomingData.SafeClearPacket

End Sub

''
' Handles the ShowMessageBox message.

Private Sub HandleShowMessageBox()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim mensaje As String

    mensaje = incomingData.ReadASCIIString()

    Select Case QueRender

        Case 0
            frmMensaje.msg.Caption = mensaje
            frmMensaje.Show , frmMain

        Case 1
            Call Sound.Sound_Play(SND_EXCLAMACION)
            Call TextoAlAsistente(mensaje)
            Call Long_2_RGBAList(textcolorAsistente, -1)

        Case 2
            frmMensaje.Show
            frmMensaje.msg.Caption = mensaje
        
        Case 3
            frmMensaje.Show , frmConnect
            frmMensaje.msg.Caption = mensaje

    End Select
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowMessageBox", Erl)
    Call incomingData.SafeClearPacket

End Sub

Private Sub HandleMostrarCuenta()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    ' FrmCuenta.Show
    AlphaNiebla = 30
    frmConnect.Visible = True
    QueRender = 2
    
    'UserMap = 323
    
    'Call SwitchMap(UserMap)
    
    SugerenciaAMostrar = RandomNumber(1, NumSug)
        
    ' LogeoAlgunaVez = True
    Call Sound.Sound_Play(192)
    
    Call Sound.Sound_Stop(SND_LLUVIAIN)
    '  Sound.NextMusic = 2
    '  Sound.Fading = 350
      
    Call Graficos_Particulas.Particle_Group_Remove_All
    Call Graficos_Particulas.Engine_Select_Particle_Set(203)
    ParticleLluviaDorada = Graficos_Particulas.General_Particle_Create(208, -1, -1)
    
    frmConnect.relampago.Enabled = False
            
    If FrmLogear.Visible Then
        Unload FrmLogear

        'Unload frmConnect
    End If
    
    If frmMain.Visible Then
        '  frmMain.Visible = False
        
        UserParalizado = False
        UserInmovilizado = False
        UserStopped = False
        
        InvasionActual = 0
        frmMain.Evento.Enabled = False
     
        'BUG CLONES
        Dim i As Integer

        For i = 1 To LastChar
            Call EraseChar(i)
        Next i
        
        frmMain.personaje(1).Visible = False
        frmMain.personaje(2).Visible = False
        frmMain.personaje(3).Visible = False
        frmMain.personaje(4).Visible = False
        frmMain.personaje(5).Visible = False

    End If
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleMostrarCuenta", Erl)
    Call incomingData.SafeClearPacket

End Sub

''
' Handles the UserIndexInServer message.

Private Sub HandleUserIndexInServer()
    
    On Error GoTo HandleUserIndexInServer_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    userIndex = incomingData.ReadInteger()
    
    Exit Sub

HandleUserIndexInServer_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUserIndexInServer", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the UserCharIndexInServer message.

Private Sub HandleUserCharIndexInServer()
    
    On Error GoTo HandleUserCharIndexInServer_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    UserCharIndex = incomingData.ReadInteger()
    UserPos = charlist(UserCharIndex).pos
    
    'Are we under a roof?
    bTecho = HayTecho(UserPos.x, UserPos.y)
    
    LastMove = FrameTime
    
    frmMain.Coord.Caption = UserMap & "-" & UserPos.x & "-" & UserPos.y
    Call frmMain.SetMinimapPosition(0, UserPos.x, UserPos.y)
    
    If frmMapaGrande.Visible Then
        Call frmMapaGrande.ActualizarPosicionMapa
    End If
    
    Exit Sub

HandleUserCharIndexInServer_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUserCharIndexInServer", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the CharacterCreate message.

Private Sub HandleCharacterCreate()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim charindex     As Integer

    Dim Body          As Integer

    Dim Head          As Integer

    Dim Heading       As E_Heading

    Dim x             As Byte

    Dim y             As Byte

    Dim weapon        As Integer

    Dim shield        As Integer

    Dim helmet        As Integer

    Dim privs         As Integer

    Dim AuraParticula As Byte

    Dim ParticulaFx   As Byte

    Dim appear        As Byte

    Dim group_index   As Integer
    
    charindex = incomingData.ReadInteger()
    Body = incomingData.ReadInteger()
    Head = incomingData.ReadInteger()
    Heading = incomingData.ReadByte()
    x = incomingData.ReadByte()
    y = incomingData.ReadByte()
    weapon = incomingData.ReadInteger()
    shield = incomingData.ReadInteger()
    helmet = incomingData.ReadInteger()
    
    With charlist(charindex)
        'Call SetCharacterFx(charindex, incomingData.ReadInteger(), incomingData.ReadInteger())
        .FxIndex = incomingData.ReadInteger
        
        incomingData.ReadInteger 'Ignore loops
        
        If .FxIndex > 0 Then
            Call InitGrh(.fX, FxData(.FxIndex).Animacion)

        End If
        
        Dim NombreYClan As String
        NombreYClan = incomingData.ReadASCIIString()
        
        Dim pos As Integer
        pos = InStr(NombreYClan, "<")

        If pos = 0 Then pos = InStr(NombreYClan, "[")
        If pos = 0 Then pos = Len(NombreYClan) + 2
        
        .nombre = Left$(NombreYClan, pos - 2)
        .clan = mid$(NombreYClan, pos)
        
        .status = incomingData.ReadByte()
        
        privs = incomingData.ReadByte()
        ParticulaFx = incomingData.ReadByte()
        .Head_Aura = incomingData.ReadASCIIString()
        .Arma_Aura = incomingData.ReadASCIIString()
        .Body_Aura = incomingData.ReadASCIIString()
        .DM_Aura = incomingData.ReadASCIIString()
        .RM_Aura = incomingData.ReadASCIIString()
        .Otra_Aura = incomingData.ReadASCIIString()
        .Escudo_Aura = incomingData.ReadASCIIString()
        .Speeding = incomingData.ReadSingle()
        
        Dim FlagNpc As Byte
        FlagNpc = incomingData.ReadByte()
        
        .EsNpc = FlagNpc > 0
        .EsMascota = FlagNpc = 2
        
        .Donador = incomingData.ReadByte()
        .appear = incomingData.ReadByte()
        appear = .appear
        .group_index = incomingData.ReadInteger()
        .clan_index = incomingData.ReadInteger()
        .clan_nivel = incomingData.ReadByte()
        .UserMinHp = incomingData.ReadLong()
        .UserMaxHp = incomingData.ReadLong()
        .UserMinMAN = incomingData.ReadLong()
        .UserMaxMAN = incomingData.ReadLong()
        .simbolo = incomingData.ReadByte()
        .Idle = incomingData.ReadBoolean()
        .Navegando = incomingData.ReadBoolean()
        
        If (.pos.x <> 0 And .pos.y <> 0) Then
            If MapData(.pos.x, .pos.y).charindex = charindex Then
                'Erase the old character from map
                MapData(charlist(charindex).pos.x, charlist(charindex).pos.y).charindex = 0

            End If

        End If

        If privs <> 0 Then

            'If the player belongs to a council AND is an admin, only whos as an admin
            If (privs And PlayerType.ChaosCouncil) <> 0 And (privs And PlayerType.User) = 0 Then
                privs = privs Xor PlayerType.ChaosCouncil

            End If
            
            If (privs And PlayerType.RoyalCouncil) <> 0 And (privs And PlayerType.User) = 0 Then
                privs = privs Xor PlayerType.RoyalCouncil

            End If
            
            'If the player is a RM, ignore other flags
            If privs And PlayerType.RoleMaster Then
                privs = PlayerType.RoleMaster

            End If
            
            'Log2 of the bit flags sent by the server gives our numbers ^^
            .priv = Log(privs) / Log(2)
        Else
            .priv = 0

        End If

        .Muerto = (Body = CASPER_BODY_IDLE)
        '.AlphaPJ = 255
    
        Call MakeChar(charindex, Body, Head, Heading, x, y, weapon, shield, helmet, ParticulaFx, appear)
        
        If .Idle Or .Navegando Then
            'Start animation
            .Body.Walk(.Heading).Started = FrameTime

        End If
        
    End With
    
    Call RefreshAllChars
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCharacterCreate", Erl)
    Call incomingData.SafeClearPacket

End Sub

''
' Handles the CharacterRemove message.

Private Sub HandleCharacterRemove()
    
    On Error GoTo HandleCharacterRemove_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim charindex   As Integer

    Dim Desvanecido As Boolean
    
    charindex = incomingData.ReadInteger()
    Desvanecido = incomingData.ReadBoolean()
    
    If Desvanecido And charlist(charindex).EsNpc = True Then
        Call CrearFantasma(charindex)

    End If

    Call EraseChar(charindex)
    Call RefreshAllChars
    
    Exit Sub

HandleCharacterRemove_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCharacterRemove", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the CharacterMove message.

Private Sub HandleCharacterMove()
    
    On Error GoTo HandleCharacterMove_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim charindex As Integer
    Dim x         As Byte
    Dim y         As Byte
    
    charindex = incomingData.ReadInteger()
    x = incomingData.ReadByte()
    y = incomingData.ReadByte()
    
    With charlist(charindex)
        
        ' Play steps sounds if the user is not an admin of any kind
        If .priv <> 1 And .priv <> 2 And .priv <> 3 And .priv <> 5 And .priv <> 25 Then
            Call DoPasosFx(charindex)

        End If

    End With
    
    Call Char_Move_by_Pos(charindex, x, y)
    
    Call RefreshAllChars
    
    Exit Sub

HandleCharacterMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCharacterMove", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the ForceCharMove message.

Private Sub HandleForceCharMove()
    
    On Error GoTo HandleForceCharMove_Err
    
    Dim Direccion As Byte
    Direccion = incomingData.ReadByte()
    
    Moviendose = True
    
    Call MainTimer.Restart(TimersIndex.Walk)

    Call Char_Move_by_Head(UserCharIndex, Direccion)
    Call MoveScreen(Direccion)
    
    Call frmMain.SetMinimapPosition(0, UserPos.x, UserPos.y)
    
    frmMain.Coord.Caption = UserMap & "-" & UserPos.x & "-" & UserPos.y

    If frmMapaGrande.Visible Then
        Call frmMapaGrande.ActualizarPosicionMapa

    End If
    
    Call RefreshAllChars
    
    Exit Sub

HandleForceCharMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleForceCharMove", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the CharacterChange message.

Private Sub HandleCharacterChange()
    
    On Error GoTo HandleCharacterChange_Err
    
    Dim charindex As Integer

    Dim tempint   As Integer

    Dim headIndex As Integer

    charindex = incomingData.ReadInteger()
    
    With charlist(charindex)
        tempint = incomingData.ReadInteger()

        If tempint < LBound(BodyData()) Or tempint > UBound(BodyData()) Then
            .Body = BodyData(0)
        Else
            .Body = BodyData(tempint)
            .iBody = tempint

        End If
        
        headIndex = incomingData.ReadInteger()

        If headIndex < LBound(HeadData()) Or headIndex > UBound(HeadData()) Then
            .Head = HeadData(0)
            .IHead = 0
            
        Else
            .Head = HeadData(headIndex)
            .IHead = headIndex

        End If

        .Muerto = (.iBody = CASPER_BODY_IDLE)
        
        .Heading = incomingData.ReadByte()
        
        tempint = incomingData.ReadInteger()

        If tempint <> 0 Then .Arma = WeaponAnimData(tempint)

        tempint = incomingData.ReadInteger()

        If tempint <> 0 Then .Escudo = ShieldAnimData(tempint)
        
        tempint = incomingData.ReadInteger()

        If tempint <> 0 Then .Casco = CascoAnimData(tempint)
                
        If .Body.HeadOffset.y = -26 Then
            .EsEnano = True
        Else
            .EsEnano = False

        End If
        
        'Call SetCharacterFx(charindex, incomingData.ReadInteger(), incomingData.ReadInteger())
        .FxIndex = incomingData.ReadInteger
        
        incomingData.ReadInteger 'Ignore loops
        
        If .FxIndex > 0 Then
            Call InitGrh(.fX, FxData(.FxIndex).Animacion)

        End If
        
        .Idle = incomingData.ReadBoolean
        
        .Navegando = incomingData.ReadBoolean
        
        If .Idle Or .Navegando Then
            'Start animation
            .Body.Walk(.Heading).Started = FrameTime

        End If

    End With
    
    Exit Sub

HandleCharacterChange_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCharacterChange", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the ObjectCreate message.

Private Sub HandleObjectCreate()
    
    On Error GoTo HandleObjectCreate_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim x        As Byte

    Dim y        As Byte

    Dim OBJIndex As Integer
    
    Dim Amount   As Integer

    Dim Color    As RGBA

    Dim Rango    As Byte

    Dim ID       As Long
    
    x = incomingData.ReadByte()
    y = incomingData.ReadByte()
    
    OBJIndex = incomingData.ReadInteger()
    
    Amount = incomingData.ReadInteger
    
    MapData(x, y).ObjGrh.GrhIndex = ObjData(OBJIndex).GrhIndex
    
    MapData(x, y).OBJInfo.OBJIndex = OBJIndex
    
    MapData(x, y).OBJInfo.Amount = Amount
    
    Call InitGrh(MapData(x, y).ObjGrh, MapData(x, y).ObjGrh.GrhIndex)
    
    If ObjData(OBJIndex).CreaLuz <> "" Then
        Call Long_2_RGBA(Color, Val(ReadField(2, ObjData(OBJIndex).CreaLuz, Asc(":"))))
        Rango = Val(ReadField(1, ObjData(OBJIndex).CreaLuz, Asc(":")))
        MapData(x, y).luz.Color = Color
        MapData(x, y).luz.Rango = Rango
        
        If Rango < 100 Then
            ID = x & y
            LucesCuadradas.Light_Create x, y, Color, Rango, ID
            LucesCuadradas.Light_Render_All
        Else
            LucesRedondas.Create_Light_To_Map x, y, Color, Rango - 99
            LucesRedondas.LightRenderAll
            LucesCuadradas.Light_Render_All

        End If
        
    End If
        
    If ObjData(OBJIndex).CreaParticulaPiso <> 0 Then
        MapData(x, y).particle_group = 0
        General_Particle_Create ObjData(OBJIndex).CreaParticulaPiso, x, y, -1

    End If
    
    Exit Sub

HandleObjectCreate_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleObjectCreate", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleFxPiso()
    
    On Error GoTo HandleFxPiso_Err

    '***************************************************
    'Ladder
    '30/5/10
    '***************************************************
    
    Dim x  As Byte

    Dim y  As Byte

    Dim fX As Byte

    x = incomingData.ReadByte()
    y = incomingData.ReadByte()
    fX = incomingData.ReadInteger()
    
    Call SetMapFx(x, y, fX, 0)
    
    Exit Sub

HandleFxPiso_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleFxPiso", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the ObjectDelete message.

Private Sub HandleObjectDelete()
    
    On Error GoTo HandleObjectDelete_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim x  As Byte

    Dim y  As Byte

    Dim ID As Long
    
    x = incomingData.ReadByte()
    y = incomingData.ReadByte()
    
    If ObjData(MapData(x, y).OBJInfo.OBJIndex).CreaLuz <> "" Then
        ID = LucesCuadradas.Light_Find(x & y)
        LucesCuadradas.Light_Remove ID
        MapData(x, y).luz.Color = COLOR_EMPTY
        MapData(x, y).luz.Rango = 0
        LucesCuadradas.Light_Render_All

    End If
    
    MapData(x, y).ObjGrh.GrhIndex = 0
    MapData(x, y).OBJInfo.OBJIndex = 0
    
    If ObjData(MapData(x, y).OBJInfo.OBJIndex).CreaParticulaPiso <> 0 Then
        Graficos_Particulas.Particle_Group_Remove (MapData(x, y).particle_group)

    End If
    
    Exit Sub

HandleObjectDelete_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleObjectDelete", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the BlockPosition message.

Private Sub HandleBlockPosition()
    
    On Error GoTo HandleBlockPosition_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim x As Byte, y As Byte, B As Byte
    
    x = incomingData.ReadByte()
    y = incomingData.ReadByte()
    B = incomingData.ReadByte()

    MapData(x, y).Blocked = MapData(x, y).Blocked And Not eBlock.ALL_SIDES
    MapData(x, y).Blocked = MapData(x, y).Blocked Or B
    
    Exit Sub

HandleBlockPosition_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBlockPosition", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the PlayMIDI message.

Private Sub HandlePlayMIDI()
    
    On Error GoTo HandlePlayMIDI_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim currentMidi As Byte
    
    currentMidi = incomingData.ReadByte()
    
    If currentMidi And mp3 = 0 Then
        ' SEngine.Music_MP3_Empty
        '  Call Audio.PlayMIDI(CStr(currentMidi) & ".mid", incomingData.ReadInteger())
    Else
        'Remove the bytes to prevent errors
        Call incomingData.ReadInteger

    End If
    
    Exit Sub

HandlePlayMIDI_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePlayMIDI", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the PlayWave message.

Private Sub HandlePlayWave()
    
    On Error GoTo HandlePlayWave_Err

    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 08/14/07
    'Last Modified by: Rapsodius
    'Added support for 3D Sounds.
    '***************************************************
        
    Dim wave As Integer

    Dim srcX As Byte

    Dim srcY As Byte
    
    wave = incomingData.ReadInteger()
    srcX = incomingData.ReadByte()
    srcY = incomingData.ReadByte()
    
    If wave = 400 And MapDat.niebla = 0 Then Exit Sub
    If wave = 401 And MapDat.niebla = 0 Then Exit Sub
    If wave = 402 And MapDat.niebla = 0 Then Exit Sub
    If wave = 403 And MapDat.niebla = 0 Then Exit Sub
    If wave = 404 And MapDat.niebla = 0 Then Exit Sub
    
    If srcX = 0 Or srcY = 0 Then
        Call Sound.Sound_Play(CStr(wave), False, 0, 0)
    Else

        If Not EstaEnArea(srcX, srcY) Then
        Else
            Call Sound.Sound_Play(CStr(wave), False, Sound.Calculate_Volume(srcX, srcY), Sound.Calculate_Pan(srcX, srcY))

        End If

    End If
    
    ' Call Audio.PlayWave(CStr(wave) & ".wav", srcX, srcY)
    
    Exit Sub

HandlePlayWave_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePlayWave", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandlePosLLamadaDeClan()
    
    On Error GoTo HandlePosLLamadaDeClan_Err

    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 08/14/07
    'Last Modified by: Rapsodius
    'Added support for 3D Sounds.
    '***************************************************
        
    Dim map  As Integer

    Dim srcX As Byte

    Dim srcY As Byte
    
    map = incomingData.ReadInteger()
    srcX = incomingData.ReadByte()
    srcY = incomingData.ReadByte()

    Dim idmap As Integer

    LLamadaDeclanMapa = map
    idmap = ObtenerIdMapaDeLlamadaDeClan(map)

    Dim x As Long

    Dim y As Long
    
    x = (idmap - 1) Mod 14
    y = Int((idmap - 1) / 14)

    'frmMapaGrande.lblAllies.Top = Y * 32
    'frmMapaGrande.lblAllies.Left = X * 32

    frmMapaGrande.llamadadeclan.Top = y * 32 + (srcX / 4.5)
    frmMapaGrande.llamadadeclan.Left = x * 32 + (srcY / 4.5)

    frmMapaGrande.llamadadeclan.Visible = True

    frmMain.LlamaDeclan.Enabled = True

    frmMapaGrande.Shape2.Visible = True

    frmMapaGrande.Shape2.Top = y * 32
    frmMapaGrande.Shape2.Left = x * 32

    LLamadaDeclanX = srcX
    LLamadaDeclanY = srcY

    HayLLamadaDeclan = True
    
    ' Call Audio.PlayWave(CStr(wave) & ".wav", srcX, srcY)
    
    Exit Sub

HandlePosLLamadaDeClan_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePosLLamadaDeClan", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleCharUpdateHP()
    
    On Error GoTo HandleCharUpdateHP_Err

    Dim charindex As Integer

    Dim minhp     As Long

    Dim maxhp     As Long
    
    charindex = incomingData.ReadInteger()
    minhp = incomingData.ReadLong()
    maxhp = incomingData.ReadLong()

    charlist(charindex).UserMinHp = minhp
    charlist(charindex).UserMaxHp = maxhp
    
    Exit Sub

HandleCharUpdateHP_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCharUpdateHP", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleCharUpdateMAN()
    
    On Error GoTo HandleCharUpdateHP_Err

    Dim charindex As Integer

    Dim minman     As Long

    Dim maxman     As Long
    
    charindex = incomingData.ReadInteger()
    minman = incomingData.ReadLong()
    maxman = incomingData.ReadLong()

    charlist(charindex).UserMinMAN = minman
    charlist(charindex).UserMaxMAN = maxman
    
    Exit Sub

HandleCharUpdateHP_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCharUpdateMAN", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleArmaMov()
    
    On Error GoTo HandleArmaMov_Err

    '***************************************************

    Dim charindex As Integer

    charindex = incomingData.ReadInteger()

    With charlist(charindex)

        If Not .Moving Then
            .MovArmaEscudo = True
            .Arma.WeaponWalk(.Heading).Started = FrameTime
            .Arma.WeaponWalk(.Heading).Loops = 0

        End If

    End With
    
    Exit Sub

HandleArmaMov_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleArmaMov", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleEscudoMov()
    
    On Error GoTo HandleEscudoMov_Err

    '***************************************************

    Dim charindex As Integer

    charindex = incomingData.ReadInteger()

    With charlist(charindex)

        If Not .Moving Then
            .MovArmaEscudo = True
            .Escudo.ShieldWalk(.Heading).Started = FrameTime
            .Escudo.ShieldWalk(.Heading).Loops = 0

        End If

    End With
    
    Exit Sub

HandleEscudoMov_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleEscudoMov", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the GuildList message.

Private Sub HandleGuildList()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    'Clear guild's list
    frmGuildAdm.guildslist.Clear
    
    Dim guildsStr As String
    guildsStr = incomingData.ReadASCIIString()
    
    If Len(guildsStr) > 0 Then

        Dim guilds() As String
        guilds = Split(guildsStr, SEPARATOR)
        
        ReDim ClanesList(0 To UBound(guilds())) As Tclan
        
        ListaClanes = True
        
        Dim i As Long

        For i = 0 To UBound(guilds())
            ClanesList(i).nombre = ReadField(1, guilds(i), Asc("-"))
            ClanesList(i).Alineacion = Val(ReadField(2, guilds(i), Asc("-")))
            ClanesList(i).indice = i
        Next i
        
        For i = 0 To UBound(guilds())
            'If ClanesList(i).Alineacion = 0 Then
            Call frmGuildAdm.guildslist.AddItem(ClanesList(i).nombre)
            'End If
        Next i

    End If
    
    COLOR_AZUL = RGB(0, 0, 0)
    
    Call Establecer_Borde(frmGuildAdm.guildslist, frmGuildAdm, COLOR_AZUL, 0, 0)

    Call frmGuildAdm.Show(vbModeless, frmMain)
    
    Exit Sub
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildList", Erl)
    Call incomingData.SafeClearPacket

End Sub

''
' Handles the AreaChanged message.

Private Sub HandleAreaChanged()
    
    On Error GoTo HandleAreaChanged_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim x As Byte

    Dim y As Byte
    
    x = incomingData.ReadByte()
    y = incomingData.ReadByte()
        
    Call CambioDeArea(x, y)
    
    Exit Sub

HandleAreaChanged_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleAreaChanged", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the PauseToggle message.

Private Sub HandlePauseToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandlePauseToggle_Err
    
    pausa = Not pausa
    
    Exit Sub

HandlePauseToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePauseToggle", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the RainToggle message.

Private Sub HandleRainToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleRainToggle_Err
    
    If Not InMapBounds(UserPos.x, UserPos.y) Then Exit Sub
            
    If bRain Then
        If MapDat.LLUVIA Then
            
            If bTecho Then
                Call Sound.Sound_Play(192)
            Else
                Call Sound.Sound_Play(195)

            End If
            
            Call Sound.Ambient_Stop
            
            Call Graficos_Particulas.Engine_MeteoParticle_Set(-1)

        End If

    Else

        If MapDat.LLUVIA Then
        
            Call Graficos_Particulas.Engine_MeteoParticle_Set(Particula_Lluvia)

        End If

        ' Call Audio.StopWave(AmbientalesBufferIndex)
    End If
    
    bRain = Not bRain
    
    Exit Sub

HandleRainToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleRainToggle", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleTrofeoToggleOn()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleTrofeoToggleOn_Err

    MostrarTrofeo = True
    
    Exit Sub

HandleTrofeoToggleOn_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleTrofeoToggleOn", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleTrofeoToggleOff()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleTrofeoToggleOff_Err

    MostrarTrofeo = False
    
    Exit Sub

HandleTrofeoToggleOff_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleTrofeoToggleOff", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the CreateFX message.

Private Sub HandleCreateFX()
    
    On Error GoTo HandleCreateFX_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim charindex As Integer

    Dim fX        As Integer

    Dim Loops     As Integer
    
    charindex = incomingData.ReadInteger()
    fX = incomingData.ReadInteger()
    Loops = incomingData.ReadInteger()
    
    If fX = 0 Then
        charlist(charindex).fX.AnimacionContador = 29
        Exit Sub

    End If
    
    Call SetCharacterFx(charindex, fX, Loops)
    
    Exit Sub

HandleCreateFX_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCreateFX", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the UpdateUserStats message.

Private Sub HandleUpdateUserStats()
    
    On Error GoTo HandleUpdateUserStats_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    UserMaxHp = incomingData.ReadInteger()
    UserMinHp = incomingData.ReadInteger()
    UserMaxMAN = incomingData.ReadInteger()
    UserMinMAN = incomingData.ReadInteger()
    UserMaxSTA = incomingData.ReadInteger()
    UserMinSTA = incomingData.ReadInteger()
    UserGLD = incomingData.ReadLong()
    UserLvl = incomingData.ReadByte()
    UserPasarNivel = incomingData.ReadLong()
    UserExp = incomingData.ReadLong()
    UserClase = incomingData.ReadByte()
    
    If UserPasarNivel > 0 Then
        frmMain.lblPorcLvl.Caption = Round(UserExp * (100 / UserPasarNivel), 2) & "%"
        frmMain.exp.Caption = PonerPuntos(UserExp) & "/" & PonerPuntos(UserPasarNivel)
        frmMain.EXPBAR.Width = UserExp / UserPasarNivel * 235
    Else
        frmMain.EXPBAR.Width = 235
        frmMain.lblPorcLvl.Caption = "¡Nivel máximo!" 'nivel maximo
        frmMain.exp.Caption = "¡Nivel máximo!"

    End If
    
    frmMain.Hpshp.Width = UserMinHp / UserMaxHp * 216
    frmMain.HpBar.Caption = UserMinHp & " / " & UserMaxHp

    If QuePestañaInferior = 0 Then
        frmMain.Hpshp.Visible = (UserMinHp > 0)

    End If

    If UserMaxMAN > 0 Then
        frmMain.MANShp.Width = UserMinMAN / UserMaxMAN * 216
        frmMain.manabar.Caption = UserMinMAN & " / " & UserMaxMAN

        If QuePestañaInferior = 0 Then
            frmMain.MANShp.Visible = (UserMinMAN > 0)
            frmMain.manabar.Visible = True

        End If

    Else
        frmMain.manabar.Visible = False
        frmMain.MANShp.Width = 0
        frmMain.MANShp.Visible = False

    End If
    
    frmMain.STAShp.Width = UserMinSTA / UserMaxSTA * 89
    frmMain.stabar.Caption = UserMinSTA & " / " & UserMaxSTA
    
    If QuePestañaInferior = 0 Then
        frmMain.STAShp.Visible = (UserMinSTA > 0)

    End If

    frmMain.GldLbl.Caption = PonerPuntos(UserGLD)
    frmMain.lblLvl.Caption = ListaClases(UserClase) & " - Nivel " & UserLvl
    
    If UserMinHp = 0 Then
        UserEstado = 1
    Else
        UserEstado = 0

    End If
    
    Exit Sub

HandleUpdateUserStats_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateUserStats", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the WorkRequestTarget message.

Private Sub HandleWorkRequestTarget()
    
    On Error GoTo HandleWorkRequestTarget_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    Dim UsingSkillREcibido As Byte
     
    UsingSkillREcibido = incomingData.ReadByte()

    If UsingSkillREcibido = 0 Then
        frmMain.MousePointer = 0
        Call FormParser.Parse_Form(frmMain, E_NORMAL)
        UsingSkill = UsingSkillREcibido
        Exit Sub

    End If

    If UsingSkillREcibido = UsingSkill Then Exit Sub
   
    UsingSkill = UsingSkillREcibido
    frmMain.MousePointer = 2

    If ShowMacros = 1 Then
        If OcultarMacrosAlCastear Then
            OcultarMacro = True

        End If

    End If
    
    Select Case UsingSkill

        Case magia
            Call FormParser.Parse_Form(frmMain, E_CAST)
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MAGIA, 100, 100, 120, 0, 0)

        Case Robar
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_ROBAR, 100, 100, 120, 0, 0)
            Call FormParser.Parse_Form(frmMain, E_SHOOT)

        Case FundirMetal
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_FUNDIRMETAL, 100, 100, 120, 0, 0)
            Call FormParser.Parse_Form(frmMain, E_SHOOT)

        Case Proyectiles
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PROYECTILES, 100, 100, 120, 0, 0)
            Call FormParser.Parse_Form(frmMain, E_ARROW)

        Case eSkill.Talar, eSkill.Alquimia, eSkill.Carpinteria, eSkill.Herreria, eSkill.Mineria, eSkill.Pescar
            Call AddtoRichTextBox(frmMain.RecTxt, "Has click donde deseas trabajar...", 100, 100, 120, 0, 0)
            Call FormParser.Parse_Form(frmMain, E_SHOOT)

        Case Grupo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MAGIA, 100, 100, 120, 0, 0)
            Call FormParser.Parse_Form(frmMain, E_SHOOT)

        Case MarcaDeClan
            Call AddtoRichTextBox(frmMain.RecTxt, "Seleccione el personaje que desea marcar..", 100, 100, 120, 0, 0)
            Call FormParser.Parse_Form(frmMain, E_SHOOT)

        Case MarcaDeGM
            Call AddtoRichTextBox(frmMain.RecTxt, "Seleccione el personaje que desea marcar..", 100, 100, 120, 0, 0)
            Call FormParser.Parse_Form(frmMain, E_SHOOT)

    End Select
    
    Exit Sub

HandleWorkRequestTarget_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleWorkRequestTarget", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the ChangeInventorySlot message.

Private Sub HandleChangeInventorySlot()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim Slot        As Byte
    Dim OBJIndex    As Integer
    Dim Name        As String
    Dim Amount      As Integer
    Dim Equipped    As Boolean
    Dim GrhIndex    As Long
    Dim ObjType     As Byte
    Dim MaxHit      As Integer
    Dim MinHit      As Integer
    Dim MaxDef      As Integer
    Dim MinDef      As Integer
    Dim Value       As Single
    Dim podrausarlo As Byte

    Slot = incomingData.ReadByte()
    OBJIndex = incomingData.ReadInteger()
    Amount = incomingData.ReadInteger()
    Equipped = incomingData.ReadBoolean()
    Value = incomingData.ReadSingle()
    podrausarlo = incomingData.ReadByte()

    Name = ObjData(OBJIndex).Name
    GrhIndex = ObjData(OBJIndex).GrhIndex
    ObjType = ObjData(OBJIndex).ObjType
    MaxHit = ObjData(OBJIndex).MaxHit
    MinHit = ObjData(OBJIndex).MinHit
    MaxDef = ObjData(OBJIndex).MaxDef
    MinDef = ObjData(OBJIndex).MinDef

    If Equipped Then

        Select Case ObjType

            Case eObjType.otWeapon
                frmMain.lblWeapon = MinHit & "/" & MaxHit
                UserWeaponEqpSlot = Slot

            Case eObjType.otNudillos
                frmMain.lblWeapon = MinHit & "/" & MaxHit
                UserWeaponEqpSlot = Slot

            Case eObjType.otArmadura
                frmMain.lblArmor = MinDef & "/" & MaxDef
                UserArmourEqpSlot = Slot

            Case eObjType.otESCUDO
                frmMain.lblShielder = MinDef & "/" & MaxDef
                UserHelmEqpSlot = Slot

            Case eObjType.otCASCO
                frmMain.lblHelm = MinDef & "/" & MaxDef
                UserShieldEqpSlot = Slot

        End Select
        
    Else

        Select Case Slot

            Case UserWeaponEqpSlot
                frmMain.lblWeapon = "0/0"
                UserWeaponEqpSlot = 0

            Case UserArmourEqpSlot
                frmMain.lblArmor = "0/0"
                UserArmourEqpSlot = 0

            Case UserHelmEqpSlot
                frmMain.lblShielder = "0/0"
                UserHelmEqpSlot = 0

            Case UserShieldEqpSlot
                frmMain.lblHelm = "0/0"
                UserShieldEqpSlot = 0

        End Select

    End If

    Call frmMain.Inventario.SetItem(Slot, OBJIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, MinDef, Value, Name, podrausarlo)

    If frmComerciar.Visible Then
        Call frmComerciar.InvComUsu.SetItem(Slot, OBJIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, MinDef, Value, Name, podrausarlo)

    ElseIf frmBancoObj.Visible Then
        Call frmBancoObj.InvBankUsu.SetItem(Slot, OBJIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, MinDef, Value, Name, podrausarlo)
        
    ElseIf frmBancoCuenta.Visible Then
        Call frmBancoCuenta.InvBankUsuCuenta.SetItem(Slot, OBJIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, MinDef, Value, Name, podrausarlo)
    
    ElseIf frmCrafteo.Visible Then
        Call frmCrafteo.InvCraftUser.SetItem(Slot, OBJIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, MinDef, Value, Name, podrausarlo)
    End If

    Exit Sub
    
ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChangeInventorySlot", Erl)
    Call incomingData.SafeClearPacket

End Sub

' Handles the InventoryUnlockSlots message.
Private Sub HandleInventoryUnlockSlots()
    '***************************************************
    'Author: Ruthnar
    'Last Modification: 30/09/20
    '
    '***************************************************
    
    On Error GoTo HandleInventoryUnlockSlots_Err
    
    Dim i As Integer
    
    UserInvUnlocked = incomingData.ReadByte
    
    For i = 1 To UserInvUnlocked
    
        frmMain.imgInvLock(i - 1).Picture = LoadInterface("inventoryunlocked.bmp")
    
    Next i

    Exit Sub

HandleInventoryUnlockSlots_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleInventoryUnlockSlots", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleRefreshAllInventorySlot()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim Slot             As Byte

    Dim OBJIndex         As Integer

    Dim Name             As String

    Dim Amount           As Integer

    Dim Equipped         As Boolean

    Dim GrhIndex         As Long

    Dim ObjType          As Byte

    Dim MaxHit           As Integer

    Dim MinHit           As Integer

    Dim defense          As Integer

    Dim Value            As Single

    Dim PuedeUsar        As Byte

    Dim rdata            As String

    Dim todo             As String
    
    Dim slotNum(1 To 25) As String
    
    todo = incomingData.ReadASCIIString()
    
    todo = Right$(todo, Len(todo))
    
    Dim i As Byte
    
    For i = 1 To 25
    
        slotNum(i) = ReadField(i, todo, Asc("*")) 'Nick
    
    Next i
        
    For i = 1 To 20
    
        slotNum(i) = Right$(slotNum(i), Len(slotNum(i)))
        Slot = ReadField(1, slotNum(i), Asc("@"))
        Call frmMain.Inventario.SetItem(Slot, ReadField(2, slotNum(i), Asc("@")), ReadField(4, slotNum(i), Asc("@")), ReadField(5, slotNum(i), Asc("@")), ReadField(6, slotNum(i), Asc("@")), ReadField(7, slotNum(i), Asc("@")), ReadField(8, slotNum(i), Asc("@")), ReadField(9, slotNum(i), Asc("@")), ReadField(10, slotNum(i), Asc("@")), ReadField(11, slotNum(i), Asc("@")), ReadField(3, slotNum(i), Asc("@")), 0)
    
        With frmMain.Inventario

            If frmComerciar.Visible Then
                Call frmComerciar.InvComUsu.SetItem(Slot, .OBJIndex(Slot), .Amount(Slot), .Equipped(Slot), .GrhIndex(Slot), .ObjType(Slot), .MaxHit(Slot), .MinHit(Slot), .Def(Slot), .Valor(Slot), .ItemName(Slot), .PuedeUsar(Slot))
            ElseIf frmBancoObj.Visible Then
                Call frmBancoObj.InvBankUsu.SetItem(Slot, .OBJIndex(Slot), .Amount(Slot), .Equipped(Slot), .GrhIndex(Slot), .ObjType(Slot), .MaxHit(Slot), .MinHit(Slot), .Def(Slot), .Valor(Slot), .ItemName(Slot), .PuedeUsar(Slot))
            ElseIf frmBancoCuenta.Visible Then
                Call frmBancoCuenta.InvBankUsuCuenta.SetItem(Slot, .OBJIndex(Slot), .Amount(Slot), .Equipped(Slot), .GrhIndex(Slot), .ObjType(Slot), .MaxHit(Slot), .MinHit(Slot), .Def(Slot), .Valor(Slot), .ItemName(Slot), .PuedeUsar(Slot))

            End If

        End With
    
    Next i

    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleRefreshAllInventorySlot", Erl)
    Call incomingData.SafeClearPacket

End Sub

''
' Handles the ChangeBankSlot message.

Private Sub HandleChangeBankSlot()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim Slot As Byte
    Slot = incomingData.ReadByte()
    
    Dim BankSlot As Inventory
    
    With BankSlot
    
        .OBJIndex = incomingData.ReadInteger()
        .Name = ObjData(.OBJIndex).Name
        .Amount = incomingData.ReadInteger()
        .GrhIndex = ObjData(.OBJIndex).GrhIndex
        .ObjType = ObjData(.OBJIndex).ObjType
        .MaxHit = ObjData(.OBJIndex).MaxHit
        .MinHit = ObjData(.OBJIndex).MinHit
        .Def = ObjData(.OBJIndex).MaxDef
        .Valor = incomingData.ReadLong()
        .PuedeUsar = incomingData.ReadByte()
        
        Call frmBancoObj.InvBoveda.SetItem(Slot, .OBJIndex, .Amount, .Equipped, .GrhIndex, .ObjType, .MaxHit, .MinHit, .Def, .Valor, .Name, .PuedeUsar)

    End With
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChangeBankSlot", Erl)
    Call incomingData.SafeClearPacket

End Sub

''
' Handles the ChangeSpellSlot message

Private Sub HandleChangeSpellSlot()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim Slot     As Byte

    Dim Index    As Byte

    Dim cooldown As Integer

    Slot = incomingData.ReadByte()
    
    UserHechizos(Slot) = incomingData.ReadInteger()
    Index = incomingData.ReadByte()

    If Index < 254 Then
    
        If Slot <= frmMain.hlst.ListCount Then
            frmMain.hlst.List(Slot - 1) = HechizoData(Index).nombre
        Else
            Call frmMain.hlst.AddItem(HechizoData(Index).nombre)

        End If

    Else
    
        If Slot <= frmMain.hlst.ListCount Then
            frmMain.hlst.List(Slot - 1) = "(Vacio)"
        Else
            Call frmMain.hlst.AddItem("(Vacio)")

        End If
    
    End If
    
    Exit Sub
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChangeSpellSlot", Erl)
    Call incomingData.SafeClearPacket

End Sub

''
' Handles the Attributes message.

Private Sub HandleAtributes()
    
    On Error GoTo HandleAtributes_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim i As Long
    
    For i = 1 To NUMATRIBUTES
        UserAtributos(i) = incomingData.ReadByte()
    Next i
    
    'Show them in character creation
    If EstadoLogin = E_MODO.Dados Then

        With frmCrearPersonaje

            If .Visible Then
                .lbFuerza.Caption = UserAtributos(eAtributos.Fuerza)
                .lbAgilidad.Caption = UserAtributos(eAtributos.Agilidad)
                .lbInteligencia.Caption = UserAtributos(eAtributos.Inteligencia)
                .lbConstitucion.Caption = UserAtributos(eAtributos.Constitucion)
                .lbCarisma = UserAtributos(eAtributos.Carisma)

            End If

        End With

    Else

        If LlegaronSkills And LlegaronStats Then
            Alocados = SkillPoints
            frmEstadisticas.puntos.Caption = SkillPoints
            frmEstadisticas.Iniciar_Labels
            frmEstadisticas.Picture = LoadInterface("ventanaestadisticas.bmp")
            frmEstadisticas.Show , frmMain
        Else
            LlegaronAtrib = True

        End If

    End If
    
    Exit Sub

HandleAtributes_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleAtributes", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the BlacksmithWeapons message.

Private Sub HandleBlacksmithWeapons()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim count As Integer

    Dim i     As Long

    Dim tmp   As String
    
    count = incomingData.ReadInteger()
    
    Call frmHerrero.lstArmas.Clear
    
    For i = 1 To count
        ArmasHerrero(i).Index = incomingData.ReadInteger()
        ' tmp = ObjData(ArmasHerrero(i).Index).name        'Get the object's name
        ArmasHerrero(i).LHierro = incomingData.ReadInteger()  'The iron needed
        ArmasHerrero(i).LPlata = incomingData.ReadInteger()    'The silver needed
        ArmasHerrero(i).LOro = incomingData.ReadInteger()    'The gold needed
        
        ' Call frmHerrero.lstArmas.AddItem(tmp)
    Next i
    
    For i = i To UBound(ArmasHerrero())
        ArmasHerrero(i).Index = 0
    Next i
    
    i = 0
    
    Exit Sub
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBlacksmithWeapons", Erl)
    Call incomingData.SafeClearPacket

End Sub

''
' Handles the BlacksmithArmors message.

Private Sub HandleBlacksmithArmors()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim count As Integer

    Dim i     As Long

    Dim tmp   As String
    
    count = incomingData.ReadInteger()
    
    'Call frmHerrero.lstArmaduras.Clear
    
    For i = 1 To count
        tmp = incomingData.ReadASCIIString()         'Get the object's name
        DefensasHerrero(i).LHierro = incomingData.ReadInteger()   'The iron needed
        DefensasHerrero(i).LPlata = incomingData.ReadInteger()   'The silver needed
        DefensasHerrero(i).LOro = incomingData.ReadInteger()   'The gold needed
        
        ' Call frmHerrero.lstArmaduras.AddItem(tmp)
        DefensasHerrero(i).Index = incomingData.ReadInteger()
    Next i
        
    Dim A      As Byte
    Dim e      As Byte
    Dim c      As Byte
    Dim tmpObj As ObjDatas

    A = 0
    e = 0
    c = 0
    
    For i = 1 To UBound(DefensasHerrero())

        If DefensasHerrero(i).Index = 0 Then Exit For
        
        tmpObj = ObjData(DefensasHerrero(i).Index)
        
        If tmpObj.ObjType = 3 Then
           
            ArmadurasHerrero(A).Index = DefensasHerrero(i).Index
            ArmadurasHerrero(A).LHierro = DefensasHerrero(i).LHierro
            ArmadurasHerrero(A).LPlata = DefensasHerrero(i).LPlata
            ArmadurasHerrero(A).LOro = DefensasHerrero(i).LOro
            A = A + 1

        End If
        
        ' Escudos (16), Objetos Magicos (21) y Anillos (35) van en la misma lista
        If tmpObj.ObjType = 16 Or tmpObj.ObjType = 35 Or tmpObj.ObjType = 21 Then
            EscudosHerrero(e).Index = DefensasHerrero(i).Index
            EscudosHerrero(e).LHierro = DefensasHerrero(i).LHierro
            EscudosHerrero(e).LPlata = DefensasHerrero(i).LPlata
            EscudosHerrero(e).LOro = DefensasHerrero(i).LOro
            e = e + 1

        End If

        If tmpObj.ObjType = 17 Then
            CascosHerrero(c).Index = DefensasHerrero(i).Index
            CascosHerrero(c).LHierro = DefensasHerrero(i).LHierro
            CascosHerrero(c).LPlata = DefensasHerrero(i).LPlata
            CascosHerrero(c).LOro = DefensasHerrero(i).LOro
            c = c + 1

        End If

    Next i
    
    Exit Sub
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBlacksmithArmors", Erl)
    Call incomingData.SafeClearPacket

End Sub

''
' Handles the CarpenterObjects message.

Private Sub HandleCarpenterObjects()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim count As Integer

    Dim i     As Long

    Dim tmp   As String
    
    count = incomingData.ReadByte()
    
    Call frmCarp.lstArmas.Clear
    
    For i = 1 To count
        ObjCarpintero(i) = incomingData.ReadInteger()
        
        Call frmCarp.lstArmas.AddItem(ObjData(ObjCarpintero(i)).Name)
    Next i
    
    For i = i To UBound(ObjCarpintero())
        ObjCarpintero(i) = 0
    Next i
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCarpenterObjects", Erl)
    Call incomingData.SafeClearPacket

End Sub

Private Sub HandleSastreObjects()

    '***************************************************
    'Author: Ladder
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim count As Integer

    Dim i     As Long

    Dim tmp   As String
    
    count = incomingData.ReadInteger()
    
    For i = i To UBound(ObjSastre())
        ObjSastre(i).Index = 0
    Next i
    
    i = 0
    
    For i = 1 To count
        ObjSastre(i).Index = incomingData.ReadInteger()
        
        ObjSastre(i).PielLobo = ObjData(ObjSastre(i).Index).PielLobo
        ObjSastre(i).PielOsoPardo = ObjData(ObjSastre(i).Index).PielOsoPardo
        ObjSastre(i).PielOsoPolar = ObjData(ObjSastre(i).Index).PielOsoPolar

    Next i
    
    Dim r As Byte

    Dim G As Byte
    
    i = 0
    r = 1
    G = 1
    
    For i = i To UBound(ObjSastre())
    
        If ObjData(ObjSastre(i).Index).ObjType = 3 Then
        
            SastreRopas(r).Index = ObjSastre(i).Index
            SastreRopas(r).PielLobo = ObjSastre(i).PielLobo
            SastreRopas(r).PielOsoPardo = ObjSastre(i).PielOsoPardo
            SastreRopas(r).PielOsoPolar = ObjSastre(i).PielOsoPolar
            r = r + 1

        End If

        If ObjData(ObjSastre(i).Index).ObjType = 17 Then
            SastreGorros(G).Index = ObjSastre(i).Index
            SastreGorros(G).PielLobo = ObjSastre(i).PielLobo
            SastreGorros(G).PielOsoPardo = ObjSastre(i).PielOsoPardo
            SastreGorros(G).PielOsoPolar = ObjSastre(i).PielOsoPolar
            G = G + 1

        End If

    Next i
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSastreObjects", Erl)
    Call incomingData.SafeClearPacket

End Sub

''
Private Sub HandleAlquimiaObjects()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim count As Integer

    Dim i     As Long

    Dim tmp   As String
    
    Dim Obj   As Integer

    count = incomingData.ReadInteger()
    
    Call frmAlqui.lstArmas.Clear
    
    For i = 1 To count
        Obj = incomingData.ReadInteger()
        tmp = ObjData(Obj).Name        'Get the object's name

        ObjAlquimista(i) = Obj
        Call frmAlqui.lstArmas.AddItem(tmp)
    Next i
    
    For i = i To UBound(ObjAlquimista())
        ObjAlquimista(i) = 0
    Next i
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleAlquimiaObjects", Erl)
    Call incomingData.SafeClearPacket

End Sub

''
' Handles the RestOK message.

Private Sub HandleRestOK()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleRestOK_Err
    
    UserDescansar = Not UserDescansar
    
    Exit Sub

HandleRestOK_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleRestOK", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the ErrorMessage message.

Private Sub HandleErrorMessage()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Call MsgBox(incomingData.ReadASCIIString())
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleErrorMessage", Erl)
    Call incomingData.SafeClearPacket

End Sub

''
' Handles the Blind message.

Private Sub HandleBlind()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleBlind_Err
    
    UserCiego = True
    
    Call SetRGBA(global_light, 4, 4, 4)
    Call MapUpdateGlobalLight
    
    Exit Sub

HandleBlind_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBlind", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the Dumb message.

Private Sub HandleDumb()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleDumb_Err
    
    UserEstupido = True
    
    Exit Sub

HandleDumb_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleDumb", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the ShowSignal message.
'Optimizacion de protocolo por Ladder

Private Sub HandleShowSignal()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim tmp As String
    Dim grh As Integer

    tmp = ObjData(incomingData.ReadInteger()).Texto
    grh = incomingData.ReadInteger()
    
    Call InitCartel(tmp, grh)
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowSignal", Erl)
    Call incomingData.SafeClearPacket

End Sub

''
' Handles the ChangeNPCInventorySlot message.

Private Sub HandleChangeNPCInventorySlot()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim Slot As Byte
    Slot = incomingData.ReadByte()
    
    Dim SlotInv As NpCinV

    With SlotInv
        .OBJIndex = incomingData.ReadInteger()
        .Name = ObjData(.OBJIndex).Name
        .Amount = incomingData.ReadInteger()
        .Valor = incomingData.ReadSingle()
        .GrhIndex = ObjData(.OBJIndex).GrhIndex
        .ObjType = ObjData(.OBJIndex).ObjType
        .MaxHit = ObjData(.OBJIndex).MaxHit
        .MinHit = ObjData(.OBJIndex).MinHit
        .Def = ObjData(.OBJIndex).MaxDef
        .PuedeUsar = incomingData.ReadByte()
        
        Call frmComerciar.InvComNpc.SetItem(Slot, .OBJIndex, .Amount, 0, .GrhIndex, .ObjType, .MaxHit, .MinHit, .Def, .Valor, .Name, .PuedeUsar)
        
    End With
    
    Exit Sub
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChangeNPCInventorySlot", Erl)
    Call incomingData.SafeClearPacket

End Sub

''
' Handles the UpdateHungerAndThirst message.

Private Sub HandleUpdateHungerAndThirst()
    
    On Error GoTo HandleUpdateHungerAndThirst_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    UserMaxAGU = incomingData.ReadByte()
    UserMinAGU = incomingData.ReadByte()
    UserMaxHAM = incomingData.ReadByte()
    UserMinHAM = incomingData.ReadByte()
    frmMain.AGUAsp.Width = UserMinAGU / UserMaxAGU * 32
    frmMain.COMIDAsp.Width = UserMinHAM / UserMaxHAM * 32
    frmMain.AGUbar.Caption = UserMinAGU '& " / " & UserMaxAGU
    frmMain.hambar.Caption = UserMinHAM ' & " / " & UserMaxHAM
    
    If QuePestañaInferior = 0 Then
        frmMain.AGUAsp.Visible = (UserMinAGU > 0)
        frmMain.COMIDAsp.Visible = (UserMinHAM > 0)

    End If

    Exit Sub

HandleUpdateHungerAndThirst_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateHungerAndThirst", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleHora()
    '***************************************************
    
    On Error GoTo HandleHora_Err

    HoraMundo = (timeGetTime And &H7FFFFFFF) - incomingData.ReadLong()
    DuracionDia = incomingData.ReadLong()
    
    If Not Connected Then
        Call RevisarHoraMundo(True)

    End If
    
    Exit Sub

HandleHora_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleHora", Erl)
    Call incomingData.SafeClearPacket
    
End Sub
 
Private Sub HandleLight()
    
    On Error GoTo HandleLight_Err
 
    Dim Color As String
    
    Color = incomingData.ReadASCIIString()

    'Call SetGlobalLight(Map_light_base)
    
    Exit Sub

HandleLight_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleLight", Erl)
    Call incomingData.SafeClearPacket
    
End Sub
 
Private Sub HandleFYA()
    
    On Error GoTo HandleFYA_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    UserAtributos(eAtributos.Fuerza) = incomingData.ReadByte()
    UserAtributos(eAtributos.Agilidad) = incomingData.ReadByte()
    
    DrogaCounter = incomingData.ReadInteger()
    
    If DrogaCounter > 0 Then
        frmMain.Contadores.Enabled = True

    End If
    
    If UserAtributos(eAtributos.Fuerza) >= 35 Then
        frmMain.Fuerzalbl.ForeColor = RGB(204, 0, 0)
    ElseIf UserAtributos(eAtributos.Fuerza) >= 25 Then
        frmMain.Fuerzalbl.ForeColor = RGB(204, 100, 100)
    Else
        frmMain.Fuerzalbl.ForeColor = vbWhite

    End If
    
    If UserAtributos(eAtributos.Agilidad) >= 35 Then
        frmMain.AgilidadLbl.ForeColor = RGB(204, 0, 0)
    ElseIf UserAtributos(eAtributos.Agilidad) >= 25 Then
        frmMain.AgilidadLbl.ForeColor = RGB(204, 100, 100)
    Else
        frmMain.AgilidadLbl.ForeColor = vbWhite

    End If

    frmMain.Fuerzalbl.Caption = UserAtributos(eAtributos.Fuerza)
    frmMain.AgilidadLbl.Caption = UserAtributos(eAtributos.Agilidad)
    
    Exit Sub

HandleFYA_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleFYA", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleUpdateNPCSimbolo()
    
    On Error GoTo HandleUpdateNPCSimbolo_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim NpcIndex As Integer

    Dim simbolo  As Byte
    
    NpcIndex = incomingData.ReadInteger()
    
    simbolo = incomingData.ReadByte()

    charlist(NpcIndex).simbolo = simbolo
    
    Exit Sub

HandleUpdateNPCSimbolo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateNPCSimbolo", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleCerrarleCliente()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleCerrarleCliente_Err
    
    EngineRun = False

    Call CloseClient
    
    Exit Sub

HandleCerrarleCliente_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCerrarleCliente", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleContadores()
    
    On Error GoTo HandleContadores_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    InviCounter = incomingData.ReadInteger()
    ScrollExpCounter = incomingData.ReadInteger()
    ScrollOroCounter = incomingData.ReadInteger()
    OxigenoCounter = incomingData.ReadInteger()
    DrogaCounter = incomingData.ReadInteger()
    
    ScrollExpCounter = ScrollExpCounter
    ScrollOroCounter = ScrollOroCounter
    OxigenoCounter = OxigenoCounter
    
    'Debug.Print ScrollExpCounter
    
    frmMain.Contadores.Enabled = True
    
    Exit Sub

HandleContadores_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleContadores", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleOxigeno()
    
    On Error GoTo HandleOxigeno_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim oxigeno As Integer

    oxigeno = incomingData.ReadInteger()
    
    Dim TextoOxigenoCounter As String

    Dim HR                  As Integer

    Dim ms                  As Integer

    Dim SS                  As Integer

    Dim secs                As Integer
    
    secs = oxigeno
    HR = secs \ 3600
    ms = (secs Mod 3600) \ 60
    SS = (secs Mod 3600) Mod 60

    If SS > 9 Then
        TextoOxigenoCounter = ms & ":" & SS
    Else
        TextoOxigenoCounter = ms & ":0" & SS

    End If

    If ms < 1 Then
        frmMain.oxigenolbl = SS
        frmMain.oxigenolbl.ForeColor = vbRed
    Else
        frmMain.oxigenolbl = ms
        frmMain.oxigenolbl.ForeColor = vbWhite

    End If
    
    Exit Sub

HandleOxigeno_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleOxigeno", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the MiniStats message.
Private Sub HandleFlashScreen()
    
    On Error GoTo HandleEfectToScreen_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim Color As Long, duracion As Long, ignorar As Boolean
    
    Color = incomingData.ReadLong()
    duracion = incomingData.ReadLong()
    ignorar = incomingData.ReadBoolean()
    
    Dim r, G, B As Byte

    B = (Color And 16711680) / 65536
    G = (Color And 65280) / 256
    r = Color And 255
    Color = D3DColorARGB(255, r, G, B)

    If Not MapDat.niebla = 1 And Not ignorar Then
        'Debug.Print "trueno cancelado"
       
        Exit Sub

    End If

    Call EfectoEnPantalla(Color, duracion)
    
    Exit Sub

HandleEfectToScreen_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleEfectToScreen", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleMiniStats()
    
    On Error GoTo HandleMiniStats_Err
    
    With UserEstadisticas
        .CiudadanosMatados = incomingData.ReadLong()
        .CriminalesMatados = incomingData.ReadLong()
        .Alineacion = incomingData.ReadByte()
        
        .NpcsMatados = incomingData.ReadInteger()
        .Clase = ListaClases(incomingData.ReadByte())
        .PenaCarcel = incomingData.ReadLong()
        .VecesQueMoriste = incomingData.ReadLong()
        .Genero = incomingData.ReadByte()

        If .Genero = 1 Then
            .Genero = "Hombre"
        Else
            .Genero = "Mujer"

        End If

        .Raza = incomingData.ReadByte()
        .Raza = ListaRazas(.Raza)
        
        .Donador = incomingData.ReadByte()
        .CreditoDonador = incomingData.ReadLong()
        .DiasRestantes = incomingData.ReadInteger()

    End With
    
    If LlegaronAtrib And LlegaronSkills Then
        Alocados = SkillPoints
        frmEstadisticas.puntos.Caption = SkillPoints
        frmEstadisticas.Iniciar_Labels
        frmEstadisticas.Picture = LoadInterface("ventanaestadisticas.bmp")
        frmEstadisticas.Show , frmMain
    Else
        LlegaronStats = True

    End If
    
    Exit Sub

HandleMiniStats_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleMiniStats", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the LevelUp message.

Private Sub HandleLevelUp()
    
    On Error GoTo HandleLevelUp_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    SkillPoints = incomingData.ReadInteger()
    
    Exit Sub

HandleLevelUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleLevelUp", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the AddForumMessage message.

Private Sub HandleAddForumMessage()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim title   As String

    Dim Message As String
    
    title = incomingData.ReadASCIIString()
    Message = incomingData.ReadASCIIString()

    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleAddForumMessage", Erl)
    Call incomingData.SafeClearPacket

End Sub

''
' Handles the ShowForumForm message.

Private Sub HandleShowForumForm()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleShowForumForm_Err
    
    ' If Not frmForo.Visible Then
    '   frmForo.Show , frmMain
    ' End If
    
    Exit Sub

HandleShowForumForm_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowForumForm", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the SetInvisible message.

Private Sub HandleSetInvisible()
    
    On Error GoTo HandleSetInvisible_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim charindex As Integer
    
    charindex = incomingData.ReadInteger()
    charlist(charindex).Invisible = incomingData.ReadBoolean()
    charlist(charindex).TimerI = 0
    
    Exit Sub

HandleSetInvisible_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSetInvisible", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleSetEscribiendo()
    
    On Error GoTo HandleSetEscribiendo_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim charindex As Integer
    
    charindex = incomingData.ReadInteger()
    charlist(charindex).Escribiendo = incomingData.ReadBoolean()
    
    Exit Sub

HandleSetEscribiendo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSetEscribiendo", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the DiceRoll message.

Private Sub HandleDiceRoll()
    
    On Error GoTo HandleDiceRoll_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    UserAtributos(eAtributos.Fuerza) = incomingData.ReadByte()
    UserAtributos(eAtributos.Agilidad) = incomingData.ReadByte()
    UserAtributos(eAtributos.Inteligencia) = incomingData.ReadByte()
    UserAtributos(eAtributos.Constitucion) = incomingData.ReadByte()
    UserAtributos(eAtributos.Carisma) = incomingData.ReadByte()
    
    frmCrearPersonaje.lbFuerza = UserAtributos(eAtributos.Fuerza)
    frmCrearPersonaje.lbAgilidad = UserAtributos(eAtributos.Agilidad)
    frmCrearPersonaje.lbInteligencia = UserAtributos(eAtributos.Inteligencia)
    frmCrearPersonaje.lbConstitucion = UserAtributos(eAtributos.Constitucion)
    frmCrearPersonaje.lbCarisma = UserAtributos(eAtributos.Carisma)
    
    Exit Sub

HandleDiceRoll_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleDiceRoll", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the MeditateToggle message.

Private Sub HandleMeditateToggle()
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleMeditateToggle_Err
    
    Dim charindex As Integer, fX As Integer
    
    charindex = incomingData.ReadInteger
    fX = incomingData.ReadInteger
    
    If charindex = UserCharIndex Then
        UserMeditar = (fX <> 0)
        
        If UserMeditar Then

            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("Comienzas a meditar.", .red, .green, .blue, .bold, .italic)

            End With

        Else

            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("Has dejado de meditar.", .red, .green, .blue, .bold, .italic)

            End With

        End If

    End If
    
    With charlist(charindex)

        If fX <> 0 Then
            Call InitGrh(.fX, FxData(fX).Animacion)

        End If
        
        .FxIndex = fX
        .fX.Loops = -1
        .fX.AnimacionContador = 0

    End With
    
    Exit Sub

HandleMeditateToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleMeditateToggle", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the BlindNoMore message.

Private Sub HandleBlindNoMore()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleBlindNoMore_Err
    
    UserCiego = False
    
    Call RestaurarLuz
    
    Exit Sub

HandleBlindNoMore_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBlindNoMore", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the DumbNoMore message.

Private Sub HandleDumbNoMore()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleDumbNoMore_Err
    
    UserEstupido = False
    
    Exit Sub

HandleDumbNoMore_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleDumbNoMore", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the SendSkills message.

Private Sub HandleSendSkills()
    
    On Error GoTo HandleSendSkills_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim i As Long
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = incomingData.ReadByte()
        'frmEstadisticas.skills(i).Caption = SkillsNames(i)
    Next i

    If LlegaronAtrib And LlegaronStats Then
        Alocados = SkillPoints
        frmEstadisticas.puntos.Caption = SkillPoints
        frmEstadisticas.Iniciar_Labels
        frmEstadisticas.Picture = LoadInterface("VentanaEstadisticas.bmp")
        frmEstadisticas.Show , frmMain
    Else
        LlegaronSkills = True

    End If
    
    Exit Sub

HandleSendSkills_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSendSkills", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the TrainerCreatureList message.

Private Sub HandleTrainerCreatureList()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim creatures() As String

    Dim i           As Long
    
    creatures = Split(incomingData.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To UBound(creatures())
        Call frmEntrenador.lstCriaturas.AddItem(creatures(i))
    Next i

    frmEntrenador.Show , frmMain
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleTrainerCreatureList", Erl)
    Call incomingData.SafeClearPacket

End Sub

''
' Handles the GuildNews message.

Private Sub HandleGuildNews()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    ' Dim guildList() As String
    Dim List()      As String

    Dim i           As Long
    
    Dim ClanNivel   As Byte

    Dim expacu      As Integer

    Dim ExpNe       As Integer

    Dim guildList() As String
        
    frmGuildNews.news = incomingData.ReadASCIIString()
    
    'Get list of existing guilds
    List = Split(incomingData.ReadASCIIString(), SEPARATOR)
        
    'Empty the list
    Call frmGuildNews.guildslist.Clear
        
    For i = 0 To UBound(List())
        Call frmGuildNews.guildslist.AddItem(ReadField(1, List(i), Asc("-")))
    Next i
    
    'Get  guilds list member
    guildList = Split(incomingData.ReadASCIIString(), SEPARATOR)
    
    Dim cantidad As String

    cantidad = CStr(UBound(guildList()) + 1)
        
    Call frmGuildNews.Miembros.Clear
        
    For i = 0 To UBound(guildList())

        If i = 0 Then
            Call frmGuildNews.Miembros.AddItem(guildList(i) & "(Lider)")
        Else
            Call frmGuildNews.Miembros.AddItem(guildList(i))

        End If

        'Debug.Print guildList(i)
    Next i
    
    ClanNivel = incomingData.ReadByte()
    expacu = incomingData.ReadInteger()
    ExpNe = incomingData.ReadInteger()
     
    With frmGuildNews
        .Frame4.Caption = "Total: " & cantidad & " miembros" '"Lista de miembros" ' - " & cantidad & " totales"
     
        .expcount.Caption = expacu & "/" & ExpNe
        .EXPBAR.Width = (((expacu + 1 / 100) / (ExpNe + 1 / 100)) * 2370)
        .nivel = "Nivel: " & ClanNivel

        If ExpNe > 0 Then
       
            .porciento.Caption = Round(CDbl(expacu) * CDbl(100) / CDbl(ExpNe), 0) & "%"
        Else
            .porciento.Caption = "¡Nivel Maximo!"
            .expcount.Caption = "¡Nivel Maximo!"

        End If
        
        '.expne = "Experiencia necesaria: " & expne
        
        Select Case ClanNivel

            Case 1
                .beneficios = "No atacarse / Chat de clan." & vbCrLf & "Max miembros: 5"

            Case 2
                .beneficios = "No atacarse / Chat de clan / Pedir ayuda (G)." & vbCrLf & "Max miembros: 10"

            Case 3
                .beneficios = "No atacarse / Chat de clan / Pedir ayuda (G) / Verse Invisible." & vbCrLf & "Max miembros: 15"

            Case 4
                .beneficios = "No atacarse / Chat de clan / Pedir ayuda (G) / Verse Invisible / Marca de clan (V)." & vbCrLf & "Max miembros: 20"

            Case 5
                .beneficios = "No atacarse / Chat de clan / Pedir ayuda (G) / Verse Invisible / Marca de clan (V) / Verse vida." & vbCrLf & " Max miembros: 25"
        
        End Select
    
    End With
    
    frmGuildNews.Show vbModeless, frmMain
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildNews", Erl)
    Call incomingData.SafeClearPacket

End Sub

''
' Handles the OfferDetails message.

Private Sub HandleOfferDetails()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Call frmUserRequest.recievePeticion(incomingData.ReadASCIIString())
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleOfferDetails", Erl)
    Call incomingData.SafeClearPacket

End Sub

''
' Handles the AlianceProposalsList message.

Private Sub HandleAlianceProposalsList()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim guildList() As String

    Dim i           As Long
    
    guildList = Split(incomingData.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To UBound(guildList())
        Call frmPeaceProp.lista.AddItem(guildList(i))
    Next i
    
    frmPeaceProp.ProposalType = TIPO_PROPUESTA.ALIANZA
    Call frmPeaceProp.Show(vbModeless, frmMain)
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleAlianceProposalsList", Erl)
    Call incomingData.SafeClearPacket

End Sub

''
' Handles the PeaceProposalsList message.

Private Sub HandlePeaceProposalsList()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim guildList() As String

    Dim i           As Long
    
    guildList = Split(incomingData.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To UBound(guildList())
        Call frmPeaceProp.lista.AddItem(guildList(i))
    Next i
    
    frmPeaceProp.ProposalType = TIPO_PROPUESTA.PAZ
    Call frmPeaceProp.Show(vbModeless, frmMain)
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePeaceProposalsList", Erl)
    Call incomingData.SafeClearPacket

End Sub

''
' Handles the CharacterInfo message.

Private Sub HandleCharacterInfo()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    With frmCharInfo

        If .frmType = CharInfoFrmType.frmMembers Then
            .Rechazar.Visible = False
            .Aceptar.Visible = False
            .Echar.Visible = True
            .desc.Visible = False
        Else
            .Rechazar.Visible = True
            .Aceptar.Visible = True
            .Echar.Visible = False
            .desc.Visible = True

        End If
    
        If incomingData.ReadByte() = 1 Then
            .Genero.Caption = "Genero: Hombre"
        Else
            .Genero.Caption = "Genero: Mujer"
        End If
            
        .nombre.Caption = "Nombre: " & incomingData.ReadASCIIString()
        .Raza.Caption = "Raza: " & ListaRazas(incomingData.ReadByte())
        .Clase.Caption = "Clase: " & ListaClases(incomingData.ReadByte())

        .nivel.Caption = "Nivel: " & incomingData.ReadByte()
        .oro.Caption = "Oro: " & incomingData.ReadLong()
        .Banco.Caption = "Banco: " & incomingData.ReadLong()
    
        .txtPeticiones.Text = incomingData.ReadASCIIString()
        .guildactual.Caption = "Clan: " & incomingData.ReadASCIIString()
        .txtMiembro.Text = incomingData.ReadASCIIString()
            
        Dim armada As Boolean
    
        Dim caos   As Boolean
            
        armada = incomingData.ReadBoolean()
        caos = incomingData.ReadBoolean()
            
        If armada Then
            .ejercito.Caption = "Ejército: Armada Real"
        ElseIf caos Then
            .ejercito.Caption = "Ejército: Legión Oscura"
    
        End If
            
        .ciudadanos.Caption = "Ciudadanos asesinados: " & CStr(incomingData.ReadLong())
        .Criminales.Caption = "Criminales asesinados: " & CStr(incomingData.ReadLong())
    
        Call .Show(vbModeless, frmMain)
    
    End With
        
    Exit Sub
    
ErrHandler:
    
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCharacterInfo", Erl)
    Call incomingData.SafeClearPacket

End Sub

''
' Handles the GuildLeaderInfo message.

Private Sub HandleGuildLeaderInfo()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim List() As String

    Dim i      As Long
    
    With frmGuildLeader
        'Get list of existing guilds
        List = Split(incomingData.ReadASCIIString(), SEPARATOR)
        
        'Empty the list
        Call .guildslist.Clear
        
        For i = 0 To UBound(List())
            Call .guildslist.AddItem(ReadField(1, List(i), Asc("-")))
        Next i
        
        'Get list of guild's members
        List = Split(incomingData.ReadASCIIString(), SEPARATOR)
        .Miembros.Caption = CStr(UBound(List()) + 1)
        
        'Empty the list
        Call .members.Clear
        
        For i = 0 To UBound(List())
            Call .members.AddItem(List(i))
        Next i
        
        .txtguildnews = incomingData.ReadASCIIString()
        
        'Get list of join requests
        List = Split(incomingData.ReadASCIIString(), SEPARATOR)
        
        'Empty the list
        Call .solicitudes.Clear
        
        For i = 0 To UBound(List())
            Call .solicitudes.AddItem(List(i))
        Next i
        
        Dim expacu As Integer

        Dim ExpNe  As Integer

        Dim nivel  As Byte
         
        nivel = incomingData.ReadByte()
        .nivel = "Nivel: " & nivel
        
        expacu = incomingData.ReadInteger()
        ExpNe = incomingData.ReadInteger()
        'barra
        .expcount.Caption = expacu & "/" & ExpNe
        .EXPBAR.Width = expacu / ExpNe * 239
        
        If ExpNe > 0 Then
       
            .porciento.Caption = Round(expacu / ExpNe * 100#, 0) & "%"
        Else
            .porciento.Caption = "¡Nivel máximo!"
            .expcount.Caption = "¡Nivel máximo!"
        End If

        Select Case nivel

            Case 1
                .beneficios = "Chat de clan + Verse en minimapa."
                .maxMiembros = "15"

            Case 2
                .beneficios = "Chat de clan + Verse en minimapa + Pedir ayuda (G)."
                .maxMiembros = "20"

            Case 3
                .beneficios = "Chat de clan + Verse en minimapa + Pedir ayuda (G) + Marca de clan (V)."
                .maxMiembros = "25"

            Case Else
                .beneficios = "Chat de clan + Verse en minimapa + Pedir ayuda (G) + Marca de clan (V) + Ver vidas y maná de tus compañeros."
                .maxMiembros = "30"
        
        End Select
        
        .Show , frmMain

    End With
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildLeaderInfo", Erl)
    Call incomingData.SafeClearPacket

End Sub

''
' Handles the GuildDetails message.

Private Sub HandleGuildDetails()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    With frmGuildBrief

        If Not .EsLeader Then

        End If
        
        .nombre.Caption = "Nombre:" & incomingData.ReadASCIIString()
        .fundador.Caption = "Fundador:" & incomingData.ReadASCIIString()
        .creacion.Caption = "Fecha de creacion:" & incomingData.ReadASCIIString()
        .lider.Caption = "Líder:" & incomingData.ReadASCIIString()
        .Miembros.Caption = "Miembros:" & incomingData.ReadInteger()
        
        .lblAlineacion.Caption = "Alineación: " & incomingData.ReadASCIIString()
        
        .desc.Text = incomingData.ReadASCIIString()
        .nivel.Caption = "Nivel de clan: " & incomingData.ReadByte()

    End With
    
    frmGuildBrief.Show vbModeless, frmMain
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildDetails", Erl)
    Call incomingData.SafeClearPacket

End Sub

''
' Handles the ShowGuildFundationForm message.

Private Sub HandleShowGuildFundationForm()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleShowGuildFundationForm_Err
    
    CreandoClan = True
    frmGuildDetails.Show , frmMain
    
    Exit Sub

HandleShowGuildFundationForm_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowGuildFundationForm", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the ParalizeOK message.

Private Sub HandleParalizeOK()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleParalizeOK_Err
    
    UserParalizado = Not UserParalizado
    
    Exit Sub

HandleParalizeOK_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleParalizeOK", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleInmovilizadoOK()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleInmovilizadoOK_Err
    
    UserInmovilizado = Not UserInmovilizado
    
    Exit Sub

HandleInmovilizadoOK_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleInmovilizadoOK", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the ShowUserRequest message.

Private Sub HandleShowUserRequest()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Call frmUserRequest.recievePeticion(incomingData.ReadASCIIString())
    Call frmUserRequest.Show(vbModeless, frmMain)
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowUserRequest", Erl)
    Call incomingData.SafeClearPacket

End Sub

''
' Handles the ChangeUserTradeSlot message.

Private Sub HandleChangeUserTradeSlot()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim miOferta As Boolean
    
    miOferta = incomingData.ReadBoolean
    Dim i          As Byte
    Dim nombreItem As String
    Dim cantidad   As Integer
    Dim grhItem    As Long
    Dim OBJIndex   As Integer

    If miOferta Then
        Dim OroAEnviar As Long
        OroAEnviar = incomingData.ReadLong
        frmComerciarUsu.lblOroMiOferta.Caption = PonerPuntos(OroAEnviar)
        frmComerciarUsu.lblMyGold.Caption = PonerPuntos(Val(frmMain.GldLbl.Caption - OroAEnviar))

        For i = 1 To 6

            With OtroInventario(i)
                OBJIndex = incomingData.ReadInteger
                nombreItem = incomingData.ReadASCIIString
                grhItem = incomingData.ReadLong
                cantidad = incomingData.ReadLong

                If cantidad > 0 Then
                    Call frmComerciarUsu.InvUserSell.SetItem(i, OBJIndex, cantidad, 0, grhItem, 0, 0, 0, 0, 0, nombreItem, 0)

                End If

            End With

        Next i
        
        Call frmComerciarUsu.InvUserSell.ReDraw
    Else
        frmComerciarUsu.lblOro.Caption = PonerPuntos(incomingData.ReadLong)

        ' frmComerciarUsu.List2.Clear
        For i = 1 To 6
            
            With OtroInventario(i)
                OBJIndex = incomingData.ReadInteger
                nombreItem = incomingData.ReadASCIIString
                grhItem = incomingData.ReadLong
                cantidad = incomingData.ReadLong

                If cantidad > 0 Then
                    Call frmComerciarUsu.InvOtherSell.SetItem(i, OBJIndex, cantidad, 0, grhItem, 0, 0, 0, 0, 0, nombreItem, 0)

                End If

            End With

        Next i
        
        Call frmComerciarUsu.InvOtherSell.ReDraw
    
    End If
    
    frmComerciarUsu.lblEstadoResp.Visible = False
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChangeUserTradeSlot", Erl)
    Call incomingData.SafeClearPacket

End Sub

''
' Handles the SpawnList message.

Private Sub HandleSpawnList()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim creatureList() As String

    creatureList = Split(incomingData.ReadASCIIString(), SEPARATOR)

    Call frmSpawnList.FillList

    frmSpawnList.Show , frmMain
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSpawnList", Erl)
    Call incomingData.SafeClearPacket

End Sub

''
' Handles the ShowSOSForm message.

Private Sub HandleShowSOSForm()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim sosList()      As String

    Dim i              As Long

    Dim nombre         As String

    Dim Consulta       As String

    Dim TipoDeConsulta As String
    
    sosList = Split(incomingData.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To UBound(sosList())
        nombre = ReadField(1, sosList(i), Asc("Ø"))
        Consulta = ReadField(2, sosList(i), Asc("Ø"))
        TipoDeConsulta = ReadField(3, sosList(i), Asc("Ø"))
        frmPanelgm.List1.AddItem nombre & "(" & TipoDeConsulta & ")"
        frmPanelgm.List2.AddItem Consulta
    Next i
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowSOSForm", Erl)
    Call incomingData.SafeClearPacket

End Sub

''
' Handles the ShowMOTDEditionForm message.

Private Sub HandleShowMOTDEditionForm()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    frmCambiaMotd.txtMotd.Text = incomingData.ReadASCIIString()
    frmCambiaMotd.Show , frmMain
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowMOTDEditionForm", Erl)
    Call incomingData.SafeClearPacket

End Sub

''
' Handles the ShowGMPanelForm message.

Private Sub HandleShowGMPanelForm()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleShowGMPanelForm_Err
    
    frmPanelgm.txtHeadNumero = incomingData.ReadInteger
    frmPanelgm.txtBodyYo = incomingData.ReadInteger
    frmPanelgm.txtCasco = incomingData.ReadInteger
    frmPanelgm.txtArma = incomingData.ReadInteger
    frmPanelgm.txtEscudo = incomingData.ReadInteger
    frmPanelgm.Show vbModeless, frmMain
    
    Exit Sub

HandleShowGMPanelForm_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowGMPanelForm", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleShowFundarClanForm()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleShowFundarClanForm_Err
    
    CreandoClan = True
    frmGuildDetails.Show vbModeless, frmMain
    
    Exit Sub

HandleShowFundarClanForm_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowFundarClanForm", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the UserNameList message.

Private Sub HandleUserNameList()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim userList() As String

    Dim i          As Long
    
    userList = Split(incomingData.ReadASCIIString(), SEPARATOR)
    
    If frmPanelgm.Visible Then
        frmPanelgm.cboListaUsus.Clear

        For i = 0 To UBound(userList())
            Call frmPanelgm.cboListaUsus.AddItem(userList(i))
        Next i

        If frmPanelgm.cboListaUsus.ListCount > 0 Then frmPanelgm.cboListaUsus.ListIndex = 0

    End If
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUserNameList", Erl)
    Call incomingData.SafeClearPacket

End Sub

''
' Handles the Pong message.

Private Sub HandlePong()
    
    On Error GoTo HandlePong_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim Time As Long
    Time = incomingData.ReadLong()

    PingRender = (timeGetTime And &H7FFFFFFF) - Time
    pingTime = 0
    
    Exit Sub

HandlePong_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePong", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Handles the UpdateTag message.

Private Sub HandleUpdateTagAndStatus()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim charindex   As Integer

    Dim status      As Byte

    Dim NombreYClan As String

    Dim group_index As Integer
    
    charindex = incomingData.ReadInteger()
    status = incomingData.ReadByte()
    NombreYClan = incomingData.ReadASCIIString()
        
    Dim pos As Integer
    pos = InStr(NombreYClan, "<")

    If pos = 0 Then pos = InStr(NombreYClan, "[")
    If pos = 0 Then pos = Len(NombreYClan) + 2
    
    charlist(charindex).nombre = Left$(NombreYClan, pos - 2)
    charlist(charindex).clan = mid$(NombreYClan, pos)
    
    group_index = incomingData.ReadInteger()
    
    'Update char status adn tag!
    charlist(charindex).status = status
    
    charlist(charindex).group_index = group_index
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateTagAndStatus", Erl)
    Call incomingData.SafeClearPacket

End Sub

''
' Flushes the outgoing data buffer of the user.
'
' @param    UserIndex User whose outgoing data buffer will be flushed.

Public Sub FlushBuffer()
    
    On Error GoTo FlushBuffer_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Sends all data existing in the buffer
    '***************************************************

    With outgoingData

        If .Length = 0 Then Exit Sub

        OutBytes = OutBytes + .Length

        Call SendData(.ReadAll)

    End With
    
    Exit Sub

FlushBuffer_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.FlushBuffer", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Sends the data using the socket controls in the MainForm.
'
' @param    Data  The data to be sent to the server.

Private Sub SendData(ByRef Data() As Byte)
    
    On Error GoTo SendData_Err

    If frmMain.MainSocket.State <> sckConnected Then Exit Sub

    #If AntiExternos = 1 Then
        Call Security.XorData(Data, UBound(Data), XorIndexOut)
    #End If
 
    Call frmMain.MainSocket.SendData(Data)
    
    Exit Sub

SendData_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.SendData", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandlePersonajesDeCuenta()

    On Error GoTo ErrHandler
    
    CantidadDePersonajesEnCuenta = incomingData.ReadByte()

    Dim ii As Byte
     
    For ii = 1 To 10
        Pjs(ii).Body = 0
        Pjs(ii).Head = 0
        Pjs(ii).Mapa = 0
        Pjs(ii).PosX = 0
        Pjs(ii).PosY = 0
        Pjs(ii).nivel = 0
        Pjs(ii).nombre = ""
        Pjs(ii).Criminal = 0
        Pjs(ii).Clase = 0
        Pjs(ii).NameMapa = ""
        Pjs(ii).Casco = 0
        Pjs(ii).Escudo = 0
        Pjs(ii).Arma = 0
        Pjs(ii).ClanName = ""
    Next ii

    For ii = 1 To CantidadDePersonajesEnCuenta
        Pjs(ii).nombre = incomingData.ReadASCIIString()
        Pjs(ii).nivel = incomingData.ReadByte()
        Pjs(ii).Mapa = incomingData.ReadInteger()
        Pjs(ii).PosX = incomingData.ReadInteger()
        Pjs(ii).PosY = incomingData.ReadInteger()
        
        Pjs(ii).Body = incomingData.ReadInteger()
        
        Pjs(ii).Head = incomingData.ReadInteger()
        Pjs(ii).Criminal = incomingData.ReadByte()
        Pjs(ii).Clase = incomingData.ReadByte()
       
        Pjs(ii).Casco = incomingData.ReadInteger()
        Pjs(ii).Escudo = incomingData.ReadInteger()
        Pjs(ii).Arma = incomingData.ReadInteger()
        Pjs(ii).ClanName = "<" & incomingData.ReadASCIIString() & ">"
       
        ' Pjs(ii).NameMapa = Pjs(ii).mapa
        Pjs(ii).NameMapa = NameMaps(Pjs(ii).Mapa).Name

    Next ii
    
    CuentaDonador = incomingData.ReadByte()
    
    Dim i As Integer

    For i = 1 To CantidadDePersonajesEnCuenta

        Select Case Pjs(i).Criminal

            Case 0 'Criminal
                Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(50).r, ColoresPJ(50).G, ColoresPJ(50).B)
                Pjs(i).priv = 0

            Case 1 'Ciudadano
                Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(49).r, ColoresPJ(49).G, ColoresPJ(49).B)
                Pjs(i).priv = 0

            Case 2 'Caos
                Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(6).r, ColoresPJ(6).G, ColoresPJ(6).B)
                Pjs(i).priv = 0

            Case 3 'Armada
                Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(8).r, ColoresPJ(8).G, ColoresPJ(8).B)
                Pjs(i).priv = 0

            Case 4 'EsConsejero
                Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(1).r, ColoresPJ(1).G, ColoresPJ(1).B)
                Pjs(i).ClanName = "<Game Master>"
                Pjs(i).priv = 1
                EsGM = True

            Case 5 ' EsSemiDios
                Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(2).r, ColoresPJ(2).G, ColoresPJ(2).B)
                Pjs(i).ClanName = "<Game Master>"
                Pjs(i).priv = 2
                EsGM = True

            Case 6 ' EsDios
                Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(3).r, ColoresPJ(3).G, ColoresPJ(3).B)
                Pjs(i).ClanName = "<Game Master>"
                Pjs(i).priv = 3
                EsGM = True

            Case 7 ' EsAdmin
                Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(4).r, ColoresPJ(4).G, ColoresPJ(4).B)
                Pjs(i).ClanName = "<Game Master>"
                Pjs(i).priv = 4
                EsGM = True

            Case Else

        End Select

    Next i
    
    AlphaRenderCuenta = MAX_ALPHA_RENDER_CUENTA
   
    If CantidadDePersonajesEnCuenta > 0 Then
        PJSeleccionado = 1
        LastPJSeleccionado = 1
        
        If Pjs(1).Mapa <> 0 Then
            Call SwitchMap(Pjs(1).Mapa)
            RenderCuenta_PosX = Pjs(1).PosX
            RenderCuenta_PosY = Pjs(1).PosY
        End If
    End If
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePersonajesDeCuenta", Erl)
    Call incomingData.SafeClearPacket

End Sub

Private Sub HandleUserOnline()
    
    On Error GoTo ErrHandler

    Dim rdata As Integer
    
    rdata = incomingData.ReadInteger()
    
    usersOnline = rdata
    frmMain.onlines = "Online: " & usersOnline
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUserOnline", Erl)
    Call incomingData.SafeClearPacket

End Sub

Private Sub HandleParticleFXToFloor()
    
    On Error GoTo HandleParticleFXToFloor_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim x              As Byte

    Dim y              As Byte

    Dim ParticulaIndex As Byte

    Dim Time           As Long

    Dim Borrar         As Boolean
     
    x = incomingData.ReadByte()
    y = incomingData.ReadByte()
    ParticulaIndex = incomingData.ReadInteger()
    Time = incomingData.ReadLong()

    If Time = 1 Then
        Time = -1

    End If
    
    If Time = 0 Then
        Borrar = True

    End If

    If Borrar Then
        Graficos_Particulas.Particle_Group_Remove (MapData(x, y).particle_group)
    Else

        If MapData(x, y).particle_group = 0 Then
            MapData(x, y).particle_group = 0
            General_Particle_Create ParticulaIndex, x, y, Time
        Else
            Call General_Char_Particle_Create(ParticulaIndex, MapData(x, y).charindex, Time)

        End If

    End If
    
    Exit Sub

HandleParticleFXToFloor_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleParticleFXToFloor", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleLightToFloor()
    
    On Error GoTo HandleLightToFloor_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim x           As Byte

    Dim y           As Byte

    Dim Color       As Long
    
    Dim color_value As RGBA

    Dim Rango       As Byte
     
    x = incomingData.ReadByte()
    y = incomingData.ReadByte()
    Color = incomingData.ReadLong()
    Rango = incomingData.ReadByte()
    
    Call Long_2_RGBA(color_value, Color)

    Dim ID  As Long

    Dim id2 As Long

    If Color = 0 Then
   
        If MapData(x, y).luz.Rango > 100 Then
            LucesRedondas.Delete_Light_To_Map x, y
   
            LucesCuadradas.Light_Render_All
            LucesRedondas.LightRenderAll
            Exit Sub
        Else
            ID = LucesCuadradas.Light_Find(x & y)
            LucesCuadradas.Light_Remove ID
            MapData(x, y).luz.Color = COLOR_EMPTY
            MapData(x, y).luz.Rango = 0
            LucesCuadradas.Light_Render_All
            Exit Sub

        End If

    End If
    
    MapData(x, y).luz.Color = color_value
    MapData(x, y).luz.Rango = Rango
    
    If Rango < 100 Then
        ID = x & y
        LucesCuadradas.Light_Create x, y, color_value, Rango, ID
        LucesRedondas.LightRenderAll
        LucesCuadradas.Light_Render_All
    Else

        LucesRedondas.Create_Light_To_Map x, y, color_value, Rango - 99
        LucesRedondas.LightRenderAll
        LucesCuadradas.Light_Render_All

    End If
    
    Exit Sub

HandleLightToFloor_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleLightToFloor", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleParticleFX()
    
    On Error GoTo HandleParticleFX_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim charindex      As Integer

    Dim ParticulaIndex As Integer

    Dim Time           As Long

    Dim Remove         As Boolean
    Dim grh            As Long
     
    charindex = incomingData.ReadInteger()
    ParticulaIndex = incomingData.ReadInteger()
    Time = incomingData.ReadLong()
    Remove = incomingData.ReadBoolean()
    grh = incomingData.ReadLong()
    
    If Remove Then
        Call Char_Particle_Group_Remove(charindex, ParticulaIndex)
        charlist(charindex).Particula = 0
    
    Else
        charlist(charindex).Particula = ParticulaIndex
        charlist(charindex).ParticulaTime = Time
        If grh > 0 Then
            Call General_Char_Particle_Create(ParticulaIndex, charindex, Time, grh)
        Else
            Call General_Char_Particle_Create(ParticulaIndex, charindex, Time)
        End If

    End If
    
    Exit Sub

HandleParticleFX_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleParticleFX", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleParticleFXWithDestino()
    
    On Error GoTo HandleParticleFXWithDestino_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim Emisor         As Integer

    Dim receptor       As Integer

    Dim ParticulaViaje As Integer

    Dim ParticulaFinal As Integer

    Dim Time           As Long

    Dim wav            As Integer

    Dim fX             As Integer
     
    Emisor = incomingData.ReadInteger()
    receptor = incomingData.ReadInteger()
    ParticulaViaje = incomingData.ReadInteger()
    ParticulaFinal = incomingData.ReadInteger()

    Time = incomingData.ReadLong()
    wav = incomingData.ReadInteger()
    fX = incomingData.ReadInteger()

    Engine_spell_Particle_Set (ParticulaViaje)

    Call Effect_Begin(ParticulaViaje, 9, Get_Pixelx_Of_Char(Emisor), Get_PixelY_Of_Char(Emisor), ParticulaFinal, Time, receptor, Emisor, wav, fX)

    ' charlist(charindex).Particula = ParticulaIndex
    ' charlist(charindex).ParticulaTime = time

    ' Call General_Char_Particle_Create(ParticulaIndex, charindex, time)
    
    Exit Sub

HandleParticleFXWithDestino_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleParticleFXWithDestino", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleParticleFXWithDestinoXY()
    
    On Error GoTo HandleParticleFXWithDestinoXY_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim Emisor         As Integer

    Dim ParticulaViaje As Integer

    Dim ParticulaFinal As Integer

    Dim Time           As Long

    Dim wav            As Integer

    Dim fX             As Integer

    Dim x              As Byte

    Dim y              As Byte
     
    Emisor = incomingData.ReadInteger()
    ParticulaViaje = incomingData.ReadInteger()
    ParticulaFinal = incomingData.ReadInteger()

    Time = incomingData.ReadLong()
    wav = incomingData.ReadInteger()
    fX = incomingData.ReadInteger()
    
    x = incomingData.ReadByte()
    y = incomingData.ReadByte()
    
    ' Debug.Print "RECIBI FX= " & fX

    Engine_spell_Particle_Set (ParticulaViaje)

    Call Effect_BeginXY(ParticulaViaje, 9, Get_Pixelx_Of_Char(Emisor), Get_PixelY_Of_Char(Emisor), x, y, ParticulaFinal, Time, Emisor, wav, fX)

    ' charlist(charindex).Particula = ParticulaIndex
    ' charlist(charindex).ParticulaTime = time

    ' Call General_Char_Particle_Create(ParticulaIndex, charindex, time)
    
    Exit Sub

HandleParticleFXWithDestinoXY_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleParticleFXWithDestinoXY", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleAuraToChar()
    
    On Error GoTo HandleAuraToChar_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/0
    '
    '***************************************************
    
    Dim charindex      As Integer

    Dim ParticulaIndex As String

    Dim Remove         As Boolean

    Dim TIPO           As Byte
     
    charindex = incomingData.ReadInteger()
    ParticulaIndex = incomingData.ReadASCIIString()

    Remove = incomingData.ReadBoolean()
    TIPO = incomingData.ReadByte()
    
    If TIPO = 1 Then
        charlist(charindex).Arma_Aura = ParticulaIndex
    ElseIf TIPO = 2 Then
        charlist(charindex).Body_Aura = ParticulaIndex
    ElseIf TIPO = 3 Then
        charlist(charindex).Escudo_Aura = ParticulaIndex
    ElseIf TIPO = 4 Then
        charlist(charindex).Head_Aura = ParticulaIndex
    ElseIf TIPO = 5 Then
        charlist(charindex).Otra_Aura = ParticulaIndex
    ElseIf TIPO = 6 Then
        charlist(charindex).DM_Aura = ParticulaIndex
    Else
        charlist(charindex).RM_Aura = ParticulaIndex

    End If
    
    Exit Sub

HandleAuraToChar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleAuraToChar", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleSpeedToChar()
    
    On Error GoTo HandleSpeedToChar_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/0
    '
    '***************************************************
    
    Dim charindex As Integer

    Dim Speeding  As Single
     
    charindex = incomingData.ReadInteger()
    Speeding = incomingData.ReadSingle()
   
    charlist(charindex).Speeding = Speeding
    
    Exit Sub

HandleSpeedToChar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSpeedToChar", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleNieveToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleNieveToggle_Err
    
    If Not InMapBounds(UserPos.x, UserPos.y) Then Exit Sub
            
    If MapDat.NIEVE Then
        Engine_MeteoParticle_Set (Particula_Nieve)

    End If

    bNieve = Not bNieve
    
    Exit Sub

HandleNieveToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleNieveToggle", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleNieblaToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleNieblaToggle_Err
    
    MaxAlphaNiebla = incomingData.ReadByte()
            
    bNiebla = Not bNiebla
    frmMain.TimerNiebla.Enabled = True
    
    Exit Sub

HandleNieblaToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleNieblaToggle", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleFamiliar()
    
    On Error GoTo HandleFamiliar_Err
    
    Exit Sub

HandleFamiliar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleFamiliar", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleBindKeys()
    
    On Error GoTo HandleBindKeys_Err

    '***************************************************
    'Macros
    'Pablo Mercavides
    '***************************************************
    
    ChatCombate = incomingData.ReadByte()
    ChatGlobal = incomingData.ReadByte()

    If ChatCombate = 1 Then
        frmMain.CombateIcon.Picture = LoadInterface("infoapretado.bmp")
    Else
        frmMain.CombateIcon.Picture = LoadInterface("info.bmp")

    End If

    If ChatGlobal = 1 Then
        frmMain.globalIcon.Picture = LoadInterface("globalapretado.bmp")
    Else
        frmMain.CombateIcon.Picture = LoadInterface("global.bmp")

    End If
    
    Exit Sub

HandleBindKeys_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBindKeys", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleLogros()
    
    On Error GoTo HandleLogros_Err

    '***************************************************
    'Pablo Mercavides
    '***************************************************
    
    NPcLogros.nombre = incomingData.ReadASCIIString()
    NPcLogros.desc = incomingData.ReadASCIIString()
    NPcLogros.cant = incomingData.ReadInteger()
    NPcLogros.TipoRecompensa = incomingData.ReadByte()

    If NPcLogros.TipoRecompensa = 1 Then
        NPcLogros.ObjRecompensa = incomingData.ReadASCIIString()

    End If
    
    If NPcLogros.TipoRecompensa = 2 Then
    
        NPcLogros.OroRecompensa = incomingData.ReadLong()

    End If
    
    If NPcLogros.TipoRecompensa = 3 Then
        NPcLogros.ExpRecompensa = incomingData.ReadLong()

    End If
    
    If NPcLogros.TipoRecompensa = 4 Then
        NPcLogros.HechizoRecompensa = incomingData.ReadByte()

    End If
    
    NPcLogros.NpcsMatados = incomingData.ReadInteger()
    
    NPcLogros.Finalizada = incomingData.ReadBoolean()
    
    UserLogros.nombre = incomingData.ReadASCIIString()
    UserLogros.desc = incomingData.ReadASCIIString()
    UserLogros.cant = incomingData.ReadInteger()
    UserLogros.TipoRecompensa = incomingData.ReadInteger()
    UserLogros.UserMatados = incomingData.ReadInteger()
    
    If UserLogros.TipoRecompensa = 1 Then
        UserLogros.ObjRecompensa = incomingData.ReadASCIIString()

    End If
    
    If UserLogros.TipoRecompensa = 2 Then
    
        UserLogros.OroRecompensa = incomingData.ReadLong()

    End If
    
    If UserLogros.TipoRecompensa = 3 Then
        UserLogros.ExpRecompensa = incomingData.ReadLong()

    End If
    
    If UserLogros.TipoRecompensa = 4 Then
        UserLogros.HechizoRecompensa = incomingData.ReadByte()

    End If
    
    UserLogros.Finalizada = incomingData.ReadBoolean()
        
    LevelLogros.nombre = incomingData.ReadASCIIString()
    LevelLogros.desc = incomingData.ReadASCIIString()
    LevelLogros.cant = incomingData.ReadInteger()
    LevelLogros.TipoRecompensa = incomingData.ReadInteger()
    LevelLogros.NivelUser = incomingData.ReadByte()
    
    If LevelLogros.TipoRecompensa = 1 Then
        LevelLogros.ObjRecompensa = incomingData.ReadASCIIString()

    End If
    
    If LevelLogros.TipoRecompensa = 2 Then
    
        LevelLogros.OroRecompensa = incomingData.ReadLong()

    End If
    
    If LevelLogros.TipoRecompensa = 3 Then
        LevelLogros.ExpRecompensa = incomingData.ReadLong()

    End If

    If LevelLogros.TipoRecompensa = 4 Then
        LevelLogros.HechizoRecompensa = incomingData.ReadByte()

    End If
    
    LevelLogros.Finalizada = incomingData.ReadBoolean()
    
    If FrmLogros.Visible Then
        Unload FrmLogros

    End If
    
    FrmLogros.Show , frmMain
    
    Exit Sub

HandleLogros_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleLogros", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleBarFx()
    
    On Error GoTo HandleBarFx_Err

    '***************************************************
    'Author: Pablo Mercavides
    '***************************************************
    
    Dim charindex As Integer

    Dim BarTime   As Integer

    Dim BarAccion As Byte
    
    charindex = incomingData.ReadInteger()
    BarTime = incomingData.ReadInteger()
    BarAccion = incomingData.ReadByte()
    
    charlist(charindex).BarTime = 0
    charlist(charindex).BarAccion = BarAccion
    charlist(charindex).MaxBarTime = BarTime / engineBaseSpeed
    
    Exit Sub

HandleBarFx_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBarFx", Erl)
    Call incomingData.SafeClearPacket
    
End Sub
 
Private Sub HandleQuestDetails()

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Recibe y maneja el paquete QuestDetails del servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    
    On Error GoTo ErrHandler
    
    Dim tmpStr         As String

    Dim tmpByte        As Byte

    Dim QuestEmpezada  As Boolean

    Dim i              As Integer
    
    Dim cantidadnpc    As Integer

    Dim NpcIndex       As Integer
    
    Dim cantidadobj    As Integer

    Dim OBJIndex       As Integer
    
    Dim AmountHave     As Integer
    
    Dim QuestIndex     As Integer
    
    Dim LevelRequerido As Byte
    Dim QuestRequerida As Integer
    
    FrmQuests.ListView2.ListItems.Clear
    FrmQuests.ListView1.ListItems.Clear
    FrmQuestInfo.ListView2.ListItems.Clear
    FrmQuestInfo.ListView1.ListItems.Clear
    
    FrmQuests.PlayerView.BackColor = RGB(11, 11, 11)
    FrmQuests.picture1.BackColor = RGB(19, 14, 11)
    FrmQuests.PlayerView.Refresh
    FrmQuests.picture1.Refresh
    FrmQuests.npclbl.Caption = ""
    FrmQuests.objetolbl.Caption = ""
    
    With incomingData
        'Nos fijamos si se trata de una quest empezada, para poder leer los NPCs que se han matado.
        QuestEmpezada = IIf(.ReadByte, True, False)
        
        If Not QuestEmpezada Then
        
            QuestIndex = .ReadInteger
        
            FrmQuestInfo.titulo.Caption = QuestList(QuestIndex).nombre
           
            'tmpStr = "Mision: " & .ReadASCIIString & vbCrLf
            
            LevelRequerido = .ReadByte
            QuestRequerida = .ReadInteger
           
            If QuestRequerida <> 0 Then
                FrmQuestInfo.Text1.Text = QuestList(QuestIndex).desc & vbCrLf & vbCrLf & "Requisitos" & vbCrLf & "Nivel requerido: " & LevelRequerido & vbCrLf & "Quest:" & QuestList(QuestRequerida).RequiredQuest
            Else
            
                FrmQuestInfo.Text1.Text = QuestList(QuestIndex).desc & vbCrLf & vbCrLf & "Requisitos" & vbCrLf & "Nivel requerido: " & LevelRequerido & vbCrLf
            
            End If
           
            tmpByte = .ReadByte

            If tmpByte Then 'Hay NPCs
                If tmpByte > 5 Then
                    FrmQuestInfo.ListView1.FlatScrollBar = False
                Else
                    FrmQuestInfo.ListView1.FlatScrollBar = True
           
                End If

                For i = 1 To tmpByte
                    cantidadnpc = .ReadInteger
                    NpcIndex = .ReadInteger
               
                    ' tmpStr = tmpStr & "*) Matar " & .ReadInteger & " " & .ReadASCIIString & "."
                    If QuestEmpezada Then
                        tmpStr = tmpStr & " (Has matado " & .ReadInteger & ")" & vbCrLf
                    Else
                        tmpStr = tmpStr & vbCrLf
                       
                        Dim subelemento As ListItem

                        Set subelemento = FrmQuestInfo.ListView1.ListItems.Add(, , NpcData(NpcIndex).Name)
                       
                        subelemento.SubItems(1) = cantidadnpc
                        subelemento.SubItems(2) = NpcIndex
                        subelemento.SubItems(3) = 0

                    End If

                Next i

            End If
           
            tmpByte = .ReadByte

            If tmpByte Then 'Hay OBJs

                For i = 1 To tmpByte
               
                    cantidadobj = .ReadInteger
                    OBJIndex = .ReadInteger
                    
                    AmountHave = .ReadInteger
                   
                    Set subelemento = FrmQuestInfo.ListView1.ListItems.Add(, , ObjData(OBJIndex).Name)
                    subelemento.SubItems(1) = AmountHave & "/" & cantidadobj
                    subelemento.SubItems(2) = OBJIndex
                    subelemento.SubItems(3) = 1
                Next i

            End If
    
            tmpStr = tmpStr & vbCrLf & "RECOMPENSAS" & vbCrLf
            'tmpStr = tmpStr & "*) Oro: " & .ReadLong & " monedas de oro." & vbCrLf
            'tmpStr = tmpStr & "*) Experiencia: " & .ReadLong & " puntos de experiencia." & vbCrLf
           
            Set subelemento = FrmQuestInfo.ListView2.ListItems.Add(, , "Oro")

            subelemento.SubItems(1) = BeautifyBigNumber(.ReadLong)
            subelemento.SubItems(2) = 12
            subelemento.SubItems(3) = 0

            Set subelemento = FrmQuestInfo.ListView2.ListItems.Add(, , "Experiencia")

            subelemento.SubItems(1) = BeautifyBigNumber(.ReadLong)
            subelemento.SubItems(2) = 608
            subelemento.SubItems(3) = 1
           
            tmpByte = .ReadByte

            If tmpByte Then

                For i = 1 To tmpByte
                    'tmpStr = tmpStr & "*) " & .ReadInteger & " " & .ReadInteger & vbCrLf
                   
                    Dim cantidadobjs As Integer

                    Dim obindex      As Integer
                   
                    cantidadobjs = .ReadInteger
                    obindex = .ReadInteger
                   
                    Set subelemento = FrmQuestInfo.ListView2.ListItems.Add(, , ObjData(obindex).Name)
                       
                    subelemento.SubItems(1) = cantidadobjs
                    subelemento.SubItems(2) = obindex
                    subelemento.SubItems(3) = 1

           
                Next i

            End If

        Else
        
            QuestIndex = .ReadInteger
        
            FrmQuests.titulo.Caption = QuestList(QuestIndex).nombre
           
            LevelRequerido = .ReadByte
            QuestRequerida = .ReadInteger
           
            FrmQuests.detalle.Text = QuestList(QuestIndex).desc & vbCrLf & vbCrLf & "Requisitos" & vbCrLf & "Nivel requerido: " & LevelRequerido & vbCrLf

            If QuestRequerida <> 0 Then
                FrmQuests.detalle.Text = FrmQuests.detalle.Text & vbCrLf & "Quest: " & QuestList(QuestRequerida).nombre

            End If

           
            tmpStr = tmpStr & vbCrLf & "OBJETIVOS" & vbCrLf
           
            tmpByte = .ReadByte

            If tmpByte Then 'Hay NPCs

                For i = 1 To tmpByte
                    cantidadnpc = .ReadInteger
                    NpcIndex = .ReadInteger
               
                    Dim matados As Integer
               
                    matados = .ReadInteger
                                     
                    Set subelemento = FrmQuests.ListView1.ListItems.Add(, , NpcData(NpcIndex).Name)
                       
                    Dim cantok As Integer

                    cantok = cantidadnpc - matados
                       
                    If cantok = 0 Then
                        subelemento.SubItems(1) = "OK"
                    Else
                        subelemento.SubItems(1) = matados & "/" & cantidadnpc

                    End If
                        
                    ' subelemento.SubItems(1) = cantidadnpc - matados
                    subelemento.SubItems(2) = NpcIndex
                    subelemento.SubItems(3) = 0
                    'End If
                Next i

            End If
           
            tmpByte = .ReadByte

            If tmpByte Then 'Hay OBJs

                For i = 1 To tmpByte
               
                    cantidadobj = .ReadInteger
                    OBJIndex = .ReadInteger
                    
                    AmountHave = .ReadInteger
                   
                    Set subelemento = FrmQuests.ListView1.ListItems.Add(, , ObjData(OBJIndex).Name)
                    subelemento.SubItems(1) = AmountHave & "/" & cantidadobj
                    subelemento.SubItems(2) = OBJIndex
                    subelemento.SubItems(3) = 1
                Next i

            End If
    
            tmpStr = tmpStr & vbCrLf & "RECOMPENSAS" & vbCrLf

            Dim tmplong As Long
           
            tmplong = .ReadLong
           
            If tmplong <> 0 Then
                Set subelemento = FrmQuests.ListView2.ListItems.Add(, , "Oro")
                subelemento.SubItems(1) = BeautifyBigNumber(tmplong)
                subelemento.SubItems(2) = 12
                subelemento.SubItems(3) = 0

            End If
            
            tmplong = .ReadLong
           
            If tmplong <> 0 Then
                Set subelemento = FrmQuests.ListView2.ListItems.Add(, , "Experiencia")
                           
                subelemento.SubItems(1) = BeautifyBigNumber(tmplong)
                subelemento.SubItems(2) = 608
                subelemento.SubItems(3) = 1

            End If
           
            tmpByte = .ReadByte

            If tmpByte Then

                For i = 1 To tmpByte
                    cantidadobjs = .ReadInteger
                    obindex = .ReadInteger
                   
                    Set subelemento = FrmQuests.ListView2.ListItems.Add(, , ObjData(obindex).Name)
                       
                    subelemento.SubItems(1) = cantidadobjs
                    subelemento.SubItems(2) = obindex
                    subelemento.SubItems(3) = 1

           
                Next i

            End If
        
        End If

    End With
    
    'Determinamos que formulario se muestra, segï¿½n si recibimos la informaciï¿½n y la quest estï¿½ empezada o no.
    If QuestEmpezada Then
        FrmQuests.txtInfo.Text = tmpStr
        Call FrmQuests.ListView1_Click
        Call FrmQuests.ListView2_Click
        Call FrmQuests.lstQuests.SetFocus
    Else

        FrmQuestInfo.Show vbModeless, frmMain
        FrmQuestInfo.Picture = LoadInterface("ventananuevamision.bmp")
        Call FrmQuestInfo.ListView1_Click
        Call FrmQuestInfo.ListView2_Click

    End If
    
    Exit Sub
    
ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleQuestDetails", Erl)
    Call incomingData.SafeClearPacket

End Sub
 
Public Sub HandleQuestListSend()

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Recibe y maneja el paquete QuestListSend del servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    
    On Error GoTo ErrHandler
    
    Dim i       As Integer
    Dim tmpByte As Byte
    Dim tmpStr  As String
     
    'Leemos la cantidad de quests que tiene el usuario
    tmpByte = incomingData.ReadByte
    
    'Limpiamos el ListBox y el TextBox del formulario
    FrmQuests.lstQuests.Clear
    FrmQuests.txtInfo.Text = vbNullString
        
    'Si el usuario tiene quests entonces hacemos el handle
    If tmpByte Then
        'Leemos el string
        tmpStr = incomingData.ReadASCIIString
        
        'Agregamos los items
        For i = 1 To tmpByte
            FrmQuests.lstQuests.AddItem ReadField(i, tmpStr, 45)
        Next i

    End If
    
    'Mostramos el formulario
    
    COLOR_AZUL = RGB(0, 0, 0)
    Call Establecer_Borde(FrmQuests.lstQuests, FrmQuests, COLOR_AZUL, 0, 0)
    FrmQuests.Picture = LoadInterface("ventanadetallemision.bmp")
    FrmQuests.Show vbModeless, frmMain
    
    'Pedimos la informacion de la primer quest (si la hay)
    If tmpByte Then Call WriteQuestDetailsRequest(1)

    Exit Sub
    
ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleQuestListSend", Erl)
    Call incomingData.SafeClearPacket

End Sub

Public Sub HandleNpcQuestListSend()

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Recibe y maneja el paquete QuestListSend del servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    
    On Error GoTo ErrHandler

    Dim tmpStr         As String
    Dim tmpByte        As Byte
    Dim QuestEmpezada  As Boolean
    Dim i              As Integer
    Dim j              As Byte
    Dim cantidadnpc    As Integer
    Dim NpcIndex       As Integer
    Dim cantidadobj    As Integer
    Dim OBJIndex       As Integer
    Dim QuestIndex     As Integer
    Dim estado         As Byte
    Dim LevelRequerido As Byte
    Dim QuestRequerida As Integer
    Dim CantidadQuest  As Byte
    Dim subelemento    As ListItem
    
    FrmQuestInfo.ListView2.ListItems.Clear
    FrmQuestInfo.ListView1.ListItems.Clear
    
    With incomingData

        CantidadQuest = .ReadByte
            
        For j = 1 To CantidadQuest
        
            QuestIndex = .ReadInteger
            
            FrmQuestInfo.titulo.Caption = QuestList(QuestIndex).nombre
                              
            QuestList(QuestIndex).RequiredLevel = .ReadByte
            QuestList(QuestIndex).RequiredQuest = .ReadInteger
            
            tmpByte = .ReadByte
    
            If tmpByte Then 'Hay NPCs
            
                If tmpByte > 5 Then
                    FrmQuestInfo.ListView1.FlatScrollBar = False
                Else
                    FrmQuestInfo.ListView1.FlatScrollBar = True
               
                End If
                    
                ReDim QuestList(QuestIndex).RequiredNPC(1 To tmpByte)
                    
                For i = 1 To tmpByte
                                                
                    QuestList(QuestIndex).RequiredNPC(i).Amount = .ReadInteger
                    QuestList(QuestIndex).RequiredNPC(i).NpcIndex = .ReadInteger

                Next i

            Else
                ReDim QuestList(QuestIndex).RequiredNPC(0)

            End If
               
            tmpByte = .ReadByte
    
            If tmpByte Then 'Hay OBJs
                ReDim QuestList(QuestIndex).RequiredOBJ(1 To tmpByte)
    
                For i = 1 To tmpByte
                   
                    QuestList(QuestIndex).RequiredOBJ(i).Amount = .ReadInteger
                    QuestList(QuestIndex).RequiredOBJ(i).OBJIndex = .ReadInteger

                Next i

            Else
                ReDim QuestList(QuestIndex).RequiredOBJ(0)
    
            End If
               
            QuestList(QuestIndex).RewardGLD = .ReadLong
            QuestList(QuestIndex).RewardEXP = .ReadLong

            tmpByte = .ReadByte
    
            If tmpByte Then
                
                ReDim QuestList(QuestIndex).RewardOBJ(1 To tmpByte)
    
                For i = 1 To tmpByte
                                              
                    QuestList(QuestIndex).RewardOBJ(i).Amount = .ReadInteger
                    QuestList(QuestIndex).RewardOBJ(i).OBJIndex = .ReadInteger
               
                Next i

            Else
                ReDim QuestList(QuestIndex).RewardOBJ(0)
    
            End If
                
            estado = .ReadByte
                
            Select Case estado
                
                Case 0
                    Set subelemento = FrmQuestInfo.ListViewQuest.ListItems.Add(, , QuestList(QuestIndex).nombre)
                        
                    subelemento.SubItems(1) = "Disponible"
                    subelemento.SubItems(2) = QuestIndex
                    subelemento.ForeColor = vbWhite
                    subelemento.ListSubItems(1).ForeColor = vbWhite

                Case 1
                    Set subelemento = FrmQuestInfo.ListViewQuest.ListItems.Add(, , QuestList(QuestIndex).nombre)
                        
                    subelemento.SubItems(1) = "En Curso"
                    subelemento.ForeColor = RGB(255, 175, 10)
                    subelemento.SubItems(2) = QuestIndex
                    subelemento.ListSubItems(1).ForeColor = RGB(255, 175, 10)
                    FrmQuestInfo.ListViewQuest.Refresh

                Case 2
                    Set subelemento = FrmQuestInfo.ListViewQuest.ListItems.Add(, , QuestList(QuestIndex).nombre)
                        
                    subelemento.SubItems(1) = "Finalizada"
                    subelemento.SubItems(2) = QuestIndex
                    subelemento.ForeColor = RGB(15, 140, 50)
                    subelemento.ListSubItems(1).ForeColor = RGB(15, 140, 50)
                    FrmQuestInfo.ListViewQuest.Refresh

                Case 3
                    Set subelemento = FrmQuestInfo.ListViewQuest.ListItems.Add(, , QuestList(QuestIndex).nombre)
                        
                    subelemento.SubItems(1) = "No disponible"
                    subelemento.SubItems(2) = QuestIndex
                    subelemento.ForeColor = RGB(255, 10, 10)
                    subelemento.ListSubItems(1).ForeColor = RGB(255, 10, 10)
                    FrmQuestInfo.ListViewQuest.Refresh
                
            End Select
                
        Next j

    End With
    
    'Determinamos que formulario se muestra, segun si recibimos la informacion y la quest estï¿½ empezada o no.
    FrmQuestInfo.Show vbModeless, frmMain
    FrmQuestInfo.Picture = LoadInterface("ventananuevamision.bmp")
    Call FrmQuestInfo.ShowQuest(1)
    
    Exit Sub
    
ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleNpcQuestListSend", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleListaCorreo()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim cant       As Byte
    Dim i          As Byte
    Dim Actualizar As Boolean

    cant = incomingData.ReadByte()
    
    FrmCorreo.lstMsg.Clear
    FrmCorreo.ListAdjuntos.Clear
    FrmCorreo.txMensaje.Text = vbNullString
    FrmCorreo.lbFecha.Caption = vbNullString
    FrmCorreo.lbItem.Caption = vbNullString

    If cant > 0 Then

        For i = 1 To cant
        
            CorreoMsj(i).Remitente = incomingData.ReadASCIIString()
            CorreoMsj(i).mensaje = incomingData.ReadASCIIString()
            CorreoMsj(i).ItemCount = incomingData.ReadByte()
            CorreoMsj(i).ItemArray = incomingData.ReadASCIIString()
            CorreoMsj(i).Leido = incomingData.ReadByte()
            CorreoMsj(i).Fecha = incomingData.ReadASCIIString()
            
            FrmCorreo.lstMsg.AddItem CorreoMsj(i).Remitente
            FrmCorreo.lstMsg.Enabled = True
            
            FrmCorreo.txMensaje.Enabled = True
        Next i

    Else
    
        FrmCorreo.lstMsg.AddItem ("Sin mensajes")
        FrmCorreo.txMensaje.Enabled = False

    End If
        
    Call FrmCorreo.lstInv.Clear

    'Fill the inventory list
    For i = 1 To MAX_INVENTORY_SLOTS

        If frmMain.Inventario.OBJIndex(i) <> 0 Then
            FrmCorreo.lstInv.AddItem frmMain.Inventario.ItemName(i)
            
        Else
            FrmCorreo.lstInv.AddItem "Vacio"

        End If

    Next i
    
    Actualizar = incomingData.ReadBoolean()

    ' FrmCorreo.lstMsg.AddItem
    If Not Actualizar Then
        FrmCorreo.Picture = LoadInterface("ventanacorreo.bmp")
        COLOR_AZUL = RGB(0, 0, 0)
        
        ' establece el borde al listbox
        Call Establecer_Borde(FrmCorreo.lstMsg, FrmCorreo, COLOR_AZUL, 0, 0)
        Call Establecer_Borde(FrmCorreo.ListAdjuntos, FrmCorreo, COLOR_AZUL, 0, 0)
        Call Establecer_Borde(FrmCorreo.ListaAenviar, FrmCorreo, COLOR_AZUL, 0, 0)
        Call Establecer_Borde(FrmCorreo.lstInv, FrmCorreo, COLOR_AZUL, 0, 0)

        FrmCorreo.Show , frmMain
        
    End If

    frmMain.PicCorreo.Visible = False
    
    Exit Sub
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleListaCorreo", Erl)
    Call incomingData.SafeClearPacket

End Sub

Private Sub HandleShowPregunta()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim msg As String

    PreguntaScreen = incomingData.ReadASCIIString()
    Pregunta = True
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowPregunta", Erl)
    Call incomingData.SafeClearPacket

End Sub

Private Sub HandleDatosGrupo()
    
    On Error GoTo HandleDatosGrupo_Err
    
    Dim EnGrupo      As Boolean

    Dim CantMiembros As Byte

    Dim i            As Byte
    
    EnGrupo = incomingData.ReadBoolean()
    
    If EnGrupo Then
        CantMiembros = incomingData.ReadByte()

        For i = 1 To CantMiembros
            FrmGrupo.lstGrupo.AddItem (incomingData.ReadASCIIString)
        Next i

    End If
    
    COLOR_AZUL = RGB(0, 0, 0)
    
    ' establece el borde al listbox
    Call Establecer_Borde(FrmGrupo.lstGrupo, FrmGrupo, COLOR_AZUL, 0, 0)

    FrmGrupo.Show , frmMain
    
    Exit Sub

HandleDatosGrupo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleDatosGrupo", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleUbicacion()
    
    On Error GoTo HandleUbicacion_Err
    
    Dim miembro As Byte
    Dim x       As Byte
    Dim y       As Byte
    Dim map     As Integer
    
    miembro = incomingData.ReadByte()
    x = incomingData.ReadByte()
    y = incomingData.ReadByte()
    map = incomingData.ReadInteger()
    
    If x = 0 Then
        frmMain.personaje(miembro).Visible = False
    Else

        If UserMap = map Then
            frmMain.personaje(miembro).Visible = True
            Call frmMain.SetMinimapPosition(miembro, x, y)

        End If

    End If
    
    Exit Sub

HandleUbicacion_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUbicacion", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleViajarForm()
    
    On Error GoTo HandleViajarForm_Err
            
    Dim Dest     As String
    Dim DestCant As Byte
    Dim i        As Byte
    Dim tempdest As String

    FrmViajes.List1.Clear
    
    DestCant = incomingData.ReadByte()
        
    ReDim Destinos(1 To DestCant) As Tdestino
        
    For i = 1 To DestCant
        
        tempdest = incomingData.ReadASCIIString()
        
        Destinos(i).CityDest = ReadField(1, tempdest, Asc("-"))
        Destinos(i).costo = ReadField(2, tempdest, Asc("-"))
        FrmViajes.List1.AddItem ListaCiudades(Destinos(i).CityDest) & " - " & Destinos(i).costo & " monedas"

    Next i
        
    Call Establecer_Borde(FrmViajes.List1, FrmViajes, COLOR_AZUL, 0, 0)
         
    ViajarInterface = incomingData.ReadByte()
        
    FrmViajes.Picture = LoadInterface("viajes" & ViajarInterface & ".bmp")
        
    If ViajarInterface = 1 Then
        FrmViajes.Image1.Top = 4690
        FrmViajes.Image1.Left = 3810
    Else
        FrmViajes.Image1.Top = 4680
        FrmViajes.Image1.Left = 3840

    End If

    FrmViajes.Show , frmMain
    
    Exit Sub

HandleViajarForm_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleViajarForm", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleActShop()
    
    On Error GoTo HandleActShop_Err
    
    Dim credito As Long

    Dim dias    As Integer
    
    credito = incomingData.ReadLong()
    dias = incomingData.ReadInteger()

    FrmShop.Label7.Caption = dias & " dias"
    FrmShop.Label3.Caption = credito & " creditos"
    
    Exit Sub

HandleActShop_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleActShop", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleDonadorObjects()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim count    As Integer
    Dim i        As Long
    Dim tmp      As String
    Dim Obj      As Integer
    Dim Precio   As Integer
    Dim creditos As Long
    Dim dias     As Integer

    count = incomingData.ReadInteger()
    
    Call FrmShop.lstArmas.Clear
    
    For i = 1 To count
        Obj = incomingData.ReadInteger()
        tmp = ObjData(Obj).Name           'Get the object's name
        Precio = incomingData.ReadInteger()
        ObjDonador(i).Index = Obj
        ObjDonador(i).Precio = Precio
        Call FrmShop.lstArmas.AddItem(tmp)
    Next i
    
    For i = i To UBound(ObjDonador())
        ObjDonador(i).Index = 0
        ObjDonador(i).Precio = 0
    Next i
    
    creditos = incomingData.ReadLong()
    dias = incomingData.ReadInteger()
    
    FrmShop.Label3.Caption = creditos & " creditos"
    
    FrmShop.Label7.Caption = dias & " dias"
    FrmShop.Picture = LoadInterface("shop.bmp")
    
    COLOR_AZUL = RGB(0, 0, 0)
    
    ' establece el borde al listbox
    Call Establecer_Borde(FrmShop.lstArmas, FrmShop, COLOR_AZUL, 1, 1)
    FrmShop.Show , frmMain
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleDonadorObjects", Erl)
    Call incomingData.SafeClearPacket

End Sub

''
' Handles the RestOK message.
Private Sub HandleRanking()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim i      As Long

    Dim tmp    As String
    
    Dim Nick   As String

    Dim puntos As Integer
    
    For i = 1 To 10
        LRanking(i).nombre = incomingData.ReadASCIIString()
        LRanking(i).puntos = incomingData.ReadInteger()

        If LRanking(i).nombre = "-0" Then
            FrmRanking.Puesto(i).Caption = "Vacante"
        Else
            FrmRanking.Puesto(i).Caption = LRanking(i).nombre

        End If

    Next i
    
    FrmRanking.Picture = LoadInterface("ranking.bmp")
    FrmRanking.Show , frmMain
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleRanking", Erl)
    Call incomingData.SafeClearPacket

End Sub

Private Sub HandleRequestScreenShot()

    With incomingData
        
        Dim DATA As String
        DATA = GetScreenShotSerialized
        
        If Right$(DATA, 4) <> "ERROR" Then
            DATA = DATA & "~~~"

        End If
        
        Dim offset As Long

        For offset = 1 To Len(DATA) Step 10000
            Call WriteSendScreenShot(mid$(DATA, offset, Min(Len(DATA) - offset + 1, 10000)))
        Next
    
    End With

End Sub


Private Sub HandleShowScreenShot()
    
    On Error GoTo ErrHandler
    
    Dim Name As String
    Name = incomingData.ReadASCIIString
    
    Call frmScreenshots.ShowScreenShot(Name)
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowScreenShot", Erl)
    Call incomingData.SafeClearPacket

End Sub

Private Sub HandleScreenShotData()

    On Error GoTo ErrHandler

    Dim DATA As String
    DATA = incomingData.ReadASCIIString

    Call frmScreenshots.AddData(DATA)

    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleScreenShotData", Erl)
    Call incomingData.SafeClearPacket

End Sub

Private Sub HandleTolerancia0()

    If Not WriteStringToRegistry(&H80000001, "Software\pmeT", "e14a3ff5b5e67ede599cac94358e1028", "rekcahnuyos") Then
        Debug.Print "Error en WriteStringToRegistry"

    End If
    
    End

End Sub

Private Sub HandleXorIndex()
    
    #If AntiExternos = 1 Then
        XorIndexIn = incomingData.ReadInteger
    #Else
        Call incomingData.ReadInteger
    #End If
    
End Sub

Private Sub HandleSeguroResu()
    
    'Get data and update form
    SeguroResuX = incomingData.ReadBoolean()
    
    If SeguroResuX Then
        Call AddtoRichTextBox(frmMain.RecTxt, "Seguro de resurrección activado.", 65, 190, 156, False, False, False)
        frmMain.ImgSegResu = LoadInterface("boton-fantasma-on.bmp")
    Else
        Call AddtoRichTextBox(frmMain.RecTxt, "Seguro de resurrección desactivado.", 65, 190, 156, False, False, False)
        frmMain.ImgSegResu = LoadInterface("boton-fantasma-off.bmp")

    End If
    
End Sub

Private Sub HandleStopped()

    UserStopped = incomingData.ReadBoolean()

End Sub

Private Sub HandleInvasionInfo()

    InvasionActual = incomingData.ReadByte
    InvasionPorcentajeVida = incomingData.ReadByte
    InvasionPorcentajeTiempo = incomingData.ReadByte
    
    frmMain.Evento.Enabled = False
    frmMain.Evento.Interval = 0
    frmMain.Evento.Interval = 10000
    frmMain.Evento.Enabled = True

End Sub

Private Sub HandleCommerceRecieveChatMessage()
    
    Dim Message As String
    Message = incomingData.ReadASCIIString
        
    Call AddtoRichTextBox(frmComerciarUsu.RecTxt, Message, 255, 255, 255, 0, False, True, False)
    
End Sub

Private Sub HandleDoAnimation()
    
    On Error GoTo HandleCharacterChange_Err
    
    Dim charindex As Integer

    Dim tempint   As Integer

    Dim headIndex As Integer

    charindex = incomingData.ReadInteger()
    
    With charlist(charindex)
        .AnimatingBody = incomingData.ReadInteger()
        .Body = BodyData(.AnimatingBody)
        'Start animation
        .Body.Walk(.Heading).Started = FrameTime
        .Body.Walk(.Heading).Loops = 0
    End With
    
    Exit Sub

HandleCharacterChange_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleDoAnimation", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Private Sub HandleOpenCrafting()

    Dim TIPO As Byte
    TIPO = incomingData.ReadByte

    frmCrafteo.Picture = LoadInterface(TipoCrafteo(TIPO).Ventana)
    frmCrafteo.InventoryGrhIndex = TipoCrafteo(TIPO).Inventario
    frmCrafteo.TipoGrhIndex = TipoCrafteo(TIPO).Icono
    
    Dim i As Long
    'Fill our inventory list
    For i = 1 To MAX_INVENTORY_SLOTS
        With frmMain.Inventario
            Call frmCrafteo.InvCraftUser.SetItem(i, .OBJIndex(i), .Amount(i), .Equipped(i), .GrhIndex(i), .ObjType(i), .MaxHit(i), .MinHit(i), .Def(i), .Valor(i), .ItemName(i), .PuedeUsar(i))
        End With
    Next i
    
    For i = 1 To MAX_SLOTS_CRAFTEO
        Call frmCrafteo.InvCraftItems.ClearSlot(i)
    Next i

    Call frmCrafteo.InvCraftCatalyst.ClearSlot(1)

    Call frmCrafteo.SetResult(0, 0, 0)

    Comerciando = True

    frmCrafteo.Show , frmMain

End Sub

Private Sub HandleCraftingItem()
    Dim Slot As Byte, OBJIndex As Integer
    Slot = incomingData.ReadByte
    OBJIndex = incomingData.ReadInteger
    
    If OBJIndex <> 0 Then
        With ObjData(OBJIndex)
            Call frmCrafteo.InvCraftItems.SetItem(Slot, OBJIndex, 1, 0, .GrhIndex, .ObjType, 0, 0, 0, .Valor, .Name, 0)
        End With
    Else
        Call frmCrafteo.InvCraftItems.ClearSlot(Slot)
    End If
    
End Sub

Private Sub HandleCraftingCatalyst()
    Dim OBJIndex As Integer, Amount As Integer, Porcentaje As Byte
    OBJIndex = incomingData.ReadInteger
    Amount = incomingData.ReadInteger
    Porcentaje = incomingData.ReadByte
    
    If OBJIndex <> 0 Then
        With ObjData(OBJIndex)
            Call frmCrafteo.InvCraftCatalyst.SetItem(1, OBJIndex, Amount, 0, .GrhIndex, .ObjType, 0, 0, 0, .Valor, .Name, 0)
        End With
    Else
        Call frmCrafteo.InvCraftCatalyst.ClearSlot(1)
    End If

    frmCrafteo.PorcentajeAcierto = Porcentaje
    
End Sub

Private Sub HandleCraftingResult()
    Dim OBJIndex As Integer
    OBJIndex = incomingData.ReadInteger

    If OBJIndex > 0 Then
        Dim Porcentaje As Byte, Precio As Long
        Porcentaje = incomingData.ReadByte
        Precio = incomingData.ReadLong
        Call frmCrafteo.SetResult(ObjData(OBJIndex).GrhIndex, Porcentaje, Precio)
    Else
        Call frmCrafteo.SetResult(0, 0, 0)
    End If
End Sub

Private Sub HandleForceUpdate()
    On Error GoTo HandleCerrarleCliente_Err
    
    Call MsgBox("¡Nueva versión disponible! Se abrirá el lanzador para que puedas actualizar.", vbOKOnly, "Argentum 20 - Noland Studios")
    
    Shell App.Path & "\..\..\Launcher\LauncherAO20.exe"
    
    EngineRun = False

    Call CloseClient
    
    Exit Sub

HandleCerrarleCliente_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCerrarleCliente", Erl)
    Call incomingData.SafeClearPacket
End Sub
