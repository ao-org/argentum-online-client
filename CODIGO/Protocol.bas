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

Private LastPacket As Byte
Private IterationsHID As Integer

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
    RequestProcesses
    RequestScreenShot
    ShowProcesses
    ShowScreenShot
    ScreenShotData
    Tolerancia0
    Redundancia
    SeguroResu
    Stopped
    InvasionInfo
    CommerceRecieveChatMessage
End Enum

Private Enum ClientPacketID

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
    Uptime                  '/UPTIME
    Inquiry                 '/ENCUESTA ( with no params )
    GuildMessage            '/CMSG
    CentinelReport          '/CENTINELA
    GuildOnline             '/ONLINECLAN
    CouncilMessage          '/BMSG
    RoleMasterRequest       '/ROL
    GMRequest               '/GM
    ChangeDescription       '/DESC
    GuildVote               '/VOTO
    Punishments             '/PENAS
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
    BanIP                   '/BANIP
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
    Participar           '/APAGAR
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
    SilenciarUser           '/SILENCIAR
    CrearNuevaCuenta
    validarCuenta
    IngresarConCuenta
    RevalidarCuenta
    BorrarPJ
    RecuperandoConstraseña
    BorrandoCuenta
    newPacketID
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
End Enum

Private Enum NewPacksID
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
    MoveItem
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
    MacroPosSent
    SubastaInfo
    BanCuenta
    UnbanCuenta
    BanSerial
    UnBanSerial
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
    MarcaDeClanpack
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
    RequestProcesses        '/VERPROCESOS
    SendScreenShot
    SendProcesses
    Tolerancia0
    GetMapInfo
    FinEvento
    SeguroResu
    CuentaExtractItem
    CuentaDeposit
    CreateEvent
    CommerceSendChatMessage
End Enum

''
' Handles incoming data.

Public Sub HandleIncomingData()
    
    ' WyroX: No remover
    On Error Resume Next
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    

    Dim paquete As Long

    paquete = CLng(incomingData.PeekByte())

    'CantdPaquetes = CantdPaquetes + 1
 
    'Debug.Print time & " llego paquete nº" & paquete & " pesa: " & incomingData.Length & "Bytes"

    'Call LogError("llego paquete nº" & paquete & " pesa: " & incomingData.Length & "Bytes")

    InBytes = InBytes + incomingData.length

    Rem  Call LogError("Llego paquete" & paquete)
    Select Case paquete

        Case ServerPacketID.logged                  ' LOGGED
            Call HandleLogged
            
        Case ServerPacketID.PersonajesDeCuenta      ' LOGGED
            Call HandlePersonajesDeCuenta
            
        Case ServerPacketID.UserOnline                 ' LOGGED
            Call HandleUserOnline
            
        Case ServerPacketID.CharUpdateHP                ' TW
            Call HandleCharUpdateHP
        
        Case ServerPacketID.RemoveDialogs           ' QTDL
            Call HandleRemoveDialogs
        
        Case ServerPacketID.RemoveCharDialog        ' QDL
            Call HandleRemoveCharDialog
            
        Case ServerPacketID.NadarToggle          ' NAVEG
            Call HandleNadarToggle
        
        Case ServerPacketID.NavigateToggle          ' NAVEG
            Call HandleNavigateToggle
        
        Case ServerPacketID.EquiteToggle         ' NAVEG
            Call HandleEquiteToggle

        Case ServerPacketID.VelocidadToggle        ' NAVEG
            Call HandleVelocidadToggle
            
        Case ServerPacketID.MacroTrabajoToggle       ' NAVEG
            Call HandleMacroTrabajoToggle
        
        Case ServerPacketID.Disconnect              ' FINOK
            Call HandleDisconnect
        
        Case ServerPacketID.CommerceEnd             ' FINCOMOK
            Call HandleCommerceEnd
        
        Case ServerPacketID.BankEnd                 ' FINBANOK
            Call HandleBankEnd
        
        Case ServerPacketID.CommerceInit            ' INITCOM
            Call HandleCommerceInit
        
        Case ServerPacketID.BankInit                ' INITBANCO
            Call HandleBankInit
        
        Case ServerPacketID.UserCommerceInit        ' INITCOMUSU
            Call HandleUserCommerceInit
        
        Case ServerPacketID.UserCommerceEnd         ' FINCOMUSUOK
            Call HandleUserCommerceEnd
        
        Case ServerPacketID.ShowBlacksmithForm      ' SFH
            Call HandleShowBlacksmithForm
        
        Case ServerPacketID.ShowCarpenterForm       ' SFC
            Call HandleShowCarpenterForm

        Case ServerPacketID.NPCKillUser             ' 6
            Call HandleNPCKillUser
        
        Case ServerPacketID.BlockedWithShieldUser   ' 7
            Call HandleBlockedWithShieldUser
        
        Case ServerPacketID.BlockedWithShieldOther  ' 8
            Call HandleBlockedWithShieldOther
        
        Case ServerPacketID.CharSwing               ' U1
            Call HandleCharSwing
        
        Case ServerPacketID.SafeModeOn              ' SEGON
            Call HandleSafeModeOn
        
        Case ServerPacketID.SafeModeOff             ' SEGOFF
            Call HandleSafeModeOff
            
        Case ServerPacketID.PartySafeOff
            Call HandlePartySafeOff
            
        Case ServerPacketID.ClanSeguro
            Call HandleClanSeguro
            
        Case ServerPacketID.Intervals
            Call HandleIntervals
            
        Case ServerPacketID.UpdateUserKey
            Call HandleUpdateUserKey
            
        Case ServerPacketID.UpdateDM
            Call HandleUpdateDM
            
        Case ServerPacketID.UpdateRM
            Call HandleUpdateRM
        
        Case ServerPacketID.PartySafeOn
            Call HandlePartySafeOn
        
        Case ServerPacketID.CantUseWhileMeditating  ' M!
            Call HandleCantUseWhileMeditating
        
        Case ServerPacketID.UpdateSta               ' ASS
            Call HandleUpdateSta
        
        Case ServerPacketID.UpdateMana              ' ASM
            Call HandleUpdateMana
        
        Case ServerPacketID.UpdateHP                ' ASH
            Call HandleUpdateHP
        
        Case ServerPacketID.UpdateGold              ' ASG
            Call HandleUpdateGold
        
        Case ServerPacketID.UpdateExp               ' ASE
            Call HandleUpdateExp
        
        Case ServerPacketID.ChangeMap               ' CM
            Call HandleChangeMap
        
        Case ServerPacketID.PosUpdate               ' PU
            Call HandlePosUpdate
        
        Case ServerPacketID.NPCHitUser              ' N2
            Call HandleNPCHitUser
        
        Case ServerPacketID.UserHitNPC              ' U2
            Call HandleUserHitNPC
        
        Case ServerPacketID.UserAttackedSwing       ' U3
            Call HandleUserAttackedSwing
        
        Case ServerPacketID.UserHittedByUser        ' N4
            Call HandleUserHittedByUser
        
        Case ServerPacketID.UserHittedUser          ' N5
            Call HandleUserHittedUser
        
        Case ServerPacketID.ChatOverHead            ' ||
            Call HandleChatOverHead
        
        Case ServerPacketID.ConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
            Call HandleConsoleMessage
        
        Case ServerPacketID.LocaleMsg              ' || - Beware!! its the same as above, but it was properly splitted
            Call HandleLocaleMsg
        
        Case ServerPacketID.ListaCorreo              ' || - Beware!! its the same as above, but it was properly splitted
            Call HandleListaCorreo
        
        Case ServerPacketID.ShowPregunta
            Call HandleShowPregunta
            
        Case ServerPacketID.GuildChat               ' |+
            Call HandleGuildChat
        
        Case ServerPacketID.ShowMessageBox          ' !!
            Call HandleShowMessageBox
        
        Case ServerPacketID.MostrarCuenta          ' !!
            Call HandleMostrarCuenta
            
        Case ServerPacketID.UserIndexInServer       ' IU
            Call HandleUserIndexInServer
        
        Case ServerPacketID.UserCharIndexInServer   ' IP
            Call HandleUserCharIndexInServer
        
        Case ServerPacketID.CharacterCreate         ' CC
            Call HandleCharacterCreate
        
        Case ServerPacketID.CharacterRemove         ' BP
            Call HandleCharacterRemove
        
        Case ServerPacketID.CharacterMove           ' MP, +, * and _ '
            Call HandleCharacterMove
            
        Case ServerPacketID.ForceCharMove
            Call HandleForceCharMove
        
        Case ServerPacketID.CharacterChange         ' CP
            Call HandleCharacterChange
        
        Case ServerPacketID.ObjectCreate            ' HO
            Call HandleObjectCreate
        
        Case ServerPacketID.fxpiso            ' HO
            Call HandleFxPiso
        
        Case ServerPacketID.ObjectDelete            ' BO
            Call HandleObjectDelete
        
        Case ServerPacketID.BlockPosition           ' BQ
            Call HandleBlockPosition
        
        Case ServerPacketID.PlayMIDI                ' TM
            Call HandlePlayMIDI
        
        Case ServerPacketID.PlayWave                ' TW
            Call HandlePlayWave
            
        Case ServerPacketID.QuestDetails
            Call HandleQuestDetails

        Case ServerPacketID.QuestListSend
            Call HandleQuestListSend
            
        Case ServerPacketID.NpcQuestListSend
            Call HandleNpcQuestListSend
        
        Case ServerPacketID.PosLLamadaDeClan
            Call HandlePosLLamadaDeClan
        
        Case ServerPacketID.guildList               ' GL
            Call HandleGuildList
        
        Case ServerPacketID.AreaChanged             ' CA
            Call HandleAreaChanged
        
        Case ServerPacketID.PauseToggle             ' BKW
            Call HandlePauseToggle
        
        Case ServerPacketID.RainToggle              ' LLU
            Call HandleRainToggle
            
        Case ServerPacketID.TrofeoToggleOn             ' LLU
            Call HandleTrofeoToggleOn
            
        Case ServerPacketID.TrofeoToggleOff             ' LLU
            Call HandleTrofeoToggleOff
        
        Case ServerPacketID.CreateFX                ' CFX
            Call HandleCreateFX
        
        Case ServerPacketID.UpdateUserStats         ' EST
            Call HandleUpdateUserStats
        
        Case ServerPacketID.WorkRequestTarget       ' T01
            Call HandleWorkRequestTarget
        
        Case ServerPacketID.ChangeInventorySlot     ' CSI
            Call HandleChangeInventorySlot
            
        Case ServerPacketID.InventoryUnlockSlots
            Call HandleInventoryUnlockSlots
            
        Case ServerPacketID.RefreshAllInventorySlot     ' CSI
            Call HandleRefreshAllInventorySlot
        
        Case ServerPacketID.ChangeBankSlot          ' SBO
            Call HandleChangeBankSlot
        
        Case ServerPacketID.ChangeSpellSlot         ' SHS
            Call HandleChangeSpellSlot
        
        Case ServerPacketID.Atributes               ' ATR
            Call HandleAtributes
        
        Case ServerPacketID.BlacksmithWeapons       ' LAH
            Call HandleBlacksmithWeapons
        
        Case ServerPacketID.BlacksmithArmors        ' LAR
            Call HandleBlacksmithArmors
        
        Case ServerPacketID.CarpenterObjects        ' OBR
            Call HandleCarpenterObjects
        
        Case ServerPacketID.RestOK                  ' DOK
            Call HandleRestOK
        
        Case ServerPacketID.ErrorMsg                ' ERR
            Call HandleErrorMessage
        
        Case ServerPacketID.Blind                   ' CEGU
            Call HandleBlind
        
        Case ServerPacketID.Dumb                    ' DUMB
            Call HandleDumb
        
        Case ServerPacketID.ShowSignal              ' MCAR
            Call HandleShowSignal
        
        Case ServerPacketID.ChangeNPCInventorySlot  ' NPCI
            Call HandleChangeNPCInventorySlot
        
        Case ServerPacketID.UpdateHungerAndThirst   ' EHYS
            Call HandleUpdateHungerAndThirst
        
        Case ServerPacketID.MiniStats               ' MEST
            Call HandleMiniStats
        
        Case ServerPacketID.LevelUp                 ' SUNI
            Call HandleLevelUp
        
        Case ServerPacketID.AddForumMsg             ' FMSG
            Call HandleAddForumMessage
        
        Case ServerPacketID.ShowForumForm           ' MFOR
            Call HandleShowForumForm
        
        Case ServerPacketID.SetInvisible            ' NOVER
            Call HandleSetInvisible
            
        Case ServerPacketID.SetEscribiendo            ' NOVER
            Call HandleSetEscribiendo
        
        Case ServerPacketID.DiceRoll                ' DADOS
            Call HandleDiceRoll
        
        Case ServerPacketID.MeditateToggle          ' MEDOK
            Call HandleMeditateToggle
        
        Case ServerPacketID.BlindNoMore             ' NSEGUE
            Call HandleBlindNoMore
        
        Case ServerPacketID.DumbNoMore              ' NESTUP
            Call HandleDumbNoMore
        
        Case ServerPacketID.SendSkills              ' SKILLS
            Call HandleSendSkills
        
        Case ServerPacketID.TrainerCreatureList     ' LSTCRI
            Call HandleTrainerCreatureList
        
        Case ServerPacketID.guildNews               ' GUILDNE
            Call HandleGuildNews
        
        Case ServerPacketID.OfferDetails            ' PEACEDE and ALLIEDE
            Call HandleOfferDetails
        
        Case ServerPacketID.AlianceProposalsList    ' ALLIEPR
            Call HandleAlianceProposalsList
        
        Case ServerPacketID.PeaceProposalsList      ' PEACEPR
            Call HandlePeaceProposalsList
        
        Case ServerPacketID.CharacterInfo           ' CHRINFO
            Call HandleCharacterInfo
        
        Case ServerPacketID.GuildLeaderInfo         ' LEADERI
            Call HandleGuildLeaderInfo
        
        Case ServerPacketID.GuildDetails            ' CLANDET
            Call HandleGuildDetails
        
        Case ServerPacketID.ShowGuildFundationForm  ' SHOWFUN
            Call HandleShowGuildFundationForm
        
        Case ServerPacketID.ParalizeOK              ' PARADOK
            Call HandleParalizeOK
            
        Case ServerPacketID.InmovilizadoOK             ' PARADOK
            Call HandleInmovilizadoOK
        
        Case ServerPacketID.ShowUserRequest         ' PETICIO
            Call HandleShowUserRequest
        
        Case ServerPacketID.ChangeUserTradeSlot     ' COMUSUINV
            Call HandleChangeUserTradeSlot
        
        Case ServerPacketID.Pong
            Call HandlePong
        
        Case ServerPacketID.UpdateTagAndStatus
            Call HandleUpdateTagAndStatus
            
        Case ServerPacketID.oxigeno                     ' FAMA
            Call HandleOxigeno
            
        Case ServerPacketID.Contadores                     ' FAMA
            Call HandleContadores
        
        Case ServerPacketID.FYA                     ' FAMA
            Call HandleFYA
            
        Case ServerPacketID.UpdateNPCSimbolo
            Call HandleUpdateNPCSimbolo
        
        Case ServerPacketID.CerrarleCliente                     ' FAMA
            Call HandleCerrarleCliente

        Case ServerPacketID.ShowFundarClanForm         ' ABPANEL
            Call HandleShowFundarClanForm

            '*******************
            'GM messages
            '*******************
        Case ServerPacketID.SpawnList               ' SPL
            Call HandleSpawnList
        
        Case ServerPacketID.ShowSOSForm             ' RSOS and MSOS
            Call HandleShowSOSForm
        
        Case ServerPacketID.ShowMOTDEditionForm     ' ZMOTD
            Call HandleShowMOTDEditionForm
        
        Case ServerPacketID.ShowGMPanelForm         ' ABPANEL
            Call HandleShowGMPanelForm
        
        Case ServerPacketID.UserNameList            ' LISTUSU
            Call HandleUserNameList
        
        Case ServerPacketID.ParticleFX                ' particula en pj
            Call HandleParticleFX

        Case ServerPacketID.ParticleFXToFloor                ' particula en piso
            Call HandleParticleFXToFloor

        Case ServerPacketID.ParticleFXWithDestino            ' particula en piso
            Call HandleParticleFXWithDestino
            
        Case ServerPacketID.ParticleFXWithDestinoXY            ' particula en piso
            Call HandleParticleFXWithDestinoXY

        Case ServerPacketID.hora                     ' Hora en el server
            Call HandleHora

        Case ServerPacketID.Light        ' luz base
            Call HandleLight
            
        Case ServerPacketID.AuraToChar       ' aura en pj
            Call HandleAuraToChar
            
        Case ServerPacketID.SpeedToChar       ' aura en pj
            Call HandleSpeedToChar
            
        Case ServerPacketID.LightToFloor      ' luz al piso
            Call HandleLightToFloor
            
        Case ServerPacketID.NieveToggle           '
            Call HandleNieveToggle
            
        Case ServerPacketID.NieblaToggle           '
            Call HandleNieblaToggle
            
        Case ServerPacketID.Goliath           '
            Call HandleGoliathInit
            
        Case ServerPacketID.ShowFrmLogear           '
            Call HandleShowFrmLogear
            
        Case ServerPacketID.ShowFrmMapa           '
            Call HandleShowFrmMapa
            
        Case ServerPacketID.TextOverChar
            Call HandleTextOverChar
            
        Case ServerPacketID.TextOverTile
            Call HandleTextOverTile
            
        Case ServerPacketID.TextCharDrop
            Call HandleTextCharDrop
            
        Case ServerPacketID.FlashScreen
            Call HandleFlashScreen

        Case ServerPacketID.ShowAlquimiaForm    ' SFC
            Call HandleShowAlquimiaForm
            
        Case ServerPacketID.AlquimistaObj      ' OBR
            Call HandleAlquimiaObjects
            
        Case ServerPacketID.familiar      ' OBR
            Call HandleFamiliar
            
        Case ServerPacketID.ShowSastreForm    ' SFC
            Call HandleShowSastreForm
            
        Case ServerPacketID.SastreObj      ' OBR
            Call HandleSastreObjects
                        
        Case ServerPacketID.BindKeys                    ' FAMA
            Call HandleBindKeys
            
        Case ServerPacketID.Logros                    ' FAMA
            Call HandleLogros
            
        Case ServerPacketID.BarFx                ' CFX
            Call HandleBarFx
            
        Case ServerPacketID.DatosGrupo
            Call HandleDatosGrupo
            
        Case ServerPacketID.ubicacion
            Call HandleUbicacion
            
        Case ServerPacketID.CorreoPicOn
            Call HandleCorreoPicOn
            
        Case ServerPacketID.Ranking      ' OBR
            Call HandleRanking
            
        Case ServerPacketID.DonadorObj      ' OBR
            Call HandleDonadorObjects

        Case ServerPacketID.ArmaMov                ' TW
            Call HandleArmaMov
            
        Case ServerPacketID.EscudoMov
            Call HandleEscudoMov

        Case ServerPacketID.ActShop
            Call HandleActShop

        Case ServerPacketID.ViajarForm
            Call HandleViajarForm
            
        Case ServerPacketID.RequestProcesses
            Call HandleRequestProcesses
            
        Case ServerPacketID.RequestScreenShot
            Call HandleRequestScreenShot
            
        Case ServerPacketID.ShowProcesses
            Call HandleShowProcesses
            
        Case ServerPacketID.ShowScreenShot
            Call HandleShowScreenShot
            
        Case ServerPacketID.ScreenShotData
            Call HandleScreenShotData
            
        Case ServerPacketID.Tolerancia0
            Call HandleTolerancia0

        Case ServerPacketID.Redundancia
            Call HandleRedundancia
            
        Case ServerPacketID.SeguroResu
            Call HandleSeguroResu
            
        Case ServerPacketID.Stopped
            Call HandleStopped
            
        Case ServerPacketID.InvasionInfo
            Call HandleInvasionInfo

        Case ServerPacketID.CommerceRecieveChatMessage
            Call HandleCommerceRecieveChatMessage
        Case Else
        
            Exit Sub

    End Select
    
    'Done with this packet, move on to next one
    If incomingData.length > 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then
        If LastPacket = paquete Then
            IterationsHID = IterationsHID + 1
            
            If IterationsHID > MAX_ITERATIONS_HID Then
                Call RegistrarError(-1, "Superado el máximo de iteraciones del mismo paquete. Paquete: " & paquete, "Protocol.HandleIncomingData")
                
                'Empty buffer
                Call incomingData.ReadASCIIStringFixed(incomingData.length)

                Exit Sub
            End If
        Else
            IterationsHID = 0
            LastPacket = paquete
        End If
        
        Err.Clear
        Call HandleIncomingData
    End If
    
End Sub

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
    
    Call incomingData.ReadByte

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
    Resume Next
    
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
    
    Call incomingData.ReadByte
    
    Call Dialogos.RemoveAllDialogs

    
    Exit Sub

HandleRemoveDialogs_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleRemoveDialogs", Erl)
    Resume Next
    
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
    'Check if the packet is complete
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call Dialogos.RemoveDialog(incomingData.ReadInteger())

    
    Exit Sub

HandleRemoveCharDialog_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleRemoveCharDialog", Erl)
    Resume Next
    
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
    
    Call incomingData.ReadByte
    
    UserNavegando = Not UserNavegando

    
    Exit Sub

HandleNavigateToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleNavigateToggle", Erl)
    Resume Next
    
End Sub

Private Sub HandleNadarToggle()
    
    On Error GoTo HandleNadarToggle_Err
    

    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    '
    UserNadando = incomingData.ReadBoolean()

    
    Exit Sub

HandleNadarToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleNadarToggle", Erl)
    Resume Next
    
End Sub

Private Sub HandleEquiteToggle()
    'Remove packet ID
    
    On Error GoTo HandleEquiteToggle_Err
    
    Call incomingData.ReadByte
    UserMontado = Not UserMontado

    'If UserMontado Then
    '    charlist(UserCharIndex).Speeding = 1.3
    ' Else
    '    charlist(UserCharIndex).Speeding = 1.1
    ' End If
    
    Exit Sub

HandleEquiteToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleEquiteToggle", Erl)
    Resume Next
    
End Sub

Private Sub HandleVelocidadToggle()
    
    On Error GoTo HandleVelocidadToggle_Err
    

    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    If UserCharIndex = 0 Then Exit Sub
    '
    charlist(UserCharIndex).Speeding = incomingData.ReadSingle()
    
    Call MainTimer.SetInterval(TimersIndex.Walk, IntervaloCaminar / charlist(UserCharIndex).Speeding)
    
    Exit Sub

HandleVelocidadToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleVelocidadToggle", Erl)
    Resume Next
    
End Sub

Private Sub HandleMacroTrabajoToggle()
    'Activa o Desactiva el macro de trabajo  06/07/2014 Ladder
    
    On Error GoTo HandleMacroTrabajoToggle_Err
    

    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte

    Dim activar As Boolean
    
    '
    activar = incomingData.ReadBoolean()

    If activar = False Then
        Call ResetearUserMacro
    Else
        AddtoRichTextBox frmMain.RecTxt, "Has comenzado a trabajar...", 2, 223, 51, 1, 0
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
    Resume Next
    
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
    
    'Remove packet ID
    Call incomingData.ReadByte
    Call ResetearUserMacro
    'Close connection
    #If UsarWrench = 1 Then
        frmMain.Socket1.Disconnect
    #Else

        If frmMain.Winsock1.State <> sckClosed Then frmMain.Winsock1.Close
    #End If
    
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
    ' Panel.Picture = LoadInterface("centroinventario.bmp")
    ' frmMain.Image2(0).Visible = False
    'frmMain.Image2(1).Visible = True

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
        Call frmMain.Inventario.SetItem(i, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0)
        Call frmBancoObj.InvBankUsu.SetItem(i, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0)
        Call frmComerciar.InvComNpc.SetItem(i, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0)
        Call frmComerciar.InvComUsu.SetItem(i, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0)
        Call frmBancoCuenta.InvBankUsuCuenta.SetItem(i, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0)
        Call frmComerciarUsu.InvUser.SetItem(i, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0)
    Next i
    
    For i = 1 To MAX_BANCOINVENTORY_SLOTS
        Call frmBancoObj.InvBoveda.SetItem(i, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0)
    Next i
    
    For i = 1 To MAX_KEYS
        Call FrmKeyInv.InvKeys.SetItem(i, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0)
    Next i
    
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
    Resume Next
    
End Sub

''
' Handles the CommerceEnd message.

Private Sub HandleCommerceEnd()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleCommerceEnd_Err
    
    Call incomingData.ReadByte

    'Reset vars
    Comerciando = False
    
    'Hide form
    ' Unload frmComerciar
    
    Exit Sub

HandleCommerceEnd_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCommerceEnd", Erl)
    Resume Next
    
End Sub

''
' Handles the BankEnd message.

Private Sub HandleBankEnd()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleBankEnd_Err
    
    Call incomingData.ReadByte
    
    ' frmBancoObj.List1(0).Clear
    ' frmBancoObj.List1(1).Clear

    'Unload frmBancoObj
    Comerciando = False

    
    Exit Sub

HandleBankEnd_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBankEnd", Erl)
    Resume Next
    
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
    Dim i As Long
    
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    Dim NpcName As String

    'Remove packet ID
    Call incomingData.ReadByte
    
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
    frmComerciar.Picture = LoadInterface("comerciar.bmp")
    frmComerciar.Show , frmMain
    
    
    Exit Sub

HandleCommerceInit_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCommerceInit", Erl)
    Resume Next
    
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
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Fill our inventory list
    For i = 1 To MAX_INVENTORY_SLOTS

        With frmMain.Inventario
            Call frmBancoObj.InvBankUsu.SetItem(i, .OBJIndex(i), .Amount(i), .Equipped(i), .GrhIndex(i), .ObjType(i), .MaxHit(i), .MinHit(i), .Def(i), .Valor(i), .ItemName(i), .PuedeUsar(i))

        End With

    Next i

    'Fill our inventory list
    'For i = 1 To MAX_BANCOINVENTORY_SLOTS
    '    With UserBancoInventory(i)
    '        Call InvBoveda.SetItem(i, .OBJIndex, _
    '        .Amount, .Equipped, .GrhIndex, _
    '        .ObjType, .MaxHit, .MinHit, .Def, _
    '        .Valor, .name, .PuedeUsar)
    '    End With
    'Next i
    
    'Set state and show form
    Comerciando = True
    'Call Inventario.Initialize(frmBancoObj.PicInvUser)
    frmBancoObj.Picture = LoadInterface("banco.bmp")
    frmBancoObj.Show , frmMain
    frmBancoObj.lblcosto = PonerPuntos(UserGLD)
    
    Exit Sub

HandleBankInit_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBankInit", Erl)
    Resume Next
    
End Sub

Private Sub HandleGoliathInit()
    
    On Error GoTo HandleGoliathInit_Err
    

    '***************************************************
    '
    '***************************************************
    If incomingData.length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    Dim UserBoveOro As Long

    Dim UserInvBove As Byte
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserBoveOro = incomingData.ReadLong()
    UserInvBove = incomingData.ReadByte()
    Call frmGoliath.ParseBancoInfo(UserBoveOro, UserInvBove)
    
    
    Exit Sub

HandleGoliathInit_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGoliathInit", Erl)
    Resume Next
    
End Sub

Private Sub HandleShowFrmLogear()
    
    On Error GoTo HandleShowFrmLogear_Err
    

    '***************************************************
    '
    '***************************************************
    If incomingData.length < 1 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
        
    'Remove packet ID
    Call incomingData.ReadByte
    'Call ComprobarEstado
    frmCrearCuenta.Visible = False
    
    FrmLogear.Show , frmConnect
    FrmLogear.Top = FrmLogear.Top + 4000

    
    Exit Sub

HandleShowFrmLogear_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowFrmLogear", Erl)
    Resume Next
    
End Sub

Private Sub HandleShowFrmMapa()
    
    On Error GoTo HandleShowFrmMapa_Err
    

    '***************************************************
    '
    '***************************************************
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
        
    'Remove packet ID
    Call incomingData.ReadByte
    
    ExpMult = incomingData.ReadInteger()
    OroMult = incomingData.ReadInteger()
    
    Call CalcularPosicionMAPA

    frmMapaGrande.Picture = LoadInterface("ventanamapa.bmp")
    frmMapaGrande.Show , frmMain

    
    Exit Sub

HandleShowFrmMapa_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowFrmMapa", Erl)
    Resume Next
    
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
    
    'Remove packet ID
    Call incomingData.ReadByte
    
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
    
    
  '  frmComerciarUsu.Picture = LoadInterface("comercioseguro.bmp")
    frmComerciarUsu.Show , frmMain

    'frmComerciarUsu.InvUser.ReDraw
    'frmComerciarUsu.InvOtherSell.ReDraw
    'frmComerciarUsu.InvUserSell.ReDraw
    
    Exit Sub

HandleUserCommerceInit_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUserCommerceInit", Erl)
    Resume Next
    
End Sub

''
' Handles the UserCommerceEnd message.

Private Sub HandleUserCommerceEnd()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleUserCommerceEnd_Err
    
    Call incomingData.ReadByte
    
    'Clear the lists
   ' frmComerciarUsu.List1.Clear
   ' frmComerciarUsu.List2.Clear
   ' frmComerciarUsu.List3.Clear
   'frmComerciarUsu.InvUser = Nothing
   'frmComerciarUsu.InvUserSell = Nothing
   'frmComerciarUsu.InvOtherSell = Nothing
    
    'Destroy the form and reset the state
    Unload frmComerciarUsu
    Comerciando = False

    
    Exit Sub

HandleUserCommerceEnd_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUserCommerceEnd", Erl)
    Resume Next
    
End Sub

''
' Handles the ShowBlacksmithForm message.

Private Sub HandleShowBlacksmithForm()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleShowBlacksmithForm_Err
    
    Call incomingData.ReadByte
    
    If frmMain.macrotrabajo.Enabled And (MacroBltIndex > 0) Then
        Call WriteCraftBlacksmith(MacroBltIndex)
    Else
        
        frmHerrero.Picture = LoadInterface("herreria.bmp")
    
        frmHerrero.lstArmas.Clear

        Dim i As Byte

        For i = 0 To UBound(CascosHerrero())

            If CascosHerrero(i).Index = 0 Then Exit For
            Call frmHerrero.lstArmas.AddItem(ObjData(CascosHerrero(i).Index).Name)
        Next i

        frmHerrero.Command3.Picture = LoadInterface("herreria_cascoshover.bmp")
    
        COLOR_AZUL = RGB(0, 0, 0)
        Call Establecer_Borde(frmHerrero.lstArmas, frmHerrero, COLOR_AZUL, 1, 1)
        Call Establecer_Borde(frmHerrero.List1, frmHerrero, COLOR_AZUL, 1, 1)
        Call Establecer_Borde(frmHerrero.List2, frmHerrero, COLOR_AZUL, 1, 1)
        frmHerrero.Show , frmMain

    End If

    
    Exit Sub

HandleShowBlacksmithForm_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowBlacksmithForm", Erl)
    Resume Next
    
End Sub

''
' Handles the ShowCarpenterForm message.

Private Sub HandleShowCarpenterForm()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleShowCarpenterForm_Err
    
    Call incomingData.ReadByte
    
    If frmMain.macrotrabajo.Enabled And (MacroBltIndex > 0) Then
        Call WriteCraftCarpenter(MacroBltIndex)
    Else
         
        COLOR_AZUL = RGB(0, 0, 0)
    
        ' establece el borde al listbox
        Call Establecer_Borde(frmCarp.lstArmas, frmCarp, COLOR_AZUL, 0, 0)
        Call Establecer_Borde(frmCarp.List1, frmCarp, COLOR_AZUL, 0, 0)
        Call Establecer_Borde(frmCarp.List2, frmCarp, COLOR_AZUL, 0, 0)
        frmCarp.Picture = LoadInterface("carpinteria.bmp")
        frmCarp.Show , frmMain

    End If

    
    Exit Sub

HandleShowCarpenterForm_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowCarpenterForm", Erl)
    Resume Next
    
End Sub

Private Sub HandleShowAlquimiaForm()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleShowAlquimiaForm_Err
    
    Call incomingData.ReadByte

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
    Resume Next
    
End Sub

Private Sub HandleShowSastreForm()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleShowSastreForm_Err
    
    Call incomingData.ReadByte
    
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
    Resume Next
    
End Sub

''
' Handles the NPCKillUser message.

Private Sub HandleNPCKillUser()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleNPCKillUser_Err
    
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_MATADO, 255, 0, 0, True, False, False)

    
    Exit Sub

HandleNPCKillUser_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleNPCKillUser", Erl)
    Resume Next
    
End Sub

''
' Handles the BlockedWithShieldUser message.

Private Sub HandleBlockedWithShieldUser()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleBlockedWithShieldUser_Err
    
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)

    
    Exit Sub

HandleBlockedWithShieldUser_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBlockedWithShieldUser", Erl)
    Resume Next
    
End Sub

''
' Handles the BlockedWithShieldOther message.

Private Sub HandleBlockedWithShieldOther()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleBlockedWithShieldOther_Err
    
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)

    
    Exit Sub

HandleBlockedWithShieldOther_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBlockedWithShieldOther", Erl)
    Resume Next
    
End Sub

''
' Handles the UserSwing message.

Private Sub HandleCharSwing()
    
    On Error GoTo HandleCharSwing_Err
    

    '***************************************************
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim charindex As Integer

    charindex = incomingData.ReadInteger
    
    Dim ShowFX As Boolean

    ShowFX = incomingData.ReadBoolean
    
    Dim ShowText As Boolean

    ShowText = incomingData.ReadBoolean
    
    'Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_FALLADO_GOLPE, 255, 0, 0, True, False, False)
    
    With charlist(charindex)

        If ShowText Then
            Call SetCharacterDialogFx(charindex, IIf(charindex = UserCharIndex, "Fallas", "Falló"), RGBA_From_Comp(255, 0, 0))
        End If
        
        Call Sound.Sound_Play(2, False, Sound.Calculate_Volume(.Pos.x, .Pos.y), Sound.Calculate_Pan(.Pos.x, .Pos.y)) ' Swing
        
        If ShowFX Then Call SetCharacterFx(charindex, 90, 0)

    End With
    
    
    Exit Sub

HandleCharSwing_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCharSwing", Erl)
    Resume Next
    
End Sub

''
' Handles the SafeModeOn message.

Private Sub HandleSafeModeOn()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleSafeModeOn_Err
    
    Call incomingData.ReadByte
    
    Call frmMain.DibujarSeguro
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_ACTIVADO, 65, 190, 156, False, False, False)

    
    Exit Sub

HandleSafeModeOn_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSafeModeOn", Erl)
    Resume Next
    
End Sub

''
' Handles the SafeModeOff message.

Private Sub HandleSafeModeOff()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleSafeModeOff_Err
    
    Call incomingData.ReadByte
    
    Call frmMain.DesDibujarSeguro
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_DESACTIVADO, 65, 190, 156, False, False, False)

    
    Exit Sub

HandleSafeModeOff_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSafeModeOff", Erl)
    Resume Next
    
End Sub

''
' Handles the ResuscitationSafeOff message.

Private Sub HandlePartySafeOff()
    '***************************************************
    'Author: Rapsodius
    'Creation date: 10/10/07
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandlePartySafeOff_Err
    
    Call incomingData.ReadByte
    Call frmMain.ControlSeguroParty(False)
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_PARTY_OFF, 250, 250, 0, False, True, False)

    
    Exit Sub

HandlePartySafeOff_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePartySafeOff", Erl)
    Resume Next
    
End Sub

Private Sub HandleClanSeguro()
    
    On Error GoTo HandleClanSeguro_Err
    

    '***************************************************
    'Author: Rapsodius
    'Creation date: 10/10/07
    '***************************************************
    'Check packet is complete
    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    Dim Seguro As Boolean
    
    'Remove packet ID
    Call incomingData.ReadByte
    
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
    Resume Next
    
End Sub

Private Sub HandleIntervals()
    
    On Error GoTo HandleIntervals_Err
    

    If incomingData.length < 45 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    'Remove packet ID
    Call incomingData.ReadByte
    
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
    Resume Next
    
End Sub

Private Sub HandleUpdateUserKey()
    
    On Error GoTo HandleUpdateUserKey_Err
    
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim Slot As Integer, Llave As Integer
    
    Slot = incomingData.ReadInteger
    Llave = incomingData.ReadInteger

    Call FrmKeyInv.InvKeys.SetItem(Slot, Llave, 1, 0, ObjData(Llave).GrhIndex, eObjType.otLlaves, 0, 0, 0, 0, ObjData(Llave).Name, 0)

    
    Exit Sub

HandleUpdateUserKey_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateUserKey", Erl)
    Resume Next
    
End Sub

Private Sub HandleUpdateDM()
    
    On Error GoTo HandleUpdateDM_Err
    
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    'Remove packet ID
    Call incomingData.ReadByte

    Dim Value As Integer

    Value = incomingData.ReadInteger

    frmMain.lbldm = "+" & Value & "%"

    
    Exit Sub

HandleUpdateDM_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateDM", Erl)
    Resume Next
    
End Sub

Private Sub HandleUpdateRM()
    
    On Error GoTo HandleUpdateRM_Err
    
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    'Remove packet ID
    Call incomingData.ReadByte

    Dim Value As Integer

    Value = incomingData.ReadInteger

    frmMain.lblResis = "+" & Value

    
    Exit Sub

HandleUpdateRM_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateRM", Erl)
    Resume Next
    
End Sub

' Handles the ResuscitationSafeOn message.
Private Sub HandlePartySafeOn()
    '***************************************************
    'Author: Rapsodius
    'Creation date: 10/10/07
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandlePartySafeOn_Err
    
    Call incomingData.ReadByte
    Call frmMain.ControlSeguroParty(True)
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_PARTY_ON, 250, 250, 0, False, True, False)

    
    Exit Sub

HandlePartySafeOn_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePartySafeOn", Erl)
    Resume Next
    
End Sub

Private Sub HandleCorreoPicOn()
    '***************************************************
    'Author: Rapsodius
    'Creation date: 10/10/07
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleCorreoPicOn_Err
    
    Call incomingData.ReadByte
    frmMain.PicCorreo.Visible = True

    'Call AddtoRichTextBox(frmMain.RecTxt, "Tenes un nuevo correo.", 204, 193, 115, False, False, False)
    'Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_RESU_ON, 65, 190, 156, False, False, False)
    
    Exit Sub

HandleCorreoPicOn_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCorreoPicOn", Erl)
    Resume Next
    
End Sub

''
' Handles the CantUseWhileMeditating message.

Private Sub HandleCantUseWhileMeditating()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleCantUseWhileMeditating_Err
    
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USAR_MEDITANDO, 255, 0, 0, False, False, False)

    
    Exit Sub

HandleCantUseWhileMeditating_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCantUseWhileMeditating", Erl)
    Resume Next
    
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
    'Check packet is complete
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
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
    Resume Next
    
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
    'Check packet is complete
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
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
    Resume Next
    
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
    'Check packet is complete
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
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
    Resume Next
    
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
    'Check packet is complete
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Get data and update form
    UserGLD = incomingData.ReadLong()
    
    frmMain.GldLbl.Caption = PonerPuntos(UserGLD)

    
    Exit Sub

HandleUpdateGold_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateGold", Erl)
    Resume Next
    
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
    'Check packet is complete
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
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
    Resume Next
    
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
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserMap = incomingData.ReadInteger()
    
    'TODO: Once on-the-fly editor is implemented check for map version before loading....
    'For now we just drop it
    Call incomingData.ReadInteger
    
    If bRain Then
        If Not MapDat.LLUVIA Then
            '  Call Audio.StopWave(RainBufferIndex)
            ' RainBufferIndex = 0
            frmMain.IsPlaying = PlayLoop.plNone

        End If

    End If

    If frmComerciar.Visible Then
        Unload frmComerciar
    End If
    
    If frmBancoObj.Visible Then
        Unload frmBancoObj
    End If
    
    If FrmShop.Visible Then
        Unload FrmShop
    End If
        
    If frmEstadisticas.Visible Then
        Unload frmEstadisticas
    End If
    
    If frmHerrero.Visible Then
        Unload frmHerrero
    End If
    
    If FrmSastre.Visible Then
        Unload FrmSastre
    End If

    If frmAlqui.Visible Then
        Unload frmAlqui
    End If

    If frmCarp.Visible Then
        Unload frmCarp
    End If
    
    If FrmGrupo.Visible Then
        Unload FrmGrupo
    End If
    
    If FrmCorreo.Visible Then
        Unload FrmCorreo
    End If
    
    If frmGoliath.Visible Then
        Unload frmGoliath
    End If
       
    If FrmViajes.Visible Then
        Unload FrmViajes
    End If
    
    If frmCantidad.Visible Then
        Unload frmCantidad
    End If
    
    If FrmRanking.Visible Then
        Unload FrmRanking
    End If
    
    If frmMapaGrande.Visible Then
        Call CalcularPosicionMAPA
    End If

    Call SwitchMap(UserMap)

    
    Exit Sub

HandleChangeMap_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChangeMap", Erl)
    Resume Next
    
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
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Remove char from old position
    If MapData(UserPos.x, UserPos.y).charindex = UserCharIndex Then
        MapData(UserPos.x, UserPos.y).charindex = 0

    End If
    
    'Set new pos
    UserPos.x = incomingData.ReadByte()
    UserPos.y = incomingData.ReadByte()

    'Set char
    MapData(UserPos.x, UserPos.y).charindex = UserCharIndex
    charlist(UserCharIndex).Pos = UserPos
        
    'Are we under a roof?
    bTecho = HayTecho(UserPos.x, UserPos.y)
                
    'Update pos label and minimap
    frmMain.Coord.Caption = UserMap & "-" & UserPos.x & "-" & UserPos.y
    Call frmMain.SetMinimapPosition(0, UserPos.x, UserPos.y)

    Call RefreshAllChars
    
    Exit Sub

HandlePosUpdate_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePosUpdate", Erl)
    Resume Next
    
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
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
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
    Resume Next
    
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
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CRIATURA_1 & PonerPuntos(incomingData.ReadLong()) & MENSAJE_2, 255, 0, 0, True, False, False)

    
    Exit Sub

HandleUserHitNPC_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUserHitNPC", Erl)
    Resume Next
    
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
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & charlist(incomingData.ReadInteger()).nombre & MENSAJE_ATAQUE_FALLO, 255, 0, 0, True, False, False)

    
    Exit Sub

HandleUserAttackedSwing_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUserAttackedSwing", Erl)
    Resume Next
    
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
    If incomingData.length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim attacker As String
    
    Dim intt     As Integer
    
    intt = incomingData.ReadInteger()
    'attacker = charlist().Nombre
    
    Dim Pos As String

    Pos = InStr(charlist(intt).nombre, "<")
    
    If Pos = 0 Then Pos = Len(charlist(intt).nombre) + 2
    
    attacker = Left$(charlist(intt).nombre, Pos - 2)
    
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
    Resume Next
    
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
    If incomingData.length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim victim As String
    
    Dim intt   As Integer
    
    intt = incomingData.ReadInteger()
    'attacker = charlist().Nombre
    
    Dim Pos As String

    Pos = InStr(charlist(intt).nombre, "<")
    
    If Pos = 0 Then Pos = Len(charlist(intt).nombre) + 2
    
    victim = Left$(charlist(intt).nombre, Pos - 2)
    
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
    Resume Next
    
End Sub

''
' Handles the ChatOverHead message.

Private Sub HandleChatOverHead()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 12 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim chat       As String

    Dim charindex  As Integer

    Dim r          As Byte

    Dim G          As Byte

    Dim B          As Byte

    Dim colortexto As Long

    Dim QueEs      As String

    chat = buffer.ReadASCIIString()
    charindex = buffer.ReadInteger()
    
    r = buffer.ReadByte()
    G = buffer.ReadByte()
    B = buffer.ReadByte()
    
    colortexto = vbColor_2_Long(buffer.ReadLong())

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
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)

    Exit Sub
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

Private Sub HandleTextOverChar()

    If incomingData.length < 9 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim chat      As String

    Dim charindex As Integer

    Dim Color     As Long
    
    chat = buffer.ReadASCIIString()
    charindex = buffer.ReadInteger()
    
    Color = buffer.ReadLong()
    
    Call SetCharacterDialogFx(charindex, chat, RGBA_From_vbColor(Color))
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)

errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

Private Sub HandleTextOverTile()

    If incomingData.length < 11 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim Text      As String

    Dim x As Integer, y As Integer

    Dim Color     As Long
    
    Text = buffer.ReadASCIIString()
    x = buffer.ReadInteger()
    y = buffer.ReadInteger()
    Color = buffer.ReadLong()
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
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

errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

Private Sub HandleTextCharDrop()

    If incomingData.length < 9 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim Text      As String

    Dim charindex As Integer

    Dim Color     As Long
    
    Text = buffer.ReadASCIIString()
    charindex = buffer.ReadInteger()
    Color = buffer.ReadLong()
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
    If charindex = 0 Then Exit Sub

    Dim x As Integer, y As Integer, OffsetX As Integer, OffsetY As Integer
    
    With charlist(charindex)
        x = .Pos.x
        y = .Pos.y
        
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

errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the ConsoleMessage message.

Private Sub HandleConsoleMessage()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue
    Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim chat      As String
    Dim fontIndex As Integer
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

    chat = buffer.ReadASCIIString()
    fontIndex = buffer.ReadByte()
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
    If ChatGlobal = 0 And fontIndex = FontTypeNames.FONTTYPE_GLOBAL Then Exit Sub

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
            chat = "------------< Información del hechizo >------------" & vbCrLf & _
                    "Nombre: " & HechizoData(Hechizo).nombre & vbCrLf & _
                    "Descripción: " & HechizoData(Hechizo).desc & vbCrLf & _
                    "Skill requerido: " & HechizoData(Hechizo).MinSkill & " de magia." & vbCrLf & _
                    "Mana necesario: " & HechizoData(Hechizo).ManaRequerido & " puntos." & vbCrLf & _
                    "Stamina necesaria: " & HechizoData(Hechizo).StaRequerido & " puntos."

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

            Dim id    As Integer
            Dim extra As String

            id = ReadField(2, chat, Asc("*"))
            extra = ReadField(3, chat, Asc("*"))
                
            chat = Locale_Parse_ServerMessage(id, extra)
           
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

        With FontTypes(fontIndex)
            Call AddtoRichTextBox(frmMain.RecTxt, chat, .red, .green, .blue, .bold, .italic)
        End With

    End If
    
    Exit Sub
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

Private Sub HandleLocaleMsg()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim chat      As String

    Dim fontIndex As Integer

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

    Dim id        As Integer

    id = buffer.ReadInteger()
    chat = buffer.ReadASCIIString()
    fontIndex = buffer.ReadByte()
    
    ' If Not CHATGLOBAL And fontIndex = FontTypeNames.FONTTYPE_GLOBAL Then
    '    Call incomingData.CopyBuffer(Buffer)
    '    Set Buffer = Nothing
    '    Exit Sub
    ' End If
   
    chat = Locale_Parse_ServerMessage(id, chat)
    
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

        With FontTypes(fontIndex)
            Call AddtoRichTextBox(frmMain.RecTxt, chat, .red, .green, .blue, .bold, .italic)

        End With

    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the GuildChat message.

Private Sub HandleGuildChat()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/07/08 (NicoNZ)
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim chat As String

    Dim str  As String

    Dim r    As Byte

    Dim G    As Byte

    Dim B    As Byte

    Dim tmp  As Integer

    Dim Cont As Integer
    
    chat = buffer.ReadASCIIString()
    
    Rem If Not DialogosClanes.Activo Then
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
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the ShowMessageBox message.

Private Sub HandleShowMessageBox()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim mensaje As String

    mensaje = buffer.ReadASCIIString()

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
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

Private Sub HandleMostrarCuenta()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 1 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
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
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

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
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    userIndex = incomingData.ReadInteger()

    
    Exit Sub

HandleUserIndexInServer_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUserIndexInServer", Erl)
    Resume Next
    
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
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserCharIndex = incomingData.ReadInteger()
    UserPos = charlist(UserCharIndex).Pos
    
    'Are we under a roof?
    bTecho = HayTecho(UserPos.x, UserPos.y)
    
    LastMove = FrameTime
    
    frmMain.Coord.Caption = UserMap & "-" & UserPos.x & "-" & UserPos.y
    Call frmMain.SetMinimapPosition(0, UserPos.x, UserPos.y)
    
    If frmMapaGrande.Visible Then
        Call CalcularPosicionMAPA
    End If
    
    Exit Sub

HandleUserCharIndexInServer_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUserCharIndexInServer", Erl)
    Resume Next
    
End Sub

''
' Handles the CharacterCreate message.

Private Sub HandleCharacterCreate()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 62 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
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
    
    charindex = buffer.ReadInteger()
    Body = buffer.ReadInteger()
    Head = buffer.ReadInteger()
    Heading = buffer.ReadByte()
    x = buffer.ReadByte()
    y = buffer.ReadByte()
    weapon = buffer.ReadInteger()
    shield = buffer.ReadInteger()
    helmet = buffer.ReadInteger()
    
    With charlist(charindex)
        'Call SetCharacterFx(charindex, buffer.ReadInteger(), buffer.ReadInteger())
        .FxIndex = buffer.ReadInteger
        
        buffer.ReadInteger 'Ignore loops
        
        If .FxIndex > 0 Then
            Call InitGrh(.fX, FxData(.FxIndex).Animacion)
        End If
        
        Dim NombreYClan As String
        NombreYClan = buffer.ReadASCIIString()
        
        Dim Pos As Integer
        Pos = InStr(NombreYClan, "<")
        If Pos = 0 Then Pos = InStr(NombreYClan, "[")
        If Pos = 0 Then Pos = Len(NombreYClan) + 2
        
        .nombre = Left$(NombreYClan, Pos - 2)
        .clan = mid$(NombreYClan, Pos)
        
        .status = buffer.ReadByte()
        
        privs = buffer.ReadByte()
        ParticulaFx = buffer.ReadByte()
        .Head_Aura = buffer.ReadASCIIString()
        .Arma_Aura = buffer.ReadASCIIString()
        .Body_Aura = buffer.ReadASCIIString()
        .DM_Aura = buffer.ReadASCIIString()
        .RM_Aura = buffer.ReadASCIIString()
        .Otra_Aura = buffer.ReadASCIIString()
        .Escudo_Aura = buffer.ReadASCIIString()
        .Speeding = buffer.ReadSingle()
        
        Dim FlagNpc As Byte
        FlagNpc = buffer.ReadByte()
        
        .EsNpc = FlagNpc > 0
        .EsMascota = FlagNpc = 2
        
        .Donador = buffer.ReadByte()
        .appear = buffer.ReadByte()
        appear = .appear
        .group_index = buffer.ReadInteger()
        .clan_index = buffer.ReadInteger()
        .clan_nivel = buffer.ReadByte()
        .UserMinHp = buffer.ReadLong()
        .UserMaxHp = buffer.ReadLong()
        .simbolo = buffer.ReadByte()
        .Idle = buffer.ReadBoolean()
        .Navegando = buffer.ReadBoolean()
        
        If (.Pos.x <> 0 And .Pos.y <> 0) Then
            If MapData(.Pos.x, .Pos.y).charindex = charindex Then
                'Erase the old character from map
                MapData(charlist(charindex).Pos.x, charlist(charindex).Pos.y).charindex = 0

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
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

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
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
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
    Resume Next
    
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
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
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
    Resume Next
    
End Sub

''
' Handles the ForceCharMove message.

Private Sub HandleForceCharMove()
    
    On Error GoTo HandleForceCharMove_Err
    
    
    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim Direccion As Byte: Direccion = incomingData.ReadByte()
    
    Moviendose = True
    
    Call MainTimer.Restart(TimersIndex.Walk)

    Call Char_Move_by_Head(UserCharIndex, Direccion)
    Call MoveScreen(Direccion)
    
    Call frmMain.SetMinimapPosition(0, UserPos.x, UserPos.y)
    
    frmMain.Coord.Caption = UserMap & "-" & UserPos.x & "-" & UserPos.y

    If frmMapaGrande.Visible Then
        Call CalcularPosicionMAPA
    End If
    
    Call RefreshAllChars
    
    
    Exit Sub

HandleForceCharMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleForceCharMove", Erl)
    Resume Next
    
End Sub

''
' Handles the CharacterChange message.

Private Sub HandleCharacterChange()
    
    On Error GoTo HandleCharacterChange_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 19 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
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
    
    Call RefreshAllChars

    
    Exit Sub

HandleCharacterChange_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCharacterChange", Erl)
    Resume Next
    
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
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim x        As Byte

    Dim y        As Byte

    Dim OBJIndex As Integer

    Dim Color    As RGBA

    Dim Rango    As Byte

    Dim id       As Long
    
    x = incomingData.ReadByte()
    y = incomingData.ReadByte()
    
    OBJIndex = incomingData.ReadInteger()
    
    MapData(x, y).ObjGrh.GrhIndex = ObjData(OBJIndex).GrhIndex
    
    MapData(x, y).OBJInfo.OBJIndex = OBJIndex
    
    Call InitGrh(MapData(x, y).ObjGrh, MapData(x, y).ObjGrh.GrhIndex)
    
    If ObjData(OBJIndex).CreaLuz <> "" Then
        Call Long_2_RGBA(Color, Val(ReadField(2, ObjData(OBJIndex).CreaLuz, Asc(":"))))
        Rango = Val(ReadField(1, ObjData(OBJIndex).CreaLuz, Asc(":")))
        MapData(x, y).luz.Color = Color
        MapData(x, y).luz.Rango = Rango
        
        If Rango < 100 Then
            id = x & y
            LucesCuadradas.Light_Create x, y, Color, Rango, id
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
    Resume Next
    
End Sub

Private Sub HandleFxPiso()
    
    On Error GoTo HandleFxPiso_Err
    

    '***************************************************
    'Ladder
    '30/5/10
    '***************************************************
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
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
    Resume Next
    
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
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim x  As Byte

    Dim y  As Byte

    Dim id As Long
    
    x = incomingData.ReadByte()
    y = incomingData.ReadByte()
    
    If ObjData(MapData(x, y).OBJInfo.OBJIndex).CreaLuz <> "" Then
        id = LucesCuadradas.Light_Find(x & y)
        LucesCuadradas.Light_Remove id
        MapData(x, y).luz.Color = COLOR_EMPTY
        MapData(x, y).luz.Rango = 0
        LucesCuadradas.Light_Render_All

    End If
    
    MapData(x, y).ObjGrh.GrhIndex = 0
    
    If ObjData(MapData(x, y).OBJInfo.OBJIndex).CreaParticulaPiso <> 0 Then
        Graficos_Particulas.Particle_Group_Remove (MapData(x, y).particle_group)

    End If

    
    Exit Sub

HandleObjectDelete_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleObjectDelete", Erl)
    Resume Next
    
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
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim x As Byte, y As Byte, B As Byte
    
    x = incomingData.ReadByte()
    y = incomingData.ReadByte()
    B = incomingData.ReadByte()

    MapData(x, y).Blocked = MapData(x, y).Blocked And Not eBlock.ALL_SIDES
    MapData(x, y).Blocked = MapData(x, y).Blocked Or B

    
    Exit Sub

HandleBlockPosition_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBlockPosition", Erl)
    Resume Next
    
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
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    Dim currentMidi As Byte

    'Remove packet ID
    Call incomingData.ReadByte
    
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
    Resume Next
    
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
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
        
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
    Resume Next
    
End Sub

Private Sub HandlePosLLamadaDeClan()
    
    On Error GoTo HandlePosLLamadaDeClan_Err
    

    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 08/14/07
    'Last Modified by: Rapsodius
    'Added support for 3D Sounds.
    '***************************************************
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
        
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
    Resume Next
    
End Sub

Private Sub HandleCharUpdateHP()
    
    On Error GoTo HandleCharUpdateHP_Err
    

    '***************************************************
    'Autor: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 08/14/07
    'Last Modified by: Rapsodius
    'Added support for 3D Sounds.
    '***************************************************
    If incomingData.length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
        
    Dim charindex As Integer

    Dim minhp     As Integer

    Dim maxhp     As Integer
    
    charindex = incomingData.ReadInteger()
    minhp = incomingData.ReadInteger()
    maxhp = incomingData.ReadInteger()

    charlist(charindex).UserMinHp = minhp
    charlist(charindex).UserMaxHp = maxhp
    
    ' Call Audio.PlayWave(CStr(wave) & ".wav", srcX, srcY)
    
    Exit Sub

HandleCharUpdateHP_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCharUpdateHP", Erl)
    Resume Next
    
End Sub

Private Sub HandleArmaMov()
    
    On Error GoTo HandleArmaMov_Err
    

    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    'Remove packet ID
    Call incomingData.ReadByte

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
    Resume Next
    
End Sub

Private Sub HandleEscudoMov()
    
    On Error GoTo HandleEscudoMov_Err
    

    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    'Remove packet ID
    Call incomingData.ReadByte

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
    Resume Next
    
End Sub

''
' Handles the GuildList message.

Private Sub HandleGuildList()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue
    Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    'Clear guild's list
    frmGuildAdm.guildslist.Clear
    
    Dim guildsStr As String: guildsStr = buffer.ReadASCIIString()
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
    If Len(guildsStr) > 0 Then

        Dim guilds() As String: guilds = Split(guildsStr, SEPARATOR)
        
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
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

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
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim x As Byte

    Dim y As Byte
    
    x = incomingData.ReadByte()
    y = incomingData.ReadByte()
        
    Call CambioDeArea(x, y)

    
    Exit Sub

HandleAreaChanged_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleAreaChanged", Erl)
    Resume Next
    
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
    
    Call incomingData.ReadByte
    
    pausa = Not pausa

    
    Exit Sub

HandlePauseToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePauseToggle", Erl)
    Resume Next
    
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
    
    
    Call incomingData.ReadByte
    
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
    Resume Next
    
End Sub

Private Sub HandleTrofeoToggleOn()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleTrofeoToggleOn_Err
    
    
    Call incomingData.ReadByte

    MostrarTrofeo = True
  
    
    Exit Sub

HandleTrofeoToggleOn_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleTrofeoToggleOn", Erl)
    Resume Next
    
End Sub

Private Sub HandleTrofeoToggleOff()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleTrofeoToggleOff_Err
    
    
    Call incomingData.ReadByte

    MostrarTrofeo = False
  
    
    Exit Sub

HandleTrofeoToggleOff_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleTrofeoToggleOff", Erl)
    Resume Next
    
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
    If incomingData.length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
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
    Resume Next
    
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
    If incomingData.length < 27 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
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
    Resume Next
    
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
    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte

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
    Resume Next
    
End Sub

''
' Handles the ChangeInventorySlot message.

Private Sub HandleChangeInventorySlot()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 12 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue
    Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim Slot              As Byte
    Dim OBJIndex          As Integer
    Dim Name              As String
    Dim Amount            As Integer
    Dim Equipped          As Boolean
    Dim GrhIndex          As Long
    Dim ObjType           As Byte
    Dim MaxHit            As Integer
    Dim MinHit            As Integer
    Dim MaxDef            As Integer
    Dim MinDef            As Integer
    Dim Value             As Single
    Dim podrausarlo       As Byte

    Slot = buffer.ReadByte()
    OBJIndex = buffer.ReadInteger()
    Amount = buffer.ReadInteger()
    Equipped = buffer.ReadBoolean()
    Value = buffer.ReadSingle()
    podrausarlo = buffer.ReadByte()

    Call incomingData.CopyBuffer(buffer)

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

    Call frmComerciar.InvComUsu.SetItem(Slot, OBJIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, MinDef, Value, Name, podrausarlo)

    Call frmBancoObj.InvBankUsu.SetItem(Slot, OBJIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, MinDef, Value, Name, podrausarlo)
    
    
    Call frmBancoCuenta.InvBankUsuCuenta.SetItem(Slot, OBJIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, MinDef, Value, Name, podrausarlo)

    Exit Sub
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long: Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

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

    Call incomingData.ReadByte
    
    UserInvUnlocked = incomingData.ReadByte
    
    For i = 1 To UserInvUnlocked
    
        frmMain.imgInvLock(i - 1).Picture = LoadInterface("inventoryunlocked.bmp")
    
    Next i
    
    'Call Inventario.DrawInventory
    
    
    Exit Sub

HandleInventoryUnlockSlots_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleInventoryUnlockSlots", Erl)
    Resume Next
    
End Sub

Private Sub HandleRefreshAllInventorySlot()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
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
    
    todo = buffer.ReadASCIIString()
    
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
    
    ' For i = 1 To 10
    Rem slotNum(1) = Right$(slotNum(1), Len(slotNum(1)))

    'slot = ReadField(1, slotNum(1), Asc("@")) 'Nick
    ' OBJIndex = ReadField(2, slotNum(1), Asc("@")) 'Nick
    ' Name = ReadField(3, slotNum(1), Asc("@")) 'Nick
    ' Amount = ReadField(4, slotNum(1), Asc("@")) 'Nick
    ' Equipped = ReadField(5, slotNum(1), Asc("@")) 'Nick
    'grhindex = ReadField(6, slotNum(1), Asc("@")) 'Nick
    ' ObjType = ReadField(7, slotNum(1), Asc("@")) 'Nick
    ' MaxHit = ReadField(8, slotNum(1), Asc("@")) 'Nick
    ' MinHit = ReadField(9, slotNum(1), Asc("@")) 'Nick
    ' defense = ReadField(10, slotNum(1), Asc("@")) 'Nick
    '  value = ReadField(11, slotNum(1), Asc("@")) 'Nick

    ' Next i
     
    Rem  Call Inventario.SetItem(slot, OBJIndex, Amount, Equipped, grhindex, ObjType, MaxHit, MinHit, defense, value, Name, PuedeUsar)
    
    ' For i = 1 To 25
    ' slotNum(i) = Right$(slotNum(i), Len(slotNum(i)))
    '  slot = ReadField(1, slotNum(i), Asc("@")) 'Nick
    '   OBJIndex = ReadField(2, slotNum(i), Asc("@")) 'Nick
    '  Name = ReadField(3, slotNum(i), Asc("@")) 'Nick
    '  Amount = ReadField(4, slotNum(i), Asc("@")) 'Nick
    '  Equipped = ReadField(5, slotNum(i), Asc("@")) 'Nick
    '  grhindex = ReadField(6, slotNum(i), Asc("@")) 'Nick
    '    ObjType = ReadField(7, slotNum(i), Asc("@")) 'Nick
    '  MaxHit = ReadField(8, slotNum(i), Asc("@")) 'Nick
    '  MinHit = ReadField(9, slotNum(i), Asc("@")) 'Nick
    '   defense = ReadField(10, slotNum(i), Asc("@")) 'Nick
    '   value = ReadField(11, slotNum(i), Asc("@")) 'Nick
    '  Debug.Print Name
    Rem  Call Inventario.SetItem(i, OBJIndex, Amount, Equipped, grhindex, ObjType, MaxHit, MinHit, defense, value, Name, PuedeUsar)
    
    ' If slot < 16 Then

    '  Else
    'If we got here then packet is complete, copy data back to original queue

    '  End If
    Call incomingData.CopyBuffer(buffer)
     
    Exit Sub
     
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the ChangeBankSlot message.

Private Sub HandleChangeBankSlot()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 11 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue
    Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim Slot As Byte: Slot = buffer.ReadByte()
    
    Dim BankSlot As Inventory
    
    With BankSlot
    
        .OBJIndex = buffer.ReadInteger()
        .Name = ObjData(.OBJIndex).Name
        .Amount = buffer.ReadInteger()
        .GrhIndex = ObjData(.OBJIndex).GrhIndex
        .ObjType = ObjData(.OBJIndex).ObjType
        .MaxHit = ObjData(.OBJIndex).MaxHit
        .MinHit = ObjData(.OBJIndex).MinHit
        .Def = ObjData(.OBJIndex).MaxDef
        .Valor = buffer.ReadLong()
        .PuedeUsar = buffer.ReadByte()
        
        Call frmBancoObj.InvBoveda.SetItem(Slot, .OBJIndex, .Amount, .Equipped, .GrhIndex, .ObjType, .MaxHit, .MinHit, .Def, .Valor, .Name, .PuedeUsar)

    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the ChangeSpellSlot message

Private Sub HandleChangeSpellSlot()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim Slot     As Byte

    Dim Index    As Byte

    Dim cooldown As Integer

    Slot = buffer.ReadByte()
    
    UserHechizos(Slot) = buffer.ReadInteger()
    Index = buffer.ReadByte()

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
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
    Exit Sub
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

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
    If incomingData.length < 1 + NUMATRIBUTES Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
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
    Resume Next
    
End Sub

''
' Handles the BlacksmithWeapons message.

Private Sub HandleBlacksmithWeapons()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 9 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim count As Integer

    Dim i     As Long

    Dim tmp   As String
    
    count = buffer.ReadInteger()
    
    Call frmHerrero.lstArmas.Clear
    
    For i = 1 To count
        ArmasHerrero(i).Index = buffer.ReadInteger()
        ' tmp = ObjData(ArmasHerrero(i).Index).name        'Get the object's name
        ArmasHerrero(i).LHierro = buffer.ReadInteger()  'The iron needed
        ArmasHerrero(i).LPlata = buffer.ReadInteger()    'The silver needed
        ArmasHerrero(i).LOro = buffer.ReadInteger()    'The gold needed
        
        ' Call frmHerrero.lstArmas.AddItem(tmp)
    Next i
    
    For i = i To UBound(ArmasHerrero())
        ArmasHerrero(i).Index = 0
    Next i
    
    i = 0
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
    Exit Sub
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the BlacksmithArmors message.

Private Sub HandleBlacksmithArmors()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim count As Integer

    Dim i     As Long

    Dim tmp   As String
    
    count = buffer.ReadInteger()
    
    'Call frmHerrero.lstArmaduras.Clear
    
    For i = 1 To count
        tmp = buffer.ReadASCIIString()         'Get the object's name
        DefensasHerrero(i).LHierro = buffer.ReadInteger()   'The iron needed
        DefensasHerrero(i).LPlata = buffer.ReadInteger()   'The silver needed
        DefensasHerrero(i).LOro = buffer.ReadInteger()   'The gold needed
        
        ' Call frmHerrero.lstArmaduras.AddItem(tmp)
        DefensasHerrero(i).Index = buffer.ReadInteger()
    Next i
    
    For i = i To UBound(DefensasHerrero())
        DefensasHerrero(i).Index = 0
    Next i
    
    Dim A As Byte

    Dim e As Byte

    Dim c As Byte

    A = 0
    e = 0
    c = 0
    
    For i = 1 To UBound(DefensasHerrero())
    
        If DefensasHerrero(i).Index = 0 Then Exit For
        If ObjData(DefensasHerrero(i).Index).ObjType = 3 Then
           
            ArmadurasHerrero(A).Index = DefensasHerrero(i).Index
            ArmadurasHerrero(A).LHierro = DefensasHerrero(i).LHierro
            ArmadurasHerrero(A).LPlata = DefensasHerrero(i).LPlata
            ArmadurasHerrero(A).LOro = DefensasHerrero(i).LOro
            A = A + 1

        End If
        
        If ObjData(DefensasHerrero(i).Index).ObjType = 16 Then
            EscudosHerrero(e).Index = DefensasHerrero(i).Index
            EscudosHerrero(e).LHierro = DefensasHerrero(i).LHierro
            EscudosHerrero(e).LPlata = DefensasHerrero(i).LPlata
            EscudosHerrero(e).LOro = DefensasHerrero(i).LOro
            e = e + 1

        End If

        If ObjData(DefensasHerrero(i).Index).ObjType = 17 Then
            CascosHerrero(c).Index = DefensasHerrero(i).Index
            CascosHerrero(c).LHierro = DefensasHerrero(i).LHierro
            CascosHerrero(c).LPlata = DefensasHerrero(i).LPlata
            CascosHerrero(c).LOro = DefensasHerrero(i).LOro
            c = c + 1

        End If

    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
    Exit Sub
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the CarpenterObjects message.

Private Sub HandleCarpenterObjects()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim count As Integer

    Dim i     As Long

    Dim tmp   As String
    
    count = buffer.ReadByte()
    
    Call frmCarp.lstArmas.Clear
    
    For i = 1 To count
        ObjCarpintero(i) = buffer.ReadInteger()
        
        Call frmCarp.lstArmas.AddItem(ObjData(ObjCarpintero(i)).Name)
    Next i
    
    For i = i To UBound(ObjCarpintero())
        ObjCarpintero(i) = 0
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

Private Sub HandleSastreObjects()

    '***************************************************
    'Author: Ladder
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim count As Integer

    Dim i     As Long

    Dim tmp   As String
    
    count = buffer.ReadInteger()
    
    For i = i To UBound(ObjSastre())
        ObjSastre(i).Index = 0
    Next i
    
    i = 0
    
    For i = 1 To count
        ObjSastre(i).Index = buffer.ReadInteger()
        
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
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
Private Sub HandleAlquimiaObjects()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim count As Integer

    Dim i     As Long

    Dim tmp   As String
    
    Dim Obj   As Integer

    count = buffer.ReadInteger()
    
    Call frmAlqui.lstArmas.Clear
    
    For i = 1 To count
        Obj = buffer.ReadInteger()
        tmp = ObjData(Obj).Name        'Get the object's name

        ObjAlquimista(i) = Obj
        Call frmAlqui.lstArmas.AddItem(tmp)
    Next i
    
    For i = i To UBound(ObjAlquimista())
        ObjAlquimista(i) = 0
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

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
    
    Call incomingData.ReadByte
    
    UserDescansar = Not UserDescansar

    
    Exit Sub

HandleRestOK_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleRestOK", Erl)
    Resume Next
    
End Sub

''
' Handles the ErrorMessage message.

Private Sub HandleErrorMessage()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Call MsgBox(buffer.ReadASCIIString())
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

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
    
    Call incomingData.ReadByte
    
    UserCiego = True
    
    Call SetRGBA(global_light, 4, 4, 4)
    Call MapUpdateGlobalLight
    
    
    Exit Sub

HandleBlind_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBlind", Erl)
    Resume Next
    
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
    
    Call incomingData.ReadByte
    
    UserEstupido = True

    
    Exit Sub

HandleDumb_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleDumb", Erl)
    Resume Next
    
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
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim tmp As String

    Dim grh As Integer

    tmp = ObjData(buffer.ReadInteger()).Texto
    grh = buffer.ReadInteger()
    Call InitCartel(tmp, grh)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the ChangeNPCInventorySlot message.

Private Sub HandleChangeNPCInventorySlot()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 11 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As clsByteQueue
    Set buffer = New clsByteQueue
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim Slot As Byte: Slot = buffer.ReadByte()
    
    Dim SlotInv As NpCinV
    With SlotInv
        .OBJIndex = buffer.ReadInteger()
        .Name = ObjData(.OBJIndex).Name
        .Amount = buffer.ReadInteger()
        .Valor = buffer.ReadSingle()
        .GrhIndex = ObjData(.OBJIndex).GrhIndex
        .ObjType = ObjData(.OBJIndex).ObjType
        .MaxHit = ObjData(.OBJIndex).MaxHit
        .MinHit = ObjData(.OBJIndex).MinHit
        .Def = ObjData(.OBJIndex).MaxDef
        .PuedeUsar = buffer.ReadByte()
        
        Call frmComerciar.InvComNpc.SetItem(Slot, .OBJIndex, .Amount, 0, .GrhIndex, .ObjType, .MaxHit, .MinHit, .Def, .Valor, .Name, .PuedeUsar)
        
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
    Exit Sub
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

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
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
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
    Resume Next
    
End Sub

Private Sub HandleHora()
    '***************************************************
    
    On Error GoTo HandleHora_Err
    
    If incomingData.length < 9 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte

    HoraMundo = (timeGetTime And &H7FFFFFFF) - incomingData.ReadLong()
    DuracionDia = incomingData.ReadLong()
    
    If Not Connected Then
        Call RevisarHoraMundo(True)
    End If

    
    Exit Sub

HandleHora_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleHora", Erl)
    Resume Next
    
End Sub
 
Private Sub HandleLight()
    
    On Error GoTo HandleLight_Err
    
 
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
 
    Dim Color As String

    Call incomingData.ReadByte
    Color = incomingData.ReadASCIIString()

    'Call SetGlobalLight(Map_light_base)
 
    
    Exit Sub

HandleLight_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleLight", Erl)
    Resume Next
    
End Sub
 
Private Sub HandleFYA()
    
    On Error GoTo HandleFYA_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
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
    Resume Next
    
End Sub

Private Sub HandleUpdateNPCSimbolo()
    
    On Error GoTo HandleUpdateNPCSimbolo_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim NpcIndex As Integer

    Dim simbolo  As Byte
    
    NpcIndex = incomingData.ReadInteger()
    
    simbolo = incomingData.ReadByte()

    charlist(NpcIndex).simbolo = simbolo

    
    Exit Sub

HandleUpdateNPCSimbolo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateNPCSimbolo", Erl)
    Resume Next
    
End Sub

Private Sub HandleCerrarleCliente()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleCerrarleCliente_Err
    
    Call incomingData.ReadByte
    
    EngineRun = False

    Call CloseClient

    
    Exit Sub

HandleCerrarleCliente_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCerrarleCliente", Erl)
    Resume Next
    
End Sub

Private Sub HandleContadores()
    
    On Error GoTo HandleContadores_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 9 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
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
    Resume Next
    
End Sub

Private Sub HandleOxigeno()
    
    On Error GoTo HandleOxigeno_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
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
    Resume Next
    
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

    If incomingData.length < 10 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
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
    Resume Next
    
End Sub

Private Sub HandleMiniStats()
    
    On Error GoTo HandleMiniStats_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 30 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
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
    Resume Next
    
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
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    SkillPoints = incomingData.ReadInteger()

    
    Exit Sub

HandleLevelUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleLevelUp", Erl)
    Resume Next
    
End Sub

''
' Handles the AddForumMessage message.

Private Sub HandleAddForumMessage()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim title   As String

    Dim Message As String
    
    title = buffer.ReadASCIIString()
    Message = buffer.ReadASCIIString()
    
    'Call frmForo.List.AddItem(title)
    ' frmForo.Text(frmForo.List.ListCount - 1).Text = Message
    ' Call Load(frmForo.Text(frmForo.List.ListCount))
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

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
    
    Call incomingData.ReadByte
    
    ' If Not frmForo.Visible Then
    '   frmForo.Show , frmMain
    ' End If
    
    Exit Sub

HandleShowForumForm_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowForumForm", Erl)
    Resume Next
    
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
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim charindex As Integer
    
    charindex = incomingData.ReadInteger()
    charlist(charindex).Invisible = incomingData.ReadBoolean()
    charlist(charindex).TimerI = 0

    
    Exit Sub

HandleSetInvisible_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSetInvisible", Erl)
    Resume Next
    
End Sub

Private Sub HandleSetEscribiendo()
    
    On Error GoTo HandleSetEscribiendo_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim charindex As Integer
    
    charindex = incomingData.ReadInteger()
    charlist(charindex).Escribiendo = incomingData.ReadBoolean()

    
    Exit Sub

HandleSetEscribiendo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSetEscribiendo", Erl)
    Resume Next
    
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
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
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
    Resume Next
    
End Sub

''
' Handles the MeditateToggle message.

Private Sub HandleMeditateToggle()
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleMeditateToggle_Err
    
    Call incomingData.ReadByte
    
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
    Resume Next
    
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
    
    Call incomingData.ReadByte
    UserCiego = False
    
    Call RestaurarLuz
    
    Exit Sub

HandleBlindNoMore_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBlindNoMore", Erl)
    Resume Next
    
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
    
    Call incomingData.ReadByte
    
    UserEstupido = False

    
    Exit Sub

HandleDumbNoMore_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleDumbNoMore", Erl)
    Resume Next
    
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
    If incomingData.length < 1 + NUMSKILLS Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
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
    Resume Next
    
End Sub

''
' Handles the TrainerCreatureList message.

Private Sub HandleTrainerCreatureList()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim creatures() As String

    Dim i           As Long
    
    creatures = Split(buffer.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To UBound(creatures())
        Call frmEntrenador.lstCriaturas.AddItem(creatures(i))
    Next i

    frmEntrenador.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the GuildNews message.

Private Sub HandleGuildNews()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 12 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    ' Dim guildList() As String
    Dim List()      As String

    Dim i           As Long
    
    Dim ClanNivel   As Byte

    Dim expacu      As Integer

    Dim ExpNe       As Integer

    Dim guildList() As String
        
    frmGuildNews.news = buffer.ReadASCIIString()
    
    'Get list of existing guilds
    List = Split(buffer.ReadASCIIString(), SEPARATOR)
        
    'Empty the list
    Call frmGuildNews.guildslist.Clear
        
    For i = 0 To UBound(List())
        Call frmGuildNews.guildslist.AddItem(ReadField(1, List(i), Asc("-")))
    Next i
    
    'Get  guilds list member
    guildList = Split(buffer.ReadASCIIString(), SEPARATOR)
    
    Dim cantidad As String

    cantidad = CStr(UBound(guildList()) + 1)
        
    Call frmGuildNews.miembros.Clear
        
    For i = 0 To UBound(guildList())

        If i = 0 Then
            Call frmGuildNews.miembros.AddItem(guildList(i) & "(Lider)")
        Else
            Call frmGuildNews.miembros.AddItem(guildList(i))

        End If

        'Debug.Print guildList(i)
    Next i
    
    ClanNivel = buffer.ReadByte()
    expacu = buffer.ReadInteger()
    ExpNe = buffer.ReadInteger()
     
    With frmGuildNews
        .Frame4.Caption = "Total: " & cantidad & " miembros" '"Lista de miembros" ' - " & cantidad & " totales"
     
        .expcount.Caption = expacu & "/" & ExpNe
        .EXPBAR.Width = (((expacu + 1 / 100) / (ExpNe + 1 / 100)) * 2370)
        .nivel = "Nivel: " & ClanNivel
        
        ' frmMain.exp.Caption = UserExp & "/" & UserPasarNivel
        ' frmMain.ExpBar.Width = (((UserExp / 100) / (UserPasarNivel / 100)) * 165)
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
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the OfferDetails message.

Private Sub HandleOfferDetails()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Call frmUserRequest.recievePeticion(buffer.ReadASCIIString())
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the AlianceProposalsList message.

Private Sub HandleAlianceProposalsList()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim guildList() As String

    Dim i           As Long
    
    guildList = Split(buffer.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To UBound(guildList())
        Call frmPeaceProp.lista.AddItem(guildList(i))
    Next i
    
    frmPeaceProp.ProposalType = TIPO_PROPUESTA.ALIANZA
    Call frmPeaceProp.Show(vbModeless, frmMain)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the PeaceProposalsList message.

Private Sub HandlePeaceProposalsList()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim guildList() As String

    Dim i           As Long
    
    guildList = Split(buffer.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To UBound(guildList())
        Call frmPeaceProp.lista.AddItem(guildList(i))
    Next i
    
    frmPeaceProp.ProposalType = TIPO_PROPUESTA.PAZ
    Call frmPeaceProp.Show(vbModeless, frmMain)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the CharacterInfo message.

Private Sub HandleCharacterInfo()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 31 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
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
        
        .nombre.Caption = "Nombre: " & buffer.ReadASCIIString()
        .Raza.Caption = "Raza: " & ListaRazas(buffer.ReadByte())
        .Clase.Caption = "Clase: " & ListaClases(buffer.ReadByte())
        
        If buffer.ReadByte() = 1 Then
            .Genero.Caption = "Genero: Hombre"
        Else
            .Genero.Caption = "Genero: Mujer"

        End If
        
        .nivel.Caption = "Nivel: " & buffer.ReadByte()
        .oro.Caption = "Oro: " & buffer.ReadLong()
        .Banco.Caption = "Banco: " & buffer.ReadLong()
        
        ' Dim reputation As Long
        'reputation = buffer.ReadLong()
        
        '.reputacion.Caption = "Reputación: " & reputation
        
        .txtPeticiones.Text = buffer.ReadASCIIString()
        .guildactual.Caption = "Clan: " & buffer.ReadASCIIString()
        .txtMiembro.Text = buffer.ReadASCIIString()
        
        Dim armada As Boolean

        Dim caos   As Boolean
        
        armada = buffer.ReadBoolean()
        caos = buffer.ReadBoolean()
        
        If armada Then
            .ejercito.Caption = "Ejército: Armada Real"
        ElseIf caos Then
            .ejercito.Caption = "Ejército: Legión Oscura"

        End If
        
        .ciudadanos.Caption = "Ciudadanos asesinados: " & CStr(buffer.ReadLong())
        .Criminales.Caption = "Criminales asesinados: " & CStr(buffer.ReadLong())
        
        '   If reputation > 0 Then
        '   .status.Caption = " (Ciudadano)"
        '     .status.ForeColor = vbBlue
        ' Else
        '    .status.Caption = " (Criminal)"
        '    .status.ForeColor = vbRed
        '  End If
        
        Call .Show(vbModeless, frmMain)

    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the GuildLeaderInfo message.

Private Sub HandleGuildLeaderInfo()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 14 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim List() As String

    Dim i      As Long
    
    With frmGuildLeader
        'Get list of existing guilds
        List = Split(buffer.ReadASCIIString(), SEPARATOR)
        
        'Empty the list
        Call .guildslist.Clear
        
        For i = 0 To UBound(List())
            Call .guildslist.AddItem(ReadField(1, List(i), Asc("-")))
        Next i
        
        'Get list of guild's members
        List = Split(buffer.ReadASCIIString(), SEPARATOR)
        .miembros.Caption = "El clan cuenta con " & CStr(UBound(List()) + 1) & " miembros."
        
        'Empty the list
        Call .members.Clear
        
        For i = 0 To UBound(List())
            Call .members.AddItem(List(i))
        Next i
        
        .txtguildnews = buffer.ReadASCIIString()
        
        'Get list of join requests
        List = Split(buffer.ReadASCIIString(), SEPARATOR)
        
        'Empty the list
        Call .solicitudes.Clear
        
        For i = 0 To UBound(List())
            Call .solicitudes.AddItem(List(i))
        Next i
        
        Dim expacu As Integer

        Dim ExpNe  As Integer

        Dim nivel  As Byte
         
        nivel = buffer.ReadByte()
        .nivel = "Nivel: " & nivel
        
        expacu = buffer.ReadInteger()
        ExpNe = buffer.ReadInteger()
        '.expacu = "Experiencia acumulada: " & expacu
        'barra
        .expcount.Caption = expacu & "/" & ExpNe
        .EXPBAR.Width = expacu / ExpNe * 2370
        
        If ExpNe > 0 Then
       
            .porciento.Caption = Round(expacu / ExpNe * 100#, 0) & "%"
        Else
            .porciento.Caption = "¡Nivel máximo!"
            .expcount.Caption = "¡Nivel máximo!"

        End If

        Select Case nivel

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
        
        .Show , frmMain

    End With

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the GuildDetails message.

Private Sub HandleGuildDetails()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 16 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    With frmGuildBrief

        If Not .EsLeader Then

        End If
        
        .nombre.Caption = "Nombre:" & buffer.ReadASCIIString()
        .fundador.Caption = "Fundador:" & buffer.ReadASCIIString()
        .creacion.Caption = "Fecha de creacion:" & buffer.ReadASCIIString()
        .lider.Caption = "Líder:" & buffer.ReadASCIIString()
        .miembros.Caption = "Miembros:" & buffer.ReadInteger()
        
        .lblAlineacion.Caption = "Alineación: " & buffer.ReadASCIIString()
        
        .desc.Text = buffer.ReadASCIIString()
        .nivel.Caption = "Nivel de clan: " & buffer.ReadByte()

    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
    frmGuildBrief.Show vbModeless, frmMain
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

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
    
    Call incomingData.ReadByte
    
    CreandoClan = True
    frmGuildDetails.Show , frmMain

    
    Exit Sub

HandleShowGuildFundationForm_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowGuildFundationForm", Erl)
    Resume Next
    
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
    
    Call incomingData.ReadByte
    
    UserParalizado = Not UserParalizado

    
    Exit Sub

HandleParalizeOK_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleParalizeOK", Erl)
    Resume Next
    
End Sub

Private Sub HandleInmovilizadoOK()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleInmovilizadoOK_Err
    
    Call incomingData.ReadByte
    
    UserInmovilizado = Not UserInmovilizado

    
    Exit Sub

HandleInmovilizadoOK_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleInmovilizadoOK", Erl)
    Resume Next
    
End Sub

''
' Handles the ShowUserRequest message.

Private Sub HandleShowUserRequest()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Call frmUserRequest.recievePeticion(buffer.ReadASCIIString())
    Call frmUserRequest.Show(vbModeless, frmMain)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the ChangeUserTradeSlot message.

Private Sub HandleChangeUserTradeSlot()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 22 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    Dim miOferta As Boolean
    
    miOferta = buffer.ReadBoolean
    Dim i As Byte
    Dim nombreItem As String
    Dim cantidad As Integer
    Dim grhItem As Long
    Dim OBJIndex As Integer
    If miOferta Then
        Dim OroAEnviar As Long
        OroAEnviar = buffer.ReadLong
        frmComerciarUsu.lblOroMiOferta.Caption = PonerPuntos(OroAEnviar)
        frmComerciarUsu.lblMyGold.Caption = PonerPuntos(Val(frmMain.GldLbl.Caption - OroAEnviar))
        For i = 1 To 6
            With OtroInventario(i)
                OBJIndex = buffer.ReadInteger
                nombreItem = buffer.ReadASCIIString
                grhItem = buffer.ReadLong
                cantidad = buffer.ReadLong
                If cantidad > 0 Then
                    Call frmComerciarUsu.InvUserSell.SetItem(i, OBJIndex, cantidad, 0, grhItem, 0, 0, 0, 0, 0, nombreItem, 0)
                End If
            End With
        Next i
        
        Call frmComerciarUsu.InvUserSell.ReDraw
    Else
        frmComerciarUsu.lblOro.Caption = PonerPuntos(buffer.ReadLong)
       ' frmComerciarUsu.List2.Clear
        For i = 1 To 6
            
            With OtroInventario(i)
                 OBJIndex = buffer.ReadInteger
                nombreItem = buffer.ReadASCIIString
                grhItem = buffer.ReadLong
                cantidad = buffer.ReadLong
                If cantidad > 0 Then
                    Call frmComerciarUsu.InvOtherSell.SetItem(i, OBJIndex, cantidad, 0, grhItem, 0, 0, 0, 0, 0, nombreItem, 0)
                End If
            End With
        Next i
        
        Call frmComerciarUsu.InvOtherSell.ReDraw
    
    End If
    
    frmComerciarUsu.lblEstadoResp.Visible = False
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the SpawnList message.

Private Sub HandleSpawnList()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim creatureList() As String

    creatureList = Split(buffer.ReadASCIIString(), SEPARATOR)

    Call frmSpawnList.FillList

    frmSpawnList.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the ShowSOSForm message.

Private Sub HandleShowSOSForm()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim sosList()      As String

    Dim i              As Long

    Dim nombre         As String

    Dim Consulta       As String

    Dim TipoDeConsulta As String
    
    sosList = Split(buffer.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To UBound(sosList())
        nombre = ReadField(1, sosList(i), Asc("Ø"))
        Consulta = ReadField(2, sosList(i), Asc("Ø"))
        TipoDeConsulta = ReadField(3, sosList(i), Asc("Ø"))
        frmPanelgm.List1.AddItem nombre & "(" & TipoDeConsulta & ")"
        frmPanelgm.List2.AddItem Consulta
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the ShowMOTDEditionForm message.

Private Sub HandleShowMOTDEditionForm()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    frmCambiaMotd.txtMotd.Text = buffer.ReadASCIIString()
    frmCambiaMotd.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

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
    
    Call incomingData.ReadByte
    frmPanelgm.txtHeadNumero = incomingData.ReadInteger
    frmPanelgm.txtBodyYo = incomingData.ReadInteger
    
    frmPanelgm.Show vbModeless, frmMain

    
    Exit Sub

HandleShowGMPanelForm_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowGMPanelForm", Erl)
    Resume Next
    
End Sub

Private Sub HandleShowFundarClanForm()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleShowFundarClanForm_Err
    
    Call incomingData.ReadByte
    
    CreandoClan = True
    frmGuildDetails.Show vbModeless, frmMain

    
    Exit Sub

HandleShowFundarClanForm_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowFundarClanForm", Erl)
    Resume Next
    
End Sub

''
' Handles the UserNameList message.

Private Sub HandleUserNameList()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim userList() As String

    Dim i          As Long
    
    userList = Split(buffer.ReadASCIIString(), SEPARATOR)
    
    If frmPanelgm.Visible Then
        frmPanelgm.cboListaUsus.Clear

        For i = 0 To UBound(userList())
            Call frmPanelgm.cboListaUsus.AddItem(userList(i))
        Next i

        If frmPanelgm.cboListaUsus.ListCount > 0 Then frmPanelgm.cboListaUsus.ListIndex = 0

    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

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
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim Time As Long

    Time = incomingData.ReadLong()
    'Call AddtoRichTextBox(frmMain.RecTxt, "El ping anterior seria " & (GetTickCount - pingTime) & " ms.", 255, 0, 0, True, False, False)
    'Call AddtoRichTextBox(frmMain.RecTxt, "El ping es " & (GetTickCount - time) & " ms.", 255, 0, 0, True, False, False)
    'timeGetTime -pingTime
    PingRender = (timeGetTime And &H7FFFFFFF) - Time
    pingTime = 0

    
    Exit Sub

HandlePong_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePong", Erl)
    Resume Next
    
End Sub

''
' Handles the UpdateTag message.

Private Sub HandleUpdateTagAndStatus()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 8 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim charindex   As Integer

    Dim status      As Byte

    Dim NombreYClan As String

    Dim group_index As Integer
    
    charindex = buffer.ReadInteger()
    status = buffer.ReadByte()
    NombreYClan = buffer.ReadASCIIString()
        
    Dim Pos As Integer
    Pos = InStr(NombreYClan, "<")
    If Pos = 0 Then Pos = InStr(NombreYClan, "[")
    If Pos = 0 Then Pos = Len(NombreYClan) + 2
    
    charlist(charindex).nombre = Left$(NombreYClan, Pos - 2)
    charlist(charindex).clan = mid$(NombreYClan, Pos)
    
    group_index = buffer.ReadInteger()
    
    'Update char status adn tag!
    charlist(charindex).status = status
    
    charlist(charindex).group_index = group_index
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Writes the "LoginExistingChar" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoginExistingChar()
    
    On Error GoTo WriteLoginExistingChar_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "LoginExistingChar" message to the outgoing data buffer
    '***************************************************
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.LoginExistingChar)
        Call .WriteASCIIString(CuentaEmail)
        Call .WriteASCIIString(SEncriptar(CuentaPassword))
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(MacAdress)  'Seguridad
        Call .WriteLong(HDserial)  'SeguridadHDserial
        Call .WriteASCIIString(CheckMD5)
        
    End With

    
    Exit Sub

WriteLoginExistingChar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteLoginExistingChar", Erl)
    Resume Next
    
End Sub

''
' Writes the "LoginNewChar" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoginNewChar()
    
    On Error GoTo WriteLoginNewChar_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "LoginNewChar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.LoginNewChar)
        Call .WriteASCIIString(CuentaEmail)
        Call .WriteASCIIString(SEncriptar(CuentaPassword))
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
        Call .WriteASCIIString(UserName)
        Call .WriteByte(UserRaza)
        Call .WriteByte(UserSexo)
        Call .WriteByte(UserClase)
        Call .WriteInteger(MiCabeza)
        Call .WriteByte(UserHogar)
        Call .WriteASCIIString(MacAdress)  'Seguridad
        Call .WriteLong(HDserial)  'SeguridadHDserial
        Call .WriteASCIIString(CheckMD5)

    End With

    
    Exit Sub

WriteLoginNewChar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteLoginNewChar", Erl)
    Resume Next
    
End Sub

''
' Writes the "Talk" message to the outgoing data buffer.
'
' @param    chat The chat text to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTalk(ByVal chat As String)
    
    On Error GoTo WriteTalk_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Talk" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Talk)
        
        Call .WriteASCIIString(chat)

    End With

    
    Exit Sub

WriteTalk_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteTalk", Erl)
    Resume Next
    
End Sub

''
' Writes the "Yell" message to the outgoing data buffer.
'
' @param    chat The chat text to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteYell(ByVal chat As String)
    
    On Error GoTo WriteYell_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Yell" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Yell)
        
        Call .WriteASCIIString(chat)

    End With

    
    Exit Sub

WriteYell_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteYell", Erl)
    Resume Next
    
End Sub

''
' Writes the "Whisper" message to the outgoing data buffer.
'
' @param    charIndex The index of the char to whom to whisper.
' @param    chat The chat text to be sent to the user.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWhisper(ByVal nombre As String, ByVal chat As String)
    
    On Error GoTo WriteWhisper_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Whisper" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Whisper)
        
        Call .WriteASCIIString(nombre)
        
        Call .WriteASCIIString(chat)

    End With

    
    Exit Sub

WriteWhisper_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteWhisper", Erl)
    Resume Next
    
End Sub

''
' Writes the "Walk" message to the outgoing data buffer.
'
' @param    heading The direction in wich the user is moving.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWalk(ByVal Heading As E_Heading)
    
    On Error GoTo WriteWalk_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Walk" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Walk)
        
        Call .WriteByte(Heading)

    End With

    
    Exit Sub

WriteWalk_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteWalk", Erl)
    Resume Next
    
End Sub

''
' Writes the "RequestPositionUpdate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestPositionUpdate()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestPositionUpdate" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteRequestPositionUpdate_Err
    
    Call outgoingData.WriteByte(ClientPacketID.RequestPositionUpdate)

    
    Exit Sub

WriteRequestPositionUpdate_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestPositionUpdate", Erl)
    Resume Next
    
End Sub

''
' Writes the "Attack" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAttack()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Attack" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteAttack_Err
    
    Call outgoingData.WriteByte(ClientPacketID.Attack)

    
    Exit Sub

WriteAttack_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteAttack", Erl)
    Resume Next
    
End Sub

''
' Writes the "PickUp" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePickUp()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PickUp" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WritePickUp_Err
    
    Call outgoingData.WriteByte(ClientPacketID.PickUp)

    
    Exit Sub

WritePickUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WritePickUp", Erl)
    Resume Next
    
End Sub

''
' Writes the "SafeToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSafeToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SafeToggle" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteSafeToggle_Err
    
    Call outgoingData.WriteByte(ClientPacketID.SafeToggle)

    
    Exit Sub

WriteSafeToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSafeToggle", Erl)
    Resume Next
    
End Sub

Public Sub WriteSeguroClan()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SafeToggle" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteSeguroClan_Err
    
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.SeguroClan)

    
    Exit Sub

WriteSeguroClan_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSeguroClan", Erl)
    Resume Next
    
End Sub

Public Sub WriteTraerBoveda()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SafeToggle" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteTraerBoveda_Err
    
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.TraerBoveda)

    
    Exit Sub

WriteTraerBoveda_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteTraerBoveda", Erl)
    Resume Next
    
End Sub

''
' Writes the "CreatePretorianClan" message to the outgoing data buffer.
'
' @param    Map         The map in which create the pretorian clan.
' @param    X           The x pos where the king is settled.
' @param    Y           The y pos where the king is settled.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreatePretorianClan(ByVal map As Integer, ByVal x As Byte, ByVal y As Byte)
'***************************************************
'Author: ZaMa
'Last Modification: 29/10/2010
'Writes the "CreatePretorianClan" message to the outgoing data buffer
'***************************************************
    
    On Error GoTo WriteCreatePretorianClan_Err
    
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.CreatePretorianClan)
        Call .WriteInteger(map)
        Call .WriteByte(x)
        Call .WriteByte(y)
    End With
    
    Exit Sub

WriteCreatePretorianClan_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCreatePretorianClan", Erl)
    Resume Next
    
End Sub

''
' Writes the "DeletePretorianClan" message to the outgoing data buffer.
'
' @param    Map         The map which contains the pretorian clan to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDeletePretorianClan(ByVal map As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 29/10/2010
'Writes the "DeletePretorianClan" message to the outgoing data buffer
'***************************************************
    
    On Error GoTo WriteDeletePretorianClan_Err
    
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.RemovePretorianClan)
        Call .WriteInteger(map)
    End With
    
    Exit Sub

WriteDeletePretorianClan_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteDeletePretorianClan", Erl)
    Resume Next
    
End Sub

''
' Writes the "PartySafeToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteParyToggle()
    '**************************************************************
    'Author: Rapsodius
    'Creation Date: 10/10/07
    'Writes the Resuscitation safe toggle packet to the outgoing data buffer.
    '**************************************************************
    
    On Error GoTo WriteParyToggle_Err
    
    Call outgoingData.WriteByte(ClientPacketID.PartySafeToggle)

    
    Exit Sub

WriteParyToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteParyToggle", Erl)
    Resume Next
    
End Sub

''
' Writes the "SeguroResu" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSeguroResu()
    '**************************************************************
    'Author: Rapsodius
    'Creation Date: 10/10/07
    'Writes the Resuscitation safe toggle packet to the outgoing data buffer.
    '**************************************************************
    
    On Error GoTo WriteSeguroResu_Err
    
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.SeguroResu)

    
    Exit Sub

WriteSeguroResu_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSeguroResu", Erl)
    Resume Next
    
End Sub

''
' Writes the "RequestGuildLeaderInfo" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestGuildLeaderInfo()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestGuildLeaderInfo" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteRequestGuildLeaderInfo_Err
    
    Call outgoingData.WriteByte(ClientPacketID.RequestGuildLeaderInfo)

    
    Exit Sub

WriteRequestGuildLeaderInfo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestGuildLeaderInfo", Erl)
    Resume Next
    
End Sub

''
' Writes the "RequestAtributes" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestAtributes()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestAtributes" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteRequestAtributes_Err
    
    Call outgoingData.WriteByte(ClientPacketID.RequestAtributes)

    
    Exit Sub

WriteRequestAtributes_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestAtributes", Erl)
    Resume Next
    
End Sub

Public Sub WriteRequestFamiliar()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestFamiliar" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteRequestFamiliar_Err
    
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.RequestFamiliar)

    
    Exit Sub

WriteRequestFamiliar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestFamiliar", Erl)
    Resume Next
    
End Sub

Public Sub WriteRequestGrupo()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestFamiliar" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteRequestGrupo_Err
    
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.RequestGrupo)

    
    Exit Sub

WriteRequestGrupo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestGrupo", Erl)
    Resume Next
    
End Sub

''
' Writes the "RequestSkills" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestSkills()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestSkills" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteRequestSkills_Err
    
    Call outgoingData.WriteByte(ClientPacketID.RequestSkills)

    
    Exit Sub

WriteRequestSkills_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestSkills", Erl)
    Resume Next
    
End Sub

''
' Writes the "RequestMiniStats" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestMiniStats()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestMiniStats" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteRequestMiniStats_Err
    
    Call outgoingData.WriteByte(ClientPacketID.RequestMiniStats)

    
    Exit Sub

WriteRequestMiniStats_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestMiniStats", Erl)
    Resume Next
    
End Sub

''
' Writes the "CommerceEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceEnd()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceEnd" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteCommerceEnd_Err
    
    Call outgoingData.WriteByte(ClientPacketID.CommerceEnd)

    
    Exit Sub

WriteCommerceEnd_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCommerceEnd", Erl)
    Resume Next
    
End Sub

''
' Writes the "UserCommerceEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceEnd()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserCommerceEnd" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteUserCommerceEnd_Err
    
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceEnd)

    
    Exit Sub

WriteUserCommerceEnd_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteUserCommerceEnd", Erl)
    Resume Next
    
End Sub

''
' Writes the "BankEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankEnd()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankEnd" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteBankEnd_Err
    
    Call outgoingData.WriteByte(ClientPacketID.BankEnd)

    
    Exit Sub

WriteBankEnd_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBankEnd", Erl)
    Resume Next
    
End Sub

''
' Writes the "UserCommerceOk" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceOk()
    '***************************************************
    'Author: Fredy Horacio Treboux (liquid)
    'Last Modification: 01/10/07
    'Writes the "UserCommerceOk" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteUserCommerceOk_Err
    
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceOk)

    
    Exit Sub

WriteUserCommerceOk_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteUserCommerceOk", Erl)
    Resume Next
    
End Sub

''
' Writes the "UserCommerceReject" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceReject()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserCommerceReject" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteUserCommerceReject_Err
    
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceReject)

    
    Exit Sub

WriteUserCommerceReject_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteUserCommerceReject", Erl)
    Resume Next
    
End Sub

''
' Writes the "Drop" message to the outgoing data buffer.
'
' @param    slot Inventory slot where the item to drop is.
' @param    amount Number of items to drop.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDrop(ByVal Slot As Byte, ByVal Amount As Long)
    
    On Error GoTo WriteDrop_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Drop" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Drop)
        Call .WriteByte(Slot)
        Call .WriteLong(Amount)

    End With

    
    Exit Sub

WriteDrop_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteDrop", Erl)
    Resume Next
    
End Sub

''
' Writes the "CastSpell" message to the outgoing data buffer.
'
' @param    slot Spell List slot where the spell to cast is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCastSpell(ByVal Slot As Byte)
    
    On Error GoTo WriteCastSpell_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CastSpell" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CastSpell)
        
        Call .WriteByte(Slot)

    End With

    
    Exit Sub

WriteCastSpell_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCastSpell", Erl)
    Resume Next
    
End Sub

Public Sub WriteInvitarGrupo()
    
    On Error GoTo WriteInvitarGrupo_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CastSpell" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.InvitarGrupo)

    End With

    
    Exit Sub

WriteInvitarGrupo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteInvitarGrupo", Erl)
    Resume Next
    
End Sub

Public Sub WriteMarcaDeClan()
    
    On Error GoTo WriteMarcaDeClan_Err
    

    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 23/08/2020
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.MarcaDeClanpack)

    End With

    
    Exit Sub

WriteMarcaDeClan_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteMarcaDeClan", Erl)
    Resume Next
    
End Sub

Public Sub WriteMarcaDeGm()
    
    On Error GoTo WriteMarcaDeGm_Err
    

    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 23/08/2020
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.MarcaDeGMPack)

    End With

    
    Exit Sub

WriteMarcaDeGm_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteMarcaDeGm", Erl)
    Resume Next
    
End Sub

Public Sub WriteAbandonarGrupo()
    
    On Error GoTo WriteAbandonarGrupo_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CastSpell" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.AbandonarGrupo)

    End With

    
    Exit Sub

WriteAbandonarGrupo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteAbandonarGrupo", Erl)
    Resume Next
    
End Sub

Public Sub WriteHecharDeGrupo(ByVal indice As Byte)
    
    On Error GoTo WriteHecharDeGrupo_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CastSpell" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.HecharDeGrupo)
        Call .WriteByte(indice)

    End With

    
    Exit Sub

WriteHecharDeGrupo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteHecharDeGrupo", Erl)
    Resume Next
    
End Sub

''
' Writes the "LeftClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLeftClick(ByVal x As Byte, ByVal y As Byte)
    
    On Error GoTo WriteLeftClick_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "LeftClick" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.LeftClick)
        
        Call .WriteByte(x)
        Call .WriteByte(y)

    End With

    
    Exit Sub

WriteLeftClick_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteLeftClick", Erl)
    Resume Next
    
End Sub

''
' Writes the "DoubleClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDoubleClick(ByVal x As Byte, ByVal y As Byte)
    
    On Error GoTo WriteDoubleClick_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "DoubleClick" message to the outgoing data buffer
    '***************************************************
    'Call IntervaloPermiteClick(True)
    With outgoingData
        Call .WriteByte(ClientPacketID.DoubleClick)
        
        Call .WriteByte(x)
        Call .WriteByte(y)

    End With

    
    Exit Sub

WriteDoubleClick_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteDoubleClick", Erl)
    Resume Next
    
End Sub

''
' Writes the "Work" message to the outgoing data buffer.
'
' @param    skill The skill which the user attempts to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWork(ByVal Skill As eSkill)
    
    On Error GoTo WriteWork_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Work" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Work)
        
        Call .WriteByte(Skill)

    End With

    
    Exit Sub

WriteWork_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteWork", Erl)
    Resume Next
    
End Sub

Public Sub WriteThrowDice()
    
    On Error GoTo WriteThrowDice_Err
    
    Call outgoingData.WriteByte(ClientPacketID.ThrowDice)

    
    Exit Sub

WriteThrowDice_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteThrowDice", Erl)
    Resume Next
    
End Sub

''
' Writes the "UseSpellMacro" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUseSpellMacro()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UseSpellMacro" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteUseSpellMacro_Err
    
    Call outgoingData.WriteByte(ClientPacketID.UseSpellMacro)

    
    Exit Sub

WriteUseSpellMacro_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteUseSpellMacro", Erl)
    Resume Next
    
End Sub

''
' Writes the "UseItem" message to the outgoing data buffer.
'
' @param    slot Invetory slot where the item to use is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUseItem(ByVal Slot As Byte)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UseItem" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteUseItem_Err
    

    With outgoingData
        Call .WriteByte(ClientPacketID.UseItem)
        Call .WriteByte(Slot)

    End With

    
    Exit Sub

WriteUseItem_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteUseItem", Erl)
    Resume Next
    
End Sub

''
' Writes the "CraftBlacksmith" message to the outgoing data buffer.
'
' @param    item Index of the item to craft in the list sent by the server.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCraftBlacksmith(ByVal Item As Integer)
    
    On Error GoTo WriteCraftBlacksmith_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CraftBlacksmith" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CraftBlacksmith)
        
        Call .WriteInteger(Item)

    End With

    
    Exit Sub

WriteCraftBlacksmith_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCraftBlacksmith", Erl)
    Resume Next
    
End Sub

''
' Writes the "CraftCarpenter" message to the outgoing data buffer.
'
' @param    item Index of the item to craft in the list sent by the server.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCraftCarpenter(ByVal Item As Integer)
    
    On Error GoTo WriteCraftCarpenter_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CraftCarpenter" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CraftCarpenter)
        
        Call .WriteInteger(Item)

    End With

    
    Exit Sub

WriteCraftCarpenter_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCraftCarpenter", Erl)
    Resume Next
    
End Sub

Public Sub WriteCraftAlquimista(ByVal Item As Integer)
    
    On Error GoTo WriteCraftAlquimista_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CraftCarpenter" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.CraftAlquimista)
        Call .WriteInteger(Item)

    End With

    
    Exit Sub

WriteCraftAlquimista_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCraftAlquimista", Erl)
    Resume Next
    
End Sub

Public Sub WriteCraftSastre(ByVal Item As Integer)
    
    On Error GoTo WriteCraftSastre_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CraftCarpenter" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.CraftSastre)
        Call .WriteInteger(Item)

    End With

    
    Exit Sub

WriteCraftSastre_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCraftSastre", Erl)
    Resume Next
    
End Sub

''
' Writes the "WorkLeftClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @param    skill The skill which the user attempts to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorkLeftClick(ByVal x As Byte, ByVal y As Byte, ByVal Skill As eSkill)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "WorkLeftClick" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteWorkLeftClick_Err
    
    
    If pausa Then Exit Sub
    

    'Call IntervaloPermiteClick(True)
    With outgoingData
        Call .WriteByte(ClientPacketID.WorkLeftClick)
        
        Call .WriteByte(x)
        Call .WriteByte(y)
        Call .WriteByte(Skill)

    End With

    
    Exit Sub

WriteWorkLeftClick_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteWorkLeftClick", Erl)
    Resume Next
    
End Sub

''
' Writes the "CreateNewGuild" message to the outgoing data buffer.
'
' @param    desc    The guild's description
' @param    name    The guild's name
' @param    site    The guild's website
' @param    codex   Array of all rules of the guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNewGuild(ByVal desc As String, ByVal Name As String, ByVal Alineacion As Byte)
    
    On Error GoTo WriteCreateNewGuild_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CreateNewGuild" message to the outgoing data buffer
    '***************************************************
    Dim temp As String

    Dim i    As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.CreateNewGuild)
        
        Call .WriteASCIIString(desc)
        Call .WriteASCIIString(Name)
        
        Call .WriteByte(Alineacion)

    End With

    
    Exit Sub

WriteCreateNewGuild_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCreateNewGuild", Erl)
    Resume Next
    
End Sub

''
' Writes the "SpellInfo" message to the outgoing data buffer.
'
' @param    slot Spell List slot where the spell which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpellInfo(ByVal Slot As Byte)
    
    On Error GoTo WriteSpellInfo_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SpellInfo" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.SpellInfo)
        
        Call .WriteByte(Slot)

    End With

    
    Exit Sub

WriteSpellInfo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSpellInfo", Erl)
    Resume Next
    
End Sub

''
' Writes the "EquipItem" message to the outgoing data buffer.
'
' @param    slot Invetory slot where the item to equip is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEquipItem(ByVal Slot As Byte)
    
    On Error GoTo WriteEquipItem_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "EquipItem" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.EquipItem)
        
        Call .WriteByte(Slot)

    End With

    
    Exit Sub

WriteEquipItem_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteEquipItem", Erl)
    Resume Next
    
End Sub

''
' Writes the "ChangeHeading" message to the outgoing data buffer.
'
' @param    heading The direction in wich the user is moving.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeHeading(ByVal Heading As E_Heading)
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeHeading" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteChangeHeading_Err
    

    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeHeading)
        
        Call .WriteByte(Heading)

    End With

    
    Exit Sub

WriteChangeHeading_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteChangeHeading", Erl)
    Resume Next
    
End Sub

''
' Writes the "ModifySkills" message to the outgoing data buffer.
'
' @param    skillEdt a-based array containing for each skill the number of points to add to it.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteModifySkills(ByRef skillEdt() As Byte)
    
    On Error GoTo WriteModifySkills_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ModifySkills" message to the outgoing data buffer
    '***************************************************
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.ModifySkills)
        
        For i = 1 To NUMSKILLS
            Call .WriteByte(skillEdt(i))
        Next i

    End With

    
    Exit Sub

WriteModifySkills_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteModifySkills", Erl)
    Resume Next
    
End Sub

''
' Writes the "Train" message to the outgoing data buffer.
'
' @param    creature Position within the list provided by the server of the creature to train against.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrain(ByVal creature As Byte)
    
    On Error GoTo WriteTrain_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Train" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Train)
        
        Call .WriteByte(creature)

    End With

    
    Exit Sub

WriteTrain_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteTrain", Erl)
    Resume Next
    
End Sub

''
' Writes the "CommerceBuy" message to the outgoing data buffer.
'
' @param    slot Position within the NPC's inventory in which the desired item is.
' @param    amount Number of items to buy.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceBuy(ByVal Slot As Byte, ByVal Amount As Integer)
    
    On Error GoTo WriteCommerceBuy_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceBuy" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CommerceBuy)
        
        Call .WriteByte(Slot)
        Call .WriteInteger(Amount)

    End With

    
    Exit Sub

WriteCommerceBuy_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCommerceBuy", Erl)
    Resume Next
    
End Sub

Public Sub WriteUseKey(ByVal Slot As Byte)
    
    On Error GoTo WriteUseKey_Err
    

    With outgoingData
        Call .WriteByte(ClientPacketID.UseKey)
        Call .WriteByte(Slot)
    End With

    
    Exit Sub

WriteUseKey_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteUseKey", Erl)
    Resume Next
    
End Sub

''
' Writes the "BankExtractItem" message to the outgoing data buffer.
'
' @param    slot Position within the bank in which the desired item is.
' @param    amount Number of items to extract.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankExtractItem(ByVal Slot As Byte, ByVal Amount As Integer, ByVal slotdestino As Byte)
    
    On Error GoTo WriteBankExtractItem_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankExtractItem" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankExtractItem)
        
        Call .WriteByte(Slot)
        Call .WriteInteger(Amount)
        Call .WriteByte(slotdestino)
        
    End With

    
    Exit Sub

WriteBankExtractItem_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBankExtractItem", Erl)
    Resume Next
    
End Sub

''
' Writes the "CommerceSell" message to the outgoing data buffer.
'
' @param    slot Position within user inventory in which the desired item is.
' @param    amount Number of items to sell.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceSell(ByVal Slot As Byte, ByVal Amount As Integer)
    
    On Error GoTo WriteCommerceSell_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceSell" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CommerceSell)
        
        Call .WriteByte(Slot)
        Call .WriteInteger(Amount)

    End With

    
    Exit Sub

WriteCommerceSell_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCommerceSell", Erl)
    Resume Next
    
End Sub

''
' Writes the "BankDeposit" message to the outgoing data buffer.
'
' @param    slot Position within the user inventory in which the desired item is.
' @param    amount Number of items to deposit.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankDeposit(ByVal Slot As Byte, ByVal Amount As Integer, ByVal slotdestino As Byte)
    
    On Error GoTo WriteBankDeposit_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankDeposit" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankDeposit)
        Call .WriteByte(Slot)
        Call .WriteInteger(Amount)
        Call .WriteByte(slotdestino)

    End With

    
    Exit Sub

WriteBankDeposit_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBankDeposit", Erl)
    Resume Next
    
End Sub

''
' Writes the "ForumPost" message to the outgoing data buffer.
'
' @param    title The message's title.
' @param    message The body of the message.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForumPost(ByVal title As String, ByVal Message As String)
    
    On Error GoTo WriteForumPost_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ForumPost" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ForumPost)
        
        Call .WriteASCIIString(title)
        Call .WriteASCIIString(Message)

    End With

    
    Exit Sub

WriteForumPost_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteForumPost", Erl)
    Resume Next
    
End Sub

''
' Writes the "MoveSpell" message to the outgoing data buffer.
'
' @param    upwards True if the spell will be moved up in the list, False if it will be moved downwards.
' @param    slot Spell List slot where the spell which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMoveSpell(ByVal upwards As Boolean, ByVal Slot As Byte)
    
    On Error GoTo WriteMoveSpell_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "MoveSpell" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MoveSpell)
        
        Call .WriteBoolean(upwards)
        Call .WriteByte(Slot)

    End With

    
    Exit Sub

WriteMoveSpell_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteMoveSpell", Erl)
    Resume Next
    
End Sub

''
' Writes the "ClanCodexUpdate" message to the outgoing data buffer.
'
' @param    desc New description of the clan.
' @param    codex New codex of the clan.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteClanCodexUpdate(ByVal desc As String)
    
    On Error GoTo WriteClanCodexUpdate_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ClanCodexUpdate" message to the outgoing data buffer
    '***************************************************
    Dim temp As String

    Dim i    As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.ClanCodexUpdate)
        Call .WriteASCIIString(desc)

    End With

    
    Exit Sub

WriteClanCodexUpdate_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteClanCodexUpdate", Erl)
    Resume Next
    
End Sub

''
' Writes the "UserCommerceOffer" message to the outgoing data buffer.
'
' @param    slot Position within user inventory in which the desired item is.
' @param    amount Number of items to offer.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceOffer(ByVal Slot As Byte, ByVal Amount As Long)
    
    On Error GoTo WriteUserCommerceOffer_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UserCommerceOffer" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.UserCommerceOffer)
        
        Call .WriteByte(Slot)
        Call .WriteLong(Amount)

    End With

    
    Exit Sub

WriteUserCommerceOffer_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteUserCommerceOffer", Erl)
    Resume Next
    
End Sub

''
' Writes the "GuildAcceptPeace" message to the outgoing data buffer.
'
' @param    guild The guild whose peace offer is accepted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAcceptPeace(ByVal guild As String)
    
    On Error GoTo WriteGuildAcceptPeace_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildAcceptPeace" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAcceptPeace)
        
        Call .WriteASCIIString(guild)

    End With

    
    Exit Sub

WriteGuildAcceptPeace_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildAcceptPeace", Erl)
    Resume Next
    
End Sub

''
' Writes the "GuildRejectAlliance" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance offer is rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRejectAlliance(ByVal guild As String)
    
    On Error GoTo WriteGuildRejectAlliance_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRejectAlliance" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRejectAlliance)
        
        Call .WriteASCIIString(guild)

    End With

    
    Exit Sub

WriteGuildRejectAlliance_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildRejectAlliance", Erl)
    Resume Next
    
End Sub

''
' Writes the "GuildRejectPeace" message to the outgoing data buffer.
'
' @param    guild The guild whose peace offer is rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRejectPeace(ByVal guild As String)
    
    On Error GoTo WriteGuildRejectPeace_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRejectPeace" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRejectPeace)
        
        Call .WriteASCIIString(guild)

    End With

    
    Exit Sub

WriteGuildRejectPeace_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildRejectPeace", Erl)
    Resume Next
    
End Sub

''
' Writes the "GuildAcceptAlliance" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance offer is accepted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAcceptAlliance(ByVal guild As String)
    
    On Error GoTo WriteGuildAcceptAlliance_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildAcceptAlliance" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAcceptAlliance)
        
        Call .WriteASCIIString(guild)

    End With

    
    Exit Sub

WriteGuildAcceptAlliance_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildAcceptAlliance", Erl)
    Resume Next
    
End Sub

''
' Writes the "GuildOfferPeace" message to the outgoing data buffer.
'
' @param    guild The guild to whom peace is offered.
' @param    proposal The text to send with the proposal.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOfferPeace(ByVal guild As String, ByVal proposal As String)
    
    On Error GoTo WriteGuildOfferPeace_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildOfferPeace" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildOfferPeace)
        
        Call .WriteASCIIString(guild)
        Call .WriteASCIIString(proposal)

    End With

    
    Exit Sub

WriteGuildOfferPeace_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildOfferPeace", Erl)
    Resume Next
    
End Sub

''
' Writes the "GuildOfferAlliance" message to the outgoing data buffer.
'
' @param    guild The guild to whom an aliance is offered.
' @param    proposal The text to send with the proposal.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOfferAlliance(ByVal guild As String, ByVal proposal As String)
    
    On Error GoTo WriteGuildOfferAlliance_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildOfferAlliance" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildOfferAlliance)
        
        Call .WriteASCIIString(guild)
        Call .WriteASCIIString(proposal)

    End With

    
    Exit Sub

WriteGuildOfferAlliance_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildOfferAlliance", Erl)
    Resume Next
    
End Sub

''
' Writes the "GuildAllianceDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance proposal's details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAllianceDetails(ByVal guild As String)
    
    On Error GoTo WriteGuildAllianceDetails_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildAllianceDetails" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAllianceDetails)
        
        Call .WriteASCIIString(guild)

    End With

    
    Exit Sub

WriteGuildAllianceDetails_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildAllianceDetails", Erl)
    Resume Next
    
End Sub

''
' Writes the "GuildPeaceDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose peace proposal's details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildPeaceDetails(ByVal guild As String)
    
    On Error GoTo WriteGuildPeaceDetails_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildPeaceDetails" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildPeaceDetails)
        
        Call .WriteASCIIString(guild)

    End With

    
    Exit Sub

WriteGuildPeaceDetails_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildPeaceDetails", Erl)
    Resume Next
    
End Sub

''
' Writes the "GuildRequestJoinerInfo" message to the outgoing data buffer.
'
' @param    username The user who wants to join the guild whose info is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestJoinerInfo(ByVal UserName As String)
    
    On Error GoTo WriteGuildRequestJoinerInfo_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRequestJoinerInfo" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRequestJoinerInfo)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteGuildRequestJoinerInfo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildRequestJoinerInfo", Erl)
    Resume Next
    
End Sub

''
' Writes the "GuildAlliancePropList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAlliancePropList()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildAlliancePropList" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteGuildAlliancePropList_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GuildAlliancePropList)

    
    Exit Sub

WriteGuildAlliancePropList_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildAlliancePropList", Erl)
    Resume Next
    
End Sub

''
' Writes the "GuildPeacePropList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildPeacePropList()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildPeacePropList" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteGuildPeacePropList_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GuildPeacePropList)

    
    Exit Sub

WriteGuildPeacePropList_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildPeacePropList", Erl)
    Resume Next
    
End Sub

''
' Writes the "GuildDeclareWar" message to the outgoing data buffer.
'
' @param    guild The guild to which to declare war.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildDeclareWar(ByVal guild As String)
    
    On Error GoTo WriteGuildDeclareWar_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildDeclareWar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildDeclareWar)
        
        Call .WriteASCIIString(guild)

    End With

    
    Exit Sub

WriteGuildDeclareWar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildDeclareWar", Erl)
    Resume Next
    
End Sub

''
' Writes the "GuildNewWebsite" message to the outgoing data buffer.
'
' @param    url The guild's new website's URL.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildNewWebsite(ByVal url As String)
    
    On Error GoTo WriteGuildNewWebsite_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildNewWebsite" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildNewWebsite)
        
        Call .WriteASCIIString(url)

    End With

    
    Exit Sub

WriteGuildNewWebsite_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildNewWebsite", Erl)
    Resume Next
    
End Sub

''
' Writes the "GuildAcceptNewMember" message to the outgoing data buffer.
'
' @param    username The name of the accepted player.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAcceptNewMember(ByVal UserName As String)
    
    On Error GoTo WriteGuildAcceptNewMember_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildAcceptNewMember" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAcceptNewMember)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteGuildAcceptNewMember_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildAcceptNewMember", Erl)
    Resume Next
    
End Sub

''
' Writes the "GuildRejectNewMember" message to the outgoing data buffer.
'
' @param    username The name of the rejected player.
' @param    reason The reason for which the player was rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRejectNewMember(ByVal UserName As String, ByVal reason As String)
    
    On Error GoTo WriteGuildRejectNewMember_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRejectNewMember" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRejectNewMember)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(reason)

    End With

    
    Exit Sub

WriteGuildRejectNewMember_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildRejectNewMember", Erl)
    Resume Next
    
End Sub

''
' Writes the "GuildKickMember" message to the outgoing data buffer.
'
' @param    username The name of the kicked player.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildKickMember(ByVal UserName As String)
    
    On Error GoTo WriteGuildKickMember_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildKickMember" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildKickMember)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteGuildKickMember_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildKickMember", Erl)
    Resume Next
    
End Sub

''
' Writes the "GuildUpdateNews" message to the outgoing data buffer.
'
' @param    news The news to be posted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildUpdateNews(ByVal news As String)
    
    On Error GoTo WriteGuildUpdateNews_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildUpdateNews" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildUpdateNews)
        
        Call .WriteASCIIString(news)

    End With

    
    Exit Sub

WriteGuildUpdateNews_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildUpdateNews", Erl)
    Resume Next
    
End Sub

''
' Writes the "GuildMemberInfo" message to the outgoing data buffer.
'
' @param    username The user whose info is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMemberInfo(ByVal UserName As String)
    
    On Error GoTo WriteGuildMemberInfo_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildMemberInfo" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildMemberInfo)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteGuildMemberInfo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildMemberInfo", Erl)
    Resume Next
    
End Sub

''
' Writes the "GuildOpenElections" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOpenElections()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildOpenElections" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteGuildOpenElections_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GuildOpenElections)

    
    Exit Sub

WriteGuildOpenElections_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildOpenElections", Erl)
    Resume Next
    
End Sub

''
' Writes the "GuildRequestMembership" message to the outgoing data buffer.
'
' @param    guild The guild to which to request membership.
' @param    application The user's application sheet.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestMembership(ByVal guild As String, ByVal Application As String)
    
    On Error GoTo WriteGuildRequestMembership_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRequestMembership" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRequestMembership)
        
        Call .WriteASCIIString(guild)
        Call .WriteASCIIString(Application)

    End With

    
    Exit Sub

WriteGuildRequestMembership_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildRequestMembership", Erl)
    Resume Next
    
End Sub

''
' Writes the "GuildRequestDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestDetails(ByVal guild As String)
    
    On Error GoTo WriteGuildRequestDetails_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRequestDetails" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRequestDetails)
        
        Call .WriteASCIIString(guild)

    End With

    
    Exit Sub

WriteGuildRequestDetails_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildRequestDetails", Erl)
    Resume Next
    
End Sub

''
' Writes the "Online" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnline()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Online" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteOnline_Err
    
    Call outgoingData.WriteByte(ClientPacketID.Online)

    
    Exit Sub

WriteOnline_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteOnline", Erl)
    Resume Next
    
End Sub

''
' Writes the "Quit" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteQuit()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Quit" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteQuit_Err
    
    Call outgoingData.WriteByte(ClientPacketID.Quit)
    UserSaliendo = True

    Rem  MostrarCuenta = True
    
    Exit Sub

WriteQuit_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteQuit", Erl)
    Resume Next
    
End Sub

''
' Writes the "GuildLeave" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildLeave()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildLeave" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteGuildLeave_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GuildLeave)

    
    Exit Sub

WriteGuildLeave_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildLeave", Erl)
    Resume Next
    
End Sub

''
' Writes the "RequestAccountState" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestAccountState()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestAccountState" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteRequestAccountState_Err
    
    Call outgoingData.WriteByte(ClientPacketID.RequestAccountState)

    
    Exit Sub

WriteRequestAccountState_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestAccountState", Erl)
    Resume Next
    
End Sub

''
' Writes the "PetStand" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePetStand()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "PetStand" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WritePetStand_Err
    
    Call outgoingData.WriteByte(ClientPacketID.PetStand)

    
    Exit Sub

WritePetStand_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WritePetStand", Erl)
    Resume Next
    
End Sub

''
' Writes the "PetFollow" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePetFollow()
    '***************************************************
    'Writes the "PetFollow" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WritePetStand_Err
    
    Call outgoingData.WriteByte(ClientPacketID.PetFollow)

    
    Exit Sub

WritePetStand_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WritePetFollow", Erl)
    Resume Next
    
End Sub

''
' Writes the "PetLeave" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePetLeave()
    '***************************************************
    'Writes the "PetLeave" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WritePetStand_Err
    
    Call outgoingData.WriteByte(ClientPacketID.PetLeave)

    
    Exit Sub

WritePetStand_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WritePetLeave", Erl)
    Resume Next
    
End Sub

''
' Writes the "GrupoMsg" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGrupoMsg(ByVal Message As String)
    
    On Error GoTo WriteGrupoMsg_Err
    

    With outgoingData
        Call .WriteByte(ClientPacketID.GrupoMsg)
        
        Call .WriteASCIIString(Message)

    End With

    
    Exit Sub

WriteGrupoMsg_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGrupoMsg", Erl)
    Resume Next
    
End Sub

''
' Writes the "TrainList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrainList()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TrainList" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteTrainList_Err
    
    Call outgoingData.WriteByte(ClientPacketID.TrainList)

    
    Exit Sub

WriteTrainList_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteTrainList", Erl)
    Resume Next
    
End Sub

''
' Writes the "Rest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRest()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Rest" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteRest_Err
    
    Call outgoingData.WriteByte(ClientPacketID.Rest)

    
    Exit Sub

WriteRest_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRest", Erl)
    Resume Next
    
End Sub

''
' Writes the "Meditate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMeditate()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Meditate" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteMeditate_Err
    
    Call outgoingData.WriteByte(ClientPacketID.Meditate)

    
    Exit Sub

WriteMeditate_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteMeditate", Erl)
    Resume Next
    
End Sub

''
' Writes the "Resucitate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResucitate()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Resucitate" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteResucitate_Err
    
    Call outgoingData.WriteByte(ClientPacketID.Resucitate)

    
    Exit Sub

WriteResucitate_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteResucitate", Erl)
    Resume Next
    
End Sub

''
' Writes the "Heal" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHeal()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Heal" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteHeal_Err
    
    Call outgoingData.WriteByte(ClientPacketID.Heal)

    
    Exit Sub

WriteHeal_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteHeal", Erl)
    Resume Next
    
End Sub

''
' Writes the "Help" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHelp()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Help" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteHelp_Err
    
    Call outgoingData.WriteByte(ClientPacketID.Help)

    
    Exit Sub

WriteHelp_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteHelp", Erl)
    Resume Next
    
End Sub

''
' Writes the "RequestStats" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestStats()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestStats" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteRequestStats_Err
    
    Call outgoingData.WriteByte(ClientPacketID.RequestStats)

    
    Exit Sub

WriteRequestStats_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestStats", Erl)
    Resume Next
    
End Sub

''
' Writes the "Promedio" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePromedio()
    '***************************************************
    'Writes the "Promedio" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo Handle
    
    Call outgoingData.WriteByte(ClientPacketID.Promedio)

    Exit Sub

Handle:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WritePromedio", Erl)
    Resume Next
    
End Sub

''
' Writes the "GiveItem" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGiveItem(UserName As String, ByVal OBJIndex As Integer, ByVal cantidad As Integer, Motivo As String)
    '***************************************************
    'Writes the "GiveItem" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo Handle
    
    With outgoingData
        Call .WriteByte(ClientPacketID.GiveItem)
        Call .WriteASCIIString(UserName)
        Call .WriteInteger(OBJIndex)
        Call .WriteInteger(cantidad)
        Call .WriteASCIIString(Motivo)
    End With
    

    Exit Sub

Handle:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGiveItem", Erl)
    Resume Next
    
End Sub

''
' Writes the "CommerceStart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceStart()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CommerceStart" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteCommerceStart_Err
    
    Call outgoingData.WriteByte(ClientPacketID.CommerceStart)

    
    Exit Sub

WriteCommerceStart_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCommerceStart", Erl)
    Resume Next
    
End Sub

''
' Writes the "BankStart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankStart()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankStart" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteBankStart_Err
    
    Call outgoingData.WriteByte(ClientPacketID.BankStart)

    
    Exit Sub

WriteBankStart_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBankStart", Erl)
    Resume Next
    
End Sub

''
' Writes the "Enlist" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEnlist()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Enlist" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteEnlist_Err
    
    Call outgoingData.WriteByte(ClientPacketID.Enlist)

    
    Exit Sub

WriteEnlist_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteEnlist", Erl)
    Resume Next
    
End Sub

''
' Writes the "Information" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInformation()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Information" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteInformation_Err
    
    Call outgoingData.WriteByte(ClientPacketID.Information)

    
    Exit Sub

WriteInformation_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteInformation", Erl)
    Resume Next
    
End Sub

''
' Writes the "Reward" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReward()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Reward" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteReward_Err
    
    Call outgoingData.WriteByte(ClientPacketID.Reward)

    
    Exit Sub

WriteReward_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteReward", Erl)
    Resume Next
    
End Sub

''
' Writes the "RequestMOTD" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestMOTD()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestMOTD" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteRequestMOTD_Err
    
    Call outgoingData.WriteByte(ClientPacketID.RequestMOTD)

    
    Exit Sub

WriteRequestMOTD_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestMOTD", Erl)
    Resume Next
    
End Sub

''
' Writes the "UpTime" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUpTime()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UpTime" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteUpTime_Err
    
    Call outgoingData.WriteByte(ClientPacketID.Uptime)

    
    Exit Sub

WriteUpTime_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteUpTime", Erl)
    Resume Next
    
End Sub

''
' Writes the "Inquiry" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInquiry()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Inquiry" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteInquiry_Err
    
    Call outgoingData.WriteByte(ClientPacketID.Inquiry)

    
    Exit Sub

WriteInquiry_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteInquiry", Erl)
    Resume Next
    
End Sub

''
' Writes the "GuildMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMessage(ByVal Message As String)
    
    On Error GoTo WriteGuildMessage_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRequestDetails" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildMessage)
        
        Call .WriteASCIIString(Message)

    End With

    
    Exit Sub

WriteGuildMessage_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildMessage", Erl)
    Resume Next
    
End Sub

''
' Writes the "CentinelReport" message to the outgoing data buffer.
'
' @param    number The number to report to the centinel.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCentinelReport(ByVal Number As Integer)
    
    On Error GoTo WriteCentinelReport_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CentinelReport" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CentinelReport)
        
        Call .WriteInteger(Number)

    End With

    
    Exit Sub

WriteCentinelReport_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCentinelReport", Erl)
    Resume Next
    
End Sub

''
' Writes the "GuildOnline" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOnline()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildOnline" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteGuildOnline_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GuildOnline)

    
    Exit Sub

WriteGuildOnline_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildOnline", Erl)
    Resume Next
    
End Sub

''
' Writes the "CouncilMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the other council members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCouncilMessage(ByVal Message As String)
    
    On Error GoTo WriteCouncilMessage_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CouncilMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CouncilMessage)
        
        Call .WriteASCIIString(Message)

    End With

    
    Exit Sub

WriteCouncilMessage_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCouncilMessage", Erl)
    Resume Next
    
End Sub

''
' Writes the "RoleMasterRequest" message to the outgoing data buffer.
'
' @param    message The message to send to the role masters.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoleMasterRequest(ByVal Message As String)
    
    On Error GoTo WriteRoleMasterRequest_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RoleMasterRequest" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RoleMasterRequest)
        
        Call .WriteASCIIString(Message)

    End With

    
    Exit Sub

WriteRoleMasterRequest_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRoleMasterRequest", Erl)
    Resume Next
    
End Sub

''
' Writes the "GMRequest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMRequest()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GMRequest" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteGMRequest_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMRequest)

    
    Exit Sub

WriteGMRequest_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGMRequest", Erl)
    Resume Next
    
End Sub

''
' Writes the "ChangeDescription" message to the outgoing data buffer.
'
' @param    desc The new description of the user's character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeDescription(ByVal desc As String)
    
    On Error GoTo WriteChangeDescription_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeDescription" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeDescription)
        
        Call .WriteASCIIString(desc)

    End With

    
    Exit Sub

WriteChangeDescription_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteChangeDescription", Erl)
    Resume Next
    
End Sub

''
' Writes the "GuildVote" message to the outgoing data buffer.
'
' @param    username The user to vote for clan leader.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildVote(ByVal UserName As String)
    
    On Error GoTo WriteGuildVote_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildVote" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildVote)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteGuildVote_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildVote", Erl)
    Resume Next
    
End Sub

''
' Writes the "Punishments" message to the outgoing data buffer.
'
' @param    username The user whose's  punishments are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePunishments(ByVal UserName As String)
    
    On Error GoTo WritePunishments_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Punishments" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Punishments)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WritePunishments_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WritePunishments", Erl)
    Resume Next
    
End Sub

''
' Writes the "ChangePassword" message to the outgoing data buffer.
'
' @param    oldPass Previous password.
' @param    newPass New password.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangePassword(ByRef oldPass As String, ByRef newPass As String)
    
    On Error GoTo WriteChangePassword_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 10/10/07
    'Last Modified By: Rapsodius
    'Writes the "ChangePassword" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangePassword)

        Call .WriteASCIIString(SEncriptar(oldPass))
        Call .WriteASCIIString(SEncriptar(newPass))

    End With

    
    Exit Sub

WriteChangePassword_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteChangePassword", Erl)
    Resume Next
    
End Sub

''
' Writes the "Gamble" message to the outgoing data buffer.
'
' @param    amount The amount to gamble.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGamble(ByVal Amount As Integer)
    
    On Error GoTo WriteGamble_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Gamble" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Gamble)
        
        Call .WriteInteger(Amount)

    End With

    
    Exit Sub

WriteGamble_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGamble", Erl)
    Resume Next
    
End Sub

''
' Writes the "InquiryVote" message to the outgoing data buffer.
'
' @param    opt The chosen option to vote for.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInquiryVote(ByVal opt As Byte)
    
    On Error GoTo WriteInquiryVote_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "InquiryVote" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.InquiryVote)
        
        Call .WriteByte(opt)

    End With

    
    Exit Sub

WriteInquiryVote_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteInquiryVote", Erl)
    Resume Next
    
End Sub

''
' Writes the "LeaveFaction" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLeaveFaction()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "LeaveFaction" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteLeaveFaction_Err
    
    Call outgoingData.WriteByte(ClientPacketID.LeaveFaction)

    
    Exit Sub

WriteLeaveFaction_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteLeaveFaction", Erl)
    Resume Next
    
End Sub

''
' Writes the "BankExtractGold" message to the outgoing data buffer.
'
' @param    amount The amount of money to extract from the bank.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankExtractGold(ByVal Amount As Long)
    
    On Error GoTo WriteBankExtractGold_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankExtractGold" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankExtractGold)
        
        Call .WriteLong(Amount)

    End With

    
    Exit Sub

WriteBankExtractGold_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBankExtractGold", Erl)
    Resume Next
    
End Sub

''
' Writes the "BankDepositGold" message to the outgoing data buffer.
'
' @param    amount The amount of money to deposit in the bank.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankDepositGold(ByVal Amount As Long)
    
    On Error GoTo WriteBankDepositGold_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankDepositGold" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankDepositGold)
        
        Call .WriteLong(Amount)

    End With

    
    Exit Sub

WriteBankDepositGold_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBankDepositGold", Erl)
    Resume Next
    
End Sub

Public Sub WriteTransFerGold(ByVal Amount As Long, ByVal destino As String)
    
    On Error GoTo WriteTransFerGold_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankDepositGold" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.TransFerGold)
        Call .WriteLong(Amount)
        Call .WriteASCIIString(destino)

    End With

    
    Exit Sub

WriteTransFerGold_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteTransFerGold", Erl)
    Resume Next
    
End Sub

Public Sub WriteItemMove(ByVal SlotActual As Byte, ByVal SlotNuevo As Byte)
    
    On Error GoTo WriteItemMove_Err
    

    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.MoveItem)
        Call .WriteByte(SlotActual)
        Call .WriteByte(SlotNuevo)

    End With

    
    Exit Sub

WriteItemMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteItemMove", Erl)
    Resume Next
    
End Sub

Public Sub WriteBovedaItemMove(ByVal SlotActual As Byte, ByVal SlotNuevo As Byte)
    
    On Error GoTo WriteBovedaItemMove_Err
    

    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.BovedaMoveItem)
        Call .WriteByte(SlotActual)
        Call .WriteByte(SlotNuevo)

    End With

    
    Exit Sub

WriteBovedaItemMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBovedaItemMove", Erl)
    Resume Next
    
End Sub

''
' Writes the "FinEvento" message to the outgoing data buffer.
'
' @param    message The message to send with the denounce.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteFinEvento()
    
    On Error GoTo WriteFinEvento_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "FinEvento" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.FinEvento)
    End With

    
    Exit Sub

WriteFinEvento_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteFinEvento", Erl)
    Resume Next
    
End Sub

''
' Writes the "Denounce" message to the outgoing data buffer.
'
' @param    message The message to send with the denounce.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDenounce(Name As String)
    
    On Error GoTo WriteDenounce_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Denounce" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Denounce)
        Call .WriteASCIIString(Name)
    End With

    
    Exit Sub

WriteDenounce_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteDenounce", Erl)
    Resume Next
    
End Sub

Public Sub WriteQuieroFundarClan()
    
    On Error GoTo WriteQuieroFundarClan_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Denounce" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.QuieroFundarClan)

    End With

    
    Exit Sub

WriteQuieroFundarClan_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteQuieroFundarClan", Erl)
    Resume Next
    
End Sub

''
' Writes the "GuildMemberList" message to the outgoing data buffer.
'
' @param    guild The guild whose member list is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMemberList(ByVal guild As String)
    
    On Error GoTo WriteGuildMemberList_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildMemberList" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildMemberList)
        
        Call .WriteASCIIString(guild)

    End With

    
    Exit Sub

WriteGuildMemberList_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildMemberList", Erl)
    Resume Next
    
End Sub

'ladder
Public Sub WriteCasamiento(ByVal UserName As String)
    
    On Error GoTo WriteCasamiento_Err
    

    '***************************************************
    'Ladder
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.Casarse)
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteCasamiento_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCasamiento", Erl)
    Resume Next
    
End Sub

Public Sub WriteMacroPos()
    
    On Error GoTo WriteMacroPos_Err
    

    '***************************************************
    'Ladder
    'Macros
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.MacroPosSent)
        Call .WriteByte(ChatCombate)
        Call .WriteByte(ChatGlobal)

    End With

    
    Exit Sub

WriteMacroPos_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteMacroPos", Erl)
    Resume Next
    
End Sub

Public Sub WriteSubastaInfo()
    
    On Error GoTo WriteSubastaInfo_Err
    

    '***************************************************
    'Ladder
    'Macros
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.SubastaInfo)

    End With

    
    Exit Sub

WriteSubastaInfo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSubastaInfo", Erl)
    Resume Next
    
End Sub

Public Sub WriteScrollInfo()
    
    On Error GoTo WriteScrollInfo_Err
    

    '***************************************************
    'Ladder
    'Macros
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.SCROLLINFO)

    End With

    
    Exit Sub

WriteScrollInfo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteScrollInfo", Erl)
    Resume Next
    
End Sub

Public Sub WriteCancelarExit()
    '***************************************************
    'Ladder
    'Cancelar Salida
    '***************************************************
    
    On Error GoTo WriteCancelarExit_Err
    
    UserSaliendo = False

    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.CancelarExit)

    End With

    
    Exit Sub

WriteCancelarExit_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCancelarExit", Erl)
    Resume Next
    
End Sub

Public Sub WriteEventoInfo()
    
    On Error GoTo WriteEventoInfo_Err
    

    '***************************************************
    'Ladder
    'Macros
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.EventoInfo)

    End With

    
    Exit Sub

WriteEventoInfo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteEventoInfo", Erl)
    Resume Next
    
End Sub

Public Sub WriteFlagTrabajar()
    
    On Error GoTo WriteFlagTrabajar_Err
    

    '***************************************************
    'Ladder
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.FlagTrabajar)

    End With

    
    Exit Sub

WriteFlagTrabajar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteFlagTrabajar", Erl)
    Resume Next
    
End Sub

''
' Writes the "GMMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to the other GMs online.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteEscribiendo()
    
    On Error GoTo WriteEscribiendo_Err
    
    If MostrarEscribiendo = 0 Then Exit Sub
    

    '***************************************************
    'Ladder
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.Escribiendo)

    End With

    
    Exit Sub

WriteEscribiendo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteEscribiendo", Erl)
    Resume Next
    
End Sub

Public Sub WriteReclamarRecompensa(ByVal Index As Byte)
    
    On Error GoTo WriteReclamarRecompensa_Err
    

    '***************************************************
    'Ladder
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.ReclamarRecompensa)
        Call .WriteByte(Index)

    End With

    
    Exit Sub

WriteReclamarRecompensa_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteReclamarRecompensa", Erl)
    Resume Next
    
End Sub

Public Sub WriteGMMessage(ByVal Message As String)
    
    On Error GoTo WriteGMMessage_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GMMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMMessage)
        
        Call .WriteASCIIString(Message)

    End With

    
    Exit Sub

WriteGMMessage_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGMMessage", Erl)
    Resume Next
    
End Sub

''
' Writes the "ShowName" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowName()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowName" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteShowName_Err
    
    Call outgoingData.WriteByte(ClientPacketID.showName)

    
    Exit Sub

WriteShowName_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteShowName", Erl)
    Resume Next
    
End Sub

''
' Writes the "OnlineRoyalArmy" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineRoyalArmy()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "OnlineRoyalArmy" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteOnlineRoyalArmy_Err
    
    Call outgoingData.WriteByte(ClientPacketID.OnlineRoyalArmy)

    
    Exit Sub

WriteOnlineRoyalArmy_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteOnlineRoyalArmy", Erl)
    Resume Next
    
End Sub

''
' Writes the "OnlineChaosLegion" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineChaosLegion()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "OnlineChaosLegion" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteOnlineChaosLegion_Err
    
    Call outgoingData.WriteByte(ClientPacketID.OnlineChaosLegion)

    
    Exit Sub

WriteOnlineChaosLegion_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteOnlineChaosLegion", Erl)
    Resume Next
    
End Sub

''
' Writes the "GoNearby" message to the outgoing data buffer.
'
' @param    username The suer to approach.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGoNearby(ByVal UserName As String)
    
    On Error GoTo WriteGoNearby_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GoNearby" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GoNearby)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteGoNearby_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGoNearby", Erl)
    Resume Next
    
End Sub

''
' Writes the "Comment" message to the outgoing data buffer.
'
' @param    message The message to leave in the log as a comment.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteComment(ByVal Message As String)
    
    On Error GoTo WriteComment_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Comment" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.comment)
        
        Call .WriteASCIIString(Message)

    End With

    
    Exit Sub

WriteComment_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteComment", Erl)
    Resume Next
    
End Sub

''
' Writes the "ServerTime" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerTime()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ServerTime" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteServerTime_Err
    
    Call outgoingData.WriteByte(ClientPacketID.serverTime)

    
    Exit Sub

WriteServerTime_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteServerTime", Erl)
    Resume Next
    
End Sub

''
' Writes the "Where" message to the outgoing data buffer.
'
' @param    username The user whose position is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWhere(ByVal UserName As String)
    
    On Error GoTo WriteWhere_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Where" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Where)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteWhere_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteWhere", Erl)
    Resume Next
    
End Sub

''
' Writes the "CreaturesInMap" message to the outgoing data buffer.
'
' @param    map The map in which to check for the existing creatures.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreaturesInMap(ByVal map As Integer)
    
    On Error GoTo WriteCreaturesInMap_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CreaturesInMap" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CreaturesInMap)
        
        Call .WriteInteger(map)

    End With

    
    Exit Sub

WriteCreaturesInMap_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCreaturesInMap", Erl)
    Resume Next
    
End Sub

''
' Writes the "WarpMeToTarget" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarpMeToTarget()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "WarpMeToTarget" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteWarpMeToTarget_Err
    
    Call outgoingData.WriteByte(ClientPacketID.WarpMeToTarget)

    
    Exit Sub

WriteWarpMeToTarget_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteWarpMeToTarget", Erl)
    Resume Next
    
End Sub

''
' Writes the "WarpChar" message to the outgoing data buffer.
'
' @param    username The user to be warped. "YO" represent's the user's char.
' @param    map The map to which to warp the character.
' @param    x The x position in the map to which to waro the character.
' @param    y The y position in the map to which to waro the character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarpChar(ByVal UserName As String, ByVal map As Integer, ByVal x As Byte, ByVal y As Byte)
    
    On Error GoTo WriteWarpChar_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "WarpChar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.WarpChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteInteger(map)
        
        Call .WriteByte(x)
        Call .WriteByte(y)

    End With

    
    Exit Sub

WriteWarpChar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteWarpChar", Erl)
    Resume Next
    
End Sub

''
' Writes the "Silence" message to the outgoing data buffer.
'
' @param    username The user to silence.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSilence(ByVal UserName As String)
    
    On Error GoTo WriteSilence_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Silence" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Silence)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteSilence_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSilence", Erl)
    Resume Next
    
End Sub

Public Sub WriteCuentaRegresiva(ByVal Second As Byte)
    
    On Error GoTo WriteCuentaRegresiva_Err
    

    '***************************************************
    'Writer by Ladder
    '/Cuentaregresiva <Segundos>
    '04-12-08
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.CuentaRegresiva)
        Call .WriteByte(Second)

    End With

    
    Exit Sub

WriteCuentaRegresiva_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCuentaRegresiva", Erl)
    Resume Next
    
End Sub

Public Sub WritePossUser(ByVal UserName As String)
    '***************************************************
    'Write by Ladder
    '03-12-08
    'Guarda la posición donde estamos parados, como la posición del personaje.
    'Esta pensado exclusivamente para deslogear PJs.
    '***************************************************
    
    On Error GoTo WritePossUser_Err
    

    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.PossUser)
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WritePossUser_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WritePossUser", Erl)
    Resume Next
    
End Sub

''
' Writes the "SOSShowList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSOSShowList()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SOSShowList" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteSOSShowList_Err
    
    Call outgoingData.WriteByte(ClientPacketID.SOSShowList)

    
    Exit Sub

WriteSOSShowList_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSOSShowList", Erl)
    Resume Next
    
End Sub

''
' Writes the "SOSRemove" message to the outgoing data buffer.
'
' @param    username The user whose SOS call has been already attended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSOSRemove(ByVal UserName As String)
    
    On Error GoTo WriteSOSRemove_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SOSRemove" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.SOSRemove)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteSOSRemove_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSOSRemove", Erl)
    Resume Next
    
End Sub

''
' Writes the "GoToChar" message to the outgoing data buffer.
'
' @param    username The user to be approached.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGoToChar(ByVal UserName As String)
    
    On Error GoTo WriteGoToChar_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GoToChar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GoToChar)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteGoToChar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGoToChar", Erl)
    Resume Next
    
End Sub

Public Sub WriteDesbuggear(ByVal Params As String)
    
    On Error GoTo WriteDesbuggear_Err
    

    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Desbuggear)
        Call .WriteASCIIString(Params)

    End With

    
    Exit Sub

WriteDesbuggear_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteDesbuggear", Erl)
    Resume Next
    
End Sub

Public Sub WriteDarLlaveAUsuario(ByVal User As String, ByVal Llave As Integer)
    
    On Error GoTo WriteDarLlaveAUsuario_Err
    

    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.DarLlaveAUsuario)
        Call .WriteASCIIString(User)
        Call .WriteInteger(Llave)
    End With

    
    Exit Sub

WriteDarLlaveAUsuario_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteDarLlaveAUsuario", Erl)
    Resume Next
    
End Sub

Public Sub WriteSacarLlave(ByVal Llave As Integer)
    
    On Error GoTo WriteSacarLlave_Err
    

    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.SacarLlave)
        Call .WriteInteger(Llave)
    End With

    
    Exit Sub

WriteSacarLlave_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSacarLlave", Erl)
    Resume Next
    
End Sub

Public Sub WriteVerLlaves()
    
    On Error GoTo WriteVerLlaves_Err
    

    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.VerLlaves)
    End With

    
    Exit Sub

WriteVerLlaves_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteVerLlaves", Erl)
    Resume Next
    
End Sub

''
' Writes the "Invisible" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInvisible()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Invisible" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteInvisible_Err
    
    Call outgoingData.WriteByte(ClientPacketID.Invisible)

    
    Exit Sub

WriteInvisible_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteInvisible", Erl)
    Resume Next
    
End Sub

''
' Writes the "GMPanel" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGMPanel()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GMPanel" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteGMPanel_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GMPanel)

    
    Exit Sub

WriteGMPanel_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGMPanel", Erl)
    Resume Next
    
End Sub

''
' Writes the "RequestUserList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestUserList()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestUserList" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteRequestUserList_Err
    
    Call outgoingData.WriteByte(ClientPacketID.RequestUserList)

    
    Exit Sub

WriteRequestUserList_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestUserList", Erl)
    Resume Next
    
End Sub

''
' Writes the "Working" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWorking()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Working" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteWorking_Err
    
    Call outgoingData.WriteByte(ClientPacketID.Working)

    
    Exit Sub

WriteWorking_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteWorking", Erl)
    Resume Next
    
End Sub

''
' Writes the "Hiding" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteHiding()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Hiding" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteHiding_Err
    
    Call outgoingData.WriteByte(ClientPacketID.Hiding)

    
    Exit Sub

WriteHiding_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteHiding", Erl)
    Resume Next
    
End Sub

''
' Writes the "Jail" message to the outgoing data buffer.
'
' @param    username The user to be sent to jail.
' @param    reason The reason for which to send him to jail.
' @param    time The time (in minutes) the user will have to spend there.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteJail(ByVal UserName As String, ByVal reason As String, ByVal Time As Byte)
    
    On Error GoTo WriteJail_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Jail" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Jail)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(reason)
        
        Call .WriteByte(Time)

    End With

    
    Exit Sub

WriteJail_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteJail", Erl)
    Resume Next
    
End Sub

Public Sub WriteCrearEvento(ByVal TIPO As Byte, ByVal duracion As Byte, ByVal multiplicacion As Byte)
    
    On Error GoTo WriteCrearEvento_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Jail" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.CrearEvento)
        
        Call .WriteByte(TIPO)
        Call .WriteByte(duracion)
        
        Call .WriteByte(multiplicacion)

    End With

    
    Exit Sub

WriteCrearEvento_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCrearEvento", Erl)
    Resume Next
    
End Sub

''
' Writes the "KillNPC" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillNPC()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "KillNPC" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteKillNPC_Err
    
    Call outgoingData.WriteByte(ClientPacketID.KillNPC)

    
    Exit Sub

WriteKillNPC_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteKillNPC", Erl)
    Resume Next
    
End Sub

''
' Writes the "WarnUser" message to the outgoing data buffer.
'
' @param    username The user to be warned.
' @param    reason Reason for the warning.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarnUser(ByVal UserName As String, ByVal reason As String)
    
    On Error GoTo WriteWarnUser_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "WarnUser" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.WarnUser)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(reason)

    End With

    
    Exit Sub

WriteWarnUser_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteWarnUser", Erl)
    Resume Next
    
End Sub

Public Sub WriteMensajeUser(ByVal UserName As String, ByVal mensaje As String)
    
    On Error GoTo WriteMensajeUser_Err
    

    '***************************************************
    'Author: Ladder
    'Last Modification: 04/jun/2014
    'Escribe un mensaje al usuario
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.MensajeUser)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(mensaje)

    End With

    
    Exit Sub

WriteMensajeUser_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteMensajeUser", Erl)
    Resume Next
    
End Sub

''
' Writes the "EditChar" message to the outgoing data buffer.
'
' @param    UserName    The user to be edited.
' @param    editOption  Indicates what to edit in the char.
' @param    arg1        Additional argument 1. Contents depend on editoption.
' @param    arg2        Additional argument 2. Contents depend on editoption.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEditChar(ByVal UserName As String, ByVal editOption As eEditOptions, ByVal arg1 As String, ByVal arg2 As String)
    
    On Error GoTo WriteEditChar_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "EditChar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.EditChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteByte(editOption)
        
        Call .WriteASCIIString(arg1)
        Call .WriteASCIIString(arg2)

    End With

    
    Exit Sub

WriteEditChar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteEditChar", Erl)
    Resume Next
    
End Sub

''
' Writes the "RequestCharInfo" message to the outgoing data buffer.
'
' @param    username The user whose information is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharInfo(ByVal UserName As String)
    
    On Error GoTo WriteRequestCharInfo_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharInfo" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RequestCharInfo)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteRequestCharInfo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestCharInfo", Erl)
    Resume Next
    
End Sub

''
' Writes the "RequestCharStats" message to the outgoing data buffer.
'
' @param    username The user whose stats are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharStats(ByVal UserName As String)
    
    On Error GoTo WriteRequestCharStats_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharStats" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RequestCharStats)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteRequestCharStats_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestCharStats", Erl)
    Resume Next
    
End Sub

''
' Writes the "RequestCharGold" message to the outgoing data buffer.
'
' @param    username The user whose gold is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharGold(ByVal UserName As String)
    
    On Error GoTo WriteRequestCharGold_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharGold" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RequestCharGold)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteRequestCharGold_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestCharGold", Erl)
    Resume Next
    
End Sub
    
''
' Writes the "RequestCharInventory" message to the outgoing data buffer.
'
' @param    username The user whose inventory is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharInventory(ByVal UserName As String)
    
    On Error GoTo WriteRequestCharInventory_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharInventory" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RequestCharInventory)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteRequestCharInventory_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestCharInventory", Erl)
    Resume Next
    
End Sub

''
' Writes the "RequestCharBank" message to the outgoing data buffer.
'
' @param    username The user whose banking information is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharBank(ByVal UserName As String)
    
    On Error GoTo WriteRequestCharBank_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharBank" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RequestCharBank)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteRequestCharBank_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestCharBank", Erl)
    Resume Next
    
End Sub

''
' Writes the "RequestCharSkills" message to the outgoing data buffer.
'
' @param    username The user whose skills are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharSkills(ByVal UserName As String)
    
    On Error GoTo WriteRequestCharSkills_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharSkills" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RequestCharSkills)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteRequestCharSkills_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestCharSkills", Erl)
    Resume Next
    
End Sub

''
' Writes the "ReviveChar" message to the outgoing data buffer.
'
' @param    username The user to eb revived.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReviveChar(ByVal UserName As String)
    
    On Error GoTo WriteReviveChar_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ReviveChar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ReviveChar)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteReviveChar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteReviveChar", Erl)
    Resume Next
    
End Sub

''
' Writes the "OnlineGM" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineGM()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "OnlineGM" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteOnlineGM_Err
    
    Call outgoingData.WriteByte(ClientPacketID.OnlineGM)

    
    Exit Sub

WriteOnlineGM_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteOnlineGM", Erl)
    Resume Next
    
End Sub

''
' Writes the "OnlineMap" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteOnlineMap()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "OnlineMap" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteOnlineMap_Err
    
    Call outgoingData.WriteByte(ClientPacketID.OnlineMap)

    
    Exit Sub

WriteOnlineMap_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteOnlineMap", Erl)
    Resume Next
    
End Sub

''
' Writes the "Forgive" message to the outgoing data buffer.
'
' @param    username The user to be forgiven.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForgive()
    
    On Error GoTo WriteForgive_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Forgive" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Forgive)
        
        '  Call .WriteASCIIString(UserName)
    End With

    
    Exit Sub

WriteForgive_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteForgive", Erl)
    Resume Next
    
End Sub

Public Sub WriteDonateGold(ByVal oro As Long)
    
    On Error GoTo WriteForgive_Err
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.DonateGold)
        Call .WriteLong(oro)
    End With

    
    Exit Sub

WriteForgive_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteDonateGold", Erl)
    Resume Next
    
End Sub

''
' Writes the "Kick" message to the outgoing data buffer.
'
' @param    username The user to be kicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKick(ByVal UserName As String)
    
    On Error GoTo WriteKick_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Kick" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Kick)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteKick_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteKick", Erl)
    Resume Next
    
End Sub

''
' Writes the "Execute" message to the outgoing data buffer.
'
' @param    username The user to be executed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteExecute(ByVal UserName As String)
    
    On Error GoTo WriteExecute_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Execute" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Execute)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteExecute_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteExecute", Erl)
    Resume Next
    
End Sub

''
' Writes the "BanChar" message to the outgoing data buffer.
'
' @param    username The user to be banned.
' @param    reason The reson for which the user is to be banned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBanChar(ByVal UserName As String, ByVal reason As String)
    
    On Error GoTo WriteBanChar_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BanChar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BanChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteASCIIString(reason)

    End With

    
    Exit Sub

WriteBanChar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBanChar", Erl)
    Resume Next
    
End Sub

Public Sub WriteBanCuenta(ByVal UserName As String, ByVal reason As String)
    
    On Error GoTo WriteBanCuenta_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BanCuenta" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.BanCuenta)
    
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(reason)

    End With

    
    Exit Sub

WriteBanCuenta_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBanCuenta", Erl)
    Resume Next
    
End Sub

Public Sub WriteUnBanCuenta(ByVal UserName As String)
    
    On Error GoTo WriteUnBanCuenta_Err
    

    '***************************************************
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.UnbanCuenta)
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteUnBanCuenta_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteUnBanCuenta", Erl)
    Resume Next
    
End Sub

Public Sub WriteBanSerial(ByVal UserName As String)
    
    On Error GoTo WriteBanSerial_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BanCuenta" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.BanSerial)
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteBanSerial_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBanSerial", Erl)
    Resume Next
    
End Sub

Public Sub WriteUnBanSerial(ByVal UserName As String, ByVal reason As String)
    
    On Error GoTo WriteUnBanSerial_Err
    

    '***************************************************
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.UnBanSerial)
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteUnBanSerial_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteUnBanSerial", Erl)
    Resume Next
    
End Sub

Public Sub WriteCerraCliente(ByVal UserName As String)
    
    On Error GoTo WriteCerraCliente_Err
    

    '***************************************************
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.CerrarCliente)
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteCerraCliente_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCerraCliente", Erl)
    Resume Next
    
End Sub

Public Sub WriteBanTemporal(ByVal UserName As String, ByVal reason As String, ByVal dias As Byte)
    
    On Error GoTo WriteBanTemporal_Err
    

    '***************************************************
    'Writes the "BanTemporal" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.BanTemporal)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(reason)
        Call .WriteByte(dias)

    End With

    
    Exit Sub

WriteBanTemporal_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBanTemporal", Erl)
    Resume Next
    
End Sub

Public Sub WriteSilenciarUser(ByVal UserName As String, ByVal Time As Byte)
    
    On Error GoTo WriteSilenciarUser_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BanChar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.SilenciarUser)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteByte(Time)

    End With

    
    Exit Sub

WriteSilenciarUser_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSilenciarUser", Erl)
    Resume Next
    
End Sub

''
' Writes the "UnbanChar" message to the outgoing data buffer.
'
' @param    username The user to be unbanned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUnbanChar(ByVal UserName As String)
    
    On Error GoTo WriteUnbanChar_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UnbanChar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.UnbanChar)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteUnbanChar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteUnbanChar", Erl)
    Resume Next
    
End Sub

''
' Writes the "NPCFollow" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNPCFollow()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "NPCFollow" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteNPCFollow_Err
    
    Call outgoingData.WriteByte(ClientPacketID.NPCFollow)

    
    Exit Sub

WriteNPCFollow_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteNPCFollow", Erl)
    Resume Next
    
End Sub

''
' Writes the "SummonChar" message to the outgoing data buffer.
'
' @param    username The user to be summoned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSummonChar(ByVal UserName As String)
    
    On Error GoTo WriteSummonChar_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SummonChar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.SummonChar)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteSummonChar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSummonChar", Erl)
    Resume Next
    
End Sub

''
' Writes the "SpawnListRequest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnListRequest()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SpawnListRequest" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteSpawnListRequest_Err
    
    Call outgoingData.WriteByte(ClientPacketID.SpawnListRequest)

    
    Exit Sub

WriteSpawnListRequest_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSpawnListRequest", Erl)
    Resume Next
    
End Sub

''
' Writes the "SpawnCreature" message to the outgoing data buffer.
'
' @param    creatureIndex The index of the creature in the spawn list to be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnCreature(ByVal creatureIndex As Integer)
    
    On Error GoTo WriteSpawnCreature_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SpawnCreature" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.SpawnCreature)
        
        Call .WriteInteger(creatureIndex)

    End With

    
    Exit Sub

WriteSpawnCreature_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSpawnCreature", Erl)
    Resume Next
    
End Sub

''
' Writes the "ResetNPCInventory" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetNPCInventory()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ResetNPCInventory" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteResetNPCInventory_Err
    
    Call outgoingData.WriteByte(ClientPacketID.ResetNPCInventory)

    
    Exit Sub

WriteResetNPCInventory_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteResetNPCInventory", Erl)
    Resume Next
    
End Sub

''
' Writes the "CleanWorld" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCleanWorld()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CleanWorld" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteCleanWorld_Err
    
    Call outgoingData.WriteByte(ClientPacketID.CleanWorld)

    
    Exit Sub

WriteCleanWorld_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCleanWorld", Erl)
    Resume Next
    
End Sub

''
' Writes the "ServerMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerMessage(ByVal Message As String)
    
    On Error GoTo WriteServerMessage_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ServerMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ServerMessage)
        
        Call .WriteASCIIString(Message)

    End With

    
    Exit Sub

WriteServerMessage_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteServerMessage", Erl)
    Resume Next
    
End Sub

''
' Writes the "NickToIP" message to the outgoing data buffer.
'
' @param    username The user whose IP is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNickToIP(ByVal UserName As String)
    
    On Error GoTo WriteNickToIP_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "NickToIP" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.NickToIP)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteNickToIP_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteNickToIP", Erl)
    Resume Next
    
End Sub

''
' Writes the "IPToNick" message to the outgoing data buffer.
'
' @param    IP The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteIPToNick(ByRef IP() As Byte)
    
    On Error GoTo WriteIPToNick_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "IPToNick" message to the outgoing data buffer
    '***************************************************
    If UBound(IP()) - LBound(IP()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.IPToNick)
        
        For i = LBound(IP()) To UBound(IP())
            Call .WriteByte(IP(i))
        Next i

    End With

    
    Exit Sub

WriteIPToNick_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteIPToNick", Erl)
    Resume Next
    
End Sub

''
' Writes the "GuildOnlineMembers" message to the outgoing data buffer.
'
' @param    guild The guild whose online player list is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOnlineMembers(ByVal guild As String)
    
    On Error GoTo WriteGuildOnlineMembers_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildOnlineMembers" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildOnlineMembers)
        
        Call .WriteASCIIString(guild)

    End With

    
    Exit Sub

WriteGuildOnlineMembers_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildOnlineMembers", Erl)
    Resume Next
    
End Sub

''
' Writes the "TeleportCreate" message to the outgoing data buffer.
'
' @param    map the map to which the teleport will lead.
' @param    x The position in the x axis to which the teleport will lead.
' @param    y The position in the y axis to which the teleport will lead.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTeleportCreate(ByVal map As Integer, ByVal x As Byte, ByVal y As Byte)
    
    On Error GoTo WriteTeleportCreate_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TeleportCreate" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.TeleportCreate)
        
        Call .WriteInteger(map)
        
        Call .WriteByte(x)
        Call .WriteByte(y)

    End With

    
    Exit Sub

WriteTeleportCreate_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteTeleportCreate", Erl)
    Resume Next
    
End Sub

''
' Writes the "TeleportDestroy" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTeleportDestroy()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TeleportDestroy" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteTeleportDestroy_Err
    
    Call outgoingData.WriteByte(ClientPacketID.TeleportDestroy)

    
    Exit Sub

WriteTeleportDestroy_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteTeleportDestroy", Erl)
    Resume Next
    
End Sub

''
' Writes the "RainToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRainToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RainToggle" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteRainToggle_Err
    
    Call outgoingData.WriteByte(ClientPacketID.RainToggle)

    
    Exit Sub

WriteRainToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRainToggle", Erl)
    Resume Next
    
End Sub

''
' Writes the "SetCharDescription" message to the outgoing data buffer.
'
' @param    desc The description to set to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetCharDescription(ByVal desc As String)
    
    On Error GoTo WriteSetCharDescription_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SetCharDescription" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.SetCharDescription)
        
        Call .WriteASCIIString(desc)

    End With

    
    Exit Sub

WriteSetCharDescription_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSetCharDescription", Erl)
    Resume Next
    
End Sub

''
' Writes the "ForceMIDIToMap" message to the outgoing data buffer.
'
' @param    midiID The ID of the midi file to play.
' @param    map The map in which to play the given midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceMIDIToMap(ByVal midiID As Byte, ByVal map As Integer)
    
    On Error GoTo WriteForceMIDIToMap_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ForceMIDIToMap" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ForceMIDIToMap)
        
        Call .WriteByte(midiID)
        
        Call .WriteInteger(map)

    End With

    
    Exit Sub

WriteForceMIDIToMap_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteForceMIDIToMap", Erl)
    Resume Next
    
End Sub

''
' Writes the "ForceWAVEToMap" message to the outgoing data buffer.
'
' @param    waveID  The ID of the wave file to play.
' @param    Map     The map into which to play the given wave.
' @param    x       The position in the x axis in which to play the given wave.
' @param    y       The position in the y axis in which to play the given wave.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceWAVEToMap(ByVal waveID As Byte, ByVal map As Integer, ByVal x As Byte, ByVal y As Byte)
    
    On Error GoTo WriteForceWAVEToMap_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ForceWAVEToMap" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ForceWAVEToMap)
        
        Call .WriteByte(waveID)
        
        Call .WriteInteger(map)
        
        Call .WriteByte(x)
        Call .WriteByte(y)

    End With

    
    Exit Sub

WriteForceWAVEToMap_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteForceWAVEToMap", Erl)
    Resume Next
    
End Sub

''
' Writes the "RoyalArmyMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the royal army members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoyalArmyMessage(ByVal Message As String)
    
    On Error GoTo WriteRoyalArmyMessage_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RoyalArmyMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RoyalArmyMessage)
        
        Call .WriteASCIIString(Message)

    End With

    
    Exit Sub

WriteRoyalArmyMessage_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRoyalArmyMessage", Erl)
    Resume Next
    
End Sub

''
' Writes the "ChaosLegionMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the chaos legion member.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosLegionMessage(ByVal Message As String)
    
    On Error GoTo WriteChaosLegionMessage_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChaosLegionMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChaosLegionMessage)
        
        Call .WriteASCIIString(Message)

    End With

    
    Exit Sub

WriteChaosLegionMessage_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteChaosLegionMessage", Erl)
    Resume Next
    
End Sub

''
' Writes the "CitizenMessage" message to the outgoing data buffer.
'
' @param    message The message to send to citizens.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCitizenMessage(ByVal Message As String)
    
    On Error GoTo WriteCitizenMessage_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CitizenMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CitizenMessage)
        
        Call .WriteASCIIString(Message)

    End With

    
    Exit Sub

WriteCitizenMessage_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCitizenMessage", Erl)
    Resume Next
    
End Sub

''
' Writes the "CriminalMessage" message to the outgoing data buffer.
'
' @param    message The message to send to criminals.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCriminalMessage(ByVal Message As String)
    
    On Error GoTo WriteCriminalMessage_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CriminalMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CriminalMessage)
        
        Call .WriteASCIIString(Message)

    End With

    
    Exit Sub

WriteCriminalMessage_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCriminalMessage", Erl)
    Resume Next
    
End Sub

''
' Writes the "TalkAsNPC" message to the outgoing data buffer.
'
' @param    message The message to send to the royal army members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTalkAsNPC(ByVal Message As String)
    
    On Error GoTo WriteTalkAsNPC_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TalkAsNPC" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.TalkAsNPC)
        
        Call .WriteASCIIString(Message)

    End With

    
    Exit Sub

WriteTalkAsNPC_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteTalkAsNPC", Erl)
    Resume Next
    
End Sub

''
' Writes the "DestroyAllItemsInArea" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDestroyAllItemsInArea()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "DestroyAllItemsInArea" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteDestroyAllItemsInArea_Err
    
    Call outgoingData.WriteByte(ClientPacketID.DestroyAllItemsInArea)

    
    Exit Sub

WriteDestroyAllItemsInArea_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteDestroyAllItemsInArea", Erl)
    Resume Next
    
End Sub

''
' Writes the "AcceptRoyalCouncilMember" message to the outgoing data buffer.
'
' @param    username The name of the user to be accepted into the royal army council.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAcceptRoyalCouncilMember(ByVal UserName As String)
    
    On Error GoTo WriteAcceptRoyalCouncilMember_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AcceptRoyalCouncilMember" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.AcceptRoyalCouncilMember)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteAcceptRoyalCouncilMember_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteAcceptRoyalCouncilMember", Erl)
    Resume Next
    
End Sub

''
' Writes the "AcceptChaosCouncilMember" message to the outgoing data buffer.
'
' @param    username The name of the user to be accepted as a chaos council member.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAcceptChaosCouncilMember(ByVal UserName As String)
    
    On Error GoTo WriteAcceptChaosCouncilMember_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AcceptChaosCouncilMember" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.AcceptChaosCouncilMember)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteAcceptChaosCouncilMember_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteAcceptChaosCouncilMember", Erl)
    Resume Next
    
End Sub

''
' Writes the "ItemsInTheFloor" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteItemsInTheFloor()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ItemsInTheFloor" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteItemsInTheFloor_Err
    
    Call outgoingData.WriteByte(ClientPacketID.ItemsInTheFloor)

    
    Exit Sub

WriteItemsInTheFloor_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteItemsInTheFloor", Erl)
    Resume Next
    
End Sub

''
' Writes the "MakeDumb" message to the outgoing data buffer.
'
' @param    username The name of the user to be made dumb.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMakeDumb(ByVal UserName As String)
    
    On Error GoTo WriteMakeDumb_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "MakeDumb" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MakeDumb)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteMakeDumb_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteMakeDumb", Erl)
    Resume Next
    
End Sub

''
' Writes the "MakeDumbNoMore" message to the outgoing data buffer.
'
' @param    username The name of the user who will no longer be dumb.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMakeDumbNoMore(ByVal UserName As String)
    
    On Error GoTo WriteMakeDumbNoMore_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "MakeDumbNoMore" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MakeDumbNoMore)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteMakeDumbNoMore_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteMakeDumbNoMore", Erl)
    Resume Next
    
End Sub

''
' Writes the "DumpIPTables" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDumpIPTables()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "DumpIPTables" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteDumpIPTables_Err
    
    Call outgoingData.WriteByte(ClientPacketID.DumpIPTables)

    
    Exit Sub

WriteDumpIPTables_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteDumpIPTables", Erl)
    Resume Next
    
End Sub

''
' Writes the "CouncilKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the council.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCouncilKick(ByVal UserName As String)
    
    On Error GoTo WriteCouncilKick_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CouncilKick" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CouncilKick)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteCouncilKick_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCouncilKick", Erl)
    Resume Next
    
End Sub

''
' Writes the "SetTrigger" message to the outgoing data buffer.
'
' @param    trigger The type of trigger to be set to the tile.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetTrigger(ByVal Trigger As eTrigger)
    
    On Error GoTo WriteSetTrigger_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SetTrigger" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.SetTrigger)
        
        Call .WriteByte(Trigger)

    End With

    
    Exit Sub

WriteSetTrigger_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSetTrigger", Erl)
    Resume Next
    
End Sub

''
' Writes the "AskTrigger" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAskTrigger()
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 04/13/07
    'Writes the "AskTrigger" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteAskTrigger_Err
    
    Call outgoingData.WriteByte(ClientPacketID.AskTrigger)

    
    Exit Sub

WriteAskTrigger_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteAskTrigger", Erl)
    Resume Next
    
End Sub

''
' Writes the "BannedIPList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBannedIPList()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BannedIPList" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteBannedIPList_Err
    
    Call outgoingData.WriteByte(ClientPacketID.BannedIPList)

    
    Exit Sub

WriteBannedIPList_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBannedIPList", Erl)
    Resume Next
    
End Sub

''
' Writes the "BannedIPReload" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBannedIPReload()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BannedIPReload" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteBannedIPReload_Err
    
    Call outgoingData.WriteByte(ClientPacketID.BannedIPReload)

    
    Exit Sub

WriteBannedIPReload_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBannedIPReload", Erl)
    Resume Next
    
End Sub

''
' Writes the "GuildBan" message to the outgoing data buffer.
'
' @param    guild The guild whose members will be banned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildBan(ByVal guild As String)
    
    On Error GoTo WriteGuildBan_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildBan" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildBan)
        
        Call .WriteASCIIString(guild)

    End With

    
    Exit Sub

WriteGuildBan_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildBan", Erl)
    Resume Next
    
End Sub

''
' Writes the "BanIP" message to the outgoing data buffer.
'
' @param    byIp    If set to true, we are banning by IP, otherwise the ip of a given character.
' @param    IP      The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @param    nick    The nick of the player whose ip will be banned.
' @param    reason  The reason for the ban.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBanIP(ByVal byIp As Boolean, ByRef IP() As Byte, ByVal Nick As String, ByVal reason As String)
    
    On Error GoTo WriteBanIP_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BanIP" message to the outgoing data buffer
    '***************************************************
    If byIp And UBound(IP()) - LBound(IP()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.BanIP)
        
        Call .WriteBoolean(byIp)
        
        If byIp Then

            For i = LBound(IP()) To UBound(IP())
                Call .WriteByte(IP(i))
            Next i

        Else
            Call .WriteASCIIString(Nick)

        End If
        
        Call .WriteASCIIString(reason)

    End With

    
    Exit Sub

WriteBanIP_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBanIP", Erl)
    Resume Next
    
End Sub

''
' Writes the "UnbanIP" message to the outgoing data buffer.
'
' @param    IP The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUnbanIP(ByRef IP() As Byte)
    
    On Error GoTo WriteUnbanIP_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UnbanIP" message to the outgoing data buffer
    '***************************************************
    If UBound(IP()) - LBound(IP()) + 1 <> 4 Then Exit Sub   'Invalid IP
    
    Dim i As Long
    
    With outgoingData
        Call .WriteByte(ClientPacketID.UnBanIp)
        
        For i = LBound(IP()) To UBound(IP())
            Call .WriteByte(IP(i))
        Next i

    End With

    
    Exit Sub

WriteUnbanIP_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteUnbanIP", Erl)
    Resume Next
    
End Sub

''
' Writes the "CreateItem" message to the outgoing data buffer.
'
' @param    itemIndex The index of the item to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateItem(ByVal ItemIndex As Long, ByVal cantidad As Integer)
    
    On Error GoTo WriteCreateItem_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CreateItem" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CreateItem)
        
        Call .WriteInteger(ItemIndex)
        Call .WriteInteger(cantidad)

    End With

    
    Exit Sub

WriteCreateItem_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCreateItem", Erl)
    Resume Next
    
End Sub

''
' Writes the "DestroyItems" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDestroyItems()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "DestroyItems" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteDestroyItems_Err
    
    Call outgoingData.WriteByte(ClientPacketID.DestroyItems)

    
    Exit Sub

WriteDestroyItems_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteDestroyItems", Erl)
    Resume Next
    
End Sub

''
' Writes the "ChaosLegionKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the Chaos Legion.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosLegionKick(ByVal UserName As String)
    
    On Error GoTo WriteChaosLegionKick_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChaosLegionKick" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChaosLegionKick)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteChaosLegionKick_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteChaosLegionKick", Erl)
    Resume Next
    
End Sub

''
' Writes the "RoyalArmyKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the Royal Army.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoyalArmyKick(ByVal UserName As String)
    
    On Error GoTo WriteRoyalArmyKick_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RoyalArmyKick" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RoyalArmyKick)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteRoyalArmyKick_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRoyalArmyKick", Erl)
    Resume Next
    
End Sub

''
' Writes the "ForceMIDIAll" message to the outgoing data buffer.
'
' @param    midiID The id of the midi file to play.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceMIDIAll(ByVal midiID As Byte)
    
    On Error GoTo WriteForceMIDIAll_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ForceMIDIAll" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ForceMIDIAll)
        
        Call .WriteByte(midiID)

    End With

    
    Exit Sub

WriteForceMIDIAll_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteForceMIDIAll", Erl)
    Resume Next
    
End Sub

''
' Writes the "ForceWAVEAll" message to the outgoing data buffer.
'
' @param    waveID The id of the wave file to play.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceWAVEAll(ByVal waveID As Byte)
    
    On Error GoTo WriteForceWAVEAll_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ForceWAVEAll" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ForceWAVEAll)
        
        Call .WriteByte(waveID)

    End With

    
    Exit Sub

WriteForceWAVEAll_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteForceWAVEAll", Erl)
    Resume Next
    
End Sub

''
' Writes the "RemovePunishment" message to the outgoing data buffer.
'
' @param    username The user whose punishments will be altered.
' @param    punishment The id of the punishment to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemovePunishment(ByVal UserName As String, ByVal punishment As Byte, ByVal NewText As String)
    
    On Error GoTo WriteRemovePunishment_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RemovePunishment" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RemovePunishment)
        
        Call .WriteASCIIString(UserName)
        Call .WriteByte(punishment)
        Call .WriteASCIIString(NewText)

    End With

    
    Exit Sub

WriteRemovePunishment_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRemovePunishment", Erl)
    Resume Next
    
End Sub

''
' Writes the "TileBlockedToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTileBlockedToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TileBlockedToggle" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteTileBlockedToggle_Err
    
    Call outgoingData.WriteByte(ClientPacketID.TileBlockedToggle)

    
    Exit Sub

WriteTileBlockedToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteTileBlockedToggle", Erl)
    Resume Next
    
End Sub

''
' Writes the "KillNPCNoRespawn" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillNPCNoRespawn()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "KillNPCNoRespawn" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteKillNPCNoRespawn_Err
    
    Call outgoingData.WriteByte(ClientPacketID.KillNPCNoRespawn)

    
    Exit Sub

WriteKillNPCNoRespawn_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteKillNPCNoRespawn", Erl)
    Resume Next
    
End Sub

''
' Writes the "KillAllNearbyNPCs" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKillAllNearbyNPCs()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "KillAllNearbyNPCs" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteKillAllNearbyNPCs_Err
    
    Call outgoingData.WriteByte(ClientPacketID.KillAllNearbyNPCs)

    
    Exit Sub

WriteKillAllNearbyNPCs_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteKillAllNearbyNPCs", Erl)
    Resume Next
    
End Sub

''
' Writes the "LastIP" message to the outgoing data buffer.
'
' @param    username The user whose last IPs are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLastIP(ByVal UserName As String)
    
    On Error GoTo WriteLastIP_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "LastIP" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.LastIP)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteLastIP_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteLastIP", Erl)
    Resume Next
    
End Sub

''
' Writes the "ChangeMOTD" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMOTD()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeMOTD" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteChangeMOTD_Err
    
    Call outgoingData.WriteByte(ClientPacketID.ChangeMOTD)

    
    Exit Sub

WriteChangeMOTD_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteChangeMOTD", Erl)
    Resume Next
    
End Sub

''
' Writes the "SetMOTD" message to the outgoing data buffer.
'
' @param    message The message to be set as the new MOTD.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetMOTD(ByVal Message As String)
    
    On Error GoTo WriteSetMOTD_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SetMOTD" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.SetMOTD)
        
        Call .WriteASCIIString(Message)

    End With

    
    Exit Sub

WriteSetMOTD_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSetMOTD", Erl)
    Resume Next
    
End Sub

''
' Writes the "SystemMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to all players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSystemMessage(ByVal Message As String)
    
    On Error GoTo WriteSystemMessage_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SystemMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.SystemMessage)
        
        Call .WriteASCIIString(Message)

    End With

    
    Exit Sub

WriteSystemMessage_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSystemMessage", Erl)
    Resume Next
    
End Sub

''
' Writes the "CreateNPC" message to the outgoing data buffer.
'
' @param    npcIndex The index of the NPC to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNPC(ByVal NpcIndex As Integer)
    
    On Error GoTo WriteCreateNPC_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CreateNPC" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CreateNPC)
        
        Call .WriteInteger(NpcIndex)

    End With

    
    Exit Sub

WriteCreateNPC_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCreateNPC", Erl)
    Resume Next
    
End Sub

''
' Writes the "CreateNPCWithRespawn" message to the outgoing data buffer.
'
' @param    npcIndex The index of the NPC to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNPCWithRespawn(ByVal NpcIndex As Integer)
    
    On Error GoTo WriteCreateNPCWithRespawn_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CreateNPCWithRespawn" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CreateNPCWithRespawn)
        
        Call .WriteInteger(NpcIndex)

    End With

    
    Exit Sub

WriteCreateNPCWithRespawn_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCreateNPCWithRespawn", Erl)
    Resume Next
    
End Sub

''
' Writes the "ImperialArmour" message to the outgoing data buffer.
'
' @param    armourIndex The index of imperial armour to be altered.
' @param    objectIndex The index of the new object to be set as the imperial armour.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteImperialArmour(ByVal armourIndex As Byte, ByVal objectIndex As Integer)
    
    On Error GoTo WriteImperialArmour_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ImperialArmour" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ImperialArmour)
        
        Call .WriteByte(armourIndex)
        
        Call .WriteInteger(objectIndex)

    End With

    
    Exit Sub

WriteImperialArmour_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteImperialArmour", Erl)
    Resume Next
    
End Sub

''
' Writes the "ChaosArmour" message to the outgoing data buffer.
'
' @param    armourIndex The index of chaos armour to be altered.
' @param    objectIndex The index of the new object to be set as the chaos armour.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosArmour(ByVal armourIndex As Byte, ByVal objectIndex As Integer)
    
    On Error GoTo WriteChaosArmour_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChaosArmour" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChaosArmour)
        
        Call .WriteByte(armourIndex)
        
        Call .WriteInteger(objectIndex)

    End With

    
    Exit Sub

WriteChaosArmour_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteChaosArmour", Erl)
    Resume Next
    
End Sub

''
' Writes the "NavigateToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNavigateToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "NavigateToggle" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteNavigateToggle_Err
    
    Call outgoingData.WriteByte(ClientPacketID.NavigateToggle)

    
    Exit Sub

WriteNavigateToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteNavigateToggle", Erl)
    Resume Next
    
End Sub

' Writes the "ServerOpenToUsersToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerOpenToUsersToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ServerOpenToUsersToggle" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteServerOpenToUsersToggle_Err
    
    Call outgoingData.WriteByte(ClientPacketID.ServerOpenToUsersToggle)

    
    Exit Sub

WriteServerOpenToUsersToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteServerOpenToUsersToggle", Erl)
    Resume Next
    
End Sub

''
' Writes the "Participar" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteParticipar()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TurnOffServer" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteParticipar_Err
    
    Call outgoingData.WriteByte(ClientPacketID.Participar)

    
    Exit Sub

WriteParticipar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteParticipar", Erl)
    Resume Next
    
End Sub

''
' Writes the "TurnCriminal" message to the outgoing data buffer.
'
' @param    username The name of the user to turn into criminal.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTurnCriminal(ByVal UserName As String)
    
    On Error GoTo WriteTurnCriminal_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TurnCriminal" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.TurnCriminal)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteTurnCriminal_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteTurnCriminal", Erl)
    Resume Next
    
End Sub

''
' Writes the "ResetFactions" message to the outgoing data buffer.
'
' @param    username The name of the user who will be removed from any faction.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetFactions(ByVal UserName As String)
    
    On Error GoTo WriteResetFactions_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ResetFactions" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ResetFactions)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteResetFactions_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteResetFactions", Erl)
    Resume Next
    
End Sub

''
' Writes the "RemoveCharFromGuild" message to the outgoing data buffer.
'
' @param    username The name of the user who will be removed from any guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveCharFromGuild(ByVal UserName As String)
    
    On Error GoTo WriteRemoveCharFromGuild_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RemoveCharFromGuild" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RemoveCharFromGuild)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteRemoveCharFromGuild_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRemoveCharFromGuild", Erl)
    Resume Next
    
End Sub

''
' Writes the "RequestCharMail" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharMail(ByVal UserName As String)
    
    On Error GoTo WriteRequestCharMail_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharMail" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RequestCharMail)
        
        Call .WriteASCIIString(UserName)

    End With

    
    Exit Sub

WriteRequestCharMail_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestCharMail", Erl)
    Resume Next
    
End Sub

''
' Writes the "AlterPassword" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    copyFrom The name of the user from which to copy the password.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterPassword(ByVal UserName As String, ByVal CopyFrom As String)
    
    On Error GoTo WriteAlterPassword_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AlterPassword" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.AlterPassword)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(CopyFrom)

    End With

    
    Exit Sub

WriteAlterPassword_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteAlterPassword", Erl)
    Resume Next
    
End Sub

''
' Writes the "AlterMail" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    newMail The new email of the player.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterMail(ByVal UserName As String, ByVal newMail As String)
    
    On Error GoTo WriteAlterMail_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AlterMail" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.AlterMail)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(newMail)

    End With

    
    Exit Sub

WriteAlterMail_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteAlterMail", Erl)
    Resume Next
    
End Sub

''
' Writes the "AlterName" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    newName The new user name.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterName(ByVal UserName As String, ByVal newName As String)
    
    On Error GoTo WriteAlterName_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AlterName" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.AlterName)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(newName)

    End With

    
    Exit Sub

WriteAlterName_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteAlterName", Erl)
    Resume Next
    
End Sub

''
' Writes the "DoBackup" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDoBackup()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "DoBackup" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteDoBackup_Err
    
    Call outgoingData.WriteByte(ClientPacketID.DoBackUp)

    
    Exit Sub

WriteDoBackup_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteDoBackup", Erl)
    Resume Next
    
End Sub

''
' Writes the "ShowGuildMessages" message to the outgoing data buffer.
'
' @param    guild The guild to listen to.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGuildMessages(ByVal guild As String)
    
    On Error GoTo WriteShowGuildMessages_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowGuildMessages" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ShowGuildMessages)
        
        Call .WriteASCIIString(guild)

    End With

    
    Exit Sub

WriteShowGuildMessages_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteShowGuildMessages", Erl)
    Resume Next
    
End Sub

''
' Writes the "SaveMap" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSaveMap()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SaveMap" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteSaveMap_Err
    
    Call outgoingData.WriteByte(ClientPacketID.SaveMap)

    
    Exit Sub

WriteSaveMap_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSaveMap", Erl)
    Resume Next
    
End Sub

''
' Writes the "ChangeMapInfoPK" message to the outgoing data buffer.
'
' @param    isPK True if the map is PK, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoPK(ByVal isPK As Boolean)
    
    On Error GoTo WriteChangeMapInfoPK_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeMapInfoPK" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeMapInfoPK)
        
        Call .WriteBoolean(isPK)

    End With

    
    Exit Sub

WriteChangeMapInfoPK_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteChangeMapInfoPK", Erl)
    Resume Next
    
End Sub

''
' Writes the "ChangeMapInfoBackup" message to the outgoing data buffer.
'
' @param    backup True if the map is to be backuped, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoBackup(ByVal backup As Boolean)
    
    On Error GoTo WriteChangeMapInfoBackup_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeMapInfoBackup" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeMapInfoBackup)
        
        Call .WriteBoolean(backup)

    End With

    
    Exit Sub

WriteChangeMapInfoBackup_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteChangeMapInfoBackup", Erl)
    Resume Next
    
End Sub

''
' Writes the "ChangeMapInfoRestricted" message to the outgoing data buffer.
'
' @param    restrict NEWBIES (only newbies), NO (everyone), ARMADA (just Armadas), CAOS (just caos) or FACCION (Armadas & caos only)
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoRestricted(ByVal restrict As String)
    
    On Error GoTo WriteChangeMapInfoRestricted_Err
    

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Writes the "ChangeMapInfoRestricted" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeMapInfoRestricted)
        
        Call .WriteASCIIString(restrict)

    End With

    
    Exit Sub

WriteChangeMapInfoRestricted_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteChangeMapInfoRestricted", Erl)
    Resume Next
    
End Sub

''
' Writes the "ChangeMapInfoNoMagic" message to the outgoing data buffer.
'
' @param    nomagic TRUE if no magic is to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoMagic(ByVal nomagic As Boolean)
    
    On Error GoTo WriteChangeMapInfoNoMagic_Err
    

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Writes the "ChangeMapInfoNoMagic" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeMapInfoNoMagic)
        
        Call .WriteBoolean(nomagic)

    End With

    
    Exit Sub

WriteChangeMapInfoNoMagic_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteChangeMapInfoNoMagic", Erl)
    Resume Next
    
End Sub

''
' Writes the "ChangeMapInfoNoInvi" message to the outgoing data buffer.
'
' @param    noinvi TRUE if invisibility is not to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoInvi(ByVal noinvi As Boolean)
    
    On Error GoTo WriteChangeMapInfoNoInvi_Err
    

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Writes the "ChangeMapInfoNoInvi" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeMapInfoNoInvi)
        
        Call .WriteBoolean(noinvi)

    End With

    
    Exit Sub

WriteChangeMapInfoNoInvi_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteChangeMapInfoNoInvi", Erl)
    Resume Next
    
End Sub
                            
''
' Writes the "ChangeMapInfoNoResu" message to the outgoing data buffer.
'
' @param    noresu TRUE if resurection is not to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoResu(ByVal noresu As Boolean)
    
    On Error GoTo WriteChangeMapInfoNoResu_Err
    

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Writes the "ChangeMapInfoNoResu" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeMapInfoNoResu)
        
        Call .WriteBoolean(noresu)

    End With

    
    Exit Sub

WriteChangeMapInfoNoResu_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteChangeMapInfoNoResu", Erl)
    Resume Next
    
End Sub
                        
''
' Writes the "ChangeMapInfoLand" message to the outgoing data buffer.
'
' @param    land options: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoLand(ByVal lAnd As String)
    
    On Error GoTo WriteChangeMapInfoLand_Err
    

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Writes the "ChangeMapInfoLand" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeMapInfoLand)
        
        Call .WriteASCIIString(lAnd)

    End With

    
    Exit Sub

WriteChangeMapInfoLand_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteChangeMapInfoLand", Erl)
    Resume Next
    
End Sub
                        
''
' Writes the "ChangeMapInfoZone" message to the outgoing data buffer.
'
' @param    zone options: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoZone(ByVal zone As String)
    
    On Error GoTo WriteChangeMapInfoZone_Err
    

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Writes the "ChangeMapInfoZone" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeMapInfoZone)
        
        Call .WriteASCIIString(zone)

    End With

    
    Exit Sub

WriteChangeMapInfoZone_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteChangeMapInfoZone", Erl)
    Resume Next
    
End Sub

''
' Writes the "SaveChars" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSaveChars()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SaveChars" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteSaveChars_Err
    
    Call outgoingData.WriteByte(ClientPacketID.SaveChars)

    
    Exit Sub

WriteSaveChars_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSaveChars", Erl)
    Resume Next
    
End Sub

''
' Writes the "CleanSOS" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCleanSOS()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CleanSOS" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteCleanSOS_Err
    
    Call outgoingData.WriteByte(ClientPacketID.CleanSOS)

    
    Exit Sub

WriteCleanSOS_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCleanSOS", Erl)
    Resume Next
    
End Sub

''
' Writes the "ShowServerForm" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowServerForm()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowServerForm" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteShowServerForm_Err
    
    Call outgoingData.WriteByte(ClientPacketID.ShowServerForm)

    
    Exit Sub

WriteShowServerForm_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteShowServerForm", Erl)
    Resume Next
    
End Sub

''
' Writes the "Night" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNight()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Night" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteNight_Err
    
    Call outgoingData.WriteByte(ClientPacketID.night)

    
    Exit Sub

WriteNight_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteNight", Erl)
    Resume Next
    
End Sub

Public Sub WriteDay()
    
    On Error GoTo WriteDay_Err
    
    Call outgoingData.WriteByte(ClientPacketID.Day)
    
    Exit Sub

WriteDay_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteDay", Erl)
    Resume Next
    
End Sub

Public Sub WriteSetTime(ByVal Time As Long)
    
    On Error GoTo WriteSetTime_Err
    
    With outgoingData
        Call .WriteByte(ClientPacketID.SetTime)
        Call .WriteLong(Time)
    End With
    
    Exit Sub

WriteSetTime_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSetTime", Erl)
    Resume Next
    
End Sub

''
' Writes the "KickAllChars" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKickAllChars()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "KickAllChars" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteKickAllChars_Err
    
    Call outgoingData.WriteByte(ClientPacketID.KickAllChars)

    
    Exit Sub

WriteKickAllChars_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteKickAllChars", Erl)
    Resume Next
    
End Sub

''
' Writes the "RequestTCPStats" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestTCPStats()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestTCPStats" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteRequestTCPStats_Err
    
    Call outgoingData.WriteByte(ClientPacketID.RequestTCPStats)

    
    Exit Sub

WriteRequestTCPStats_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestTCPStats", Erl)
    Resume Next
    
End Sub

''
' Writes the "ReloadNPCs" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadNPCs()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ReloadNPCs" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteReloadNPCs_Err
    
    Call outgoingData.WriteByte(ClientPacketID.ReloadNPCs)

    
    Exit Sub

WriteReloadNPCs_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteReloadNPCs", Erl)
    Resume Next
    
End Sub

''
' Writes the "ReloadServerIni" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadServerIni()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ReloadServerIni" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteReloadServerIni_Err
    
    Call outgoingData.WriteByte(ClientPacketID.ReloadServerIni)

    
    Exit Sub

WriteReloadServerIni_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteReloadServerIni", Erl)
    Resume Next
    
End Sub

''
' Writes the "ReloadSpells" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadSpells()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ReloadSpells" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteReloadSpells_Err
    
    Call outgoingData.WriteByte(ClientPacketID.ReloadSpells)

    
    Exit Sub

WriteReloadSpells_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteReloadSpells", Erl)
    Resume Next
    
End Sub

''
' Writes the "ReloadObjects" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReloadObjects()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ReloadObjects" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteReloadObjects_Err
    
    Call outgoingData.WriteByte(ClientPacketID.ReloadObjects)

    
    Exit Sub

WriteReloadObjects_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteReloadObjects", Erl)
    Resume Next
    
End Sub

''
' Writes the "Restart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRestart()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Restart" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteRestart_Err
    
    Call outgoingData.WriteByte(ClientPacketID.Restart)

    
    Exit Sub

WriteRestart_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRestart", Erl)
    Resume Next
    
End Sub

''
' Writes the "ResetAutoUpdate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetAutoUpdate()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ResetAutoUpdate" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteResetAutoUpdate_Err
    
    Call outgoingData.WriteByte(ClientPacketID.ResetAutoUpdate)

    
    Exit Sub

WriteResetAutoUpdate_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteResetAutoUpdate", Erl)
    Resume Next
    
End Sub

''
' Writes the "ChatColor" message to the outgoing data buffer.
'
' @param    r The red component of the new chat color.
' @param    g The green component of the new chat color.
' @param    b The blue component of the new chat color.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChatColor(ByVal r As Byte, ByVal G As Byte, ByVal B As Byte)
    
    On Error GoTo WriteChatColor_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChatColor" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChatColor)
        
        Call .WriteByte(r)
        Call .WriteByte(G)
        Call .WriteByte(B)

    End With

    
    Exit Sub

WriteChatColor_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteChatColor", Erl)
    Resume Next
    
End Sub

''
' Writes the "Ignored" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteIgnored()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Ignored" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteIgnored_Err
    
    Call outgoingData.WriteByte(ClientPacketID.Ignored)

    
    Exit Sub

WriteIgnored_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteIgnored", Erl)
    Resume Next
    
End Sub

''
' Writes the "CheckSlot" message to the outgoing data buffer.
'
' @param    UserName    The name of the char whose slot will be checked.
' @param    slot        The slot to be checked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCheckSlot(ByVal UserName As String, ByVal Slot As Byte)
    
    On Error GoTo WriteCheckSlot_Err
    

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Writes the "CheckSlot" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CheckSlot)
        Call .WriteASCIIString(UserName)
        Call .WriteByte(Slot)

    End With

    
    Exit Sub

WriteCheckSlot_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCheckSlot", Erl)
    Resume Next
    
End Sub

''
' Writes the "Ping" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePing()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 26/01/2007
    'Writes the "Ping" message to the outgoing data buffer
    '***************************************************
    'Prevent the timer from being cut
    '   If pingTime <> 0 Then Exit Sub
    
    On Error GoTo WritePing_Err
    

    Call outgoingData.WriteByte(ClientPacketID.Ping)
    pingTime = timeGetTime And &H7FFFFFFF
    Call outgoingData.WriteLong(pingTime)
    
    ' Avoid computing errors due to frame rate
    Call FlushBuffer
    'DoEvents
    
    
    Exit Sub

WritePing_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WritePing", Erl)
    Resume Next
    
End Sub

Public Sub WriteLlamadadeClan()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 26/01/2007
    'Writes the "Ping" message to the outgoing data buffer
    '***************************************************
    'Prevent the timer from being cut
    '   If pingTime <> 0 Then Exit Sub
    
    On Error GoTo WriteLlamadadeClan_Err
    

    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.llamadadeclan)
    
    ' Avoid computing errors due to frame rate
    Call FlushBuffer
    'DoEvents
    
    
    Exit Sub

WriteLlamadadeClan_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteLlamadadeClan", Erl)
    Resume Next
    
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
    Dim sndData As String
    
    With outgoingData

        If .length = 0 Then Exit Sub
        '   Debug.Print "Salio paquete con peso de: " & .Length & " bytes"
        OutBytes = OutBytes + .length
        ' InBytes = 0
        sndData = .ReadASCIIStringFixed(.length)
        
        Call SendData(sndData)

    End With

    
    Exit Sub

FlushBuffer_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.FlushBuffer", Erl)
    Resume Next
    
End Sub

''
' Sends the data using the socket controls in the MainForm.
'
' @param    sdData  The data to be sent to the server.

Private Sub SendData(ByRef sdData As String)
    
    On Error GoTo SendData_Err
    
    #If UsarWrench = 1 Then

        If Not frmMain.Socket1.IsWritable Then
            'Put data back in the bytequeue
            Call outgoingData.WriteASCIIStringFixed(sdData)
            Exit Sub

        End If
   
        If Not frmMain.Socket1.Connected Then Exit Sub
    #Else

        If frmMain.Winsock1.State <> sckConnected Then Exit Sub
    #End If
 
    #If AntiExternos Then
        Security.Redundance = CLng(Security.Redundance * Security.MultiplicationFactor) Mod 255

        Dim Data() As Byte: Data = StrConv(sdData, vbFromUnicode)
        Call Security.NAC_E_Byte(Data, Security.Redundance)
        
        sdData = StrConv(Data, vbUnicode)

    #End If
 
    #If UsarWrench = 1 Then
        Call frmMain.Socket1.Write(sdData, Len(sdData))
    #Else
        Call frmMain.Winsock1.SendData(sdData)
    #End If
 
    
    Exit Sub

SendData_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.SendData", Erl)
    Resume Next
    
End Sub

Public Sub WriteQuestionGM(ByVal Consulta As String, ByVal TipoDeConsulta As String)
    
    On Error GoTo WriteQuestionGM_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ForumPost" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.QuestionGM)
        Call .WriteASCIIString(Consulta)
        Call .WriteASCIIString(TipoDeConsulta)

    End With

    
    Exit Sub

WriteQuestionGM_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteQuestionGM", Erl)
    Resume Next
    
End Sub

Public Sub WriteOfertaInicial(ByVal Oferta As Long)
    
    On Error GoTo WriteOfertaInicial_Err
    

    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.OfertaInicial)
        Call .WriteLong(Oferta)

    End With

    
    Exit Sub

WriteOfertaInicial_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteOfertaInicial", Erl)
    Resume Next
    
End Sub

Public Sub WriteOferta(ByVal OfertaDeSubasta As Long)
    
    On Error GoTo WriteOferta_Err
    

    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.OfertaDeSubasta)
        Call .WriteLong(OfertaDeSubasta)

    End With

    
    Exit Sub

WriteOferta_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteOferta", Erl)
    Resume Next
    
End Sub

Public Sub WriteGlobalMessage(ByVal Message As String)
    
    On Error GoTo WriteGlobalMessage_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRequestDetails" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GlobalMessage)
        
        Call .WriteASCIIString(Message)
        
    End With

    
    Exit Sub

WriteGlobalMessage_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGlobalMessage", Erl)
    Resume Next
    
End Sub

Public Sub WriteGlobalOnOff()
    
    On Error GoTo WriteGlobalOnOff_Err
    
    Call outgoingData.WriteByte(ClientPacketID.GlobalOnOff)

    
    Exit Sub

WriteGlobalOnOff_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGlobalOnOff", Erl)
    Resume Next
    
End Sub

Public Sub WriteNuevaCuenta()
    
    On Error GoTo WriteNuevaCuenta_Err
    

    With outgoingData
        Call .WriteByte(ClientPacketID.CrearNuevaCuenta)
    
        Call .WriteASCIIString(CuentaEmail)
        Call .WriteASCIIString(SEncriptar(CuentaPassword))
        Call .WriteASCIIString(CuentaEmail)

    End With

    
    Exit Sub

WriteNuevaCuenta_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteNuevaCuenta", Erl)
    Resume Next
    
End Sub

Public Sub WriteValidarCuenta()
    
    On Error GoTo WriteValidarCuenta_Err
    

    With outgoingData
        Call .WriteByte(ClientPacketID.validarCuenta)
        Call .WriteASCIIString(CuentaEmail)
        Call .WriteASCIIString(ValidacionCode)

    End With

    
    Exit Sub

WriteValidarCuenta_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteValidarCuenta", Erl)
    Resume Next
    
End Sub

Public Sub WriteReValidarCuenta()
    
    On Error GoTo WriteReValidarCuenta_Err
    

    With outgoingData
        Call .WriteByte(ClientPacketID.RevalidarCuenta)
        Call .WriteASCIIString(CuentaEmail)

    End With

    
    Exit Sub

WriteReValidarCuenta_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteReValidarCuenta", Erl)
    Resume Next
    
End Sub

Public Sub WriteRecuperandoConstraseña()
    
    On Error GoTo WriteRecuperandoConstraseña_Err
    

    With outgoingData
        Call .WriteByte(ClientPacketID.RecuperandoConstraseña)
        Call .WriteASCIIString(CuentaEmail)

    End With

    
    Exit Sub

WriteRecuperandoConstraseña_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRecuperandoConstraseña", Erl)
    Resume Next
    
End Sub

Public Sub WriteBorrandoCuenta()
    
    On Error GoTo WriteBorrandoCuenta_Err
    

    With outgoingData
        Call .WriteByte(ClientPacketID.BorrandoCuenta)
        Call .WriteASCIIString(CuentaEmail)
        Call .WriteASCIIString(SEncriptar(CuentaPassword))
        Call .WriteASCIIString(CheckMD5)

    End With

    
    Exit Sub

WriteBorrandoCuenta_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBorrandoCuenta", Erl)
    Resume Next
    
End Sub

Public Sub WriteBorrandoPJ()
    
    On Error GoTo WriteBorrandoPJ_Err
    

    With outgoingData
        Call .WriteByte(ClientPacketID.BorrarPJ)
        Call .WriteASCIIString(DeleteUser)
        Call .WriteASCIIString(CuentaEmail)
        Call .WriteASCIIString(SEncriptar(CuentaPassword))
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
        Call .WriteASCIIString(MacAdress)  'Seguridad
        Call .WriteLong(HDserial)  'SeguridadHDserial
        Call .WriteASCIIString(CheckMD5)

    End With

    
    Exit Sub

WriteBorrandoPJ_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBorrandoPJ", Erl)
    Resume Next
    
End Sub

Public Sub WriteIngresandoConCuenta()
    
    On Error GoTo WriteIngresandoConCuenta_Err
    

    With outgoingData
        Call .WriteByte(ClientPacketID.IngresarConCuenta)
        Call .WriteASCIIString(CuentaEmail)
        Call .WriteASCIIString(SEncriptar(CuentaPassword))
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
        Call .WriteASCIIString(MacAdress)  'Seguridad
        Call .WriteLong(HDserial)  'SeguridadHDserial
        Call .WriteASCIIString(CheckMD5)
        
    End With

    
    Exit Sub

WriteIngresandoConCuenta_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteIngresandoConCuenta", Erl)
    Resume Next
    
End Sub

Private Sub HandlePersonajesDeCuenta()

    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    'Remove packet ID
    Call buffer.ReadByte
    
    CantidadDePersonajesEnCuenta = buffer.ReadByte()

    Dim ii As Byte
     
    For ii = 1 To 10
        Pjs(ii).Body = 0
        Pjs(ii).Head = 0
        Pjs(ii).Mapa = 0
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
        Pjs(ii).nombre = buffer.ReadASCIIString()
        Pjs(ii).nivel = buffer.ReadByte()
        Pjs(ii).Mapa = buffer.ReadInteger()
        Pjs(ii).Body = buffer.ReadInteger()
        
        Pjs(ii).Head = buffer.ReadInteger()
        Pjs(ii).Criminal = buffer.ReadByte()
        Pjs(ii).Clase = buffer.ReadByte()
       
        Pjs(ii).Casco = buffer.ReadInteger()
        Pjs(ii).Escudo = buffer.ReadInteger()
        Pjs(ii).Arma = buffer.ReadInteger()
        Pjs(ii).ClanName = "<" & buffer.ReadASCIIString() & ">"
       
        ' Pjs(ii).NameMapa = Pjs(ii).mapa
        Pjs(ii).NameMapa = NameMaps(Pjs(ii).Mapa).Name

    Next ii
    
    CuentaDonador = buffer.ReadByte()
    
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
   
    If CantidadDePersonajesEnCuenta > 0 Then
        PJSeleccionado = 1

    End If

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

Private Sub HandleUserOnline()

    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    'Remove packet ID
    Call buffer.ReadByte

    Dim rdata As Integer
    
    rdata = buffer.ReadInteger()
    
    usersOnline = rdata
    frmMain.onlines = "Online: " & usersOnline
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

Private Sub HandleParticleFXToFloor()
    
    On Error GoTo HandleParticleFXToFloor_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 8 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
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
    Resume Next
    
End Sub

Private Sub HandleLightToFloor()
    
    On Error GoTo HandleLightToFloor_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 8 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim x     As Byte

    Dim y     As Byte

    Dim Color As Long
    
    Dim color_value As RGBA

    Dim Rango As Byte
     
    x = incomingData.ReadByte()
    y = incomingData.ReadByte()
    Color = incomingData.ReadLong()
    Rango = incomingData.ReadByte()
    
    Call Long_2_RGBA(color_value, Color)

    Dim id  As Long

    Dim id2 As Long

    If Color = 0 Then
   
        If MapData(x, y).luz.Rango > 100 Then
            LucesRedondas.Delete_Light_To_Map x, y
   
            LucesCuadradas.Light_Render_All
            LucesRedondas.LightRenderAll
            Exit Sub
        Else
            id = LucesCuadradas.Light_Find(x & y)
            LucesCuadradas.Light_Remove id
            MapData(x, y).luz.Color = COLOR_EMPTY
            MapData(x, y).luz.Rango = 0
            LucesCuadradas.Light_Render_All
            Exit Sub

        End If

    End If
    
    MapData(x, y).luz.Color = color_value
    MapData(x, y).luz.Rango = Rango
    
    If Rango < 100 Then
        id = x & y
        LucesCuadradas.Light_Create x, y, color_value, Rango, id
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
    Resume Next
    
End Sub

Private Sub HandleParticleFX()
    
    On Error GoTo HandleParticleFX_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 10 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim charindex      As Integer

    Dim ParticulaIndex As Integer

    Dim Time           As Long

    Dim Remove         As Boolean
     
    charindex = incomingData.ReadInteger()
    ParticulaIndex = incomingData.ReadInteger()
    Time = incomingData.ReadLong()
    Remove = incomingData.ReadBoolean()
    
    If Remove Then
        Call Char_Particle_Group_Remove(charindex, ParticulaIndex)
        charlist(charindex).Particula = 0
    
    Else
        charlist(charindex).Particula = ParticulaIndex
        charlist(charindex).ParticulaTime = Time
     
        Call General_Char_Particle_Create(ParticulaIndex, charindex, Time)

    End If
    
    
    Exit Sub

HandleParticleFX_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleParticleFX", Erl)
    Resume Next
    
End Sub

Private Sub HandleParticleFXWithDestino()
    
    On Error GoTo HandleParticleFXWithDestino_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 17 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
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
    Resume Next
    
End Sub

Private Sub HandleParticleFXWithDestinoXY()
    
    On Error GoTo HandleParticleFXWithDestinoXY_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 17 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
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
    Resume Next
    
End Sub

Private Sub HandleAuraToChar()
    
    On Error GoTo HandleAuraToChar_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/0
    '
    '***************************************************
    If incomingData.length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
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
    Resume Next
    
End Sub

Private Sub HandleSpeedToChar()
    
    On Error GoTo HandleSpeedToChar_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/0
    '
    '***************************************************
    If incomingData.length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim charindex As Integer

    Dim Speeding  As Single
     
    charindex = incomingData.ReadInteger()
    Speeding = incomingData.ReadSingle()
   
    charlist(charindex).Speeding = Speeding

    
    Exit Sub

HandleSpeedToChar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSpeedToChar", Erl)
    Resume Next
    
End Sub

Public Sub WriteNieveToggle()
    
    On Error GoTo WriteNieveToggle_Err
    

    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.NieveToggle)

    End With

    
    Exit Sub

WriteNieveToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteNieveToggle", Erl)
    Resume Next
    
End Sub

Private Sub HandleNieveToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleNieveToggle_Err
    
    
    Call incomingData.ReadByte
    
    If Not InMapBounds(UserPos.x, UserPos.y) Then Exit Sub
            
    If MapDat.NIEVE Then
        Engine_MeteoParticle_Set (Particula_Nieve)

    End If

    bNieve = Not bNieve
  
    
    Exit Sub

HandleNieveToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleNieveToggle", Erl)
    Resume Next
    
End Sub

Private Sub HandleNieblaToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleNieblaToggle_Err
    
    
    Call incomingData.ReadByte
    
    MaxAlphaNiebla = incomingData.ReadByte()
            
    bNiebla = Not bNiebla
    frmMain.TimerNiebla.Enabled = True
  
    
    Exit Sub

HandleNieblaToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleNieblaToggle", Erl)
    Resume Next
    
End Sub

Public Sub WriteNieblaToggle()
    
    On Error GoTo WriteNieblaToggle_Err
    

    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.NieblaToggle)
    End With

    
    Exit Sub

WriteNieblaToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteNieblaToggle", Erl)
    Resume Next
    
End Sub

Public Sub WriteGenio()
    '***************************************************
    '/GENIO
    'Ladder
    '***************************************************
    
    On Error GoTo WriteGenio_Err
    
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.Genio)

    
    Exit Sub

WriteGenio_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGenio", Erl)
    Resume Next
    
End Sub

Private Sub HandleFamiliar()
    
    On Error GoTo HandleFamiliar_Err
    

    If incomingData.length < 1 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    
    Exit Sub

HandleFamiliar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleFamiliar", Erl)
    Resume Next
    
End Sub

Private Sub HandleBindKeys()
    
    On Error GoTo HandleBindKeys_Err
    

    '***************************************************
    'Macros
    'Pablo Mercavides
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
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
    Resume Next
    
End Sub

Private Sub HandleLogros()
    
    On Error GoTo HandleLogros_Err
    

    '***************************************************
    'Pablo Mercavides
    '***************************************************
    If incomingData.length < 40 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
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
    Resume Next
    
End Sub

Private Sub HandleBarFx()
    
    On Error GoTo HandleBarFx_Err
    

    '***************************************************
    'Author: Pablo Mercavides
    '***************************************************
    If incomingData.length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim charindex As Integer

    Dim BarTime   As Integer

    Dim BarAccion As Byte
    
    charindex = incomingData.ReadInteger()
    BarTime = incomingData.ReadInteger()
    BarAccion = incomingData.ReadByte()
    
    charlist(charindex).BarTime = 0
    charlist(charindex).BarAccion = BarAccion
    charlist(charindex).MaxBarTime = BarTime

    
    Exit Sub

HandleBarFx_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBarFx", Erl)
    Resume Next
    
End Sub

Public Sub WriteCompletarAccion(ByVal Accion As Byte)
    
    On Error GoTo WriteCompletarAccion_Err
    

    '***************************************************
    'Author: Pablo Mercavides
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.CompletarAccion)
        Call .WriteByte(Accion)

    End With

    
    Exit Sub

WriteCompletarAccion_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCompletarAccion", Erl)
    Resume Next
    
End Sub

Public Sub WriteTraerRecompensas()
    
    On Error GoTo WriteTraerRecompensas_Err
    

    '***************************************************
    'Ladder
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.TraerRecompensas)

    End With

    
    Exit Sub

WriteTraerRecompensas_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteTraerRecompensas", Erl)
    Resume Next
    
End Sub

Public Sub WriteTraerShop()
    
    On Error GoTo WriteTraerShop_Err
    

    '***************************************************
    'Ladder
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.Traershop)

    End With

    
    Exit Sub

WriteTraerShop_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteTraerShop", Erl)
    Resume Next
    
End Sub

Public Sub WriteTraerRanking()
    
    On Error GoTo WriteTraerRanking_Err
    

    '***************************************************
    'Ladder
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.TraerRanking)

    End With

    
    Exit Sub

WriteTraerRanking_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteTraerRanking", Erl)
    Resume Next
    
End Sub

Public Sub WritePareja()
    
    On Error GoTo WritePareja_Err
    

    '***************************************************
    'Ladder
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.Pareja)

    End With

    
    Exit Sub

WritePareja_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WritePareja", Erl)
    Resume Next
    
End Sub

Public Sub WriteQuest()
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Escribe el paquete Quest al servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    
    On Error GoTo WriteQuest_Err
    
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.Quest)

    
    Exit Sub

WriteQuest_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteQuest", Erl)
    Resume Next
    
End Sub
 
Public Sub WriteQuestDetailsRequest(ByVal QuestSlot As Byte)
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Escribe el paquete QuestDetailsRequest al servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    
    On Error GoTo WriteQuestDetailsRequest_Err
    
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.QuestDetailsRequest)
    Call outgoingData.WriteByte(QuestSlot)

    
    Exit Sub

WriteQuestDetailsRequest_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteQuestDetailsRequest", Erl)
    Resume Next
    
End Sub
 
Public Sub WriteQuestAccept(ByVal ListInd As Byte)
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Escribe el paquete QuestAccept al servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    
    On Error GoTo WriteQuestAccept_Err
    
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.QuestAccept)
    Call outgoingData.WriteByte(ListInd)

    
    Exit Sub

WriteQuestAccept_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteQuestAccept", Erl)
    Resume Next
    
End Sub
 
Private Sub HandleQuestDetails()

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Recibe y maneja el paquete QuestDetails del servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If incomingData.length < 15 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    Dim tmpStr        As String

    Dim tmpByte       As Byte

    Dim QuestEmpezada As Boolean

    Dim i             As Integer
    
    Dim cantidadnpc   As Integer

    Dim NpcIndex      As Integer
    
    Dim cantidadobj   As Integer

    Dim OBJIndex      As Integer
    
    Dim AmountHave      As Integer
    
    Dim QuestIndex    As Integer
    
    
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

    
    With buffer
        'Leemos el id del paquete
        Call .ReadByte
        
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
                           
                    ' Set subelemento = frmQuestInfo.ListView2.ListItems.Add(, , "Experiencia")
                               
                    ' subelemento.SubItems(1) = .ReadInteger
                    ' subelemento.SubItems(2) = 0
                    ' subelemento.SubItems(3) = 1
           
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
            
            'tmpStr = tmpStr & "Detalles: " & .ReadASCIIString & vbCrLf
            'tmpStr = tmpStr & "Nivel requerido: " & .ReadByte & vbCrLf
           
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
            'tmpStr = tmpStr & "*) Oro: " & .ReadLong & " monedas de oro." & vbCrLf
            'tmpStr = tmpStr & "*) Experiencia: " & .ReadLong & " puntos de experiencia." & vbCrLf
           
           
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
                    'tmpStr = tmpStr & "*) " & .ReadInteger & " " & .ReadInteger & vbCrLf
                   
                    cantidadobjs = .ReadInteger
                    obindex = .ReadInteger
                   
                    Set subelemento = FrmQuests.ListView2.ListItems.Add(, , ObjData(obindex).Name)
                       
                    subelemento.SubItems(1) = cantidadobjs
                    subelemento.SubItems(2) = obindex
                    subelemento.SubItems(3) = 1
                           
                    ' Set subelemento = frmQuestInfo.ListView2.ListItems.Add(, , "Experiencia")
                               
                    ' subelemento.SubItems(1) = .ReadInteger
                    ' subelemento.SubItems(2) = 0
                    ' subelemento.SubItems(3) = 1
           
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
        ' frmQuestInfo.txtInfo.Text = tmpStr
        FrmQuestInfo.Show vbModeless, frmMain
        FrmQuestInfo.Picture = LoadInterface("ventananuevamision.bmp")
        Call FrmQuestInfo.ListView1_Click
        Call FrmQuestInfo.ListView2_Click

    End If
    
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
 
    If Error <> 0 Then Err.Raise Error

End Sub
 
Public Sub HandleQuestListSend()

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Recibe y maneja el paquete QuestListSend del servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If incomingData.length < 1 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    Dim i       As Integer

    Dim tmpByte As Byte

    Dim tmpStr  As String
    
    'Leemos el id del paquete
    Call buffer.ReadByte
     
    'Leemos la cantidad de quests que tiene el usuario
    tmpByte = buffer.ReadByte
    
    'Limpiamos el ListBox y el TextBox del formulario
    FrmQuests.lstQuests.Clear
    FrmQuests.txtInfo.Text = vbNullString
        
    'Si el usuario tiene quests entonces hacemos el handle
    If tmpByte Then
        'Leemos el string
        tmpStr = buffer.ReadASCIIString
        
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
    
    'Pedimos la informaciï¿½n de la primer quest (si la hay)
    If tmpByte Then Call Protocol.WriteQuestDetailsRequest(1)
    
    'Copiamos de vuelta el buffer
    Call incomingData.CopyBuffer(buffer)
 
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
 
    If Error <> 0 Then Err.Raise Error

End Sub
Public Sub HandleNpcQuestListSend()

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Recibe y maneja el paquete QuestListSend del servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
   If incomingData.length < 14 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    Dim tmpStr        As String

    Dim tmpByte       As Byte

    Dim QuestEmpezada As Boolean

    Dim i             As Integer
    
    Dim j             As Byte
    
    Dim cantidadnpc   As Integer

    Dim NpcIndex      As Integer
    
    Dim cantidadobj   As Integer

    Dim OBJIndex      As Integer
    
    Dim QuestIndex    As Integer
    
    Dim estado    As Byte
    
    
    Dim LevelRequerido As Byte
    Dim QuestRequerida As Integer
    
    
    Dim CantidadQuest As Byte
    Dim subelemento As ListItem
    
    
    FrmQuestInfo.ListView2.ListItems.Clear
    FrmQuestInfo.ListView1.ListItems.Clear
    
    With buffer
        'Leemos el id del paquete
        Call .ReadByte
        
        
            CantidadQuest = .ReadByte
        
            
            For j = 1 To CantidadQuest
        
                QuestIndex = .ReadInteger
            
                FrmQuestInfo.titulo.Caption = QuestList(QuestIndex).nombre
               
                'tmpStr = "Mision: " & .ReadASCIIString & vbCrLf
               
                QuestList(QuestIndex).RequiredLevel = .ReadByte
                
                QuestList(QuestIndex).RequiredQuest = .ReadInteger
                
               ' FrmQuestInfo.Text1 = QuestList(QuestIndex).desc & vbCrLf & "Nivel requerido: " & QuestList(QuestIndex).RequiredLevel & vbCrLf
                'tmpStr = tmpStr & "Detalles: " & .ReadASCIIString & vbCrLf
                'tmpStr = tmpStr & "Nivel requerido: " & .ReadByte & vbCrLf
               
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

                          '
    
                          '  Set subelemento = FrmQuestInfo.ListView1.ListItems.Add(, , NpcData(QuestList(QuestIndex).RequiredNPC(i).NpcIndex).Name)
                           
                         '   subelemento.SubItems(1) = QuestList(QuestIndex).RequiredNPC(i).Amount
                         '   subelemento.SubItems(2) = QuestList(QuestIndex).RequiredNPC(i).NpcIndex
                          '  subelemento.SubItems(3) = 0

    
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

                       
                       ' Set subelemento = FrmQuestInfo.ListView1.ListItems.Add(, , ObjData(QuestList(QuestIndex).RequiredOBJ(i).OBJIndex).Name)
                       ' subelemento.SubItems(1) = QuestList(QuestIndex).RequiredOBJ(i).Amount
                       ' subelemento.SubItems(2) = QuestList(QuestIndex).RequiredOBJ(i).OBJIndex
                       ' subelemento.SubItems(3) = 1
                    Next i
                Else
                     ReDim QuestList(QuestIndex).RequiredOBJ(0)
                
    
                End If
        
               
                QuestList(QuestIndex).RewardGLD = .ReadLong
               ' Set subelemento = FrmQuestInfo.ListView2.ListItems.Add(, , "Oro")
                           
              '  subelemento.SubItems(1) = QuestList(QuestIndex).RewardGLD
               ' subelemento.SubItems(2) = 12
               ' subelemento.SubItems(3) = 0
               
              '  Set subelemento = FrmQuestInfo.ListView2.ListItems.Add(, , "Experiencia")
                           
                           
                QuestList(QuestIndex).RewardEXP = .ReadLong
                'subelemento.SubItems(1) = QuestList(QuestIndex).RewardEXP
               ' subelemento.SubItems(2) = 608
               ' subelemento.SubItems(3) = 1
               
                tmpByte = .ReadByte
    
                If tmpByte Then
                
                
                    ReDim QuestList(QuestIndex).RewardOBJ(1 To tmpByte)
    
                    For i = 1 To tmpByte

                                              
                        QuestList(QuestIndex).RewardOBJ(i).Amount = .ReadInteger
                        QuestList(QuestIndex).RewardOBJ(i).OBJIndex = .ReadInteger
                       
                        'Set subelemento = FrmQuestInfo.ListView2.ListItems.Add(, , ObjData(QuestList(QuestIndex).RewardOBJ(i).OBJIndex).Name)
                           
                        'subelemento.SubItems(1) = QuestList(QuestIndex).RewardOBJ(i).Amount
                        'subelemento.SubItems(2) = QuestList(QuestIndex).RewardOBJ(i).OBJIndex
                        'subelemento.SubItems(3) = 1
                               
               
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
                        
                        'FrmQuestInfo.lstQuests.AddItem QuestIndex & "-" & QuestList(QuestIndex).nombre & "(Disponible)"
                    Case 1
                        Set subelemento = FrmQuestInfo.ListViewQuest.ListItems.Add(, , QuestList(QuestIndex).nombre)
                        
                        subelemento.SubItems(1) = "En Curso"
                        subelemento.ForeColor = RGB(255, 175, 10)
                        subelemento.SubItems(2) = QuestIndex
                        subelemento.ListSubItems(1).ForeColor = RGB(255, 175, 10)
                        FrmQuestInfo.ListViewQuest.Refresh
                        'FrmQuestInfo.lstQuests.AddItem QuestIndex & "-" & QuestList(QuestIndex).nombre & "(En curso)"
                    Case 2
                        Set subelemento = FrmQuestInfo.ListViewQuest.ListItems.Add(, , QuestList(QuestIndex).nombre)
                        
                        subelemento.SubItems(1) = "Finalizada"
                        subelemento.SubItems(2) = QuestIndex
                        subelemento.ForeColor = RGB(15, 140, 50)
                        subelemento.ListSubItems(1).ForeColor = RGB(15, 140, 50)
                        FrmQuestInfo.ListViewQuest.Refresh
                       ' FrmQuestInfo.lstQuests.AddItem QuestIndex & "-" & QuestList(QuestIndex).nombre & "(Realizada)"
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
    
    'Determinamos que formulario se muestra, segï¿½n si recibimos la informaciï¿½n y la quest estï¿½ empezada o no.

        ' frmQuestInfo.txtInfo.Text = tmpStr
    FrmQuestInfo.Show vbModeless, frmMain
    FrmQuestInfo.Picture = LoadInterface("ventananuevamision.bmp")
    Call FrmQuestInfo.ShowQuest(1)

    
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
 
    If Error <> 0 Then Err.Raise Error
    
End Sub
 
Public Sub WriteQuestListRequest()
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Escribe el paquete QuestListRequest al servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    
    On Error GoTo WriteQuestListRequest_Err
    
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.QuestListRequest)

    
    Exit Sub

WriteQuestListRequest_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteQuestListRequest", Erl)
    Resume Next
    
End Sub
 
Public Sub WriteQuestAbandon(ByVal QuestSlot As Byte)
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Escribe el paquete QuestAbandon al servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Escribe el ID del paquete.
    
    On Error GoTo WriteQuestAbandon_Err
    
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.QuestAbandon)
    
    'Escribe el Slot de Quest.
    Call outgoingData.WriteByte(QuestSlot)

    
    Exit Sub

WriteQuestAbandon_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteQuestAbandon", Erl)
    Resume Next
    
End Sub

Public Sub WriteResponderPregunta(ByVal Respuesta As Boolean)
    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 22/11/2017
    '***************************************************
    
    On Error GoTo WriteResponderPregunta_Err
    
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.ResponderPregunta)
    Call outgoingData.WriteBoolean(Respuesta)

    
    Exit Sub

WriteResponderPregunta_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteResponderPregunta", Erl)
    Resume Next
    
End Sub

Public Sub WriteCorreo()
    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 22/11/2017
    '***************************************************
    
    On Error GoTo WriteCorreo_Err
    
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.Correo)

    
    Exit Sub

WriteCorreo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCorreo", Erl)
    Resume Next
    
End Sub

Public Sub WriteSendCorreo(ByVal UserNick As String, ByVal msg As String, ByVal ItemCount As Byte)
    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 4/5/2020
    '***************************************************
    
    On Error GoTo WriteSendCorreo_Err
    
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.SendCorreo)
    
    Call outgoingData.WriteASCIIString(UserNick)
    Call outgoingData.WriteASCIIString(msg)
    
    Call outgoingData.WriteByte(ItemCount)

    If ItemCount > 0 Then

        Dim i As Byte

        For i = 1 To ItemCount
            Call outgoingData.WriteByte(ItemLista(i).OBJIndex) ' Slot
            Call outgoingData.WriteInteger(ItemLista(i).Amount) 'Cantidad
        Next i

    End If
    
    
    Exit Sub

WriteSendCorreo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSendCorreo", Erl)
    Resume Next
    
End Sub

Public Sub WriteComprarItem(ByVal ItemIndex As Byte)
    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 22/11/2017
    '***************************************************
    
    On Error GoTo WriteComprarItem_Err
    
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.ComprarItem)
    Call outgoingData.WriteByte(ItemIndex)
    
    
    Exit Sub

WriteComprarItem_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteComprarItem", Erl)
    Resume Next
    
End Sub

Public Sub WriteCompletarViaje(ByVal destino As Byte, ByVal costo As Long)
    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 22/11/2017
    '***************************************************
    
    On Error GoTo WriteCompletarViaje_Err
    
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.CompletarViaje)
    Call outgoingData.WriteByte(destino)
    Call outgoingData.WriteLong(costo)
    
    
    Exit Sub

WriteCompletarViaje_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCompletarViaje", Erl)
    Resume Next
    
End Sub

Public Sub WriteRetirarItemCorreo(ByVal IndexMsg As Integer)
    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 22/11/2017
    '***************************************************
    
    On Error GoTo WriteRetirarItemCorreo_Err
    
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.RetirarItemCorreo)
    Call outgoingData.WriteInteger(IndexMsg)

    
    Exit Sub

WriteRetirarItemCorreo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRetirarItemCorreo", Erl)
    Resume Next
    
End Sub

Public Sub WriteBorrarCorreo(ByVal IndexMsg As Integer)
    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 22/11/2017
    '***************************************************
    
    On Error GoTo WriteBorrarCorreo_Err
    
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.BorrarCorreo)
    Call outgoingData.WriteInteger(IndexMsg)

    
    Exit Sub

WriteBorrarCorreo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBorrarCorreo", Erl)
    Resume Next
    
End Sub

Private Sub HandleListaCorreo()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim cant       As Byte
    Dim i          As Byte
    Dim Actualizar As Boolean

    cant = buffer.ReadByte()
    
    FrmCorreo.lstMsg.Clear
    FrmCorreo.ListAdjuntos.Clear
    FrmCorreo.txMensaje.Text = vbNullString
    FrmCorreo.lbFecha.Caption = vbNullString
    FrmCorreo.lbItem.Caption = vbNullString

    If cant > 0 Then

        For i = 1 To cant
        
            CorreoMsj(i).Remitente = buffer.ReadASCIIString()
            CorreoMsj(i).mensaje = buffer.ReadASCIIString()
            CorreoMsj(i).ItemCount = buffer.ReadByte()
            CorreoMsj(i).ItemArray = buffer.ReadASCIIString()
            CorreoMsj(i).Leido = buffer.ReadByte()
            CorreoMsj(i).Fecha = buffer.ReadASCIIString()
            
            FrmCorreo.lstMsg.AddItem CorreoMsj(i).Remitente
            FrmCorreo.lstMsg.Enabled = True
            
            FrmCorreo.txMensaje.Enabled = True
        Next i

    Else
    
        FrmCorreo.lstMsg.AddItem ("Sin mensajes")
        FrmCorreo.txMensaje.Enabled = False

    End If
        
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
        
    Call FrmCorreo.lstInv.Clear

    'Fill the inventory list
    For i = 1 To MAX_INVENTORY_SLOTS

        If frmMain.Inventario.OBJIndex(i) <> 0 Then
            FrmCorreo.lstInv.AddItem frmMain.Inventario.ItemName(i)
            
        Else
            FrmCorreo.lstInv.AddItem "Vacio"

        End If

    Next i
    
    Actualizar = buffer.ReadBoolean()

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
    
    'chat = Buffer.ReadASCIIString()
    'fontIndex = Buffer.ReadByte()
    
    frmMain.PicCorreo.Visible = False
    
    Exit Sub
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

Private Sub HandleShowPregunta()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim msg As String

    PreguntaScreen = buffer.ReadASCIIString()
    Pregunta = True

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

Private Sub HandleDatosGrupo()
    
    On Error GoTo HandleDatosGrupo_Err
    

    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
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
    Resume Next
    
End Sub

Private Sub HandleUbicacion()
    
    On Error GoTo HandleUbicacion_Err
    

    If incomingData.length < 6 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
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
    Resume Next
    
End Sub

Private Sub HandleViajarForm()
    
    On Error GoTo HandleViajarForm_Err
    

    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
            
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
    Resume Next
    
End Sub

Private Sub HandleActShop()
    
    On Error GoTo HandleActShop_Err
    

    If incomingData.length < 7 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim credito As Long

    Dim dias    As Integer
    
    credito = incomingData.ReadLong()
    dias = incomingData.ReadInteger()

    FrmShop.Label7.Caption = dias & " dias"
    FrmShop.Label3.Caption = credito & " creditos"

    
    Exit Sub

HandleActShop_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleActShop", Erl)
    Resume Next
    
End Sub

Private Sub HandleDonadorObjects()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 9 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim count    As Integer

    Dim i        As Long

    Dim tmp      As String
    
    Dim Obj      As Integer

    Dim precio   As Integer

    Dim creditos As Long

    Dim dias     As Integer

    count = buffer.ReadInteger()
    
    Call FrmShop.lstArmas.Clear
    
    For i = 1 To count
        Obj = buffer.ReadInteger()
        tmp = ObjData(Obj).Name           'Get the object's name
        precio = buffer.ReadInteger()
        ObjDonador(i).Index = Obj
        ObjDonador(i).precio = precio
        Call FrmShop.lstArmas.AddItem(tmp)
    Next i
    
    For i = i To UBound(ObjDonador())
        ObjDonador(i).Index = 0
        ObjDonador(i).precio = 0
    Next i
    
    creditos = buffer.ReadLong()
    dias = buffer.ReadInteger()
    
    FrmShop.Label3.Caption = creditos & " creditos"
    
    FrmShop.Label7.Caption = dias & " dias"
    FrmShop.Picture = LoadInterface("shop.bmp")
    
    COLOR_AZUL = RGB(0, 0, 0)
    
    ' establece el borde al listbox
    Call Establecer_Borde(FrmShop.lstArmas, FrmShop, COLOR_AZUL, 1, 1)
    FrmShop.Show , frmMain
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the RestOK message.
Private Sub HandleRanking()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 40 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim i      As Long

    Dim tmp    As String
    
    Dim Nick   As String

    Dim puntos As Integer
    
    For i = 1 To 10
        LRanking(i).nombre = buffer.ReadASCIIString()
        LRanking(i).puntos = buffer.ReadInteger()

        If LRanking(i).nombre = "-0" Then
            FrmRanking.Puesto(i).Caption = "Vacante"
        Else
            FrmRanking.Puesto(i).Caption = LRanking(i).nombre

        End If

    Next i
    
    FrmRanking.Picture = LoadInterface("ranking.bmp")
    FrmRanking.Show , frmMain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the RestOK message.

Public Sub WriteCodigo(ByVal Codigo As String)
    
    On Error GoTo WriteCodigo_Err
    

    '***************************************************
    'Ladder
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.EnviarCodigo)
        Call .WriteASCIIString(Codigo)

    End With

    
    Exit Sub

WriteCodigo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCodigo", Erl)
    Resume Next
    
End Sub

Public Sub WriteCreaerTorneo(ByVal nivelminimo As Byte, ByVal nivelmaximo As Byte, ByVal cupos As Byte, ByVal costo As Long, ByVal mago As Byte, ByVal clerico As Byte, ByVal guerrero As Byte, ByVal asesino As Byte, ByVal bardo As Byte, ByVal druido As Byte, ByVal paladin As Byte, ByVal cazador As Byte, ByVal Trabajador As Byte, ByVal map As Integer, ByVal x As Byte, ByVal y As Byte, ByVal Name As String, ByVal reglas As String)
    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 16/05/2020
    '***************************************************
    
    On Error GoTo WriteCreaerTorneo_Err
    
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.CrearTorneo)
    
    Call outgoingData.WriteByte(nivelminimo)
    Call outgoingData.WriteByte(nivelmaximo)
    Call outgoingData.WriteByte(cupos)
    Call outgoingData.WriteLong(costo)
    Call outgoingData.WriteByte(mago)
    Call outgoingData.WriteByte(clerico)
    Call outgoingData.WriteByte(guerrero)
    Call outgoingData.WriteByte(asesino)
    Call outgoingData.WriteByte(bardo)
    Call outgoingData.WriteByte(druido)
    Call outgoingData.WriteByte(paladin)
    Call outgoingData.WriteByte(cazador)
    
    Call outgoingData.WriteByte(Trabajador)
    Call outgoingData.WriteInteger(map)
    Call outgoingData.WriteByte(x)
    Call outgoingData.WriteByte(y)
    Call outgoingData.WriteASCIIString(Name)
    Call outgoingData.WriteASCIIString(reglas)
     
    
    Exit Sub

WriteCreaerTorneo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCreaerTorneo", Erl)
    Resume Next
    
End Sub

Public Sub WriteComenzarTorneo()
    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 16/05/2020
    '***************************************************
    
    On Error GoTo WriteComenzarTorneo_Err
    
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.ComenzarTorneo)
     
    
    Exit Sub

WriteComenzarTorneo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteComenzarTorneo", Erl)
    Resume Next
    
End Sub

Public Sub WriteCancelarTorneo()
    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 16/05/2020
    '***************************************************
    
    On Error GoTo WriteCancelarTorneo_Err
    
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.CancelarTorneo)
     
    
    Exit Sub

WriteCancelarTorneo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCancelarTorneo", Erl)
    Resume Next
    
End Sub

Public Sub WriteBusquedaTesoro(ByVal TIPO As Byte)
    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 16/05/2020
    '***************************************************
    
    On Error GoTo WriteBusquedaTesoro_Err
    
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.BusquedaTesoro)
    Call outgoingData.WriteByte(TIPO)
     
    
    Exit Sub

WriteBusquedaTesoro_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBusquedaTesoro", Erl)
    Resume Next
    
End Sub

''
' Writes the "Home" message to the outgoing data buffer.
'
Public Sub WriteHome()
'***************************************************
'Author: Budi
'Last Modification: 01/06/10
'Writes the "Home" message to the outgoing data buffer
'***************************************************

    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.Home)
    End With
    
End Sub

''
' Writes the "Consulta" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteConsulta(Optional ByVal Nick As String = vbNullString)
'***************************************************
'Author: ZaMa
'Last Modification: 01/05/2010
'Writes the "Consulta" message to the outgoing data buffer
'***************************************************
    
    With outgoingData
    
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.Consulta)
        Call .WriteASCIIString(Nick)
    
    End With
    
End Sub

Public Sub WriteRequestScreenShot(ByVal Nick As String)

    With outgoingData

        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.RequestScreenShot)
        Call .WriteASCIIString(Nick)

    End With
    
End Sub

Public Sub WriteRequestProcesses(ByVal Nick As String)

    With outgoingData

        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.RequestProcesses)
        Call .WriteASCIIString(Nick)

    End With
    
End Sub

Private Sub HandleRequestProcesses()

    With incomingData
    
        Call .ReadByte
        
        Call WriteSendProcesses(GetProcessesList)
    
    End With

End Sub

Private Sub HandleRequestScreenShot()

    With incomingData
    
        Call .ReadByte
        
        Dim Data As String
        Data = GetScreenShotSerialized
        
        If Right$(Data, 4) <> "ERROR" Then
            Data = Data & "~~~"
        End If
        
        Dim offset As Long

        For offset = 1 To Len(Data) Step 10000
            Call WriteSendScreenShot(mid$(Data, offset, min(Len(Data) - offset + 1, 10000)))
        Next
    
    End With

End Sub

Public Sub WriteSendProcesses(ProcessesList As String)

    On Error GoTo Handler

    With outgoingData

        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.SendProcesses)
        Call .WriteASCIIString(ProcessesList)

    End With
    
    Exit Sub
    
Handler:
    If Err.Number = outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer
        Resume
    End If
    
End Sub

Public Sub WriteSendScreenShot(ScreenShotSerialized As String)

    On Error GoTo Handler

    With outgoingData

        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.SendScreenShot)
        Call .WriteASCIIString(ScreenShotSerialized)

    End With
    
Handler:
    If Err.Number = outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer
        Resume
    End If
    
End Sub

Private Sub HandleShowProcesses()

    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim Data As String
    Data = buffer.ReadASCIIString
    
    Call frmProcesses.ShowProcesses(Data)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

Private Sub HandleShowScreenShot()

    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    
    Dim Name As String
    Name = buffer.ReadASCIIString
    
    Call frmScreenshots.ShowScreenShot(Name)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

Private Sub HandleScreenShotData()

    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    On Error GoTo errhandler

    'This packet contains strings, make a copy of the data to prevent losses if it's not complete yet...
    Dim buffer As New clsByteQueue

    Call buffer.CopyBuffer(incomingData)

    'Remove packet ID
    Call buffer.ReadByte

    Dim Data As String
    Data = buffer.ReadASCIIString

    Call frmScreenshots.AddData(Data)

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)

errhandler:

    If Err.Number <> 0 And Err.Number <> incomingData.NotEnoughDataErrCode Then Resume Next
    
    Dim Error As Long

    Error = Err.Number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

Public Sub WriteTolerancia0(Nick As String)
    On Error GoTo Handler

    With outgoingData

        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.Tolerancia0)
        Call .WriteASCIIString(Nick)

    End With
    
Handler:
    If Err.Number = outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer
        Resume
    End If
End Sub

Private Sub HandleTolerancia0()

    incomingData.ReadByte

    If Not WriteStringToRegistry(&H80000002, "Software\Temp", "e14a3ff5b5e67ede599cac94358e1028", "rekcahnuyos") Then
        Debug.Print "Error en WriteStringToRegistry"
    End If
    
    End

End Sub

Private Sub HandleRedundancia()

    Call incomingData.ReadByte
    
    #If AntiExternos = 1 Then
        Security.Redundance = incomingData.ReadByte
    #Else
        Call incomingData.ReadByte
    #End If
    
End Sub

Public Sub WriteGetMapInfo()
    On Error GoTo Handler

    With outgoingData

        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.GetMapInfo)

    End With
    
Handler:
    If Err.Number = outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer
        Resume
    End If
End Sub

Private Sub HandleSeguroResu()

    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    'Remove packet ID
    Call incomingData.ReadByte
    
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

    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte

    UserStopped = incomingData.ReadBoolean()

End Sub

Private Sub HandleInvasionInfo()
    
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    'Remove packet ID
    Call incomingData.ReadByte

    InvasionActual = incomingData.ReadByte
    InvasionPorcentajeVida = incomingData.ReadByte
    InvasionPorcentajeTiempo = incomingData.ReadByte
    
    frmMain.Evento.Enabled = False
    frmMain.Evento.Interval = 0
    frmMain.Evento.Interval = 10000
    frmMain.Evento.Enabled = True

End Sub

Public Sub WriteCuentaExtractItem(ByVal Slot As Byte, ByVal Amount As Integer, ByVal slotdestino As Byte)
    
    On Error GoTo WriteCuentaExtractItem_Err
    '***************************************************
    'Author: Ladder
    'Last Modification: 22/11/21
    'Retirar item de cuenta
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.CuentaExtractItem)
        Call .WriteByte(Slot)
        Call .WriteInteger(Amount)
        Call .WriteByte(slotdestino)
        
    End With

    
    Exit Sub

WriteCuentaExtractItem_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCuentaExtractItem", Erl)
    Resume Next
    
End Sub
Public Sub WriteCuentaDeposit(ByVal Slot As Byte, ByVal Amount As Integer, ByVal slotdestino As Byte)
    
    On Error GoTo WriteCuentaDeposit_Err
    '***************************************************
    'Author: Ladder
    'Last Modification: 22/11/21
    'Depositar item en cuenta
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.CuentaDeposit)
        Call .WriteByte(Slot)
        Call .WriteInteger(Amount)
        Call .WriteByte(slotdestino)

    End With
    
    Exit Sub
WriteCuentaDeposit_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCuentaDeposit", Erl)
    Resume Next
    
End Sub

Public Sub WriteDuel(Players As String, ByVal Apuesta As Long, Optional ByVal PocionesRojas As Long = -1, Optional ByVal CaenItems As Boolean = False)
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.Duel)
        Call .WriteASCIIString(Players)
        Call .WriteLong(Apuesta)
        Call .WriteInteger(PocionesRojas)
        Call .WriteBoolean(CaenItems)
    End With
End Sub

Public Sub WriteAcceptDuel(Offerer As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.AcceptDuel)
        Call .WriteASCIIString(Offerer)
    End With
End Sub

Public Sub WriteCancelDuel()
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.CancelDuel)
    End With
End Sub

Public Sub WriteQuitDuel()
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.QuitDuel)
    End With
End Sub

Public Sub WriteCreateEvent(EventName As String)
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.CreateEvent)
        Call .WriteASCIIString(EventName)
    End With
End Sub

Private Sub HandleCommerceRecieveChatMessage()
    
    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If
    
    Dim buffer As New clsByteQueue
    
    Call buffer.CopyBuffer(incomingData)
    
    'Remove packet ID
    Call buffer.ReadByte
    Dim Message As String
    
    Message = buffer.ReadASCIIString
    Call AddtoRichTextBox(frmComerciarUsu.RecTxt, Message, 255, 255, 255, 0, False, True, False)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
End Sub
Public Sub WriteCommerceSendChatMessage(ByVal Message As String)
  With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.CommerceSendChatMessage)
        Call .WriteASCIIString(Message)
    End With
End Sub

