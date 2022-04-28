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

    Connected
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
    CharacterCreate         ' CC
    CharacterRemove         ' BP
    CharacterMove           ' MP, +, * and _ '
    UserIndexInServer       ' IU
    UserCharIndexInServer   ' IP
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
    'SendNight              ' NOC
    Pong
    UpdateTagAndStatus
    FYA
    CerrarleCliente
    Contadores
    ShowPapiro
    
    'GM messages
    SpawnListt               ' SPL
    ShowSOSForm             ' MSOS
    ShowMOTDEditionForm     ' ZMOTD
    ShowGMPanelForm         ' ABPANEL
    UserNameList            ' LISTUSU
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
    SastreObj
    ShowSastreForm ' 126
    VelocidadToggle
    MacroTrabajoToggle
    BindKeys
    ShowFrmLogear
    ShowFrmMapa
    InmovilizadoOK
    BarFx
    LocaleMsg
    ShowPregunta
    DatosGrupo
    ubicacion
    ArmaMov
    EscudoMov
    ViajarForm
    NadarToggle
    ShowFundarClanForm
    CharUpdateHP
    CharUpdateMAN
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
    GuardNotice
    AnswerReset
    ObjQuestListSend
    UpdateBankGld
    PelearConPezEspecial
    Privilegios
    ShopInit
    UpdateShopClienteCredits
    SensuiRetrasado
    [PacketCount]
End Enum

Public Enum ClientPacketID
    LoginExistingChar       'OLOGIN
    LoginNewChar            'NLOGIN
    Talk                    ';
    Yell                    '-
    Whisper                 '\
    Walk                    'M
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
    '--------------------
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
    GuildMessage            '/CMSG
    CentinelReport          '/CENTINELA
    GuildOnline             '/ONLINECLAN
    CouncilMessage          '/BMSG
    RoleMasterRequest       '/ROL
    ChangeDescription       '/DESC
    GuildVote               '/VOTO
    punishments             '/PENAS
    Gamble                  '/APOSTAR
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
    CouncilKick             '/KICKCONSE
    SetTrigger              '/TRIGGER
    AskTrigger              '/TRIGGER with no args
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
    AlterName               '/ANAME
    DoBackUp                '/DOBACKUP
    ShowGuildMessages       '/SHOWCMSG
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
    ReloadNPCs              '/RELOADNPCS
    ReloadServerIni         '/RELOADSINI
    ReloadSpells            '/RELOADHECHIZOS
    ReloadObjects           '/RELOADOBJ
    ChatColor               '/CHATCOLOR
    Ignored                 '/IGNORADO
    CheckSlot               '/SLOT
    
    'Nuevas Ladder
    SetSpeed                '/SPEED
    GlobalMessage           '/CONSOLA
    GlobalOnOff
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
    Genio
    Casarse
    CraftAlquimista
    FlagTrabajar
    CraftSastre
    MensajeUser
    TraerBoveda
    CompletarAccion
    InvitarGrupo
    ResponderPregunta
    RequestGrupo
    AbandonarGrupo
    HecharDeGrupo
    MacroPossent
    SubastaInfo
    BanCuenta
    UnbanCuenta
    CerrarCliente
    EventoInfo
    CrearEvento
    BanTemporal
    CancelarExit
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
    Quest
    QuestAccept
    QuestListRequest
    QuestDetailsRequest
    QuestAbandon
    SeguroClan
    Home                    '/HOGAR
    Consulta                '/CONSULTA
    GetMapInfo              '/MAPINFO
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
    ResetChar              '/RESET NICK
    ResetearPersonaje
    DeleteItem
    FinalizarPescaEspecial
    RomperCania
    UseItemU
    RepeatMacro
    BuyShopItem
    [PacketCount]
End Enum

Private Reader As Network.Reader

''
' Handles incoming data.

Public Function HandleIncomingData(ByVal Message As Network.Reader) As Boolean
On Error GoTo HandleIncomingData_Err

    Set Reader = Message
    
    Dim PacketId As Long
    PacketId = Reader.ReadInt16
    
    #If DEBUGGING Then
        'Debug.Print PacketId
    #End If
    
    Select Case PacketID
        Case ServerPacketID.Connected
            Call HandleConnected
        Case ServerPacketID.logged
            Call HandleLogged
        Case ServerPacketID.RemoveDialogs
            Call HandleRemoveDialogs
        Case ServerPacketID.RemoveCharDialog
            Call HandleRemoveCharDialog
        Case ServerPacketID.NavigateToggle
            Call HandleNavigateToggle
        Case ServerPacketID.EquiteToggle
            Call HandleEquiteToggle
        Case ServerPacketID.Disconnect
            Call HandleDisconnect
        Case ServerPacketID.CommerceEnd
            Call HandleCommerceEnd
        Case ServerPacketID.BankEnd
            Call HandleBankEnd
        Case ServerPacketID.CommerceInit
            Call HandleCommerceInit
        Case ServerPacketID.BankInit
            Call HandleBankInit
        Case ServerPacketID.UserCommerceInit
            Call HandleUserCommerceInit
        Case ServerPacketID.UserCommerceEnd
            Call HandleUserCommerceEnd
        Case ServerPacketID.ShowBlacksmithForm
            Call HandleShowBlacksmithForm
        Case ServerPacketID.ShowCarpenterForm
            Call HandleShowCarpenterForm
        Case ServerPacketID.NPCKillUser
            Call HandleNPCKillUser
        Case ServerPacketID.BlockedWithShieldUser
            Call HandleBlockedWithShieldUser
        Case ServerPacketID.BlockedWithShieldOther
            Call HandleBlockedWithShieldOther
        Case ServerPacketID.CharSwing
            Call HandleCharSwing
        Case ServerPacketID.SafeModeOn
            Call HandleSafeModeOn
        Case ServerPacketID.SafeModeOff
            Call HandleSafeModeOff
        Case ServerPacketID.PartySafeOn
            Call HandlePartySafeOn
        Case ServerPacketID.PartySafeOff
            Call HandlePartySafeOff
        Case ServerPacketID.CantUseWhileMeditating
            Call HandleCantUseWhileMeditating
        Case ServerPacketID.UpdateSta
            Call HandleUpdateSta
        Case ServerPacketID.UpdateMana
            Call HandleUpdateMana
        Case ServerPacketID.UpdateHP
            Call HandleUpdateHP
        Case ServerPacketID.UpdateGold
            Call HandleUpdateGold
        Case ServerPacketID.UpdateExp
            Call HandleUpdateExp
        Case ServerPacketID.ChangeMap
            Call HandleChangeMap
        Case ServerPacketID.PosUpdate
            Call HandlePosUpdate
        Case ServerPacketID.NPCHitUser
            Call HandleNPCHitUser
        Case ServerPacketID.UserHitNPC
            Call HandleUserHitNPC
        Case ServerPacketID.UserAttackedSwing
            Call HandleUserAttackedSwing
        Case ServerPacketID.UserHittedByUser
            Call HandleUserHittedByUser
        Case ServerPacketID.UserHittedUser
            Call HandleUserHittedUser
        Case ServerPacketID.ChatOverHead
            Call HandleChatOverHead
        Case ServerPacketID.ConsoleMsg
            Call HandleConsoleMessage
        Case ServerPacketID.GuildChat
            Call HandleGuildChat
        Case ServerPacketID.ShowMessageBox
            Call HandleShowMessageBox
        Case ServerPacketID.MostrarCuenta
            Call HandleMostrarCuenta
        Case ServerPacketID.CharacterCreate
            Call HandleCharacterCreate
        Case ServerPacketID.CharacterRemove
            Call HandleCharacterRemove
        Case ServerPacketID.CharacterMove
            Call HandleCharacterMove
        Case ServerPacketID.UserIndexInServer
            Call HandleUserIndexInServer
        Case ServerPacketID.UserCharIndexInServer
            Call HandleUserCharIndexInServer
        Case ServerPacketID.ForceCharMove
            Call HandleForceCharMove
        Case ServerPacketID.CharacterChange
            Call HandleCharacterChange
        Case ServerPacketID.ObjectCreate
            Call HandleObjectCreate
        Case ServerPacketID.fxpiso
            Call HandleFxPiso
        Case ServerPacketID.ObjectDelete
            Call HandleObjectDelete
        Case ServerPacketID.BlockPosition
            Call HandleBlockPosition
        Case ServerPacketID.PlayMIDI
            Call HandlePlayMIDI
        Case ServerPacketID.PlayWave
            Call HandlePlayWave
        Case ServerPacketID.guildList
            Call HandleGuildList
        Case ServerPacketID.AreaChanged
            Call HandleAreaChanged
        Case ServerPacketID.PauseToggle
            Call HandlePauseToggle
        Case ServerPacketID.RainToggle
            Call HandleRainToggle
        Case ServerPacketID.CreateFX
            Call HandleCreateFX
        Case ServerPacketID.UpdateUserStats
            Call HandleUpdateUserStats
        Case ServerPacketID.WorkRequestTarget
            Call HandleWorkRequestTarget
        Case ServerPacketID.ChangeInventorySlot
            Call HandleChangeInventorySlot
        Case ServerPacketID.InventoryUnlockSlots
            Call HandleInventoryUnlockSlots
        Case ServerPacketID.ChangeBankSlot
            Call HandleChangeBankSlot
        Case ServerPacketID.ChangeSpellSlot
            Call HandleChangeSpellSlot
        Case ServerPacketID.Atributes
            Call HandleAtributes
        Case ServerPacketID.BlacksmithWeapons
            Call HandleBlacksmithWeapons
        Case ServerPacketID.BlacksmithArmors
            Call HandleBlacksmithArmors
        Case ServerPacketID.CarpenterObjects
            Call HandleCarpenterObjects
        Case ServerPacketID.RestOK
            Call HandleRestOK
        Case ServerPacketID.ErrorMsg
            Call HandleErrorMessage
        Case ServerPacketID.Blind
            Call HandleBlind
        Case ServerPacketID.Dumb
            Call HandleDumb
        Case ServerPacketID.ShowSignal
            Call HandleShowSignal
        Case ServerPacketID.ChangeNPCInventorySlot
            Call HandleChangeNPCInventorySlot
        Case ServerPacketID.UpdateHungerAndThirst
            Call HandleUpdateHungerAndThirst
        Case ServerPacketID.MiniStats
            Call HandleMiniStats
        Case ServerPacketID.LevelUp
            Call HandleLevelUp
        Case ServerPacketID.AddForumMsg
            Call HandleAddForumMessage
        Case ServerPacketID.ShowForumForm
            Call HandleShowForumForm
        Case ServerPacketID.SetInvisible
            Call HandleSetInvisible
        Case ServerPacketID.MeditateToggle
            Call HandleMeditateToggle
        Case ServerPacketID.BlindNoMore
            Call HandleBlindNoMore
        Case ServerPacketID.DumbNoMore
            Call HandleDumbNoMore
        Case ServerPacketID.SendSkills
            Call HandleSendSkills
        Case ServerPacketID.TrainerCreatureList
            Call HandleTrainerCreatureList
        Case ServerPacketID.guildNews
            Call HandleGuildNews
        Case ServerPacketID.OfferDetails
            Call HandleOfferDetails
        Case ServerPacketID.AlianceProposalsList
            Call HandleAlianceProposalsList
        Case ServerPacketID.PeaceProposalsList
            Call HandlePeaceProposalsList
        Case ServerPacketID.CharacterInfo
            Call HandleCharacterInfo
        Case ServerPacketID.GuildLeaderInfo
            Call HandleGuildLeaderInfo
        Case ServerPacketID.GuildDetails
            Call HandleGuildDetails
        Case ServerPacketID.ShowGuildFundationForm
            Call HandleShowGuildFundationForm
        Case ServerPacketID.ParalizeOK
            Call HandleParalizeOK
        Case ServerPacketID.ShowUserRequest
            Call HandleShowUserRequest
        Case ServerPacketID.ChangeUserTradeSlot
            Call HandleChangeUserTradeSlot
        'Case ServerPacketID.SendNight
        '    Call HandleSendNight
        Case ServerPacketID.Pong
            Call HandlePong
        Case ServerPacketID.UpdateTagAndStatus
            Call HandleUpdateTagAndStatus
        Case ServerPacketID.FYA
            Call HandleFYA
        Case ServerPacketID.CerrarleCliente
            Call HandleCerrarleCliente
        Case ServerPacketID.Contadores
            Call HandleContadores
        Case ServerPacketID.ShowPapiro
            Call HandleShowPapiro
        Case ServerPacketID.SpawnListt
            Call HandleSpawnList
        Case ServerPacketID.ShowSOSForm
            Call HandleShowSOSForm
        Case ServerPacketID.ShowMOTDEditionForm
            Call HandleShowMOTDEditionForm
        Case ServerPacketID.ShowGMPanelForm
            Call HandleShowGMPanelForm
        Case ServerPacketID.UserNameList
            Call HandleUserNameList
        Case ServerPacketID.UserOnline
            Call HandleUserOnline
        Case ServerPacketID.ParticleFX
            Call HandleParticleFX
        Case ServerPacketID.ParticleFXToFloor
            Call HandleParticleFXToFloor
        Case ServerPacketID.ParticleFXWithDestino
            Call HandleParticleFXWithDestino
        Case ServerPacketID.ParticleFXWithDestinoXY
            Call HandleParticleFXWithDestinoXY
        Case ServerPacketID.hora
            Call HandleHora
        Case ServerPacketID.Light
            Call HandleLight
        Case ServerPacketID.AuraToChar
            Call HandleAuraToChar
        Case ServerPacketID.SpeedToChar
            Call HandleSpeedToChar
        Case ServerPacketID.LightToFloor
            Call HandleLightToFloor
        Case ServerPacketID.NieveToggle
            Call HandleNieveToggle
        Case ServerPacketID.NieblaToggle
            Call HandleNieblaToggle
        Case ServerPacketID.Goliath
            Call HandleGoliath
        Case ServerPacketID.TextOverChar
            Call HandleTextOverChar
        Case ServerPacketID.TextOverTile
            Call HandleTextOverTile
        Case ServerPacketID.TextCharDrop
            Call HandleTextCharDrop
        Case ServerPacketID.FlashScreen
            Call HandleFlashScreen
        Case ServerPacketID.AlquimistaObj
            Call HandleAlquimiaObjects
        Case ServerPacketID.ShowAlquimiaForm
            Call HandleShowAlquimiaForm
        Case ServerPacketID.SastreObj
            Call HandleSastreObjects
        Case ServerPacketID.ShowSastreForm
            Call HandleShowSastreForm
        Case ServerPacketID.VelocidadToggle
            Call HandleVelocidadToggle
        Case ServerPacketID.MacroTrabajoToggle
            Call HandleMacroTrabajoToggle
        Case ServerPacketID.BindKeys
            Call HandleBindKeys
        Case ServerPacketID.ShowFrmLogear
            Call HandleShowFrmLogear
        Case ServerPacketID.ShowFrmMapa
            Call HandleShowFrmMapa
        Case ServerPacketID.InmovilizadoOK
            Call HandleInmovilizadoOK
        Case ServerPacketID.BarFx
            Call HandleBarFx
        Case ServerPacketID.LocaleMsg
            Call HandleLocaleMsg
        Case ServerPacketID.ShowPregunta
            Call HandleShowPregunta
        Case ServerPacketID.DatosGrupo
            Call HandleDatosGrupo
        Case ServerPacketID.ubicacion
            Call HandleUbicacion
        Case ServerPacketID.ArmaMov
            Call HandleArmaMov
        Case ServerPacketID.EscudoMov
            Call HandleEscudoMov
        Case ServerPacketID.ViajarForm
            Call HandleViajarForm
        Case ServerPacketID.NadarToggle
            Call HandleNadarToggle
        Case ServerPacketID.ShowFundarClanForm
            Call HandleShowFundarClanForm
        Case ServerPacketID.CharUpdateHP
            Call HandleCharUpdateHP
        Case ServerPacketID.CharUpdateMAN
            Call HandleCharUpdateMAN
        Case ServerPacketID.PosLLamadaDeClan
            Call HandlePosLLamadaDeClan
        Case ServerPacketID.QuestDetails
            Call HandleQuestDetails
        Case ServerPacketID.QuestListSend
            Call HandleQuestListSend
        Case ServerPacketID.NpcQuestListSend
            Call HandleNpcQuestListSend
        Case ServerPacketID.UpdateNPCSimbolo
            Call HandleUpdateNPCSimbolo
        Case ServerPacketID.ClanSeguro
            Call HandleClanSeguro
        Case ServerPacketID.Intervals
            Call HandleIntervals
        Case ServerPacketID.UpdateUserKey
            Call HandleUpdateUserKey
        Case ServerPacketID.UpdateRM
            Call HandleUpdateRM
        Case ServerPacketID.UpdateDM
            Call HandleUpdateDM
        Case ServerPacketID.SeguroResu
            Call HandleSeguroResu
        Case ServerPacketID.Stopped
            Call HandleStopped
        Case ServerPacketID.InvasionInfo
            Call HandleInvasionInfo
        Case ServerPacketID.CommerceRecieveChatMessage
            Call HandleCommerceRecieveChatMessage
        Case ServerPacketID.DoAnimation
            Call HandleDoAnimation
        Case ServerPacketID.OpenCrafting
            Call HandleOpenCrafting
        Case ServerPacketID.CraftingItem
            Call HandleCraftingItem
        Case ServerPacketID.CraftingCatalyst
            Call HandleCraftingCatalyst
        Case ServerPacketID.CraftingResult
            Call HandleCraftingResult
        Case ServerPacketID.ForceUpdate
            Call HandleForceUpdate
        Case ServerPacketID.AnswerReset
            Call HandleAnswerReset
        Case ServerPacketID.ObjQuestListSend
            Call HandleObjQuestListSend
        Case ServerPacketID.UpdateBankGld
            Call HandleUpdateBankGld
        Case ServerPacketID.PelearConPezEspecial
            Call HandlePelearConPezEspecial
        Case ServerPacketID.Privilegios
            Call HandlePrivilegios
        Case ServerPacketID.ShopInit
            Call HandleShopInit
        Case ServerPacketID.UpdateShopClienteCredits
            Call HandleUpdateShopClienteCredits
        Case ServerPacketID.SensuiRetrasado
            Call HandleSensuiRetrasado
        Case Else
            Err.Raise &HDEADBEEF, "Invalid Message"
    End Select
    
    If (Message.GetAvailable() > 0) Then
        Err.Raise &HDEADBEEF, "HandleIncomingData", "El paquete '" & PacketID & "' se encuentra en mal estado con '" & Message.GetAvailable() & "' bytes de mas"
    End If

    HandleIncomingData = True
    
HandleIncomingData_Err:
    
    Set Reader = Nothing

    If Err.Number <> 0 Then
        Call RegistrarError(Err.Number, Err.Description & ". PacketID: " & PacketID, "Protocol.HandleIncomingData", Erl)
        Call modNetwork.Disconnect
        
        HandleIncomingData = False
    End If

End Function

''
' Handles the Connected message.

Private Sub HandleConnected()

    frmMain.ShowFPS.Enabled = True

    Call Login
    
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
    frmMain.imgDeleteItem.Visible = True
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

    Call Dialogos.RemoveDialog(Reader.ReadInt16())
    
    Exit Sub

HandleRemoveCharDialog_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleRemoveCharDialog", Erl)
    
    
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
    
    
End Sub

Private Sub HandleNadarToggle()
    
    On Error GoTo HandleNadarToggle_Err

    UserNadando = Reader.ReadBool()
    UserNadandoTrajeCaucho = Reader.ReadBool()
    
    Exit Sub

HandleNadarToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleNadarToggle", Erl)
    
    
End Sub

Private Sub HandleEquiteToggle()
 
    On Error GoTo HandleEquiteToggle_Err
    
    UserMontado = Not UserMontado

    Exit Sub

HandleEquiteToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleEquiteToggle", Erl)
    
    
End Sub

Private Sub HandleVelocidadToggle()
    
    On Error GoTo HandleVelocidadToggle_Err

    If UserCharIndex = 0 Then Exit Sub
    
    charlist(UserCharIndex).Speeding = Reader.ReadReal32()
    
    Call MainTimer.SetInterval(TimersIndex.Walk, IntervaloCaminar / charlist(UserCharIndex).Speeding)
    
    Exit Sub

HandleVelocidadToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleVelocidadToggle", Erl)
    
    
End Sub

Private Sub HandleMacroTrabajoToggle()
    'Activa o Desactiva el macro de trabajo  06/07/2014 Ladder
    
    On Error GoTo HandleMacroTrabajoToggle_Err

    Dim activar As Boolean
    activar = Reader.ReadBool()

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
    
    
End Sub

''
' Handles the Disconnect message.

Public Sub HandleDisconnect()
    
    On Error GoTo HandleDisconnect_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim i As Long
    
    If (Not Reader Is Nothing) Then
    FullLogout = Reader.ReadBool
    End If

    Mod_Declaraciones.Connected = False
    
    Call ResetearUserMacro

    'Close connection
    Call modNetwork.Disconnect
    
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
    frmMain.Fuerzalbl.Visible = True
    frmMain.AgilidadLbl.Visible = True
    frmMain.oxigenolbl.Visible = True
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
    initPacketControl
    'Stop audio
    If Sonido Then
        Sound.Sound_Stop_All
        Sound.Ambient_Stop

    End If

    Call CleanDialogs
    
    'frmMain.IsPlaying = PlayLoop.plNone
    
    'Show connection form
    UserMap = 1
    
    EntradaY = 1
    EntradaX = 1
    Call EraseChar(UserCharIndex, True)
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
    UserNadandoTrajeCaucho = False
    bRain = False
    AlphaNiebla = 30
    frmMain.TimerNiebla.Enabled = False
    bNiebla = False
    bNieve = False
    bFogata = False
    SkillPoints = 0
    UserEstado = 0
    
    InviCounter = 0
    DrogaCounter = 0
    EscribeRetrasadoSensui = False
    frmMain.timerRetrasadoSensui = False
     
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
    
    If Not FullLogout Then
        'Si no es un deslogueo completo, envío nuevamente la lista de Pjs.
        Call connectToLoginServer
    End If
    
    Exit Sub
HandleDisconnect_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleDisconnect", Erl)
    
    
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

    NpcName = Reader.ReadString8()

    'Fill our inventory list
    For i = 1 To MAX_INVENTORY_SLOTS

        With frmMain.Inventario
            Call frmComerciar.InvComUsu.SetItem(i, .ObjIndex(i), .Amount(i), .Equipped(i), .GrhIndex(i), .ObjType(i), .MaxHit(i), .MinHit(i), .Def(i), .Valor(i), .ItemName(i), .PuedeUsar(i))

        End With

    Next i

    'Set state and show form
    Comerciando = True
    'Call Inventario.Initialize(frmComerciar.PicInvUser)
    frmComerciar.Show , frmMain
    frmComerciar.Refresh
    
    Exit Sub

HandleCommerceInit_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCommerceInit", Erl)
    
    
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
            Call frmBancoObj.InvBankUsu.SetItem(i, .ObjIndex(i), .Amount(i), .Equipped(i), .GrhIndex(i), .ObjType(i), .MaxHit(i), .MinHit(i), .Def(i), .Valor(i), .ItemName(i), .PuedeUsar(i))

        End With

    Next i

    'Set state and show form
    Comerciando = True

    frmBancoObj.lblcosto = PonerPuntos(UserGLD)
    frmBancoObj.Show , frmMain
    frmBancoObj.Refresh
    
    Exit Sub

HandleBankInit_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBankInit", Erl)
    
    
End Sub

Private Sub HandleGoliath()
    
    On Error GoTo HandleGoliathInit_Err

    '***************************************************
    '
    '***************************************************

    Dim UserBoveOro As Long

    Dim UserInvBove As Byte
    
    UserBoveOro = Reader.ReadInt32()
    UserInvBove = Reader.ReadInt8()
    Call frmGoliath.ParseBancoInfo(UserBoveOro, UserInvBove)
    
    Exit Sub

HandleGoliathInit_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGoliathInit", Erl)
    
    
End Sub

Private Sub HandleShowFrmLogear()
    
    On Error GoTo HandleShowFrmLogear_Err

    '***************************************************
    '
    '***************************************************
    FrmLogear.Show , frmConnect
    
    Exit Sub

HandleShowFrmLogear_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowFrmLogear", Erl)
    
    
End Sub

Private Sub HandleShowFrmMapa()
    
    On Error GoTo HandleShowFrmMapa_Err

    '***************************************************
    '
    '***************************************************
    ExpMult = Reader.ReadInt16()
    OroMult = Reader.ReadInt16()
    
    Call frmMapaGrande.CalcularPosicionMAPA

    frmMapaGrande.Picture = LoadInterface("ventanamapa.bmp")
    frmMapaGrande.Show , frmMain
    
    Exit Sub

HandleShowFrmMapa_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowFrmMapa", Erl)
    
    
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
            frmComerciarUsu.InvUser.SetItem i, .ObjIndex(i), .Amount(i), .Equipped(i), .GrhIndex(i), .ObjType(i), 0, 0, 0, 0, .ItemName(i), 0
        Next i

    End With
        
    frmComerciarUsu.lblMyGold.Caption = PonerPuntos(UserGLD)
    
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
    
    
End Sub

''
' Handles the UserSwing message.

Private Sub HandleCharSwing()
    
    On Error GoTo HandleCharSwing_Err
    
    Dim charindex As Integer

    charindex = Reader.ReadInt16
    
    Dim ShowFX As Boolean

    ShowFX = Reader.ReadBool
    
    Dim ShowText As Boolean

    ShowText = Reader.ReadBool
        
    With charlist(charindex)

        If ShowText Then
            Call SetCharacterDialogFx(charindex, IIf(charindex = UserCharIndex, "Fallas", "Falló"), RGBA_From_Comp(255, 0, 0))

        End If
        
        Call Sound.Sound_Play(2, False, Sound.Calculate_Volume(.Pos.x, .Pos.y), Sound.Calculate_Pan(.Pos.x, .Pos.y)) ' Swing
        
        ' If ShowFX And .Invisible = False Then Call SetCharacterFx(charindex, 90, 0)
         
        
    End With
    
    Exit Sub

HandleCharSwing_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCharSwing", Erl)
    
    
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
    
    
End Sub

Private Sub HandleClanSeguro()
    
    On Error GoTo HandleClanSeguro_Err

    '***************************************************
    'Author: Rapsodius
    'Creation date: 10/10/07
    '***************************************************
    Dim Seguro As Boolean
    
    'Get data and update form
    Seguro = Reader.ReadBool()
    
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
    
    
End Sub

Private Sub HandleIntervals()
    
    On Error GoTo HandleIntervals_Err

    IntervaloArco = Reader.ReadInt32()
    IntervaloCaminar = Reader.ReadInt32()
    IntervaloGolpe = Reader.ReadInt32()
    IntervaloGolpeMagia = Reader.ReadInt32()
    IntervaloMagia = Reader.ReadInt32()
    IntervaloMagiaGolpe = Reader.ReadInt32()
    IntervaloGolpeUsar = Reader.ReadInt32()
    IntervaloTrabajoExtraer = Reader.ReadInt32()
    IntervaloTrabajoConstruir = Reader.ReadInt32()
    IntervaloUsarU = Reader.ReadInt32()
    IntervaloUsarClic = Reader.ReadInt32()
    IntervaloTirar = Reader.ReadInt32()
    
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
    
    
End Sub

Private Sub HandleUpdateUserKey()
    
    On Error GoTo HandleUpdateUserKey_Err
 
    Dim Slot As Integer, Llave As Integer
    
    Slot = Reader.ReadInt16
    Llave = Reader.ReadInt16

    Call FrmKeyInv.InvKeys.SetItem(Slot, Llave, 1, 0, ObjData(Llave).GrhIndex, eObjType.otLlaves, 0, 0, 0, 0, ObjData(Llave).Name, 0)
    
    Exit Sub

HandleUpdateUserKey_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateUserKey", Erl)
    
    
End Sub

Private Sub HandleUpdateDM()
    
    On Error GoTo HandleUpdateDM_Err
 
    Dim Value As Integer

    Value = Reader.ReadInt16

    frmMain.lbldm = "+" & Value & "%"
    
    Exit Sub

HandleUpdateDM_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateDM", Erl)
    
    
End Sub

Private Sub HandleUpdateRM()
    
    On Error GoTo HandleUpdateRM_Err
 
    Dim Value As Integer

    Value = Reader.ReadInt16

    frmMain.lblResis = "+" & Value
    
    Exit Sub

HandleUpdateRM_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateRM", Erl)
    
    
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
    UserMinSTA = Reader.ReadInt16()
    frmMain.STAShp.Width = UserMinSTA / UserMaxSTA * 89
    frmMain.stabar.Caption = UserMinSTA & " / " & UserMaxSTA

    If QuePestañaInferior = 0 Then
        frmMain.STAShp.Visible = (UserMinSTA > 0)

    End If
    
    Exit Sub

HandleUpdateSta_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateSta", Erl)
    
    
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
    UserMinMAN = Reader.ReadInt16()
    
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
    NuevoValor = Reader.ReadInt16()
    
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
    UserGLD = Reader.ReadInt32()
    OroPorNivel = Reader.ReadInt32()
    
    frmMain.GldLbl.Caption = PonerPuntos(UserGLD)
    
    'If UserGLD > UserLvl * OroPorNivel Then
    If UserGLD <= 100000 Then
        frmMain.GldLbl.ForeColor = vbRed
    Else
        frmMain.GldLbl.ForeColor = &H80FFFF
    End If
    
    Exit Sub

HandleUpdateGold_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateGold", Erl)
    
    
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
    UserExp = Reader.ReadInt32()

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
    
    
End Sub

Private Sub HandleChangeMap()
    On Error GoTo HandleChangeMap_Err
    UserMap = Reader.ReadInt16()
    If bRain Then
        If Not MapDat.LLUVIA Then
            frmMain.IsPlaying = PlayLoop.plNone
        End If
    End If
    If frmComerciar.Visible Then Unload frmComerciar
    If frmBancoObj.Visible Then Unload frmBancoObj
    If frmEstadisticas.Visible Then Unload frmEstadisticas
    If frmStatistics.Visible Then Unload frmStatistics
    If frmHerrero.Visible Then Unload frmHerrero
    If FrmSastre.Visible Then Unload FrmSastre
    If frmAlqui.Visible Then Unload frmAlqui
    If frmCarp.Visible Then Unload frmCarp
    If FrmGrupo.Visible Then Unload FrmGrupo
    If frmGoliath.Visible Then Unload frmGoliath
    If FrmViajes.Visible Then Unload FrmViajes
    If frmCantidad.Visible Then Unload frmCantidad
    Call SwitchMap(UserMap)
    Exit Sub

HandleChangeMap_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChangeMap", Erl)
    
    
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
    UserPos.x = Reader.ReadInt8()
    UserPos.y = Reader.ReadInt8()

    'Set char
    MapData(UserPos.x, UserPos.y).charindex = UserCharIndex
    charlist(UserCharIndex).Pos = UserPos
        
    'Are we under a roof?
    bTecho = HayTecho(UserPos.x, UserPos.y)
                
    'Update pos label and minimap
    frmMain.Coord.Caption = UserMap & "-" & UserPos.x & "-" & UserPos.y
    Call frmMain.SetMinimapPosition(0, UserPos.x, UserPos.y)
    
    If MapDat.Seguro = 1 Then
        frmMain.Coord.ForeColor = RGB(0, 170, 0)
    Else
        frmMain.Coord.ForeColor = RGB(170, 0, 0)
    End If

    Call RefreshAllChars
    
    Exit Sub

HandlePosUpdate_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePosUpdate", Erl)
    
    
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
    
    Lugar = Reader.ReadInt8()

    DañoStr = PonerPuntos(Reader.ReadInt16)

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
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CRIATURA_1 & PonerPuntos(Reader.ReadInt32()) & MENSAJE_2, 255, 0, 0, True, False, False)
    
    Exit Sub

HandleUserHitNPC_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUserHitNPC", Erl)
    
    
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
    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & charlist(Reader.ReadInt16()).nombre & MENSAJE_ATAQUE_FALLO, 255, 0, 0, True, False, False)
    
    Exit Sub

HandleUserAttackedSwing_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUserAttackedSwing", Erl)
    
    
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
    
    intt = Reader.ReadInt16()
    
    Dim Pos As String

    Pos = InStr(charlist(intt).nombre, "<")
    
    If Pos = 0 Then Pos = Len(charlist(intt).nombre) + 2
    
    attacker = Left$(charlist(intt).nombre, Pos - 2)
    
    Dim Lugar As Byte
    Lugar = Reader.ReadInt8
    
    Dim DañoStr As String
    DañoStr = PonerPuntos(Reader.ReadInt16())
    
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
    
    intt = Reader.ReadInt16()
    'attacker = charlist().Nombre
    
    Dim Pos As String

    Pos = InStr(charlist(intt).nombre, "<")
    
    If Pos = 0 Then Pos = Len(charlist(intt).nombre) + 2
    
    victim = Left$(charlist(intt).nombre, Pos - 2)
    
    Dim Lugar As Byte
    Lugar = Reader.ReadInt8()
    
    Dim DañoStr As String
    DañoStr = PonerPuntos(Reader.ReadInt16())
    
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

    Dim EsSpell    As Boolean
    chat = Reader.ReadString8()
    charindex = Reader.ReadInt16()
    
    r = Reader.ReadInt8()
    G = Reader.ReadInt8()
    B = Reader.ReadInt8()
    
    colortexto = vbColor_2_Long(Reader.ReadInt32())
    EsSpell = Reader.ReadBool()

    'Optimizacion de protocolo por Ladder
    QueEs = ReadField(1, chat, Asc("*"))
    
    Dim copiar As Boolean

    copiar = True
    
    Dim duracion As Integer

    duracion = 250
    
    Select Case QueEs

        Case "NPCDESC"
            chat = NpcData(ReadField(2, chat, Asc("*"))).desc
            copiar = False

        Case "PMAG"
            chat = HechizoData(ReadField(2, chat, Asc("*"))).PalabrasMagicas
            copiar = False
            duracion = 20
            
        Case "QUESTFIN"
            chat = QuestList(ReadField(2, chat, Asc("*"))).DescFinal
            copiar = False
            duracion = 20
            
        Case "QUESTNEXT"
            chat = QuestList(ReadField(2, chat, Asc("*"))).NextQuest
            copiar = False
            duracion = 20
            
            If LenB(chat) = 0 Then
                chat = "Ya has completado esa misión para mí."

            End If
            
        Case "NOCONSOLA" ' El chat no sale en la consola
            chat = ReadField(2, chat, Asc("*"))
            copiar = False
            duracion = 20
        
    End Select
            
    'Only add the chat if the character exists (a CharacterRemove may have been sent to the PC / NPC area before the buffer was flushed)
    If charlist(charindex).active Then

        Call Char_Dialog_Set(charindex, chat, colortexto, duracion, 30, 1, EsSpell)

    End If
    
    If charlist(charindex).EsNpc = False Then
         
        If CopiarDialogoAConsola = 1 And copiar Then
    
            Call WriteChatOverHeadInConsole(charindex, chat, r, G, B)

        End If

    End If

    Exit Sub
    
ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChatOverHead", Erl)
    

End Sub

Private Sub HandleTextOverChar()

    On Error GoTo ErrHandler
    
    Dim chat      As String

    Dim charindex As Integer

    Dim Color     As Long
    
    chat = Reader.ReadString8()
    charindex = Reader.ReadInt16()
    
    Color = Reader.ReadInt32()
    
    Call SetCharacterDialogFx(charindex, chat, RGBA_From_vbColor(Color))

    Exit Sub
    
ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleTextOverChar", Erl)
    

End Sub

Private Sub HandleTextOverTile()

    On Error GoTo ErrHandler
    
    Dim Text  As String

    Dim x     As Integer, y As Integer

    Dim Color As Long
    
    Text = Reader.ReadString8()
    x = Reader.ReadInt16()
    y = Reader.ReadInt16()
    Color = Reader.ReadInt32()
    
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
    

End Sub

Private Sub HandleTextCharDrop()

    On Error GoTo ErrHandler
    
    Dim Text      As String

    Dim charindex As Integer

    Dim Color     As Long
    
    Text = Reader.ReadString8()
    charindex = Reader.ReadInt16()
    Color = Reader.ReadInt32()
    
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

    Exit Sub
    
ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleTextCharDrop", Erl)
    

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

    chat = Reader.ReadString8()
    FontIndex = Reader.ReadInt8()
    
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

        With FontTypes(FontIndex)
            Call AddtoRichTextBox(frmMain.RecTxt, chat, .red, .green, .blue, .bold, .italic)

        End With

    End If
    
    Exit Sub
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleConsoleMessage", Erl)
    

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

    Dim id        As Integer

    id = Reader.ReadInt16()
    chat = Reader.ReadString8()
    FontIndex = Reader.ReadInt8()

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

        With FontTypes(FontIndex)
            Call AddtoRichTextBox(frmMain.RecTxt, chat, .red, .green, .blue, .bold, .italic)

        End With

    End If
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleLocaleMsg", Erl)
    

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
    
    status = Reader.ReadInt8()
    chat = Reader.ReadString8()
    
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

    mensaje = Reader.ReadString8()

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
    
    userIndex = Reader.ReadInt16()
    
    Exit Sub

HandleUserIndexInServer_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUserIndexInServer", Erl)
    
    
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
    
    UserCharIndex = Reader.ReadInt16()
    'Debug.Print "UserCharIndex " & UserCharIndex
    UserPos = charlist(UserCharIndex).Pos
    
    'Are we under a roof?
    bTecho = HayTecho(UserPos.x, UserPos.y)
    
    LastMove = FrameTime
    
    frmMain.Coord.Caption = UserMap & "-" & UserPos.x & "-" & UserPos.y
    
    If MapDat.Seguro = 1 Then
        frmMain.Coord.ForeColor = RGB(0, 170, 0)
    Else
        frmMain.Coord.ForeColor = RGB(170, 0, 0)
    End If
    
    Call frmMain.SetMinimapPosition(0, UserPos.x, UserPos.y)
    
    If frmMapaGrande.Visible Then
        Call frmMapaGrande.ActualizarPosicionMapa
    End If
    
    Exit Sub

HandleUserCharIndexInServer_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUserCharIndexInServer", Erl)
    
    
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
    
    charindex = Reader.ReadInt16()
  
    Body = Reader.ReadInt16()
    Head = Reader.ReadInt16()
    Heading = Reader.ReadInt8()
    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
    weapon = Reader.ReadInt16()
    shield = Reader.ReadInt16()
    helmet = Reader.ReadInt16()
    
    With charlist(charindex)
        'Call SetCharacterFx(charindex, Reader.ReadInt16(), Reader.ReadInt16())
        .FxIndex = Reader.ReadInt16
        
        Reader.ReadInt16 'Ignore loops
        
        If .FxIndex > 0 Then
            Call InitGrh(.fX, FxData(.FxIndex).Animacion)

        End If
        
        Dim NombreYClan As String
        NombreYClan = Reader.ReadString8()
     
   ' Debug.Print "HandleCharacterCreate " & charindex & " " & NombreYClan & " x:" & x & " y:" & y
    
         
        Dim Pos As Integer
        Pos = InStr(NombreYClan, "<")

        If Pos = 0 Then Pos = InStr(NombreYClan, "[")
        If Pos = 0 Then Pos = Len(NombreYClan) + 2
        
        .nombre = Left$(NombreYClan, Pos - 2)
        .clan = mid$(NombreYClan, Pos)
        
        .status = Reader.ReadInt8()
        
        privs = Reader.ReadInt8()
        ParticulaFx = Reader.ReadInt8()
        .Head_Aura = Reader.ReadString8()
        .Arma_Aura = Reader.ReadString8()
        .Body_Aura = Reader.ReadString8()
        .DM_Aura = Reader.ReadString8()
        .RM_Aura = Reader.ReadString8()
        .Otra_Aura = Reader.ReadString8()
        .Escudo_Aura = Reader.ReadString8()
        .Speeding = Reader.ReadReal32()
        
        Dim FlagNpc As Byte
        FlagNpc = Reader.ReadInt8()
        
        .EsNpc = FlagNpc > 0
        .EsMascota = FlagNpc = 2
        
        .appear = Reader.ReadInt8()
        appear = .appear
        .group_index = Reader.ReadInt16()
        .clan_index = Reader.ReadInt16()
        .clan_nivel = Reader.ReadInt8()
        .UserMinHp = Reader.ReadInt32()
        .UserMaxHp = Reader.ReadInt32()
        .UserMinMAN = Reader.ReadInt32()
        .UserMaxMAN = Reader.ReadInt32()
        .simbolo = Reader.ReadInt8()
         Dim flags As Byte
        
        flags = Reader.ReadInt8()
        
                
        .Idle = flags And &O1
        
        .Navegando = flags And &O2
        .tipoUsuario = Reader.ReadInt8()
        
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
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCharacterCreate", Erl)
    

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
    Dim dbgid As Integer
    
    Dim Desvanecido As Boolean
    Dim fueWarp As Boolean
    charindex = Reader.ReadInt16()
    Desvanecido = Reader.ReadBool()
    fueWarp = Reader.ReadBool()
    If Desvanecido And charlist(charindex).EsNpc = True Then
        Call CrearFantasma(charindex)
    End If

    Call EraseChar(charindex, fueWarp)
    Call RefreshAllChars
    
    Exit Sub

HandleCharacterRemove_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCharacterRemove", Erl)
    
    
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
    
    charindex = Reader.ReadInt16()
    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
    
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
    
    
End Sub

''
' Handles the ForceCharMove message.

Private Sub HandleForceCharMove()
    
    On Error GoTo HandleForceCharMove_Err
    
    Dim Direccion As Byte
    Direccion = Reader.ReadInt8()
    
    Moviendose = True
    
    Call MainTimer.Restart(TimersIndex.Walk)

    Call Char_Move_by_Head(UserCharIndex, Direccion)
    Call MoveScreen(Direccion)
    
    Call frmMain.SetMinimapPosition(0, UserPos.x, UserPos.y)
    
    frmMain.Coord.Caption = UserMap & "-" & UserPos.x & "-" & UserPos.y
    
    If MapDat.Seguro = 1 Then
        frmMain.Coord.ForeColor = RGB(0, 170, 0)
    Else
        frmMain.Coord.ForeColor = RGB(170, 0, 0)
    End If

    If frmMapaGrande.Visible Then
        Call frmMapaGrande.ActualizarPosicionMapa

    End If
    
    Call RefreshAllChars
    
    Exit Sub

HandleForceCharMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleForceCharMove", Erl)
    
    
End Sub

''
' Handles the CharacterChange message.

Private Sub HandleCharacterChange()
    
    On Error GoTo HandleCharacterChange_Err
    
    Dim charindex As Integer

    Dim tempint   As Integer

    Dim headIndex As Integer

    charindex = Reader.ReadInt16()
    
    With charlist(charindex)
        tempint = Reader.ReadInt16()

        If tempint < LBound(BodyData()) Or tempint > UBound(BodyData()) Then
            .Body = BodyData(0)
        Else
            .Body = BodyData(tempint)
            .iBody = tempint

        End If
        
        headIndex = Reader.ReadInt16()

        If headIndex < LBound(HeadData()) Or headIndex > UBound(HeadData()) Then
            .Head = HeadData(0)
            .IHead = 0
            
        Else
            .Head = HeadData(headIndex)
            .IHead = headIndex

        End If

        .Muerto = (.iBody = CASPER_BODY_IDLE)
        
        .Heading = Reader.ReadInt8()
        
        tempint = Reader.ReadInt16()

        If TempInt <> 0 And TempInt <= UBound(WeaponAnimData) Then
            .Arma = WeaponAnimData(TempInt)
        End If

        tempint = Reader.ReadInt16()

        If TempInt <> 0 And TempInt <= UBound(ShieldAnimData) Then
            .Escudo = ShieldAnimData(TempInt)
        End If
        
        tempint = Reader.ReadInt16()

        If TempInt <> 0 And TempInt <= UBound(CascoAnimData) Then
            .Casco = CascoAnimData(TempInt)
        End If
                
        If .Body.HeadOffset.y = -26 Then
            .EsEnano = True
        Else
            .EsEnano = False

        End If
        
        'Call SetCharacterFx(charindex, Reader.ReadInt16(), Reader.ReadInt16())
        .FxIndex = Reader.ReadInt16
        
        Reader.ReadInt16 'Ignore loops
        
        If .FxIndex > 0 Then
            Call InitGrh(.fX, FxData(.FxIndex).Animacion)

        End If
        
        Dim flags As Byte
        
        flags = Reader.ReadInt8()
        
        .Idle = flags And &O1
        
        .Navegando = flags And &O2
        
        If .Idle Or .Navegando Then
            'Start animation
            .Body.Walk(.Heading).Started = FrameTime

        End If

    End With
    
    Exit Sub

HandleCharacterChange_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCharacterChange", Erl)
    
    
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

    Dim ObjIndex As Integer
    
    Dim Amount   As Integer

    Dim Color    As RGBA

    Dim Rango    As Byte

    Dim id       As Long
    
    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
    
    ObjIndex = Reader.ReadInt16()
    
    Amount = Reader.ReadInt16
    
    MapData(x, y).ObjGrh.GrhIndex = ObjData(ObjIndex).GrhIndex
    
    MapData(x, y).OBJInfo.ObjIndex = ObjIndex
    
    MapData(x, y).OBJInfo.Amount = Amount
    
    Call InitGrh(MapData(x, y).ObjGrh, MapData(x, y).ObjGrh.GrhIndex)
    
    If ObjData(ObjIndex).CreaLuz <> "" Then
        Call Long_2_RGBA(Color, Val(ReadField(2, ObjData(ObjIndex).CreaLuz, Asc(":"))))
        Rango = Val(ReadField(1, ObjData(ObjIndex).CreaLuz, Asc(":")))
        MapData(x, y).luz.Color = Color
        MapData(x, y).luz.Rango = Rango
        
        If Rango < 100 Then
            id = x & y
            LucesCuadradas.Light_Create x, y, Color, Rango, id
        Else
            LucesRedondas.Create_Light_To_Map x, y, Color, Rango - 99
        End If
        
    End If
        
    If ObjData(ObjIndex).CreaParticulaPiso <> 0 Then
        MapData(x, y).particle_group = 0
        General_Particle_Create ObjData(ObjIndex).CreaParticulaPiso, x, y, -1

    End If
    
    Exit Sub

HandleObjectCreate_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleObjectCreate", Erl)
    
    
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

    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
    fX = Reader.ReadInt16()
    
    Call SetMapFx(x, y, fX, 0)
    
    Exit Sub

HandleFxPiso_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleFxPiso", Erl)
    
    
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

    Dim id As Long
    
    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
    
    If ObjData(MapData(x, y).OBJInfo.ObjIndex).CreaLuz <> "" Then
        id = LucesCuadradas.Light_Find(x & y)
        LucesCuadradas.Light_Remove id
        MapData(x, y).luz.Color = COLOR_EMPTY
        MapData(x, y).luz.Rango = 0
       ' LucesCuadradas.Light_Render_All

    End If
    
    MapData(x, y).ObjGrh.GrhIndex = 0
    MapData(x, y).OBJInfo.ObjIndex = 0
    
    If ObjData(MapData(x, y).OBJInfo.ObjIndex).CreaParticulaPiso <> 0 Then
        Graficos_Particulas.Particle_Group_Remove (MapData(x, y).particle_group)

    End If
    
    Exit Sub

HandleObjectDelete_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleObjectDelete", Erl)
    
    
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
    
    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
    B = Reader.ReadInt8()

    MapData(x, y).Blocked = MapData(x, y).Blocked And Not eBlock.ALL_SIDES
    MapData(x, y).Blocked = MapData(x, y).Blocked Or B
    
    Exit Sub

HandleBlockPosition_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBlockPosition", Erl)
    
    
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
    
    Call Reader.ReadInt8   ' File
    Call Reader.ReadInt16  ' Loop
    
    Exit Sub

HandlePlayMIDI_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePlayMIDI", Erl)
    
    
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
    Dim cancelLastWave As Byte
    
    wave = Reader.ReadInt16()
    srcX = Reader.ReadInt8()
    srcY = Reader.ReadInt8()
    cancelLastWave = Reader.ReadInt8()
    
    If wave = 400 And MapDat.niebla = 0 Then Exit Sub
    If wave = 401 And MapDat.niebla = 0 Then Exit Sub
    If wave = 402 And MapDat.niebla = 0 Then Exit Sub
    If wave = 403 And MapDat.niebla = 0 Then Exit Sub
    If wave = 404 And MapDat.niebla = 0 Then Exit Sub
    
    If cancelLastWave Then
        Call Sound.Sound_Stop(CStr(wave))
        If cancelLastWave = 2 Then Exit Sub
    End If
    
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
    
    map = Reader.ReadInt16()
    srcX = Reader.ReadInt8()
    srcY = Reader.ReadInt8()

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
    
    
End Sub

Private Sub HandleCharUpdateHP()
    
    On Error GoTo HandleCharUpdateHP_Err

    Dim charindex As Integer

    Dim minhp     As Long

    Dim maxhp     As Long
    
    charindex = Reader.ReadInt16()
    minhp = Reader.ReadInt32()
    maxhp = Reader.ReadInt32()

    charlist(charindex).UserMinHp = minhp
    charlist(charindex).UserMaxHp = maxhp
    
    Exit Sub

HandleCharUpdateHP_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCharUpdateHP", Erl)
    
    
End Sub

Private Sub HandleCharUpdateMAN()
    
    On Error GoTo HandleCharUpdateHP_Err

    Dim charindex As Integer

    Dim minman     As Long

    Dim maxman     As Long
    
    charindex = Reader.ReadInt16()
    minman = Reader.ReadInt32()
    maxman = Reader.ReadInt32()

    charlist(charindex).UserMinMAN = minman
    charlist(charindex).UserMaxMAN = maxman
    
    Exit Sub

HandleCharUpdateHP_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCharUpdateMAN", Erl)
    
    
End Sub

Private Sub HandleArmaMov()
    
    On Error GoTo HandleArmaMov_Err

    '***************************************************

    Dim charindex As Integer

    charindex = Reader.ReadInt16()

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
    
    
End Sub

Private Sub HandleEscudoMov()
    
    On Error GoTo HandleEscudoMov_Err

    '***************************************************

    Dim charindex As Integer

    charindex = Reader.ReadInt16()

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
    guildsStr = Reader.ReadString8()
    
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
    
    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
        
    Call CambioDeArea(x, y)
    
    Exit Sub

HandleAreaChanged_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleAreaChanged", Erl)
    
    
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
    
    charindex = Reader.ReadInt16()
    fX = Reader.ReadInt16()
    Loops = Reader.ReadInt16()
    
    If fX = 0 Then
        charlist(charindex).fX.AnimacionContador = 29
        Exit Sub

    End If
    
    Call SetCharacterFx(charindex, fX, Loops)
    
    Exit Sub

HandleCreateFX_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCreateFX", Erl)
    
    
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
    
    UserMaxHp = Reader.ReadInt16()
    UserMinHp = Reader.ReadInt16()
    UserMaxMAN = Reader.ReadInt16()
    UserMinMAN = Reader.ReadInt16()
    UserMaxSTA = Reader.ReadInt16()
    UserMinSTA = Reader.ReadInt16()
    UserGLD = Reader.ReadInt32()
    OroPorNivel = Reader.ReadInt32()
    UserLvl = Reader.ReadInt8()
    UserPasarNivel = Reader.ReadInt32()
    UserExp = Reader.ReadInt32()
    UserClase = Reader.ReadInt8()
    
    If UserPasarNivel > 0 Then
        frmMain.lblPorcLvl.Caption = Round(UserExp * (100 / UserPasarNivel), 2) & "%"
        frmMain.exp.Caption = PonerPuntos(UserExp) & "/" & PonerPuntos(UserPasarNivel)
        frmMain.EXPBAR.Width = UserExp / UserPasarNivel * 235
    Else
        frmMain.EXPBAR.Width = 235
        frmMain.lblPorcLvl.Caption = "¡Nivel máximo!" 'nivel maximo
        frmMain.exp.Caption = "¡Nivel máximo!"

    End If
    
    If UserMaxHp > 0 Then
        frmMain.Hpshp.Width = UserMinHp / UserMaxHp * 216
    Else
        frmMain.Hpshp.Width = 0
    End If

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
    
    If UserMaxSTA > 0 Then
        frmMain.STAShp.Width = UserMinSTA / UserMaxSTA * 89
    Else
        frmMain.STAShp.Width = 0
    End If

    frmMain.stabar.Caption = UserMinSTA & " / " & UserMaxSTA
    
    If QuePestañaInferior = 0 Then
        frmMain.STAShp.Visible = (UserMinSTA > 0)

    End If
    
    'If UserGLD > UserLvl * OroPorNivel Then
    If UserGLD <= 100000 Then
        frmMain.GldLbl.ForeColor = vbRed
    Else
        frmMain.GldLbl.ForeColor = &H80FFFF
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
    
    UsingSkillREcibido = Reader.ReadInt8()
    casteaArea = Reader.ReadBool()
    RadioHechizoArea = Reader.ReadInt8()
    'RadioHechizoArea = RadioHechizoArea / 2

    If UsingSkillREcibido = 0 Then
        frmMain.MousePointer = 0
        Call FormParser.Parse_Form(frmMain, E_NORMAL)
        UsingSkill = UsingSkillREcibido
        Exit Sub

    End If

    If UsingSkillREcibido = UsingSkill Then Exit Sub
   
    UsingSkill = UsingSkillREcibido
    frmMain.MousePointer = 2
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
    Dim ObjIndex    As Integer
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

    Slot = Reader.ReadInt8()
    ObjIndex = Reader.ReadInt16()
    Amount = Reader.ReadInt16()
    Equipped = Reader.ReadBool()
    Value = Reader.ReadReal32()
    podrausarlo = Reader.ReadInt8()

    Name = ObjData(ObjIndex).Name
    GrhIndex = ObjData(ObjIndex).GrhIndex
    ObjType = ObjData(ObjIndex).ObjType
    MaxHit = ObjData(ObjIndex).MaxHit
    MinHit = ObjData(ObjIndex).MinHit
    MaxDef = ObjData(ObjIndex).MaxDef
    MinDef = ObjData(ObjIndex).MinDef

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

    Call frmMain.Inventario.SetItem(Slot, ObjIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, MinDef, Value, Name, podrausarlo)

    If frmComerciar.Visible Then
        Call frmComerciar.InvComUsu.SetItem(Slot, ObjIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, MinDef, Value, Name, podrausarlo)

    ElseIf frmBancoObj.Visible Then
        Call frmBancoObj.InvBankUsu.SetItem(Slot, ObjIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, MinDef, Value, Name, podrausarlo)
        
    ElseIf frmBancoCuenta.Visible Then
        Call frmBancoCuenta.InvBankUsuCuenta.SetItem(Slot, ObjIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, MinDef, Value, Name, podrausarlo)
    
    ElseIf frmCrafteo.Visible Then
        Call frmCrafteo.InvCraftUser.SetItem(Slot, ObjIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, MinDef, Value, Name, podrausarlo)
    End If

    Exit Sub
    
ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChangeInventorySlot", Erl)
    

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
    
    UserInvUnlocked = Reader.ReadInt8
    
    For i = 1 To UserInvUnlocked
    
        frmMain.imgInvLock(i - 1).Picture = LoadInterface("inventoryunlocked.bmp")
    
    Next i

    Exit Sub

HandleInventoryUnlockSlots_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleInventoryUnlockSlots", Erl)
    
    
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
    Dim BankSlot As Inventory
    
    With BankSlot
        Slot = Reader.ReadInt8()
        .ObjIndex = Reader.ReadInt16()
        .Amount = Reader.ReadInt16()
        .Valor = Reader.ReadInt32()
        .PuedeUsar = Reader.ReadInt8()
        
        If .ObjIndex > 0 Then
            .Name = ObjData(.ObjIndex).Name
            .GrhIndex = ObjData(.ObjIndex).GrhIndex
            .ObjType = ObjData(.ObjIndex).ObjType
            .MaxHit = ObjData(.ObjIndex).MaxHit
            .MinHit = ObjData(.ObjIndex).MinHit
            .Def = ObjData(.ObjIndex).MaxDef
        End If
        
        Call frmBancoObj.InvBoveda.SetItem(Slot, .ObjIndex, .Amount, .Equipped, .GrhIndex, .ObjType, .MaxHit, .MinHit, .Def, .Valor, .Name, .PuedeUsar)

    End With
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChangeBankSlot", Erl)
    

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

    Slot = Reader.ReadInt8()
    
    UserHechizos(Slot) = Reader.ReadInt16()
    Index = Reader.ReadInt8()

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
        UserAtributos(i) = Reader.ReadInt8()
    Next i
    
    'Show them in character creation

    
    If LlegaronStats Then
        frmStatistics.Iniciar_Labels
        frmStatistics.Picture = LoadInterface("ventanaestadisticas_personaje.bmp")
        frmStatistics.Show , frmMain
    Else
        LlegaronAtrib = True
    End If
    

    
    Exit Sub

HandleAtributes_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleAtributes", Erl)
    
    
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
    
    count = Reader.ReadInt16()
    
    Call frmHerrero.lstArmas.Clear
    
    For i = 1 To count
        ArmasHerrero(i).Index = Reader.ReadInt16()
        ' tmp = ObjData(ArmasHerrero(i).Index).name        'Get the object's name
        ArmasHerrero(i).LHierro = Reader.ReadInt16()  'The iron needed
        ArmasHerrero(i).LPlata = Reader.ReadInt16()    'The silver needed
        ArmasHerrero(i).LOro = Reader.ReadInt16()    'The gold needed
        
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
    
    count = Reader.ReadInt16()
    
    'Call frmHerrero.lstArmaduras.Clear
    
    For i = 1 To count
        tmp = Reader.ReadString8()         'Get the object's name
        DefensasHerrero(i).LHierro = Reader.ReadInt16()   'The iron needed
        DefensasHerrero(i).LPlata = Reader.ReadInt16()   'The silver needed
        DefensasHerrero(i).LOro = Reader.ReadInt16()   'The gold needed
        
        ' Call frmHerrero.lstArmaduras.AddItem(tmp)
        DefensasHerrero(i).Index = Reader.ReadInt16()
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
        If tmpObj.ObjType = 16 Or tmpObj.ObjType = 35 Or tmpObj.ObjType = 21 Or tmpObj.ObjType = 100 Then
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
    
    count = Reader.ReadInt8()
    
    Call frmCarp.lstArmas.Clear
    
    For i = 1 To count
        ObjCarpintero(i) = Reader.ReadInt16()
        
        Call frmCarp.lstArmas.AddItem(ObjData(ObjCarpintero(i)).Name)
    Next i
    
    For i = i To UBound(ObjCarpintero())
        ObjCarpintero(i) = 0
    Next i
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCarpenterObjects", Erl)
    

End Sub

Private Sub HandleSastreObjects()

    '***************************************************
    'Author: Ladder
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim count As Integer

    Dim i     As Long

    Dim tmp   As String
    
    count = Reader.ReadInt16()
    
    For i = i To UBound(ObjSastre())
        ObjSastre(i).Index = 0
    Next i
    
    i = 0
    
    For i = 1 To count
        ObjSastre(i).Index = Reader.ReadInt16()
        
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
    
        If ObjData(ObjSastre(i).Index).ObjType = 3 Or ObjData(ObjSastre(i).Index).ObjType = 100 Then
        
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

    count = Reader.ReadInt16()
    
    Call frmAlqui.lstArmas.Clear
    
    For i = 1 To count
        Obj = Reader.ReadInt16()
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
    
    Call MsgBox(Reader.ReadString8())
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleErrorMessage", Erl)
    

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

    tmp = ObjData(Reader.ReadInt16()).Texto
    grh = Reader.ReadInt16()
    
    Call InitCartel(tmp, grh)
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowSignal", Erl)
    

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
    Slot = Reader.ReadInt8()
    
    Dim SlotInv As NpCinV

    With SlotInv
        .ObjIndex = Reader.ReadInt16()
        .Name = ObjData(.ObjIndex).Name
        .Amount = Reader.ReadInt16()
        .Valor = Reader.ReadReal32()
        .GrhIndex = ObjData(.ObjIndex).GrhIndex
        .ObjType = ObjData(.ObjIndex).ObjType
        .MaxHit = ObjData(.ObjIndex).MaxHit
        .MinHit = ObjData(.ObjIndex).MinHit
        .Def = ObjData(.ObjIndex).MaxDef
        .PuedeUsar = Reader.ReadInt8()
        
        Call frmComerciar.InvComNpc.SetItem(Slot, .ObjIndex, .Amount, 0, .GrhIndex, .ObjType, .MaxHit, .MinHit, .Def, .Valor, .Name, .PuedeUsar)
        
    End With
    
    Exit Sub
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChangeNPCInventorySlot", Erl)
    

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
    
    UserMaxAGU = Reader.ReadInt8()
    UserMinAGU = Reader.ReadInt8()
    UserMaxHAM = Reader.ReadInt8()
    UserMinHAM = Reader.ReadInt8()
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
    
    
End Sub

Private Sub HandleHora()
    '***************************************************
    
    On Error GoTo HandleHora_Err

    HoraMundo = GetTickCount() - Reader.ReadInt32()
    DuracionDia = Reader.ReadInt32()
    
    If Not Connected Then
        Call RevisarHoraMundo(True)

    End If
    
    Exit Sub

HandleHora_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleHora", Erl)
    
    
End Sub
 
Private Sub HandleLight()
    
    On Error GoTo HandleLight_Err
 
    Dim Color As String
    
    Color = Reader.ReadString8()

    'Call SetGlobalLight(Map_light_base)
    
    Exit Sub

HandleLight_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleLight", Erl)
    
    
End Sub
 
Private Sub HandleFYA()
    
    On Error GoTo HandleFYA_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    UserAtributos(eAtributos.Fuerza) = Reader.ReadInt8()
    UserAtributos(eAtributos.Agilidad) = Reader.ReadInt8()
    
    DrogaCounter = Reader.ReadInt16()
    
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
    
    NpcIndex = Reader.ReadInt16()
    
    simbolo = Reader.ReadInt8()

    charlist(NpcIndex).simbolo = simbolo
    
    Exit Sub

HandleUpdateNPCSimbolo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateNPCSimbolo", Erl)
    
    
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
    
    
End Sub

Private Sub HandleContadores()
    
    On Error GoTo HandleContadores_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    InviCounter = Reader.ReadInt16()
    DrogaCounter = Reader.ReadInt16()
    

    frmMain.Contadores.Enabled = True
    
    Exit Sub

HandleContadores_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleContadores", Erl)
    
    
End Sub

Private Sub HandleShowPapiro()
    On Error GoTo HandleShowPapiro_Err
    
    frmMensajePapiro.Show , frmMain
    
    'incomingdata papiromessage
    Exit Sub

HandleShowPapiro_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowPapiro", Erl)
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
    
    Color = Reader.ReadInt32()
    duracion = Reader.ReadInt32()
    ignorar = Reader.ReadBool()
    
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
    
    
End Sub

Private Sub HandleMiniStats()
    
    On Error GoTo HandleMiniStats_Err
    
    With UserEstadisticas
        .CiudadanosMatados = Reader.ReadInt32()
        .CriminalesMatados = Reader.ReadInt32()
        .Alineacion = Reader.ReadInt8()
        
        .NpcsMatados = Reader.ReadInt16()
        .Clase = ListaClases(Reader.ReadInt8())
        .PenaCarcel = Reader.ReadInt32()
        .VecesQueMoriste = Reader.ReadInt32()
        .Genero = Reader.ReadInt8()
        .PuntosPesca = Reader.ReadInt32()

        If .Genero = 1 Then
            .Genero = "Hombre"
        Else
            .Genero = "Mujer"

        End If

        .Raza = Reader.ReadInt8()
        .Raza = ListaRazas(.Raza)
    End With
    
    If LlegaronAtrib Then
        frmStatistics.Iniciar_Labels
        frmStatistics.Picture = LoadInterface("ventanaestadisticas_personaje.bmp")
        frmStatistics.Show , frmMain
    Else
        LlegaronStats = True
    End If
    
    Exit Sub

HandleMiniStats_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleMiniStats", Erl)
    
    
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
    
    SkillPoints = Reader.ReadInt16()
    
    Exit Sub

HandleLevelUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleLevelUp", Erl)
    
    
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
    
    title = Reader.ReadString8()
    Message = Reader.ReadString8()

    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleAddForumMessage", Erl)
    

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
    
    charindex = Reader.ReadInt16()
    charlist(charindex).Invisible = Reader.ReadBool()
    charlist(charindex).TimerI = 0
    
    Exit Sub

HandleSetInvisible_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSetInvisible", Erl)
    
    
End Sub


''
' Handles the MeditateToggle message.

Private Sub HandleMeditateToggle()
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleMeditateToggle_Err
    
    Dim charindex As Integer, fX As Integer
    
    charindex = Reader.ReadInt16
    fX = Reader.ReadInt16
    
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
        UserSkills(i) = Reader.ReadInt8()
        'frmEstadisticas.skills(i).Caption = SkillsNames(i)
    Next i

    If LlegaronSkills Then
        Alocados = SkillPoints
        frmEstadisticas.puntos.Caption = SkillPoints
        frmEstadisticas.Iniciar_Labels
        frmEstadisticas.Picture = LoadInterface("ventanaskills.bmp")
        frmEstadisticas.Show , frmMain
        LlegaronSkills = False
    End If
    
    Exit Sub

HandleSendSkills_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSendSkills", Erl)
    
    
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
    
    creatures = Split(Reader.ReadString8(), SEPARATOR)
    
    For i = 0 To UBound(creatures())
        Call frmEntrenador.lstCriaturas.AddItem(creatures(i))
    Next i

    frmEntrenador.Show , frmMain
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleTrainerCreatureList", Erl)
    

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
        
    frmGuildNews.news = Reader.ReadString8()
    
    'Get list of existing guilds
    List = Split(Reader.ReadString8(), SEPARATOR)
        
    'Empty the list
    Call frmGuildNews.guildslist.Clear
        
    For i = 0 To UBound(List())
        Call frmGuildNews.guildslist.AddItem(ReadField(1, List(i), Asc("-")))
    Next i
    
    'Get  guilds list member
    guildList = Split(Reader.ReadString8(), SEPARATOR)
    
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
    
    ClanNivel = Reader.ReadInt8()
    expacu = Reader.ReadInt16()
    ExpNe = Reader.ReadInt16()
     
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
                .beneficios = "Max miembros: 5"

            Case 2
                .beneficios = "Pedir ayuda (G) / Max miembros: 7"

            Case 3
                .beneficios = "Pedir ayuda (G) / Seguro de clan." & vbCrLf & "Max miembros: 7"

            Case 4
                .beneficios = "Pedir ayuda (G) / Seguro de clan. " & vbCrLf & "Max miembros: 12"

            Case 5
                .beneficios = "Pedir ayuda (G) / Seguro de clan /  Ver vida y mana." & vbCrLf & " Max miembros: 15"
                
            Case 6
                .beneficios = "Pedir ayuda (G) / Seguro de clan / Ver vida y mana/ Verse invisible." & vbCrLf & " Max miembros: 20"
        
        End Select
    
    End With
    
    frmGuildNews.Show vbModeless, frmMain
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildNews", Erl)
    

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
    
    Call frmUserRequest.recievePeticion(Reader.ReadString8())
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleOfferDetails", Erl)
    

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
    
    guildList = Split(Reader.ReadString8(), SEPARATOR)
    
    For i = 0 To UBound(guildList())
        Call frmPeaceProp.lista.AddItem(guildList(i))
    Next i
    
    frmPeaceProp.ProposalType = TIPO_PROPUESTA.ALIANZA
    Call frmPeaceProp.Show(vbModeless, frmMain)
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleAlianceProposalsList", Erl)
    

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
    
    guildList = Split(Reader.ReadString8(), SEPARATOR)
    
    For i = 0 To UBound(guildList())
        Call frmPeaceProp.lista.AddItem(guildList(i))
    Next i
    
    frmPeaceProp.ProposalType = TIPO_PROPUESTA.PAZ
    Call frmPeaceProp.Show(vbModeless, frmMain)
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePeaceProposalsList", Erl)
    

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
    
        If Reader.ReadInt8() = 1 Then
            .Genero.Caption = "Genero: Hombre"
        Else
            .Genero.Caption = "Genero: Mujer"
        End If
            
        .nombre.Caption = "Nombre: " & Reader.ReadString8()
        .Raza.Caption = "Raza: " & ListaRazas(Reader.ReadInt8())
        .Clase.Caption = "Clase: " & ListaClases(Reader.ReadInt8())

        .nivel.Caption = "Nivel: " & Reader.ReadInt8()
        .oro.Caption = "Oro: " & Reader.ReadInt32()
        .Banco.Caption = "Banco: " & Reader.ReadInt32()
    
        .txtPeticiones.Text = Reader.ReadString8()
        .guildactual.Caption = "Clan: " & Reader.ReadString8()
        .txtMiembro.Text = Reader.ReadString8()
            
        Dim armada As Boolean
    
        Dim caos   As Boolean
            
        armada = Reader.ReadBool()
        caos = Reader.ReadBool()
            
        If armada Then
            .ejercito.Caption = "Ejército: Armada Real"
        ElseIf caos Then
            .ejercito.Caption = "Ejército: Legión Oscura"
    
        End If
            
        .ciudadanos.Caption = "Ciudadanos asesinados: " & CStr(Reader.ReadInt32())
        .Criminales.Caption = "Criminales asesinados: " & CStr(Reader.ReadInt32())
    
        Call .Show(vbModeless, frmMain)
    
    End With
        
    Exit Sub
    
ErrHandler:
    
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCharacterInfo", Erl)
    

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
    
    Dim str As String
    
    Dim List() As String

    Dim i      As Long
    
    With frmGuildLeader
        'Empty the list
        Call .guildslist.Clear
    
        str = Reader.ReadString8()
    
        If LenB(str) > 0 Then
            'Get list of existing guilds
            List = Split(str, SEPARATOR)

            For i = 0 To UBound(List())
                Call .guildslist.AddItem(ReadField(1, List(i), Asc("-")))
            Next i
        End If
        
        'Empty the list
        Call .members.Clear
        
        str = Reader.ReadString8()
        
        If LenB(str) > 0 Then
            'Get list of guild's members
            List = Split(str, SEPARATOR)
            .miembros.Caption = CStr(UBound(List()) + 1)

            For i = 0 To UBound(List())
                Call .members.AddItem(List(i))
            Next i
        End If
        
        .txtguildnews = Reader.ReadString8()
        
        'Empty the list
        Call .solicitudes.Clear
        
        str = Reader.ReadString8()
        
        If LenB(str) > 0 Then
            'Get list of join requests
            List = Split(str, SEPARATOR)
        
            For i = 0 To UBound(List())
                Call .solicitudes.AddItem(List(i))
            Next i
        End If
        
        Dim expacu As Integer

        Dim ExpNe  As Integer

        Dim nivel  As Byte
         
        nivel = Reader.ReadInt8()
        .nivel = "Nivel: " & nivel
        
        expacu = Reader.ReadInt16()
        ExpNe = Reader.ReadInt16()
        'barra
        .expcount.Caption = expacu & "/" & ExpNe
        
        If ExpNe > 0 Then
            .EXPBAR.Width = expacu / ExpNe * 239
            .porciento.Caption = Round(expacu / ExpNe * 100#, 0) & "%"
        Else
            .EXPBAR.Width = 239
            .porciento.Caption = "¡Nivel máximo!"
            .expcount.Caption = "¡Nivel máximo!"
        End If

        Select Case nivel

               Case 1
                .beneficios = "Max miembros: 5"
                .maxMiembros = 5
            Case 2
                .beneficios = "Pedir ayuda (G) / Max miembros: 7"
                .maxMiembros = 7

            Case 3
                .beneficios = "Pedir ayuda (G) / Seguro de clan." & vbCrLf & "Max miembros: 7"
                .maxMiembros = 7

            Case 4
                .beneficios = "Pedir ayuda (G) / Seguro de clan. " & vbCrLf & "Max miembros: 12"
                .maxMiembros = 12

            Case 5
                .beneficios = "Pedir ayuda (G) / Seguro de clan /  Ver vida y mana." & vbCrLf & " Max miembros: 15"
                .maxMiembros = 15
                
            Case 6
                .beneficios = "Pedir ayuda (G) / Seguro de clan / Ver vida y mana/ Verse invisible." & vbCrLf & " Max miembros: 20"
                .maxMiembros = 20
        End Select
        
        .Show , frmMain

    End With
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildLeaderInfo", Erl)
    

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
        
        .nombre.Caption = "Nombre:" & Reader.ReadString8()
        .fundador.Caption = "Fundador:" & Reader.ReadString8()
        .creacion.Caption = "Fecha de creacion:" & Reader.ReadString8()
        .lider.Caption = "Líder:" & Reader.ReadString8()
        .miembros.Caption = "Miembros:" & Reader.ReadInt16()
        
        .lblAlineacion.Caption = "Alineación: " & Reader.ReadString8()
        
        .desc.Text = Reader.ReadString8()
        .nivel.Caption = "Nivel de clan: " & Reader.ReadInt8()

    End With
    
    frmGuildBrief.Show vbModeless, frmMain
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildDetails", Erl)
    

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
    
    Call frmUserRequest.recievePeticion(Reader.ReadString8())
    Call frmUserRequest.Show(vbModeless, frmMain)
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowUserRequest", Erl)
    

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
    
    miOferta = Reader.ReadBool
    Dim i          As Byte
    Dim nombreItem As String
    Dim cantidad   As Integer
    Dim grhItem    As Long
    Dim ObjIndex   As Integer

    If miOferta Then
        Dim OroAEnviar As Long
        OroAEnviar = Reader.ReadInt32
        frmComerciarUsu.lblOroMiOferta.Caption = PonerPuntos(OroAEnviar)
        frmComerciarUsu.lblMyGold.Caption = PonerPuntos(UserGLD - OroAEnviar)

        For i = 1 To 6

            With OtroInventario(i)
                ObjIndex = Reader.ReadInt16
                nombreItem = Reader.ReadString8
                grhItem = Reader.ReadInt32
                cantidad = Reader.ReadInt32

                If cantidad > 0 Then
                    Call frmComerciarUsu.InvUserSell.SetItem(i, ObjIndex, cantidad, 0, grhItem, 0, 0, 0, 0, 0, nombreItem, 0)

                End If

            End With

        Next i
        
        Call frmComerciarUsu.InvUserSell.ReDraw
    Else
        frmComerciarUsu.lblOro.Caption = PonerPuntos(Reader.ReadInt32)

        ' frmComerciarUsu.List2.Clear
        For i = 1 To 6
            
            With OtroInventario(i)
                ObjIndex = Reader.ReadInt16
                nombreItem = Reader.ReadString8
                grhItem = Reader.ReadInt32
                cantidad = Reader.ReadInt32

                If cantidad > 0 Then
                    Call frmComerciarUsu.InvOtherSell.SetItem(i, ObjIndex, cantidad, 0, grhItem, 0, 0, 0, 0, 0, nombreItem, 0)

                End If

            End With

        Next i
        
        Call frmComerciarUsu.InvOtherSell.ReDraw
    
    End If
    
    frmComerciarUsu.lblEstadoResp.Visible = False
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChangeUserTradeSlot", Erl)
    

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
    
    frmSpawnList.ListaCompleta = Reader.ReadBool

    Call frmSpawnList.FillList

    frmSpawnList.Show , frmMain
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSpawnList", Erl)
    

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
    
    sosList = Split(Reader.ReadString8(), SEPARATOR)
    
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
    
    frmCambiaMotd.txtMotd.Text = Reader.ReadString8()
    frmCambiaMotd.Show , frmMain
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowMOTDEditionForm", Erl)
    

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
    
    Dim MiCargo As Integer
    
    
    frmPanelgm.txtHeadNumero = Reader.ReadInt16
    frmPanelgm.txtBodyYo = Reader.ReadInt16
    frmPanelgm.txtCasco = Reader.ReadInt16
    frmPanelgm.txtArma = Reader.ReadInt16
    frmPanelgm.txtEscudo = Reader.ReadInt16
    frmPanelgm.Show vbModeless, frmMain
    
    MiCargo = charlist(UserCharIndex).priv
    
    Select Case MiCargo ' ReyarB ajustar privilejios
    
        Case 1
        frmPanelgm.mnuChar.Visible = False
        frmPanelgm.cmdHerramientas.Visible = False
        frmPanelgm.Admin(0).Visible = False
        
        Case 2 'Consejeros
        frmPanelgm.mnuChar.Visible = False
        frmPanelgm.cmdHerramientas.Visible = False
        frmPanelgm.Admin(0).Visible = False
        frmPanelgm.cmdConsulta.Visible = False
        frmPanelgm.cmdMatarNPC.Visible = False
        frmPanelgm.cmdEventos.Visible = False
        frmPanelgm.cmdBody0(2).Visible = False
        frmPanelgm.cmdHead0.Visible = False
        frmPanelgm.SendGlobal.Visible = False
        frmPanelgm.Mensajeria.Visible = False
        frmPanelgm.cmdMapeo.Visible = False
        frmPanelgm.cmdMapeo.Enabled = False
        frmPanelgm.cmdcrearevento.Enabled = False
        frmPanelgm.cmdcrearevento.Visible = False
        frmPanelgm.txtMod.Width = 4580
        frmPanelgm.Height = 7580
        frmPanelgm.mnuTraer.Visible = False
        frmPanelgm.mnuIra.Visible = False
                
        Case 3 ' Semidios
        frmPanelgm.mnuChar.Visible = False
        frmPanelgm.mnuChar.Visible = False
        frmPanelgm.cmdHerramientas.Visible = True
        frmPanelgm.Admin(0).Visible = False
        frmPanelgm.cmdcrearevento.Enabled = False
        frmPanelgm.cmdcrearevento.Visible = False
        frmPanelgm.mnuHerramientas(23).Visible = False
        
        Case 4 ' Dios
        frmPanelgm.mnuChar.Visible = True
        frmPanelgm.mnuChar.Visible = True
        frmPanelgm.cmdHerramientas.Visible = True
        frmPanelgm.Admin(0).Visible = False
        
        Case 5
        
    
    End Select
    
    Exit Sub

HandleShowGMPanelForm_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowGMPanelForm", Erl)
    
    
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
    
    userList = Split(Reader.ReadString8(), SEPARATOR)
    
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
    Time = Reader.ReadInt32()

    PingRender = GetTickCount() - Time

    Exit Sub

HandlePong_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePong", Erl)
    
    
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
    
    charindex = Reader.ReadInt16()
    status = Reader.ReadInt8()
    NombreYClan = Reader.ReadString8()
        
    Dim Pos As Integer
    Pos = InStr(NombreYClan, "<")

    If Pos = 0 Then Pos = InStr(NombreYClan, "[")
    If Pos = 0 Then Pos = Len(NombreYClan) + 2
    
    charlist(charindex).nombre = Left$(NombreYClan, Pos - 2)
    charlist(charindex).clan = mid$(NombreYClan, Pos)
    
    group_index = Reader.ReadInt16()
    
    'Update char status adn tag!
    charlist(charindex).status = status
    
    charlist(charindex).group_index = group_index
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateTagAndStatus", Erl)
    

End Sub

Private Sub HandleUserOnline()
    
    On Error GoTo ErrHandler

    Dim rdata As Integer
    
    rdata = Reader.ReadInt16()
    
    usersOnline = rdata
    frmMain.onlines = "Online: " & usersOnline
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUserOnline", Erl)
    

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
     
    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
    ParticulaIndex = Reader.ReadInt16()
    Time = Reader.ReadInt32()

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
     
    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
    Color = Reader.ReadInt32()
    Rango = Reader.ReadInt8()
    
    Call Long_2_RGBA(color_value, Color)

    Dim id  As Long

    Dim id2 As Long

    If Color = 0 Then
   
        If MapData(x, y).luz.Rango > 100 Then
            LucesRedondas.Delete_Light_To_Map x, y
            Exit Sub
        Else
            id = LucesCuadradas.Light_Find(x & y)
            LucesCuadradas.Light_Remove id
            MapData(x, y).luz.Color = COLOR_EMPTY
            MapData(x, y).luz.Rango = 0
            Exit Sub

        End If

    End If
    
    MapData(x, y).luz.Color = color_value
    MapData(x, y).luz.Rango = Rango
    
    If Rango < 100 Then
        id = x & y
        LucesCuadradas.Light_Create x, y, color_value, Rango, id
    Else

        LucesRedondas.Create_Light_To_Map x, y, color_value, Rango - 99
    End If
    
    Exit Sub

HandleLightToFloor_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleLightToFloor", Erl)
    
    
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
     
    charindex = Reader.ReadInt16()
    ParticulaIndex = Reader.ReadInt16()
    Time = Reader.ReadInt32()
    Remove = Reader.ReadBool()
    grh = Reader.ReadInt32()
    
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
     
    Emisor = Reader.ReadInt16()
    receptor = Reader.ReadInt16()
    ParticulaViaje = Reader.ReadInt16()
    ParticulaFinal = Reader.ReadInt16()

    Time = Reader.ReadInt32()
    wav = Reader.ReadInt16()
    fX = Reader.ReadInt16()

    Engine_spell_Particle_Set (ParticulaViaje)

    Call Effect_Begin(ParticulaViaje, 9, Get_Pixelx_Of_Char(Emisor), Get_PixelY_Of_Char(Emisor), ParticulaFinal, Time, receptor, Emisor, wav, fX)

    ' charlist(charindex).Particula = ParticulaIndex
    ' charlist(charindex).ParticulaTime = time

    ' Call General_Char_Particle_Create(ParticulaIndex, charindex, time)
    
    Exit Sub

HandleParticleFXWithDestino_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleParticleFXWithDestino", Erl)
    
    
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
     
    Emisor = Reader.ReadInt16()
    ParticulaViaje = Reader.ReadInt16()
    ParticulaFinal = Reader.ReadInt16()

    Time = Reader.ReadInt32()
    wav = Reader.ReadInt16()
    fX = Reader.ReadInt16()
    
    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
    
    ' Debug.Print "RECIBI FX= " & fX

    Engine_spell_Particle_Set (ParticulaViaje)

    Call Effect_BeginXY(ParticulaViaje, 9, Get_Pixelx_Of_Char(Emisor), Get_PixelY_Of_Char(Emisor), x, y, ParticulaFinal, Time, Emisor, wav, fX)

    ' charlist(charindex).Particula = ParticulaIndex
    ' charlist(charindex).ParticulaTime = time

    ' Call General_Char_Particle_Create(ParticulaIndex, charindex, time)
    
    Exit Sub

HandleParticleFXWithDestinoXY_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleParticleFXWithDestinoXY", Erl)
    
    
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
     
    charindex = Reader.ReadInt16()
    ParticulaIndex = Reader.ReadString8()

    Remove = Reader.ReadBool()
    TIPO = Reader.ReadInt8()
    
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
     
    charindex = Reader.ReadInt16()
    Speeding = Reader.ReadReal32()
   
    charlist(charindex).Speeding = Speeding
    
    Exit Sub

HandleSpeedToChar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSpeedToChar", Erl)
    
    
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
    
    
End Sub

Private Sub HandleNieblaToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleNieblaToggle_Err
    
    MaxAlphaNiebla = Reader.ReadInt8()
            
    bNiebla = Not bNiebla
    frmMain.TimerNiebla.Enabled = True
    
    Exit Sub

HandleNieblaToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleNieblaToggle", Erl)
    
    
End Sub


Private Sub HandleBindKeys()
    
    On Error GoTo HandleBindKeys_Err

    '***************************************************
    'Macros
    'Pablo Mercavides
    '***************************************************
    
    ChatCombate = Reader.ReadInt8()
    ChatGlobal = Reader.ReadInt8()

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
    
    
End Sub

Private Sub HandleBarFx()
    
    On Error GoTo HandleBarFx_Err

    '***************************************************
    'Author: Pablo Mercavides
    '***************************************************
    
    Dim charindex As Integer

    Dim BarTime   As Integer

    Dim BarAccion As Byte
    
    charindex = Reader.ReadInt16()
    BarTime = Reader.ReadInt16()
    BarAccion = Reader.ReadInt8()
    
    charlist(charindex).BarTime = 0
    charlist(charindex).BarAccion = BarAccion
    charlist(charindex).MaxBarTime = BarTime
    
    Exit Sub

HandleBarFx_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBarFx", Erl)
    
    
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

    Dim ObjIndex       As Integer
    
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
    
        'Nos fijamos si se trata de una quest empezada, para poder leer los NPCs que se han matado.
        QuestEmpezada = IIf(Reader.ReadInt8, True, False)
        
        If Not QuestEmpezada Then
        
            QuestIndex = Reader.ReadInt16
        
            FrmQuestInfo.titulo.Caption = QuestList(QuestIndex).nombre
           
            'tmpStr = "Mision: " & .ReadString8 & vbCrLf
            
            LevelRequerido = Reader.ReadInt8
            QuestRequerida = Reader.ReadInt16
           
            If QuestRequerida <> 0 Then
                FrmQuestInfo.Text1.Text = ""
               Call AddtoRichTextBox(FrmQuestInfo.Text1, QuestList(QuestIndex).desc & vbCrLf & vbCrLf & "Requisitos" & vbCrLf & "Nivel requerido: " & LevelRequerido & vbCrLf & "Quest:" & QuestList(QuestRequerida).RequiredQuest, 128, 128, 128)
            Else
                
                FrmQuestInfo.Text1.Text = ""
                Call AddtoRichTextBox(FrmQuestInfo.Text1, QuestList(QuestIndex).desc & vbCrLf & vbCrLf & "Requisitos" & vbCrLf & "Nivel requerido: " & LevelRequerido & vbCrLf, 128, 128, 128)
            End If
           
            tmpByte = Reader.ReadInt8

            If tmpByte Then 'Hay NPCs
                If tmpByte > 5 Then
                    FrmQuestInfo.ListView1.FlatScrollBar = False
                Else
                    FrmQuestInfo.ListView1.FlatScrollBar = True
           
                End If

                For i = 1 To tmpByte
                    cantidadnpc = Reader.ReadInt16
                    NpcIndex = Reader.ReadInt16
               
                    ' tmpStr = tmpStr & "*) Matar " & .ReadInt16 & " " & .ReadString8 & "."
                    If QuestEmpezada Then
                        tmpStr = tmpStr & " (Has matado " & Reader.ReadInt16 & ")" & vbCrLf
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
           
            tmpByte = Reader.ReadInt8

            If tmpByte Then 'Hay OBJs

                For i = 1 To tmpByte
               
                    cantidadobj = Reader.ReadInt16
                    ObjIndex = Reader.ReadInt16
                    
                    AmountHave = Reader.ReadInt16
                   
                    Set subelemento = FrmQuestInfo.ListView1.ListItems.Add(, , ObjData(ObjIndex).Name)
                    subelemento.SubItems(1) = AmountHave & "/" & cantidadobj
                    subelemento.SubItems(2) = ObjIndex
                    subelemento.SubItems(3) = 1
                Next i

            End If
    
            tmpStr = tmpStr & vbCrLf & "RECOMPENSAS" & vbCrLf
            'tmpStr = tmpStr & "*) Oro: " & .ReadInt32 & " monedas de oro." & vbCrLf
            'tmpStr = tmpStr & "*) Experiencia: " & .ReadInt32 & " puntos de experiencia." & vbCrLf
           
            Set subelemento = FrmQuestInfo.ListView2.ListItems.Add(, , "Oro")

            subelemento.SubItems(1) = BeautifyBigNumber(Reader.ReadInt32)
            subelemento.SubItems(2) = 12
            subelemento.SubItems(3) = 0

            Set subelemento = FrmQuestInfo.ListView2.ListItems.Add(, , "Experiencia")

            subelemento.SubItems(1) = BeautifyBigNumber(Reader.ReadInt32)
            subelemento.SubItems(2) = 608
            subelemento.SubItems(3) = 1
           
            tmpByte = Reader.ReadInt8

            If tmpByte Then

                For i = 1 To tmpByte
                    'tmpStr = tmpStr & "*) " & .ReadInt16 & " " & .ReadInt16 & vbCrLf
                   
                    Dim cantidadobjs As Integer

                    Dim obindex      As Integer
                   
                    cantidadobjs = Reader.ReadInt16
                    obindex = Reader.ReadInt16
                   
                    Set subelemento = FrmQuestInfo.ListView2.ListItems.Add(, , ObjData(obindex).Name)
                       
                    subelemento.SubItems(1) = cantidadobjs
                    subelemento.SubItems(2) = obindex
                    subelemento.SubItems(3) = 1

           
                Next i

            End If

        Else
        
            QuestIndex = Reader.ReadInt16
        
            FrmQuests.titulo.Caption = QuestList(QuestIndex).nombre
           
            LevelRequerido = Reader.ReadInt8
            QuestRequerida = Reader.ReadInt16
           
            FrmQuests.detalle.Text = QuestList(QuestIndex).desc & vbCrLf & vbCrLf & "Requisitos" & vbCrLf & "Nivel requerido: " & LevelRequerido & vbCrLf

            If QuestRequerida <> 0 Then
                FrmQuests.detalle.Text = FrmQuests.detalle.Text & vbCrLf & "Quest: " & QuestList(QuestRequerida).nombre

            End If

           
            tmpStr = tmpStr & vbCrLf & "OBJETIVOS" & vbCrLf
           
            tmpByte = Reader.ReadInt8

            If tmpByte Then 'Hay NPCs

                For i = 1 To tmpByte
                    cantidadnpc = Reader.ReadInt16
                    NpcIndex = Reader.ReadInt16
               
                    Dim matados As Integer
               
                    matados = Reader.ReadInt16
                                     
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
           
            tmpByte = Reader.ReadInt8

            If tmpByte Then 'Hay OBJs

                For i = 1 To tmpByte
               
                    cantidadobj = Reader.ReadInt16
                    ObjIndex = Reader.ReadInt16
                    
                    AmountHave = Reader.ReadInt16
                   
                    Set subelemento = FrmQuests.ListView1.ListItems.Add(, , ObjData(ObjIndex).Name)
                    subelemento.SubItems(1) = AmountHave & "/" & cantidadobj
                    subelemento.SubItems(2) = ObjIndex
                    subelemento.SubItems(3) = 1
                Next i

            End If
    
            tmpStr = tmpStr & vbCrLf & "RECOMPENSAS" & vbCrLf

            Dim tmplong As Long
           
            tmplong = Reader.ReadInt32
           
            If tmplong <> 0 Then
                Set subelemento = FrmQuests.ListView2.ListItems.Add(, , "Oro")
                subelemento.SubItems(1) = BeautifyBigNumber(tmplong)
                subelemento.SubItems(2) = 12
                subelemento.SubItems(3) = 0

            End If
            
            tmplong = Reader.ReadInt32
           
            If tmplong <> 0 Then
                Set subelemento = FrmQuests.ListView2.ListItems.Add(, , "Experiencia")
                           
                subelemento.SubItems(1) = BeautifyBigNumber(tmplong)
                subelemento.SubItems(2) = 608
                subelemento.SubItems(3) = 1

            End If
           
            tmpByte = Reader.ReadInt8

            If tmpByte Then

                For i = 1 To tmpByte
                    cantidadobjs = Reader.ReadInt16
                    obindex = Reader.ReadInt16
                   
                    Set subelemento = FrmQuests.ListView2.ListItems.Add(, , ObjData(obindex).Name)
                       
                    subelemento.SubItems(1) = cantidadobjs
                    subelemento.SubItems(2) = obindex
                    subelemento.SubItems(3) = 1

           
                Next i

            End If
        
        End If

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
    tmpByte = Reader.ReadInt8
    
    'Limpiamos el ListBox y el TextBox del formulario
    FrmQuests.lstQuests.Clear
    FrmQuests.txtInfo.Text = vbNullString
        
    'Si el usuario tiene quests entonces hacemos el handle
    If tmpByte Then
        'Leemos el string
        tmpStr = Reader.ReadString8
        
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
    Dim ObjIndex       As Integer
    Dim QuestIndex     As Integer
    Dim estado         As Byte
    Dim LevelRequerido As Byte
    Dim QuestRequerida As Integer
    Dim CantidadQuest  As Byte
    Dim Repetible      As Boolean
    Dim subelemento    As ListItem
    
    FrmQuestInfo.ListView2.ListItems.Clear
    FrmQuestInfo.ListView1.ListItems.Clear

        CantidadQuest = Reader.ReadInt8
            
        For j = 1 To CantidadQuest
        
            QuestIndex = Reader.ReadInt16
            
            FrmQuestInfo.titulo.Caption = QuestList(QuestIndex).nombre
                              
            QuestList(QuestIndex).RequiredLevel = Reader.ReadInt8
            QuestList(QuestIndex).RequiredQuest = Reader.ReadInt16
            
            tmpByte = Reader.ReadInt8
    
            If tmpByte Then 'Hay NPCs
            
                If tmpByte > 5 Then
                    FrmQuestInfo.ListView1.FlatScrollBar = False
                Else
                    FrmQuestInfo.ListView1.FlatScrollBar = True
               
                End If
                    
                ReDim QuestList(QuestIndex).RequiredNPC(1 To tmpByte)
                    
                For i = 1 To tmpByte
                                                
                    QuestList(QuestIndex).RequiredNPC(i).Amount = Reader.ReadInt16
                    QuestList(QuestIndex).RequiredNPC(i).NpcIndex = Reader.ReadInt16

                Next i

            Else
                ReDim QuestList(QuestIndex).RequiredNPC(0)

            End If
               
            tmpByte = Reader.ReadInt8
    
            If tmpByte Then 'Hay OBJs
                ReDim QuestList(QuestIndex).RequiredOBJ(1 To tmpByte)
    
                For i = 1 To tmpByte
                   
                    QuestList(QuestIndex).RequiredOBJ(i).Amount = Reader.ReadInt16
                    QuestList(QuestIndex).RequiredOBJ(i).ObjIndex = Reader.ReadInt16

                Next i

            Else
                ReDim QuestList(QuestIndex).RequiredOBJ(0)
    
            End If
               
            QuestList(QuestIndex).RewardGLD = Reader.ReadInt32
            QuestList(QuestIndex).RewardEXP = Reader.ReadInt32

            tmpByte = Reader.ReadInt8
    
            If tmpByte Then
                
                ReDim QuestList(QuestIndex).RewardOBJ(1 To tmpByte)
    
                For i = 1 To tmpByte
                                              
                    QuestList(QuestIndex).RewardOBJ(i).Amount = Reader.ReadInt16
                    QuestList(QuestIndex).RewardOBJ(i).ObjIndex = Reader.ReadInt16
               
                Next i

            Else
                ReDim QuestList(QuestIndex).RewardOBJ(0)
    
            End If
                
            estado = Reader.ReadInt8
            Repetible = QuestList(QuestIndex).Repetible = 1
            
            Set subelemento = FrmQuestInfo.ListViewQuest.ListItems.Add(, , QuestList(QuestIndex).nombre & IIf(Repetible, " (R)", ""))
            subelemento.SubItems(2) = QuestIndex
  
            Select Case estado
                
                Case 0
                    subelemento.SubItems(1) = "Disponible"
                    subelemento.ForeColor = vbWhite
                    subelemento.ListSubItems(1).ForeColor = vbWhite

                Case 1
                    subelemento.SubItems(1) = "En Curso"
                    subelemento.ForeColor = RGB(255, 175, 10)
                    subelemento.ListSubItems(1).ForeColor = RGB(255, 175, 10)

                Case 2
                    If Repetible Then
                        subelemento.SubItems(1) = "Repetible"
                        subelemento.ForeColor = RGB(180, 180, 180)
                        subelemento.ListSubItems(1).ForeColor = RGB(180, 180, 180)
                    Else
                        subelemento.SubItems(1) = "Finalizada"
                        subelemento.ForeColor = RGB(15, 140, 50)
                        subelemento.ListSubItems(1).ForeColor = RGB(15, 140, 50)
                    End If

                Case 3
                    subelemento.SubItems(1) = "No disponible"
                    subelemento.ForeColor = RGB(255, 10, 10)
                    subelemento.ListSubItems(1).ForeColor = RGB(255, 10, 10)
            End Select
            
            FrmQuestInfo.ListViewQuest.Refresh
                
        Next j

    'Determinamos que formulario se muestra, segun si recibimos la informacion y la quest estï¿½ empezada o no.
    FrmQuestInfo.Show vbModeless, frmMain
    FrmQuestInfo.Picture = LoadInterface("ventananuevamision.bmp")
    Call FrmQuestInfo.ShowQuest(1)
    
    Exit Sub
    
ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleNpcQuestListSend", Erl)
    
    
End Sub

Private Sub HandleShowPregunta()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo ErrHandler
    
    Dim msg As String

    PreguntaScreen = Reader.ReadString8()
    Pregunta = True
    
    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowPregunta", Erl)
    

End Sub

Private Sub HandleDatosGrupo()
    
    On Error GoTo HandleDatosGrupo_Err
    
    Dim EnGrupo      As Boolean

    Dim CantMiembros As Byte

    Dim i            As Byte
    
    EnGrupo = Reader.ReadBool()
    
    If EnGrupo Then
        CantMiembros = Reader.ReadInt8()

        For i = 1 To CantMiembros
            FrmGrupo.lstGrupo.AddItem (Reader.ReadString8)
        Next i

    End If
    
    COLOR_AZUL = RGB(0, 0, 0)
    
    ' establece el borde al listbox
    Call Establecer_Borde(FrmGrupo.lstGrupo, FrmGrupo, COLOR_AZUL, 0, 0)

    FrmGrupo.Show , frmMain
    
    Exit Sub

HandleDatosGrupo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleDatosGrupo", Erl)
    
    
End Sub

Private Sub HandleUbicacion()
    
    On Error GoTo HandleUbicacion_Err
    
    Dim miembro As Byte
    Dim x       As Byte
    Dim y       As Byte
    Dim map     As Integer
    
    miembro = Reader.ReadInt8()
    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
    map = Reader.ReadInt16()
    
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
    
    
End Sub

Private Sub HandleViajarForm()
    
    On Error GoTo HandleViajarForm_Err
            
    Dim Dest     As String
    Dim DestCant As Byte
    Dim i        As Byte
    Dim tempdest As String

    FrmViajes.List1.Clear
    
    DestCant = Reader.ReadInt8()
        
    ReDim Destinos(1 To DestCant) As Tdestino
        
    For i = 1 To DestCant
        
        tempdest = Reader.ReadString8()
        
        Destinos(i).CityDest = ReadField(1, tempdest, Asc("-"))
        Destinos(i).costo = ReadField(2, tempdest, Asc("-"))
        FrmViajes.List1.AddItem ListaCiudades(Destinos(i).CityDest) & " - " & Destinos(i).costo & " monedas"

    Next i
        
    Call Establecer_Borde(FrmViajes.List1, FrmViajes, COLOR_AZUL, 0, 0)
         
    ViajarInterface = Reader.ReadInt8()
        
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
    
    
End Sub

Private Sub HandleSeguroResu()
    
    'Get data and update form
    SeguroResuX = Reader.ReadBool()
    
    If SeguroResuX Then
        Call AddtoRichTextBox(frmMain.RecTxt, "Seguro de resurrección activado.", 65, 190, 156, False, False, False)
        frmMain.ImgSegResu = LoadInterface("boton-fantasma-on.bmp")
    Else
        Call AddtoRichTextBox(frmMain.RecTxt, "Seguro de resurrección desactivado.", 65, 190, 156, False, False, False)
        frmMain.ImgSegResu = LoadInterface("boton-fantasma-off.bmp")

    End If
    
End Sub

Private Sub HandleStopped()

    UserStopped = Reader.ReadBool()

End Sub

Private Sub HandleInvasionInfo()

    InvasionActual = Reader.ReadInt8
    InvasionPorcentajeVida = Reader.ReadInt8
    InvasionPorcentajeTiempo = Reader.ReadInt8
    
    frmMain.Evento.Enabled = False
    frmMain.Evento.Interval = 0
    frmMain.Evento.Interval = 10000
    frmMain.Evento.Enabled = True

End Sub

Private Sub HandleCommerceRecieveChatMessage()
    
    Dim Message As String
    Message = Reader.ReadString8
        
    Call AddtoRichTextBox(frmComerciarUsu.RecTxt, Message, 255, 255, 255, 0, False, True, False)
    
End Sub

Private Sub HandleDoAnimation()
    
    On Error GoTo HandleCharacterChange_Err
    
    Dim charindex As Integer

    Dim tempint   As Integer

    Dim headIndex As Integer

    charindex = Reader.ReadInt16()
    
    With charlist(charindex)
        .AnimatingBody = Reader.ReadInt16()
        .Body = BodyData(.AnimatingBody)
        'Start animation
        .Body.Walk(.Heading).Started = FrameTime
        .Body.Walk(.Heading).Loops = 0
    End With
    
    Exit Sub

HandleCharacterChange_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleDoAnimation", Erl)
    
    
End Sub

Private Sub HandleOpenCrafting()

    Dim TIPO As Byte
    TIPO = Reader.ReadInt8

    frmCrafteo.Picture = LoadInterface(TipoCrafteo(TIPO).Ventana)
    frmCrafteo.InventoryGrhIndex = TipoCrafteo(TIPO).Inventario
    frmCrafteo.TipoGrhIndex = TipoCrafteo(TIPO).Icono
    
    Dim i As Long
    'Fill our inventory list
    For i = 1 To MAX_INVENTORY_SLOTS
        With frmMain.Inventario
            Call frmCrafteo.InvCraftUser.SetItem(i, .ObjIndex(i), .Amount(i), .Equipped(i), .GrhIndex(i), .ObjType(i), .MaxHit(i), .MinHit(i), .Def(i), .Valor(i), .ItemName(i), .PuedeUsar(i))
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
    Dim Slot As Byte, ObjIndex As Integer
    Slot = Reader.ReadInt8
    ObjIndex = Reader.ReadInt16
    
    If ObjIndex <> 0 Then
        With ObjData(ObjIndex)
            Call frmCrafteo.InvCraftItems.SetItem(Slot, ObjIndex, 1, 0, .GrhIndex, .ObjType, 0, 0, 0, .Valor, .Name, 0)
        End With
    Else
        Call frmCrafteo.InvCraftItems.ClearSlot(Slot)
    End If
    
End Sub

Private Sub HandleCraftingCatalyst()
    Dim ObjIndex As Integer, Amount As Integer, Porcentaje As Byte
    ObjIndex = Reader.ReadInt16
    Amount = Reader.ReadInt16
    Porcentaje = Reader.ReadInt8
    
    If ObjIndex <> 0 Then
        With ObjData(ObjIndex)
            Call frmCrafteo.InvCraftCatalyst.SetItem(1, ObjIndex, Amount, 0, .GrhIndex, .ObjType, 0, 0, 0, .Valor, .Name, 0)
        End With
    Else
        Call frmCrafteo.InvCraftCatalyst.ClearSlot(1)
    End If

    frmCrafteo.PorcentajeAcierto = Porcentaje
    
End Sub

Private Sub HandleCraftingResult()
    Dim ObjIndex As Integer
    ObjIndex = Reader.ReadInt16

    If ObjIndex > 0 Then
        Dim Porcentaje As Byte, Precio As Long
        Porcentaje = Reader.ReadInt8
        Precio = Reader.ReadInt32
        Call frmCrafteo.SetResult(ObjData(ObjIndex).GrhIndex, Porcentaje, Precio)
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
    
End Sub

Public Sub HandleAnswerReset()
    On Error GoTo ErrHandler

    If MsgBox("¿Está seguro que desea resetear el personaje? Los items que no sean depositados se perderán.", vbYesNo, "Resetear personaje") = vbYes Then
        Call WriteResetearPersonaje
    End If

    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleAnswerReset", Erl)
End Sub
Public Sub HandleUpdateBankGld()

    On Error GoTo ErrHandler
    
    Dim UserBoveOro As Long
        
    UserBoveOro = Reader.ReadInt32
    
    Call frmGoliath.UpdateBankGld(UserBoveOro)
    Exit Sub
ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateBankGld", Erl)

End Sub

Public Sub HandlePelearConPezEspecial()
    On Error GoTo ErrHandler
    
    PosicionBarra = 1
    DireccionBarra = 1
    Dim i As Integer
    
    For i = 1 To MAX_INTENTOS
        intentosPesca(i) = 0
    Next i
    PescandoEspecial = True
    ContadorIntentosPescaEspecial_Fallados = 0
    ContadorIntentosPescaEspecial_Acertados = 0
    startTimePezEspecial = GetTickCount()
    Call Char_Dialog_Set(UserCharIndex, "Oh! Creo que tengo un super pez en mi linea, intentare obtenerlo con la letra P", &H1FFFF, 200, 130)
    Exit Sub
ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePelearConPezEspecial", Erl)
End Sub

Public Sub HandlePrivilegios()
    On Error GoTo errhandler
    
    EsGM = Reader.ReadBool
    If EsGM Then
        frmMain.panelGM.Visible = True
        frmMain.createObj.Visible = True
        frmMain.btnInvisible.Visible = True
        frmMain.btnSpawn.Visible = True
    Else
        frmMain.panelGM.Visible = False
        frmMain.createObj.Visible = False
        frmMain.btnInvisible.Visible = False
        frmMain.btnSpawn.Visible = False
    End If
    Exit Sub
errhandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePrivilegios", Erl)
End Sub

Public Sub HandleShopInit()
    
    Dim cant_obj_shop As Long, i As Long
    
    cant_obj_shop = Reader.ReadInt16
    
    credits_shopAO20 = Reader.ReadInt32
    frmShopAO20.lblCredits.Caption = credits_shopAO20
    
    ReDim ObjShop(1 To cant_obj_shop) As ObjDatas
    
    For i = 1 To cant_obj_shop
        ObjShop(i).objNum = Reader.ReadInt32
        ObjShop(i).Valor = Reader.ReadInt32
        ObjShop(i).Name = Reader.ReadString8
         
        Call frmShopAO20.lstItemShopFilter.AddItem(ObjShop(i).Name & " (Valor: " & ObjShop(i).Valor & ")", i - 1)
    Next i
    
    frmShopAO20.Show , frmMain
End Sub

Public Sub HandleUpdateShopClienteCredits()
    credits_shopAO20 = Reader.ReadInt32
    frmShopAO20.lblCredits.Caption = credits_shopAO20
End Sub

Public Sub HandleSensuiRetrasado()
    EscribeRetrasadoSensui = True
    frmMain.timerRetrasadoSensui.Enabled = True
    
End Sub

Public Sub HandleObjQuestListSend()

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Recibe y maneja el paquete QuestListSend del servidor.
    'Last modified: 29/08/2021 by HarThaoS
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
    Dim QuestIndex     As Integer
    Dim estado         As Byte
    Dim LevelRequerido As Byte
    Dim QuestRequerida As Integer
    Dim CantidadQuest  As Byte
    Dim Repetible      As Boolean
    Dim subelemento    As ListItem

    FrmQuestInfo.ListView2.ListItems.Clear
    FrmQuestInfo.ListView1.ListItems.Clear


    QuestIndex = Reader.ReadInt16

    FrmQuestInfo.titulo.Caption = QuestList(QuestIndex).nombre

    QuestList(QuestIndex).RequiredLevel = Reader.ReadInt8
    QuestList(QuestIndex).RequiredQuest = Reader.ReadInt16


    tmpByte = Reader.ReadInt8

    If tmpByte Then 'Hay NPCs

        If tmpByte > 5 Then
            FrmQuestInfo.ListView1.FlatScrollBar = False
        Else
            FrmQuestInfo.ListView1.FlatScrollBar = True

        End If

        ReDim QuestList(QuestIndex).RequiredNPC(1 To tmpByte)

        For i = 1 To tmpByte

            QuestList(QuestIndex).RequiredNPC(i).Amount = Reader.ReadInt16
            QuestList(QuestIndex).RequiredNPC(i).NpcIndex = Reader.ReadInt16

        Next i

    Else
        ReDim QuestList(QuestIndex).RequiredNPC(0)

    End If

    tmpByte = Reader.ReadInt8

    If tmpByte Then 'Hay OBJs
        ReDim QuestList(QuestIndex).RequiredOBJ(1 To tmpByte)

        For i = 1 To tmpByte

            QuestList(QuestIndex).RequiredOBJ(i).Amount = Reader.ReadInt16
            QuestList(QuestIndex).RequiredOBJ(i).OBJIndex = Reader.ReadInt16

        Next i

    Else
        ReDim QuestList(QuestIndex).RequiredOBJ(0)

    End If

    QuestList(QuestIndex).RewardGLD = Reader.ReadInt32
    QuestList(QuestIndex).RewardEXP = Reader.ReadInt32

    tmpByte = Reader.ReadInt8

    If tmpByte Then

        ReDim QuestList(QuestIndex).RewardOBJ(1 To tmpByte)

        For i = 1 To tmpByte

            QuestList(QuestIndex).RewardOBJ(i).Amount = Reader.ReadInt16
            QuestList(QuestIndex).RewardOBJ(i).OBJIndex = Reader.ReadInt16

        Next i

    Else
        ReDim QuestList(QuestIndex).RewardOBJ(0)

    End If

    estado = Reader.ReadInt8
    Repetible = QuestList(QuestIndex).Repetible = 1

    Set subelemento = FrmQuestInfo.ListViewQuest.ListItems.Add(, , QuestList(QuestIndex).nombre & IIf(Repetible, " (R)", ""))
    subelemento.SubItems(2) = QuestIndex

    Select Case estado

        Case 0
            subelemento.SubItems(1) = "Disponible"
            subelemento.ForeColor = vbWhite
            subelemento.ListSubItems(1).ForeColor = vbWhite

        Case 1
            subelemento.SubItems(1) = "En Curso"
            subelemento.ForeColor = RGB(255, 175, 10)
            subelemento.ListSubItems(1).ForeColor = RGB(255, 175, 10)

        Case 2
            If Repetible Then
                subelemento.SubItems(1) = "Repetible"
                subelemento.ForeColor = RGB(180, 180, 180)
                subelemento.ListSubItems(1).ForeColor = RGB(180, 180, 180)
            Else
                subelemento.SubItems(1) = "Finalizada"
                subelemento.ForeColor = RGB(15, 140, 50)
                subelemento.ListSubItems(1).ForeColor = RGB(15, 140, 50)
            End If

        Case 3
            subelemento.SubItems(1) = "No disponible"
            subelemento.ForeColor = RGB(255, 10, 10)
            subelemento.ListSubItems(1).ForeColor = RGB(255, 10, 10)
    End Select

    FrmQuestInfo.ListViewQuest.Refresh

    'Determinamos que formulario se muestra, segun si recibimos la informacion y la quest estï¿½ empezada o no.
    FrmQuestInfo.Show vbModeless, frmMain
    FrmQuestInfo.Picture = LoadInterface("ventananuevamision.bmp")
    Call FrmQuestInfo.ShowQuest(1)

    Exit Sub

ErrHandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleNpcQuestListSend", Erl)


End Sub
