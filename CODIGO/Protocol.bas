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

Private Enum ServerPacketID

    logged                  ' LOGGED  0
    RemoveDialogs           ' QTDL
    RemoveCharDialog        ' QDL
    NavigateToggle          ' NAVEG
    EquiteToggle
    CreateRenderText
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
    EfectOverHEad '120
    EfectToScreen
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
    ExpOverHEad
    OroOverHEad
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
    UpdateNPCSimbolo
    ClanSeguro
    Intervals
    UpdateUserKey
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
    invisible               '/INVISIBLE
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
    ToggleCentinelActivated '/CENTINELAACTIVADO
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
End Enum

Private Enum NewPacksID

    OfertaInicial
    OfertaDeSubasta
    QuestionGM
    CuentaRegresiva
    PossUser
    Duelo
    NieveToggle
    NieblaToggle
    TransFerGold
    MoveItem
    Genio                 '/GENIO
    Casarse
    CraftAlquimista
    DropItem
    RequestFamiliar
    FlagTrabajar
    CraftSastre
    MensajeUser
    TraerBoveda
    CompletarAccion
    Escribiendo
    TraerRecompensas
    ReclamarRecompensa
    DecimeLaHora
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
    ScrollInfo
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

End Enum

''
' Handles incoming data.

Public Sub HandleIncomingData()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    On Error Resume Next

    Dim paquete As Byte

    paquete = incomingData.PeekByte()

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
            
        Case ServerPacketID.CreateRenderText       ' CDMG ' GSZAO
            Call HandleCreateRenderValue
            
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
            
        Case ServerPacketID.EfectOverHEad
            Call HandleEfectOverHead
            
        Case ServerPacketID.ExpOverHEad
            Call HandleExpOverHead
            
        Case ServerPacketID.OroOverHEad
            Call HandleOroOverHead
        
        Case ServerPacketID.EfectToScreen
            Call HandleEfectToScreen
            
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

        Case Else
        
            Exit Sub

    End Select
    
    'Done with this packet, move on to next one
    If incomingData.length > 0 And Err.number <> incomingData.NotEnoughDataErrCode Then
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
    Call incomingData.ReadByte
    
    '#If AntiExternos Then
    'Security.Redundance = incomingData.ReadByte()
    '#End If
    
    ' Variable initialization
    UserCiego = False
    EngineRun = True
    UserDescansar = False
    Nombres = True
    Pregunta = False

    frmmain.stabar.Visible = True
    frmmain.HpBar.Visible = True

    If UserMaxMAN <> 0 Then
        frmmain.manabar.Visible = True
    End If

    frmmain.hambar.Visible = True
    frmmain.AGUbar.Visible = True
    frmmain.Hpshp.Visible = (UserMinHp > 0)
    frmmain.MANShp.Visible = (UserMinMAN > 0)
    frmmain.STAShp.Visible = (UserMinSTA > 0)
    frmmain.AGUAsp.Visible = (UserMinAGU > 0)
    frmmain.COMIDAsp.Visible = (UserMinHAM > 0)
    frmmain.GldLbl.Visible = True
    ' frmMain.Label6.Visible = True
    frmmain.Fuerzalbl.Visible = True
    frmmain.AgilidadLbl.Visible = True
    frmmain.oxigenolbl.Visible = True
    QueRender = 0
    'Set connected state
    
    Call SetConnected
    
    'Show tip
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
    Call incomingData.ReadByte
    
    Call Dialogos.RemoveAllDialogs

End Sub

''
' Handles the RemoveCharDialog message.

Private Sub HandleRemoveCharDialog()

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
    Call incomingData.ReadByte
    
    UserNavegando = Not UserNavegando

End Sub

Private Sub HandleNadarToggle()

    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    '
    UserNadando = incomingData.ReadBoolean()

End Sub

Private Sub HandleCreateRenderValue()

    '***************************************************
    'Author: maTih.-
    'Last Modification: 09/06/2012 - ^[GS]^
    '
    '***************************************************
    With incomingData
        .ReadByte
        Call modRenderValue.Create(.ReadByte(), .ReadByte(), 0, .ReadDouble(), .ReadByte())

    End With

End Sub

Private Sub HandleEquiteToggle()
    'Remove packet ID
    Call incomingData.ReadByte
    UserMontado = Not UserMontado

    'If UserMontado Then
    '    charlist(UserCharIndex).Speeding = 1.3
    ' Else
    '    charlist(UserCharIndex).Speeding = 1.1
    ' End If
End Sub

Private Sub HandleVelocidadToggle()

    If incomingData.length < 2 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    '
    charlist(UserCharIndex).Speeding = incomingData.ReadSingle()

End Sub

Private Sub HandleMacroTrabajoToggle()
    'Activa o Desactiva el macro de trabajo  06/07/2014 Ladder

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
        AddtoRichTextBox frmmain.RecTxt, "Has comenzado a trabajar...", 2, 223, 51, 1, 0
        frmmain.MacroLadder.Interval = IntervaloTrabajo
        frmmain.MacroLadder.Enabled = True
        UserMacro.Intervalo = IntervaloTrabajo
        UserMacro.Activado = True
        UserMacro.cantidad = 999
        UserMacro.TIPO = 6
        
        TargetXMacro = tX
        TargetYMacro = tY

    End If

End Sub

''
' Handles the Disconnect message.

Private Sub HandleDisconnect()

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
        frmmain.Socket1.Disconnect
    #Else

        If frmmain.Winsock1.State <> sckClosed Then frmmain.Winsock1.Close
    #End If
    
    'Hide main form
    'FrmCuenta.Visible = True
    
    frmConnect.Visible = True
    QueRender = 2

    Call Graficos_Particulas.Particle_Group_Remove_All
    Call Graficos_Particulas.Engine_Select_Particle_Set(203)
    
    ParticleLluviaDorada = General_Particle_Create(208, -1, -1)

    frmmain.hlst.Visible = False
    frmmain.Timerping.Enabled = False
    
    frmmain.Visible = False
    
    OpcionMenu = 0
    ' Panel.Picture = LoadInterface("centroinventario.bmp")
    ' frmMain.Image2(0).Visible = False
    'frmMain.Image2(1).Visible = True

    frmmain.picInv.Visible = True
    frmmain.hlst.Visible = False

    frmmain.cmdlanzar.Visible = False
    'frmMain.lblrefuerzolanzar.Visible = False
    frmmain.cmdMoverHechi(0).Visible = False
    frmmain.cmdMoverHechi(1).Visible = False
    
    frmmain.PicResu.Visible = True
    frmmain.PicResuOn.Visible = False
     
    frmmain.PicSegClanOn.Visible = True
    
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
    
    frmmain.HoraFantasiaTimer.Enabled = False
    Call SwitchMapIAO(UserMap)
    
    frmmain.personaje(1).Visible = False
    frmmain.personaje(2).Visible = False
    frmmain.personaje(3).Visible = False
    frmmain.personaje(4).Visible = False
    frmmain.personaje(5).Visible = False
    
    UserClase = 0
    UserSexo = 0
    UserRaza = 0
    MiCabeza = 0
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i
    
    For i = 1 To UserInvUnlocked
        frmmain.imgInvLock(i - 1).Picture = Nothing
    Next i
    
    For i = 1 To MAX_INVENTORY_SLOTS
        Call frmmain.Inventario.SetItem(i, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0)
        Call frmBancoObj.InvBankUsu.SetItem(i, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0)
        Call frmComerciar.InvComNpc.SetItem(i, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0)
        Call frmComerciar.InvComUsu.SetItem(i, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0)
    Next i
    
    For i = 1 To MAX_BANCOINVENTORY_SLOTS
        Call frmBancoObj.InvBoveda.SetItem(i, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0)
    Next i
    
    UserInvUnlocked = 0

    Alocados = 0

    'Reset global vars
    UserParalizado = False
    UserSaliendo = False
    UserInmovilizado = False
    pausa = False
    UserMeditar = False
    UserDescansar = False
    UserNavegando = False
    UserMontado = False
    UserNadando = False
    bRain = False
    AlphaNiebla = 30
    frmmain.TimerNiebla.Enabled = False
    bNiebla = False
    MostrarTrofeo = False
    bNieve = False
    bFogata = False
    SkillPoints = 0
    meteo_estado = 0
    UserEstado = 0
    
    InviCounter = 0
    ScrollExpCounter = 0
    ScrollOroCounter = 0
    DrogaCounter = 0
    OxigenoCounter = 0
     
    frmmain.Contadores.Enabled = False
     
    'Delete all kind of dialogs
    
    'Reset some char variables...
    For i = 1 To LastChar + 1
        charlist(i).invisible = False
        charlist(i).Arma_Aura = ""
        charlist(i).Body_Aura = ""
        charlist(i).Escudo_Aura = ""
        charlist(i).Otra_Aura = ""
        charlist(i).Head_Aura = ""
        charlist(i).Speeding = 0
        charlist(i).AuraAngle = 0
    Next i

    For i = 1 To LastChar + 1
        charlist(i).dialog = ""
    Next i
        
    'Unload all forms except frmMain and frmConnect
    Dim frm As Form
    
    For Each frm In Forms

        If frm.name <> frmmain.name And frm.name <> frmConnect.name And frm.name <> frmMensaje.name Then
            Unload frm

        End If

    Next
    
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
    Call incomingData.ReadByte

    'Reset vars
    Comerciando = False
    
    'Hide form
    ' Unload frmComerciar
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
    Call incomingData.ReadByte
    
    ' frmBancoObj.List1(0).Clear
    ' frmBancoObj.List1(1).Clear

    'Unload frmBancoObj
    Comerciando = False

End Sub

''
' Handles the CommerceInit message.

Private Sub HandleCommerceInit()

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

        With frmmain.Inventario
            Call frmComerciar.InvComUsu.SetItem(i, .OBJIndex(i), .Amount(i), .Equipped(i), .GrhIndex(i), .ObjType(i), .MaxHit(i), .MinHit(i), .Def(i), .Valor(i), .ItemName(i), .PuedeUsar(i))

        End With

    Next i

    'Set state and show form
    Comerciando = True
    'Call Inventario.Initialize(frmComerciar.PicInvUser)
    frmComerciar.Picture = LoadInterface("comerciar.bmp")
    HayFormularioAbierto = True
    frmComerciar.Show , frmmain
    
End Sub

''
' Handles the BankInit message.

Private Sub HandleBankInit()

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

        With frmmain.Inventario
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
    frmBancoObj.Show , frmmain
    frmBancoObj.lblCosto = PonerPuntos(UserGLD)
    HayFormularioAbierto = True

End Sub

Private Sub HandleGoliathInit()

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
    
End Sub

Private Sub HandleShowFrmLogear()

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

End Sub

Private Sub HandleShowFrmMapa()

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
    HayFormularioAbierto = True
    frmMapaGrande.Show , frmmain

End Sub

''
' Handles the UserCommerceInit message.

Private Sub HandleUserCommerceInit()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim i As Long
    
    'Remove packet ID
    Call incomingData.ReadByte
    
    'Clears lists if necessary
    If frmComerciarUsu.List1.ListCount > 0 Then frmComerciarUsu.List1.Clear
    If frmComerciarUsu.List2.ListCount > 0 Then frmComerciarUsu.List2.Clear
    
    'Fill inventory list
    For i = 1 To MAX_INVENTORY_SLOTS

        If frmmain.Inventario.OBJIndex(i) <> 0 Then
            frmComerciarUsu.List1.AddItem frmmain.Inventario.ItemName(i)
            frmComerciarUsu.List1.ItemData(frmComerciarUsu.List1.NewIndex) = frmmain.Inventario.Amount(i)
        Else
            frmComerciarUsu.List1.AddItem ""
            frmComerciarUsu.List1.ItemData(frmComerciarUsu.List1.NewIndex) = 0

        End If

    Next i
    
    'Set state and show form
    Comerciando = True
    
    COLOR_AZUL = RGB(0, 0, 0)
    Call Establecer_Borde(frmComerciarUsu.List1, frmComerciarUsu, COLOR_AZUL, 0, 0)
    Call Establecer_Borde(frmComerciarUsu.List2, frmComerciarUsu, COLOR_AZUL, 0, 0)
    
    frmComerciarUsu.Picture = LoadInterface("comercioseguro.bmp")
    frmComerciarUsu.Image1.Picture = LoadInterface("comercioseguro_opbjeto.bmp")
    frmComerciarUsu.Show , frmmain

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
    Call incomingData.ReadByte
    
    'Clear the lists
    frmComerciarUsu.List1.Clear
    frmComerciarUsu.List2.Clear
    
    'Destroy the form and reset the state
    Unload frmComerciarUsu
    Comerciando = False

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
    Call incomingData.ReadByte
    
    If frmmain.macrotrabajo.Enabled And (MacroBltIndex > 0) Then
        Call WriteCraftBlacksmith(MacroBltIndex)
    Else
        
        frmHerrero.Picture = LoadInterface("herreria.bmp")
    
        frmHerrero.lstArmas.Clear

        Dim i As Byte

        For i = 0 To UBound(CascosHerrero())

            If CascosHerrero(i).Index = 0 Then Exit For
            Call frmHerrero.lstArmas.AddItem(ObjData(CascosHerrero(i).Index).name)
        Next i

        frmHerrero.Command3.Picture = LoadInterface("herreria_cascoshover.bmp")
    
        COLOR_AZUL = RGB(0, 0, 0)
        Call Establecer_Borde(frmHerrero.lstArmas, frmHerrero, COLOR_AZUL, 1, 1)
        Call Establecer_Borde(frmHerrero.List1, frmHerrero, COLOR_AZUL, 1, 1)
        Call Establecer_Borde(frmHerrero.List2, frmHerrero, COLOR_AZUL, 1, 1)
        HayFormularioAbierto = True
        frmHerrero.Show , frmmain

    End If

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
    Call incomingData.ReadByte
    
    If frmmain.macrotrabajo.Enabled And (MacroBltIndex > 0) Then
        Call WriteCraftCarpenter(MacroBltIndex)
    Else
         
        COLOR_AZUL = RGB(0, 0, 0)
    
        ' establece el borde al listbox
        Call Establecer_Borde(frmCarp.lstArmas, frmCarp, COLOR_AZUL, 0, 0)
        Call Establecer_Borde(frmCarp.List1, frmCarp, COLOR_AZUL, 0, 0)
        Call Establecer_Borde(frmCarp.List2, frmCarp, COLOR_AZUL, 0, 0)
        frmCarp.Picture = LoadInterface("carpinteria.bmp")
        frmCarp.Show , frmmain
        HayFormularioAbierto = True

    End If

End Sub

Private Sub HandleShowAlquimiaForm()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte

    If frmmain.macrotrabajo.Enabled And (MacroBltIndex > 0) Then
        Call WriteCraftAlquimista(MacroBltIndex)
    Else
        frmAlqui.Picture = LoadInterface("alquimia.bmp")
    
        COLOR_AZUL = RGB(0, 0, 0)
        
        ' establece el borde al listbox
        Call Establecer_Borde(frmAlqui.lstArmas, frmAlqui, COLOR_AZUL, 1, 1)
        Call Establecer_Borde(frmAlqui.List1, frmAlqui, COLOR_AZUL, 1, 1)
        Call Establecer_Borde(frmAlqui.List2, frmAlqui, COLOR_AZUL, 1, 1)

        frmAlqui.Show , frmmain
        HayFormularioAbierto = True

    End If

End Sub

Private Sub HandleShowSastreForm()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    If frmmain.macrotrabajo.Enabled And (MacroBltIndex > 0) Then
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
            FrmSastre.lstArmas.AddItem (ObjData(SastreRopas(i).Index).name)
        Next i
    
        FrmSastre.Command1.Picture = LoadInterface("sastreria_vestimentahover.bmp")
        FrmSastre.Show , frmmain
        HayFormularioAbierto = True

    End If

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
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmmain.RecTxt, MENSAJE_CRIATURA_MATADO, 255, 0, 0, True, False, False)

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
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmmain.RecTxt, MENSAJE_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)

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
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmmain.RecTxt, MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)

End Sub

''
' Handles the UserSwing message.

Private Sub HandleCharSwing()

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
            .dialogEfec = IIf(charindex = UserCharIndex, "Fallas", "Falló")
            .SubeEfecto = 100
            .dialog_Efect_color.r = 255
            .dialog_Efect_color.g = 0
            .dialog_Efect_color.b = 0
            .dialog_Efect_color.a = 255

        End If
        
        Call Sound.Sound_Play(2, False, Sound.Calculate_Volume(.Pos.x, .Pos.y), Sound.Calculate_Pan(.Pos.x, .Pos.y)) ' Swing
        
        If ShowFX Then Call SetCharacterFx(charindex, 90, 0)

    End With
    
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
    Call incomingData.ReadByte
    
    Call frmmain.DibujarSeguro
    Call AddtoRichTextBox(frmmain.RecTxt, MENSAJE_SEGURO_ACTIVADO, 65, 190, 156, False, False, False)

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
    Call incomingData.ReadByte
    
    Call frmmain.DesDibujarSeguro
    Call AddtoRichTextBox(frmmain.RecTxt, MENSAJE_SEGURO_DESACTIVADO, 65, 190, 156, False, False, False)

End Sub

''
' Handles the ResuscitationSafeOff message.

Private Sub HandlePartySafeOff()
    '***************************************************
    'Author: Rapsodius
    'Creation date: 10/10/07
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    Call frmmain.ControlSeguroParty(False)
    Call AddtoRichTextBox(frmmain.RecTxt, MENSAJE_SEGURO_PARTY_OFF, 250, 250, 0, False, True, False)

End Sub

Private Sub HandleClanSeguro()

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
    
    If Seguro Then
        Call AddtoRichTextBox(frmmain.RecTxt, "Seguro de clan desactivado.", 65, 190, 156, False, False, False)
        frmmain.PicSegClanOn.Visible = False
    Else
        Call AddtoRichTextBox(frmmain.RecTxt, "Seguro de clan activado.", 65, 190, 156, False, False, False)
        frmmain.PicSegClanOn.Visible = True
    
    End If

End Sub

Private Sub HandleIntervals()

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
    IntervaloTrabajo = incomingData.ReadLong()
    IntervaloUsarU = incomingData.ReadLong()
    IntervaloUsarClic = incomingData.ReadLong()
    IntervaloTirar = incomingData.ReadLong()
    
    'Set the intervals of timers
    Call MainTimer.SetInterval(TimersIndex.Attack, IntervaloGolpe)
    Call MainTimer.SetInterval(TimersIndex.Work, IntervaloTrabajo)
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
    
    frmmain.macrotrabajo.Interval = IntervaloTrabajo
    frmmain.macrotrabajo.Enabled = False

    'Init timers
    Call MainTimer.Start(TimersIndex.Attack)
    Call MainTimer.Start(TimersIndex.Work)
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

End Sub

Private Sub HandleUpdateUserKey()
    If incomingData.length < 5 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub
    End If

    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim Slot As Integer, Llave As Integer
    
    Slot = incomingData.ReadInteger
    Llave = incomingData.ReadInteger

    Call FrmKeyInv.InvKeys.SetItem(Slot, Llave, 1, 0, ObjData(Llave).GrhIndex, eObjType.otLlaves, 0, 0, 0, 0, ObjData(Llave).name, 0)

End Sub

' Handles the ResuscitationSafeOn message.
Private Sub HandlePartySafeOn()
    '***************************************************
    'Author: Rapsodius
    'Creation date: 10/10/07
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    Call frmmain.ControlSeguroParty(True)
    Call AddtoRichTextBox(frmmain.RecTxt, MENSAJE_SEGURO_PARTY_ON, 250, 250, 0, False, True, False)

End Sub

Private Sub HandleCorreoPicOn()
    '***************************************************
    'Author: Rapsodius
    'Creation date: 10/10/07
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    frmmain.PicCorreo.Visible = True

    'Call AddtoRichTextBox(frmMain.RecTxt, "Tenes un nuevo correo.", 204, 193, 115, False, False, False)
    'Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_RESU_ON, 65, 190, 156, False, False, False)
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
    Call incomingData.ReadByte
    
    Call AddtoRichTextBox(frmmain.RecTxt, MENSAJE_USAR_MEDITANDO, 255, 0, 0, False, False, False)

End Sub

''
' Handles the UpdateSta message.

Private Sub HandleUpdateSta()

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
    frmmain.STAShp.Width = UserMinSTA / UserMaxSTA * 89
    frmmain.stabar.Caption = UserMinSTA & " / " & UserMaxSTA
    frmmain.STAShp.Visible = (UserMinSTA > 0)

End Sub

''
' Handles the UpdateMana message.

Private Sub HandleUpdateMana()

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
    UserMinMAN = incomingData.ReadInteger()
    
    If UserMaxMAN > 0 Then
        frmmain.MANShp.Width = UserMinMAN / UserMaxMAN * 216
        frmmain.manabar.Caption = UserMinMAN & " / " & UserMaxMAN
        frmmain.MANShp.Visible = (UserMinMAN > 0)
    Else
        frmmain.MANShp.Width = 0
        frmmain.manabar.Visible = False
        frmmain.MANShp.Visible = False
    End If

End Sub

''
' Handles the UpdateHP message.

Private Sub HandleUpdateHP()

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
    UserMinHp = incomingData.ReadInteger()
    frmmain.Hpshp.Width = UserMinHp / UserMaxHp * 216
    frmmain.HpBar.Caption = UserMinHp & " / " & UserMaxHp
    frmmain.Hpshp.Visible = (UserMinHp > 0)
    
    'Velocidad de la musica
    
    'Is the user alive??
    If UserMinHp = 0 Then
        UserEstado = 1
        meteo_estado = 0
        Meteo_Engine.SetNuevoEstado 5
    Else
    
        ' Dim Rojo As Byte, Verde As Byte, Azul As Byte

        ' If MapDat.base_light = 16777215 Then
        '    Map_light_base = D3DColorARGB(255, 255, 255, 255)
        '     ColorAmbiente.r = 255
        '     ColorAmbiente.b = 255
        '     ColorAmbiente.g = 255
        '    ColorAmbiente.a = 255
        '    Call Map_Base_Light_Set(map_base_light)
        ' Else
        ' Call Obtener_RGB(MapDat.base_light, Rojo, Verde, Azul)
        ' ColorAmbiente.r = Rojo
        ' ColorAmbiente.b = Azul
        '  ColorAmbiente.g = Verde
        '  ColorAmbiente.a = 255
        '  Map_light_base = D3DColorARGB(255, ColorAmbiente.r, ColorAmbiente.g, ColorAmbiente.b)
        'Call Map_Base_Light_Set(Map_light_base)
        'End If
        UserEstado = 0

    End If

End Sub

''
' Handles the UpdateGold message.

Private Sub HandleUpdateGold()

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
    
    frmmain.GldLbl.Caption = PonerPuntos(UserGLD)

End Sub

''
' Handles the UpdateExp message.

Private Sub HandleUpdateExp()

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

    frmmain.exp.Caption = PonerPuntos(UserExp) & "/" & PonerPuntos(UserPasarNivel)
    If UserPasarNivel > 0 Then
        frmmain.EXPBAR.Width = UserExp / UserPasarNivel * 204
        frmmain.lblPorcLvl.Caption = Round(UserExp * 100 / UserPasarNivel, 0) & "%"
    Else
        frmmain.EXPBAR.Width = 204
        frmmain.lblPorcLvl.Caption = "¡Nivel máximo!"
    End If

End Sub

''
' Handles the ChangeMap message.

Private Sub HandleChangeMap()

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
            frmmain.IsPlaying = PlayLoop.plNone

        End If

    End If
        
    If HayFormularioAbierto Then
        If frmComerciar.Visible Then
            Unload frmComerciar
            HayFormularioAbierto = False

            ' Exit Sub
        End If
        
        If frmBancoObj.Visible Then
            Unload frmBancoObj
            HayFormularioAbierto = False

            ' Exit Sub
        End If
        
        If FrmShop.Visible Then
            Unload FrmShop
            HayFormularioAbierto = False

            'Exit Sub
        End If
            
        If frmEstadisticas.Visible Then
            Unload frmEstadisticas
            HayFormularioAbierto = False

            ' Exit Sub
        End If
        
        If frmHerrero.Visible Then
            Unload frmHerrero
            HayFormularioAbierto = False

            ' Exit Sub
        End If
        
        If FrmSastre.Visible Then
            Unload FrmSastre
            HayFormularioAbierto = False

            '  Exit Sub
        End If

        If frmAlqui.Visible Then
            Unload frmAlqui
            HayFormularioAbierto = False

            ' Exit Sub
        End If

        If frmCarp.Visible Then
            Unload frmCarp
            HayFormularioAbierto = False

            ' Exit Sub
        End If
        
        If FrmGrupo.Visible Then
            Unload FrmGrupo
            HayFormularioAbierto = False

            ' Exit Sub
        End If
        
        If FrmCorreo.Visible Then
            Unload FrmCorreo
            HayFormularioAbierto = False

            ' Exit Sub
        End If
        
        If frmGoliath.Visible Then
            Unload frmGoliath
            HayFormularioAbierto = False

            'Exit Sub
        End If
           
        If FrmViajes.Visible Then
            Unload FrmViajes
            HayFormularioAbierto = False

            ' Exit Sub
        End If
        
        If frmCantidad.Visible Then
            Unload frmCantidad
            HayFormularioAbierto = False

            ' Exit Sub
        End If
        
        If FrmRanking.Visible Then
            Unload FrmRanking
            HayFormularioAbierto = False

            ' Exit Sub
        End If
        
        If frmMapaGrande.Visible Then
            Call CalcularPosicionMAPA

        End If

    End If

    Call SwitchMapIAO(UserMap)

End Sub

''
' Handles the PosUpdate message.

Private Sub HandlePosUpdate()

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
    bTecho = IIf(MapData(UserPos.x, UserPos.y).Trigger = 1 Or MapData(UserPos.x, UserPos.y).Trigger = 2 Or MapData(UserPos.x, UserPos.y).Trigger = 6 Or MapData(UserPos.x, UserPos.y).Trigger > 9 Or MapData(UserPos.x, UserPos.y).Trigger = 4, True, False)
                
    'Update pos label and minimap
    frmmain.Coord.Caption = UserMap & "-" & UserPos.x & "-" & UserPos.y
    frmmain.personaje(0).Left = UserPos.x - 5
    frmmain.personaje(0).Top = UserPos.y - 4

End Sub

''
' Handles the NPCHitUser message.

Private Sub HandleNPCHitUser()

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
    
    Select Case incomingData.ReadByte()

        Case bCabeza
            Call AddtoRichTextBox(frmmain.RecTxt, MENSAJE_GOLPE_CABEZA & CStr(incomingData.ReadInteger()), 255, 0, 0, True, False, False)

        Case bBrazoIzquierdo
            Call AddtoRichTextBox(frmmain.RecTxt, MENSAJE_GOLPE_BRAZO_IZQ & CStr(incomingData.ReadInteger()), 255, 0, 0, True, False, False)

        Case bBrazoDerecho
            Call AddtoRichTextBox(frmmain.RecTxt, MENSAJE_GOLPE_BRAZO_DER & CStr(incomingData.ReadInteger()), 255, 0, 0, True, False, False)

        Case bPiernaIzquierda
            Call AddtoRichTextBox(frmmain.RecTxt, MENSAJE_GOLPE_PIERNA_IZQ & CStr(incomingData.ReadInteger()), 255, 0, 0, True, False, False)

        Case bPiernaDerecha
            Call AddtoRichTextBox(frmmain.RecTxt, MENSAJE_GOLPE_PIERNA_DER & CStr(incomingData.ReadInteger()), 255, 0, 0, True, False, False)

        Case bTorso
            Call AddtoRichTextBox(frmmain.RecTxt, MENSAJE_GOLPE_TORSO & CStr(incomingData.ReadInteger()), 255, 0, 0, True, False, False)

    End Select

End Sub

''
' Handles the UserHitNPC message.

Private Sub HandleUserHitNPC()

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
    
    Call AddtoRichTextBox(frmmain.RecTxt, MENSAJE_GOLPE_CRIATURA_1 & CStr(incomingData.ReadLong()) & MENSAJE_2, 255, 0, 0, True, False, False)

End Sub

''
' Handles the UserAttackedSwing message.

Private Sub HandleUserAttackedSwing()

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
    
    Call AddtoRichTextBox(frmmain.RecTxt, MENSAJE_1 & charlist(incomingData.ReadInteger()).nombre & MENSAJE_ATAQUE_FALLO, 255, 0, 0, True, False, False)

End Sub

''
' Handles the UserHittingByUser message.

Private Sub HandleUserHittedByUser()

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
    
    Select Case incomingData.ReadByte

        Case bCabeza
            Call AddtoRichTextBox(frmmain.RecTxt, attacker & MENSAJE_RECIVE_IMPACTO_CABEZA & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)

        Case bBrazoIzquierdo
            Call AddtoRichTextBox(frmmain.RecTxt, attacker & MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)

        Case bBrazoDerecho
            Call AddtoRichTextBox(frmmain.RecTxt, attacker & MENSAJE_RECIVE_IMPACTO_BRAZO_DER & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)

        Case bPiernaIzquierda
            Call AddtoRichTextBox(frmmain.RecTxt, attacker & MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)

        Case bPiernaDerecha
            Call AddtoRichTextBox(frmmain.RecTxt, attacker & MENSAJE_RECIVE_IMPACTO_PIERNA_DER & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)

        Case bTorso
            Call AddtoRichTextBox(frmmain.RecTxt, attacker & MENSAJE_RECIVE_IMPACTO_TORSO & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)

    End Select

End Sub

''
' Handles the UserHittedUser message.

Private Sub HandleUserHittedUser()

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
    
    Select Case incomingData.ReadByte

        Case bCabeza
            Call AddtoRichTextBox(frmmain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_CABEZA & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)

        Case bBrazoIzquierdo
            Call AddtoRichTextBox(frmmain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)

        Case bBrazoDerecho
            Call AddtoRichTextBox(frmmain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_BRAZO_DER & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)

        Case bPiernaIzquierda
            Call AddtoRichTextBox(frmmain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)

        Case bPiernaDerecha
            Call AddtoRichTextBox(frmmain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_PIERNA_DER & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)

        Case bTorso
            Call AddtoRichTextBox(frmmain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & victim & MENSAJE_PRODUCE_IMPACTO_TORSO & CStr(incomingData.ReadInteger() & MENSAJE_2), 255, 0, 0, True, False, False)

    End Select

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

    Dim g          As Byte

    Dim b          As Byte

    Dim colortexto As Long

    Dim QueEs      As String

    chat = buffer.ReadASCIIString()
    charindex = buffer.ReadInteger()
    
    r = buffer.ReadByte()
    g = buffer.ReadByte()
    b = buffer.ReadByte()
    
    colortexto = buffer.ReadLong()

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
            chat = DESCFINAL(ReadField(2, chat, Asc("*")))
            copiar = True
            duracion = 20
            
        Case "QUESTNEXT"
            chat = NEXTQUEST(ReadField(2, chat, Asc("*")))
            copiar = True
            duracion = 20
        
    End Select
            
    'Only add the chat if the character exists (a CharacterRemove may have been sent to the PC / NPC area before the buffer was flushed)
    If charlist(charindex).active Then

        Call Char_Dialog_Set(charindex, chat, colortexto, duracion, 30)

    End If
    
    If charlist(charindex).EsNpc = False Then
         
        If CopiarDialogoAConsola = 1 And Not copiar Then
    
            'Call CopiarDialogoToConsola(charlist(charindex).nombre, chat, r & g & b)
            Call WriteChatOverHeadInConsole(charindex, chat, r, g, b)

        End If

    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)

errhandler:

    Dim Error As Long

    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

Private Sub HandleEfectOverHead()

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
    
    Dim chat      As String

    Dim charindex As Integer

    Dim color     As Long
    
    chat = buffer.ReadASCIIString()
    charindex = buffer.ReadInteger()
    
    color = buffer.ReadLong()
    
    ' Debug.Print color

    charlist(charindex).dialogEfec = chat
    charlist(charindex).SubeEfecto = 100
    
    Dim r, g, b As Byte
    
    b = (color And 16711680) / 65536
    g = (color And 65280) / 256
    r = color And 255
    
    charlist(charindex).dialog_Efect_color.r = b
    charlist(charindex).dialog_Efect_color.g = g
    charlist(charindex).dialog_Efect_color.b = r
    charlist(charindex).dialog_Efect_color.a = 255
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)

errhandler:

    Dim Error As Long

    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

Private Sub HandleExpOverHead()

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
    
    Dim chat      As String

    Dim charindex As Integer

    chat = buffer.ReadASCIIString()
    charindex = buffer.ReadInteger()

    If charlist(charindex).active Then

        charlist(charindex).dialogExp = "+" & chat
        charlist(charindex).SubeExp = 255
        charlist(charindex).dialog_Exp_color.r = 42
        charlist(charindex).dialog_Exp_color.g = 169
        charlist(charindex).dialog_Exp_color.b = 222
            
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)

errhandler:

    Dim Error As Long

    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing
    
    If Error <> 0 Then Err.Raise Error

End Sub

Private Sub HandleOroOverHead()

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
    
    Dim chat      As String

    Dim charindex As Integer

    chat = buffer.ReadASCIIString()
    charindex = buffer.ReadInteger()

    If charlist(charindex).active Then

        charlist(charindex).dialogOro = "+" & chat
        charlist(charindex).SubeOro = 255
        charlist(charindex).dialog_Oro_color.r = 204
        charlist(charindex).dialog_Oro_color.g = 193
        charlist(charindex).dialog_Oro_color.b = 115
            
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)

errhandler:

    Dim Error As Long

    Error = Err.number

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
    Dim g         As Byte
    Dim b         As Byte
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
            NpcName = NpcData(ReadField(2, chat, Asc("*"))).name
            chat = NpcName & ReadField(3, chat, Asc("*"))

        Case "O" 'OBJETO
            objname = ObjData(ReadField(2, chat, Asc("*"))).name
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
            g = 255
        Else
            g = Val(str)

        End If
            
        str = ReadField(4, chat, 126)

        If Val(str) > 255 Then
            b = 255
        Else
            b = Val(str)

        End If
            
        Call AddtoRichTextBox(frmmain.RecTxt, Left$(chat, InStr(1, chat, "~") - 1), r, g, b, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
    
    Else

        With FontTypes(fontIndex)
            Call AddtoRichTextBox(frmmain.RecTxt, chat, .red, .green, .blue, .bold, .italic)
        End With

    End If
    
    Exit Sub
    
errhandler:

    Dim Error As Long

    Error = Err.number

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

    Dim g         As Byte

    Dim b         As Byte

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
            g = 255
        Else
            g = Val(str)

        End If
            
        str = ReadField(4, chat, 126)

        If Val(str) > 255 Then
            b = 255
        Else
            b = Val(str)

        End If
            
        Call AddtoRichTextBox(frmmain.RecTxt, Left$(chat, InStr(1, chat, "~") - 1), r, g, b, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
    Else

        With FontTypes(fontIndex)
            Call AddtoRichTextBox(frmmain.RecTxt, chat, .red, .green, .blue, .bold, .italic)

        End With

    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    Dim Error As Long

    Error = Err.number

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

    Dim g    As Byte

    Dim b    As Byte

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
            g = 255
        Else
            g = Val(str)

        End If
            
        str = ReadField(4, chat, 126)

        If Val(str) > 255 Then
            b = 255
        Else
            b = Val(str)

        End If
            
        Call AddtoRichTextBox(frmmain.RecTxt, Left$(chat, InStr(1, chat, "~") - 1), r, g, b, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
    Else

        With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
            Call AddtoRichTextBox(frmmain.RecTxt, chat, .red, .green, .blue, .bold, .italic)

        End With

    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    Dim Error As Long

    Error = Err.number

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

    If QueRender = 0 Then
        frmMensaje.msg.Caption = mensaje
        frmMensaje.Show , frmmain
    ElseIf QueRender = 1 Then
        Call Sound.Sound_Play(SND_EXCLAMACION)
        Call TextoAlAsistente(mensaje)
        textcolorAsistente(0) = D3DColorXRGB(255, 255, 255)
        textcolorAsistente(1) = textcolorAsistente(0)
        textcolorAsistente(2) = textcolorAsistente(0)
        textcolorAsistente(3) = textcolorAsistente(0)
    ElseIf QueRender = 2 Then
        frmMensaje.Show
        frmMensaje.msg.Caption = mensaje

    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    Dim Error As Long

    Error = Err.number

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
    
    'Call SwitchMapIAO(UserMap)
    
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
    
    If frmmain.Visible Then
        '  frmMain.Visible = False
        
        UserParalizado = False
        UserInmovilizado = False
     
        'BUG CLONES
        Dim i As Integer

        For i = 1 To LastChar
            Call EraseChar(i)
        Next i
        
        frmmain.personaje(1).Visible = False
        frmmain.personaje(2).Visible = False
        frmmain.personaje(3).Visible = False
        frmmain.personaje(4).Visible = False
        frmmain.personaje(5).Visible = False

    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    Dim Error As Long

    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the UserIndexInServer message.

Private Sub HandleUserIndexInServer()

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
    
    userindex = incomingData.ReadInteger()

End Sub

''
' Handles the UserCharIndexInServer message.

Private Sub HandleUserCharIndexInServer()

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
    bTecho = IIf(MapData(UserPos.x, UserPos.y).Trigger = 1 Or MapData(UserPos.x, UserPos.y).Trigger = 2 Or MapData(UserPos.x, UserPos.y).Trigger = 6 Or MapData(UserPos.x, UserPos.y).Trigger > 9 Or MapData(UserPos.x, UserPos.y).Trigger = 4, True, False)
    frmmain.Coord.Caption = UserMap & "-" & UserPos.x & "-" & UserPos.y
    frmmain.personaje(0).Left = UserPos.x - 5
    frmmain.personaje(0).Top = UserPos.y - 4

End Sub

''
' Handles the CharacterCreate message.

Private Sub HandleCharacterCreate()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 48 Then
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
        
        .nombre = buffer.ReadASCIIString()
        .status = buffer.ReadByte()
        
        privs = buffer.ReadByte()
        ParticulaFx = buffer.ReadByte()
        .Head_Aura = buffer.ReadASCIIString()
        .Arma_Aura = buffer.ReadASCIIString()
        .Body_Aura = buffer.ReadASCIIString()
        .Otra_Aura = buffer.ReadASCIIString()
        .Escudo_Aura = buffer.ReadASCIIString()
        .Speeding = buffer.ReadSingle()
        
        .EsNpc = buffer.ReadBoolean()
        .Donador = buffer.ReadByte()
        .appear = buffer.ReadByte()
        appear = .appear
        .group_index = buffer.ReadInteger()
        .clan_index = buffer.ReadInteger()
        .clan_nivel = buffer.ReadByte()
        .UserMinHp = buffer.ReadLong()
        .UserMaxHp = buffer.ReadLong()
        .simbolo = buffer.ReadByte()
        
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

        .MUERTO = (Head = CASPER_HEAD)
        '.AlphaPJ = 255
        
    End With
    
    Call MakeChar(charindex, Body, Head, Heading, x, y, weapon, shield, helmet, ParticulaFx, appear)

    Call RefreshAllChars
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    Dim Error As Long

    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the CharacterRemove message.

Private Sub HandleCharacterRemove()

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

End Sub

''
' Handles the CharacterMove message.

Private Sub HandleCharacterMove()

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

End Sub

''
' Handles the CharacterChange message.

Private Sub HandleCharacterChange()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 18 Then
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
            .iHead = 0
            
        Else
            .Head = HeadData(headIndex)
            .iHead = headIndex

        End If

        .MUERTO = (headIndex = CASPER_HEAD)
        
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

    End With
    
    Call RefreshAllChars

End Sub

''
' Handles the ObjectCreate message.

Private Sub HandleObjectCreate()

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

    Dim color    As Long

    Dim Rango    As Byte

    Dim id       As Long
    
    x = incomingData.ReadByte()
    y = incomingData.ReadByte()
    
    OBJIndex = incomingData.ReadInteger()
    
    MapData(x, y).ObjGrh.GrhIndex = ObjData(OBJIndex).GrhIndex
    
    MapData(x, y).OBJInfo.OBJIndex = OBJIndex
    
    Call InitGrh(MapData(x, y).ObjGrh, MapData(x, y).ObjGrh.GrhIndex)
    
    If ObjData(OBJIndex).CreaLuz <> "" Then
        color = Val(ReadField(2, ObjData(OBJIndex).CreaLuz, Asc(":")))
        Rango = Val(ReadField(1, ObjData(OBJIndex).CreaLuz, Asc(":")))
        MapData(x, y).luz.color = color
        MapData(x, y).luz.Rango = Rango
        
        If Rango < 100 Then
            id = x & y
            LucesCuadradas.Light_Create x, y, color, Rango, id
            LucesCuadradas.Light_Render_All
        Else

            Dim r, g, b As Byte

            b = (color And 16711680) / 65536
            g = (color And 65280) / 256
            r = color And 255
            LucesRedondas.Create_Light_To_Map x, y, Rango - 99, b, g, r
            LucesRedondas.LightRenderAll
            LucesCuadradas.Light_Render_All

        End If
        
    End If
        
    If ObjData(OBJIndex).CreaParticulaPiso <> 0 Then
        MapData(x, y).particle_group = 0
        General_Particle_Create ObjData(OBJIndex).CreaParticulaPiso, x, y, -1

    End If
    
End Sub

Private Sub HandleFxPiso()

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
    
End Sub

''
' Handles the ObjectDelete message.

Private Sub HandleObjectDelete()

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
        MapData(x, y).luz.color = 0
        MapData(x, y).luz.Rango = 0
        LucesCuadradas.Light_Render_All

    End If
    
    MapData(x, y).ObjGrh.GrhIndex = 0
    
    If ObjData(MapData(x, y).OBJInfo.OBJIndex).CreaParticulaPiso <> 0 Then
        Graficos_Particulas.Particle_Group_Remove (MapData(x, y).particle_group)

    End If

End Sub

''
' Handles the BlockPosition message.

Private Sub HandleBlockPosition()

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
    
    Dim x As Byte

    Dim y As Byte
    
    x = incomingData.ReadByte()
    y = incomingData.ReadByte()
    
    If incomingData.ReadBoolean() Then
        MapData(x, y).Blocked = 1
    Else
        MapData(x, y).Blocked = 0

    End If

End Sub

''
' Handles the PlayMIDI message.

Private Sub HandlePlayMIDI()

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

End Sub

''
' Handles the PlayWave message.

Private Sub HandlePlayWave()

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
End Sub

Private Sub HandlePosLLamadaDeClan()

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

    frmmain.LlamaDeclan.Enabled = True

    frmMapaGrande.Shape2.Visible = True

    frmMapaGrande.Shape2.Top = y * 32
    frmMapaGrande.Shape2.Left = x * 32

    LLamadaDeclanX = srcX
    LLamadaDeclanY = srcY

    HayLLamadaDeclan = True
    
    ' Call Audio.PlayWave(CStr(wave) & ".wav", srcX, srcY)
End Sub

Private Sub HandleCharUpdateHP()

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
End Sub

Private Sub HandleArmaMov()

    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    'Remove packet ID
    Call incomingData.ReadByte

    Dim charindex As Integer

    charindex = incomingData.ReadInteger()

    charlist(charindex).MovArmaEscudo = True
    charlist(charindex).Arma.WeaponWalk(charlist(charindex).Heading).Started = 1

End Sub

Private Sub HandleEscudoMov()

    '***************************************************
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If

    'Remove packet ID
    Call incomingData.ReadByte

    Dim charindex As Integer

    charindex = incomingData.ReadInteger()

    charlist(charindex).MovArmaEscudo = True
    charlist(charindex).Escudo.ShieldWalk(charlist(charindex).Heading).Started = 1

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
    
    HayFormularioAbierto = True
    
    Call frmGuildAdm.Show(vbModeless, frmmain)
    
    Exit Sub
    
errhandler:

    Dim Error As Long

    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the AreaChanged message.

Private Sub HandleAreaChanged()

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
    Call incomingData.ReadByte
    
    pausa = Not pausa

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
    
    Call incomingData.ReadByte
    
    If Not InMapBounds(UserPos.x, UserPos.y) Then Exit Sub
    bTecho = (MapData(UserPos.x, UserPos.y).Trigger = 1 Or MapData(UserPos.x, UserPos.y).Trigger = 6 Or MapData(UserPos.x, UserPos.y).Trigger = 2 Or MapData(UserPos.x, UserPos.y).Trigger = 4)
            
    If bRain Then
        If MapDat.LLUVIA Then
            
            If bTecho Then
                Call Sound.Sound_Play(192)
            Else
                Call Sound.Sound_Play(195)

            End If
            
            Call Sound.Ambient_Stop
            
            Call Graficos_Particulas.Engine_Meteo_Particle_Set(-1)

        End If

    Else

        If MapDat.LLUVIA Then
        
            Call Graficos_Particulas.Engine_Meteo_Particle_Set(Particula_Lluvia)

        End If

        ' Call Audio.StopWave(AmbientalesBufferIndex)
    End If
    
    bRain = Not bRain
    Call Meteo_Engine.CargarClima
  
End Sub

Private Sub HandleTrofeoToggleOn()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    Call incomingData.ReadByte

    MostrarTrofeo = True
  
End Sub

Private Sub HandleTrofeoToggleOff()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    Call incomingData.ReadByte

    MostrarTrofeo = False
  
End Sub

''
' Handles the CreateFX message.

Private Sub HandleCreateFX()

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

End Sub

''
' Handles the UpdateUserStats message.

Private Sub HandleUpdateUserStats()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 26 Then
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
    
    If UserPasarNivel > 0 Then
        frmmain.lblPorcLvl.Caption = Round(UserExp * 100 / UserPasarNivel, 0) & "%"
        frmmain.exp.Caption = PonerPuntos(UserExp) & "/" & PonerPuntos(UserPasarNivel)
        frmmain.EXPBAR.Width = UserExp / UserPasarNivel * 204
    Else
        frmmain.EXPBAR.Width = 204
        frmmain.lblPorcLvl.Caption = "" 'nivel maximo
        frmmain.exp.Caption = "¡Nivel máximo!"

    End If
    
    frmmain.Hpshp.Width = UserMinHp / UserMaxHp * 216
    frmmain.HpBar.Caption = UserMinHp & " / " & UserMaxHp
    frmmain.Hpshp.Visible = (UserMinHp > 0)

    If UserMaxMAN > 0 Then
        frmmain.MANShp.Width = UserMinMAN / UserMaxMAN * 216
        frmmain.manabar.Caption = UserMinMAN & " / " & UserMaxMAN
        frmmain.MANShp.Visible = (UserMinMAN > 0)
    Else
        frmmain.manabar.Visible = False
        frmmain.MANShp.Width = 0
        frmmain.MANShp.Visible = False
    End If
    
    frmmain.STAShp.Width = UserMinSTA / UserMaxSTA * 89
    frmmain.stabar.Caption = UserMinSTA & " / " & UserMaxSTA
    frmmain.STAShp.Visible = (UserMinSTA > 0)
    
    frmmain.GldLbl.Caption = PonerPuntos(UserGLD)
    frmmain.lblLvl.Caption = UserLvl
    
    If UserMinHp = 0 Then
        UserEstado = 1
        meteo_estado = 5
        Meteo_Engine.SetNuevoEstado 5
    Else
        UserEstado = 0

    End If

End Sub

''
' Handles the WorkRequestTarget message.

Private Sub HandleWorkRequestTarget()

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
        frmmain.MousePointer = 0
        Call FormParser.Parse_Form(frmmain, E_NORMAL)
        UsingSkill = UsingSkillREcibido
        Exit Sub

    End If

    If UsingSkillREcibido = UsingSkill Then Exit Sub
   
    UsingSkill = UsingSkillREcibido
    frmmain.MousePointer = 2

    If ShowMacros = 1 Then
        If OcultarMacrosAlCastear Then
            OcultarMacro = True

        End If

    End If
    
    Select Case UsingSkill

        Case magia
            Call FormParser.Parse_Form(frmmain, E_CAST)
            Call AddtoRichTextBox(frmmain.RecTxt, MENSAJE_TRABAJO_MAGIA, 100, 100, 120, 0, 0)

        Case Robar
            Call AddtoRichTextBox(frmmain.RecTxt, MENSAJE_TRABAJO_ROBAR, 100, 100, 120, 0, 0)
            Call FormParser.Parse_Form(frmmain, E_SHOOT)

        Case FundirMetal
            Call AddtoRichTextBox(frmmain.RecTxt, MENSAJE_TRABAJO_FUNDIRMETAL, 100, 100, 120, 0, 0)
            Call FormParser.Parse_Form(frmmain, E_SHOOT)

        Case Proyectiles
            Call AddtoRichTextBox(frmmain.RecTxt, MENSAJE_TRABAJO_PROYECTILES, 100, 100, 120, 0, 0)
            Call FormParser.Parse_Form(frmmain, E_ARROW)

        Case eSkill.Talar, eSkill.Alquimia, eSkill.Carpinteria, eSkill.Herreria, eSkill.Mineria, eSkill.Pescar
            Call AddtoRichTextBox(frmmain.RecTxt, "Has click donde deseas trabajar...", 100, 100, 120, 0, 0)
            Call FormParser.Parse_Form(frmmain, E_SHOOT)

        Case Grupo
            Call AddtoRichTextBox(frmmain.RecTxt, MENSAJE_TRABAJO_MAGIA, 100, 100, 120, 0, 0)
            Call FormParser.Parse_Form(frmmain, E_SHOOT)

        Case MarcaDeClan
            Call AddtoRichTextBox(frmmain.RecTxt, "Seleccione el personaje que desea marcar..", 100, 100, 120, 0, 0)
            Call FormParser.Parse_Form(frmmain, E_SHOOT)

        Case MarcaDeGM
            Call AddtoRichTextBox(frmmain.RecTxt, "Seleccione el personaje que desea marcar..", 100, 100, 120, 0, 0)
            Call FormParser.Parse_Form(frmmain, E_SHOOT)

    End Select

End Sub

''
' Handles the ChangeInventorySlot message.

Private Sub HandleChangeInventorySlot()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 13 Then
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
    Dim name              As String
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
    Dim ResistenciaMagica As Byte
    Dim DañoMagico As Byte
    
    Slot = buffer.ReadByte()
    OBJIndex = buffer.ReadInteger()
    Amount = buffer.ReadInteger()
    Equipped = buffer.ReadBoolean()
    Value = buffer.ReadSingle()
    podrausarlo = buffer.ReadByte()
    ResistenciaMagica = buffer.ReadByte()
    DañoMagico = buffer.ReadByte()
    
    Call incomingData.CopyBuffer(buffer)
    
    name = ObjData(OBJIndex).name
    GrhIndex = ObjData(OBJIndex).GrhIndex
    ObjType = ObjData(OBJIndex).ObjType
    MaxHit = ObjData(OBJIndex).MaxHit
    MinHit = ObjData(OBJIndex).MinHit
    MaxDef = ObjData(OBJIndex).MaxDef
    MinDef = ObjData(OBJIndex).MinDef
    
    If Equipped Then

        Select Case ObjType

            Case eObjType.otWeapon
                frmmain.lblWeapon = MinHit & "/" & MaxHit
                UserWeaponEqpSlot = Slot

            Case eObjType.otNudillos
                frmmain.lblWeapon = MinHit & "/" & MaxHit
                UserWeaponEqpSlot = Slot

            Case eObjType.otArmadura
                frmmain.lblArmor = MinDef & "/" & MaxDef
                UserArmourEqpSlot = Slot

            Case eObjType.otESCUDO
                frmmain.lblShielder = MinDef & "/" & MaxDef
                UserHelmEqpSlot = Slot

            Case eObjType.otCASCO
                frmmain.lblHelm = MinDef & "/" & MaxDef
                UserShieldEqpSlot = Slot

        End Select
        
    Else

        Select Case Slot

            Case UserWeaponEqpSlot
                frmmain.lblWeapon = "0/0"
                UserWeaponEqpSlot = 0

            Case UserArmourEqpSlot
                frmmain.lblArmor = "0/0"
                UserArmourEqpSlot = 0

            Case UserHelmEqpSlot
                frmmain.lblShielder = "0/0"
                UserHelmEqpSlot = 0

            Case UserShieldEqpSlot
                frmmain.lblHelm = "0/0"
                UserShieldEqpSlot = 0

        End Select

    End If

    frmmain.lblResis = ResistenciaMagica & "%"
    frmmain.lbldm = DañoMagico & "%"
    
    Call frmmain.Inventario.SetItem(Slot, OBJIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, MinDef, Value, name, podrausarlo)
    
    Call frmComerciar.InvComUsu.SetItem(Slot, OBJIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, MinDef, Value, name, podrausarlo)
    
    Call frmBancoObj.InvBankUsu.SetItem(Slot, OBJIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, MinDef, Value, name, podrausarlo)
    
    Exit Sub
    
errhandler:

    Dim Error As Long: Error = Err.number

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
    
    Dim i As Integer

    Call incomingData.ReadByte
    
    UserInvUnlocked = incomingData.ReadByte
    
    For i = 1 To UserInvUnlocked
    
        frmmain.imgInvLock(i - 1).Picture = LoadInterface("inventoryunlocked.bmp")
    
    Next i
    
    'Call Inventario.DrawInventory
    
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

    Dim name             As String

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
        Call frmmain.Inventario.SetItem(Slot, ReadField(2, slotNum(i), Asc("@")), ReadField(4, slotNum(i), Asc("@")), ReadField(5, slotNum(i), Asc("@")), ReadField(6, slotNum(i), Asc("@")), ReadField(7, slotNum(i), Asc("@")), ReadField(8, slotNum(i), Asc("@")), ReadField(9, slotNum(i), Asc("@")), ReadField(10, slotNum(i), Asc("@")), ReadField(11, slotNum(i), Asc("@")), ReadField(3, slotNum(i), Asc("@")), 0)
    
        With frmmain.Inventario

            If frmComerciar.Visible Then
                Call frmComerciar.InvComUsu.SetItem(Slot, .OBJIndex(Slot), .Amount(Slot), .Equipped(Slot), .GrhIndex(Slot), .ObjType(Slot), .MaxHit(Slot), .MinHit(Slot), .Def(Slot), .Valor(Slot), .ItemName(Slot), .PuedeUsar(Slot))
            ElseIf frmBancoObj.Visible Then
                Call frmBancoObj.InvBankUsu.SetItem(Slot, .OBJIndex(Slot), .Amount(Slot), .Equipped(Slot), .GrhIndex(Slot), .ObjType(Slot), .MaxHit(Slot), .MinHit(Slot), .Def(Slot), .Valor(Slot), .ItemName(Slot), .PuedeUsar(Slot))

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

    Dim Error As Long

    Error = Err.number

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
        .name = ObjData(.OBJIndex).name
        .Amount = buffer.ReadInteger()
        .GrhIndex = ObjData(.OBJIndex).GrhIndex
        .ObjType = ObjData(.OBJIndex).ObjType
        .MaxHit = ObjData(.OBJIndex).MaxHit
        .MinHit = ObjData(.OBJIndex).MinHit
        .Def = ObjData(.OBJIndex).MaxDef
        .Valor = buffer.ReadLong()
        .PuedeUsar = buffer.ReadByte()
        
        Call frmBancoObj.InvBoveda.SetItem(Slot, .OBJIndex, .Amount, .Equipped, .GrhIndex, .ObjType, .MaxHit, .MinHit, .Def, .Valor, .name, .PuedeUsar)

    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    Dim Error As Long

    Error = Err.number

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
    
        If Slot <= frmmain.hlst.ListCount Then
            frmmain.hlst.List(Slot - 1) = HechizoData(Index).nombre
        Else
            Call frmmain.hlst.AddItem(HechizoData(Index).nombre)

        End If

    Else
    
        If Slot <= frmmain.hlst.ListCount Then
            frmmain.hlst.List(Slot - 1) = "(Vacio)"
        Else
            Call frmmain.hlst.AddItem("(Vacio)")

        End If
    
    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
    Exit Sub
    
errhandler:

    Dim Error As Long

    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the Attributes message.

Private Sub HandleAtributes()

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
            HayFormularioAbierto = True
            frmEstadisticas.Show , frmmain
        Else
            LlegaronAtrib = True
        End If
    End If

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

    Dim Error As Long

    Error = Err.number

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
    
    Dim a As Byte

    Dim e As Byte

    Dim c As Byte

    a = 0
    e = 0
    c = 0
    
    For i = 1 To UBound(DefensasHerrero())
    
        If DefensasHerrero(i).Index = 0 Then Exit For
        If ObjData(DefensasHerrero(i).Index).ObjType = 3 Then
           
            ArmadurasHerrero(a).Index = DefensasHerrero(i).Index
            ArmadurasHerrero(a).LHierro = DefensasHerrero(i).LHierro
            ArmadurasHerrero(a).LPlata = DefensasHerrero(i).LPlata
            ArmadurasHerrero(a).LOro = DefensasHerrero(i).LOro
            a = a + 1

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

    Dim Error As Long

    Error = Err.number

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
        
        Call frmCarp.lstArmas.AddItem(ObjData(ObjCarpintero(i)).name)
    Next i
    
    For i = i To UBound(ObjCarpintero())
        ObjCarpintero(i) = 0
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    Dim Error As Long

    Error = Err.number

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

    Dim g As Byte
    
    i = 0
    r = 1
    g = 1
    
    For i = i To UBound(ObjSastre())
    
        If ObjData(ObjSastre(i).Index).ObjType = 3 Then
        
            SastreRopas(r).Index = ObjSastre(i).Index
            SastreRopas(r).PielLobo = ObjSastre(i).PielLobo
            SastreRopas(r).PielOsoPardo = ObjSastre(i).PielOsoPardo
            SastreRopas(r).PielOsoPolar = ObjSastre(i).PielOsoPolar
            r = r + 1

        End If

        If ObjData(ObjSastre(i).Index).ObjType = 17 Then
            SastreGorros(g).Index = ObjSastre(i).Index
            SastreGorros(g).PielLobo = ObjSastre(i).PielLobo
            SastreGorros(g).PielOsoPardo = ObjSastre(i).PielOsoPardo
            SastreGorros(g).PielOsoPolar = ObjSastre(i).PielOsoPolar
            g = g + 1

        End If

    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    Dim Error As Long

    Error = Err.number

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
        tmp = ObjData(Obj).name        'Get the object's name

        ObjAlquimista(i) = Obj
        Call frmAlqui.lstArmas.AddItem(tmp)
    Next i
    
    For i = i To UBound(ObjAlquimista())
        ObjAlquimista(i) = 0
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    Dim Error As Long

    Error = Err.number

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
    Call incomingData.ReadByte
    
    UserDescansar = Not UserDescansar

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

    Dim Error As Long

    Error = Err.number

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
    Call incomingData.ReadByte
    
    UserCiego = True
    
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
    Call incomingData.ReadByte
    
    UserEstupido = True

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

    Dim Error As Long

    Error = Err.number

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
        .name = ObjData(.OBJIndex).name
        .Amount = buffer.ReadInteger()
        .Valor = buffer.ReadSingle()
        .GrhIndex = ObjData(.OBJIndex).GrhIndex
        .ObjType = ObjData(.OBJIndex).ObjType
        .MaxHit = ObjData(.OBJIndex).MaxHit
        .MinHit = ObjData(.OBJIndex).MinHit
        .Def = ObjData(.OBJIndex).MaxDef
        .PuedeUsar = buffer.ReadByte()
        
        Call frmComerciar.InvComNpc.SetItem(Slot, .OBJIndex, .Amount, 0, .GrhIndex, .ObjType, .MaxHit, .MinHit, .Def, .Valor, .name, .PuedeUsar)
        
    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
    Exit Sub
    
errhandler:

    Dim Error As Long

    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the UpdateHungerAndThirst message.

Private Sub HandleUpdateHungerAndThirst()

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
    frmmain.AGUAsp.Width = UserMinAGU / UserMaxAGU * 32
    frmmain.COMIDAsp.Width = UserMinHAM / UserMaxHAM * 32
    frmmain.AGUbar.Caption = UserMinAGU '& " / " & UserMaxAGU
    frmmain.hambar.Caption = UserMinHAM ' & " / " & UserMaxHAM
    frmmain.AGUAsp.Visible = (UserMinAGU > 0)
    frmmain.COMIDAsp.Visible = (UserMinHAM > 0)
End Sub

Private Sub HandleHora()

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
    
    Dim HoraServer        As Integer

    Dim TimerHoraFantasia As Integer

    HoraServer = incomingData.ReadInteger()
    TimerHoraFantasia = incomingData.ReadInteger()
    
    frmmain.HoraFantasiaTimer.Interval = TimerHoraFantasia
    frmmain.HoraFantasiaTimer.Enabled = True
    
    HoraFantasia = HoraServer
        
    'frmMain.lblHoraFantasia.Caption = GetTimeFormated(HoraFantasia)
    
    If UserEstado = 1 Then Exit Sub
    If Map_light_base <> -1 Then
        Map_Base_Light_Set (Map_light_base)
            
    End If
    
    meteo_estado = 0

    If meteo_estado = 0 Then
        If HoraServer > 299 And HoraServer < 661 Then
            Meteo_Engine.SetNuevoEstado (4)
            meteo_estado = 4
            Exit Sub
                
        ElseIf HoraServer > 659 And HoraServer < 1021 Then
            Meteo_Engine.SetNuevoEstado (1)
                
            meteo_estado = 1
        ElseIf HoraServer > 1019 And HoraServer < 1201 Then
            Meteo_Engine.SetNuevoEstado (2)
            meteo_estado = 2
        Else
            Meteo_Engine.SetNuevoEstado (3)
            meteo_estado = 3
                        
        End If
    
    Else
    
        If HoraServer = 300 Then
            If meteo_estado = 1 Then Exit Sub
            Meteo_Engine.NextEstado
            meteo_estado = 1
        ElseIf HoraServer = 720 Then

            If meteo_estado = 2 Then Exit Sub
            Meteo_Engine.NextEstado
            meteo_estado = 2
        ElseIf HoraServer = 1080 Then

            If meteo_estado = 3 Then Exit Sub
            Meteo_Engine.NextEstado
            meteo_estado = 3
        Else

            If meteo_estado = 4 Then Exit Sub
            Meteo_Engine.NextEstado
            meteo_estado = 4

        End If

    End If

End Sub
 
Private Sub HandleLight()
 
    If incomingData.length < 3 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
 
    Dim color As String

    Call incomingData.ReadByte
    color = incomingData.ReadASCIIString()

    If color = "" Then
        Map_light_base = 0
        Exit Sub

    End If

    Dim r, g, b As Byte

    b = (color And 16711680) / 65536
    g = (color And 65280) / 256
    r = color And 255
    Map_light_base = D3DColorARGB(255, r, g, b)
    ColorAmbiente.r = r
    ColorAmbiente.b = b
    ColorAmbiente.g = g
    ColorAmbiente.a = 255
    Call Map_Base_Light_Set(Map_light_base)
 
End Sub
 
Private Sub HandleFYA()

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
        frmmain.Contadores.Enabled = True

    End If
    
    If UserAtributos(eAtributos.Fuerza) >= 33 Then
        frmmain.Fuerzalbl.ForeColor = RGB(204, 0, 0)
    ElseIf UserAtributos(eAtributos.Fuerza) >= 25 Then
        frmmain.Fuerzalbl.ForeColor = RGB(204, 100, 100)
    Else
        frmmain.Fuerzalbl.ForeColor = vbWhite
    End If
    
    If UserAtributos(eAtributos.Agilidad) >= 33 Then
        frmmain.AgilidadLbl.ForeColor = RGB(204, 0, 0)
    ElseIf UserAtributos(eAtributos.Agilidad) >= 25 Then
        frmmain.AgilidadLbl.ForeColor = RGB(204, 100, 100)
    Else
        frmmain.AgilidadLbl.ForeColor = vbWhite
    End If

    frmmain.Fuerzalbl.Caption = UserAtributos(eAtributos.Fuerza)
    frmmain.AgilidadLbl.Caption = UserAtributos(eAtributos.Agilidad)

End Sub

Private Sub HandleUpdateNPCSimbolo()

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

End Sub

Private Sub HandleCerrarleCliente()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    EngineRun = False

    Call CloseClient

End Sub

Private Sub HandleContadores()

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
    
    frmmain.Contadores.Enabled = True
    
End Sub

Private Sub HandleOxigeno()

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
        frmmain.oxigenolbl = SS
        frmmain.oxigenolbl.ForeColor = vbRed
    Else
        frmmain.oxigenolbl = ms
        frmmain.oxigenolbl.ForeColor = vbWhite

    End If
    
End Sub

''
' Handles the MiniStats message.
Private Sub HandleEfectToScreen()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    Dim color As Long, duracion As Long, ignorar As Boolean

    If incomingData.length < 10 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    color = incomingData.ReadLong()
    duracion = incomingData.ReadLong()
    ignorar = incomingData.ReadBoolean()
    
    Dim r, g, b As Byte

    b = (color And 16711680) / 65536
    g = (color And 65280) / 256
    r = color And 255
    color = D3DColorARGB(255, r, g, b)

    If Not MapDat.niebla = 1 And Not ignorar Then
        'Debug.Print "trueno cancelado"
       
        Exit Sub

    End If

    Call EfectoEnPantalla(color, duracion)
    
End Sub

Private Sub HandleMiniStats()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    If incomingData.length < 34 Then
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
        .BattlePuntos = incomingData.ReadLong()

    End With
    
    If LlegaronAtrib And LlegaronSkills Then
        Alocados = SkillPoints
        frmEstadisticas.puntos.Caption = SkillPoints
        frmEstadisticas.Iniciar_Labels
        frmEstadisticas.Picture = LoadInterface("ventanaestadisticas.bmp")
        HayFormularioAbierto = True
        frmEstadisticas.Show , frmmain
    Else
        LlegaronStats = True
    End If

End Sub

''
' Handles the LevelUp message.

Private Sub HandleLevelUp()

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

    Dim Error As Long

    Error = Err.number

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
    Call incomingData.ReadByte
    
    ' If Not frmForo.Visible Then
    '   frmForo.Show , frmMain
    ' End If
End Sub

''
' Handles the SetInvisible message.

Private Sub HandleSetInvisible()

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
    charlist(charindex).invisible = incomingData.ReadBoolean()
    charlist(charindex).TimerI = 0
    
    #If SeguridadAlkon Then

        If charlist(charindex).invisible Then
            Call MI(CualMI).SetInvisible(charindex)
        Else
            Call MI(CualMI).ResetInvisible(charindex)

        End If

    #End If

End Sub

Private Sub HandleSetEscribiendo()

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

End Sub

''
' Handles the DiceRoll message.

Private Sub HandleDiceRoll()

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

End Sub

''
' Handles the MeditateToggle message.

Private Sub HandleMeditateToggle()
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    Dim charindex As Integer, fX As Integer
    
    charindex = incomingData.ReadInteger
    fX = incomingData.ReadInteger
    
    If charindex = UserCharIndex Then
        UserMeditar = (fX <> 0)
    End If
    
    With charlist(charindex)
        If fX <> 0 Then
            Call InitGrh(.fX, FxData(fX).Animacion)
        End If
        
        .FxIndex = fX
        .fX.Loops = -1
        .fX.AnimacionContador = 0
    End With

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
    Call incomingData.ReadByte
    UserCiego = False
    
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
    Call incomingData.ReadByte
    
    UserEstupido = False

End Sub

''
' Handles the SendSkills message.

Private Sub HandleSendSkills()

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
        HayFormularioAbierto = True
        frmEstadisticas.Show , frmmain
    Else
        LlegaronSkills = True
    End If

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

    frmEntrenador.Show , frmmain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    Dim Error As Long

    Error = Err.number

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
    
    frmGuildNews.Show vbModeless, frmmain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    Dim Error As Long

    Error = Err.number

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

    Dim Error As Long

    Error = Err.number

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
    Call frmPeaceProp.Show(vbModeless, frmmain)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    Dim Error As Long

    Error = Err.number

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
    Call frmPeaceProp.Show(vbModeless, frmmain)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    Dim Error As Long

    Error = Err.number

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
        
        Call .Show(vbModeless, frmmain)

    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    Dim Error As Long

    Error = Err.number

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
       
            .porciento.Caption = Round(expacu * 100 / ExpNe, 0) & "%"
        Else
            .porciento.Caption = "¡Nivel Maximo!"
            .expcount.Caption = "¡Nivel Maximo!"

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
        
        .Show , frmmain

    End With

    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    Dim Error As Long

    Error = Err.number

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
    
    frmGuildBrief.Show vbModeless, frmmain
    
errhandler:

    Dim Error As Long

    Error = Err.number

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
    Call incomingData.ReadByte
    
    CreandoClan = True
    frmGuildDetails.Show , frmmain

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
    Call incomingData.ReadByte
    
    UserParalizado = Not UserParalizado

End Sub

Private Sub HandleInmovilizadoOK()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    UserInmovilizado = Not UserInmovilizado

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
    Call frmUserRequest.Show(vbModeless, frmmain)
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    Dim Error As Long

    Error = Err.number

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
    
    With OtroInventario(1)
        .OBJIndex = buffer.ReadInteger()
        .name = buffer.ReadASCIIString()
        .Amount = buffer.ReadLong()
        .GrhIndex = buffer.ReadInteger()
        .ObjType = buffer.ReadByte()
        .MaxHit = buffer.ReadInteger()
        .MinHit = buffer.ReadInteger()
        .Def = buffer.ReadInteger()
        .Valor = buffer.ReadLong()
        
        frmComerciarUsu.List2.Clear
        
        Call frmComerciarUsu.List2.AddItem(.name)
        frmComerciarUsu.List2.ItemData(frmComerciarUsu.List2.NewIndex) = .Amount
        
        frmComerciarUsu.lblEstadoResp.Visible = False

    End With
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    Dim Error As Long

    Error = Err.number

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

    Dim i              As Long
    
    creatureList = Split(buffer.ReadASCIIString(), SEPARATOR)
    
    For i = 0 To UBound(creatureList())
        Call frmSpawnList.lstCriaturas.AddItem(NpcData(creatureList(i)).name)
    Next i

    frmSpawnList.Show , frmmain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    Dim Error As Long

    Error = Err.number

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
        frmPanelGm.List1.AddItem nombre & "(" & TipoDeConsulta & ")"
        frmPanelGm.List2.AddItem Consulta
    Next i
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    Dim Error As Long

    Error = Err.number

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
    frmCambiaMotd.Show , frmmain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    Dim Error As Long

    Error = Err.number

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
    Call incomingData.ReadByte
    
    frmPanelGm.Show vbModeless, frmmain

End Sub

Private Sub HandleShowFundarClanForm()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    Call incomingData.ReadByte
    
    CreandoClan = True
    frmGuildDetails.Show vbModeless, frmmain

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
    
    If frmPanelGm.Visible Then
        frmPanelGm.cboListaUsus.Clear

        For i = 0 To UBound(userList())
            Call frmPanelGm.cboListaUsus.AddItem(userList(i))
        Next i

        If frmPanelGm.cboListaUsus.ListCount > 0 Then frmPanelGm.cboListaUsus.ListIndex = 0

    End If
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    Dim Error As Long

    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the Pong message.

Private Sub HandlePong()

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
    
    Dim time As Long

    time = incomingData.ReadLong()
    'Call AddtoRichTextBox(frmMain.RecTxt, "El ping anterior seria " & (GetTickCount - pingTime) & " ms.", 255, 0, 0, True, False, False)
    'Call AddtoRichTextBox(frmMain.RecTxt, "El ping es " & (GetTickCount - time) & " ms.", 255, 0, 0, True, False, False)
    'timeGetTime -pingTime
    PingRender = timeGetTime - time
    pingTime = 0

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

    Dim userTag     As String

    Dim group_index As Integer
    
    charindex = buffer.ReadInteger()
    status = buffer.ReadByte()
    userTag = buffer.ReadASCIIString()
    
    group_index = buffer.ReadInteger()
    
    'Update char status adn tag!
    charlist(charindex).status = status
    charlist(charindex).nombre = userTag
    
    charlist(charindex).group_index = group_index
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    Dim Error As Long

    Error = Err.number

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
        
    End With

End Sub

''
' Writes the "LoginNewChar" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLoginNewChar()

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
        Call .WriteASCIIString(MacAdress)  'Seguridad
        Call .WriteLong(HDserial)  'SeguridadHDserial

    End With

End Sub

''
' Writes the "Talk" message to the outgoing data buffer.
'
' @param    chat The chat text to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTalk(ByVal chat As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Talk" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Talk)
        
        Call .WriteASCIIString(chat)

    End With

End Sub

''
' Writes the "Yell" message to the outgoing data buffer.
'
' @param    chat The chat text to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteYell(ByVal chat As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Yell" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Yell)
        
        Call .WriteASCIIString(chat)

    End With

End Sub

''
' Writes the "Whisper" message to the outgoing data buffer.
'
' @param    charIndex The index of the char to whom to whisper.
' @param    chat The chat text to be sent to the user.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWhisper(ByVal nombre As String, ByVal chat As String)

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

End Sub

''
' Writes the "Walk" message to the outgoing data buffer.
'
' @param    heading The direction in wich the user is moving.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWalk(ByVal Heading As E_Heading)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Walk" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Walk)
        
        Call .WriteByte(Heading)

    End With

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
    Call outgoingData.WriteByte(ClientPacketID.RequestPositionUpdate)

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
    Call outgoingData.WriteByte(ClientPacketID.Attack)

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
    Call outgoingData.WriteByte(ClientPacketID.PickUp)

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
    Call outgoingData.WriteByte(ClientPacketID.SafeToggle)

End Sub

Public Sub WriteSeguroClan()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SafeToggle" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.SeguroClan)

End Sub

Public Sub WriteTraerBoveda()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SafeToggle" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.TraerBoveda)

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
    Call outgoingData.WriteByte(ClientPacketID.PartySafeToggle)

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
    Call outgoingData.WriteByte(ClientPacketID.RequestGuildLeaderInfo)

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
    Call outgoingData.WriteByte(ClientPacketID.RequestAtributes)

End Sub

Public Sub WriteRequestFamiliar()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestFamiliar" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.RequestFamiliar)

End Sub

Public Sub WriteRequestGrupo()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestFamiliar" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.RequestGrupo)

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
    Call outgoingData.WriteByte(ClientPacketID.RequestSkills)

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
    Call outgoingData.WriteByte(ClientPacketID.RequestMiniStats)

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
    Call outgoingData.WriteByte(ClientPacketID.CommerceEnd)

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
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceEnd)

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
    Call outgoingData.WriteByte(ClientPacketID.BankEnd)

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
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceOk)

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
    Call outgoingData.WriteByte(ClientPacketID.UserCommerceReject)

End Sub

''
' Writes the "Drop" message to the outgoing data buffer.
'
' @param    slot Inventory slot where the item to drop is.
' @param    amount Number of items to drop.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDrop(ByVal Slot As Byte, ByVal Amount As Long)

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

End Sub

''
' Writes the "CastSpell" message to the outgoing data buffer.
'
' @param    slot Spell List slot where the spell to cast is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCastSpell(ByVal Slot As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CastSpell" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CastSpell)
        
        Call .WriteByte(Slot)

    End With

End Sub

Public Sub WriteInvitarGrupo()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CastSpell" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.InvitarGrupo)

    End With

End Sub

Public Sub WriteMarcaDeClan()

    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 23/08/2020
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.MarcaDeClanpack)

    End With

End Sub

Public Sub WriteMarcaDeGm()

    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 23/08/2020
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.MarcaDeGMPack)

    End With

End Sub

Public Sub WriteAbandonarGrupo()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CastSpell" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.AbandonarGrupo)

    End With

End Sub

Public Sub WriteHecharDeGrupo(ByVal indice As Byte)

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

End Sub

''
' Writes the "LeftClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLeftClick(ByVal x As Byte, ByVal y As Byte)

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

End Sub

''
' Writes the "DoubleClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDoubleClick(ByVal x As Byte, ByVal y As Byte)

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

End Sub

''
' Writes the "Work" message to the outgoing data buffer.
'
' @param    skill The skill which the user attempts to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWork(ByVal Skill As eSkill)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Work" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Work)
        
        Call .WriteByte(Skill)

    End With

End Sub

Public Sub WriteThrowDice()
    Call outgoingData.WriteByte(ClientPacketID.ThrowDice)

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
    Call outgoingData.WriteByte(ClientPacketID.UseSpellMacro)

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

    With outgoingData
        Call .WriteByte(ClientPacketID.UseItem)
        Call .WriteByte(Slot)

    End With

End Sub

''
' Writes the "CraftBlacksmith" message to the outgoing data buffer.
'
' @param    item Index of the item to craft in the list sent by the server.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCraftBlacksmith(ByVal Item As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CraftBlacksmith" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CraftBlacksmith)
        
        Call .WriteInteger(Item)

    End With

End Sub

''
' Writes the "CraftCarpenter" message to the outgoing data buffer.
'
' @param    item Index of the item to craft in the list sent by the server.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCraftCarpenter(ByVal Item As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CraftCarpenter" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CraftCarpenter)
        
        Call .WriteInteger(Item)

    End With

End Sub

Public Sub WriteCraftAlquimista(ByVal Item As Integer)

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

End Sub

Public Sub WriteCraftSastre(ByVal Item As Integer)

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

    'Call IntervaloPermiteClick(True)
    With outgoingData
        Call .WriteByte(ClientPacketID.WorkLeftClick)
        
        Call .WriteByte(x)
        Call .WriteByte(y)
        
        Call .WriteByte(Skill)

    End With

End Sub

''
' Writes the "CreateNewGuild" message to the outgoing data buffer.
'
' @param    desc    The guild's description
' @param    name    The guild's name
' @param    site    The guild's website
' @param    codex   Array of all rules of the guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNewGuild(ByVal desc As String, ByVal name As String, ByVal Alineacion As Byte)

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
        Call .WriteASCIIString(name)
        
        Call .WriteByte(Alineacion)

    End With

End Sub

''
' Writes the "SpellInfo" message to the outgoing data buffer.
'
' @param    slot Spell List slot where the spell which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpellInfo(ByVal Slot As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SpellInfo" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.SpellInfo)
        
        Call .WriteByte(Slot)

    End With

End Sub

''
' Writes the "EquipItem" message to the outgoing data buffer.
'
' @param    slot Invetory slot where the item to equip is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteEquipItem(ByVal Slot As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "EquipItem" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.EquipItem)
        
        Call .WriteByte(Slot)

    End With

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

    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeHeading)
        
        Call .WriteByte(Heading)

    End With

End Sub

''
' Writes the "ModifySkills" message to the outgoing data buffer.
'
' @param    skillEdt a-based array containing for each skill the number of points to add to it.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteModifySkills(ByRef skillEdt() As Byte)

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

End Sub

''
' Writes the "Train" message to the outgoing data buffer.
'
' @param    creature Position within the list provided by the server of the creature to train against.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTrain(ByVal creature As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Train" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Train)
        
        Call .WriteByte(creature)

    End With

End Sub

''
' Writes the "CommerceBuy" message to the outgoing data buffer.
'
' @param    slot Position within the NPC's inventory in which the desired item is.
' @param    amount Number of items to buy.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceBuy(ByVal Slot As Byte, ByVal Amount As Integer)

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

End Sub

Public Sub WriteUseKey(ByVal Slot As Byte)

    With outgoingData
        Call .WriteByte(ClientPacketID.UseKey)
        Call .WriteByte(Slot)
    End With

End Sub

''
' Writes the "BankExtractItem" message to the outgoing data buffer.
'
' @param    slot Position within the bank in which the desired item is.
' @param    amount Number of items to extract.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankExtractItem(ByVal Slot As Byte, ByVal Amount As Integer, ByVal slotdestino As Byte)

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

End Sub

''
' Writes the "CommerceSell" message to the outgoing data buffer.
'
' @param    slot Position within user inventory in which the desired item is.
' @param    amount Number of items to sell.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCommerceSell(ByVal Slot As Byte, ByVal Amount As Integer)

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

End Sub

''
' Writes the "BankDeposit" message to the outgoing data buffer.
'
' @param    slot Position within the user inventory in which the desired item is.
' @param    amount Number of items to deposit.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankDeposit(ByVal Slot As Byte, ByVal Amount As Integer, ByVal slotdestino As Byte)

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

End Sub

''
' Writes the "ForumPost" message to the outgoing data buffer.
'
' @param    title The message's title.
' @param    message The body of the message.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForumPost(ByVal title As String, ByVal Message As String)

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

End Sub

''
' Writes the "MoveSpell" message to the outgoing data buffer.
'
' @param    upwards True if the spell will be moved up in the list, False if it will be moved downwards.
' @param    slot Spell List slot where the spell which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMoveSpell(ByVal upwards As Boolean, ByVal Slot As Byte)

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

End Sub

''
' Writes the "ClanCodexUpdate" message to the outgoing data buffer.
'
' @param    desc New description of the clan.
' @param    codex New codex of the clan.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteClanCodexUpdate(ByVal desc As String)

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

End Sub

''
' Writes the "UserCommerceOffer" message to the outgoing data buffer.
'
' @param    slot Position within user inventory in which the desired item is.
' @param    amount Number of items to offer.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUserCommerceOffer(ByVal Slot As Byte, ByVal Amount As Long)

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

End Sub

''
' Writes the "GuildAcceptPeace" message to the outgoing data buffer.
'
' @param    guild The guild whose peace offer is accepted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAcceptPeace(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildAcceptPeace" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAcceptPeace)
        
        Call .WriteASCIIString(guild)

    End With

End Sub

''
' Writes the "GuildRejectAlliance" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance offer is rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRejectAlliance(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRejectAlliance" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRejectAlliance)
        
        Call .WriteASCIIString(guild)

    End With

End Sub

''
' Writes the "GuildRejectPeace" message to the outgoing data buffer.
'
' @param    guild The guild whose peace offer is rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRejectPeace(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRejectPeace" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRejectPeace)
        
        Call .WriteASCIIString(guild)

    End With

End Sub

''
' Writes the "GuildAcceptAlliance" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance offer is accepted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAcceptAlliance(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildAcceptAlliance" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAcceptAlliance)
        
        Call .WriteASCIIString(guild)

    End With

End Sub

''
' Writes the "GuildOfferPeace" message to the outgoing data buffer.
'
' @param    guild The guild to whom peace is offered.
' @param    proposal The text to send with the proposal.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOfferPeace(ByVal guild As String, ByVal proposal As String)

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

End Sub

''
' Writes the "GuildOfferAlliance" message to the outgoing data buffer.
'
' @param    guild The guild to whom an aliance is offered.
' @param    proposal The text to send with the proposal.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOfferAlliance(ByVal guild As String, ByVal proposal As String)

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

End Sub

''
' Writes the "GuildAllianceDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance proposal's details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAllianceDetails(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildAllianceDetails" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAllianceDetails)
        
        Call .WriteASCIIString(guild)

    End With

End Sub

''
' Writes the "GuildPeaceDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose peace proposal's details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildPeaceDetails(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildPeaceDetails" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildPeaceDetails)
        
        Call .WriteASCIIString(guild)

    End With

End Sub

''
' Writes the "GuildRequestJoinerInfo" message to the outgoing data buffer.
'
' @param    username The user who wants to join the guild whose info is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestJoinerInfo(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRequestJoinerInfo" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRequestJoinerInfo)
        
        Call .WriteASCIIString(UserName)

    End With

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
    Call outgoingData.WriteByte(ClientPacketID.GuildAlliancePropList)

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
    Call outgoingData.WriteByte(ClientPacketID.GuildPeacePropList)

End Sub

''
' Writes the "GuildDeclareWar" message to the outgoing data buffer.
'
' @param    guild The guild to which to declare war.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildDeclareWar(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildDeclareWar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildDeclareWar)
        
        Call .WriteASCIIString(guild)

    End With

End Sub

''
' Writes the "GuildNewWebsite" message to the outgoing data buffer.
'
' @param    url The guild's new website's URL.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildNewWebsite(ByVal url As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildNewWebsite" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildNewWebsite)
        
        Call .WriteASCIIString(url)

    End With

End Sub

''
' Writes the "GuildAcceptNewMember" message to the outgoing data buffer.
'
' @param    username The name of the accepted player.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildAcceptNewMember(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildAcceptNewMember" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildAcceptNewMember)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "GuildRejectNewMember" message to the outgoing data buffer.
'
' @param    username The name of the rejected player.
' @param    reason The reason for which the player was rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRejectNewMember(ByVal UserName As String, ByVal reason As String)

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

End Sub

''
' Writes the "GuildKickMember" message to the outgoing data buffer.
'
' @param    username The name of the kicked player.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildKickMember(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildKickMember" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildKickMember)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "GuildUpdateNews" message to the outgoing data buffer.
'
' @param    news The news to be posted.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildUpdateNews(ByVal news As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildUpdateNews" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildUpdateNews)
        
        Call .WriteASCIIString(news)

    End With

End Sub

''
' Writes the "GuildMemberInfo" message to the outgoing data buffer.
'
' @param    username The user whose info is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMemberInfo(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildMemberInfo" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildMemberInfo)
        
        Call .WriteASCIIString(UserName)

    End With

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
    Call outgoingData.WriteByte(ClientPacketID.GuildOpenElections)

End Sub

''
' Writes the "GuildRequestMembership" message to the outgoing data buffer.
'
' @param    guild The guild to which to request membership.
' @param    application The user's application sheet.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestMembership(ByVal guild As String, ByVal Application As String)

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

End Sub

''
' Writes the "GuildRequestDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildRequestDetails(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRequestDetails" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildRequestDetails)
        
        Call .WriteASCIIString(guild)

    End With

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
    Call outgoingData.WriteByte(ClientPacketID.Online)

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
    Call outgoingData.WriteByte(ClientPacketID.Quit)
    UserSaliendo = True

    Rem  MostrarCuenta = True
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
    Call outgoingData.WriteByte(ClientPacketID.GuildLeave)

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
    Call outgoingData.WriteByte(ClientPacketID.RequestAccountState)

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
    Call outgoingData.WriteByte(ClientPacketID.PetStand)

End Sub

''
' Writes the "GrupoMsg" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGrupoMsg(ByVal Message As String)

    With outgoingData
        Call .WriteByte(ClientPacketID.GrupoMsg)
        
        Call .WriteASCIIString(Message)

    End With

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
    Call outgoingData.WriteByte(ClientPacketID.TrainList)

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
    Call outgoingData.WriteByte(ClientPacketID.Rest)

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
    Call outgoingData.WriteByte(ClientPacketID.Meditate)

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
    Call outgoingData.WriteByte(ClientPacketID.Resucitate)

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
    Call outgoingData.WriteByte(ClientPacketID.Heal)

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
    Call outgoingData.WriteByte(ClientPacketID.Help)

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
    Call outgoingData.WriteByte(ClientPacketID.RequestStats)

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
    Call outgoingData.WriteByte(ClientPacketID.CommerceStart)

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
    Call outgoingData.WriteByte(ClientPacketID.BankStart)

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
    Call outgoingData.WriteByte(ClientPacketID.Enlist)

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
    Call outgoingData.WriteByte(ClientPacketID.Information)

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
    Call outgoingData.WriteByte(ClientPacketID.Reward)

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
    Call outgoingData.WriteByte(ClientPacketID.RequestMOTD)

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
    Call outgoingData.WriteByte(ClientPacketID.Uptime)

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
    Call outgoingData.WriteByte(ClientPacketID.Inquiry)

End Sub

''
' Writes the "GuildMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMessage(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRequestDetails" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildMessage)
        
        Call .WriteASCIIString(Message)

    End With

End Sub

''
' Writes the "CentinelReport" message to the outgoing data buffer.
'
' @param    number The number to report to the centinel.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCentinelReport(ByVal number As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CentinelReport" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CentinelReport)
        
        Call .WriteInteger(number)

    End With

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
    Call outgoingData.WriteByte(ClientPacketID.GuildOnline)

End Sub

''
' Writes the "CouncilMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the other council members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCouncilMessage(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CouncilMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CouncilMessage)
        
        Call .WriteASCIIString(Message)

    End With

End Sub

''
' Writes the "RoleMasterRequest" message to the outgoing data buffer.
'
' @param    message The message to send to the role masters.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoleMasterRequest(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RoleMasterRequest" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RoleMasterRequest)
        
        Call .WriteASCIIString(Message)

    End With

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
    Call outgoingData.WriteByte(ClientPacketID.GMRequest)

End Sub

''
' Writes the "ChangeDescription" message to the outgoing data buffer.
'
' @param    desc The new description of the user's character.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeDescription(ByVal desc As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeDescription" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeDescription)
        
        Call .WriteASCIIString(desc)

    End With

End Sub

''
' Writes the "GuildVote" message to the outgoing data buffer.
'
' @param    username The user to vote for clan leader.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildVote(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildVote" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildVote)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "Punishments" message to the outgoing data buffer.
'
' @param    username The user whose's  punishments are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WritePunishments(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Punishments" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Punishments)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "ChangePassword" message to the outgoing data buffer.
'
' @param    oldPass Previous password.
' @param    newPass New password.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangePassword(ByRef oldPass As String, ByRef newPass As String)

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

End Sub

''
' Writes the "Gamble" message to the outgoing data buffer.
'
' @param    amount The amount to gamble.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGamble(ByVal Amount As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Gamble" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Gamble)
        
        Call .WriteInteger(Amount)

    End With

End Sub

''
' Writes the "InquiryVote" message to the outgoing data buffer.
'
' @param    opt The chosen option to vote for.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInquiryVote(ByVal opt As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "InquiryVote" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.InquiryVote)
        
        Call .WriteByte(opt)

    End With

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
    Call outgoingData.WriteByte(ClientPacketID.LeaveFaction)

End Sub

''
' Writes the "BankExtractGold" message to the outgoing data buffer.
'
' @param    amount The amount of money to extract from the bank.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankExtractGold(ByVal Amount As Long)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankExtractGold" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankExtractGold)
        
        Call .WriteLong(Amount)

    End With

End Sub

''
' Writes the "BankDepositGold" message to the outgoing data buffer.
'
' @param    amount The amount of money to deposit in the bank.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBankDepositGold(ByVal Amount As Long)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankDepositGold" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.BankDepositGold)
        
        Call .WriteLong(Amount)

    End With

End Sub

Public Sub WriteTransFerGold(ByVal Amount As Long, ByVal destino As String)

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

End Sub

Public Sub WriteItemMove(ByVal SlotActual As Byte, ByVal SlotNuevo As Byte)

    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.MoveItem)
        Call .WriteByte(SlotActual)
        Call .WriteByte(SlotNuevo)

    End With

End Sub

Public Sub WriteBovedaItemMove(ByVal SlotActual As Byte, ByVal SlotNuevo As Byte)

    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.BovedaMoveItem)
        Call .WriteByte(SlotActual)
        Call .WriteByte(SlotNuevo)

    End With

End Sub

''
' Writes the "Denounce" message to the outgoing data buffer.
'
' @param    message The message to send with the denounce.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteDenounce()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Denounce" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Denounce)

    End With

End Sub

Public Sub WriteQuieroFundarClan()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Denounce" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.QuieroFundarClan)

    End With

End Sub

''
' Writes the "GuildMemberList" message to the outgoing data buffer.
'
' @param    guild The guild whose member list is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildMemberList(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildMemberList" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildMemberList)
        
        Call .WriteASCIIString(guild)

    End With

End Sub

'ladder
Public Sub WriteCasamiento(ByVal UserName As String)

    '***************************************************
    'Ladder
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.Casarse)
        Call .WriteASCIIString(UserName)

    End With

End Sub

Public Sub WriteDropItem(ByVal Item As Byte, ByVal x As Byte, ByVal y As Byte, ByVal DropItem As Integer)
    '***************************************************
    'Ladder
    '***************************************************

    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.DropItem)
        Call .WriteByte(Item)
        Call .WriteByte(x)
        Call .WriteByte(y)
        Call .WriteInteger(DropItem)

    End With

End Sub

Public Sub WriteMacroPos()

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

End Sub

Public Sub WriteSubastaInfo()

    '***************************************************
    'Ladder
    'Macros
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.SubastaInfo)

    End With

End Sub

Public Sub WriteScrollInfo()

    '***************************************************
    'Ladder
    'Macros
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.ScrollInfo)

    End With

End Sub

Public Sub WriteCancelarExit()
    '***************************************************
    'Ladder
    'Cancelar Salida
    '***************************************************
    UserSaliendo = False

    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.CancelarExit)

    End With

End Sub

Public Sub WriteEventoInfo()

    '***************************************************
    'Ladder
    'Macros
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.EventoInfo)

    End With

End Sub

Public Sub WriteFlagTrabajar()

    '***************************************************
    'Ladder
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.FlagTrabajar)

    End With

End Sub

''
' Writes the "GMMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to the other GMs online.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteEscribiendo()

    '***************************************************
    'Ladder
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.Escribiendo)

    End With

End Sub

Public Sub WriteReclamarRecompensa(ByVal Index As Byte)

    '***************************************************
    'Ladder
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.ReclamarRecompensa)
        Call .WriteByte(Index)

    End With

End Sub

Public Sub WriteGMMessage(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GMMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GMMessage)
        
        Call .WriteASCIIString(Message)

    End With

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
    Call outgoingData.WriteByte(ClientPacketID.showName)

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
    Call outgoingData.WriteByte(ClientPacketID.OnlineRoyalArmy)

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
    Call outgoingData.WriteByte(ClientPacketID.OnlineChaosLegion)

End Sub

''
' Writes the "GoNearby" message to the outgoing data buffer.
'
' @param    username The suer to approach.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGoNearby(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GoNearby" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GoNearby)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "Comment" message to the outgoing data buffer.
'
' @param    message The message to leave in the log as a comment.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteComment(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Comment" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.comment)
        
        Call .WriteASCIIString(Message)

    End With

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
    Call outgoingData.WriteByte(ClientPacketID.serverTime)

End Sub

''
' Writes the "Where" message to the outgoing data buffer.
'
' @param    username The user whose position is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWhere(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Where" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Where)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "CreaturesInMap" message to the outgoing data buffer.
'
' @param    map The map in which to check for the existing creatures.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreaturesInMap(ByVal map As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CreaturesInMap" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CreaturesInMap)
        
        Call .WriteInteger(map)

    End With

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
    Call outgoingData.WriteByte(ClientPacketID.WarpMeToTarget)

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

End Sub

''
' Writes the "Silence" message to the outgoing data buffer.
'
' @param    username The user to silence.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSilence(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Silence" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Silence)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

Public Sub WriteCuentaRegresiva(ByVal Second As Byte)

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

End Sub

Public Sub WritePossUser(ByVal UserName As String)
    '***************************************************
    'Write by Ladder
    '03-12-08
    'Guarda la posición donde estamos parados, como la posición del personaje.
    'Esta pensado exclusivamente para deslogear PJs.
    '***************************************************

    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.PossUser)
        Call .WriteASCIIString(UserName)

    End With

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
    Call outgoingData.WriteByte(ClientPacketID.SOSShowList)

End Sub

''
' Writes the "SOSRemove" message to the outgoing data buffer.
'
' @param    username The user whose SOS call has been already attended.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSOSRemove(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SOSRemove" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.SOSRemove)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "GoToChar" message to the outgoing data buffer.
'
' @param    username The user to be approached.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGoToChar(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GoToChar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GoToChar)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

Public Sub WriteDesbuggear(ByVal Params As String)

    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Desbuggear)
        Call .WriteASCIIString(Params)

    End With

End Sub

Public Sub WriteDarLlaveAUsuario(ByVal User As String, ByVal Llave As Integer)

    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.DarLlaveAUsuario)
        Call .WriteASCIIString(User)
        Call .WriteInteger(Llave)
    End With

End Sub

Public Sub WriteSacarLlave(ByVal Llave As Integer)

    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.SacarLlave)
        Call .WriteInteger(Llave)
    End With

End Sub

Public Sub WriteVerLlaves()

    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.VerLlaves)
    End With

End Sub

''
' Writes the "invisible" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteInvisible()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "invisible" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.invisible)

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
    Call outgoingData.WriteByte(ClientPacketID.GMPanel)

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
    Call outgoingData.WriteByte(ClientPacketID.RequestUserList)

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
    Call outgoingData.WriteByte(ClientPacketID.Working)

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
    Call outgoingData.WriteByte(ClientPacketID.Hiding)

End Sub

''
' Writes the "Jail" message to the outgoing data buffer.
'
' @param    username The user to be sent to jail.
' @param    reason The reason for which to send him to jail.
' @param    time The time (in minutes) the user will have to spend there.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteJail(ByVal UserName As String, ByVal reason As String, ByVal time As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Jail" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Jail)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(reason)
        
        Call .WriteByte(time)

    End With

End Sub

Public Sub WriteCrearEvento(ByVal TIPO As Byte, ByVal duracion As Byte, ByVal multiplicacion As Byte)

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
    Call outgoingData.WriteByte(ClientPacketID.KillNPC)

End Sub

''
' Writes the "WarnUser" message to the outgoing data buffer.
'
' @param    username The user to be warned.
' @param    reason Reason for the warning.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteWarnUser(ByVal UserName As String, ByVal reason As String)

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

End Sub

Public Sub WriteMensajeUser(ByVal UserName As String, ByVal mensaje As String)

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

End Sub

''
' Writes the "RequestCharInfo" message to the outgoing data buffer.
'
' @param    username The user whose information is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharInfo(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharInfo" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RequestCharInfo)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "RequestCharStats" message to the outgoing data buffer.
'
' @param    username The user whose stats are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharStats(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharStats" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RequestCharStats)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "RequestCharGold" message to the outgoing data buffer.
'
' @param    username The user whose gold is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharGold(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharGold" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RequestCharGold)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub
    
''
' Writes the "RequestCharInventory" message to the outgoing data buffer.
'
' @param    username The user whose inventory is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharInventory(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharInventory" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RequestCharInventory)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "RequestCharBank" message to the outgoing data buffer.
'
' @param    username The user whose banking information is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharBank(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharBank" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RequestCharBank)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "RequestCharSkills" message to the outgoing data buffer.
'
' @param    username The user whose skills are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharSkills(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharSkills" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RequestCharSkills)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "ReviveChar" message to the outgoing data buffer.
'
' @param    username The user to eb revived.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteReviveChar(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ReviveChar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ReviveChar)
        
        Call .WriteASCIIString(UserName)

    End With

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
    Call outgoingData.WriteByte(ClientPacketID.OnlineGM)

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
    Call outgoingData.WriteByte(ClientPacketID.OnlineMap)

End Sub

''
' Writes the "Forgive" message to the outgoing data buffer.
'
' @param    username The user to be forgiven.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForgive()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Forgive" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Forgive)
        
        '  Call .WriteASCIIString(UserName)
    End With

End Sub

''
' Writes the "Kick" message to the outgoing data buffer.
'
' @param    username The user to be kicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteKick(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Kick" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Kick)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "Execute" message to the outgoing data buffer.
'
' @param    username The user to be executed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteExecute(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Execute" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.Execute)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "BanChar" message to the outgoing data buffer.
'
' @param    username The user to be banned.
' @param    reason The reson for which the user is to be banned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteBanChar(ByVal UserName As String, ByVal reason As String)

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

End Sub

Public Sub WriteBanCuenta(ByVal UserName As String, ByVal reason As String)

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

End Sub

Public Sub WriteUnBanCuenta(ByVal UserName As String)

    '***************************************************
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.UnbanCuenta)
        Call .WriteASCIIString(UserName)

    End With

End Sub

Public Sub WriteBanSerial(ByVal UserName As String)

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

End Sub

Public Sub WriteUnBanSerial(ByVal UserName As String, ByVal reason As String)

    '***************************************************
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.UnBanSerial)
        Call .WriteASCIIString(UserName)

    End With

End Sub

Public Sub WriteCerraCliente(ByVal UserName As String)

    '***************************************************
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.CerrarCliente)
        Call .WriteASCIIString(UserName)

    End With

End Sub

Public Sub WriteBanTemporal(ByVal UserName As String, ByVal reason As String, ByVal dias As Byte)

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

End Sub

Public Sub WriteSilenciarUser(ByVal UserName As String, ByVal time As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BanChar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.SilenciarUser)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteByte(time)

    End With

End Sub

''
' Writes the "UnbanChar" message to the outgoing data buffer.
'
' @param    username The user to be unbanned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUnbanChar(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "UnbanChar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.UnbanChar)
        
        Call .WriteASCIIString(UserName)

    End With

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
    Call outgoingData.WriteByte(ClientPacketID.NPCFollow)

End Sub

''
' Writes the "SummonChar" message to the outgoing data buffer.
'
' @param    username The user to be summoned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSummonChar(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SummonChar" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.SummonChar)
        
        Call .WriteASCIIString(UserName)

    End With

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
    Call outgoingData.WriteByte(ClientPacketID.SpawnListRequest)

End Sub

''
' Writes the "SpawnCreature" message to the outgoing data buffer.
'
' @param    creatureIndex The index of the creature in the spawn list to be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSpawnCreature(ByVal creatureIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SpawnCreature" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.SpawnCreature)
        
        Call .WriteInteger(creatureIndex)

    End With

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
    Call outgoingData.WriteByte(ClientPacketID.ResetNPCInventory)

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
    Call outgoingData.WriteByte(ClientPacketID.CleanWorld)

End Sub

''
' Writes the "ServerMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteServerMessage(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ServerMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ServerMessage)
        
        Call .WriteASCIIString(Message)

    End With

End Sub

''
' Writes the "NickToIP" message to the outgoing data buffer.
'
' @param    username The user whose IP is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteNickToIP(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "NickToIP" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.NickToIP)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "IPToNick" message to the outgoing data buffer.
'
' @param    IP The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteIPToNick(ByRef IP() As Byte)

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

End Sub

''
' Writes the "GuildOnlineMembers" message to the outgoing data buffer.
'
' @param    guild The guild whose online player list is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildOnlineMembers(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildOnlineMembers" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildOnlineMembers)
        
        Call .WriteASCIIString(guild)

    End With

End Sub

''
' Writes the "TeleportCreate" message to the outgoing data buffer.
'
' @param    map the map to which the teleport will lead.
' @param    x The position in the x axis to which the teleport will lead.
' @param    y The position in the y axis to which the teleport will lead.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTeleportCreate(ByVal map As Integer, ByVal x As Byte, ByVal y As Byte)

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
    Call outgoingData.WriteByte(ClientPacketID.TeleportDestroy)

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
    Call outgoingData.WriteByte(ClientPacketID.RainToggle)

End Sub

''
' Writes the "SetCharDescription" message to the outgoing data buffer.
'
' @param    desc The description to set to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetCharDescription(ByVal desc As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SetCharDescription" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.SetCharDescription)
        
        Call .WriteASCIIString(desc)

    End With

End Sub

''
' Writes the "ForceMIDIToMap" message to the outgoing data buffer.
'
' @param    midiID The ID of the midi file to play.
' @param    map The map in which to play the given midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceMIDIToMap(ByVal midiID As Byte, ByVal map As Integer)

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

End Sub

''
' Writes the "RoyalArmyMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the royal army members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoyalArmyMessage(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RoyalArmyMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RoyalArmyMessage)
        
        Call .WriteASCIIString(Message)

    End With

End Sub

''
' Writes the "ChaosLegionMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the chaos legion member.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosLegionMessage(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChaosLegionMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChaosLegionMessage)
        
        Call .WriteASCIIString(Message)

    End With

End Sub

''
' Writes the "CitizenMessage" message to the outgoing data buffer.
'
' @param    message The message to send to citizens.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCitizenMessage(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CitizenMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CitizenMessage)
        
        Call .WriteASCIIString(Message)

    End With

End Sub

''
' Writes the "CriminalMessage" message to the outgoing data buffer.
'
' @param    message The message to send to criminals.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCriminalMessage(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CriminalMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CriminalMessage)
        
        Call .WriteASCIIString(Message)

    End With

End Sub

''
' Writes the "TalkAsNPC" message to the outgoing data buffer.
'
' @param    message The message to send to the royal army members.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTalkAsNPC(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TalkAsNPC" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.TalkAsNPC)
        
        Call .WriteASCIIString(Message)

    End With

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
    Call outgoingData.WriteByte(ClientPacketID.DestroyAllItemsInArea)

End Sub

''
' Writes the "AcceptRoyalCouncilMember" message to the outgoing data buffer.
'
' @param    username The name of the user to be accepted into the royal army council.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAcceptRoyalCouncilMember(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AcceptRoyalCouncilMember" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.AcceptRoyalCouncilMember)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "AcceptChaosCouncilMember" message to the outgoing data buffer.
'
' @param    username The name of the user to be accepted as a chaos council member.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAcceptChaosCouncilMember(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "AcceptChaosCouncilMember" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.AcceptChaosCouncilMember)
        
        Call .WriteASCIIString(UserName)

    End With

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
    Call outgoingData.WriteByte(ClientPacketID.ItemsInTheFloor)

End Sub

''
' Writes the "MakeDumb" message to the outgoing data buffer.
'
' @param    username The name of the user to be made dumb.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMakeDumb(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "MakeDumb" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MakeDumb)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "MakeDumbNoMore" message to the outgoing data buffer.
'
' @param    username The name of the user who will no longer be dumb.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteMakeDumbNoMore(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "MakeDumbNoMore" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.MakeDumbNoMore)
        
        Call .WriteASCIIString(UserName)

    End With

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
    Call outgoingData.WriteByte(ClientPacketID.DumpIPTables)

End Sub

''
' Writes the "CouncilKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the council.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCouncilKick(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CouncilKick" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CouncilKick)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "SetTrigger" message to the outgoing data buffer.
'
' @param    trigger The type of trigger to be set to the tile.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetTrigger(ByVal Trigger As eTrigger)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SetTrigger" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.SetTrigger)
        
        Call .WriteByte(Trigger)

    End With

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
    Call outgoingData.WriteByte(ClientPacketID.AskTrigger)

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
    Call outgoingData.WriteByte(ClientPacketID.BannedIPList)

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
    Call outgoingData.WriteByte(ClientPacketID.BannedIPReload)

End Sub

''
' Writes the "GuildBan" message to the outgoing data buffer.
'
' @param    guild The guild whose members will be banned.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGuildBan(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildBan" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GuildBan)
        
        Call .WriteASCIIString(guild)

    End With

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

End Sub

''
' Writes the "UnbanIP" message to the outgoing data buffer.
'
' @param    IP The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteUnbanIP(ByRef IP() As Byte)

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

End Sub

''
' Writes the "CreateItem" message to the outgoing data buffer.
'
' @param    itemIndex The index of the item to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateItem(ByVal ItemIndex As Long, ByVal cantidad As Integer)

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
    Call outgoingData.WriteByte(ClientPacketID.DestroyItems)

End Sub

''
' Writes the "ChaosLegionKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the Chaos Legion.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosLegionKick(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChaosLegionKick" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChaosLegionKick)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "RoyalArmyKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the Royal Army.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRoyalArmyKick(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RoyalArmyKick" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RoyalArmyKick)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "ForceMIDIAll" message to the outgoing data buffer.
'
' @param    midiID The id of the midi file to play.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceMIDIAll(ByVal midiID As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ForceMIDIAll" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ForceMIDIAll)
        
        Call .WriteByte(midiID)

    End With

End Sub

''
' Writes the "ForceWAVEAll" message to the outgoing data buffer.
'
' @param    waveID The id of the wave file to play.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteForceWAVEAll(ByVal waveID As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ForceWAVEAll" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ForceWAVEAll)
        
        Call .WriteByte(waveID)

    End With

End Sub

''
' Writes the "RemovePunishment" message to the outgoing data buffer.
'
' @param    username The user whose punishments will be altered.
' @param    punishment The id of the punishment to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemovePunishment(ByVal UserName As String, ByVal punishment As Byte, ByVal NewText As String)

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
    Call outgoingData.WriteByte(ClientPacketID.TileBlockedToggle)

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
    Call outgoingData.WriteByte(ClientPacketID.KillNPCNoRespawn)

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
    Call outgoingData.WriteByte(ClientPacketID.KillAllNearbyNPCs)

End Sub

''
' Writes the "LastIP" message to the outgoing data buffer.
'
' @param    username The user whose last IPs are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteLastIP(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "LastIP" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.LastIP)
        
        Call .WriteASCIIString(UserName)

    End With

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
    Call outgoingData.WriteByte(ClientPacketID.ChangeMOTD)

End Sub

''
' Writes the "SetMOTD" message to the outgoing data buffer.
'
' @param    message The message to be set as the new MOTD.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSetMOTD(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SetMOTD" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.SetMOTD)
        
        Call .WriteASCIIString(Message)

    End With

End Sub

''
' Writes the "SystemMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to all players.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSystemMessage(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SystemMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.SystemMessage)
        
        Call .WriteASCIIString(Message)

    End With

End Sub

''
' Writes the "CreateNPC" message to the outgoing data buffer.
'
' @param    npcIndex The index of the NPC to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNPC(ByVal NpcIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CreateNPC" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CreateNPC)
        
        Call .WriteInteger(NpcIndex)

    End With

End Sub

''
' Writes the "CreateNPCWithRespawn" message to the outgoing data buffer.
'
' @param    npcIndex The index of the NPC to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCreateNPCWithRespawn(ByVal NpcIndex As Integer)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CreateNPCWithRespawn" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.CreateNPCWithRespawn)
        
        Call .WriteInteger(NpcIndex)

    End With

End Sub

''
' Writes the "ImperialArmour" message to the outgoing data buffer.
'
' @param    armourIndex The index of imperial armour to be altered.
' @param    objectIndex The index of the new object to be set as the imperial armour.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteImperialArmour(ByVal armourIndex As Byte, ByVal objectIndex As Integer)

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

End Sub

''
' Writes the "ChaosArmour" message to the outgoing data buffer.
'
' @param    armourIndex The index of chaos armour to be altered.
' @param    objectIndex The index of the new object to be set as the chaos armour.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChaosArmour(ByVal armourIndex As Byte, ByVal objectIndex As Integer)

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
    Call outgoingData.WriteByte(ClientPacketID.NavigateToggle)

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
    Call outgoingData.WriteByte(ClientPacketID.ServerOpenToUsersToggle)

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
    Call outgoingData.WriteByte(ClientPacketID.Participar)

End Sub

''
' Writes the "TurnCriminal" message to the outgoing data buffer.
'
' @param    username The name of the user to turn into criminal.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteTurnCriminal(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "TurnCriminal" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.TurnCriminal)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "ResetFactions" message to the outgoing data buffer.
'
' @param    username The name of the user who will be removed from any faction.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteResetFactions(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ResetFactions" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ResetFactions)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "RemoveCharFromGuild" message to the outgoing data buffer.
'
' @param    username The name of the user who will be removed from any guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRemoveCharFromGuild(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RemoveCharFromGuild" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RemoveCharFromGuild)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "RequestCharMail" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteRequestCharMail(ByVal UserName As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestCharMail" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.RequestCharMail)
        
        Call .WriteASCIIString(UserName)

    End With

End Sub

''
' Writes the "AlterPassword" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    copyFrom The name of the user from which to copy the password.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterPassword(ByVal UserName As String, ByVal CopyFrom As String)

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

End Sub

''
' Writes the "AlterMail" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    newMail The new email of the player.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterMail(ByVal UserName As String, ByVal newMail As String)

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

End Sub

''
' Writes the "AlterName" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    newName The new user name.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteAlterName(ByVal UserName As String, ByVal newName As String)

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

End Sub

''
' Writes the "ToggleCentinelActivated" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteToggleCentinelActivated()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ToggleCentinelActivated" message to the outgoing data buffer
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.ToggleCentinelActivated)

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
    Call outgoingData.WriteByte(ClientPacketID.DoBackUp)

End Sub

''
' Writes the "ShowGuildMessages" message to the outgoing data buffer.
'
' @param    guild The guild to listen to.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteShowGuildMessages(ByVal guild As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ShowGuildMessages" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ShowGuildMessages)
        
        Call .WriteASCIIString(guild)

    End With

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
    Call outgoingData.WriteByte(ClientPacketID.SaveMap)

End Sub

''
' Writes the "ChangeMapInfoPK" message to the outgoing data buffer.
'
' @param    isPK True if the map is PK, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoPK(ByVal isPK As Boolean)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeMapInfoPK" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeMapInfoPK)
        
        Call .WriteBoolean(isPK)

    End With

End Sub

''
' Writes the "ChangeMapInfoBackup" message to the outgoing data buffer.
'
' @param    backup True if the map is to be backuped, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoBackup(ByVal backup As Boolean)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChangeMapInfoBackup" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeMapInfoBackup)
        
        Call .WriteBoolean(backup)

    End With

End Sub

''
' Writes the "ChangeMapInfoRestricted" message to the outgoing data buffer.
'
' @param    restrict NEWBIES (only newbies), NO (everyone), ARMADA (just Armadas), CAOS (just caos) or FACCION (Armadas & caos only)
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoRestricted(ByVal restrict As String)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Writes the "ChangeMapInfoRestricted" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeMapInfoRestricted)
        
        Call .WriteASCIIString(restrict)

    End With

End Sub

''
' Writes the "ChangeMapInfoNoMagic" message to the outgoing data buffer.
'
' @param    nomagic TRUE if no magic is to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoMagic(ByVal nomagic As Boolean)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Writes the "ChangeMapInfoNoMagic" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeMapInfoNoMagic)
        
        Call .WriteBoolean(nomagic)

    End With

End Sub

''
' Writes the "ChangeMapInfoNoInvi" message to the outgoing data buffer.
'
' @param    noinvi TRUE if invisibility is not to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoInvi(ByVal noinvi As Boolean)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Writes the "ChangeMapInfoNoInvi" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeMapInfoNoInvi)
        
        Call .WriteBoolean(noinvi)

    End With

End Sub
                            
''
' Writes the "ChangeMapInfoNoResu" message to the outgoing data buffer.
'
' @param    noresu TRUE if resurection is not to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoNoResu(ByVal noresu As Boolean)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Writes the "ChangeMapInfoNoResu" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeMapInfoNoResu)
        
        Call .WriteBoolean(noresu)

    End With

End Sub
                        
''
' Writes the "ChangeMapInfoLand" message to the outgoing data buffer.
'
' @param    land options: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoLand(ByVal land As String)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Writes the "ChangeMapInfoLand" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeMapInfoLand)
        
        Call .WriteASCIIString(land)

    End With

End Sub
                        
''
' Writes the "ChangeMapInfoZone" message to the outgoing data buffer.
'
' @param    zone options: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChangeMapInfoZone(ByVal zone As String)

    '***************************************************
    'Author: Pablo (ToxicWaste)
    'Last Modification: 26/01/2007
    'Writes the "ChangeMapInfoZone" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChangeMapInfoZone)
        
        Call .WriteASCIIString(zone)

    End With

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
    Call outgoingData.WriteByte(ClientPacketID.SaveChars)

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
    Call outgoingData.WriteByte(ClientPacketID.CleanSOS)

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
    Call outgoingData.WriteByte(ClientPacketID.ShowServerForm)

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
    Call outgoingData.WriteByte(ClientPacketID.night)

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
    Call outgoingData.WriteByte(ClientPacketID.KickAllChars)

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
    Call outgoingData.WriteByte(ClientPacketID.RequestTCPStats)

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
    Call outgoingData.WriteByte(ClientPacketID.ReloadNPCs)

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
    Call outgoingData.WriteByte(ClientPacketID.ReloadServerIni)

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
    Call outgoingData.WriteByte(ClientPacketID.ReloadSpells)

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
    Call outgoingData.WriteByte(ClientPacketID.ReloadObjects)

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
    Call outgoingData.WriteByte(ClientPacketID.Restart)

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
    Call outgoingData.WriteByte(ClientPacketID.ResetAutoUpdate)

End Sub

''
' Writes the "ChatColor" message to the outgoing data buffer.
'
' @param    r The red component of the new chat color.
' @param    g The green component of the new chat color.
' @param    b The blue component of the new chat color.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteChatColor(ByVal r As Byte, ByVal g As Byte, ByVal b As Byte)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ChatColor" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.ChatColor)
        
        Call .WriteByte(r)
        Call .WriteByte(g)
        Call .WriteByte(b)

    End With

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
    Call outgoingData.WriteByte(ClientPacketID.Ignored)

End Sub

''
' Writes the "CheckSlot" message to the outgoing data buffer.
'
' @param    UserName    The name of the char whose slot will be checked.
' @param    slot        The slot to be checked.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteCheckSlot(ByVal UserName As String, ByVal Slot As Byte)

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

    Call outgoingData.WriteByte(ClientPacketID.Ping)
    pingTime = timeGetTime
    Call outgoingData.WriteLong(pingTime)
    
    ' Avoid computing errors due to frame rate
    Call FlushBuffer
    'DoEvents
    
End Sub

Public Sub WriteLlamadadeClan()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 26/01/2007
    'Writes the "Ping" message to the outgoing data buffer
    '***************************************************
    'Prevent the timer from being cut
    '   If pingTime <> 0 Then Exit Sub

    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.llamadadeclan)
    
    ' Avoid computing errors due to frame rate
    Call FlushBuffer
    'DoEvents
    
End Sub

''
' Flushes the outgoing data buffer of the user.
'
' @param    UserIndex User whose outgoing data buffer will be flushed.

Public Sub FlushBuffer()

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

End Sub

''
' Sends the data using the socket controls in the MainForm.
'
' @param    sdData  The data to be sent to the server.

Private Sub SendData(ByRef sdData As String)
    #If UsarWrench = 1 Then

        If Not frmmain.Socket1.IsWritable Then
            'Put data back in the bytequeue
            Call outgoingData.WriteASCIIStringFixed(sdData)
            Exit Sub

        End If
   
        If Not frmmain.Socket1.Connected Then Exit Sub
    #Else

        If frmmain.Winsock1.State <> sckConnected Then Exit Sub
    #End If
 
    #If AntiExternos Then

        Dim Data() As Byte

        Data = StrConv(sdData, vbFromUnicode)
        Security.NAC_E_Byte Data, Security.Redundance
        sdData = StrConv(Data, vbUnicode)
        'sdData = Security.NAC_E_String(sdData, Security.Redundance)
    #End If
 
    #If UsarWrench = 1 Then
        Call frmmain.Socket1.Write(sdData, Len(sdData))
    #Else
        Call frmmain.Winsock1.SendData(sdData)
    #End If
 
End Sub

Public Sub WriteQuestionGM(ByVal Consulta As String, ByVal TipoDeConsulta As String)

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

End Sub

Public Sub WriteOfertaInicial(ByVal Oferta As Long)

    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.OfertaInicial)
        Call .WriteLong(Oferta)

    End With

End Sub

Public Sub WriteOferta(ByVal OfertaDeSubasta As Long)

    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.OfertaDeSubasta)
        Call .WriteLong(OfertaDeSubasta)

    End With

End Sub

Public Sub WriteGlobalMessage(ByVal Message As String)

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRequestDetails" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.GlobalMessage)
        
        Call .WriteASCIIString(Message)
        
    End With

End Sub

Public Sub WriteGlobalOnOff()
    Call outgoingData.WriteByte(ClientPacketID.GlobalOnOff)

End Sub

Public Sub WriteNuevaCuenta()

    With outgoingData
        Call .WriteByte(ClientPacketID.CrearNuevaCuenta)
    
        Call .WriteASCIIString(CuentaEmail)
        Call .WriteASCIIString(SEncriptar(CuentaPassword))
        Call .WriteASCIIString(CuentaEmail)

    End With

End Sub

Public Sub WriteValidarCuenta()

    With outgoingData
        Call .WriteByte(ClientPacketID.validarCuenta)
        Call .WriteASCIIString(CuentaEmail)
        Call .WriteASCIIString(ValidacionCode)

    End With

End Sub

Public Sub WriteReValidarCuenta()

    With outgoingData
        Call .WriteByte(ClientPacketID.RevalidarCuenta)
        Call .WriteASCIIString(CuentaEmail)

    End With

End Sub

Public Sub WriteRecuperandoConstraseña()

    With outgoingData
        Call .WriteByte(ClientPacketID.RecuperandoConstraseña)
        Call .WriteASCIIString(CuentaEmail)

    End With

End Sub

Public Sub WriteBorrandoCuenta()

    With outgoingData
        Call .WriteByte(ClientPacketID.BorrandoCuenta)
        Call .WriteASCIIString(CuentaEmail)
        Call .WriteASCIIString(SEncriptar(CuentaPassword))

    End With

End Sub

Public Sub WriteBorrandoPJ()

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

    End With

End Sub

Public Sub WriteIngresandoConCuenta()

    With outgoingData
        Call .WriteByte(ClientPacketID.IngresarConCuenta)
        Call .WriteASCIIString(CuentaEmail)
        Call .WriteASCIIString(SEncriptar(CuentaPassword))
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
        Call .WriteASCIIString(MacAdress)  'Seguridad
        Call .WriteLong(HDserial)  'SeguridadHDserial
        
    End With

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
        Pjs(ii).mapa = 0
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
        Pjs(ii).mapa = buffer.ReadInteger()
        Pjs(ii).Body = buffer.ReadInteger()
        
        Pjs(ii).Head = buffer.ReadInteger()
        Pjs(ii).Criminal = buffer.ReadByte()
        Pjs(ii).Clase = buffer.ReadByte()
       
        Pjs(ii).Casco = buffer.ReadInteger()
        Pjs(ii).Escudo = buffer.ReadInteger()
        Pjs(ii).Arma = buffer.ReadInteger()
        Pjs(ii).ClanName = "<" & buffer.ReadASCIIString() & ">"
       
        ' Pjs(ii).NameMapa = Pjs(ii).mapa
        Pjs(ii).NameMapa = NameMaps(Pjs(ii).mapa).name

    Next ii
    
    CuentaDonador = buffer.ReadByte()
    
    Dim i As Integer

    For i = 1 To CantidadDePersonajesEnCuenta

        Select Case Pjs(i).Criminal

            Case 0 'Criminal
                Pjs(i).LetraColor = RGB(ColoresPJ(50).r, ColoresPJ(50).g, ColoresPJ(50).b)
                Pjs(i).priv = 0

            Case 1 'Ciudadano
                Pjs(i).LetraColor = RGB(ColoresPJ(49).r, ColoresPJ(49).g, ColoresPJ(49).b)
                Pjs(i).priv = 0

            Case 2 'Caos
                Pjs(i).LetraColor = RGB(179, 0, 4)
                Pjs(i).priv = 0

            Case 3 'Armada
                Pjs(i).LetraColor = RGB(31, 139, 139)
                Pjs(i).priv = 0

            Case 4 'EsConsejero
                Pjs(i).LetraColor = RGB(2, 161, 38)
                Pjs(i).ClanName = "<Game Design>"
                Pjs(i).priv = 1
                EsGM = True

            Case 5 ' EsSemiDios
                Pjs(i).LetraColor = RGB(2, 161, 38)
                Pjs(i).ClanName = "<Game Master>"
                Pjs(i).priv = 2
                EsGM = True

            Case 6 ' EsDios
                Pjs(i).LetraColor = RGB(217, 164, 32)
                Pjs(i).ClanName = "<Administrador>"
                Pjs(i).priv = 3
                EsGM = True

            Case 7 ' EsAdmin
                Pjs(i).LetraColor = RGB(217, 164, 32)
                Pjs(i).ClanName = "<Administrador>"
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

    Dim Error As Long

    Error = Err.number

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
    frmmain.onlines = "Onlines: " & usersOnline
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    Dim Error As Long

    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

Private Sub HandleParticleFXToFloor()

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

    Dim time           As Long

    Dim Borrar         As Boolean
     
    x = incomingData.ReadByte()
    y = incomingData.ReadByte()
    ParticulaIndex = incomingData.ReadInteger()
    time = incomingData.ReadLong()

    If time = 1 Then
        time = -1

    End If
    
    If time = 0 Then
        Borrar = True

    End If

    If Borrar Then
        Graficos_Particulas.Particle_Group_Remove (MapData(x, y).particle_group)
    Else

        If MapData(x, y).particle_group = 0 Then
            MapData(x, y).particle_group = 0
            General_Particle_Create ParticulaIndex, x, y, time
        Else
            Call General_Char_Particle_Create(ParticulaIndex, MapData(x, y).charindex, time)

        End If

    End If

End Sub

Private Sub HandleLightToFloor()

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

    Dim color As Long

    Dim Rango As Byte
     
    x = incomingData.ReadByte()
    y = incomingData.ReadByte()
    color = incomingData.ReadLong()
    Rango = incomingData.ReadByte()

    Dim id  As Long

    Dim id2 As Long

    If color = 0 Then
   
        If MapData(x, y).luz.Rango > 100 Then
            LucesRedondas.Delete_Light_To_Map x, y
   
            LucesCuadradas.Light_Render_All
            LucesRedondas.LightRenderAll
            Exit Sub
        Else
            id = LucesCuadradas.Light_Find(x & y)
            LucesCuadradas.Light_Remove id
            MapData(x, y).luz.color = color
            MapData(x, y).luz.Rango = 0
            LucesCuadradas.Light_Render_All
            Exit Sub

        End If

    End If
    
    MapData(x, y).luz.color = color
    MapData(x, y).luz.Rango = Rango
    
    If Rango < 100 Then
        id = x & y
        LucesCuadradas.Light_Create x, y, color, Rango, id
        LucesCuadradas.Light_Render_All
    Else

        Dim r, g, b As Byte

        b = (color And 16711680) / 65536
        g = (color And 65280) / 256
        r = color And 255
        LucesRedondas.Create_Light_To_Map x, y, Rango - 99, b, g, r

    End If

End Sub

Private Sub HandleParticleFX()

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

    Dim time           As Long

    Dim Remove         As Boolean
     
    charindex = incomingData.ReadInteger()
    ParticulaIndex = incomingData.ReadInteger()
    time = incomingData.ReadLong()
    Remove = incomingData.ReadBoolean()
    
    If Remove Then
        Call Char_Particle_Group_Remove(charindex, ParticulaIndex)
        charlist(charindex).Particula = 0
    
    Else
        charlist(charindex).Particula = ParticulaIndex
        charlist(charindex).ParticulaTime = time
     
        Call General_Char_Particle_Create(ParticulaIndex, charindex, time)

    End If
    
End Sub

Private Sub HandleParticleFXWithDestino()

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

    Dim time           As Long

    Dim wav            As Integer

    Dim fX             As Integer
     
    Emisor = incomingData.ReadInteger()
    receptor = incomingData.ReadInteger()
    ParticulaViaje = incomingData.ReadInteger()
    ParticulaFinal = incomingData.ReadInteger()

    time = incomingData.ReadLong()
    wav = incomingData.ReadInteger()
    fX = incomingData.ReadInteger()

    Engine_spell_Particle_Set (ParticulaViaje)

    Call Effect_Begin(ParticulaViaje, 9, Get_Pixelx_Of_Char(Emisor), Get_PixelY_Of_Char(Emisor), ParticulaFinal, time, receptor, Emisor, wav, fX)

    ' charlist(charindex).Particula = ParticulaIndex
    ' charlist(charindex).ParticulaTime = time

    ' Call General_Char_Particle_Create(ParticulaIndex, charindex, time)
    
End Sub

Private Sub HandleParticleFXWithDestinoXY()

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

    Dim time           As Long

    Dim wav            As Integer

    Dim fX             As Integer

    Dim x              As Byte

    Dim y              As Byte
     
    Emisor = incomingData.ReadInteger()
    ParticulaViaje = incomingData.ReadInteger()
    ParticulaFinal = incomingData.ReadInteger()

    time = incomingData.ReadLong()
    wav = incomingData.ReadInteger()
    fX = incomingData.ReadInteger()
    
    x = incomingData.ReadByte()
    y = incomingData.ReadByte()
    
    ' Debug.Print "RECIBI FX= " & fX

    Engine_spell_Particle_Set (ParticulaViaje)

    Call Effect_BeginXY(ParticulaViaje, 9, Get_Pixelx_Of_Char(Emisor), Get_PixelY_Of_Char(Emisor), x, y, ParticulaFinal, time, Emisor, wav, fX)

    ' charlist(charindex).Particula = ParticulaIndex
    ' charlist(charindex).ParticulaTime = time

    ' Call General_Char_Particle_Create(ParticulaIndex, charindex, time)
    
End Sub

Private Sub HandleAuraToChar()

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
      
    End If

End Sub

Private Sub HandleSpeedToChar()

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

End Sub

Public Sub WriteDuelo()
    '***************************************************
    '***************************************************

    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.Duelo)

    End With
    
End Sub

Public Sub WriteNieveToggle()

    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.NieveToggle)

    End With

End Sub

Private Sub HandleNieveToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    Call incomingData.ReadByte
    
    If Not InMapBounds(UserPos.x, UserPos.y) Then Exit Sub
    bTecho = (MapData(UserPos.x, UserPos.y).Trigger = 1 Or MapData(UserPos.x, UserPos.y).Trigger = 2 Or MapData(UserPos.x, UserPos.y).Trigger > 9 Or MapData(UserPos.x, UserPos.y).Trigger = 6 Or MapData(UserPos.x, UserPos.y).Trigger = 4)
            
    If MapDat.NIEVE Then
        Engine_Meteo_Particle_Set (Particula_Nieve)

    End If

    bNieve = Not bNieve
  
End Sub

Private Sub HandleNieblaToggle()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    Call incomingData.ReadByte
    
    MaxAlphaNiebla = incomingData.ReadByte()
            
    bNiebla = Not bNiebla
    frmmain.TimerNiebla.Enabled = True
    
    Call Meteo_Engine.CargarClima
  
End Sub

Public Sub WriteNieblaToggle()

    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.NieblaToggle)

    End With

End Sub

Public Sub WriteGenio()
    '***************************************************
    '/GENIO
    'Ladder
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.Genio)

End Sub

Private Sub HandleFamiliar()

    If incomingData.length < 1 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
    
End Sub

Private Sub HandleBindKeys()

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
        frmmain.CombateIcon.Picture = LoadInterface("infoapretado.bmp")
    Else
        frmmain.CombateIcon.Picture = LoadInterface("info.bmp")

    End If

    If ChatGlobal = 1 Then
        frmmain.globalIcon.Picture = LoadInterface("globalapretado.bmp")
    Else
        frmmain.CombateIcon.Picture = LoadInterface("global.bmp")

    End If

End Sub

Private Sub HandleLogros()

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
    
    FrmLogros.Show , frmmain

End Sub

Private Sub HandleBarFx()

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

End Sub

Public Sub WriteCompletarAccion(ByVal Accion As Byte)

    '***************************************************
    'Author: Pablo Mercavides
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.CompletarAccion)
        Call .WriteByte(Accion)

    End With

End Sub

Public Sub WriteTraerRecompensas()

    '***************************************************
    'Ladder
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.TraerRecompensas)

    End With

End Sub

Public Sub WriteTraerShop()

    '***************************************************
    'Ladder
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.Traershop)

    End With

End Sub

Public Sub WriteTraerRanking()

    '***************************************************
    'Ladder
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.TraerRanking)

    End With

End Sub

Public Sub WritePareja()

    '***************************************************
    'Ladder
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.Pareja)

    End With

End Sub

Public Sub WriteQuest()
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Escribe el paquete Quest al servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.Quest)

End Sub
 
Public Sub WriteQuestDetailsRequest(ByVal QuestSlot As Byte)
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Escribe el paquete QuestDetailsRequest al servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.QuestDetailsRequest)
    Call outgoingData.WriteByte(QuestSlot)

End Sub
 
Public Sub WriteQuestAccept()
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Escribe el paquete QuestAccept al servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.QuestAccept)

End Sub
 
Private Sub HandleQuestDetails()

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Recibe y maneja el paquete QuestDetails del servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    If incomingData.length < 13 Then
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
    
    Dim questindex    As Integer
    
    FrmQuests.ListView2.ListItems.Clear
    FrmQuests.ListView1.ListItems.Clear
    FrmQuestInfo.ListView2.ListItems.Clear
    FrmQuestInfo.ListView1.ListItems.Clear
    
    With buffer
        'Leemos el id del paquete
        Call .ReadByte
        
        'Nos fijamos si se trata de una quest empezada, para poder leer los NPCs que se han matado.
        QuestEmpezada = IIf(.ReadByte, True, False)
        
        If Not QuestEmpezada Then
        
            questindex = .ReadInteger
        
            FrmQuestInfo.titulo.Caption = Quest_Name(questindex)
           
            'tmpStr = "Mision: " & .ReadASCIIString & vbCrLf
           
            FrmQuestInfo.detalle.Caption = Quest_Desc(questindex) & vbCrLf & "Nivel requerido: " & .ReadByte & vbCrLf
            'tmpStr = tmpStr & "Detalles: " & .ReadASCIIString & vbCrLf
            'tmpStr = tmpStr & "Nivel requerido: " & .ReadByte & vbCrLf
           
            tmpStr = tmpStr & vbCrLf & "OBJETIVOS" & vbCrLf
           
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

                        Set subelemento = FrmQuestInfo.ListView1.ListItems.Add(, , NpcData(NpcIndex).name)
                       
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
                   
                    Set subelemento = FrmQuestInfo.ListView1.ListItems.Add(, , ObjData(OBJIndex).name)
                    subelemento.SubItems(1) = cantidadobj
                    subelemento.SubItems(2) = OBJIndex
                    subelemento.SubItems(3) = 1
                Next i

            End If
    
            tmpStr = tmpStr & vbCrLf & "RECOMPENSAS" & vbCrLf
            'tmpStr = tmpStr & "*) Oro: " & .ReadLong & " monedas de oro." & vbCrLf
            'tmpStr = tmpStr & "*) Experiencia: " & .ReadLong & " puntos de experiencia." & vbCrLf
           
            Set subelemento = FrmQuestInfo.ListView2.ListItems.Add(, , "Oro")
                       
            subelemento.SubItems(1) = .ReadLong
            subelemento.SubItems(2) = 12
            subelemento.SubItems(3) = 0
           
            Set subelemento = FrmQuestInfo.ListView2.ListItems.Add(, , "Experiencia")
                       
            subelemento.SubItems(1) = .ReadLong
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
                   
                    Set subelemento = FrmQuestInfo.ListView2.ListItems.Add(, , ObjData(obindex).name)
                       
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
        
            questindex = .ReadInteger
        
            FrmQuests.titulo.Caption = Quest_Name(questindex)
           
            'tmpStr = "Mision: " & .ReadASCIIString & vbCrLf
           
            FrmQuests.detalle.Caption = Quest_Desc(questindex) & vbCrLf & "Nivel requerido: " & .ReadByte & vbCrLf
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
                                     
                    Set subelemento = FrmQuests.ListView1.ListItems.Add(, , NpcData(NpcIndex).name)
                       
                    Dim cantok As Integer

                    cantok = cantidadnpc - matados
                       
                    If cantok = 0 Then
                        subelemento.SubItems(1) = "OK"
                    Else
                        subelemento.SubItems(1) = cantok

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
                   
                    Set subelemento = FrmQuests.ListView1.ListItems.Add(, , ObjData(OBJIndex).name)
                    subelemento.SubItems(1) = cantidadobj
                    subelemento.SubItems(2) = OBJIndex
                    subelemento.SubItems(3) = 1
                Next i

            End If
    
            tmpStr = tmpStr & vbCrLf & "RECOMPENSAS" & vbCrLf
            'tmpStr = tmpStr & "*) Oro: " & .ReadLong & " monedas de oro." & vbCrLf
            'tmpStr = tmpStr & "*) Experiencia: " & .ReadLong & " puntos de experiencia." & vbCrLf
           
            Set subelemento = FrmQuests.ListView2.ListItems.Add(, , "Oro")
                       
            subelemento.SubItems(1) = .ReadLong
            subelemento.SubItems(2) = 12
            subelemento.SubItems(3) = 0
           
            Set subelemento = FrmQuests.ListView2.ListItems.Add(, , "Experiencia")
                       
            subelemento.SubItems(1) = .ReadLong
            subelemento.SubItems(2) = 608
            subelemento.SubItems(3) = 1
           
            tmpByte = .ReadByte

            If tmpByte Then

                For i = 1 To tmpByte
                    'tmpStr = tmpStr & "*) " & .ReadInteger & " " & .ReadInteger & vbCrLf
                   
                    cantidadobjs = .ReadInteger
                    obindex = .ReadInteger
                   
                    Set subelemento = FrmQuests.ListView2.ListItems.Add(, , ObjData(obindex).name)
                       
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
        FrmQuestInfo.Show vbModeless, frmmain
        FrmQuestInfo.Picture = LoadInterface("mision.bmp")
        Call FrmQuestInfo.ListView1_Click
        Call FrmQuestInfo.ListView2_Click

    End If
    
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    Dim Error As Long

    Error = Err.number

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
    FrmQuests.Picture = LoadInterface("encurso.bmp")
    FrmQuests.Show vbModeless, frmmain
    
    'Pedimos la informaciï¿½n de la primer quest (si la hay)
    If tmpByte Then Call Protocol.WriteQuestDetailsRequest(1)
    
    'Copiamos de vuelta el buffer
    Call incomingData.CopyBuffer(buffer)
 
errhandler:

    Dim Error As Long

    Error = Err.number

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
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.QuestListRequest)

End Sub
 
Public Sub WriteQuestAbandon(ByVal QuestSlot As Byte)
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Escribe el paquete QuestAbandon al servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Escribe el ID del paquete.
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.QuestAbandon)
    
    'Escribe el Slot de Quest.
    Call outgoingData.WriteByte(QuestSlot)

End Sub

Public Sub WriteDecimeLaHora()
    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 22/11/2017
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.DecimeLaHora)

End Sub

Public Sub WriteResponderPregunta(ByVal Respuesta As Boolean)
    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 22/11/2017
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.ResponderPregunta)
    Call outgoingData.WriteBoolean(Respuesta)

End Sub

Public Sub WriteCorreo()
    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 22/11/2017
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.Correo)

End Sub

Public Sub WriteSendCorreo(ByVal UserNick As String, ByVal msg As String, ByVal ItemCount As Byte)
    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 4/5/2020
    '***************************************************
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
    
End Sub

Public Sub WriteComprarItem(ByVal ItemIndex As Byte)
    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 22/11/2017
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.ComprarItem)
    Call outgoingData.WriteByte(ItemIndex)
    
End Sub

Public Sub WriteCompletarViaje(ByVal destino As Byte, ByVal costo As Long)
    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 22/11/2017
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.CompletarViaje)
    Call outgoingData.WriteByte(destino)
    Call outgoingData.WriteLong(costo)
    
End Sub

Public Sub WriteRetirarItemCorreo(ByVal IndexMsg As Integer)
    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 22/11/2017
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.RetirarItemCorreo)
    Call outgoingData.WriteInteger(IndexMsg)

End Sub

Public Sub WriteBorrarCorreo(ByVal IndexMsg As Integer)
    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 22/11/2017
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.BorrarCorreo)
    Call outgoingData.WriteInteger(IndexMsg)

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
    FrmCorreo.txMensaje.Text = ""
    FrmCorreo.lbFecha.Caption = ""
    FrmCorreo.lbItem.Caption = ""

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
    
    Call FrmCorreo.lstInv.Clear

    'Fill the inventory list
    For i = 1 To MAX_INVENTORY_SLOTS

        If frmmain.Inventario.OBJIndex(i) <> 0 Then
            FrmCorreo.lstInv.AddItem frmmain.Inventario.ItemName(i)
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

        HayFormularioAbierto = True
        FrmCorreo.Show , frmmain
        
    End If
    
    'chat = Buffer.ReadASCIIString()
    '  fontIndex = Buffer.ReadByte()
    
    frmmain.PicCorreo.Visible = False
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    Dim Error As Long

    Error = Err.number

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

    Dim Error As Long

    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

Private Sub HandleDatosGrupo()

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

    FrmGrupo.Picture = LoadInterface("grupo.bmp")
    FrmGrupo.Show , frmmain
    HayFormularioAbierto = True

End Sub

Private Sub HandleUbicacion()

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
        frmmain.personaje(miembro).Visible = False
    Else

        If UserMap = map Then
            frmmain.personaje(miembro).Visible = True
            frmmain.personaje(miembro).Left = x - 4
            frmmain.personaje(miembro).Top = y - 2

        End If

    End If

End Sub

Private Sub HandleViajarForm()

    If incomingData.length < 4 Then
        Err.Raise incomingData.NotEnoughDataErrCode
        Exit Sub

    End If
    
    'Remove packet ID
    Call incomingData.ReadByte
            
    Dim dest     As String

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
    HayFormularioAbierto = True
        
    If ViajarInterface = 1 Then
        FrmViajes.Image1.Top = 4690
        FrmViajes.Image1.Left = 3810
    Else
        FrmViajes.Image1.Top = 4680
        FrmViajes.Image1.Left = 3840

    End If

    FrmViajes.Show , frmmain

End Sub

Private Sub HandleActShop()

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
        tmp = ObjData(Obj).name           'Get the object's name
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
    FrmShop.Show , frmmain
    HayFormularioAbierto = True
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    Dim Error As Long

    Error = Err.number

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
    
    HayFormularioAbierto = True
    
    FrmRanking.Picture = LoadInterface("ranking.bmp")
    FrmRanking.Show , frmmain
    
    'If we got here then packet is complete, copy data back to original queue
    Call incomingData.CopyBuffer(buffer)
    
errhandler:

    Dim Error As Long

    Error = Err.number

    On Error GoTo 0
    
    'Destroy auxiliar buffer
    Set buffer = Nothing

    If Error <> 0 Then Err.Raise Error

End Sub

''
' Handles the RestOK message.

Public Sub WriteCodigo(ByVal Codigo As String)

    '***************************************************
    'Ladder
    '***************************************************
    With outgoingData
        Call .WriteByte(ClientPacketID.newPacketID)
        Call .WriteByte(NewPacksID.EnviarCodigo)
        Call .WriteASCIIString(Codigo)

    End With

End Sub

Public Sub WriteCreaerTorneo(ByVal nivelminimo As Byte, ByVal nivelmaximo As Byte, ByVal cupos As Byte, ByVal costo As Long, ByVal mago As Byte, ByVal clerico As Byte, ByVal guerrero As Byte, ByVal asesino As Byte, ByVal bardo As Byte, ByVal druido As Byte, ByVal paladin As Byte, ByVal cazador As Byte, ByVal Trabajador As Byte, ByVal map As Integer, ByVal x As Byte, ByVal y As Byte, ByVal name As String, ByVal reglas As String)
    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 16/05/2020
    '***************************************************
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
    Call outgoingData.WriteASCIIString(name)
    Call outgoingData.WriteASCIIString(reglas)
     
End Sub

Public Sub WriteComenzarTorneo()
    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 16/05/2020
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.ComenzarTorneo)
     
End Sub

Public Sub WriteCancelarTorneo()
    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 16/05/2020
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.CancelarTorneo)
     
End Sub

Public Sub WriteBusquedaTesoro(ByVal TIPO As Byte)
    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 16/05/2020
    '***************************************************
    Call outgoingData.WriteByte(ClientPacketID.newPacketID)
    Call outgoingData.WriteByte(NewPacksID.BusquedaTesoro)
    Call outgoingData.WriteByte(TIPO)
     
End Sub

