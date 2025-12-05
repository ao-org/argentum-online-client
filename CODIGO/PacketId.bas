Attribute VB_Name = "PacketId"

Public Enum ServerPacketID
    eMinPacket
    eConnected
    elogged                  ' LOGGED  0
    eRemoveDialogs           ' QTDL
    eRemoveCharDialog        ' QDL
    eNavigateToggle          ' NAVEG
    eEquiteToggle
    eDisconnect              ' FINOK
    eCommerceEnd             ' FINCOMOK
    eBankEnd                 ' FINBANOK
    eCommerceInit            ' INITCOM
    eBankInit                ' INITBANCO
    eUserCommerceInit        ' INITCOMUSU   10
    eUserCommerceEnd         ' FINCOMUSUOK
    eShowBlacksmithForm      ' SFH
    eShowCarpenterForm       ' SFC
    eNPCKillUser             ' 6
    eBlockedWithShieldUser   ' 7
    eBlockedWithShieldOther  ' 8
    eCharSwing               ' U1
    eSafeModeOn              ' SEGON
    eSafeModeOff             ' SEGOFF 20
    ePartySafeOn
    ePartySafeOff
    eCantUseWhileMeditating  ' M!
    eUpdateSta               ' ASS
    eUpdateMana              ' ASM
    eUpdateHP                ' ASH
    eUpdateGold              ' ASG
    eUpdateExp               ' ASE 30
    eChangeMap               ' CM
    ePosUpdate               ' PU
    eNPCHitUser              ' N2
    eUserHittedByUser        ' N4
    eUserHittedUser          ' N5
    eChatOverHead            ' ||
    eLocaleChatOverHead
    eConsoleMsg              ' || - Beware!! its the same as above, but it was properly splitted
    eConsoleFactionMessage
    eGuildChat               ' |+   40
    eShowMessageBox          ' !!
    eMostrarCuenta
    eCharacterCreate         ' CC
    eCharacterRemove         ' BP
    eCharacterMove           ' MP, +, * and _ '
    eCharacterTranslate
    eUserIndexInServer       ' IU
    eUserCharIndexInServer   ' IP
    eForceCharMove
    eCharacterChange         ' CP
    eObjectCreate            ' HO
    efxpiso
    eObjectDelete            ' BO  50
    eBlockPosition           ' BQ
    ePlayMIDI                ' TM
    ePlayWave                ' TW
    eguildList               ' GL
    eAreaChanged             ' CA
    ePauseToggle             ' BKW
    eRainToggle              ' LLU
    eCreateFX                ' CFX
    eUpdateUserStats         ' EST
    eWorkRequestTarget       ' T01 60
    eChangeInventorySlot     ' CSI
    eInventoryUnlockSlots
    eChangeBankSlot          ' SBO
    eChangeSpellSlot         ' SHS
    eAtributes               ' ATR
    eBlacksmithWeapons       ' LAH
    eBlacksmithArmors        ' LAR
    eBlacksmithExtraObjects
    eCarpenterObjects        ' OBR
    eRestOK                  ' DOK
    eErrorMsg                ' ERR
    eBlind                   ' CEGU 70
    eDumb                    ' DUMB
    eShowSignal              ' MCAR
    eChangeNPCInventorySlot  ' NPCI
    eUpdateHungerAndThirst   ' EHYS
    eMiniStats               ' MEST
    eLevelUp                 ' SUNI
    eAddForumMsg             ' FMSG
    eShowForumForm           ' MFOR
    eSetInvisible            ' NOVER 80
    eMeditateToggle          ' MEDOK
    eBlindNoMore             ' NSEGUE
    eDumbNoMore              ' NESTUP
    eSendSkills              ' SKILLS
    eTrainerCreatureList     ' LSTCRI
    eguildNews               ' GUILDNE
    eOfferDetails            ' PEACEDE & ALLIEDE
    eAlianceProposalsList    ' ALLIEPR
    ePeaceProposalsList      ' PEACEPR 90
    eCharacterInfo           ' CHRINFO
    eGuildLeaderInfo         ' LEADERI
    eGuildDetails            ' CLANDET
    eShowGuildFundationForm  ' SHOWFUN
    eParalizeOK              ' PARADOK
    eStunStart               ' Stun start time
    eShowUserRequest         ' PETICIO
    eChangeUserTradeSlot     ' COMUSUINV
    'SendNight              ' NOC
    eUpdateTagAndStatus
    eFYA
    eCerrarleCliente
    eContadores
    eShowPapiro
    eUpdateCooldownType
    'GM messages
    eSpawnListt               ' SPL
    eShowSOSForm             ' MSOS
    eShowMOTDEditionForm     ' ZMOTD
    eShowGMPanelForm         ' ABPANEL
    eUserNameList            ' LISTUSU
    eUserOnline '110
    eParticleFX
    eParticleFXToFloor
    eParticleFXWithDestino
    eParticleFXWithDestinoXY
    ehora
    eLight
    eAuraToChar
    eSpeedToChar
    eLightToFloor
    eNieveToggle
    eNieblaToggle
    eGoliath
    eTextOverChar
    eTextOverTile
    eTextCharDrop
    eConsoleCharText
    eFlashScreen
    eAlquimistaObj
    eShowAlquimiaForm
    eSastreObj
    eShowSastreForm ' 126
    eVelocidadToggle
    eMacroTrabajoToggle
    eBindKeys
    eShowFrmLogear
    eShowFrmMapa
    eInmovilizadoOK
    eBarFx
    eLocaleMsg
    eShowPregunta
    eDatosGrupo
    eubicacion
    eArmaMov
    eEscudoMov
    eViajarForm
    eNadarToggle
    eShowFundarClanForm
    eCharUpdateHP
    eCharUpdateMAN
    ePosLLamadaDeClan
    eQuestDetails
    eQuestListSend
    eNpcQuestListSend
    eUpdateNPCSimbolo
    eClanSeguro
    eIntervals
    eUpdateUserKey
    eUpdateRM
    eUpdateDM
    eSeguroResu
    eLegionarySecure
    eStopped
    eInvasionInfo
    eCommerceRecieveChatMessage
    eDoAnimation
    eOpenCrafting
    eCraftingItem
    eCraftingCatalyst
    eCraftingResult
    eGuardNotice
    eAnswerReset
    eObjQuestListSend
    eUpdateBankGld
    ePelearConPezEspecial
    ePrivilegios
    eShopInit
    eUpdateShopClienteCredits
    eSendSkillCdUpdate
    eUpdateFlag
    eCharAtaca
    eNotificarClienteSeguido
    eGetInventarioHechizos
    eNotificarClienteCasteo
    ePosUpdateUserChar
    ePosUpdateChar
    ePlayWaveStep
    eShopPjsInit
    eDebugDataResponse
    eCreateProjectile
    eUpdateTrap
    eUpdateGroupInfo
    eUpdateCharValue 'updates some char index value based on enum
    eSendClientToggles 'Get active feature Toggles from server
    eAntiCheatMessage
    eAntiCheatStartSession
    eReportLobbyList
#If PYMMO = 0 Then
    eAccountCharacterList
#End If
    eChangeSkinSlot
    eGuildConfig
    eShowPickUpObj
    eMaxPacket
    [PacketCount]
End Enum

Public Enum ClientPacketID
    eMinPacket
    '--------------------
    eCraftCarpenter          'CNC
    eWorkLeftClick           'WLC
    eCreateNewGuild          'CIG
    eSpellInfo               'INFS
    eEquipItem               'EQUI
    eChangeHeading           'CHEA
    eModifySkills            'SKSE
    eTrain                   'ENTR
    eCommerceBuy             'COMP
    eBankExtractItem         'RETI
    eCommerceSell            'VEND
    eBankDeposit             'DEPO
    eForumPost               'DEMSG
    eMoveSpell               'DESPHE
    eClanCodexUpdate         'DESCOD
    eUserCommerceOffer       'OFRECER
    eGuildAcceptPeace        'ACEPPEAT
    eGuildRejectAlliance     'RECPALIA
    eGuildRejectPeace        'RECPPEAT
    eGuildAcceptAlliance     'ACEPALIA
    eGuildOfferPeace         'PEACEOFF
    eGuildOfferAlliance      'ALLIEOFF
    eGuildAllianceDetails    'ALLIEDET
    eGuildPeaceDetails       'PEACEDET
    eGuildRequestJoinerInfo  'ENVCOMEN
    eGuildAlliancePropList   'ENVALPRO
    eGuildPeacePropList      'ENVPROPP
    eGuildDeclareWar         'DECGUERR
    eGuildNewWebsite         'NEWWEBSI
    eGuildAcceptNewMember    'ACEPTARI
    eGuildRejectNewMember    'RECHAZAR
    eGuildKickMember         'ECHARCLA
    eGuildUpdateNews         'ACTGNEWS
    eGuildMemberInfo         '1HRINFO<
    eGuildOpenElections      'ABREELEC
    eGuildRequestMembership  'SOLICITUD
    eGuildRequestDetails     'CLANDETAILS
    eOnline                  '/ONLINE
    eQuit                    '/SALIR
    eGuildLeave              '/SALIRCLAN
    eRequestAccountState     '/BALANCE
    ePetStand                '/QUIETO
    ePetFollow               '/ACOMPAÑAR
    ePetLeave                '/LIBERAR
    eGrupoMsg                '/GrupoMsg
    eTrainList               '/ENTRENAR
    eRest                    '/DESCANSAR
    eMeditate                '/MEDITAR
    eResucitate              '/RESUCITAR
    eHeal                    '/CURAR
    eHelp                    '/AYUDA
    eRequestStats            '/EST
    eCommerceStart           '/COMERCIAR
    eBankStart               '/BOVEDA
    eEnlist                  '/ENLISTAR
    eInformation             '/INFORMACION
    eReward                  '/RECOMPENSA
    eRequestMOTD             '/MOTD
    eUpTime                  '/UPTIME
    eGuildMessage            '/CMSG
    eGuildOnline             '/ONLINECLAN
    eCouncilMessage          '/BMSG
    eFactionMessage          '/FMSG
    eRoleMasterRequest       '/ROL
    eChangeDescription       '/DESC
    eGuildVote               '/VOTO
    epunishments             '/PENAS
    eGamble                  '/APOSTAR
    eMapPriceEntrance        '/ARENA
    eLeaveFaction            '/RETIRAR ( with no arguments )
    eBankExtractGold         '/RETIRAR ( with arguments )
    eBankDepositGold         '/DEPOSITAR
    eDenounce                '/DENUNCIAR
    eLoginExistingChar       'OLOGIN
    eLoginNewChar            'NLOGIN
    eTalk                    ';
    eYell                    '-
    eWhisper                 '\
    eWalk                    'M
    eRequestPositionUpdate   'RPU
    eAttack                  'AT
    ePickUp                  'AG
    eSafeToggle              '/SEG & SEG  (SEG's behaviour has to be coded in the client)
    ePartySafeToggle
    eRequestGuildLeaderInfo  'GLINFO
    eRequestAtributes        'ATR
    eRequestSkills           'ESKI
    eRequestMiniStats        'FEST
    eCommerceEnd             'FINCOM
    eUserCommerceEnd         'FINCOMUSU
    eBankEnd                 'FINBAN
    eUserCommerceOk          'COMUSUOK
    eUserCommerceReject      'COMUSUNO
    eDrop                    'TI
    eCastSpell               'LH
    eLeftClick               'LC
    eDoubleClick             'RC
    eWork                    'UK
    eUseSpellMacro           'UMH
    eUseItem                 'USA
    eCraftBlacksmith         'CNS
    'GM messages
    eGMMessage               '/GMSG
    eshowName                '/SHOWNAME
    eOnlineRoyalArmy         '/ONLINEREAL
    eOnlineChaosLegion       '/ONLINECAOS
    eGoNearby                '/IRCERCA
    ecomment                 '/REM
    eWhere                   '/DONDE
    eCreaturesInMap          '/NENE
    eWarpMeToTarget          '/TELEPLOC
    eWarpChar                '/TELEP
    eSilence                 '/SILENCIAR
    eSOSShowList             '/SHOW SOS
    eSOSRemove               'SOSDONE
    eGoToChar                '/IRA
    einvisible               '/INVISIBLE
    eGMPanel                 '/PANELGM
    eRequestUserList         'LISTUSU
    eWorking                 '/TRABAJANDO
    eHiding                  '/OCULTANDO
    eJail                    '/CARCEL
    eKillNPC                 '/RMATA
    eWarnUser                '/ADVERTENCIA
    eEditChar                '/MOD
    eRequestCharInfo         '/INFO
    eRequestCharStats        '/STAT
    eRequestCharGold         '/BAL
    eRequestCharInventory    '/INV
    eRequestCharBank         '/BOV
    eRequestCharSkills       '/SKILLS
    eReviveChar              '/REVIVIR
    eOnlineGM                '/ONLINEGM
    eOnlineMap               '/ONLINEMAP
    eForgive                 '/PERDON
    eKick                    '/ECHAR
    eExecute                 '/EJECUTAR
    eBanChar                 '/BAN
    eUnbanChar               '/UNBAN
    eNPCFollow               '/SEGUIR
    eSummonChar              '/SUM
    eSpawnListRequest        '/CC
    eSpawnCreature           'SPA
    eResetNPCInventory       '/RESETINV
    eCleanWorld              '/LIMPIAR
    eServerMessage           '/RMSG
    eNickToIP                '/NICK2IP
    eIPToNick                '/IP2NICK
    eGuildOnlineMembers      '/ONCLAN
    eTeleportCreate          '/CT
    eTeleportDestroy         '/DT
    eRainToggle              '/LLUVIA
    eSetCharDescription      '/SETDESC
    eForceMIDIToMap          '/FORCEMIDIMAP
    eForceWAVEToMap          '/FORCEWAVMAP
    eRoyalArmyMessage        '/REALMSG
    eChaosLegionMessage      '/CAOSMSG
    eTalkAsNPC               '/TALKAS
    eDestroyAllItemsInArea   '/MASSDEST
    eAcceptRoyalCouncilMember '/ACEPTCONSE
    eAcceptChaosCouncilMember '/ACEPTCONSECAOS
    eItemsInTheFloor         '/PISO
    eMakeDumb                '/ESTUPIDO
    eMakeDumbNoMore          '/NOESTUPIDO
    eCouncilKick             '/KICKCONSE
    eSetTrigger              '/TRIGGER
    eAskTrigger              '/TRIGGER with no args
    eGuildMemberList         '/MIEMBROSCLAN
    eGuildBan                '/BANCLAN
    eCreateItem              '/CI
    eDestroyItems            '/DEST
    eChaosLegionKick         '/NOCAOS
    eRoyalArmyKick           '/NOREAL
    eForceMIDIAll            '/FORCEMIDI
    eForceWAVEAll            '/FORCEWAV
    eRemovePunishment        '/BORRARPENA
    eTileBlockedToggle       '/BLOQ
    eKillNPCNoRespawn        '/MATA
    eKillAllNearbyNPCs       '/MASSKILL
    eLastIP                  '/LASTIP
    eChangeMOTD              '/MOTDCAMBIA
    eSetMOTD                 'ZMOTD
    eSystemMessage           '/SMSG
    eCreateNPC               '/ACC
    eCreateNPCWithRespawn    '/RACC
    eImperialArmour          '/AI1 - 4
    eChaosArmour             '/AC1 - 4
    eNavigateToggle          '/NAVE
    eServerOpenToUsersToggle '/HABILITAR
    eParticipar              '/PARTICIPAR
    eTurnCriminal            '/CONDEN
    eResetFactions           '/RAJAR
    eRemoveCharFromGuild     '/RAJARCLAN
    eAlterName               '/ANAME
    eDoBackUp                '/DOBACKUP
    eShowGuildMessages       '/SHOWCMSG
    eChangeMapInfoPK         '/MODMAPINFO PK
    eChangeMapInfoBackup     '/MODMAPINFO BACKUP
    eChangeMapInfoRestricted '/MODMAPINFO RESTRINGIR
    eChangeMapInfoNoMagic    '/MODMAPINFO MAGIASINEFECTO
    eChangeMapInfoNoInvi     '/MODMAPINFO INVISINEFECTO
    eChangeMapInfoNoResu     '/MODMAPINFO RESUSINEFECTO
    eChangeMapInfoLand       '/MODMAPINFO TERRENO
    eChangeMapInfoZone       '/MODMAPINFO ZONA
    eChangeMapSetting        '/MODSETTING setting value
    eSaveChars               '/GRABAR
    eCleanSOS                '/BORRAR SOS
    eShowServerForm          '/SHOW INT
    eKickAllChars            '/ECHARTODOSPJS
    eChatColor               '/CHATCOLOR
    eIgnored                 '/IGNORADO
    eCheckSlot               '/SLOT
    eSetSpeed                '/SPEED
    eGlobalMessage           '/CONSOLA
    eGlobalOnOff
    eUseKey
    eDonateGold              '/DONAR
    ePromedio                '/PROMEDIO
    eGiveItem                '/DAR
    eOfertaInicial
    eOfertaDeSubasta
    eQuestionGM
    eCuentaRegresiva
    ePossUser
    eDuel
    eAcceptDuel
    eCancelDuel
    eQuitDuel
    eNieveToggle
    eNieblaToggle
    eTransFerGold
    eMoveitem
    eGenio
    eCasarse
    eCraftAlquimista
    eFlagTrabajar
    eCraftSastre
    eMensajeUser
    eTraerBoveda
    eCompletarAccion
    eInvitarGrupo
    eResponderPregunta
    eRequestGrupo
    eAbandonarGrupo
    eHecharDeGrupo
    eMacroPossent
    eSubastaInfo
    eBanCuenta
    eUnbanCuenta
    eCerrarCliente
    eEventoInfo
    eCrearEvento
    eBanTemporal
    eCancelarExit
    eCrearTorneo
    eComenzarTorneo
    eCancelarTorneo
    eBusquedaTesoro
    eCompletarViaje
    eBovedaMoveItem
    eQuieroFundarClan
    ellamadadeclan
    eMarcaDeClanPack
    eMarcaDeGMPack
    eQuest
    eQuestAccept
    eQuestListRequest
    eQuestDetailsRequest
    eQuestAbandon
    eSeguroClan
    ehome                    '/HOGAR
    eConsulta                '/CONSULTA
    eGetMapInfo              '/MAPINFO
    eFinEvento
    eSeguroResu
    eLegionarySecure
    eCuentaExtractItem
    eCuentaDeposit
    eCreateEvent
    eCommerceSendChatMessage
    eLogMacroClickHechizo
    eAddItemCrafting
    eRemoveItemCrafting
    eAddCatalyst
    eRemoveCatalyst
    eCraftItem
    eCloseCrafting
    eMoveCraftItem
    ePetLeaveAll
    eResetChar              '/RESET NICK
    eResetearPersonaje
    eDeleteItem
    eFinalizarPescaEspecial
    eRomperCania
    eUseItemU
    eRepeatMacro
    eBuyShopItem
    ePerdonFaccion              '/PERDONFACCION NAME
    eStartEvent           '/EVENTO CAPTURA/LOBBY
    eCancelarEvento          '/CANCELAREVENTO
    eNotifyInventarioHechizos
    ePublicarPersonajeMAO
    eEventoFaccionario
    eRequestDebug '/RequestDebug consulta info debug al server, para gms
    eLobbyCommand
    eFeatureToggle
    eActionOnGroupFrame
    eSetHotkeySlot
    eUseHKeySlot
    eAntiCheatMessage
    eRequestLobbyList
    #If PYMMO = 0 Then
        eCreateAccount
        eLoginAccount
        eDeleteCharacter
    #End If
    eChangeSkinSlot
    eMaxPacket
    [PacketCount]
End Enum
