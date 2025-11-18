Attribute VB_Name = "Protocol"
' Argentum 20 Game Client
'
'    Copyright (C) 2023 Noland Studios LTD
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
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 Márquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'
Option Explicit
''
'When we have a list of strings, we use this to separate them and prevent
'having too many string lengths in the queue. Yes, each string is NULL-terminated :P
Private Const SEPARATOR As String * 1 = vbNullChar
Private LastPacket      As Byte
Private IterationsHID   As Integer
Private Const MAX_ITERATIONS_HID = 200
#If DIRECT_PLAY = 0 Then
    Private Reader As Network.Reader

Public Function HandleIncomingData(ByVal message As Network.Reader) As Boolean
#Else
    Private Reader As New clsNetReader
    Public Function HandleIncomingData(dpnotify As DxVBLibA.DPNMSG_RECEIVE) As Boolean
    #End If
    On Error GoTo HandleIncomingData_Err
    #If DIRECT_PLAY = 0 Then
        Set Reader = message
    #Else
        Reader.set_data dpnotify
    #End If
    Dim PacketId As Long
    PacketId = Reader.ReadInt16
    #If REMOTE_CLOSE = 1 Then
        Select Case PacketId
            Case ServerPacketID.eConnected
                Call HandleConnected
                Call SaveStringInFile("Authenticated with server OK", "remote_debug.txt")
            Case ServerPacketID.elogged
                frmDebug.add_text_tracebox "Logged"
                Dim dummy As Boolean
                dummy = Reader.ReadBool
                Call SaveStringInFile("Logged with character " & CharacterRemote, "remote_debug.txt")
                InitiateShutdownProcess = True
                ShutdownProcessTimer.start
            Case ServerPacketID.eLocaleMsg
                Dim chat      As String
                Dim FontIndex As Integer
                Dim str       As String
                PacketId = Reader.ReadInt16()
                chat = Reader.ReadString8()
                FontIndex = Reader.ReadInt8()
                Call SaveStringInFile(chat, "remote_debug.txt")
            Case Else
                'don't care, just consume
                Do While (message.GetAvailable() > 0)
                    PacketId = Reader.ReadInt8
                Loop
        End Select
    #Else
        Select Case PacketId
            Case ServerPacketID.eConnected
                Call HandleConnected
            Case ServerPacketID.elogged
                Call HandleLogged
            Case ServerPacketID.eRemoveDialogs
                Call HandleRemoveDialogs
            Case ServerPacketID.eRemoveCharDialog
                Call HandleRemoveCharDialog
            Case ServerPacketID.eNavigateToggle
                Call HandleNavigateToggle
            Case ServerPacketID.eEquiteToggle
                Call HandleEquiteToggle
            Case ServerPacketID.eDisconnect
                Call HandleDisconnect
            Case ServerPacketID.eCommerceEnd
                Call HandleCommerceEnd
            Case ServerPacketID.eBankEnd
                Call HandleBankEnd
            Case ServerPacketID.eCommerceInit
                Call HandleCommerceInit
            Case ServerPacketID.eBankInit
                Call HandleBankInit
            Case ServerPacketID.eUserCommerceInit
                Call HandleUserCommerceInit
            Case ServerPacketID.eUserCommerceEnd
                Call HandleUserCommerceEnd
            Case ServerPacketID.eShowBlacksmithForm
                Call HandleShowBlacksmithForm
            Case ServerPacketID.eShowCarpenterForm
                Call HandleShowCarpenterForm
            Case ServerPacketID.eNPCKillUser
                Call HandleNPCKillUser
            Case ServerPacketID.eBlockedWithShieldUser
                Call HandleBlockedWithShieldUser
            Case ServerPacketID.eBlockedWithShieldOther
                Call HandleBlockedWithShieldOther
            Case ServerPacketID.eCharSwing
                Call HandleCharSwing
            Case ServerPacketID.eSafeModeOn
                Call HandleSafeModeOn
            Case ServerPacketID.eSafeModeOff
                Call HandleSafeModeOff
            Case ServerPacketID.ePartySafeOn
                Call HandlePartySafeOn
            Case ServerPacketID.ePartySafeOff
                Call HandlePartySafeOff
            Case ServerPacketID.eCantUseWhileMeditating
                Call HandleCantUseWhileMeditating
            Case ServerPacketID.eUpdateSta
                Call HandleUpdateSta
            Case ServerPacketID.eUpdateMana
                Call HandleUpdateMana
            Case ServerPacketID.eUpdateHP
                Call HandleUpdateHP
            Case ServerPacketID.eUpdateGold
                Call HandleUpdateGold
            Case ServerPacketID.eUpdateExp
                Call HandleUpdateExp
            Case ServerPacketID.eChangeMap
                Call HandleChangeMap
            Case ServerPacketID.ePosUpdate
                Call HandlePosUpdate
            Case ServerPacketID.ePosUpdateUserChar
                Call HandlePosUpdateUserChar
            Case ServerPacketID.ePosUpdateChar
                Call HandlePosUpdateChar
            Case ServerPacketID.eNPCHitUser
                Call HandleNPCHitUser
            Case ServerPacketID.eUserHittedByUser
                Call HandleUserHittedByUser
            Case ServerPacketID.eUserHittedUser
                Call HandleUserHittedUser
            Case ServerPacketID.eChatOverHead
                Call HandleChatOverHead
            Case ServerPacketID.eLocaleChatOverHead
                Call HandleLocaleChatOverHead
            Case ServerPacketID.eConsoleMsg
                Call HandleConsoleMessage
            Case ServerPacketID.eConsoleFactionMessage
                Call HandleConsoleFactionMessage
            Case ServerPacketID.eGuildChat
                Call HandleGuildChat
            Case ServerPacketID.eShowMessageBox
                Call HandleShowMessageBox
            Case ServerPacketID.eMostrarCuenta
                Call HandleMostrarCuenta
            Case ServerPacketID.eCharacterCreate
                Call HandleCharacterCreate
            Case ServerPacketID.eUpdateFlag
                Call HandleUpdateFlag
            Case ServerPacketID.eCharacterRemove
                Call HandleCharacterRemove
            Case ServerPacketID.eCharacterMove
                Call HandleCharacterMove
            Case ServerPacketID.eCharacterTranslate
                Call HandleCharacterTranslate
            Case ServerPacketID.eUserIndexInServer
                Call HandleUserIndexInServer
            Case ServerPacketID.eUserCharIndexInServer
                Call HandleUserCharIndexInServer
            Case ServerPacketID.eForceCharMove
                Call HandleForceCharMove
            Case ServerPacketID.eCharacterChange
                Call HandleCharacterChange
            Case ServerPacketID.eObjectCreate
                Call HandleObjectCreate
            Case ServerPacketID.efxpiso
                Call HandleFxPiso
            Case ServerPacketID.eObjectDelete
                Call HandleObjectDelete
            Case ServerPacketID.eBlockPosition
                Call HandleBlockPosition
            Case ServerPacketID.ePlayMIDI
                Call HandlePlayMIDI
            Case ServerPacketID.ePlayWave
                Call HandlePlayWave
            Case ServerPacketID.ePlayWaveStep
                Call HandlePlayWaveStep
            Case ServerPacketID.eguildList
                Call HandleGuildList
            Case ServerPacketID.eAreaChanged
                Call HandleAreaChanged
            Case ServerPacketID.ePauseToggle
                Call HandlePauseToggle
            Case ServerPacketID.eRainToggle
                Call HandleRainToggle
            Case ServerPacketID.eCreateFX
                Call HandleCreateFX
            Case ServerPacketID.eCharAtaca
                Call HandleCharAtaca
            Case ServerPacketID.eGetInventarioHechizos
                Call HandleGetInventarioHechizos
            Case ServerPacketID.eNotificarClienteCasteo
                Call HandleNotificarClienteCasteo
            Case ServerPacketID.eNotificarClienteSeguido
                Call HandleNotificarClienteSeguido
            Case ServerPacketID.eUpdateUserStats
                Call HandleUpdateUserStats
            Case ServerPacketID.eWorkRequestTarget
                Call HandleWorkRequestTarget
            Case ServerPacketID.eChangeInventorySlot
                Call HandleChangeInventorySlot
            Case ServerPacketID.eInventoryUnlockSlots
                Call HandleInventoryUnlockSlots
            Case ServerPacketID.eChangeBankSlot
                Call HandleChangeBankSlot
            Case ServerPacketID.eChangeSpellSlot
                Call HandleChangeSpellSlot
            Case ServerPacketID.eAtributes
                Call HandleAtributes
            Case ServerPacketID.eBlacksmithWeapons
                Call HandleBlacksmithWeapons
            Case ServerPacketID.eBlacksmithArmors
                Call HandleBlacksmithArmors
            Case ServerPacketID.eBlacksmithExtraObjects
                Call HandleBlacksmithExtraObjects
            Case ServerPacketID.eCarpenterObjects
                Call HandleCarpenterObjects
            Case ServerPacketID.eRestOK
                Call HandleRestOK
            Case ServerPacketID.eErrorMsg
                Call HandleErrorMessage
            Case ServerPacketID.eBlind
                Call HandleBlind
            Case ServerPacketID.eDumb
                Call HandleDumb
            Case ServerPacketID.eShowSignal
                Call HandleShowSignal
            Case ServerPacketID.eChangeNPCInventorySlot
                Call HandleChangeNPCInventorySlot
            Case ServerPacketID.eUpdateHungerAndThirst
                Call HandleUpdateHungerAndThirst
            Case ServerPacketID.eMiniStats
                Call HandleMiniStats
            Case ServerPacketID.eLevelUp
                Call HandleLevelUp
            Case ServerPacketID.eAddForumMsg
                Call HandleAddForumMessage
            Case ServerPacketID.eShowForumForm
                Call HandleShowForumForm
            Case ServerPacketID.eSetInvisible
                Call HandleSetInvisible
            Case ServerPacketID.eMeditateToggle
                Call HandleMeditateToggle
            Case ServerPacketID.eBlindNoMore
                Call HandleBlindNoMore
            Case ServerPacketID.eDumbNoMore
                Call HandleDumbNoMore
            Case ServerPacketID.eSendSkills
                Call HandleSendSkills
            Case ServerPacketID.eTrainerCreatureList
                Call HandleTrainerCreatureList
            Case ServerPacketID.eguildNews
                Call HandleGuildNews
            Case ServerPacketID.eOfferDetails
                Call HandleOfferDetails
            Case ServerPacketID.eAlianceProposalsList
                Call HandleAlianceProposalsList
            Case ServerPacketID.ePeaceProposalsList
                Call HandlePeaceProposalsList
            Case ServerPacketID.eCharacterInfo
                Call HandleCharacterInfo
            Case ServerPacketID.eGuildLeaderInfo
                Call HandleGuildLeaderInfo
            Case ServerPacketID.eGuildDetails
                Call HandleGuildDetails
            Case ServerPacketID.eShowGuildFundationForm
                Call HandleShowGuildFundationForm
            Case ServerPacketID.eParalizeOK
                Call HandleParalizeOK
            Case ServerPacketID.eStunStart
                Call HandleStunStart
            Case ServerPacketID.eShowUserRequest
                Call HandleShowUserRequest
            Case ServerPacketID.eChangeUserTradeSlot
                Call HandleChangeUserTradeSlot
            Case ServerPacketID.eUpdateTagAndStatus
                Call HandleUpdateTagAndStatus
            Case ServerPacketID.eFYA
                Call HandleFYA
            Case ServerPacketID.eCerrarleCliente
                Call HandleCerrarleCliente
            Case ServerPacketID.eContadores
                Call HandleContadores
            Case ServerPacketID.eShowPapiro
                Call HandleShowPapiro
            Case ServerPacketID.eUpdateCooldownType
                Call HandleUpdateCooldownType
            Case ServerPacketID.eSpawnListt
                Call HandleSpawnList
            Case ServerPacketID.eShowSOSForm
                Call HandleShowSOSForm
            Case ServerPacketID.eShowMOTDEditionForm
                Call HandleShowMOTDEditionForm
            Case ServerPacketID.eShowGMPanelForm
                Call HandleShowGMPanelForm
            Case ServerPacketID.eUserNameList
                Call HandleUserNameList
            Case ServerPacketID.eUserOnline
                Call HandleUserOnline
            Case ServerPacketID.eParticleFX
                Call HandleParticleFX
            Case ServerPacketID.eParticleFXToFloor
                Call HandleParticleFXToFloor
            Case ServerPacketID.eParticleFXWithDestino
                Call HandleParticleFXWithDestino
            Case ServerPacketID.eParticleFXWithDestinoXY
                Call HandleParticleFXWithDestinoXY
            Case ServerPacketID.ehora
                Call HandleHora
            Case ServerPacketID.eLight
                Call HandleLight
            Case ServerPacketID.eAuraToChar
                Call HandleAuraToChar
            Case ServerPacketID.eSpeedToChar
                Call HandleSpeedToChar
            Case ServerPacketID.eLightToFloor
                Call HandleLightToFloor
            Case ServerPacketID.eNieveToggle
                Call HandleNieveToggle
            Case ServerPacketID.eNieblaToggle
                Call HandleNieblaToggle
            Case ServerPacketID.eGoliath
                Call HandleGoliath
            Case ServerPacketID.eTextOverChar
                Call HandleTextOverChar
            Case ServerPacketID.eTextOverTile
                Call HandleTextOverTile
            Case ServerPacketID.eTextCharDrop
                Call HandleTextCharDrop
            Case ServerPacketID.eConsoleCharText
                Call HandleConsoleCharText
            Case ServerPacketID.eFlashScreen
                Call HandleFlashScreen
            Case ServerPacketID.eAlquimistaObj
                Call HandleAlquimiaObjects
            Case ServerPacketID.eShowAlquimiaForm
                Call HandleShowAlquimiaForm
            Case ServerPacketID.eSastreObj
                Call HandleSastreObjects
            Case ServerPacketID.eShowSastreForm
                Call HandleShowSastreForm
            Case ServerPacketID.eVelocidadToggle
                Call HandleVelocidadToggle
            Case ServerPacketID.eMacroTrabajoToggle
                Call HandleMacroTrabajoToggle
            Case ServerPacketID.eBindKeys
                Call HandleBindKeys
            Case ServerPacketID.eShowFrmLogear
                Call HandleShowFrmLogear
            Case ServerPacketID.eShowFrmMapa
                Call HandleShowFrmMapa
            Case ServerPacketID.eInmovilizadoOK
                Call HandleInmovilizadoOK
            Case ServerPacketID.eBarFx
                Call HandleBarFx
            Case ServerPacketID.eLocaleMsg
                Call HandleLocaleMsg
            Case ServerPacketID.eShowPregunta
                Call HandleShowPregunta
            Case ServerPacketID.eDatosGrupo
                Call HandleDatosGrupo
            Case ServerPacketID.eubicacion
                Call HandleUbicacion
            Case ServerPacketID.eArmaMov
                Call HandleArmaMov
            Case ServerPacketID.eEscudoMov
                Call HandleEscudoMov
            Case ServerPacketID.eViajarForm
                Call HandleViajarForm
            Case ServerPacketID.eNadarToggle
                Call HandleNadarToggle
            Case ServerPacketID.eShowFundarClanForm
                Call HandleShowFundarClanForm
            Case ServerPacketID.eCharUpdateHP
                Call HandleCharUpdateHP
            Case ServerPacketID.eCharUpdateMAN
                Call HandleCharUpdateMAN
            Case ServerPacketID.ePosLLamadaDeClan
                Call HandlePosLLamadaDeClan
            Case ServerPacketID.eQuestDetails
                Call HandleQuestDetails
            Case ServerPacketID.eQuestListSend
                Call HandleQuestListSend
            Case ServerPacketID.eNpcQuestListSend
                Call HandleNpcQuestListSend
            Case ServerPacketID.eUpdateNPCSimbolo
                Call HandleUpdateNPCSimbolo
            Case ServerPacketID.eClanSeguro
                Call HandleClanSeguro
            Case ServerPacketID.eIntervals
                Call HandleIntervals
            Case ServerPacketID.eUpdateUserKey
                Call HandleUpdateUserKey
            Case ServerPacketID.eUpdateRM
                Call HandleUpdateRM
            Case ServerPacketID.eUpdateDM
                Call HandleUpdateDM
            Case ServerPacketID.eSeguroResu
                Call HandleSeguroResu
            Case ServerPacketID.eLegionarySecure
                Call HandleLegionarySecure
            Case ServerPacketID.eStopped
                Call HandleStopped
            Case ServerPacketID.eInvasionInfo
                Call HandleInvasionInfo
            Case ServerPacketID.eCommerceRecieveChatMessage
                Call HandleCommerceRecieveChatMessage
            Case ServerPacketID.eDoAnimation
                Call HandleDoAnimation
            Case ServerPacketID.eOpenCrafting
                Call HandleOpenCrafting
            Case ServerPacketID.eCraftingItem
                Call HandleCraftingItem
            Case ServerPacketID.eCraftingCatalyst
                Call HandleCraftingCatalyst
            Case ServerPacketID.eCraftingResult
                Call HandleCraftingResult
            Case ServerPacketID.eAnswerReset
                Call HandleAnswerReset
            Case ServerPacketID.eObjQuestListSend
                Call HandleObjQuestListSend
            Case ServerPacketID.eUpdateBankGld
                Call HandleUpdateBankGld
            Case ServerPacketID.ePelearConPezEspecial
                Call HandlePelearConPezEspecial
            Case ServerPacketID.ePrivilegios
                Call HandlePrivilegios
            Case ServerPacketID.eShopInit
                Call HandleShopInit
            Case ServerPacketID.eShopPjsInit
                Call HandleShopPjsInit
            Case ServerPacketID.eUpdateShopClienteCredits
                Call HandleUpdateShopClienteCredits
            Case ServerPacketID.eSendSkillCdUpdate
                Call HandleSendSkillCdUpdate
            Case ServerPacketID.eDebugDataResponse
                Call HandleDebugDataResponse
            Case ServerPacketID.eCreateProjectile
                Call HandleCreateProjectile
            Case ServerPacketID.eUpdateTrap
                Call HandleUpdateTrapState
            Case ServerPacketID.eUpdateGroupInfo
                Call HandleUpdateGroupInfo
            Case ServerPacketID.eUpdateCharValue
                Call HandleUpdateCharValue
            Case ServerPacketID.eSendClientToggles
                Call HandleSendClientToggles
            Case ServerPacketID.eAntiCheatMessage
                Call HandleAntiCheatMessage
            Case ServerPacketID.eAntiCheatStartSession
                Call HandleAntiCheatStartSession
            Case ServerPacketID.eReportLobbyList
                Call HandleReportLobbyList
            Case ServerPacketID.eChangeSkinSlot
                Call HandleChangeSkinSlot
                #If PYMMO = 0 Then
                Case ServerPacketID.eAccountCharacterList
                    Call HandleAccountCharacterList
                #End If
            Case Else
                ' Invalid Message
        End Select
    #End If
    ' —————————————————————————————
    ' Detect both (a) extra bytes from known packets
    '         and (b) any packet where we had NO handler
    ' In either case, Reader.GetAvailable() > 0
    If (Reader.GetAvailable() > 0) Then
        Call RegistrarError(&HDEADBEEF, "Server message ID: " & PacketId & " unhandled or too many bytes; " & Reader.GetAvailable() & " extra bytes found", _
                "Protocol.HandleIncomingData", Erl)
        Do While (Reader.GetAvailable() > 0)
            Dim dummy As Byte
            dummy = Reader.ReadInt8
        Loop
    End If
    HandleIncomingData = True
HandleIncomingData_Err:
    Set Reader = Nothing
    If Err.Number <> 0 Then
        Call RegistrarError(Err.Number, Err.Description & ". PacketID: " & PacketId, "Protocol.HandleIncomingData", Erl)
        HandleIncomingData = False
    End If
End Function

Private Sub HandleConnected()
    #If DEBUGGING = 1 Then
        Dim i        As Integer
        Dim values() As Byte
        Reader.ReadSafeArrayInt8 values
        For i = LBound(values) To UBound(values)
            Debug.Assert values(i) = i
        Next i
    #End If
    #If REMOTE_CLOSE = 0 Then
        frmMain.ShowFPS.enabled = True
    #End If
    #If DIRECT_PLAY = 0 Then
        'We already sent the LoginExistingChar message with the double click event
        Call Login
    #End If
End Sub

Private Sub HandleLogged()
    On Error GoTo HandleLogged_Err
    newUser = Reader.ReadBool
    UserCiego = False
    EngineRun = True
    UserDescansar = False
    Nombres = True
    Pregunta = False
    frmMain.stabar.visible = True
    frmMain.panelInf.Picture = LoadInterface("ventanaprincipal_stats.bmp")
    frmMain.HpBar.visible = True
    If UserStats.maxman <> 0 Then
        frmMain.manabar.visible = True
    End If
    frmMain.hambar.visible = True
    frmMain.AGUbar.visible = True
    frmMain.Hpshp.visible = (UserStats.MinHp > 0)
    frmMain.shieldBar.visible = (UserStats.HpShield > 0)
    frmMain.MANShp.visible = (UserStats.minman > 0)
    frmMain.STAShp.visible = (UserStats.MinSTA > 0)
    frmMain.AGUAsp.visible = (UserStats.MinAGU > 0)
    frmMain.COMIDAsp.visible = (UserStats.MinHAM > 0)
    frmMain.GldLbl.visible = True
    frmMain.Fuerzalbl.visible = True
    frmMain.AgilidadLbl.visible = True
    frmMain.oxigenolbl.visible = True
    frmMain.imgDeleteItem.visible = True
    frmMain.oxigenolbl.visible = False
    Call NameMapa(ResourceMap)
    lFrameTimer = 0
    FramesPerSecCounter = 0
    frmMain.ImgSegParty = LoadInterface("boton-seguro-party-on.bmp")
    frmMain.ImgSegClan = LoadInterface("boton-seguro-clan-on.bmp")
    frmMain.ImgSegResu = LoadInterface("boton-fantasma-on.bmp")
    frmMain.ImgLegionarySecure = LoadInterface("boton-demonio-on.bmp")
    SeguroParty = True
    SeguroClanX = True
    SeguroResuX = True
    LegionarySecureX = True
    Call ResetAllCd
    Call SetConnected
    g_game_state.State = e_state_gameplay_screen
    Exit Sub
HandleLogged_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleLogged", Erl)
End Sub

''
' Handles the RemoveDialogs message.
Private Sub HandleRemoveDialogs()
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
    Call Dialogos.RemoveDialog(Reader.ReadInt16())
    Exit Sub
HandleRemoveCharDialog_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleRemoveCharDialog", Erl)
End Sub

''
' Handles the NavigateToggle message.
Private Sub HandleNavigateToggle()
    'Remove packet ID
    On Error GoTo HandleNavigateToggle_Err
    UserNavegando = Reader.ReadBool()
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
    Dim Speeding As Single
    Speeding = Reader.ReadReal32()
    If UserCharIndex = 0 Then Exit Sub
    charlist(UserCharIndex).Speeding = Speeding
    Call ApplySpeedingToChar(UserCharIndex)
    Call MainTimer.SetInterval(TimersIndex.Walk, gIntervals.Walk / charlist(UserCharIndex).Speeding)
    Exit Sub
HandleVelocidadToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleVelocidadToggle", Erl)
End Sub

Private Sub HandleMacroTrabajoToggle()
    'Activa o Desactiva el macro de trabajo
    On Error GoTo HandleMacroTrabajoToggle_Err
    Dim activar As Boolean
    activar = Reader.ReadBool()
    If activar = False Then
        Call ResetearUserMacro
    Else
        Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_TRABAJO_INICIA"), 2, 223, 51, 1, 0)
        frmMain.MacroLadder.Interval = gIntervals.BuildWork
        frmMain.MacroLadder.enabled = True
        UserMacro.Intervalo = gIntervals.BuildWork
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
    Dim i As Long
    If (Not Reader Is Nothing) Then
        FullLogout = Reader.ReadBool
    End If
    Call SaveSetting("OPCIONES", "LastScroll", hlst.Scroll)
    Mod_Declaraciones.Connected = False
    Call ResetearUserMacro
    'Close connection
    Call modNetwork.Disconnect
    'Hide main form
    Call resetearCartel
    frmConnect.visible = True
    isLogged = False
    Call Graficos_Particulas.Particle_Group_Remove_All
    Call Graficos_Particulas.Engine_Select_Particle_Set(203)
    ParticleLluviaDorada = General_Particle_Create(208, -1, -1)
    frmMain.picHechiz.visible = False
    frmMain.UpdateLight.enabled = False
    frmMain.UpdateDaytime.enabled = False
    frmMain.visible = False
    Seguido = False
    OpcionMenu = 0
    frmMain.picInv.visible = True
    frmMain.picHechiz.visible = False
    frmMain.cmdlanzar.visible = False
    'frmMain.lblrefuerzolanzar.Visible = False
    frmMain.cmdMoverHechi(0).visible = False
    frmMain.cmdMoverHechi(1).visible = False
    QuePestañaInferior = 0
    frmMain.stabar.visible = True
    frmMain.HpBar.visible = True
    frmMain.manabar.visible = True
    frmMain.hambar.visible = True
    frmMain.AGUbar.visible = True
    frmMain.shieldBar.visible = (UserStats.HpShield > 0)
    frmMain.Hpshp.visible = True
    frmMain.shieldBar.visible = True
    frmMain.MANShp.visible = True
    frmMain.STAShp.visible = True
    frmMain.AGUAsp.visible = True
    frmMain.COMIDAsp.visible = True
    frmMain.GldLbl.visible = True
    frmMain.Fuerzalbl.visible = True
    frmMain.AgilidadLbl.visible = True
    frmMain.oxigenolbl.visible = True
    frmMain.QuestBoton.visible = False
    frmMain.ImgHogar.visible = False
    frmMain.lblWeapon.visible = True
    frmMain.lblShielder.visible = True
    frmMain.lblHelm.visible = True
    frmMain.lblArmor.visible = True
    frmMain.lblResis.visible = True
    frmMain.lbldm.visible = True
    frmMain.imgBugReport.visible = False
    frmMain.oxigenolbl.visible = False
    frmMain.panelinferior(0).Picture = Nothing
    frmMain.panelinferior(1).Picture = Nothing
    frmMain.buttonskins.visible = False
    frmMain.Image5.visible = False
    frmMain.clanimg.visible = False
    frmMain.cmdLlavero.visible = False
    frmMain.QuestBoton.visible = False
    frmMain.ImgSeg.visible = False
    frmMain.ImgSegParty.visible = False
    frmMain.ImgSegClan.visible = False
    frmMain.ImgSegResu.visible = False
    frmMain.ImgLegionarySecure.visible = False
    initPacketControl
    Call ao20audio.StopAllPlayback
    Call CleanDialogs
    'Show connection form
    UserMap = 1
    EntradaY = 1
    EntradaX = 1
    Call EraseChar(UserCharIndex, True)
    Call SwitchMap(UserMap)
    frmMain.personaje(1).visible = False
    frmMain.personaje(2).visible = False
    frmMain.personaje(3).visible = False
    frmMain.personaje(4).visible = False
    frmMain.personaje(5).visible = False
    UserStats.Clase = 0
    UserStats.Sexo = 0
    UserStats.Raza = 0
    MiCabeza = 0
    UserStats.Hogar = 0
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i
    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i
    For i = 1 To UserInvUnlocked
        frmMain.imgInvLock(i - 1).Picture = Nothing
    Next i
    Dim EmptySlot As Slot
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
    For i = 1 To MAX_SKINSINVENTORY_SLOTS
        Call frmSkins.InvSkins.ClearSlot(i)
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
    frmMain.TimerNiebla.enabled = False
    bNiebla = False
    bNieve = False
    bFogata = False
    SkillPoints = 0
    UserStats.estado = 0
    Group.Clear
    InviCounter = 0
    DrogaCounter = 0
    frmMain.Contadores.enabled = False
    InvasionActual = 0
    frmMain.Evento.enabled = False
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
    Call EndAntiCheatSession
    Call ClearHotkeys
    'Unload all forms except frmMain and frmConnect
    Dim Frm As Form
    For Each Frm In Forms
        If ShouldUnloadForm(Frm.Name) Then
            Unload Frm
        End If
    Next
    #If PYMMO = 1 Then
        If g_game_state.State <> e_state_createchar_screen Then
            g_game_state.State = e_state_account_screen
        End If
        If Not FullLogout Then
            'Si no es un deslogueo completo, envío nuevamente la lista de Pjs.
            Call connectToLoginServer
        End If
    #ElseIf PYMMO = 0 Then
        If prgRun Then
            Call General_Set_Connect
        End If
    #End If
    Exit Sub
HandleDisconnect_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleDisconnect", Erl)
End Sub

Private Function ShouldUnloadForm(ByVal FormName As String) As Boolean
    If FormName = frmMain.Name Then Exit Function
    If FormName = frmConnect.Name Then Exit Function
    If FormName = frmMensaje.Name Then Exit Function
    ShouldUnloadForm = True
End Function

''
' Handles the CommerceEnd message.
Private Sub HandleCommerceEnd()
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
    Dim i       As Long
    Dim NpcName As String
    NpcName = Reader.ReadString8()
    'Fill our inventory list
    For i = 1 To MAX_INVENTORY_SLOTS
        With frmMain.Inventario
            Call frmComerciar.InvComUsu.SetItem(i, .ObjIndex(i), .Amount(i), .Equipped(i), .GrhIndex(i), .ObjType(i), .MaxHit(i), .MinHit(i), .Def(i), .Valor(i), .ItemName(i), _
                    .ElementalTags(i), .PuedeUsar(i))
        End With
    Next i
    'Set state and show form
    Comerciando = True
    'Call Inventario.Initialize(frmComerciar.PicInvUser)
    frmComerciar.Show , GetGameplayForm()
    frmComerciar.Refresh
    Exit Sub
HandleCommerceInit_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCommerceInit", Erl)
End Sub

''
' Handles the BankInit message.
Private Sub HandleBankInit()
    On Error GoTo HandleBankInit_Err
    Dim i As Long
    'Fill our inventory list
    For i = 1 To MAX_INVENTORY_SLOTS
        With frmMain.Inventario
            Call frmBancoObj.InvBankUsu.SetItem(i, .ObjIndex(i), .Amount(i), .Equipped(i), .GrhIndex(i), .ObjType(i), .MaxHit(i), .MinHit(i), .Def(i), .Valor(i), .ItemName(i), _
                    .ElementalTags(i), .PuedeUsar(i))
        End With
    Next i
    'Set state and show form
    Comerciando = True
    frmBancoObj.lblcosto = PonerPuntos(UserStats.GLD)
    frmBancoObj.Show , GetGameplayForm()
    frmBancoObj.Refresh
    Exit Sub
HandleBankInit_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBankInit", Erl)
End Sub

Private Sub HandleGoliath()
    On Error GoTo HandleGoliathInit_Err
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
    '
    FrmLogear.Show , frmConnect
    Exit Sub
HandleShowFrmLogear_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowFrmLogear", Erl)
End Sub

Private Sub HandleShowFrmMapa()
    On Error GoTo HandleShowFrmMapa_Err
    '
    ExpMult = Reader.ReadInt16()
    OroMult = Reader.ReadInt16()
    Call frmMapaGrande.CalcularPosicionMAPA
    frmMapaGrande.Picture = LoadInterface("ventanamapa.bmp")
    frmMapaGrande.Show , GetGameplayForm()
    Exit Sub
HandleShowFrmMapa_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowFrmMapa", Erl)
End Sub

''
' Handles the UserCommerceInit message.
Private Sub HandleUserCommerceInit()
    On Error GoTo HandleUserCommerceInit_Err
    Dim i As Long
    'Clears lists if necessary
    'Fill inventory list
    For i = 1 To MAX_INVENTORY_SLOTS
        With frmMain.Inventario
            Call frmComerciarUsu.InvUser.SetItem(i, .ObjIndex(i), .Amount(i), .Equipped(i), .GrhIndex(i), .ObjType(i), .MaxHit(i), .MinHit(i), .Def(i), .Valor(i), .ItemName(i), _
                    .ElementalTags(i), .PuedeUsar(i))
        End With
    Next i
    frmComerciarUsu.lblMyGold.Caption = PonerPuntos(UserStats.GLD)
    Dim J As Byte
    For J = 1 To 6
        Call frmComerciarUsu.InvOtherSell.SetItem(J, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0, 0)
        Call frmComerciarUsu.InvUserSell.SetItem(J, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0, 0)
    Next J
    'Set state and show form
    Comerciando = True
    frmComerciarUsu.Show , GetGameplayForm()
    Exit Sub
HandleUserCommerceInit_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUserCommerceInit", Erl)
End Sub

''
' Handles the UserCommerceEnd message.
Private Sub HandleUserCommerceEnd()
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
    On Error GoTo HandleShowBlacksmithForm_Err
    If frmMain.macrotrabajo.enabled And (MacroBltIndex > 0) Then
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
        frmHerrero.Show , GetGameplayForm()
    End If
    Exit Sub
HandleShowBlacksmithForm_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowBlacksmithForm", Erl)
End Sub

''
' Handles the ShowCarpenterForm message.
Private Sub HandleShowCarpenterForm()
    On Error GoTo HandleShowCarpenterForm_Err
    COLOR_AZUL = RGB(0, 0, 0)
    ' establece el borde al listbox
    Call Establecer_Borde(frmCarp.lstArmas, frmCarp, COLOR_AZUL, 0, 0)
    Call Establecer_Borde(frmCarp.List1, frmCarp, COLOR_AZUL, 0, 0)
    Call Establecer_Borde(frmCarp.List2, frmCarp, COLOR_AZUL, 0, 0)
    frmCarp.Show , GetGameplayForm()
    ' End If
    Exit Sub
HandleShowCarpenterForm_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowCarpenterForm", Erl)
End Sub

Private Sub HandleShowAlquimiaForm()
    On Error GoTo HandleShowAlquimiaForm_Err
    If frmMain.macrotrabajo.enabled And (MacroBltIndex > 0) Then
        Call WriteCraftAlquimista(MacroBltIndex)
    Else
        frmAlqui.Picture = LoadInterface("ventanaalquimia.bmp")
        COLOR_AZUL = RGB(0, 0, 0)
        ' establece el borde al listbox
        Call Establecer_Borde(frmAlqui.lstArmas, frmAlqui, COLOR_AZUL, 1, 1)
        Call Establecer_Borde(frmAlqui.List1, frmAlqui, COLOR_AZUL, 1, 1)
        Call Establecer_Borde(frmAlqui.List2, frmAlqui, COLOR_AZUL, 1, 1)
        frmAlqui.Show , GetGameplayForm()
    End If
    Exit Sub
HandleShowAlquimiaForm_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowAlquimiaForm", Erl)
End Sub

Private Sub HandleShowSastreForm()
    On Error GoTo HandleShowSastreForm_Err
    If frmMain.macrotrabajo.enabled And (MacroBltIndex > 0) Then
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
        FrmSastre.Show , GetGameplayForm()
    End If
    Exit Sub
HandleShowSastreForm_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowSastreForm", Erl)
End Sub

''
' Handles the NPCKillUser message.
Private Sub HandleNPCKillUser()
    On Error GoTo HandleNPCKillUser_Err
    Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_CRIATURA_MATADO"), 255, 0, 0, True, False, False)
    Exit Sub
HandleNPCKillUser_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleNPCKillUser", Erl)
End Sub

''
' Handles the BlockedWithShieldUser message.
Private Sub HandleBlockedWithShieldUser()
    On Error GoTo HandleBlockedWithShieldUser_Err
    Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_RECHAZO_ATAQUE_ESCUDO"), 255, 0, 0, True, False, False)
    Exit Sub
HandleBlockedWithShieldUser_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBlockedWithShieldUser", Erl)
End Sub

''
' Handles the BlockedWithShieldOther message.
Private Sub HandleBlockedWithShieldOther()
    On Error GoTo HandleBlockedWithShieldOther_Err
    Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO"), 255, 0, 0, True, False, False)
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
    Dim NotificoTexto As Boolean
    NotificoTexto = Reader.ReadBool
    With charlist(charindex)
        If ShowText And NotificoTexto Then
            Call SetCharacterDialogFx(charindex, IIf(charindex = UserCharIndex, (JsonLanguage.Item("MENSAJE_FALLAS")), (JsonLanguage.Item("MENSAJE_FALLO"))), RGBA_From_Comp(255, _
                    0, 0))
        End If
        If EstaPCarea(charindex) Then
            Call ao20audio.PlayWav(2, False, ao20audio.ComputeCharFxVolume(.Pos), ao20audio.ComputeCharFxPan(.Pos))
        End If
    End With
    Exit Sub
HandleCharSwing_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCharSwing", Erl)
End Sub

''
' Handles the SafeModeOn message.
Private Sub HandleSafeModeOn()
    On Error GoTo HandleSafeModeOn_Err
    SeguroGame = True
    Call frmMain.DibujarSeguro
    Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_SEGURO_ACTIVADO"), 65, 190, 156, False, False, False)
    Exit Sub
HandleSafeModeOn_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSafeModeOn", Erl)
End Sub

''
' Handles the SafeModeOff message.
Private Sub HandleSafeModeOff()
    On Error GoTo HandleSafeModeOff_Err
    SeguroGame = False
    Call frmMain.DesDibujarSeguro
    Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_SEGURO_DESACTIVADO"), 65, 190, 156, False, False, False)
    Exit Sub
HandleSafeModeOff_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSafeModeOff", Erl)
End Sub

''
' Handles the ResuscitationSafeOff message.
Private Sub HandlePartySafeOff()
    On Error GoTo HandlePartySafeOff_Err
    Call frmMain.ControlSeguroParty(False)
    Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_SEGURO_PARTY_OFF"), 250, 250, 0, False, True, False)
    Exit Sub
HandlePartySafeOff_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePartySafeOff", Erl)
End Sub

Private Sub HandleClanSeguro()
    On Error GoTo HandleClanSeguro_Err
    Dim Seguro As Boolean
    'Get data and update form
    Seguro = Reader.ReadBool()
    If SeguroClanX Then
        Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_SEGURO_CLAN_DESACTIVADO"), 65, 190, 156, False, False, False)
        frmMain.ImgSegClan = LoadInterface("boton-seguro-clan-off.bmp")
        SeguroClanX = False
    Else
        Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_SEGURO_CLAN_ACTIVADO"), 65, 190, 156, False, False, False)
        frmMain.ImgSegClan = LoadInterface("boton-seguro-clan-on.bmp")
        SeguroClanX = True
    End If
    Exit Sub
HandleClanSeguro_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleClanSeguro", Erl)
End Sub

Private Sub HandleIntervals()
    On Error GoTo HandleIntervals_Err
    gIntervals.Bow = Reader.ReadInt32()
    gIntervals.Walk = Reader.ReadInt32()
    gIntervals.Hit = Reader.ReadInt32()
    gIntervals.HitMagic = Reader.ReadInt32()
    gIntervals.Magic = Reader.ReadInt32()
    gIntervals.MagicHit = Reader.ReadInt32()
    gIntervals.HitUseItem = Reader.ReadInt32()
    gIntervals.ExtractWork = Reader.ReadInt32()
    gIntervals.BuildWork = Reader.ReadInt32()
    gIntervals.UseItemKey = Reader.ReadInt32()
    gIntervals.UseItemClick = Reader.ReadInt32()
    gIntervals.DropItem = Reader.ReadInt32()
    'Set the intervals of timers
    Call MainTimer.SetInterval(TimersIndex.Attack, gIntervals.Hit)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithU, gIntervals.UseItemKey)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithDblClick, gIntervals.UseItemClick)
    Call MainTimer.SetInterval(TimersIndex.SendRPU, INT_SENTRPU)
    Call MainTimer.SetInterval(TimersIndex.CastSpell, gIntervals.Magic)
    Call MainTimer.SetInterval(TimersIndex.Arrows, gIntervals.Bow)
    Call MainTimer.SetInterval(TimersIndex.CastAttack, gIntervals.MagicHit)
    Call MainTimer.SetInterval(TimersIndex.AttackSpell, gIntervals.HitMagic)
    Call MainTimer.SetInterval(TimersIndex.AttackUse, gIntervals.HitUseItem)
    Call MainTimer.SetInterval(TimersIndex.Drop, gIntervals.DropItem)
    Call MainTimer.SetInterval(TimersIndex.Walk, gIntervals.Walk)
    'Init timers
    Call MainTimer.start(TimersIndex.Attack)
    Call MainTimer.start(TimersIndex.UseItemWithU)
    Call MainTimer.start(TimersIndex.UseItemWithDblClick)
    Call MainTimer.start(TimersIndex.SendRPU)
    Call MainTimer.start(TimersIndex.CastSpell)
    Call MainTimer.start(TimersIndex.Arrows)
    Call MainTimer.start(TimersIndex.CastAttack)
    Call MainTimer.start(TimersIndex.AttackSpell)
    Call MainTimer.start(TimersIndex.AttackUse)
    Call MainTimer.start(TimersIndex.Drop)
    Call MainTimer.start(TimersIndex.Walk)
    Exit Sub
HandleIntervals_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleIntervals", Erl)
End Sub

Private Sub HandleUpdateUserKey()
    On Error GoTo HandleUpdateUserKey_Err
    Dim Slot As Integer, Llave As Integer
    Slot = Reader.ReadInt16
    Llave = Reader.ReadInt16
    Call FrmKeyInv.InvKeys.SetItem(Slot, Llave, 1, 0, ObjData(Llave).GrhIndex, eObjType.otLlaves, 0, 0, 0, 0, ObjData(Llave).Name, 0, 0)
    Exit Sub
HandleUpdateUserKey_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateUserKey", Erl)
End Sub

Private Sub HandleUpdateDM()
    On Error GoTo HandleUpdateDM_Err
    Dim value As Integer
    value = Reader.ReadInt16
    frmMain.lbldm = "+" & value & "%"
    Exit Sub
HandleUpdateDM_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateDM", Erl)
End Sub

Private Sub HandleUpdateRM()
    On Error GoTo HandleUpdateRM_Err
    Dim value As Integer
    value = Reader.ReadInt16
    frmMain.lblResis = "+" & value & "%"
    Exit Sub
HandleUpdateRM_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateRM", Erl)
End Sub

' Handles the ResuscitationSafeOn message.
Private Sub HandlePartySafeOn()
    On Error GoTo HandlePartySafeOn_Err
    Call frmMain.ControlSeguroParty(True)
    Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_SEGURO_PARTY_ON"), 250, 250, 0, False, True, False)
    Exit Sub
HandlePartySafeOn_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePartySafeOn", Erl)
End Sub

''
' Handles the CantUseWhileMeditating message.
Private Sub HandleCantUseWhileMeditating()
    On Error GoTo HandleCantUseWhileMeditating_Err
    Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_USAR_MEDITANDO"), 255, 0, 0, False, False, False)
    Exit Sub
HandleCantUseWhileMeditating_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCantUseWhileMeditating", Erl)
End Sub

''
' Handles the UpdateSta message.
Private Sub HandleUpdateSta()
    On Error GoTo HandleUpdateSta_Err
    'Get data and update form
    UserStats.MinSTA = Reader.ReadInt16()
    Call frmMain.UpdateStamina
    Exit Sub
HandleUpdateSta_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateSta", Erl)
End Sub

''
' Handles the UpdateMana message.
Private Sub HandleUpdateMana()
    On Error GoTo HandleUpdateMana_Err
    Dim OldMana As Integer
    OldMana = UserStats.minman
    'Get data and update form
    UserStats.minman = Reader.ReadInt16()
    If UserMeditar And UserStats.minman - OldMana > 0 And ChatCombate = 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg(JsonLanguage.Item("MENSAJE_GANAR_MANA") & (UserStats.minman - OldMana) & JsonLanguage.Item("MENSAJE_DE_MANA"), .red, .green, .blue, .bold, .italic)
        End With
    End If
    Call frmMain.UpdateManaBar
    Exit Sub
HandleUpdateMana_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateMana", Erl)
End Sub

Private Sub HandleUpdateHP()
    On Error GoTo HandleUpdateHP_Err
    Dim NuevoValor As Long
    Dim Shield     As Long
    NuevoValor = Reader.ReadInt16()
    Shield = Reader.ReadInt32
    ' Si perdió vida, mostramos los stats en el frmMain
    If NuevoValor < UserStats.MinHp Or Shield < UserStats.HpShield Then
        Call frmMain.ShowStats
    End If
    'Get data and update form
    UserStats.MinHp = NuevoValor
    UserStats.HpShield = Shield
    Call frmMain.UpdateHpBar
    'Is the user alive??
    If UserStats.MinHp = 0 Then
        #If DEBUGGING = 0 Then
            Call svb_unlock_achivement("Memento Mori")
        #End If
        UserStats.estado = 1
        charlist(UserCharIndex).Invisible = False
        If MostrarTutorial And tutorial_index <= 0 Then
            If tutorial(e_tutorialIndex.TUTORIAL_Muerto).Activo = 1 Then
                tutorial_index = e_tutorialIndex.TUTORIAL_Muerto
                Call mostrarCartel(tutorial(tutorial_index).titulo, tutorial(tutorial_index).textos(1), tutorial(tutorial_index).Grh, -1, &H164B8A, , , False, 100, 479, 100, _
                        535, 640, 530, 50, 100)
            End If
        End If
        DrogaCounter = 0
        Call deleteCharIndexs
    Else
        UserStats.estado = 0
    End If
    Exit Sub
HandleUpdateHP_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateHP", Erl)
End Sub

''
' Handles the UpdateGold message.
Private Sub HandleUpdateGold()
    On Error GoTo HandleUpdateGold_Err
    'Get data and update form
    UserStats.GLD = Reader.ReadInt32()
    UserStats.OroPorNivel = Reader.ReadInt32()
    Call frmMain.UpdateGoldState
    Exit Sub
HandleUpdateGold_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateGold", Erl)
End Sub

''
' Handles the UpdateExp message.
Private Sub HandleUpdateExp()
    On Error GoTo HandleUpdateExp_Err
    'Get data and update form
    UserStats.exp = Reader.ReadInt32()
    Call frmMain.UpdateExpBar
    Exit Sub
HandleUpdateExp_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateExp", Erl)
End Sub

Private Sub HandleChangeMap()
    On Error GoTo HandleChangeMap_Err
    UserMap = Reader.ReadInt16()
    ResourceMap = Reader.ReadInt16()
    If frmComerciar.visible Then Unload frmComerciar
    If frmBancoObj.visible Then Unload frmBancoObj
    If frmEstadisticas.visible Then Unload frmEstadisticas
    If frmStatistics.visible Then Unload frmStatistics
    If frmHerrero.visible Then Unload frmHerrero
    If FrmSastre.visible Then Unload FrmSastre
    If frmAlqui.visible Then Unload frmAlqui
    If frmCarp.visible Then Unload frmCarp
    If FrmGrupo.visible Then Unload FrmGrupo
    If frmGoliath.visible Then Unload frmGoliath
    If FrmViajes.visible Then Unload FrmViajes
    If frmCantidad.visible Then Unload frmCantidad
    Call SwitchMap(UserMap, ResourceMap)
    Exit Sub
HandleChangeMap_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChangeMap", Erl)
End Sub

''
' Handles the PosUpdate message.
Private Sub HandlePosUpdate()
    On Error GoTo HandlePosUpdate_Err
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
    UpdatePlayerRoof
    'Update pos label and minimap
    Call UpdateMapPos
    Call RefreshAllChars
    Exit Sub
HandlePosUpdate_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePosUpdate", Erl)
End Sub

''
' Handles the PosUpdate message.
Private Sub HandlePosUpdateUserChar()
    On Error GoTo HandlePosUpdateUserChar_Err
    Dim temp_x As Byte, temp_y As Byte
    temp_x = UserPos.x
    temp_y = UserPos.y
    'Set new pos
    UserPos.x = Reader.ReadInt8()
    UserPos.y = Reader.ReadInt8()
    Dim charindex As Integer
    charindex = Reader.ReadInt16()
    'Remove char from old position
    If MapData(temp_x, temp_y).charindex = charindex Then
        MapData(temp_x, temp_y).charindex = 0
    End If
    'Set char
    MapData(UserPos.x, UserPos.y).charindex = charindex
    charlist(charindex).Pos = UserPos
    'Are we under a roof?
    UpdatePlayerRoof
    Call RefreshAllChars
    Exit Sub
HandlePosUpdateUserChar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePosUpdateUserChar", Erl)
End Sub

''
' Handles the NPCHitUser message.
''
' Handles the PosUpdate message.
Private Sub HandlePosUpdateChar()
    On Error GoTo HandlePosUpdateChar_Err
    Dim charindex As Integer
    Dim x         As Byte, y As Byte
    'Set new pos
    charindex = Reader.ReadInt16()
    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
    If charindex = 0 Then Exit Sub
    If charlist(charindex).Pos.x > 0 And charlist(charindex).Pos.y > 0 Then
        If MapData(charlist(charindex).Pos.x, charlist(charindex).Pos.y).charindex = charindex Then
            MapData(charlist(charindex).Pos.x, charlist(charindex).Pos.y).charindex = 0
        End If
        MapData(x, y).charindex = charindex
        charlist(charindex).Pos.x = x
        charlist(charindex).Pos.y = y
        charlist(charindex).MoveOffsetX = 0
        charlist(charindex).MoveOffsetY = 0
    End If
    Exit Sub
HandlePosUpdateChar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePosUpdateChar", Erl)
End Sub

''
' Handles the NPCHitUser message.
Private Sub HandleNPCHitUser()
    On Error GoTo HandleNPCHitUser_Err
    Dim Lugar As Byte, DañoStr As String
    Lugar = Reader.ReadInt8()
    DañoStr = PonerPuntos(Reader.ReadInt16)
    Select Case Lugar
        Case bCabeza
            Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_GOLPE_CABEZA") & DañoStr, 255, 0, 0, True, False, False)
        Case bBrazoIzquierdo
            Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_GOLPE_BRAZO_IZQ") & DañoStr, 255, 0, 0, True, False, False)
        Case bBrazoDerecho
            Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_GOLPE_BRAZO_DER") & DañoStr, 255, 0, 0, True, False, False)
        Case bPiernaIzquierda
            Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_GOLPE_PIERNA_IZQ") & DañoStr, 255, 0, 0, True, False, False)
        Case bPiernaDerecha
            Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_GOLPE_PIERNA_DER") & DañoStr, 255, 0, 0, True, False, False)
        Case bTorso
            Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_GOLPE_TORSO") & DañoStr, 255, 0, 0, True, False, False)
    End Select
    #If DEBUGGING = 0 Then
        Call svb_unlock_achivement("Small victory")
    #End If
    Exit Sub
HandleNPCHitUser_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleNPCHitUser", Erl)
End Sub

''
' Handles the UserHittingByUser message.
Private Sub HandleUserHittedByUser()
    On Error GoTo HandleUserHittedByUser_Err
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
            Call AddtoRichTextBox(frmMain.RecTxt, attacker & JsonLanguage.Item("MENSAJE_RECIVE_IMPACTO_CABEZA") & DañoStr & JsonLanguage.Item("MENSAJE_2"), 255, 0, 0, True, _
                    False, False)
        Case bBrazoIzquierdo
            Call AddtoRichTextBox(frmMain.RecTxt, attacker & JsonLanguage.Item("MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ") & DañoStr & JsonLanguage.Item("MENSAJE_2"), 255, 0, 0, True, _
                    False, False)
        Case bBrazoDerecho
            Call AddtoRichTextBox(frmMain.RecTxt, attacker & JsonLanguage.Item("MENSAJE_RECIVE_IMPACTO_BRAZO_DER") & DañoStr & JsonLanguage.Item("MENSAJE_2"), 255, 0, 0, True, _
                    False, False)
        Case bPiernaIzquierda
            Call AddtoRichTextBox(frmMain.RecTxt, attacker & JsonLanguage.Item("MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ") & DañoStr & JsonLanguage.Item("MENSAJE_2"), 255, 0, 0, True, _
                    False, False)
        Case bPiernaDerecha
            Call AddtoRichTextBox(frmMain.RecTxt, attacker & JsonLanguage.Item("MENSAJE_RECIVE_IMPACTO_PIERNA_DER") & DañoStr & JsonLanguage.Item("MENSAJE_2"), 255, 0, 0, True, _
                    False, False)
        Case bTorso
            Call AddtoRichTextBox(frmMain.RecTxt, attacker & JsonLanguage.Item("MENSAJE_RECIVE_IMPACTO_TORSO") & DañoStr & JsonLanguage.Item("MENSAJE_2"), 255, 0, 0, True, _
                    False, False)
    End Select
    Exit Sub
HandleUserHittedByUser_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUserHittedByUser", Erl)
End Sub

''
' Handles the UserHittedUser message.
Private Sub HandleUserHittedUser()
    On Error GoTo HandleUserHittedUser_Err
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
            Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_PRODUCE_IMPACTO_1") & victim & JsonLanguage.Item("MENSAJE_PRODUCE_IMPACTO_CABEZA") & DañoStr & _
                    JsonLanguage.Item("MENSAJE_2"), 255, 0, 0, True, False, False)
        Case bBrazoIzquierdo
            Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_PRODUCE_IMPACTO_1") & victim & JsonLanguage.Item("MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ") & DañoStr & _
                    JsonLanguage.Item("MENSAJE_2"), 255, 0, 0, True, False, False)
        Case bBrazoDerecho
            Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_PRODUCE_IMPACTO_1") & victim & JsonLanguage.Item("MENSAJE_PRODUCE_IMPACTO_BRAZO_DER") & DañoStr & _
                    JsonLanguage.Item("MENSAJE_2"), 255, 0, 0, True, False, False)
        Case bPiernaIzquierda
            Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_PRODUCE_IMPACTO_1") & victim & JsonLanguage.Item("MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ") & DañoStr & _
                    JsonLanguage.Item("MENSAJE_2"), 255, 0, 0, True, False, False)
        Case bPiernaDerecha
            Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_PRODUCE_IMPACTO_1") & victim & JsonLanguage.Item("MENSAJE_PRODUCE_IMPACTO_PIERNA_DER") & DañoStr & _
                    JsonLanguage.Item("MENSAJE_2"), 255, 0, 0, True, False, False)
        Case bTorso
            Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_PRODUCE_IMPACTO_1") & victim & JsonLanguage.Item("MENSAJE_PRODUCE_IMPACTO_TORSO") & DañoStr & _
                    JsonLanguage.Item("MENSAJE_2"), 255, 0, 0, True, False, False)
    End Select
    Exit Sub
HandleUserHittedUser_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUserHittedUser", Erl)
End Sub

Private Sub HandleChatOverHeadImpl(ByVal chat As String, _
                                   ByVal charindex As Integer, _
                                   ByVal color As Long, _
                                   ByVal EsSpell As Boolean, _
                                   ByVal x As Byte, _
                                   ByVal y As Byte, _
                                   ByVal RequiredMinDisplayTime As Integer, _
                                   ByVal MaxDisplayTime As Integer)
    On Error GoTo errhandler
    Dim QueEs As String
    If x + y > 0 Then
        With charlist(charindex)
            If .Invisible And charindex <> UserCharIndex Then
                If MapData(.Pos.x, .Pos.y).charindex = charindex Then MapData(.Pos.x, .Pos.y).charindex = 0
                .Pos.x = x
                .Pos.y = y
                MapData(x, y).charindex = charindex
            End If
        End With
    End If
    QueEs = ReadField(1, chat, Asc("*"))
    Dim TextColor As RGBA
    TextColor = RGBA_From_Long(color)
    Dim copiar As Boolean
    copiar = True
    Dim duracion As Integer
    duracion = 250
    Dim text As String
    text = ReadField(2, chat, Asc("*"))
    Select Case QueEs
        Case "LOCMSG"
            ' text = "2082*NombreDelNpc¬OtroValor"
            Dim MsgID    As Integer
            Dim extraStr As String
            MsgID = val(ReadField(1, text, Asc("*")))             ' 2082
            extraStr = ReadField(2, Text, Asc("*"))               ' "Nombre¬OtroValor"
            chat = Locale_Parse_ServerMessage(MsgID, extraStr)
            copiar = False
            duracion = 20
        Case "NPCDESC"
            chat = NpcData(text).desc
            copiar = False
            If npcs_en_render And tutorial_index <= 0 Then
                Dim headGrh As Long
                Dim bodyGrh As Long
                headGrh = HeadData(NpcData(text).Head).Head(3).GrhIndex
                bodyGrh = GrhData(BodyData(NpcData(text).Body).Walk(3).GrhIndex).Frames(1)
                If headGrh = 0 Then
                    Call mostrarCartel(Split(NpcData(text).Name, " <")(0), NpcData(text).desc, bodyGrh, 200 + 30 * Len(chat), &H164B8A, , , True, 100, 479, 100, 535, 20, 500, _
                            50, 80, bodyGrh, 1)
                Else
                    Dim HeadOffsetY As Integer
                    HeadOffsetY = CInt(BodyData(NpcData(text).Body).HeadOffset.y) - 30
                    Call mostrarCartel(Split(NpcData(text).Name, " <")(0), NpcData(text).desc, headGrh, 200 + 30 * Len(chat), &H164B8A, , , True, 100, 479, 100, 535, 20, 500, _
                            50, 100, bodyGrh, HeadOffsetY)
                End If
            End If
        Case "PMAG"
            chat = HechizoData(ReadField(2, chat, Asc("*"))).PalabrasMagicas
            If charlist(UserCharIndex).Muerto = True Then chat = ""
            copiar = False
            duracion = 20
        Case "QUESTFIN"
            chat = QuestList(ReadField(2, chat, Asc("*"))).DescFinal
            copiar = False
            duracion = 20
        Case "NOCONSOLA" ' El chat no sale en la consola
            chat = ReadField(2, chat, Asc("*"))
            copiar = False
            duracion = 20
    End Select
    'Only add the chat if the character exists (a CharacterRemove may have been sent to the PC / NPC area before the buffer was flushed)
    If charlist(charindex).active = 1 Then
        Call Char_Dialog_Set(charindex, chat, color, duracion, 30, 1, EsSpell, RequiredMinDisplayTime, MaxDisplayTime)
    End If
    If charlist(charindex).EsNpc = False Then
        If CopiarDialogoAConsola = 1 And copiar Then
            Call WriteChatOverHeadInConsole(charindex, chat, TextColor.R, TextColor.G, TextColor.B)
        End If
    End If
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChatOverHeadImpl", Erl)
End Sub

' Handles the ChatOverHead message.
Private Sub HandleLocaleChatOverHead()
    On Error GoTo errhandler
    Dim ChatId        As Integer
    Dim Params        As String
    Dim charindex     As Integer
    Dim TextColor     As Long
    Dim IsSpell       As Boolean
    Dim x             As Byte, y As Byte
    Dim MinChatTime   As Integer
    Dim MaxChatTime   As Integer
    Dim LocalizedText As String
    ChatId = Reader.ReadInt16
    Params = Reader.ReadString8
    charindex = Reader.ReadInt16()
    TextColor = vbColor_2_Long(Reader.ReadInt32())
    IsSpell = Reader.ReadBool()
    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
    MinChatTime = Reader.ReadInt16()
    MaxChatTime = Reader.ReadInt16()
    LocalizedText = Locale_Parse_ServerMessage(ChatId, Params)
    Call HandleChatOverHeadImpl(LocalizedText, charindex, TextColor, IsSpell, x, y, MinChatTime, MaxChatTime)
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleLocaleChatOverHead", Erl)
End Sub

' Handles the ChatOverHead message.
Private Sub HandleChatOverHead()
    On Error GoTo errhandler
    Dim chat        As String
    Dim charindex   As Integer
    Dim colortexto  As Long
    Dim QueEs       As String
    Dim EsSpell     As Boolean
    Dim x           As Byte, y As Byte
    Dim MinChatTime As Integer
    Dim MaxChatTime As Integer
    chat = Reader.ReadString8()
    charindex = Reader.ReadInt16()
    colortexto = vbColor_2_Long(Reader.ReadInt32())
    EsSpell = Reader.ReadBool()
    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
    MinChatTime = Reader.ReadInt16()
    MaxChatTime = Reader.ReadInt16()
    Call HandleChatOverHeadImpl(chat, charindex, colortexto, EsSpell, x, y, MinChatTime, MaxChatTime)
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChatOverHead", Erl)
End Sub

Private Sub HandleTextOverChar()
    On Error GoTo errhandler
    Dim chat      As String
    Dim charindex As Integer
    Dim color     As Long
    chat = Reader.ReadString8()
    charindex = Reader.ReadInt16()
    color = Reader.ReadInt32()
    Call SetCharacterDialogFx(charindex, chat, RGBA_From_vbColor(color))
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleTextOverChar", Erl)
End Sub

Private Sub HandleTextOverTile()
    On Error GoTo errhandler
    Dim text     As String
    Dim x        As Integer
    Dim y        As Integer
    Dim color    As Long
    Dim duration As Integer
    Dim OffsetY  As Integer
    Dim Animated As Boolean
    text = Reader.ReadString8()
    x = Reader.ReadInt16()
    y = Reader.ReadInt16()
    color = Reader.ReadInt32()
    duration = Reader.ReadInt16()
    OffsetY = Reader.ReadInt16()
    Animated = Reader.ReadBool()
    If InMapBounds(x, y) Then
        With MapData(x, y)
            Dim Index As Integer
            If UBound(.DialogEffects) = 0 Then
                ReDim .DialogEffects(1 To 1)
                Index = 1
            Else
                For Index = 1 To UBound(.DialogEffects)
                    If .DialogEffects(Index).text = vbNullString Then
                        Exit For
                    End If
                Next
                If Index > UBound(.DialogEffects) Then
                    ReDim Preserve .DialogEffects(1 To UBound(.DialogEffects) + 1)
                End If
            End If
            With .DialogEffects(Index)
                .color = RGBA_From_vbColor(color)
                .start = FrameTime
                .text = text
                .offset.x = 0
                .offset.y = OffsetY
                .duration = duration
                .Animated = Animated
            End With
        End With
    End If
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleTextOverTile", Erl)
End Sub

Private Sub HandleTextCharDrop()
    On Error GoTo errhandler
    Dim text      As String
    Dim charindex As Integer
    Dim color     As Long
    Dim duration  As Integer
    Dim Animated  As Boolean
    text = Reader.ReadString8()
    charindex = Reader.ReadInt16()
    color = Reader.ReadInt32()
    duration = Reader.ReadInt16()
    Animated = Reader.ReadBool()
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
                    If .DialogEffects(Index).text = vbNullString Then
                        Exit For
                    End If
                Next
                If Index > UBound(.DialogEffects) Then
                    ReDim .DialogEffects(1 To UBound(.DialogEffects) + 1)
                End If
            End If
            With .DialogEffects(Index)
                .color = RGBA_From_vbColor(color)
                .start = FrameTime
                .text = text
                .offset.x = OffsetX
                .offset.y = OffsetY
                .duration = duration
                .Animated = Animated
            End With
        End With
    End If
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleTextCharDrop", Erl)
End Sub

Private Sub HandleConsoleCharText()
    Dim text         As String
    Dim color        As Long
    Dim SourceName   As String
    Dim SourceStatus As Integer
    Dim Privileges   As Integer
    text = Reader.ReadString8()
    color = Reader.ReadInt32()
    SourceName = Reader.ReadString8()
    SourceStatus = Reader.ReadInt16()
    Privileges = Reader.ReadInt16()
    If Privileges > 0 Then
        Privileges = Log(Privileges) / Log(2)
    End If
    Dim TextColor As RGBA
    TextColor = RGBA_From_vbColor(color)
    Call WriteConsoleUserChat(text, SourceName, TextColor.R, TextColor.G, TextColor.B, SourceStatus, Privileges)
End Sub

''
' Handles the ConsoleMessage message.
Private Sub HandleConsoleMessage()
    On Error GoTo errhandler
    Dim chat           As String
    Dim FontIndex      As Integer
    Dim str            As String
    Dim R              As Byte
    Dim G              As Byte
    Dim B              As Byte
    Dim QueEs          As String
    Dim NpcName        As String
    Dim npcElementTags As Long
    Dim objname        As String
    Dim ElementalTags  As Long
    Dim quantity       As Integer
    Dim Hechizo        As Integer
    Dim userName       As String
    Dim Valor          As String
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
            'natural item elemental tags logical or with rune imbued item
            ElementalTags = CLng(val(ObjData(ReadField(2, chat, Asc("*"))).ElementalTags)) Or CLng(val(ReadField(4, chat, Asc("*"))))
            chat = objname & " " & ElementalTagsToTxtParser(ElementalTags) & ReadField(3, chat, Asc("*"))
        Case "HECINF"
            Hechizo = ReadField(2, chat, Asc("*"))
            chat = "------------< Información del hechizo >------------" & vbCrLf & "Nombre: " & HechizoData(Hechizo).nombre & vbCrLf & "Descripción: " & HechizoData( _
                    Hechizo).desc & vbCrLf & "Skill requerido: " & HechizoData(Hechizo).MinSkill & " de magia." & vbCrLf & "Mana necesario: " & HechizoData( _
                    Hechizo).ManaRequerido & " puntos." & vbCrLf & "Stamina necesaria: " & HechizoData(Hechizo).StaRequerido & " puntos."
        Case "ProMSG"
            Hechizo = ReadField(2, chat, Asc("*"))
            chat = HechizoData(Hechizo).PropioMsg
        Case "HecMSG"
            Hechizo = ReadField(2, chat, Asc("*"))
            chat = HechizoData(Hechizo).HechizeroMsg & " la criatura."
        Case "HecMSGU"
            Hechizo = ReadField(2, chat, Asc("*"))
            userName = ReadField(3, chat, Asc("*"))
            chat = HechizoData(Hechizo).HechizeroMsg & " " & userName & "."
        Case "HecMSGA"
            Hechizo = ReadField(2, chat, Asc("*"))
            userName = ReadField(3, chat, Asc("*"))
            chat = userName & " " & HechizoData(Hechizo).TargetMsg
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
        If val(str) > 255 Then
            R = 255
        Else
            R = val(str)
        End If
        str = ReadField(3, chat, 126)
        If val(str) > 255 Then
            G = 255
        Else
            G = val(str)
        End If
        str = ReadField(4, chat, 126)
        If val(str) > 255 Then
            B = 255
        Else
            B = val(str)
        End If
        Call AddtoRichTextBox(frmMain.RecTxt, Left$(chat, InStr(1, chat, "~") - 1), R, G, B, val(ReadField(5, chat, 126)) <> 0, val(ReadField(6, chat, 126)) <> 0)
    Else
        With FontTypes(FontIndex)
            Call AddtoRichTextBox(frmMain.RecTxt, chat, .red, .green, .blue, .bold, .italic)
            If EsGM Then
                Call frmPanelgm.CadenaChat(chat)
            End If
        End With
    End If
    Exit Sub
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleConsoleMessage", Erl)
End Sub

Private Sub HandleConsoleFactionMessage()
    On Error GoTo errhandler
    Dim chat         As String
    Dim FontIndex    As Integer
    Dim factionLabel As String
    chat = Reader.ReadString8()
    FontIndex = Reader.ReadInt8()
    factionLabel = Reader.ReadString8()
    'Si tiene el chat global desactivado, no se le muestran los mensajes faccionarios tampoco
    If ChatGlobal = 0 Then Exit Sub
    With FontTypes(FontIndex)
        Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item(factionLabel) & chat, .red, .green, .blue, .bold, .italic)
    End With
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleConsoleFactionMessage", Erl)
End Sub

Private Sub HandleLocaleMsg()
    On Error GoTo errhandler
    Dim chat      As String
    Dim FontIndex As Integer
    Dim str       As String
    Dim R         As Byte
    Dim G         As Byte
    Dim B         As Byte
    Dim id        As Integer
    id = Reader.ReadInt16()
    chat = Reader.ReadString8()
    FontIndex = Reader.ReadInt8()
    chat = Locale_Parse_ServerMessage(id, chat)
    If InStr(1, chat, "~") Then
        str = ReadField(2, chat, 126)
        If val(str) > 255 Then
            R = 255
        Else
            R = val(str)
        End If
        str = ReadField(3, chat, 126)
        If val(str) > 255 Then
            G = 255
        Else
            G = val(str)
        End If
        str = ReadField(4, chat, 126)
        If val(str) > 255 Then
            B = 255
        Else
            B = val(str)
        End If
        Call AddtoRichTextBox(frmMain.RecTxt, Left$(chat, InStr(1, chat, "~") - 1), R, G, B, val(ReadField(5, chat, 126)) <> 0, val(ReadField(6, chat, 126)) <> 0)
    Else
        With FontTypes(FontIndex)
            Call AddtoRichTextBox(frmMain.RecTxt, chat, .red, .green, .blue, .bold, .italic)
        End With
    End If
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleLocaleMsg", Erl)
End Sub

Private Sub HandleGuildChat()
    On Error GoTo errhandler
    Dim finalChat   As String
    Dim status      As Byte
    Dim R           As Byte, G As Byte, B As Byte
    Dim colorRed    As String, colorGreen As String, colorBlue As String
    Dim boldFlag    As Boolean, italicFlag As Boolean
    Dim prefix      As String
    Dim messageText As String
    status = Reader.ReadInt8()
    finalChat = Reader.ReadString8()
    ' Check for localized message format: Msg####¬Nombre
    If Left$(finalChat, 3) = "Msg" Then
        prefix = ReadField(1, finalChat, 172)
        finalChat = Locale_Parse_ServerMessage(val(mid$(prefix, 4)), mid$(finalChat, Len(prefix) + 2))
    End If
    ' If guild chat dialog is inactive, display in main chat window
    If Not DialogosClanes.Activo Then
        If InStr(1, finalChat, "~") > 0 Then
            ' Color-formatted chat: split fields and apply custom RGB
            colorRed = ReadField(2, finalChat, 126)
            colorGreen = ReadField(3, finalChat, 126)
            colorBlue = ReadField(4, finalChat, 126)
            R = IIf(val(colorRed) > 255, 255, val(colorRed))
            G = IIf(val(colorGreen) > 255, 255, val(colorGreen))
            B = IIf(val(colorBlue) > 255, 255, val(colorBlue))
            boldFlag = (val(ReadField(5, finalChat, 126)) <> 0)
            italicFlag = (val(ReadField(6, finalChat, 126)) <> 0)
            messageText = Left$(finalChat, InStr(1, finalChat, "~") - 1)
            Call AddtoRichTextBox(frmMain.RecTxt, messageText, R, G, B, boldFlag, italicFlag)
        Else
            ' Use default font style for guild messages
            With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
                Call AddtoRichTextBox(frmMain.RecTxt, finalChat, .red, .green, .blue, .bold, .italic)
            End With
        End If
    Else
        ' Chat is redirected to the guild dialog
        Call DialogosClanes.PushBackText(ReadField(1, finalChat, 126), status)
    End If
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildChat", Erl)
End Sub

Private Sub HandleShowMessageBox()
    On Error GoTo errhandler
    Dim mensaje   As String
    Dim MessageID As Integer
    Dim extra     As String
    ' Obtener el ID del mensaje desde el servidor
    MessageID = Reader.ReadInt16()
    extra = Reader.ReadString8()
    ' Obtener el mensaje a partir del archivo de localización usando el ID
    mensaje = Locale_Parse_ServerMessage(MessageID, extra)
    Select Case g_game_state.State()
        Case e_state_gameplay_screen
            frmMensaje.msg.Caption = mensaje
            frmMensaje.Show , GetGameplayForm()
        Case e_state_connect_screen
            Call ao20audio.PlayWav(SND_EXCLAMACION)
            Call TextoAlAsistente(mensaje, False, False)
            Call Long_2_RGBAList(textcolorAsistente, -1)
        Case e_state_account_screen
            frmMensaje.Show
            frmMensaje.msg.Caption = mensaje
        Case e_state_createchar_screen
            frmMensaje.Show , frmConnect
            frmMensaje.msg.Caption = mensaje
    End Select
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowMessageBox", Erl)
End Sub

Private Sub HandleMostrarCuenta()
    On Error GoTo errhandler
    AlphaNiebla = 30
    frmConnect.visible = True
    g_game_state.State = e_state_account_screen
    SugerenciaAMostrar = RandomNumber(1, NumSug)
    Call Graficos_Particulas.Particle_Group_Remove_All
    Call Graficos_Particulas.Engine_Select_Particle_Set(203)
    ParticleLluviaDorada = Graficos_Particulas.General_Particle_Create(208, -1, -1)
    If FrmLogear.visible Then
        Unload FrmLogear
        'Unload frmConnect
    End If
    If frmMain.visible Then
        '  frmMain.Visible = False
        UserParalizado = False
        UserInmovilizado = False
        UserStopped = False
        InvasionActual = 0
        frmMain.Evento.enabled = False
        'BUG CLONES
        Dim i As Integer
        For i = 1 To LastChar
            Call EraseChar(i)
        Next i
        frmMain.personaje(1).visible = False
        frmMain.personaje(2).visible = False
        frmMain.personaje(3).visible = False
        frmMain.personaje(4).visible = False
        frmMain.personaje(5).visible = False
    End If
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleMostrarCuenta", Erl)
End Sub

''
' Handles the UserIndexInServer message.
Private Sub HandleUserIndexInServer()
    On Error GoTo HandleUserIndexInServer_Err
    userIndex = Reader.ReadInt16()
    Exit Sub
HandleUserIndexInServer_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUserIndexInServer", Erl)
End Sub

''
' Handles the UserCharIndexInServer message.
Private Sub HandleUserCharIndexInServer()
    On Error GoTo HandleUserCharIndexInServer_Err
    UserCharIndex = Reader.ReadInt16()
    'frmdebug.add_text_tracebox "UserCharIndex " & UserCharIndex
    UserPos = charlist(UserCharIndex).Pos
    'Are we under a roof?
    UpdatePlayerRoof
    lastMove = FrameTime
    If MapDat.Seguro = 1 Then
        frmMain.Coord.ForeColor = RGB(0, 170, 0)
    Else
        frmMain.Coord.ForeColor = RGB(170, 0, 0)
    End If
    Call UpdateMapPos
    Exit Sub
HandleUserCharIndexInServer_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUserCharIndexInServer", Erl)
End Sub

''
' Handles the CharacterCreate message.
Private Sub HandleCharacterCreate()
    On Error GoTo errhandler
    Dim charindex     As Integer
    Dim Body          As Integer
    Dim Head          As Integer
    Dim Heading       As E_Heading
    Dim x             As Byte
    Dim y             As Byte
    Dim weapon        As Integer
    Dim Shield        As Integer
    Dim helmet        As Integer
    Dim privs         As Integer
    Dim Cart          As Integer
    Dim Backpack      As Integer
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
    Shield = Reader.ReadInt16()
    helmet = Reader.ReadInt16()
    Cart = Reader.ReadInt16()
    Backpack = Reader.ReadInt16()
    
    With charlist(charindex)
        Dim loopC, Fx As Integer
        Fx = Reader.ReadInt16
        loopC = Reader.ReadInt16
        Call StartFx(.ActiveAnimation, Fx, loopC)
        .Meditating = Fx <> 0
        Dim NombreYClan As String
        NombreYClan = Reader.ReadString8()   '
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
        .Invisible = False
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
        .DontBlockTile = False
        flags = Reader.ReadInt8()
        .Idle = flags And &O1
        .Navegando = flags And &O2
        .tipoUsuario = Reader.ReadInt8()
        .Team = Reader.ReadInt8()
        .banderaIndex = Reader.ReadInt8()
        .AnimAtaque1 = Reader.ReadInt16()
        
        If Backpack > 0 Then
            .Backpack = BodyData(Backpack)
            .tmpBackPack = Backpack
            .HasBackpack = True
        End If
        
        'dwarven exoesqueleton exception
        If .Body.BodyIndex = DwarvenExoesqueletonBody Then
            weapon = NO_WEAPON
            Shield = NO_SHIELD
            helmet = NO_HELMET
            Cart = NO_CART
            Backpack = NO_BACKPACK
            Head = 0
        End If
        If (.Pos.x <> 0 And .Pos.y <> 0) Then
            If MapData(.Pos.x, .Pos.y).charindex = charindex Then
                'Erase the old character from map
                MapData(charlist(charindex).Pos.x, charlist(charindex).Pos.y).charindex = 0
            End If
        End If
        If privs <> 0 Then
            'Log2 of the bit flags sent by the server gives our numbers ^^
            .priv = Log(privs) / Log(2)
        Else
            .priv = 0
        End If
        .Muerto = (Body = CASPER_BODY_IDLE)
        Call MakeChar(charindex, Body, Head, Heading, x, y, weapon, Shield, helmet, Cart, Backpack, ParticulaFx, appear)
        If .Navegando = False Or UserNadandoTrajeCaucho = True Then
            If .Body.AnimateOnIdle = 0 Then
                .Body.Walk(.Heading).started = 0
            ElseIf .Body.Walk(.Heading).started = 0 Then
                .Body.Walk(.Heading).started = FrameTime
            End If
            If Not .MovArmaEscudo Then
                .Arma.WeaponWalk(.Heading).started = 0
                .Escudo.ShieldWalk(.Heading).started = 0
            End If
            If .Body.IdleBody > 0 Then
                .Body = BodyData(.Body.IdleBody)
                .Body.Walk(.Heading).started = FrameTime
            End If
        End If
    End With
    Call RefreshAllChars
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCharacterCreate", Erl)
End Sub

Private Sub HandleUpdateFlag()
    On Error GoTo errhandler
    Dim charindex As Integer
    Dim flag      As Long
    charindex = Reader.ReadInt16()
    flag = Reader.ReadInt8()
    With charlist(charindex)
        .banderaIndex = flag
    End With
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleTextOverChar", Erl)
End Sub

Private Sub HandleCharacterRemove()
    On Error GoTo HandleCharacterRemove_Err
    Dim charindex   As Integer
    Dim Desvanecido As Boolean
    Dim fueWarp     As Boolean
    charindex = Reader.ReadInt16()
    Desvanecido = Reader.ReadBool()
    fueWarp = Reader.ReadBool()
    If Desvanecido And charlist(charindex).EsNpc = True Then
        Call CrearFantasma(charindex)
    End If
    Call EraseChar(charindex, fueWarp)
    Call RefreshAllChars
    Call ao20audio.StopAllWavsMatchingLabel("meditate" & CStr(charindex))
    Exit Sub
HandleCharacterRemove_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCharacterRemove", Erl)
End Sub

''
' Handles the CharacterMove message.
Private Sub HandleCharacterMove()
    On Error GoTo HandleCharacterMove_Err
    Dim charindex As Integer
    Dim x         As Byte
    Dim y         As Byte
    Dim dir       As Byte
    charindex = Reader.ReadInt16()
    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
    Call Char_Move_by_Pos(charindex, x, y)
    Call RefreshAllChars
    With charlist(charindex)
        ' Play steps sounds if the user is not an admin of any kind
        If .priv <> 1 And .priv <> 2 And .priv <> 3 And .priv <> 5 And .priv <> 25 Then
            Call DoPasosFx(charindex)
        End If
    End With
    Exit Sub
HandleCharacterMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCharacterMove", Erl)
End Sub

Private Sub HandleCharacterTranslate()
    On Error GoTo HandleCharacterTranslate_Err
    Dim charindex       As Integer
    Dim TileX           As Byte
    Dim TileY           As Byte
    Dim TranslationTime As Long
    charindex = Reader.ReadInt16()
    TileX = Reader.ReadInt8()
    TileY = Reader.ReadInt8()
    TranslationTime = Reader.ReadInt32()
    Call TranslateCharacterToPos(charindex, TileX, TileY, TranslationTime)
    Call RefreshAllChars
    Exit Sub
HandleCharacterTranslate_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCharacterTranslate", Erl)
End Sub

''
' Handles the ForceCharMove message.
Private Sub HandleForceCharMove()
    On Error GoTo HandleForceCharMove_Err
    Dim direccion As Byte
    direccion = Reader.ReadInt8()
    Moviendose = True
    Call MainTimer.Restart(TimersIndex.Walk)
    Call Char_Move_by_Head(UserCharIndex, direccion)
    Call MoveScreen(direccion)
    Call UpdateMapPos
    If MapDat.Seguro = 1 Then
        frmMain.Coord.ForeColor = RGB(0, 170, 0)
    Else
        frmMain.Coord.ForeColor = RGB(170, 0, 0)
    End If
    Call RefreshAllChars
    Exit Sub
HandleForceCharMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleForceCharMove", Erl)
End Sub


' Handles the CharacterChange message.
Private Sub HandleCharacterChange()
    
    On Error GoTo HandleCharacterChange_Err
    Dim charindex As Integer
    Dim TempInt   As Integer
    Dim headIndex As Integer
    charindex = Reader.ReadInt16()
    With charlist(charindex)
        ' ===== Preservar estado previo para fase =====
        Dim wasMoving        As Boolean: wasMoving = .Moving
        Dim oldHeading       As E_Heading: oldHeading = .Heading
        Dim prevWalk         As Grh: prevWalk = .Body.Walk(oldHeading)
        Dim prevWeaponWalk   As Grh: prevWeaponWalk = .Arma.WeaponWalk(oldHeading)
        Dim prevShieldWalk   As Grh: prevShieldWalk = .Escudo.ShieldWalk(oldHeading)
        Dim hadMovArmaEscudo As Boolean: hadMovArmaEscudo = .MovArmaEscudo
        Dim keepStartIdle    As Long
        Dim newGi            As Long
        ' ============================================
        ' Body
        TempInt = Reader.ReadInt16()
        If TempInt < LBound(BodyData()) Or TempInt > UBound(BodyData()) Then
            .Body = BodyData(0)
            .iBody = 0
        Else
            .Body = BodyData(TempInt)
            .iBody = TempInt
        End If
        ' Head
        headIndex = Reader.ReadInt16()
        If headIndex < LBound(HeadData()) Or headIndex > UBound(HeadData()) Then
            .Head = HeadData(0)
            .IHead = 0
        Else
            .Head = HeadData(headIndex)
            .IHead = headIndex
        End If
        .Muerto = (.iBody = CASPER_BODY_IDLE)
        ' Heading nuevo
        .Heading = Reader.ReadInt8()
        ' Arma / Escudo / Casco
        TempInt = Reader.ReadInt16()
        If TempInt <> 0 And TempInt <= UBound(WeaponAnimData) Then .Arma = WeaponAnimData(TempInt)
        TempInt = Reader.ReadInt16()
        If TempInt <> 0 And TempInt <= UBound(ShieldAnimData) Then .Escudo = ShieldAnimData(TempInt)
        TempInt = Reader.ReadInt16()
        If TempInt <> 0 And TempInt <= UBound(CascoAnimData) Then .Casco = CascoAnimData(TempInt)
        TempInt = Reader.ReadInt16()
        If TempInt <= 2 Or TempInt > UBound(BodyData()) Then
            .HasCart = False
        Else
            .Cart = BodyData(TempInt)
            .HasCart = True
        End If
        TempInt = Reader.ReadInt16()
        If TempInt <= 2 Or TempInt > UBound(BodyData()) Then
            .HasBackpack = False
            .tmpBackPack = 0
        Else
            .Backpack = BodyData(TempInt)
            .tmpBackPack = TempInt
            .HasBackpack = True
        End If
        .EsEnano = (.Body.HeadOffset.y = -26)
        ' FX
        Dim Fx As Integer: Fx = Reader.ReadInt16
        Call StartFx(.ActiveAnimation, Fx)
        .Meditating = (Fx <> 0)
        Reader.ReadInt16 ' Ignore loops
        ' Flags
        Dim flags As Byte
        flags = Reader.ReadInt8()
        .Idle = (flags And &O1)
        .Navegando = (flags And &O2)
        'exception for dwarven exoesqueleton
        If .iBody = DwarvenExoesqueletonBody Then
            .Head = HeadData(0)
            .HasBackpack = False
            .HasCart = False
            .Casco = CascoAnimData(NO_HELMET)
            .Escudo = ShieldAnimData(NO_SHIELD)
            .Arma = WeaponAnimData(NO_WEAPON)
        End If
        ' ==================== ANIMACIÓN / FASE ====================
        If .Idle Then
            ' --- IDLE ---
            If .Navegando = False Or UserNadandoTrajeCaucho = True Then
                If .Body.AnimateOnIdle = 0 Then
                    ' Idle sin anim: parar
                    .Body.Walk(.Heading).started = 0
                Else
                    ' Idle con anim: si cambia a IdleBody, preservá fase si venía animando
                    If .Body.IdleBody > 0 Then
                        newGi = BodyData(.Body.IdleBody).Walk(.Heading).GrhIndex
                        If prevWalk.started > 0 And wasMoving Then
                            keepStartIdle = SyncGrhPhase(prevWalk, newGi)
                        Else
                            keepStartIdle = FrameTime
                        End If
                        .Body = BodyData(.Body.IdleBody)
                        .Body.Walk(.Heading).started = keepStartIdle
                    ElseIf .Body.Walk(.Heading).started = 0 Then
                        If .Body.Walk(.Heading).started = 0 Then
                            .Body.Walk(.Heading).started = FrameTime
                        End If
                    End If
                End If
                ' Arma/Escudo en idle: respetá tu regla
                If Not .MovArmaEscudo Then
                    .Arma.WeaponWalk(.Heading).started = 0
                    .Escudo.ShieldWalk(.Heading).started = 0
                End If
            End If
        Else
            ' --- NO IDLE (camina / se mueve) ---
            Dim keepStart As Long
            Dim targetGi  As Long
            targetGi = .Body.Walk(.Heading).GrhIndex
            If wasMoving And prevWalk.started > 0 Then
                keepStart = SyncGrhPhase(prevWalk, targetGi)
            ElseIf .Body.Walk(.Heading).started > 0 Then
                keepStart = .Body.Walk(.Heading).started
            Else
                keepStart = FrameTime
            End If
            .Body.Walk(.Heading).started = keepStart
            targetGi = .Backpack.Walk(.Heading).GrhIndex
            If wasMoving And prevWalk.started > 0 Then
                keepStart = SyncGrhPhase(prevWalk, targetGi)
            ElseIf .Backpack.Walk(.Heading).started > 0 Then
                keepStart = .Body.Walk(.Heading).started
            Else
                keepStart = FrameTime
            End If
            .Backpack.Walk(.Heading).started = keepStart
            ' Arma/Escudo: mantener en fase con el cuerpo
            If .MovArmaEscudo Then
                Dim keepW As Long, keepS As Long
                If hadMovArmaEscudo And prevWeaponWalk.started > 0 Then
                    keepW = SyncGrhPhase(prevWeaponWalk, .Arma.WeaponWalk(.Heading).GrhIndex)
                Else
                    keepW = keepStart
                End If
                If hadMovArmaEscudo And prevShieldWalk.started > 0 Then
                    keepS = SyncGrhPhase(prevShieldWalk, .Escudo.ShieldWalk(.Heading).GrhIndex)
                Else
                    keepS = keepStart
                End If
                If .Arma.WeaponWalk(.Heading).started = 0 Then .Arma.WeaponWalk(.Heading).started = keepW
                If .Escudo.ShieldWalk(.Heading).started = 0 Then .Escudo.ShieldWalk(.Heading).started = keepS
            Else
                .Arma.WeaponWalk(.Heading).started = 0
                .Escudo.ShieldWalk(.Heading).started = 0
            End If
        End If
        ' ===========================================================
    End With
    Exit Sub
HandleCharacterChange_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCharacterChange", Erl)
End Sub

''
' Handles the ObjectCreate message.
Private Sub HandleObjectCreate()
    On Error GoTo HandleObjectCreate_Err
    Dim x             As Byte
    Dim y             As Byte
    Dim ObjIndex      As Integer
    Dim Amount        As Integer
    Dim color         As RGBA
    Dim Rango         As Byte
    Dim id            As Long
    Dim ElementalTags As Long
    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
    ObjIndex = Reader.ReadInt16()
    Amount = Reader.ReadInt16
    ElementalTags = Reader.ReadInt32
    MapData(x, y).ObjGrh.GrhIndex = ObjData(ObjIndex).GrhIndex
    MapData(x, y).OBJInfo.ObjIndex = ObjIndex
    MapData(x, y).OBJInfo.Amount = Amount
    MapData(x, y).OBJInfo.ElementalTags = ElementalTags
    Call InitGrh(MapData(x, y).ObjGrh, MapData(x, y).ObjGrh.GrhIndex)
    If ObjData(ObjIndex).CreaLuz <> "" Then
        Call Long_2_RGBA(color, val(ReadField(2, ObjData(ObjIndex).CreaLuz, Asc(":"))))
        Rango = val(ReadField(1, ObjData(ObjIndex).CreaLuz, Asc(":")))
        MapData(x, y).luz.color = color
        MapData(x, y).luz.Rango = Rango
        If Rango < 100 Then
            id = x & y
            LucesCuadradas.Light_Create x, y, color, Rango, id
        Else
            LucesRedondas.Create_Light_To_Map x, y, color, Rango - 99
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
    Dim x  As Byte
    Dim y  As Byte
    Dim Fx As Byte
    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
    Fx = Reader.ReadInt16()
    Call SetMapFx(x, y, Fx, 0)
    Exit Sub
HandleFxPiso_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleFxPiso", Erl)
End Sub

''
' Handles the ObjectDelete message.
Private Sub HandleObjectDelete()
    On Error GoTo HandleObjectDelete_Err
    Dim x  As Byte
    Dim y  As Byte
    Dim id As Long
    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
    If ObjData(MapData(x, y).OBJInfo.ObjIndex).CreaLuz <> "" Then
        id = LucesCuadradas.Light_Find(x & y)
        LucesCuadradas.Light_Remove id
        MapData(x, y).luz.color = COLOR_EMPTY
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
    Call Reader.ReadInt8   ' File
    Call Reader.ReadInt16  ' Loop
    Exit Sub
HandlePlayMIDI_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePlayMIDI", Erl)
End Sub

Private Sub HandlePlayWave()
    On Error GoTo HandlePlayWave_Err
    '=== Read packet data ===
    Dim wave           As Integer           ' ID of the WAV to play
    Dim srcX           As Byte              ' Source X coordinate for 3D positioning
    Dim srcY           As Byte              ' Source Y coordinate for 3D positioning
    Dim cancelLastWave As Byte    ' 0 = no cancel, 1 = stop previous, 2 = stop & do not play new
    Dim Localize       As Byte          ' 1 = prefix filename for localization, 0 = use default
    Dim filename       As String
    wave = Reader.ReadInt16()
    srcX = Reader.ReadInt8()
    srcY = Reader.ReadInt8()
    cancelLastWave = Reader.ReadInt8()
    Localize = Reader.ReadInt8()
    '=== Determine filename, applying localization if requested ===
    ' If Localize=1, prepend the user’s language code (e.g. "pt") so that
    ' pt_123.wav will be used instead of 123.wav when playing the clip.
    If Localize = 1 And language <> Spanish Then
        Dim langPrefix As String
        langPrefix = GetLanguagePrefix(language)
        filename = langPrefix & "_" & CStr(wave)
    Else
        filename = CStr(wave)
    End If
    '=== Handle special “fog of war” waves: IDs 400–404 only play if MapDat.niebla <> 0 ===
    Select Case wave
        Case 400 To 404
            If MapDat.niebla = 0 Then
                Exit Sub   ' Skip playing these if fog is disabled
            End If
    End Select
    '=== Cancel previous wave if requested ===
    If cancelLastWave <> 0 Then
        ao20audio.StopWav CStr(wave)
        If cancelLastWave = 2 Then
            Exit Sub   ' Don’t play the new wave if flag=2
        End If
    End If
    '=== Play the wave, either with spatial positioning or at default volume ===
    If srcX = 0 Or srcY = 0 Then
        ' No position provided: play at default volume & center pan
        ao20audio.PlayWav filename, False, 0, 0
    Else
        ' Only play if the source position is within audible area
        If EstaEnArea(srcX, srcY) Then
            Dim p As Position
            p.x = srcX
            p.y = srcY
            ' Compute volume & pan based on distance and orientation
            ao20audio.PlayWav filename, False, ao20audio.ComputeCharFxVolume(p), ao20audio.ComputeCharFxPan(p)
        End If
    End If
    Exit Sub
HandlePlayWave_Err:
    ' Log any runtime error for diagnostics
    RegistrarError Err.Number, Err.Description, "Protocol.HandlePlayWave", Erl
End Sub

''
' Handles the PlayWave message.
Private Sub HandlePlayWaveStep()
    On Error GoTo HandlePlayWaveStep_Err
    Dim charindex As Integer
    Dim Grh       As Long
    Dim Grh2      As Long
    Dim distance  As Byte
    Dim balance   As Integer
    Dim step      As Boolean
    charindex = Reader.ReadInt16()
    Grh = Reader.ReadInt32()
    Grh2 = Reader.ReadInt32()
    distance = Reader.ReadInt8()
    balance = Reader.ReadInt16()
    step = Reader.ReadBool()
    Call DoPasosInvi(Grh, Grh2, distance, balance, step)
    With charlist(charindex)
        ' Esta invisible, lo sacamos del mapa para que no tosquee
        If MapData(.Pos.x, .Pos.y).charindex = charindex Then
            MapData(.Pos.x, .Pos.y).charindex = 0
        End If
    End With
    Exit Sub
HandlePlayWaveStep_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePlayWaveStep", Erl)
End Sub

Private Sub HandlePosLLamadaDeClan()
    On Error GoTo HandlePosLLamadaDeClan_Err
    Dim map  As Integer
    Dim srcX As Byte
    Dim srcY As Byte
    map = Reader.ReadInt16()
    srcX = Reader.ReadInt8()
    srcY = Reader.ReadInt8()
    LLamadaDeclanX = srcX
    LLamadaDeclanY = srcY
    Call frmMapaGrande.ShowClanCall(map, srcX, srcY)
    Exit Sub
HandlePosLLamadaDeClan_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePosLLamadaDeClan", Erl)
End Sub

Private Sub HandleCharUpdateHP()
    On Error GoTo HandleCharUpdateHP_Err
    Dim charindex As Integer
    Dim MinHp     As Long
    Dim MaxHp     As Long
    Dim Shield    As Long
    charindex = Reader.ReadInt16()
    MinHp = Reader.ReadInt32()
    MaxHp = Reader.ReadInt32()
    Shield = Reader.ReadInt32()
    If Group.GroupSize > 0 Then
        Dim i As Integer
        For i = 0 To Group.GroupSize - 1
            If Group.GroupMembers(i).charindex = charindex Then
                Group.GroupMembers(i).MinHp = MinHp
                Group.GroupMembers(i).MaxHp = MaxHp
                Group.GroupMembers(i).Shield = Shield
            End If
        Next i
    End If
    charlist(charindex).UserMinHp = MinHp
    charlist(charindex).UserMaxHp = MaxHp
    charlist(charindex).Shield = Shield
    Exit Sub
HandleCharUpdateHP_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCharUpdateHP", Erl)
End Sub

Private Sub HandleCharUpdateMAN()
    On Error GoTo HandleCharUpdateHP_Err
    Dim charindex As Integer
    Dim minman    As Long
    Dim maxman    As Long
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
    Dim charindex As Integer
    Dim isRanged  As Byte
    charindex = Reader.ReadInt16()
    isRanged = Reader.ReadInt8()
    With charlist(charindex)
        If Not .Moving Then
            .MovArmaEscudo = True
            .Arma.WeaponWalk(.Heading).started = FrameTime
            .Arma.WeaponWalk(.Heading).Loops = 0
        End If
    End With
    Exit Sub
HandleArmaMov_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleArmaMov", Erl)
End Sub

Private Sub HandleCreateProjectile()
    On Error GoTo HandleCreateProjectile_Err
    Dim x, y, endX, endY, projectileType As Byte
    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
    endX = Reader.ReadInt8()
    endY = Reader.ReadInt8()
    projectileType = Reader.ReadInt8()
    Call InitializeProjectile(GProjectile, x, y, endX, endY, projectileType)
    Exit Sub
HandleCreateProjectile_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCreateProjectile", Erl)
End Sub

Private Sub HandleUpdateTrapState()
    Dim x, y As Byte
    Dim State As Byte
    State = Reader.ReadInt8()
    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
    If State > 0 Then
        Call InitGrh(MapData(x, y).Trap, 38370)
    Else
        MapData(x, y).Trap.GrhIndex = 0
    End If
End Sub

Private Sub HandleUpdateGroupInfo()
    On Error GoTo HandleUpdateGroupInfo_Err
    Group.GroupSize = Reader.ReadInt8
    Dim i As Integer
    If Group.GroupSize > 1 Then
        ReDim Group.GroupMembers(Group.GroupSize) As t_GroupEntry
    End If
    For i = 0 To Group.GroupSize - 1
        Group.GroupMembers(i).Name = Reader.ReadString8
        Group.GroupMembers(i).Name = mid$(Group.GroupMembers(i).Name, 1, min(Len(Group.GroupMembers(i).Name), 10))
        Group.GroupMembers(i).charindex = Reader.ReadInt16
        Group.GroupMembers(i).Head = HeadData(Reader.ReadInt16)
        Group.GroupMembers(i).GroupId = i + 1
        Group.GroupMembers(i).MinHp = Reader.ReadInt16
        Group.GroupMembers(i).MaxHp = Reader.ReadInt16
    Next i
    Call UpdateRenderArea
    Exit Sub
HandleUpdateGroupInfo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateGroupInfo", Erl)
End Sub

Private Sub HandleUpdateCharValue()
    On Error GoTo HandleUpdateGroupInfo_Err
    Dim charindex     As Integer
    Dim CharValueType As Integer
    Dim value         As Long
    charindex = Reader.ReadInt16
    CharValueType = Reader.ReadInt16
    value = Reader.ReadInt32
    Select Case CharValueType
        Case e_CharValue.eDontBlockTile
            charlist(charindex).DontBlockTile = value
    End Select
    Exit Sub
HandleUpdateGroupInfo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateGroupInfo", Erl)
End Sub

Private Sub HandleStunStart()
    On Error GoTo HandleStunStart_Err
    Dim duration As Integer
    duration = Reader.ReadInt16()
    StunEndTime = GetTickCount() + duration
    TotalStunTime = duration
    Exit Sub
HandleStunStart_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleStunStart", Erl)
End Sub

Private Sub HandleEscudoMov()
    On Error GoTo HandleEscudoMov_Err
    Dim charindex As Integer
    charindex = Reader.ReadInt16()
    With charlist(charindex)
        If Not .Moving Then
            .MovArmaEscudo = True
            .Escudo.ShieldWalk(.Heading).started = FrameTime
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
    On Error GoTo errhandler
    'Clear guild's list
    frmGuildList.GuildsList.Clear
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
            ClanesList(i).Alineacion = val(ReadField(2, guilds(i), Asc("-")))
            ClanesList(i).indice = i
        Next i
        For i = 0 To UBound(guilds())
            'If ClanesList(i).Alineacion = 0 Then
            Call frmGuildList.GuildsList.AddItem(ClanesList(i).nombre)
            'End If
        Next i
    End If
    COLOR_AZUL = RGB(0, 0, 0)
    Call Establecer_Borde(frmGuildList.GuildsList, frmGuildList, COLOR_AZUL, 0, 0)
    Call frmGuildList.Show(vbModeless, GetGameplayForm())
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildList", Erl)
End Sub

''
' Handles the AreaChanged message.
Private Sub HandleAreaChanged()
    On Error GoTo HandleAreaChanged_Err
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
    'Remove packet ID
    On Error GoTo HandlePauseToggle_Err
    pausa = Not pausa
    Exit Sub
HandlePauseToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePauseToggle", Erl)
End Sub

Private Sub HandleRainToggle()
    On Error GoTo HandleRainToggle_Err
    bRain = Reader.ReadBool
    If Not InMapBounds(UserPos.x, UserPos.y) Then Exit Sub
    If MapDat.LLUVIA = 0 Then Exit Sub
    If bRain Then
        Call ao20audio.PlayWeatherAudio(IIf(bTecho, SND_RAIN_IN_LOOP, SND_RAIN_OUT_LOOP))
        Call Graficos_Particulas.Engine_MeteoParticle_Set(Particula_Lluvia, False)
    Else
        Call ao20audio.PlayAmbientAudio(UserMap)
        Call ao20audio.PlayAmbientWav(IIf(bTecho, SND_RAIN_IN_END, SND_RAIN_OUT_END))
        Call Graficos_Particulas.Engine_MeteoParticle_Set(-1)
    End If
    Exit Sub
HandleRainToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleRainToggle", Erl)
End Sub

''
' Handles the CreateFX message.
Private Sub HandleCreateFX()
    On Error GoTo HandleCreateFX_Err
    Dim charindex As Integer
    Dim Fx        As Integer
    Dim Loops     As Integer
    Dim x         As Byte, y As Byte
    charindex = Reader.ReadInt16()
    Fx = Reader.ReadInt16()
    Loops = Reader.ReadInt16()
    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
    If x + y > 0 Then
        With charlist(charindex)
            If .Invisible And charindex <> UserCharIndex Then
                If MapData(.Pos.x, .Pos.y).charindex = charindex Then MapData(.Pos.x, .Pos.y).charindex = 0
                .Pos.x = x
                .Pos.y = y
                MapData(x, y).charindex = charindex
            End If
        End With
    End If
    Call SetCharacterFx(charindex, Fx, Loops)
    Exit Sub
HandleCreateFX_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCreateFX", Erl)
End Sub

''
' Handles the CharAtaca message.
Private Sub HandleCharAtaca()
    On Error GoTo HandleCharAtaca_Err
    Dim NpcIndex    As Integer
    Dim VictimIndex As Integer
    Dim danio       As Long
    Dim AnimAttack  As Integer
    NpcIndex = Reader.ReadInt16()
    VictimIndex = Reader.ReadInt16()
    danio = Reader.ReadInt32()
    AnimAttack = Reader.ReadInt16()
    Dim oldWalk    As Grh
    Dim keepStart  As Long
    With charlist(NpcIndex)
        If AnimAttack > 0 Then
            oldWalk = .Body.Walk(.Heading)
            .AnimatingBody = AnimAttack
            .Idle = False
            .Body = BodyData(AnimAttack)
            .Body.Walk(.Heading).Loops = 0
            If oldWalk.started > 0 And .Moving Then
                keepStart = SyncGrhPhase(oldWalk, .Body.Walk(.Heading).GrhIndex)
            Else
                keepStart = FrameTime
            End If
            .Body.Walk(.Heading).started = keepStart
            If Not .MovArmaEscudo Then
                If .Arma.WeaponWalk(.Heading).GrhIndex <> 0 Then
                    .Arma.WeaponWalk(.Heading).Loops = 0
                    .Arma.WeaponWalk(.Heading).started = .Body.Walk(.Heading).started
                End If
                If .Escudo.ShieldWalk(.Heading).GrhIndex <> 0 Then
                    .Escudo.ShieldWalk(.Heading).Loops = 0
                    .Escudo.ShieldWalk(.Heading).started = .Body.Walk(.Heading).started
                End If
            End If
        Else
            If Not .Moving Then
                .MovArmaEscudo = True
                .Arma.WeaponWalk(.Heading).started = FrameTime
                .Arma.WeaponWalk(.Heading).Loops = 0
            End If
        End If
    End With
    'renderizo sangre si está sin montar ni navegar
    If danio > 0 And charlist(VictimIndex).Navegando = 0 Then Call SetCharacterFx(VictimIndex, 14, 0)
    If charlist(UserCharIndex).Muerto = False And EstaPCarea(NpcIndex) Then
        Call ao20audio.PlayWav(CStr(IIf(danio = -1, 2, 10)), False, ao20audio.ComputeCharFxVolume(charlist(NpcIndex).Pos), ao20audio.ComputeCharFxPan(charlist(NpcIndex).Pos))
    End If
    Exit Sub
HandleCharAtaca_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCharAtaca", Erl)
End Sub

''
' Handles the CharAtaca message.
Private Sub HandleNotificarClienteSeguido()
    On Error GoTo NotificarClienteSeguido_Err
    Seguido = Reader.ReadInt8
    LastSentPosX = -1
    LastSentPosY = -1
    Exit Sub
NotificarClienteSeguido_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.NotificarClienteSeguido", Erl)
End Sub


Private Sub HandleGetInventarioHechizos()
    On Error GoTo GetInventarioHechizos_Err
    Dim inventario_o_hechizos As Byte
    Dim hechiSel              As Byte
    Dim scrollSel             As Byte
    inventario_o_hechizos = Reader.ReadInt8()
    hechiSel = Reader.ReadInt8()
    scrollSel = Reader.ReadInt8()
    'Clickeó en inventario
    If inventario_o_hechizos = 1 Then
        Call frmMain.inventoryClick
        'Clickeó en hechizos
    ElseIf inventario_o_hechizos = 2 Then
        Call frmMain.hechizosClick
        hlst.Scroll = scrollSel
        hlst.ListIndex = hechiSel
    End If
    Exit Sub
GetInventarioHechizos_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.GetInventarioHechizos", Erl)
End Sub

Private Sub HandleNotificarClienteCasteo()
    On Error GoTo NotificarClienteCasteo_Err
    Dim value As Byte
    value = Reader.ReadInt8()
    'Clickeó en inventario
    If value = 1 Then
        frmMain.shapexy.BackColor = RGB(0, 170, 0)
        'Clickeó en hechizos
    Else
        frmMain.shapexy.BackColor = RGB(170, 0, 0)
    End If
    Exit Sub
NotificarClienteCasteo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.NotificarClienteCasteo", Erl)
End Sub


Private Sub HandleUpdateUserStats()
    On Error GoTo HandleUpdateUserStats_Err
    UserStats.MaxHp = Reader.ReadInt16()
    UserStats.MinHp = Reader.ReadInt16()
    UserStats.HpShield = Reader.ReadInt32()
    UserStats.maxman = Reader.ReadInt16()
    UserStats.minman = Reader.ReadInt16()
    UserStats.MaxSTA = Reader.ReadInt16()
    UserStats.MinSTA = Reader.ReadInt16()
    UserStats.GLD = Reader.ReadInt32()
    UserStats.OroPorNivel = Reader.ReadInt32()
    UserStats.Lvl = Reader.ReadInt8()
    #If DEBUGGING = 0 Then
        Select Case UserStats.Lvl:
            Case 10
                Call svb_unlock_achivement("Adventurer")
            Case 20
                Call svb_unlock_achivement("Seasoned adventurer")
            Case 30
                Call svb_unlock_achivement("Big shot!")
            Case 40
                Call svb_unlock_achivement("Oh! You mean business!")
            Case Else
                'Nothing
        End Select
    #End If
    UserStats.PasarNivel = Reader.ReadInt32()
    UserStats.exp = Reader.ReadInt32()
    UserStats.Clase = Reader.ReadInt8()
    If UserStats.MinHp = 0 Then
        UserStats.estado = 1
        charlist(UserCharIndex).Invisible = False
        DrogaCounter = 0
    Else
        UserStats.estado = 0
    End If
    Call frmMain.UpdateStatsLayout
    Exit Sub
HandleUpdateUserStats_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateUserStats", Erl)
End Sub

''
' Handles the WorkRequestTarget message.
Private Sub HandleWorkRequestTarget()
    On Error GoTo HandleWorkRequestTarget_Err
    Dim UsingSkillREcibido As Byte
    UsingSkillREcibido = Reader.ReadInt8()
    casteaArea = Reader.ReadBool()
    RadioHechizoArea = Reader.ReadInt8()
    Dim Frm As Form
    Set Frm = GetGameplayForm
    If UsingSkillREcibido = 0 Then
        Frm.MousePointer = 0
        Call FormParser.Parse_Form(frmMain, E_NORMAL)
        UsingSkill = UsingSkillREcibido
        Exit Sub
    End If
    If PescandoEspecial = True Then Exit Sub
    If UsingSkillREcibido = UsingSkill Then Exit Sub
    UsingSkill = UsingSkillREcibido
    Frm.MousePointer = 2
    Select Case UsingSkill
        Case magia, eSkill.TargetableItem
            Call FormParser.Parse_Form(Frm, E_CAST)
            Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_TRABAJO_MAGIA"), 100, 100, 120, 0, 0)
        Case Robar
            Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_TRABAJO_ROBAR"), 100, 100, 120, 0, 0)
            Call FormParser.Parse_Form(Frm, E_SHOOT)
        Case FundirMetal
            Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_TRABAJO_FUNDIRMETAL"), 100, 100, 120, 0, 0)
            Call FormParser.Parse_Form(Frm, E_SHOOT)
        Case Proyectiles
            Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_TRABAJO_PROYECTILES"), 100, 100, 120, 0, 0)
            Call FormParser.Parse_Form(Frm, E_ARROW)
        Case eSkill.Talar, eSkill.Alquimia, eSkill.Carpinteria, eSkill.Herreria, eSkill.Mineria, eSkill.Pescar
            Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_CLICK_DONDE_DESEAS_TRABAJAR"), 100, 100, 120, 0, 0)
            Call FormParser.Parse_Form(Frm, E_SHOOT)
        Case Grupo
            Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_TRABAJO_MAGIA"), 100, 100, 120, 0, 0)
            Call FormParser.Parse_Form(Frm, E_SHOOT)
        Case MarcaDeClan
            Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_SELECCIONA_PERSONAJE_A_MARCAR"), 100, 100, 120, 0, 0)
            Call FormParser.Parse_Form(Frm, E_SHOOT)
        Case MarcaDeGM
            Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_SELECCIONA_PERSONAJE_A_MARCAR"), 100, 100, 120, 0, 0)
            Call FormParser.Parse_Form(Frm, E_SHOOT)
        Case Domar
            Call FormParser.Parse_Form(Frm, E_SHOOT)
    End Select
    Exit Sub
HandleWorkRequestTarget_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleWorkRequestTarget", Erl)
End Sub

''
' Handles the ChangeInventorySlot message.
Private Sub HandleChangeInventorySlot()
    On Error GoTo errhandler
    Dim Slot          As Byte
    Dim ObjIndex      As Integer
    Dim Name          As String
    Dim Amount        As Integer
    Dim Equipped      As Boolean
    Dim GrhIndex      As Long
    Dim ObjType       As Byte
    Dim MaxHit        As Integer
    Dim MinHit        As Integer
    Dim MaxDef        As Integer
    Dim MinDef        As Integer
    Dim value         As Single
    Dim podrausarlo   As Byte
    Dim IsBindable    As Boolean
    Dim ElementalTags As Long
    Slot = Reader.ReadInt8()
    ObjIndex = Reader.ReadInt16()
    Amount = Reader.ReadInt16()
    Equipped = Reader.ReadBool()
    value = Reader.ReadReal32()
    podrausarlo = Reader.ReadInt8()
    ElementalTags = Reader.ReadInt32()
    IsBindable = Reader.ReadBool()
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
    Call ModGameplayUI.SetInvItem(Slot, ObjIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, MinDef, value, Name, podrausarlo, ElementalTags, IsBindable)
    If frmComerciar.visible Then
        Call frmComerciar.InvComUsu.SetItem(Slot, ObjIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, MinDef, value, Name, ElementalTags, podrausarlo)
    ElseIf frmBancoObj.visible Then
        Call frmBancoObj.InvBankUsu.SetItem(Slot, ObjIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, MinDef, value, Name, ElementalTags, podrausarlo)
    ElseIf frmBancoCuenta.visible Then
        Call frmBancoCuenta.InvBankUsuCuenta.SetItem(Slot, ObjIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, MinDef, value, Name, ElementalTags, podrausarlo)
    ElseIf frmCrafteo.visible Then
        Call frmCrafteo.InvCraftUser.SetItem(Slot, ObjIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, MinDef, value, Name, ElementalTags, podrausarlo)
    End If
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChangeInventorySlot", Erl)
End Sub

' Handles the InventoryUnlockSlots message.
Private Sub HandleInventoryUnlockSlots()
    On Error GoTo HandleInventoryUnlockSlots_Err
    UserInvUnlocked = Reader.ReadInt8
    Call frmMain.UnlockInvslot(UserInvUnlocked)
    Exit Sub
HandleInventoryUnlockSlots_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleInventoryUnlockSlots", Erl)
End Sub

''
' Handles the ChangeBankSlot message.
Private Sub HandleChangeBankSlot()
    On Error GoTo errhandler
    Dim Slot     As Byte
    Dim BankSlot As Slot
    With BankSlot
        Slot = Reader.ReadInt8()
        .ObjIndex = Reader.ReadInt16()
        .ElementalTags = Reader.ReadInt32()
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
        Call frmBancoObj.InvBoveda.SetItem(Slot, .ObjIndex, .Amount, .Equipped, .GrhIndex, .ObjType, .MaxHit, .MinHit, .Def, .Valor, .Name, .ElementalTags, .PuedeUsar)
    End With
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChangeBankSlot", Erl)
End Sub

''
' Handles the ChangeSpellSlot message
Private Sub HandleChangeSpellSlot()
    On Error GoTo errhandler
    Dim Slot       As Byte
    Dim Index      As Integer
    Dim Cooldown   As Integer
    Dim IsBindable As Boolean
    Slot = Reader.ReadInt8()
    UserHechizos(Slot) = Reader.ReadInt16()
    Index = Reader.ReadInt16()
    IsBindable = Reader.ReadBool()
    If Index >= 0 Then
        HechizoData(Index).IsBindable = IsBindable
        If Slot <= hlst.ListCount Then
            hlst.List(Slot - 1) = HechizoData(Index).nombre
        Else
            Call hlst.AddItem(HechizoData(Index).nombre)
            hlst.Scroll = LastScroll
        End If
    Else
        If Slot <= hlst.ListCount Then
            hlst.List(Slot - 1) = JsonLanguage.Item("EMPTY_LABEL")
        Else
            Call hlst.AddItem(JsonLanguage.Item("EMPTY_LABEL"))
            hlst.Scroll = LastScroll
        End If
    End If
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChangeSpellSlot", Erl)
End Sub

''
' Handles the Attributes message.
Private Sub HandleAtributes()
    On Error GoTo HandleAtributes_Err
    Dim i As Long
    For i = 1 To NUMATRIBUTES
        UserAtributos(i) = Reader.ReadInt8()
    Next i
    'Show them in character creation
    If LlegaronStats Then
        frmStatistics.Iniciar_Labels
        frmStatistics.Picture = LoadInterface("ventanaestadisticas_personaje.bmp")
        frmStatistics.Show , GetGameplayForm()
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
    On Error GoTo errhandler
    Dim count As Integer
    Dim i     As Long
    Dim tmp   As String
    count = Reader.ReadInt16()
    Call frmHerrero.lstArmas.Clear
    For i = 1 To count
        ArmasHerrero(i).Index = Reader.ReadInt16()
    Next i
    For i = i To UBound(ArmasHerrero())
        ArmasHerrero(i).Index = 0
    Next i
    i = 0
    Exit Sub
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBlacksmithWeapons", Erl)
End Sub

''
' Handles the BlacksmithArmors message.
Private Sub HandleBlacksmithArmors()
    On Error GoTo errhandler
    Dim count As Integer
    Dim i     As Long
    Dim tmp   As String
    count = Reader.ReadInt16()
    'Call frmHerrero.lstArmaduras.Clear
    For i = 1 To count
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
            A = A + 1
        End If
        ' Escudos (16), Objetos Magicos (21) y Anillos (35) van en la misma lista
        If tmpObj.ObjType = 16 Or tmpObj.ObjType = 35 Or tmpObj.ObjType = 21 Or tmpObj.ObjType = 100 Or tmpObj.ObjType = 30 Then
            EscudosHerrero(e).Index = DefensasHerrero(i).Index
            e = e + 1
        End If
        If tmpObj.ObjType = 17 Then
            CascosHerrero(c).Index = DefensasHerrero(i).Index
            c = c + 1
        End If
    Next i
    Exit Sub
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBlacksmithArmors", Erl)
End Sub

Private Sub HandleBlacksmithExtraObjects()
    On Error GoTo errhandler
    Dim count As Integer
    Dim i     As Long
    Dim tmp   As String
    count = Reader.ReadInt16()
    Call frmHerrero.lstArmas.Clear
    For i = 1 To count
        RunasElementalesHerrero(i).Index = Reader.ReadInt16()
    Next i
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBlacksmithExtraObjects", Erl)
End Sub

''
' Handles the CarpenterObjects message.
Private Sub HandleCarpenterObjects()
    On Error GoTo errhandler
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
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCarpenterObjects", Erl)
End Sub

Private Sub HandleSastreObjects()
    On Error GoTo errhandler
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
        ObjSastre(i).PielLoboNegro = ObjData(ObjSastre(i).Index).PielLoboNegro
    Next i
    Dim R As Byte
    Dim G As Byte
    i = 0
    R = 1
    G = 1
    For i = i To UBound(ObjSastre())
        If ObjData(ObjSastre(i).Index).ObjType = 3 Or ObjData(ObjSastre(i).Index).ObjType = 100 Then
            SastreRopas(R).Index = ObjSastre(i).Index
            SastreRopas(R).PielLobo = ObjSastre(i).PielLobo
            SastreRopas(R).PielOsoPardo = ObjSastre(i).PielOsoPardo
            SastreRopas(R).PielOsoPolar = ObjSastre(i).PielOsoPolar
            SastreRopas(G).PielLoboNegro = ObjSastre(i).PielLoboNegro
            R = R + 1
        End If
        If ObjData(ObjSastre(i).Index).ObjType = 17 Then
            SastreGorros(G).Index = ObjSastre(i).Index
            SastreGorros(G).PielLobo = ObjSastre(i).PielLobo
            SastreGorros(G).PielOsoPardo = ObjSastre(i).PielOsoPardo
            SastreGorros(G).PielOsoPolar = ObjSastre(i).PielOsoPolar
            SastreGorros(G).PielLoboNegro = ObjSastre(i).PielLoboNegro
            G = G + 1
        End If
    Next i
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSastreObjects", Erl)
End Sub

Private Sub HandleAlquimiaObjects()
    On Error GoTo errhandler
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
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleAlquimiaObjects", Erl)
End Sub

''
' Handles the RestOK message.
Private Sub HandleRestOK()
    'Remove packet ID
    On Error GoTo HandleRestOK_Err
    UserDescansar = Not UserDescansar
    Exit Sub
HandleRestOK_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleRestOK", Erl)
End Sub

Private Sub HandleErrorMessage()
    On Error GoTo errhandler
    GetRemoteError = True
    Dim str As String
    str = Reader.ReadString8()
    Call MsgBox(str)
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleErrorMessage", Erl)
End Sub

''
' Handles the Blind message.
Private Sub HandleBlind()
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
    'Remove packet ID
    On Error GoTo HandleDumb_Err
    UserEstupido = True
    Exit Sub
HandleDumb_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleDumb", Erl)
End Sub

''
' Handles the ShowSignal message.
Private Sub HandleShowSignal()
    On Error GoTo errhandler
    Dim tmp As String
    Dim Grh As Integer
    tmp = ObjData(Reader.ReadInt16()).Texto
    Grh = Reader.ReadInt16()
    Call InitCartel(tmp, Grh)
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowSignal", Erl)
End Sub

''
' Handles the ChangeNPCInventorySlot message.
Private Sub HandleChangeNPCInventorySlot()
    On Error GoTo errhandler
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
        .ElementalTags = Reader.ReadInt32()
        .PuedeUsar = Reader.ReadInt8()
        Call frmComerciar.InvComNpc.SetItem(Slot, .ObjIndex, .Amount, 0, .GrhIndex, .ObjType, .MaxHit, .MinHit, .Def, .Valor, .Name, .ElementalTags, .PuedeUsar)
    End With
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChangeNPCInventorySlot", Erl)
End Sub

''
' Handles the UpdateHungerAndThirst message.
Private Sub HandleUpdateHungerAndThirst()
    On Error GoTo HandleUpdateHungerAndThirst_Err
    UserStats.MaxAGU = Reader.ReadInt8()
    UserStats.MinAGU = Reader.ReadInt8()
    UserStats.MaxHAM = Reader.ReadInt8()
    UserStats.MinHAM = Reader.ReadInt8()
    Call frmMain.UpdateFoodState
    Exit Sub
HandleUpdateHungerAndThirst_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateHungerAndThirst", Erl)
End Sub



Private Sub HandleHora()
    On Error GoTo HandleHora_Err

    Dim elapsedFromServer As Long
    Dim dayLen As Long
    elapsedFromServer = Reader.ReadInt32()
    dayLen = Reader.ReadInt32()
    

    WorldTime_HandleHora elapsedFromServer, dayLen

    If Not Connected Then
        RevisarHoraMundo True
    End If
    Exit Sub

HandleHora_Err:
    RegistrarError Err.Number, Err.Description, "Protocol.HandleHora", Erl
End Sub

 
Private Sub HandleLight()
    On Error GoTo HandleLight_Err
    Dim color As String
    color = Reader.ReadString8()
    'Call SetGlobalLight(Map_light_base)
    Exit Sub
HandleLight_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleLight", Erl)
End Sub
 
Private Sub HandleFYA()
    On Error GoTo HandleFYA_Err
    UserAtributos(eAtributos.Fuerza) = Reader.ReadInt8()
    UserAtributos(eAtributos.Agilidad) = Reader.ReadInt8()
    UserStats.str = UserAtributos(eAtributos.Fuerza)
    UserStats.Agi = UserAtributos(eAtributos.Agilidad)
    DrogaCounter = Reader.ReadInt16()
    If UserStats.str >= 35 Then
        UserStats.StrState = eHighBuff
    ElseIf UserStats.str >= 25 Then
        UserStats.StrState = eMinBuff
    Else
        UserStats.StrState = eNormal
    End If
    If UserStats.Agi >= 35 Then
        UserStats.AgiState = eHighBuff
    ElseIf UserStats.Agi >= 25 Then
        UserStats.AgiState = eMinBuff
    Else
        UserStats.AgiState = eNormal
    End If
    If DrogaCounter > 0 Then
        frmMain.Contadores.enabled = True
    End If
    Call frmMain.UpdateBuff
    Exit Sub
HandleFYA_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleFYA", Erl)
End Sub

Private Sub HandleUpdateNPCSimbolo()
    On Error GoTo HandleUpdateNPCSimbolo_Err
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
    On Error GoTo HandleCerrarleCliente_Err
    EngineRun = False
    Call CloseClient
    Exit Sub
HandleCerrarleCliente_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCerrarleCliente", Erl)
End Sub

Private Sub HandleContadores()
    On Error GoTo HandleContadores_Err
    InviCounter = Reader.ReadInt16()
    DrogaCounter = Reader.ReadInt16()
    frmMain.Contadores.enabled = True
    Exit Sub
HandleContadores_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleContadores", Erl)
End Sub

Private Sub HandleShowPapiro()
    On Error GoTo HandleShowPapiro_Err
    frmMensajePapiro.Show , GetGameplayForm()
    Exit Sub
HandleShowPapiro_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowPapiro", Erl)
End Sub

Private Sub HandleUpdateCooldownType()
    On Error GoTo HandleUpdateCooldownType_Err
    Dim CDType As Byte
    CDType = Reader.ReadInt8()
    CdTimes(CDType) = GetTickCount()
    Exit Sub
HandleUpdateCooldownType_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateCooldownType", Erl)
End Sub

''
' Handles the MiniStats message.
Private Sub HandleFlashScreen()
    On Error GoTo HandleEfectToScreen_Err
    Dim color As Long, duracion As Long, ignorar As Boolean
    color = Reader.ReadInt32()
    duracion = Reader.ReadInt32()
    ignorar = Reader.ReadBool()
    Dim R, G, B As Byte
    B = (color And 16711680) / 65536
    G = (color And 65280) / 256
    R = color And 255
    color = D3DColorARGB(255, R, G, B)
    If Not MapDat.niebla = 1 And Not ignorar Then
        'frmdebug.add_text_tracebox "trueno cancelado"
        Exit Sub
    End If
    Call EfectoEnPantalla(color, duracion)
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
            .Genero = JsonLanguage.Item("MENSAJE_576")
        Else
            .Genero = JsonLanguage.Item("MENSAJE_577")
        End If
        .Raza = Reader.ReadInt8()
        .Raza = ListaRazas(.Raza)
    End With
    If LlegaronAtrib Then
    
        #If DXUI Then
            frmStatistics.visible = False
            Call Unload(frmStatistics)
            
        
        #Else
            frmStatistics.Iniciar_Labels
            frmStatistics.Picture = LoadInterface("ventanaestadisticas_personaje.bmp")
            frmStatistics.Show , GetGameplayForm()
        #End If
    

        
        
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
    SkillPoints = Reader.ReadInt16()
    #If DEBUGGING = 0 Then
        Call svb_unlock_achivement("Newbie's fate")
    #End If
    Exit Sub
HandleLevelUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleLevelUp", Erl)
End Sub

''
' Handles the AddForumMessage message.
Private Sub HandleAddForumMessage()
    On Error GoTo errhandler
    Dim title   As String
    Dim message As String
    title = Reader.ReadString8()
    message = Reader.ReadString8()
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleAddForumMessage", Erl)
End Sub

''
' Handles the ShowForumForm message.
Private Sub HandleShowForumForm()
    On Error GoTo HandleShowForumForm_Err
    Exit Sub
HandleShowForumForm_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowForumForm", Erl)
End Sub

''
' Handles the SetInvisible message.
Private Sub HandleSetInvisible()
    On Error GoTo HandleSetInvisible_Err
    Dim charindex As Integer
    Dim x         As Byte, y As Byte
    charindex = Reader.ReadInt16()
    charlist(charindex).Invisible = Reader.ReadBool()
    charlist(charindex).TimerI = 0
    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
    If x + y > 0 Then
        With charlist(charindex)
            If charindex <> UserCharIndex Then
                If .Invisible Then
                    If Not IsCharVisible(charindex) And General_Distance_Get(x, y, UserPos.x, UserPos.y) > DISTANCIA_ENVIO_DATOS Then
                        If .clan_index > 0 Then
                            If .clan_index = charlist(UserCharIndex).clan_index And charindex <> UserCharIndex And .Muerto = 0 Then
                                If .clan_nivel >= 6 Then
                                    Exit Sub
                                End If
                            End If
                        End If
                        If .Meditating Then Exit Sub
                        If MapData(.Pos.x, .Pos.y).charindex = charindex Then MapData(.Pos.x, .Pos.y).charindex = 0
                        .MoveOffsetX = 0
                        .MoveOffsetY = 0
                    End If
                Else
                    If MapData(.Pos.x, .Pos.y).charindex = charindex Then MapData(.Pos.x, .Pos.y).charindex = 0
                    .Pos.x = x
                    .Pos.y = y
                    MapData(x, y).charindex = charindex
                    If Abs(.MoveOffsetX) > 32 Or Abs(.MoveOffsetY) > 32 Or (.MoveOffsetX <> 0 And .MoveOffsetY <> 0) Then
                        .MoveOffsetX = 0
                        .MoveOffsetY = 0
                    End If
                End If
            End If
        End With
    End If
    Exit Sub
HandleSetInvisible_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSetInvisible", Erl)
End Sub

Private Sub HandleMeditateToggle()
    On Error GoTo HandleMeditateToggle_Err
    Dim charindex As Integer, Fx As Integer
    Dim x         As Byte, y As Byte
    charindex = Reader.ReadInt16
    Fx = Reader.ReadInt16
    x = Reader.ReadInt8
    y = Reader.ReadInt8
    charlist(charindex).Meditating = Fx <> 0
    If x + y > 0 Then
        With charlist(charindex)
            If .Invisible And charindex <> UserCharIndex Then
                If MapData(.Pos.x, .Pos.y).charindex = charindex Then MapData(.Pos.x, .Pos.y).charindex = 0
                .Pos.x = x
                .Pos.y = y
                MapData(x, y).charindex = charindex
            End If
        End With
    End If
    If charindex = UserCharIndex Then
        UserMeditar = (Fx <> 0)
        If ChatCombate = 1 Then 'Si la pestaña "INFO" esta activada muestra mensajes de meditacion
            If UserMeditar Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg(JsonLanguage.Item("MENSAJE_COMIENZAS_A_MEDITAR"), .red, .green, .blue, .bold, .italic)
                End With
            Else
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg(JsonLanguage.Item("MENSAJE_HAS_DEJADO_DE_MEDITAR"), .red, .green, .blue, .bold, .italic)
                End With
            End If
        End If
    End If
    With charlist(charindex)
        If Fx <> 0 Then
            Call StartFx(.ActiveAnimation, Fx, -1)
            ' Play sound only in PC area
            If EstaPCarea(charindex) Then
                Call ao20audio.PlayWav(SND_MEDITATE, True, ao20audio.ComputeCharFxVolume(.Pos), ao20audio.ComputeCharFxPan(.Pos), "meditate" & CStr(charindex))
            End If
        Else
            Call ao20audio.StopWav(SND_MEDITATE, "meditate" & CStr(charindex))
            Call ChangeToClip(.ActiveAnimation, 3)
        End If
    End With
    Exit Sub
HandleMeditateToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleMeditateToggle", Erl)
End Sub

''
' Handles the BlindNoMore message.
Private Sub HandleBlindNoMore()
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
    Dim i As Long
    For i = 1 To NUMSKILLS
        UserSkills(i) = Reader.ReadInt8()
        'frmEstadisticas.skills(i).Caption = SkillsNames(i)
    Next i
    If LlegaronSkills Then
        Alocados = SkillPoints
        frmEstadisticas.Puntos.Caption = SkillPoints
        frmEstadisticas.Iniciar_Labels
        frmEstadisticas.Picture = LoadInterface("ventanaskills.bmp")
        frmEstadisticas.Show , GetGameplayForm()
        LlegaronSkills = False
    End If
    Exit Sub
HandleSendSkills_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSendSkills", Erl)
End Sub

''
' Handles the TrainerCreatureList message.
Private Sub HandleTrainerCreatureList()
    On Error GoTo errhandler
    Dim creatures() As String
    Dim i           As Long
    creatures = Split(Reader.ReadString8(), SEPARATOR)
    For i = 0 To UBound(creatures())
        Call frmEntrenador.lstCriaturas.AddItem(creatures(i))
    Next i
    frmEntrenador.Show , GetGameplayForm()
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleTrainerCreatureList", Erl)
End Sub

''
' Handles the GuildNews message.
Private Sub HandleGuildNews()
    On Error GoTo errhandler
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
        'frmdebug.add_text_tracebox guildList(i)
    Next i
    ClanNivel = Reader.ReadInt8()
    expacu = Reader.ReadInt16()
    ExpNe = Reader.ReadInt16()
    With frmGuildNews
        .lblMiembros.Caption = cantidad
        .expcount.Caption = expacu & "/" & ExpNe
        .EXPBAR.Width = (((expacu + 1 / 100) / (ExpNe + 1 / 100)) * 2370)
        .lblNivel = ClanNivel
        If ExpNe > 0 Then
            .porciento.Caption = Round(CDbl(expacu) * CDbl(100) / CDbl(ExpNe), 0) & "%"
        Else
            .porciento.Caption = "¡Nivel Máximo!"
            .expcount.Caption = "¡Nivel Máximo!"
        End If
        '.expne = "Experiencia necesaria: " & expne
        Select Case ClanNivel
            Case 1
                .beneficios = JsonLanguage.Item("MENSAJE_BENEFICIOS_MAX_MIEMBROS")
            Case 2
                .beneficios = JsonLanguage.Item("MENSAJE_PEDIR_AYUDA") & vbCrLf & JsonLanguage.Item("MENSAJE_MAX_MIEMBROS_8")
            Case 3
                .beneficios = JsonLanguage.Item("MENSAJE_PEDIR_AYUDA") & vbCrLf & JsonLanguage.Item("MENSAJE_SEGURO_CLAN") & vbCrLf & JsonLanguage.Item("MENSAJE_MAX_MIEMBROS_11")
            Case 4
                .beneficios = JsonLanguage.Item("MENSAJE_PEDIR_AYUDA") & vbCrLf & JsonLanguage.Item("MENSAJE_SEGURO_CLAN") & vbCrLf & JsonLanguage.Item("MENSAJE_MAX_MIEMBROS_14")
            Case 5
                .beneficios = JsonLanguage.Item("MENSAJE_PEDIR_AYUDA") & vbCrLf & JsonLanguage.Item("MENSAJE_SEGURO_CLAN") & vbCrLf & JsonLanguage.Item("MENSAJE_VER_VIDA_MANA") _
                        & vbCrLf & JsonLanguage.Item("MENSAJE_MAX_MIEMBROS_17")
            Case 6
                .beneficios = JsonLanguage.Item("MENSAJE_PEDIR_AYUDA") & vbCrLf & JsonLanguage.Item("MENSAJE_SEGURO_CLAN") & vbCrLf & JsonLanguage.Item("MENSAJE_VER_VIDA_MANA") _
                        & vbCrLf & JsonLanguage.Item("MENSAJE_VERSE_INVISIBLE") & vbCrLf & JsonLanguage.Item("MENSAJE_MAX_MIEMBROS_20")
        End Select
    End With
    frmGuildNews.Show vbModeless, GetGameplayForm()
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildNews", Erl)
End Sub

''
' Handles the OfferDetails message.
Private Sub HandleOfferDetails()
    On Error GoTo errhandler
    Call frmUserRequest.recievePeticion(Reader.ReadString8())
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleOfferDetails", Erl)
End Sub

''
' Handles the AlianceProposalsList message.
Private Sub HandleAlianceProposalsList()
    On Error GoTo errhandler
    Dim guildList() As String
    Dim i           As Long
    guildList = Split(Reader.ReadString8(), SEPARATOR)
    For i = 0 To UBound(guildList())
        Call frmPeaceProp.lista.AddItem(guildList(i))
    Next i
    frmPeaceProp.ProposalType = TIPO_PROPUESTA.ALIANZA
    Call frmPeaceProp.Show(vbModeless, GetGameplayForm())
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleAlianceProposalsList", Erl)
End Sub

''
' Handles the PeaceProposalsList message.
Private Sub HandlePeaceProposalsList()
    On Error GoTo errhandler
    Dim guildList() As String
    Dim i           As Long
    guildList = Split(Reader.ReadString8(), SEPARATOR)
    For i = 0 To UBound(guildList())
        Call frmPeaceProp.lista.AddItem(guildList(i))
    Next i
    frmPeaceProp.ProposalType = TIPO_PROPUESTA.PAZ
    Call frmPeaceProp.Show(vbModeless, GetGameplayForm())
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePeaceProposalsList", Erl)
End Sub

''
' Handles the CharacterInfo message.
Private Sub HandleCharacterInfo()
    On Error GoTo errhandler
    With frmCharInfo
        If .frmType = CharInfoFrmType.frmMembers Then
            .Rechazar.visible = False
            .Aceptar.visible = False
            .Echar.visible = True
            .desc.visible = False
        Else
            .Rechazar.visible = True
            .Aceptar.visible = True
            .Echar.visible = False
            .desc.visible = True
        End If
        If Reader.ReadInt8() = 1 Then
            .Genero.Caption = JsonLanguage.Item("MENSAJE_GENERO_HOMBRE")
        Else
            .Genero.Caption = JsonLanguage.Item("MENSAJE_GENERO_MUJER")
        End If
        .nombre.Caption = JsonLanguage.Item("MENSAJE_NOMBRE") & ": " & Reader.ReadString8()
        .Raza.Caption = JsonLanguage.Item("MENSAJE_RAZA") & ": " & ListaRazas(Reader.ReadInt8())
        .Clase.Caption = JsonLanguage.Item("MENSAJE_CLASE") & ": " & ListaClases(Reader.ReadInt8())
        .Nivel.Caption = JsonLanguage.Item("MENSAJE_NIVEL") & ": " & Reader.ReadInt8()
        .oro.Caption = JsonLanguage.Item("MENSAJE_ORO") & ": " & Reader.ReadInt32()
        .Banco.Caption = JsonLanguage.Item("MENSAJE_BANCO") & ": " & Reader.ReadInt32()
        .txtPeticiones.text = Reader.ReadString8()
        .guildactual.Caption = JsonLanguage.Item("MENSAJE_CLAN") & ": " & Reader.ReadString8()
        .txtMiembro.text = Reader.ReadString8()
        Dim armada As Boolean
        Dim caos   As Boolean
        armada = Reader.ReadBool()
        caos = Reader.ReadBool()
        If armada Then
            .ejercito.Caption = JsonLanguage.Item("MENSAJE_EJERCITO_ARMADA_REAL")
        ElseIf caos Then
            .ejercito.Caption = JsonLanguage.Item("MENSAJE_EJERCITO_LEGION_OSCURA")
        End If
        .ciudadanos.Caption = JsonLanguage.Item("MENSAJE_CIUDADANOS_ASESINADOS") & ": " & CStr(Reader.ReadInt32())
        .Criminales.Caption = JsonLanguage.Item("MENSAJE_CRIMINALES_ASESINADOS") & ": " & CStr(Reader.ReadInt32())
        Call .Show(vbModeless, GetGameplayForm())
    End With
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCharacterInfo", Erl)
End Sub

''
' Handles the GuildLeaderInfo message.
Private Sub HandleGuildLeaderInfo()
    On Error GoTo errhandler
    Dim str    As String
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
        Dim Nivel  As Byte
        Nivel = Reader.ReadInt8()
        .Nivel = Nivel
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
        Dim Padding As String
        Padding = Space$(27)
        Select Case Nivel
            Case 1
                .beneficios = Padding & JsonLanguage.Item("MENSAJE_BENEFICIOS_MAX_MIEMBROS")
                .maxMiembros = 5
            Case 2
                .beneficios = Padding & JsonLanguage.Item("MENSAJE_MAX_MIEMBROS_8") & " / " & JsonLanguage.Item("MENSAJE_PEDIR_AYUDA")
                .maxMiembros = 8
            Case 3
                .beneficios = Padding & JsonLanguage.Item("MENSAJE_MAX_MIEMBROS_11") & " / " & JsonLanguage.Item("MENSAJE_PEDIR_AYUDA") & " / " & JsonLanguage.Item( _
                        "MENSAJE_SEGURO_CLAN")
                .maxMiembros = 11
            Case 4
                .beneficios = Padding & JsonLanguage.Item("MENSAJE_MAX_MIEMBROS_14") & " / " & JsonLanguage.Item("MENSAJE_PEDIR_AYUDA") & " / " & JsonLanguage.Item( _
                        "MENSAJE_SEGURO_CLAN")
                .maxMiembros = 14
            Case 5
                .beneficios = Padding & JsonLanguage.Item("MENSAJE_MAX_MIEMBROS_17") & " / " & JsonLanguage.Item("MENSAJE_PEDIR_AYUDA") & " / " & JsonLanguage.Item( _
                        "MENSAJE_SEGURO_CLAN") & " / " & JsonLanguage.Item("MENSAJE_VER_VIDA_MANA")
                .maxMiembros = 17
            Case 6
                .beneficios = Padding & JsonLanguage.Item("MENSAJE_MAX_MIEMBROS_20") & " / " & JsonLanguage.Item("MENSAJE_PEDIR_AYUDA") & " / " & JsonLanguage.Item( _
                        "MENSAJE_SEGURO_CLAN") & " / " & JsonLanguage.Item("MENSAJE_VER_VIDA_MANA") & " / " & JsonLanguage.Item("MENSAJE_VERSE_INVISIBLE")
                .maxMiembros = 20
        End Select
        .Show , GetGameplayForm()
    End With
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildLeaderInfo", Erl)
End Sub

''
' Handles the GuildDetails message.
Private Sub HandleGuildDetails()
    On Error GoTo errhandler
    With frmGuildBrief
        If Not .EsLeader Then
        End If
        Dim GuildDetails As t_GuildInfo
        GuildDetails.Name = Reader.ReadString8()
        GuildDetails.Founder = Reader.ReadString8
        GuildDetails.CreationDate = Reader.ReadString8()
        GuildDetails.Leader = Reader.ReadString8()
        GuildDetails.MemberCount = Reader.ReadInt16()
        GuildDetails.Aligment = Reader.ReadString8()
        GuildDetails.Description = Reader.ReadString8()
        GuildDetails.level = Reader.ReadInt8()
        .nombre.Caption = GuildDetails.Name
        .fundador.Caption = GuildDetails.Founder
        .creacion.Caption = GuildDetails.CreationDate
        .lider.Caption = GuildDetails.Founder 'Provisoriamente hacemos que se muestre el fundador como lider
        .miembros.Caption = GuildDetails.MemberCount
        .lblAlineacion.Caption = GuildDetails.Aligment
        .desc.text = GuildDetails.Description
        .Nivel.Caption = GuildDetails.level
    End With
    frmGuildBrief.Show vbModeless, GetGameplayForm()
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildDetails", Erl)
End Sub

''
' Handles the ShowGuildFundationForm message.
Private Sub HandleShowGuildFundationForm()
    'Remove packet ID
    On Error GoTo HandleShowGuildFundationForm_Err
    CreandoClan = True
    frmGuildDetails.Show , GetGameplayForm()
    Exit Sub
HandleShowGuildFundationForm_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowGuildFundationForm", Erl)
End Sub

''
' Handles the ParalizeOK message.
Private Sub HandleParalizeOK()
    'Remove packet ID
    On Error GoTo HandleParalizeOK_Err
    UserParalizado = Not UserParalizado
    Exit Sub
HandleParalizeOK_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleParalizeOK", Erl)
End Sub

Private Sub HandleInmovilizadoOK()
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
    On Error GoTo errhandler
    Call frmUserRequest.recievePeticion(Reader.ReadString8())
    Call frmUserRequest.Show(vbModeless, GetGameplayForm())
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowUserRequest", Erl)
End Sub

''
' Handles the ChangeUserTradeSlot message.
Private Sub HandleChangeUserTradeSlot()
    On Error GoTo errhandler
    Dim miOferta As Boolean
    miOferta = Reader.ReadBool
    Dim i             As Byte
    Dim nombreItem    As String
    Dim cantidad      As Integer
    Dim grhItem       As Long
    Dim ObjIndex      As Integer
    Dim ElementalTags As Long
    If miOferta Then
        Dim OroAEnviar As Long
        OroAEnviar = Reader.ReadInt32
        frmComerciarUsu.lblOroMiOferta.Caption = PonerPuntos(OroAEnviar)
        frmComerciarUsu.lblMyGold.Caption = PonerPuntos(UserStats.GLD - OroAEnviar)
        For i = 1 To 6
            With OtroInventario(i)
                ObjIndex = Reader.ReadInt16
                nombreItem = Reader.ReadString8
                grhItem = Reader.ReadInt32
                cantidad = Reader.ReadInt32
                ElementalTags = Reader.ReadInt32
                If cantidad > 0 Then
                    Call frmComerciarUsu.InvUserSell.SetItem(i, ObjIndex, cantidad, 0, grhItem, 0, 0, 0, 0, 0, nombreItem, ElementalTags, 0)
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
                ElementalTags = Reader.ReadInt32
                If cantidad > 0 Then
                    Call frmComerciarUsu.InvOtherSell.SetItem(i, ObjIndex, cantidad, 0, grhItem, 0, 0, 0, 0, 0, nombreItem, ElementalTags, 0)
                End If
            End With
        Next i
        Call frmComerciarUsu.InvOtherSell.ReDraw
    End If
    frmComerciarUsu.lblEstadoResp.visible = False
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChangeUserTradeSlot", Erl)
End Sub

''
' Handles the SpawnList message.
Private Sub HandleSpawnList()
    On Error GoTo errhandler
    frmSpawnList.ListaCompleta = Reader.ReadBool
    Call frmSpawnList.FillList
    frmSpawnList.Show , GetGameplayForm()
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSpawnList", Erl)
End Sub

''
' Handles the ShowSOSForm message.
Private Sub HandleShowSOSForm()
    On Error GoTo errhandler
    Dim sosList()           As String
    Dim i                   As Long
    Dim nombre              As String
    Dim Consulta            As String
    Dim TipoDeConsulta      As String
    Dim FechaHoraDeConsulta As Date
    sosList = Split(Reader.ReadString8(), SEPARATOR)
    For i = 0 To UBound(sosList())
        nombre = ReadField(1, sosList(i), Asc("Ø"))
        Consulta = ReadField(2, sosList(i), Asc("Ø"))
        TipoDeConsulta = ReadField(3, sosList(i), Asc("Ø"))
        FechaHoraDeConsulta = ReadField(4, sosList(i), Asc("Ø"))
        frmPanelgm.List1.AddItem nombre & "(" & TipoDeConsulta & ") - " & format(FechaHoraDeConsulta, "dd/MM/yyyy hh:mm AM/PM")
        frmPanelgm.List2.AddItem Consulta
    Next i
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowSOSForm", Erl)
End Sub

''
' Handles the ShowMOTDEditionForm message.
Private Sub HandleShowMOTDEditionForm()
    On Error GoTo errhandler
    frmCambiaMotd.txtMotd.text = Reader.ReadString8()
    frmCambiaMotd.Show , GetGameplayForm()
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowMOTDEditionForm", Erl)
End Sub

''
' Handles the ShowGMPanelForm message.
Private Sub HandleShowGMPanelForm()
    'Remove packet ID
    On Error GoTo HandleShowGMPanelForm_Err
    Dim MiCargo As Integer
    #If DEBUGGING = 1 Then
        frmPanelgm.FraControlMacros.visible = False
    #End If
    frmPanelgm.txtHeadNumero = Reader.ReadInt16
    frmPanelgm.txtBodyYo = Reader.ReadInt16
    frmPanelgm.txtCasco = Reader.ReadInt16
    frmPanelgm.txtArma = Reader.ReadInt16
    frmPanelgm.txtEscudo = Reader.ReadInt16
    frmPanelgm.Show vbModeless, GetGameplayForm()
    MiCargo = charlist(UserCharIndex).priv
    Select Case MiCargo ' Ajustar privilejios
        Case 1
            frmPanelgm.mnuChar.visible = False
            frmPanelgm.cmdHerramientas.visible = False
            frmPanelgm.Admin(0).visible = False
        Case 2 'Consejeros
            frmPanelgm.mnuChar.visible = False
            frmPanelgm.cmdHerramientas.visible = False
            frmPanelgm.Admin(0).visible = False
            frmPanelgm.cmdConsulta.visible = False
            frmPanelgm.cmdMatarNPC.visible = False
            frmPanelgm.cmdEventos.visible = False
            frmPanelgm.cmdMapaSeguro.visible = False
            frmPanelgm.cmdInvisible.visible = False
            frmPanelgm.SendGlobal.visible = False
            frmPanelgm.Mensajeria.visible = False
            frmPanelgm.cmdMapeo.visible = False
            frmPanelgm.cmdMapeo.enabled = False
            frmPanelgm.cmdCerrarCliente.visible = False
            frmPanelgm.cmdCerrarCliente.enabled = False
            frmPanelgm.cmdcrearevento.enabled = False
            frmPanelgm.cmdcrearevento.visible = False
            frmPanelgm.txtMod.Width = 4560
            frmPanelgm.Height = 5080
            frmPanelgm.mnuTraer.visible = False
            frmPanelgm.mnuIra.visible = False
        Case 3 ' Semidios
            frmPanelgm.Admin(0).visible = False
            frmPanelgm.cmdcrearevento.enabled = False
            frmPanelgm.cmdcrearevento.visible = False
            frmPanelgm.cmdMapeo.visible = False
            frmPanelgm.cmdMapeo.enabled = False
            frmPanelgm.cmdCerrarCliente.visible = False
            frmPanelgm.cmdCerrarCliente.enabled = False
        Case 4 ' Dios
            frmPanelgm.Admin(0).visible = False
            frmPanelgm.cmdMapeo.visible = False
            frmPanelgm.cmdMapeo.enabled = False
            frmPanelgm.cmdCerrarCliente.visible = False
            frmPanelgm.cmdCerrarCliente.enabled = False
        Case 5
    End Select
    Exit Sub
HandleShowGMPanelForm_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowGMPanelForm", Erl)
End Sub

Private Sub HandleShowFundarClanForm()
    'Remove packet ID
    On Error GoTo HandleShowFundarClanForm_Err
    CreandoClan = True
    frmGuildDetails.Show vbModeless, GetGameplayForm()
    Exit Sub
HandleShowFundarClanForm_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShowFundarClanForm", Erl)
End Sub

''
' Handles the UserNameList message.
Private Sub HandleUserNameList()
    On Error GoTo errhandler
    Dim userList() As String
    Dim i          As Long
    userList = Split(Reader.ReadString8(), SEPARATOR)
    If frmPanelgm.visible Then
        frmPanelgm.cboListaUsus.Clear
        For i = 0 To UBound(userList())
            Call frmPanelgm.cboListaUsus.AddItem(userList(i))
        Next i
        If frmPanelgm.cboListaUsus.ListCount > 0 Then frmPanelgm.cboListaUsus.ListIndex = 0
    End If
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUserNameList", Erl)
End Sub

Private Sub HandleUpdateTagAndStatus()
    On Error GoTo errhandler
    Dim charindex   As Integer
    Dim status      As Byte
    Dim NombreYClan As String
    Dim group_index As Integer
    charindex = Reader.ReadInt16()
    status = Reader.ReadInt8()
    NombreYClan = Reader.ReadString8()
    group_index = Reader.ReadInt16()
    Dim Pos As Integer
    Pos = InStr(NombreYClan, "<")
    If Pos = 0 Then Pos = InStr(NombreYClan, "[")
    If Pos = 0 Then Pos = Len(NombreYClan) + 2
    charlist(charindex).nombre = Left$(NombreYClan, Pos - 2)
    charlist(charindex).clan = mid$(NombreYClan, Pos)
    'Update char status adn tag!
    charlist(charindex).status = status
    charlist(charindex).group_index = group_index
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateTagAndStatus", Erl)
End Sub

Private Sub HandleUserOnline()
    On Error GoTo errhandler
    Dim rdata As Integer
    rdata = Reader.ReadInt16()
    usersOnline = rdata
    frmMain.onlines = "Online: " & usersOnline
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUserOnline", Erl)
End Sub

Private Sub HandleParticleFXToFloor()
    On Error GoTo HandleParticleFXToFloor_Err
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
    Dim x           As Byte
    Dim y           As Byte
    Dim color       As Long
    Dim color_value As RGBA
    Dim Rango       As Byte
    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
    color = Reader.ReadInt32()
    Rango = Reader.ReadInt8()
    Call Long_2_RGBA(color_value, color)
    Dim id  As Long
    Dim id2 As Long
    If color = 0 Then
        If MapData(x, y).luz.Rango > 100 Then
            LucesRedondas.Delete_Light_To_Map x, y
            Exit Sub
        Else
            id = LucesCuadradas.Light_Find(x & y)
            LucesCuadradas.Light_Remove id
            MapData(x, y).luz.color = COLOR_EMPTY
            MapData(x, y).luz.Rango = 0
            Exit Sub
        End If
    End If
    MapData(x, y).luz.color = color_value
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
    Dim charindex      As Integer
    Dim ParticulaIndex As Integer
    Dim Time           As Long
    Dim Remove         As Boolean
    Dim Grh            As Long
    Dim x              As Byte, y As Byte
    charindex = Reader.ReadInt16()
    ParticulaIndex = Reader.ReadInt16()
    Time = Reader.ReadInt32()
    Remove = Reader.ReadBool()
    Grh = Reader.ReadInt32()
    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
    If x + y > 0 Then
        With charlist(charindex)
            If .Invisible And charindex <> UserCharIndex Then
                If MapData(.Pos.x, .Pos.y).charindex = charindex Then MapData(.Pos.x, .Pos.y).charindex = 0
                .Pos.x = x
                .Pos.y = y
                MapData(x, y).charindex = charindex
            End If
        End With
    End If
    If Remove Then
        Call Char_Particle_Group_Remove(charindex, ParticulaIndex)
        charlist(charindex).Particula = 0
    Else
        charlist(charindex).Particula = ParticulaIndex
        charlist(charindex).ParticulaTime = Time
        If Grh > 0 Then
            Call General_Char_Particle_Create(ParticulaIndex, charindex, Time, Grh)
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
    Dim Emisor         As Integer
    Dim receptor       As Integer
    Dim ParticulaViaje As Integer
    Dim ParticulaFinal As Integer
    Dim Time           As Long
    Dim wav            As Integer
    Dim Fx             As Integer
    Dim x              As Byte, y As Byte
    Emisor = Reader.ReadInt16()
    receptor = Reader.ReadInt16()
    ParticulaViaje = Reader.ReadInt16()
    ParticulaFinal = Reader.ReadInt16()
    Time = Reader.ReadInt32()
    wav = Reader.ReadInt16()
    Fx = Reader.ReadInt16()
    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
    If x + y > 0 Then
        With charlist(receptor)
            If .Invisible And receptor <> UserCharIndex Then
                If MapData(.Pos.x, .Pos.y).charindex = receptor Then MapData(.Pos.x, .Pos.y).charindex = 0
                .Pos.x = x
                .Pos.y = y
                MapData(x, y).charindex = receptor
            End If
        End With
    End If
    Engine_spell_Particle_Set (ParticulaViaje)
    Call Effect_Begin(ParticulaViaje, 9, Get_Pixelx_Of_Char(Emisor), Get_PixelY_Of_Char(Emisor), ParticulaFinal, Time, receptor, Emisor, wav, Fx)
    ' charlist(charindex).Particula = ParticulaIndex
    ' charlist(charindex).ParticulaTime = time
    ' Call General_Char_Particle_Create(ParticulaIndex, charindex, time)
    Exit Sub
HandleParticleFXWithDestino_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleParticleFXWithDestino", Erl)
End Sub

Private Sub HandleParticleFXWithDestinoXY()
    On Error GoTo HandleParticleFXWithDestinoXY_Err
    Dim Emisor         As Integer
    Dim ParticulaViaje As Integer
    Dim ParticulaFinal As Integer
    Dim Time           As Long
    Dim wav            As Integer
    Dim Fx             As Integer
    Dim x              As Byte
    Dim y              As Byte
    Emisor = Reader.ReadInt16()
    ParticulaViaje = Reader.ReadInt16()
    ParticulaFinal = Reader.ReadInt16()
    Time = Reader.ReadInt32()
    wav = Reader.ReadInt16()
    Fx = Reader.ReadInt16()
    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
    ' frmdebug.add_text_tracebox "RECIBI FX= " & fX
    Engine_spell_Particle_Set (ParticulaViaje)
    Call Effect_BeginXY(ParticulaViaje, 9, Get_Pixelx_Of_Char(Emisor), Get_PixelY_Of_Char(Emisor), x, y, ParticulaFinal, Time, Emisor, wav, Fx)
    ' charlist(charindex).Particula = ParticulaIndex
    ' charlist(charindex).ParticulaTime = time
    ' Call General_Char_Particle_Create(ParticulaIndex, charindex, time)
    Exit Sub
HandleParticleFXWithDestinoXY_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleParticleFXWithDestinoXY", Erl)
End Sub

Private Sub HandleAuraToChar()
    On Error GoTo HandleAuraToChar_Err
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
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleAuraToChar " & charindex, Erl)
End Sub

Private Sub HandleSpeedToChar()
    On Error GoTo HandleSpeedToChar_Err
    Dim charindex As Integer
    Dim Speeding  As Single
    charindex = Reader.ReadInt16()
    Speeding = Reader.ReadReal32()
    ' (Opcional defensivo)
    If charindex < LBound(charlist) Or charindex > UBound(charlist) Then Exit Sub
    charlist(charindex).Speeding = Speeding
    Call ApplySpeedingToChar(charindex)   ' <- actualiza .speed de las anims
    Exit Sub
HandleSpeedToChar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSpeedToChar", Erl)
End Sub

Private Sub HandleNieveToggle()
    'Remove packet ID
    On Error GoTo HandleNieveToggle_Err
    bNieve = Reader.ReadBool
    If Not InMapBounds(UserPos.x, UserPos.y) Then Exit Sub
    If MapDat.NIEVE = 0 Then Exit Sub
    If bNieve Then
        Call ao20audio.PlayWeatherAudio(IIf(bTecho, SND_NIEVEIN, SND_NIEVEOUT))
        Call Graficos_Particulas.Engine_MeteoParticle_Set(Particula_Nieve, False)
    Else
        Call ao20audio.PlayAmbientAudio(UserMap)
        Call ao20audio.PlayAmbientWav(IIf(bTecho, SND_RAIN_IN_END, SND_RAIN_OUT_END))
        Call Graficos_Particulas.Engine_MeteoParticle_Set(-1)
    End If
    Exit Sub
HandleNieveToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleNieveToggle", Erl)
End Sub

Private Sub HandleNieblaToggle()
    'Remove packet ID
    On Error GoTo HandleNieblaToggle_Err
    MaxAlphaNiebla = Reader.ReadInt8()
    bNiebla = Not bNiebla
    frmMain.TimerNiebla.enabled = True
    Exit Sub
HandleNieblaToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleNieblaToggle", Erl)
End Sub

Private Sub HandleBindKeys()
    On Error GoTo HandleBindKeys_Err
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
    'Recibe y maneja el paquete QuestDetails del servidor.
    On Error GoTo errhandler
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
    FrmQuests.ListView1.ColumnHeaders(2).Width = 780 'Agrando el ancho de la columna para que entre la cantidad de npcs correctamente
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
            FrmQuestInfo.Text1.text = ""
            Call AddtoRichTextBox(FrmQuestInfo.Text1, QuestList(QuestIndex).desc & vbCrLf & vbCrLf & JsonLanguage.Item("MENSAJE_QUEST_REQUISITOS") & vbCrLf & JsonLanguage.Item( _
                    "MENSAJE_QUEST_NIVEL_REQUERIDO") & LevelRequerido & vbCrLf & JsonLanguage.Item("MENSAJE_QUEST_REQUERIDA") & QuestList(QuestRequerida).RequiredQuest, 128, _
                    128, 128)
        Else
            FrmQuestInfo.Text1.text = ""
            Call AddtoRichTextBox(FrmQuestInfo.Text1, QuestList(QuestIndex).desc & vbCrLf & vbCrLf & JsonLanguage.Item("MENSAJE_QUEST_REQUISITOS") & vbCrLf & JsonLanguage.Item( _
                    "MENSAJE_QUEST_NIVEL_REQUERIDO") & LevelRequerido & vbCrLf, 128, 128, 128)
        End If
        tmpByte = Reader.ReadInt8
        If tmpByte Then 'Hay NPCs
            If tmpByte > 5 Then
                FrmQuestInfo.ListView1.FlatScrollBar = False
                FrmQuestInfo.ListView1.ColumnHeaders.Item(1).Width = 1550
            Else
                FrmQuestInfo.ListView1.FlatScrollBar = True
                FrmQuestInfo.ListView1.ColumnHeaders.Item(1).Width = 1800
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
        tmpStr = tmpStr & vbCrLf & JsonLanguage.Item("MENSAJE_RECOMPENSAS") & vbCrLf
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
        FrmQuests.lblRepetible.visible = QuestList(QuestIndex).Repetible = 1
        LevelRequerido = Reader.ReadInt8
        QuestRequerida = Reader.ReadInt16
        FrmQuests.detalle.text = QuestList(QuestIndex).desc & vbCrLf & vbCrLf & JsonLanguage.Item("MENSAJE_QUEST_REQUISITOS") & vbCrLf & JsonLanguage.Item( _
                "MENSAJE_QUEST_NIVEL_REQUERIDO") & LevelRequerido & vbCrLf
        If QuestRequerida <> 0 Then
            FrmQuests.detalle.text = FrmQuests.detalle.text & vbCrLf & JsonLanguage.Item("MENSAJE_QUEST_REQUERIDA") & QuestList(QuestRequerida).nombre
        End If
        tmpStr = tmpStr & vbCrLf & JsonLanguage.Item("MENSAJE_OBJETIVOS") & vbCrLf
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
                If cantok > 0 Then
                    subelemento.SubItems(1) = matados & "/" & cantidadnpc
                Else
                    subelemento.SubItems(1) = "OK"
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
        Dim RequiredSkill, RequiredValue As Byte
        RequiredSkill = Reader.ReadInt8
        RequiredValue = Reader.ReadInt8
        If RequiredSkill > 0 Then
            FrmQuests.detalle.text = FrmQuests.detalle.text & SkillsNames(RequiredSkill) & ": " & RequiredValue
        End If
        tmpStr = tmpStr & vbCrLf & JsonLanguage.Item("MENSAJE_RECOMPENSAS") & vbCrLf
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
        tmpByte = Reader.ReadInt8 'skills
        For i = 1 To tmpByte
            obindex = Reader.ReadInt16
            Set subelemento = FrmQuests.ListView2.ListItems.Add(, , HechizoData(obindex).nombre)
            subelemento.SubItems(1) = 1
            subelemento.SubItems(2) = obindex
            subelemento.SubItems(3) = 1
        Next i
    End If
    'Determinamos que formulario se muestra, según si recibimos la información y la quest está empezada o no.
    If QuestEmpezada Then
        FrmQuests.txtInfo.text = tmpStr
        Call FrmQuests.ListView1_Click
        Call FrmQuests.ListView2_Click
        Call FrmQuests.lstQuests.SetFocus
    Else
        FrmQuestInfo.Show vbModeless, GetGameplayForm()
        FrmQuestInfo.Picture = LoadInterface("ventananuevamision.bmp")
        Call FrmQuestInfo.ListView1_Click
        Call FrmQuestInfo.ListView2_Click
    End If
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleQuestDetails", Erl)
End Sub
 
Public Sub HandleQuestListSend()
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Recibe y maneja el paquete QuestListSend del servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    On Error GoTo errhandler
    Dim i       As Integer
    Dim tmpByte As Byte
    Dim tmpStr  As String
    'Leemos la cantidad de quests que tiene el usuario
    tmpByte = Reader.ReadInt8
    'Limpiamos el ListBox y el TextBox del formulario
    FrmQuests.lstQuests.Clear
    FrmQuests.txtInfo.text = vbNullString
    'Si el usuario tiene quests entonces hacemos el handle
    If tmpByte Then
        'Leemos el string
        tmpStr = Reader.ReadString8
        'Agregamos los items
        For i = 1 To tmpByte
            FrmQuests.lstQuests.AddItem ReadField(i, QuestList(tmpStr).nombre, 59)
        Next i
    End If
    'Mostramos el formulario
    COLOR_AZUL = RGB(0, 0, 0)
    Call Establecer_Borde(FrmQuests.lstQuests, FrmQuests, COLOR_AZUL, 0, 0)
    FrmQuests.Picture = LoadInterface("ventanadetallemision.bmp")
    FrmQuests.Show vbModeless, GetGameplayForm()
    'Pedimos la informacion de la primer quest (si la hay)
    If tmpByte Then Call WriteQuestDetailsRequest(1)
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleQuestListSend", Erl)
End Sub

Public Sub HandleNpcQuestListSend()
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Recibe y maneja el paquete QuestListSend del servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    On Error GoTo errhandler
    Dim tmpStr         As String
    Dim tmpByte        As Byte
    Dim QuestEmpezada  As Boolean
    Dim i              As Integer
    Dim J              As Byte
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
    For J = 1 To CantidadQuest
        QuestIndex = Reader.ReadInt16
        FrmQuestInfo.titulo.Caption = QuestList(QuestIndex).nombre
        QuestList(QuestIndex).RequiredLevel = Reader.ReadInt8
        QuestList(QuestIndex).RequiredQuest = Reader.ReadInt16
        QuestList(QuestIndex).RequiredClass = Reader.ReadInt8
        QuestList(QuestIndex).LimitLevel = Reader.ReadInt8
        tmpByte = Reader.ReadInt8
        If tmpByte Then 'Hay NPCs
            If tmpByte > 5 Then
                FrmQuestInfo.ListView1.FlatScrollBar = False
                FrmQuestInfo.ListView1.ColumnHeaders.Item(1).Width = 1550
            Else
                FrmQuestInfo.ListView1.FlatScrollBar = True
                FrmQuestInfo.ListView1.ColumnHeaders.Item(1).Width = 1800
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
        tmpByte = Reader.ReadInt8 ' required spells
        If tmpByte Then
            ReDim QuestList(QuestIndex).RequiredSpellList(1 To tmpByte)
            For i = 1 To tmpByte
                QuestList(QuestIndex).RequiredSpellList(i) = Reader.ReadInt16
            Next i
        Else
            ReDim QuestList(QuestIndex).RequiredSpellList(0)
        End If
        QuestList(QuestIndex).RequiredSkill.SkillType = Reader.ReadInt8
        QuestList(QuestIndex).RequiredSkill.RequiredValue = Reader.ReadInt8
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
        QuestList(QuestIndex).RewardSkillCount = Reader.ReadInt8
        If QuestList(QuestIndex).RewardSkillCount > 0 Then
            ReDim QuestList(QuestIndex).RewardSkill(1 To QuestList(QuestIndex).RewardSkillCount)
            For i = 1 To QuestList(QuestIndex).RewardSkillCount
                QuestList(QuestIndex).RewardSkill(i) = Reader.ReadInt16
            Next i
        End If
        estado = Reader.ReadInt8
        Repetible = QuestList(QuestIndex).Repetible = 1
        Set subelemento = FrmQuestInfo.ListViewQuest.ListItems.Add(, , QuestList(QuestIndex).nombre & IIf(Repetible, " (R)", ""))
        subelemento.SubItems(2) = QuestIndex
        Select Case estado
            Case 0
                subelemento.SubItems(1) = JsonLanguage.Item("MENSAJE_DISPONIBLE")
                subelemento.ForeColor = vbWhite
                subelemento.ListSubItems(1).ForeColor = vbWhite
            Case 1
                subelemento.SubItems(1) = JsonLanguage.Item("MENSAJE_EN_CURSO")
                subelemento.ForeColor = RGB(255, 175, 10)
                subelemento.ListSubItems(1).ForeColor = RGB(255, 175, 10)
            Case 2
                If Repetible Then
                    subelemento.SubItems(1) = JsonLanguage.Item("MENSAJE_REPETIBLE")
                    subelemento.ForeColor = RGB(180, 180, 180)
                    subelemento.ListSubItems(1).ForeColor = RGB(180, 180, 180)
                Else
                    subelemento.SubItems(1) = JsonLanguage.Item("MENSAJE_FINALIZADA")
                    subelemento.ForeColor = RGB(15, 140, 50)
                    subelemento.ListSubItems(1).ForeColor = RGB(15, 140, 50)
                End If
            Case 3
                subelemento.SubItems(1) = JsonLanguage.Item("MENSAJE_NO_DISPONIBLE")
                subelemento.ForeColor = RGB(255, 10, 10)
                subelemento.ListSubItems(1).ForeColor = RGB(255, 10, 10)
        End Select
        FrmQuestInfo.ListViewQuest.Refresh
    Next J
    'Determinamos que formulario se muestra, segun si recibimos la informacion y la quest está empezada o no.
    FrmQuestInfo.Show vbModeless, GetGameplayForm()
    FrmQuestInfo.Picture = LoadInterface("ventananuevamision.bmp")
    Call FrmQuestInfo.ShowQuest(1)
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleNpcQuestListSend", Erl)
End Sub

Private Sub HandleShowPregunta()
    On Error GoTo errhandler
    Dim MsgID As Integer
    Dim param As String
    MsgID = Reader.ReadInt16()
    param = Reader.ReadString8()
    PreguntaScreen = Locale_Parse_ServerMessage(MsgID, param)
    Pregunta = True
    Exit Sub
errhandler:
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
    FrmGrupo.Show , GetGameplayForm()
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
        frmMain.personaje(miembro).visible = False
    Else
        If UserMap = map Then
            frmMain.personaje(miembro).visible = True
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
    FrmViajes.Show , GetGameplayForm()
    Exit Sub
HandleViajarForm_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleViajarForm", Erl)
End Sub

Private Sub HandleSeguroResu()
    'Get data and update form
    SeguroResuX = Reader.ReadBool()
    If SeguroResuX Then
        Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_SEGURO_RESURRECCION_ACTIVADO"), 65, 190, 156, False, False, False)
        frmMain.ImgSegResu = LoadInterface("boton-fantasma-on.bmp")
    Else
        Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_SEGURO_RESURRECCION_DESACTIVADO"), 65, 190, 156, False, False, False)
        frmMain.ImgSegResu = LoadInterface("boton-fantasma-off.bmp")
    End If
End Sub

Private Sub HandleLegionarySecure()
    'Get data and update form
    LegionarySecureX = Reader.ReadBool()
    If LegionarySecureX Then
        Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_SEGURO_LEGION_ACTIVADO"), 65, 190, 156, False, False, False)
        frmMain.ImgLegionarySecure = LoadInterface("boton-demonio-on.bmp")
        'SeguroFaccX = True
    Else
        Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_SEGURO_LEGION_DESACTIVADO"), 65, 190, 156, False, False, False)
        frmMain.ImgLegionarySecure = LoadInterface("boton-demonio-off.bmp")
    End If
End Sub

Private Sub HandleStopped()
    UserStopped = Reader.ReadBool()
End Sub

Private Sub HandleInvasionInfo()
    InvasionActual = Reader.ReadInt8
    InvasionPorcentajeVida = Reader.ReadInt8
    InvasionPorcentajeTiempo = Reader.ReadInt8
    frmMain.Evento.enabled = False
    frmMain.Evento.Interval = 0
    frmMain.Evento.Interval = 10000
    frmMain.Evento.enabled = True
End Sub

Private Sub HandleCommerceRecieveChatMessage()
    Dim message As String
    message = Reader.ReadString8
    Call AddtoRichTextBox(frmComerciarUsu.RecTxt, message, 255, 255, 255, 0, False, True)
End Sub

Private Sub HandleDoAnimation()
    On Error GoTo HandleDoAnimation_Err
    Dim charindex As Integer
    Dim oldWalk   As Grh
    Dim keepStart As Long
    charindex = Reader.ReadInt16()
    With charlist(charindex)
        ' Guardar el walk anterior ANTES de cambiar el Body
        oldWalk = .Body.Walk(.Heading)
        .AnimatingBody = Reader.ReadInt16()
        ' Calcular el "started" preservando fase si ya estaba animando
        If oldWalk.started > 0 And .Moving Then
            keepStart = SyncGrhPhase(oldWalk, BodyData(.AnimatingBody).Walk(.Heading).GrhIndex)
        Else
            keepStart = FrameTime
        End If
        ' Aplicar el cambio de Body y setear started preservado
        .Body = BodyData(.AnimatingBody)
        If .Body.Walk(.Heading).started = 0 Or keepStart <> 0 Then
            .Body.Walk(.Heading).started = keepStart
            ' Hacer que la animación de casteo sea de UNA SOLA pasada.
            ' Si queda en INFINITE_LOOPS nunca vuelve al idle.
            .Body.Walk(.Heading).Loops = 0
            ' (opcional) evitar que arma/escudo queden loopeando durante el cast
            If .Arma.WeaponWalk(.Heading).GrhIndex <> 0 Then .Arma.WeaponWalk(.Heading).Loops = 0
            If .Escudo.ShieldWalk(.Heading).GrhIndex <> 0 Then .Escudo.ShieldWalk(.Heading).Loops = 0
        End If
        ' Mantener arma/escudo en fase con el cuerpo (solo si están “apagados”)
        If .Arma.WeaponWalk(.Heading).started = 0 Then
            .Arma.WeaponWalk(.Heading).started = .Body.Walk(.Heading).started
        End If
        If .Escudo.ShieldWalk(.Heading).started = 0 Then
            .Escudo.ShieldWalk(.Heading).started = .Body.Walk(.Heading).started
        End If
        .Idle = False
    End With
    Exit Sub
HandleDoAnimation_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleDoAnimation", Erl)
End Sub

Private Sub HandleOpenCrafting()
    Dim TIPO As Byte
    TIPO = Reader.ReadInt8
    frmCrafteo.Picture = LoadInterface(TipoCrafteo(TIPO).Ventana)
    frmCrafteo.InventoryGrhIndex = TipoCrafteo(TIPO).Inventario
    frmCrafteo.TipoGrhIndex = TipoCrafteo(TIPO).Icono
    Dim i As Long
    For i = 1 To MAX_INVENTORY_SLOTS
        With frmMain.Inventario
            Call frmCrafteo.InvCraftUser.SetItem(i, .ObjIndex(i), .Amount(i), .Equipped(i), .GrhIndex(i), .ObjType(i), .MaxHit(i), .MinHit(i), .Def(i), .Valor(i), .ItemName(i), _
                    .ElementalTags(i), .PuedeUsar(i))
        End With
    Next i
    For i = 1 To MAX_SLOTS_CRAFTEO
        Call frmCrafteo.InvCraftItems.ClearSlot(i)
    Next i
    Call frmCrafteo.InvCraftCatalyst.ClearSlot(1)
    Call frmCrafteo.SetResult(0, 0, 0)
    Comerciando = True
    frmCrafteo.Show , GetGameplayForm()
End Sub

Private Sub HandleCraftingItem()
    Dim Slot As Byte, ObjIndex As Integer
    Slot = Reader.ReadInt8
    ObjIndex = Reader.ReadInt16
    If ObjIndex <> 0 Then
        With ObjData(ObjIndex)
            Call frmCrafteo.InvCraftItems.SetItem(Slot, ObjIndex, 1, 0, .GrhIndex, .ObjType, 0, 0, 0, .Valor, .Name, .ElementalTags, 0)
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
            Call frmCrafteo.InvCraftCatalyst.SetItem(1, ObjIndex, Amount, 0, .GrhIndex, .ObjType, 0, 0, 0, .Valor, .Name, .ElementalTags, 0)
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

Public Sub HandleAnswerReset()
    On Error GoTo errhandler
    If MsgBox(JsonLanguage.Item("MENSAJEBOX_RESETEAR_PERSONAJE"), vbYesNo, "Resetear personaje") = vbYes Then
        Call WriteResetearPersonaje
    End If
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleAnswerReset", Erl)
End Sub

Public Sub HandleUpdateBankGld()
    On Error GoTo errhandler
    Dim UserBoveOro As Long
    UserBoveOro = Reader.ReadInt32
    Call frmGoliath.UpdateBankGld(UserBoveOro)
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateBankGld", Erl)
End Sub

Public Sub HandlePelearConPezEspecial()
    On Error GoTo errhandler
    PosicionBarra = 1
    DireccionBarra = 1
    Dim i As Integer
    For i = 1 To MAX_INTENTOS
        intentosPesca(i) = 0
    Next i
    PescandoEspecial = True
    Call ao20audio.PlayWav(55)
    ContadorIntentosPescaEspecial_Fallados = 0
    ContadorIntentosPescaEspecial_Acertados = 0
    startTimePezEspecial = GetTickCount()
    Call Char_Dialog_Set(UserCharIndex, JsonLanguage.Item("MENSAJE_SUPER_PEZ"), &H1FFFF, 200, 130)
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePelearConPezEspecial", Erl)
End Sub

Public Sub HandlePrivilegios()
    On Error GoTo errhandler
    EsGM = Reader.ReadBool
    If EsGM Then
        frmMain.panelGM.visible = True
        frmMain.createObj.visible = True
        frmMain.btnInvisible.visible = True
        frmMain.btnSpawn.visible = True
        frmMain.onlines.visible = True
    Else
        frmMain.panelGM.visible = False
        frmMain.createObj.visible = False
        frmMain.btnInvisible.visible = False
        frmMain.btnSpawn.visible = False
        frmMain.onlines.visible = False
    End If
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePrivilegios", Erl)
End Sub

Public Sub HandleShopPjsInit()
    frmShopPjsAO20.Show , GetGameplayForm()
End Sub
Public Sub HandleShopInit()
    On Error GoTo HandleShopInit_Err
    Dim cant_obj_shop As Long, i As Long, J As Long
    Dim tmp As ObjDatas
    cant_obj_shop = Reader.ReadInt16
    credits_shopAO20 = Reader.ReadInt32
    frmShopAO20.lblCredits.Caption = credits_shopAO20
    ReDim ObjShop(1 To cant_obj_shop) As ObjDatas

    ' Leer todos los objetos
    For i = 1 To cant_obj_shop
        ObjShop(i).ObjNum = Reader.ReadInt32
        ObjShop(i).Valor = Reader.ReadInt32
        ObjShop(i).Name = Reader.ReadString8
    Next i

    ' Ordenar por ObjType y luego por Name (alfabético)
    For i = 1 To cant_obj_shop - 1
        For J = i + 1 To cant_obj_shop
    
            Dim typeI As Long
            Dim typeJ As Long
    
            typeI = ObjData(ObjShop(i).ObjNum).ObjType
            typeJ = ObjData(ObjShop(J).ObjNum).ObjType
    
            ' Si el ObjType es mayor, intercambiar
            If typeI > typeJ Then
                tmp = ObjShop(i)
                ObjShop(i) = ObjShop(J)
                ObjShop(J) = tmp
    
            ' Si el ObjType es igual, ordenar por Name
            ElseIf typeI = typeJ Then
                If StrComp(ObjShop(i).Name, ObjShop(J).Name, vbTextCompare) > 0 Then
                    tmp = ObjShop(i)
                    ObjShop(i) = ObjShop(J)
                    ObjShop(J) = tmp
                End If
            End If
        Next J
    Next i

    ' Agregar al ListBox ya ordenado
    For i = 1 To cant_obj_shop
        With frmShopAO20.lstItemShopFilter
            .AddItem ObjShop(i).Name
            .ItemData(.NewIndex) = i
        End With
    Next i

    frmShopAO20.Show , GetGameplayForm()
    frmShopAO20.ResetShopPreview
    Exit Sub
HandleShopInit_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleShopInit", Erl)
End Sub
Public Function GetObjRopajeHumano(ByVal objNum As Long) As Long
    On Error GoTo GetObjRopajeHumano_Err
    If objNum < LBound(ObjData) Or objNum > UBound(ObjData) Then Exit Function
    GetObjRopajeHumano = ObjData(objNum).RopajeHumano
    If GetObjRopajeHumano <> 0 Then Exit Function

    GetObjRopajeHumano = FetchRopajeHumanoFromIndex(objNum)
    If GetObjRopajeHumano <> 0 Then
        ObjData(objNum).RopajeHumano = GetObjRopajeHumano
        Exit Function
    End If

    GetObjRopajeHumano = ParseRopajeHumanoDescriptor(ObjData(objNum).info)
    If GetObjRopajeHumano <> 0 Then Exit Function
    GetObjRopajeHumano = ParseRopajeHumanoDescriptor(ObjData(objNum).Texto)
    If GetObjRopajeHumano <> 0 Then Exit Function
    GetObjRopajeHumano = ParseRopajeHumanoDescriptor(ObjData(objNum).en_texto)
    Exit Function

GetObjRopajeHumano_Err:
    GetObjRopajeHumano = 0
End Function

Private Function FetchRopajeHumanoFromIndex(ByVal objNum As Long) As Long
    On Error GoTo FetchRopajeHumanoFromIndex_Err
    If ObjIndexData Is Nothing Then Exit Function
    Dim rawValue As String
    rawValue = ObjIndexData.GetValue("OBJ" & objNum, "RopajeHumano")
    If LenB(rawValue) = 0 Then Exit Function
    FetchRopajeHumanoFromIndex = val(rawValue)
    If FetchRopajeHumanoFromIndex <> 0 Then
        Debug.Print "[ShopRopaje] ObjNum=" & objNum & _
                    " RopajeHumanoRaw=""" & rawValue & """" & _
                    " Parsed=" & FetchRopajeHumanoFromIndex
    End If
    Exit Function

FetchRopajeHumanoFromIndex_Err:
    Debug.Print "[ShopRopaje] FetchRopajeHumanoFromIndex error " & Err.Number & " - " & Err.Description
End Function

Private Function ParseRopajeHumanoDescriptor(ByVal descriptor As String) As Long
    On Error GoTo ParseRopajeHumanoDescriptor_Err
    If LenB(descriptor) = 0 Then Exit Function
    Dim searchStart As Long
    Dim tokenPos As Long
    Dim tokenLength As Long
    searchStart = 1
    Do
        tokenPos = FindNextRopajeHumanoToken(descriptor, searchStart, tokenLength)
        If tokenPos = 0 Then Exit Do
        ParseRopajeHumanoDescriptor = ExtractNumberAfterToken(descriptor, tokenPos + tokenLength)
        If ParseRopajeHumanoDescriptor <> 0 Then Exit Function
        searchStart = tokenPos + tokenLength
    Loop
    Exit Function

ParseRopajeHumanoDescriptor_Err:
    ParseRopajeHumanoDescriptor = 0
End Function

Private Function FindNextRopajeHumanoToken(ByVal descriptor As String, _
                                           ByVal startPos As Long, _
                                           ByRef tokenLength As Long) As Long
    Dim descriptorLower As String
    Dim candidatePos As Long
    Dim bestPos As Long
    descriptorLower = LCase$(descriptor)
    bestPos = 0
    tokenLength = 0

    candidatePos = InStr(startPos, descriptorLower, "ropajehumano", vbBinaryCompare)
    If candidatePos <> 0 Then
        bestPos = candidatePos
        tokenLength = Len("ropajehumano")
    End If

    candidatePos = InStr(startPos, descriptorLower, "ropaje humano", vbBinaryCompare)
    If candidatePos <> 0 Then
        If bestPos = 0 Or candidatePos < bestPos Then
            bestPos = candidatePos
            tokenLength = Len("ropaje humano")
        End If
    End If

    candidatePos = InStr(startPos, descriptorLower, "ropaje_humano", vbBinaryCompare)
    If candidatePos <> 0 Then
        If bestPos = 0 Or candidatePos < bestPos Then
            bestPos = candidatePos
            tokenLength = Len("ropaje_humano")
        End If
    End If

    candidatePos = InStr(startPos, descriptorLower, "ropaje-humano", vbBinaryCompare)
    If candidatePos <> 0 Then
        If bestPos = 0 Or candidatePos < bestPos Then
            bestPos = candidatePos
            tokenLength = Len("ropaje-humano")
        End If
    End If

    FindNextRopajeHumanoToken = bestPos
End Function

Private Function ExtractNumberAfterToken(ByVal source As String, ByVal startPos As Long) As Long
    Dim pos As Long
    Dim descriptorLength As Long
    descriptorLength = Len(source)
    pos = startPos
    Do While pos <= descriptorLength
        Dim ch As String * 1
        ch = Mid$(source, pos, 1)
        Select Case ch
            Case "0" To "9"
                Exit Do
            Case " ", vbTab, "=", ":", "-", "_", "(", ")", "[", "]", "{", "}", ".", ","
                pos = pos + 1
            Case Else
                Dim code As Integer
                code = Asc(ch)
                If (code >= Asc("A") And code <= Asc("Z")) Or _
                   (code >= Asc("a") And code <= Asc("z")) Then
                    Exit Function
                End If
                pos = pos + 1
        End Select
    Loop
    If pos > descriptorLength Then Exit Function
    ExtractNumberAfterToken = Val(Mid$(source, pos))
End Function

Public Sub HandleUpdateShopClienteCredits()
    credits_shopAO20 = Reader.ReadInt32
    frmShopAO20.lblCredits.Caption = credits_shopAO20
End Sub

Public Sub HandleSendSkillCdUpdate()
    On Error GoTo errhandler
    Dim Effect      As t_ActiveEffect
    Dim ElapsedTime As Long
    Effect.TypeId = Reader.ReadInt16
    Effect.id = Reader.ReadInt32
    ElapsedTime = Reader.ReadInt32
    Effect.duration = Reader.ReadInt32
    Effect.EffectType = Reader.ReadInt8
    Effect.Grh = EffectResources(Effect.TypeId).GrhId
    Effect.startTime = GetTickCount() - (Effect.duration - ElapsedTime)
    Effect.StackCount = Reader.ReadInt16()
    If Effect.EffectType = eBuff Then
        Call AddOrUpdateEffect(BuffList, Effect)
    End If
    If Effect.EffectType = eDebuff Then
        Call AddOrUpdateEffect(DeBuffList, Effect)
    End If
    If Effect.EffectType = eCD Then
        Call AddOrUpdateEffect(CDList, Effect)
    End If
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSendSkillCdUpdate " & Effect.TypeId, Erl)
End Sub

Public Sub HandleSendClientToggles()
    On Error GoTo errhandler
    Dim ToggleCount As Integer
    ToggleCount = Reader.ReadInt16
    Dim i          As Integer
    Dim ToggleName As String
    For i = 0 To ToggleCount - 1
        ToggleName = Reader.ReadString8
        If ToggleName = "hotokey-enabled" Then
            Call SetMask(FeatureToggles, eEnableHotkeys)
        End If
    Next i
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSendClientToggles", Erl)
End Sub

Public Sub HandleObjQuestListSend()
    'Recibe y maneja el paquete QuestListSend del servidor.
    On Error GoTo errhandler
    Dim tmpStr         As String
    Dim tmpByte        As Byte
    Dim QuestEmpezada  As Boolean
    Dim i              As Integer
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
    QuestIndex = Reader.ReadInt16
    FrmQuestInfo.titulo.Caption = QuestList(QuestIndex).nombre
    QuestList(QuestIndex).RequiredLevel = Reader.ReadInt8
    QuestList(QuestIndex).RequiredQuest = Reader.ReadInt16
    tmpByte = Reader.ReadInt8
    If tmpByte Then 'Hay NPCs
        If tmpByte > 5 Then
            FrmQuestInfo.ListView1.FlatScrollBar = False
            FrmQuestInfo.ListView1.ColumnHeaders.Item(1).Width = 1550
        Else
            FrmQuestInfo.ListView1.FlatScrollBar = True
            FrmQuestInfo.ListView1.ColumnHeaders.Item(1).Width = 1800
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
    tmpByte = Reader.ReadInt8 ' required spells
    If tmpByte Then
        ReDim QuestList(QuestIndex).RequiredSpellList(1 To tmpByte)
        For i = 1 To tmpByte
            QuestList(QuestIndex).RequiredSpellList(i) = Reader.ReadInt16
        Next i
    Else
        ReDim QuestList(QuestIndex).RequiredSpellList(0)
    End If
    QuestList(QuestIndex).RequiredSkill.SkillType = Reader.ReadInt8
    QuestList(QuestIndex).RequiredSkill.RequiredValue = Reader.ReadInt8
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
            subelemento.SubItems(1) = JsonLanguage.Item("MENSAJE_DISPONIBLE")
            subelemento.ForeColor = vbWhite
            subelemento.ListSubItems(1).ForeColor = vbWhite
        Case 1
            subelemento.SubItems(1) = JsonLanguage.Item("MENSAJE_EN_CURSO")
            subelemento.ForeColor = RGB(255, 175, 10)
            subelemento.ListSubItems(1).ForeColor = RGB(255, 175, 10)
        Case 2
            If Repetible Then
                subelemento.SubItems(1) = JsonLanguage.Item("MENSAJE_REPETIBLE")
                subelemento.ForeColor = RGB(180, 180, 180)
                subelemento.ListSubItems(1).ForeColor = RGB(180, 180, 180)
            Else
                subelemento.SubItems(1) = JsonLanguage.Item("MENSAJE_FINALIZADA")
                subelemento.ForeColor = RGB(15, 140, 50)
                subelemento.ListSubItems(1).ForeColor = RGB(15, 140, 50)
            End If
        Case 3
            subelemento.SubItems(1) = JsonLanguage.Item("MENSAJE_NO_DISPONIBLE")
            subelemento.ForeColor = RGB(255, 10, 10)
            subelemento.ListSubItems(1).ForeColor = RGB(255, 10, 10)
    End Select
    FrmQuestInfo.ListViewQuest.Refresh
    'Determinamos que formulario se muestra, segun si recibimos la informacion y la quest está empezada o no.
    FrmQuestInfo.Show vbModeless, GetGameplayForm()
    FrmQuestInfo.Picture = LoadInterface("ventananuevamision.bmp")
    Call FrmQuestInfo.ShowQuest(1)
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleNpcQuestListSend", Erl)
End Sub

Public Sub HandleDebugDataResponse()
    On Error GoTo HandleDebugResponse_Err
    Dim cantidadDeMensajes As Integer
    Dim mensaje            As String
    Dim i                  As Integer
    cantidadDeMensajes = Reader.ReadInt16
    Dim File As Integer: File = FreeFile
    Open App.path & "\logs\RemoteError.log" For Append As #File
    For i = 1 To cantidadDeMensajes
        mensaje = Reader.ReadString8
        If LenB(mensaje) <> 0 Then
            Print #File, mensaje
            Print #File, vbNullString
        End If
    Next i
    Close #File
    Exit Sub
HandleDebugResponse_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleDebugResponse", Erl)
End Sub

Public Sub HandleAntiCheatStartSession()
    On Error GoTo HandleAntiCheatStartSession_Err
    Call BeginAntiCheatSession
    Exit Sub
HandleAntiCheatStartSession_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleAntiCheatStartSession", Erl)
End Sub

Public Sub HandleAntiCheatMessage()
    On Error GoTo HandleAntiCheatMessage_Err
    Dim Buffer() As Byte
    Call Reader.ReadSafeArrayInt8(Buffer)
    Call HandleAntiCheatServerMessage(Buffer)
    Exit Sub
HandleAntiCheatMessage_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleAntiCheatMessage", Erl)
End Sub

Public Sub HandleReportLobbyList()
    On Error GoTo HandleReportLobbyList_Err
    Dim OpenLobbyCount As Integer
    Dim LobbyList()    As t_LobbyData
    OpenLobbyCount = Reader.ReadInt16
    ReDim LobbyList(OpenLobbyCount) As t_LobbyData
    Dim i As Integer
    For i = 1 To OpenLobbyCount
        LobbyList(i).Index = i - 1
        LobbyList(i).id = Reader.ReadInt16
        LobbyList(i).Description = Reader.ReadString8
        LobbyList(i).ScenarioType = Reader.ReadString8
        LobbyList(i).MinLevel = Reader.ReadInt16
        LobbyList(i).MaxLevel = Reader.ReadInt16
        LobbyList(i).MinPlayers = Reader.ReadInt16
        LobbyList(i).MaxPlayers = Reader.ReadInt16
        LobbyList(i).RegisteredPlayers = Reader.ReadInt16
        LobbyList(i).TeamSize = Reader.ReadInt16
        LobbyList(i).TeamType = Reader.ReadInt16
        LobbyList(i).InscriptionPrice = Reader.ReadInt32
        LobbyList(i).IsPrivate = Reader.ReadInt8
    Next i
    Call frmLobbyBattleground.SetLobbyList(LobbyList)
    frmLobbyBattleground.Show
    Exit Sub
HandleReportLobbyList_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleDebugResponse", Erl)
End Sub

Private Sub HandleChangeSkinSlot()

Dim Slot                        As Byte
Dim ObjIndex                    As Integer
Dim GrhIndex                    As Long
Dim Amount                      As Integer
Dim Equipped                    As Boolean
Dim Name                        As String
Dim ObjType                     As Byte

    On Error GoTo HandleChangeSkinSlot_Error

    With Reader
        Slot = .ReadInt8
        ObjIndex = .ReadInt16
        Equipped = .ReadBool
        
        GrhIndex = .ReadInt32
        ObjType = .ReadInt8
        Name = .ReadString8

        Debug.Print "Skin SLOT: " & Slot & " objIndex: " & ObjIndex & " Amount: 1 Time: " & Time & " Date: " & Date

        If Slot > 0 Then
            With a_Skins(Slot)
                .Amount = 1
                .ObjIndex = ObjIndex
                .ObjType = ObjType
                .Equipped = Equipped
                .Name = Name
                .PuedeUsar = 0
                .GrhIndex = GrhIndex
            End With
        End If

        Call Load(frmSkins)
        Call frmSkins.InvSkins.SetItem(Slot, ObjIndex, 1, CByte(Equipped), GrhIndex, ObjType, 0, 0, 0, 0, Name, 0, 0)
    End With

    On Error GoTo 0
    Exit Sub

HandleChangeSkinSlot_Error:

    Call LogError("Error " & Err.Number & " (" & Err.Description & ") en el procedimiento HandleChangeSkinSlot del módulo Módulo Protocol en la línea: " & Erl())

End Sub

#If PYMMO = 0 Then
    Public Sub HandleAccountCharacterList()
        CantidadDePersonajesEnCuenta = Reader.ReadInt
        Dim ii As Byte
        For ii = 1 To MAX_PERSONAJES_EN_CUENTA
            Pjs(ii).nombre = ""
            Pjs(ii).Head = 0 ' si is_sailing o muerto, cabeza en 0
            Pjs(ii).Clase = 0
            Pjs(ii).Body = 0
            Pjs(ii).Mapa = 0
            Pjs(ii).PosX = 0
            Pjs(ii).PosY = 0
            Pjs(ii).Nivel = 0
            Pjs(ii).Criminal = 0
            Pjs(ii).Casco = 0
            Pjs(ii).Escudo = 0
            Pjs(ii).Arma = 0
            Pjs(ii).ClanName = ""
            Pjs(ii).NameMapa = ""
            Pjs(ii).Backpack = 0
        Next ii
        For ii = 1 To min(CantidadDePersonajesEnCuenta, MAX_PERSONAJES_EN_CUENTA)
            Pjs(ii).nombre = Reader.ReadString8
            Pjs(ii).Body = Reader.ReadInt
            Pjs(ii).Head = Reader.ReadInt
            Pjs(ii).Clase = Reader.ReadInt
            Pjs(ii).Mapa = Reader.ReadInt
            Pjs(ii).PosX = Reader.ReadInt
            Pjs(ii).PosY = Reader.ReadInt
            Pjs(ii).Nivel = Reader.ReadInt
            Pjs(ii).Criminal = Reader.ReadInt
            Pjs(ii).Casco = Reader.ReadInt
            Pjs(ii).Escudo = Reader.ReadInt
            Pjs(ii).Arma = Reader.ReadInt
            Pjs(ii).Backpack = Reader.ReadInt
            Pjs(ii).ClanName = ""
        Next ii
        For ii = 1 To min(CantidadDePersonajesEnCuenta, MAX_PERSONAJES_EN_CUENTA)
            If Pjs(ii).Body = DwarvenExoesqueletonBody Then
                Pjs(ii).Head = 0
                Pjs(ii).Casco = 0
                Pjs(ii).Escudo = 0
                Pjs(ii).Arma = 0
                Pjs(ii).Backpack = 0
            End If
        Next ii
        Dim i As Long
        For i = 1 To min(CantidadDePersonajesEnCuenta, MAX_PERSONAJES_EN_CUENTA)
            Select Case Pjs(i).Criminal
                Case 0 'Criminal
                    Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(50).R, ColoresPJ(50).G, ColoresPJ(50).B)
                    Pjs(i).priv = 0
                Case 1 'Ciudadano
                    Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(49).R, ColoresPJ(49).G, ColoresPJ(49).B)
                    Pjs(i).priv = 0
                Case 2 'Caos
                    Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(6).R, ColoresPJ(6).G, ColoresPJ(6).B)
                    Pjs(i).priv = 0
                Case 3 'Armada
                    Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(8).R, ColoresPJ(8).G, ColoresPJ(8).B)
                    Pjs(i).priv = 0
                Case 4 'consejero
                    Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(1).R, ColoresPJ(1).G, ColoresPJ(1).B)
                Case 5 'semi dios
                    Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(3).R, ColoresPJ(3).G, ColoresPJ(3).B)
                Case 6 'dios
                    Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(4).R, ColoresPJ(4).G, ColoresPJ(4).B)
                Case 7 'admin
                    Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(5).R, ColoresPJ(5).G, ColoresPJ(5).B)
                Case Else 'es raro o rolemaster
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
        Call LoadCharacterSelectionScreen
    End Sub

#End If
