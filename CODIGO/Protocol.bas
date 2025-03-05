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
            Debug.Print "Logged"
            Dim dummy As Boolean
            dummy = Reader.ReadBool
            Call SaveStringInFile("Logged with character " & CharacterRemote, "remote_debug.txt")
            InitiateShutdownProcess = True
            ShutdownProcessTimer.Start
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
        Case ServerPacketID.eForceCharMoveSiguiendo
            Call HandleForceCharMoveSiguiendo
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
        Case ServerPacketID.eRecievePosSeguimiento
            Call HandleRecievePosSeguimiento
        Case ServerPacketID.eCancelarSeguimiento
            Call HandleCancelarSeguimiento
        Case ServerPacketID.eGetInventarioHechizos
            Call HandleGetInventarioHechizos
        Case ServerPacketID.eNotificarClienteCasteo
            Call HandleNotificarClienteCasteo
        Case ServerPacketID.eSendFollowingCharindex
            Call HandleSendFollowingCharindex
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
        Case ServerPacketID.eForceUpdate
            Call HandleForceUpdate
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
        Case ServerPacketID.eRequestTelemetry
            Call HandleRequestTelemetry
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
        #If PYMMO = 0 Then
        Case ServerPacketID.eAccountCharacterList
            Call HandleAccountCharacterList
        #End If
        Case Else
            Err.Raise &HDEADBEEF, "Invalid Message"
    End Select
#End If
    
    If (Reader.GetAvailable() > 0) Then
        Err.Raise &HDEADBEEF, "HandleIncomingData", "El paquete '" & PacketId & "' se encuentra en mal estado con '" & Reader.GetAvailable() & "' bytes de mas"
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
    Dim i As Integer
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
    SeguroParty = True
    SeguroClanX = True
    SeguroResuX = True
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

    If UserCharIndex = 0 Then Exit Sub
    
    charlist(UserCharIndex).Speeding = Reader.ReadReal32()
    
    Call MainTimer.SetInterval(TimersIndex.Walk, gIntervals.Walk / charlist(UserCharIndex).Speeding)
    
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

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
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
    frmMain.mapMundo.visible = False
    frmMain.Image5.visible = False
    frmMain.clanimg.visible = False
    frmMain.cmdLlavero.visible = False
    frmMain.QuestBoton.visible = False
    frmMain.ImgSeg.visible = False
    frmMain.ImgSegParty.visible = False
    frmMain.ImgSegClan.visible = False
    frmMain.ImgSegResu.visible = False
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
                Call frmBancoObj.InvBankUsu.SetItem(i, .ObjIndex(i), .Amount(i), .Equipped(i), .GrhIndex(i), .ObjType(i), .MaxHit(i), .MinHit(i), .Def(i), .Valor(i), .ItemName(i), .PuedeUsar(i))
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
                Call frmComerciarUsu.InvUser.SetItem(i, .ObjIndex(i), .Amount(i), .Equipped(i), .GrhIndex(i), .ObjType(i), .MaxHit(i), .MinHit(i), .Def(i), .Valor(i), .ItemName(i), .PuedeUsar(i))
            End With

    Next i
    frmComerciarUsu.lblMyGold.Caption = PonerPuntos(UserStats.GLD)
    
    Dim J As Byte

    For J = 1 To 6
        Call frmComerciarUsu.InvOtherSell.SetItem(J, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0)
        Call frmComerciarUsu.InvUserSell.SetItem(J, 0, 0, 0, 0, 0, 0, 0, 0, 0, "", 0)
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

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    On Error GoTo HandleShowCarpenterForm_Err
        
   ' If frmMain.macrotrabajo.Enabled And (MacroBltIndex > 0) Then
    
        'Call WriteCraftCarpenter(MacroBltIndex)
        
   ' Else
         
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
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
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
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
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
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************

    On Error GoTo HandleNPCKillUser_Err
        
    Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_CRIATURA_MATADO"), 255, 0, 0, True, False, False)
    
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
        
    Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_RECHAZO_ATAQUE_ESCUDO"), 255, 0, 0, True, False, False)
    
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
        
    Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO"), 255, 0, 0, True, False, False)
    
    Exit Sub

HandleBlockedWithShieldOther_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBlockedWithShieldOther", Erl)
    
    
End Sub

''
' Handles the UserSwing message.

Private Sub HandleCharSwing()
    
    On Error GoTo HandleCharSwing_Err
    
    Dim CharIndex As Integer

    CharIndex = Reader.ReadInt16
    
    Dim ShowFX As Boolean

    ShowFX = Reader.ReadBool
    
    Dim ShowText As Boolean

    ShowText = Reader.ReadBool
    
    Dim NotificoTexto As Boolean
    
    NotificoTexto = Reader.ReadBool
        
    With charlist(CharIndex)

        If ShowText And NotificoTexto Then
            Call SetCharacterDialogFx(CharIndex, IIf(CharIndex = UserCharIndex, "Fallas", "Falló"), RGBA_From_Comp(255, 0, 0))
        End If
        
        If EstaPCarea(CharIndex) Then
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
        Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_SEGURO_CLAN_ACTIVADO."), 65, 190, 156, False, False, False)
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
    
    If UserMeditar And UserStats.minman - OldMana > 0 Then

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
    Dim Shield As Long
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
        Call svb_unlock_achivement("Memento Mori")
        UserStats.estado = 1
        charlist(UserCharIndex).Invisible = False
        If MostrarTutorial And tutorial_index <= 0 Then
            If tutorial(e_tutorialIndex.TUTORIAL_Muerto).Activo = 1 Then
                tutorial_index = e_tutorialIndex.TUTORIAL_Muerto
                Call mostrarCartel(tutorial(tutorial_index).titulo, tutorial(tutorial_index).textos(1), tutorial(tutorial_index).grh, -1, &H164B8A, , , False, 100, 479, 100, 535, 640, 530, 50, 100)
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
    If MapData(UserPos.x, UserPos.y).CharIndex = UserCharIndex Then
        MapData(UserPos.x, UserPos.y).CharIndex = 0

    End If
    
    'Set new pos
    UserPos.x = Reader.ReadInt8()
    UserPos.y = Reader.ReadInt8()

    'Set char
    MapData(UserPos.x, UserPos.y).CharIndex = UserCharIndex
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

    Dim CharIndex As Integer
    CharIndex = Reader.ReadInt16()
    
        'Remove char from old position
    If MapData(temp_x, temp_y).CharIndex = CharIndex Then
        MapData(temp_x, temp_y).CharIndex = 0
    End If
    'Set char
    MapData(UserPos.x, UserPos.y).CharIndex = CharIndex
    charlist(CharIndex).Pos = UserPos
        
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

    Dim CharIndex As Integer
    Dim x As Byte, y As Byte
    
    'Set new pos
    CharIndex = Reader.ReadInt16()
    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
    
    If CharIndex = 0 Then Exit Sub
    
    If charlist(CharIndex).Pos.x > 0 And charlist(CharIndex).Pos.y > 0 Then
    
        If MapData(charlist(CharIndex).Pos.x, charlist(CharIndex).Pos.y).CharIndex = CharIndex Then
            MapData(charlist(CharIndex).Pos.x, charlist(CharIndex).Pos.y).CharIndex = 0
        End If
        
        MapData(x, y).CharIndex = CharIndex
        charlist(CharIndex).Pos.x = x
        charlist(CharIndex).Pos.y = y
        charlist(CharIndex).MoveOffsetX = 0
        charlist(CharIndex).MoveOffsetY = 0
    
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
    
    Call svb_unlock_achivement("Small victory")
    
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
            Call AddtoRichTextBox(frmMain.RecTxt, attacker & JsonLanguage.Item("MENSAJE_RECIVE_IMPACTO_CABEZA") & DañoStr & JsonLanguage.Item("MENSAJE_2"), 255, 0, 0, True, False, False)

        Case bBrazoIzquierdo
            Call AddtoRichTextBox(frmMain.RecTxt, attacker & JsonLanguage.Item("MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ") & DañoStr & JsonLanguage.Item("MENSAJE_2"), 255, 0, 0, True, False, False)

        Case bBrazoDerecho
            Call AddtoRichTextBox(frmMain.RecTxt, attacker & JsonLanguage.Item("MENSAJE_RECIVE_IMPACTO_BRAZO_DER") & DañoStr & JsonLanguage.Item("MENSAJE_2"), 255, 0, 0, True, False, False)

        Case bPiernaIzquierda
            Call AddtoRichTextBox(frmMain.RecTxt, attacker & JsonLanguage.Item("MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ") & DañoStr & JsonLanguage.Item("MENSAJE_2"), 255, 0, 0, True, False, False)

        Case bPiernaDerecha
            Call AddtoRichTextBox(frmMain.RecTxt, attacker & JsonLanguage.Item("MENSAJE_RECIVE_IMPACTO_PIERNA_DER") & DañoStr & JsonLanguage.Item("MENSAJE_2"), 255, 0, 0, True, False, False)

        Case bTorso
            Call AddtoRichTextBox(frmMain.RecTxt, attacker & JsonLanguage.Item("MENSAJE_RECIVE_IMPACTO_TORSO") & DañoStr & JsonLanguage.Item("MENSAJE_2"), 255, 0, 0, True, False, False)

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
            Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_PRODUCE_IMPACTO_1") & victim & JsonLanguage.Item("MENSAJE_PRODUCE_IMPACTO_CABEZA") & DañoStr & JsonLanguage.Item("MENSAJE_2"), 255, 0, 0, True, False, False)

        Case bBrazoIzquierdo
            Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_PRODUCE_IMPACTO_1") & victim & JsonLanguage.Item("MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ") & DañoStr & JsonLanguage.Item("MENSAJE_2"), 255, 0, 0, True, False, False)

        Case bBrazoDerecho
            Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_PRODUCE_IMPACTO_1") & victim & JsonLanguage.Item("MENSAJE_PRODUCE_IMPACTO_BRAZO_DER") & DañoStr & JsonLanguage.Item("MENSAJE_2"), 255, 0, 0, True, False, False)

        Case bPiernaIzquierda
            Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_PRODUCE_IMPACTO_1") & victim & JsonLanguage.Item("MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ") & DañoStr & JsonLanguage.Item("MENSAJE_2"), 255, 0, 0, True, False, False)

        Case bPiernaDerecha
            Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_PRODUCE_IMPACTO_1") & victim & JsonLanguage.Item("MENSAJE_PRODUCE_IMPACTO_PIERNA_DER") & DañoStr & JsonLanguage.Item("MENSAJE_2"), 255, 0, 0, True, False, False)

        Case bTorso
            Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_PRODUCE_IMPACTO_1") & victim & JsonLanguage.Item("MENSAJE_PRODUCE_IMPACTO_TORSO") & DañoStr & JsonLanguage.Item("MENSAJE_2"), 255, 0, 0, True, False, False)

    End Select
    
    Exit Sub

HandleUserHittedUser_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUserHittedUser", Erl)
    
    
End Sub

Private Sub HandleChatOverHeadImpl(ByVal chat As String, _
                                    ByVal CharIndex As Integer, _
                                    ByVal Color As Long, _
                                    ByVal EsSpell As Boolean, _
                                    ByVal x As Byte, _
                                    ByVal y As Byte, _
                                    ByVal RequiredMinDisplayTime As Integer, _
                                    ByVal MaxDisplayTime As Integer)
On Error GoTo errhandler
    Dim QueEs      As String
    If x + y > 0 Then
        With charlist(CharIndex)
            If .Invisible And CharIndex <> UserCharIndex Then
                If MapData(.Pos.x, .Pos.y).CharIndex = CharIndex Then MapData(.Pos.x, .Pos.y).CharIndex = 0
                .Pos.x = x
                .Pos.y = y
                MapData(x, y).CharIndex = CharIndex
            End If
        End With
    End If
    
    'Optimizacion de protocolo por Ladder
    QueEs = ReadField(1, chat, Asc("*"))
    Dim TextColor As RGBA
    TextColor = RGBA_From_Long(Color)
    Dim copiar As Boolean

    copiar = True
    
    Dim duracion As Integer

    duracion = 250
    
    Dim Text As String
    Text = ReadField(2, chat, Asc("*"))
    
    Select Case QueEs
        Case "NPCDESC"
            chat = NpcData(Text).desc
            copiar = False
            
            If npcs_en_render And tutorial_index <= 0 Then
                Dim headGrh As Long
                Dim bodyGrh As Long
                headGrh = HeadData(NpcData(Text).Head).Head(3).GrhIndex
                bodyGrh = GrhData(BodyData(NpcData(Text).Body).Walk(3).GrhIndex).Frames(1)
                
                If headGrh = 0 Then
                    Call mostrarCartel(Split(NpcData(Text).Name, " <")(0), NpcData(Text).desc, bodyGrh, 200 + 30 * Len(chat), &H164B8A, , , True, 100, 479, 100, 535, 20, 500, 50, 80, bodyGrh, 1)
                Else
                    Dim HeadOffsetY As Integer
                    HeadOffsetY = CInt(BodyData(NpcData(Text).Body).HeadOffset.y)
                    Call mostrarCartel(Split(NpcData(Text).Name, " <")(0), NpcData(Text).desc, headGrh, 200 + 30 * Len(chat), &H164B8A, , , True, 100, 479, 100, 535, 20, 500, 50, 100, bodyGrh, HeadOffsetY)
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
    If charlist(CharIndex).active = 1 Then
        Call Char_Dialog_Set(CharIndex, chat, Color, duracion, 30, 1, EsSpell, RequiredMinDisplayTime, MaxDisplayTime)
    End If
    
    If charlist(CharIndex).EsNpc = False Then
        If CopiarDialogoAConsola = 1 And copiar Then
            Call WriteChatOverHeadInConsole(CharIndex, chat, TextColor.r, TextColor.G, TextColor.b)
        End If
    End If
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChatOverHeadImpl", Erl)
End Sub

' Handles the ChatOverHead message.
Private Sub HandleLocaleChatOverHead()
    On Error GoTo errhandler
    Dim ChatId       As Integer
    Dim Params As String
    Dim CharIndex  As Integer
    Dim TextColor As Long
    Dim IsSpell    As Boolean
    Dim x As Byte, y As Byte
    Dim MinChatTime As Integer
    Dim MaxChatTime As Integer
    Dim LocalizedText As String
    ChatId = Reader.ReadInt16
    Params = Reader.ReadString8
    CharIndex = Reader.ReadInt16()
    TextColor = vbColor_2_Long(Reader.ReadInt32())
    IsSpell = Reader.ReadBool()
    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
    MinChatTime = Reader.ReadInt16()
    MaxChatTime = Reader.ReadInt16()
    LocalizedText = Locale_Parse_ServerMessage(ChatId, Params)
    Call HandleChatOverHeadImpl(LocalizedText, CharIndex, TextColor, IsSpell, x, y, MinChatTime, MaxChatTime)
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleLocaleChatOverHead", Erl)
End Sub
' Handles the ChatOverHead message.
Private Sub HandleChatOverHead()

    On Error GoTo errhandler

    Dim chat       As String
    Dim CharIndex  As Integer
    Dim colortexto As Long
    Dim QueEs      As String
    Dim EsSpell    As Boolean
    Dim x As Byte, y As Byte
    Dim MinChatTime As Integer
    Dim MaxChatTime As Integer
    chat = Reader.ReadString8()
    CharIndex = Reader.ReadInt16()
    colortexto = vbColor_2_Long(Reader.ReadInt32())
    EsSpell = Reader.ReadBool()
    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
    MinChatTime = Reader.ReadInt16()
    MaxChatTime = Reader.ReadInt16()
    Call HandleChatOverHeadImpl(chat, CharIndex, colortexto, EsSpell, x, y, MinChatTime, MaxChatTime)
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleChatOverHead", Erl)
End Sub

Private Sub HandleTextOverChar()

    On Error GoTo errhandler
    
    Dim chat      As String

    Dim CharIndex As Integer

    Dim Color     As Long
    
    chat = Reader.ReadString8()
    CharIndex = Reader.ReadInt16()
    
    Color = Reader.ReadInt32()
    
    Call SetCharacterDialogFx(CharIndex, chat, RGBA_From_vbColor(Color))

    Exit Sub
    
errhandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleTextOverChar", Erl)
    

End Sub

Private Sub HandleTextOverTile()

    On Error GoTo errhandler
    
    Dim Text  As String
    Dim x As Integer
    Dim y As Integer
    Dim Color As Long
    Dim duration As Integer
    Dim OffsetY As Integer
    Dim Animated As Boolean
    Text = Reader.ReadString8()
    x = Reader.ReadInt16()
    y = Reader.ReadInt16()
    Color = Reader.ReadInt32()
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
                    If .DialogEffects(Index).Text = vbNullString Then
                        Exit For
                    End If
                Next
                If Index > UBound(.DialogEffects) Then
                    ReDim Preserve .DialogEffects(1 To UBound(.DialogEffects) + 1)
                End If
            End If
            With .DialogEffects(Index)
                .Color = RGBA_From_vbColor(Color)
                .Start = FrameTime
                .Text = Text
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
    
    Dim Text      As String
    Dim CharIndex As Integer
    Dim Color     As Long
    Dim duration As Integer
    Dim Animated As Boolean
    
    Text = Reader.ReadString8()
    CharIndex = Reader.ReadInt16()
    Color = Reader.ReadInt32()
    duration = Reader.ReadInt16()
    Animated = Reader.ReadBool()
    If CharIndex = 0 Then Exit Sub
    Dim x As Integer, y As Integer, OffsetX As Integer, OffsetY As Integer
    
    With charlist(CharIndex)
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
    Dim Text      As String
    Dim Color     As Long
    Dim SourceName As String
    Dim SourceStatus As Integer
    Dim Privileges As Integer
    Text = Reader.ReadString8()
    Color = Reader.ReadInt32()
    SourceName = Reader.ReadString8()
    SourceStatus = Reader.ReadInt16()
    Privileges = Reader.ReadInt16()
    If Privileges > 0 Then
        Privileges = Log(Privileges) / Log(2)
    End If
    Dim TextColor As RGBA
    TextColor = RGBA_From_vbColor(Color)
    
    Call WriteConsoleUserChat(Text, SourceName, TextColor.r, TextColor.G, TextColor.b, SourceStatus, Privileges)
End Sub
''
' Handles the ConsoleMessage message.

Private Sub HandleConsoleMessage()
    
    On Error GoTo errhandler
    
    Dim chat      As String
    Dim FontIndex As Integer
    Dim str       As String
    Dim r         As Byte
    Dim G         As Byte
    Dim b         As Byte
    Dim QueEs     As String
    Dim NpcName   As String
    Dim objname   As String
    Dim Hechizo   As Integer
    Dim userName  As String
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
            If language = Spanish Then
                Hechizo = ReadField(2, chat, Asc("*"))
                chat = "------------< Información del hechizo >------------" & vbCrLf & "Nombre: " & HechizoData(Hechizo).nombre & vbCrLf & "Descripción: " & HechizoData(Hechizo).desc & vbCrLf & "Skill requerido: " & HechizoData(Hechizo).MinSkill & " de magia." & vbCrLf & "Mana necesario: " & HechizoData(Hechizo).ManaRequerido & " puntos." & vbCrLf & "Stamina necesaria: " & HechizoData(Hechizo).StaRequerido & " puntos."
            Else
                Hechizo = ReadField(2, chat, Asc("*"))
                chat = "------------< Spell information >------------" & vbCrLf & "Name: " & HechizoData(Hechizo).nombre & vbCrLf & "Description: " & HechizoData(Hechizo).desc & vbCrLf & "Required skill: " & HechizoData(Hechizo).MinSkill & " of magic." & vbCrLf & "Mana needed: " & HechizoData(Hechizo).ManaRequerido & " points." & vbCrLf & "Stamina needed: " & HechizoData(Hechizo).StaRequerido & " points."
            End If
            
            Case "ProMSG"
            If language = Spanish Then
                Hechizo = ReadField(2, chat, Asc("*"))
                chat = HechizoData(Hechizo).PropioMsg
            Else
                Hechizo = ReadField(2, chat, Asc("*"))
                chat = HechizoData(Hechizo).en_PropioMsg
            End If
    
            Case "HecMSG"
            If language = Spanish Then
                Hechizo = ReadField(2, chat, Asc("*"))
                chat = HechizoData(Hechizo).HechizeroMsg & " la criatura."
            Else
                Hechizo = ReadField(2, chat, Asc("*"))
                chat = HechizoData(Hechizo).HechizeroMsg & " the creature."
            End If
    
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
            b = 255
        Else
            b = Val(str)

        End If
            
        Call AddtoRichTextBox(frmMain.RecTxt, Left$(chat, InStr(1, chat, "~") - 1), r, G, b, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
    
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

Private Sub HandleLocaleMsg()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errhandler
    
    Dim chat      As String

    Dim FontIndex As Integer

    Dim str       As String

    Dim r         As Byte

    Dim G         As Byte

    Dim b         As Byte

    Dim QueEs     As String

    Dim NpcName   As String

    Dim objname   As String

    Dim Hechizo   As Byte

    Dim userName  As String

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
            b = 255
        Else
            b = Val(str)

        End If
            
        Call AddtoRichTextBox(frmMain.RecTxt, Left$(chat, InStr(1, chat, "~") - 1), r, G, b, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
    Else

        With FontTypes(FontIndex)
            Call AddtoRichTextBox(frmMain.RecTxt, chat, .red, .green, .blue, .bold, .italic)

        End With

    End If
    
    Exit Sub

errhandler:

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
    
    On Error GoTo errhandler
    
    Dim chat As String

    Dim status As Byte
    
    Dim str  As String

    Dim r    As Byte

    Dim G    As Byte

    Dim b    As Byte

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
                b = 255
            Else
                b = Val(str)
            End If
                
            Call AddtoRichTextBox(frmMain.RecTxt, Left$(chat, InStr(1, chat, "~") - 1), r, G, b, Val(ReadField(5, chat, 126)) <> 0, Val(ReadField(6, chat, 126)) <> 0)
        Else
            With FontTypes(FontTypeNames.FONTTYPE_GUILDMSG)
                Call AddtoRichTextBox(frmMain.RecTxt, chat, .red, .green, .blue, .bold, .italic)
            End With
        End If
    Else
        Call DialogosClanes.PushBackText(ReadField(1, chat, 126), status)
    End If
    
    Exit Sub

errhandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleGuildChat", Erl)
    

End Sub

Private Sub HandleShowMessageBox()
On Error GoTo errhandler
    
    Dim mensaje As String

    mensaje = Reader.ReadString8()

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

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errhandler
    
    Dim CharIndex     As Integer
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
    Dim AuraParticula As Byte
    Dim ParticulaFx   As Byte
    Dim appear        As Byte
    Dim group_index   As Integer
    
    CharIndex = Reader.ReadInt16()

    Body = Reader.ReadInt16()
    Head = Reader.ReadInt16()
    Heading = Reader.ReadInt8()
    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
    weapon = Reader.ReadInt16()
    Shield = Reader.ReadInt16()
    helmet = Reader.ReadInt16()
    Cart = Reader.ReadInt16()
    
    With charlist(CharIndex)
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
        
          
        If (.Pos.x <> 0 And .Pos.y <> 0) Then
            If MapData(.Pos.x, .Pos.y).CharIndex = CharIndex Then
                'Erase the old character from map
                MapData(charlist(CharIndex).Pos.x, charlist(CharIndex).Pos.y).CharIndex = 0

            End If

        End If

        If privs <> 0 Then
            'Log2 of the bit flags sent by the server gives our numbers ^^
            .priv = Log(privs) / Log(2)
        Else
            .priv = 0

        End If

        .Muerto = (Body = CASPER_BODY_IDLE)
        Call MakeChar(CharIndex, Body, Head, Heading, x, y, weapon, Shield, helmet, Cart, ParticulaFx, appear)
        
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

    Dim CharIndex As Integer
    Dim flag As Long
    
    
    CharIndex = Reader.ReadInt16()
    flag = Reader.ReadInt8()
    
    With charlist(CharIndex)
        .banderaIndex = flag
    End With
    

    Exit Sub
    
errhandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleTextOverChar", Erl)
    

End Sub

Private Sub HandleCharacterRemove()
On Error GoTo HandleCharacterRemove_Err
    Dim CharIndex   As Integer
    Dim Desvanecido As Boolean
    Dim fueWarp As Boolean
    
    CharIndex = Reader.ReadInt16()
    Desvanecido = Reader.ReadBool()
    fueWarp = Reader.ReadBool()
    
    If Desvanecido And charlist(CharIndex).EsNpc = True Then
        Call CrearFantasma(CharIndex)
    End If

    Call EraseChar(CharIndex, fueWarp)
    Call RefreshAllChars
    Call ao20audio.StopAllWavsMatchingLabel("meditate" & CStr(CharIndex))
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
    
    Dim CharIndex As Integer
    Dim x         As Byte
    Dim y         As Byte
    Dim dir       As Byte
    CharIndex = Reader.ReadInt16()
    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
    
    Call Char_Move_by_Pos(CharIndex, x, y)
    Call RefreshAllChars
    
    With charlist(CharIndex)
        ' Play steps sounds if the user is not an admin of any kind
        If .priv <> 1 And .priv <> 2 And .priv <> 3 And .priv <> 5 And .priv <> 25 Then
            Call DoPasosFx(CharIndex)
        End If
    End With

    Exit Sub

HandleCharacterMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCharacterMove", Erl)
End Sub

Private Sub HandleCharacterTranslate()
On Error GoTo HandleCharacterTranslate_Err
    Dim CharIndex As Integer
    Dim TileX     As Byte
    Dim TileY     As Byte
    Dim TranslationTime As Long
    CharIndex = Reader.ReadInt16()
    TileX = Reader.ReadInt8()
    TileY = Reader.ReadInt8()
    TranslationTime = Reader.ReadInt32()
    Call TranslateCharacterToPos(CharIndex, TileX, TileY, TranslationTime)
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

''
' Handles the ForceCharMove message.

Private Sub HandleForceCharMoveSiguiendo()
    
    On Error GoTo HandleForceCharMoveSiguiendo_Err
    
    Dim direccion As Byte
    direccion = Reader.ReadInt8()
    Moviendose = True
    
    Call MainTimer.Restart(TimersIndex.Walk)
    'Capaz hay que eliminar el char_move_by_head
    
    UserPos.x = charlist(CharindexSeguido).Pos.x
    UserPos.y = charlist(CharindexSeguido).Pos.y

    Call UpdateMapPos
    
    If MapDat.Seguro = 1 Then
        frmMain.Coord.ForeColor = RGB(0, 170, 0)
    Else
        frmMain.Coord.ForeColor = RGB(170, 0, 0)
    End If
    
    Call Char_Move_by_Head(CharindexSeguido, direccion)
    Call MoveScreen(direccion)
    'Call RefreshAllChars
    
    Exit Sub

HandleForceCharMoveSiguiendo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleForceCharMoveSiguiendo", Erl)
    
    
End Sub

''
' Handles the CharacterChange message.

Private Sub HandleCharacterChange()
    
    On Error GoTo HandleCharacterChange_Err
    
    Dim CharIndex As Integer

    Dim TempInt   As Integer

    Dim headIndex As Integer

    CharIndex = Reader.ReadInt16()

    With charlist(CharIndex)
        TempInt = Reader.ReadInt16()

        If TempInt < LBound(BodyData()) Or TempInt > UBound(BodyData()) Then
            .Body = BodyData(0)
        Else
            .Body = BodyData(TempInt)
            .iBody = TempInt

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
        
        TempInt = Reader.ReadInt16()

        If TempInt <> 0 And TempInt <= UBound(WeaponAnimData) Then
            .Arma = WeaponAnimData(TempInt)
        End If

        TempInt = Reader.ReadInt16()

        If TempInt <> 0 And TempInt <= UBound(ShieldAnimData) Then
            .Escudo = ShieldAnimData(TempInt)
        End If
        
        TempInt = Reader.ReadInt16()

        If TempInt <> 0 And TempInt <= UBound(CascoAnimData) Then
            .Casco = CascoAnimData(TempInt)
        End If
        
        TempInt = Reader.ReadInt16()
        
        If TempInt <= 2 Or TempInt > UBound(BodyData()) Then
            .HasCart = False
        Else
            .Cart = BodyData(TempInt)
            .HasCart = True
        End If
                
        If .Body.HeadOffset.y = -26 Then
            .EsEnano = True
        Else
            .EsEnano = False

        End If
        
        Dim Fx As Integer: Fx = Reader.ReadInt16
        Call StartFx(.ActiveAnimation, Fx)
        
        .Meditating = Fx <> 0
        
        Reader.ReadInt16 'Ignore loops
        
        Dim flags As Byte
        
        flags = Reader.ReadInt8()
        
        .Idle = flags And &O1
        .Navegando = flags And &O2
        
        If .Idle Then
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
    
    Dim x As Byte, y As Byte, b As Byte
    
    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
    b = Reader.ReadInt8()

    MapData(x, y).Blocked = MapData(x, y).Blocked And Not eBlock.ALL_SIDES
    MapData(x, y).Blocked = MapData(x, y).Blocked Or b
    
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

Private Sub HandlePlayWave()
On Error GoTo HandlePlayWave_Err
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
        Call ao20audio.StopWav(CStr(wave))
        If cancelLastWave = 2 Then Exit Sub
    End If
    
    If srcX = 0 Or srcY = 0 Then
        Call ao20audio.PlayWav(CStr(wave), False, 0, 0)
    Else
        If EstaEnArea(srcX, srcY) Then
            Dim p As Position
            p.x = srcX
            p.y = srcY
            Call ao20audio.PlayWav(CStr(wave), False, ao20audio.ComputeCharFxVolume(p), ao20audio.ComputeCharFxPan(p))
        End If
    End If
    Exit Sub
HandlePlayWave_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandlePlayWave", Erl)
End Sub

''
' Handles the PlayWave message.

Private Sub HandlePlayWaveStep()
    
    On Error GoTo HandlePlayWaveStep_Err
    
    Dim CharIndex As Integer
    Dim grh As Long
    Dim Grh2 As Long
    Dim distance As Byte
    Dim balance As Integer
    Dim step As Boolean
    
100 CharIndex = Reader.ReadInt16()
102 grh = Reader.ReadInt32()
104 Grh2 = Reader.ReadInt32()
106 distance = Reader.ReadInt8()
108 balance = Reader.ReadInt16()
110 step = Reader.ReadBool()
    
112 Call DoPasosInvi(grh, Grh2, distance, balance, step)

    With charlist(CharIndex)
        ' Esta invisible, lo sacamos del mapa para que no tosquee
114     If MapData(.Pos.x, .Pos.y).CharIndex = CharIndex Then
116         MapData(.Pos.x, .Pos.y).CharIndex = 0
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
    Dim CharIndex As Integer
    Dim MinHp     As Long
    Dim MaxHp     As Long
    Dim Shield As Long
    CharIndex = Reader.ReadInt16()
    MinHp = Reader.ReadInt32()
    MaxHp = Reader.ReadInt32()
    Shield = Reader.ReadInt32()
    If Group.GroupSize > 0 Then
        Dim i As Integer
        For i = 0 To Group.GroupSize - 1
            If Group.GroupMembers(i).CharIndex = CharIndex Then
                Group.GroupMembers(i).MinHp = MinHp
                Group.GroupMembers(i).MaxHp = MaxHp
                Group.GroupMembers(i).Shield = Shield
            End If
        Next i
    End If
    charlist(CharIndex).UserMinHp = MinHp
    charlist(CharIndex).UserMaxHp = MaxHp
    charlist(CharIndex).Shield = Shield
    Exit Sub

HandleCharUpdateHP_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCharUpdateHP", Erl)
End Sub

Private Sub HandleCharUpdateMAN()
    
    On Error GoTo HandleCharUpdateHP_Err

    Dim CharIndex As Integer

    Dim minman     As Long

    Dim maxman     As Long
    
    CharIndex = Reader.ReadInt16()
    minman = Reader.ReadInt32()
    maxman = Reader.ReadInt32()

    charlist(CharIndex).UserMinMAN = minman
    charlist(CharIndex).UserMaxMAN = maxman
    
    Exit Sub

HandleCharUpdateHP_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCharUpdateMAN", Erl)
    
    
End Sub

Private Sub HandleArmaMov()
    
    On Error GoTo HandleArmaMov_Err

    '***************************************************

    Dim CharIndex As Integer
    Dim isRanged As Byte
    CharIndex = Reader.ReadInt16()
    isRanged = Reader.ReadInt8()

    With charlist(CharIndex)

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

    '***************************************************
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
        Group.GroupMembers(i).CharIndex = Reader.ReadInt16
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

Private Sub HandleRequestTelemetry()
    On Error GoTo HandleUpdateGroupInfo_Err
    Exit Sub
HandleUpdateGroupInfo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateGroupInfo", Erl)
End Sub

Private Sub HandleUpdateCharValue()
On Error GoTo HandleUpdateGroupInfo_Err
    Dim CharIndex As Integer
    Dim CharValueType As Integer
    Dim Value As Long
    CharIndex = Reader.ReadInt16
    CharValueType = Reader.ReadInt16
    Value = Reader.ReadInt32
    Select Case CharValueType
        Case e_CharValue.eDontBlockTile
            charlist(CharIndex).DontBlockTile = Value
    End Select
    Exit Sub
HandleUpdateGroupInfo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleUpdateGroupInfo", Erl)
End Sub

Private Sub HandleStunStart()
On Error GoTo HandleStunStart_Err
    If EstaSiguiendo Then Exit Sub
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

    '***************************************************

    Dim CharIndex As Integer

    CharIndex = Reader.ReadInt16()

    With charlist(CharIndex)

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
    Call frmGuildAdm.Show(vbModeless, GetGameplayForm())
    Exit Sub

errhandler:

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

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim CharIndex As Integer

    Dim Fx        As Integer

    Dim Loops     As Integer
    
    Dim x As Byte, y As Byte
        
    CharIndex = Reader.ReadInt16()
    Fx = Reader.ReadInt16()
    Loops = Reader.ReadInt16()
    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
    
    If x + y > 0 Then
        With charlist(CharIndex)
            If .Invisible And CharIndex <> UserCharIndex Then
                If MapData(.Pos.x, .Pos.y).CharIndex = CharIndex Then MapData(.Pos.x, .Pos.y).CharIndex = 0
                .Pos.x = x
                .Pos.y = y
                MapData(x, y).CharIndex = CharIndex
            End If
        End With
    End If
    
    Call SetCharacterFx(CharIndex, Fx, Loops)
    
    Exit Sub

HandleCreateFX_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCreateFX", Erl)
    
    
End Sub

''
' Handles the CharAtaca message.

Private Sub HandleCharAtaca()
    
    On Error GoTo HandleCharAtaca_Err
    
    Dim NpcIndex As Integer
    Dim VictimIndex As Integer
    Dim danio     As Long
    Dim AnimAttack As Integer
    
    NpcIndex = Reader.ReadInt16()
    VictimIndex = Reader.ReadInt16()
    danio = Reader.ReadInt32()
    AnimAttack = Reader.ReadInt16()
    
    Dim grh As grh
    With charlist(NpcIndex)
        If AnimAttack > 0 Then
            .Body = BodyData(AnimAttack)
            .Body.Walk(.Heading).started = FrameTime
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
''
' Handles the UpdateUserStats message.
''
' Handles the CharAtaca message.

Private Sub HandleRecievePosSeguimiento()
    
    On Error GoTo RecievePosSeguimiento_Err
    
    Dim PosX As Integer
    Dim PosY As Integer
    
    PosX = Reader.ReadInt16()
    PosY = Reader.ReadInt16()
    
    frmMain.shapexy.Left = PosX - 6

    Exit Sub
    

RecievePosSeguimiento_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.RecievePosSeguimiento", Erl)
    
    
End Sub

Private Sub HandleCancelarSeguimiento()
    
    On Error GoTo CancelarSeguimiento_Err
    
    frmMain.shapexy.Left = 1200
    frmMain.shapexy.Top = 1200
    CharindexSeguido = 0
    OffsetLimitScreen = 32
    Exit Sub
    

CancelarSeguimiento_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.CancelarSeguimiento", Erl)
    
End Sub

Private Sub HandleGetInventarioHechizos()
    
    On Error GoTo GetInventarioHechizos_Err
    
    Dim inventario_o_hechizos As Byte
    Dim hechiSel As Byte
    Dim scrollSel As Byte
    
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
    
    Dim Value As Byte
    
    Value = Reader.ReadInt8()

    'Clickeó en inventario
    If Value = 1 Then
        frmMain.shapexy.BackColor = RGB(0, 170, 0)
    'Clickeó en hechizos
    Else
        frmMain.shapexy.BackColor = RGB(170, 0, 0)
    End If

    Exit Sub
    

NotificarClienteCasteo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.NotificarClienteCasteo", Erl)
    
End Sub



Private Sub HandleSendFollowingCharindex()
    
    On Error GoTo SendFollowingCharindex_Err
    
    Dim CharIndex As Integer
    CharIndex = Reader.ReadInt16()
    UserCharIndex = CharIndex
    CharindexSeguido = CharIndex
    OffsetLimitScreen = 31
    Exit Sub
    

SendFollowingCharindex_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.SendFollowingCharindex", Erl)
    
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
    If EstaSiguiendo Then Exit Sub
    If UsingSkillREcibido = 0 Then
        Frm.MousePointer = 0
        Call FormParser.Parse_Form(frmMain, E_NORMAL)
        UsingSkill = UsingSkillREcibido
        Exit Sub
    End If

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
    
    On Error GoTo errhandler
    
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
    Dim IsBindable As Boolean

    Slot = Reader.ReadInt8()
    ObjIndex = Reader.ReadInt16()
    Amount = Reader.ReadInt16()
    Equipped = Reader.ReadBool()
    Value = Reader.ReadReal32()
    podrausarlo = Reader.ReadInt8()
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
    
    Call ModGameplayUI.SetInvItem(Slot, ObjIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, MinDef, Value, Name, podrausarlo, IsBindable)
    
    If frmComerciar.visible Then
        Call frmComerciar.InvComUsu.SetItem(Slot, ObjIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, MinDef, Value, Name, podrausarlo)

    ElseIf frmBancoObj.visible Then
        Call frmBancoObj.InvBankUsu.SetItem(Slot, ObjIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, MinDef, Value, Name, podrausarlo)
        
    ElseIf frmBancoCuenta.visible Then
        Call frmBancoCuenta.InvBankUsuCuenta.SetItem(Slot, ObjIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, MinDef, Value, Name, podrausarlo)
    
    ElseIf frmCrafteo.visible Then
        Call frmCrafteo.InvCraftUser.SetItem(Slot, ObjIndex, Amount, Equipped, GrhIndex, ObjType, MaxHit, MinHit, MinDef, Value, Name, podrausarlo)
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
    
    Dim Slot As Byte
    Dim BankSlot As Slot
    
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

errhandler:

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
    
    On Error GoTo errhandler
    
    Dim Slot     As Byte
    Dim Index    As Integer
    Dim Cooldown As Integer
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
            hlst.List(Slot - 1) = "(Vacio)"
        Else
            Call hlst.AddItem("(Vacio)")
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
        ' tmp = ObjData(ArmasHerrero(i).Index).name        'Get the object's name
        ArmasHerrero(i).LHierro = Reader.ReadInt16()  'The iron needed
        ArmasHerrero(i).LPlata = Reader.ReadInt16()    'The silver needed
        ArmasHerrero(i).LOro = Reader.ReadInt16()    'The gold needed
        ArmasHerrero(i).Coal = Reader.ReadInt16()   'The coal needed
        ' Call frmHerrero.lstArmas.AddItem(tmp)
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
        tmp = Reader.ReadString8()         'Get the object's name
        DefensasHerrero(i).LHierro = Reader.ReadInt16()   'The iron needed
        DefensasHerrero(i).LPlata = Reader.ReadInt16()   'The silver needed
        DefensasHerrero(i).LOro = Reader.ReadInt16()   'The gold needed
        DefensasHerrero(i).Coal = Reader.ReadInt16()   'The coal needed
        ' Call frmHerrero.lstArmaduras.AddItem(tmp)
        DefensasHerrero(i).Index = Reader.ReadInt16()
    Next i
        
    Dim a      As Byte
    Dim e      As Byte
    Dim c      As Byte
    Dim tmpObj As ObjDatas

    a = 0
    e = 0
    c = 0
    
    For i = 1 To UBound(DefensasHerrero())

        If DefensasHerrero(i).Index = 0 Then Exit For
        
        tmpObj = ObjData(DefensasHerrero(i).Index)
        
        If tmpObj.ObjType = 3 Then
           
            ArmadurasHerrero(a).Index = DefensasHerrero(i).Index
            ArmadurasHerrero(a).LHierro = DefensasHerrero(i).LHierro
            ArmadurasHerrero(a).LPlata = DefensasHerrero(i).LPlata
            ArmadurasHerrero(a).LOro = DefensasHerrero(i).LOro
            a = a + 1

        End If
        
        ' Escudos (16), Objetos Magicos (21) y Anillos (35) van en la misma lista
        If tmpObj.ObjType = 16 Or tmpObj.ObjType = 35 Or tmpObj.ObjType = 21 Or tmpObj.ObjType = 100 Or tmpObj.ObjType = 30 Then
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

errhandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBlacksmithArmors", Erl)
    

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
            SastreRopas(G).PielLoboNegro = ObjSastre(i).PielLoboNegro
            r = r + 1

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
    
    On Error GoTo errhandler
    
    Dim tmp As String
    Dim grh As Integer

    tmp = ObjData(Reader.ReadInt16()).Texto
    grh = Reader.ReadInt16()
    
    Call InitCartel(tmp, grh)
    
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
        .PuedeUsar = Reader.ReadInt8()
        
        Call frmComerciar.InvComNpc.SetItem(Slot, .ObjIndex, .Amount, 0, .GrhIndex, .ObjType, .MaxHit, .MinHit, .Def, .Valor, .Name, .PuedeUsar)
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

    Dim Color As Long, duracion As Long, ignorar As Boolean
    
    Color = Reader.ReadInt32()
    duracion = Reader.ReadInt32()
    ignorar = Reader.ReadBool()
    
    Dim r, G, b As Byte

    b = (Color And 16711680) / 65536
    G = (Color And 65280) / 256
    r = Color And 255
    Color = D3DColorARGB(255, r, G, b)

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
        frmStatistics.Show , GetGameplayForm()
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
    Call svb_unlock_achivement("Newbie's fate")
    
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

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    Dim CharIndex As Integer
    Dim x As Byte, y As Byte
    CharIndex = Reader.ReadInt16()
    charlist(CharIndex).Invisible = Reader.ReadBool()
    charlist(CharIndex).TimerI = 0
    
    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
    
    If x + y > 0 Then
        With charlist(CharIndex)
            If CharIndex <> UserCharIndex Then
                If .Invisible Then
                    If Not IsCharVisible(CharIndex) And General_Distance_Get(x, y, UserPos.x, UserPos.y) > DISTANCIA_ENVIO_DATOS Then
                        If .clan_index > 0 Then
                            If .clan_index = charlist(UserCharIndex).clan_index And CharIndex <> UserCharIndex And .Muerto = 0 Then
                                If .clan_nivel >= 6 Then
                                    Exit Sub
                                End If
                            End If
                        End If
                        If .Meditating Then Exit Sub
                        If MapData(.Pos.x, .Pos.y).CharIndex = CharIndex Then MapData(.Pos.x, .Pos.y).CharIndex = 0
                        .MoveOffsetX = 0
                        .MoveOffsetY = 0
                    End If
                Else
                    If MapData(.Pos.x, .Pos.y).CharIndex = CharIndex Then MapData(.Pos.x, .Pos.y).CharIndex = 0
                    .Pos.x = x
                    .Pos.y = y
                    MapData(x, y).CharIndex = CharIndex
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
    
    Dim CharIndex As Integer, Fx As Integer
    Dim x As Byte, y As Byte
    
    CharIndex = Reader.ReadInt16
    Fx = Reader.ReadInt16
    x = Reader.ReadInt8
    y = Reader.ReadInt8
    
    charlist(CharIndex).Meditating = Fx <> 0
    
    If x + y > 0 Then
        With charlist(CharIndex)
            If .Invisible And CharIndex <> UserCharIndex Then
                If MapData(.Pos.x, .Pos.y).CharIndex = CharIndex Then MapData(.Pos.x, .Pos.y).CharIndex = 0
                .Pos.x = x
                .Pos.y = y
                MapData(x, y).CharIndex = CharIndex
            End If
        End With
    End If
    
    If CharIndex = UserCharIndex Then
        UserMeditar = (Fx <> 0)
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
    
    With charlist(CharIndex)
        If Fx <> 0 Then
            Call StartFx(.ActiveAnimation, Fx, -1)
            ' Play sound only in PC area
            If EstaPCarea(CharIndex) Then
                Call ao20audio.PlayWav(SND_MEDITATE, True, ao20audio.ComputeCharFxVolume(.Pos), ao20audio.ComputeCharFxPan(.Pos), "meditate" & CStr(CharIndex))
            End If
        Else
            Call ao20audio.StopWav(SND_MEDITATE, "meditate" & CStr(CharIndex))
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

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
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

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
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

        'Debug.Print guildList(i)
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
                .beneficios = "Max miembros: 5"

            Case 2
                .beneficios = "Pedir ayuda (G)" & vbCrLf & "Max miembros: 8"

            Case 3
                .beneficios = "Pedir ayuda (G)" & vbCrLf & "Seguro de clan" & vbCrLf & "Max miembros: 11"

            Case 4
                .beneficios = "Pedir ayuda (G)" & vbCrLf & "Seguro de clan" & vbCrLf & "Max miembros: 14"

            Case 5
                .beneficios = "Pedir ayuda (G)" & vbCrLf & "Seguro de clan" & vbCrLf & "Ver vida y mana" & vbCrLf & " Max miembros: 17"
                
            Case 6
                .beneficios = "Pedir ayuda (G)" & vbCrLf & "Seguro de clan" & vbCrLf & "Ver vida y mana" & vbCrLf & "Verse invisible" & vbCrLf & " Max miembros: 20"
        
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

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errhandler
    
    Call frmUserRequest.recievePeticion(Reader.ReadString8())
    
    Exit Sub

errhandler:

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

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
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

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
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
            .Genero.Caption = "Genero: Hombre"
        Else
            .Genero.Caption = "Genero: Mujer"
        End If
            
        .nombre.Caption = "Nombre: " & Reader.ReadString8()
        .Raza.Caption = "Raza: " & ListaRazas(Reader.ReadInt8())
        .Clase.Caption = "Clase: " & ListaClases(Reader.ReadInt8())

        .Nivel.Caption = "Nivel: " & Reader.ReadInt8()
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
    
        Call .Show(vbModeless, GetGameplayForm())
    
    End With
        
    Exit Sub
    
errhandler:
    
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
    
    On Error GoTo errhandler
    
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

        Dim Nivel  As Byte
         
        Nivel = Reader.ReadInt8()
        .Nivel = "Nivel: " & Nivel
        
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
        Padding = Space$(19)

        Select Case Nivel

               Case 1
                .beneficios = Padding & "Max miembros: 5"
                .maxMiembros = 5
            Case 2
                .beneficios = Padding & "Pedir ayuda (G) / Max miembros: 7"
                .maxMiembros = 7

            Case 3
                .beneficios = Padding & "Pedir ayuda (G) / Seguro de clan." & vbCrLf & "Max miembros: 7"
                .maxMiembros = 7

            Case 4
                .beneficios = Padding & "Pedir ayuda (G) / Seguro de clan. " & vbCrLf & "Max miembros: 12"
                .maxMiembros = 12

            Case 5
                .beneficios = Padding & "Pedir ayuda (G) / Seguro de clan /  Ver vida y mana." & vbCrLf & "Max miembros: 15"
                .maxMiembros = 15
                
            Case 6
                .beneficios = Padding & "Pedir ayuda (G) / Seguro de clan / Ver vida y mana / Verse invisible." & vbCrLf & "Max miembros: 20"
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

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
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
        
        .desc.Text = GuildDetails.Description
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
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
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
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    'Remove packet ID
    
    On Error GoTo HandleParalizeOK_Err
    If EstaSiguiendo Then Exit Sub
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
    If EstaSiguiendo Then Exit Sub
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

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errhandler
    
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
        frmComerciarUsu.lblMyGold.Caption = PonerPuntos(UserStats.GLD - OroAEnviar)

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
    
    frmComerciarUsu.lblEstadoResp.visible = False
    
    Exit Sub

errhandler:

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

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errhandler
    
    Dim sosList()      As String

    Dim i              As Long

    Dim nombre         As String

    Dim Consulta       As String

    Dim TipoDeConsulta As String
    
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

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errhandler
    
    frmCambiaMotd.txtMotd.Text = Reader.ReadString8()
    frmCambiaMotd.Show , GetGameplayForm()
    
    Exit Sub

errhandler:

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
    
    Select Case MiCargo ' ReyarB ajustar privilejios
    
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
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
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

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
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

''
' Handles the UpdateTag message.

Private Sub HandleUpdateTagAndStatus()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errhandler
    
    Dim CharIndex   As Integer

    Dim status      As Byte

    Dim NombreYClan As String

    Dim group_index As Integer
    
    CharIndex = Reader.ReadInt16()
    status = Reader.ReadInt8()
    NombreYClan = Reader.ReadString8()
        
    Dim Pos As Integer
    Pos = InStr(NombreYClan, "<")

    If Pos = 0 Then Pos = InStr(NombreYClan, "[")
    If Pos = 0 Then Pos = Len(NombreYClan) + 2
    
    charlist(CharIndex).nombre = Left$(NombreYClan, Pos - 2)
    charlist(CharIndex).clan = mid$(NombreYClan, Pos)
    
    group_index = Reader.ReadInt16()
    
    'Update char status adn tag!
    charlist(CharIndex).status = status
    
    charlist(CharIndex).group_index = group_index
    
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
            Call General_Char_Particle_Create(ParticulaIndex, MapData(x, y).CharIndex, Time)

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
    
    Dim CharIndex      As Integer

    Dim ParticulaIndex As Integer

    Dim Time           As Long

    Dim Remove         As Boolean
    Dim grh            As Long
    
    Dim x As Byte, y As Byte
    
    CharIndex = Reader.ReadInt16()
    ParticulaIndex = Reader.ReadInt16()
    Time = Reader.ReadInt32()
    Remove = Reader.ReadBool()
    grh = Reader.ReadInt32()
    
    x = Reader.ReadInt8()
    y = Reader.ReadInt8()
    
    If x + y > 0 Then
        With charlist(CharIndex)
            If .Invisible And CharIndex <> UserCharIndex Then
                If MapData(.Pos.x, .Pos.y).CharIndex = CharIndex Then MapData(.Pos.x, .Pos.y).CharIndex = 0
                .Pos.x = x
                .Pos.y = y
                MapData(x, y).CharIndex = CharIndex
            End If
        End With
    End If
    If Remove Then
        Call Char_Particle_Group_Remove(CharIndex, ParticulaIndex)
        charlist(CharIndex).Particula = 0
    
    Else
        charlist(CharIndex).Particula = ParticulaIndex
        charlist(CharIndex).ParticulaTime = Time
        If grh > 0 Then
            Call General_Char_Particle_Create(ParticulaIndex, CharIndex, Time, grh)
        Else
            Call General_Char_Particle_Create(ParticulaIndex, CharIndex, Time)
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

    Dim Fx             As Integer
    
    Dim x As Byte, y As Byte
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
                If MapData(.Pos.x, .Pos.y).CharIndex = receptor Then MapData(.Pos.x, .Pos.y).CharIndex = 0
                .Pos.x = x
                .Pos.y = y
                MapData(x, y).CharIndex = receptor
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
    
    ' Debug.Print "RECIBI FX= " & fX

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

        '***************************************************
        'Author: Juan Martín Sotuyo Dodero (Maraxus)
        'Last Modification: 05/17/0
        '
        '***************************************************
    
        Dim CharIndex      As Integer

        Dim ParticulaIndex As String

        Dim Remove         As Boolean

        Dim TIPO           As Byte
     
100     CharIndex = Reader.ReadInt16()
102     ParticulaIndex = Reader.ReadString8()

104     Remove = Reader.ReadBool()
106     TIPO = Reader.ReadInt8()
    
108     If TIPO = 1 Then
110         charlist(CharIndex).Arma_Aura = ParticulaIndex
112     ElseIf TIPO = 2 Then
114         charlist(CharIndex).Body_Aura = ParticulaIndex
116     ElseIf TIPO = 3 Then
118         charlist(CharIndex).Escudo_Aura = ParticulaIndex
120     ElseIf TIPO = 4 Then
122         charlist(CharIndex).Head_Aura = ParticulaIndex
124     ElseIf TIPO = 5 Then
126         charlist(CharIndex).Otra_Aura = ParticulaIndex
128     ElseIf TIPO = 6 Then
130         charlist(CharIndex).DM_Aura = ParticulaIndex
        Else
132         charlist(CharIndex).RM_Aura = ParticulaIndex

        End If
    
        Exit Sub

HandleAuraToChar_Err:
134     Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleAuraToChar " & CharIndex, Erl)
    
    
End Sub

Private Sub HandleSpeedToChar()
    
    On Error GoTo HandleSpeedToChar_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/0
    '
    '***************************************************
    
    Dim CharIndex As Integer

    Dim Speeding  As Single
     
    CharIndex = Reader.ReadInt16()
    Speeding = Reader.ReadReal32()
   
    charlist(CharIndex).Speeding = Speeding
    
    Exit Sub

HandleSpeedToChar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSpeedToChar", Erl)
    
    
End Sub
Private Sub HandleNieveToggle()
    '**
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '**
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
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
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

    Dim CharIndex As Integer

    Dim BarTime   As Integer

    Dim BarAccion As Byte
    
    CharIndex = Reader.ReadInt16()
    BarTime = Reader.ReadInt16()
    BarAccion = Reader.ReadInt8()
    
    charlist(CharIndex).BarTime = 0
    charlist(CharIndex).BarAccion = BarAccion
    charlist(CharIndex).MaxBarTime = BarTime
    
    Exit Sub

HandleBarFx_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleBarFx", Erl)
    
    
End Sub
 
Private Sub HandleQuestDetails()

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Recibe y maneja el paquete QuestDetails del servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    
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
            Dim RequiredSkill, RequiredValue As Byte
            RequiredSkill = Reader.ReadInt8
            RequiredValue = Reader.ReadInt8
            If RequiredSkill > 0 Then
                FrmQuests.detalle.Text = FrmQuests.detalle.Text & SkillsNames(RequiredSkill) & ": " & RequiredValue
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
            tmpByte = Reader.ReadInt8 'skills
            For i = 1 To tmpByte
                obindex = Reader.ReadInt16
                Set subelemento = FrmQuests.ListView2.ListItems.Add(, , HechizoData(obindex).nombre)
                subelemento.SubItems(1) = 1
                subelemento.SubItems(2) = obindex
                subelemento.SubItems(3) = 1
            Next i
        End If

    'Determinamos que formulario se muestra, segï¿½n si recibimos la informaciï¿½n y la quest estï¿½ empezada o no.
    If QuestEmpezada Then
        FrmQuests.txtInfo.Text = tmpStr
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
    FrmQuests.txtInfo.Text = vbNullString
        
    'Si el usuario tiene quests entonces hacemos el handle
    If tmpByte Then
        'Leemos el string
        tmpStr = Reader.ReadString8
        
        'Agregamos los items
        For i = 1 To tmpByte
            FrmQuests.lstQuests.AddItem ReadField(i, tmpStr, 59)
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
                
        Next J

    'Determinamos que formulario se muestra, segun si recibimos la informacion y la quest estï¿½ empezada o no.
    FrmQuestInfo.Show vbModeless, GetGameplayForm()
    FrmQuestInfo.Picture = LoadInterface("ventananuevamision.bmp")
    Call FrmQuestInfo.ShowQuest(1)
    
    Exit Sub
    
errhandler:

    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleNpcQuestListSend", Erl)
    
    
End Sub

Private Sub HandleShowPregunta()

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    '
    '***************************************************
    
    On Error GoTo errhandler
    
    Dim msg As String

    msg = Reader.ReadString8()
    PreguntaScreen = msg
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
    
    On Error GoTo HandleCharacterChange_Err
    
    Dim CharIndex As Integer

    Dim TempInt   As Integer

    Dim headIndex As Integer

    CharIndex = Reader.ReadInt16()
    
    With charlist(CharIndex)
        .AnimatingBody = Reader.ReadInt16()
        .Body = BodyData(.AnimatingBody)
        'Start animation
        .Body.Walk(.Heading).started = FrameTime
        .Body.Walk(.Heading).Loops = 0
        .Idle = False
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

    frmCrafteo.Show , GetGameplayForm()

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
    
    Call MsgBox(JsonLanguage.Item("MENSAJEBOX_NUEVA_VERSION"), vbOKOnly, "Argentum 20 - Noland Studios")
    
    Shell App.path & "\..\..\Launcher\LauncherAO20.exe"
    
    EngineRun = False

    Call CloseClient
    
    Exit Sub

HandleCerrarleCliente_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleCerrarleCliente", Erl)
    
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
    Call Char_Dialog_Set(UserCharIndex, "Oh! Creo que tengo un super pez en mi linea, intentare obtenerlo con la letra P", &H1FFFF, 200, 130)
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
    
    Dim cant_obj_shop As Long, i As Long
    
    cant_obj_shop = Reader.ReadInt16
    
    credits_shopAO20 = Reader.ReadInt32
    frmShopAO20.lblCredits.Caption = credits_shopAO20
    
        ReDim ObjShop(1 To cant_obj_shop) As ObjDatas
        
        For i = 1 To cant_obj_shop
            ObjShop(i).ObjNum = Reader.ReadInt32
            ObjShop(i).Valor = Reader.ReadInt32
            ObjShop(i).Name = Reader.ReadString8
             
            Call frmShopAO20.lstItemShopFilter.AddItem(ObjShop(i).Name & " (Valor: " & ObjShop(i).Valor & ")", i - 1)
        Next i
        frmShopAO20.Show , GetGameplayForm()
 
End Sub

Public Sub HandleUpdateShopClienteCredits()
    credits_shopAO20 = Reader.ReadInt32
    frmShopAO20.lblCredits.Caption = credits_shopAO20
End Sub

Public Sub HandleSendSkillCdUpdate()
    On Error GoTo errhandler
        Dim Effect As t_ActiveEffect
        Dim ElapsedTime As Long
100     Effect.TypeId = Reader.ReadInt16
102     Effect.id = Reader.ReadInt32
104     ElapsedTime = Reader.ReadInt32
106     Effect.duration = Reader.ReadInt32
108     Effect.EffectType = Reader.ReadInt8
110     Effect.grh = EffectResources(Effect.TypeId).GrhId
112     Effect.startTime = GetTickCount() - (Effect.duration - ElapsedTime)
114     Effect.StackCount = Reader.ReadInt16()
116     If Effect.EffectType = eBuff Then
118         Call AddOrUpdateEffect(BuffList, Effect)
        End If
120     If Effect.EffectType = eDebuff Then
122         Call AddOrUpdateEffect(DeBuffList, Effect)
        End If
124     If Effect.EffectType = eCD Then
130         Call AddOrUpdateEffect(CDList, Effect)
        End If
        Exit Sub
errhandler:
132     Call RegistrarError(Err.Number, Err.Description, "Protocol.HandleSendSkillCdUpdate " & Effect.TypeId, Erl)
End Sub

Public Sub HandleSendClientToggles()
On Error GoTo errhandler
    Dim ToggleCount As Integer
    ToggleCount = Reader.ReadInt16
    Dim i As Integer
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

    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Recibe y maneja el paquete QuestListSend del servidor.
    'Last modified: 29/08/2021 by HarThaoS
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

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
    Dim mensaje As String
    Dim i As Integer

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
    Dim LobbyList() As t_LobbyData
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

#If PYMMO = 0 Then
    
Public Sub HandleAccountCharacterList()

    CantidadDePersonajesEnCuenta = Reader.ReadInt

    Dim ii As Byte
     'name, head_id, class_id, body_id, pos_map, pos_x, pos_y, level, status, helmet_id, shield_id, weapon_id, guild_index, is_dead, is_sailing
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
        Pjs(ii).ClanName = "" ' "<" & "pepito" & ">"

    Next ii
    
    Dim i As Long
    For i = 1 To min(CantidadDePersonajesEnCuenta, MAX_PERSONAJES_EN_CUENTA)
        Select Case Pjs(i).Criminal
            Case 0 'Criminal
                Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(50).r, ColoresPJ(50).G, ColoresPJ(50).b)
                Pjs(i).priv = 0
            Case 1 'Ciudadano
                Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(49).r, ColoresPJ(49).G, ColoresPJ(49).b)
                Pjs(i).priv = 0
            Case 2 'Caos
                Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(6).r, ColoresPJ(6).G, ColoresPJ(6).b)
                Pjs(i).priv = 0
            Case 3 'Armada
                Call SetRGBA(Pjs(i).LetraColor, ColoresPJ(8).r, ColoresPJ(8).G, ColoresPJ(8).b)
                Pjs(i).priv = 0
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

    Call LoadCharacterSelectionScreen
End Sub
#End If

