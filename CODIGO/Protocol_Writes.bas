Attribute VB_Name = "Protocol_Writes"
Option Explicit

Private Writer As Network.Writer

Public Sub Initialize()
    Set Writer = New Network.Writer
End Sub

Public Sub Clear()
    Call Writer.Clear
End Sub

Public Sub Flush(ByVal Connection As Network.Client)
    If (Writer.GetOffset() = 0) Then
        Exit Sub
    End If
    
    Call Connection.Send(False, Writer)
    Call Connection.Flush
    
    Call Writer.Clear
End Sub

''
' Writes the "LoginExistingChar" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteLoginExistingChar()
    Call Writer.WriteInt(ClientPacketID.LoginExistingChar)
    Call Writer.WriteString8(CuentaEmail)
    Call Writer.WriteString8(SEncriptar(CuentaPassword))
    Call Writer.WriteInt8(App.Major)
    Call Writer.WriteInt8(App.Minor)
    Call Writer.WriteInt8(App.Revision)
    Call Writer.WriteString8(UserName)
    Call Writer.WriteString8(MacAdress)  'Seguridad
    Call Writer.WriteInt32(HDserial)  'SeguridadHDserial
    Call Writer.WriteString8(CheckMD5)
End Sub

''
' Writes the "LoginNewChar" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteLoginNewChar()
    Call Writer.WriteInt(ClientPacketID.LoginNewChar)
    Call Writer.WriteString8(CuentaEmail)
    Call Writer.WriteString8(SEncriptar(CuentaPassword))
    Call Writer.WriteInt8(App.Major)
    Call Writer.WriteInt8(App.Minor)
    Call Writer.WriteInt8(App.Revision)
    Call Writer.WriteString8(UserName)
    Call Writer.WriteInt8(UserRaza)
    Call Writer.WriteInt8(UserSexo)
    Call Writer.WriteInt8(UserClase)
    Call Writer.WriteInt16(MiCabeza)
    Call Writer.WriteInt8(UserHogar)
    Call Writer.WriteString8(MacAdress)  'Seguridad
    Call Writer.WriteInt32(HDserial)  'SeguridadHDserial
    Call Writer.WriteString8(CheckMD5)
End Sub

''
' Writes the "Talk" message to the outgoing data buffer.
'
' @param    chat The chat text to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteTalk(ByVal chat As String)
    Call Writer.WriteInt(ClientPacketID.Talk)
    Call Writer.WriteString8(chat)
End Sub

''
' Writes the "Yell" message to the outgoing data buffer.
'
' @param    chat The chat text to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteYell(ByVal chat As String)
    Call Writer.WriteInt(ClientPacketID.Yell)
    Call Writer.WriteString8(chat)
End Sub

''
' Writes the "Whisper" message to the outgoing data buffer.
'
' @param    charIndex The index of the char to whom to whisper.
' @param    chat The chat text to be sent to the user.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteWhisper(ByVal nombre As String, ByVal chat As String)
    Call Writer.WriteInt(ClientPacketID.Whisper)
    Call Writer.WriteString8(nombre)
    Call Writer.WriteString8(chat)
End Sub

''
' Writes the "Walk" message to the outgoing data buffer.
'
' @param    heading The direction in wich the user is moving.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteWalk(ByVal Heading As E_Heading)
    Call Writer.WriteInt(ClientPacketID.Walk)
    Call Writer.WriteInt8(Heading)
End Sub

''
' Writes the "RequestPositionUpdate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRequestPositionUpdate()
    Call Writer.WriteInt(ClientPacketID.RequestPositionUpdate)
End Sub

''
' Writes the "Attack" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteAttack()
    Call Writer.WriteInt(ClientPacketID.Attack)
End Sub

''
' Writes the "PickUp" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePickUp()
    Call Writer.WriteInt(ClientPacketID.PickUp)
End Sub

''
' Writes the "SafeToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSafeToggle()
    Call Writer.WriteInt(ClientPacketID.SafeToggle)
End Sub

Public Sub WriteSeguroClan()
    Call Writer.WriteInt(ClientPacketID.SeguroClan)
End Sub

Public Sub WriteTraerBoveda()
    Call Writer.WriteInt(ClientPacketID.TraerBoveda)
End Sub

''
' Writes the "CreatePretorianClan" message to the outgoing data buffer.
'
' @param    Map         The map in which create the pretorian clan.
' @param    X           The x pos where the king is settled.
' @param    Y           The y pos where the king is settled.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCreatePretorianClan(ByVal map As Integer, _
                                    ByVal x As Byte, _
                                    ByVal y As Byte)
    Call Writer.WriteInt(ClientPacketID.CreatePretorianClan)
    Call Writer.WriteInt16(map)
    Call Writer.WriteInt8(x)
    Call Writer.WriteInt8(y)
End Sub

''
' Writes the "DeletePretorianClan" message to the outgoing data buffer.
'
' @param    Map         The map which contains the pretorian clan to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteDeletePretorianClan(ByVal map As Integer)
    Call Writer.WriteInt(ClientPacketID.RemovePretorianClan)
    Call Writer.WriteInt16(map)
End Sub

''
' Writes the "PartySafeToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteParyToggle()
    Call Writer.WriteInt(ClientPacketID.PartySafeToggle)
End Sub

''
' Writes the "SeguroResu" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSeguroResu()
    Call Writer.WriteInt(ClientPacketID.SeguroResu)
End Sub

''
' Writes the "RequestGuildLeaderInfo" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRequestGuildLeaderInfo()
    Call Writer.WriteInt(ClientPacketID.RequestGuildLeaderInfo)
End Sub

''
' Writes the "RequestAtributes" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRequestAtributes()
    Call Writer.WriteInt(ClientPacketID.RequestAtributes)
End Sub

Public Sub WriteRequestFamiliar()
    Call Writer.WriteInt(ClientPacketID.RequestFamiliar)
End Sub

Public Sub WriteRequestGrupo()
    Call Writer.WriteInt(ClientPacketID.RequestGrupo)
End Sub

''
' Writes the "RequestSkills" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRequestSkills()
    Call Writer.WriteInt(ClientPacketID.RequestSkills)
End Sub

''
' Writes the "RequestMiniStats" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRequestMiniStats()
    Call Writer.WriteInt(ClientPacketID.RequestMiniStats)
End Sub

''
' Writes the "CommerceEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCommerceEnd()
    Call Writer.WriteInt(ClientPacketID.CommerceEnd)
End Sub

''
' Writes the "UserCommerceEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserCommerceEnd()
    Call Writer.WriteInt(ClientPacketID.UserCommerceEnd)
End Sub

''
' Writes the "BankEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBankEnd()
    Call Writer.WriteInt(ClientPacketID.BankEnd)
End Sub

''
' Writes the "UserCommerceOk" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserCommerceOk()
    Call Writer.WriteInt(ClientPacketID.UserCommerceOk)
End Sub

''
' Writes the "UserCommerceReject" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserCommerceReject()
    Call Writer.WriteInt(ClientPacketID.UserCommerceReject)
End Sub

''
' Writes the "Drop" message to the outgoing data buffer.
'
' @param    slot Inventory slot where the item to drop is.
' @param    amount Number of items to drop.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteDrop(ByVal Slot As Byte, ByVal Amount As Long)
    Call Writer.WriteInt(ClientPacketID.Drop)
    Call Writer.WriteInt8(Slot)
    Call Writer.WriteInt32(Amount)
End Sub

''
' Writes the "CastSpell" message to the outgoing data buffer.
'
' @param    slot Spell List slot where the spell to cast is.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCastSpell(ByVal Slot As Byte)
    Call Writer.WriteInt(ClientPacketID.CastSpell)
    Call Writer.WriteInt8(Slot)
End Sub

Public Sub WriteInvitarGrupo()
    Call Writer.WriteInt(ClientPacketID.InvitarGrupo)
End Sub

Public Sub WriteMarcaDeClan()
    Call Writer.WriteInt(ClientPacketID.MarcaDeClanPack)
End Sub

Public Sub WriteMarcaDeGm()
    Call Writer.WriteInt(ClientPacketID.MarcaDeGMPack)
End Sub

Public Sub WriteAbandonarGrupo()
    Call Writer.WriteInt(ClientPacketID.AbandonarGrupo)
End Sub

Public Sub WriteEcharDeGrupo(ByVal indice As Byte)
    Call Writer.WriteInt(ClientPacketID.HecharDeGrupo)
    Call Writer.WriteInt8(indice)
End Sub

''
' Writes the "LeftClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteLeftClick(ByVal x As Byte, ByVal y As Byte)
    Call Writer.WriteInt(ClientPacketID.LeftClick)
    Call Writer.WriteInt8(x)
    Call Writer.WriteInt8(y)
End Sub

''
' Writes the "DoubleClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteDoubleClick(ByVal x As Byte, ByVal y As Byte)
    Call Writer.WriteInt(ClientPacketID.DoubleClick)
    Call Writer.WriteInt8(x)
    Call Writer.WriteInt8(y)
End Sub

''
' Writes the "Work" message to the outgoing data buffer.
'
' @param    skill The skill which the user attempts to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteWork(ByVal Skill As eSkill)
    Call Writer.WriteInt(ClientPacketID.Work)
    Call Writer.WriteInt8(Skill)
End Sub

Public Sub WriteThrowDice()
    Call Writer.WriteInt(ClientPacketID.ThrowDice)
End Sub

''
' Writes the "UseSpellMacro" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUseSpellMacro()
    Call Writer.WriteInt(ClientPacketID.UseSpellMacro)
End Sub

''
' Writes the "UseItem" message to the outgoing data buffer.
'
' @param    slot Invetory slot where the item to use is.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUseItem(ByVal Slot As Byte)
    Call Writer.WriteInt(ClientPacketID.UseItem)
    Call Writer.WriteInt8(Slot)
End Sub

''
' Writes the "CraftBlacksmith" message to the outgoing data buffer.
'
' @param    item Index of the item to craft in the list sent by the server.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCraftBlacksmith(ByVal Item As Integer)
    Call Writer.WriteInt(ClientPacketID.CraftBlacksmith)
    Call Writer.WriteInt16(Item)
End Sub

''
' Writes the "CraftCarpenter" message to the outgoing data buffer.
'
' @param    item Index of the item to craft in the list sent by the server.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCraftCarpenter(ByVal Item As Integer)
    Call Writer.WriteInt(ClientPacketID.CraftCarpenter)
    Call Writer.WriteInt16(Item)
End Sub

Public Sub WriteCraftAlquimista(ByVal Item As Integer)
    Call Writer.WriteInt(ClientPacketID.CraftAlquimista)
    Call Writer.WriteInt16(Item)
End Sub

Public Sub WriteCraftSastre(ByVal Item As Integer)
    Call Writer.WriteInt(ClientPacketID.CraftSastre)
    Call Writer.WriteInt16(Item)
End Sub

''
' Writes the "WorkLeftClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @param    skill The skill which the user attempts to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteWorkLeftClick(ByVal x As Byte, ByVal y As Byte, ByVal Skill As eSkill)
    Call Writer.WriteInt(ClientPacketID.WorkLeftClick)
    Call Writer.WriteInt8(x)
    Call Writer.WriteInt8(y)
    Call Writer.WriteInt8(Skill)
End Sub

''
' Writes the "CreateNewGuild" message to the outgoing data buffer.
'
' @param    desc    The guild's description
' @param    name    The guild's name
' @param    site    The guild's website
' @param    codex   Array of all rules of the guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCreateNewGuild(ByVal desc As String, _
                               ByVal Name As String, _
                               ByVal Alineacion As Byte)
    Call Writer.WriteInt(ClientPacketID.CreateNewGuild)
    Call Writer.WriteString8(desc)
    Call Writer.WriteString8(Name)
    Call Writer.WriteInt8(Alineacion)
End Sub

''
' Writes the "SpellInfo" message to the outgoing data buffer.
'
' @param    slot Spell List slot where the spell which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSpellInfo(ByVal Slot As Byte)
    Call Writer.WriteInt(ClientPacketID.SpellInfo)
    Call Writer.WriteInt8(Slot)
End Sub

''
' Writes the "EquipItem" message to the outgoing data buffer.
'
' @param    slot Invetory slot where the item to equip is.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteEquipItem(ByVal Slot As Byte)
    Call Writer.WriteInt(ClientPacketID.EquipItem)
    Call Writer.WriteInt8(Slot)
End Sub

''
' Writes the "ChangeHeading" message to the outgoing data buffer.
'
' @param    heading The direction in wich the user is moving.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeHeading(ByVal Heading As E_Heading)
    Call Writer.WriteInt(ClientPacketID.ChangeHeading)
    Call Writer.WriteInt8(Heading)
End Sub

''
' Writes the "ModifySkills" message to the outgoing data buffer.
'
' @param    skillEdt a-based array containing for each skill the number of points to add to it.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteModifySkills(ByRef skillEdt() As Byte)
    Call Writer.WriteInt(ClientPacketID.ModifySkills)
    
    Dim i As Long
    
    For i = 1 To NUMSKILLS
        Call Writer.WriteInt8(skillEdt(i))
    Next i

End Sub

''
' Writes the "Train" message to the outgoing data buffer.
'
' @param    creature Position within the list provided by the server of the creature to train against.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteTrain(ByVal creature As Byte)
    Call Writer.WriteInt(ClientPacketID.Train)
    Call Writer.WriteInt8(creature)
End Sub

''
' Writes the "CommerceBuy" message to the outgoing data buffer.
'
' @param    slot Position within the NPC's inventory in which the desired item is.
' @param    amount Number of items to buy.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCommerceBuy(ByVal Slot As Byte, ByVal Amount As Integer)
    Call Writer.WriteInt(ClientPacketID.CommerceBuy)
    Call Writer.WriteInt8(Slot)
    Call Writer.WriteInt16(Amount)
End Sub

Public Sub WriteUseKey(ByVal Slot As Byte)
    Call Writer.WriteInt(ClientPacketID.UseKey)
    Call Writer.WriteInt8(Slot)
End Sub

''
' Writes the "BankExtractItem" message to the outgoing data buffer.
'
' @param    slot Position within the bank in which the desired item is.
' @param    amount Number of items to extract.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBankExtractItem(ByVal Slot As Byte, _
                                ByVal Amount As Integer, _
                                ByVal slotdestino As Byte)
    Call Writer.WriteInt(ClientPacketID.BankExtractItem)
    Call Writer.WriteInt8(Slot)
    Call Writer.WriteInt16(Amount)
    Call Writer.WriteInt8(slotdestino)
End Sub

''
' Writes the "CommerceSell" message to the outgoing data buffer.
'
' @param    slot Position within user inventory in which the desired item is.
' @param    amount Number of items to sell.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCommerceSell(ByVal Slot As Byte, ByVal Amount As Integer)
    Call Writer.WriteInt(ClientPacketID.CommerceSell)
    Call Writer.WriteInt8(Slot)
    Call Writer.WriteInt16(Amount)
End Sub

''
' Writes the "BankDeposit" message to the outgoing data buffer.
'
' @param    slot Position within the user inventory in which the desired item is.
' @param    amount Number of items to deposit.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBankDeposit(ByVal Slot As Byte, _
                            ByVal Amount As Integer, _
                            ByVal slotdestino As Byte)
    Call Writer.WriteInt(ClientPacketID.BankDeposit)
    Call Writer.WriteInt8(Slot)
    Call Writer.WriteInt16(Amount)
    Call Writer.WriteInt8(slotdestino)
End Sub

''
' Writes the "ForumPost" message to the outgoing data buffer.
'
' @param    title The message's title.
' @param    message The body of the message.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteForumPost(ByVal title As String, ByVal Message As String)
    Call Writer.WriteInt(ClientPacketID.ForumPost)
    Call Writer.WriteString8(title)
    Call Writer.WriteString8(Message)
End Sub

''
' Writes the "MoveSpell" message to the outgoing data buffer.
'
' @param    upwards True if the spell will be moved up in the list, False if it will be moved downwards.
' @param    slot Spell List slot where the spell which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteMoveSpell(ByVal upwards As Boolean, ByVal Slot As Byte)
    Call Writer.WriteInt(ClientPacketID.MoveSpell)
    Call Writer.WriteBool(upwards)
    Call Writer.WriteInt8(Slot)
End Sub

''
' Writes the "ClanCodexUpdate" message to the outgoing data buffer.
'
' @param    desc New description of the clan.
' @param    codex New codex of the clan.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteClanCodexUpdate(ByVal desc As String)
    Call Writer.WriteInt(ClientPacketID.ClanCodexUpdate)
    Call Writer.WriteString8(desc)
End Sub

''
' Writes the "UserCommerceOffer" message to the outgoing data buffer.
'
' @param    slot Position within user inventory in which the desired item is.
' @param    amount Number of items to offer.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserCommerceOffer(ByVal Slot As Byte, ByVal Amount As Long)
    Call Writer.WriteInt(ClientPacketID.UserCommerceOffer)
    Call Writer.WriteInt8(Slot)
    Call Writer.WriteInt32(Amount)
End Sub

''
' Writes the "GuildAcceptPeace" message to the outgoing data buffer.
'
' @param    guild The guild whose peace offer is accepted.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildAcceptPeace(ByVal guild As String)
    Call Writer.WriteInt(ClientPacketID.GuildAcceptPeace)
    Call Writer.WriteString8(guild)
End Sub

''
' Writes the "GuildRejectAlliance" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance offer is rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildRejectAlliance(ByVal guild As String)
    Call Writer.WriteInt(ClientPacketID.GuildRejectAlliance)
    Call Writer.WriteString8(guild)
End Sub

''
' Writes the "GuildRejectPeace" message to the outgoing data buffer.
'
' @param    guild The guild whose peace offer is rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildRejectPeace(ByVal guild As String)
    Call Writer.WriteInt(ClientPacketID.GuildRejectPeace)
    Call Writer.WriteString8(guild)
End Sub

''
' Writes the "GuildAcceptAlliance" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance offer is accepted.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildAcceptAlliance(ByVal guild As String)
    Call Writer.WriteInt(ClientPacketID.GuildAcceptAlliance)
    Call Writer.WriteString8(guild)
End Sub

''
' Writes the "GuildOfferPeace" message to the outgoing data buffer.
'
' @param    guild The guild to whom peace is offered.
' @param    proposal The text to s the proposal.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildOfferPeace(ByVal guild As String, ByVal proposal As String)
    Call Writer.WriteInt(ClientPacketID.GuildOfferPeace)
    Call Writer.WriteString8(guild)
    Call Writer.WriteString8(proposal)
End Sub

''
' Writes the "GuildOfferAlliance" message to the outgoing data buffer.
'
' @param    guild The guild to whom an aliance is offered.
' @param    proposal The text to s the proposal.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildOfferAlliance(ByVal guild As String, ByVal proposal As String)
    Call Writer.WriteInt(ClientPacketID.GuildOfferAlliance)
    Call Writer.WriteString8(guild)
    Call Writer.WriteString8(proposal)
End Sub

''
' Writes the "GuildAllianceDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance proposal's details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildAllianceDetails(ByVal guild As String)
    Call Writer.WriteInt(ClientPacketID.GuildAllianceDetails)
    Call Writer.WriteString8(guild)
End Sub

''
' Writes the "GuildPeaceDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose peace proposal's details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildPeaceDetails(ByVal guild As String)
    Call Writer.WriteInt(ClientPacketID.GuildPeaceDetails)
    Call Writer.WriteString8(guild)
End Sub

''
' Writes the "GuildRequestJoinerInfo" message to the outgoing data buffer.
'
' @param    username The user who wants to join the guild whose info is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildRequestJoinerInfo(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.GuildRequestJoinerInfo)
    Call Writer.WriteString8(UserName)
End Sub

''
' Writes the "GuildAlliancePropList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildAlliancePropList()
    Call Writer.WriteInt(ClientPacketID.GuildAlliancePropList)
End Sub

''
' Writes the "GuildPeacePropList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildPeacePropList()
    Call Writer.WriteInt(ClientPacketID.GuildPeacePropList)
End Sub

''
' Writes the "GuildDeclareWar" message to the outgoing data buffer.
'
' @param    guild The guild to which to declare war.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildDeclareWar(ByVal guild As String)
    Call Writer.WriteInt(ClientPacketID.GuildDeclareWar)
    Call Writer.WriteString8(guild)
End Sub

''
' Writes the "GuildNewWebsite" message to the outgoing data buffer.
'
' @param    url The guild's new website's URL.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildNewWebsite(ByVal url As String)
    Call Writer.WriteInt(ClientPacketID.GuildNewWebsite)
    Call Writer.WriteString8(url)
End Sub

''
' Writes the "GuildAcceptNewMember" message to the outgoing data buffer.
'
' @param    username The name of the accepted player.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildAcceptNewMember(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.GuildAcceptNewMember)
    Call Writer.WriteString8(UserName)
End Sub

''
' Writes the "GuildRejectNewMember" message to the outgoing data buffer.
'
' @param    username The name of the rejected player.
' @param    reason The reason for which the player was rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildRejectNewMember(ByVal UserName As String, ByVal reason As String)
    Call Writer.WriteInt(ClientPacketID.GuildRejectNewMember)
    Call Writer.WriteString8(UserName)
    Call Writer.WriteString8(reason)
End Sub

''
' Writes the "GuildKickMember" message to the outgoing data buffer.
'
' @param    username The name of the kicked player.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildKickMember(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.GuildKickMember)
    Call Writer.WriteString8(UserName)
End Sub

''
' Writes the "GuildUpdateNews" message to the outgoing data buffer.
'
' @param    news The news to be posted.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildUpdateNews(ByVal news As String)
    Call Writer.WriteInt(ClientPacketID.GuildUpdateNews)
    Call Writer.WriteString8(news)
End Sub

''
' Writes the "GuildMemberInfo" message to the outgoing data buffer.
'
' @param    username The user whose info is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildMemberInfo(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.GuildMemberInfo)
    Call Writer.WriteString8(UserName)
End Sub

''
' Writes the "GuildOpenElections" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildOpenElections()
    Call Writer.WriteInt(ClientPacketID.GuildOpenElections)
End Sub

''
' Writes the "GuildRequestMembership" message to the outgoing data buffer.
'
' @param    guild The guild to which to request membership.
' @param    application The user's application sheet.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildRequestMembership(ByVal guild As String, ByVal Application As String)
    Call Writer.WriteInt(ClientPacketID.GuildRequestMembership)
    Call Writer.WriteString8(guild)
    Call Writer.WriteString8(Application)
End Sub

''
' Writes the "GuildRequestDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildRequestDetails(ByVal guild As String)
    Call Writer.WriteInt(ClientPacketID.GuildRequestDetails)
    Call Writer.WriteString8(guild)
End Sub

''
' Writes the "Online" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteOnline()
    Call Writer.WriteInt(ClientPacketID.Online)
End Sub

''
' Writes the "Quit" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteQuit()
    Call Writer.WriteInt(ClientPacketID.Quit)
    UserSaliendo = True
End Sub

''
' Writes the "GuildLeave" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildLeave()
    Call Writer.WriteInt(ClientPacketID.GuildLeave)
End Sub

''
' Writes the "RequestAccountState" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRequestAccountState()
    Call Writer.WriteInt(ClientPacketID.RequestAccountState)
End Sub

''
' Writes the "PetStand" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePetStand()
    Call Writer.WriteInt(ClientPacketID.PetStand)
End Sub

''
' Writes the "PetFollow" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePetFollow()
    Call Writer.WriteInt(ClientPacketID.PetFollow)
End Sub

''
' Writes the "PetLeave" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePetLeave()
    Call Writer.WriteInt(ClientPacketID.PetLeave)
End Sub

''
' Writes the "GrupoMsg" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGrupoMsg(ByVal Message As String)
    Call Writer.WriteInt(ClientPacketID.GrupoMsg)
    Call Writer.WriteString8(Message)
End Sub

''
' Writes the "TrainList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteTrainList()
    Call Writer.WriteInt(ClientPacketID.TrainList)
End Sub

''
' Writes the "Rest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRest()
    Call Writer.WriteInt(ClientPacketID.Rest)
End Sub

''
' Writes the "Meditate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteMeditate()
    Call Writer.WriteInt(ClientPacketID.Meditate)
End Sub

''
' Writes the "Resucitate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteResucitate()
    Call Writer.WriteInt(ClientPacketID.Resucitate)
End Sub

''
' Writes the "Heal" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteHeal()
    Call Writer.WriteInt(ClientPacketID.Heal)
End Sub

''
' Writes the "Help" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteHelp()
    Call Writer.WriteInt(ClientPacketID.Help)
End Sub

''
' Writes the "RequestStats" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRequestStats()
    Call Writer.WriteInt(ClientPacketID.RequestStats)
End Sub

''
' Writes the "Promedio" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePromedio()
    Call Writer.WriteInt(ClientPacketID.Promedio)
End Sub

''
' Writes the "GiveItem" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGiveItem(UserName As String, _
                         ByVal OBJIndex As Integer, _
                         ByVal cantidad As Integer, _
                         Motivo As String)
    Call Writer.WriteInt(ClientPacketID.GiveItem)
    Call Writer.WriteString8(UserName)
    Call Writer.WriteInt16(OBJIndex)
    Call Writer.WriteInt16(cantidad)
    Call Writer.WriteString8(Motivo)
End Sub

''
' Writes the "CommerceStart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCommerceStart()
    Call Writer.WriteInt(ClientPacketID.CommerceStart)
End Sub

''
' Writes the "BankStart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBankStart()
    Call Writer.WriteInt(ClientPacketID.BankStart)
End Sub

''
' Writes the "Enlist" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteEnlist()
    Call Writer.WriteInt(ClientPacketID.Enlist)
End Sub

''
' Writes the "Information" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteInformation()
    Call Writer.WriteInt(ClientPacketID.Information)
End Sub

''
' Writes the "Reward" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteReward()
    Call Writer.WriteInt(ClientPacketID.Reward)
End Sub

''
' Writes the "RequestMOTD" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRequestMOTD()
    Call Writer.WriteInt(ClientPacketID.RequestMOTD)
End Sub

''
' Writes the "UpTime" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUpTime()
    Call Writer.WriteInt(ClientPacketID.UpTime)
End Sub

''
' Writes the "Inquiry" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteInquiry()
    Call Writer.WriteInt(ClientPacketID.Inquiry)
End Sub

''
' Writes the "GuildMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildMessage(ByVal Message As String)
    Call Writer.WriteInt(ClientPacketID.GuildMessage)
    Call Writer.WriteString8(Message)
End Sub

''
' Writes the "CentinelReport" message to the outgoing data buffer.
'
' @param    number The number to report to the centinel.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCentinelReport(ByVal Number As Integer)
    Call Writer.WriteInt(ClientPacketID.CentinelReport)
    Call Writer.WriteInt16(Number)
End Sub

''
' Writes the "GuildOnline" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildOnline()
    Call Writer.WriteInt(ClientPacketID.GuildOnline)
End Sub

''
' Writes the "CouncilMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the other council members.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCouncilMessage(ByVal Message As String)
    Call Writer.WriteInt(ClientPacketID.CouncilMessage)
    Call Writer.WriteString8(Message)
End Sub

''
' Writes the "RoleMasterRequest" message to the outgoing data buffer.
'
' @param    message The message to send to the role masters.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRoleMasterRequest(ByVal Message As String)
    Call Writer.WriteInt(ClientPacketID.RoleMasterRequest)
    Call Writer.WriteString8(Message)
End Sub

''
' Writes the "ChangeDescription" message to the outgoing data buffer.
'
' @param    desc The new description of the user's character.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeDescription(ByVal desc As String)
    Call Writer.WriteInt(ClientPacketID.ChangeDescription)
    Call Writer.WriteString8(desc)
End Sub

''
' Writes the "GuildVote" message to the outgoing data buffer.
'
' @param    username The user to vote for clan leader.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildVote(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.GuildVote)
    Call Writer.WriteString8(UserName)
End Sub

''
' Writes the "Punishments" message to the outgoing data buffer.
'
' @param    username The user whose's  punishments are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePunishments(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.punishments)
    Call Writer.WriteString8(UserName)
End Sub

''
' Writes the "ChangePassword" message to the outgoing data buffer.
'
' @param    oldPass Previous password.
' @param    newPass New password.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangePassword(ByRef oldPass As String, ByRef newPass As String)
    Call Writer.WriteInt(ClientPacketID.ChangePassword)
    Call Writer.WriteString8(SEncriptar(oldPass))
    Call Writer.WriteString8(SEncriptar(newPass))
End Sub

''
' Writes the "Gamble" message to the outgoing data buffer.
'
' @param    amount The amount to gamble.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGamble(ByVal Amount As Integer)
    Call Writer.WriteInt(ClientPacketID.Gamble)
    Call Writer.WriteInt16(Amount)
End Sub

''
' Writes the "InquiryVote" message to the outgoing data buffer.
'
' @param    opt The chosen option to vote for.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteInquiryVote(ByVal opt As Byte)
    Call Writer.WriteInt(ClientPacketID.InquiryVote)
    Call Writer.WriteInt8(opt)
End Sub

''
' Writes the "LeaveFaction" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteLeaveFaction()
    Call Writer.WriteInt(ClientPacketID.LeaveFaction)
End Sub

''
' Writes the "BankExtractGold" message to the outgoing data buffer.
'
' @param    amount The amount of money to extract from the bank.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBankExtractGold(ByVal Amount As Long)
    Call Writer.WriteInt(ClientPacketID.BankExtractGold)
    Call Writer.WriteInt32(Amount)
End Sub

''
' Writes the "BankDepositGold" message to the outgoing data buffer.
'
' @param    amount The amount of money to deposit in the bank.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBankDepositGold(ByVal Amount As Long)
    Call Writer.WriteInt(ClientPacketID.BankDepositGold)
    Call Writer.WriteInt32(Amount)
End Sub

Public Sub WriteTransFerGold(ByVal Amount As Long, ByVal destino As String)
    Call Writer.WriteInt(ClientPacketID.TransFerGold)
    Call Writer.WriteInt32(Amount)
    Call Writer.WriteString8(destino)
End Sub

Public Sub WriteItemMove(ByVal SlotActual As Byte, ByVal SlotNuevo As Byte)
    Call Writer.WriteInt(ClientPacketID.Moveitem)
    Call Writer.WriteInt8(SlotActual)
    Call Writer.WriteInt8(SlotNuevo)
End Sub

Public Sub WriteBovedaItemMove(ByVal SlotActual As Byte, ByVal SlotNuevo As Byte)
    Call Writer.WriteInt(ClientPacketID.BovedaMoveItem)
    Call Writer.WriteInt8(SlotActual)
    Call Writer.WriteInt8(SlotNuevo)
End Sub

''
' Writes the "FinEvento" message to the outgoing data buffer.
'
' @param    message The message to s the denounce.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteFinEvento()
    Call Writer.WriteInt(ClientPacketID.FinEvento)
End Sub

''
' Writes the "Denounce" message to the outgoing data buffer.
'
' @param    message The message to s the denounce.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteDenounce(Name As String)
    Call Writer.WriteInt(ClientPacketID.Denounce)
    Call Writer.WriteString8(Name)
End Sub

Public Sub WriteQuieroFundarClan()
    Call Writer.WriteInt(ClientPacketID.QuieroFundarClan)
End Sub

''
' Writes the "GuildMemberList" message to the outgoing data buffer.
'
' @param    guild The guild whose member list is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildMemberList(ByVal guild As String)
    Call Writer.WriteInt(ClientPacketID.GuildMemberList)
    Call Writer.WriteString8(guild)
End Sub

Public Sub WriteCasamiento(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.Casarse)
    Call Writer.WriteString8(UserName)
End Sub

Public Sub WriteMacroPos()
    Call Writer.WriteInt(ClientPacketID.MacroPossent)
    Call Writer.WriteInt8(ChatCombate)
    Call Writer.WriteInt8(ChatGlobal)
End Sub

Public Sub WriteSubastaInfo()
    Call Writer.WriteInt(ClientPacketID.SubastaInfo)
End Sub

Public Sub WriteScrollInfo()
    Call Writer.WriteInt(ClientPacketID.SCROLLINFO)
End Sub

Public Sub WriteCancelarExit()
    UserSaliendo = False
    Call Writer.WriteInt(ClientPacketID.CancelarExit)
End Sub

Public Sub WriteEventoInfo()
    Call Writer.WriteInt(ClientPacketID.EventoInfo)
End Sub

Public Sub WriteFlagTrabajar()
    Call Writer.WriteInt(ClientPacketID.FlagTrabajar)
End Sub

''
' Writes the "GMMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to the other GMs online.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteEscribiendo()
    Call Writer.WriteInt(ClientPacketID.Escribiendo)
End Sub

Public Sub WriteReclamarRecompensa(ByVal Index As Byte)
    Call Writer.WriteInt(ClientPacketID.ReclamarRecompensa)
    Call Writer.WriteInt8(Index)
End Sub

Public Sub WriteGMMessage(ByVal Message As String)
    Call Writer.WriteInt(ClientPacketID.GMMessage)
    Call Writer.WriteString8(Message)
End Sub

''
' Writes the "ShowName" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowName()
    Call Writer.WriteInt(ClientPacketID.showName)
End Sub

''
' Writes the "OnlineRoyalArmy" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteOnlineRoyalArmy()
    Call Writer.WriteInt(ClientPacketID.OnlineRoyalArmy)
End Sub

''
' Writes the "OnlineChaosLegion" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteOnlineChaosLegion()
    Call Writer.WriteInt(ClientPacketID.OnlineChaosLegion)
End Sub

''
' Writes the "GoNearby" message to the outgoing data buffer.
'
' @param    username The suer to approach.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGoNearby(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.GoNearby)
    Call Writer.WriteString8(UserName)
End Sub

''
' Writes the "Comment" message to the outgoing data buffer.
'
' @param    message The message to leave in the log as a comment.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteComment(ByVal Message As String)
    Call Writer.WriteInt(ClientPacketID.comment)
    Call Writer.WriteString8(Message)
End Sub

''
' Writes the "ServerTime" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteServerTime()
    Call Writer.WriteInt(ClientPacketID.serverTime)
End Sub

''
' Writes the "Where" message to the outgoing data buffer.
'
' @param    username The user whose position is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteWhere(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.Where)
    Call Writer.WriteString8(UserName)
End Sub

''
' Writes the "CreaturesInMap" message to the outgoing data buffer.
'
' @param    map The map in which to check for the existing creatures.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCreaturesInMap(ByVal map As Integer)
    Call Writer.WriteInt(ClientPacketID.CreaturesInMap)
    Call Writer.WriteInt16(map)
End Sub

''
' Writes the "WarpMeToTarget" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteWarpMeToTarget()
    Call Writer.WriteInt(ClientPacketID.WarpMeToTarget)
End Sub

''
' Writes the "WarpChar" message to the outgoing data buffer.
'
' @param    username The user to be warped. "YO" represent's the user's char.
' @param    map The map to which to warp the character.
' @param    x The x position in the map to which to waro the character.
' @param    y The y position in the map to which to waro the character.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteWarpChar(ByVal UserName As String, _
                         ByVal map As Integer, _
                         ByVal x As Byte, _
                         ByVal y As Byte)
    Call Writer.WriteInt(ClientPacketID.WarpChar)
    Call Writer.WriteString8(UserName)
    Call Writer.WriteInt16(map)
    Call Writer.WriteInt8(x)
    Call Writer.WriteInt8(y)
End Sub

''
' Writes the "Silence" message to the outgoing data buffer.
'
' @param    username The user to silence.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSilence(ByVal UserName As String, ByVal Minutos As Integer)
    Call Writer.WriteInt(ClientPacketID.Silence)
    Call Writer.WriteString8(UserName)
    Call Writer.WriteInt16(Minutos)
End Sub

Public Sub WriteCuentaRegresiva(ByVal Second As Byte)
    Call Writer.WriteInt(ClientPacketID.CuentaRegresiva)
    Call Writer.WriteInt8(Second)
End Sub

Public Sub WritePossUser(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.PossUser)
    Call Writer.WriteString8(UserName)
End Sub

''
' Writes the "SOSShowList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSOSShowList()
    Call Writer.WriteInt(ClientPacketID.SOSShowList)
End Sub

''
' Writes the "SOSRemove" message to the outgoing data buffer.
'
' @param    username The user whose SOS call has been already attended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSOSRemove(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.SOSRemove)
    Call Writer.WriteString8(UserName)
End Sub

''
' Writes the "GoToChar" message to the outgoing data buffer.
'
' @param    username The user to be approached.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGoToChar(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.GoToChar)
    Call Writer.WriteString8(UserName)
End Sub

Public Sub WriteDesbuggear(ByVal Params As String)
    Call Writer.WriteInt(ClientPacketID.Desbuggear)
    Call Writer.WriteString8(Params)
End Sub

Public Sub WriteDarLlaveAUsuario(ByVal User As String, ByVal Llave As Integer)
    Call Writer.WriteInt(ClientPacketID.DarLlaveAUsuario)
    Call Writer.WriteString8(User)
    Call Writer.WriteInt16(Llave)
End Sub

Public Sub WriteSacarLlave(ByVal Llave As Integer)
    Call Writer.WriteInt(ClientPacketID.SacarLlave)
    Call Writer.WriteInt16(Llave)
End Sub

Public Sub WriteVerLlaves()
    Call Writer.WriteInt(ClientPacketID.VerLlaves)
End Sub

''
' Writes the "Invisible" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteInvisible()
    Call Writer.WriteInt(ClientPacketID.Invisible)
End Sub

''
' Writes the "GMPanel" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGMPanel()
    Call Writer.WriteInt(ClientPacketID.GMPanel)
End Sub

''
' Writes the "RequestUserList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRequestUserList()
    Call Writer.WriteInt(ClientPacketID.RequestUserList)
End Sub

''
' Writes the "Working" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteWorking()
    Call Writer.WriteInt(ClientPacketID.Working)
End Sub

''
' Writes the "Hiding" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteHiding()
    Call Writer.WriteInt(ClientPacketID.Hiding)
End Sub

''
' Writes the "Jail" message to the outgoing data buffer.
'
' @param    username The user to be sent to jail.
' @param    reason The reason for which to send him to jail.
' @param    time The time (in minutes) the user will have to spend there.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteJail(ByVal UserName As String, ByVal reason As String, ByVal Time As Byte)
    Call Writer.WriteInt(ClientPacketID.Jail)
    Call Writer.WriteString8(UserName)
    Call Writer.WriteString8(reason)
    Call Writer.WriteInt8(Time)
End Sub

Public Sub WriteCrearEvento(ByVal TIPO As Byte, _
                            ByVal duracion As Byte, _
                            ByVal multiplicacion As Byte)
    Call Writer.WriteInt(ClientPacketID.CrearEvento)
    Call Writer.WriteInt8(TIPO)
    Call Writer.WriteInt8(duracion)
    Call Writer.WriteInt8(multiplicacion)
End Sub

''
' Writes the "KillNPC" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteKillNPC()
    Call Writer.WriteInt(ClientPacketID.KillNPC)
End Sub

''
' Writes the "WarnUser" message to the outgoing data buffer.
'
' @param    username The user to be warned.
' @param    reason Reason for the warning.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteWarnUser(ByVal UserName As String, ByVal reason As String)
    Call Writer.WriteInt(ClientPacketID.WarnUser)
    Call Writer.WriteString8(UserName)
    Call Writer.WriteString8(reason)
End Sub

Public Sub WriteMensajeUser(ByVal UserName As String, ByVal mensaje As String)
    Call Writer.WriteInt(ClientPacketID.MensajeUser)
    Call Writer.WriteString8(UserName)
    Call Writer.WriteString8(mensaje)
End Sub

''
' Writes the "EditChar" message to the outgoing data buffer.
'
' @param    UserName    The user to be edited.
' @param    editOption  Indicates what to edit in the char.
' @param    arg1        Additional argument 1. Contents depend on editoption.
' @param    arg2        Additional argument 2. Contents depend on editoption.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteEditChar(ByVal UserName As String, _
                         ByVal editOption As eEditOptions, _
                         ByVal arg1 As String, _
                         ByVal arg2 As String)
    Call Writer.WriteInt(ClientPacketID.EditChar)
    Call Writer.WriteString8(UserName)
    Call Writer.WriteInt8(editOption)
    Call Writer.WriteString8(arg1)
    Call Writer.WriteString8(arg2)
End Sub

''
' Writes the "RequestCharInfo" message to the outgoing data buffer.
'
' @param    username The user whose information is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRequestCharInfo(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.RequestCharInfo)
    Call Writer.WriteString8(UserName)
End Sub

''
' Writes the "RequestCharStats" message to the outgoing data buffer.
'
' @param    username The user whose stats are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRequestCharStats(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.RequestCharStats)
    Call Writer.WriteString8(UserName)
End Sub

''
' Writes the "RequestCharGold" message to the outgoing data buffer.
'
' @param    username The user whose gold is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRequestCharGold(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.RequestCharGold)
    Call Writer.WriteString8(UserName)
End Sub
    
''
' Writes the "RequestCharInventory" message to the outgoing data buffer.
'
' @param    username The user whose inventory is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRequestCharInventory(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.RequestCharInventory)
    Call Writer.WriteString8(UserName)
End Sub

''
' Writes the "RequestCharBank" message to the outgoing data buffer.
'
' @param    username The user whose banking information is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRequestCharBank(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.RequestCharBank)
    Call Writer.WriteString8(UserName)
End Sub

''
' Writes the "RequestCharSkills" message to the outgoing data buffer.
'
' @param    username The user whose skills are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRequestCharSkills(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.RequestCharSkills)
    Call Writer.WriteString8(UserName)
End Sub

''
' Writes the "ReviveChar" message to the outgoing data buffer.
'
' @param    username The user to eb revived.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteReviveChar(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.ReviveChar)
    Call Writer.WriteString8(UserName)
End Sub

''
' Writes the "OnlineGM" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteOnlineGM()
    Call Writer.WriteInt(ClientPacketID.OnlineGM)
End Sub

''
' Writes the "OnlineMap" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteOnlineMap()
    Call Writer.WriteInt(ClientPacketID.OnlineMap)
End Sub

''
' Writes the "Forgive" message to the outgoing data buffer.
'
' @param    username The user to be forgiven.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteForgive()
    Call Writer.WriteInt(ClientPacketID.Forgive)
End Sub

Public Sub WriteDonateGold(ByVal oro As Long)
    Call Writer.WriteInt(ClientPacketID.DonateGold)
    Call Writer.WriteInt32(oro)
End Sub

''
' Writes the "Kick" message to the outgoing data buffer.
'
' @param    username The user to be kicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteKick(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.Kick)
    Call Writer.WriteString8(UserName)
End Sub

''
' Writes the "Execute" message to the outgoing data buffer.
'
' @param    username The user to be executed.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteExecute(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.Execute)
    Call Writer.WriteString8(UserName)
End Sub

''
' Writes the "BanChar" message to the outgoing data buffer.
'
' @param    username The user to be banned.
' @param    reason The reson for which the user is to be banned.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBanChar(ByVal UserName As String, ByVal reason As String)
    Call Writer.WriteInt(ClientPacketID.BanChar)
    Call Writer.WriteString8(UserName)
    Call Writer.WriteString8(reason)
End Sub

Public Sub WriteBanCuenta(ByVal UserName As String, ByVal reason As String)
    Call Writer.WriteInt(ClientPacketID.BanCuenta)
    Call Writer.WriteString8(UserName)
    Call Writer.WriteString8(reason)
End Sub

Public Sub WriteUnBanCuenta(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.UnbanCuenta)
    Call Writer.WriteString8(UserName)
End Sub

Public Sub WriteBanSerial(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.BanSerial)
    Call Writer.WriteString8(UserName)
End Sub

Public Sub WriteUnBanSerial(ByVal UserName As String, ByVal reason As String)
    Call Writer.WriteInt(ClientPacketID.unBanSerial)
    Call Writer.WriteString8(UserName)
End Sub

Public Sub WriteCerraCliente(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.CerrarCliente)
    Call Writer.WriteString8(UserName)
End Sub

Public Sub WriteBanTemporal(ByVal UserName As String, _
                            ByVal reason As String, _
                            ByVal dias As Byte)
    Call Writer.WriteInt(ClientPacketID.BanTemporal)
    Call Writer.WriteString8(UserName)
    Call Writer.WriteString8(reason)
    Call Writer.WriteInt8(dias)
End Sub

''
' Writes the "UnbanChar" message to the outgoing data buffer.
'
' @param    username The user to be unbanned.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUnbanChar(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.UnbanChar)
    Call Writer.WriteString8(UserName)
End Sub

''
' Writes the "NPCFollow" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteNPCFollow()
    Call Writer.WriteInt(ClientPacketID.NPCFollow)
End Sub

''
' Writes the "SummonChar" message to the outgoing data buffer.
'
' @param    username The user to be summoned.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSummonChar(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.SummonChar)
    Call Writer.WriteString8(UserName)
End Sub

''
' Writes the "SpawnListRequest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSpawnListRequest()
    Call Writer.WriteInt(ClientPacketID.SpawnListRequest)
End Sub

''
' Writes the "SpawnCreature" message to the outgoing data buffer.
'
' @param    creatureIndex The index of the creature in the spawn list to be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSpawnCreature(ByVal creatureIndex As Integer)
    Call Writer.WriteInt(ClientPacketID.SpawnCreature)
    Call Writer.WriteInt16(creatureIndex)
End Sub

''
' Writes the "ResetNPCInventory" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteResetNPCInventory()
    Call Writer.WriteInt(ClientPacketID.ResetNPCInventory)
End Sub

''
' Writes the "CleanWorld" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCleanWorld()
    Call Writer.WriteInt(ClientPacketID.CleanWorld)
End Sub

''
' Writes the "ServerMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteServerMessage(ByVal Message As String)
    Call Writer.WriteInt(ClientPacketID.ServerMessage)
    Call Writer.WriteString8(Message)
End Sub

''
' Writes the "NickToIP" message to the outgoing data buffer.
'
' @param    username The user whose IP is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteNickToIP(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.NickToIP)
    Call Writer.WriteString8(UserName)
End Sub

''
' Writes the "IPToNick" message to the outgoing data buffer.
'
' @param    IP The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteIPToNick(ByRef IP() As Byte)

    If UBound(IP()) - LBound(IP()) + 1 <> 4 Then Exit Sub   'Invalid IP

    Dim i As Long

    Call Writer.WriteInt(ClientPacketID.IPToNick)

    For i = LBound(IP()) To UBound(IP())
        Call Writer.WriteInt8(IP(i))
    Next i

End Sub

''
' Writes the "GuildOnlineMembers" message to the outgoing data buffer.
'
' @param    guild The guild whose online player list is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildOnlineMembers(ByVal guild As String)
    Call Writer.WriteInt(ClientPacketID.GuildOnlineMembers)
    Call Writer.WriteString8(guild)
End Sub

''
' Writes the "TeleportCreate" message to the outgoing data buffer.
'
' @param    map the map to which the teleport will lead.
' @param    x The position in the x axis to which the teleport will lead.
' @param    y The position in the y axis to which the teleport will lead.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteTeleportCreate(ByVal map As Integer, _
                               ByVal x As Byte, _
                               ByVal y As Byte, _
                               ByVal Motivo As String)
    Call Writer.WriteInt(ClientPacketID.TeleportCreate)
    Call Writer.WriteInt16(map)
    Call Writer.WriteInt8(x)
    Call Writer.WriteInt8(y)
    Call Writer.WriteString8(Motivo)
End Sub

''
' Writes the "TeleportDestroy" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteTeleportDestroy()
    Call Writer.WriteInt(ClientPacketID.TeleportDestroy)
End Sub

''
' Writes the "RainToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRainToggle()
    Call Writer.WriteInt(ClientPacketID.RainToggle)
End Sub

''
' Writes the "SetCharDescription" message to the outgoing data buffer.
'
' @param    desc The description to set to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSetCharDescription(ByVal desc As String)
    Call Writer.WriteInt(ClientPacketID.SetCharDescription)
    Call Writer.WriteString8(desc)
End Sub

''
' Writes the "ForceMIDIToMap" message to the outgoing data buffer.
'
' @param    midiID The ID of the midi file to play.
' @param    map The map in which to play the given midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteForceMIDIToMap(ByVal midiID As Byte, ByVal map As Integer)
    Call Writer.WriteInt(ClientPacketID.ForceMIDIToMap)
    Call Writer.WriteInt8(midiID)
    Call Writer.WriteInt16(map)
End Sub

''
' Writes the "ForceWAVEToMap" message to the outgoing data buffer.
'
' @param    waveID  The ID of the wave file to play.
' @param    Map     The map into which to play the given wave.
' @param    x       The position in the x axis in which to play the given wave.
' @param    y       The position in the y axis in which to play the given wave.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteForceWAVEToMap(ByVal waveID As Byte, _
                               ByVal map As Integer, _
                               ByVal x As Byte, _
                               ByVal y As Byte)
    Call Writer.WriteInt(ClientPacketID.ForceWAVEToMap)
    Call Writer.WriteInt8(waveID)
    Call Writer.WriteInt16(map)
    Call Writer.WriteInt8(x)
    Call Writer.WriteInt8(y)
End Sub

''
' Writes the "RoyalArmyMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the royal army members.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRoyalArmyMessage(ByVal Message As String)
    Call Writer.WriteInt(ClientPacketID.RoyalArmyMessage)
    Call Writer.WriteString8(Message)
End Sub

''
' Writes the "ChaosLegionMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the chaos legion member.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChaosLegionMessage(ByVal Message As String)
    Call Writer.WriteInt(ClientPacketID.ChaosLegionMessage)
    Call Writer.WriteString8(Message)
End Sub

''
' Writes the "CitizenMessage" message to the outgoing data buffer.
'
' @param    message The message to send to citizens.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCitizenMessage(ByVal Message As String)
    Call Writer.WriteInt(ClientPacketID.CitizenMessage)
    Call Writer.WriteString8(Message)
End Sub

''
' Writes the "CriminalMessage" message to the outgoing data buffer.
'
' @param    message The message to send to criminals.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCriminalMessage(ByVal Message As String)
    Call Writer.WriteInt(ClientPacketID.CriminalMessage)
    Call Writer.WriteString8(Message)
End Sub

''
' Writes the "TalkAsNPC" message to the outgoing data buffer.
'
' @param    message The message to send to the royal army members.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteTalkAsNPC(ByVal Message As String)
    Call Writer.WriteInt(ClientPacketID.TalkAsNPC)
    Call Writer.WriteString8(Message)
End Sub

''
' Writes the "DestroyAllItemsInArea" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteDestroyAllItemsInArea()
    Call Writer.WriteInt(ClientPacketID.DestroyAllItemsInArea)
End Sub

''
' Writes the "AcceptRoyalCouncilMember" message to the outgoing data buffer.
'
' @param    username The name of the user to be accepted into the royal army council.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteAcceptRoyalCouncilMember(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.AcceptRoyalCouncilMember)
    Call Writer.WriteString8(UserName)
End Sub

''
' Writes the "AcceptChaosCouncilMember" message to the outgoing data buffer.
'
' @param    username The name of the user to be accepted as a chaos council member.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteAcceptChaosCouncilMember(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.AcceptChaosCouncilMember)
    Call Writer.WriteString8(UserName)
End Sub

''
' Writes the "ItemsInTheFloor" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteItemsInTheFloor()
    Call Writer.WriteInt(ClientPacketID.ItemsInTheFloor)
End Sub

''
' Writes the "MakeDumb" message to the outgoing data buffer.
'
' @param    username The name of the user to be made dumb.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteMakeDumb(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.MakeDumb)
    Call Writer.WriteString8(UserName)
End Sub

''
' Writes the "MakeDumbNoMore" message to the outgoing data buffer.
'
' @param    username The name of the user who will no longer be dumb.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteMakeDumbNoMore(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.MakeDumbNoMore)
    Call Writer.WriteString8(UserName)
End Sub

''
' Writes the "DumpIPTables" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteDumpIPTables()
    Call Writer.WriteInt(ClientPacketID.DumpIPTables)
End Sub

''
' Writes the "CouncilKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the council.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCouncilKick(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.CouncilKick)
    Call Writer.WriteString8(UserName)
End Sub

''
' Writes the "SetTrigger" message to the outgoing data buffer.
'
' @param    trigger The type of trigger to be set to the tile.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSetTrigger(ByVal Trigger As eTrigger)
    Call Writer.WriteInt(ClientPacketID.SetTrigger)
    Call Writer.WriteInt8(Trigger)
End Sub

''
' Writes the "AskTrigger" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteAskTrigger()
    Call Writer.WriteInt(ClientPacketID.AskTrigger)
End Sub

''
' Writes the "BannedIPList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBannedIPList()
    Call Writer.WriteInt(ClientPacketID.BannedIPList)
End Sub

''
' Writes the "BannedIPReload" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBannedIPReload()
    Call Writer.WriteInt(ClientPacketID.BannedIPReload)
End Sub

''
' Writes the "GuildBan" message to the outgoing data buffer.
'
' @param    guild The guild whose members will be banned.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildBan(ByVal guild As String)
    Call Writer.WriteInt(ClientPacketID.GuildBan)
    Call Writer.WriteString8(guild)
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
Public Sub WriteBanIP(ByVal NickOrIP As String, ByVal reason As String)
    Call Writer.WriteInt(ClientPacketID.banip)
    Call Writer.WriteString8(NickOrIP)
    Call Writer.WriteString8(reason)
End Sub

''
' Writes the "UnbanIP" message to the outgoing data buffer.
'
' @param    IP The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUnbanIP(ByRef IP() As Byte)

    If UBound(IP()) - LBound(IP()) + 1 <> 4 Then Exit Sub   'Invalid IP

    Dim i As Long

    Call Writer.WriteInt(ClientPacketID.UnBanIp)

    For i = LBound(IP()) To UBound(IP())
        Call Writer.WriteInt8(IP(i))
    Next i

End Sub

''
' Writes the "CreateItem" message to the outgoing data buffer.
'
' @param    itemIndex The index of the item to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCreateItem(ByVal ItemIndex As Long, ByVal cantidad As Integer)
    Call Writer.WriteInt(ClientPacketID.CreateItem)
    Call Writer.WriteInt16(ItemIndex)
    Call Writer.WriteInt16(cantidad)
End Sub

''
' Writes the "DestroyItems" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteDestroyItems()
    Call Writer.WriteInt(ClientPacketID.DestroyItems)
End Sub

''
' Writes the "ChaosLegionKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the Chaos Legion.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChaosLegionKick(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.ChaosLegionKick)
    Call Writer.WriteString8(UserName)
End Sub

''
' Writes the "RoyalArmyKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the Royal Army.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRoyalArmyKick(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.RoyalArmyKick)
    Call Writer.WriteString8(UserName)
End Sub

''
' Writes the "ForceMIDIAll" message to the outgoing data buffer.
'
' @param    midiID The id of the midi file to play.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteForceMIDIAll(ByVal midiID As Byte)
    Call Writer.WriteInt(ClientPacketID.ForceMIDIAll)
    Call Writer.WriteInt8(midiID)
End Sub

''
' Writes the "ForceWAVEAll" message to the outgoing data buffer.
'
' @param    waveID The id of the wave file to play.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteForceWAVEAll(ByVal waveID As Byte)
    Call Writer.WriteInt(ClientPacketID.ForceWAVEAll)
    Call Writer.WriteInt8(waveID)
End Sub

''
' Writes the "RemovePunishment" message to the outgoing data buffer.
'
' @param    username The user whose punishments will be altered.
' @param    punishment The id of the punishment to be removed.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRemovePunishment(ByVal UserName As String, _
                                 ByVal punishment As Byte, _
                                 ByVal NewText As String)
    Call Writer.WriteInt(ClientPacketID.RemovePunishment)
    Call Writer.WriteString8(UserName)
    Call Writer.WriteInt8(punishment)
    Call Writer.WriteString8(NewText)
End Sub

''
' Writes the "TileBlockedToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteTileBlockedToggle()
    Call Writer.WriteInt(ClientPacketID.TileBlockedToggle)
End Sub

''
' Writes the "KillNPCNoRespawn" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteKillNPCNoRespawn()
    Call Writer.WriteInt(ClientPacketID.KillNPCNoRespawn)
End Sub

''
' Writes the "KillAllNearbyNPCs" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteKillAllNearbyNPCs()
    Call Writer.WriteInt(ClientPacketID.KillAllNearbyNPCs)
End Sub

''
' Writes the "LastIP" message to the outgoing data buffer.
'
' @param    username The user whose last IPs are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteLastIP(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.LastIP)
    Call Writer.WriteString8(UserName)
End Sub

''
' Writes the "ChangeMOTD" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeMOTD()
    Call Writer.WriteInt(ClientPacketID.ChangeMOTD)
End Sub

''
' Writes the "SetMOTD" message to the outgoing data buffer.
'
' @param    message The message to be set as the new MOTD.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSetMOTD(ByVal Message As String)
    Call Writer.WriteInt(ClientPacketID.SetMOTD)
    Call Writer.WriteString8(Message)
End Sub

''
' Writes the "SystemMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to all players.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSystemMessage(ByVal Message As String)
    Call Writer.WriteInt(ClientPacketID.SystemMessage)
    Call Writer.WriteString8(Message)
End Sub

''
' Writes the "CreateNPC" message to the outgoing data buffer.
'
' @param    npcIndex The index of the NPC to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCreateNPC(ByVal NpcIndex As Integer)
    Call Writer.WriteInt(ClientPacketID.CreateNPC)
    Call Writer.WriteInt16(NpcIndex)
End Sub

''
' Writes the "CreateNPCWithRespawn" message to the outgoing data buffer.
'
' @param    npcIndex The index of the NPC to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCreateNPCWithRespawn(ByVal NpcIndex As Integer)
    Call Writer.WriteInt(ClientPacketID.CreateNPCWithRespawn)
    Call Writer.WriteInt16(NpcIndex)
End Sub

''
' Writes the "ImperialArmour" message to the outgoing data buffer.
'
' @param    armourIndex The index of imperial armour to be altered.
' @param    objectIndex The index of the new object to be set as the imperial armour.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteImperialArmour(ByVal armourIndex As Byte, ByVal objectIndex As Integer)
    Call Writer.WriteInt(ClientPacketID.ImperialArmour)
    Call Writer.WriteInt8(armourIndex)
    Call Writer.WriteInt16(objectIndex)
End Sub

''
' Writes the "ChaosArmour" message to the outgoing data buffer.
'
' @param    armourIndex The index of chaos armour to be altered.
' @param    objectIndex The index of the new object to be set as the chaos armour.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChaosArmour(ByVal armourIndex As Byte, ByVal objectIndex As Integer)
    Call Writer.WriteInt(ClientPacketID.ChaosArmour)
    Call Writer.WriteInt8(armourIndex)
    Call Writer.WriteInt16(objectIndex)
End Sub

''
' Writes the "NavigateToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteNavigateToggle()
    Call Writer.WriteInt(ClientPacketID.NavigateToggle)
End Sub

' Writes the "ServerOpenToUsersToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteServerOpenToUsersToggle()
    Call Writer.WriteInt(ClientPacketID.ServerOpenToUsersToggle)
End Sub

''
' Writes the "Participar" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteParticipar()
    Call Writer.WriteInt(ClientPacketID.Participar)
End Sub

''
' Writes the "TurnCriminal" message to the outgoing data buffer.
'
' @param    username The name of the user to turn into criminal.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteTurnCriminal(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.TurnCriminal)
    Call Writer.WriteString8(UserName)
End Sub

''
' Writes the "ResetFactions" message to the outgoing data buffer.
'
' @param    username The name of the user who will be removed from any faction.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteResetFactions(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.ResetFactions)
    Call Writer.WriteString8(UserName)
End Sub

''
' Writes the "RemoveCharFromGuild" message to the outgoing data buffer.
'
' @param    username The name of the user who will be removed from any guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRemoveCharFromGuild(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.RemoveCharFromGuild)
    Call Writer.WriteString8(UserName)
End Sub

''
' Writes the "RequestCharMail" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRequestCharMail(ByVal UserName As String)
    Call Writer.WriteInt(ClientPacketID.RequestCharMail)
    Call Writer.WriteString8(UserName)
End Sub

''
' Writes the "AlterPassword" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    copyFrom The name of the user from which to copy the password.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteAlterPassword(ByVal UserName As String, ByVal CopyFrom As String)
    Call Writer.WriteInt(ClientPacketID.AlterPassword)
    Call Writer.WriteString8(UserName)
    Call Writer.WriteString8(CopyFrom)
End Sub

''
' Writes the "AlterMail" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    newMail The new email of the player.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteAlterMail(ByVal UserName As String, ByVal newMail As String)
    Call Writer.WriteInt(ClientPacketID.AlterMail)
    Call Writer.WriteString8(UserName)
    Call Writer.WriteString8(newMail)
End Sub

''
' Writes the "AlterName" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    newName The new user name.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteAlterName(ByVal UserName As String, ByVal newName As String)
    Call Writer.WriteInt(ClientPacketID.AlterName)
    Call Writer.WriteString8(UserName)
    Call Writer.WriteString8(newName)
End Sub

''
' Writes the "DoBackup" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteDoBackup()
    Call Writer.WriteInt(ClientPacketID.DoBackUp)
End Sub

''
' Writes the "ShowGuildMessages" message to the outgoing data buffer.
'
' @param    guild The guild to listen to.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowGuildMessages(ByVal guild As String)
    Call Writer.WriteInt(ClientPacketID.ShowGuildMessages)
    Call Writer.WriteString8(guild)
End Sub

''
' Writes the "SaveMap" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSaveMap()
    Call Writer.WriteInt(ClientPacketID.SaveMap)
End Sub

''
' Writes the "ChangeMapInfoPK" message to the outgoing data buffer.
'
' @param    isPK True if the map is PK, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeMapInfoPK(ByVal isPK As Boolean)
    Call Writer.WriteInt(ClientPacketID.ChangeMapInfoPK)
    Call Writer.WriteBool(isPK)
End Sub

''
' Writes the "ChangeMapInfoBackup" message to the outgoing data buffer.
'
' @param    backup True if the map is to be backuped, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeMapInfoBackup(ByVal backup As Boolean)
    Call Writer.WriteInt(ClientPacketID.ChangeMapInfoBackup)
    Call Writer.WriteBool(backup)
End Sub

''
' Writes the "ChangeMapInfoRestricted" message to the outgoing data buffer.
'
' @param    restrict NEWBIES (only newbies), NO (everyone), ARMADA (just Armadas), CAOS (just caos) or FACCION (Armadas & caos only)
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeMapInfoRestricted(ByVal restrict As String)
    Call Writer.WriteInt(ClientPacketID.ChangeMapInfoRestricted)
    Call Writer.WriteString8(restrict)
End Sub

''
' Writes the "ChangeMapInfoNoMagic" message to the outgoing data buffer.
'
' @param    nomagic TRUE if no magic is to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeMapInfoNoMagic(ByVal nomagic As Boolean)
    Call Writer.WriteInt(ClientPacketID.ChangeMapInfoNoMagic)
    Call Writer.WriteBool(nomagic)
End Sub

''
' Writes the "ChangeMapInfoNoInvi" message to the outgoing data buffer.
'
' @param    noinvi TRUE if invisibility is not to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeMapInfoNoInvi(ByVal noinvi As Boolean)
    Call Writer.WriteInt(ClientPacketID.ChangeMapInfoNoInvi)
    Call Writer.WriteBool(noinvi)
End Sub
                            
''
' Writes the "ChangeMapInfoNoResu" message to the outgoing data buffer.
'
' @param    noresu TRUE if resurection is not to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeMapInfoNoResu(ByVal noresu As Boolean)
    Call Writer.WriteInt(ClientPacketID.ChangeMapInfoNoResu)
    Call Writer.WriteBool(noresu)
End Sub
                        
''
' Writes the "ChangeMapInfoLand" message to the outgoing data buffer.
'
' @param    land options: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeMapInfoLand(ByVal lAnd As String)
    Call Writer.WriteInt(ClientPacketID.ChangeMapInfoLand)
    Call Writer.WriteString8(lAnd)
End Sub
                        
''
' Writes the "ChangeMapInfoZone" message to the outgoing data buffer.
'
' @param    zone options: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeMapInfoZone(ByVal zone As String)
    Call Writer.WriteInt(ClientPacketID.ChangeMapInfoZone)
    Call Writer.WriteString8(zone)
End Sub

''
' Writes the "SaveChars" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSaveChars()
    Call Writer.WriteInt(ClientPacketID.SaveChars)
End Sub

''
' Writes the "CleanSOS" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCleanSOS()
    Call Writer.WriteInt(ClientPacketID.CleanSOS)
End Sub

''
' Writes the "ShowServerForm" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowServerForm()
    Call Writer.WriteInt(ClientPacketID.ShowServerForm)
End Sub

''
' Writes the "Night" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteNight()
    Call Writer.WriteInt(ClientPacketID.night)
End Sub

Public Sub WriteDay()
    Call Writer.WriteInt(ClientPacketID.Day)
End Sub

Public Sub WriteSetTime(ByVal Time As Long)
    Call Writer.WriteInt(ClientPacketID.SetTime)
    Call Writer.WriteInt32(Time)
End Sub

''
' Writes the "KickAllChars" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteKickAllChars()
    Call Writer.WriteInt(ClientPacketID.KickAllChars)
End Sub

''
' Writes the "RequestTCPStats" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRequestTCPStats()
    Call Writer.WriteInt(ClientPacketID.RequestTCPStats)
End Sub

''
' Writes the "ReloadNPCs" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteReloadNPCs()
    Call Writer.WriteInt(ClientPacketID.ReloadNPCs)
End Sub

''
' Writes the "ReloadServerIni" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteReloadServerIni()
    Call Writer.WriteInt(ClientPacketID.ReloadServerIni)
End Sub

''
' Writes the "ReloadSpells" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteReloadSpells()
    Call Writer.WriteInt(ClientPacketID.ReloadSpells)
End Sub

''
' Writes the "ReloadObjects" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteReloadObjects()
    Call Writer.WriteInt(ClientPacketID.ReloadObjects)
End Sub

''
' Writes the "Restart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRestart()
    Call Writer.WriteInt(ClientPacketID.Restart)
End Sub

''
' Writes the "ResetAutoUpdate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteResetAutoUpdate()
    Call Writer.WriteInt(ClientPacketID.ResetAutoUpdate)
End Sub

''
' Writes the "ChatColor" message to the outgoing data buffer.
'
' @param    r The red component of the new chat color.
' @param    g The green component of the new chat color.
' @param    b The blue component of the new chat color.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChatColor(ByVal r As Byte, ByVal G As Byte, ByVal B As Byte)
    Call Writer.WriteInt(ClientPacketID.ChatColor)
    Call Writer.WriteInt8(r)
    Call Writer.WriteInt8(G)
    Call Writer.WriteInt8(B)
End Sub

''
' Writes the "Ignored" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteIgnored()
    Call Writer.WriteInt(ClientPacketID.Ignored)
End Sub

''
' Writes the "CheckSlot" message to the outgoing data buffer.
'
' @param    UserName    The name of the char whose slot will be checked.
' @param    slot        The slot to be checked.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCheckSlot(ByVal UserName As String, ByVal Slot As Byte)
    Call Writer.WriteInt(ClientPacketID.CheckSlot)
    Call Writer.WriteString8(UserName)
    Call Writer.WriteInt8(Slot)
End Sub

''
' Writes the "Ping" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePing()
    Call Writer.WriteInt(ClientPacketID.Ping)
    pingTime = GetTickCount()
    Call Writer.WriteInt32(pingTime)
    
    ' Avoid computing errors due to frame rate
    Call modNetwork.Poll
End Sub

Public Sub WriteLlamadadeClan()
    Call Writer.WriteInt(ClientPacketID.llamadadeclan)
End Sub

Public Sub WriteQuestionGM(ByVal Consulta As String, ByVal TipoDeConsulta As String)
    Call Writer.WriteInt(ClientPacketID.QuestionGM)
    Call Writer.WriteString8(Consulta)
    Call Writer.WriteString8(TipoDeConsulta)
End Sub

Public Sub WriteOfertaInicial(ByVal Oferta As Long)
    Call Writer.WriteInt(ClientPacketID.OfertaInicial)
    Call Writer.WriteInt32(Oferta)
End Sub

Public Sub WriteOferta(ByVal OfertaDeSubasta As Long)
    Call Writer.WriteInt(ClientPacketID.OfertaDeSubasta)
    Call Writer.WriteInt32(OfertaDeSubasta)
End Sub

Public Sub WriteGlobalMessage(ByVal Message As String)
    Call Writer.WriteInt(ClientPacketID.GlobalMessage)
    Call Writer.WriteString8(Message)
End Sub

Public Sub WriteGlobalOnOff()
    Call Writer.WriteInt(ClientPacketID.GlobalOnOff)
End Sub

Public Sub WriteBorrandoPJ()
    'TODO_WOLF: Pure
    Call Writer.WriteInt(ClientPacketID.BorrarPJ)
    Call Writer.WriteString8(DeleteUser)
    Call Writer.WriteString8(CuentaEmail)
    Call Writer.WriteString8(SEncriptar(CuentaPassword))
    Call Writer.WriteInt8(App.Major)
    Call Writer.WriteInt8(App.Minor)
    Call Writer.WriteInt8(App.Revision)
    Call Writer.WriteString8(MacAdress)  'Seguridad
    Call Writer.WriteInt32(HDserial)  'SeguridadHDserial
    Call Writer.WriteString8(CheckMD5)
End Sub

Public Sub WriteIngresandoConCuenta()
    'TODO_WOLF: Pure
    Call Writer.WriteInt(ClientPacketID.IngresarConCuenta)
    Call Writer.WriteString8(CuentaEmail)
    Call Writer.WriteString8(SEncriptar(CuentaPassword))
    Call Writer.WriteInt8(App.Major)
    Call Writer.WriteInt8(App.Minor)
    Call Writer.WriteInt8(App.Revision)
    Call Writer.WriteString8(MacAdress)  'Seguridad
    Call Writer.WriteInt32(HDserial)  'SeguridadHDserial
    Call Writer.WriteString8(CheckMD5)
End Sub

Public Sub WriteNieblaToggle()
    Call Writer.WriteInt(ClientPacketID.NieblaToggle)
End Sub

Public Sub WriteGenio()
    Call Writer.WriteInt(ClientPacketID.Genio)
End Sub

Public Sub WriteTraerRecompensas()
    Call Writer.WriteInt(ClientPacketID.TraerRecompensas)
End Sub

Public Sub WriteTraerShop()
    Call Writer.WriteInt(ClientPacketID.Traershop)
End Sub

Public Sub WriteTraerRanking()
    Call Writer.WriteInt(ClientPacketID.TraerRanking)
End Sub

Public Sub WritePareja()
    Call Writer.WriteInt(ClientPacketID.Pareja)
End Sub

Public Sub WriteQuest()
    Call Writer.WriteInt(ClientPacketID.Quest)
End Sub
 
Public Sub WriteQuestDetailsRequest(ByVal QuestSlot As Byte)
    Call Writer.WriteInt(ClientPacketID.QuestDetailsRequest)
    Call Writer.WriteInt8(QuestSlot)
End Sub
 
Public Sub WriteQuestAccept(ByVal ListInd As Byte)
    Call Writer.WriteInt(ClientPacketID.QuestAccept)
    Call Writer.WriteInt8(ListInd)
End Sub

Public Sub WriteQuestListRequest()
    Call Writer.WriteInt(ClientPacketID.QuestListRequest)
End Sub
 
Public Sub WriteQuestAbandon(ByVal QuestSlot As Byte)
    Call Writer.WriteInt(ClientPacketID.QuestAbandon)
    'Escribe el Slot de Quest.
    Call Writer.WriteInt8(QuestSlot)
End Sub

Public Sub WriteResponderPregunta(ByVal Respuesta As Boolean)
    Call Writer.WriteInt(ClientPacketID.ResponderPregunta)
    Call Writer.WriteBool(Respuesta)
End Sub

Public Sub WriteCorreo()
    Call Writer.WriteInt(ClientPacketID.Correo)
End Sub

Public Sub WriteSendCorreo(ByVal UserNick As String, _
                           ByVal msg As String, _
                           ByVal ItemCount As Byte)
    Call Writer.WriteInt(ClientPacketID.SendCorreo)
    Call Writer.WriteString8(UserNick)
    Call Writer.WriteString8(msg)
    Call Writer.WriteInt8(ItemCount)

    If ItemCount > 0 Then

        Dim i As Byte

        For i = 1 To ItemCount
            Call Writer.WriteInt8(ItemLista(i).OBJIndex) ' Slot
            Call Writer.WriteInt16(ItemLista(i).Amount) 'Cantidad
        Next i

    End If

End Sub

Public Sub WriteComprarItem(ByVal ItemIndex As Byte)
    Call Writer.WriteInt(ClientPacketID.ComprarItem)
    Call Writer.WriteInt8(ItemIndex)
End Sub

Public Sub WriteCompletarViaje(ByVal destino As Byte, ByVal costo As Long)
    Call Writer.WriteInt(ClientPacketID.CompletarViaje)
    Call Writer.WriteInt8(destino)
    Call Writer.WriteInt32(costo)
End Sub

Public Sub WriteRetirarItemCorreo(ByVal IndexMsg As Integer)
    Call Writer.WriteInt(ClientPacketID.RetirarItemCorreo)
    Call Writer.WriteInt16(IndexMsg)
End Sub

Public Sub WriteBorrarCorreo(ByVal IndexMsg As Integer)
    Call Writer.WriteInt(ClientPacketID.BorrarCorreo)
    Call Writer.WriteInt16(IndexMsg)
End Sub

Public Sub WriteCodigo(ByVal Codigo As String)
    Call Writer.WriteInt(ClientPacketID.EnviarCodigo)
    Call Writer.WriteString8(Codigo)
End Sub

Public Sub WriteCreaerTorneo(ByVal nivelminimo As Byte, _
                             ByVal nivelmaximo As Byte, _
                             ByVal cupos As Byte, _
                             ByVal costo As Long, _
                             ByVal mago As Byte, _
                             ByVal clerico As Byte, _
                             ByVal guerrero As Byte, _
                             ByVal asesino As Byte, _
                             ByVal bardo As Byte, _
                             ByVal druido As Byte, _
                             ByVal paladin As Byte, _
                             ByVal cazador As Byte, _
                             ByVal Trabajador As Byte, _
                             ByVal Pirata As Byte, _
                             ByVal Ladron As Byte, _
                             ByVal Bandido As Byte, _
                             ByVal map As Integer, _
                             ByVal x As Byte, _
                             ByVal y As Byte, _
                             ByVal Name As String, _
                             ByVal reglas As String)
    Call Writer.WriteInt(ClientPacketID.CrearTorneo)
    Call Writer.WriteInt8(nivelminimo)
    Call Writer.WriteInt8(nivelmaximo)
    Call Writer.WriteInt8(cupos)
    Call Writer.WriteInt32(costo)
    Call Writer.WriteInt8(mago)
    Call Writer.WriteInt8(clerico)
    Call Writer.WriteInt8(guerrero)
    Call Writer.WriteInt8(asesino)
    Call Writer.WriteInt8(bardo)
    Call Writer.WriteInt8(druido)
    Call Writer.WriteInt8(paladin)
    Call Writer.WriteInt8(cazador)
    Call Writer.WriteInt8(Trabajador)
    Call Writer.WriteInt8(Pirata)
    Call Writer.WriteInt8(Ladron)
    Call Writer.WriteInt8(Bandido)
    Call Writer.WriteInt16(map)
    Call Writer.WriteInt8(x)
    Call Writer.WriteInt8(y)
    Call Writer.WriteString8(Name)
    Call Writer.WriteString8(reglas)
End Sub

Public Sub WriteComenzarTorneo()
    Call Writer.WriteInt(ClientPacketID.ComenzarTorneo)
End Sub

Public Sub WriteCancelarTorneo()
    Call Writer.WriteInt(ClientPacketID.CancelarTorneo)
End Sub

Public Sub WriteBusquedaTesoro(ByVal TIPO As Byte)
    Call Writer.WriteInt(ClientPacketID.BusquedaTesoro)
    Call Writer.WriteInt8(TIPO)
End Sub

''
' Writes the "Home" message to the outgoing data buffer.
'
Public Sub WriteHome()
    Call Writer.WriteInt(ClientPacketID.Home)
End Sub

''
' Writes the "Consulta" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteConsulta(Optional ByVal Nick As String = vbNullString)
    Call Writer.WriteInt(ClientPacketID.Consulta)
    Call Writer.WriteString8(Nick)
End Sub

Public Sub WriteCuentaExtractItem(ByVal Slot As Byte, _
                                  ByVal Amount As Integer, _
                                  ByVal slotdestino As Byte)
    Call Writer.WriteInt(ClientPacketID.CuentaExtractItem)
    Call Writer.WriteInt8(Slot)
    Call Writer.WriteInt16(Amount)
    Call Writer.WriteInt8(slotdestino)
End Sub

Public Sub WriteCuentaDeposit(ByVal Slot As Byte, _
                              ByVal Amount As Integer, _
                              ByVal slotdestino As Byte)
    Call Writer.WriteInt(ClientPacketID.CuentaDeposit)
    Call Writer.WriteInt8(Slot)
    Call Writer.WriteInt16(Amount)
    Call Writer.WriteInt8(slotdestino)
End Sub

Public Sub WriteDuel(Players As String, _
                     ByVal Apuesta As Long, _
                     Optional ByVal PocionesRojas As Long = -1, _
                     Optional ByVal CaenItems As Boolean = False)
    Call Writer.WriteInt(ClientPacketID.Duel)
    Call Writer.WriteString8(Players)
    Call Writer.WriteInt32(Apuesta)
    Call Writer.WriteInt16(PocionesRojas)
    Call Writer.WriteBool(CaenItems)
End Sub

Public Sub WriteAcceptDuel(Offerer As String)
    Call Writer.WriteInt(ClientPacketID.AcceptDuel)
    Call Writer.WriteString8(Offerer)
End Sub

Public Sub WriteCancelDuel()
    Call Writer.WriteInt(ClientPacketID.CancelDuel)
End Sub

Public Sub WriteQuitDuel()
    Call Writer.WriteInt(ClientPacketID.QuitDuel)
End Sub

Public Sub WriteCreateEvent(EventName As String)
    Call Writer.WriteInt(ClientPacketID.CreateEvent)
    Call Writer.WriteString8(EventName)
End Sub

Public Sub WriteCommerceSendChatMessage(ByVal Message As String)
    Call Writer.WriteInt(ClientPacketID.CommerceSendChatMessage)
    Call Writer.WriteString8(Message)
End Sub

Public Sub WriteLogMacroClickHechizo()
    Call Writer.WriteInt(ClientPacketID.LogMacroClickHechizo)
End Sub

Public Sub WriteNieveToggle()
    Call Writer.WriteInt(ClientPacketID.NieveToggle)
End Sub

Public Sub WriteCompletarAccion(ByVal Accion As Byte)
    Call Writer.WriteInt(ClientPacketID.CompletarAccion)
    Call Writer.WriteInt8(Accion)
End Sub

Public Sub WriteTolerancia0(Nick As String)
    Call Writer.WriteInt(ClientPacketID.Tolerancia0)
    Call Writer.WriteString8(Nick)
End Sub

Public Sub WriteGetMapInfo()
    Call Writer.WriteInt(ClientPacketID.GetMapInfo)
End Sub

Public Sub WriteAddItemCrafting(ByVal SlotInv As Byte, ByVal SlotCraft As Byte)
    Call Writer.WriteInt(ClientPacketID.AddItemCrafting)
    Call Writer.WriteInt8(SlotInv)
    Call Writer.WriteInt8(SlotCraft)
End Sub
    
Public Sub WriteRemoveItemCrafting(ByVal SlotCraft As Byte, ByVal SlotInv As Byte)
    Call Writer.WriteInt(ClientPacketID.RemoveItemCrafting)
    Call Writer.WriteInt8(SlotCraft)
    Call Writer.WriteInt8(SlotInv)
End Sub

Public Sub WriteAddCatalyst(ByVal SlotInv As Byte)
    Call Writer.WriteInt(ClientPacketID.AddCatalyst)
    Call Writer.WriteInt8(SlotInv)
End Sub

Public Sub WriteRemoveCatalyst(ByVal SlotInv As Byte)
    Call Writer.WriteInt(ClientPacketID.RemoveCatalyst)
    Call Writer.WriteInt8(SlotInv)
End Sub

Public Sub WriteCraftItem()
    Call Writer.WriteInt(ClientPacketID.CraftItem)
End Sub

Public Sub WriteMoveCraftItem(ByVal Drag As Byte, ByVal Drop As Byte)
    Call Writer.WriteInt(ClientPacketID.MoveCraftItem)
    Call Writer.WriteInt8(Drag)
    Call Writer.WriteInt8(Drop)
End Sub

Public Sub WriteCloseCrafting()
    Call Writer.WriteInt(ClientPacketID.CloseCrafting)
End Sub

Public Sub WritePetLeaveAll()
    Call Writer.WriteInt(ClientPacketID.PetLeaveAll)
End Sub

Public Sub WriteGuardNoticeResponse(ByVal Codigo As String)
    Call Writer.WriteInt(ClientPacketID.GuardNoticeResponse)
    Call Writer.WriteString8(Codigo)
End Sub

Public Sub WriteResendVerificationCode(ByVal Codigo As String)
    Call Writer.WriteInt(ClientPacketID.GuardResendVerificationCode)
End Sub
