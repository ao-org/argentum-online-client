Attribute VB_Name = "Protocol_Writes"
Option Explicit


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
        Call .WriteID(ClientPacketID.LoginExistingChar)
        Call .WriteASCIIString(CuentaEmail)
        Call .WriteASCIIString(SEncriptar(CuentaPassword))
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(MacAdress)  'Seguridad
        Call .WriteLong(HDserial)  'SeguridadHDserial
        Call .WriteASCIIString(CheckMD5)
        Call .EndPacket
    End With
    
    Exit Sub

WriteLoginExistingChar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteLoginExistingChar", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.LoginNewChar)
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
        Call .EndPacket
    End With
    
    Exit Sub

WriteLoginNewChar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteLoginNewChar", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.Talk)
        
        Call .WriteASCIIString(chat)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteTalk_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteTalk", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.Yell)
        
        Call .WriteASCIIString(chat)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteYell_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteYell", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.Whisper)
        
        Call .WriteASCIIString(nombre)
        Call .WriteASCIIString(chat)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteWhisper_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteWhisper", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.Walk)
        
        Call .WriteByte(Heading)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteWalk_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteWalk", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.RequestPositionUpdate)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteRequestPositionUpdate_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestPositionUpdate", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.Attack)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteAttack_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteAttack", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.PickUp)
    Call outgoingData.EndPacket
    
    Exit Sub

WritePickUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WritePickUp", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.SafeToggle)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteSafeToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSafeToggle", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteSeguroClan()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SafeToggle" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteSeguroClan_Err
    
    
    Call outgoingData.WriteByte(ClientPacketID.SeguroClan)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteSeguroClan_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSeguroClan", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteTraerBoveda()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "SafeToggle" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteTraerBoveda_Err
    
    
    Call outgoingData.WriteByte(ClientPacketID.TraerBoveda)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteTraerBoveda_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteTraerBoveda", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.CreatePretorianClan)
        Call .WriteInteger(map)
        Call .WriteByte(x)
        Call .WriteByte(y)
        Call .EndPacket
    End With
    
    Exit Sub

WriteCreatePretorianClan_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCreatePretorianClan", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.RemovePretorianClan)
        Call .WriteInteger(map)
        Call .EndPacket
        
    End With
    
    Exit Sub

WriteDeletePretorianClan_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteDeletePretorianClan", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.PartySafeToggle)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteParyToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteParyToggle", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    
    Call outgoingData.WriteByte(ClientPacketID.SeguroResu)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteSeguroResu_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSeguroResu", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.RequestGuildLeaderInfo)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteRequestGuildLeaderInfo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestGuildLeaderInfo", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.RequestAtributes)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteRequestAtributes_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestAtributes", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteRequestFamiliar()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestFamiliar" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteRequestFamiliar_Err
    
    
    Call outgoingData.WriteByte(ClientPacketID.RequestFamiliar)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteRequestFamiliar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestFamiliar", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteRequestGrupo()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "RequestFamiliar" message to the outgoing data buffer
    '***************************************************
    
    On Error GoTo WriteRequestGrupo_Err
    
    
    Call outgoingData.WriteByte(ClientPacketID.RequestGrupo)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteRequestGrupo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestGrupo", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.RequestSkills)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteRequestSkills_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestSkills", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.RequestMiniStats)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteRequestMiniStats_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestMiniStats", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.CommerceEnd)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteCommerceEnd_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCommerceEnd", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.UserCommerceEnd)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteUserCommerceEnd_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteUserCommerceEnd", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.BankEnd)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteBankEnd_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBankEnd", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.UserCommerceOk)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteUserCommerceOk_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteUserCommerceOk", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.UserCommerceReject)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteUserCommerceReject_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteUserCommerceReject", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.Drop)
        Call .WriteByte(Slot)
        Call .WriteLong(Amount)
        Call .EndPacket
    End With
    
    Exit Sub

WriteDrop_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteDrop", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.CastSpell)
        
        Call .WriteByte(Slot)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteCastSpell_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCastSpell", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteInvitarGrupo()
    
    On Error GoTo WriteInvitarGrupo_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CastSpell" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.InvitarGrupo)
        Call .EndPacket
    End With
    
    Exit Sub

WriteInvitarGrupo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteInvitarGrupo", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteMarcaDeClan()
    
    On Error GoTo WriteMarcaDeClan_Err

    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 23/08/2020
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.MarcaDeClanPack)
        Call .EndPacket
    End With
    
    Exit Sub

WriteMarcaDeClan_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteMarcaDeClan", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteMarcaDeGm()
    
    On Error GoTo WriteMarcaDeGm_Err

    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 23/08/2020
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.MarcaDeGMPack)
        Call .EndPacket
    End With
    
    Exit Sub

WriteMarcaDeGm_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteMarcaDeGm", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteAbandonarGrupo()
    
    On Error GoTo WriteAbandonarGrupo_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CastSpell" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.AbandonarGrupo)
        Call .EndPacket
    End With
    
    Exit Sub

WriteAbandonarGrupo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteAbandonarGrupo", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteHecharDeGrupo(ByVal indice As Byte)
    
    On Error GoTo WriteHecharDeGrupo_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CastSpell" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.HecharDeGrupo)
        Call .WriteByte(indice)
        Call .EndPacket
    End With
    
    Exit Sub

WriteHecharDeGrupo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteHecharDeGrupo", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.LeftClick)
        
        Call .WriteByte(x)
        Call .WriteByte(y)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteLeftClick_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteLeftClick", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.DoubleClick)
        
        Call .WriteByte(x)
        Call .WriteByte(y)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteDoubleClick_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteDoubleClick", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.Work)
        
        Call .WriteByte(Skill)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteWork_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteWork", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteThrowDice()
    
    On Error GoTo WriteThrowDice_Err
    
    Call outgoingData.WriteID(ClientPacketID.ThrowDice)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteThrowDice_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteThrowDice", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.UseSpellMacro)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteUseSpellMacro_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteUseSpellMacro", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.UseItem)
        Call .WriteByte(Slot)
        Call .EndPacket
    End With
    
    Exit Sub

WriteUseItem_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteUseItem", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.CraftBlacksmith)
        
        Call .WriteInteger(Item)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteCraftBlacksmith_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCraftBlacksmith", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.CraftCarpenter)
        
        Call .WriteInteger(Item)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteCraftCarpenter_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCraftCarpenter", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteCraftAlquimista(ByVal Item As Integer)
    
    On Error GoTo WriteCraftAlquimista_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CraftCarpenter" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.CraftAlquimista)
        Call .WriteInteger(Item)
        Call .EndPacket
    End With
    
    Exit Sub

WriteCraftAlquimista_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCraftAlquimista", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteCraftSastre(ByVal Item As Integer)
    
    On Error GoTo WriteCraftSastre_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "CraftCarpenter" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.CraftSastre)
        Call .WriteInteger(Item)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteCraftSastre_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCraftSastre", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.WorkLeftClick)
        
        Call .WriteByte(x)
        Call .WriteByte(y)
        Call .WriteByte(Skill)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteWorkLeftClick_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteWorkLeftClick", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.CreateNewGuild)
        
        Call .WriteASCIIString(desc)
        Call .WriteASCIIString(Name)
        Call .WriteByte(Alineacion)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteCreateNewGuild_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCreateNewGuild", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.SpellInfo)
        
        Call .WriteByte(Slot)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteSpellInfo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSpellInfo", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.EquipItem)
        
        Call .WriteByte(Slot)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteEquipItem_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteEquipItem", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.ChangeHeading)
        
        Call .WriteByte(Heading)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteChangeHeading_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteChangeHeading", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.ModifySkills)
        
        For i = 1 To NUMSKILLS
            Call .WriteByte(skillEdt(i))
        Next i
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteModifySkills_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteModifySkills", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.Train)
        
        Call .WriteByte(creature)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteTrain_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteTrain", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.CommerceBuy)
        
        Call .WriteByte(Slot)
        Call .WriteInteger(Amount)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteCommerceBuy_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCommerceBuy", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteUseKey(ByVal Slot As Byte)
    
    On Error GoTo WriteUseKey_Err

    With outgoingData
        Call .WriteID(ClientPacketID.UseKey)
        Call .WriteByte(Slot)
        Call .EndPacket
    End With
    
    Exit Sub

WriteUseKey_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteUseKey", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.BankExtractItem)
        
        Call .WriteByte(Slot)
        Call .WriteInteger(Amount)
        Call .WriteByte(slotdestino)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteBankExtractItem_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBankExtractItem", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.CommerceSell)
        
        Call .WriteByte(Slot)
        Call .WriteInteger(Amount)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteCommerceSell_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCommerceSell", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.BankDeposit)
        Call .WriteByte(Slot)
        Call .WriteInteger(Amount)
        Call .WriteByte(slotdestino)
        Call .EndPacket
    End With
    
    Exit Sub

WriteBankDeposit_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBankDeposit", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.ForumPost)
        
        Call .WriteASCIIString(title)
        Call .WriteASCIIString(Message)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteForumPost_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteForumPost", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.MoveSpell)
        
        Call .WriteBoolean(upwards)
        Call .WriteByte(Slot)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteMoveSpell_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteMoveSpell", Erl)
    Call incomingData.SafeClearPacket
    
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

    With outgoingData
        Call .WriteID(ClientPacketID.ClanCodexUpdate)
        Call .WriteASCIIString(desc)
        Call .EndPacket
    End With
    
    Exit Sub

WriteClanCodexUpdate_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteClanCodexUpdate", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.UserCommerceOffer)
        
        Call .WriteByte(Slot)
        Call .WriteLong(Amount)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteUserCommerceOffer_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteUserCommerceOffer", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.GuildAcceptPeace)
        
        Call .WriteASCIIString(guild)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteGuildAcceptPeace_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildAcceptPeace", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.GuildRejectAlliance)
        
        Call .WriteASCIIString(guild)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteGuildRejectAlliance_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildRejectAlliance", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.GuildRejectPeace)
        
        Call .WriteASCIIString(guild)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteGuildRejectPeace_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildRejectPeace", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.GuildAcceptAlliance)
        
        Call .WriteASCIIString(guild)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteGuildAcceptAlliance_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildAcceptAlliance", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.GuildOfferPeace)
        
        Call .WriteASCIIString(guild)
        Call .WriteASCIIString(proposal)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteGuildOfferPeace_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildOfferPeace", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.GuildOfferAlliance)
        
        Call .WriteASCIIString(guild)
        Call .WriteASCIIString(proposal)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteGuildOfferAlliance_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildOfferAlliance", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.GuildAllianceDetails)
        
        Call .WriteASCIIString(guild)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteGuildAllianceDetails_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildAllianceDetails", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.GuildPeaceDetails)
        
        Call .WriteASCIIString(guild)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteGuildPeaceDetails_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildPeaceDetails", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.GuildRequestJoinerInfo)
        
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteGuildRequestJoinerInfo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildRequestJoinerInfo", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.GuildAlliancePropList)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteGuildAlliancePropList_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildAlliancePropList", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.GuildPeacePropList)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteGuildPeacePropList_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildPeacePropList", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.GuildDeclareWar)
        
        Call .WriteASCIIString(guild)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteGuildDeclareWar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildDeclareWar", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.GuildNewWebsite)
        
        Call .WriteASCIIString(url)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteGuildNewWebsite_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildNewWebsite", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.GuildAcceptNewMember)
        
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteGuildAcceptNewMember_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildAcceptNewMember", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.GuildRejectNewMember)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(reason)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteGuildRejectNewMember_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildRejectNewMember", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.GuildKickMember)
        
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteGuildKickMember_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildKickMember", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.GuildUpdateNews)
        
        Call .WriteASCIIString(news)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteGuildUpdateNews_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildUpdateNews", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.GuildMemberInfo)
        
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteGuildMemberInfo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildMemberInfo", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.GuildOpenElections)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteGuildOpenElections_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildOpenElections", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.GuildRequestMembership)
        
        Call .WriteASCIIString(guild)
        Call .WriteASCIIString(Application)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteGuildRequestMembership_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildRequestMembership", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.GuildRequestDetails)
        
        Call .WriteASCIIString(guild)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteGuildRequestDetails_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildRequestDetails", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.Online)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteOnline_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteOnline", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.Quit)
    UserSaliendo = True
    Call outgoingData.EndPacket

    Exit Sub

WriteQuit_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteQuit", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.GuildLeave)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteGuildLeave_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildLeave", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.RequestAccountState)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteRequestAccountState_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestAccountState", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.PetStand)
    Call outgoingData.EndPacket
    
    Exit Sub

WritePetStand_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WritePetStand", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.PetFollow)
    Call outgoingData.EndPacket
    
    Exit Sub

WritePetStand_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WritePetFollow", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.PetLeave)
    Call outgoingData.EndPacket
    
    Exit Sub

WritePetStand_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WritePetLeave", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Writes the "GrupoMsg" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteGrupoMsg(ByVal Message As String)
    
    On Error GoTo WriteGrupoMsg_Err

    With outgoingData
        Call .WriteID(ClientPacketID.GrupoMsg)
        
        Call .WriteASCIIString(Message)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteGrupoMsg_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGrupoMsg", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.TrainList)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteTrainList_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteTrainList", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.Rest)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteRest_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRest", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.Meditate)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteMeditate_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteMeditate", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.Resucitate)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteResucitate_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteResucitate", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.Heal)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteHeal_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteHeal", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.Help)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteHelp_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteHelp", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.RequestStats)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteRequestStats_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestStats", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.Promedio)
    Call outgoingData.EndPacket
    
    Exit Sub

Handle:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WritePromedio", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.GiveItem)
        Call .WriteASCIIString(UserName)
        Call .WriteInteger(OBJIndex)
        Call .WriteInteger(cantidad)
        Call .WriteASCIIString(Motivo)
        Call .EndPacket
    End With

    Exit Sub

Handle:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGiveItem", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.CommerceStart)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteCommerceStart_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCommerceStart", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.BankStart)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteBankStart_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBankStart", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.Enlist)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteEnlist_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteEnlist", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.Information)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteInformation_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteInformation", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.Reward)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteReward_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteReward", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.RequestMOTD)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteRequestMOTD_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestMOTD", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.UpTime)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteUpTime_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteUpTime", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.Inquiry)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteInquiry_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteInquiry", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.GuildMessage)
        
        Call .WriteASCIIString(Message)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteGuildMessage_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildMessage", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.CentinelReport)
        
        Call .WriteInteger(Number)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteCentinelReport_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCentinelReport", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.GuildOnline)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteGuildOnline_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildOnline", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.CouncilMessage)
        
        Call .WriteASCIIString(Message)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteCouncilMessage_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCouncilMessage", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.RoleMasterRequest)
        
        Call .WriteASCIIString(Message)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteRoleMasterRequest_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRoleMasterRequest", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.GMRequest)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteGMRequest_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGMRequest", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.ChangeDescription)
        
        Call .WriteASCIIString(desc)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteChangeDescription_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteChangeDescription", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.GuildVote)
        
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteGuildVote_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildVote", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.punishments)
        
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WritePunishments_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WritePunishments", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.ChangePassword)

        Call .WriteASCIIString(SEncriptar(oldPass))
        Call .WriteASCIIString(SEncriptar(newPass))
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteChangePassword_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteChangePassword", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.Gamble)
        
        Call .WriteInteger(Amount)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteGamble_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGamble", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.InquiryVote)
        
        Call .WriteByte(opt)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteInquiryVote_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteInquiryVote", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.LeaveFaction)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteLeaveFaction_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteLeaveFaction", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.BankExtractGold)
        
        Call .WriteLong(Amount)
        Call .EndPacket
    End With
    
    Exit Sub

WriteBankExtractGold_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBankExtractGold", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.BankDepositGold)
        
        Call .WriteLong(Amount)
        Call .EndPacket
    End With
    
    Exit Sub

WriteBankDepositGold_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBankDepositGold", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteTransFerGold(ByVal Amount As Long, ByVal destino As String)
    
    On Error GoTo WriteTransFerGold_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BankDepositGold" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.TransFerGold)
        Call .WriteLong(Amount)
        Call .WriteASCIIString(destino)
        Call .EndPacket
    End With
    
    Exit Sub

WriteTransFerGold_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteTransFerGold", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteItemMove(ByVal SlotActual As Byte, ByVal SlotNuevo As Byte)
    
    On Error GoTo WriteItemMove_Err

    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.Moveitem)
        Call .WriteByte(SlotActual)
        Call .WriteByte(SlotNuevo)
        Call .EndPacket
    End With
    
    Exit Sub

WriteItemMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteItemMove", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteBovedaItemMove(ByVal SlotActual As Byte, ByVal SlotNuevo As Byte)
    
    On Error GoTo WriteBovedaItemMove_Err

    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.BovedaMoveItem)
        Call .WriteByte(SlotActual)
        Call .WriteByte(SlotNuevo)
        Call .EndPacket
    End With
    
    Exit Sub

WriteBovedaItemMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBovedaItemMove", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.FinEvento)
        Call .EndPacket
    End With
    
    Exit Sub

WriteFinEvento_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteFinEvento", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.Denounce)
        Call .WriteASCIIString(Name)
        Call .EndPacket
    End With
    
    Exit Sub

WriteDenounce_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteDenounce", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteQuieroFundarClan()
    
    On Error GoTo WriteQuieroFundarClan_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Denounce" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.QuieroFundarClan)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteQuieroFundarClan_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteQuieroFundarClan", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.GuildMemberList)
        
        Call .WriteASCIIString(guild)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteGuildMemberList_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildMemberList", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

'ladder
Public Sub WriteCasamiento(ByVal UserName As String)
    
    On Error GoTo WriteCasamiento_Err

    '***************************************************
    'Ladder
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.Casarse)
        Call .WriteASCIIString(UserName)
        Call .EndPacket
    End With
    
    Exit Sub

WriteCasamiento_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCasamiento", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteMacroPos()
    
    On Error GoTo WriteMacroPos_Err

    '***************************************************
    'Ladder
    'Macros
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.MacroPossent)
        Call .WriteByte(ChatCombate)
        Call .WriteByte(ChatGlobal)
        Call .EndPacket
    End With
    
    Exit Sub

WriteMacroPos_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteMacroPos", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteSubastaInfo()
    
    On Error GoTo WriteSubastaInfo_Err

    '***************************************************
    'Ladder
    'Macros
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.SubastaInfo)
        Call .EndPacket
    End With
    
    Exit Sub

WriteSubastaInfo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSubastaInfo", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteScrollInfo()
    
    On Error GoTo WriteScrollInfo_Err

    '***************************************************
    'Ladder
    'Macros
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.SCROLLINFO)
        Call .EndPacket
    End With
    
    Exit Sub

WriteScrollInfo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteScrollInfo", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteCancelarExit()
    '***************************************************
    'Ladder
    'Cancelar Salida
    '***************************************************
    
    On Error GoTo WriteCancelarExit_Err
    
    UserSaliendo = False

    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.CancelarExit)
        Call .EndPacket
    End With
    
    Exit Sub

WriteCancelarExit_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCancelarExit", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteEventoInfo()
    
    On Error GoTo WriteEventoInfo_Err

    '***************************************************
    'Ladder
    'Macros
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.EventoInfo)
        Call .EndPacket
    End With
    
    Exit Sub

WriteEventoInfo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteEventoInfo", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteFlagTrabajar()
    
    On Error GoTo WriteFlagTrabajar_Err

    '***************************************************
    'Ladder
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.FlagTrabajar)
        Call .EndPacket
    End With
    
    Exit Sub

WriteFlagTrabajar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteFlagTrabajar", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.Escribiendo)
        Call .EndPacket
    End With
    
    Exit Sub

WriteEscribiendo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteEscribiendo", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteReclamarRecompensa(ByVal Index As Byte)
    
    On Error GoTo WriteReclamarRecompensa_Err

    '***************************************************
    'Ladder
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.ReclamarRecompensa)
        Call .WriteByte(Index)
        Call .EndPacket
    End With
    
    Exit Sub

WriteReclamarRecompensa_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteReclamarRecompensa", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteGMMessage(ByVal Message As String)
    
    On Error GoTo WriteGMMessage_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GMMessage" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.GMMessage)
        
        Call .WriteASCIIString(Message)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteGMMessage_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGMMessage", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.showName)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteShowName_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteShowName", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.OnlineRoyalArmy)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteOnlineRoyalArmy_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteOnlineRoyalArmy", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.OnlineChaosLegion)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteOnlineChaosLegion_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteOnlineChaosLegion", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.GoNearby)
        
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteGoNearby_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGoNearby", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.comment)
        
        Call .WriteASCIIString(Message)
    
        Call .EndPacket
    End With
    
    Exit Sub

WriteComment_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteComment", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.serverTime)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteServerTime_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteServerTime", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.Where)
        
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteWhere_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteWhere", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.CreaturesInMap)
        
        Call .WriteInteger(map)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteCreaturesInMap_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCreaturesInMap", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.WarpMeToTarget)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteWarpMeToTarget_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteWarpMeToTarget", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.WarpChar)
        
        Call .WriteASCIIString(UserName)
        Call .WriteInteger(map)
        Call .WriteByte(x)
        Call .WriteByte(y)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteWarpChar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteWarpChar", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

''
' Writes the "Silence" message to the outgoing data buffer.
'
' @param    username The user to silence.
' @remarks  The data is not actually sent until the buffer is properly flushed.

Public Sub WriteSilence(ByVal UserName As String, ByVal Minutos As Integer)
    
    On Error GoTo WriteSilence_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Silence" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.Silence)
        
        Call .WriteASCIIString(UserName)
        Call .WriteInteger(Minutos)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteSilence_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSilence", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteCuentaRegresiva(ByVal Second As Byte)
    
    On Error GoTo WriteCuentaRegresiva_Err

    '***************************************************
    'Writer by Ladder
    '/Cuentaregresiva <Segundos>
    '04-12-08
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.CuentaRegresiva)
        Call .WriteByte(Second)
        Call .EndPacket
    End With
    
    Exit Sub

WriteCuentaRegresiva_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCuentaRegresiva", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.PossUser)
        Call .WriteASCIIString(UserName)
        Call .EndPacket
    End With
    
    Exit Sub

WritePossUser_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WritePossUser", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.SOSShowList)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteSOSShowList_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSOSShowList", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.SOSRemove)
        
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
        
    End With
    
    Exit Sub

WriteSOSRemove_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSOSRemove", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.GoToChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteGoToChar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGoToChar", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteDesbuggear(ByVal Params As String)
    
    On Error GoTo WriteDesbuggear_Err

    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.Desbuggear)
        Call .WriteASCIIString(Params)
        Call .EndPacket
    End With
    
    Exit Sub

WriteDesbuggear_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteDesbuggear", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteDarLlaveAUsuario(ByVal User As String, ByVal Llave As Integer)
    
    On Error GoTo WriteDarLlaveAUsuario_Err

    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.DarLlaveAUsuario)
        Call .WriteASCIIString(User)
        Call .WriteInteger(Llave)
        Call .EndPacket
    End With
    
    Exit Sub

WriteDarLlaveAUsuario_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteDarLlaveAUsuario", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteSacarLlave(ByVal Llave As Integer)
    
    On Error GoTo WriteSacarLlave_Err

    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.SacarLlave)
        Call .WriteInteger(Llave)
        Call .EndPacket
    End With
    
    Exit Sub

WriteSacarLlave_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSacarLlave", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteVerLlaves()
    
    On Error GoTo WriteVerLlaves_Err

    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.VerLlaves)
        Call .EndPacket
    End With
    
    Exit Sub

WriteVerLlaves_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteVerLlaves", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.Invisible)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteInvisible_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteInvisible", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.GMPanel)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteGMPanel_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGMPanel", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.RequestUserList)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteRequestUserList_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestUserList", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.Working)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteWorking_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteWorking", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.Hiding)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteHiding_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteHiding", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.Jail)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(reason)
        Call .WriteByte(Time)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteJail_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteJail", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteCrearEvento(ByVal TIPO As Byte, ByVal duracion As Byte, ByVal multiplicacion As Byte)
    
    On Error GoTo WriteCrearEvento_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "Jail" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.CrearEvento)
        
        Call .WriteByte(TIPO)
        Call .WriteByte(duracion)
        
        Call .WriteByte(multiplicacion)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteCrearEvento_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCrearEvento", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.KillNPC)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteKillNPC_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteKillNPC", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.WarnUser)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(reason)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteWarnUser_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteWarnUser", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteMensajeUser(ByVal UserName As String, ByVal mensaje As String)
    
    On Error GoTo WriteMensajeUser_Err

    '***************************************************
    'Author: Ladder
    'Last Modification: 04/jun/2014
    'Escribe un mensaje al usuario
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.MensajeUser)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(mensaje)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteMensajeUser_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteMensajeUser", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.EditChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .WriteByte(editOption)
        
        Call .WriteASCIIString(arg1)
        Call .WriteASCIIString(arg2)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteEditChar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteEditChar", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.RequestCharInfo)
        
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteRequestCharInfo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestCharInfo", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.RequestCharStats)
        
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteRequestCharStats_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestCharStats", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.RequestCharGold)
        
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteRequestCharGold_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestCharGold", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.RequestCharInventory)
        
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteRequestCharInventory_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestCharInventory", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.RequestCharBank)
        
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteRequestCharBank_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestCharBank", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.RequestCharSkills)
        
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteRequestCharSkills_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestCharSkills", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.ReviveChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteReviveChar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteReviveChar", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.OnlineGM)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteOnlineGM_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteOnlineGM", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.OnlineMap)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteOnlineMap_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteOnlineMap", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.Forgive)
        Call .EndPacket
    End With
    
    Exit Sub

WriteForgive_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteForgive", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteDonateGold(ByVal oro As Long)
    
    On Error GoTo WriteForgive_Err

    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.DonateGold)
        Call .WriteLong(oro)
        Call .EndPacket
    End With
    
    Exit Sub

WriteForgive_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteDonateGold", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.Kick)
        
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteKick_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteKick", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.Execute)
        
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteExecute_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteExecute", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.BanChar)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(reason)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteBanChar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBanChar", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteBanCuenta(ByVal UserName As String, ByVal reason As String)
    
    On Error GoTo WriteBanCuenta_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BanCuenta" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.BanCuenta)
    
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(reason)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteBanCuenta_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBanCuenta", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteUnBanCuenta(ByVal UserName As String)
    
    On Error GoTo WriteUnBanCuenta_Err

    '***************************************************
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.UnbanCuenta)
        Call .WriteASCIIString(UserName)
        Call .EndPacket
    End With
    
    Exit Sub

WriteUnBanCuenta_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteUnBanCuenta", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteBanSerial(ByVal UserName As String)
    
    On Error GoTo WriteBanSerial_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BanCuenta" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.BanSerial)
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteBanSerial_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBanSerial", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteUnBanSerial(ByVal UserName As String, ByVal reason As String)
    
    On Error GoTo WriteUnBanSerial_Err

    '***************************************************
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.unBanSerial)
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteUnBanSerial_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteUnBanSerial", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteCerraCliente(ByVal UserName As String)
    
    On Error GoTo WriteCerraCliente_Err

    '***************************************************
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.CerrarCliente)
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteCerraCliente_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCerraCliente", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteBanTemporal(ByVal UserName As String, ByVal reason As String, ByVal dias As Byte)
    
    On Error GoTo WriteBanTemporal_Err

    '***************************************************
    'Writes the "BanTemporal" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.BanTemporal)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(reason)
        Call .WriteByte(dias)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteBanTemporal_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBanTemporal", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.UnbanChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteUnbanChar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteUnbanChar", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.NPCFollow)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteNPCFollow_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteNPCFollow", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.SummonChar)
        
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteSummonChar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSummonChar", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.SpawnListRequest)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteSpawnListRequest_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSpawnListRequest", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.SpawnCreature)
        
        Call .WriteInteger(creatureIndex)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteSpawnCreature_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSpawnCreature", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.ResetNPCInventory)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteResetNPCInventory_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteResetNPCInventory", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.CleanWorld)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteCleanWorld_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCleanWorld", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.ServerMessage)
        
        Call .WriteASCIIString(Message)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteServerMessage_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteServerMessage", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.NickToIP)
        
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteNickToIP_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteNickToIP", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.IPToNick)
        
        For i = LBound(IP()) To UBound(IP())
            Call .WriteByte(IP(i))
        Next i
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteIPToNick_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteIPToNick", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.GuildOnlineMembers)
        
        Call .WriteASCIIString(guild)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteGuildOnlineMembers_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildOnlineMembers", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.TeleportCreate)
        
        Call .WriteInteger(map)
        Call .WriteByte(x)
        Call .WriteByte(y)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteTeleportCreate_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteTeleportCreate", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.TeleportDestroy)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteTeleportDestroy_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteTeleportDestroy", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.RainToggle)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteRainToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRainToggle", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.SetCharDescription)
        
        Call .WriteASCIIString(desc)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteSetCharDescription_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSetCharDescription", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.ForceMIDIToMap)
        
        Call .WriteByte(midiID)
        Call .WriteInteger(map)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteForceMIDIToMap_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteForceMIDIToMap", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.ForceWAVEToMap)
        
        Call .WriteByte(waveID)
        Call .WriteInteger(map)
        Call .WriteByte(x)
        Call .WriteByte(y)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteForceWAVEToMap_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteForceWAVEToMap", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.RoyalArmyMessage)
        
        Call .WriteASCIIString(Message)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteRoyalArmyMessage_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRoyalArmyMessage", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.ChaosLegionMessage)
        
        Call .WriteASCIIString(Message)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteChaosLegionMessage_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteChaosLegionMessage", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.CitizenMessage)
        
        Call .WriteASCIIString(Message)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteCitizenMessage_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCitizenMessage", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.CriminalMessage)
        
        Call .WriteASCIIString(Message)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteCriminalMessage_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCriminalMessage", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.TalkAsNPC)
        
        Call .WriteASCIIString(Message)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteTalkAsNPC_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteTalkAsNPC", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.DestroyAllItemsInArea)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteDestroyAllItemsInArea_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteDestroyAllItemsInArea", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.AcceptRoyalCouncilMember)
        
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteAcceptRoyalCouncilMember_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteAcceptRoyalCouncilMember", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.AcceptChaosCouncilMember)
        
        Call .WriteASCIIString(UserName)
    
        Call .EndPacket
        
    End With
    
    Exit Sub

WriteAcceptChaosCouncilMember_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteAcceptChaosCouncilMember", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.ItemsInTheFloor)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteItemsInTheFloor_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteItemsInTheFloor", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.MakeDumb)
        
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteMakeDumb_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteMakeDumb", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.MakeDumbNoMore)
        
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteMakeDumbNoMore_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteMakeDumbNoMore", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.DumpIPTables)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteDumpIPTables_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteDumpIPTables", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.CouncilKick)
        
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteCouncilKick_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCouncilKick", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.SetTrigger)
        
        Call .WriteByte(Trigger)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteSetTrigger_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSetTrigger", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.AskTrigger)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteAskTrigger_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteAskTrigger", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.BannedIPList)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteBannedIPList_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBannedIPList", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.BannedIPReload)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteBannedIPReload_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBannedIPReload", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.GuildBan)
        
        Call .WriteASCIIString(guild)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteGuildBan_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGuildBan", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    On Error GoTo WriteBanIP_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "BanIP" message to the outgoing data buffer
    '***************************************************

    With outgoingData
        Call .WriteID(ClientPacketID.banip)

        Call .WriteASCIIString(NickOrIP)
        Call .WriteASCIIString(reason)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteBanIP_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBanIP", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.UnBanIp)
        
        For i = LBound(IP()) To UBound(IP())
            Call .WriteByte(IP(i))
        Next i
        
        Call .EndPacket
    End With
 
    Exit Sub

WriteUnbanIP_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteUnbanIP", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.CreateItem)
        
        Call .WriteInteger(ItemIndex)
        Call .WriteInteger(cantidad)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteCreateItem_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCreateItem", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.DestroyItems)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteDestroyItems_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteDestroyItems", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.ChaosLegionKick)
        
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteChaosLegionKick_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteChaosLegionKick", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.RoyalArmyKick)
        
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteRoyalArmyKick_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRoyalArmyKick", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.ForceMIDIAll)
        
        Call .WriteByte(midiID)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteForceMIDIAll_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteForceMIDIAll", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.ForceWAVEAll)
        
        Call .WriteByte(waveID)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteForceWAVEAll_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteForceWAVEAll", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.RemovePunishment)
        
        Call .WriteASCIIString(UserName)
        Call .WriteByte(punishment)
        Call .WriteASCIIString(NewText)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteRemovePunishment_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRemovePunishment", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.TileBlockedToggle)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteTileBlockedToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteTileBlockedToggle", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.KillNPCNoRespawn)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteKillNPCNoRespawn_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteKillNPCNoRespawn", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.KillAllNearbyNPCs)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteKillAllNearbyNPCs_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteKillAllNearbyNPCs", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.LastIP)
        
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteLastIP_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteLastIP", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.ChangeMOTD)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteChangeMOTD_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteChangeMOTD", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.SetMOTD)
        
        Call .WriteASCIIString(Message)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteSetMOTD_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSetMOTD", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.SystemMessage)
        
        Call .WriteASCIIString(Message)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteSystemMessage_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSystemMessage", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.CreateNPC)
        
        Call .WriteInteger(NpcIndex)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteCreateNPC_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCreateNPC", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.CreateNPCWithRespawn)
        
        Call .WriteInteger(NpcIndex)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteCreateNPCWithRespawn_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCreateNPCWithRespawn", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.ImperialArmour)
        
        Call .WriteByte(armourIndex)
        Call .WriteInteger(objectIndex)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteImperialArmour_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteImperialArmour", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.ChaosArmour)
        
        Call .WriteByte(armourIndex)
        Call .WriteInteger(objectIndex)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteChaosArmour_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteChaosArmour", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.NavigateToggle)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteNavigateToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteNavigateToggle", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.ServerOpenToUsersToggle)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteServerOpenToUsersToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteServerOpenToUsersToggle", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.Participar)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteParticipar_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteParticipar", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.TurnCriminal)
        
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteTurnCriminal_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteTurnCriminal", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.ResetFactions)
        
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteResetFactions_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteResetFactions", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.RemoveCharFromGuild)
        
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteRemoveCharFromGuild_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRemoveCharFromGuild", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.RequestCharMail)
        
        Call .WriteASCIIString(UserName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteRequestCharMail_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestCharMail", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.AlterPassword)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(CopyFrom)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteAlterPassword_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteAlterPassword", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.AlterMail)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(newMail)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteAlterMail_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteAlterMail", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.AlterName)
        
        Call .WriteASCIIString(UserName)
        Call .WriteASCIIString(newName)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteAlterName_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteAlterName", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.DoBackUp)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteDoBackup_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteDoBackup", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.ShowGuildMessages)
        
        Call .WriteASCIIString(guild)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteShowGuildMessages_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteShowGuildMessages", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.SaveMap)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteSaveMap_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSaveMap", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.ChangeMapInfoPK)
        
        Call .WriteBoolean(isPK)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteChangeMapInfoPK_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteChangeMapInfoPK", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.ChangeMapInfoBackup)
        
        Call .WriteBoolean(backup)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteChangeMapInfoBackup_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteChangeMapInfoBackup", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.ChangeMapInfoRestricted)
        
        Call .WriteASCIIString(restrict)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteChangeMapInfoRestricted_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteChangeMapInfoRestricted", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.ChangeMapInfoNoMagic)
        
        Call .WriteBoolean(nomagic)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteChangeMapInfoNoMagic_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteChangeMapInfoNoMagic", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.ChangeMapInfoNoInvi)
        
        Call .WriteBoolean(noinvi)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteChangeMapInfoNoInvi_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteChangeMapInfoNoInvi", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.ChangeMapInfoNoResu)
        
        Call .WriteBoolean(noresu)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteChangeMapInfoNoResu_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteChangeMapInfoNoResu", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.ChangeMapInfoLand)
        
        Call .WriteASCIIString(lAnd)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteChangeMapInfoLand_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteChangeMapInfoLand", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.ChangeMapInfoZone)
        
        Call .WriteASCIIString(zone)
        
        Call .EndPacket
    End With
    
    Exit Sub

WriteChangeMapInfoZone_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteChangeMapInfoZone", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.SaveChars)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteSaveChars_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSaveChars", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.CleanSOS)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteCleanSOS_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCleanSOS", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.ShowServerForm)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteShowServerForm_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteShowServerForm", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.night)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteNight_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteNight", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteDay()
    
    On Error GoTo WriteDay_Err
    
    Call outgoingData.WriteID(ClientPacketID.Day)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteDay_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteDay", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteSetTime(ByVal Time As Long)
    
    On Error GoTo WriteSetTime_Err
    
    With outgoingData
        Call .WriteID(ClientPacketID.SetTime)
        Call .WriteLong(Time)
        Call .EndPacket
    End With
    
    Exit Sub

WriteSetTime_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSetTime", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.KickAllChars)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteKickAllChars_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteKickAllChars", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.RequestTCPStats)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteRequestTCPStats_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRequestTCPStats", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.ReloadNPCs)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteReloadNPCs_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteReloadNPCs", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.ReloadServerIni)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteReloadServerIni_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteReloadServerIni", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.ReloadSpells)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteReloadSpells_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteReloadSpells", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.ReloadObjects)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteReloadObjects_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteReloadObjects", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.Restart)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteRestart_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRestart", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.ResetAutoUpdate)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteResetAutoUpdate_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteResetAutoUpdate", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.ChatColor)
        
        Call .WriteByte(r)
        Call .WriteByte(G)
        Call .WriteByte(B)
        
        Call .EndPacket
        
    End With
    
    Exit Sub

WriteChatColor_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteChatColor", Erl)
    Call incomingData.SafeClearPacket
    
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
    
    Call outgoingData.WriteID(ClientPacketID.Ignored)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteIgnored_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteIgnored", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.CheckSlot)
        Call .WriteASCIIString(UserName)
        Call .WriteByte(Slot)
        Call .EndPacket
    End With
    
    Exit Sub

WriteCheckSlot_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCheckSlot", Erl)
    Call incomingData.SafeClearPacket
    
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

    Call outgoingData.WriteID(ClientPacketID.Ping)
    pingTime = timeGetTime And &H7FFFFFFF
    Call outgoingData.WriteLong(pingTime)
    Call outgoingData.EndPacket
    
    ' Avoid computing errors due to frame rate
    Call FlushBuffer
    'DoEvents
    
    Exit Sub

WritePing_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WritePing", Erl)
    Call incomingData.SafeClearPacket
    
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

    
    Call outgoingData.WriteByte(ClientPacketID.llamadadeclan)
    Call outgoingData.EndPacket
    
    ' Avoid computing errors due to frame rate
    Call FlushBuffer
    'DoEvents
    
    Exit Sub

WriteLlamadadeClan_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteLlamadadeClan", Erl)
    Call incomingData.SafeClearPacket
    
End Sub


Public Sub WriteQuestionGM(ByVal Consulta As String, ByVal TipoDeConsulta As String)
    
    On Error GoTo WriteQuestionGM_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "ForumPost" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.QuestionGM)
        Call .WriteASCIIString(Consulta)
        Call .WriteASCIIString(TipoDeConsulta)
        Call .EndPacket
        
    End With
    
    Exit Sub

WriteQuestionGM_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteQuestionGM", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteOfertaInicial(ByVal Oferta As Long)
    
    On Error GoTo WriteOfertaInicial_Err

    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.OfertaInicial)
        Call .WriteLong(Oferta)
        Call .EndPacket
        
    End With
    
    Exit Sub

WriteOfertaInicial_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteOfertaInicial", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteOferta(ByVal OfertaDeSubasta As Long)
    
    On Error GoTo WriteOferta_Err

    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.OfertaDeSubasta)
        Call .WriteLong(OfertaDeSubasta)
        Call .EndPacket
        
    End With
    
    Exit Sub

WriteOferta_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteOferta", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteGlobalMessage(ByVal Message As String)
    
    On Error GoTo WriteGlobalMessage_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 05/17/06
    'Writes the "GuildRequestDetails" message to the outgoing data buffer
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.GlobalMessage)
        
        Call .WriteASCIIString(Message)
        
        Call .EndPacket
        
    End With
    
    Exit Sub

WriteGlobalMessage_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGlobalMessage", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteGlobalOnOff()
    
    On Error GoTo WriteGlobalOnOff_Err
    
    Call outgoingData.WriteID(ClientPacketID.GlobalOnOff)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteGlobalOnOff_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGlobalOnOff", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteBorrandoPJ()
    
    On Error GoTo WriteBorrandoPJ_Err

    With outgoingData
        Call .WriteID(ClientPacketID.BorrarPJ)
        Call .WriteASCIIString(DeleteUser)
        Call .WriteASCIIString(CuentaEmail)
        Call .WriteASCIIString(SEncriptar(CuentaPassword))
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
        Call .WriteASCIIString(MacAdress)  'Seguridad
        Call .WriteLong(HDserial)  'SeguridadHDserial
        Call .WriteASCIIString(CheckMD5)
        Call .EndPacket
        
    End With
    
    Exit Sub

WriteBorrandoPJ_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBorrandoPJ", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteIngresandoConCuenta()
    
    On Error GoTo WriteIngresandoConCuenta_Err

    With outgoingData
        Call .WriteID(ClientPacketID.IngresarConCuenta)
        Call .WriteASCIIString(CuentaEmail)
        Call .WriteASCIIString(SEncriptar(CuentaPassword))
        Call .WriteByte(App.Major)
        Call .WriteByte(App.Minor)
        Call .WriteByte(App.Revision)
        Call .WriteASCIIString(MacAdress)  'Seguridad
        Call .WriteLong(HDserial)  'SeguridadHDserial
        Call .WriteASCIIString(CheckMD5)
        Call .EndPacket
        
    End With
    
    Exit Sub

WriteIngresandoConCuenta_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteIngresandoConCuenta", Erl)
    Call incomingData.SafeClearPacket
    
End Sub


Public Sub WriteNieblaToggle()
    
    On Error GoTo WriteNieblaToggle_Err

    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.NieblaToggle)
        Call .EndPacket
    End With
    
    Exit Sub

WriteNieblaToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteNieblaToggle", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteGenio()
    '***************************************************
    '/GENIO
    'Ladder
    '***************************************************
    
    On Error GoTo WriteGenio_Err
    
    
    Call outgoingData.WriteByte(ClientPacketID.Genio)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteGenio_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteGenio", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteTraerRecompensas()
    
    On Error GoTo WriteTraerRecompensas_Err

    '***************************************************
    'Ladder
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.TraerRecompensas)
        Call .EndPacket
    End With
    
    Exit Sub

WriteTraerRecompensas_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteTraerRecompensas", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteTraerShop()
    
    On Error GoTo WriteTraerShop_Err

    '***************************************************
    'Ladder
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.Traershop)
        Call .EndPacket
    End With
    
    Exit Sub

WriteTraerShop_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteTraerShop", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteTraerRanking()
    
    On Error GoTo WriteTraerRanking_Err

    '***************************************************
    'Ladder
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.TraerRanking)
        Call .EndPacket
        
    End With
    
    Exit Sub

WriteTraerRanking_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteTraerRanking", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WritePareja()
    
    On Error GoTo WritePareja_Err

    '***************************************************
    'Ladder
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.Pareja)
        Call .EndPacket
        
    End With
    
    Exit Sub

WritePareja_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WritePareja", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteQuest()
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Escribe el paquete Quest al servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    
    On Error GoTo WriteQuest_Err
    
    
    Call outgoingData.WriteByte(ClientPacketID.Quest)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteQuest_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteQuest", Erl)
    Call incomingData.SafeClearPacket
    
End Sub
 
Public Sub WriteQuestDetailsRequest(ByVal QuestSlot As Byte)
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Escribe el paquete QuestDetailsRequest al servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    
    On Error GoTo WriteQuestDetailsRequest_Err
    
    
    Call outgoingData.WriteByte(ClientPacketID.QuestDetailsRequest)
    Call outgoingData.WriteByte(QuestSlot)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteQuestDetailsRequest_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteQuestDetailsRequest", Erl)
    Call incomingData.SafeClearPacket
    
End Sub
 
Public Sub WriteQuestAccept(ByVal ListInd As Byte)
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Escribe el paquete QuestAccept al servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    
    On Error GoTo WriteQuestAccept_Err
    
    
    Call outgoingData.WriteByte(ClientPacketID.QuestAccept)
    Call outgoingData.WriteByte(ListInd)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteQuestAccept_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteQuestAccept", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

 
Public Sub WriteQuestListRequest()
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Escribe el paquete QuestListRequest al servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    
    On Error GoTo WriteQuestListRequest_Err
    
    
    Call outgoingData.WriteByte(ClientPacketID.QuestListRequest)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteQuestListRequest_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteQuestListRequest", Erl)
    Call incomingData.SafeClearPacket
    
End Sub
 
Public Sub WriteQuestAbandon(ByVal QuestSlot As Byte)
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Escribe el paquete QuestAbandon al servidor.
    'Last modified: 31/01/2010 by Amraphen
    '$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    'Escribe el ID del paquete.
    
    On Error GoTo WriteQuestAbandon_Err
    
    
    Call outgoingData.WriteByte(ClientPacketID.QuestAbandon)
    
    'Escribe el Slot de Quest.
    Call outgoingData.WriteByte(QuestSlot)
    
    Call outgoingData.EndPacket
    
    Exit Sub

WriteQuestAbandon_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteQuestAbandon", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteResponderPregunta(ByVal Respuesta As Boolean)
    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 22/11/2017
    '***************************************************
    
    On Error GoTo WriteResponderPregunta_Err
    
    
    Call outgoingData.WriteByte(ClientPacketID.ResponderPregunta)
    Call outgoingData.WriteBoolean(Respuesta)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteResponderPregunta_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteResponderPregunta", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteCorreo()
    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 22/11/2017
    '***************************************************
    
    On Error GoTo WriteCorreo_Err
    
    
    Call outgoingData.WriteByte(ClientPacketID.Correo)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteCorreo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCorreo", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteSendCorreo(ByVal UserNick As String, ByVal msg As String, ByVal ItemCount As Byte)
    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 4/5/2020
    '***************************************************
    
    On Error GoTo WriteSendCorreo_Err
    
    
    Call outgoingData.WriteByte(ClientPacketID.SendCorreo)
    
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
    
    Call outgoingData.EndPacket
    
    Exit Sub

WriteSendCorreo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteSendCorreo", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteComprarItem(ByVal ItemIndex As Byte)
    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 22/11/2017
    '***************************************************
    
    On Error GoTo WriteComprarItem_Err
    
    
    Call outgoingData.WriteByte(ClientPacketID.ComprarItem)
    Call outgoingData.WriteByte(ItemIndex)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteComprarItem_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteComprarItem", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteCompletarViaje(ByVal destino As Byte, ByVal costo As Long)
    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 22/11/2017
    '***************************************************
    
    On Error GoTo WriteCompletarViaje_Err
    
    
    Call outgoingData.WriteByte(ClientPacketID.CompletarViaje)
    Call outgoingData.WriteByte(destino)
    Call outgoingData.WriteLong(costo)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteCompletarViaje_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCompletarViaje", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteRetirarItemCorreo(ByVal IndexMsg As Integer)
    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 22/11/2017
    '***************************************************
    
    On Error GoTo WriteRetirarItemCorreo_Err
    
    
    Call outgoingData.WriteByte(ClientPacketID.RetirarItemCorreo)
    Call outgoingData.WriteInteger(IndexMsg)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteRetirarItemCorreo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteRetirarItemCorreo", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteBorrarCorreo(ByVal IndexMsg As Integer)
    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 22/11/2017
    '***************************************************
    
    On Error GoTo WriteBorrarCorreo_Err
    
    
    Call outgoingData.WriteByte(ClientPacketID.BorrarCorreo)
    Call outgoingData.WriteInteger(IndexMsg)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteBorrarCorreo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBorrarCorreo", Erl)
    Call incomingData.SafeClearPacket
    
End Sub


''
' Handles the RestOK message.

Public Sub WriteCodigo(ByVal Codigo As String)
    
    On Error GoTo WriteCodigo_Err

    '***************************************************
    'Ladder
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.EnviarCodigo)
        Call .WriteASCIIString(Codigo)
        Call .EndPacket
        
    End With

    Exit Sub

WriteCodigo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCodigo", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteCreaerTorneo(ByVal nivelminimo As Byte, ByVal nivelmaximo As Byte, ByVal cupos As Byte, ByVal costo As Long, ByVal mago As Byte, ByVal clerico As Byte, ByVal guerrero As Byte, ByVal asesino As Byte, ByVal bardo As Byte, ByVal druido As Byte, ByVal paladin As Byte, ByVal cazador As Byte, ByVal Trabajador As Byte, ByVal map As Integer, ByVal x As Byte, ByVal y As Byte, ByVal Name As String, ByVal reglas As String)
    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 16/05/2020
    '***************************************************
    
    On Error GoTo WriteCreaerTorneo_Err
    
    
    Call outgoingData.WriteByte(ClientPacketID.CrearTorneo)
    
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
    
    
    Call outgoingData.EndPacket
    
    Exit Sub

WriteCreaerTorneo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCreaerTorneo", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteComenzarTorneo()
    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 16/05/2020
    '***************************************************
    
    On Error GoTo WriteComenzarTorneo_Err
    
    
    Call outgoingData.WriteByte(ClientPacketID.ComenzarTorneo)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteComenzarTorneo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteComenzarTorneo", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteCancelarTorneo()
    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 16/05/2020
    '***************************************************
    
    On Error GoTo WriteCancelarTorneo_Err
    
    
    Call outgoingData.WriteByte(ClientPacketID.CancelarTorneo)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteCancelarTorneo_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCancelarTorneo", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteBusquedaTesoro(ByVal TIPO As Byte)
    '***************************************************
    'Author: Pablo Mercavides
    'Last Modification: 16/05/2020
    '***************************************************
    
    On Error GoTo WriteBusquedaTesoro_Err
    
    
    Call outgoingData.WriteByte(ClientPacketID.BusquedaTesoro)
    Call outgoingData.WriteByte(TIPO)
    Call outgoingData.EndPacket
    
    Exit Sub

WriteBusquedaTesoro_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteBusquedaTesoro", Erl)
    Call incomingData.SafeClearPacket
    
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
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.Home)
        Call .EndPacket
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
    
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.Consulta)
        Call .WriteASCIIString(Nick)
        Call .EndPacket
        
    End With
    
End Sub

Public Sub WriteRequestScreenShot(ByVal Nick As String)

    With outgoingData

        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.RequestScreenShot)
        Call .WriteASCIIString(Nick)
        Call .EndPacket
        
    End With
    
End Sub

Public Sub WriteSendScreenShot(ScreenShotSerialized As String)

    On Error GoTo Handler

    With outgoingData

        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.SendScreenShot)
        Call .WriteASCIIString(ScreenShotSerialized)
        Call .EndPacket
    End With
    
Handler:

    If outgoingData.errNumber = outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer
        Resume

    End If
    
End Sub

Public Sub WriteCuentaExtractItem(ByVal Slot As Byte, ByVal Amount As Integer, ByVal slotdestino As Byte)
    
    On Error GoTo WriteCuentaExtractItem_Err

    '***************************************************
    'Author: Ladder
    'Last Modification: 22/11/21
    'Retirar item de cuenta
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.CuentaExtractItem)
        Call .WriteByte(Slot)
        Call .WriteInteger(Amount)
        Call .WriteByte(slotdestino)
        Call .EndPacket
    End With
    
    Exit Sub

WriteCuentaExtractItem_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCuentaExtractItem", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteCuentaDeposit(ByVal Slot As Byte, ByVal Amount As Integer, ByVal slotdestino As Byte)
    
    On Error GoTo WriteCuentaDeposit_Err

    '***************************************************
    'Author: Ladder
    'Last Modification: 22/11/21
    'Depositar item en cuenta
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.CuentaDeposit)
        Call .WriteByte(Slot)
        Call .WriteInteger(Amount)
        Call .WriteByte(slotdestino)
        Call .EndPacket
    End With
    
    Exit Sub
WriteCuentaDeposit_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCuentaDeposit", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteDuel(Players As String, ByVal Apuesta As Long, Optional ByVal PocionesRojas As Long = -1, Optional ByVal CaenItems As Boolean = False)

    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.Duel)
        Call .WriteASCIIString(Players)
        Call .WriteLong(Apuesta)
        Call .WriteInteger(PocionesRojas)
        Call .WriteBoolean(CaenItems)
        Call .EndPacket
    End With

End Sub

Public Sub WriteAcceptDuel(Offerer As String)

    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.AcceptDuel)
        Call .WriteASCIIString(Offerer)
        Call .EndPacket
    End With

End Sub

Public Sub WriteCancelDuel()

    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.CancelDuel)
        Call .EndPacket
    End With

End Sub

Public Sub WriteQuitDuel()

    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.QuitDuel)
        Call .EndPacket
    End With

End Sub

Public Sub WriteCreateEvent(EventName As String)

    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.CreateEvent)
        Call .WriteASCIIString(EventName)
        Call .EndPacket
    End With

End Sub


Public Sub WriteCommerceSendChatMessage(ByVal Message As String)

    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.CommerceSendChatMessage)
        Call .WriteASCIIString(Message)
        Call .EndPacket
    End With

End Sub

Public Sub WriteLogMacroClickHechizo()

    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.LogMacroClickHechizo)
        Call .EndPacket
    End With

End Sub

Public Sub WriteNieveToggle()
    
    On Error GoTo WriteNieveToggle_Err

    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.NieveToggle)
        Call .EndPacket

    End With
    
    Exit Sub

WriteNieveToggle_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteNieveToggle", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteCompletarAccion(ByVal Accion As Byte)
    
    On Error GoTo WriteCompletarAccion_Err

    '***************************************************
    'Author: Pablo Mercavides
    '***************************************************
    With outgoingData
        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.CompletarAccion)
        Call .WriteByte(Accion)
        Call .EndPacket

    End With
    
    Exit Sub

WriteCompletarAccion_Err:
    Call RegistrarError(Err.Number, Err.Description, "Protocol.WriteCompletarAccion", Erl)
    Call incomingData.SafeClearPacket
    
End Sub

Public Sub WriteTolerancia0(Nick As String)

    On Error GoTo Handler

    With outgoingData

        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.Tolerancia0)
        Call .WriteASCIIString(Nick)
        Call .EndPacket

    End With
    
Handler:

    If outgoingData.errNumber = outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer
        Resume

    End If

End Sub

Public Sub WriteGetMapInfo()

    On Error GoTo Handler

    With outgoingData

        Call .WriteID(ClientPacketID.NewPacketID)
        Call .WriteByte(ClientPacketID.GetMapInfo)
        Call .EndPacket

    End With
    
Handler:

    If outgoingData.errNumber = outgoingData.NotEnoughSpaceErrCode Then
        Call FlushBuffer
        Resume

    End If

End Sub
