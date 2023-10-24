Attribute VB_Name = "Protocol_Writes"
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
Private Writer As Network.Writer

Public Function writer_is_nothing() As Boolean
writer_is_nothing = Writer Is Nothing
End Function
Public Sub Initialize()
100     Set Writer = New Network.Writer
End Sub

Public Sub Clear()
100     Call Writer.Clear
End Sub



#If PYMMO = 1 Then
''
' Writes the "LoginExistingChar" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteLoginExistingChar()
        '<EhHeader>
        On Error GoTo WriteLoginExistingChar_Err
        
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eLoginExistingChar)
102     Call Writer.WriteString8(encrypted_session_token)


        Dim encrypted_username_b64 As String
        encrypted_username_b64 = AO20CryptoSysWrapper.Encrypt(cnvHexStrFromBytes(public_key), username)
        
104     Call Writer.WriteString8(encrypted_username_b64)
106     Call Writer.WriteInt8(App.Major)
108     Call Writer.WriteInt8(App.Minor)
110     Call Writer.WriteInt8(App.Revision)
118     Call Writer.WriteString8(CheckMD5)
            
120     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteLoginExistingChar_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteLoginExistingChar", Erl)
        '</EhFooter>
End Sub

''
' Writes the "LoginNewChar" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteLoginNewChar(ByVal Name As String, ByVal Race As Integer, ByVal Gender As Integer, ByVal Class As Integer, ByVal Head As Integer, ByVal HomeCity As Integer)
        '<EhHeader>
        On Error GoTo WriteLoginNewChar_Err
        '</EhHeader>
        
        Dim encrypted_username_b64 As String
        encrypted_username_b64 = AO20CryptoSysWrapper.Encrypt(cnvHexStrFromBytes(public_key), Name)
        
100     Call Writer.WriteInt16(ClientPacketID.eLoginNewChar)
102     Call Writer.WriteString8(encrypted_session_token)
104     Call Writer.WriteString8(encrypted_username_b64)
106     Call Writer.WriteInt8(App.Major)
108     Call Writer.WriteInt8(App.Minor)
110     Call Writer.WriteInt8(App.Revision)
128     Call Writer.WriteString8(CheckMD5)
114     Call Writer.WriteInt8(Race)
116     Call Writer.WriteInt8(Gender)
118     Call Writer.WriteInt8(Class)
120     Call Writer.WriteInt16(Head)
122     Call Writer.WriteInt8(HomeCity)
    
130     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteLoginNewChar_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteLoginNewChar", Erl)
        '</EhFooter>
End Sub
#End If

#If PYMMO = 0 Then

Public Sub WriteCreateAccount()
        '<EhHeader>
        On Error GoTo WriteCreateAccount_Err
        
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eCreateAccount)

104     Call Writer.WriteString8(CuentaEmail)
        Call Writer.WriteString8(CuentaPassword)

120     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCreateAccount_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCreateAccount", Erl)
        '</EhFooter>
End Sub

Public Sub WriteLoginAccount()
        '<EhHeader>
        On Error GoTo WriteLoginAccount_Err
        
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eLoginAccount)

104     Call Writer.WriteString8(CuentaEmail)
        Call Writer.WriteString8(CuentaPassword)

120     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteLoginAccount_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteLoginAccount", Erl)
        '</EhFooter>
End Sub

Public Sub WriteDeleteCharacter()

End Sub
''
' Writes the "LoginExistingChar" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteLoginExistingChar()
        '<EhHeader>
        On Error GoTo WriteLoginExistingChar_Err
        
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eLoginExistingChar)

104     Call Writer.WriteString8(username)
            
120     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteLoginExistingChar_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteLoginExistingChar", Erl)
        '</EhFooter>
End Sub

''
' Writes the "LoginNewChar" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteLoginNewChar(ByVal Name As String, ByVal Race As Integer, ByVal Gender As Integer, ByVal Class As Integer, ByVal Head As Integer, ByVal HomeCity As Integer)
        '<EhHeader>
        On Error GoTo WriteLoginNewChar_Err
        '</EhHeader>

100     Call Writer.WriteInt16(ClientPacketID.eLoginNewChar)
104     Call Writer.WriteString8(Name)
114     Call Writer.WriteInt(Race)
116     Call Writer.WriteInt(Gender)
118     Call Writer.WriteInt(Class)
120     Call Writer.WriteInt(Head)
122     Call Writer.WriteInt(HomeCity)
    
130     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteLoginNewChar_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteLoginNewChar", Erl)
        '</EhFooter>
End Sub
#End If

''
' Writes the "Talk" message to the outgoing data buffer.
'
' @param    chat The chat text to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteTalk(ByVal chat As String)
        '<EhHeader>
        On Error GoTo WriteTalk_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eTalk)
102     Call Writer.WriteString8(chat)
        packetCounters.TS_Talk = packetCounters.TS_Talk + 1
        Call Writer.WriteInt32(packetCounters.TS_Talk)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteTalk_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteTalk", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Yell" message to the outgoing data buffer.
'
' @param    chat The chat text to be sent.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteYell(ByVal chat As String)
        '<EhHeader>
        On Error GoTo WriteYell_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eYell)
102     Call Writer.WriteString8(chat)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteYell_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteYell", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Whisper" message to the outgoing data buffer.
'
' @param    charIndex The index of the char to whom to whisper.
' @param    chat The chat text to be sent to the user.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteWhisper(ByVal nombre As String, ByVal chat As String)
        '<EhHeader>
        On Error GoTo WriteWhisper_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eWhisper)
102     Call Writer.WriteString8(nombre)
104     Call Writer.WriteString8(chat)
    
106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteWhisper_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteWhisper", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Walk" message to the outgoing data buffer.
'
' @param    heading The direction in wich the user is moving.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteWalk(ByVal Heading As E_Heading)
        '<EhHeader>
        On Error GoTo WriteWalk_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eWalk)
102     Call Writer.WriteInt8(Heading)
        packetCounters.TS_Walk = packetCounters.TS_Walk + 1
        Call Writer.WriteInt32(packetCounters.TS_Walk)
        
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteWalk_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteWalk", Erl)
        '</EhFooter>
End Sub

''
' Writes the "RequestPositionUpdate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRequestPositionUpdate()
        '<EhHeader>
        On Error GoTo WriteRequestPositionUpdate_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eRequestPositionUpdate)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteRequestPositionUpdate_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteRequestPositionUpdate", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Attack" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteAttack()
        '<EhHeader>
        On Error GoTo WriteAttack_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eAttack)
        packetCounters.TS_Attack = packetCounters.TS_Attack + 1
        Call Writer.WriteInt32(packetCounters.TS_Attack)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteAttack_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteAttack", Erl)
        '</EhFooter>
End Sub

''
' Writes the "PickUp" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePickUp()
        '<EhHeader>
        On Error GoTo WritePickUp_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.ePickUp)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WritePickUp_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WritePickUp", Erl)
        '</EhFooter>
End Sub

''
' Writes the "SafeToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSafeToggle()
        '<EhHeader>
        On Error GoTo WriteSafeToggle_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eSafeToggle)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteSafeToggle_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteSafeToggle", Erl)
        '</EhFooter>
End Sub

Public Sub WriteSeguroClan()
        '<EhHeader>
        On Error GoTo WriteSeguroClan_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eSeguroClan)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteSeguroClan_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteSeguroClan", Erl)
        '</EhFooter>
End Sub

Public Sub WriteTraerBoveda()
        '<EhHeader>
        On Error GoTo WriteTraerBoveda_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eTraerBoveda)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteTraerBoveda_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteTraerBoveda", Erl)
        '</EhFooter>
End Sub


''
' Writes the "PartySafeToggle" message to the outgoing data buffer.
'PartySafeOn
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteParyToggle()
        '<EhHeader>
        On Error GoTo WriteParyToggle_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.ePartySafeToggle)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteParyToggle_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteParyToggle", Erl)
        '</EhFooter>
End Sub

''
' Writes the "SeguroResu" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSeguroResu()
        '<EhHeader>
        On Error GoTo WriteSeguroResu_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eSeguroResu)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteSeguroResu_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteSeguroResu", Erl)
        '</EhFooter>
End Sub

''
' Writes the "RequestGuildLeaderInfo" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRequestGuildLeaderInfo()
        '<EhHeader>
        On Error GoTo WriteRequestGuildLeaderInfo_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eRequestGuildLeaderInfo)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteRequestGuildLeaderInfo_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteRequestGuildLeaderInfo", Erl)
        '</EhFooter>
End Sub

''
' Writes the "RequestAtributes" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRequestAtributes()
        '<EhHeader>
        On Error GoTo WriteRequestAtributes_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eRequestAtributes)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteRequestAtributes_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteRequestAtributes", Erl)
        '</EhFooter>
End Sub


Public Sub WriteRequestGrupo()
        '<EhHeader>
        On Error GoTo WriteRequestGrupo_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eRequestGrupo)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteRequestGrupo_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteRequestGrupo", Erl)
        '</EhFooter>
End Sub

''
' Writes the "RequestSkills" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRequestSkills()
        '<EhHeader>
        On Error GoTo WriteRequestSkills_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eRequestSkills)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteRequestSkills_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteRequestSkills", Erl)
        '</EhFooter>
End Sub

''
' Writes the "RequestMiniStats" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRequestMiniStats()
        '<EhHeader>
        On Error GoTo WriteRequestMiniStats_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eRequestMiniStats)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteRequestMiniStats_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteRequestMiniStats", Erl)
        '</EhFooter>
End Sub

''
' Writes the "CommerceEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCommerceEnd()
        '<EhHeader>
        On Error GoTo WriteCommerceEnd_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eCommerceEnd)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCommerceEnd_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCommerceEnd", Erl)
        '</EhFooter>
End Sub

''
' Writes the "UserCommerceEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserCommerceEnd()
        '<EhHeader>
        On Error GoTo WriteUserCommerceEnd_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eUserCommerceEnd)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteUserCommerceEnd_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteUserCommerceEnd", Erl)
        '</EhFooter>
End Sub

''
' Writes the "BankEnd" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBankEnd()
        '<EhHeader>
        On Error GoTo WriteBankEnd_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eBankEnd)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteBankEnd_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteBankEnd", Erl)
        '</EhFooter>
End Sub

''
' Writes the "UserCommerceOk" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserCommerceOk()
        '<EhHeader>
        On Error GoTo WriteUserCommerceOk_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eUserCommerceOk)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteUserCommerceOk_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteUserCommerceOk", Erl)
        '</EhFooter>
End Sub

''
' Writes the "UserCommerceReject" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserCommerceReject()
        '<EhHeader>
        On Error GoTo WriteUserCommerceReject_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eUserCommerceReject)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteUserCommerceReject_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteUserCommerceReject", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Drop" message to the outgoing data buffer.
'
' @param    slot Inventory slot where the item to drop is.
' @param    amount Number of items to drop.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteDrop(ByVal Slot As Byte, ByVal Amount As Long)
        '<EhHeader>
        On Error GoTo WriteDrop_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eDrop)
102     Call Writer.WriteInt8(Slot)
104     Call Writer.WriteInt32(Amount)
        packetCounters.TS_Drop = packetCounters.TS_Drop + 1
        Call Writer.WriteInt32(packetCounters.TS_Drop)
    
106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteDrop_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteDrop", Erl)
        '</EhFooter>
End Sub

''
' Writes the "CastSpell" message to the outgoing data buffer.
'
' @param    slot Spell List slot where the spell to cast is.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCastSpell(ByVal Slot As Byte)
        '<EhHeader>
        On Error GoTo WriteCastSpell_Err
        '</EhHeader>
       ' Dim arr() As Byte
       ' Dim packet_crc As Long
        
100     Call Writer.WriteInt16(ClientPacketID.eCastSpell)
102     Call Writer.WriteInt8(Slot)
        packetCounters.TS_CastSpell = packetCounters.TS_CastSpell + 1
        Call Writer.WriteInt32(packetCounters.TS_CastSpell)
        
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCastSpell_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCastSpell", Erl)
        '</EhFooter>
End Sub

Public Sub WriteInvitarGrupo()
        '<EhHeader>
        On Error GoTo WriteInvitarGrupo_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eInvitarGrupo)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteInvitarGrupo_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteInvitarGrupo", Erl)
        '</EhFooter>
End Sub

Public Sub WriteMarcaDeClan()
        '<EhHeader>
        On Error GoTo WriteMarcaDeClan_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eMarcaDeClanPack)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteMarcaDeClan_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteMarcaDeClan", Erl)
        '</EhFooter>
End Sub

Public Sub WriteMarcaDeGm()
        '<EhHeader>
        On Error GoTo WriteMarcaDeGm_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eMarcaDeGMPack)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteMarcaDeGm_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteMarcaDeGm", Erl)
        '</EhFooter>
End Sub

Public Sub WriteAbandonarGrupo()
        '<EhHeader>
        On Error GoTo WriteAbandonarGrupo_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eAbandonarGrupo)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteAbandonarGrupo_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteAbandonarGrupo", Erl)
        '</EhFooter>
End Sub

Public Sub WriteEcharDeGrupo(ByVal indice As Byte)
        '<EhHeader>
        On Error GoTo WriteEcharDeGrupo_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eHecharDeGrupo)
102     Call Writer.WriteInt8(indice)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteEcharDeGrupo_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteEcharDeGrupo", Erl)
        '</EhFooter>
End Sub

''
' Writes the "LeftClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteLeftClick(ByVal x As Byte, ByVal y As Byte)
        '<EhHeader>
        On Error GoTo WriteLeftClick_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eLeftClick)
102     Call Writer.WriteInt8(x)
104     Call Writer.WriteInt8(y)
        packetCounters.TS_LeftClick = packetCounters.TS_LeftClick + 1
        'Debug.Print packetCounters.TS_LeftClick
        Call Writer.WriteInt32(packetCounters.TS_LeftClick)
    
106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteLeftClick_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteLeftClick", Erl)
        '</EhFooter>
End Sub

''
' Writes the "DoubleClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteDoubleClick(ByVal x As Byte, ByVal y As Byte)
        '<EhHeader>
        On Error GoTo WriteDoubleClick_Err
        '</EhHeader>
        
100     Call Writer.WriteInt16(ClientPacketID.eDoubleClick)
102     Call Writer.WriteInt8(x)
104     Call Writer.WriteInt8(y)
    
106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteDoubleClick_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteDoubleClick", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Work" message to the outgoing data buffer.
'
' @param    skill The skill which the user attempts to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteWork(ByVal Skill As eSkill)
        '<EhHeader>
        On Error GoTo WriteWork_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eWork)
102     Call Writer.WriteInt8(Skill)
        packetCounters.TS_Work = packetCounters.TS_Work + 1
        Call Writer.WriteInt32(packetCounters.TS_Work)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteWork_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteWork", Erl)
        '</EhFooter>
End Sub


''
' Writes the "UseSpellMacro" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUseSpellMacro()
        '<EhHeader>
        On Error GoTo WriteUseSpellMacro_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eUseSpellMacro)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteUseSpellMacro_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteUseSpellMacro", Erl)
        '</EhFooter>
End Sub
''
' Writes the "UseItem" message to the outgoing data buffer.
'
' @param    slot Invetory slot where the item to use is.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUseItem(ByVal Slot As Byte)

        'If LastUseItemTimeStamp > 0 Then
        '    If (GetTickCount - LastUseItemTimeStamp) < 100 Then Exit Sub
        'End If
        
        'LastUseItemTimeStamp = GetTickCount
        '<EhHeader>
        
        On Error GoTo WriteUseItem_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eUseItem)
102     Call Writer.WriteInt8(Slot)
        Call Writer.WriteInt8(ActiveInventoryTab = eInventory)
        
        packetCounters.TS_UseItem = packetCounters.TS_UseItem + 1
        Call Writer.WriteInt32(packetCounters.TS_UseItem)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteUseItem_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteUseItem", Erl)
        '</EhFooter>
End Sub
''
' Writes the "UseItem" message to the outgoing data buffer.
'
' @param    slot Invetory slot where the item to use is.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUseItemU(ByVal Slot As Byte)

        'If LastUseItemTimeStampU > 0 Then
        '    If (GetTickCount - LastUseItemTimeStampU) < 100 Then Exit Sub
        'End If
        
        'LastUseItemTimeStampU = GetTickCount
        '<EhHeader>
        
        On Error GoTo WriteUseItemU_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eUseItemU)
102     Call Writer.WriteInt8(Slot)
        
        packetCounters.TS_UseItemU = packetCounters.TS_UseItemU + 1
        Call Writer.WriteInt32(packetCounters.TS_UseItemU)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteUseItemU_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteUseItemU", Erl)
        '</EhFooter>
End Sub
''
' Writes the "UseItem" message to the outgoing data buffer.
'
' @param    slot Invetory slot where the item to use is.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRepeatMacro()
        
        On Error GoTo WriteRepeatMacro_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eRepeatMacro)
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteRepeatMacro_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteRepeatMacro", Erl)
        '</EhFooter>
End Sub
''
' Writes the "UseItem" message to the outgoing data buffer.
'
' @param    slot Invetory slot where the item to use is.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub writeBuyShopItem(ByVal objNum As Long)
        
        On Error GoTo writeBuyShopItem_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eBuyShopItem)
        Call Writer.WriteInt32(objNum)
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

writeBuyShopItem_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.writeBuyShopItem", Erl)
        '</EhFooter>
End Sub

''
' Writes the "CraftBlacksmith" message to the outgoing data buffer.
'
' @param    item Index of the item to craft in the list sent by the server.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCraftBlacksmith(ByVal Item As Integer)
        '<EhHeader>
        On Error GoTo WriteCraftBlacksmith_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eCraftBlacksmith)
102     Call Writer.WriteInt16(Item)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCraftBlacksmith_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCraftBlacksmith", Erl)
        '</EhFooter>
End Sub

' HarThaoS: Arreglo las cagadas de Ladder. A tu casa, a mimir.
Public Sub WriteCraftCarpenter(ByVal Item As Integer, ByVal cantidad As Long)
        '<EhHeader>
        On Error GoTo WriteCraftCarpenter_Err
        '</EhHeader><
100     Call Writer.WriteInt16(ClientPacketID.eCraftCarpenter)
102     Call Writer.WriteInt16(Item)
103     Call Writer.WriteInt32(cantidad)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCraftCarpenter_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCraftCarpenter", Erl)
        '</EhFooter>
End Sub

Public Sub WriteCraftAlquimista(ByVal Item As Integer)
        '<EhHeader>
        On Error GoTo WriteCraftAlquimista_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eCraftAlquimista)
102     Call Writer.WriteInt16(Item)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCraftAlquimista_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCraftAlquimista", Erl)
        '</EhFooter>
End Sub

Public Sub WriteCraftSastre(ByVal Item As Integer)
        '<EhHeader>
        On Error GoTo WriteCraftSastre_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eCraftSastre)
102     Call Writer.WriteInt16(Item)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCraftSastre_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCraftSastre", Erl)
        '</EhFooter>
End Sub

''
' Writes the "WorkLeftClick" message to the outgoing data buffer.
'
' @param    x Tile coord in the x-axis in which the user clicked.
' @param    y Tile coord in the y-axis in which the user clicked.
' @param    skill The skill which the user attempts to use.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteWorkLeftClick(ByVal x As Byte, ByVal y As Byte, ByVal Skill As eSkill)
        '<EhHeader>
        On Error GoTo WriteWorkLeftClick_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eWorkLeftClick)
102     Call Writer.WriteInt8(x)
104     Call Writer.WriteInt8(y)
106     Call Writer.WriteInt8(Skill)

        packetCounters.TS_WorkLeftClick = packetCounters.TS_WorkLeftClick + 1
        Call Writer.WriteInt32(packetCounters.TS_WorkLeftClick)
    
108     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteWorkLeftClick_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteWorkLeftClick", Erl)
        '</EhFooter>
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
        '<EhHeader>
        On Error GoTo WriteCreateNewGuild_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eCreateNewGuild)
102     Call Writer.WriteString8(desc)
104     Call Writer.WriteString8(Name)
106     Call Writer.WriteInt8(Alineacion)
    
108     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCreateNewGuild_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCreateNewGuild", Erl)
        '</EhFooter>
End Sub

''
' Writes the "SpellInfo" message to the outgoing data buffer.
'
' @param    slot Spell List slot where the spell which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSpellInfo(ByVal Slot As Byte)
        '<EhHeader>
        On Error GoTo WriteSpellInfo_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eSpellInfo)
102     Call Writer.WriteInt8(Slot)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteSpellInfo_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteSpellInfo", Erl)
        '</EhFooter>
End Sub

''
' Writes the "EquipItem" message to the outgoing data buffer.
'
' @param    slot Invetory slot where the item to equip is.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteEquipItem(ByVal Slot As Byte)
        '<EhHeader>
        On Error GoTo WriteEquipItem_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eEquipItem)
102     Call Writer.WriteInt8(Slot)
        packetCounters.TS_EquipItem = packetCounters.TS_EquipItem + 1
        Call Writer.WriteInt32(packetCounters.TS_EquipItem)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteEquipItem_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteEquipItem", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ChangeHeading" message to the outgoing data buffer.
'
' @param    heading The direction in wich the user is moving.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeHeading(ByVal Heading As E_Heading)
        '<EhHeader>
        On Error GoTo WriteChangeHeading_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eChangeHeading)
102     Call Writer.WriteInt8(Heading)
        packetCounters.TS_ChangeHeading = packetCounters.TS_ChangeHeading + 1
        Call Writer.WriteInt32(packetCounters.TS_ChangeHeading)
    
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteChangeHeading_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteChangeHeading", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ModifySkills" message to the outgoing data buffer.
'
' @param    skillEdt a-based array containing for each skill the number of points to add to it.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteModifySkills(ByRef skillEdt() As Byte)
        '<EhHeader>
        On Error GoTo WriteModifySkills_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eModifySkills)
    
        Dim i As Long
    
102     For i = 1 To NUMSKILLS
104         Call Writer.WriteInt8(skillEdt(i))
106     Next i
    
108     Call modNetwork.Send(Writer)

        '<EhFooter>
        Exit Sub

WriteModifySkills_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteModifySkills", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Train" message to the outgoing data buffer.
'
' @param    creature Position within the list provided by the server of the creature to train against.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteTrain(ByVal creature As Byte)
        '<EhHeader>
        On Error GoTo WriteTrain_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eTrain)
102     Call Writer.WriteInt8(creature)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteTrain_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteTrain", Erl)
        '</EhFooter>
End Sub

''
' Writes the "CommerceBuy" message to the outgoing data buffer.
'
' @param    slot Position within the NPC's inventory in which the desired item is.
' @param    amount Number of items to buy.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCommerceBuy(ByVal Slot As Byte, ByVal Amount As Integer)
        '<EhHeader>
        On Error GoTo WriteCommerceBuy_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eCommerceBuy)
102     Call Writer.WriteInt8(Slot)
104     Call Writer.WriteInt16(Amount)
    
106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCommerceBuy_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCommerceBuy", Erl)
        '</EhFooter>
End Sub

Public Sub WriteUseKey(ByVal Slot As Byte)
        '<EhHeader>
        On Error GoTo WriteUseKey_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eUseKey)
102     Call Writer.WriteInt8(Slot)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteUseKey_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteUseKey", Erl)
        '</EhFooter>
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
        '<EhHeader>
        On Error GoTo WriteBankExtractItem_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eBankExtractItem)
102     Call Writer.WriteInt8(Slot)
104     Call Writer.WriteInt16(Amount)
106     Call Writer.WriteInt8(slotdestino)
    
108     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteBankExtractItem_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteBankExtractItem", Erl)
        '</EhFooter>
End Sub

''
' Writes the "CommerceSell" message to the outgoing data buffer.
'
' @param    slot Position within user inventory in which the desired item is.
' @param    amount Number of items to sell.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCommerceSell(ByVal Slot As Byte, ByVal Amount As Integer)
        '<EhHeader>
        On Error GoTo WriteCommerceSell_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eCommerceSell)
102     Call Writer.WriteInt8(Slot)
104     Call Writer.WriteInt16(Amount)
    
106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCommerceSell_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCommerceSell", Erl)
        '</EhFooter>
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
        '<EhHeader>
        On Error GoTo WriteBankDeposit_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eBankDeposit)
102     Call Writer.WriteInt8(Slot)
104     Call Writer.WriteInt16(Amount)
106     Call Writer.WriteInt8(slotdestino)
    
108     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteBankDeposit_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteBankDeposit", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ForumPost" message to the outgoing data buffer.
'
' @param    title The message's title.
' @param    message The body of the message.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteForumPost(ByVal title As String, ByVal Message As String)
        '<EhHeader>
        On Error GoTo WriteForumPost_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eForumPost)
102     Call Writer.WriteString8(title)
104     Call Writer.WriteString8(Message)
    
106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteForumPost_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteForumPost", Erl)
        '</EhFooter>
End Sub

''
' Writes the "MoveSpell" message to the outgoing data buffer.
'
' @param    upwards True if the spell will be moved up in the list, False if it will be moved downwards.
' @param    slot Spell List slot where the spell which's info is requested is.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteMoveSpell(ByVal upwards As Boolean, ByVal Slot As Byte)
        '<EhHeader>
        On Error GoTo WriteMoveSpell_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eMoveSpell)
102     Call Writer.WriteBool(upwards)
104     Call Writer.WriteInt8(Slot)
    
106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteMoveSpell_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteMoveSpell", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ClanCodexUpdate" message to the outgoing data buffer.
'
' @param    desc New description of the clan.
' @param    codex New codex of the clan.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteClanCodexUpdate(ByVal desc As String)
        '<EhHeader>
        On Error GoTo WriteClanCodexUpdate_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eClanCodexUpdate)
102     Call Writer.WriteString8(desc)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteClanCodexUpdate_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteClanCodexUpdate", Erl)
        '</EhFooter>
End Sub

''
' Writes the "UserCommerceOffer" message to the outgoing data buffer.
'
' @param    slot Position within user inventory in which the desired item is.
' @param    amount Number of items to offer.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUserCommerceOffer(ByVal Slot As Byte, ByVal Amount As Long)
        '<EhHeader>
        On Error GoTo WriteUserCommerceOffer_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eUserCommerceOffer)
102     Call Writer.WriteInt8(Slot)
104     Call Writer.WriteInt32(Amount)
    
106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteUserCommerceOffer_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteUserCommerceOffer", Erl)
        '</EhFooter>
End Sub

''
' Writes the "GuildAcceptPeace" message to the outgoing data buffer.
'
' @param    guild The guild whose peace offer is accepted.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildAcceptPeace(ByVal guild As String)
        '<EhHeader>
        On Error GoTo WriteGuildAcceptPeace_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGuildAcceptPeace)
102     Call Writer.WriteString8(guild)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGuildAcceptPeace_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGuildAcceptPeace", Erl)
        '</EhFooter>
End Sub

''
' Writes the "GuildRejectAlliance" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance offer is rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildRejectAlliance(ByVal guild As String)
        '<EhHeader>
        On Error GoTo WriteGuildRejectAlliance_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGuildRejectAlliance)
102     Call Writer.WriteString8(guild)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGuildRejectAlliance_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGuildRejectAlliance", Erl)
        '</EhFooter>
End Sub

''
' Writes the "GuildRejectPeace" message to the outgoing data buffer.
'
' @param    guild The guild whose peace offer is rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildRejectPeace(ByVal guild As String)
        '<EhHeader>
        On Error GoTo WriteGuildRejectPeace_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGuildRejectPeace)
102     Call Writer.WriteString8(guild)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGuildRejectPeace_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGuildRejectPeace", Erl)
        '</EhFooter>
End Sub

''
' Writes the "GuildAcceptAlliance" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance offer is accepted.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildAcceptAlliance(ByVal guild As String)
        '<EhHeader>
        On Error GoTo WriteGuildAcceptAlliance_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGuildAcceptAlliance)
102     Call Writer.WriteString8(guild)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGuildAcceptAlliance_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGuildAcceptAlliance", Erl)
        '</EhFooter>
End Sub

''
' Writes the "GuildOfferPeace" message to the outgoing data buffer.
'
' @param    guild The guild to whom peace is offered.
' @param    proposal The text to s the proposal.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildOfferPeace(ByVal guild As String, ByVal proposal As String)
        '<EhHeader>
        On Error GoTo WriteGuildOfferPeace_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGuildOfferPeace)
102     Call Writer.WriteString8(guild)
104     Call Writer.WriteString8(proposal)
    
106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGuildOfferPeace_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGuildOfferPeace", Erl)
        '</EhFooter>
End Sub

''
' Writes the "GuildOfferAlliance" message to the outgoing data buffer.
'
' @param    guild The guild to whom an aliance is offered.
' @param    proposal The text to s the proposal.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildOfferAlliance(ByVal guild As String, ByVal proposal As String)
        '<EhHeader>
        On Error GoTo WriteGuildOfferAlliance_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGuildOfferAlliance)
102     Call Writer.WriteString8(guild)
104     Call Writer.WriteString8(proposal)
    
106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGuildOfferAlliance_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGuildOfferAlliance", Erl)
        '</EhFooter>
End Sub

''
' Writes the "GuildAllianceDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose aliance proposal's details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildAllianceDetails(ByVal guild As String)
        '<EhHeader>
        On Error GoTo WriteGuildAllianceDetails_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGuildAllianceDetails)
102     Call Writer.WriteString8(guild)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGuildAllianceDetails_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGuildAllianceDetails", Erl)
        '</EhFooter>
End Sub

''
' Writes the "GuildPeaceDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose peace proposal's details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildPeaceDetails(ByVal guild As String)
        '<EhHeader>
        On Error GoTo WriteGuildPeaceDetails_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGuildPeaceDetails)
102     Call Writer.WriteString8(guild)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGuildPeaceDetails_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGuildPeaceDetails", Erl)
        '</EhFooter>
End Sub

''
' Writes the "GuildRequestJoinerInfo" message to the outgoing data buffer.
'
' @param    username The user who wants to join the guild whose info is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildRequestJoinerInfo(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WriteGuildRequestJoinerInfo_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGuildRequestJoinerInfo)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGuildRequestJoinerInfo_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGuildRequestJoinerInfo", Erl)
        '</EhFooter>
End Sub

''
' Writes the "GuildAlliancePropList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildAlliancePropList()
        '<EhHeader>
        On Error GoTo WriteGuildAlliancePropList_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGuildAlliancePropList)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGuildAlliancePropList_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGuildAlliancePropList", Erl)
        '</EhFooter>
End Sub

''
' Writes the "GuildPeacePropList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildPeacePropList()
        '<EhHeader>
        On Error GoTo WriteGuildPeacePropList_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGuildPeacePropList)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGuildPeacePropList_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGuildPeacePropList", Erl)
        '</EhFooter>
End Sub

''
' Writes the "GuildDeclareWar" message to the outgoing data buffer.
'
' @param    guild The guild to which to declare war.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildDeclareWar(ByVal guild As String)
        '<EhHeader>
        On Error GoTo WriteGuildDeclareWar_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGuildDeclareWar)
102     Call Writer.WriteString8(guild)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGuildDeclareWar_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGuildDeclareWar", Erl)
        '</EhFooter>
End Sub

''
' Writes the "GuildNewWebsite" message to the outgoing data buffer.
'
' @param    url The guild's new website's URL.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildNewWebsite(ByVal url As String)
        '<EhHeader>
        On Error GoTo WriteGuildNewWebsite_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGuildNewWebsite)
102     Call Writer.WriteString8(url)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGuildNewWebsite_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGuildNewWebsite", Erl)
        '</EhFooter>
End Sub

''
' Writes the "GuildAcceptNewMember" message to the outgoing data buffer.
'
' @param    username The name of the accepted player.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildAcceptNewMember(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WriteGuildAcceptNewMember_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGuildAcceptNewMember)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGuildAcceptNewMember_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGuildAcceptNewMember", Erl)
        '</EhFooter>
End Sub

''
' Writes the "GuildRejectNewMember" message to the outgoing data buffer.
'
' @param    username The name of the rejected player.
' @param    reason The reason for which the player was rejected.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildRejectNewMember(ByVal UserName As String, ByVal reason As String)
        '<EhHeader>
        On Error GoTo WriteGuildRejectNewMember_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGuildRejectNewMember)
102     Call Writer.WriteString8(UserName)
104     Call Writer.WriteString8(reason)
    
106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGuildRejectNewMember_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGuildRejectNewMember", Erl)
        '</EhFooter>
End Sub

''
' Writes the "GuildKickMember" message to the outgoing data buffer.
'
' @param    username The name of the kicked player.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildKickMember(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WriteGuildKickMember_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGuildKickMember)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGuildKickMember_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGuildKickMember", Erl)
        '</EhFooter>
End Sub

''
' Writes the "GuildUpdateNews" message to the outgoing data buffer.
'
' @param    news The news to be posted.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildUpdateNews(ByVal news As String)
        '<EhHeader>
        On Error GoTo WriteGuildUpdateNews_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGuildUpdateNews)
102     Call Writer.WriteString8(news)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGuildUpdateNews_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGuildUpdateNews", Erl)
        '</EhFooter>
End Sub

''
' Writes the "GuildMemberInfo" message to the outgoing data buffer.
'
' @param    username The user whose info is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildMemberInfo(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WriteGuildMemberInfo_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGuildMemberInfo)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGuildMemberInfo_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGuildMemberInfo", Erl)
        '</EhFooter>
End Sub

''
' Writes the "GuildOpenElections" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildOpenElections()
        '<EhHeader>
        On Error GoTo WriteGuildOpenElections_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGuildOpenElections)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGuildOpenElections_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGuildOpenElections", Erl)
        '</EhFooter>
End Sub

''
' Writes the "GuildRequestMembership" message to the outgoing data buffer.
'
' @param    guild The guild to which to request membership.
' @param    application The user's application sheet.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildRequestMembership(ByVal guild As String, ByVal Application As String)
        '<EhHeader>
        On Error GoTo WriteGuildRequestMembership_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGuildRequestMembership)
102     Call Writer.WriteString8(guild)
104     Call Writer.WriteString8(Application)
    
106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGuildRequestMembership_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGuildRequestMembership", Erl)
        '</EhFooter>
End Sub

''
' Writes the "GuildRequestDetails" message to the outgoing data buffer.
'
' @param    guild The guild whose details are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildRequestDetails(ByVal guild As String)
        '<EhHeader>
        On Error GoTo WriteGuildRequestDetails_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGuildRequestDetails)
102     Call Writer.WriteString8(guild)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGuildRequestDetails_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGuildRequestDetails", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Online" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteOnline()
        '<EhHeader>
        On Error GoTo WriteOnline_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eOnline)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteOnline_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteOnline", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Quit" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteQuit()
        '<EhHeader>
        On Error GoTo WriteQuit_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eQuit)
102     Call modNetwork.Send(Writer)
    
104     UserSaliendo = True
        '<EhFooter>
        Exit Sub

WriteQuit_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteQuit", Erl)
        '</EhFooter>
End Sub

''
' Writes the "GuildLeave" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildLeave()
        '<EhHeader>
        On Error GoTo WriteGuildLeave_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGuildLeave)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGuildLeave_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGuildLeave", Erl)
        '</EhFooter>
End Sub

''
' Writes the "RequestAccountState" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRequestAccountState()
        '<EhHeader>
        On Error GoTo WriteRequestAccountState_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eRequestAccountState)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteRequestAccountState_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteRequestAccountState", Erl)
        '</EhFooter>
End Sub

''
' Writes the "PetStand" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePetStand()
        '<EhHeader>
        On Error GoTo WritePetStand_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.ePetStand)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WritePetStand_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WritePetStand", Erl)
        '</EhFooter>
End Sub

''
' Writes the "PetFollow" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePetFollow()
        '<EhHeader>
        On Error GoTo WritePetFollow_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.ePetFollow)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WritePetFollow_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WritePetFollow", Erl)
        '</EhFooter>
End Sub

''
' Writes the "PetLeave" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePetLeave()
        '<EhHeader>
        On Error GoTo WritePetLeave_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.ePetLeave)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WritePetLeave_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WritePetLeave", Erl)
        '</EhFooter>
End Sub

''
' Writes the "GrupoMsg" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGrupoMsg(ByVal Message As String)
        '<EhHeader>
        On Error GoTo WriteGrupoMsg_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGrupoMsg)
102     Call Writer.WriteString8(Message)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGrupoMsg_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGrupoMsg", Erl)
        '</EhFooter>
End Sub

''
' Writes the "TrainList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteTrainList()
        '<EhHeader>
        On Error GoTo WriteTrainList_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eTrainList)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteTrainList_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteTrainList", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Rest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRest()
        '<EhHeader>
        On Error GoTo WriteRest_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eRest)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteRest_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteRest", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Meditate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteMeditate()
        '<EhHeader>
        On Error GoTo WriteMeditate_Err
        '</EhHeader>
        
        If UserMoving Then Exit Sub
        
100     Call Writer.WriteInt16(ClientPacketID.eMeditate)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteMeditate_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteMeditate", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Resucitate" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteResucitate()
        '<EhHeader>
        On Error GoTo WriteResucitate_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eResucitate)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteResucitate_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteResucitate", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Heal" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteHeal()
        '<EhHeader>
        On Error GoTo WriteHeal_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eHeal)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteHeal_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteHeal", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Help" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteHelp()
        '<EhHeader>
        On Error GoTo WriteHelp_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eHelp)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteHelp_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteHelp", Erl)
        '</EhFooter>
End Sub


Public Sub WriteEventoFaccionario()
        On Error GoTo WriteEventoFaccionario_Err
100     Call Writer.WriteInt16(ClientPacketID.eEventoFaccionario)
    
102     Call modNetwork.Send(Writer)
        Exit Sub

WriteEventoFaccionario_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteEventoFaccionario", Erl)
        '</EhFooter>
End Sub

''
' Writes the "RequestStats" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRequestStats()
        '<EhHeader>
        On Error GoTo WriteRequestStats_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eRequestStats)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteRequestStats_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteRequestStats", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Promedio" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePromedio()
        '<EhHeader>
        On Error GoTo WritePromedio_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.ePromedio)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WritePromedio_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WritePromedio", Erl)
        '</EhFooter>
End Sub

''
' Writes the "GiveItem" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGiveItem(UserName As String, _
                         ByVal ObjIndex As Integer, _
                         ByVal cantidad As Integer, _
                         Motivo As String)
        '<EhHeader>
        On Error GoTo WriteGiveItem_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGiveItem)
102     Call Writer.WriteString8(UserName)
104     Call Writer.WriteInt16(ObjIndex)
106     Call Writer.WriteInt16(cantidad)
108     Call Writer.WriteString8(Motivo)
    
110     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGiveItem_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGiveItem", Erl)
        '</EhFooter>
End Sub

''
' Writes the "CommerceStart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCommerceStart()
        '<EhHeader>
        On Error GoTo WriteCommerceStart_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eCommerceStart)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCommerceStart_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCommerceStart", Erl)
        '</EhFooter>
End Sub

''
' Writes the "BankStart" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBankStart()
        '<EhHeader>
        On Error GoTo WriteBankStart_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eBankStart)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteBankStart_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteBankStart", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Enlist" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteEnlist()
        '<EhHeader>
        On Error GoTo WriteEnlist_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eEnlist)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteEnlist_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteEnlist", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Information" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteInformation()
        '<EhHeader>
        On Error GoTo WriteInformation_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eInformation)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteInformation_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteInformation", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Reward" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteReward()
        '<EhHeader>
        On Error GoTo WriteReward_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eReward)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteReward_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteReward", Erl)
        '</EhFooter>
End Sub

''
' Writes the "RequestMOTD" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRequestMOTD()
        '<EhHeader>
        On Error GoTo WriteRequestMOTD_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eRequestMOTD)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteRequestMOTD_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteRequestMOTD", Erl)
        '</EhFooter>
End Sub

''
' Writes the "UpTime" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUpTime()
        '<EhHeader>
        On Error GoTo WriteUpTime_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eUpTime)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteUpTime_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteUpTime", Erl)
        '</EhFooter>
End Sub


''
' Writes the "GuildMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildMessage(ByVal Message As String)
        '<EhHeader>
        On Error GoTo WriteGuildMessage_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGuildMessage)
102     Call Writer.WriteString8(Message)
        packetCounters.TS_GuildMessage = packetCounters.TS_GuildMessage + 1
        Call Writer.WriteInt32(packetCounters.TS_GuildMessage)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGuildMessage_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGuildMessage", Erl)
        '</EhFooter>
End Sub

''
' Writes the "GuildOnline" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildOnline()
        '<EhHeader>
        On Error GoTo WriteGuildOnline_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGuildOnline)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGuildOnline_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGuildOnline", Erl)
        '</EhFooter>
End Sub

''
' Writes the "CouncilMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the other council members.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCouncilMessage(ByVal Message As String)
        '<EhHeader>
        On Error GoTo WriteCouncilMessage_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eCouncilMessage)
102     Call Writer.WriteString8(Message)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCouncilMessage_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCouncilMessage", Erl)
        '</EhFooter>
End Sub

''
' Writes the "RoleMasterRequest" message to the outgoing data buffer.
'
' @param    message The message to send to the role masters.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRoleMasterRequest(ByVal Message As String)
        '<EhHeader>
        On Error GoTo WriteRoleMasterRequest_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eRoleMasterRequest)
102     Call Writer.WriteString8(Message)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteRoleMasterRequest_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteRoleMasterRequest", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ChangeDescription" message to the outgoing data buffer.
'
' @param    desc The new description of the user's character.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeDescription(ByVal desc As String)
        '<EhHeader>
        On Error GoTo WriteChangeDescription_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eChangeDescription)
102     Call Writer.WriteString8(desc)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteChangeDescription_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteChangeDescription", Erl)
        '</EhFooter>
End Sub

''
' Writes the "GuildVote" message to the outgoing data buffer.
'
' @param    username The user to vote for clan leader.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildVote(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WriteGuildVote_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGuildVote)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGuildVote_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGuildVote", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Punishments" message to the outgoing data buffer.
'
' @param    username The user whose's  punishments are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WritePunishments(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WritePunishments_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.epunishments)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WritePunishments_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WritePunishments", Erl)
        '</EhFooter>
End Sub


''
' Writes the "Gamble" message to the outgoing data buffer.
'
' @param    amount The amount to gamble.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGamble(ByVal Amount As Integer)
        '<EhHeader>
        On Error GoTo WriteGamble_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGamble)
102     Call Writer.WriteInt16(Amount)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGamble_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGamble", Erl)
        '</EhFooter>
End Sub


''
' Writes the "LeaveFaction" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteLeaveFaction()
        '<EhHeader>
        On Error GoTo WriteLeaveFaction_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eLeaveFaction)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteLeaveFaction_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteLeaveFaction", Erl)
        '</EhFooter>
End Sub

''
' Writes the "BankExtractGold" message to the outgoing data buffer.
'
' @param    amount The amount of money to extract from the bank.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBankExtractGold(ByVal Amount As Long)
        '<EhHeader>
        On Error GoTo WriteBankExtractGold_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eBankExtractGold)
102     Call Writer.WriteInt32(Amount)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteBankExtractGold_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteBankExtractGold", Erl)
        '</EhFooter>
End Sub

''
' Writes the "BankDepositGold" message to the outgoing data buffer.
'
' @param    amount The amount of money to deposit in the bank.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBankDepositGold(ByVal Amount As Long)
        '<EhHeader>
        On Error GoTo WriteBankDepositGold_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eBankDepositGold)
102     Call Writer.WriteInt32(Amount)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteBankDepositGold_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteBankDepositGold", Erl)
        '</EhFooter>
End Sub

Public Sub WriteTransFerGold(ByVal Amount As Long, ByVal destino As String)
        '<EhHeader>
        On Error GoTo WriteTransFerGold_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eTransFerGold)
102     Call Writer.WriteInt32(Amount)
104     Call Writer.WriteString8(destino)
    Call modNetwork.Send(Writer)
    
        '<EhFooter>
        Exit Sub

WriteTransFerGold_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteTransFerGold", Erl)
        '</EhFooter>
End Sub

Public Sub WriteItemMove(ByVal SlotActual As Byte, ByVal SlotNuevo As Byte)
        '<EhHeader>
        On Error GoTo WriteItemMove_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eMoveitem)
102     Call Writer.WriteInt8(SlotActual)
104     Call Writer.WriteInt8(SlotNuevo)
    
106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteItemMove_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteItemMove", Erl)
        '</EhFooter>
End Sub

Public Sub WriteNotifyInventarioHechizos(ByVal value As Byte, ByVal hechiSel As Byte, ByVal scrollSel As Byte)
        '<EhHeader>
        On Error GoTo NotifyInventarioHechizos_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eNotifyInventarioHechizos)
104     Call Writer.WriteInt8(value)
        Call Writer.WriteInt8(hechiSel)
        Call Writer.WriteInt8(scrollSel)
    
106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

NotifyInventarioHechizos_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.NotifyInventarioHechizos", Erl)
        '</EhFooter>
End Sub



Public Sub WriteBovedaItemMove(ByVal SlotActual As Byte, ByVal SlotNuevo As Byte)
        '<EhHeader>
        On Error GoTo WriteBovedaItemMove_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eBovedaMoveItem)
102     Call Writer.WriteInt8(SlotActual)
104     Call Writer.WriteInt8(SlotNuevo)
    
106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteBovedaItemMove_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteBovedaItemMove", Erl)
        '</EhFooter>
End Sub

''
' Writes the "FinEvento" message to the outgoing data buffer.
'
' @param    message The message to s the denounce.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteFinEvento()
        '<EhHeader>
        On Error GoTo WriteFinEvento_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eFinEvento)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteFinEvento_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteFinEvento", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Denounce" message to the outgoing data buffer.
'
' @param    message The message to s the denounce.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteDenounce(Name As String)
        '<EhHeader>
        On Error GoTo WriteDenounce_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eDenounce)
102     Call Writer.WriteString8(Name)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteDenounce_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteDenounce", Erl)
        '</EhFooter>
End Sub

Public Sub WriteQuieroFundarClan()
        '<EhHeader>
        On Error GoTo WriteQuieroFundarClan_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eQuieroFundarClan)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteQuieroFundarClan_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteQuieroFundarClan", Erl)
        '</EhFooter>
End Sub

''
' Writes the "GuildMemberList" message to the outgoing data buffer.
'
' @param    guild The guild whose member list is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildMemberList(ByVal guild As String)
        '<EhHeader>
        On Error GoTo WriteGuildMemberList_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGuildMemberList)
102     Call Writer.WriteString8(guild)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGuildMemberList_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGuildMemberList", Erl)
        '</EhFooter>
End Sub

Public Sub WriteCasamiento(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WriteCasamiento_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eCasarse)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCasamiento_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCasamiento", Erl)
        '</EhFooter>
End Sub

Public Sub WriteMacroPos()
        '<EhHeader>
        On Error GoTo WriteMacroPos_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eMacroPossent)
102     Call Writer.WriteInt8(ChatCombate)
104     Call Writer.WriteInt8(ChatGlobal)
    
106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteMacroPos_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteMacroPos", Erl)
        '</EhFooter>
End Sub

Public Sub WriteSubastaInfo()
        '<EhHeader>
        On Error GoTo WriteSubastaInfo_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eSubastaInfo)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteSubastaInfo_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteSubastaInfo", Erl)
        '</EhFooter>
End Sub

Public Sub WriteCancelarExit()
        '<EhHeader>
        On Error GoTo WriteCancelarExit_Err
        '</EhHeader>
100     UserSaliendo = False
102     Call Writer.WriteInt16(ClientPacketID.eCancelarExit)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCancelarExit_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCancelarExit", Erl)
        '</EhFooter>
End Sub

Public Sub WriteEventoInfo()
        '<EhHeader>
        On Error GoTo WriteEventoInfo_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eEventoInfo)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteEventoInfo_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteEventoInfo", Erl)
        '</EhFooter>
End Sub

Public Sub WriteFlagTrabajar()
        '<EhHeader>
        On Error GoTo WriteFlagTrabajar_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eFlagTrabajar)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteFlagTrabajar_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteFlagTrabajar", Erl)
        '</EhFooter>
End Sub


Public Sub WriteGMMessage(ByVal Message As String)
        '<EhHeader>
        On Error GoTo WriteGMMessage_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGMMessage)
102     Call Writer.WriteString8(Message)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGMMessage_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGMMessage", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ShowName" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowName()
        '<EhHeader>
        On Error GoTo WriteShowName_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eshowName)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteShowName_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteShowName", Erl)
        '</EhFooter>
End Sub

''
' Writes the "OnlineRoyalArmy" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteOnlineRoyalArmy()
        '<EhHeader>
        On Error GoTo WriteOnlineRoyalArmy_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eOnlineRoyalArmy)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteOnlineRoyalArmy_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteOnlineRoyalArmy", Erl)
        '</EhFooter>
End Sub

''
' Writes the "OnlineChaosLegion" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteOnlineChaosLegion()
        '<EhHeader>
        On Error GoTo WriteOnlineChaosLegion_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eOnlineChaosLegion)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteOnlineChaosLegion_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteOnlineChaosLegion", Erl)
        '</EhFooter>
End Sub

''
' Writes the "GoNearby" message to the outgoing data buffer.
'
' @param    username The suer to approach.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGoNearby(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WriteGoNearby_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGoNearby)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGoNearby_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGoNearby", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Comment" message to the outgoing data buffer.
'
' @param    message The message to leave in the log as a comment.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteComment(ByVal Message As String)
        '<EhHeader>
        On Error GoTo WriteComment_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.ecomment)
102     Call Writer.WriteString8(Message)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteComment_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteComment", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ServerTime" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteServerTime()
        '<EhHeader>
        On Error GoTo WriteServerTime_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eserverTime)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteServerTime_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteServerTime", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Where" message to the outgoing data buffer.
'
' @param    username The user whose position is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteWhere(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WriteWhere_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eWhere)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteWhere_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteWhere", Erl)
        '</EhFooter>
End Sub

''
' Writes the "CreaturesInMap" message to the outgoing data buffer.
'
' @param    map The map in which to check for the existing creatures.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCreaturesInMap(ByVal map As Integer)
        '<EhHeader>
        On Error GoTo WriteCreaturesInMap_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eCreaturesInMap)
102     Call Writer.WriteInt16(map)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCreaturesInMap_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCreaturesInMap", Erl)
        '</EhFooter>
End Sub

''
' Writes the "WarpMeToTarget" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteWarpMeToTarget()
        '<EhHeader>
        On Error GoTo WriteWarpMeToTarget_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eWarpMeToTarget)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteWarpMeToTarget_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteWarpMeToTarget", Erl)
        '</EhFooter>
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
           
        If EstaSiguiendo() Then Exit Sub
        
        '<EhHeader>
        On Error GoTo WriteWarpChar_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eWarpChar)
102     Call Writer.WriteString8(UserName)
104     Call Writer.WriteInt16(map)
106     Call Writer.WriteInt8(x)
108     Call Writer.WriteInt8(y)
    
110     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteWarpChar_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteWarpChar", Erl)
        '</EhFooter>
End Sub

Public Sub WriteStartLobby(ByVal LobbyType As Byte, ByRef LobbySettings As t_NewScenearioSettings, ByVal Description As String, ByVal Password As String)
On Error GoTo WriteStartLobby_Err
100     Call Writer.WriteInt16(ClientPacketID.eStartEvent)
102     Call Writer.WriteInt8(lobbyType)
103     Call Writer.WriteInt8(LobbySettings.ScenearioType)
104     Call Writer.WriteInt8(LobbySettings.MinLevel)
106     Call Writer.WriteInt8(LobbySettings.MaxLevel)
108     Call Writer.WriteInt8(LobbySettings.MinPlayers)
110     Call Writer.WriteInt8(LobbySettings.MaxPlayers)
112     Call Writer.WriteInt8(LobbySettings.TeamSize)
114     Call Writer.WriteInt8(LobbySettings.TeamType)
115     Call Writer.WriteInt8(LobbySettings.RoundAmount)
116     Call Writer.WriteInt32(LobbySettings.InscriptionFee)
118     Call Writer.WriteString8(Description)
120     Call Writer.WriteString8(Password)
122     Call modNetwork.Send(Writer)
        Exit Sub
WriteStartLobby_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteStartLobby", Erl)
    
End Sub

Public Sub WriteCancelarEvento()
On Error GoTo WriteCancelarCaptura_Err
   
100     Call Writer.WriteInt16(ClientPacketID.eCancelarEvento)
110     Call modNetwork.Send(Writer)
        Exit Sub

WriteCancelarCaptura_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCancelarCaptura", Erl)
End Sub

''
' Writes the "Silence" message to the outgoing data buffer.
'
' @param    username The user to silence.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSilence(ByVal UserName As String, ByVal Minutos As Integer)
        '<EhHeader>
        On Error GoTo WriteSilence_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eSilence)
102     Call Writer.WriteString8(UserName)
104     Call Writer.WriteInt16(Minutos)
    
106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteSilence_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteSilence", Erl)
        '</EhFooter>
End Sub

Public Sub WriteCuentaRegresiva(ByVal Second As Byte)
        '<EhHeader>
        On Error GoTo WriteCuentaRegresiva_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eCuentaRegresiva)
102     Call Writer.WriteInt8(Second)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCuentaRegresiva_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCuentaRegresiva", Erl)
        '</EhFooter>
End Sub

Public Sub WritePossUser(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WritePossUser_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.ePossUser)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WritePossUser_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WritePossUser", Erl)
        '</EhFooter>
End Sub

''
' Writes the "SOSShowList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSOSShowList()
        '<EhHeader>
        On Error GoTo WriteSOSShowList_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eSOSShowList)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteSOSShowList_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteSOSShowList", Erl)
        '</EhFooter>
End Sub

''
' Writes the "SOSRemove" message to the outgoing data buffer.
'
' @param    username The user whose SOS call has been already attended.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSOSRemove(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WriteSOSRemove_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eSOSRemove)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteSOSRemove_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteSOSRemove", Erl)
        '</EhFooter>
End Sub

''
' Writes the "GoToChar" message to the outgoing data buffer.
'
' @param    username The user to be approached.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGoToChar(ByVal UserName As String)

        
        If EstaSiguiendo() Then Exit Sub
        '<EhHeader>
        On Error GoTo WriteGoToChar_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGoToChar)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGoToChar_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGoToChar", Erl)
        '</EhFooter>
End Sub


''
' Writes the "Invisible" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteInvisible()
        '<EhHeader>
        On Error GoTo WriteInvisible_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eInvisible)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteInvisible_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteInvisible", Erl)
        '</EhFooter>
End Sub

''
' Writes the "GMPanel" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGMPanel()
        '<EhHeader>
        On Error GoTo WriteGMPanel_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGMPanel)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGMPanel_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGMPanel", Erl)
        '</EhFooter>
End Sub

''
' Writes the "RequestUserList" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRequestUserList()
        '<EhHeader>
        On Error GoTo WriteRequestUserList_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eRequestUserList)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteRequestUserList_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteRequestUserList", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Working" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteWorking()
        '<EhHeader>
        On Error GoTo WriteWorking_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eWorking)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteWorking_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteWorking", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Hiding" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteHiding()
        '<EhHeader>
        On Error GoTo WriteHiding_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eHiding)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteHiding_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteHiding", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Jail" message to the outgoing data buffer.
'
' @param    username The user to be sent to jail.
' @param    reason The reason for which to send him to jail.
' @param    time The time (in minutes) the user will have to spend there.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteJail(ByVal UserName As String, ByVal reason As String, ByVal Time As Byte)
        '<EhHeader>
        On Error GoTo WriteJail_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eJail)
102     Call Writer.WriteString8(UserName)
104     Call Writer.WriteString8(reason)
106     Call Writer.WriteInt8(Time)
    
108     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteJail_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteJail", Erl)
        '</EhFooter>
End Sub

Public Sub WriteCrearEvento(ByVal TIPO As Byte, _
                            ByVal duracion As Byte, _
                            ByVal multiplicacion As Byte)
        '<EhHeader>
        On Error GoTo WriteCrearEvento_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eCrearEvento)
102     Call Writer.WriteInt8(TIPO)
104     Call Writer.WriteInt8(duracion)
106     Call Writer.WriteInt8(multiplicacion)
    
108     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCrearEvento_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCrearEvento", Erl)
        '</EhFooter>
End Sub

''
' Writes the "KillNPC" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteKillNPC()
        '<EhHeader>
        On Error GoTo WriteKillNPC_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eKillNPC)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteKillNPC_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteKillNPC", Erl)
        '</EhFooter>
End Sub

''
' Writes the "WarnUser" message to the outgoing data buffer.
'
' @param    username The user to be warned.
' @param    reason Reason for the warning.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteWarnUser(ByVal UserName As String, ByVal reason As String)
        '<EhHeader>
        On Error GoTo WriteWarnUser_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eWarnUser)
102     Call Writer.WriteString8(UserName)
104     Call Writer.WriteString8(reason)
    
106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteWarnUser_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteWarnUser", Erl)
        '</EhFooter>
End Sub

Public Sub WriteMensajeUser(ByVal UserName As String, ByVal mensaje As String)
        '<EhHeader>
        On Error GoTo WriteMensajeUser_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eMensajeUser)
102     Call Writer.WriteString8(UserName)
104     Call Writer.WriteString8(mensaje)
    
106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteMensajeUser_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteMensajeUser", Erl)
        '</EhFooter>
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
        '<EhHeader>
        On Error GoTo WriteEditChar_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eEditChar)
102     Call Writer.WriteString8(UserName)
104     Call Writer.WriteInt8(editOption)
106     Call Writer.WriteString8(arg1)
108     Call Writer.WriteString8(arg2)
    
110     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteEditChar_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteEditChar", Erl)
        '</EhFooter>
End Sub

''
' Writes the "RequestCharInfo" message to the outgoing data buffer.
'
' @param    username The user whose information is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRequestCharInfo(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WriteRequestCharInfo_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eRequestCharInfo)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteRequestCharInfo_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteRequestCharInfo", Erl)
        '</EhFooter>
End Sub

''
' Writes the "RequestCharStats" message to the outgoing data buffer.
'
' @param    username The user whose stats are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRequestCharStats(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WriteRequestCharStats_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eRequestCharStats)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteRequestCharStats_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteRequestCharStats", Erl)
        '</EhFooter>
End Sub

''
' Writes the "RequestCharGold" message to the outgoing data buffer.
'
' @param    username The user whose gold is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRequestCharGold(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WriteRequestCharGold_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eRequestCharGold)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteRequestCharGold_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteRequestCharGold", Erl)
        '</EhFooter>
End Sub
    
''
' Writes the "RequestCharInventory" message to the outgoing data buffer.
'
' @param    username The user whose inventory is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRequestCharInventory(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WriteRequestCharInventory_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eRequestCharInventory)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteRequestCharInventory_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteRequestCharInventory", Erl)
        '</EhFooter>
End Sub

''
' Writes the "RequestCharBank" message to the outgoing data buffer.
'
' @param    username The user whose banking information is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRequestCharBank(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WriteRequestCharBank_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eRequestCharBank)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteRequestCharBank_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteRequestCharBank", Erl)
        '</EhFooter>
End Sub

''
' Writes the "RequestCharSkills" message to the outgoing data buffer.
'
' @param    username The user whose skills are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRequestCharSkills(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WriteRequestCharSkills_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eRequestCharSkills)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteRequestCharSkills_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteRequestCharSkills", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ReviveChar" message to the outgoing data buffer.
'
' @param    username The user to eb revived.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteReviveChar(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WriteReviveChar_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eReviveChar)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteReviveChar_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteReviveChar", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ReviveChar" message to the outgoing data buffer.
'
' @param    username The user to eb revived.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSeguirMouse(ByVal username As String)
        '<EhHeader>
        On Error GoTo WriteSeguirMouse_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eSeguirMouse)
102     Call Writer.WriteString8(username)

104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteSeguirMouse_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteSeguirMouse", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ReviveChar" message to the outgoing data buffer.
'
' @param    username The user to eb revived.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSendPosSeguimiento(ByVal Cheat_X As Integer, ByVal Cheat_Y As Integer)
        '<EhHeader>
        On Error GoTo WriteSendPosSeguimiento_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eSendPosSeguimiento)
102     Call Writer.WriteString16(Cheat_X)
103     Call Writer.WriteString16(Cheat_Y)

104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteSendPosSeguimiento_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteSendPosSeguimiento", Erl)
        '</EhFooter>
End Sub


'HarThaoS: Perdón caótico
Public Sub WritePerdonFaccion(ByVal username As String)
        '<EhHeader>
        On Error GoTo WritePerdonFaccion_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.ePerdonFaccion)
102     Call Writer.WriteString8(username)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WritePerdonFaccion_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WritePerdonFaccion", Erl)
        '</EhFooter>
End Sub

''
' Writes the "OnlineGM" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteOnlineGM()
        '<EhHeader>
        On Error GoTo WriteOnlineGM_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eOnlineGM)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteOnlineGM_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteOnlineGM", Erl)
        '</EhFooter>
End Sub

''
' Writes the "OnlineMap" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteOnlineMap()
        '<EhHeader>
        On Error GoTo WriteOnlineMap_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eOnlineMap)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteOnlineMap_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteOnlineMap", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Forgive" message to the outgoing data buffer.
'
' @param    username The user to be forgiven.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteForgive()
        '<EhHeader>
        On Error GoTo WriteForgive_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eForgive)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteForgive_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteForgive", Erl)
        '</EhFooter>
End Sub

Public Sub WriteDonateGold(ByVal oro As Long)
        '<EhHeader>
        On Error GoTo WriteDonateGold_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eDonateGold)
102     Call Writer.WriteInt32(oro)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteDonateGold_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteDonateGold", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Kick" message to the outgoing data buffer.
'
' @param    username The user to be kicked.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteKick(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WriteKick_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eKick)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteKick_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteKick", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Execute" message to the outgoing data buffer.
'
' @param    username The user to be executed.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteExecute(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WriteExecute_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eExecute)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteExecute_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteExecute", Erl)
        '</EhFooter>
End Sub

''
' Writes the "BanChar" message to the outgoing data buffer.
'
' @param    username The user to be banned.
' @param    reason The reson for which the user is to be banned.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteBanChar(ByVal UserName As String, ByVal reason As String)
        '<EhHeader>
        On Error GoTo WriteBanChar_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eBanChar)
102     Call Writer.WriteString8(UserName)
104     Call Writer.WriteString8(reason)
    
106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteBanChar_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteBanChar", Erl)
        '</EhFooter>
End Sub

Public Sub WriteBanCuenta(ByVal UserName As String, ByVal reason As String)
        '<EhHeader>
        On Error GoTo WriteBanCuenta_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eBanCuenta)
102     Call Writer.WriteString8(UserName)
104     Call Writer.WriteString8(reason)
    
106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteBanCuenta_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteBanCuenta", Erl)
        '</EhFooter>
End Sub

Public Sub WriteUnBanCuenta(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WriteUnBanCuenta_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eUnbanCuenta)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteUnBanCuenta_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteUnBanCuenta", Erl)
        '</EhFooter>
End Sub

Public Sub WriteCerraCliente(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WriteCerraCliente_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eCerrarCliente)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCerraCliente_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCerraCliente", Erl)
        '</EhFooter>
End Sub

Public Sub WriteBanTemporal(ByVal UserName As String, _
                            ByVal reason As String, _
                            ByVal dias As Byte)
        '<EhHeader>
        On Error GoTo WriteBanTemporal_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eBanTemporal)
102     Call Writer.WriteString8(UserName)
104     Call Writer.WriteString8(reason)
106     Call Writer.WriteInt8(dias)
    
108     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteBanTemporal_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteBanTemporal", Erl)
        '</EhFooter>
End Sub

''
' Writes the "UnbanChar" message to the outgoing data buffer.
'
' @param    username The user to be unbanned.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteUnbanChar(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WriteUnbanChar_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eUnbanChar)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteUnbanChar_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteUnbanChar", Erl)
        '</EhFooter>
End Sub

''
' Writes the "NPCFollow" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteNPCFollow()
        '<EhHeader>
        On Error GoTo WriteNPCFollow_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eNPCFollow)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteNPCFollow_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteNPCFollow", Erl)
        '</EhFooter>
End Sub

''
' Writes the "SummonChar" message to the outgoing data buffer.
'
' @param    username The user to be summoned.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSummonChar(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WriteSummonChar_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eSummonChar)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteSummonChar_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteSummonChar", Erl)
        '</EhFooter>
End Sub

''
' Writes the "SpawnListRequest" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSpawnListRequest()
        '<EhHeader>
        On Error GoTo WriteSpawnListRequest_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eSpawnListRequest)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteSpawnListRequest_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteSpawnListRequest", Erl)
        '</EhFooter>
End Sub

''
' Writes the "SpawnCreature" message to the outgoing data buffer.
'
' @param    creatureIndex The index of the creature in the spawn list to be spawned.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSpawnCreature(ByVal creatureIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteSpawnCreature_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eSpawnCreature)
102     Call Writer.WriteInt16(creatureIndex)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteSpawnCreature_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteSpawnCreature", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ResetNPCInventory" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteResetNPCInventory()
        '<EhHeader>
        On Error GoTo WriteResetNPCInventory_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eResetNPCInventory)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteResetNPCInventory_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteResetNPCInventory", Erl)
        '</EhFooter>
End Sub

''
' Writes the "CleanWorld" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCleanWorld()
        '<EhHeader>
        On Error GoTo WriteCleanWorld_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eCleanWorld)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCleanWorld_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCleanWorld", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ServerMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteServerMessage(ByVal Message As String)
        '<EhHeader>
        On Error GoTo WriteServerMessage_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eServerMessage)
102     Call Writer.WriteString8(Message)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteServerMessage_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteServerMessage", Erl)
        '</EhFooter>
End Sub

''
' Writes the "NickToIP" message to the outgoing data buffer.
'
' @param    username The user whose IP is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteNickToIP(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WriteNickToIP_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eNickToIP)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteNickToIP_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteNickToIP", Erl)
        '</EhFooter>
End Sub

''
' Writes the "IPToNick" message to the outgoing data buffer.
'
' @param    IP The IP for which to search for players. Must be an array of 4 elements with the 4 components of the IP.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteIPToNick(ByRef IP() As Byte)
        '<EhHeader>
        On Error GoTo WriteIPToNick_Err
        '</EhHeader>

100     If UBound(IP()) - LBound(IP()) + 1 <> 4 Then Exit Sub   'Invalid IP

        Dim i As Long

102     Call Writer.WriteInt16(ClientPacketID.eIPToNick)

104     For i = LBound(IP()) To UBound(IP())
106         Call Writer.WriteInt8(IP(i))
108     Next i

110     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteIPToNick_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteIPToNick", Erl)
        '</EhFooter>
End Sub

''
' Writes the "GuildOnlineMembers" message to the outgoing data buffer.
'
' @param    guild The guild whose online player list is requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildOnlineMembers(ByVal guild As String)
        '<EhHeader>
        On Error GoTo WriteGuildOnlineMembers_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGuildOnlineMembers)
102     Call Writer.WriteString8(guild)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGuildOnlineMembers_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGuildOnlineMembers", Erl)
        '</EhFooter>
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
                               ByVal Radio As Byte, _
                               ByVal Motivo As String)
        '<EhHeader>
        On Error GoTo WriteTeleportCreate_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eTeleportCreate)
102     Call Writer.WriteInt16(map)
104     Call Writer.WriteInt8(x)
106     Call Writer.WriteInt8(y)
107     Call Writer.WriteInt8(Radio)
108     Call Writer.WriteString8(Motivo)
    
110     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteTeleportCreate_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteTeleportCreate", Erl)
        '</EhFooter>
End Sub

''
' Writes the "TeleportDestroy" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteTeleportDestroy()
        '<EhHeader>
        On Error GoTo WriteTeleportDestroy_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eTeleportDestroy)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteTeleportDestroy_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteTeleportDestroy", Erl)
        '</EhFooter>
End Sub

''
' Writes the "RainToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRainToggle()
        '<EhHeader>
        On Error GoTo WriteRainToggle_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eRainToggle)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteRainToggle_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteRainToggle", Erl)
        '</EhFooter>
End Sub

''
' Writes the "SetCharDescription" message to the outgoing data buffer.
'
' @param    desc The description to set to players.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSetCharDescription(ByVal desc As String)
        '<EhHeader>
        On Error GoTo WriteSetCharDescription_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eSetCharDescription)
102     Call Writer.WriteString8(desc)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteSetCharDescription_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteSetCharDescription", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ForceMIDIToMap" message to the outgoing data buffer.
'
' @param    midiID The ID of the midi file to play.
' @param    map The map in which to play the given midi.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteForceMIDIToMap(ByVal midiID As Byte, ByVal map As Integer)
        '<EhHeader>
        On Error GoTo WriteForceMIDIToMap_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eForceMIDIToMap)
102     Call Writer.WriteInt8(midiID)
104     Call Writer.WriteInt16(map)
    
106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteForceMIDIToMap_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteForceMIDIToMap", Erl)
        '</EhFooter>
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
        '<EhHeader>
        On Error GoTo WriteForceWAVEToMap_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eForceWAVEToMap)
102     Call Writer.WriteInt8(waveID)
104     Call Writer.WriteInt16(map)
106     Call Writer.WriteInt8(x)
108     Call Writer.WriteInt8(y)
    
110     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteForceWAVEToMap_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteForceWAVEToMap", Erl)
        '</EhFooter>
End Sub

''
' Writes the "RoyalArmyMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the royal army members.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRoyalArmyMessage(ByVal Message As String)
        '<EhHeader>
        On Error GoTo WriteRoyalArmyMessage_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eRoyalArmyMessage)
102     Call Writer.WriteString8(Message)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteRoyalArmyMessage_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteRoyalArmyMessage", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ChaosLegionMessage" message to the outgoing data buffer.
'
' @param    message The message to send to the chaos legion member.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChaosLegionMessage(ByVal Message As String)
        '<EhHeader>
        On Error GoTo WriteChaosLegionMessage_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eChaosLegionMessage)
102     Call Writer.WriteString8(Message)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteChaosLegionMessage_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteChaosLegionMessage", Erl)
        '</EhFooter>
End Sub

''
' Writes the "TalkAsNPC" message to the outgoing data buffer.
'
' @param    message The message to send to the royal army members.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteTalkAsNPC(ByVal Message As String)
        '<EhHeader>
        On Error GoTo WriteTalkAsNPC_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eTalkAsNPC)
102     Call Writer.WriteString8(Message)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteTalkAsNPC_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteTalkAsNPC", Erl)
        '</EhFooter>
End Sub

''
' Writes the "DestroyAllItemsInArea" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteDestroyAllItemsInArea()
        '<EhHeader>
        On Error GoTo WriteDestroyAllItemsInArea_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eDestroyAllItemsInArea)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteDestroyAllItemsInArea_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteDestroyAllItemsInArea", Erl)
        '</EhFooter>
End Sub

''
' Writes the "AcceptRoyalCouncilMember" message to the outgoing data buffer.
'
' @param    username The name of the user to be accepted into the royal army council.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteAcceptRoyalCouncilMember(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WriteAcceptRoyalCouncilMember_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eAcceptRoyalCouncilMember)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteAcceptRoyalCouncilMember_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteAcceptRoyalCouncilMember", Erl)
        '</EhFooter>
End Sub

''
' Writes the "AcceptChaosCouncilMember" message to the outgoing data buffer.
'
' @param    username The name of the user to be accepted as a chaos council member.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteAcceptChaosCouncilMember(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WriteAcceptChaosCouncilMember_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eAcceptChaosCouncilMember)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteAcceptChaosCouncilMember_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteAcceptChaosCouncilMember", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ItemsInTheFloor" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteItemsInTheFloor()
        '<EhHeader>
        On Error GoTo WriteItemsInTheFloor_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eItemsInTheFloor)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteItemsInTheFloor_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteItemsInTheFloor", Erl)
        '</EhFooter>
End Sub

''
' Writes the "MakeDumb" message to the outgoing data buffer.
'
' @param    username The name of the user to be made dumb.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteMakeDumb(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WriteMakeDumb_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eMakeDumb)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteMakeDumb_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteMakeDumb", Erl)
        '</EhFooter>
End Sub

''
' Writes the "MakeDumbNoMore" message to the outgoing data buffer.
'
' @param    username The name of the user who will no longer be dumb.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteMakeDumbNoMore(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WriteMakeDumbNoMore_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eMakeDumbNoMore)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteMakeDumbNoMore_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteMakeDumbNoMore", Erl)
        '</EhFooter>
End Sub

''
' Writes the "CouncilKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the council.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCouncilKick(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WriteCouncilKick_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eCouncilKick)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCouncilKick_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCouncilKick", Erl)
        '</EhFooter>
End Sub

''
' Writes the "SetTrigger" message to the outgoing data buffer.
'
' @param    trigger The type of trigger to be set to the tile.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSetTrigger(ByVal Trigger As eTrigger)
        '<EhHeader>
        On Error GoTo WriteSetTrigger_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eSetTrigger)
102     Call Writer.WriteInt8(Trigger)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteSetTrigger_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteSetTrigger", Erl)
        '</EhFooter>
End Sub

''
' Writes the "AskTrigger" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteAskTrigger()
        '<EhHeader>
        On Error GoTo WriteAskTrigger_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eAskTrigger)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteAskTrigger_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteAskTrigger", Erl)
        '</EhFooter>
End Sub



''
' Writes the "GuildBan" message to the outgoing data buffer.
'
' @param    guild The guild whose members will be banned.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteGuildBan(ByVal guild As String)
        '<EhHeader>
        On Error GoTo WriteGuildBan_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGuildBan)
102     Call Writer.WriteString8(guild)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGuildBan_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGuildBan", Erl)
        '</EhFooter>
End Sub



''
' Writes the "CreateItem" message to the outgoing data buffer.
'
' @param    itemIndex The index of the item to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCreateItem(ByVal ItemIndex As Long, ByVal cantidad As Integer)
        '<EhHeader>
        On Error GoTo WriteCreateItem_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eCreateItem)
102     Call Writer.WriteInt16(ItemIndex)
104     Call Writer.WriteInt16(cantidad)
    
106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCreateItem_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCreateItem", Erl)
        '</EhFooter>
End Sub

''
' Writes the "DestroyItems" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteDestroyItems()
        '<EhHeader>
        On Error GoTo WriteDestroyItems_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eDestroyItems)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteDestroyItems_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteDestroyItems", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ChaosLegionKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the Chaos Legion.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChaosLegionKick(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WriteChaosLegionKick_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eChaosLegionKick)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteChaosLegionKick_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteChaosLegionKick", Erl)
        '</EhFooter>
End Sub

''
' Writes the "RoyalArmyKick" message to the outgoing data buffer.
'
' @param    username The name of the user to be kicked from the Royal Army.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRoyalArmyKick(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WriteRoyalArmyKick_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eRoyalArmyKick)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteRoyalArmyKick_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteRoyalArmyKick", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ForceMIDIAll" message to the outgoing data buffer.
'
' @param    midiID The id of the midi file to play.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteForceMIDIAll(ByVal midiID As Byte)
        '<EhHeader>
        On Error GoTo WriteForceMIDIAll_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eForceMIDIAll)
102     Call Writer.WriteInt8(midiID)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteForceMIDIAll_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteForceMIDIAll", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ForceWAVEAll" message to the outgoing data buffer.
'
' @param    waveID The id of the wave file to play.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteForceWAVEAll(ByVal waveID As Byte)
        '<EhHeader>
        On Error GoTo WriteForceWAVEAll_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eForceWAVEAll)
102     Call Writer.WriteInt8(waveID)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteForceWAVEAll_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteForceWAVEAll", Erl)
        '</EhFooter>
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
        '<EhHeader>
        On Error GoTo WriteRemovePunishment_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eRemovePunishment)
102     Call Writer.WriteString8(UserName)
104     Call Writer.WriteInt8(punishment)
106     Call Writer.WriteString8(NewText)
    
108     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteRemovePunishment_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteRemovePunishment", Erl)
        '</EhFooter>
End Sub

''
' Writes the "TileBlockedToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteTileBlockedToggle()
        '<EhHeader>
        On Error GoTo WriteTileBlockedToggle_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eTileBlockedToggle)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteTileBlockedToggle_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteTileBlockedToggle", Erl)
        '</EhFooter>
End Sub

''
' Writes the "KillNPCNoRespawn" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteKillNPCNoRespawn()
        '<EhHeader>
        On Error GoTo WriteKillNPCNoRespawn_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eKillNPCNoRespawn)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteKillNPCNoRespawn_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteKillNPCNoRespawn", Erl)
        '</EhFooter>
End Sub

''
' Writes the "KillAllNearbyNPCs" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteKillAllNearbyNPCs()
        '<EhHeader>
        On Error GoTo WriteKillAllNearbyNPCs_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eKillAllNearbyNPCs)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteKillAllNearbyNPCs_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteKillAllNearbyNPCs", Erl)
        '</EhFooter>
End Sub

''
' Writes the "LastIP" message to the outgoing data buffer.
'
' @param    username The user whose last IPs are requested.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteLastIP(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WriteLastIP_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eLastIP)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteLastIP_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteLastIP", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ChangeMOTD" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeMOTD()
        '<EhHeader>
        On Error GoTo WriteChangeMOTD_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eChangeMOTD)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteChangeMOTD_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteChangeMOTD", Erl)
        '</EhFooter>
End Sub

''
' Writes the "SetMOTD" message to the outgoing data buffer.
'
' @param    message The message to be set as the new MOTD.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSetMOTD(ByVal Message As String)
        '<EhHeader>
        On Error GoTo WriteSetMOTD_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eSetMOTD)
102     Call Writer.WriteString8(Message)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteSetMOTD_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteSetMOTD", Erl)
        '</EhFooter>
End Sub

''
' Writes the "SystemMessage" message to the outgoing data buffer.
'
' @param    message The message to be sent to all players.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSystemMessage(ByVal Message As String)
        '<EhHeader>
        On Error GoTo WriteSystemMessage_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eSystemMessage)
102     Call Writer.WriteString8(Message)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteSystemMessage_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteSystemMessage", Erl)
        '</EhFooter>
End Sub

''
' Writes the "CreateNPC" message to the outgoing data buffer.
'
' @param    npcIndex The index of the NPC to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCreateNPC(ByVal NpcIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteCreateNPC_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eCreateNPC)
102     Call Writer.WriteInt16(NpcIndex)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCreateNPC_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCreateNPC", Erl)
        '</EhFooter>
End Sub

''
' Writes the "CreateNPCWithRespawn" message to the outgoing data buffer.
'
' @param    npcIndex The index of the NPC to be created.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCreateNPCWithRespawn(ByVal NpcIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteCreateNPCWithRespawn_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eCreateNPCWithRespawn)
102     Call Writer.WriteInt16(NpcIndex)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCreateNPCWithRespawn_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCreateNPCWithRespawn", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ImperialArmour" message to the outgoing data buffer.
'
' @param    armourIndex The index of imperial armour to be altered.
' @param    objectIndex The index of the new object to be set as the imperial armour.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteImperialArmour(ByVal armourIndex As Byte, ByVal objectIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteImperialArmour_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eImperialArmour)
102     Call Writer.WriteInt8(armourIndex)
104     Call Writer.WriteInt16(objectIndex)
    
106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteImperialArmour_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteImperialArmour", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ChaosArmour" message to the outgoing data buffer.
'
' @param    armourIndex The index of chaos armour to be altered.
' @param    objectIndex The index of the new object to be set as the chaos armour.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChaosArmour(ByVal armourIndex As Byte, ByVal objectIndex As Integer)
        '<EhHeader>
        On Error GoTo WriteChaosArmour_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eChaosArmour)
102     Call Writer.WriteInt8(armourIndex)
104     Call Writer.WriteInt16(objectIndex)
    
106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteChaosArmour_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteChaosArmour", Erl)
        '</EhFooter>
End Sub

''
' Writes the "NavigateToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteNavigateToggle()
        '<EhHeader>
        On Error GoTo WriteNavigateToggle_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eNavigateToggle)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteNavigateToggle_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteNavigateToggle", Erl)
        '</EhFooter>
End Sub

' Writes the "ServerOpenToUsersToggle" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteServerOpenToUsersToggle()
        '<EhHeader>
        On Error GoTo WriteServerOpenToUsersToggle_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eServerOpenToUsersToggle)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteServerOpenToUsersToggle_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteServerOpenToUsersToggle", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Participar" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteParticipar(ByVal RoomId As Integer, ByVal Password As String)
        '<EhHeader>
        On Error GoTo WriteParticipar_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eParticipar)
102     Call Writer.WriteInt16(RoomId)
104     Call Writer.WriteString8(Password)
106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteParticipar_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteParticipar", Erl)
        '</EhFooter>
End Sub

''
' Writes the "TurnCriminal" message to the outgoing data buffer.
'
' @param    username The name of the user to turn into criminal.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteTurnCriminal(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WriteTurnCriminal_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eTurnCriminal)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteTurnCriminal_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteTurnCriminal", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ResetFactions" message to the outgoing data buffer.
'
' @param    username The name of the user who will be removed from any faction.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteResetFactions(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WriteResetFactions_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eResetFactions)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteResetFactions_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteResetFactions", Erl)
        '</EhFooter>
End Sub

''
' Writes the "RemoveCharFromGuild" message to the outgoing data buffer.
'
' @param    username The name of the user who will be removed from any guild.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteRemoveCharFromGuild(ByVal UserName As String)
        '<EhHeader>
        On Error GoTo WriteRemoveCharFromGuild_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eRemoveCharFromGuild)
102     Call Writer.WriteString8(UserName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteRemoveCharFromGuild_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteRemoveCharFromGuild", Erl)
        '</EhFooter>
End Sub

''
' Writes the "AlterName" message to the outgoing data buffer.
'
' @param    username The name of the user whose mail is requested.
' @param    newName The new user name.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteAlterName(ByVal UserName As String, ByVal newName As String)
        '<EhHeader>
        On Error GoTo WriteAlterName_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eAlterName)
102     Call Writer.WriteString8(UserName)
104     Call Writer.WriteString8(newName)
    
106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteAlterName_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteAlterName", Erl)
        '</EhFooter>
End Sub

''
' Writes the "DoBackup" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteDoBackup()
        '<EhHeader>
        On Error GoTo WriteDoBackup_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eDoBackUp)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteDoBackup_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteDoBackup", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ShowGuildMessages" message to the outgoing data buffer.
'
' @param    guild The guild to listen to.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowGuildMessages(ByVal guild As String)
        '<EhHeader>
        On Error GoTo WriteShowGuildMessages_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eShowGuildMessages)
102     Call Writer.WriteString8(guild)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteShowGuildMessages_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteShowGuildMessages", Erl)
        '</EhFooter>
End Sub


''
' Writes the "ChangeMapInfoPK" message to the outgoing data buffer.
'
' @param    isPK True if the map is PK, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeMapInfoPK(ByVal isPK As Boolean)
        '<EhHeader>
        On Error GoTo WriteChangeMapInfoPK_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eChangeMapInfoPK)
102     Call Writer.WriteBool(isPK)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteChangeMapInfoPK_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteChangeMapInfoPK", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ChangeMapInfoBackup" message to the outgoing data buffer.
'
' @param    backup True if the map is to be backuped, False otherwise.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeMapInfoBackup(ByVal backup As Boolean)
        '<EhHeader>
        On Error GoTo WriteChangeMapInfoBackup_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eChangeMapInfoBackup)
102     Call Writer.WriteBool(backup)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteChangeMapInfoBackup_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteChangeMapInfoBackup", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ChangeMapInfoRestricted" message to the outgoing data buffer.
'
' @param    restrict NEWBIES (only newbies), NO (everyone), ARMADA (just Armadas), CAOS (just caos) or FACCION (Armadas & caos only)
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeMapInfoRestricted(ByVal restrict As String)
        '<EhHeader>
        On Error GoTo WriteChangeMapInfoRestricted_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eChangeMapInfoRestricted)
102     Call Writer.WriteString8(restrict)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteChangeMapInfoRestricted_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteChangeMapInfoRestricted", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ChangeMapInfoNoMagic" message to the outgoing data buffer.
'
' @param    nomagic TRUE if no magic is to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeMapInfoNoMagic(ByVal nomagic As Boolean)
        '<EhHeader>
        On Error GoTo WriteChangeMapInfoNoMagic_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eChangeMapInfoNoMagic)
102     Call Writer.WriteBool(nomagic)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteChangeMapInfoNoMagic_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteChangeMapInfoNoMagic", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ChangeMapInfoNoInvi" message to the outgoing data buffer.
'
' @param    noinvi TRUE if invisibility is not to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeMapInfoNoInvi(ByVal noinvi As Boolean)
        '<EhHeader>
        On Error GoTo WriteChangeMapInfoNoInvi_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eChangeMapInfoNoInvi)
102     Call Writer.WriteBool(noinvi)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteChangeMapInfoNoInvi_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteChangeMapInfoNoInvi", Erl)
        '</EhFooter>
End Sub
                            
''
' Writes the "ChangeMapInfoNoResu" message to the outgoing data buffer.
'
' @param    noresu TRUE if resurection is not to be allowed in the map.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeMapInfoNoResu(ByVal noresu As Boolean)
        '<EhHeader>
        On Error GoTo WriteChangeMapInfoNoResu_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eChangeMapInfoNoResu)
102     Call Writer.WriteBool(noresu)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteChangeMapInfoNoResu_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteChangeMapInfoNoResu", Erl)
        '</EhFooter>
End Sub
                        
''
' Writes the "ChangeMapInfoLand" message to the outgoing data buffer.
'
' @param    land options: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeMapInfoLand(ByVal lAnd As String)
        '<EhHeader>
        On Error GoTo WriteChangeMapInfoLand_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eChangeMapInfoLand)
102     Call Writer.WriteString8(lAnd)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteChangeMapInfoLand_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteChangeMapInfoLand", Erl)
        '</EhFooter>
End Sub
                        
''
' Writes the "ChangeMapInfoZone" message to the outgoing data buffer.
'
' @param    zone options: "BOSQUE", "NIEVE", "DESIERTO", "CIUDAD", "CAMPO", "DUNGEON".
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChangeMapInfoZone(ByVal zone As String)
        '<EhHeader>
        On Error GoTo WriteChangeMapInfoZone_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eChangeMapInfoZone)
102     Call Writer.WriteString8(zone)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteChangeMapInfoZone_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteChangeMapInfoZone", Erl)
        '</EhFooter>
End Sub

Public Sub WriteChangeMapSetting(ByVal setting As Byte, ByVal value As Byte)
        '<EhHeader>
        On Error GoTo WriteChangeMapSetting_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eChangeMapSetting)
102     Call Writer.WriteInt8(setting)
104     Call Writer.WriteInt8(value)

106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteChangeMapSetting_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteChangeMapSetting", Erl)
        '</EhFooter>
End Sub

''
' Writes the "SaveChars" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteSaveChars()
        '<EhHeader>
        On Error GoTo WriteSaveChars_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eSaveChars)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteSaveChars_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteSaveChars", Erl)
        '</EhFooter>
End Sub

''
' Writes the "CleanSOS" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCleanSOS()
        '<EhHeader>
        On Error GoTo WriteCleanSOS_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eCleanSOS)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCleanSOS_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCleanSOS", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ShowServerForm" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteShowServerForm()
        '<EhHeader>
        On Error GoTo WriteShowServerForm_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eShowServerForm)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteShowServerForm_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteShowServerForm", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Night" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteNight()
        '<EhHeader>
        On Error GoTo WriteNight_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.enight)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteNight_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteNight", Erl)
        '</EhFooter>
End Sub

Public Sub WriteDay()
        '<EhHeader>
        On Error GoTo WriteDay_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eDay)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteDay_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteDay", Erl)
        '</EhFooter>
End Sub

Public Sub WriteSetTime(ByVal Time As Long)
        '<EhHeader>
        On Error GoTo WriteSetTime_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eSetTime)
102     Call Writer.WriteInt32(Time)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteSetTime_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteSetTime", Erl)
        '</EhFooter>
End Sub

''
' Writes the "KickAllChars" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteKickAllChars()
        '<EhHeader>
        On Error GoTo WriteKickAllChars_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eKickAllChars)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteKickAllChars_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteKickAllChars", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ReloadNPCs" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteReloadNPCs()
        '<EhHeader>
        On Error GoTo WriteReloadNPCs_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eReloadNPCs)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteReloadNPCs_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteReloadNPCs", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ReloadServerIni" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteReloadServerIni()
        '<EhHeader>
        On Error GoTo WriteReloadServerIni_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eReloadServerIni)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteReloadServerIni_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteReloadServerIni", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ReloadSpells" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteReloadSpells()
        '<EhHeader>
        On Error GoTo WriteReloadSpells_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eReloadSpells)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteReloadSpells_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteReloadSpells", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ReloadObjects" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteReloadObjects()
        '<EhHeader>
        On Error GoTo WriteReloadObjects_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eReloadObjects)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteReloadObjects_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteReloadObjects", Erl)
        '</EhFooter>
End Sub

''
' Writes the "ChatColor" message to the outgoing data buffer.
'
' @param    r The red component of the new chat color.
' @param    g The green component of the new chat color.
' @param    b The blue component of the new chat color.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteChatColor(ByVal r As Byte, ByVal G As Byte, ByVal B As Byte)
        '<EhHeader>
        On Error GoTo WriteChatColor_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eChatColor)
102     Call Writer.WriteInt8(r)
104     Call Writer.WriteInt8(G)
106     Call Writer.WriteInt8(B)
    
108     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteChatColor_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteChatColor", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Ignored" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteIgnored()
        '<EhHeader>
        On Error GoTo WriteIgnored_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eIgnored)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteIgnored_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteIgnored", Erl)
        '</EhFooter>
End Sub

''
' Writes the "CheckSlot" message to the outgoing data buffer.
'
' @param    UserName    The name of the char whose slot will be checked.
' @param    slot        The slot to be checked.
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteCheckSlot(ByVal UserName As String, ByVal Slot As Byte)
        '<EhHeader>
        On Error GoTo WriteCheckSlot_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eCheckSlot)
102     Call Writer.WriteString8(UserName)
104     Call Writer.WriteInt8(Slot)
    
106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCheckSlot_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCheckSlot", Erl)
        '</EhFooter>
End Sub


Public Sub WriteLlamadadeClan()
        '<EhHeader>
        On Error GoTo WriteLlamadadeClan_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.ellamadadeclan)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteLlamadadeClan_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteLlamadadeClan", Erl)
        '</EhFooter>
End Sub

Public Sub WriteQuestionGM(ByVal Consulta As String, ByVal TipoDeConsulta As String)
        '<EhHeader>
        On Error GoTo WriteQuestionGM_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eQuestionGM)
102     Call Writer.WriteString8(Consulta)
104     Call Writer.WriteString8(TipoDeConsulta)
        packetCounters.TS_QuestionGM = packetCounters.TS_QuestionGM + 1
        Call Writer.WriteInt32(packetCounters.TS_QuestionGM)
    
106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteQuestionGM_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteQuestionGM", Erl)
        '</EhFooter>
End Sub

Public Sub WriteOfertaInicial(ByVal Oferta As Long)
        '<EhHeader>
        On Error GoTo WriteOfertaInicial_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eOfertaInicial)
102     Call Writer.WriteInt32(Oferta)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteOfertaInicial_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteOfertaInicial", Erl)
        '</EhFooter>
End Sub

Public Sub WriteOferta(ByVal OfertaDeSubasta As Long)
        '<EhHeader>
        On Error GoTo WriteOferta_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eOfertaDeSubasta)
102     Call Writer.WriteInt32(OfertaDeSubasta)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteOferta_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteOferta", Erl)
        '</EhFooter>
End Sub

Public Sub WriteSetSpeed(ByVal speed As Single)
        '<EhHeader>
        On Error GoTo WriteSetSpeed_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eSetSpeed)
102     Call Writer.WriteReal32(speed)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteSetSpeed_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteSetSpeed", Erl)
        '</EhFooter>
End Sub

Public Sub WriteGlobalMessage(ByVal Message As String)
        '<EhHeader>
        On Error GoTo WriteGlobalMessage_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGlobalMessage)
102     Call Writer.WriteString8(Message)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGlobalMessage_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGlobalMessage", Erl)
        '</EhFooter>
End Sub

Public Sub WriteGlobalOnOff()
        '<EhHeader>
        On Error GoTo WriteGlobalOnOff_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGlobalOnOff)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGlobalOnOff_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGlobalOnOff", Erl)
        '</EhFooter>
End Sub

Public Sub WriteNieblaToggle()
        '<EhHeader>
        On Error GoTo WriteNieblaToggle_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eNieblaToggle)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteNieblaToggle_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteNieblaToggle", Erl)
        '</EhFooter>
End Sub

Public Sub WriteGenio()
        '<EhHeader>
        On Error GoTo WriteGenio_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGenio)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGenio_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGenio", Erl)
        '</EhFooter>
End Sub

Public Sub WriteQuest()
        '<EhHeader>
        On Error GoTo WriteQuest_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eQuest)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteQuest_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteQuest", Erl)
        '</EhFooter>
End Sub
 
Public Sub WriteQuestDetailsRequest(ByVal QuestSlot As Byte)
        '<EhHeader>
        On Error GoTo WriteQuestDetailsRequest_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eQuestDetailsRequest)
102     Call Writer.WriteInt8(QuestSlot)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteQuestDetailsRequest_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteQuestDetailsRequest", Erl)
        '</EhFooter>
End Sub
 
Public Sub WriteQuestAccept(ByVal ListInd As Byte)
        '<EhHeader>
        On Error GoTo WriteQuestAccept_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eQuestAccept)
102     Call Writer.WriteInt8(ListInd)
        
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteQuestAccept_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteQuestAccept", Erl)
        '</EhFooter>
End Sub

Public Sub WriteQuestListRequest()
        '<EhHeader>
        On Error GoTo WriteQuestListRequest_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eQuestListRequest)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteQuestListRequest_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteQuestListRequest", Erl)
        '</EhFooter>
End Sub
 
Public Sub WriteQuestAbandon(ByVal QuestSlot As Byte)
        '<EhHeader>
        On Error GoTo WriteQuestAbandon_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eQuestAbandon)
        'Escribe el Slot de Quest.
102     Call Writer.WriteInt8(QuestSlot)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteQuestAbandon_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteQuestAbandon", Erl)
        '</EhFooter>
End Sub

Public Sub WriteResponderPregunta(ByVal Respuesta As Boolean)
        '<EhHeader>
        On Error GoTo WriteResponderPregunta_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eResponderPregunta)
102     Call Writer.WriteBool(Respuesta)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteResponderPregunta_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteResponderPregunta", Erl)
        '</EhFooter>
End Sub

Public Sub WriteCompletarViaje(ByVal destino As Byte, ByVal costo As Long)
        '<EhHeader>
        On Error GoTo WriteCompletarViaje_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eCompletarViaje)
102     Call Writer.WriteInt8(destino)
104     Call Writer.WriteInt32(costo)
    
106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCompletarViaje_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCompletarViaje", Erl)
        '</EhFooter>
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
        '<EhHeader>
        On Error GoTo WriteCreaerTorneo_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eCrearTorneo)
102     Call Writer.WriteInt8(nivelminimo)
104     Call Writer.WriteInt8(nivelmaximo)
106     Call Writer.WriteInt8(cupos)
108     Call Writer.WriteInt32(costo)
110     Call Writer.WriteInt8(mago)
112     Call Writer.WriteInt8(clerico)
114     Call Writer.WriteInt8(guerrero)
116     Call Writer.WriteInt8(asesino)
118     Call Writer.WriteInt8(bardo)
120     Call Writer.WriteInt8(druido)
122     Call Writer.WriteInt8(paladin)
124     Call Writer.WriteInt8(cazador)
126     Call Writer.WriteInt8(Trabajador)
128     Call Writer.WriteInt8(Pirata)
130     Call Writer.WriteInt8(Ladron)
132     Call Writer.WriteInt8(Bandido)
134     Call Writer.WriteInt16(map)
136     Call Writer.WriteInt8(x)
138     Call Writer.WriteInt8(y)
140     Call Writer.WriteString8(Name)
142     Call Writer.WriteString8(reglas)
    
144     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCreaerTorneo_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCreaerTorneo", Erl)
        '</EhFooter>
End Sub

Public Sub WriteComenzarTorneo()
        '<EhHeader>
        On Error GoTo WriteComenzarTorneo_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eComenzarTorneo)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteComenzarTorneo_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteComenzarTorneo", Erl)
        '</EhFooter>
End Sub

Public Sub WriteCancelarTorneo()
        '<EhHeader>
        On Error GoTo WriteCancelarTorneo_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eCancelarTorneo)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCancelarTorneo_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCancelarTorneo", Erl)
        '</EhFooter>
End Sub

Public Sub WriteBusquedaTesoro(ByVal TIPO As Byte)
        '<EhHeader>
        On Error GoTo WriteBusquedaTesoro_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eBusquedaTesoro)
102     Call Writer.WriteInt8(TIPO)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteBusquedaTesoro_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteBusquedaTesoro", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Home" message to the outgoing data buffer.
'
Public Sub WriteHome()
        '<EhHeader>
        On Error GoTo WriteHome_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eHome)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteHome_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteHome", Erl)
        '</EhFooter>
End Sub

''
' Writes the "Consulta" message to the outgoing data buffer.
'
' @remarks  The data is not actually sent until the buffer is properly flushed.
Public Sub WriteConsulta(Optional ByVal Nick As String = vbNullString)
        '<EhHeader>
        On Error GoTo WriteConsulta_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eConsulta)
102     Call Writer.WriteString8(Nick)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteConsulta_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteConsulta", Erl)
        '</EhFooter>
End Sub

Public Sub WriteCuentaExtractItem(ByVal Slot As Byte, _
                                  ByVal Amount As Integer, _
                                  ByVal slotdestino As Byte)
        '<EhHeader>
        On Error GoTo WriteCuentaExtractItem_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eCuentaExtractItem)
102     Call Writer.WriteInt8(Slot)
104     Call Writer.WriteInt16(Amount)
106     Call Writer.WriteInt8(slotdestino)
    
108     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCuentaExtractItem_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCuentaExtractItem", Erl)
        '</EhFooter>
End Sub

Public Sub WriteCuentaDeposit(ByVal Slot As Byte, _
                              ByVal Amount As Integer, _
                              ByVal slotdestino As Byte)
        '<EhHeader>
        On Error GoTo WriteCuentaDeposit_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eCuentaDeposit)
102     Call Writer.WriteInt8(Slot)
104     Call Writer.WriteInt16(Amount)
106     Call Writer.WriteInt8(slotdestino)
    
108     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCuentaDeposit_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCuentaDeposit", Erl)
        '</EhFooter>
End Sub

Public Sub WriteDuel(Players As String, _
                     ByVal Apuesta As Long, _
                     Optional ByVal PocionesRojas As Long = -1, _
                     Optional ByVal CaenItems As Boolean = False)
        '<EhHeader>
        On Error GoTo WriteDuel_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eDuel)
102     Call Writer.WriteString8(Players)
104     Call Writer.WriteInt32(Apuesta)
106     Call Writer.WriteInt16(PocionesRojas)
108     Call Writer.WriteBool(CaenItems)
    
110     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteDuel_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteDuel", Erl)
        '</EhFooter>
End Sub

Public Sub WriteAcceptDuel(Offerer As String)
        '<EhHeader>
        On Error GoTo WriteAcceptDuel_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eAcceptDuel)
102     Call Writer.WriteString8(Offerer)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteAcceptDuel_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteAcceptDuel", Erl)
        '</EhFooter>
End Sub

Public Sub WriteCancelDuel()
        '<EhHeader>
        On Error GoTo WriteCancelDuel_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eCancelDuel)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCancelDuel_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCancelDuel", Erl)
        '</EhFooter>
End Sub

Public Sub WriteQuitDuel()
        '<EhHeader>
        On Error GoTo WriteQuitDuel_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eQuitDuel)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteQuitDuel_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteQuitDuel", Erl)
        '</EhFooter>
End Sub

Public Sub WriteCreateEvent(EventName As String)
        '<EhHeader>
        On Error GoTo WriteCreateEvent_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eCreateEvent)
102     Call Writer.WriteString8(EventName)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCreateEvent_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCreateEvent", Erl)
        '</EhFooter>
End Sub

Public Sub WriteCommerceSendChatMessage(ByVal Message As String)
        '<EhHeader>
        On Error GoTo WriteCommerceSendChatMessage_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eCommerceSendChatMessage)
102     Call Writer.WriteString8(Message)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCommerceSendChatMessage_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCommerceSendChatMessage", Erl)
        '</EhFooter>
End Sub

Public Sub WriteLogMacroClickHechizo(ByVal tipo As Byte, Optional ByVal clicks As Long = 1)
        '<EhHeader>
        On Error GoTo WriteLogMacroClickHechizo_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eLogMacroClickHechizo)
101     Call Writer.WriteInt8(tipo)
103     Call Writer.WriteInt32(clicks)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteLogMacroClickHechizo_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteLogMacroClickHechizo", Erl)
        '</EhFooter>
End Sub

Public Sub WriteNieveToggle()
        '<EhHeader>
        On Error GoTo WriteNieveToggle_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eNieveToggle)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteNieveToggle_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteNieveToggle", Erl)
        '</EhFooter>
End Sub

Public Sub WriteCompletarAccion(ByVal Accion As Byte)
        '<EhHeader>
        On Error GoTo WriteCompletarAccion_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eCompletarAccion)
102     Call Writer.WriteInt8(Accion)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCompletarAccion_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCompletarAccion", Erl)
        '</EhFooter>
End Sub

Public Sub WriteGetMapInfo()
        '<EhHeader>
        On Error GoTo WriteGetMapInfo_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eGetMapInfo)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteGetMapInfo_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteGetMapInfo", Erl)
        '</EhFooter>
End Sub

Public Sub WriteAddItemCrafting(ByVal SlotInv As Byte, ByVal SlotCraft As Byte)
        '<EhHeader>
        On Error GoTo WriteAddItemCrafting_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eAddItemCrafting)
102     Call Writer.WriteInt8(SlotInv)
104     Call Writer.WriteInt8(SlotCraft)
    
106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteAddItemCrafting_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteAddItemCrafting", Erl)
        '</EhFooter>
End Sub
    
Public Sub WriteRemoveItemCrafting(ByVal SlotCraft As Byte, ByVal SlotInv As Byte)
        '<EhHeader>
        On Error GoTo WriteRemoveItemCrafting_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eRemoveItemCrafting)
102     Call Writer.WriteInt8(SlotCraft)
104     Call Writer.WriteInt8(SlotInv)
    
106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteRemoveItemCrafting_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteRemoveItemCrafting", Erl)
        '</EhFooter>
End Sub

Public Sub WriteAddCatalyst(ByVal SlotInv As Byte)
        '<EhHeader>
        On Error GoTo WriteAddCatalyst_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eAddCatalyst)
102     Call Writer.WriteInt8(SlotInv)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteAddCatalyst_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteAddCatalyst", Erl)
        '</EhFooter>
End Sub

Public Sub WriteRemoveCatalyst(ByVal SlotInv As Byte)
        '<EhHeader>
        On Error GoTo WriteRemoveCatalyst_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eRemoveCatalyst)
102     Call Writer.WriteInt8(SlotInv)
    
104     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteRemoveCatalyst_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteRemoveCatalyst", Erl)
        '</EhFooter>
End Sub

Public Sub WriteCraftItem()
        '<EhHeader>
        On Error GoTo WriteCraftItem_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eCraftItem)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCraftItem_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCraftItem", Erl)
        '</EhFooter>
End Sub

Public Sub WriteMoveCraftItem(ByVal Drag As Byte, ByVal Drop As Byte)
        '<EhHeader>
        On Error GoTo WriteMoveCraftItem_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eMoveCraftItem)
102     Call Writer.WriteInt8(Drag)
104     Call Writer.WriteInt8(Drop)
    
106     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteMoveCraftItem_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteMoveCraftItem", Erl)
        '</EhFooter>
End Sub

Public Sub WriteCloseCrafting()
        '<EhHeader>
        On Error GoTo WriteCloseCrafting_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eCloseCrafting)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteCloseCrafting_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteCloseCrafting", Erl)
        '</EhFooter>
End Sub

Public Sub WritePetLeaveAll()
        '<EhHeader>
        On Error GoTo WritePetLeaveAll_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.ePetLeaveAll)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WritePetLeaveAll_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WritePetLeaveAll", Erl)
        '</EhFooter>
End Sub



Public Sub WriteResetChar(ByVal Nick As String)
    On Error GoTo WriteResetChar_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eResetChar)
        Call Writer.WriteString8(Nick)
    
102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteResetChar_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteResetChar", Erl)
End Sub

Public Sub WriteResetearPersonaje()
         On Error GoTo WriteResetearPersonaje_Err

100     Call Writer.WriteInt16(ClientPacketID.eResetearPersonaje)

102     Call modNetwork.Send(Writer)
        Exit Sub

WriteResetearPersonaje_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteResetearPersonaje", Erl)
End Sub

Public Sub WriteDeleteItem(ByVal Slot As Byte)
     On Error GoTo WriteDeleteItem_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eDeleteItem)
        Call Writer.WriteInt8(Slot)

102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteDeleteItem_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteDeleteItem", Erl)
End Sub


Public Sub WriteFinalizarPescaEspecial()
     On Error GoTo WriteFinalizarPescaEspecial_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eFinalizarPescaEspecial)

102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteFinalizarPescaEspecial_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteFinalizarPescaEspecial", Erl)
End Sub


Public Sub WriteRomperCania()
     On Error GoTo WriteRomperCania_Err
        '</EhHeader>
100     Call Writer.WriteInt16(ClientPacketID.eRomperCania)

102     Call modNetwork.Send(Writer)
        '<EhFooter>
        Exit Sub

WriteRomperCania_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteRomperCania", Erl)
End Sub


Public Sub writePublicarPersonajeMAO(ByVal valor As Long)
     On Error GoTo writePublicarPersonajeMAO_Err
        
100     Call Writer.WriteInt16(ClientPacketID.ePublicarPersonajeMAO)
        Call Writer.WriteInt32(valor)
102     Call modNetwork.Send(Writer)
        
        Exit Sub

writePublicarPersonajeMAO_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.writePublicarPersonajeMAO", Erl)
End Sub

Public Sub WriteRequestDebug(ByVal debugType As Byte, ByRef arguments() As String, ByVal argCount As Integer)
    On Error GoTo WriteRequestDebug_Err
        
100     Call Writer.WriteInt16(ClientPacketID.eRequestDebug)
        Call Writer.WriteInt8(debugType)
        If debugType = e_DebugCommands.eConnectionState Then
            Writer.WriteString8 (arguments(0))
        End If
        
102     Call modNetwork.Send(Writer)
        
        Exit Sub

WriteRequestDebug_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.writePublicarPersonajeMAO", Erl)
End Sub

Public Sub WriteLobbyCommand(ByVal command As Byte, Optional ByVal Params As String = "")
    On Error GoTo WriteLobbyCommand_Err
        
100     Call Writer.WriteInt16(ClientPacketID.eLobbyCommand)
        Call Writer.WriteInt8(command)
        Call Writer.WriteString8(Params)
102     Call modNetwork.Send(Writer)
        Exit Sub

WriteLobbyCommand_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteLobbyCommand", Erl)
End Sub

Public Sub WriteFeatureEnable(ByVal name As String, ByVal Value As Byte)
    On Error GoTo WriteFeatureEnable_Err
        
100     Call Writer.WriteInt16(ClientPacketID.eFeatureToggle)
        Call Writer.WriteInt8(value)
        Call Writer.WriteString8(name)
102     Call modNetwork.Send(Writer)
        Exit Sub

WriteFeatureEnable_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteFeatureEnable", Erl)
End Sub

Public Sub WriteActionOnGroupFrame(ByVal GruopIndex As Byte)
    On Error GoTo WriteFeatureEnable_Err
        
100     Call Writer.WriteInt16(ClientPacketID.eActionOnGroupFrame)
        Call Writer.WriteInt8(GruopIndex)
102     Call modNetwork.Send(Writer)
        Exit Sub

WriteFeatureEnable_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteFeatureEnable", Erl)
End Sub

Public Sub WriteSendTelemetry(ByRef TelemetryData() As Byte, ByVal DataSize As Long)
    On Error GoTo WriteFeatureEnable_Err
        
100     Call Writer.WriteInt16(ClientPacketID.eSendTelemetry)
        Call Writer.WriteInt32(DataSize)
        Dim i As Long
        For i = 0 To DataSize - 1
            Writer.WriteInt8 (TelemetryData(i))
        Next i
        Call modNetwork.Send(Writer)
        Exit Sub

WriteFeatureEnable_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteFeatureEnable", Erl)
End Sub

Public Sub WriteSetHotkeySlot(ByVal SlotIndex As Byte, ByVal Index As Integer, ByVal LastKnownSlot As Integer, ByVal HotkeyType As e_HotkeyType)
On Error GoTo WriteSetHotkeySlot_Err
        
100     Call Writer.WriteInt16(ClientPacketID.eSetHotkeySlot)
        Call Writer.WriteInt8(SlotIndex)
        Call Writer.WriteInt16(Index)
        Call Writer.WriteInt16(LastKnownSlot)
        Call Writer.WriteInt8(HotkeyType)
        Call modNetwork.Send(Writer)
        Exit Sub

WriteSetHotkeySlot_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteSetHotkeySlot", Erl)
End Sub

Public Sub WriteUseHKeySlot(ByVal SlotIndex As Byte)
On Error GoTo WriteUseHKeySlot_Err
        
100     Call Writer.WriteInt16(ClientPacketID.eUseHKeySlot)
        Call Writer.WriteInt8(SlotIndex)
        Call modNetwork.Send(Writer)
        Exit Sub

WriteUseHKeySlot_Err:
        Call Writer.Clear
        Call RegistrarError(Err.Number, Err.Description, "Argentum20.Protocol_Writes.WriteUseHKeySlot", Erl)
End Sub
