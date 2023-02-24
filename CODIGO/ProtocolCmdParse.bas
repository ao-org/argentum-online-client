Attribute VB_Name = "ProtocolCmdParse"
'    Argentum 20 - Game Client Program
'    Copyright (C) 2022 - Noland Studios
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
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'
Option Explicit

Public Enum eNumber_Types

    ent_Byte
    ent_Integer
    ent_Long
    ent_Trigger

End Enum

Public Enum e_LobbyCommandId
    eSetSpawnPos
    eSetMaxLevel
    eSetMinLevel
    eSetClassLimit
    eRegisterPlayer
    eSummonSinglePlayer
    eSummonAll
    eReturnSinglePlayer
    eReturnAllSummoned
    eOpenLobby
    eStartEvent
    eEndEvent
    eCancelEvent
    eListPlayers
    eKickPlayer
    eForceReset
    eSetTeamSize
    eAddPlayer
End Enum

Public Enum e_DebugCommands
    eGetLastLogs
    eConnectionState
End Enum

''
' Interpreta, valida y ejecuta el comando ingresado .
'
' @param    RawCommand El comando en version String
' @remarks  None Known.
Public Sub ParseUserCommand(ByVal RawCommand As String)
    
    On Error GoTo ParseUserCommand_Err
    

    '***************************************************
    'Author: Alejandro Santos (AlejoLp)
    'Last Modification: 12/20/06
    'Interpreta, valida y ejecuta el comando ingresado
    '***************************************************
    Dim TmpArgos()         As String
    
    Dim Comando            As String

    Dim ArgumentosAll()    As String

    Dim ArgumentosRaw      As String

    Dim Argumentos2()      As String

    Dim Argumentos3()      As String

    Dim Argumentos4()      As String

    Dim CantidadArgumentos As Long

    Dim notNullArguments   As Boolean
    
    Dim tmpArr()           As String

    Dim tmpInt             As Integer
    
    ' TmpArgs: Un array de a lo sumo dos elementos,
    ' el primero es el comando (hasta el primer espacio)
    ' y el segundo elemento es el resto. Si no hay argumentos
    ' devuelve un array de un solo elemento
    TmpArgos = Split(RawCommand, " ", 2)
    
    Comando = Trim$(UCase$(TmpArgos(0)))
    
    If UBound(TmpArgos) > 0 Then
        ' El string en crudo que este despues del primer espacio
        ArgumentosRaw = TmpArgos(1)
        
        'veo que los argumentos no sean nulos
        notNullArguments = LenB(Trim$(ArgumentosRaw))
        
        ' Un array separado por blancos, con tantos elementos como
        ' se pueda
        ArgumentosAll = Split(TmpArgos(1), " ")
        
        ' Cantidad de argumentos. En ESTE PUNTO el minimo es 1
        CantidadArgumentos = UBound(ArgumentosAll) + 1
        
        ' Los siguientes arrays tienen A LO SUMO, COMO MAXIMO
        ' 2, 3 y 4 elementos respectivamente. Eso significa
        ' que pueden tener menos, por lo que es imperativo
        ' preguntar por CantidadArgumentos.
        
        Argumentos2 = Split(TmpArgos(1), " ", 2)
        Argumentos3 = Split(TmpArgos(1), " ", 3)
        Argumentos4 = Split(TmpArgos(1), " ", 4)
    Else
        CantidadArgumentos = 0

    End If
    
    ' Sacar cartel APESTA!! (y es ilógico, estás diciendo una pausa/espacio  :rolleyes: )
    If Comando = "" Then Comando = " "
    
    If Left$(Comando, 1) = "/" Then
        ' Comando normal
        
        Select Case Comando

            Case "/SEG"
                Call WriteSafeToggle

            Case "/ONLINE"
                Call WriteOnline
                
            Case "/SALIR"
                Call WriteQuit
                
            Case "/SALIRCLAN"
                Call WriteGuildLeave
                
            Case "/BALANCE"
                If UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

                    End With

                    Exit Sub

                End If

                Call WriteRequestAccountState
                
            Case "/QUIETO"
                If UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                    End With

                    Exit Sub

                End If

                Call WritePetStand

            Case "/ACOMPAÑAR"
                If UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                    End With

                    Exit Sub

                End If

                Call WritePetFollow
                
            Case "/LIBERAR"
                If UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                    End With

                    Exit Sub

                End If

                Call WritePetLeave
            
            Case "/LIBERARTODOS", "/LIBERARTODAS"
                If UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                    End With

                    Exit Sub

                End If

                Call WritePetLeaveAll
                                
            Case "/ENTRENAR"
                If UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

                    End With

                    Exit Sub

                End If

                Call WriteTrainList
                
            Case "/DESCANSAR"
                If UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

                    End With

                    Exit Sub

                End If

                Call WriteRest
                
            Case "/MEDITAR"
                If UserMinMAN = UserMaxMAN Then
                    With FontTypes(FontTypeNames.FONTTYPE_INFOBOLD)
                        Call ShowConsoleMsg("¡Tu maná está completo!", .red, .green, .blue, .bold, .italic)
                    End With

                    Exit Sub
                End If
                
                If UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
                    End With

                    Exit Sub

                End If

                Call WriteMeditate
        
            Case "/RESUCITAR"
                Call WriteResucitate
                
            Case "/CURAR"
                Call WriteHeal
                              
            Case "/EST"
                Call WriteRequestStats
                
            Case "/PROMEDIO"
                Call WritePromedio
            
            Case "/AYUDA"
                Call WriteHelp
            
            Case "/EVENTOFACCIONARIO"
                Call WriteEventoFaccionario
            Case "/SUBASTA"
                Call WriteSubastaInfo
                
            Case "/EVENTO"
                Call WriteEventoInfo
                
            Case "/COMERCIAR"
                If UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

                    End With

                    Exit Sub
                
                ElseIf Comerciando Then 'Comerciando

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("Ya estás comerciando", .red, .green, .blue, .bold, .italic)

                    End With

                    Exit Sub

                End If

                Call WriteCommerceStart
                
            Case "/BOVEDA"
                If UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

                    End With

                    Exit Sub

                End If

                Call WriteBankStart
                
            Case "/ENLISTAR"
                Call WriteEnlist
                    
            Case "/INFORMACION"
                Call WriteInformation
                
            Case "/RECOMPENSA"
                Call WriteReward
                
            Case "/MOTD"
                Call WriteRequestMOTD
                
            Case "/UPTIME"
                Call WriteUpTime
        
            Case "/CMSG"
                'Ojo, no usar notNullArguments porque se usa el string Vacío para borrar cartel.
                If CantidadArgumentos > 0 Then
                    Call WriteGuildMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")

                End If
                
            Case "/GRUPO"
                'Ojo, no usar notNullArguments porque se usa el string Vacío para borrar cartel.
                If CantidadArgumentos > 0 Then
                    Call WriteGrupoMsg(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")

                End If
        
            Case "/ONLINECLAN"
                Call WriteGuildOnline
                
            Case "/BMSG"
                If notNullArguments Then
                    Call WriteCouncilMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")

                End If
                
            Case "/ROL"
                If notNullArguments Then
                    Call WriteRoleMasterRequest(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba una pregunta.")

                End If
                
            Case "/GM"
                FrmGmAyuda.Show vbModeless, frmMain
                 
            Case "/OFERTAINICIAL"
                If notNullArguments Then
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Long) Then
                        If ArgumentosRaw > 0 Then
                            Call WriteOfertaInicial(ArgumentosRaw)
                        Else
                            Call ShowConsoleMsg("Cantidad incorrecta. Utilice /OFERTAINICIAL CANTIDAD.")

                        End If

                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Cantidad incorrecta. Utilice /OFERTAINICIAL CANTIDAD.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /OFERTAINICIAL CANTIDAD.")

                End If
            
            Case "/OFERTAR"
                If notNullArguments Then
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Long) Then
                        If ArgumentosRaw > 0 Then
                            Call WriteOferta(ArgumentosRaw)
                        Else
                            Call ShowConsoleMsg("Ingrese una oferta correcta.")

                        End If

                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Ingrese una oferta correcta.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Ingrese una oferta correcta.")

                End If
                        
            Case "/DESC"
                If UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

                    End With

                    Exit Sub

                End If

                If Len(ArgumentosRaw) > 50 Then

                    With FontTypes(FontTypeNames.FONTTYPE_INFOIAO)
                        Call ShowConsoleMsg("Tu descripción no puede ser tan larga (Max. 50 caracteres).", .red, .green, .blue, .bold, .italic)

                    End With

                    Exit Sub

                End If
                
                Call WriteChangeDescription(ArgumentosRaw)
            
            Case "/VOTO"
                If notNullArguments Then
                    Call WriteGuildVote(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /voto NICKNAME.")

                End If
               
            Case "/PENAS"
                Call WritePunishments(ArgumentosRaw)

            
            Case "/APOSTAR"
                If UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

                    End With

                    Exit Sub

                End If

                If notNullArguments Then
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Integer) Then
                        Call WriteGamble(ArgumentosRaw)
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Cantidad incorrecta. Utilice /apostar CANTIDAD.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /apostar CANTIDAD.")

                End If
                
            Case "/RETIRAR"
                If UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

                    End With

                    Exit Sub

                End If

                If CantidadArgumentos = 0 Then
                    ' Version sin argumentos: LeaveFaction
                    Call ShowConsoleMsg("Ingrese la cantidad que desea retirar")
                Else

                    ' Version con argumentos: BankExtractGold
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Long) Then
                        Call WriteBankExtractGold(ArgumentosRaw)
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Cantidad incorrecta. Utilice /retirar CANTIDAD.")

                    End If

                End If
                
             Case "/RETIRARFACCION"
                If UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

                    End With

                    Exit Sub

                End If

                If CantidadArgumentos = 0 Then
                    ' Version sin argumentos: LeaveFaction
                    Call WriteLeaveFaction
                End If
    
            Case "/DEPOSITAR"
                If UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

                    End With

                    Exit Sub

                End If
                
                If notNullArguments Then
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Long) Then
                        Call WriteBankDepositGold(ArgumentosRaw)
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Cantidad incorecta. Utilice /depositar CANTIDAD.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /depositar CANTIDAD.")

                End If
                
            Case "/DENUNCIAR"
                If notNullArguments Then
                    PreguntaScreen = "¿Denunciar los mensajes de " & ArgumentosRaw & "? El uso indebido del comando es motivo de advertencia."
                    Pregunta = True
                    TargetName = ArgumentosRaw
    
                    PreguntaLocal = True
                    PreguntaNUM = 2
                Else
                    Call ShowConsoleMsg("Utilice /denunciar NICK, para reportar los mensajes de un usuario.")
                End If

            Case "/FINALIZAREVENTO"
                Call WriteFinEvento

            Case "/PROPONER"
                If notNullArguments Then
                    Call WriteCasamiento(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /PROPONER NICK")

                End If

                '
                ' BEGIN GM COMMANDS
                '
            
            Case "/GMSG"
                If notNullArguments Then
                    Call WriteGMMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")

                End If
                
            Case "/SHOWNAME"
                Call WriteShowName
                
            Case "/ONLINEREAL"
                Call WriteOnlineRoyalArmy
                
            Case "/ONLINECAOS"
                Call WriteOnlineChaosLegion
                
            Case "/IRCERCA"
                If notNullArguments Then
                    Call WriteGoNearby(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /ircerca NICKNAME.")

                End If
                
            Case "/REM", "/LOG", "/ANOTAR"
                If notNullArguments Then
                    Call WriteComment(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un comentario.")

                End If
            
            Case "/HORA"
                If notNullArguments And EsGM Then
                    Call WriteSetTime(GetTimeFromString(ArgumentosRaw))
                Else
                    Call WriteServerTime
                End If
            
            Case "/DONDE"
                If notNullArguments Then
                    Call WriteWhere(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /donde NICKNAME.")

                End If
                
            Case "/NENE"
                If notNullArguments Then
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Integer) Then
                        Call WriteCreaturesInMap(ArgumentosRaw)
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Mapa incorrecto. Utilice /nene MAPA.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /nene MAPA.")

                End If
                
            Case "/TELEPLOC"
                Call WriteWarpMeToTarget
                
            Case "/TELEP"
                If notNullArguments Then
                    If CantidadArgumentos >= 4 Then
                        If ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(3), eNumber_Types.ent_Byte) Then
                            Call WriteWarpChar(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2), ArgumentosAll(3))
                        Else
                            'No es numerico
                            Call ShowConsoleMsg("Valor incorrecto. Utilice /telep NICKNAME MAPA X Y.")

                        End If

                    End If
                    
                    If CantidadArgumentos = 2 Then
                        If ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) Then
                            Call WriteWarpChar(ArgumentosAll(0), ArgumentosAll(1), 50, 50)
                        Else
                            'No es numerico
                            Call ShowConsoleMsg("Valor incorrecto. Utilice /telep NICKNAME MAPA X Y.")

                        End If

                    End If
                    
                Else
                    
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /telep NICKNAME MAPA X Y.")

                End If
                
            
            Case "/CREAREVENTO"
                Call CreateEventCmd(ArgumentosAll, CantidadArgumentos)
            
            Case "/CONFIGLOBBY"
                Call ConfigLobby(ArgumentosAll, CantidadArgumentos)
            
            Case "/CANCELAREVENTO"
                Call WriteCancelarEvento
                
            
            Case "/SILENCIAR"
                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@", 2)

                    If UBound(tmpArr) = 1 Then
                        Call WriteSilence(tmpArr(0), tmpArr(1))
                    Else
                        'Faltan los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /SILENCIAR NICKNAME@MOTIVO.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /SILENCIAR NICKNAME@MOTIVO.")

                End If
                
            Case "/CR"
                If notNullArguments Then
                    Call WriteCuentaRegresiva(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /CR TIEMPO (En segundos).")

                End If
            
            Case "/LOG"
                If notNullArguments Then
                    Call WritePossUser(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /LOG <NICK>.")

                End If
                
            Case "/SHOW"
                If notNullArguments Then

                    Select Case UCase$(ArgumentosAll(0))

                        Case "INT"
                            Call WriteShowServerForm
                            
                    End Select

                End If
                
            Case "/IRA"
                If EsGM Then
                    Call WriteGoToChar(ArgumentosRaw)
                End If
                
            Case "/GO"
                If EsGM Then
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Integer) Then
                        Call WriteWarpChar("YO", ArgumentosRaw, 50, 50)
                    Else
                        Call WriteGoToChar(ArgumentosRaw)

                    End If
                End If
                
            Case "/LUZ"
                If EsGM Then
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Integer) Then
                        Call LucesRedondas.Create_Light_To_Map(UserPos.x, UserPos.y, COLOR_WHITE(0), Val(ArgumentosRaw))
                    Else
                        Call LucesRedondas.Create_Light_To_Map(UserPos.x, UserPos.y, COLOR_WHITE(0), 10)
                    End If
                End If
                
            'Case "/LUZMAPA"
            '    If EsGM Then
            '        If notNullArguments Then
            '            If CantidadArgumentos = 3 Then
            '                If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) And _
            '                    ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) And _
            '                    ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Integer) Then
                                
                                'Call SetGlobalLight(D3DColorXRGB(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2)))
                                'Call MapUpdateGlobalLight
            '                    Exit Sub

            '                End If
            '            End If
            '        End If

                    'Avisar que falta el parametro
            '        Call ShowConsoleMsg("Faltan parámetros. Utilice /luzmapa R G B.")
            '    End If
        
            Case "/INVISIBLE"
                Call WriteInvisible

            Case "/PANELGM"
                Call WriteSOSShowList
                Call WriteGMPanel
            
            Case "/GENIO"
                Call WriteGenio
                
            Case "/TRABAJANDO"
                Call WriteWorking
                
            Case "/OCULTANDO"
                Call WriteHiding
                
            Case "/CARCEL"
                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@")

                    If UBound(tmpArr) = 2 Then
                        If ValidNumber(tmpArr(2), eNumber_Types.ent_Byte) Then
                            Call WriteJail(tmpArr(0), tmpArr(1), tmpArr(2))
                        Else
                            'No es numerico
                            Call ShowConsoleMsg("Tiempo incorrecto. Utilice /carcel NICKNAME@MOTIVO@TIEMPO.")

                        End If

                    Else
                        'Faltan los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /carcel NICKNAME@MOTIVO@TIEMPO.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /carcel NICKNAME@MOTIVO@TIEMPO.")

                End If
                
            Case "/CREAREVENTO"

                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@")

                    If UBound(tmpArr) = 2 Then
                        If ValidNumber(tmpArr(2), eNumber_Types.ent_Byte) Then
                            Call WriteCrearEvento(tmpArr(0), tmpArr(1), tmpArr(2))
                        Else
                            'No es numerico
                            Call ShowConsoleMsg("Tiempo incorrecto. Utilice /CREAREVENTO TIPO@DURACION@MULTIPLICACION.")

                        End If

                    Else
                        'Faltan los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /CREAREVENTO TIPO@DURACION@MULTIPLICACION.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /CREAREVENTO TIPO@DURACION@MULTIPLICACION.")

                End If
                
            Case "/RMATA"
                Call WriteKillNPC
                
            Case "/ADVERTENCIA"

                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@", 2)

                    If UBound(tmpArr) = 1 Then
                        Call WriteWarnUser(tmpArr(0), tmpArr(1))
                    Else
                        'Faltan los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /advertencia NICKNAME@MOTIVO.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /advertencia NICKNAME@MOTIVO.")

                End If
                
            Case "/MENSAJEINFORMACION"

                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@", 2)

                    If UBound(tmpArr) = 1 Then
                        Call WriteMensajeUser(tmpArr(0), tmpArr(1))
                    Else
                        'Faltan los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /MENSAJEINFORMACION NICKNAME@MENSAJE.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /MENSAJEINFORMACION NICKNAME@MENSAJE.")

                End If
                
            Case "/MOD"

                If notNullArguments And CantidadArgumentos >= 3 Then

                    Select Case UCase$(ArgumentosAll(1))

                        Case "BODY", "CUERPO"
                            tmpInt = eEditOptions.eo_Body
                            
                        Case "ARMA"
                            tmpInt = eEditOptions.eo_Arma
                            
                        Case "CASCO"
                            tmpInt = eEditOptions.eo_Casco
                            
                        Case "ESCUDO"
                            tmpInt = eEditOptions.eo_Escudo
                            
                        Case "PARTICULA"
                            tmpInt = eEditOptions.eo_Particula
                            
                        Case "HEAD", "CABEZA"
                            tmpInt = eEditOptions.eo_Head
                        
                        Case "ORO"
                            tmpInt = eEditOptions.eo_Gold
                        
                        Case "LEVEL", "LVL", "ELV"
                            tmpInt = eEditOptions.eo_Level
                        
                        Case "SKILLS"
                            tmpInt = eEditOptions.eo_Skills
                        
                        Case "SKILLSLIBRES", "LIBRES"
                            tmpInt = eEditOptions.eo_SkillPointsLeft
                        
                        Case "CLASE", "CLASS"
                            tmpInt = eEditOptions.eo_Class
                        
                        Case "EXP"
                            tmpInt = eEditOptions.eo_Experience
                        
                        Case "CRI", "CRIMINALES"
                            tmpInt = eEditOptions.eo_CriminalsKilled
                        
                        Case "CIU", "CIUDADANOS"
                            tmpInt = eEditOptions.eo_CiticensKilled
                        
                        Case "SEX", "SEXO", "GENERO", "GENDER"
                            tmpInt = eEditOptions.eo_Sex
                            
                        Case "RAZA", "RACE"
                            tmpInt = eEditOptions.eo_Raza
                            
                        Case "HP", "VIDA", "SALUD"
                            tmpInt = eEditOptions.eo_Vida
                            
                        Case "MP", "MANA"
                            tmpInt = eEditOptions.eo_Mana
                            
                        Case "STA", "STAMINA", "ENERGIA"
                            tmpInt = eEditOptions.eo_Energia
                            
                        Case "MINHP", "MINVIDA"
                            tmpInt = eEditOptions.eo_MinHP
                            
                        Case "MINMP", "MINMANA"
                            tmpInt = eEditOptions.eo_MinMP
                            
                        Case "HIT", "GOLPE"
                            tmpInt = eEditOptions.eo_Hit
                            
                        Case "MINHIT", "MINGOLPE"
                            tmpInt = eEditOptions.eo_MinHit
                            
                        Case "MAXHIT", "MAXGOLPE"
                            tmpInt = eEditOptions.eo_MaxHit
                            
                        Case "DESC", "DESCRIPCION"
                            tmpInt = eEditOptions.eo_Desc
                            
                        Case "INT", "INTERVALO"
                            tmpInt = eEditOptions.eo_Intervalo
                            
                        Case "HOGAR", "CIUDAD", "CASA"
                            tmpInt = eEditOptions.eo_Hogar

                        Case Else
                            tmpInt = -1

                    End Select
                    
                    If tmpInt > 0 Then
                        If CantidadArgumentos = 3 Then
                            Call WriteEditChar(ArgumentosAll(0), tmpInt, ArgumentosAll(2), "")
                        Else
                            Call WriteEditChar(ArgumentosAll(0), tmpInt, ArgumentosAll(2), ArgumentosAll(3))

                        End If

                    Else
                        'Avisar que no exite el comando
                        Call ShowConsoleMsg("Comando incorrecto.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros.")

                End If
                
            Case "/INFO"

                If notNullArguments Then
                    Call WriteRequestCharInfo(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /info NICKNAME.")

                End If
                
            Case "/STAT"

                If notNullArguments Then
                    Call WriteRequestCharStats(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /stat NICKNAME.")

                End If
                
            Case "/BAL"

                If notNullArguments Then
                    Call WriteRequestCharGold(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /bal NICKNAME.")

                End If
                
            Case "/INV"

                If notNullArguments Then
                    Call WriteRequestCharInventory(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /inv NICKNAME.")

                End If
                
            Case "/BOV"

                If notNullArguments Then
                    Call WriteRequestCharBank(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /bov NICKNAME.")

                End If
                
            Case "/SKILLS"

                If notNullArguments Then
                    Call WriteRequestCharSkills(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /skills NICKNAME.")

                End If
                
            Case "/REVIVIR"

                If notNullArguments Then
                    Call WriteReviveChar(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /revivir NICKNAME.")

                End If
                
            Case "/SM"
                Call WriteSeguirMouse(ArgumentosRaw)
                
            Case "/PERDONFACCION"

                If notNullArguments Then
                    Call WritePerdonFaccion(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /PERDONFACCION NICKNAME.")
                End If
                
            Case "/ONLINEGM"
                Call WriteOnlineGM
                
            Case "/ONLINEMAP"
                Call WriteOnlineMap
                
            Case "/PERDON"
                Call WriteForgive
            
            Case "/DONAR"
                If notNullArguments Then
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Long) Then
                        Call WriteDonateGold(ArgumentosRaw)
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Cantidad incorecta. Utilice /donar CANTIDAD.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /donar CANTIDAD.")

                End If
                
            Case "/ECHAR"

                If notNullArguments Then
                    Call WriteKick(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /echar NICKNAME.")

                End If
                
            Case "/EJECUTAR"

                If notNullArguments Then
                    Call WriteExecute(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /ejecutar NICKNAME.")

                End If
                
            Case "/BAN"

                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@", 2)

                    If UBound(tmpArr) = 1 Then
                        Call WriteBanChar(tmpArr(0), tmpArr(1))
                    Else
                        'Faltan los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /ban NICKNAME@MOTIVO.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /ban NICKNAME@MOTIVO.")

                End If
                
            Case "/BANCUENTA"

                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@", 2)

                    If UBound(tmpArr) = 1 Then
                        Call WriteBanCuenta(tmpArr(0), tmpArr(1))
                    Else
                        'Faltan los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /BANCUENTA NICKNAME@MOTIVO.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /BANCUENTA NICKNAME@MOTIVO.")

                End If
                
            Case "/UNBANCUENTA"
                If notNullArguments Then

                    Call WriteUnBanCuenta(ArgumentosRaw)
    
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /UNBANCUENTA NICKNAME o EMAIL.")

                End If
                
            Case "/UNBAN"

                If notNullArguments Then
                    Call WriteUnbanChar(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /unban NICKNAME.")

                End If
                
            Case "/SEGUIR"
                Call WriteNPCFollow
                
            Case "/SUM"

                If EsGM Then
                    'If notNullArguments Then
                    Call WriteSummonChar(ArgumentosRaw)

                    'Else
                    'Avisar que falta el parametro
                    'Call ShowConsoleMsg("Faltan parámetros. Utilice /sum NICKNAME.")
                    ' End If
                End If
                
            Case "/CC"
                If EsGM Then
                    Call WriteSpawnListRequest
                End If
                
            Case "/CO"

                'Call WriteSpawnListRequest
                If EsGM Then
                
                    Dim i As Long

                    For i = 1 To NumOBJs

                        If ObjData(i).Name <> "" Then

                            Dim subelemento As ListItem

                            Set subelemento = FrmObjetos.ListView1.ListItems.Add(, , ObjData(i).Name)
                            
                            subelemento.SubItems(1) = i

                        End If

                    Next i

                    FrmObjetos.Show , frmMain

                End If
                
            Case "/RESETINV"
                Call WriteResetNPCInventory
                
            Case "/LIMPIAR"
                Call WriteCleanWorld
                
            Case "/RMSG"

                If notNullArguments Then
                    Call WriteServerMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")

                End If
                
            Case "/NICK2IP"

                If notNullArguments Then
                    Call WriteNickToIP(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /nick2ip NICKNAME.")

                End If
                
            Case "/IP2NICK"

                If notNullArguments Then
                    If validipv4str(ArgumentosRaw) Then
                        Call WriteIPToNick(str2ipv4l(ArgumentosRaw))
                    Else
                        'No es una IP
                        Call ShowConsoleMsg("IP incorrecta. Utilice /ip2nick IP.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /ip2nick IP.")

                End If
                
            Case "/ONCLAN"

                If notNullArguments Then
                    Call WriteGuildOnlineMembers(ArgumentosRaw)
                Else
                    'Avisar sintaxis incorrecta
                    Call ShowConsoleMsg("Utilice /onclan nombre del clan.")

                End If
                
            Case "/CT" ' 1 50 50 @motivo asd asd asd
                Dim tempStr() As String
                
                If InStr(1, ArgumentosRaw, "@") Then
                
                    tempStr = Split(ArgumentosRaw, "@")
                    
                    If notNullArguments And CantidadArgumentos > 4 And tempStr(1) <> vbNullString Then
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(3), eNumber_Types.ent_Byte) Then
                            Call WriteTeleportCreate(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2), ArgumentosAll(3), tempStr(1))
                        Else
                            'No es numerico
                            Call ShowConsoleMsg("Valor incorrecto. Utilice /ct MAPA X Y RADIO @MOTIVO. Si no requiere radio mande /CT MAPA X Y 0 @MOTIVO")
    
                        End If
    
                    Else
                        'Avisar que falta el parametro
                        Call ShowConsoleMsg("Faltan parámetros. Utilice /ct MAPA X Y @MOTIVO.")
    
                    End If
                Else
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /ct MAPA X Y @MOTIVO.")
                End If
                
            Case "/DT"
                Call WriteTeleportDestroy
                
            Case "/LLUVIA"
                Call WriteRainToggle
            
            Case "/NIEVE"
                Call WriteNieveToggle
            
            Case "/NIEBLA"
                Call WriteNieblaToggle
                
            Case "/SETDESC"
                Call WriteSetCharDescription(ArgumentosRaw)
            
            Case "/FORCEMIDIMAP"

                If notNullArguments Then

                    'elegir el mapa es opcional
                    If CantidadArgumentos = 1 Then
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) Then
                            'eviamos un mapa nulo para que tome el del usuario.
                            Call WriteForceMIDIToMap(ArgumentosAll(0), 0)
                        Else
                            'No es numerico
                            Call ShowConsoleMsg("Midi incorrecto. Utilice /forcemidimap MIDI MAPA, siendo el mapa opcional.")

                        End If

                    Else

                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) Then
                            Call WriteForceMIDIToMap(ArgumentosAll(0), ArgumentosAll(1))
                        Else
                            'No es numerico
                            Call ShowConsoleMsg("Valor incorrecto. Utilice /forcemidimap MIDI MAPA, siendo el mapa opcional.")

                        End If

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Utilice /forcemidimap MIDI MAPA, siendo el mapa opcional.")

                End If
                
            Case "/FORCEWAVMAP"

                If notNullArguments Then

                    'elegir la posicion es opcional
                    If CantidadArgumentos = 1 Then
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) Then
                            'eviamos una posicion nula para que tome la del usuario.
                            Call WriteForceWAVEToMap(ArgumentosAll(0), 0, 0, 0)
                        Else
                            'No es numerico
                            Call ShowConsoleMsg("Utilice /forcewavmap WAV MAP X Y, siendo los últimos 3 opcionales.")

                        End If

                    ElseIf CantidadArgumentos = 4 Then

                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(3), eNumber_Types.ent_Byte) Then
                            Call WriteForceWAVEToMap(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2), ArgumentosAll(3))
                        Else
                            'No es numerico
                            Call ShowConsoleMsg("Utilice /forcewavmap WAV MAP X Y, siendo los últimos 3 opcionales.")

                        End If

                    Else
                        'Avisar que falta el parametro
                        Call ShowConsoleMsg("Utilice /forcewavmap WAV MAP X Y, siendo los últimos 3 opcionales.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Utilice /forcewavmap WAV MAP X Y, siendo los últimos 3 opcionales.")

                End If
                
            Case "/REALMSG"

                If notNullArguments Then
                    Call WriteRoyalArmyMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")

                End If
                 
            Case "/CAOSMSG"

                If notNullArguments Then
                    Call WriteChaosLegionMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")

                End If
           
            Case "/TALKAS"

                If notNullArguments Then
                    Call WriteTalkAsNPC(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")

                End If
        
            Case "/MASSDEST"
                Call WriteDestroyAllItemsInArea
    
            Case "/ACEPTCONSE"

                If notNullArguments Then
                    Call WriteAcceptRoyalCouncilMember(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /aceptconse NICKNAME.")

                End If
                
            Case "/ACEPTCONSECAOS"

                If notNullArguments Then
                    Call WriteAcceptChaosCouncilMember(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /aceptconsecaos NICKNAME.")

                End If
                
            Case "/PISO"
                Call WriteItemsInTheFloor
                
            Case "/ESTUPIDO"

                If notNullArguments Then
                    Call WriteMakeDumb(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /estupido NICKNAME.")

                End If
                
            Case "/NOESTUPIDO"

                If notNullArguments Then
                    Call WriteMakeDumbNoMore(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /noestupido NICKNAME.")

                End If

            Case "/KICKCONSE"

                If notNullArguments Then
                    Call WriteCouncilKick(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /kickconse NICKNAME.")

                End If
                
            Case "/TRIGGER"

                If notNullArguments Then
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Trigger) Then
                        Call WriteSetTrigger(ArgumentosRaw)
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Numero incorrecto. Utilice /trigger NUMERO.")

                    End If

                Else
                    'Version sin parametro
                    Call WriteAskTrigger

                End If
                
            Case "/BANIPLIST"
                Call WriteBannedIPList
                
            Case "/BANIPRELOAD"
                Call WriteBannedIPReload
                
            Case "/MIEMBROSCLAN"

                If notNullArguments Then
                    Call WriteGuildMemberList(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /miembrosclan GUILDNAME.")

                End If
                
            Case "/BANCLAN"

                If notNullArguments Then
                    Call WriteGuildBan(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /banclan GUILDNAME.")

                End If
                
            Case "/BANIP"

                If CantidadArgumentos = 2 Then

                    Call WriteBanIP(ArgumentosAll(0), ArgumentosAll(1))

                Else
                
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /banip IP motivo o /banip nick motivo.")

                End If
                
            Case "/UNBANIP"

                If notNullArguments Then
                
                    If validipv4str(ArgumentosRaw) Then
                    
                        Call WriteUnbanIP(str2ipv4l(ArgumentosRaw))
                        
                    Else
                    
                        'No es una IP
                        Call ShowConsoleMsg("IP incorrecta. Utilice /unbanip IP.")

                    End If

                Else
                
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /unbanip IP.")

                End If
                
            Case "/CI"

                If EsGM Then
                    If notNullArguments Then
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) Then
                            If CantidadArgumentos = 1 Then
                                Call WriteCreateItem(ArgumentosAll(0), 1)
    
                            ElseIf CantidadArgumentos >= 2 Then

                                If ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) Then
                                    Call WriteCreateItem(ArgumentosAll(0), ArgumentosAll(1))
                                Else
                                    Call ShowConsoleMsg("Valor incorrecto. Utilice /CI OBJETO [CANTIDAD=1].")

                                End If

                            End If

                        Else
                            'No es numerico
                            Call ShowConsoleMsg("Valor incorrecto. Utilice /CI OBJETO [CANTIDAD=1].")

                        End If

                    Else
                        'Avisar que falta el parametro
                        Call ShowConsoleMsg("Faltan parámetros. Utilice /CI OBJETO [CANTIDAD=1].")

                    End If

                End If
                
            Case "/DAR"

                If EsGM Then
                    If notNullArguments Then
                        tmpArr = Split(ArgumentosRaw, "@", 4)
    
                        If UBound(tmpArr) < 2 Then
                            'Faltan los parametros con el formato propio
                             Call ShowConsoleMsg("Faltan parámetros. Utilice /DAR NOMBRE@MOTIVO@OBJETO[@CANTIDAD=1].")
    
                        Else
                            If Len(tmpArr(0)) = 0 Then
                                Call ShowConsoleMsg("Ingrese el nombre del usuario. Utilice /DAR NOMBRE@MOTIVO@OBJETO[@CANTIDAD=1].")
                            
                            ElseIf Len(tmpArr(1)) = 0 Then
                                Call ShowConsoleMsg("Ingrese el motivo para dar. Utilice /DAR NOMBRE@MOTIVO@OBJETO[@CANTIDAD=1].")
                            
                            ElseIf ValidNumber(tmpArr(2), ent_Integer) Then
                                Dim cantidad As String
                                If UBound(tmpArr) = 2 Then
                                    cantidad = 1
                                Else
                                    cantidad = tmpArr(3)
                                End If
                                
                                If ValidNumber(cantidad, ent_Integer) Then
                                    Call WriteGiveItem(tmpArr(0), tmpArr(2), cantidad, tmpArr(1))
                                Else
                                    Call ShowConsoleMsg("Cantidad inválida. Utilice /DAR NOMBRE@MOTIVO@OBJETO[@CANTIDAD=1].")
                                End If
                            Else
                                Call ShowConsoleMsg("Número de objeto inválido. Utilice /DAR NOMBRE@MOTIVO@OBJETO[@CANTIDAD=1].")
                            End If
                            
                        End If
    
                    Else
                        'Avisar que falta el parametro
                         Call ShowConsoleMsg("Faltan parámetros. Utilice /DAR NOMBRE@MOTIVO@OBJETO[@CANTIDAD=1].")
    
                    End If
                End If
                
            Case "/DEST"
                Call WriteDestroyItems
                
            Case "/NOCAOS"

                If notNullArguments Then
                    Call WriteChaosLegionKick(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /nocaos NICKNAME.")

                End If
    
            Case "/NOREAL"

                If notNullArguments Then
                    Call WriteRoyalArmyKick(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /noreal NICKNAME.")

                End If
    
            Case "/FORCEMIDI"

                If notNullArguments Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) Then
                        Call WriteForceMIDIAll(ArgumentosAll(0))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Midi incorrecto. Utilice /forcemidi MIDI.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /forcemidi MIDI.")

                End If
    
            Case "/FORCEWAV"

                If notNullArguments Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) Then
                        Call WriteForceWAVEAll(ArgumentosAll(0))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Wav incorrecto. Utilice /forcewav WAV.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /forcewav WAV.")

                End If
                
            Case "/BORRARPENA"

                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@", 3)

                    If UBound(tmpArr) = 2 Then
                        Call WriteRemovePunishment(tmpArr(0), tmpArr(1), tmpArr(2))
                    Else
                        'Faltan los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /borrarpena NICK@PENA@NuevaPena.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /borrarpena NICK@PENA@NuevaPena.")

                End If
                
            Case "/BLOQ"
                Call WriteTileBlockedToggle
                
            Case "/MATA"
                Call WriteKillNPCNoRespawn
        
            Case "/MASSKILL"
                Call WriteKillAllNearbyNPCs
                
            Case "/LASTIP"

                If notNullArguments Then
                    Call WriteLastIP(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /lastip NICKNAME.")

                End If
    
            Case "/MOTDCAMBIA"
                Call WriteChangeMOTD
                
            Case "/SMSG"

                If notNullArguments Then
                    Call WriteSystemMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")

                End If
                
            Case "/ACC"

                If notNullArguments Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) Then
                        Call WriteCreateNPC(ArgumentosAll(0))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Npc incorrecto. Utilice /acc NPC.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /acc NPC.")

                End If
                
            Case "/RACC"

                If notNullArguments Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) Then
                        Call WriteCreateNPCWithRespawn(ArgumentosAll(0))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Npc incorrecto. Utilice /racc NPC.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /racc NPC.")

                End If
        
            Case "/AI" ' ELIMINAR

                If notNullArguments And CantidadArgumentos >= 2 Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) Then
                        Call WriteImperialArmour(ArgumentosAll(0), ArgumentosAll(1))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Valor incorrecto. Utilice /ai ARMADURA OBJETO.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /ai ARMADURA OBJETO.")

                End If
                
            Case "/AC" ' ELIMINAR

                If notNullArguments And CantidadArgumentos >= 2 Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) Then
                        Call WriteChaosArmour(ArgumentosAll(0), ArgumentosAll(1))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Valor incorrecto. Utilice /ac ARMADURA OBJETO.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /ac ARMADURA OBJETO.")

                End If
                
            Case "/NAVE"
                Call WriteNavigateToggle
        
            Case "/HABILITAR"
                Call WriteServerOpenToUsersToggle
            
            Case "/PARTICIPAR"  '
                Call WriteParticipar
                
            Case "/CONDEN"

                If notNullArguments Then
                    Call WriteTurnCriminal(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /conden NICKNAME.")

                End If
                
            Case "/RAJAR"

                If notNullArguments Then
                    Call WriteResetFactions(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /rajar NICKNAME.")

                End If
                
            Case "/RAJARCLAN"

                If notNullArguments Then
                    Call WriteRemoveCharFromGuild(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /rajarclan NICKNAME.")

                End If

            Case "/ANAME"

                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@", 2)

                    If UBound(tmpArr) = 1 Then
                        Call WriteAlterName(tmpArr(0), tmpArr(1))
                    Else
                        'Faltan los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /aname ORIGEN@DESTINO.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /aname ORIGEN@DESTINO.")

                End If
                
            Case "/SLOT"

                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@", 2)

                    If UBound(tmpArr) = 1 Then
                        If ValidNumber(tmpArr(1), eNumber_Types.ent_Byte) Then
                            Call WriteCheckSlot(tmpArr(0), tmpArr(1))
                        Else
                            'Faltan o sobran los parametros con el formato propio
                            Call ShowConsoleMsg("Formato incorrecto. Utilice /slot NICK@SLOT.")

                        End If

                    Else
                        'Faltan o sobran los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /slot NICK@SLOT.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /slot NICK@SLOT.")

                End If

            Case "/DOBACKUP"
                Call WriteDoBackup
                
            Case "/SHOWCMSG"

                If notNullArguments Then
                    Call WriteShowGuildMessages(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /showcmsg GUILDNAME.")

                End If
                
            Case "/MODMAPINFO" ' PK, BACKUP

                If CantidadArgumentos > 1 Then

                    Select Case UCase$(ArgumentosAll(0))

                        Case "SEGURO" ' "/MODMAPINFO SEGURO"
                            Call WriteChangeMapInfoPK(ArgumentosAll(1) = "1")
                        
                        Case "BACKUP" ' "/MODMAPINFO BACKUP"
                            Call WriteChangeMapInfoBackup(ArgumentosAll(1) = "1")
                        
                        Case "RESTRINGIR" '/MODMAPINFO RESTRINGIR
                            Call WriteChangeMapInfoRestricted(ArgumentosAll(1))
                        
                        Case "MAGIASINEFECTO" '/MODMAPINFO MAGIASINEFECTO
                            Call WriteChangeMapInfoNoMagic(ArgumentosAll(1))
                        
                        Case "INVISINEFECTO" '/MODMAPINFO INVISINEFECTO
                            Call WriteChangeMapInfoNoInvi(ArgumentosAll(1))
                        
                        Case "RESUSINEFECTO" '/MODMAPINFO RESUSINEFECTO
                            Call WriteChangeMapInfoNoResu(ArgumentosAll(1))
                        
                        Case "TERRENO" '/MODMAPINFO TERRENO
                            Call WriteChangeMapInfoLand(ArgumentosAll(1))
                        
                        Case "ZONA" '/MODMAPINFO ZONA
                            Call WriteChangeMapInfoZone(ArgumentosAll(1))
                    End Select

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parametros. Opciones: SEGURO, BACKUP, RESTRINGIR, MAGIASINEFECTO, INVISINEFECTO, RESUSINEFECTO, TERRENO, ZONA")

                End If
            Case "/MAPSETTING"
                If EsGM Then
                    Call HandleMapSetting(ArgumentosAll, CantidadArgumentos)
                End If
                
            Case "/MAPINFO"
                If EsGM Then
                    Call WriteGetMapInfo
                End If
                
            Case "/GRABAR"
                Call WriteSaveChars
                
            Case "/BORRAR"

                If notNullArguments Then

                    Select Case UCase(ArgumentosAll(0))

                        Case "SOS" ' "/BORRAR SOS"
                            Call WriteCleanSOS
                            
                    End Select

                End If
                
            Case "/NOCHE"
                Call WriteNight
                
            Case "/DIA"
                Call WriteDay
                
            Case "/ECHARTODOSPJS"
                Call WriteKickAllChars

            Case "/RELOADNPCS"
                Call WriteReloadNPCs
                
            Case "/RELOADSINI"
                Call WriteReloadServerIni
            
            Case "/HOGAR"
                Call WriteHome
            
            Case "/RELOADHECHIZOS"
                Call WriteReloadSpells
                
            Case "/RELOADOBJ"
                Call WriteReloadObjects
                 

            Case "/CHATCOLOR"

                If notNullArguments And CantidadArgumentos >= 3 Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) Then
                        Call WriteChatColor(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Valor incorrecto. Utilice /chatcolor R G B.")

                    End If

                ElseIf Not notNullArguments Then    'Go back to default!
                    Call WriteChatColor(0, 255, 0)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /chatcolor R G B.")

                End If
            
            Case "/IGNORADO"
                Call WriteIgnored
                            
            Case "/CONSOLA"
            
                'Ojo, no usar notNullArguments porque se usa el string Vacío para borrar cartel.
                If UserEstado = 1 Then 'Muerto

                    With FontTypes(FontTypeNames.FONTTYPE_INFO)
                        Call ShowConsoleMsg("¡¡Estás muerto!! No puedes comunicarte con el mundo de los vivos.", .red, .green, .blue, .bold, .italic)

                    End With

                Else
                
                    If UserLvl < 5 Then

                        With FontTypes(FontTypeNames.FONTTYPE_GLOBAL)
                            Call ShowConsoleMsg("¡¡Debes ser nivel 5 o superior para usar el global!!!", .red, .green, .blue, .bold, .italic)

                        End With

                    Else
                
                        If CantidadArgumentos > 0 Then
                            ArgumentosRaw = Replace(Right$(ArgumentosRaw, Len(ArgumentosRaw) - 0), "~", "")
                            Call WriteGlobalMessage(ArgumentosRaw)
                        Else
                            'Avisar que falta el parametro
                            Call ShowConsoleMsg("Escriba un mensaje.")

                        End If

                    End If

                End If

            Case "/GLOBAL"
                If EsGM Then
                    Call WriteGlobalOnOff
                End If
                
            Case "/CONSULTA"
                If EsGM Then
                    Call WriteConsulta(ArgumentosRaw)
                End If

            Case "/RETAR", "/RETO"
                frmRetos.Show , frmMain

                If notNullArguments Then
                    Dim Names() As String
                    Names = Split(ArgumentosRaw, "@", frmRetos.Jugador.count - 1)

                    For i = 0 To UBound(Names)
                        frmRetos.Jugador(i + 1).Text = Names(i)
                        frmRetos.Jugador(i + 1).Visible = True
                    Next
                    
                    If UBound(Names) Mod 2 = 1 Then
                        frmRetos.Jugador(UBound(Names) + 2).Visible = True
                    End If
                End If
                
            Case "/ACEPTAR"
                If notNullArguments Then
                    Call WriteAcceptDuel(ArgumentosRaw)
                Else
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /aceptar nick.")
                End If
                
            Case "/CANCELAR"
                Call WriteCancelDuel
                
            Case "/ABANDONAR"
                Call WriteQuitDuel
                
            Case "/CE"
                If EsGM Then
                    Call WriteCreateEvent(ArgumentosRaw)
                End If
                
            Case "/RESET"
                If EsGM Then
                    If notNullArguments Then
                        Call WriteResetChar(ArgumentosRaw)
                    Else
                        Call ShowConsoleMsg("Faltan parámetros. Utilice /reset nick.")
                    End If
                End If
            Case "/REQDEBUG"
                Call HandleReqDebugCmd(ArgumentosAll, CantidadArgumentos)
            Case "/FEATURETOGGLE"
                Call HandleFeatureToggle(ArgumentosAll, CantidadArgumentos)
            Case Else
                Call ShowConsoleMsg("El comando es invalido.")

        End Select
        
    ElseIf Left$(Comando, 1) = "-" Then

        If UserEstado = 1 Then 'Muerto

            With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)

            End With

            Exit Sub

        End If

        ' Gritar
        Call WriteYell(mid$(RawCommand, 2))
        
    Else
        Call WriteTalk(RawCommand)
    End If

    
    Exit Sub

ParseUserCommand_Err:
    Call RegistrarError(Err.Number, Err.Description, "ProtocolCmdParse.ParseUserCommand", Erl)
    Resume Next
    
End Sub

Private Sub HandleMapSetting(ByRef arguments() As String, ByVal argCount As Integer)
    If argCount < 2 Then
        Call ShowConsoleMsg("Parametros incorrectos.")
        Exit Sub
    End If
    If Not ValidNumber(arguments(1), eNumber_Types.ent_Byte) Then
        Call ShowConsoleMsg("Parametros incorrectos.")
        Exit Sub
    End If
    Dim settingType As Byte
    Select Case Trim$(UCase$(arguments(0)))
        Case "DROPITEMS"
            settingType = 0
        Case "SAFEPVP"
            settingType = 1
        Case Else
            Call ShowConsoleMsg("Parametros incorrectos.")
            Exit Sub
    End Select
    Call WriteChangeMapSetting(settingType, arguments(1))
End Sub

Private Sub HandleFeatureToggle(ByRef arguments() As String, ByVal argCount As Integer)
    If EsGM Then
        If argCount < 2 Then
            Call ShowConsoleMsg("Parametros incorrectos.")
            Exit Sub
        End If
        Dim varName As String
            Dim value As Byte
            varName = arguments(0)
            If Not ValidNumber(arguments(1), eNumber_Types.ent_Byte) Then
                Call ShowConsoleMsg("Parametros incorrectos.")
                Exit Sub
            End If
            value = arguments(1)
            Call WriteFeatureEnable(varName, value)
    End If
End Sub

Private Sub HandleReqDebugCmd(ByRef arguments() As String, ByVal argCount As Integer)
    If EsGM Then
        If argCount = 0 Then
            Call WriteRequestDebug(e_DebugCommands.eGetLastLogs, arguments(), 0)
        Else
            Dim eType As String
            eType = Trim$(UCase$(arguments(0)))
            If eType = "CONNECTION" Then
                Dim username As String
                Dim i As Integer
                For i = 1 To argCount - 1
                    username = username & arguments(i)
                    If i < argCount - 1 Then
                        username = username & " "
                    End If
                Next i
                arguments(0) = username
                Call WriteRequestDebug(e_DebugCommands.eConnectionState, arguments, 1)
            Else
                Call ShowConsoleMsg("Parametros incorrectos.")
            End If
        End If
    End If
End Sub

Private Sub StartCaptureTheFlag(ByRef arguments() As String, ByVal argCount As Integer)
    If argCount >= 6 Then
        If ValidNumber(arguments(1), eNumber_Types.ent_Long) And ValidNumber(arguments(2), eNumber_Types.ent_Long) And ValidNumber(arguments(3), eNumber_Types.ent_Long) And ValidNumber(arguments(4), eNumber_Types.ent_Long) And ValidNumber(arguments(5), eNumber_Types.ent_Long) Then
            Call WrtieStartCapture(arguments(1), arguments(2), arguments(3), arguments(4), arguments(5))
        Else
            'No es numerico
            Call ShowConsoleMsg("Valor incorrecto. Utilice /CREAREVENTO CAPTURA PARTICIPANTES CANTIDAD_RONDAS NIVEL_MINIMO NIVEL_MAXIMO PRECIO.")
        End If
    Else
        Call ShowConsoleMsg("Valor incorrecto. Utilice /CREAREVENTO CAPTURA PARTICIPANTES CANTIDAD_RONDAS NIVEL_MINIMO NIVEL_MAXIMO PRECIO.")
    End If
End Sub

Private Sub StartLobby(ByRef arguments() As String, ByVal argCount As Integer)
    If argCount >= 3 Then
        If ValidNumber(arguments(1), eNumber_Types.ent_Long) And ValidNumber(arguments(2), eNumber_Types.ent_Long) And ValidNumber(arguments(3), eNumber_Types.ent_Long) Then
            Call WriteStartLobby(0, arguments(1), arguments(2), arguments(3))
        Else
            'No es numerico
            Call ShowConsoleMsg("Valor incorrecto. Utilice /CREAREVENTO LOBBY PARTICIPANTES NIVEL_MINIMO NIVEL_MAXIMO.")
        End If
    Else
        Call ShowConsoleMsg("Valor incorrecto. Utilice /CREAREVENTO LOBBY PARTICIPANTES NIVEL_MINIMO NIVEL_MAXIMO.")
    End If
End Sub

Private Sub StartCustomMap(ByVal mapType As Byte, ByVal name As String, ByRef arguments() As String, ByVal argCount As Integer)
    If argCount >= 3 Then
        If ValidNumber(arguments(1), eNumber_Types.ent_Long) And ValidNumber(arguments(2), eNumber_Types.ent_Long) And ValidNumber(arguments(3), eNumber_Types.ent_Long) Then
            Call WriteStartLobby(mapType, arguments(1), arguments(2), arguments(3))
        Else
            'No es numerico
            Call ShowConsoleMsg("Valor incorrecto. Utilice /CREAREVENTO " & name & " PARTICIPANTES NIVEL_MINIMO NIVEL_MAXIMO.")
        End If
    Else
        Call ShowConsoleMsg("Valor incorrecto. Utilice /CREAREVENTO " & name & " PARTICIPANTES NIVEL_MINIMO NIVEL_MAXIMO.")
    End If
End Sub

Private Sub CreateEventCmd(ByRef arguments() As String, ByVal argCount As Integer)
    If argCount > 0 Then
        Dim eType As String
        eType = Trim$(UCase$(arguments(0)))
        If eType = "CAPTURA" Then
            Call StartCaptureTheFlag(arguments, argCount)
        ElseIf eType = "LOBBY" Then
            Call StartLobby(arguments, argCount)
        ElseIf eType = "CACERIA" Then
            Call StartCustomMap(2, eType, arguments, argCount)
        ElseIf eType = "DEATHMATCH" Then
            Call StartCustomMap(3, eType, arguments, argCount)
        ElseIf eType = "NAVALCONQUEST" Then
            Call StartCustomMap(4, eType, arguments, argCount)
        Else
            Call ShowConsoleMsg("Invalid event type.")
        End If
        
    Else
        'Avisar que falta el parametro
        Call ShowConsoleMsg("Faltan parámetros. Utilice /CREAREVENTO CAPTURA/LOBBY/CACERIA/DEATHMATCH.")
    End If
End Sub

Private Sub ConfigLobbyClass(ByRef arguments() As String, ByVal argCount As Integer)
    If argCount > 1 Then
        Dim eType As String
        eType = Trim$(UCase$(arguments(1)))
        If eType = "MAGE" Or eType = "MAGO" Then
            Call WriteLobbyCommand(e_LobbyCommandId.eSetClassLimit, 1)
        ElseIf eType = "CLERIC" Or eType = "CLERIGO" Then
            Call WriteLobbyCommand(e_LobbyCommandId.eSetClassLimit, 2)
        ElseIf eType = "WARRIOR" Or eType = "GUERRERO" Then
            Call WriteLobbyCommand(e_LobbyCommandId.eSetClassLimit, 3)
        ElseIf eType = "ASSASIN" Or eType = "ASESINO" Then
            Call WriteLobbyCommand(e_LobbyCommandId.eSetClassLimit, 4)
        ElseIf eType = "BARD" Or eType = "BARDO" Then
            Call WriteLobbyCommand(e_LobbyCommandId.eSetClassLimit, 5)
        ElseIf eType = "DRUID" Or eType = "DRUIDA" Then
            Call WriteLobbyCommand(e_LobbyCommandId.eSetClassLimit, 6)
        ElseIf eType = "PALADIN" Or eType = "PALADIN" Then
            Call WriteLobbyCommand(e_LobbyCommandId.eSetClassLimit, 7)
        ElseIf eType = "HUNTER" Or eType = "CAZADOR" Then
            Call WriteLobbyCommand(e_LobbyCommandId.eSetClassLimit, 8)
        ElseIf eType = "WORKER" Or eType = "TRABAJADOR" Then
            Call WriteLobbyCommand(e_LobbyCommandId.eSetClassLimit, 9)
        ElseIf eType = "PIRATE" Or eType = "PIRATA" Then
            Call WriteLobbyCommand(e_LobbyCommandId.eSetClassLimit, 10)
        ElseIf eType = "THIEF" Or eType = "LADRON" Then
            Call WriteLobbyCommand(e_LobbyCommandId.eSetClassLimit, 11)
        ElseIf eType = "BANDIT" Or eType = "BANDIDO" Then
            Call WriteLobbyCommand(e_LobbyCommandId.eSetClassLimit, 12)
        End If
    Else
        Call ShowConsoleMsg("Valor incorrecto. Utilice /CONFIGLOBBY CLASS MAGE/CLERIC/WARRIOR/ASSASIN/BARD/DRUID/PALADIN/HUNTER/WORKER/PIRATE/THIEF/BANDIT")
    End If
End Sub

Private Sub ConfigLobbyTeamCount(ByRef arguments() As String, ByVal argCount As Integer)
    If argCount > 1 Then
        If ValidNumber(arguments(1), eNumber_Types.ent_Long) And ValidNumber(arguments(2), eNumber_Types.ent_Long) And ValidNumber(arguments(3), eNumber_Types.ent_Long) Then
            
        Else
            'No es numerico
            Call ShowConsoleMsg("Valor incorrecto. Utilice /CONFIGLOBBY SETTEAMCOUNT NUMBER")
        End If
    Else
        Call ShowConsoleMsg("Valor incorrecto. Utilice /CONFIGLOBBY SETTEAMCOUNT NUMBER")
    End If
End Sub

Private Sub ConfigLobbyMaxLevel(ByRef arguments() As String, ByVal argCount As Integer)
    If argCount >= 2 Then
        If ValidNumber(arguments(1), eNumber_Types.ent_Long) Then
            Call WriteLobbyCommand(e_LobbyCommandId.eSetMaxLevel, arguments(1))
        Else
            'No es numerico
            Call ShowConsoleMsg("Valor incorrecto. Utilice /CONFIGLOBBY MAXLVL LVL")
        End If
    Else
        Call ShowConsoleMsg("Valor incorrecto. Utilice /CONFIGLOBBY MAXLVL LVL")
    End If
End Sub

Private Sub ConfigLobbyMinLevel(ByRef arguments() As String, ByVal argCount As Integer)
    If argCount >= 2 Then
        If ValidNumber(arguments(1), eNumber_Types.ent_Long) Then
            Call WriteLobbyCommand(e_LobbyCommandId.eSetMinLevel, arguments(1))
        Else
            'No es numerico
            Call ShowConsoleMsg("Valor incorrecto. Utilice /CONFIGLOBBY MINLVL LVL")
        End If
    Else
        Call ShowConsoleMsg("Valor incorrecto. Utilice /CONFIGLOBBY MINLVL LVL")
    End If
End Sub

Private Sub ConfigLobbySummonPlayer(ByRef arguments() As String, ByVal argCount As Integer)
    If argCount >= 2 Then
        If ValidNumber(arguments(1), eNumber_Types.ent_Long) Then
            Call WriteLobbyCommand(e_LobbyCommandId.eSummonSinglePlayer, arguments(1))
        Else
            'No es numerico
            Call ShowConsoleMsg("Valor incorrecto. Utilice /CONFIGLOBBY SUMPLAYER LOBBY_INDEX")
        End If
    Else
        Call ShowConsoleMsg("Valor incorrecto. Utilice /CONFIGLOBBY SUMPLAYER LOBBY_INDEX")
    End If
End Sub

Private Sub ConfigLobbyReturnPlayer(ByRef arguments() As String, ByVal argCount As Integer)
    If argCount >= 2 Then
        If ValidNumber(arguments(1), eNumber_Types.ent_Long) Then
            Call WriteLobbyCommand(e_LobbyCommandId.eReturnSinglePlayer, arguments(1))
        Else
            'No es numerico
            Call ShowConsoleMsg("Valor incorrecto. Utilice /CONFIGLOBBY RETPL LOBBY_INDEX")
        End If
    Else
        Call ShowConsoleMsg("Valor incorrecto. Utilice /CONFIGLOBBY RETPL LOBBY_INDEX")
    End If
End Sub

Private Sub ConfigLobbyOpenLobby(ByRef arguments() As String, ByVal argCount As Integer)
    If argCount >= 2 Then
        If Trim$(UCase$(arguments(1))) = "PRIVATE" Then
            Call WriteLobbyCommand(e_LobbyCommandId.eOpenLobby, "0")
        Else
            Call WriteLobbyCommand(e_LobbyCommandId.eOpenLobby, "1")
        End If
    Else
        Call WriteLobbyCommand(e_LobbyCommandId.eOpenLobby, "1")
        Call ShowConsoleMsg("Valor incorrecto. Utilice /CONFIGLOBBY OPEN [PRIVATE/PUBLIC(DEFAULT)]")
    End If
End Sub

Private Sub ConfigLobbyAddPlayer(ByRef arguments() As String, ByVal argCount As Integer)
    If argCount >= 2 Then
        Dim PlayerName As String
        Dim i As Integer
        For i = 1 To argCount - 1
            PlayerName = PlayerName & arguments(i)
        Next i
        Call WriteLobbyCommand(e_LobbyCommandId.eAddPlayer, PlayerName)
    Else
        Call ShowConsoleMsg("Valor incorrecto. Utilice /CONFIGLOBBY ADDPLAYER player name")
    End If
End Sub

Private Sub ConfigLobbySetTeamSize(ByRef arguments() As String, ByVal argCount As Integer)
    If argCount >= 2 Then
        Dim premade As Byte
        premade = 1
        If argCount >= 3 Then
            If Trim$(UCase$(arguments(2))) = "PREMADE" Then
              premade = 0
            End If
        End If
        If ValidNumber(arguments(1), eNumber_Types.ent_Long) Then
            Call WriteLobbyCommand(e_LobbyCommandId.eSetTeamSize, arguments(1) & " " & premade)
        Else
            Call ShowConsoleMsg("Valor incorrecto. Utilice /CONFIGLOBBY SETTEAMSIZE NUMBER_OF_TEAMS [RANDOM/PREMADE(DEFAULT RANDOM)]")
        End If
    Else
        Call ShowConsoleMsg("Valor incorrecto. Utilice /CONFIGLOBBY SETTEAMSIZE NUMBER_OF_TEAMS [RANDOM/PREMADE(DEFAULT RANDOM)]")
    End If
End Sub

Private Sub ConfigLobbyKickPlayer(ByRef arguments() As String, ByVal argCount As Integer)
    If argCount >= 2 Then
        If ValidNumber(arguments(2), eNumber_Types.ent_Long) Then
            Call WriteLobbyCommand(e_LobbyCommandId.eKickPlayer, arguments(2))
        Else
            Call ShowConsoleMsg("Valor incorrecto. Utilice /CONFIGLOBBY KICK LOBBY_INDEX")
        End If
    Else
        Call ShowConsoleMsg("Valor incorrecto. Utilice /CONFIGLOBBY KICK LOBBY_INDEX")
    End If
End Sub

Private Sub ConfigLobby(ByRef arguments() As String, ByVal argCount As Integer)
    If argCount > 0 Then
        Dim eType As String
        eType = Trim$(UCase$(arguments(0)))
        If eType = "SPAWN" Then
            Call WriteLobbyCommand(e_LobbyCommandId.eSetSpawnPos)
        ElseIf eType = "MAXLVL" Then
            Call ConfigLobbyMaxLevel(arguments, argCount)
        ElseIf eType = "MINLVL" Then
            Call ConfigLobbyMinLevel(arguments, argCount)
        ElseIf eType = "SUMPLAYER" Then
            Call ConfigLobbySummonPlayer(arguments, argCount)
        ElseIf eType = "SUMALL" Then
            Call WriteLobbyCommand(e_LobbyCommandId.eSummonAll)
        ElseIf eType = "RETURNPLAYER" Then
            Call ConfigLobbyReturnPlayer(arguments, argCount)
        ElseIf eType = "RETALL" Then
            Call WriteLobbyCommand(e_LobbyCommandId.eReturnAllSummoned)
        ElseIf eType = "OPEN" Then
            Call ConfigLobbyOpenLobby(arguments, argCount)
        ElseIf eType = "START" Then
            Call WriteLobbyCommand(e_LobbyCommandId.eStartEvent)
        ElseIf eType = "END" Then
            Call WriteLobbyCommand(e_LobbyCommandId.eEndEvent)
        ElseIf eType = "LIST" Then
            Call WriteLobbyCommand(e_LobbyCommandId.eListPlayers)
        ElseIf eType = "CLASS" Then
            Call ConfigLobbyClass(arguments(), argCount)
        ElseIf eType = "KICK" Then
            Call ConfigLobbyKickPlayer(arguments(), argCount)
        ElseIf eType = "FORCERESET" Then
            Call WriteLobbyCommand(e_LobbyCommandId.eForceReset)
        ElseIf eType = "SETTEAMSIZE" Then
            Call ConfigLobbySetTeamSize(arguments(), argCount)
        ElseIf eType = "ADDPLAYER" Then
            Call ConfigLobbyAddPlayer(arguments(), argCount)
        Else
            Call ShowConsoleMsg("Parametro invalido. Utilice /CONFIGLOBBY SPAWN/MAXLVL/MINLVL/CLASS/SUMPLAYER/SUMALL/RETURNPLAYER/RETALL/OPEN/START/END/LIST/KICK/FORCERESET/SETTEAMSIZE/ADDPLAYER")
        End If
    Else
        'Avisar que falta el parametro
        Call ShowConsoleMsg("Faltan parámetros. Utilice /CONFIGLOBBY SPAWN/MAXLVL/MINLVL/CLASS/SUMPLAYER/SUMALL/RETURNPLAYER/RETALL/OPEN/START/END/LIST/KICK")
    End If
End Sub
''
' Show a console message.
'
' @param    Message The message to be written.
' @param    red Sets the font red color.
' @param    green Sets the font green color.
' @param    blue Sets the font blue color.
' @param    bold Sets the font bold style.
' @param    italic Sets the font italic style.

Public Sub ShowConsoleMsg(ByVal Message As String, Optional ByVal red As Integer = 255, Optional ByVal green As Integer = 255, Optional ByVal blue As Integer = 255, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False)
    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 01/03/07
    '
    '***************************************************
    
    On Error GoTo ShowConsoleMsg_Err
    
    Call AddtoRichTextBox(frmMain.RecTxt, Message, red, green, blue, bold, italic)

    
    Exit Sub

ShowConsoleMsg_Err:
    Call RegistrarError(Err.Number, Err.Description, "ProtocolCmdParse.ShowConsoleMsg", Erl)
    Resume Next
    
End Sub

''
' Returns whether the number is correct.
'
' @param    Numero The number to be checked.
' @param    Tipo The acceptable type of number.

Public Function ValidNumber(ByVal Numero As String, ByVal TIPO As eNumber_Types) As Boolean
    
    On Error GoTo ValidNumber_Err
    

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 01/06/07
    '
    '***************************************************
    Dim Minimo As Long

    Dim Maximo As Long
    
    If Not IsNumeric(Numero) Then Exit Function
    
    Select Case TIPO

        Case eNumber_Types.ent_Byte
            Minimo = 0
            Maximo = 255

        Case eNumber_Types.ent_Integer
            Minimo = -32768
            Maximo = 32767

        Case eNumber_Types.ent_Long
            Minimo = -2147483648#
            Maximo = 2147483647
        
        Case eNumber_Types.ent_Trigger
            Minimo = 0
            Maximo = 99

    End Select
    
    If Val(Numero) >= Minimo And Val(Numero) <= Maximo Then ValidNumber = True

    
    Exit Function

ValidNumber_Err:
    Call RegistrarError(Err.Number, Err.Description, "ProtocolCmdParse.ValidNumber", Erl)
    Resume Next
    
End Function

''
' Returns whether the ip format is correct.
'
' @param    IP The ip to be checked.

Private Function validipv4str(ByVal IP As String) As Boolean
    
    On Error GoTo validipv4str_Err
    

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 01/06/07
    '
    '***************************************************
    Dim tmpArr() As String
    
    tmpArr = Split(IP, ".")
    
    If UBound(tmpArr) <> 3 Then Exit Function

    If Not ValidNumber(tmpArr(0), eNumber_Types.ent_Byte) Or Not ValidNumber(tmpArr(1), eNumber_Types.ent_Byte) Or Not ValidNumber(tmpArr(2), eNumber_Types.ent_Byte) Or Not ValidNumber(tmpArr(3), eNumber_Types.ent_Byte) Then Exit Function
    
    validipv4str = True

    
    Exit Function

validipv4str_Err:
    Call RegistrarError(Err.Number, Err.Description, "ProtocolCmdParse.validipv4str", Erl)
    Resume Next
    
End Function

''
' Converts a string into the correct ip format.
'
' @param    IP The ip to be converted.

Private Function str2ipv4l(ByVal IP As String) As Byte()
    
    On Error GoTo str2ipv4l_Err
    

    '***************************************************
    'Author: Nicolas Matias Gonzalez (NIGO)
    'Last Modification: 07/26/07
    'Last Modified By: Rapsodius
    'Specify Return Type as Array of Bytes
    'Otherwise, the default is a Variant or Array of Variants, that slows down
    'the function
    '***************************************************
    Dim tmpArr() As String

    Dim bArr(3)  As Byte
    
    tmpArr = Split(IP, ".")
    
    bArr(0) = CByte(tmpArr(0))
    bArr(1) = CByte(tmpArr(1))
    bArr(2) = CByte(tmpArr(2))
    bArr(3) = CByte(tmpArr(3))

    str2ipv4l = bArr

    
    Exit Function

str2ipv4l_Err:
    Call RegistrarError(Err.Number, Err.Description, "ProtocolCmdParse.str2ipv4l", Erl)
    Resume Next
    
End Function

''
' Do an Split() in the /AEMAIL in onother way
'
' @param text All the comand without the /aemail
' @return An bidimensional array with user and mail

Private Function AEMAILSplit(ByRef Text As String) As String()
    
    On Error GoTo AEMAILSplit_Err
    

    '***************************************************
    'Author: Lucas Tavolaro Ortuz (Tavo)
    'Useful for AEMAIL BUG FIX
    'Last Modification: 07/26/07
    'Last Modified By: Rapsodius
    'Specify Return Type as Array of Strings
    'Otherwise, the default is a Variant or Array of Variants, that slows down
    'the function
    '***************************************************
    Dim tmpArr(0 To 1) As String

    Dim Pos            As Byte
    
    Pos = InStr(1, Text, "-")
    
    If Pos <> 0 Then
        tmpArr(0) = mid$(Text, 1, Pos - 1)
        tmpArr(1) = mid$(Text, Pos + 1)
    Else
        tmpArr(0) = vbNullString

    End If
    
    AEMAILSplit = tmpArr

    
    Exit Function

AEMAILSplit_Err:
    Call RegistrarError(Err.Number, Err.Description, "ProtocolCmdParse.AEMAILSplit", Erl)
    Resume Next
    
End Function
