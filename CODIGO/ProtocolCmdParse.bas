Attribute VB_Name = "ProtocolCmdParse"
'RevolucionAo 1.0
'Pablo Mercavides

Option Explicit

Public Enum eNumber_Types

    ent_Byte
    ent_Integer
    ent_Long
    ent_Trigger

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
                If UserParalizado Or UserInmovilizado Then 'Inmo

                    With FontTypes(FontTypeNames.FONTTYPE_WARNING)
                        Call ShowConsoleMsg("No puedes salir estando paralizado.", .red, .green, .blue, .bold, .italic)

                    End With

                    Exit Sub

                End If
                
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
                
            Case "/DUELO"
                Call WriteDuelo
                
            Case "/RECOMPENSA"
                Call WriteReward
                
            Case "/MOTD"
                Call WriteRequestMOTD
                
            Case "/UPTIME"
                Call WriteUpTime
                
            Case "/ENCUESTA"
                If CantidadArgumentos = 0 Then
                    ' Version sin argumentos: Inquiry
                    Call WriteInquiry
                Else

                    ' Version con argumentos: InquiryVote
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Byte) Then
                        Call WriteInquiryVote(ArgumentosRaw)
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Para votar una opcion, escribe /encuesta NUMERODEOPCION, por ejemplo para votar la opcion 1, escribe /encuesta 1.")

                    End If

                End If
        
            Case "/CMSG"
                'Ojo, no usar notNullArguments porque se usa el string vacio para borrar cartel.
                If CantidadArgumentos > 0 Then
                    Call WriteGuildMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")

                End If
                
            Case "/GRUPO"
                'Ojo, no usar notNullArguments porque se usa el string vacio para borrar cartel.
                If CantidadArgumentos > 0 Then
                    Call WriteGrupoMsg(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")

                End If
            
            Case "/CENTINELA"
                If notNullArguments Then
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Integer) Then
                        Call WriteCentinelReport(CInt(ArgumentosRaw))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("El código de verificación debe ser numerico. Utilice /centinela X, siendo X el código de verificación.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /centinela X, siendo X el código de verificación.")

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
                        Call ShowConsoleMsg("Tu descripción no puede ser tán larga (Max. 50 caracteres).", .red, .green, .blue, .bold, .italic)

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
                If notNullArguments Then
                    Call WritePunishments(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /penas NICKNAME.")

                End If
                
            Case "/CONTRASEÑA"
                Call frmNewPassword.Show(vbModal, frmMain)
            
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
                    Call WriteLeaveFaction
                Else

                    ' Version con argumentos: BankExtractGold
                    If ValidNumber(ArgumentosRaw, eNumber_Types.ent_Long) Then
                        Call WriteBankExtractGold(ArgumentosRaw)
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Cantidad incorrecta. Utilice /retirar CANTIDAD.")

                    End If

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
                    Call ShowConsoleMsg("Faltan paramtetros. Utilice /depositar CANTIDAD.")

                End If
                
            Case "/DENUNCIAR"
                'If notNullArguments Then
                '  Call WriteDenounce(ArgumentosRaw)
                'Else
                'Avisar que falta el parametro
                Call ShowConsoleMsg("Este comando fue desactivado. Utilice /GM para contactar a un administrador del juego.")
                'End If
                
            Case "/FINALIZAREVENTO"
                Call WriteDenounce

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
                
            Case "/REM"
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
                    Call Protocol.WriteServerTime
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
                
            Case "/SILENCIAR"
                If notNullArguments Then
                    Call WriteSilence(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /silenciar NICKNAME.")

                End If
                
            Case "/CUENTAREGRESIVA"
                If notNullArguments Then
                    Call WriteCuentaRegresiva(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /CUENTAREGRESIVA TIEMPO (En segundos).")

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
                    If notNullArguments Then
                        Call WriteGoToChar(ArgumentosRaw)
                    Else
                        'Avisar que falta el parametro
                        Call ShowConsoleMsg("Faltan parámetros. Utilice /ira NICKNAME.")

                    End If

                End If
                
            Case "/GO"
                If EsGM Then
                    If notNullArguments Then
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) Then
                            Call WriteWarpChar("YO", ArgumentosAll(0), 50, 50)
                        Else
                            Call WriteGoToChar(ArgumentosRaw)

                        End If

                    Else
                        'Avisar que falta el parametro
                        Call ShowConsoleMsg("Faltan parámetros. Utilice /go NICKNAME o /go MAPA.")

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
                
            Case "/LUZMAPA"
                If EsGM Then
                    If notNullArguments Then
                        If CantidadArgumentos = 3 Then
                            If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) And _
                                ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) And _
                                ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Integer) Then
                                
                                Call SetGlobalLight(D3DColorXRGB(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2)))
                                Call MapUpdateGlobalLight
                                Exit Sub

                            End If
                        End If
                    End If

                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /luzmapa R G B.")
                End If
                
            Case "/DESBUGGEAR"
                If EsGM Then
                    Call WriteDesbuggear(ArgumentosRaw)
                End If
                
            Case "/DARLLAVE"
                If EsGM Then
                    If notNullArguments Or CantidadArgumentos < 2 Then
                        If Not ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Integer) Then
                            Call ShowConsoleMsg("Número de llave inválida. Utilice /darllave NICKNAME LLAVE(ID DE OBJETO)")
                        Else
                            Call WriteDarLlaveAUsuario(ArgumentosAll(0), Val(ArgumentosAll(1)))
                        End If
                    Else
                        'Avisar que falta el parametro
                        Call ShowConsoleMsg("Faltan parámetros. Utilice /darllave NICKNAME LLAVE")
                    End If
                End If
                
            Case "/SACARLLAVE"
                If EsGM Then
                    If notNullArguments Then
                        If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) Then
                            Call WriteSacarLlave(Val(ArgumentosAll(0)))
                        Else
                            Call ShowConsoleMsg("Parámetro inválido. Utilice /sacarllave LLAVE(ID DE OBJETO)")
                        End If
                    Else
                        'Avisar que falta el parametro
                        Call ShowConsoleMsg("Faltan parámetros. Utilice /sacarllave LLAVE")
                    End If
                End If
                
            Case "/VERLLAVES"
                Call WriteVerLlaves
        
            Case "/INVISIBLE"
                Call WriteInvisible
            
            Case "/PAREJA"
                Call WritePareja
                
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

                        Case "BODY"
                            tmpInt = eEditOptions.eo_Body
                            
                        Case "ARMA"
                            tmpInt = eEditOptions.eo_Arma
                            
                        Case "CASCO"
                            tmpInt = eEditOptions.eo_Casco
                            
                        Case "ESCUDO"
                            tmpInt = eEditOptions.eo_Escudo
                            
                        Case "PARTICULA"
                            tmpInt = eEditOptions.eo_Particula
                            
                        Case "HEAD"
                            tmpInt = eEditOptions.eo_Head
                        
                        Case "ORO"
                            tmpInt = eEditOptions.eo_Gold
                        
                        Case "LEVEL"
                            tmpInt = eEditOptions.eo_Level
                        
                        Case "SKILLS"
                            tmpInt = eEditOptions.eo_Skills
                        
                        Case "SKILLSLIBRES"
                            tmpInt = eEditOptions.eo_SkillPointsLeft
                        
                        Case "CLASE"
                            tmpInt = eEditOptions.eo_Class
                        
                        Case "EXP"
                            tmpInt = eEditOptions.eo_Experience
                        
                        Case "CRI"
                            tmpInt = eEditOptions.eo_CriminalsKilled
                        
                        Case "CIU"
                            tmpInt = eEditOptions.eo_CiticensKilled
                        
                        Case "SEX"
                            tmpInt = eEditOptions.eo_Sex
                            
                        Case "RAZA"
                            tmpInt = eEditOptions.eo_Raza
                        
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
                
            Case "/SILENCIO"

                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@", 2)

                    If UBound(tmpArr) = 1 Then
                        Call WriteSilenciarUser(tmpArr(0), tmpArr(1))
                    Else
                        'Faltan los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /SILENCIO NICKNAME@MOTIVO.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /SILENCIO NICKNAME@MOTIVO.")

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
                    frmSpawnList.FillList
                    frmSpawnList.Show , frmMain
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
                
            Case "/CT"

                If notNullArguments And CantidadArgumentos = 3 Then
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) And ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Byte) And ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) Then
                        Call WriteTeleportCreate(ArgumentosAll(0), ArgumentosAll(1), ArgumentosAll(2))
                    Else
                        'No es numerico
                        Call ShowConsoleMsg("Valor incorrecto. Utilice /ct MAPA X Y.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /ct MAPA X Y.")

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
                
            Case "/CIUMSG"

                If notNullArguments Then
                    Call WriteCitizenMessage(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Escriba un mensaje.")

                End If
            
            Case "/CRIMSG"

                If notNullArguments Then
                    Call WriteCriminalMessage(ArgumentosRaw)
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
                
            Case "/DUMPSECURITY"
                Call WriteDumpIPTables
                
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

                If CantidadArgumentos >= 2 Then
                    If validipv4str(ArgumentosAll(0)) Then
                        Call WriteBanIP(True, str2ipv4l(ArgumentosAll(0)), vbNullString, Right$(ArgumentosRaw, Len(ArgumentosRaw) - Len(ArgumentosAll(0)) - 1))
                    Else
                        'No es una IP, es un nick
                        Call WriteBanIP(False, str2ipv4l("0.0.0.0"), ArgumentosAll(0), Right$(ArgumentosRaw, Len(ArgumentosRaw) - Len(ArgumentosAll(0)) - 1))

                    End If

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
                
            Case "/LASTEMAIL"

                If notNullArguments Then
                    Call WriteRequestCharMail(ArgumentosRaw)
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /lastemail NICKNAME.")

                End If
                
            Case "/APASS"

                If notNullArguments Then
                    tmpArr = Split(ArgumentosRaw, "@", 2)

                    If UBound(tmpArr) = 1 Then
                        Call WriteAlterPassword(tmpArr(0), tmpArr(1))
                    Else
                        'Faltan los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /apass PJSINPASS@PJCONPASS.")

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /apass PJSINPASS@PJCONPASS.")

                End If
                
            Case "/AEMAIL"

                If notNullArguments Then
                    tmpArr = AEMAILSplit(ArgumentosRaw)

                    If LenB(tmpArr(0)) = 0 Then
                        'Faltan los parametros con el formato propio
                        Call ShowConsoleMsg("Formato incorrecto. Utilice /aemail NICKNAME-NUEVOMAIL.")
                    Else
                        Call WriteAlterMail(tmpArr(0), tmpArr(1))

                    End If

                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Faltan parámetros. Utilice /aemail NICKNAME-NUEVOMAIL.")

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
            
            Case "/CREARPRETORIANOS"
            
                If CantidadArgumentos = 3 Then
                    
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) And _
                       ValidNumber(ArgumentosAll(1), eNumber_Types.ent_Byte) And _
                       ValidNumber(ArgumentosAll(2), eNumber_Types.ent_Byte) Then
                       
                        Call WriteCreatePretorianClan(Val(ArgumentosAll(0)), Val(ArgumentosAll(1)), _
                                                      Val(ArgumentosAll(2)))
                    Else
                        'Faltan o sobran los parametros con el formato propio
                        Call ShowConsoleMsg("Formato inválido. Es /CrearPretorianos MAPA X Y.")
                    End If
                    
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Formato inválido. Es /CrearPretorianos MAPA X Y.")
                End If
                
            Case "/ELIMINARPRETORIANOS"
            
                If CantidadArgumentos = 1 Then
                    
                    If ValidNumber(ArgumentosAll(0), eNumber_Types.ent_Integer) Then
                        Call WriteDeletePretorianClan(Val(ArgumentosAll(0)))
                    Else
                        'Faltan o sobran los parametros con el formato propio
                        Call ShowConsoleMsg("Formato inválido. Es /EliminarPretorianos MAPA")
                    End If
                    
                Else
                    'Avisar que falta el parametro
                    Call ShowConsoleMsg("Formato inválido. Es /EliminarPretorianos MAPA")
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
                
            Case "/GUARDAMAPA"
                Call WriteSaveMap
                
            Case "/MODMAPINFO" ' PK, BACKUP

                If CantidadArgumentos > 1 Then

                    Select Case UCase$(ArgumentosAll(0))

                        Case "PK" ' "/MODMAPINFO PK"
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
                    Call ShowConsoleMsg("Faltan parametros. Opciones: PK, BACKUP, RESTRINGIR, MAGIASINEFECTO, INVISINEFECTO, RESUSINEFECTO, TERRENO, ZONA")

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
                
            Case "/TCPESSTATS"
                Call WriteRequestTCPStats
                
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
                 
            Case "/REINICIAR"
                Call WriteRestart
                
            Case "/AUTOUPDATE"
                Call WriteResetAutoUpdate
            
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
            
            Case "/PING"
                Call WritePing
                
            Case "/CONSOLA"
            
                'Ojo, no usar notNullArguments porque se usa el string vacio para borrar cartel.
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
                Call WriteGlobalOnOff
                
            Case "/CONSULTA"
                Call WriteConsulta(ArgumentosRaw)
                
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
        ' Hablar
        Call WriteTalk(RawCommand)

    End If

    
    Exit Sub

ParseUserCommand_Err:
    Call RegistrarError(Err.number, Err.Description, "ProtocolCmdParse.ParseUserCommand", Erl)
    Resume Next
    
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
    Call RegistrarError(Err.number, Err.Description, "ProtocolCmdParse.ShowConsoleMsg", Erl)
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
            Maximo = 7

    End Select
    
    If Val(Numero) >= Minimo And Val(Numero) <= Maximo Then ValidNumber = True

    
    Exit Function

ValidNumber_Err:
    Call RegistrarError(Err.number, Err.Description, "ProtocolCmdParse.ValidNumber", Erl)
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
    Call RegistrarError(Err.number, Err.Description, "ProtocolCmdParse.validipv4str", Erl)
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
    Call RegistrarError(Err.number, Err.Description, "ProtocolCmdParse.str2ipv4l", Erl)
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
    Call RegistrarError(Err.number, Err.Description, "ProtocolCmdParse.AEMAILSplit", Erl)
    Resume Next
    
End Function
