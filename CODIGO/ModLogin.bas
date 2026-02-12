Attribute VB_Name = "ModLogin"
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
Dim ServerSettings As clsIniManager

Public ServerIpCount        As Integer
Public LoginServerCount     As Integer
Public GameServerIP()       As String
Public GameServerPort()     As Integer
Public LoginServerIP()      As String
Public LoginServerPort()    As Integer
Private m_ServerVectorsInitialized  As Boolean
Private m_ServerVectorsEnvironment  As String
Public CurrentGameServerIndex  As Integer
Public CurrentLoginServerIndex As Integer

   
Public Sub DoLogin(ByVal Account As String, ByVal Password As String, ByVal storeCredentials As Boolean)
    On Error GoTo DoLogin_Err
    #If REMOTE_CLOSE = 1 Then
        ModAuth.LoginOperation = e_operation.Authenticate
        Call LoginOrConnect(E_MODO.IngresandoConCuenta)
    #Else
        If IntervaloPermiteConectar Then
            CuentaEmail = Account
            CuentaPassword = Password
            If storeCredentials Then
                CuentaRecordada.nombre = CuentaEmail
                CuentaRecordada.Password = CuentaPassword
                Call GuardarCuenta(CuentaEmail, CuentaPassword)
            Else
                ' Reseteamos los datos de cuenta guardados
                Call GuardarCuenta(vbNullString, vbNullString)
            End If
            If CheckUserDataLoged() = True Then
                ModAuth.LoginOperation = e_operation.Authenticate
                Call LoginOrConnect(E_MODO.IngresandoConCuenta)
            End If
        End If
    #End If
    Exit Sub
DoLogin_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModLogin.DoLogin", Erl)
    Resume Next
End Sub

Public Sub SetActiveServer(ByVal IP As String, ByVal port As String)
    ServerIndex = IP & ":" & port
    IPdelServidor = IP
    PuertoDelServidor = port
    #If PYMMO = 0 Or DEBUGGING = 1 Then
        Call SaveSetting("INIT", "ServerIndex", IPdelServidor & ":" & PuertoDelServidor)
    #End If
    
    #If PYMMO = 1 Then
        'DEVELOPER mode is used to connect to localhost
        #If Developer = 1 Then
            IPdelServidorLogin = "127.0.0.1"
            PuertoDelServidorLogin = 4000
            IPdelServidor = IPdelServidorLogin
            PuertoDelServidor = 6501
        #Else
            #If DEBUGGING = 0 Then
                'When not in DEVELOPER mode we read the ip and port from the list
                Call SetDefaultServer
            #Else
                'Staging, set the ip and port for pymmo
                IPdelServidorLogin = "45.235.98.188"
                PuertoDelServidorLogin = 6500 '6502 Is also usable, there are 2 login servers in Staging and Prod
            #End If
        #End If
    #End If
    
    frmDebug.add_text_tracebox "Using Login Server " & IPdelServidorLogin & ":" & PuertoDelServidorLogin
    frmDebug.add_text_tracebox "Using Game Server " & IPdelServidor & ":" & PuertoDelServidor
End Sub


Private Function EnsureServerSettingsLoaded() As Boolean
    On Error GoTo errhandler

    If Not ServerSettings Is Nothing Then
        EnsureServerSettingsLoaded = True
        Exit Function
    End If

    Dim RemotesPath As String

    Set ServerSettings = New clsIniManager

    #If Compresion = 1 Then
        If Not Extract_File(Scripts, App.path & "\..\Recursos\OUTPUT\", "Remotes.ini", Windows_Temp_Dir, ResourcesPassword, False) Then
            Err.Description = "¡No se puede cargar el archivo de Remotes.ini!"
            MsgBox Err.Description
            Set ServerSettings = Nothing
            Exit Function
        End If
        RemotesPath = Windows_Temp_Dir & "Remotes.ini"
    #Else
        RemotesPath = App.path & "\..\Recursos\init\Remotes.ini"
    #End If

    Debug.Assert FileExist(RemotesPath, vbNormal)

    Call ServerSettings.Initialize(RemotesPath)

    EnsureServerSettingsLoaded = True
    Exit Function

errhandler:
    MsgBox "Error inicializando ServerSettings: " & Err.Description, vbCritical
    Set ServerSettings = Nothing
End Function

Private Sub ShuffleIndices(ByRef indices() As Integer)
    Dim i As Integer
    Dim J As Integer
    Dim tmp As Integer

    For i = LBound(indices) To UBound(indices)
        J = RandomNumber(LBound(indices), UBound(indices))

        tmp = indices(i)
        indices(i) = indices(J)
        indices(J) = tmp
    Next i
End Sub

Private Sub LoadServerVectors(ByVal environment As String)
    Dim loginServers As Integer
    Dim serverOrder() As Integer
    Dim loginOrder() As Integer
    Dim k As Integer
    Dim idx As Integer

    ' If already initialized for this environment, do nothing
    If m_ServerVectorsInitialized Then
        If StrComp(m_ServerVectorsEnvironment, environment, vbTextCompare) = 0 Then
            Exit Sub
        End If
    End If

    ' Not initialized yet (or environment changed) -> build everything
    m_ServerVectorsInitialized = False   ' reset until we succeed

    loginServers = val(ServerSettings.GetValue(environment, "LoginCount"))
    ServerIpCount = val(ServerSettings.GetValue(environment, "ServerCount"))
    LoginServerCount = loginServers

    If ServerIpCount <= 0 Or loginServers <= 0 Then
        MsgBox "Configuración de servidores inválida para el entorno: " & environment, vbCritical
        Exit Sub
    End If

    ' Prepare arrays
    ReDim GameServerIP(1 To ServerIpCount)
    ReDim GameServerPort(1 To ServerIpCount)

    ReDim LoginServerIP(1 To loginServers)
    ReDim LoginServerPort(1 To loginServers)

    ' Randomized game server order
    ReDim serverOrder(1 To ServerIpCount)
    For k = 1 To ServerIpCount
        serverOrder(k) = k
    Next k
    ShuffleIndices serverOrder

    ' Randomized login server order
    ReDim loginOrder(1 To loginServers)
    For k = 1 To loginServers
        loginOrder(k) = k
    Next k
    ShuffleIndices loginOrder

    ' Fill game server vectors
    For k = 1 To ServerIpCount
        idx = serverOrder(k)
        GameServerIP(k) = ServerSettings.GetValue(environment, "ServerIp" & idx)
        GameServerPort(k) = val(ServerSettings.GetValue(environment, "PortPort" & idx))
    Next k

    ' Fill login server vectors
    For k = 1 To loginServers
        idx = loginOrder(k)
        LoginServerIP(k) = ServerSettings.GetValue(environment, "LoginIp" & idx)
        LoginServerPort(k) = val(ServerSettings.GetValue(environment, "LoginPort" & idx))
    Next k

    ' Mark as initialized for this environment
    m_ServerVectorsInitialized = True
    m_ServerVectorsEnvironment = environment
End Sub

Public Sub SetActiveEnvironment(ByVal environment As String)
    If Not EnsureServerSettingsLoaded() Then Exit Sub

    #If Developer = 0 And DEBUGGING = 0 Then
        environment = "Production"
    #End If

    ' Build server vectors (only once per environment inside this helper)
    Call LoadServerVectors(environment)

    ' Choose initial indices using the helpers (0 means "no previous")
    If CurrentGameServerIndex = 0 Then
        CurrentGameServerIndex = GetNextGameServer(0)
    End If
    If CurrentLoginServerIndex = 0 Then
        CurrentLoginServerIndex = GetNextLoginServer(0)
    End If
    
    If CurrentGameServerIndex = -1 Or CurrentLoginServerIndex = -1 Then
        MsgBox "No hay servidores configurados para el entorno: " & environment, vbCritical
        Exit Sub
    End If

    ' Set active Game server
    IPdelServidor = GameServerIP(CurrentGameServerIndex)
    PuertoDelServidor = GameServerPort(CurrentGameServerIndex)

    ' Set active Login server
    IPdelServidorLogin = LoginServerIP(CurrentLoginServerIndex)
    PuertoDelServidorLogin = LoginServerPort(CurrentLoginServerIndex)

    frmDebug.add_text_tracebox "Using Login Server #" & CurrentLoginServerIndex & _
                               " -> " & IPdelServidorLogin & ":" & PuertoDelServidorLogin
    frmDebug.add_text_tracebox "Using Game Server #" & CurrentGameServerIndex & _
                               " -> " & IPdelServidor & ":" & PuertoDelServidor
End Sub

Public Function GetNextGameServer(ByVal prevIndex As Integer) As Integer
    If ServerIpCount <= 0 Then
        GetNextGameServer = -1
        Exit Function
    End If

    Dim nextIndex As Integer

    If prevIndex < 1 Or prevIndex > ServerIpCount Then
        nextIndex = 1
    Else
        nextIndex = prevIndex + 1
    End If

    If nextIndex > ServerIpCount Then
        nextIndex = 1
    End If

    GetNextGameServer = nextIndex
End Function


Public Function GetNextLoginServer(ByVal prevIndex As Integer) As Integer
    If LoginServerCount <= 0 Then
        GetNextLoginServer = -1
        Exit Function
    End If

    Dim nextIndex As Integer

    If prevIndex < 1 Or prevIndex > LoginServerCount Then
        nextIndex = 1
    Else
        nextIndex = prevIndex + 1
    End If

    If nextIndex > LoginServerCount Then
        nextIndex = 1
    End If

    GetNextLoginServer = nextIndex
End Function
Public Sub SwitchToNextGameServer(ByRef CurrentIndex As Integer)
    CurrentIndex = GetNextGameServer(CurrentIndex)
    If CurrentIndex = -1 Then Exit Sub

    IPdelServidor = GameServerIP(CurrentIndex)
    PuertoDelServidor = GameServerPort(CurrentIndex)

    frmDebug.add_text_tracebox "Switching to Game Server #" & CurrentIndex & _
                               " -> " & IPdelServidor & ":" & PuertoDelServidor
End Sub


Public Sub SwitchToNextLoginServer(ByRef CurrentIndex As Integer)
    CurrentIndex = GetNextLoginServer(CurrentIndex)
    If CurrentIndex = -1 Then Exit Sub

    IPdelServidorLogin = LoginServerIP(CurrentIndex)
    PuertoDelServidorLogin = LoginServerPort(CurrentIndex)

    frmDebug.add_text_tracebox "Switching to Login Server #" & CurrentIndex & _
                               " -> " & IPdelServidorLogin & ":" & PuertoDelServidorLogin
End Sub


Public Sub CreateAccount(ByVal Name As String, ByVal Surname As String, ByVal Email As String, ByVal Password As String)
    NewAccountData.Name = Name
    NewAccountData.Surname = Surname
    NewAccountData.Email = Email
    NewAccountData.Password = Password
    #If PYMMO = 1 Then
        ModAuth.LoginOperation = e_operation.SignUp
        Call connectToLoginServer
    #Else
        CuentaEmail = Email
        CuentaPassword = Password
        Call LoginOrConnect(CreandoCuenta)
    #End If
End Sub

Public Sub LoadCharacterSelectionScreen()
    AlphaNiebla = 30
    frmConnect.visible = True
    g_game_state.State = e_state_account_screen
    SugerenciaAMostrar = RandomNumber(1, NumSug)
    Call ao20audio.PlayWav(192)
    Call Graficos_Particulas.Particle_Group_Remove_All
    Call Graficos_Particulas.Engine_Select_Particle_Set(203)
    ParticleLluviaDorada = Graficos_Particulas.General_Particle_Create(208, -1, -1)
    If FrmLogear.visible Then Unload FrmLogear
    If frmNewAccount.visible Then Unload frmNewAccount
    If frmMain.visible Then
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
End Sub

Public Sub GoToLogIn()
    g_game_state.State = e_state_connect_screen
End Sub

Public Sub LogOut()
    frmDebug.add_text_tracebox "Vuelvo al login, debería borrar el token"
    Auth_state = e_state.Idle
    Call ComprobarEstado
    UserSaliendo = True
    Call modNetwork.Disconnect
    CantidadDePersonajesEnCuenta = 0
    Dim i As Integer
    For i = 1 To MAX_PERSONAJES_EN_CUENTA
        Pjs(i).Body = 0
        Pjs(i).Head = 0
        Pjs(i).Mapa = 0
        Pjs(i).Nivel = 0
        Pjs(i).nombre = ""
        Pjs(i).Clase = 0
        Pjs(i).Criminal = 0
        Pjs(i).NameMapa = ""
    Next i
    General_Set_Connect
End Sub

Public Sub ResendValidationCode(ByVal Email As String)
    CuentaEmail = Email
    ModAuth.LoginOperation = e_operation.RequestVerificationCode
    Call connectToLoginServer
End Sub

Public Sub ValidateCode(ByVal Email As String, ByVal code As String)
    CuentaEmail = Email
    ValidationCode = code
    ModAuth.LoginOperation = e_operation.ValidateAccount
    Call connectToLoginServer
End Sub

Public Sub RequestPasswordReset(ByVal Email As String)
    CuentaEmail = Email
    ModAuth.LoginOperation = e_operation.ForgotPassword
    Call connectToLoginServer
End Sub

Public Sub RequestNewPassword(ByVal Email As String, ByVal newPassword As String, ByVal code As String)
    CuentaEmail = Email
    ValidationCode = code
    CuentaPassword = newPassword
    ModAuth.LoginOperation = e_operation.ResetPassword
    Auth_state = e_state.RequestResetPassword
    Call connectToLoginServer
End Sub

Public Sub LoginCharacter(ByVal Name As String)
    On Error GoTo LogearPersonaje_Err
    userName = Name
    If Connected Then
        frmMain.ShowFPS.enabled = (FPSFLAG = 1)
        frmMain.fps.visible = (FPSFLAG = 1)
    End If
    #If PYMMO = 0 Then
        Call Protocol_Writes.WriteLoginExistingChar
    #End If
    #If PYMMO = 1 Then
        Call modNetwork.Connect(IPdelServidor, PuertoDelServidor)
        ModAuth.LoginOperation = e_operation.Authenticate
        Call LoginOrConnect(E_MODO.Normal)
        
        #If REMOTE_CLOSE = 0 Then
            frmConnecting.Show False, frmConnect
        #Else
            Call SaveStringInFile("LoginCharacter: " & Name, "remote_debug.txt")
        #End If
    #End If
    Exit Sub
LogearPersonaje_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModLogin.LogearPersonaje", Erl)
    Resume Next
End Sub

Public Sub ShowLogin()
    frmConnect.Show
    Dim patchNotes As String
    patchNotes = GetPatchNotes()
    If Not patchNotes = "" Then
        frmPatchNotes.SetNotes (patchNotes)
        frmPatchNotes.Show , frmConnect
    Else
        FrmLogear.Show , frmConnect
    End If
End Sub

Public Sub ShowScharSelection()
    Call connectToLoginServer
End Sub

Public Sub CreateCharacter(ByVal Name As String, ByVal Race As Integer, ByVal Gender As Integer, ByVal Class As Integer, ByVal Head As Integer, ByVal HomeCity As Integer)
    userName = Name
    UserStats.Raza = Race
    UserStats.Sexo = Gender
    UserStats.Clase = Class
    MiCabeza = Head
    UserStats.Hogar = HomeCity
    #If PYMMO = 1 Then
        Call modNetwork.Connect(IPdelServidor, PuertoDelServidor)
        Call LoginOrConnect(E_MODO.CrearNuevoPj)
    #Else
        Call Protocol_Writes.WriteLoginNewChar(userName, UserStats.Raza, UserStats.Sexo, UserStats.Clase, MiCabeza, UserStats.Hogar)
    #End If
End Sub

Public Sub RequestDeleteCharacter()
    #If PYMMO = 1 Then
        ModAuth.LoginOperation = e_operation.deletechar
        Call connectToLoginServer
    #Else
        Call DisplayError("Unsoported on localhost.", "")
    #End If
End Sub

Public Sub DeleteCharRequestCode()
    MsgBox ("Se ha enviado un código de verificación al mail proporcionado.")
End Sub

Public Sub TransferChar(ByVal Name As String, ByVal DestinationAccunt As String)
    TransferCharNewOwner = DestinationAccunt
    TransferCharname = Name
    Debug.Assert Len(TransferCharNewOwner) > 0
    Debug.Assert Len(Name) > 0
    ModAuth.LoginOperation = e_operation.transfercharacter
    Call connectToLoginServer
End Sub




Public Sub OnClientDisconnect(ByVal Error As Long)
    On Error GoTo OnClientDisconnect_Err
    
    Const WSAETIMEDOUT       As Long = 10060
    Const WSAECONNREFUSED    As Long = 10061
    Const WSAECONNRESET      As Long = 10054
    Const WSAENETUNREACH     As Long = 10051
    Const WSAEHOSTUNREACH    As Long = 10065

    #If REMOTE_CLOSE = 0 Then

        frmConnect.MousePointer = 1
        frmMain.ShowFPS.enabled = False

        Select Case Error

            '----------------------------------------------------------
            ' SERVER REFUSED CONNECTION (immediate failure)
            '----------------------------------------------------------
            Case WSAECONNREFUSED
                If frmConnect.visible Then
                    Call DisplayError(JsonLanguage.Item("MENSAJE_ERROR_CONEXION"), "connection-failure")
                Else
                    Call DisplayError(JsonLanguage.Item("MENSAJE_ERROR_CONEXION_SERVIDOR"), "connection-failure")
                End If
                Connected = False
                Exit Sub


            '----------------------------------------------------------
            ' TIMEOUT – retry with next Game Server
            '----------------------------------------------------------
            Case WSAETIMEDOUT
                Dim nextGame As Integer
                nextGame = GetNextGameServer(CurrentGameServerIndex)

                If nextGame = -1 Or nextGame = CurrentGameServerIndex Then
                    frmDebug.add_text_tracebox "No hay más Game Servers para reintentar (timeout)."
                    Call DisplayError(JsonLanguage.Item("MENSAJE_ERROR_CONEXION_SERVIDOR"), "connection-failure")
                    If Connected Then Call HandleDisconnect
                    Exit Sub
                End If

                ' Switch to next game server
                CurrentGameServerIndex = nextGame
                IPdelServidor = GameServerIP(CurrentGameServerIndex)
                PuertoDelServidor = GameServerPort(CurrentGameServerIndex)

                frmDebug.add_text_tracebox _
                    "Timeout ? cambiando a Game Server #" & CurrentGameServerIndex & _
                    " -> " & IPdelServidor & ":" & PuertoDelServidor

                Call modNetwork.Reconnect(IPdelServidor, PuertoDelServidor)
                Call Login
                Exit Sub


            '----------------------------------------------------------
            ' ANY NONZERO ERROR except normal closure ? unexpected
            '----------------------------------------------------------
            Case Is <> 0, 2
                Call DisplayError(JsonLanguage.Item("MENSAJE_ERROR_CONEXION_SERVIDOR"), "connection-failure")

                If frmConnect.visible Then
                    Connected = False
                ElseIf Connected Then
                    Call HandleDisconnect
                End If

                Exit Sub


            '----------------------------------------------------------
            ' NORMAL CLOSURE OR LESS SERIOUS ERROR
            '----------------------------------------------------------
            Case Else
                If Error <> 0 Then
                    Call RegistrarError(Error, "Conexion cerrada", "OnClientDisconnect")
                End If

                If frmConnect.visible Then
                    Connected = False
                ElseIf Connected Then
                    Call HandleDisconnect
                End If

                If Not GetRemoteError And Error > 0 Then
                    Call DisplayError(JsonLanguage.Item("MENSAJE_CONEXION_CERRADA"), "connection-closed")
                End If

                Exit Sub

        End Select

    #Else

        frmDebug.add_text_tracebox "OnClientDisconnect " & Error
        Call SaveStringInFile("OnClientDisconnect " & Error, "remote_debug.txt")
        prgRun = False
    #End If
    Exit Sub
OnClientDisconnect_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModLogin.OnClientDisconnect", Erl)
    Resume Next
End Sub
