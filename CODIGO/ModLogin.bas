Attribute VB_Name = "ModLogin"
Option Explicit

Dim ServerSettings As clsIniManager
    
Public Sub DoLogin(ByVal account As String, ByVal password As String, ByVal storeCredentials As Boolean)
    On Error GoTo DoLogin_Err
    
    If IntervaloPermiteConectar Then
        CuentaEmail = account
        CuentaPassword = password

        If storeCredentials Then
            CuentaRecordada.nombre = CuentaEmail
            CuentaRecordada.password = CuentaPassword
            
            Call GuardarCuenta(CuentaEmail, CuentaPassword)
        Else
            ' Reseteamos los datos de cuenta guardados
            Call GuardarCuenta(vbNullString, vbNullString)
        End If

        If CheckUserDataLoged() = True Then
            ModAuth.LoginOperation = e_operation.Authenticate
            Call LoginOrConnect(E_MODO.IngresandoConCuenta)
        End If
        Call SaveRAOInit

    End If

    Exit Sub

DoLogin_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModLogin.DoLogin", Erl)
    Resume Next
End Sub

Public Sub SetActiveServer(ByVal IP As String, ByVal port As String, Optional IgnoreHardcode As Boolean = False)
    ServerIndex = ip & ":" & port
    IPdelServidor = ip
    PuertoDelServidor = port
    
    #If PYMMO = 0 Or DEBUGGING = 1 Then
            Call SaveSetting("INIT", "ServerIndex", IPdelServidor & ":" & PuertoDelServidor)
    #End If
    If Not IgnoreHardcode Then
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
    End If
    Debug.Print "Using Login Server " & IPdelServidorLogin & ":" & PuertoDelServidorLogin
    Debug.Print "Using Game Server " & IPdelServidor & ":" & PuertoDelServidor
End Sub

Public Sub SetActiveEnvironment(ByVal environment As String)
    If ServerSettings Is Nothing Then
        Dim RemotesPath As String
        Set ServerSettings = New clsIniManager
#If Compresion = 1 Then
        If Not Extract_File(Scripts, App.path & "\..\Recursos\OUTPUT\", "Remotes.ini", Windows_Temp_Dir, ResourcesPassword, False) Then
            Err.Description = "¡No se puede cargar el archivo de Remotes.ini!"
            MsgBox Err.Description
            Exit Sub
        End If
        RemotesPath = Windows_Temp_Dir & "Remotes.ini"
#Else
        RemotesPath = App.path & "\..\Recursos\init\Remotes.ini"
#End If
        Debug.Assert FileExist(RemotesPath, vbNormal)
        Call ServerSettings.Initialize(RemotesPath)
        
    End If
#If Developer = 0 And DEBUGGING = 0 Then
    Environment = "Production"
#End If
    Dim loginServers As Integer
    Dim serverCount As Integer
    loginServers = Val(ServerSettings.GetValue(environment, "LoginCount"))
    serverCount = Val(ServerSettings.GetValue(environment, "ServerCount"))
    Dim loginOpt, serverOpt, k As Integer
    For k = 1 To 100
        serverOpt = RandomNumber(1, serverCount)
    Next k
    For k = 1 To 100
        loginOpt = RandomNumber(1, loginServers)
    Next k
    IPdelServidor = ServerSettings.GetValue(environment, "ServerIp" & serverOpt)
    PuertoDelServidor = ServerSettings.GetValue(environment, "PortPort" & serverOpt)
    IPdelServidorLogin = ServerSettings.GetValue(environment, "LoginIp" & serverOpt)
    PuertoDelServidorLogin = ServerSettings.GetValue(environment, "LoginPort" & serverOpt)
    Debug.Print "Using Login Server " & IPdelServidorLogin & ":" & PuertoDelServidorLogin
    Debug.Print "Using Game Server " & IPdelServidor & ":" & PuertoDelServidor
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
    If BabelUI.BabelInitialized Then
        Call SendLoginCharacters(Pjs, CantidadDePersonajesEnCuenta)
        If g_game_state.state <> e_state_createchar_screen Then
            Call BabelUI.SetActiveScreen("character-selection")
            g_game_state.state = e_state_account_screen
        Else
            Call BabelUI.SetActiveScreen("create-character")
        End If
    Else
        frmConnect.visible = True
        g_game_state.state = e_state_account_screen
    End If
    
    SugerenciaAMostrar = RandomNumber(1, NumSug)
    Call Sound.Sound_Play(192)
    Call Sound.Sound_Stop(SND_LLUVIAIN)
    Call Graficos_Particulas.Particle_Group_Remove_All
    Call Graficos_Particulas.Engine_Select_Particle_Set(203)
    ParticleLluviaDorada = Graficos_Particulas.General_Particle_Create(208, -1, -1)
    If FrmLogear.visible Then
        Unload FrmLogear
    End If
    
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
    g_game_state.state = e_state_connect_screen
    If BabelUI.BabelInitialized Then
        Call BabelUI.SetActiveScreen("login")
    End If
End Sub

Public Sub LogOut()
    Debug.Print "Vuelvo al login, debería borrar el token"
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
        Pjs(i).nivel = 0
        Pjs(i).nombre = ""
        Pjs(i).Clase = 0
        Pjs(i).Criminal = 0
        Pjs(i).NameMapa = ""
    Next i
    General_Set_Connect
End Sub

Public Sub ResendValidationCode(ByVal email As String)
    CuentaEmail = email
    ModAuth.LoginOperation = e_operation.RequestVerificationCode
    Call connectToLoginServer
End Sub

Public Sub ValidateCode(ByVal email As String, ByVal code As String)
    CuentaEmail = email
    validationCode = code
    ModAuth.LoginOperation = e_operation.ValidateAccount
    Call connectToLoginServer
End Sub

Public Sub RequestPasswordReset(ByVal email As String)
    CuentaEmail = email
    ModAuth.LoginOperation = e_operation.ForgotPassword
    Call connectToLoginServer
End Sub

Public Sub RequestNewPassword(ByVal email As String, ByVal newPassword As String, ByVal code As String)
    CuentaEmail = email
    validationCode = code
    CuentaPassword = newPassword
    ModAuth.LoginOperation = e_operation.ResetPassword
    Auth_state = e_state.RequestResetPassword
    Call connectToLoginServer
End Sub

Public Sub LoginCharacter(ByVal Name As String)
On Error GoTo LogearPersonaje_Err
    username = Name
    If Connected Then
        frmMain.ShowFPS.enabled = True
    End If
#If PYMMO = 0 Then
    Call Protocol_Writes.WriteLoginExistingChar
#End If
#If PYMMO = 1 Then
    Call modNetwork.Connect(IPdelServidor, PuertoDelServidor)
    ModAuth.LoginOperation = e_operation.Authenticate
    Call LoginOrConnect(E_MODO.Normal)
#End If
    Exit Sub
LogearPersonaje_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModLogin.LogearPersonaje", Erl)
    Resume Next
End Sub

Public Sub ShowLogin()
    If UseBabelUI Then
        frmBabelLogin.Show
        BabelUI.SetActiveScreen ("login")
    Else
        frmConnect.Show
        Dim patchNotes As String
        patchNotes = GetPatchNotes()
        If Not patchNotes = "" Then
            frmPatchNotes.SetNotes (patchNotes)
            frmPatchNotes.Show , frmConnect
        Else
            FrmLogear.Show , frmConnect
        End If
    End If
End Sub

Public Sub ShowScharSelection()
    If UseBabelUI Then
        frmBabelLogin.Show
        BabelUI.SetActiveScreen ("charcter-selection")
    Else
        Call connectToLoginServer
    End If
End Sub

Public Sub CreateCharacter(ByVal name As String, ByVal Race As Integer, ByVal Gender As Integer, ByVal Class As Integer, ByVal Head As Integer, ByVal HomeCity As Integer)
    userName = name
    UserRaza = Race
    UserSexo = Gender
    UserClase = Class
    MiCabeza = Head
    UserHogar = HomeCity
#If PYMMO = 1 Then
    Call modNetwork.Connect(IPdelServidor, PuertoDelServidor)
    Call LoginOrConnect(E_MODO.CrearNuevoPj)
#Else
    Call Protocol_Writes.WriteLoginNewChar(userName, UserRaza, UserSexo, UserClase, MiCabeza, UserHogar)
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
    If UseBabelUI Then
        Call RequestDeleteCode
    Else
        MsgBox ("Se ha enviado un código de verificación al mail proporcionado.")
    End If
End Sub

Public Sub TransferChar(ByVal Name As String, ByVal DestinationAccunt As String)
    TransferCharNewOwner = DestinationAccunt
    TransferCharname = name
    Debug.Assert Len(TransferCharNewOwner) > 0
    Debug.Assert Len(Name) > 0
    ModAuth.LoginOperation = e_operation.TransferCharacter
    Call connectToLoginServer
End Sub
