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
    
Public Sub DoLogin(ByVal Account As String, ByVal Password As String, ByVal storeCredentials As Boolean)
    On Error Goto DoLogin_Err
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
    Exit Sub
DoLogin_Err:
    Call TraceError(Err.Number, Err.Description, "ModLogin.DoLogin", Erl)
End Sub

Public Sub SetActiveServer(ByVal IP As String, ByVal port As String, Optional IgnoreHardcode As Boolean = False)
    On Error Goto SetActiveServer_Err
    ServerIndex = IP & ":" & port
    IPdelServidor = IP
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
    frmDebug.add_text_tracebox "Using Login Server " & IPdelServidorLogin & ":" & PuertoDelServidorLogin
    frmDebug.add_text_tracebox "Using Game Server " & IPdelServidor & ":" & PuertoDelServidor
    Exit Sub
SetActiveServer_Err:
    Call TraceError(Err.Number, Err.Description, "ModLogin.SetActiveServer", Erl)
End Sub

Public Sub SetActiveEnvironment(ByVal environment As String)
    On Error Goto SetActiveEnvironment_Err
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
    environment = "Production"
#End If
    Dim loginServers As Integer
    loginServers = Val(ServerSettings.GetValue(environment, "LoginCount"))
    ServerIpCount = Val(ServerSettings.GetValue(environment, "ServerCount"))
    Dim loginOpt, serverOpt, k As Integer
    For k = 1 To 100
        serverOpt = RandomNumber(1, ServerIpCount)
    Next k
    For k = 1 To 100
        loginOpt = RandomNumber(1, loginServers)
    Next k
    IPdelServidor = ServerSettings.GetValue(environment, "ServerIp" & serverOpt)
    PuertoDelServidor = ServerSettings.GetValue(environment, "PortPort" & serverOpt)
    IPdelServidorLogin = ServerSettings.GetValue(environment, "LoginIp" & loginOpt)
    PuertoDelServidorLogin = ServerSettings.GetValue(environment, "LoginPort" & loginOpt)
    frmDebug.add_text_tracebox "Using Login Server " & IPdelServidorLogin & ":" & PuertoDelServidorLogin
    frmDebug.add_text_tracebox "Using Game Server " & IPdelServidor & ":" & PuertoDelServidor
    Exit Sub
SetActiveEnvironment_Err:
    Call TraceError(Err.Number, Err.Description, "ModLogin.SetActiveEnvironment", Erl)
End Sub

Public Sub CreateAccount(ByVal Name As String, ByVal Surname As String, ByVal Email As String, ByVal Password As String)
    On Error Goto CreateAccount_Err
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
    Exit Sub
CreateAccount_Err:
    Call TraceError(Err.Number, Err.Description, "ModLogin.CreateAccount", Erl)
End Sub

Public Sub LoadCharacterSelectionScreen()
    On Error Goto LoadCharacterSelectionScreen_Err
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
    Exit Sub
LoadCharacterSelectionScreen_Err:
    Call TraceError(Err.Number, Err.Description, "ModLogin.LoadCharacterSelectionScreen", Erl)
End Sub

Public Sub GoToLogIn()
    On Error Goto GoToLogIn_Err
    g_game_state.State = e_state_connect_screen
    Exit Sub
GoToLogIn_Err:
    Call TraceError(Err.Number, Err.Description, "ModLogin.GoToLogIn", Erl)
End Sub

Public Sub LogOut()
    On Error Goto LogOut_Err
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
    Exit Sub
LogOut_Err:
    Call TraceError(Err.Number, Err.Description, "ModLogin.LogOut", Erl)
End Sub

Public Sub ResendValidationCode(ByVal Email As String)
    On Error Goto ResendValidationCode_Err
    CuentaEmail = Email
    ModAuth.LoginOperation = e_operation.RequestVerificationCode
    Call connectToLoginServer
    Exit Sub
ResendValidationCode_Err:
    Call TraceError(Err.Number, Err.Description, "ModLogin.ResendValidationCode", Erl)
End Sub

Public Sub ValidateCode(ByVal Email As String, ByVal code As String)
    On Error Goto ValidateCode_Err
    CuentaEmail = Email
    ValidationCode = code
    ModAuth.LoginOperation = e_operation.ValidateAccount
    Call connectToLoginServer
    Exit Sub
ValidateCode_Err:
    Call TraceError(Err.Number, Err.Description, "ModLogin.ValidateCode", Erl)
End Sub

Public Sub RequestPasswordReset(ByVal Email As String)
    On Error Goto RequestPasswordReset_Err
    CuentaEmail = Email
    ModAuth.LoginOperation = e_operation.ForgotPassword
    Call connectToLoginServer
    Exit Sub
RequestPasswordReset_Err:
    Call TraceError(Err.Number, Err.Description, "ModLogin.RequestPasswordReset", Erl)
End Sub

Public Sub RequestNewPassword(ByVal Email As String, ByVal newPassword As String, ByVal code As String)
    On Error Goto RequestNewPassword_Err
    CuentaEmail = Email
    ValidationCode = code
    CuentaPassword = newPassword
    ModAuth.LoginOperation = e_operation.ResetPassword
    Auth_state = e_state.RequestResetPassword
    Call connectToLoginServer
    Exit Sub
RequestNewPassword_Err:
    Call TraceError(Err.Number, Err.Description, "ModLogin.RequestNewPassword", Erl)
End Sub

Public Sub LoginCharacter(ByVal Name As String)
    On Error Goto LoginCharacter_Err
On Error GoTo LogearPersonaje_Err
    userName = Name
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
    Exit Sub
LoginCharacter_Err:
    Call TraceError(Err.Number, Err.Description, "ModLogin.LoginCharacter", Erl)
End Sub

Public Sub ShowLogin()
    On Error Goto ShowLogin_Err
        frmConnect.Show
        Dim patchNotes As String
        patchNotes = GetPatchNotes()
        If Not patchNotes = "" Then
            frmPatchNotes.SetNotes (patchNotes)
            frmPatchNotes.Show , frmConnect
        Else
            FrmLogear.Show , frmConnect
        End If
    Exit Sub
ShowLogin_Err:
    Call TraceError(Err.Number, Err.Description, "ModLogin.ShowLogin", Erl)
End Sub

Public Sub ShowScharSelection()
    On Error Goto ShowScharSelection_Err
        Call connectToLoginServer
    Exit Sub
ShowScharSelection_Err:
    Call TraceError(Err.Number, Err.Description, "ModLogin.ShowScharSelection", Erl)
End Sub

Public Sub CreateCharacter(ByVal Name As String, ByVal Race As Integer, ByVal Gender As Integer, ByVal Class As Integer, ByVal Head As Integer, ByVal HomeCity As Integer)
    On Error Goto CreateCharacter_Err
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
    
    Exit Sub
CreateCharacter_Err:
    Call TraceError(Err.Number, Err.Description, "ModLogin.CreateCharacter", Erl)
End Sub

Public Sub RequestDeleteCharacter()
    On Error Goto RequestDeleteCharacter_Err
#If PYMMO = 1 Then
    ModAuth.LoginOperation = e_operation.deletechar
    Call connectToLoginServer
#Else
    Call DisplayError("Unsoported on localhost.", "")
#End If
    Exit Sub
RequestDeleteCharacter_Err:
    Call TraceError(Err.Number, Err.Description, "ModLogin.RequestDeleteCharacter", Erl)
End Sub

Public Sub DeleteCharRequestCode()
    On Error Goto DeleteCharRequestCode_Err

        MsgBox ("Se ha enviado un código de verificación al mail proporcionado.")

    Exit Sub
DeleteCharRequestCode_Err:
    Call TraceError(Err.Number, Err.Description, "ModLogin.DeleteCharRequestCode", Erl)
End Sub

Public Sub TransferChar(ByVal Name As String, ByVal DestinationAccunt As String)
    On Error Goto TransferChar_Err
    TransferCharNewOwner = DestinationAccunt
    TransferCharname = Name
    Debug.Assert Len(TransferCharNewOwner) > 0
    Debug.Assert Len(Name) > 0
    ModAuth.LoginOperation = e_operation.transfercharacter
    Call connectToLoginServer
    Exit Sub
TransferChar_Err:
    Call TraceError(Err.Number, Err.Description, "ModLogin.TransferChar", Erl)
End Sub

Public Sub OnClientDisconnect(ByVal Error As Long)
    On Error Goto OnClientDisconnect_Err
    
On Error GoTo OnClientDisconnect_Err

#If REMOTE_CLOSE = 0 Then
    If (Error = 10061) Then
        If frmConnect.visible Then
            Call DisplayError(JsonLanguage.Item("MENSAJE_ERROR_CONEXION"), "connection-failure")
        Else
            Call DisplayError(JsonLanguage.Item("MENSAJE_ERROR_CONEXION_SERVIDOR"), "connection-failure")
        End If
    Else
        frmConnect.MousePointer = 1
        frmMain.ShowFPS.enabled = False
        If (Error <> 0 And Error <> 2) Then
            Call DisplayError(JsonLanguage.Item("MENSAJE_ERROR_CONEXION_SERVIDOR"), "connection-failure")
            
            If frmConnect.visible Then
                Connected = False
            Else
                If (Connected) Then
                    Call HandleDisconnect
                End If
            End If
        Else
            If Error <> 0 Then
                Call RegistrarError(Error, "Conexion cerrada", "OnClientDisconnect")
            End If
            If frmConnect.visible Then
                Connected = False
            Else
                If (Connected) Then
                    Call HandleDisconnect
                End If
            End If
            If Not GetRemoteError And Error > 0 Then
                Call DisplayError(JsonLanguage.Item("MENSAJE_CONEXION_CERRADA"), "connection-closed")
            End If
        End If
    End If
#Else
    frmDebug.add_text_tracebox "OnClientDisconnect " & Error
    Call SaveStringInFile("OnClientDisconnect " & Error, "remote_debug.txt")
    prgRun = False
#End If
    Exit Sub
OnClientDisconnect_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModLogin.OnClientDisconnect", Erl)
    Resume Next
    Exit Sub
OnClientDisconnect_Err:
    Call TraceError(Err.Number, Err.Description, "ModLogin.OnClientDisconnect", Erl)
End Sub
