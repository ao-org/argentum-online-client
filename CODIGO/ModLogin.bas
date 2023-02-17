Attribute VB_Name = "ModLogin"
Option Explicit

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

Public Sub SetActiveServer(ByVal ip As String, ByVal port As String)
    ServerIndex = ip & ":" & port
    IPdelServidor = ip
    PuertoDelServidor = port
    
    #If PYMMO = 0 Or DEBUGGING = 1 Then
            Call SaveSetting("INIT", "ServerIndex", IPdelServidor & ":" & PuertoDelServidor)
    #End If
    
    #If PYMMO = 1 Then
        #If DEVELOPER = 1 Then
            IPdelServidorLogin = "127.0.0.1"
            PuertoDelServidorLogin = 4000
            IPdelServidor = IPdelServidorLogin
            PuertoDelServidor = 7667
        #Else
            'Production and staging use this path
            #If DEBUGGING = 0 Then
                Call SetDefaultServer
            #Else
                IPdelServidorLogin = "45.235.98.31"
                PuertoDelServidorLogin = 11814
            #End If
        #End If
    #End If
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
