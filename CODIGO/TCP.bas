Attribute VB_Name = "Mod_TCP"
'RevolucionAo 1.0
'Pablo Mercavides

Option Explicit

Public Warping        As Boolean

Public LlegaronSkills As Boolean
Public LlegaronStats  As Boolean
Public LlegaronAtrib  As Boolean

Public Function PuedoQuitarFoco() As Boolean
    
    On Error GoTo PuedoQuitarFoco_Err
    
    PuedoQuitarFoco = True
    
    Exit Function

PuedoQuitarFoco_Err:
    Call RegistrarError(Err.number, Err.Description, "Mod_TCP.PuedoQuitarFoco", Erl)
    Resume Next
    
End Function
#If PYMMO = 1 Then
Sub LoginOrConnect(ByVal Modo As E_MODO)
    EstadoLogin = Modo
    
    If Auth_state = e_state.AccountLogged Then
        'Call modNetwork.Connect(IPdelServidor, PuertoDelServidor)
        Call Login
    Else
        Call connectToLoginServer
    End If
  
End Sub
#ElseIf PYMMO = 0 Then
Sub LoginOrConnect(ByVal Modo As E_MODO)
    
    EstadoLogin = Modo
    
    If (Not modNetwork.IsConnected) Then
        Call modNetwork.Connect(IPdelServidor, PuertoDelServidor)
    Else
        Call Login
    End If

End Sub
#End If

Sub Login()
    
    On Error GoTo Login_Err
    
    Select Case EstadoLogin
    
        Case E_MODO.Normal
            Call WriteLoginExistingChar
        
        Case E_MODO.CrearNuevoPj
            Call WriteLoginNewChar
#If PYMMO = 0 Then
        Case E_MODO.IngresandoConCuenta
            Call WriteLoginAccount
        
        Case E_MODO.CreandoCuenta
            Call WriteCreateAccount
#End If
    End Select

    Exit Sub

Login_Err:
    Call RegistrarError(Err.number, Err.Description, "Mod_TCP.Login", Erl)
    Resume Next
    
End Sub
