Attribute VB_Name = "Mod_TCP"
'RevolucionAo 1.0
'Pablo Mercavides

Option Explicit

Public Const SOL_SOCKET As Long = &HFFFF&
Public Const TCP_NODELAY As Long = &H1
Public Declare Function setsockopt Lib "ws2_32" ( _
    ByVal s As Long, _
    ByVal level As Long, _
    ByVal optname As Long, _
    ByRef optval As Any, _
    ByVal optlen As Long _
    ) As Long

Public Warping        As Boolean

Public LlegaronSkills As Boolean
Public LlegaronStats  As Boolean
Public LlegaronAtrib  As Boolean

Public Sub SetSocketNoDelay(ByVal Handle As Long, ByVal Status As Boolean)
    Dim TcpNoDelay As Long
    TcpNoDelay = IIf(Status, 1, 0)
    setsockopt Handle, SOL_SOCKET, TCP_NODELAY, TcpNoDelay, LenB(TcpNoDelay)
End Sub

Public Function PuedoQuitarFoco() As Boolean
    
    On Error GoTo PuedoQuitarFoco_Err
    
    PuedoQuitarFoco = True

    'PuedoQuitarFoco = Not frmEstadisticas.Visible And _
    '                 Not frmGuildAdm.Visible And _
    '                 Not frmGuildDetails.Visible And _
    '                 Not frmGuildBrief.Visible And _
    '                 Not frmGuildFoundation.Visible And _
    '                 Not frmGuildLeader.Visible And _
    '                 Not frmCharInfo.Visible And _
    '                 Not frmGuildNews.Visible And _
    '                 Not frmGuildSol.Visible And _
    '                 Not frmCommet.Visible And _
    '                 Not frmPeaceProp.Visible
    '
    
    Exit Function

PuedoQuitarFoco_Err:
    Call RegistrarError(Err.number, Err.Description, "Mod_TCP.PuedoQuitarFoco", Erl)
    Resume Next
    
End Function

Sub LoginOrConnect(ByVal Modo As E_MODO)
    EstadoLogin = Modo
    
    If (Connected) Then
        Call Login
    Else
        Call modNetwork.Connect(IPdelServidor, PuertoDelServidor)
    End If
End Sub

Sub Login()
    
    On Error GoTo Login_Err
    

    If EstadoLogin = E_MODO.Normal Then
        Call WriteLoginExistingChar
    ElseIf EstadoLogin = E_MODO.CrearNuevoPj Then
        Call WriteLoginNewChar
    ElseIf EstadoLogin = E_MODO.Dados Then
        Call WriteThrowDice
        
        If QueRender <> 3 Then
            UserMap = 37
            AlphaNiebla = 3
            'EntradaY = 90
            'EntradaX = 90
            CPHeading = 3
            CPEquipado = True
            Call SwitchMap(UserMap)
            ' frmCrearPersonaje.Show
            QueRender = 3
            
            Call IniciarCrearPj
            '      Sound.NextMusic = 3
            ' Sound.Fading = 350
            'FrmCuenta.Visible = False
            frmConnect.txtNombre.Visible = True
            frmConnect.txtNombre.SetFocus

            Call Sound.Sound_Play(SND_DICE)

        End If

    End If
    
    Exit Sub

Login_Err:
    Call RegistrarError(Err.number, Err.Description, "Mod_TCP.Login", Erl)
    Resume Next
    
End Sub
