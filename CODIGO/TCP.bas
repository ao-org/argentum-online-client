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
    Call RegistrarError(Err.Number, Err.Description, "Mod_TCP.PuedoQuitarFoco", Erl)
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
    
    Select Case EstadoLogin

        Case E_MODO.CrearNuevoPj
            Call WriteLoginNewChar
            
        Case E_MODO.Dados
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
        
        Case E_MODO.IngresandoConCuenta
            Call WriteIngresandoConCuenta
            
        Case E_MODO.BorrandoPJ
            Call WriteBorrandoPJ
    End Select

    Exit Sub

Login_Err:
    Call RegistrarError(Err.Number, Err.Description, "Mod_TCP.Login", Erl)
    Resume Next
    
End Sub
