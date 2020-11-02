Attribute VB_Name = "Mod_TCP"
'RevolucionAo 1.0
'Pablo Mercavides


Option Explicit
Public Warping As Boolean
Public LlegaronSkills As Boolean
Public LlegaronAtrib As Boolean


Public Function PuedoQuitarFoco() As Boolean
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
End Function

Sub Login()
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
            Call SwitchMapIAO(UserMap)
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
    
    DoEvents
    
    Call FlushBuffer
End Sub
