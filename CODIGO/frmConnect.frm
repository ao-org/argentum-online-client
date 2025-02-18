VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmConnect 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Argentum20"
   ClientHeight    =   11520
   ClientLeft      =   15
   ClientTop       =   105
   ClientWidth     =   15360
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock AuthSocket 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "45.235.99.71"
      RemotePort      =   4004
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   12600
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   14640
      Top             =   240
   End
   Begin VB.PictureBox render 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   11520
      Left            =   0
      ScaleHeight     =   768
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1024
      TabIndex        =   0
      Top             =   0
      Width           =   15360
      Begin VB.TextBox txtNombre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   6240
         MaxLength       =   18
         TabIndex        =   1
         Top             =   3360
         Visible         =   0   'False
         Width           =   2130
      End
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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


Private SelectedCharIndex As Byte

#If DIRECT_PLAY = 1 Then
'We need to implement the Event model for DirectPlay so we can receive callbacks
Implements DirectPlay8Event
Implements DirectPlay8LobbyEvent
Public mfGotEvent As Boolean
Public mfConnectComplete As Boolean

'We will handle all of the msgs here, and report them all back to the callback sub
'in case the caller cares what's going on
Private Sub DirectPlay8Event_AddRemovePlayerGroup(ByVal lMsgID As Long, ByVal lPlayerID As Long, ByVal lGroupID As Long, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
    
End Sub

Private Sub DirectPlay8Event_AppDesc(fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
  
End Sub

Private Sub DirectPlay8Event_AsyncOpComplete(dpnotify As DxVBLibA.DPNMSG_ASYNC_OP_COMPLETE, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlay8Event_ConnectComplete(dpnotify As DxVBLibA.DPNMSG_CONNECT_COMPLETE, fRejectMsg As Boolean)
    mfGotEvent = True
    Debug.Print " DirectPlay8Event_ConnectComplete"
    If dpnotify.hResultCode = DPNERR_SESSIONFULL Then 'Already too many people joined up
        MsgBox "The maximum number of people allowed in this session have already joined.  Please choose a different session or create your own.", vbOKOnly Or vbInformation, "Full"
    Else
        'We got our connect complete event
        mfConnectComplete = True
        modNetwork.OnClientConnect dpnotify, fRejectMsg
    End If
End Sub

Private Sub DirectPlay8Event_CreateGroup(ByVal lGroupID As Long, ByVal lOwnerID As Long, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlay8Event_CreatePlayer(ByVal lPlayerID As Long, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
    Debug.Print "DirectPlay8Event_CreatePlayer " & lPlayerID
End Sub

Private Sub DirectPlay8Event_DestroyGroup(ByVal lGroupID As Long, ByVal lReason As Long, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlay8Event_DestroyPlayer(ByVal lPlayerID As Long, ByVal lReason As Long, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
    Debug.Print "DirectPlay8Event_DestroyPlayer " & lPlayerID
End Sub

Private Sub DirectPlay8Event_EnumHostsQuery(dpnotify As DxVBLibA.DPNMSG_ENUM_HOSTS_QUERY, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlay8Event_EnumHostsResponse(dpnotify As DxVBLibA.DPNMSG_ENUM_HOSTS_RESPONSE, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlay8Event_HostMigrate(ByVal lNewHostID As Long, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlay8Event_IndicateConnect(dpnotify As DxVBLibA.DPNMSG_INDICATE_CONNECT, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlay8Event_IndicatedConnectAborted(fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlay8Event_InfoNotify(ByVal lMsgID As Long, ByVal lNotifyID As Long, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
    Debug.Print "DirectPlay8Event_InfoNotify"
End Sub

Private Sub DirectPlay8Event_Receive(dpnotify As DxVBLibA.DPNMSG_RECEIVE, fRejectMsg As Boolean)
    Call modNetwork.Receive(dpnotify, fRejectMsg)
End Sub

Private Sub DirectPlay8Event_SendComplete(dpnotify As DxVBLibA.DPNMSG_SEND_COMPLETE, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlay8Event_TerminateSession(dpnotify As DxVBLibA.DPNMSG_TERMINATE_SESSION, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
    Debug.Print "DirectPlay8Event_TerminateSession"
    Call modNetwork.OnClientDisconnect(dpnotify, fRejectMsg)
End Sub

Private Sub DirectPlay8LobbyEvent_Connect(dlNotify As DxVBLibA.DPL_MESSAGE_CONNECT, fRejectMsg As Boolean)
   Exit Sub
ErrOut:
    Debug.Print "Error:" & CStr(Err.Number) & " - " & Err.Description
End Sub

Private Sub DirectPlay8LobbyEvent_ConnectionSettings(ConnectionSettings As DxVBLibA.DPL_MESSAGE_CONNECTION_SETTINGS)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlay8LobbyEvent_Disconnect(ByVal DisconnectID As Long, ByVal lReason As Long)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlay8LobbyEvent_Receive(dlNotify As DxVBLibA.DPL_MESSAGE_RECEIVE, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlay8LobbyEvent_SessionStatus(ByVal status As Long, ByVal lHandle As Long)
    'VB requires that we must implement *every* member of this interface
End Sub

#End If




Private Sub AuthSocket_Connect()
    If Not SessionOpened Then
        Call OpenSessionRequest
        Select Case LoginOperation
            Case e_operation.Authenticate
                Auth_state = e_state.RequestAccountLogin
            Case e_operation.SignUp
                Auth_state = e_state.RequestSignUp
            Case e_operation.ValidateAccount
                Auth_state = e_state.RequestValidateAccount
            Case e_operation.ForgotPassword
                Auth_state = e_state.RequestForgotPassword
            Case e_operation.ResetPassword
                Auth_state = e_state.RequestResetPassword
            Case e_operation.deletechar
                Auth_state = e_state.RequestDeleteChar
            Case e_operation.ConfirmDeleteChar
                Auth_state = e_state.ConfirmDeleteChar
            Case e_operation.RequestVerificationCode
                Auth_state = e_state.RequestVerificationCode
            Case e_operation.transfercharacter
                Auth_state = e_state.RequestTransferCharacter
                
        End Select
    End If
    
End Sub

Private Sub AuthSocket_DataArrival(ByVal BytesTotal As Long)
    ModAuth.AuthSocket_DataArrival BytesTotal
End Sub

Private Sub AuthSocket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
#If REMOTE_CLOSE = 0 Then

    Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_SERVIDOR_OFFLINE"), False, SessionOpened)
#Else
    Debug.Print "SERVIDOR OFFLINE"
#End If
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo Form_KeyDown_Err
    

    If KeyCode = vbKeyEscape Then
        prgRun = False
        End

    End If

    
    Exit Sub

Form_KeyDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmConnect.Form_KeyDown", Erl)
    Resume Next
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    If (Not FormParser Is Nothing) Then
    Call FormParser.Parse_Form(Me)
    End If
    
    EngineRun = False
        
    Timer2.enabled = True
    Timer1.enabled = True
    
    ' Seteamos el caption hay que poner 20 aniversario
    Me.Caption = App.title
    
    ' Removemos la barra de titulo pero conservando el caption para la barra de tareas
#If Developer = 0 Then
    Call Form_RemoveTitleBar(Me)
#End If
    Debug.Assert D3DWindow.BackBufferWidth <> 0
    Debug.Assert D3DWindow.BackBufferHeight <> 0
    Me.Width = D3DWindow.BackBufferWidth * screen.TwipsPerPixelX
    Me.Height = D3DWindow.BackBufferHeight * screen.TwipsPerPixelY
    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmConnect.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub render_DblClick()
On Error GoTo render_DblClick_Err
    Form_RemoveTitleBar Me

    Select Case g_game_state.State()

        Case e_state_account_screen
            
            If PJSeleccionado < 1 Then Exit Sub

            Call ao20audio.PlayWav(SND_CLICK)

            If IntervaloPermiteConectar Then
                Call LogearPersonaje(Pjs(PJSeleccionado).nombre)

            End If

        Case e_state_createchar_screen
        
    End Select

    
    Exit Sub

render_DblClick_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmConnect.render_DblClick", Erl)
    Resume Next
    
End Sub


#If PYMMO = 1 Then
Private Sub render_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo render_MouseUp_Err
    
    Select Case g_game_state.State()

        Case e_state_createchar_screen
            If x > 282 And x < 322 And y > 428 And y < 468 Then 'Boton heading
                If CPHeading + 1 >= 5 Then
                CPHeading = 1
                Else
                    CPHeading = CPHeading + 1
                End If
            End If
            
            
            
            If x > 412 And x < 446 And y > 427 And y < 470 Then 'Boton Equipar
                If CPHeading - 1 <= 0 Then
                CPHeading = 4
            Else
                CPHeading = CPHeading - 1
                End If
            End If
                    

            If x > 325 And x < 344 And y > 371 And y < 387 Then 'Boton izquierda cabezas
                If frmCrearPersonaje.Cabeza.ListCount = 0 Then Exit Sub
                If frmCrearPersonaje.Cabeza.ListIndex > 0 Then
                    frmCrearPersonaje.Cabeza.ListIndex = frmCrearPersonaje.Cabeza.ListIndex - 1

                End If

                If frmCrearPersonaje.Cabeza.ListIndex = 0 Then
                    frmCrearPersonaje.Cabeza.ListIndex = frmCrearPersonaje.Cabeza.ListCount - 1

                End If

            End If
    
            If x > 394 And x < 411 And y > 373 And y < 386 Then 'Boton Derecha cabezas
                If frmCrearPersonaje.Cabeza.ListCount = 0 Then Exit Sub
                If (frmCrearPersonaje.Cabeza.ListIndex + 1) <> frmCrearPersonaje.Cabeza.ListCount Then
                    frmCrearPersonaje.Cabeza.ListIndex = frmCrearPersonaje.Cabeza.ListIndex + 1

                End If

                If (frmCrearPersonaje.Cabeza.ListIndex + 1) = frmCrearPersonaje.Cabeza.ListCount Then
                    frmCrearPersonaje.Cabeza.ListIndex = 0

                End If

            End If
                       
            ' Clase inicial
            ' Shugar: Arreglo los botones para seleccionar la clase inicial.
                       
            If x > 540 And x < 554 And y > 278 And y < 291 Then 'Boton izquierda clase
                
                Call Rotacion_boton_atras_clase

            End If
            
            If x > 658 And x < 671 And y > 278 And y < 291 Then 'Boton Derecha Clase
                                
                Call Rotacion_boton_adelante_clase

            End If

                
            If x > 539 And x < 553 And y > 322 And y < 335 Then 'OK
                If frmCrearPersonaje.lstRaza.ListIndex < frmCrearPersonaje.lstRaza.ListCount - 1 Then
                    frmCrearPersonaje.lstRaza.ListIndex = frmCrearPersonaje.lstRaza.ListIndex + 1
                Else
                    frmCrearPersonaje.lstRaza.ListIndex = 0

                End If

            End If
            
            If x > 657 And x < 672 And y > 321 And y < 338 Then 'OK
                If frmCrearPersonaje.lstRaza.ListIndex - 1 < 0 Then
                    frmCrearPersonaje.lstRaza.ListIndex = frmCrearPersonaje.lstRaza.ListCount - 1
                Else
                    frmCrearPersonaje.lstRaza.ListIndex = frmCrearPersonaje.lstRaza.ListIndex - 1
                End If

            End If
            
            If x > 298 And x < 314 And y > 276 And y < 291 Then 'ok
    
                If frmCrearPersonaje.lstGenero.ListIndex < frmCrearPersonaje.lstGenero.ListCount - 1 Then
                    frmCrearPersonaje.lstGenero.ListIndex = frmCrearPersonaje.lstGenero.ListIndex + 1
                Else
                    frmCrearPersonaje.lstGenero.ListIndex = 0

                End If

            End If
            
            If x > 415 And x < 431 And y > 277 And y < 295 Then 'ok
                If frmCrearPersonaje.lstGenero.ListIndex - 1 < 0 Then
                    frmCrearPersonaje.lstGenero.ListIndex = frmCrearPersonaje.lstGenero.ListCount - 1
                Else
                    frmCrearPersonaje.lstGenero.ListIndex = frmCrearPersonaje.lstGenero.ListIndex - 1

                End If

            End If

            ' Hogar inicial
            ' Shugar: Arreglo los botones para seleccionar el hogar inicial.
        
            If x > 416 And x < 433 And y > 323 And y < 338 Then
  
                Call Rotacion_boton_adelante_ciudades
                
            End If
            
            If x > 297 And x < 314 And y > 321 And y < 340 Then
            
                Call Rotacion_boton_atras_ciudades
                
            End If
            
            
            If x >= 289 And x < 289 + 160 And y >= 525 And y < 525 + 37 Then 'Boton > Volver
                Call ao20audio.PlayWav(SND_CLICK)
                AlphaNiebla = 25
                EntradaY = 1
                EntradaX = 1
                frmConnect.txtNombre.visible = False
                g_game_state.State = e_state_account_screen
                Call Graficos_Particulas.Engine_Select_Particle_Set(203)
                ParticleLluviaDorada = General_Particle_Create(208, -1, -1)

            End If
            
            
            If x >= 532 And x < 532 + 160 And y >= 525 And y < 525 + 37 Then 'Boton > Crear
                Call ao20audio.PlayWav(SND_CLICK)

                Dim k As Object

                
                userName = frmConnect.txtNombre.Text
                
                Dim Error As String
                If Not ValidarNombre(userName, Error) Then
                    frmMensaje.msg.Caption = Error
                    frmMensaje.Show , Me
                    Exit Sub
                End If

                UserStats.Raza = frmCrearPersonaje.lstRaza.ListIndex + 1
                UserStats.Sexo = frmCrearPersonaje.lstGenero.ListIndex + 1
                UserStats.Clase = frmCrearPersonaje.lstProfesion.ListIndex + 1
                UserStats.Hogar = frmCrearPersonaje.lstHogar.ListIndex + 1
               
                If frmCrearPersonaje.CheckData() Then
                    UserPassword = CuentaPassword
                    StopCreandoCuenta = True
                    If Connected Then
                        frmMain.ShowFPS.enabled = True
                    End If
                    EstadoLogin = E_MODO.CrearNuevoPj
                    frmConnecting.Show
                    Call modNetwork.Connect(IPdelServidor, PuertoDelServidor)
                End If
                

            End If

            Exit Sub

        Case e_state_account_screen
            character_screen_action = e_action_nothing_to_do
            
            If (x > 256 And x < 414) And (y > 710 And y < 747) Then
                character_screen_action = e_action_create_character
            End If
            
            If (x > 14 And x < 112) And (y > 675 And y < 708) Then
                character_screen_action = e_action_delete_character
            End If
            
            If (x > 980 And x < 1000) And (y > 675 And y < 708) Then
                character_screen_action = e_action_transfer_character
            End If
            
            If (x > 19 And x < 48) And (y > 21 And y < 45) Then
                character_screen_action = e_action_logout_account

            End If
            
            If (x > 604 And x < 759) And (y > 711 And y < 745) Then
                character_screen_action = e_action_login_character

            End If
            
            If (x > 971 And x < 1001) And (y > 21 And y < 45) Then
                character_screen_action = e_action_close_game

            End If
            
            If character_screen_action = 0 Then
                Dim NuevoSeleccionado As Byte
                NuevoSeleccionado = 0

                Dim DivX As Integer, DivY As Integer

                Dim ModX As Integer, ModY As Integer

                'Ladder: Cambie valores de posicion porque se ajusto interface (Los valores de los comentarios son los reales)
                
                ' Division entera
                DivX = Int((x - 207) / 131) ' 217 = primer pj x, 131 = offset x entre cada pj
                DivY = Int((y - 246) / 158) ' 233 = primer pj y, 158 = offset y entre cada pj
                ' Resto
                ModX = (x - 207) Mod 131 ' 217 = primer pj x, 131 = offset x entre cada pj
                ModY = (y - 246) Mod 158 ' 233 = primer pj y, 158 = offset y entre cada pj
                
                ' La division no puede ser negativa (cliqueo muy a la izquierda)
                ' ni ser mayor o igual a 5 (max. pjs por linea)
                If DivX >= 0 And DivX < 5 Then

                    ' no puede ser mayor o igual a 2 (max. lineas)
                    If DivY >= 0 And DivY < 2 Then

                        ' El resto tiene que ser menor que las dimensiones del "rectangulo" del pj
                        If ModX < 79 Then ' 64 = ancho del "rectangulo" del pj
                            If ModY < 93 Then ' 64 = alto del "rectangulo" del pj

                                ' Si todo se cumple, entonces cliqueo en un pj (dado por las divisiones)
                                NuevoSeleccionado = 1 + DivX + DivY * 5 ' 5 = cantidad de pjs por linea (+1 porque los pjs van de 1 a MAX)

                                If Pjs(NuevoSeleccionado).Mapa = 0 Then NuevoSeleccionado = 0
                                
                            End If

                        End If

                    End If

                End If
                
                If PJSeleccionado <> NuevoSeleccionado Then
                    LastPJSeleccionado = PJSeleccionado
                    PJSeleccionado = NuevoSeleccionado
                End If

            End If
            
            
                
            Select Case character_screen_action
                Case e_action_close_game
                    CloseClient
                Case e_action_create_character
                    If CantidadDePersonajesEnCuenta >= 10 Then
                        Call MensajeAdvertencia(JsonLanguage.Item("ADVERTENCIA_LIMITE_PERSONAJES"))
                        
                        Exit Sub
                    End If
                    UserMap = 37
                    AlphaNiebla = 3
                    CPHeading = 3
                    CPEquipado = True
                    Call SwitchMap(UserMap)
                    g_game_state.State = e_state_createchar_screen
                    Call IniciarCrearPj
                    frmConnect.txtNombre.visible = True
                    frmConnect.txtNombre.SetFocus
                    Call ao20audio.PlayWav(SND_DICE)
               Case e_action_transfer_character
                    If SelectedCharIndex = 0 Then Exit Sub
                    TransferCharname = Pjs(SelectedCharIndex).nombre
                    If MsgBox(JsonLanguage.Item("MENSAJEBOX_TRANSFERIR_PERSONAJE") & TransferCharname & JsonLanguage.Item("MENSAJEBOX_A_OTRA_CUENTA"), vbYesNo + vbQuestion, JsonLanguage.Item("MENSAJEBOX_TRANSFERIR_TITULO")) = vbYes Then
                        frmTransferChar.Show , frmConnect
                    End If
                Case e_action_delete_character
                    If SelectedCharIndex = 0 Then Exit Sub
                    DeleteUser = Pjs(SelectedCharIndex).nombre
                    If MsgBox(JsonLanguage.Item("MENSAJEBOX_BORRAR_PERSONAJE") & DeleteUser & JsonLanguage.Item("MENSAJEBOX_DE_LA_CUENTA"), vbYesNo + vbQuestion, JsonLanguage.Item("MENSAJEBOX_BORRAR_TITULO")) = vbYes Then
                        ModAuth.LoginOperation = e_operation.deletechar
                        Call connectToLoginServer
                        frmDeleteChar.Show , frmConnect
                    End If

                Case e_action_logout_account
                    Call LogOut
                Case e_action_login_character
                    If PJSeleccionado < 1 Then Exit Sub
                    If IntervaloPermiteConectar Then
                        Call ao20audio.PlayWav(SND_CLICK)
                        Call LogearPersonaje(Pjs(PJSeleccionado).nombre)
                    End If
            End Select
            SelectedCharIndex = PJSeleccionado
            If PJSeleccionado = 0 Then Exit Sub
            If PJSeleccionado > CantidadDePersonajesEnCuenta Then Exit Sub
        
        Case e_state_connect_screen
            While LastClickAsistente = ClickEnAsistenteRandom
                ClickEnAsistenteRandom = RandomNumber(1, 4)
            Wend
            LastClickAsistente = ClickEnAsistenteRandom
             If (x > 490 And x < 522) And (y > 297 And y < 357) Then
                If ClickEnAsistenteRandom = 1 Then
                    Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_VISITAR_FORO"), False, False)
                End If
                If ClickEnAsistenteRandom = 2 Then
                    Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_INVITAR_AMIGOS"), False, False)
                End If
                If ClickEnAsistenteRandom = 3 Then
                    Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_AYUDA_JUEGO"), False, False)
                End If
                If ClickEnAsistenteRandom = 4 Then
                    Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_CONFIGURAR_JUEGO"), False, False)
                End If
            End If
    End Select
    Exit Sub
render_MouseUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmConnect.render_MouseUp", Erl)
    Resume Next
    
End Sub

#ElseIf PYMMO = 0 Then

Private Sub render_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo render_MouseUp_Err
    

    Select Case g_game_state.State()


        Case e_state_createchar_screen
       
            If x > 282 And x < 322 And y > 428 And y < 468 Then 'Boton heading
                If CPHeading + 1 >= 5 Then
                CPHeading = 1
                Else
                    CPHeading = CPHeading + 1
                End If
            End If
            
            
            
            If x > 412 And x < 446 And y > 427 And y < 470 Then 'Boton Equipar
                If CPHeading - 1 <= 0 Then
                CPHeading = 4
            Else
                CPHeading = CPHeading - 1
                End If
            End If
                    

            If x > 325 And x < 344 And y > 371 And y < 387 Then 'Boton izquierda cabezas
                If frmCrearPersonaje.Cabeza.ListCount = 0 Then Exit Sub
                If frmCrearPersonaje.Cabeza.ListIndex > 0 Then
                    frmCrearPersonaje.Cabeza.ListIndex = frmCrearPersonaje.Cabeza.ListIndex - 1

                End If

                If frmCrearPersonaje.Cabeza.ListIndex = 0 Then
                    frmCrearPersonaje.Cabeza.ListIndex = frmCrearPersonaje.Cabeza.ListCount - 1

                End If

            End If
    
            If x > 394 And x < 411 And y > 373 And y < 386 Then 'Boton Derecha cabezas
                If frmCrearPersonaje.Cabeza.ListCount = 0 Then Exit Sub
                If (frmCrearPersonaje.Cabeza.ListIndex + 1) <> frmCrearPersonaje.Cabeza.ListCount Then
                    frmCrearPersonaje.Cabeza.ListIndex = frmCrearPersonaje.Cabeza.ListIndex + 1

                End If

                If (frmCrearPersonaje.Cabeza.ListIndex + 1) = frmCrearPersonaje.Cabeza.ListCount Then
                    frmCrearPersonaje.Cabeza.ListIndex = 0

                End If

            End If
                        
            ' Clase inicial
            ' Shugar: Arreglo los botones para seleccionar la clase inicial.
                
            If x > 540 And x < 554 And y > 278 And y < 291 Then 'Boton izquierda clase
                
                Call Rotacion_boton_atras_clase

            End If
            
            If x > 658 And x < 671 And y > 278 And y < 291 Then 'Boton Derecha clase
                                
                Call Rotacion_boton_adelante_clase

            End If
                
            If x > 539 And x < 553 And y > 322 And y < 335 Then 'OK
                If frmCrearPersonaje.lstRaza.ListIndex < frmCrearPersonaje.lstRaza.ListCount - 1 Then
                    frmCrearPersonaje.lstRaza.ListIndex = frmCrearPersonaje.lstRaza.ListIndex + 1
                Else
                    frmCrearPersonaje.lstRaza.ListIndex = 0

                End If

            End If
            
            If x > 657 And x < 672 And y > 321 And y < 338 Then 'OK
                If frmCrearPersonaje.lstRaza.ListIndex - 1 < 0 Then
                    frmCrearPersonaje.lstRaza.ListIndex = frmCrearPersonaje.lstRaza.ListCount - 1
                Else
                    frmCrearPersonaje.lstRaza.ListIndex = frmCrearPersonaje.lstRaza.ListIndex - 1

                End If

            End If
            
            If x > 298 And x < 314 And y > 276 And y < 291 Then 'ok
    
                If frmCrearPersonaje.lstGenero.ListIndex < frmCrearPersonaje.lstGenero.ListCount - 1 Then
                    frmCrearPersonaje.lstGenero.ListIndex = frmCrearPersonaje.lstGenero.ListIndex + 1
                Else
                    frmCrearPersonaje.lstGenero.ListIndex = 0

                End If

            End If
            
            
            If x > 415 And x < 431 And y > 277 And y < 295 Then 'ok
                If frmCrearPersonaje.lstGenero.ListIndex - 1 < 0 Then
                    frmCrearPersonaje.lstGenero.ListIndex = frmCrearPersonaje.lstGenero.ListCount - 1
                Else
                    frmCrearPersonaje.lstGenero.ListIndex = frmCrearPersonaje.lstGenero.ListIndex - 1

                End If

            End If
            
            ' Hogar inicial
            ' Shugar: Arreglo los botones para seleccionar el hogar inicial.
            
            If x > 416 And x < 433 And y > 323 And y < 338 Then
            
                Call Rotacion_boton_adelante_ciudades

            End If
            
            
            If x > 297 And x < 314 And y > 321 And y < 340 Then
                
                Call Rotacion_boton_atras_ciudades
                
            End If

            
            If x >= 289 And x < 289 + 160 And y >= 525 And y < 525 + 37 Then 'Boton > Volver
                Call ao20audio.PlayWav(SND_CLICK)
                'UserMap = 323
                AlphaNiebla = 25
                EntradaY = 1
                EntradaX = 1
                
                'Call SwitchMap(UserMap)
                frmConnect.txtNombre.visible = False
                g_game_state.State = e_state_account_screen
                
                Call Graficos_Particulas.Engine_Select_Particle_Set(203)
                ParticleLluviaDorada = General_Particle_Create(208, -1, -1)

            End If
            
            
            If x >= 532 And x < 532 + 160 And y >= 525 And y < 525 + 37 Then 'Boton > Crear
                Call ao20audio.PlayWav(SND_CLICK)

                Dim k As Object

                
                userName = frmConnect.txtNombre.Text
                
                Dim Error As String
                If Not ValidarNombre(userName, Error) Then
                    frmMensaje.msg.Caption = Error
                    frmMensaje.Show , Me
                    Exit Sub
                End If

                UserStats.Raza = frmCrearPersonaje.lstRaza.ListIndex + 1
                UserStats.Sexo = frmCrearPersonaje.lstGenero.ListIndex + 1
                UserStats.Clase = frmCrearPersonaje.lstProfesion.ListIndex + 1
                UserStats.Hogar = frmCrearPersonaje.lstHogar.ListIndex + 1
               
                If frmCrearPersonaje.CheckData() Then
                    UserPassword = CuentaPassword
                    StopCreandoCuenta = True

                    If Connected Then
                        frmMain.ShowFPS.enabled = True
                    End If
          
                    Call Protocol_Writes.WriteLoginNewChar(userName, UserStats.Raza, UserStats.Sexo, UserStats.Clase, MiCabeza, UserStats.Hogar)
                End If
            End If

            Exit Sub
        Case e_state_account_screen
            character_screen_action = e_action_nothing_to_do
            
            If (x > 256 And x < 414) And (y > 710 And y < 747) Then 'Boton crear pj
                character_screen_action = e_action_create_character
            End If
            
            If (x > 14 And x < 112) And (y > 675 And y < 708) Then ' Boton Borrar pj
                character_screen_action = e_action_delete_character
            End If
            
            If (x > 980 And x < 1000) And (y > 675 And y < 708) Then
                character_screen_action = e_action_transfer_character
            End If
            
            If (x > 19 And x < 48) And (y > 21 And y < 45) Then ' Boton deslogear
                character_screen_action = e_action_logout_account
            End If
            
            If (x > 604 And x < 759) And (y > 711 And y < 745) Then ' Boton logear
                character_screen_action = e_action_login_character
            End If
            
            If (x > 971 And x < 1001) And (y > 21 And y < 45) Then ' Boton Cerrar
                character_screen_action = e_action_close_game
            End If
            
            If character_screen_action = e_action_nothing_to_do Then
                Dim NuevoSeleccionado As Byte
                NuevoSeleccionado = 0
                Dim DivX As Integer, DivY As Integer
                Dim ModX As Integer, ModY As Integer
                DivX = Int((x - 207) / 131) ' 217 = primer pj x, 131 = offset x entre cada pj
                DivY = Int((y - 246) / 158) ' 233 = primer pj y, 158 = offset y entre cada pj
                ModX = (x - 207) Mod 131 ' 217 = primer pj x, 131 = offset x entre cada pj
                ModY = (y - 246) Mod 158 ' 233 = primer pj y, 158 = offset y entre cada pj
                ' La division no puede ser negativa (cliqueo muy a la izquierda)
                ' ni ser mayor o igual a 5 (max. pjs por linea)
                If DivX >= 0 And DivX < 5 Then

                    ' no puede ser mayor o igual a 2 (max. lineas)
                    If DivY >= 0 And DivY < 2 Then

                        ' El resto tiene que ser menor que las dimensiones del "rectangulo" del pj
                        If ModX < 79 Then ' 64 = ancho del "rectangulo" del pj
                            If ModY < 93 Then ' 64 = alto del "rectangulo" del pj

                                ' Si todo se cumple, entonces cliqueo en un pj (dado por las divisiones)
                                NuevoSeleccionado = 1 + DivX + DivY * 5 ' 5 = cantidad de pjs por linea (+1 porque los pjs van de 1 a MAX)

                                If Pjs(NuevoSeleccionado).Mapa = 0 Then NuevoSeleccionado = 0
                                
                            End If

                        End If

                    End If

                End If
                
                If PJSeleccionado <> NuevoSeleccionado Then
                    LastPJSeleccionado = PJSeleccionado
                    PJSeleccionado = NuevoSeleccionado
                End If

            End If
                
            Select Case character_screen_action
                Case e_action_close_game
                    CloseClient

                Case e_action_create_character

                    If CantidadDePersonajesEnCuenta >= 10 Then
                        Call MensajeAdvertencia(JsonLanguage.Item("ADVERTENCIA_LIMITE_PERSONAJES"))
                        
                        Exit Sub
                    End If
                    UserMap = 37
                    AlphaNiebla = 3
                    CPHeading = 3
                    CPEquipado = True
                    Call SwitchMap(UserMap)
                    g_game_state.State = e_state_createchar_screen
 

                    Call IniciarCrearPj
                    frmConnect.txtNombre.visible = True
                    frmConnect.txtNombre.SetFocus
        
                    Call ao20audio.PlayWav(SND_DICE)
                Case e_action_delete_character

                    If SelectedCharIndex = 0 Then Exit Sub
                    DeleteUser = Pjs(SelectedCharIndex).nombre

                    Dim tmp As String

                    If MsgBox(JsonLanguage.Item("MENSAJEBOX_BORRAR_PERSONAJE") & DeleteUser & JsonLanguage.Item("MENSAJEBOX_DE_LA_CUENTA"), vbYesNo + vbQuestion, JsonLanguage.Item("MENSAJEBOX_BORRAR")) = vbYes Then

                        frmDeleteChar.Show , frmConnect
                        
            

                    End If

                Case e_action_logout_account
                
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
   
                Case e_action_logout_account

                    If PJSeleccionado < 1 Then Exit Sub

                    If IntervaloPermiteConectar Then
                        Call ao20audio.PlayWav(SND_CLICK)
                        Call LogearPersonaje(Pjs(PJSeleccionado).nombre)

                    End If

            End Select

            SelectedCharIndex = PJSeleccionado
 
            If PJSeleccionado = 0 Then Exit Sub
            If PJSeleccionado > CantidadDePersonajesEnCuenta Then Exit Sub
 
        Case e_state_connect_screen
            
            While LastClickAsistente = ClickEnAsistenteRandom
                ClickEnAsistenteRandom = RandomNumber(1, 4)
            Wend
            
            LastClickAsistente = ClickEnAsistenteRandom
            
            
             If (x > 490 And x < 522) And (y > 297 And y < 357) Then
             
                If ClickEnAsistenteRandom = 1 Then
                    Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_VISITAR_FORO"), False, True)

                End If

                If ClickEnAsistenteRandom = 2 Then
                    Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_INVITAR_AMIGOS"), False, True)

                End If

                If ClickEnAsistenteRandom = 3 Then
                    Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_HELP_IN_GAME"), False, True)
      
                End If

                If ClickEnAsistenteRandom = 4 Then
                    Call TextoAlAsistente(JsonLanguage.Item("MENSAJEBOX_GAME_SETTINGS"), False, True)

                End If

            End If

    End Select

  
    Exit Sub

render_MouseUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmConnect.render_MouseUp", Erl)
    Resume Next
    
End Sub
#End If

Private Sub Rotacion_boton_adelante_clase()

    ' Shugar - 27/7/24
    ' Saco de la selección de clases al Ladrón y al Pirata.
    ' Botón de la derecha: es el que aumenta el index.
    ' Implementación de buffer circular, arranca en eClass.Mage
  
    Select Case frmCrearPersonaje.lstProfesion.ListIndex
        Case eClass.Mage - 1
            frmCrearPersonaje.lstProfesion.ListIndex = eClass.Druid - 1
        Case eClass.Druid - 1
            frmCrearPersonaje.lstProfesion.ListIndex = eClass.Bard - 1
        Case eClass.Bard - 1
            frmCrearPersonaje.lstProfesion.ListIndex = eClass.Cleric - 1
        Case eClass.Cleric - 1
            frmCrearPersonaje.lstProfesion.ListIndex = eClass.Assasin - 1
        Case eClass.Assasin - 1
            frmCrearPersonaje.lstProfesion.ListIndex = eClass.Bandit - 1
        Case eClass.Bandit - 1
            frmCrearPersonaje.lstProfesion.ListIndex = eClass.paladin - 1
        Case eClass.paladin - 1
            frmCrearPersonaje.lstProfesion.ListIndex = eClass.Hunter - 1
        Case eClass.Hunter - 1
            frmCrearPersonaje.lstProfesion.ListIndex = eClass.Warrior - 1
        Case eClass.Warrior - 1
            frmCrearPersonaje.lstProfesion.ListIndex = eClass.Trabajador - 1
        Case eClass.Trabajador - 1
            frmCrearPersonaje.lstProfesion.ListIndex = eClass.Mage - 1
    End Select

End Sub

Private Sub Rotacion_boton_atras_clase()
    
    ' Shugar - 27/7/24
    ' Saco de la selección de clases al Ladrón y al Pirata.
    ' Botón de la izquierda: es el que disminuye el index.
    ' Implementación de buffer circular, arranca en eClass.Mage
    
    Select Case frmCrearPersonaje.lstProfesion.ListIndex
        Case eClass.Mage - 1
            frmCrearPersonaje.lstProfesion.ListIndex = eClass.Trabajador - 1
        Case eClass.Druid - 1
            frmCrearPersonaje.lstProfesion.ListIndex = eClass.Mage - 1
        Case eClass.Bard - 1
            frmCrearPersonaje.lstProfesion.ListIndex = eClass.Druid - 1
        Case eClass.Cleric - 1
            frmCrearPersonaje.lstProfesion.ListIndex = eClass.Bard - 1
        Case eClass.Assasin - 1
            frmCrearPersonaje.lstProfesion.ListIndex = eClass.Cleric - 1
        Case eClass.Bandit - 1
            frmCrearPersonaje.lstProfesion.ListIndex = eClass.Assasin - 1
        Case eClass.paladin - 1
            frmCrearPersonaje.lstProfesion.ListIndex = eClass.Bandit - 1
        Case eClass.Hunter - 1
            frmCrearPersonaje.lstProfesion.ListIndex = eClass.paladin - 1
        Case eClass.Warrior - 1
            frmCrearPersonaje.lstProfesion.ListIndex = eClass.Hunter - 1
        Case eClass.Trabajador - 1
            frmCrearPersonaje.lstProfesion.ListIndex = eClass.Warrior - 1
    End Select

End Sub


Private Sub Rotacion_boton_adelante_ciudades()

    ' Shugar - 14/6/24
    ' Limito la selección del hogar a Ulla, Nix, Arghal y Forgat.
    ' Botón de la derecha: es el que aumenta el index.
    ' Implementación de buffer circular, arranca en eCiudad.cUllathorpe
  
    Select Case frmCrearPersonaje.lstHogar.ListIndex
        Case eCiudad.cUllathorpe - 1
            frmCrearPersonaje.lstHogar.ListIndex = eCiudad.cNix - 1
        Case eCiudad.cNix - 1
            frmCrearPersonaje.lstHogar.ListIndex = eCiudad.cArghal - 1
        Case eCiudad.cArghal - 1
            frmCrearPersonaje.lstHogar.ListIndex = eCiudad.cForgat - 1
        Case eCiudad.cForgat - 1
            frmCrearPersonaje.lstHogar.ListIndex = eCiudad.cUllathorpe - 1
    End Select

End Sub

Private Sub Rotacion_boton_atras_ciudades()
    
    ' Shugar - 14/6/24
    ' Limito la selección del hogar a Ulla, Nix, Arghal y Forgat.
    ' Botón de la izquierda: es el que disminuye el index.
    ' Implementación de buffer circular, arranca en eCiudad.cUllathorpe
    
    Select Case frmCrearPersonaje.lstHogar.ListIndex
        Case eCiudad.cUllathorpe - 1
            frmCrearPersonaje.lstHogar.ListIndex = eCiudad.cForgat - 1
        Case eCiudad.cForgat - 1
            frmCrearPersonaje.lstHogar.ListIndex = eCiudad.cArghal - 1
        Case eCiudad.cArghal - 1
            frmCrearPersonaje.lstHogar.ListIndex = eCiudad.cNix - 1
        Case eCiudad.cNix - 1
            frmCrearPersonaje.lstHogar.ListIndex = eCiudad.cUllathorpe - 1
    End Select

End Sub

Private Sub txtNombre_Change()
    
    On Error GoTo txtNombre_Change_Err
    
    CPName = txtNombre

    
    Exit Sub

txtNombre_Change_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmConnect.txtNombre_Change", Erl)
    Resume Next
    
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    
    On Error GoTo txtNombre_KeyPress_Err
    
    StopCreandoCuenta = False

    
    Exit Sub

txtNombre_KeyPress_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmConnect.txtNombre_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub LogearPersonaje(ByVal nick As String)
    Call ModLogin.LoginCharacter(nick)
End Sub
