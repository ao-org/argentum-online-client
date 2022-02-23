VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
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
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Char As Byte


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
            Case e_operation.DeleteChar
                Auth_state = e_state.RequestDeleteChar
            Case e_operation.ConfirmDeleteChar
                Auth_state = e_state.ConfirmDeleteChar
        End Select
    End If
    
End Sub

Private Sub AuthSocket_DataArrival(ByVal bytesTotal As Long)
    ModAuth.AuthSocket_DataArrival bytesTotal
End Sub

Private Sub AuthSocket_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call TextoAlAsistente("Servidor Offline, intente nuevamente.")
    Dim i As Long
    
    If Split(servers_login_connections(1), ":")(0) = IPdelServidorLogin Then
        IPdelServidorLogin = Split(servers_login_connections(2), ":")(0)
        PuertoDelServidorLogin = Split(servers_login_connections(2), ":")(1)
    Else
        IPdelServidorLogin = Split(servers_login_connections(1), ":")(0)
        PuertoDelServidorLogin = Split(servers_login_connections(1), ":")(1)
    End If
    
End Sub

Private Sub Form_Activate()
    
    On Error GoTo Form_Activate_Err
    
    Call Graficos_Particulas.Engine_Select_Particle_Set(203)
    ParticleLluviaDorada = General_Particle_Create(208, -1, -1)

    
    Exit Sub

Form_Activate_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmConnect.Form_Activate", Erl)
    Resume Next
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo Form_KeyDown_Err
    

    If KeyCode = 27 Then
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

    QueRender = 1
    
    EngineRun = False
        
    Timer2.Enabled = True
    Timer1.Enabled = True
    
    ' Seteamos el caption hay que poner 20 aniversario
    Me.Caption = "Argentum20"
    
    ' Removemos la barra de titulo pero conservando el caption para la barra de tareas
    Call Form_RemoveTitleBar(Me)
    Me.Width = 1024 * Screen.TwipsPerPixelX
    Me.Height = 768 * Screen.TwipsPerPixelY
    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmConnect.Form_Load", Erl)
    Resume Next
    
End Sub




Private Sub render_DblClick()
    
    On Error GoTo render_DblClick_Err
    

    Select Case QueRender

        Case 2
            
            If PJSeleccionado < 1 Then Exit Sub

            Call Sound.Sound_Play(SND_CLICK)

            If IntervaloPermiteConectar Then
                Call LogearPersonaje(Pjs(PJSeleccionado).nombre)

            End If

        Case 3
        
    End Select

    
    Exit Sub

render_DblClick_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmConnect.render_DblClick", Erl)
    Resume Next
    
End Sub

'Comprobación versión cliente
Public Sub AnalizarCliente()
    
    On Error GoTo Analizar_Err
    
    On Error Resume Next
    Dim json As String
    Dim jsonSplit() As String
    Dim Token As String
    
   'obtengo el MD5 del Argentum.exe
   json = Inet1.OpenURL("https://parches.ao20.com.ar/files/Version.json")
    If Left(json, 5) <> "<!DOC" Then
        If json <> "" Then
            Token = Left(Split(json, "Argentum.exe"":""")(1), 32)
        Else
            Exit Sub
        End If
    End If
    'Compruebo los MD5 con host
    If Token <> CheckMD5 Then
        Shell App.Path & "\..\..\Launcher\LauncherAO20.exe"
        End
    End If
        
    Exit Sub

Analizar_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmConnect.Analizar", Erl)
    Resume Next
End Sub


Private Sub render_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo render_MouseUp_Err
    

    Select Case QueRender

        Case 3
        
            'Debug.Print "x: " & x & " y:" & y
        
        
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
                        
                
                
            If x > 540 And x < 554 And y > 278 And y < 291 Then 'Boton izquierda clase
                If frmCrearPersonaje.lstProfesion.ListIndex < frmCrearPersonaje.lstProfesion.ListCount - 1 Then
                    frmCrearPersonaje.lstProfesion.ListIndex = frmCrearPersonaje.lstProfesion.ListIndex + 1
                Else
                    frmCrearPersonaje.lstProfesion.ListIndex = 0

                End If

            End If
            
            If x > 658 And x < 671 And y > 278 And y < 291 Then 'Boton Derecha cabezas
                If frmCrearPersonaje.lstProfesion.ListIndex - 1 < 0 Then
                    frmCrearPersonaje.lstProfesion.ListIndex = frmCrearPersonaje.lstProfesion.ListCount - 1
                Else
                    frmCrearPersonaje.lstProfesion.ListIndex = frmCrearPersonaje.lstProfesion.ListIndex - 1

                End If

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
        
        
        
        'ciudad
        
            If x > 297 And x < 314 And y > 321 And y < 340 Then 'ok
    
                If frmCrearPersonaje.lstHogar.ListIndex < frmCrearPersonaje.lstHogar.ListCount - 1 Then
                    frmCrearPersonaje.lstHogar.ListIndex = frmCrearPersonaje.lstHogar.ListIndex + 1
                Else
                    frmCrearPersonaje.lstHogar.ListIndex = 0

                End If

            End If
            
            If x > 416 And x < 433 And y > 323 And y < 338 Then 'ok
                If frmCrearPersonaje.lstHogar.ListIndex - 1 < 0 Then
                    frmCrearPersonaje.lstHogar.ListIndex = frmCrearPersonaje.lstHogar.ListCount - 1
                Else
                    frmCrearPersonaje.lstHogar.ListIndex = frmCrearPersonaje.lstHogar.ListIndex - 1

                End If

            End If
        'ciudad
            If x >= 289 And x < 289 + 160 And y >= 525 And y < 525 + 37 Then 'Boton > Volver
                Call Sound.Sound_Play(SND_CLICK)
                'UserMap = 323
                AlphaNiebla = 25
                EntradaY = 1
                EntradaX = 1
                
                'Call SwitchMap(UserMap)
                frmConnect.txtNombre.Visible = False
                QueRender = 2
                
                Call Graficos_Particulas.Engine_Select_Particle_Set(203)
                ParticleLluviaDorada = General_Particle_Create(208, -1, -1)

            End If
            
            
            If x >= 532 And x < 532 + 160 And y >= 525 And y < 525 + 37 Then 'Boton > Crear
                Call Sound.Sound_Play(SND_CLICK)

                Dim k As Object

                
                UserName = frmConnect.txtNombre.Text
                
                Dim Error As String
                If Not ValidarNombre(UserName, Error) Then
                    frmMensaje.msg.Caption = Error
                    frmMensaje.Show , Me
                    Exit Sub
                End If

                UserRaza = frmCrearPersonaje.lstRaza.ListIndex + 1
                UserSexo = frmCrearPersonaje.lstGenero.ListIndex + 1
                UserClase = frmCrearPersonaje.lstProfesion.ListIndex + 1
                
                UserHogar = frmCrearPersonaje.lstHogar.ListIndex + 1
               
                If frmCrearPersonaje.CheckData() Then
                    UserPassword = CuentaPassword
                    StopCreandoCuenta = True

                    If Connected Then
                        frmMain.ShowFPS.Enabled = True
                    End If
                    
                    'Call modNetwork.Connect(IPdelServidor, PuertoDelServidor)
                    'TODO: Mostrar ventana de creación de personaje
                    EstadoLogin = E_MODO.CrearNuevoPj
                    Call modNetwork.Connect(IPdelServidor, PuertoDelServidor)
                End If
                

            End If

            Exit Sub

        Case 2
            OpcionSeleccionada = 0

            If (x > 256 And x < 414) And (y > 710 And y < 747) Then 'Boton crear pj
                OpcionSeleccionada = 1

            End If
            
            If (x > 14 And x < 112) And (y > 675 And y < 708) Then ' Boton Borrar pj
                OpcionSeleccionada = 2

            End If
            
            If (x > 19 And x < 48) And (y > 21 And y < 45) Then ' Boton deslogear
                OpcionSeleccionada = 3

            End If
            
            If (x > 604 And x < 759) And (y > 711 And y < 745) Then ' Boton logear
                OpcionSeleccionada = 4

            End If
            
            If (x > 971 And x < 1001) And (y > 21 And y < 45) Then ' Boton Cerrar
                OpcionSeleccionada = 5

            End If
            
            If OpcionSeleccionada = 0 Then
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
                
            Select Case OpcionSeleccionada

                Case 5
                    CloseClient

                Case 1

                    If CantidadDePersonajesEnCuenta >= 10 Then
                        Call MensajeAdvertencia("Has alcanzado el limite de personajes creados por cuenta.")
                        Exit Sub

                    End If
                    UserMap = 37
                    AlphaNiebla = 3
                    CPHeading = 3
                    CPEquipado = True
                    Call SwitchMap(UserMap)
                    QueRender = 3
                    
                    Call IniciarCrearPj
                    frmConnect.txtNombre.Visible = True
                    frmConnect.txtNombre.SetFocus
        
                    Call Sound.Sound_Play(SND_DICE)
                Case 2

                    If Char = 0 Then Exit Sub
                    DeleteUser = Pjs(Char).nombre

                    Dim tmp As String

                    If MsgBox("¿Esta seguro que desea borrar el personaje " & DeleteUser & " de la cuenta?", vbYesNo + vbQuestion, "Borrar personaje") = vbYes Then
                        
                        ModAuth.LoginOperation = e_operation.DeleteChar
                        Call connectToLoginServer
                        frmDeleteChar.Show , frmConnect
                        
                        'If tmp = CuentaPassword Then
                        '    Call LoginOrConnect(E_MODO.BorrandoPJ)
                        '
                        '    If PJSeleccionado <> 0 Then
                        '        LastPJSeleccionado = PJSeleccionado
                        '        PJSeleccionado = 0
                        '    End If
                        'Else
                        '    MsgBox ("Contraseña incorrecta")

                        'End If

                    End If

                Case 3
                    Debug.Print "Vuelvo al login, debería borrar el token"
                    Auth_state = e_state.Idle
                    'Call ModAuth.LogOutRequest
                    Call ComprobarEstado

                    If Musica Then

                        'ReproducirMp3 (4)
                    End If
                
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
                    
                    'Unload Me
                Case 4

                    If PJSeleccionado < 1 Then Exit Sub

                    If IntervaloPermiteConectar Then
                        Call Sound.Sound_Play(SND_CLICK)
                        Call LogearPersonaje(Pjs(PJSeleccionado).nombre)

                    End If

            End Select

            Char = PJSeleccionado
            Rem MsgBox X & "   " & Y
 
            If PJSeleccionado = 0 Then Exit Sub
            If PJSeleccionado > CantidadDePersonajesEnCuenta Then Exit Sub
        
        Case 1
            
            While LastClickAsistente = ClickEnAsistenteRandom
                ClickEnAsistenteRandom = RandomNumber(1, 4)
            Wend
            
            LastClickAsistente = ClickEnAsistenteRandom
            
            
             If (x > 490 And x < 522) And (y > 297 And y < 357) Then
             
                If ClickEnAsistenteRandom = 1 Then
                    Call TextoAlAsistente("No te olvides de visitar nuestro foro https://www.elmesonhostigado.com/foro/")

                End If

                If ClickEnAsistenteRandom = 2 Then
                    Call TextoAlAsistente("¡Invitá a tus amigos y disfrutá en grupo tu viaje por Argentum 20!")

                End If

                If ClickEnAsistenteRandom = 3 Then
                    Call TextoAlAsistente("Si necesitás ayuda dentro del juego podés tipear /GM y escribir tu consulta")
                    
                End If

                If ClickEnAsistenteRandom = 4 Then
                    Call TextoAlAsistente("¿Sabías que podés configurar el juego a tu gusto como la respiración, modalidades del Lanzar y teclas?")
                End If

            End If

    End Select

    'ClickEnAsistente

    
    Exit Sub

render_MouseUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmConnect.render_MouseUp", Erl)
    Resume Next
    
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

Private Sub LogearPersonaje(ByVal Nick As String)
    
    On Error GoTo LogearPersonaje_Err
    
    UserName = Nick

    If Connected Then
        frmMain.ShowFPS.Enabled = True
    End If
    
    Call modNetwork.Connect(IPdelServidor, PuertoDelServidor)
    ModAuth.LoginOperation = e_operation.Authenticate
    Call LoginOrConnect(E_MODO.Normal)
    
    Exit Sub

LogearPersonaje_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmConnect.LogearPersonaje", Erl)
    Resume Next
    
End Sub
