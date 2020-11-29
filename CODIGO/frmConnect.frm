VERSION 5.00
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
   Begin VB.PictureBox render 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Begin VB.Timer RelampagoFin 
         Enabled         =   0   'False
         Left            =   13080
         Top             =   1320
      End
      Begin VB.Timer relampago 
         Enabled         =   0   'False
         Interval        =   10000
         Left            =   13080
         Top             =   480
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

Private Sub Form_Activate()
    Call Graficos_Particulas.Engine_Select_Particle_Set(203)
    ParticleLluviaDorada = General_Particle_Create(208, -1, -1)

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        prgRun = False
        End

    End If

End Sub

Private Sub Form_Load()
    Call FormParser.Parse_Form(Me)

    QueRender = 1
    relampago.Enabled = True
    
    LogeoAlgunaVez = False
    EngineRun = False
        
    Timer2.Enabled = True
    Timer1.Enabled = True
    ' Seteamos el caption hay que poner 20 aniversario
    Me.Caption = "Argentum20"
    
    ' Removemos la barra de titulo pero conservando el caption para la barra de tareas
    Call Form_RemoveTitleBar(Me)
    Me.Width = 1024 * Screen.TwipsPerPixelX
    Me.Height = 768 * Screen.TwipsPerPixelY

End Sub

Private Sub relampago_Timer()

    Dim trueno         As Byte
    
    Dim truenocolor    As Byte

    Dim duraciontrueno As Byte
    
    trueno = RandomNumber(1, 255)
    
    If trueno > 100 Then

        Dim Color As Long, duracion As Long

        duraciontrueno = RandomNumber(80, 200)

        truenocolor = RandomNumber(1, 4)

        Dim TruenoWav As Integer

        TruenoWav = RandomNumber(400, 403)

        Sound.Sound_Play CStr(TruenoWav), False, 0, 0

        Select Case truenocolor

            Case 1
                Color = &H8080

            Case 2
                Color = &HF8F8F8

            Case 3
                Color = &HEFEECB

            Case 4
                Color = &HE2B3F7

        End Select

        Dim R, G, B As Byte

        B = (Color And 16711680) / 65536
        G = (Color And 65280) / 256
        R = Color And 255
        Color = D3DColorARGB(255, R, G, B)
        
        SetGlobalLight (Color)
        RelampagoFin.Interval = duraciontrueno
        RelampagoFin.Enabled = True

    End If

End Sub

Private Sub RelampagoFin_Timer()
    Call SetGlobalLight(MapDat.base_light)
    RelampagoFin.Enabled = False

End Sub

Private Sub render_DblClick()

    Select Case QueRender

        Case 2
            
            If PJSeleccionado < 1 Then Exit Sub
            If Pjs(PJSeleccionado).nombre = "" Then
                PJSeleccionado = 0
                Exit Sub

            End If

            Call Sound.Sound_Play(SND_CLICK)

            If IntervaloPermiteConectar Then
                Call LogearPersonaje(Pjs(PJSeleccionado).nombre)

            End If

        Case 3
        
    End Select

End Sub

Private Sub render_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Select Case QueRender

        Case 3
        
        Debug.Print "x: " & X & " y:" & Y

            If X > 331 And X < 347 And Y > 412 And Y < 424 Then 'Boton izquierda cabezas
                If frmCrearPersonaje.Cabeza.ListCount = 0 Then Exit Sub
                If frmCrearPersonaje.Cabeza.ListIndex > 0 Then
                    frmCrearPersonaje.Cabeza.ListIndex = frmCrearPersonaje.Cabeza.ListIndex - 1

                End If

                If frmCrearPersonaje.Cabeza.ListIndex = 0 Then
                    frmCrearPersonaje.Cabeza.ListIndex = frmCrearPersonaje.Cabeza.ListCount - 1

                End If

            End If
    
            If X > 401 And X < 415 And Y > 412 And Y < 424 Then 'Boton Derecha cabezas
                If frmCrearPersonaje.Cabeza.ListCount = 0 Then Exit Sub
                If (frmCrearPersonaje.Cabeza.ListIndex + 1) <> frmCrearPersonaje.Cabeza.ListCount Then
                    frmCrearPersonaje.Cabeza.ListIndex = frmCrearPersonaje.Cabeza.ListIndex + 1

                End If

                If (frmCrearPersonaje.Cabeza.ListIndex + 1) = frmCrearPersonaje.Cabeza.ListCount Then
                    frmCrearPersonaje.Cabeza.ListIndex = 0

                End If

            End If
                        
                
                
            If X > 540 And X < 554 And Y > 278 And Y < 291 Then 'Boton izquierda clase
                If frmCrearPersonaje.lstProfesion.ListIndex < frmCrearPersonaje.lstProfesion.ListCount - 1 Then
                    frmCrearPersonaje.lstProfesion.ListIndex = frmCrearPersonaje.lstProfesion.ListIndex + 1
                Else
                    frmCrearPersonaje.lstProfesion.ListIndex = 0

                End If

            End If
            
            If X > 658 And X < 671 And Y > 278 And Y < 291 Then 'Boton Derecha cabezas
                If frmCrearPersonaje.lstProfesion.ListIndex - 1 < 0 Then
                    frmCrearPersonaje.lstProfesion.ListIndex = frmCrearPersonaje.lstProfesion.ListCount - 1
                Else
                    frmCrearPersonaje.lstProfesion.ListIndex = frmCrearPersonaje.lstProfesion.ListIndex - 1

                End If

            End If
                
            If X > 539 And X < 553 And Y > 322 And Y < 335 Then 'OK
                If frmCrearPersonaje.lstRaza.ListIndex < frmCrearPersonaje.lstRaza.ListCount - 1 Then
                    frmCrearPersonaje.lstRaza.ListIndex = frmCrearPersonaje.lstRaza.ListIndex + 1
                Else
                    frmCrearPersonaje.lstRaza.ListIndex = 0

                End If

            End If
            
            If X > 657 And X < 672 And Y > 321 And Y < 338 Then 'OK
                If frmCrearPersonaje.lstRaza.ListIndex - 1 < 0 Then
                    frmCrearPersonaje.lstRaza.ListIndex = frmCrearPersonaje.lstRaza.ListCount - 1
                Else
                    frmCrearPersonaje.lstRaza.ListIndex = frmCrearPersonaje.lstRaza.ListIndex - 1

                End If

            End If
            
            If X > 298 And X < 314 And Y > 276 And Y < 291 Then 'ok
    
                If frmCrearPersonaje.lstGenero.ListIndex < frmCrearPersonaje.lstGenero.ListCount - 1 Then
                    frmCrearPersonaje.lstGenero.ListIndex = frmCrearPersonaje.lstGenero.ListIndex + 1
                Else
                    frmCrearPersonaje.lstGenero.ListIndex = 0

                End If

            End If
            
            If X > 415 And X < 431 And Y > 277 And Y < 295 Then 'ok
                If frmCrearPersonaje.lstGenero.ListIndex - 1 < 0 Then
                    frmCrearPersonaje.lstGenero.ListIndex = frmCrearPersonaje.lstGenero.ListCount - 1
                Else
                    frmCrearPersonaje.lstGenero.ListIndex = frmCrearPersonaje.lstGenero.ListIndex - 1

                End If

            End If
        
        
        
        'ciudad
        
            If X > 297 And X < 314 And Y > 321 And Y < 340 Then 'ok
    
                If frmCrearPersonaje.lstHogar.ListIndex < frmCrearPersonaje.lstHogar.ListCount - 1 Then
                    frmCrearPersonaje.lstHogar.ListIndex = frmCrearPersonaje.lstHogar.ListIndex + 1
                Else
                    frmCrearPersonaje.lstHogar.ListIndex = 0

                End If

            End If
            
            If X > 416 And X < 433 And Y > 323 And Y < 338 Then 'ok
                If frmCrearPersonaje.lstHogar.ListIndex - 1 < 0 Then
                    frmCrearPersonaje.lstHogar.ListIndex = frmCrearPersonaje.lstHogar.ListCount - 1
                Else
                    frmCrearPersonaje.lstHogar.ListIndex = frmCrearPersonaje.lstHogar.ListIndex - 1

                End If

            End If
        'ciudad
            If X >= 289 And X < 289 + 160 And Y >= 525 And Y < 525 + 37 Then 'Boton > Volver
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
            
            If X >= 532 And X < 532 + 160 And Y >= 525 And Y < 525 + 37 Then 'Boton > Crear
                Call Sound.Sound_Play(SND_CLICK)

                Dim k As Object

                If StopCreandoCuenta = True Then Exit Sub
                
                UserName = frmConnect.txtNombre.Text
                
                If Right$(UserName, 1) = " " Then
                    UserName = RTrim$(UserName)

                    'MsgBox "Nombre invalido, se han removido los espacios al final del nombre"
                End If

                UserRaza = frmCrearPersonaje.lstRaza.ListIndex + 1
                UserSexo = frmCrearPersonaje.lstGenero.ListIndex + 1
                UserClase = frmCrearPersonaje.lstProfesion.ListIndex + 1
                
                UserHogar = frmCrearPersonaje.lstHogar.ListIndex + 1
               
                If frmCrearPersonaje.CheckData() Then
                    UserPassword = CuentaPassword
                    StopCreandoCuenta = True

                    If frmmain.Socket1.Connected Then
                        EstadoLogin = E_MODO.CrearNuevoPj
                        Call Login
                        frmmain.ShowFPS.Enabled = True
                        Exit Sub
                    Else
                        EstadoLogin = E_MODO.CrearNuevoPj
                        frmmain.Socket1.HostName = IPdelServidor
                        frmmain.Socket1.RemotePort = PuertoDelServidor
                        frmmain.Socket1.Connect

                    End If

                End If

            End If
            
            If X >= 658 And X < 682 + 18 And Y >= 365 And Y < 385 Then
                Call Sound.Sound_Play(SND_DICE) ' Este sonido hay que ponerlo en el evento "click" o hacer q suene menos xq rompe oidos sino
                
                If frmmain.Socket1.Connected Then
                    EstadoLogin = E_MODO.Dados
                    Call Login
                Else
                    EstadoLogin = E_MODO.Dados
                    frmmain.Socket1.HostName = IPdelServidor
                    frmmain.Socket1.RemotePort = PuertoDelServidor
                    frmmain.Socket1.Connect

                End If

            End If

            Exit Sub

        Case 2
            OpcionSeleccionada = 0

            If (X > 256 And X < 414) And (Y > 710 And Y < 747) Then 'Boton crear pj
                OpcionSeleccionada = 1

            End If
            
            If (X > 14 And X < 112) And (Y > 675 And Y < 708) Then ' Boton Borrar pj
                OpcionSeleccionada = 2

            End If
            
            If (X > 19 And X < 48) And (Y > 21 And Y < 45) Then ' Boton deslogear
                OpcionSeleccionada = 3

            End If
            
            If (X > 604 And X < 759) And (Y > 711 And Y < 745) Then ' Boton logear
                OpcionSeleccionada = 4

            End If
            
            If (X > 971 And X < 1001) And (Y > 21 And Y < 45) Then ' Boton Cerrar
                OpcionSeleccionada = 5

            End If
            
            If OpcionSeleccionada = 0 Then
                PJSeleccionado = 0

                Dim DivX As Integer, DivY As Integer

                Dim ModX As Integer, ModY As Integer

                'Ladder: Cambie valores de posicion porque se ajusto interface (Los valores de los comentarios son los reales)
                
                ' Division entera
                DivX = Int((X - 207) / 131) ' 217 = primer pj x, 131 = offset x entre cada pj
                DivY = Int((Y - 246) / 158) ' 233 = primer pj y, 158 = offset y entre cada pj
                ' Resto
                ModX = (X - 207) Mod 131 ' 217 = primer pj x, 131 = offset x entre cada pj
                ModY = (Y - 246) Mod 158 ' 233 = primer pj y, 158 = offset y entre cada pj
                
                ' La division no puede ser negativa (cliqueo muy a la izquierda)
                ' ni ser mayor o igual a 5 (max. pjs por linea)
                If DivX >= 0 And DivX < 5 Then

                    ' no puede ser mayor o igual a 2 (max. lineas)
                    If DivY >= 0 And DivY < 2 Then

                        ' El resto tiene que ser menor que las dimensiones del "rectangulo" del pj
                        If ModX < 79 Then ' 64 = ancho del "rectangulo" del pj
                            If ModY < 93 Then ' 64 = alto del "rectangulo" del pj
                                ' Si todo se cumple, entonces cliqueo en un pj (dado por las divisiones)
                                PJSeleccionado = 1 + DivX + DivY * 5 ' 5 = cantidad de pjs por linea (+1 porque los pjs van de 1 a MAX)

                            End If

                        End If

                    End If

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
                    
                    If IntervaloPermiteConectar Then
                        EstadoLogin = E_MODO.Dados

                        If Musica Then

                            '  ReproducirMp3 (2)
                            'Else
                            ' Call Audio.PlayMIDI("123.mid")
                        End If

                        If frmmain.Socket1.Connected Then
                            frmmain.Socket1.Disconnect
                            frmmain.Socket1.Cleanup
                            DoEvents

                        End If

                        frmmain.Socket1.HostName = IPdelServidor
                        frmmain.Socket1.RemotePort = PuertoDelServidor
                        frmmain.Socket1.Connect

                    End If

                Case 2

                    If Char = 0 Then Exit Sub
                    DeleteUser = Pjs(Char).nombre

                    Dim tmp As String

                    If MsgBox("¿Esta seguro que desea borrar el personaje " & DeleteUser & " de la cuenta?", vbYesNo + vbQuestion, "Borrar personaje") = vbYes Then
                        Call inputbox_Password(Me, "*")
                        tmp = InputBox("Para confirmar el borrado debe ingresar su contraseña.", App.title)
            
                        If tmp = CuentaPassword Then
                            If frmmain.Socket1.Connected Then
                                frmmain.Socket1.Disconnect
                                frmmain.Socket1.Cleanup
                                DoEvents

                            End If

                            EstadoLogin = E_MODO.BorrandoPJ
                            frmmain.Socket1.HostName = IPdelServidor
                            frmmain.Socket1.RemotePort = PuertoDelServidor
                            frmmain.Socket1.Connect
                            PJSeleccionado = 0
                        Else
                            MsgBox ("Contraseña incorrecta")

                        End If

                    End If

                Case 3
                    Call ComprobarEstado

                    If Musica Then

                        'ReproducirMp3 (4)
                    End If

                    If frmmain.Socket1.Connected Then
                        frmmain.Socket1.Disconnect
                        frmmain.Socket1.Cleanup
                        DoEvents

                    End If

                    CantidadDePersonajesEnCuenta = 0
                    CuentaDonador = 0
                
                    Dim i As Integer

                    For i = 1 To 8
                        Pjs(i).Body = 0
                        Pjs(i).Head = 0
                        Pjs(i).Mapa = 0
                        Pjs(i).nivel = 0
                        Pjs(i).nombre = ""
                        Pjs(i).Clase = 0
                        Pjs(i).Criminal = 0
                        Pjs(i).NameMapa = ""
                    Next i

                    LogeoAlgunaVez = False
                    General_Set_Connect
                    
                    'Unload Me
                Case 4

                    If PJSeleccionado < 1 Then Exit Sub
                    If Pjs(PJSeleccionado).nombre = "" Then
                        PJSeleccionado = 0
                        Exit Sub

                    End If

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

            #If DEBUGGING = 1 Then

                ' Crear cuenta a manopla
                If X >= 40 And X < 195 And Y >= 330 And Y < 365 Then
                    FrmLogear.Visible = False
    
                    If frmmain.Socket1.Connected Then
                        frmmain.Socket1.Disconnect
                        frmmain.Socket1.Cleanup
                        DoEvents

                    End If
    
                    frmMasOpciones.Show , frmConnect
                    frmMasOpciones.Top = frmMasOpciones.Top + 3000
                    Exit Sub

                End If

            #End If

            If (X > 479 And X < 501) And (Y > 341 And Y < 470) Then
 
                ClickEnAsistente = ClickEnAsistente + 1

                If ClickEnAsistente = 1 Then
                    Call TextoAlAsistente("¿En que te puedo ayudar?")

                End If

                If ClickEnAsistente = 2 Then
                    Call TextoAlAsistente("¿Ya tenes una cuenta? Logea por acá abajo.")

                End If

                If ClickEnAsistente = 4 Then
                    Call TextoAlAsistente("Si necesita ayuda dentro del juego podes usar el comando /GM y un compañero mio se acercara hacia tí.")

                End If

                If ClickEnAsistente = 5 Then
                    Call TextoAlAsistente("¡Espero tengas un bello dia.")

                End If

                If ClickEnAsistente = 20 Then
                    Call TextoAlAsistente("Bueno... listo.")

                End If

                If ClickEnAsistente = 12 Then
                    Call TextoAlAsistente("¡Auch! ¡Me haces cosquillas!")

                End If

                If ClickEnAsistente = 20 Then
                    Call TextoAlAsistente("En cualquier momento se larga....")

                End If

                If ClickEnAsistente = 25 Then
                    Call TextoAlAsistente("A Ladder le falto ponerme un paragua...")

                End If

                If ClickEnAsistente = 28 Then
                    Call TextoAlAsistente("¡Para! ¡Por favor!")

                End If

                If ClickEnAsistente = 30 Then
                    Call TextoAlAsistente("¡Me estas desconcentrando!")

                End If

                If ClickEnAsistente > 35 Then
                    Call TextoAlAsistente("")

                End If

            End If

    End Select

    'ClickEnAsistente

End Sub

Private Sub txtNombre_Change()
    CPName = txtNombre

End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    StopCreandoCuenta = False

End Sub

Private Sub LogearPersonaje(ByVal Nick As String)

    If frmmain.Socket1.Connected Then
        UserName = Nick
        frmmain.ShowFPS.Enabled = True
        EstadoLogin = Normal
        Call Login
        Exit Sub
    Else
        EstadoLogin = Normal
        UserName = Nick
        frmmain.Socket1.HostName = IPdelServidor
        frmmain.Socket1.RemotePort = PuertoDelServidor
        frmmain.Socket1.Connect
        Exit Sub

    End If

End Sub

