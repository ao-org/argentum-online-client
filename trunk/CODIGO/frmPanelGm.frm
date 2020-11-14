VERSION 5.00
Begin VB.Form frmPanelGm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Panel GM"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   -75
   ClientWidth     =   4815
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Lista 2 (Consultas)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8280
      TabIndex        =   8
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Lista 1 (Principal)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   7
      Top             =   720
      Width           =   2175
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdTarget 
      Caption         =   "Seleccionar personaje"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   4560
   End
   Begin VB.TextBox txtMsg 
      Alignment       =   2  'Center
      Height          =   1035
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2640
      Width           =   4575
   End
   Begin VB.CommandButton cmdActualiza 
      Caption         =   "&Actualiza"
      Height          =   315
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.ComboBox cboListaUsus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3675
   End
   Begin VB.ListBox List2 
      Height          =   2010
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   4575
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   120
      X2              =   4680
      Y1              =   495
      Y2              =   495
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   4680
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Menu mnuUsuario 
      Caption         =   "Usuario"
      Visible         =   0   'False
      Begin VB.Menu mnuBorrar 
         Caption         =   "Borrar mensaje"
      End
      Begin VB.Menu mnuIra 
         Caption         =   "Ir al usuario"
      End
      Begin VB.Menu mnuTraer 
         Caption         =   "Traer el usuario"
      End
      Begin VB.Menu mnuResponder 
         Caption         =   "Responder"
      End
      Begin VB.Menu mnuInvalida 
         Caption         =   "Inválida"
      End
      Begin VB.Menu mnuManual 
         Caption         =   "Manual/FAQ"
      End
   End
   Begin VB.Menu mnuChar 
      Caption         =   "Personaje"
      Begin VB.Menu Destrabar 
         Caption         =   "Destrabar"
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "Echar"
         Index           =   0
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "Sumonear"
         Index           =   2
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "Ir a"
         Index           =   3
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "Ubicación"
         Index           =   6
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "IP del personaje"
         Index           =   13
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "IP anteriores"
         Index           =   14
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "Ir cerca"
         Index           =   15
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "Email anterior"
         Index           =   16
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "Libre"
         Index           =   18
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "Ver Penas"
         Index           =   20
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "Revivir"
         Index           =   21
      End
      Begin VB.Menu cmdAccion 
         Caption         =   "Advertencias"
         Index           =   22
      End
      Begin VB.Menu Ejecutar 
         Caption         =   "Ejecutar"
      End
      Begin VB.Menu CerrarleCliente 
         Caption         =   "Cerrar Cliente"
      End
      Begin VB.Menu cmdBan 
         Caption         =   "Banear"
         Begin VB.Menu BanCuenta 
            Caption         =   "Cuenta"
         End
         Begin VB.Menu Temporal 
            Caption         =   "Temporal"
         End
         Begin VB.Menu mnuBan 
            Caption         =   "Personaje"
            Index           =   1
         End
         Begin VB.Menu mnuBan 
            Caption         =   "Personaje e IP"
            Index           =   19
         End
         Begin VB.Menu banMacYHD 
            Caption         =   "Mac & HD"
         End
      End
      Begin VB.Menu mnuDesbanear 
         Caption         =   "Desbanear"
         Begin VB.Menu UnbanPersonaje 
            Caption         =   "Personaje"
         End
         Begin VB.Menu UnbanCuenta 
            Caption         =   "Cuenta"
         End
         Begin VB.Menu UnBanIp 
            Caption         =   "IP"
         End
         Begin VB.Menu UnbanMacYHD 
            Caption         =   "Mac & HD"
         End
      End
      Begin VB.Menu mnuEncarcelar 
         Caption         =   "Encarcelar"
         Begin VB.Menu mnuCarcel 
            Caption         =   "5 Minutos"
            Index           =   5
         End
         Begin VB.Menu mnuCarcel 
            Caption         =   "15 Minutos"
            Index           =   15
         End
         Begin VB.Menu mnuCarcel 
            Caption         =   "30 Minutos"
            Index           =   30
         End
         Begin VB.Menu mnuCarcel 
            Caption         =   "Definir otro"
            Index           =   60
         End
      End
      Begin VB.Menu mnuInfo 
         Caption         =   "Información"
         Begin VB.Menu mnuAccion 
            Caption         =   "General"
            Index           =   0
         End
         Begin VB.Menu mnuAccion 
            Caption         =   "Inventario"
            Index           =   1
         End
         Begin VB.Menu mnuAccion 
            Caption         =   "Skills"
            Index           =   2
         End
         Begin VB.Menu mnuAccion 
            Caption         =   "Atributos"
            Index           =   3
         End
         Begin VB.Menu mnuAccion 
            Caption         =   "Bóveda"
            Index           =   4
         End
      End
      Begin VB.Menu mnuSilenciar 
         Caption         =   "Silenciar"
         Begin VB.Menu mnuSilencio 
            Caption         =   "5 Minutos"
            Index           =   5
         End
         Begin VB.Menu mnuSilencio 
            Caption         =   "15 Minutos"
            Index           =   15
         End
         Begin VB.Menu mnuSilencio 
            Caption         =   "30 Minutos"
            Index           =   30
         End
         Begin VB.Menu mnuSilencio 
            Caption         =   "Definir otro"
            Index           =   60
         End
      End
      Begin VB.Menu Procesos 
         Caption         =   "Procesos"
         Begin VB.Menu VerProcesos 
            Caption         =   "Ver Procesos"
         End
         Begin VB.Menu CerrarProceso 
            Caption         =   "Cerrar Proceso"
         End
      End
      Begin VB.Menu Editar 
         Caption         =   "Editar"
         Begin VB.Menu Vida 
            Caption         =   "Vida"
         End
         Begin VB.Menu Energia 
            Caption         =   "Energia"
         End
         Begin VB.Menu Mana 
            Caption         =   "Mana"
         End
         Begin VB.Menu oro 
            Caption         =   "Oro"
         End
         Begin VB.Menu SkillLibres 
            Caption         =   "Skill Libres"
         End
         Begin VB.Menu ciudadanos 
            Caption         =   "Ciudadanos"
         End
         Begin VB.Menu Criminales 
            Caption         =   "Criminales"
         End
         Begin VB.Menu Cabeza 
            Caption         =   "Cabeza"
         End
         Begin VB.Menu Cuerpo 
            Caption         =   "Cuerpo"
         End
         Begin VB.Menu Clase 
            Caption         =   "Clase"
         End
         Begin VB.Menu Raza 
            Caption         =   "Raza"
         End
      End
      Begin VB.Menu EnviarA 
         Caption         =   "Enviar a"
         Begin VB.Menu MnuEnviar 
            Caption         =   "UllaThorpe"
            Index           =   0
         End
         Begin VB.Menu MnuEnviar 
            Caption         =   "Nix"
            Index           =   1
         End
         Begin VB.Menu MnuEnviar 
            Caption         =   "Banderbille"
            Index           =   2
         End
         Begin VB.Menu MnuEnviar 
            Caption         =   "Arghal"
            Index           =   3
         End
         Begin VB.Menu MnuEnviar 
            Caption         =   "Otro"
            Index           =   4
         End
      End
   End
   Begin VB.Menu cmdHerramientas 
      Caption         =   "Herramientas"
      Begin VB.Menu torneos 
         Caption         =   "Torneos"
         Begin VB.Menu creartoneo 
            Caption         =   "Crear Nuevo"
         End
         Begin VB.Menu torneo_comenzar 
            Caption         =   "Comenzar Actual"
         End
         Begin VB.Menu torneo_cancelar 
            Caption         =   "Cancelar actual"
         End
      End
      Begin VB.Menu cmdcrearevento 
         Caption         =   "Eventos"
         Begin VB.Menu evento1 
            Caption         =   "Exp y oro x2 - 30 min"
         End
         Begin VB.Menu evento2 
            Caption         =   "Exp y oro x2 - 59 min"
         End
         Begin VB.Menu evento3 
            Caption         =   "Todo x 2 - 30 min"
         End
         Begin VB.Menu evento4 
            Caption         =   "Exp x3 - 30 min"
         End
         Begin VB.Menu personalizado 
            Caption         =   "Personalizado"
         End
         Begin VB.Menu finalizarevento 
            Caption         =   "Finalizar actual"
         End
         Begin VB.Menu BusqedaTesoro 
            Caption         =   "Busqueda del tesoro"
         End
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Insertar comentario"
         Index           =   4
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Enviar hora"
         Index           =   5
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Enemigos en mapa"
         Index           =   7
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Limpiar Mapa"
         Index           =   15
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Usuarios trabajando"
         Index           =   23
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Bloquear tile"
         Index           =   26
      End
      Begin VB.Menu mnuHerramientas 
         Caption         =   "Usuarios en el mapa"
         Index           =   30
      End
      Begin VB.Menu Spawn 
         Caption         =   "Lista de Spawn"
      End
      Begin VB.Menu Teleports 
         Caption         =   "Teleports"
         Begin VB.Menu CrearTeleport 
            Caption         =   "Crear Teleport"
         End
         Begin VB.Menu DestruirTeleport 
            Caption         =   "Destruir"
         End
      End
      Begin VB.Menu IP 
         Caption         =   "Direcciónes de IP"
         Index           =   0
         Begin VB.Menu mnuIP 
            Caption         =   "Banear una IP"
            Index           =   17
         End
         Begin VB.Menu mnuIP 
            Caption         =   "Lista de IPs baneadas"
            Index           =   25
         End
         Begin VB.Menu Desbanear 
            Caption         =   "Desbanear una Ip"
         End
      End
      Begin VB.Menu velocidadengine 
         Caption         =   "Velocidad de Engine"
         Begin VB.Menu Normal 
            Caption         =   "Normal"
         End
         Begin VB.Menu rapido 
            Caption         =   "Rapido"
         End
         Begin VB.Menu muyrapido 
            Caption         =   "Muy rapido"
         End
      End
      Begin VB.Menu usersOnline 
         Caption         =   "Usuarios Online"
      End
      Begin VB.Menu StaffOnline 
         Caption         =   "Staff Online"
      End
      Begin VB.Menu ResetPozos 
         Caption         =   "Reseter Pozos Magicos"
      End
      Begin VB.Menu quitarnpcs 
         Caption         =   "Quitar NPCs del area"
      End
   End
   Begin VB.Menu Admin 
      Caption         =   "Administración"
      Index           =   0
      Begin VB.Menu mnuAdmin 
         Caption         =   "Apagar servidor"
         Index           =   27
      End
      Begin VB.Menu mnuAdmin 
         Caption         =   "Grabar personajes"
         Index           =   28
      End
      Begin VB.Menu mnuAdmin 
         Caption         =   "Iniciar WorldSave"
         Index           =   29
      End
      Begin VB.Menu mnuAdmin 
         Caption         =   "Detener o reanudar el mundo"
         Index           =   33
      End
      Begin VB.Menu Limpiezas 
         Caption         =   "Limpiezas"
         Begin VB.Menu LimpiarVision 
            Caption         =   "Limpiar Vision"
         End
         Begin VB.Menu Limpiarmundo 
            Caption         =   "Limpiar el mundo"
         End
      End
      Begin VB.Menu mnuRecargar 
         Caption         =   "Actualizar"
         Index           =   35
         Begin VB.Menu mnuReload 
            Caption         =   "Objetos"
            Index           =   1
         End
         Begin VB.Menu mnuReload 
            Caption         =   "General"
            Index           =   2
         End
         Begin VB.Menu mnuReload 
            Caption         =   "Hechizos"
            Index           =   4
         End
         Begin VB.Menu mnuReload 
            Caption         =   "Motd"
            Index           =   5
         End
         Begin VB.Menu mnuReload 
            Caption         =   "NPCs"
            Index           =   6
         End
         Begin VB.Menu mnuReload 
            Caption         =   "Sockets"
            Index           =   7
         End
         Begin VB.Menu mnuReload 
            Caption         =   "Opciones"
            Index           =   8
         End
      End
      Begin VB.Menu Ambiente 
         Caption         =   "Estado climático"
         Index           =   0
         Begin VB.Menu mnuAmbiente 
            Caption         =   "Iniciar o detener una lluvia"
            Index           =   31
         End
         Begin VB.Menu mnuAmbiente 
            Caption         =   "Anochecer o amanecer"
            Index           =   32
         End
      End
      Begin VB.Menu Subastas 
         Caption         =   "Subastas"
         Begin VB.Menu SubastaEstado 
            Caption         =   "Habilitar/Desabilitar"
         End
         Begin VB.Menu SubastaCerrar 
            Caption         =   "Cerrar Actual"
         End
      End
      Begin VB.Menu centinela 
         Caption         =   "Centinela"
         Begin VB.Menu CentinelaEstado 
            Caption         =   "Habilitar/Desabilitar"
         End
      End
      Begin VB.Menu Global 
         Caption         =   "Global"
         Begin VB.Menu GlobalEstado 
            Caption         =   "Habilitar/Desabilitar"
         End
      End
      Begin VB.Menu mapas 
         Caption         =   "Mapas"
         Begin VB.Menu SeguroInseguro 
            Caption         =   "Mapa Seguro/Inseguro"
         End
         Begin VB.Menu GuardarMapa 
            Caption         =   "Guardar Mapa"
         End
      End
      Begin VB.Menu BorrarPersonaje 
         Caption         =   "Borrar Personaje"
      End
      Begin VB.Menu MOTD 
         Caption         =   "Cambiar MOTD"
      End
   End
   Begin VB.Menu Mensajeria 
      Caption         =   "Propios"
      Begin VB.Menu MensajeriaMenu 
         Caption         =   "Mensaje por Consola"
         Index           =   0
      End
      Begin VB.Menu MensajeriaMenu 
         Caption         =   "Mensaje por Ventana"
         Index           =   1
      End
      Begin VB.Menu MensajeriaMenu 
         Caption         =   "Mensaje a GMS"
         Index           =   2
      End
      Begin VB.Menu MensajeriaMenu 
         Caption         =   "Hablar como NPC"
         Index           =   3
      End
      Begin VB.Menu YoAcciones 
         Caption         =   "Invsible"
         Index           =   0
      End
      Begin VB.Menu YoAcciones 
         Caption         =   "Chat Color"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmPanelGm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Nick       As String

Dim tmp        As String

Public LastStr As String

Private Const MAX_GM_MSG = 300

Dim reason                      As Long

Private MisMSG(0 To MAX_GM_MSG) As String

Private Apunt(0 To MAX_GM_MSG)  As Integer

Public Sub CrearGMmSg(Nick As String, msg As String)

    If List1.ListCount < MAX_GM_MSG Then
        List1.AddItem Nick & "-" & List1.ListCount
        MisMSG(List1.ListCount - 1) = msg
        Apunt(List1.ListCount - 1) = List1.ListCount - 1

    End If

End Sub

Private Sub BanCuenta_Click()
    tmp = InputBox("¿Motivo?", "Ingrese el motivo")
    Nick = cboListaUsus.Text

    If MsgBox("¿Está seguro que desea banear la cuenta de """ & Nick & """?", vbYesNo + vbQuestion) = vbYes Then
        Call WriteBanCuenta(Nick, tmp)

    End If

End Sub

Private Sub banMacYHD_Click()
    Call WriteBanSerial(cboListaUsus.Text)

End Sub

Private Sub BorrarPersonaje_Click()

    If MsgBox("¿Está seguro que desea Borrar el personaje " & cboListaUsus.Text & "?", vbYesNo + vbQuestion) = vbYes Then

        'Call SendData("/KILLCHAR " & cboListaUsus.Text)
    End If

End Sub

Private Sub BusqedaTesoro_Click()

    tmp = InputBox("Ingrese tipo de evento:" & vbCrLf & "0: Busqueda de tesoro en continente" & vbCrLf & "1: Busqueda de tesoro en dungeon" & vbCrLf & "2: Aparicion de criatura", "Iniciar evento")

    If tmp > 255 Then Exit Sub
    If IsNumeric(tmp) Then

        Call WriteBusquedaTesoro(CByte(tmp))
    Else
        MsgBox ("Tipo invalido")

    End If

End Sub

Private Sub Cabeza_Click()
    tmp = InputBox("Ingrese el valor de cabeza que desea editar.", "Edicion de Usuarios")

    'Call SendData("/MOD " & cboListaUsus.Text & " Head " & tmp)
End Sub

Private Sub CentinelaEstado_Click()

    'Call SendData("/CENTINELAACTIVADO")
End Sub

Private Sub CerrarleCliente_Click()
    Call WriteCerraCliente(cboListaUsus.Text)

End Sub

Private Sub CerrarProceso_Click()
    tmp = InputBox("Ingrese el nombre del proceso", "Cerrar Proceso")

    If tmp <> "" Then

        'Call SendData("/CERRARPROCESO " & cboListaUsus.Text & "@" & tmp)
    End If

End Sub

Private Sub ciudadanos_Click()
    tmp = InputBox("Ingrese el valor de ciudadanos que desea editar.", "Edicion de Usuarios")

    'Call SendData("/MOD " & cboListaUsus.Text & " CIU " & tmp)
End Sub

Private Sub Clase_Click()
    tmp = InputBox("Ingrese el valor de clase Libres que desea editar.", "Edicion de Usuarios")

    'Call SendData("/MOD " & cboListaUsus.Text & " CLASE " & tmp)
End Sub

Private Sub cmdAccion_Click(Index As Integer)
    Nick = Replace(cboListaUsus.Text, " ", "+")

    Select Case Index

        Case 0 '/ECHAR NICK 0.12.1
            Call WriteKick(Nick)

        Case 1 '/BAN NICK MOTIVO 0.12.1
            tmp = InputBox("¿Motivo?", "Ingrese el motivo")

            If MsgBox("¿Está seguro que desea banear al personaje """ & cboListaUsus.Text & """?", vbYesNo + vbQuestion) = vbYes Then
                Call WriteBanChar(Nick, tmp)

            End If

        Case 2 '/SUM NICK 0.12.1

            If LenB(Nick) <> 0 Then Call WriteSummonChar(Nick)

        Case 3 '/ira NICK 0.12.1

            If LenB(Nick) <> 0 Then Call WriteGoToChar(Nick)

        Case 4 '/REM 0.12.1
            tmp = InputBox("¿Comentario?", "Ingrese comentario")
            Call WriteComment(tmp)

        Case 5 '/HORA 0.12.1
            Call Protocol.WriteServerTime

        Case 6 '/DONDE NICK 0.12.1

            If LenB(Nick) <> 0 Then Call WriteWhere(Nick)

        Case 7 '/NENE 0.12.1
            tmp = InputBox("¿En qué mapa?", "")
            Call ParseUserCommand("/NENE " & tmp)

        Case 8 '/info nick
            ' Call SendData("/INFO " & Nick)
   
        Case 9 '/inv nick
            ' Call SendData("/INV " & Nick)
   
        Case 10 '/skills nick
            ' Call SendData("/SKILLS " & Nick)
   
        Case 11 '/CARCEL NICK @ MOTIVO  0.12.1
            tmp = InputBox("¿Minutos a encarcelar? (hasta 60)", "")

            If tmp > 1 Then
                Call ParseUserCommand("/CARCEL " & Nick & "@encarcelado via panelgm@" & tmp)
           
            Else
                MsgBox ("Ingreso un tiempo invalido.")
            
            End If

        Case 13 '/nick2ip NICK 0.12.1
            Call WriteNickToIP(Nick)

        Case 14 '/Lastip NICK 0.12.1
            Call WriteLastIP(Nick)

        Case 15 '/IrCerca NICK 0.12.1

            If LenB(Nick) <> 0 Then Call WriteGoNearby(Nick)

        Case 16 '/LASTMAIL NICK 0.12.1
            Call WriteRequestCharMail(Nick)

        Case 17 '/BANIP IP 0.12.1
            tmp = InputBox("Escriba la dirección IP a banear.", "")
            reason = InputBox("Escriba el motivo del baneo.", "")

            If MsgBox("¿Esta seguro que desea banear la IP """ & tmp & ", debido a " & reason & """?", vbYesNo + vbQuestion) = vbYes Then
                Call ParseUserCommand("/BANIP " & tmp & " " & reason)

            End If

        Case 18 '/bov nick

        Case 19 '/BANED IP AND PERSONAJE 0.12.1   REVISAR
    
            If MsgBox("¿Esta seguro que desea banear la IP y el personaje """ & Nick & """?", vbYesNo + vbQuestion) = vbYes Then
        
                Call ParseUserCommand("/banip " & Nick & " panelgm")

                'Call WriteBanIP(False, str2ipv4l("0.0.0.0"), ArgumentosAll(0), Right$(ArgumentosRaw, Len(ArgumentosRaw) - Len(ArgumentosAll(0)) - 1))
            End If

        Case 20 '/PENAS NICK 0.12.1
            Call WritePunishments(Nick)

        Case 21 '/REVIVIR NICK 0.12.1
            Call WriteReviveChar(Nick)

        Case 22 'ADVERTENCIA 0.12.1
            tmp = InputBox("Escriba el motivo de la advertencia.", "Advertir a " & Nick)

            If LenB(tmp) <> 0 Then
                Call ParseUserCommand("/ADVERTENCIA " & Nick & "@" & tmp)

            End If

        Case 23 '/TRABAJANDO 0.12.1
            Call WriteWorking

        Case 25 '/BANIPLIST 0.12.1
            Call WriteBannedIPList

        Case 26 '/BLOQ 0.12.1
            Call WriteTileBlockedToggle

        Case 27 '/APAGAR 0.12.1

            'Call WriteTurnOffServer
        Case 28 '/GRABAR 0.12.1
            Call WriteSaveChars

        Case 29 '/DOBACKUP 0.12.1
            Call WriteDoBackup

        Case 30 '/ONLINEMAP 0.12.1
            Call WriteOnlineMap

        Case 31 '/LLUVIA 0.12.1
            Call WriteRainToggle

        Case 32 '/NOCHE 0.12.1
            Call WriteNight

        Case 33

            'Call SendData("/PAUSAR")
        Case 34 '/LIMPIARMUNDO 0.12.1
            Call WriteCleanWorld

        Case 35 '/SILENCIO NICK@TIEMPO

            tmp = InputBox("¿Minutos a silenciar? (hasta 255)", "")

            If MsgBox("¿Esta seguro que desea silenciar al personaje """ & Nick & """?", vbYesNo + vbQuestion) = vbYes Then
                If tmp > 255 Then Exit Sub
                Call ParseUserCommand("/SILENCIO " & cboListaUsus.Text & "@" & tmp)

            End If

    End Select

    Nick = ""

End Sub

Private Sub cmdActualiza_Click()
    Call WriteRequestUserList
    Call FlushBuffer

End Sub

Private Sub cmdcerrar_Click()

    Me.Visible = False
    List1.Clear
    List2.Clear
    txtMsg.Text = ""

End Sub

Private Sub cmdOnline_Click()

End Sub

Private Sub cmdTarget_Click()
    'Dim Usuaritio As String

    'cboListaUsus = List1.List(List1.ListIndex)
    'Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre el personaje...", 100, 100, 120, 0, 0)
    'frmMain.MousePointer = 2
    'frmMain.PanelSelect = True
    'Call SendData("TGUSER")
    Call WriteMarcaDeGm

End Sub

Private Sub Command1_Click()
    List1.Visible = True
    List2.Visible = False

End Sub

Private Sub Command2_Click()
    List1.Visible = False
    List2.Visible = True

End Sub

Private Sub CrearTeleport_Click()
    tmp = InputBox("Ingrese las cordenadas, por ejemplo para ulla: 1 50 50", "Ingrese Posiciones")
    Call ParseUserCommand("/CT " & tmp)

End Sub

Private Sub creartoneo_Click()
    FrmTorneo.Show

End Sub

Private Sub Criminales_Click()
    tmp = InputBox("Ingrese el valor de criminales que desea editar.", "Edicion de Usuarios")

    'Call SendData("/MOD " & cboListaUsus.Text & " CRI " & tmp)
End Sub

Private Sub Cuerpo_Click()
    tmp = InputBox("Ingrese el valor de cuerpo que desea editar.", "Edicion de Usuarios")

    'Call SendData("/MOD " & cboListaUsus.Text & " BODY " & tmp)
End Sub

Private Sub Desbanear_Click()
    tmp = InputBox("Escriba la dirección IP a desbanear", "")

    If MsgBox("¿Esta seguro que desea desbanear la IP """ & tmp & """?", vbYesNo + vbQuestion) = vbYes Then
        Call ParseUserCommand("/UNBANIP " & tmp)

    End If

End Sub

Private Sub Destrabar_Click()
    Nick = Replace(cboListaUsus.Text, " ", "+")
    Call WritePossUser(Nick)

End Sub

Private Sub DestruirTeleport_Click()
    Call WriteTeleportDestroy '/DT 0.12.1

End Sub

Private Sub Ejecutar_Click()
    Nick = cboListaUsus.Text
    Call WriteExecute(Nick) '/EJECUTAR NICK 0.12.1

End Sub

Private Sub Energia_Click()
    tmp = InputBox("Ingrese el valor de energia que desea editar.", "Edicion de Usuarios")

    'Call SendData("/MOD " & cboListaUsus.Text & " EN " & tmp)
End Sub

Private Sub evento1_Click()
    Call WriteCrearEvento(5, 30, 2)

End Sub

Private Sub evento2_Click()
    Call WriteCrearEvento(5, 59, 2)

End Sub

Private Sub evento3_Click()
    Call WriteCrearEvento(7, 30, 2)

End Sub

Private Sub evento4_Click()
    Call WriteCrearEvento(2, 30, 3)

End Sub

Private Sub finalizarevento_Click()

    If MsgBox("¿Esta seguro que desea finalizar el evento?", vbYesNo + vbQuestion, "¡ATENCION!") = vbYes Then
        Call WriteDenounce

    End If

End Sub

Private Sub Form_Load()
    List1.Clear
    List2.Clear
    txtMsg.Text = ""
    Call WriteRequestUserList
    Call FlushBuffer

End Sub

Private Sub lento_Click()

End Sub

Private Sub GlobalEstado_Click()

    'Call SendData("/ACTIVAR")
End Sub

Private Sub GuardarMapa_Click()

    'Call SendData("/BACK")
End Sub

Private Sub Limpiarmundo_Click()

    'Call SendData("/LIMPIARMUNDO")
End Sub

Private Sub LimpiarVision_Click()
    Call WriteDestroyAllItemsInArea

End Sub

Private Sub List1_Click()

    Dim ind As Integer

    ind = Val(ReadField(2, List1.List(List1.ListIndex), Asc("@")))
    txtMsg = List2.List(List1.ListIndex)

End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        PopUpMenu mnuUsuario

    End If

End Sub

Private Sub Mana_Click()
    tmp = InputBox("Ingrese el valor de mana que desea editar.", "Edicion de Usuarios")

    'Call SendData("/MOD " & cboListaUsus.Text & " MP " & tmp)
End Sub

Private Sub MensajeriaMenu_Click(Index As Integer)

    Select Case Index

        Case 0 'Mensaje por consola a usuarios 0.12.1
            tmp = InputBox("Ingrese el texto:", "Mensaje por consola a usuarios")
            Call WriteServerMessage(tmp)

        Case 1 'Mensaje por ventana a usuarios 0.12.1
            tmp = InputBox("Ingrese el texto:", "Mensaje del sistema a usuarios")
            Call WriteSystemMessage(tmp)

        Case 2 'Mensaje por consola a GMS 0.12.1
            tmp = InputBox("Escriba el mensaje.", "Mensaje por consola de GM")
            Call WriteGMMessage(tmp)

        Case 3 'Hablar como NPC 0.12.1
            tmp = InputBox("Escriba un Mensaje.", "Hablar por NPC")
            Call WriteTalkAsNPC(tmp)

    End Select

End Sub

Private Sub mnuBorrar_Click()

    Dim elitem          As String

    Dim ProximamentTipo As String

    Dim TIPO            As String

    elitem = List1.ListIndex

    If List1.ListIndex < 0 Then Exit Sub
    Call ReadNick
    ProximamentTipo = General_Field_Read(2, List1.List(List1.ListIndex), "(")
    TIPO = General_Field_Read(1, ProximamentTipo, ")")
    Call WriteSOSRemove(Nick & "Ø" & txtMsg & "Ø" & TIPO)
    List1.RemoveItem List1.ListIndex
    List2.RemoveItem elitem
    txtMsg.Text = ""

End Sub

Private Sub MnuEnviar_Click(Index As Integer)

    Dim Coordenadas As String

    Nick = Replace(cboListaUsus.Text, " ", "+")

    Select Case Index

        Case 0 'Ulla
            Coordenadas = "55 57 46"
            Call ParseUserCommand("/TELEP " & Nick & " " & Coordenadas)

        Case 1 'Nix
            Coordenadas = "106 30 72"
            Call ParseUserCommand("/TELEP " & Nick & " " & Coordenadas)

        Case 2 'Bander
            Coordenadas = "59 50 50"
            Call ParseUserCommand("/TELEP " & Nick & " " & Coordenadas)

        Case 3 'Arghal
            Coordenadas = "151 50 50"
            Call ParseUserCommand("/TELEP " & Nick & " " & Coordenadas)

        Case 4 'Otro

            If LenB(Nick) <> 0 Then
                Coordenadas = InputBox("Indique la posición (MAPA X Y).", "Transportar a " & Nick)

                If LenB(Coordenadas) <> 0 Then Call ParseUserCommand("/TELEP " & Nick & " " & Coordenadas)

            End If

    End Select

End Sub

Private Sub mnuIRa_Click()
    Call WriteGoToChar(ReadField(1, List1.List(List1.ListIndex), Asc("(")))

End Sub

Private Sub mnutraer_Click()
    Call WriteSummonChar(ReadField(1, List1.List(List1.ListIndex), Asc("(")))

End Sub

Private Sub mnuInvalida_Click()
    Nick = ReadField(1, List1.List(List1.ListIndex), Asc("("))
    Call ParseUserCommand("/MENSAJEINFORMACION " & Nick & "@" & "Su consulta fue rechazada debido a que esta catalogada como invalida.")

End Sub

Private Sub mnuResponder_Click()
    Nick = ReadField(1, List1.List(List1.ListIndex), Asc("("))
    tmp = InputBox("Ingrese la respuesta:", "Responder consulta")
    Call ParseUserCommand("/MENSAJEINFORMACION " & Nick & "@" & tmp)

End Sub

Private Sub mnuManual_Click()
    Nick = ReadField(1, List1.List(List1.ListIndex), Asc("("))
    Call ParseUserCommand("/MENSAJEINFORMACION " & Nick & "@" & "Su consulta fue rechazada debido a que la respuesta se encuentra en el Manual o FAQ de nuestra pagina web. Para mas información visite: www.argentum20.com.ar.")

End Sub

Private Sub mnuAccion_Click(Index As Integer)
    Nick = cboListaUsus.Text

    If LenB(Nick) <> 0 Then

        Select Case Index

            Case 0 ' Informacion General
                Call WriteRequestCharStats(Nick)

            Case 1 ' Inventario
                Call WriteRequestCharInventory(Nick)

            Case 2 'Skill
                Call WriteRequestCharSkills(Nick)

            Case 3 'Atributos
                Call WriteRequestCharInfo(Nick)

            Case 4 'Boveda
                Call WriteRequestCharBank(Nick)
                Call WriteRequestCharGold(Nick)

        End Select

    End If

End Sub

Private Sub mnuAdmin_Click(Index As Integer)
    Call cmdAccion_Click(Index)

End Sub

Private Sub mnuAmbiente_Click(Index As Integer)
    Call cmdAccion_Click(Index)

End Sub

Private Sub mnuBan_Click(Index As Integer)
    Call cmdAccion_Click(Index)

End Sub

Private Sub mnuCarcel_Click(Index As Integer)

    If Index = 60 Then
        Call cmdAccion_Click(11)
        Exit Sub

    End If

    Nick = cboListaUsus.Text

    Call ParseUserCommand("/CARCEL " & Nick & "@encarcelado via panelgm@" & Index)

End Sub

Private Sub mnuSilencio_Click(Index As Integer)

    If Index = 60 Then
        Call cmdAccion_Click(35)
        Exit Sub

    End If

    Call ParseUserCommand("/SILENCIO " & cboListaUsus.Text & "@" & Index)

End Sub

Private Sub mnuHerramientas_Click(Index As Integer)
    Call cmdAccion_Click(Index)

End Sub

Private Sub mnuIP_Click(Index As Integer)
    Call cmdAccion_Click(Index)

End Sub

Private Sub mnuReload_Click(Index As Integer)

    Select Case Index

        Case 1 'Reload objetos
            Call WriteReloadObjects

        Case 2 'Reload server.ini
            Call WriteReloadServerIni

        Case 3 'Reload mapas

            ' Call SendData("/RELOAD MAP")
        Case 4 'Reload hechizos
            Call WriteReloadSpells

        Case 5 'Reload motd

            '  Call SendData("/RELOADMOTD")
        Case 6 'Reload npcs
            Call WriteReloadNPCs

        Case 7 'Reload sockets

             If MsgBox("Al realizar esta acción reiniciará la API de Winsock. Se cerrarán todas las conexiónes.", vbYesNo, "Advertencia") = vbYes Then
               '   Call SendData("/RELOAD SOCK")
           End If

    Case 8 'Reload otros

        ' Call SendData("/RELOADOPCIONES")
End Select

End Sub

Private Sub MOTD_Click()
    Call WriteChangeMOTD 'Cambiar MOTD 0.12.1

End Sub

Private Sub muyrapido_Click()
    charlist(UserCharIndex).Speeding = 5

End Sub

Private Sub Normal_Click()
    charlist(UserCharIndex).Speeding = 1#

End Sub

Private Sub oro_Click()
    tmp = InputBox("Ingrese el valor de oro que desea editar.", "Edicion de Usuarios")

    'Call SendData("/MOD " & cboListaUsus.Text & " ORO " & tmp)
End Sub

Private Sub personalizado_Click()
    tmp = InputBox("Ingrese evento  Tipo@Duracion@Multiplicacion" & vbCrLf & vbCrLf & "Tipo 1=Multiplica Oro" & vbCrLf & "Tipo 2=Multiplica Experiencia" & vbCrLf & "Tipo 3=Multiplica Recoleccion" & vbCrLf & "Tipo 4=Multiplica Dropeo" & vbCrLf & "Tipo 5=Multiplica Oro y Experiencia" & vbCrLf & "Tipo 6=Multiplica Oro, experiencia y recoleccion" & vbCrLf & "Tipo 7=Multiplica Todo" & vbCrLf & "Duracion= Maximo: 59" & vbCrLf & "Multiplicacion= Maximo 10", "Creacion de nuevo evento")
    Call ParseUserCommand("/CREAREVENTO " & tmp)

End Sub

Private Sub quitarnpcs_Click()

    'Call SendData("/LIMPIAR")
End Sub

Private Sub rapido_Click()
    charlist(UserCharIndex).Speeding = 2

End Sub

Private Sub Raza_Click()
    tmp = InputBox("Ingrese el valor de raza que desea editar.", "Edicion de Usuarios")

    'Call SendData("/MOD " & cboListaUsus.Text & " RAZA " & tmp)
End Sub

Private Sub ResetPozos_Click()

    'Call SendData("/RESETPOZOS")
End Sub

Private Sub SeguroInseguro_Click()

    'Call SendData("/SEGURO")
End Sub

Private Sub SkillLibres_Click()
    tmp = InputBox("Ingrese el valor de skills Libres que desea editar.", "Edicion de Usuarios")

    'Call SendData("/MOD " & cboListaUsus.Text & " SKILLSLIBRES " & tmp)
End Sub

Private Sub Spawn_Click()
    Call WriteSpawnListRequest

End Sub

Private Sub StaffOnline_Click()
    Call WriteOnlineGM '/ONLINEGM 0.12.1

End Sub

Private Sub SubastaEstado_Click()

    'Call SendData("/SUBASTAACTIVADA")
End Sub

Private Sub Temporal_Click()

    Dim tmp  As String

    Dim tmp2 As Byte

    tmp2 = InputBox("¿Dias?", "Ingrese cantidad de días (Maximo 255)")
    tmp = InputBox("¿Motivo?", "Ingrese el motivo")

    If MsgBox("¿Está seguro que desea banear el personaje de """ & cboListaUsus.Text & """ por " & tmp2 & " días?", vbYesNo + vbQuestion) = vbYes Then
        Call WriteBanTemporal(cboListaUsus.Text, tmp, tmp2)

    End If

End Sub

Private Sub torneo_cancelar_Click()
    Call WriteCancelarTorneo

End Sub

Private Sub torneo_comenzar_Click()
    Call WriteComenzarTorneo

End Sub

Private Sub UnbanCuenta_Click()
    Call WriteUnBanCuenta(cboListaUsus.Text)

End Sub

Private Sub UnbanPersonaje_Click()
    Nick = cboListaUsus.Text

    If MsgBox("¿Esta seguro que desea removerle el ban al personaje """ & Nick & """?", vbYesNo + vbQuestion) = vbYes Then
        Call WriteUnbanChar(Nick)

    End If

End Sub

Private Sub usersOnline_Click()

    'Call SendData("/ONLINE")
End Sub

Private Sub VerProcesos_Click()

    'Call SendData("/VERPROCESOS " & cboListaUsus.Text)
End Sub

Private Sub Vida_Click()
    tmp = InputBox("Ingrese el valor de vida que desea editar.", "Edicion de Usuarios")

    'Call SendData("/MOD " & cboListaUsus.Text & " HP " & tmp)
End Sub

Private Sub ReadNick()

    If List1.Visible Then
        Nick = General_Field_Read(1, List1.List(List1.ListIndex), "(")

        If Nick = "" Then Exit Sub
        Nick = Left$(Nick, Len(Nick))
    Else
        Nick = General_Field_Read(1, List2.List(List2.ListIndex), "(")

        If Nick = "" Then Exit Sub
        Nick = Left$(Nick, Len(Nick))

    End If

End Sub

Private Sub YoAcciones_Click(Index As Integer)

    Select Case Index

        Case 0 '/INVISIBLE 0.12.1
            Call WriteInvisible

        Case 1 'CHATCOLOR 0.12.1
            tmp = InputBox("Defina el color (R G B). Deje en blanco para usar el default.", "Cambiar color del chat")
            Call ParseUserCommand("/CHATCOLOR " & tmp)

        Case 2

    End Select
    
End Sub
