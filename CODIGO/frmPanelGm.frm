VERSION 5.00
Begin VB.Form frmPanelgm 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Panel GM"
   ClientHeight    =   8610
   ClientLeft      =   18150
   ClientTop       =   4710
   ClientWidth     =   7200
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
   ScaleHeight     =   8610
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdButtonActualizarListaGms 
      BackColor       =   &H80000018&
      Caption         =   "Actualizar"
      Height          =   360
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton cmdInvisibilidadSi 
      BackColor       =   &H8000000A&
      Caption         =   "Invisibilidad Si/No"
      Height          =   360
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   6120
      Width           =   2295
   End
   Begin VB.CommandButton cmdMagiaSin 
      BackColor       =   &H8000000A&
      Caption         =   "Magia / Sin Magia"
      Height          =   360
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   5640
      Width           =   2295
   End
   Begin VB.CommandButton cmdRestringirMapa 
      BackColor       =   &H8000000A&
      Caption         =   "Restringir Mapa"
      Height          =   360
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   5160
      Width           =   2295
   End
   Begin VB.TextBox txtTextTriggers 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   51
      Text            =   "5"
      Top             =   4200
      Width           =   495
   End
   Begin VB.CommandButton cmdInsertarTrigger 
      BackColor       =   &H8000000A&
      Caption         =   "Insertar trigger´s"
      Height          =   360
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton cmdRecargarObjetos 
      BackColor       =   &H8000000A&
      Caption         =   "Recargar Objetos"
      Height          =   360
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   6600
      Width           =   2295
   End
   Begin VB.CommandButton cmdInseguro 
      BackColor       =   &H8000000A&
      Caption         =   "Inseguro"
      Height          =   360
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   4680
      Width           =   1110
   End
   Begin VB.CommandButton cmdSeguro 
      BackColor       =   &H8000000A&
      Caption         =   " Seguro"
      Height          =   360
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdMapeo 
      BackColor       =   &H8000000A&
      Caption         =   "Mapeo"
      Height          =   360
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton cmdBloqueoPos 
      BackColor       =   &H8000000A&
      Caption         =   "Bloquear/Desbloquear - Pos"
      Height          =   360
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton SendGlobal 
      BackColor       =   &H8000000A&
      Caption         =   "A GMs"
      Height          =   300
      Left            =   3900
      Style           =   1  'Graphical
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   4300
      Width           =   800
   End
   Begin VB.CommandButton cmdEscudo 
      Caption         =   "-"
      Height          =   360
      Index           =   1
      Left            =   1200
      TabIndex        =   43
      Top             =   5640
      Width           =   390
   End
   Begin VB.TextBox txtEscudo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   42
      Text            =   "0"
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton cmdEscudo 
      Caption         =   "+"
      Height          =   360
      Index           =   0
      Left            =   2520
      TabIndex        =   41
      Top             =   5640
      Width           =   375
   End
   Begin VB.CommandButton cmd 
      Caption         =   "-"
      Height          =   405
      Index           =   1
      Left            =   1200
      MaskColor       =   &H80000006&
      TabIndex        =   39
      Top             =   6120
      Width           =   375
   End
   Begin VB.TextBox txtArma 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1680
      TabIndex        =   38
      Text            =   "0"
      Top             =   6120
      Width           =   700
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00C0C0C0&
      Caption         =   "+"
      Height          =   405
      Index           =   0
      Left            =   2520
      MaskColor       =   &H80000006&
      TabIndex        =   37
      Top             =   6120
      Width           =   375
   End
   Begin VB.CommandButton cmdmenos 
      Caption         =   "-"
      Height          =   405
      Left            =   1200
      MaskColor       =   &H80000006&
      TabIndex        =   35
      Top             =   6600
      Width           =   375
   End
   Begin VB.CommandButton cmdmas 
      BackColor       =   &H8000000A&
      Caption         =   "+"
      Height          =   405
      Left            =   2520
      MaskColor       =   &H80000006&
      TabIndex        =   34
      Top             =   6600
      Width           =   375
   End
   Begin VB.TextBox txtCasco 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1680
      TabIndex        =   33
      Text            =   "0"
      Top             =   6600
      Width           =   700
   End
   Begin VB.CommandButton cmdMapaSeguro 
      BackColor       =   &H8000000A&
      Caption         =   "Info/Mapa"
      Height          =   360
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000A&
      Caption         =   "Destrabar"
      Height          =   375
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton cmdEventos 
      BackColor       =   &H8000000A&
      Caption         =   "Eventos"
      Height          =   360
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1590
   End
   Begin VB.CommandButton cmdIrCerca 
      BackColor       =   &H8000000A&
      Caption         =   "Ir Cerca"
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton cmdInformación 
      BackColor       =   &H8000000A&
      Caption         =   "Información General"
      Height          =   360
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   7560
      Width           =   1935
   End
   Begin VB.CommandButton cmdInvisible 
      BackColor       =   &H8000000A&
      Caption         =   "Invisible"
      Height          =   360
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdGuardarMapa 
      BackColor       =   &H8000000A&
      Caption         =   "Guardar Mapa"
      Height          =   360
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cmdMatarNPC 
      BackColor       =   &H8000000A&
      Caption         =   "Matar NPC"
      Height          =   360
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton cmdHeadMas 
      BackColor       =   &H00FFFFFF&
      Caption         =   "+"
      Height          =   405
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton cmdHeadMenos 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   405
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5160
      Width           =   375
   End
   Begin VB.CommandButton cmdHead0 
      BackColor       =   &H8000000A&
      Caption         =   "Head 0"
      Height          =   360
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5160
      Width           =   1590
   End
   Begin VB.TextBox txtHeadNumero 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1680
      TabIndex        =   17
      Text            =   "0"
      Top             =   5160
      Width           =   700
   End
   Begin VB.CommandButton cmdBody0 
      BackColor       =   &H8000000A&
      Caption         =   "Body 0"
      Height          =   360
      Index           =   2
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4680
      Width           =   1590
   End
   Begin VB.CommandButton cmdBodyMas 
      BackColor       =   &H00FFFFFF&
      Caption         =   "+"
      Height          =   405
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4680
      Width           =   375
   End
   Begin VB.CommandButton cmdBodyMenos 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   405
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4680
      Width           =   375
   End
   Begin VB.TextBox txtBodyYo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1680
      TabIndex        =   13
      Text            =   "0"
      Top             =   4680
      Width           =   700
   End
   Begin VB.CommandButton cmdConsulta 
      BackColor       =   &H8000000A&
      Caption         =   "/Consulta"
      Height          =   360
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton cmdModIntervalo 
      BackColor       =   &H8000000A&
      Caption         =   "/mod intervalo golpe"
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1695
   End
   Begin VB.TextBox txtMod 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   4320
      Width           =   3735
   End
   Begin VB.CommandButton cmdRevivir 
      BackColor       =   &H8000000A&
      Caption         =   "/Revivir"
      Height          =   360
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   7560
      Width           =   1215
   End
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
      Left            =   7320
      TabIndex        =   9
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
      Left            =   7320
      TabIndex        =   8
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton cmdCerrar 
      BackColor       =   &H8000000A&
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdSeleccionarPersonaje 
      BackColor       =   &H8000000A&
      Caption         =   "Seleccionar personaje"
      Height          =   375
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3360
      Width           =   1815
   End
   Begin VB.ListBox List1 
      BackColor       =   &H8000000A&
      Height          =   2010
      ItemData        =   "frmPanelGm.frx":0000
      Left            =   120
      List            =   "frmPanelGm.frx":0002
      TabIndex        =   3
      Top             =   110
      Width           =   4560
   End
   Begin VB.TextBox txtMsg 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      Height          =   1035
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2160
      Width           =   4575
   End
   Begin VB.CommandButton cmdActualiza 
      BackColor       =   &H8000000A&
      Caption         =   "&Actualiza"
      Height          =   315
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   7200
      Width           =   855
   End
   Begin VB.ComboBox cboListaUsus 
      BackColor       =   &H8000000A&
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
      TabIndex        =   1
      Top             =   7200
      Width           =   3675
   End
   Begin VB.ListBox List2 
      Height          =   1425
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   4560
   End
   Begin VB.Label lblEscudo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Escudo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   40
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label lblArma 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Arma"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   36
      Top             =   6120
      Width           =   735
   End
   Begin VB.Label lblHead 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Casco"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   32
      Top             =   6720
      Width           =   735
   End
   Begin VB.Label lblHead 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Body"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   29
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label lblHead 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Head"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   28
      Top             =   5160
      Width           =   885
   End
   Begin VB.Label lblDialogo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dialogo del GM + Enter"
      ForeColor       =   &H80000010&
      Height          =   195
      Left            =   120
      TabIndex        =   22
      Top             =   3960
      Width           =   1650
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
      Begin VB.Menu mnuDestrabar 
         Caption         =   "Destrabar"
      End
      Begin VB.Menu mnuBorrar 
         Caption         =   "Borrar mensaje"
      End
      Begin VB.Menu mnuIra 
         Caption         =   "Ir al usuario"
      End
      Begin VB.Menu mnuTraer 
         Caption         =   "Traer el usuario"
      End
      Begin VB.Menu mnuConsulta 
         Caption         =   "Consulta"
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
      Begin VB.Menu Investigar 
         Caption         =   "Investigar"
         Begin VB.Menu VerPantalla 
            Caption         =   "Ver pantalla cliente"
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
         Caption         =   "Torneos y Eventos"
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
            Caption         =   "Recolección x2 - 30 min"
         End
         Begin VB.Menu evento2 
            Caption         =   "Recolección x3 - 30 min"
         End
         Begin VB.Menu evento3 
            Caption         =   "Recolección x 5 - 30 min"
         End
         Begin VB.Menu personalizado 
            Caption         =   "Personalizado"
            Enabled         =   0   'False
            Visible         =   0   'False
         End
         Begin VB.Menu finalizarevento 
            Caption         =   "Finalizar actual"
         End
         Begin VB.Menu BusqedaTesoro 
            Caption         =   "Busqueda del tesoro"
            Enabled         =   0   'False
            Visible         =   0   'False
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
      Begin VB.Menu IP 
         Caption         =   "Direcciones de IP"
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
         Begin VB.Menu flash 
            Caption         =   "Flash"
         End
      End
      Begin VB.Menu usersOnline 
         Caption         =   "Usuarios Online"
      End
      Begin VB.Menu StaffOnline 
         Caption         =   "Staff Online"
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
Attribute VB_Name = "frmPanelgm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'    Argentum 20 - Game Client Program
'    Copyright (C) 2022 - Noland Studios
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
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'
Option Explicit

Dim Nick       As String

Dim tmp        As String

Dim tmpUser        As String

Public LastStr As String

Private Const MAX_GM_MSG = 300

Dim reason                      As Long

Private MisMSG(0 To MAX_GM_MSG) As String

Private Apunt(0 To MAX_GM_MSG)  As Integer

Public Sub CrearGMmSg(Nick As String, msg As String)
    
    On Error GoTo CrearGMmSg_Err
    

    If List1.ListCount < MAX_GM_MSG Then
        List1.AddItem Nick & "-" & List1.ListCount
        MisMSG(List1.ListCount - 1) = msg
        Apunt(List1.ListCount - 1) = List1.ListCount - 1

    End If

    
    Exit Sub

CrearGMmSg_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.CrearGMmSg", Erl)
    Resume Next
    
End Sub

Private Sub BanCuenta_Click()
    
    On Error GoTo BanCuenta_Click_Err
    
    tmp = InputBox("¿Motivo?", "Ingrese el motivo")
    Nick = cboListaUsus.Text

    If MsgBox("¿Estás seguro que desea banear la cuenta de """ & nick & """?", vbYesNo + vbQuestion) = vbYes Then
        Call WriteBanCuenta(Nick, tmp)

    End If

    
    Exit Sub

BanCuenta_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.BanCuenta_Click", Erl)
    Resume Next
    
End Sub

Private Sub BorrarPersonaje_Click()
    
    On Error GoTo BorrarPersonaje_Click_Err
    

    If MsgBox("¿Estás seguro que desea Borrar el personaje " & cboListaUsus.Text & "?", vbYesNo + vbQuestion) = vbYes Then

        Call ParseUserCommand("/KILLCHAR " & cboListaUsus.Text) ' ver ReyarB
    End If

    
    Exit Sub

BorrarPersonaje_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.BorrarPersonaje_Click", Erl)
    Resume Next
    
End Sub

Private Sub BusqedaTesoro_Click()
    
    On Error GoTo BusqedaTesoro_Click_Err
    

    tmp = InputBox("Ingrese tipo de evento:" & vbCrLf & "0: Busqueda de tesoro en continente" & vbCrLf & "1: Busqueda de tesoro en dungeon" & vbCrLf & "2: Aparicion de criatura", "Iniciar evento")

    If tmp >= 3 Or tmp = "" Then
        Exit Sub
    End If
    
    If IsNumeric(tmp) Then

        Call WriteBusquedaTesoro(CByte(tmp))
    Else
        MsgBox ("Tipo invalido")

    End If

    
    Exit Sub

BusqedaTesoro_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.BusqedaTesoro_Click", Erl)
    Resume Next
    
End Sub

Private Sub Cabeza_Click()
    
    On Error GoTo Cabeza_Click_Err
    
    tmp = InputBox("Ingrese el valor de cabeza que desea editar.", "Edicion de Usuarios")

    Call ParseUserCommand("/MOD " & cboListaUsus.Text & " Head " & tmp) 'ver ReyarB
    
    Exit Sub

Cabeza_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Cabeza_Click", Erl)
    Resume Next
    
End Sub

Private Sub CerrarleCliente_Click()
    
    On Error GoTo CerrarleCliente_Click_Err
    
    Call WriteCerraCliente(cboListaUsus.Text)

    
    Exit Sub

CerrarleCliente_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.CerrarleCliente_Click", Erl)
    Resume Next
    
End Sub

Private Sub CerrarProceso_Click()
    
    On Error GoTo CerrarProceso_Click_Err
    
    tmp = InputBox("Ingrese el nombre del proceso", "Cerrar Proceso")

    If tmp <> "" Then

        Call ParseUserCommand("/CERRARPROCESO " & cboListaUsus.Text & "@" & tmp)
    End If

    
    Exit Sub

CerrarProceso_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.CerrarProceso_Click", Erl)
    Resume Next
    
End Sub

Private Sub ciudadanos_Click()
    
    On Error GoTo ciudadanos_Click_Err
    
    tmp = InputBox("Ingrese el valor de ciudadanos que desea editar.", "Edicion de Usuarios")

    Call ParseUserCommand("/MOD " & cboListaUsus.Text & " CIU " & tmp) ' ver ReyarB
    
    Exit Sub

ciudadanos_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.ciudadanos_Click", Erl)
    Resume Next
    
End Sub

Private Sub Clase_Click()
    
    On Error GoTo Clase_Click_Err
    
    tmp = InputBox("Ingrese el valor de clase Libres que desea editar.", "Edicion de Usuarios")

    Call ParseUserCommand("/MOD " & cboListaUsus.Text & " CLASE " & tmp) 'ver ReyarB
    
    Exit Sub

Clase_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Clase_Click", Erl)
    Resume Next
    
End Sub

Private Sub cmd_Click(Index As Integer)
        
        On Error GoTo cmd_Click_Err
        
100     tmpUser = "yo"

102     Select Case Index
     
            Case 0
104             txtArma.Text = txtArma.Text + 1

106         Case 1
108             If txtArma.Text >= 1 Then txtArma.Text = txtArma.Text - 1
        End Select
    
110     tmp = txtArma.Text

112     Call ParseUserCommand("/MOD " & tmpUser & " Arma " & tmp)
    
114     Call frmPanelgm.txtMod.SetFocus
        
        Exit Sub

cmd_Click_Err:
        MsgBox Err.Description & vbCrLf & "in Argentum20.frmPanelgm.cmd_Click " & "at line " & Erl, vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        
End Sub

Private Sub cmdAccion_Click(Index As Integer)
    
    On Error GoTo 0
    
    Nick = Replace(cboListaUsus.Text, " ", "+")

    Select Case Index

        Case 0 '/ECHAR NICK 0.12.1
            Call WriteKick(Nick)

        Case 1 '/BAN NICK MOTIVO 0.12.1
            tmp = InputBox("¿Motivo?", "Ingrese el motivo")

            If MsgBox("¿Estás seguro que desea banear al personaje """ & cboListaUsus.Text & """?", vbYesNo + vbQuestion) = vbYes Then
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
            Call WriteServerTime

        Case 6 '/DONDE NICK 0.12.1

            If LenB(Nick) <> 0 Then Call WriteWhere(Nick)

        Case 7 '/NENE 0.12.1
            tmp = InputBox("¿En qué mapa?", "")
            Call ParseUserCommand("/NENE " & tmp)

        Case 8 '/info nick
            Call ParseUserCommand("/INFO " & Nick)
   
        Case 9 '/inv nick
            Call ParseUserCommand("/INV " & Nick)
   
        Case 10 '/skills nick
            Call ParseUserCommand("/SKILLS " & Nick)
   
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

        Case 17 '/BANIP IP 0.12.1
            tmp = InputBox("Escriba la dirección IP a banear.", "")
            reason = InputBox("Escriba el motivo del baneo.", "")

            If MsgBox("¿Estás seguro que deseas banear la IP """ & tmp & ", debido a " & reason & """?", vbYesNo + vbQuestion) = vbYes Then
                Call ParseUserCommand("/BANIP " & tmp & " " & reason)

            End If

        Case 18 '/bov nick

        Case 19 '/BANED IP AND PERSONAJE 0.12.1   REVISAR
    
            If MsgBox("¿Estás seguro que deseas banear la IP y el personaje """ & nick & """?", vbYesNo + vbQuestion) = vbYes Then
        
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

        Case 32
            Dim Elapsed As Single
            Elapsed = (FrameTime - HoraMundo) / DuracionDia
            Elapsed = (Elapsed - Fix(Elapsed)) * 24
        
            Dim HoraActual As Integer
            HoraActual = Fix(Elapsed)
            
            ' Es de noche?
            If HoraActual >= 0 And HoraActual <= 6 Then
                ' Hacemos de dia
                Call WriteDay
                ' Viceversa
            Else
                Call WriteNight
            End If

        Case 33

            Call ParseUserCommand("/PAUSAR") ' ver ReyarB

        Case 34 '/LIMPIARMUNDO 0.12.1
            Call WriteCleanWorld

        Case 35 '/SILENCIO NICK@TIEMPO

            tmp = InputBox("¿Minutos a silenciar? (hasta 255)", "")

            If MsgBox("¿Estás seguro que desea silenciar al personaje """ & nick & """?", vbYesNo + vbQuestion) = vbYes Then
                If tmp > 255 Then Exit Sub
                Call ParseUserCommand("/SILENCIAR " & cboListaUsus.Text & "@" & tmp)

            End If

    End Select

    Nick = ""
    
    Exit Sub

cmdAccion_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.cmdAccion_Click", Erl)
    Resume Next
    
End Sub

Private Sub cmdActualiza_Click()
    
    On Error GoTo cmdActualiza_Click_Err
    
    Call WriteRequestUserList
    
    Call frmPanelgm.txtMod.SetFocus
    
    Exit Sub

cmdActualiza_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.cmdActualiza_Click", Erl)
    Resume Next
    
End Sub

Private Sub cmdBloqueoPos_Click()
    Call WriteTileBlockedToggle
End Sub

Private Sub cmdBody0_Click(Index As Integer)

    tmpUser = "yo"

    Call ParseUserCommand("/MOD " & tmpUser & " Body 0")
    
    Call frmPanelgm.txtMod.SetFocus
    Exit Sub
End Sub

Private Sub cmdBodyMas_Click()

    tmpUser = "yo"
    
    txtBodyYo.Text = txtBodyYo.Text + 1
    
    tmp = txtBodyYo.Text
    

    Call ParseUserCommand("/MOD " & tmpUser & " Body " & tmp)
    
    Call frmPanelgm.txtMod.SetFocus
    Exit Sub
    
End Sub

Private Sub cmdBodyMenos_Click()

    tmpUser = "yo"
    
    If txtBodyYo.Text >= 1 Then txtBodyYo.Text = txtBodyYo.Text - 1
    
    tmp = txtBodyYo.Text
    

    Call ParseUserCommand("/MOD " & tmpUser & " Body " & tmp)
    
    Call frmPanelgm.txtMod.SetFocus
    Exit Sub
End Sub

Private Sub cmdcerrar_Click()
    
    On Error GoTo cmdcerrar_Click_Err
    

    Me.Visible = False
    List1.Clear
    List2.Clear
    txtMsg.Text = ""

    
    Exit Sub

cmdcerrar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.cmdcerrar_Click", Erl)
    Resume Next
    
End Sub

Private Sub cmdOnline_Click()

End Sub

Private Sub cmdTarget_Click()
    'Dim Usuaritio As String
    
    On Error GoTo cmdTarget_Click_Err
    

    'cboListaUsus = List1.List(List1.ListIndex)
    'Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre el personaje...", 100, 100, 120, 0, 0)
    'frmMain.MousePointer = 2
    'frmMain.PanelSelect = True
    'Call SendData("TGUSER")
    Call WriteMarcaDeGm

    
    Exit Sub

cmdTarget_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.cmdTarget_Click", Erl)
    Resume Next
    
End Sub

Private Sub cmdComandoGM_Click()
Call ParseUserCommand(txtMod.Text)
End Sub

Private Sub cmdConsulta_Click()

    tmpUser = cboListaUsus.Text
         
    Call ParseUserCommand("/CONSULTA " & tmpUser)
    Call frmPanelgm.txtMod.SetFocus
 
End Sub

Private Sub cmdEscudo_Click(Index As Integer)
                
100     tmpUser = "yo"

102     Select Case Index
     
            Case 0
104             txtEscudo.Text = txtEscudo.Text + 1

106         Case 1
108             If txtEscudo.Text >= 1 Then txtEscudo.Text = txtEscudo.Text - 1
        End Select
    
110     tmp = txtEscudo.Text

112     Call ParseUserCommand("/MOD " & tmpUser & " Escudo " & tmp)
    
114     Call frmPanelgm.txtMod.SetFocus
        
        Exit Sub
End Sub

Private Sub cmdEventos_Click()
    tmp = InputBox("Ingrese tipo de evento:" & vbCrLf & "0: Busqueda de tesoro en continente" & vbCrLf & "1: Busqueda de tesoro en dungeon" & vbCrLf & "2: Aparicion de criatura", "Iniciar evento")
    
    If IsNumeric(tmp) Then

        Call WriteBusquedaTesoro(CByte(tmp))
    Else
        MsgBox ("Tipo invalido")

    End If
    
    Call frmPanelgm.txtMod.SetFocus
    Exit Sub
End Sub

Private Sub cmdGuardarMapa_Click()
    Call ParseUserCommand("/GUARDAMAPA")
    Call frmPanelgm.txtMod.SetFocus
End Sub

Private Sub cmdHeadMenos_Click()

    tmpUser = "yo"
       
    If txtHeadNumero.Text >= 1 Then txtHeadNumero.Text = txtHeadNumero.Text - 1
    
    tmp = txtHeadNumero.Text
    

    Call ParseUserCommand("/MOD " & tmpUser & " Head " & tmp)
    
    Call frmPanelgm.txtMod.SetFocus
End Sub

Private Sub cmdHeadMas_Click()

    tmpUser = "yo"
       
    txtHeadNumero.Text = txtHeadNumero.Text + 1
    
    tmp = txtHeadNumero.Text
    

    Call ParseUserCommand("/MOD " & tmpUser & " Head " & tmp)
    
    Call frmPanelgm.txtMod.SetFocus
    Exit Sub
End Sub

Private Sub cmdHead0_Click()

    tmpUser = "yo"
    tmp = 0

    Call ParseUserCommand("/MOD " & tmpUser & " Head " & tmp)
    
    Call frmPanelgm.txtMod.SetFocus
    Exit Sub
    
End Sub

Private Sub cmdInformación_Click()
    
    tmpUser = cboListaUsus.Text
    
    Call WriteRequestCharStats(tmpUser)
    Call frmPanelgm.txtMod.SetFocus

End Sub

Private Sub cmdInseguro_Click()
Call ParseUserCommand("/MODMAPINFO SEGURO 0")
End Sub

Private Sub cmdInsertarTrigger_Click()

    Call ParseUserCommand("/TRIGGER " & txtTextTriggers.Text)

End Sub

Private Sub cmdInvisibilidadSi_Click()

    Call ParseUserCommand("/MODMAPINFO RESTRINGIR " & "SININVI")
    
End Sub

Private Sub cmdInvisible_Click()
    
    Call ParseUserCommand("/INVISIBLE")
        
    Call frmPanelgm.txtMod.SetFocus
    Exit Sub
    
End Sub

Private Sub cmdIrCerca_Click()

    tmpUser = cboListaUsus.Text
    Call WriteGoNearby(tmpUser)

End Sub

Private Sub cmdMagiaSin_Click()

    Call ParseUserCommand("/MODMAPINFO RESTRINGIR " & "SINMAGIA")
        
End Sub

Private Sub cmdMapaSeguro_Click()

    tmp = InputBox("Edicion de Mapa:" & vbCrLf & "0 : Informacion del Mapa" & vbCrLf & "1 : Pasar Mapa a Seguro" & vbCrLf & "2 : Pasar Mapa a InSeguro", "Modificar")
    
    Select Case tmp

        Case 0
            Call ParseUserCommand("/MAPINFO")

        Case 1
            Call ParseUserCommand("/MODMAPINFO SEGURO 1")

        Case 2
            Call ParseUserCommand("/MODMAPINFO SEGURO 0")
            
            
    End Select
    
End Sub

Private Sub cmdMapeo_Click()
    If frmPanelgm.Width = 7365 Then
        frmPanelgm.Width = 4860
    Else
        frmPanelgm.Width = 7365
    End If
End Sub

Private Sub cmdMas_Click()
    tmpUser = "yo"
       
    txtCasco.Text = txtCasco.Text + 1
        
    tmp = txtCasco.Text
        
    Call ParseUserCommand("/MOD " & tmpUser & " Casco " & tmp)
    
    Call frmPanelgm.txtMod.SetFocus
End Sub

Private Sub cmdMatarNPC_Click()

Call ParseUserCommand("/MATA")
    Call frmPanelgm.txtMod.SetFocus

End Sub

Private Sub cmdMenos_Click()
    tmpUser = "yo"
       
    If txtCasco.Text >= 1 Then txtCasco.Text = txtCasco.Text - 1
        
    tmp = txtCasco.Text
    
    
    Call ParseUserCommand("/MOD " & tmpUser & " Casco " & tmp)
    
    Call frmPanelgm.txtMod.SetFocus
End Sub

Private Sub cmdModIntervalo_Click()
      
    tmp = 1

Call ParseUserCommand("/MOD " & "yo" & " INTERVALO GOLPE " & tmp)

    Call frmPanelgm.txtMod.SetFocus
End Sub

Private Sub cmdRecargarObjetos_Click()
    Call WriteReloadObjects
End Sub

Private Sub cmdRestringirMapa_Click()

    Call ParseUserCommand("/MODMAPINFO RESTRINGIR " & "Newbie")
    Call ParseUserCommand("/MODMAPINFO RESTRINGIR " & "NoPKS")
    Call ParseUserCommand("/MODMAPINFO RESTRINGIR " & "NoCIUD")
    ' Wyrox Harthaos me falta Criminales , no se como restringir a todos de una
    ' luego de restringir
    ' faltaria mandar a cada uno a su hogar
    ' tambien los loguean mandarlos a su hogar.

End Sub

Private Sub cmdRevivir_Click()

    tmpUser = cboListaUsus.Text
    
    Call WriteReviveChar(tmpUser)
    Call frmPanelgm.txtMod.SetFocus
    
End Sub

Private Sub cmdSeguro_Click()
Call ParseUserCommand("/MODMAPINFO SEGURO 1")
End Sub

Private Sub cmdSeleccionarPersonaje_Click(Index As Integer)
    'cboListaUsus = List1.List(List1.ListIndex)
    'Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre el personaje...", 100, 100, 120, 0, 0)
    'frmMain.MousePointer = 2
    'frmMain.PanelSelect = True
    'Call SendData("TGUSER")
    Call WriteMarcaDeGm
    cboListaUsus = List1.List(List1.ListIndex)
    Call frmPanelgm.txtMod.SetFocus
    'txtHeadUser.Text = cboListaUsus.Text
End Sub

Private Sub Command1_Click()
    
    On Error GoTo Command1_Click_Err
    
    List1.Visible = True
    List2.Visible = False

    
    Exit Sub

Command1_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Command1_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command2_Click()
    
    On Error GoTo Command2_Click_Err
    
    List1.Visible = False
    List2.Visible = True

    
    Exit Sub

Command2_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Command2_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command3_Click()
    'Call WriteGoNearby(tmpUser)
    Call Destrabar_Click
    Call frmPanelgm.txtMod.SetFocus
End Sub

Private Sub CrearTeleport_Click()
    
    On Error GoTo CrearTeleport_Click_Err
    
    tmp = InputBox("Ingrese las cordenadas, por ejemplo para ulla: 1 50 50", "Ingrese Posiciones")
    Call ParseUserCommand("/CT " & tmp)

    
    Exit Sub

CrearTeleport_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.CrearTeleport_Click", Erl)
    Resume Next
    
End Sub

Private Sub creartoneo_Click()
    
    On Error GoTo creartoneo_Click_Err
    
    FrmTorneo.Show

    
    Exit Sub

creartoneo_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.creartoneo_Click", Erl)
    Resume Next
    
End Sub

Private Sub Criminales_Click()
    
    On Error GoTo Criminales_Click_Err
    
    tmp = InputBox("Ingrese el valor de criminales que desea editar.", "Edicion de Usuarios")

    Call ParseUserCommand("/MOD " & cboListaUsus.Text & " CRI " & tmp)
    
    Exit Sub

Criminales_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Criminales_Click", Erl)
    Resume Next
    
End Sub

Private Sub Cuerpo_Click()
    
    On Error GoTo Cuerpo_Click_Err
    
    tmp = InputBox("Ingrese el valor de cuerpo que desea editar.", "Edicion de Usuarios")

    Call ParseUserCommand("/MOD " & cboListaUsus.Text & " BODY " & tmp)
    
    Exit Sub

Cuerpo_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Cuerpo_Click", Erl)
    Resume Next
    
End Sub

Private Sub Desbanear_Click()
    
    On Error GoTo Desbanear_Click_Err
    
    tmp = InputBox("Escriba la dirección IP a desbanear", "")

    If MsgBox("¿Estás seguro que deseas desbanear la IP """ & tmp & """?", vbYesNo + vbQuestion) = vbYes Then
        Call ParseUserCommand("/UNBANIP " & tmp)

    End If

    
    Exit Sub

Desbanear_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Desbanear_Click", Erl)
    Resume Next
    
End Sub

Private Sub Destrabar_Click()
    
    On Error GoTo Destrabar_Click_Err
    
    Nick = Replace(List1.Text, " ", "+")
    Call WritePossUser(Nick)

    
    Exit Sub

Destrabar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Destrabar_Click", Erl)
    Resume Next
    
End Sub

Private Sub DestruirTeleport_Click()
    
    On Error GoTo DestruirTeleport_Click_Err
    
    Call WriteTeleportDestroy '/DT 0.12.1

    
    Exit Sub

DestruirTeleport_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.DestruirTeleport_Click", Erl)
    Resume Next
    
End Sub

Private Sub Ejecutar_Click()
    
    On Error GoTo Ejecutar_Click_Err
    
    Nick = cboListaUsus.Text
    Call WriteExecute(Nick) '/EJECUTAR NICK 0.12.1

    
    Exit Sub

Ejecutar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Ejecutar_Click", Erl)
    Resume Next
    
End Sub

Private Sub Energia_Click()
    
    On Error GoTo Energia_Click_Err
    
    tmp = InputBox("Ingrese el valor de energia que desea editar.", "Edicion de Usuarios")

    Call ParseUserCommand("/MOD " & cboListaUsus.Text & " EN " & tmp)
    
    Exit Sub

Energia_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Energia_Click", Erl)
    Resume Next
    
End Sub

Private Sub evento1_Click()
    
    On Error GoTo evento1_Click_Err
    
    Call WriteCrearEvento(3, 30, 2)

    
    Exit Sub

evento1_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.evento1_Click", Erl)
    Resume Next
    
End Sub

Private Sub evento2_Click()
    
    On Error GoTo evento2_Click_Err
    
    Call WriteCrearEvento(3, 30, 3)

    
    Exit Sub

evento2_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.evento2_Click", Erl)
    Resume Next
    
End Sub

Private Sub evento3_Click()
    
    On Error GoTo evento3_Click_Err
    
    Call WriteCrearEvento(3, 30, 5)

    
    Exit Sub

evento3_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.evento3_Click", Erl)
    Resume Next
    
End Sub

Private Sub evento4_Click()
    
    On Error GoTo evento4_Click_Err
    
    Call WriteCrearEvento(2, 30, 3)

    
    Exit Sub

evento4_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.evento4_Click", Erl)
    Resume Next
    
End Sub

Private Sub finalizarevento_Click()
    
    On Error GoTo finalizarevento_Click_Err
    

    If MsgBox("¿Estás seguro que deseas finalizar el evento?", vbYesNo + vbQuestion, "¡ATENCIÓN!") = vbYes Then
        Call WriteFinEvento
    End If

    
    Exit Sub

finalizarevento_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.finalizarevento_Click", Erl)
    Resume Next
    
End Sub

Private Sub flash_Click()
    Call WriteSetSpeed(15#)
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    List1.Clear
    List2.Clear
    txtMsg.Text = ""
    Call WriteRequestUserList
    
    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Form_Load", Erl)
    Resume Next
    
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Form_KeyPress_Err
    

    If (KeyAscii = 27) Then
        Unload Me

    End If

    
    Exit Sub

Form_KeyPress_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Form_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub lento_Click()

End Sub

Private Sub GlobalEstado_Click()

    Call ParseUserCommand("/ACTIVAR")
End Sub

Private Sub GuardarMapa_Click()

    Call ParseUserCommand("/BACK")
End Sub

Private Sub Limpiarmundo_Click()

    Call ParseUserCommand("/LIMPIARMUNDO")
End Sub

Private Sub LimpiarVision_Click()
    
    On Error GoTo LimpiarVision_Click_Err
    
    Call WriteDestroyAllItemsInArea

    
    Exit Sub

LimpiarVision_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.LimpiarVision_Click", Erl)
    Resume Next
    
End Sub

Private Sub List1_Click()
    
    On Error GoTo List1_Click_Err
    

    Dim ind As Integer

    ind = Val(ReadField(2, List1.List(List1.ListIndex), Asc("@")))
    txtMsg = List2.List(List1.ListIndex)
    
    
    Exit Sub

List1_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.List1_Click", Erl)
    Resume Next
    
End Sub

Private Sub List1_DblClick()
        tmpUser = Split(List1.Text, "(")(0)
        Call WriteGoNearby(tmpUser)
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo List1_MouseDown_Err
    

    If Button = vbRightButton Then
        PopUpMenu mnuUsuario

    End If

    
    Exit Sub

List1_MouseDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.List1_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub Mana_Click()
    
    On Error GoTo Mana_Click_Err
    
    tmp = InputBox("Ingrese el valor de mana que desea editar.", "Edicion de Usuarios")

    Call ParseUserCommand("/MOD " & cboListaUsus.Text & " MP " & tmp)
    
    Exit Sub

Mana_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Mana_Click", Erl)
    Resume Next
    
End Sub

Private Sub MensajeriaMenu_Click(Index As Integer)
    
    On Error GoTo MensajeriaMenu_Click_Err
    

    Select Case Index

        Case 0 'Mensaje por consola a usuarios 0.12.1
            tmp = InputBox("Ingrese el texto:", "Mensaje por consola a usuarios")
            If LenB(tmp) Then Call WriteServerMessage(tmp)

        Case 1 'Mensaje por ventana a usuarios 0.12.1
            tmp = InputBox("Ingrese el texto:", "Mensaje del sistema a usuarios")
            If LenB(tmp) Then Call WriteSystemMessage(tmp)

        Case 2 'Mensaje por consola a GMS 0.12.1
            tmp = InputBox("Escriba el mensaje.", "Mensaje por consola de GM")
            If LenB(tmp) Then Call WriteGMMessage(tmp)

        Case 3 'Hablar como NPC 0.12.1
            tmp = InputBox("Escriba un Mensaje.", "Hablar por NPC")
            If LenB(tmp) Then Call WriteTalkAsNPC(tmp)

    End Select

    
    Exit Sub

MensajeriaMenu_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.MensajeriaMenu_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuBorrar_Click()
    
    On Error GoTo mnuBorrar_Click_Err
    

    Dim elitem          As String
    Dim ProximamentTipo As String
    Dim TIPO            As String

    elitem = List1.ListIndex

    If List1.ListIndex < 0 Then Exit Sub
    
    Call ReadNick
    
    ProximamentTipo = General_Field_Read(2, List1.List(List1.ListIndex), "(")
    
    TIPO = General_Field_Read(1, ProximamentTipo, ")")
    
    Call WriteSOSRemove(nick & "Ø" & txtMsg & "Ø" & tipo)
    
    Call List1.RemoveItem(List1.ListIndex)
    Call List2.RemoveItem(elitem)
    
    txtMsg.Text = vbNullString

    
    Exit Sub

mnuBorrar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.mnuBorrar_Click", Erl)
    Resume Next
    
End Sub

Private Sub MnuEnviar_Click(Index As Integer)
    
    On Error GoTo MnuEnviar_Click_Err
    

    Dim Coordenadas As String

    Nick = Replace(cboListaUsus.Text, " ", "+")

    Select Case Index
            'ReyarB modifico cordenadas

         Case 0 'Ulla
            Coordenadas = "1 55 45"
            Call ParseUserCommand("/TELEP " & Nick & " " & Coordenadas)

        Case 1 'Nix
            Coordenadas = "34 40 85"
            Call ParseUserCommand("/TELEP " & Nick & " " & Coordenadas)

        Case 2 'Bander
            Coordenadas = "59 45 45"
            Call ParseUserCommand("/TELEP " & Nick & " " & Coordenadas)

        Case 3 'Arghal
            Coordenadas = "151 37 69"
            Call ParseUserCommand("/TELEP " & Nick & " " & Coordenadas)

        Case 4 'Otro

            If LenB(Nick) <> 0 Then
                Coordenadas = InputBox("Indique la posición (MAPA X Y).", "Transportar a " & nick)

                If LenB(Coordenadas) <> 0 Then Call ParseUserCommand("/TELEP " & Nick & " " & Coordenadas)

            End If

    End Select

    
    Exit Sub

MnuEnviar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.MnuEnviar_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuIRa_Click()
    
    On Error GoTo mnuIRa_Click_Err
    
    Call WriteGoToChar(ReadField(1, List1.List(List1.ListIndex), Asc("(")))

    
    Exit Sub

mnuIRa_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.mnuIRa_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuDestrabar_Click()
    On Error GoTo mnuDestrabar_Click_Err
    Nick = Replace(List1.Text, " ", "+")
    Call WritePossUser(Nick)

    
    Exit Sub

mnuDestrabar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.mnuDestrabar_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnutraer_Click()
    
    On Error GoTo mnutraer_Click_Err
    
    Call WriteSummonChar(ReadField(1, List1.List(List1.ListIndex), Asc("(")))

    
    Exit Sub

mnutraer_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.mnutraer_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuInvalida_Click()
    
    On Error GoTo mnuInvalida_Click_Err
    
    Nick = ReadField(1, List1.List(List1.ListIndex), Asc("("))
    
    Call ParseUserCommand("/MENSAJEINFORMACION " & Nick & "@" & "Su consulta fue rechazada debido a que esta fue catalogada como invalida.")

    ' Lo advertimos
    Call WriteWarnUser(nick, "Consulta a GM's inválida.")
    
    ' Borramos el mensaje de la lista.
    Call mnuBorrar_Click
    
    Exit Sub

mnuInvalida_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.mnuInvalida_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuResponder_Click()
    
    On Error GoTo mnuResponder_Click_Err
    
    Nick = ReadField(1, List1.List(List1.ListIndex), Asc("("))
    tmp = InputBox("Ingrese la respuesta:", "Responder consulta")
    Call ParseUserCommand("/MENSAJEINFORMACION " & Nick & "@" & tmp)

    
    Exit Sub

mnuResponder_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.mnuResponder_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuManual_Click()
    
    On Error GoTo mnuManual_Click_Err
    
    Nick = ReadField(1, List1.List(List1.ListIndex), Asc("("))
    Call ParseUserCommand("/MENSAJEINFORMACION " & nick & "@" & "Su consulta fue rechazada debido a que la respuesta se encuentra en el Manual o FAQ de nuestra pagina web. Para mas información visite: www.argentum20.com.ar.")

    
    Exit Sub

mnuManual_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.mnuManual_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuAccion_Click(Index As Integer)
    
    On Error GoTo mnuAccion_Click_Err
    
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

    
    Exit Sub

mnuAccion_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.mnuAccion_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuAdmin_Click(Index As Integer)
    
    On Error GoTo mnuAdmin_Click_Err
    
    Call cmdAccion_Click(Index)

    
    Exit Sub

mnuAdmin_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.mnuAdmin_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuAmbiente_Click(Index As Integer)
    
    On Error GoTo mnuAmbiente_Click_Err
    
    Call cmdAccion_Click(Index)

    
    Exit Sub

mnuAmbiente_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.mnuAmbiente_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuBan_Click(Index As Integer)
    
    On Error GoTo mnuBan_Click_Err
    
    Call cmdAccion_Click(Index)

    
    Exit Sub

mnuBan_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.mnuBan_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuCarcel_Click(Index As Integer)
    
    On Error GoTo mnuCarcel_Click_Err
    

    If Index = 60 Then
        Call cmdAccion_Click(11)
        Exit Sub

    End If

    Nick = cboListaUsus.Text

    Call ParseUserCommand("/CARCEL " & Nick & "@encarcelado via panelgm@" & Index)

    
    Exit Sub

mnuCarcel_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.mnuCarcel_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuSilencio_Click(Index As Integer)
    
    On Error GoTo mnuSilencio_Click_Err
    

    If Index = 60 Then
        Call cmdAccion_Click(35)
        Exit Sub

    End If

    Call ParseUserCommand("/SILENCIAR " & cboListaUsus.Text & "@" & Index)

    
    Exit Sub

mnuSilencio_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.mnuSilencio_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuHerramientas_Click(Index As Integer)
    
    On Error GoTo mnuHerramientas_Click_Err
    
    Call cmdAccion_Click(Index)

    
    Exit Sub

mnuHerramientas_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.mnuHerramientas_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuIP_Click(Index As Integer)
    
    On Error GoTo mnuIP_Click_Err
    
    Call cmdAccion_Click(Index)

    
    Exit Sub

mnuIP_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.mnuIP_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuReload_Click(Index As Integer)
    
    On Error GoTo mnuReload_Click_Err
    

    Select Case Index

        Case 1 'Reload objetos
            Call WriteReloadObjects

        Case 2 'Reload server.ini
            Call WriteReloadServerIni

        Case 3 'Reload mapas

            Call ParseUserCommand("/RELOAD MAP") 'Ver ReyarB
            
        Case 4 'Reload hechizos
            Call WriteReloadSpells

        Case 5 'Reload motd

            Call ParseUserCommand("/RELOADMOTD") ' ver ReyarB
        Case 6 'Reload npcs
            Call WriteReloadNPCs

        Case 7 'Reload sockets

             If MsgBox("Al realizar esta acción reiniciará la API de Winsock. Se cerrarán todas las conexiones.", vbYesNo, "Advertencia") = vbYes Then
               '   Call SendData("/RELOAD SOCK")
           End If

    Case 8 'Reload otros

        Call ParseUserCommand("/RELOADOPCIONES") 'ver ReyarB
End Select

    
    Exit Sub

mnuReload_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.mnuReload_Click", Erl)
    Resume Next
    
End Sub

Private Sub MOTD_Click()
    
    On Error GoTo MOTD_Click_Err
    
    Call WriteChangeMOTD 'Cambiar MOTD 0.12.1

    
    Exit Sub

MOTD_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.MOTD_Click", Erl)
    Resume Next
    
End Sub

Private Sub muyrapido_Click()
    
    On Error GoTo muyrapido_Click_Err
    
    Call WriteSetSpeed(5#)
    
    Exit Sub

muyrapido_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.muyrapido_Click", Erl)
    Resume Next
    
End Sub

Private Sub Normal_Click()
    
    On Error GoTo Normal_Click_Err
    
    Call WriteSetSpeed(1#)
    
    Exit Sub

Normal_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Normal_Click", Erl)
    Resume Next
    
End Sub

Private Sub oro_Click()
    
    On Error GoTo oro_Click_Err
    
    tmp = InputBox("Ingrese el valor de oro que desea editar.", "Edicion de Usuarios")

    Call ParseUserCommand("/MOD " & cboListaUsus.Text & " ORO " & tmp) ' ver ReyarB
    
    Exit Sub

oro_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.oro_Click", Erl)
    Resume Next
    
End Sub

Private Sub personalizado_Click()
    
    On Error GoTo personalizado_Click_Err
    
    tmp = InputBox("Ingrese evento  Tipo@Duracion@Multiplicacion" & vbCrLf & vbCrLf & "Tipo 1=Multiplica Oro" & vbCrLf & "Tipo 2=Multiplica Experiencia" & vbCrLf & "Tipo 3=Multiplica Recoleccion" & vbCrLf & "Tipo 4=Multiplica Dropeo" & vbCrLf & "Tipo 5=Multiplica Oro y Experiencia" & vbCrLf & "Tipo 6=Multiplica Oro, experiencia y recoleccion" & vbCrLf & "Tipo 7=Multiplica Todo" & vbCrLf & "Duracion= Maximo: 59" & vbCrLf & "Multiplicacion= Maximo 3", "Creacion de nuevo evento")
    Call ParseUserCommand("/CREAREVENTO " & tmp)

    
    Exit Sub

personalizado_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.personalizado_Click", Erl)
    Resume Next
    
End Sub

Private Sub rapido_Click()
    
    On Error GoTo rapido_Click_Err

    Call WriteSetSpeed(2#)

    Exit Sub

rapido_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.rapido_Click", Erl)
    Resume Next
    
End Sub

Private Sub Raza_Click()
    
    On Error GoTo Raza_Click_Err
    
    tmp = InputBox("Ingrese el valor de raza que desea editar.", "Edicion de Usuarios")

    Call ParseUserCommand("/MOD " & cboListaUsus.Text & " RAZA " & tmp) 'Ver ReyarB
    
    Exit Sub

Raza_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Raza_Click", Erl)
    Resume Next
    
End Sub

Private Sub ResetPozos_Click()

    Call ParseUserCommand("/RESETPOZOS") 'ver ReyarB
    
End Sub

Private Sub SeguroInseguro_Click()

    Call ParseUserCommand("/MODMAPINFO SEGURO 1")
End Sub

Private Sub SendGlobal_Click()
    If LenB(txtMod.Text) Then Call ParseUserCommand("/GMSG " & txtMod.Text)
    txtMod.Text = ""
    txtMod.SetFocus
End Sub

Private Sub SkillLibres_Click()
    
    On Error GoTo SkillLibres_Click_Err
    
    tmp = InputBox("Ingrese el valor de skills Libres que desea editar.", "Edicion de Usuarios")

    Call ParseUserCommand("/MOD " & cboListaUsus.Text & " SKILLSLIBRES " & tmp) ' ver ReyarB
    
    Exit Sub

SkillLibres_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.SkillLibres_Click", Erl)
    Resume Next
    
End Sub

Private Sub Spawn_Click()
    
    On Error GoTo Spawn_Click_Err
    
    Call WriteSpawnListRequest

    
    Exit Sub

Spawn_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Spawn_Click", Erl)
    Resume Next
    
End Sub

Private Sub StaffOnline_Click()
    
    On Error GoTo StaffOnline_Click_Err
    
    Call WriteOnlineGM '/ONLINEGM 0.12.1

    
    Exit Sub

StaffOnline_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.StaffOnline_Click", Erl)
    Resume Next
    
End Sub

Private Sub SubastaEstado_Click()

    Call ParseUserCommand("/SUBASTAACTIVADA") ' ver ReyarB
End Sub

Private Sub Temporal_Click()
    
    On Error GoTo Temporal_Click_Err

    Dim tmp  As String

    Dim tmp2 As Byte

    tmp2 = InputBox("¿Días?", "Ingrese cantidad de días (Maximo 255)")
    tmp = InputBox("¿Motivo?", "Ingrese el motivo")

    If MsgBox("¿Estás seguro que deseas banear el personaje de """ & cboListaUsus.Text & """ por " & tmp2 & " días?", vbYesNo + vbQuestion) = vbYes Then
        Call WriteBanTemporal(cboListaUsus.Text, tmp, tmp2)

    End If

    Exit Sub

Temporal_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Temporal_Click", Erl)
    Resume Next
    
End Sub

Private Sub cmdButtonActualizarListaGms_Click()
    cmdButtonActualizarListaGms.Enabled = False
    List1.Clear
    List2.Clear
    Call WriteSOSShowList
    cmdButtonActualizarListaGms.Enabled = True
End Sub

Private Sub torneo_cancelar_Click()
    
    On Error GoTo torneo_cancelar_Click_Err
    
    Call WriteCancelarTorneo
    Call ParseUserCommand("/configlobby end")
    
    Exit Sub

torneo_cancelar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.torneo_cancelar_Click", Erl)
    Resume Next
    
End Sub

Private Sub torneo_comenzar_Click()
    
    On Error GoTo torneo_comenzar_Click_Err
    
    Call WriteComenzarTorneo

    
    Exit Sub

torneo_comenzar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.torneo_comenzar_Click", Erl)
    Resume Next
    
End Sub

Private Sub txtArma_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(txtArma.Text & Chr(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0
End Sub

Private Sub txtArma_Change()
    Call ParseUserCommand("/MOD YO" & " Arma " & txtArma.Text)
End Sub


Private Sub txtBodyYo_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(txtBodyYo.Text & Chr(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0
End Sub

Private Sub txtBodyYo_Change()
    Call ParseUserCommand("/MOD YO" & " Body " & txtBodyYo.Text)
End Sub

Private Sub txtCasco_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(txtCasco.Text & Chr(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0
End Sub

Private Sub txtCasco_Change()
    Call ParseUserCommand("/MOD YO" & " Casco " & txtCasco.Text)
End Sub

Private Sub txtEscudo_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(txtEscudo.Text & Chr(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0
End Sub

Private Sub txtEscudo_Change()
    Call ParseUserCommand("/MOD YO" & " Escudo " & txtEscudo.Text)
End Sub

Private Sub txtHeadNumero_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(txtHeadNumero.Text & Chr(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0
End Sub

Private Sub txtHeadNumero_Change()
    Call ParseUserCommand("/MOD YO" & " Head " & txtHeadNumero.Text)
End Sub


Private Sub txtMod_KeyPress(KeyAscii As Integer)
    'If Not IsNumeric(txtMod.Text & Chr(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0
    
    If KeyAscii = vbKeyReturn Then
        Call ParseUserCommand(txtMod.Text)
        txtMod = ""
    End If
End Sub

Private Sub UnbanCuenta_Click()
    
    On Error GoTo UnbanCuenta_Click_Err
    
    Call WriteUnBanCuenta(cboListaUsus.Text)

    
    Exit Sub

UnbanCuenta_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.UnbanCuenta_Click", Erl)
    Resume Next
    
End Sub

Private Sub UnbanPersonaje_Click()
    
    On Error GoTo UnbanPersonaje_Click_Err
    
    Nick = cboListaUsus.Text

    If MsgBox("¿Estás seguro que deseas removerle el ban al personaje """ & nick & """?", vbYesNo + vbQuestion) = vbYes Then
        Call WriteUnbanChar(Nick)

    End If

    
    Exit Sub

UnbanPersonaje_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.UnbanPersonaje_Click", Erl)
    Resume Next
    
End Sub

Private Sub VerPantalla_Click()
    Call ParseUserCommand("/SS " & cboListaUsus.Text)
End Sub


Private Sub Vida_Click()
    
    On Error GoTo Vida_Click_Err
    
    tmp = InputBox("Ingrese el valor de vida que desea editar.", "Edicion de Usuarios")

    Call ParseUserCommand("/MOD " & cboListaUsus.Text & " HP " & tmp) 'ver ReyarB
    
    Exit Sub

Vida_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Vida_Click", Erl)
    Resume Next
    
End Sub

Private Sub ReadNick()
    
    On Error GoTo ReadNick_Err
    

    If List1.Visible Then
        Nick = General_Field_Read(1, List1.List(List1.ListIndex), "(")

        If Nick = "" Then Exit Sub
        Nick = Left$(Nick, Len(Nick))
    Else
        Nick = General_Field_Read(1, List2.List(List2.ListIndex), "(")

        If Nick = "" Then Exit Sub
        Nick = Left$(Nick, Len(Nick))

    End If

    
    Exit Sub

ReadNick_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.ReadNick", Erl)
    Resume Next
    
End Sub

Private Sub YoAcciones_Click(Index As Integer)
    
    On Error GoTo YoAcciones_Click_Err
    

    Select Case Index

        Case 0 '/INVISIBLE 0.12.1
            Call WriteInvisible

        Case 1 'CHATCOLOR 0.12.1
            tmp = InputBox("Defina el color (R G B). Deje en blanco para usar el default.", "Cambiar color del chat")
            Call ParseUserCommand("/CHATCOLOR " & tmp)

        Case 2

    End Select
    
    
    Exit Sub

YoAcciones_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.YoAcciones_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuConsulta_Click()
    
    Dim Nick As String
        Nick = ReadField(1, List1.List(List1.ListIndex), Asc("("))
    
    If Len(Nick) <> 0 Then
        
        Call WriteConsulta(Nick)
        
    End If

End Sub
