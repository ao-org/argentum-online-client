VERSION 5.00
Begin VB.Form FrmTorneo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Organizacion de evento"
   ClientHeight    =   6180
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   4428
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   4428
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraTorneosY 
      Caption         =   "Torneos y Eventos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   120
      TabIndex        =   34
      Top             =   120
      Width           =   4215
      Begin VB.OptionButton OptAbordaje 
         Caption         =   "Abordaje"
         Height          =   255
         Left            =   240
         TabIndex        =   84
         Top             =   2880
         Width           =   3735
      End
      Begin VB.OptionButton OptBusquedaDe 
         Caption         =   "Busqueda de tesoro"
         Height          =   255
         Left            =   240
         TabIndex        =   83
         Top             =   2520
         Width           =   3615
      End
      Begin VB.OptionButton OptBufones 
         Caption         =   "Bufones"
         Height          =   255
         Left            =   240
         TabIndex        =   82
         Top             =   2160
         Width           =   3135
      End
      Begin VB.CommandButton cmdCancelarTodos 
         Caption         =   "Cancelar todos los eventos con Lobby"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   81
         Top             =   5400
         Width           =   3495
      End
      Begin VB.OptionButton OptTorneo 
         Caption         =   "Torneo"
         Height          =   255
         Left            =   240
         TabIndex        =   78
         Top             =   1800
         Width           =   1935
      End
      Begin VB.OptionButton OptCapturaDe 
         Caption         =   "Captura de bandera"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   720
         Width           =   3615
      End
      Begin VB.OptionButton OptMatarCon 
         Caption         =   "Dia del Garrote"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   1080
         Width           =   3735
      End
      Begin VB.OptionButton OptElDe 
         Caption         =   "DeathMach"
         Height          =   255
         Left            =   240
         TabIndex        =   36
         Top             =   1440
         Width           =   3855
      End
      Begin VB.CommandButton cmdConfigurarE 
         Caption         =   "Configurar e Iniciar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   360
         TabIndex        =   35
         Top             =   4440
         Width           =   3495
      End
      Begin VB.Label lblSeleccionarEl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccionar el evento a realizar"
         Height          =   195
         Left            =   240
         TabIndex        =   39
         Top             =   360
         Width           =   2235
      End
   End
   Begin VB.Frame FraDeathMach 
      Caption         =   "DeathMach"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   120
      TabIndex        =   51
      Top             =   120
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton cmdVerAnotados 
         Caption         =   "Ver Anotados"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   79
         Top             =   5280
         Width           =   1335
      End
      Begin VB.CommandButton cmdAnunciarEldeath 
         Caption         =   "Anunciar el evento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   58
         Top             =   4320
         Width           =   3495
      End
      Begin VB.CommandButton cmdIniciarEldeath 
         Caption         =   "Iniciar el Evento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   57
         Top             =   4680
         Width           =   3495
      End
      Begin VB.CommandButton cmdCrearEldeath 
         Caption         =   "Crear el evento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   56
         Top             =   3960
         Width           =   3495
      End
      Begin VB.TextBox txtMinlvldeath 
         Height          =   285
         Left            =   1200
         TabIndex        =   55
         Text            =   "1"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtMaxlvldeath 
         Height          =   285
         Left            =   1200
         TabIndex        =   54
         Text            =   "47"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtPlayerDeath 
         Height          =   285
         Left            =   1200
         TabIndex        =   53
         Text            =   "1"
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelarDeach 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2880
         TabIndex        =   52
         Top             =   5280
         Width           =   990
      End
      Begin VB.Label lblNivelMinimodeath 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel Minimo"
         Height          =   195
         Left            =   120
         TabIndex        =   61
         Top             =   600
         Width           =   870
      End
      Begin VB.Label lblNivelMaximodeath 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel Maximo"
         Height          =   195
         Left            =   120
         TabIndex        =   60
         Top             =   960
         Width           =   930
      End
      Begin VB.Label lblCantidadDedeath 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad de participantes"
         Height          =   195
         Left            =   120
         TabIndex        =   59
         Top             =   1440
         Width           =   1845
      End
   End
   Begin VB.Frame FraCapturaDe 
      Caption         =   "Captura de Bandera"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   120
      TabIndex        =   62
      Top             =   120
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton Command5 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2880
         TabIndex        =   76
         Top             =   5520
         Width           =   990
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Crear el evento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   75
         Top             =   4200
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Iniciar el Evento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   74
         Top             =   4920
         Width           =   3495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Anunciar el evento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   73
         Top             =   4560
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.TextBox txtPrecio 
         Height          =   285
         Left            =   2040
         TabIndex        =   72
         Text            =   "1"
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txtNivelmax 
         Height          =   285
         Left            =   2040
         TabIndex        =   69
         Text            =   "47"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtNivelMinimo 
         Height          =   285
         Left            =   2040
         TabIndex        =   68
         Text            =   "1"
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox txtRondas 
         Height          =   285
         Left            =   2040
         TabIndex        =   66
         Text            =   "2"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtParticipantes 
         Height          =   285
         Left            =   2040
         TabIndex        =   63
         Text            =   "2"
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblPrecio 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Precio"
         Height          =   195
         Left            =   120
         TabIndex        =   71
         Top             =   2040
         Width           =   435
      End
      Begin VB.Label lblMaximo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "nivel maximo"
         Height          =   195
         Left            =   120
         TabIndex        =   70
         Top             =   1680
         Width           =   915
      End
      Begin VB.Label lblNivelMinimocapt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel Minimo"
         Height          =   195
         Left            =   120
         TabIndex        =   67
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lblCantDeRonda 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad de rondas"
         Height          =   195
         Left            =   120
         TabIndex        =   65
         Top             =   960
         Width           =   1410
      End
      Begin VB.Label lblCanPert 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad de participantes"
         Height          =   195
         Left            =   120
         TabIndex        =   64
         Top             =   600
         Width           =   1845
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dia del Garrote"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   120
      TabIndex        =   40
      Top             =   120
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton cmdVerAnotadosGarrote 
         Caption         =   "Ver Anotados"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   80
         Top             =   5280
         Width           =   1455
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2880
         TabIndex        =   50
         Top             =   5280
         Width           =   990
      End
      Begin VB.TextBox txtPlayer 
         Height          =   285
         Left            =   1200
         TabIndex        =   49
         Text            =   "1"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtNivelmaximo 
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   45
         Text            =   "47"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtNivelMinino 
         Height          =   285
         Left            =   1200
         TabIndex        =   44
         Text            =   "1"
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton cmdCrearEl 
         Caption         =   "Crear el evento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   43
         Top             =   3960
         Width           =   3495
      End
      Begin VB.CommandButton cmdIniciarEl 
         Caption         =   "Iniciar el Evento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   42
         Top             =   4680
         Width           =   3495
      End
      Begin VB.CommandButton cmdAnunciarEl 
         Caption         =   "Anunciar el evento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   41
         Top             =   4320
         Width           =   3495
      End
      Begin VB.Label lblCantidadDe 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad de participantes"
         Height          =   195
         Left            =   120
         TabIndex        =   48
         Top             =   1440
         Width           =   1845
      End
      Begin VB.Label lblNivelMaximo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel Maximo"
         Height          =   195
         Left            =   120
         TabIndex        =   47
         Top             =   960
         Width           =   930
      End
      Begin VB.Label lblNivelMinimo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel Minimo"
         Height          =   195
         Left            =   120
         TabIndex        =   46
         Top             =   600
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Torneos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton cmdCancelarTorneo 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2640
         TabIndex        =   77
         Top             =   5520
         Width           =   1335
      End
      Begin VB.CheckBox Check11 
         Caption         =   "Ladron"
         Height          =   195
         Left            =   2520
         TabIndex        =   33
         Top             =   3360
         Width           =   975
      End
      Begin VB.CheckBox Check10 
         Caption         =   "Pirata"
         Height          =   195
         Left            =   1440
         TabIndex        =   32
         Top             =   3360
         Width           =   975
      End
      Begin VB.CheckBox Check12 
         Caption         =   "Bandido"
         Height          =   195
         Left            =   240
         TabIndex        =   31
         Top             =   3360
         Width           =   975
      End
      Begin VB.TextBox reglas 
         Alignment       =   2  'Center
         Height          =   885
         Left            =   360
         ScrollBars      =   1  'Horizontal
         TabIndex        =   29
         Text            =   "Prohibido atacarse, tirar invisibilidad, etc"
         Top             =   4440
         Width           =   3375
      End
      Begin VB.TextBox nombre 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   360
         TabIndex        =   27
         Text            =   "Torneo 2vs 2"
         Top             =   3960
         Width           =   3375
      End
      Begin VB.TextBox y 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3240
         TabIndex        =   24
         Text            =   "50"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox x 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2520
         TabIndex        =   23
         Text            =   "50"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox map 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         TabIndex        =   21
         Text            =   "55"
         Top             =   720
         Width           =   615
      End
      Begin VB.CheckBox Check9 
         Caption         =   "Trabajador"
         Height          =   195
         Left            =   2520
         TabIndex        =   19
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Crear evento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   5520
         Width           =   2295
      End
      Begin VB.CheckBox Check8 
         Caption         =   "Cazador"
         Height          =   195
         Left            =   1440
         TabIndex        =   17
         Top             =   3000
         Width           =   975
      End
      Begin VB.CheckBox Check7 
         Caption         =   "Paladin"
         Height          =   195
         Left            =   2520
         TabIndex        =   16
         Top             =   3000
         Width           =   975
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Druida"
         Height          =   195
         Left            =   2520
         TabIndex        =   15
         Top             =   2640
         Width           =   975
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Bardo"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   3000
         Width           =   975
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Asesino"
         Height          =   195
         Left            =   1440
         TabIndex        =   13
         Top             =   2640
         Width           =   975
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Guerrero"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   2640
         Width           =   975
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Clerigo"
         Height          =   195
         Left            =   1440
         TabIndex        =   11
         Top             =   2280
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Mago"
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox costo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Text            =   "1000"
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox cupos 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Text            =   "8"
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox nivelmax 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Text            =   "50"
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox nivelmin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Text            =   "1"
         Top             =   450
         Width           =   975
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "Reglas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   4200
         Width           =   3375
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Nombre del evento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   3720
         Width           =   3375
      End
      Begin VB.Label Label9 
         Caption         =   "Y:"
         Height          =   255
         Left            =   3600
         TabIndex        =   26
         Top             =   840
         Width           =   255
      End
      Begin VB.Label Label8 
         Caption         =   "X:"
         Height          =   255
         Left            =   2640
         TabIndex        =   25
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Mapa"
         Height          =   255
         Left            =   2760
         TabIndex        =   22
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "Info de mapa"
         Height          =   255
         Left            =   2760
         TabIndex        =   20
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Clases"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   9
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Costo:"
         Height          =   255
         Left            =   780
         TabIndex        =   7
         Top             =   1580
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Cupos:"
         Height          =   255
         Left            =   750
         TabIndex        =   5
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Nivel Maximo:"
         Height          =   255
         Left            =   280
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Nivel Minimo:"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmTorneo"
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
Private Sub cmdAnunciarEl_Click()
Call ParseUserCommand("/configlobby open ")
End Sub

Private Sub cmdAnunciarEldeath_Click()
Call ParseUserCommand("/configlobby open ")
End Sub

Private Sub cmdCancelar_Click()
    FrmTorneo.Frame2.visible = False
    FrmTorneo.FraTorneosY.visible = True
End Sub

Private Sub cmdCancelarDeach_Click()
    FrmTorneo.FraTorneosY.visible = True
    FrmTorneo.FraDeathMach.visible = False
End Sub

Private Sub cmdCancelarTodos_Click()
    Call ParseUserCommand("/configlobby end")
End Sub

Private Sub cmdCancelarTorneo_Click()
    FrmTorneo.FraTorneosY.visible = True
    FrmTorneo.Frame1.visible = False
End Sub

Private Sub cmdConfigurarE_Click()
FrmTorneo.FraTorneosY.visible = False
    If OptMatarCon.Value = True Then
        FrmTorneo.Frame2.visible = True
    End If
    If OptElDe.Value = True Then
        FrmTorneo.FraDeathMach.visible = True
    End If
    If OptCapturaDe.Value = True Then
        FrmTorneo.FraCapturaDe.visible = True
    End If
        If OptTorneo.Value = True Then
        FrmTorneo.Frame1.visible = True
    End If
End Sub

Private Sub cmdCrearEl_Click()
Call ParseUserCommand("/crearevento caceria " & txtPlayer & " " & txtNivelMinino & " " & txtNivelmax)
End Sub

Private Sub cmdCrearEldeath_Click()
Call ParseUserCommand("/crearevento deathmatch " & txtPlayerDeath & " " & txtMinlvldeath & " " & txtMaxlvldeath)
End Sub

Private Sub cmdIniciarEl_Click()
Call ParseUserCommand("/configlobby start")
End Sub

Private Sub cmdIniciarEldeath_Click()
Call ParseUserCommand("/configlobby start")
End Sub

Private Sub cmdVerAnotados_Click()
    Call ParseUserCommand("/configlobby list")
End Sub

Private Sub cmdVerAnotadosGarrote_Click()
    Call ParseUserCommand("/configlobby list")
End Sub

Private Sub Command1_Click()
    
    On Error GoTo Command1_Click_Err
    
    Call WriteCreaerTorneo(nivelmin, nivelmax, cupos, costo.Text, Check1.Value, Check2.Value, Check3.Value, Check4.Value, Check5.Value, Check6.Value, Check7.Value, Check8.Value, Check9.Value, Check10.Value, Check11.Value, Check12.Value, map, x, y, nombre, reglas)

    
    Exit Sub

Command1_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmTorneo.Command1_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command3_Click()
    Call ParseUserCommand("/CREAREVENTO CAPTURA " & txtParticipantes & " " & txtRondas & " " & txtNivelMinimo & " " & txtNivelmax & " " & txtPrecio)
End Sub

Private Sub Command5_Click()
    FrmTorneo.FraCapturaDe.visible = False
    FrmTorneo.FraTorneosY.visible = True
End Sub

