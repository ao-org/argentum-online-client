VERSION 5.00
Begin VB.Form frmPanelgm 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Panel GM"
   ClientHeight    =   8745
   ClientLeft      =   16095
   ClientTop       =   3480
   ClientWidth     =   7155
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
   ScaleHeight     =   8745
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   Begin VB.Frame MacrosCheat 
      BackColor       =   &H80000007&
      Height          =   4095
      Left            =   4800
      TabIndex        =   75
      Top             =   1440
      Width           =   2295
      Begin VB.TextBox txtSegundos 
         Height          =   285
         Left            =   1680
         TabIndex        =   86
         Text            =   "1.5"
         Top             =   720
         Width           =   375
      End
      Begin VB.CheckBox chkCoordenadas 
         BackColor       =   &H80000007&
         Caption         =   "Coordenadas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   240
         TabIndex        =   85
         Top             =   2160
         Width           =   255
      End
      Begin VB.CheckBox chkLeftClick 
         BackColor       =   &H80000001&
         Caption         =   "LeftClick"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   240
         Left            =   240
         TabIndex        =   84
         Top             =   1200
         Width           =   255
      End
      Begin VB.CheckBox chkUsarItem 
         BackColor       =   &H80000001&
         Caption         =   "UsarItem"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   240
         TabIndex        =   83
         Top             =   960
         Width           =   255
      End
      Begin VB.CheckBox chkOcultar 
         BackColor       =   &H80000001&
         Caption         =   "Ocultar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   240
         TabIndex        =   82
         Top             =   720
         Width           =   255
      End
      Begin VB.CheckBox chkPaquetes 
         BackColor       =   &H80000001&
         Caption         =   "Paquetes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   240
         TabIndex        =   81
         Top             =   1680
         Width           =   255
      End
      Begin VB.CheckBox chkAntiCheat 
         BackColor       =   &H80000007&
         Caption         =   "AntiCheat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   240
         TabIndex        =   80
         Top             =   240
         Width           =   255
      End
      Begin VB.CheckBox chkInasistido 
         BackColor       =   &H80000007&
         Caption         =   "Inasistido"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   240
         TabIndex        =   79
         Top             =   2640
         Width           =   255
      End
      Begin VB.CheckBox chkIRAUser 
         BackColor       =   &H80000007&
         Caption         =   "IRA User"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   240
         TabIndex        =   78
         Top             =   3720
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CheckBox chkClicks 
         BackColor       =   &H80000001&
         Caption         =   "Clicks"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   240
         TabIndex        =   77
         Top             =   2400
         Width           =   255
      End
      Begin VB.CheckBox chkCarteleo 
         BackColor       =   &H80000007&
         Caption         =   "Carteleo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   240
         TabIndex        =   76
         Top             =   2880
         Width           =   255
      End
      Begin VB.Label lblIraUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ira User"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   96
         Top             =   3720
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label lblCarteleo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Carteleo"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   2
         Left            =   600
         TabIndex        =   95
         Top             =   2880
         Width           =   930
      End
      Begin VB.Label lblInasistido 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inasistido"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   94
         Top             =   2640
         Width           =   930
      End
      Begin VB.Label lblClicks 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clicks"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   93
         Top             =   2400
         Width           =   390
      End
      Begin VB.Label lblInasistido 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cordenadas"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   92
         Top             =   2160
         Width           =   930
      End
      Begin VB.Label lblPaquetes 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Paquetes"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   91
         Top             =   1680
         Width           =   690
      End
      Begin VB.Label lblLeftClick 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LeftClick"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   90
         Top             =   1200
         Width           =   690
      End
      Begin VB.Label lblUsarItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "UsarItem"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   89
         Top             =   960
         Width           =   690
      End
      Begin VB.Label lblOcultar 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ocultar"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   88
         Top             =   720
         Width           =   690
      End
      Begin VB.Label lblAnticheat 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Anticheat"
         ForeColor       =   &H80000005&
         Height          =   195
         Index           =   0
         Left            =   600
         TabIndex        =   87
         Top             =   240
         Width           =   690
      End
   End
   Begin VB.Frame FraControlMacros 
      BackColor       =   &H80000001&
      ForeColor       =   &H00C0C0C0&
      Height          =   2535
      Left            =   120
      TabIndex        =   52
      Top             =   4680
      Width           =   4575
      Begin VB.CommandButton cmdInventario 
         BackColor       =   &H8000000A&
         Caption         =   "Inventario"
         Height          =   330
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdIrCerca 
         BackColor       =   &H8000000A&
         Caption         =   "Ir Cerca"
         Height          =   330
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdCerrarCliente 
         BackColor       =   &H8000000A&
         Caption         =   "Cerrar Cliente"
         Height          =   330
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cmdBanPJ 
         BackColor       =   &H8000000A&
         Caption         =   "Ban PJ"
         Height          =   330
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton cmdCarcel 
         BackColor       =   &H8000000A&
         Caption         =   "Carcel"
         Height          =   330
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton cmdBoveda 
         BackColor       =   &H8000000A&
         Caption         =   "Boveda"
         Height          =   330
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton cmdPenas 
         BackColor       =   &H8000000A&
         Caption         =   "Penas"
         Height          =   330
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton cmdEjecutar 
         BackColor       =   &H8000000A&
         Caption         =   "Ejecutar"
         Height          =   330
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton cmdEchar 
         BackColor       =   &H8000000A&
         Caption         =   "Echar"
         Height          =   330
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton cmdSTAT 
         BackColor       =   &H8000000A&
         Caption         =   "STAT"
         Height          =   330
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton cmdInfo 
         BackColor       =   &H8000000A&
         Caption         =   "Info"
         Height          =   330
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H8000000A&
         Caption         =   "Consulta"
         Height          =   330
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton cmdSUMUser 
         BackColor       =   &H8000000A&
         Caption         =   "Traer Usuario"
         Height          =   330
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton cmdBorrarInformes 
         BackColor       =   &H8000000A&
         Caption         =   "Borrar Informes"
         Height          =   690
         Left            =   3240
         Style           =   1  'Graphical
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CheckBox chkInfoTXT 
         BackColor       =   &H80000007&
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   3360
         TabIndex        =   55
         Top             =   480
         Width           =   255
      End
      Begin VB.CommandButton cmdSeguirMouse 
         BackColor       =   &H8000000A&
         Caption         =   "Seguir Mouse"
         Height          =   330
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   240
         Width           =   1455
      End
      Begin VB.CheckBox chkAutoName 
         BackColor       =   &H80000007&
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   3360
         TabIndex        =   53
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "GrabarTXT"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   3600
         TabIndex        =   72
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "AutoName"
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   3600
         TabIndex        =   71
         Top             =   260
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdTrabajando 
      BackColor       =   &H8000000A&
      Caption         =   "Actualizar Trabajadores"
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   69
      Top             =   3840
      Width           =   2175
   End
   Begin VB.CheckBox chkVerPanel 
      BackColor       =   &H00000000&
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   3240
      TabIndex        =   67
      Top             =   7680
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CommandButton cmdButtonActualizarListaGms 
      BackColor       =   &H80000018&
      Caption         =   "Actualizar"
      Height          =   360
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   51
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
      TabIndex        =   50
      Top             =   6600
      Width           =   2295
   End
   Begin VB.CommandButton cmdMagiaSin 
      BackColor       =   &H8000000A&
      Caption         =   "Magia / Sin Magia"
      Height          =   360
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   6120
      Width           =   2295
   End
   Begin VB.CommandButton cmdRestringirMapa 
      BackColor       =   &H8000000A&
      Caption         =   "Restringir Mapa"
      Height          =   360
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   5640
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
      TabIndex        =   47
      Text            =   "5"
      Top             =   4560
      Width           =   495
   End
   Begin VB.CommandButton cmdInsertarTrigger 
      BackColor       =   &H8000000A&
      Caption         =   "Insertar trigger´s"
      Height          =   360
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton cmdRecargarNPCs 
      BackColor       =   &H8000000A&
      Caption         =   "Recargar NPCs"
      Height          =   360
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   7080
      Width           =   2295
   End
   Begin VB.CommandButton cmdInseguro 
      BackColor       =   &H8000000A&
      Caption         =   "Inseguro"
      Height          =   360
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   5040
      Width           =   1110
   End
   Begin VB.CommandButton cmdSeguro 
      BackColor       =   &H8000000A&
      Caption         =   " Seguro"
      Height          =   360
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton cmdMapeo 
      BackColor       =   &H8000000A&
      Caption         =   "Mapeo"
      Height          =   360
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdBloqueoPos 
      BackColor       =   &H8000000A&
      Caption         =   "Bloquear/Desbloquear - Pos"
      Height          =   360
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton SendGlobal 
      BackColor       =   &H8000000A&
      Caption         =   "A GMs"
      Height          =   300
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   4320
      Width           =   800
   End
   Begin VB.CommandButton cmdEscudo 
      Caption         =   "-"
      Height          =   360
      Index           =   1
      Left            =   1200
      TabIndex        =   39
      Top             =   5760
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
      TabIndex        =   38
      Text            =   "0"
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton cmdEscudo 
      Caption         =   "+"
      Height          =   360
      Index           =   0
      Left            =   2520
      TabIndex        =   37
      Top             =   5760
      Width           =   375
   End
   Begin VB.CommandButton cmd 
      Caption         =   "-"
      Height          =   405
      Index           =   1
      Left            =   1200
      MaskColor       =   &H80000006&
      TabIndex        =   35
      Top             =   6240
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
      TabIndex        =   34
      Text            =   "0"
      Top             =   6240
      Width           =   700
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00C0C0C0&
      Caption         =   "+"
      Height          =   405
      Index           =   0
      Left            =   2520
      MaskColor       =   &H80000006&
      TabIndex        =   33
      Top             =   6240
      Width           =   375
   End
   Begin VB.CommandButton cmdmenos 
      Caption         =   "-"
      Height          =   405
      Left            =   1200
      MaskColor       =   &H80000006&
      TabIndex        =   31
      Top             =   6720
      Width           =   375
   End
   Begin VB.CommandButton cmdmas 
      BackColor       =   &H8000000A&
      Caption         =   "+"
      Height          =   405
      Left            =   2520
      MaskColor       =   &H80000006&
      TabIndex        =   30
      Top             =   6720
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
      TabIndex        =   29
      Text            =   "0"
      Top             =   6720
      Width           =   700
   End
   Begin VB.CommandButton cmdMapaSeguro 
      BackColor       =   &H8000000A&
      Caption         =   "Info/Mapa"
      Height          =   360
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000A&
      Caption         =   "Destrabar"
      Height          =   375
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cmdEventos 
      BackColor       =   &H8000000A&
      Caption         =   "Eventos"
      Height          =   360
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton cmdInformación 
      BackColor       =   &H8000000A&
      Caption         =   "Información General"
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   8040
      Width           =   1935
   End
   Begin VB.CommandButton cmdInvisible 
      BackColor       =   &H8000000A&
      Caption         =   "Invisible"
      Height          =   360
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton cmdMatarNPC 
      BackColor       =   &H8000000A&
      Caption         =   "Matar NPC"
      Height          =   360
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton cmdHeadMas 
      BackColor       =   &H00FFFFFF&
      Caption         =   "+"
      Height          =   405
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5280
      Width           =   375
   End
   Begin VB.CommandButton cmdHeadMenos 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   405
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5280
      Width           =   375
   End
   Begin VB.CommandButton cmdHead0 
      BackColor       =   &H8000000A&
      Caption         =   "Head 0"
      Height          =   360
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5280
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
      TabIndex        =   16
      Text            =   "0"
      Top             =   5280
      Width           =   700
   End
   Begin VB.CommandButton cmdBody0 
      BackColor       =   &H8000000A&
      Caption         =   "Body 0"
      Height          =   360
      Index           =   2
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1590
   End
   Begin VB.CommandButton cmdBodyMas 
      BackColor       =   &H00FFFFFF&
      Caption         =   "+"
      Height          =   405
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4800
      Width           =   375
   End
   Begin VB.CommandButton cmdBodyMenos 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      Height          =   405
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4800
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
      TabIndex        =   12
      Text            =   "0"
      Top             =   4800
      Width           =   700
   End
   Begin VB.CommandButton cmdConsulta 
      BackColor       =   &H8000000A&
      Caption         =   "/Consulta"
      Height          =   360
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6720
      Width           =   1200
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
      Width           =   3615
   End
   Begin VB.CommandButton cmdRevivir 
      BackColor       =   &H8000000A&
      Caption         =   "/Revivir"
      Height          =   360
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   8040
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
      Caption         =   "/ir A"
      Height          =   320
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7320
      Width           =   855
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
      Caption         =   "Actualizar Usuarios"
      Height          =   360
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3840
      Width           =   2295
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
      Top             =   7320
      Width           =   3675
   End
   Begin VB.ListBox List2 
      Height          =   1425
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   4560
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Ver panel /MOD "
      ForeColor       =   &H00C0C0C0&
      Height          =   200
      Left            =   3500
      TabIndex        =   70
      Top             =   7710
      Width           =   1215
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
      Left            =   0
      TabIndex        =   36
      Top             =   5760
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
      TabIndex        =   32
      Top             =   6240
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
      TabIndex        =   28
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
      Left            =   240
      TabIndex        =   25
      Top             =   4920
      Width           =   735
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
      Left            =   120
      TabIndex        =   24
      Top             =   5400
      Width           =   885
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

Dim nick       As String

Dim tmp        As String

Dim tmpUser        As String

Dim Resultado  As Boolean

Public ContMacro  As Integer

Public LastStr As String

Private Const MAX_GM_MSG = 300

Dim reason                      As Long

Private MisMSG(0 To MAX_GM_MSG) As String

Private Apunt(0 To MAX_GM_MSG)  As Integer

Public Sub CrearGMmSg(nick As String, msg As String)
    On Error Goto CrearGMmSg_Err
    
    On Error GoTo CrearGMmSg_Err
    

    If List1.ListCount < MAX_GM_MSG Then
        List1.AddItem nick & "-" & List1.ListCount
        MisMSG(List1.ListCount - 1) = msg
        Apunt(List1.ListCount - 1) = List1.ListCount - 1

    End If

    
    Exit Sub

CrearGMmSg_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.CrearGMmSg", Erl)
    Resume Next
    
    Exit Sub
CrearGMmSg_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.CrearGMmSg", Erl)
End Sub

Private Sub BanCuenta_Click()
    On Error Goto BanCuenta_Click_Err
    
    On Error GoTo BanCuenta_Click_Err
    
    tmp = InputBox(JsonLanguage.Item("INPUTBOX_MOTIVO"), JsonLanguage.Item("INPUTBOX_TITULO"))
    nick = cboListaUsus.Text

    If MsgBox(JsonLanguage.Item("MENSAJEBOX_BANEAR_CUENTA") & " " & nick, vbYesNo + vbQuestion) = vbYes Then
        Call WriteBanCuenta(nick, tmp)

    End If

    
    Exit Sub

BanCuenta_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.BanCuenta_Click", Erl)
    Resume Next
    
    Exit Sub
BanCuenta_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.BanCuenta_Click", Erl)
End Sub

Private Sub BorrarPersonaje_Click()
    On Error Goto BorrarPersonaje_Click_Err
    
    On Error GoTo BorrarPersonaje_Click_Err
    

    If MsgBox(JsonLanguage.Item("MENSAJEBOX_BORRAR_PERSONAJE") & " " & cboListaUsus.Text, vbYesNo + vbQuestion) = vbYes Then


        Call ParseUserCommand("/KILLCHAR " & cboListaUsus.Text)
    End If

    
    Exit Sub

BorrarPersonaje_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.BorrarPersonaje_Click", Erl)
    Resume Next
    
    Exit Sub
BorrarPersonaje_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.BorrarPersonaje_Click", Erl)
End Sub

Private Sub BusqedaTesoro_Click()
    On Error Goto BusqedaTesoro_Click_Err
    
    On Error GoTo BusqedaTesoro_Click_Err
    

    tmp = InputBox(JsonLanguage.Item("MENSAJE_INGRESAR_TIPO_EVENTO") & vbCrLf & JsonLanguage.Item("MENSAJE_EVENTO_TESORO_CONTINENTE") & vbCrLf & JsonLanguage.Item("MENSAJE_EVENTO_TESORO_DUNGEON") & vbCrLf & JsonLanguage.Item("MENSAJE_EVENTO_APARICION_CRIATURA"), JsonLanguage.Item("MENSAJE_INICIAR_EVENTO"))

    If tmp >= 3 Or tmp = "" Then
        Exit Sub
    End If
    
    If IsNumeric(tmp) Then

        Call WriteBusquedaTesoro(CByte(tmp))
    Else
        MsgBox JsonLanguage.Item("MENSAJE_TIPO_INVALIDO"), vbExclamation, JsonLanguage.Item("TITULO_ERROR")

    End If

    
    Exit Sub

BusqedaTesoro_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.BusqedaTesoro_Click", Erl)
    Resume Next
    
    Exit Sub
BusqedaTesoro_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.BusqedaTesoro_Click", Erl)
End Sub

Private Sub Cabeza_Click()
    On Error Goto Cabeza_Click_Err
    
    On Error GoTo Cabeza_Click_Err
    
    tmp = InputBox(JsonLanguage.Item("MENSAJE_INGRESAR_VALOR_CABEZA"), JsonLanguage.Item("MENSAJE_EDICION_USUARIOS"))

    Call ParseUserCommand("/MOD " & cboListaUsus.Text & " Head " & tmp)
    
    Exit Sub

Cabeza_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Cabeza_Click", Erl)
    Resume Next
    
    Exit Sub
Cabeza_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.Cabeza_Click", Erl)
End Sub

Private Sub CerrarleCliente_Click()
    On Error Goto CerrarleCliente_Click_Err
    
    On Error GoTo CerrarleCliente_Click_Err
    
    Call WriteCerraCliente(cboListaUsus.Text)

    
    Exit Sub

CerrarleCliente_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.CerrarleCliente_Click", Erl)
    Resume Next
    
    Exit Sub
CerrarleCliente_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.CerrarleCliente_Click", Erl)
End Sub

Private Sub CerrarProceso_Click()
    On Error Goto CerrarProceso_Click_Err
    
    On Error GoTo CerrarProceso_Click_Err
    
    tmp = InputBox(JsonLanguage.Item("MENSAJE_INGRESAR_NOMBRE_PROCESO"), JsonLanguage.Item("MENSAJE_CERRAR_PROCESO"))

    If tmp <> "" Then

        Call ParseUserCommand("/CERRARPROCESO " & cboListaUsus.Text & "@" & tmp)
    End If

    
    Exit Sub

CerrarProceso_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.CerrarProceso_Click", Erl)
    Resume Next
    
    Exit Sub
CerrarProceso_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.CerrarProceso_Click", Erl)
End Sub

Private Sub chkAntiCheat_Click()
    On Error Goto chkAntiCheat_Click_Err
    If chkAntiCheat.Value = 0 Then
        chkOcultar.Value = False
        chkUsarItem.Value = False
        chkLeftClick.Value = False
        chkPaquetes.Value = False
        chkCoordenadas.Value = False
        chkClicks.Value = False
        chkInasistido.Value = False
        chkCarteleo.Value = False
    Else
        chkOcultar.Value = 1
        chkUsarItem.Value = 1
        chkLeftClick.Value = 1
        chkPaquetes.Value = 1
        chkCoordenadas.Value = 1
        chkClicks.Value = 1
        chkInasistido.Value = 1
        chkCarteleo.Value = 1
    End If
    Exit Sub
chkAntiCheat_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.chkAntiCheat_Click", Erl)
End Sub

Private Sub chkVerPanel_Click()
    On Error Goto chkVerPanel_Click_Err
    If chkVerPanel.Value = 1 Then
        FraControlMacros.visible = False
    Else
        FraControlMacros.visible = True
    End If
    Exit Sub
chkVerPanel_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.chkVerPanel_Click", Erl)
End Sub

Private Sub ciudadanos_Click()
    On Error Goto ciudadanos_Click_Err
    
    On Error GoTo ciudadanos_Click_Err
    
    tmp = InputBox("Ingrese el valor de ciudadanos que desea editar.", JsonLanguage.Item("MENSAJE_EDICION_USUARIOS"))

    Call ParseUserCommand("/MOD " & cboListaUsus.Text & " CIU " & tmp)
    
    Exit Sub

ciudadanos_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.ciudadanos_Click", Erl)
    Resume Next
    
    Exit Sub
ciudadanos_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.ciudadanos_Click", Erl)
End Sub

Private Sub Clase_Click()
    On Error Goto Clase_Click_Err
    
    On Error GoTo Clase_Click_Err
    
    tmp = InputBox("Ingrese el valor de clase Libres que desea editar.", JsonLanguage.Item("MENSAJE_EDICION_USUARIOS"))

    Call ParseUserCommand("/MOD " & cboListaUsus.Text & " CLASE " & tmp)
    
    Exit Sub

Clase_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Clase_Click", Erl)
    Resume Next
    
    Exit Sub
Clase_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.Clase_Click", Erl)
End Sub

Private Sub cmd_Click(Index As Integer)
    On Error Goto cmd_Click_Err
        
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
        
    Exit Sub
cmd_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmd_Click", Erl)
End Sub

Private Sub cmdAccion_Click(Index As Integer)
    On Error Goto cmdAccion_Click_Err
    
    On Error GoTo 0
    
    nick = Replace(cboListaUsus.Text, " ", "+")

    Select Case Index

        Case 0 '/ECHAR NICK 0.12.1
            Call WriteKick(nick)

        Case 1 '/BAN NICK MOTIVO 0.12.1
            tmp = InputBox(JsonLanguage.Item("MENSAJE_MOTIVO"), JsonLanguage.Item("TITULO_MOTIVO"))

            If MsgBox(JsonLanguage.Item("MENSAJEBOX_BANEAR_PERSONAJE") & " " & cboListaUsus.Text, vbYesNo + vbQuestion) = vbYes Then
                Call WriteBanChar(nick, tmp)

            End If

        Case 2 '/SUM NICK 0.12.1

            If LenB(nick) <> 0 Then Call WriteSummonChar(nick)

        Case 3 '/ira NICK 0.12.1

            If LenB(nick) <> 0 Then Call WriteGoToChar(nick)

        Case 4 '/REM 0.12.1
            tmp = InputBox("¿Comentario?", "Ingrese comentario")
            Call WriteComment(tmp)

        Case 5 '/HORA 0.12.1
            Call WriteServerTime

        Case 6 '/DONDE NICK 0.12.1

            If LenB(nick) <> 0 Then Call WriteWhere(nick)

        Case 7 '/NENE 0.12.1
            tmp = InputBox("¿En qué mapa?", "")
            Call ParseUserCommand("/NENE " & tmp)

        Case 8 '/info nick
            Call ParseUserCommand("/INFO " & nick)
   
        Case 9 '/inv nick
            Call ParseUserCommand("/INV " & nick)
   
        Case 10 '/skills nick
            Call ParseUserCommand("/SKILLS " & nick)
   
        Case 11 '/CARCEL NICK @ MOTIVO  0.12.1
            tmp = InputBox("¿Minutos a encarcelar? (hasta 60)", "")

            If tmp > 1 Then
                Call ParseUserCommand("/CARCEL " & nick & "@encarcelado via panelgm@" & tmp)
           
            Else
                MsgBox JsonLanguage.Item("MENSAJE_TIEMPO_INVALIDO")
            
            End If

        Case 13 '/nick2ip NICK 0.12.1
            Call WriteNickToIP(nick)

        Case 14 '/Lastip NICK 0.12.1
            Call WriteLastIP(nick)

        Case 15 '/IrCerca NICK 0.12.1

            If LenB(nick) <> 0 Then Call WriteGoNearby(nick)

        Case 17 '/BANIP IP 0.12.1
            Call ShowConsoleMsg("Not supported.")
        Case 18 '/bov nick

        Case 19 '/BANED IP AND PERSONAJE 0.12.1   REVISAR
    
            Call ShowConsoleMsg("Not supported.")

        Case 20 '/PENAS NICK 0.12.1
            Call WritePunishments(nick)

        Case 21 '/REVIVIR NICK 0.12.1
            Call WriteReviveChar(nick)

        Case 22 'ADVERTENCIA 0.12.1
            tmp = InputBox("Escriba el motivo de la advertencia.", "Advertir a " & nick)

            If LenB(tmp) <> 0 Then
                Call ParseUserCommand("/ADVERTENCIA " & nick & "@" & tmp)

            End If

        Case 23 '/TRABAJANDO 0.12.1
            Call WriteWorking

        Case 25 '/BANIPLIST 0.12.1
            Call ShowConsoleMsg("Not supported.")

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

            Call ParseUserCommand("/PAUSAR")

        Case 34 '/LIMPIARMUNDO 0.12.1
            Call WriteCleanWorld

        Case 35 '/SILENCIO NICK@TIEMPO

            tmp = InputBox("¿Minutos a silenciar? (hasta 255)", "")

            If MsgBox(JsonLanguage.Item("MENSAJE_SILENCIAR_PERSONAJE") & nick & """?", vbYesNo + vbQuestion) = vbYes Then
                If tmp > 255 Then Exit Sub
                Call ParseUserCommand("/SILENCIAR " & cboListaUsus.Text & "@" & tmp)

            End If

    End Select

    nick = ""
    
    Exit Sub

cmdAccion_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.cmdAccion_Click", Erl)
    Resume Next
    
    Exit Sub
cmdAccion_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdAccion_Click", Erl)
End Sub

Private Sub cmdActualiza_Click()
    On Error Goto cmdActualiza_Click_Err
    
    On Error GoTo cmdActualiza_Click_Err
    
    Call WriteRequestUserList
    
    Call frmPanelgm.txtMod.SetFocus
    
    Exit Sub

cmdActualiza_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.cmdActualiza_Click", Erl)
    Resume Next
    
    Exit Sub
cmdActualiza_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdActualiza_Click", Erl)
End Sub

Private Sub cmdCerrarCliente_Click()
    On Error Goto cmdCerrarCliente_Click_Err
    tmpUser = cboListaUsus.Text
    Call ParseUserCommand("/SM ")
    Call frmPanelgm.txtMod.SetFocus
    Call WriteGoNearby(tmpUser)
    Call WriteCerraCliente(tmpUser)
    Exit Sub
cmdCerrarCliente_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdCerrarCliente_Click", Erl)
End Sub

Private Sub cmdBanPJ_Click()
    On Error Goto cmdBanPJ_Click_Err
    tmpUser = cboListaUsus.Text
    Call ParseUserCommand("/BAN")
    tmp = InputBox("Escriba el motivo del BAN.", "Baneo de " & tmpUser)

    If tmp = "" Then
        InputBox ("No se puede bannear si dar motivos a " & tmpUser)
    Else
        Call WriteBanChar(tmpUser, tmp)
    End If
    Exit Sub
cmdBanPJ_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdBanPJ_Click", Erl)
End Sub

Private Sub cmdBloqueoPos_Click()
    On Error Goto cmdBloqueoPos_Click_Err
    Call WriteTileBlockedToggle
    Exit Sub
cmdBloqueoPos_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdBloqueoPos_Click", Erl)
End Sub

Private Sub cmdBody0_Click(Index As Integer)
    On Error Goto cmdBody0_Click_Err

    tmpUser = "yo"

    Call ParseUserCommand("/MOD " & tmpUser & " Body 0")
    
    Call frmPanelgm.txtMod.SetFocus
    Exit Sub
    Exit Sub
cmdBody0_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdBody0_Click", Erl)
End Sub

Private Sub cmdBodyMas_Click()
    On Error Goto cmdBodyMas_Click_Err

    tmpUser = "yo"
    
    txtBodyYo.Text = txtBodyYo.Text + 1
    
    tmp = txtBodyYo.Text
    

    Call ParseUserCommand("/MOD " & tmpUser & " Body " & tmp)
    
    Call frmPanelgm.txtMod.SetFocus
    Exit Sub
    
    Exit Sub
cmdBodyMas_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdBodyMas_Click", Erl)
End Sub

Private Sub cmdBodyMenos_Click()
    On Error Goto cmdBodyMenos_Click_Err

    tmpUser = "yo"
    
    If txtBodyYo.Text >= 1 Then txtBodyYo.Text = txtBodyYo.Text - 1
    
    tmp = txtBodyYo.Text
    

    Call ParseUserCommand("/MOD " & tmpUser & " Body " & tmp)
    
    Call frmPanelgm.txtMod.SetFocus
    Exit Sub
    Exit Sub
cmdBodyMenos_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdBodyMenos_Click", Erl)
End Sub

Private Sub cmdBorrarInformes_Click()
    On Error Goto cmdBorrarInformes_Click_Err
    Dim ruta As String
    Dim archivos As Variant
    Dim i As Integer

    ' Obtén la ruta del directorio del ejecutable
    ruta = App.path

    ' Definir los nombres de los archivos
    archivos = Array("MacroOcultar.txt", "MacroUseItemU.txt", "MacroUseItem.txt", "MacroGuildMessage.txt", "MacroLeftClick.txt", "MacroChangeHeading.txt", "MacroCoordenadas.txt", "MacroDeClick.txt", "MacroInasistido.txt", "MacroCarteleo.txt", "MacroDePaquetes.txt", "MacroTotal.txt")

    ' Verificar y eliminar cada archivo en la lista
    For i = LBound(archivos) To UBound(archivos)
        If dir(ruta & "\" & archivos(i)) <> "" Then
            Kill ruta & "\" & archivos(i)
        End If
    Next i
    Exit Sub
cmdBorrarInformes_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdBorrarInformes_Click", Erl)
End Sub


Private Sub cmdBoveda_Click()
    On Error Goto cmdBoveda_Click_Err
    tmpUser = cboListaUsus.Text
    Call WriteRequestCharBank(tmpUser)
    Call frmPanelgm.txtMod.SetFocus

    Exit Sub
cmdBoveda_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdBoveda_Click", Erl)
End Sub

Private Sub cmdCarcel_Click()
    On Error Goto cmdCarcel_Click_Err
    Dim tmp As String
    Dim tmptime As String
    tmpUser = cboListaUsus.Text
    tmp = InputBox("Escriba el motivo de Carcel .", "Carcel a " & TargetName)
    tmptime = InputBox("Escriba el tiempo de Carcel .", "Tiempo de Carcel a " & TargetName)
    If tmp = "" Or tmptime = "" Then
        MsgBox JsonLanguage.Item("MENSAJE_FALTAN_DATOS"), vbExclamation, "Error"
    Else
        Call WriteJail(tmpUser, tmp, tmptime)
    End If
    Exit Sub
cmdCarcel_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdCarcel_Click", Erl)
End Sub

Private Sub cmdcerrar_Click()
    On Error Goto cmdcerrar_Click_Err
    
    Call ParseUserCommand("/SM ")
    tmpUser = cboListaUsus.Text
    Call WriteGoNearby(tmpUser)
    
    Exit Sub
cmdcerrar_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdcerrar_Click", Erl)
End Sub

Private Sub cmdOnline_Click()
    On Error Goto cmdOnline_Click_Err

    Exit Sub
cmdOnline_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdOnline_Click", Erl)
End Sub

Private Sub cmdTarget_Click()
    On Error Goto cmdTarget_Click_Err
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
    
    Exit Sub
cmdTarget_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdTarget_Click", Erl)
End Sub

Private Sub cmdComandoGM_Click()
    On Error Goto cmdComandoGM_Click_Err
Call ParseUserCommand(txtMod.Text)
    Exit Sub
cmdComandoGM_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdComandoGM_Click", Erl)
End Sub

Private Sub cmdConsulta_Click()
    On Error Goto cmdConsulta_Click_Err

    tmpUser = cboListaUsus.Text
         
    Call ParseUserCommand("/CONSULTA " & tmpUser)
    Call frmPanelgm.txtMod.SetFocus
 
    Exit Sub
cmdConsulta_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdConsulta_Click", Erl)
End Sub

Private Sub cmdEchar_Click()
    On Error Goto cmdEchar_Click_Err
    tmpUser = cboListaUsus.Text
    Call WriteKick(tmpUser)
    Call frmPanelgm.txtMod.SetFocus

    Exit Sub
cmdEchar_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdEchar_Click", Erl)
End Sub

Private Sub cmdEjecutar_Click()
    On Error Goto cmdEjecutar_Click_Err
    tmpUser = cboListaUsus.Text
    Call WriteExecute(tmpUser)
    Call frmPanelgm.txtMod.SetFocus
    Exit Sub
cmdEjecutar_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdEjecutar_Click", Erl)
End Sub

Private Sub cmdEscudo_Click(Index As Integer)
    On Error Goto cmdEscudo_Click_Err
                
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
    Exit Sub
cmdEscudo_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdEscudo_Click", Erl)
End Sub

Private Sub cmdEventos_Click()
    On Error Goto cmdEventos_Click_Err
    tmp = InputBox(JsonLanguage.Item("MENSAJE_INGRESAR_TIPO_EVENTO") & vbCrLf & JsonLanguage.Item("MENSAJE_EVENTO_TESORO_CONTINENTE") & vbCrLf & JsonLanguage.Item("MENSAJE_EVENTO_TESORO_DUNGEON") & vbCrLf & JsonLanguage.Item("MENSAJE_EVENTO_APARICION_CRIATURA"), JsonLanguage.Item("MENSAJE_INICIAR_EVENTO"))
    
    If IsNumeric(tmp) Then

        Call WriteBusquedaTesoro(CByte(tmp))
    Else
        MsgBox JsonLanguage.Item("MENSAJE_TIPO_INVALIDO"), vbExclamation, JsonLanguage.Item("TITULO_ERROR")

    End If
    
    Call frmPanelgm.txtMod.SetFocus
    Exit Sub
    Exit Sub
cmdEventos_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdEventos_Click", Erl)
End Sub


Private Sub cmdHeadMenos_Click()
    On Error Goto cmdHeadMenos_Click_Err

    tmpUser = "yo"
       
    If txtHeadNumero.Text >= 1 Then txtHeadNumero.Text = txtHeadNumero.Text - 1
    
    tmp = txtHeadNumero.Text
    

    Call ParseUserCommand("/MOD " & tmpUser & " Head " & tmp)
    
    Call frmPanelgm.txtMod.SetFocus
    Exit Sub
cmdHeadMenos_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdHeadMenos_Click", Erl)
End Sub

Private Sub cmdHeadMas_Click()
    On Error Goto cmdHeadMas_Click_Err

    tmpUser = "yo"
       
    txtHeadNumero.Text = txtHeadNumero.Text + 1
    
    tmp = txtHeadNumero.Text
    

    Call ParseUserCommand("/MOD " & tmpUser & " Head " & tmp)
    
    Call frmPanelgm.txtMod.SetFocus
    Exit Sub
    Exit Sub
cmdHeadMas_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdHeadMas_Click", Erl)
End Sub

Private Sub cmdHead0_Click()
    On Error Goto cmdHead0_Click_Err

    tmpUser = "yo"
    tmp = 0

    Call ParseUserCommand("/MOD " & tmpUser & " Head " & tmp)
    
    Call frmPanelgm.txtMod.SetFocus
    Exit Sub
    
    Exit Sub
cmdHead0_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdHead0_Click", Erl)
End Sub

Private Sub cmdInfo_Click()
    On Error Goto cmdInfo_Click_Err

    tmpUser = cboListaUsus.Text
    Call WriteRequestCharInfo(tmpUser)
    Call frmPanelgm.txtMod.SetFocus
    Exit Sub
cmdInfo_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdInfo_Click", Erl)
End Sub

Private Sub cmdInformación_Click()
    On Error Goto cmdInformación_Click_Err
    
    tmpUser = cboListaUsus.Text
    
    Call WriteRequestCharStats(tmpUser)
    Call frmPanelgm.txtMod.SetFocus

    Exit Sub
cmdInformación_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdInformación_Click", Erl)
End Sub

Private Sub cmdInseguro_Click()
    On Error Goto cmdInseguro_Click_Err
Call ParseUserCommand("/MODMAPINFO SEGURO 0")
    Exit Sub
cmdInseguro_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdInseguro_Click", Erl)
End Sub

Private Sub cmdInsertarTrigger_Click()
    On Error Goto cmdInsertarTrigger_Click_Err

    Call ParseUserCommand("/TRIGGER " & txtTextTriggers.Text)

    Exit Sub
cmdInsertarTrigger_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdInsertarTrigger_Click", Erl)
End Sub

Private Sub cmdInventario_Click()
    On Error Goto cmdInventario_Click_Err
    tmpUser = cboListaUsus.Text
    Call ParseUserCommand("/INV " & tmpUser)
    Call frmPanelgm.txtMod.SetFocus
    Exit Sub
cmdInventario_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdInventario_Click", Erl)
End Sub

Private Sub cmdInvisibilidadSi_Click()
    On Error Goto cmdInvisibilidadSi_Click_Err

    Call ParseUserCommand("/MODMAPINFO RESTRINGIR " & "SININVI")
    
    Exit Sub
cmdInvisibilidadSi_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdInvisibilidadSi_Click", Erl)
End Sub

Private Sub cmdInvisible_Click()
    On Error Goto cmdInvisible_Click_Err
    
    Call ParseUserCommand("/INVISIBLE")
        
    Call frmPanelgm.txtMod.SetFocus
    Exit Sub
    
    Exit Sub
cmdInvisible_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdInvisible_Click", Erl)
End Sub

Private Sub cmdIrCerca_Click()
    On Error Goto cmdIrCerca_Click_Err
    Call ParseUserCommand("/SM ")
    tmpUser = cboListaUsus.Text
    Call WriteGoNearby(tmpUser)

    Exit Sub
cmdIrCerca_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdIrCerca_Click", Erl)
End Sub

Private Sub cmdMagiaSin_Click()
    On Error Goto cmdMagiaSin_Click_Err

    Call ParseUserCommand("/MODMAPINFO RESTRINGIR " & "SINMAGIA")
        
    Exit Sub
cmdMagiaSin_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdMagiaSin_Click", Erl)
End Sub

Private Sub cmdMapaSeguro_Click()
    On Error Goto cmdMapaSeguro_Click_Err

    tmp = InputBox(JsonLanguage.Item("MENSAJE_EDICION_MAPA") & vbCrLf & JsonLanguage.Item("MENSAJE_MAPA_INFORMACION") & vbCrLf & JsonLanguage.Item("MENSAJE_MAPA_SEGURO") & vbCrLf & JsonLanguage.Item("MENSAJE_MAPA_INSEGURO"), JsonLanguage.Item("MENSAJE_MODIFICAR"))
    
    Select Case tmp

        Case 0
            Call ParseUserCommand("/MAPINFO")

        Case 1
            Call ParseUserCommand("/MODMAPINFO SEGURO 1")

        Case 2
            Call ParseUserCommand("/MODMAPINFO SEGURO 0")
            
            
    End Select
    
    Exit Sub
cmdMapaSeguro_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdMapaSeguro_Click", Erl)
End Sub

Private Sub cmdMapeo_Click()
    On Error Goto cmdMapeo_Click_Err
    If frmPanelgm.Width = 7365 Then
        frmPanelgm.Width = 4860
    Else
        frmPanelgm.Width = 7365
    End If
    Exit Sub
cmdMapeo_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdMapeo_Click", Erl)
End Sub

Private Sub cmdMas_Click()
    On Error Goto cmdMas_Click_Err
    tmpUser = "yo"
       
    txtCasco.Text = txtCasco.Text + 1
        
    tmp = txtCasco.Text
        
    Call ParseUserCommand("/MOD " & tmpUser & " Casco " & tmp)
    
    Call frmPanelgm.txtMod.SetFocus
    Exit Sub
cmdMas_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdMas_Click", Erl)
End Sub

Private Sub cmdMatarNPC_Click()
    On Error Goto cmdMatarNPC_Click_Err

Call ParseUserCommand("/MATA")
    Call frmPanelgm.txtMod.SetFocus

    Exit Sub
cmdMatarNPC_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdMatarNPC_Click", Erl)
End Sub

Private Sub cmdMenos_Click()
    On Error Goto cmdMenos_Click_Err
    tmpUser = "yo"
       
    If txtCasco.Text >= 1 Then txtCasco.Text = txtCasco.Text - 1
        
    tmp = txtCasco.Text
    
    
    Call ParseUserCommand("/MOD " & tmpUser & " Casco " & tmp)
    
    Call frmPanelgm.txtMod.SetFocus
    Exit Sub
cmdMenos_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdMenos_Click", Erl)
End Sub


Private Sub cmdPenas_Click()
    On Error Goto cmdPenas_Click_Err
    tmpUser = cboListaUsus.Text
    Call WritePunishments(tmpUser)
    Call frmPanelgm.txtMod.SetFocus
    Exit Sub
cmdPenas_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdPenas_Click", Erl)
End Sub

Private Sub cmdRecargarNPCs_Click()
    On Error Goto cmdRecargarNPCs_Click_Err
    Call WriteReloadNPCs
    Exit Sub
cmdRecargarNPCs_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdRecargarNPCs_Click", Erl)
End Sub

Private Sub cmdRestringirMapa_Click()
    On Error Goto cmdRestringirMapa_Click_Err

    Call ParseUserCommand("/MODMAPINFO RESTRINGIR " & "Newbie")
    Call ParseUserCommand("/MODMAPINFO RESTRINGIR " & "NoPKS")
    Call ParseUserCommand("/MODMAPINFO RESTRINGIR " & "NoCIUD")
    ' me falta Criminales , no se como restringir a todos de una
    ' luego de restringir
    ' faltaria mandar a cada uno a su hogar
    ' tambien los loguean mandarlos a su hogar.

    Exit Sub
cmdRestringirMapa_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdRestringirMapa_Click", Erl)
End Sub

Private Sub cmdRevivir_Click()
    On Error Goto cmdRevivir_Click_Err

    tmpUser = cboListaUsus.Text
    
    Call WriteReviveChar(tmpUser)
    Call frmPanelgm.txtMod.SetFocus
    
    Exit Sub
cmdRevivir_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdRevivir_Click", Erl)
End Sub

Private Sub cmdSeguirMouse_Click()
    On Error Goto cmdSeguirMouse_Click_Err
    tmpUser = cboListaUsus.Text
    chkAutoName.Value = 0
    Call ParseUserCommand("/SM " & tmpUser)
    Call frmPanelgm.txtMod.SetFocus
    Exit Sub
cmdSeguirMouse_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdSeguirMouse_Click", Erl)
End Sub

Private Sub cmdSeguro_Click()
    On Error Goto cmdSeguro_Click_Err
Call ParseUserCommand("/MODMAPINFO SEGURO 1")
    Exit Sub
cmdSeguro_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdSeguro_Click", Erl)
End Sub

Private Sub cmdSeleccionarPersonaje_Click(Index As Integer)
    On Error Goto cmdSeleccionarPersonaje_Click_Err
    'cboListaUsus = List1.List(List1.ListIndex)
    'Call AddtoRichTextBox(frmMain.RecTxt, "Haz click sobre el personaje...", 100, 100, 120, 0, 0)
    'frmMain.MousePointer = 2
    'frmMain.PanelSelect = True
    'Call SendData("TGUSER")
    Call WriteMarcaDeGm
    cboListaUsus = List1.List(List1.ListIndex)
    Call frmPanelgm.txtMod.SetFocus
    'txtHeadUser.Text = cboListaUsus.Text
    Exit Sub
cmdSeleccionarPersonaje_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdSeleccionarPersonaje_Click", Erl)
End Sub

Private Sub cmdSTAT_Click()
    On Error Goto cmdSTAT_Click_Err
    tmpUser = cboListaUsus.Text
    Call WriteRequestCharStats(tmpUser)
    Call frmPanelgm.txtMod.SetFocus
    Exit Sub
cmdSTAT_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdSTAT_Click", Erl)
End Sub

Private Sub cmdSUMUser_Click()
    On Error Goto cmdSUMUser_Click_Err
    Call WriteSummonChar(cboListaUsus.Text)
    Exit Sub
cmdSUMUser_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdSUMUser_Click", Erl)
End Sub

Private Sub cmdTrabajando_Click()
    On Error Goto cmdTrabajando_Click_Err
    Call WriteWorking
    Call frmPanelgm.txtMod.SetFocus
    Exit Sub
cmdTrabajando_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdTrabajando_Click", Erl)
End Sub

Private Sub Command1_Click()
    On Error Goto Command1_Click_Err
    
    On Error GoTo Command1_Click_Err
    
    List1.visible = True
    List2.visible = False

    
    Exit Sub

Command1_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Command1_Click", Erl)
    Resume Next
    
    Exit Sub
Command1_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.Command1_Click", Erl)
End Sub

Private Sub Command2_Click()
    On Error Goto Command2_Click_Err
    
    On Error GoTo Command2_Click_Err
    
    List1.visible = False
    List2.visible = True

    
    Exit Sub

Command2_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Command2_Click", Erl)
    Resume Next
    
    Exit Sub
Command2_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.Command2_Click", Erl)
End Sub

Private Sub Command3_Click()
    On Error Goto Command3_Click_Err
    'Call WriteGoNearby(tmpUser)
    Call Destrabar_Click
    Call frmPanelgm.txtMod.SetFocus
    Exit Sub
Command3_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.Command3_Click", Erl)
End Sub

Private Sub CrearTeleport_Click()
    On Error Goto CrearTeleport_Click_Err
    
    On Error GoTo CrearTeleport_Click_Err
    
    tmp = InputBox(JsonLanguage.Item("MENSAJE_INGRESAR_COORDENADAS"), JsonLanguage.Item("MENSAJE_INGRESAR_POSICIONES"))
    Call ParseUserCommand("/CT " & tmp)

    
    Exit Sub

CrearTeleport_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.CrearTeleport_Click", Erl)
    Resume Next
    
    Exit Sub
CrearTeleport_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.CrearTeleport_Click", Erl)
End Sub

Private Sub Command4_Click()
    On Error Goto Command4_Click_Err
    tmpUser = cboListaUsus.Text
    Call WriteGoNearby(tmpUser)
    Call WriteConsulta(tmpUser)
    Exit Sub
Command4_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.Command4_Click", Erl)
End Sub

Private Sub creartoneo_Click()
    On Error Goto creartoneo_Click_Err
    
    On Error GoTo creartoneo_Click_Err
    
    FrmTorneo.Show

    
    Exit Sub

creartoneo_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.creartoneo_Click", Erl)
    Resume Next
    
    Exit Sub
creartoneo_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.creartoneo_Click", Erl)
End Sub

Private Sub Criminales_Click()
    On Error Goto Criminales_Click_Err
    
    On Error GoTo Criminales_Click_Err
    
    tmp = InputBox(JsonLanguage.Item("MENSAJE_INGRESAR_VALOR_CRIMINALES"), JsonLanguage.Item("MENSAJE_EDICION_USUARIOS"))

    Call ParseUserCommand("/MOD " & cboListaUsus.Text & " CRI " & tmp)
    
    Exit Sub

Criminales_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Criminales_Click", Erl)
    Resume Next
    
    Exit Sub
Criminales_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.Criminales_Click", Erl)
End Sub

Private Sub Cuerpo_Click()
    On Error Goto Cuerpo_Click_Err
    
    On Error GoTo Cuerpo_Click_Err
    
    tmp = InputBox("Ingrese el valor de cuerpo que desea editar.", JsonLanguage.Item("MENSAJE_EDICION_USUARIOS"))

    Call ParseUserCommand("/MOD " & cboListaUsus.Text & " BODY " & tmp)
    
    Exit Sub

Cuerpo_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Cuerpo_Click", Erl)
    Resume Next
    
    Exit Sub
Cuerpo_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.Cuerpo_Click", Erl)
End Sub


Private Sub Destrabar_Click()
    On Error Goto Destrabar_Click_Err
    
    On Error GoTo Destrabar_Click_Err
    
    nick = Replace(List1.Text, " ", "+")
    Call WritePossUser(nick)

    
    Exit Sub

Destrabar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Destrabar_Click", Erl)
    Resume Next
    
    Exit Sub
Destrabar_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.Destrabar_Click", Erl)
End Sub

Private Sub DestruirTeleport_Click()
    On Error Goto DestruirTeleport_Click_Err
    
    On Error GoTo DestruirTeleport_Click_Err
    
    Call WriteTeleportDestroy '/DT 0.12.1

    
    Exit Sub

DestruirTeleport_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.DestruirTeleport_Click", Erl)
    Resume Next
    
    Exit Sub
DestruirTeleport_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.DestruirTeleport_Click", Erl)
End Sub

Private Sub Ejecutar_Click()
    On Error Goto Ejecutar_Click_Err
    
    On Error GoTo Ejecutar_Click_Err
    
    nick = cboListaUsus.Text
    Call WriteExecute(nick) '/EJECUTAR NICK 0.12.1

    
    Exit Sub

Ejecutar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Ejecutar_Click", Erl)
    Resume Next
    
    Exit Sub
Ejecutar_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.Ejecutar_Click", Erl)
End Sub

Private Sub Energia_Click()
    On Error Goto Energia_Click_Err
    
    On Error GoTo Energia_Click_Err
    
    tmp = InputBox("Ingrese el valor de energia que desea editar.", JsonLanguage.Item("MENSAJE_EDICION_USUARIOS"))

    Call ParseUserCommand("/MOD " & cboListaUsus.Text & " EN " & tmp)
    
    Exit Sub

Energia_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Energia_Click", Erl)
    Resume Next
    
    Exit Sub
Energia_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.Energia_Click", Erl)
End Sub

Private Sub evento1_Click()
    On Error Goto evento1_Click_Err
    
    On Error GoTo evento1_Click_Err
    
    Call WriteCrearEvento(3, 30, 2)

    
    Exit Sub

evento1_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.evento1_Click", Erl)
    Resume Next
    
    Exit Sub
evento1_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.evento1_Click", Erl)
End Sub

Private Sub evento2_Click()
    On Error Goto evento2_Click_Err
    
    On Error GoTo evento2_Click_Err
    
    Call WriteCrearEvento(3, 30, 3)

    
    Exit Sub

evento2_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.evento2_Click", Erl)
    Resume Next
    
    Exit Sub
evento2_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.evento2_Click", Erl)
End Sub

Private Sub evento3_Click()
    On Error Goto evento3_Click_Err
    
    On Error GoTo evento3_Click_Err
    
    Call WriteCrearEvento(3, 30, 5)

    
    Exit Sub

evento3_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.evento3_Click", Erl)
    Resume Next
    
    Exit Sub
evento3_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.evento3_Click", Erl)
End Sub

Private Sub evento4_Click()
    On Error Goto evento4_Click_Err
    
    On Error GoTo evento4_Click_Err
    
    Call WriteCrearEvento(2, 30, 3)

    
    Exit Sub

evento4_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.evento4_Click", Erl)
    Resume Next
    
    Exit Sub
evento4_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.evento4_Click", Erl)
End Sub

Private Sub finalizarevento_Click()
    On Error Goto finalizarevento_Click_Err
    
    On Error GoTo finalizarevento_Click_Err
    

    If MsgBox(JsonLanguage.Item("MENSAJE_FINALIZAR_EVENTO"), vbYesNo + vbQuestion, "¡ATENCIÓN!") = vbYes Then
        Call WriteFinEvento
    End If

    
    Exit Sub

finalizarevento_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.finalizarevento_Click", Erl)
    Resume Next
    
    Exit Sub
finalizarevento_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.finalizarevento_Click", Erl)
End Sub

Private Sub flash_Click()
    On Error Goto flash_Click_Err
    Call WriteSetSpeed(15#)
    Exit Sub
flash_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.flash_Click", Erl)
End Sub

Private Sub Form_Load()
    On Error Goto Form_Load_Err
    
    On Error GoTo Form_Load_Err
    
    List1.Clear
    List2.Clear
    txtMsg.Text = ""
    Call WriteRequestUserList
    
    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Form_Load", Erl)
    Resume Next
    
    Exit Sub
Form_Load_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.Form_Load", Erl)
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Goto Form_KeyPress_Err
    
    On Error GoTo Form_KeyPress_Err
    

    If (KeyAscii = 27) Then
        Unload Me

    End If

    
    Exit Sub

Form_KeyPress_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Form_KeyPress", Erl)
    Resume Next
    
    Exit Sub
Form_KeyPress_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.Form_KeyPress", Erl)
End Sub

Private Sub lento_Click()
    On Error Goto lento_Click_Err

    Exit Sub
lento_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.lento_Click", Erl)
End Sub

Private Sub GlobalEstado_Click()
    On Error Goto GlobalEstado_Click_Err

    Call ParseUserCommand("/ACTIVAR")
    Exit Sub
GlobalEstado_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.GlobalEstado_Click", Erl)
End Sub

Private Sub GuardarMapa_Click()
    On Error Goto GuardarMapa_Click_Err

    Call ParseUserCommand("/BACK")
    Exit Sub
GuardarMapa_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.GuardarMapa_Click", Erl)
End Sub

Private Sub Limpiarmundo_Click()
    On Error Goto Limpiarmundo_Click_Err

    Call ParseUserCommand("/LIMPIARMUNDO")
    Exit Sub
Limpiarmundo_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.Limpiarmundo_Click", Erl)
End Sub

Private Sub LimpiarVision_Click()
    On Error Goto LimpiarVision_Click_Err
    
    On Error GoTo LimpiarVision_Click_Err
    
    Call WriteDestroyAllItemsInArea

    
    Exit Sub

LimpiarVision_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.LimpiarVision_Click", Erl)
    Resume Next
    
    Exit Sub
LimpiarVision_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.LimpiarVision_Click", Erl)
End Sub

Private Sub List1_Click()
    On Error Goto List1_Click_Err
    
    On Error GoTo List1_Click_Err
    

    Dim ind As Integer

    ind = Val(ReadField(2, List1.List(List1.ListIndex), Asc("@")))
    txtMsg = List2.List(List1.ListIndex)
    
    
    Exit Sub

List1_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.List1_Click", Erl)
    Resume Next
    
    Exit Sub
List1_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.List1_Click", Erl)
End Sub

Private Sub List1_DblClick()
    On Error Goto List1_DblClick_Err
        tmpUser = Split(List1.Text, "(")(0)
        Call WriteGoNearby(tmpUser)
    Exit Sub
List1_DblClick_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.List1_DblClick", Erl)
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Goto List1_MouseDown_Err
    
    On Error GoTo List1_MouseDown_Err
    

    If Button = vbRightButton Then
        PopUpMenu mnuUsuario

    End If

    
    Exit Sub

List1_MouseDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.List1_MouseDown", Erl)
    Resume Next
    
    Exit Sub
List1_MouseDown_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.List1_MouseDown", Erl)
End Sub

Private Sub Mana_Click()
    On Error Goto Mana_Click_Err
    
    On Error GoTo Mana_Click_Err
    
    tmp = InputBox(JsonLanguage.Item("MENSAJE_INGRESAR_VALOR_MANA"), JsonLanguage.Item("MENSAJE_EDICION_USUARIOS"))

    Call ParseUserCommand("/MOD " & cboListaUsus.Text & " MP " & tmp)
    
    Exit Sub

Mana_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Mana_Click", Erl)
    Resume Next
    
    Exit Sub
Mana_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.Mana_Click", Erl)
End Sub

Private Sub MensajeriaMenu_Click(Index As Integer)
    On Error Goto MensajeriaMenu_Click_Err
    
    On Error GoTo MensajeriaMenu_Click_Err
    

    Select Case Index

        Case 0 'Mensaje por consola a usuarios 0.12.1
            tmp = InputBox(JsonLanguage.Item("MENSAJE_INGRESAR_TEXTO"), JsonLanguage.Item("MENSAJE_CONSOLA_USUARIOS"))
            If LenB(tmp) Then Call WriteServerMessage(tmp)

        Case 1 'Mensaje por ventana a usuarios 0.12.1
            tmp = InputBox(JsonLanguage.Item("MENSAJE_INGRESAR_TEXTO"), JsonLanguage.Item("MENSAJE_SISTEMA_USUARIOS"))
            If LenB(tmp) Then Call WriteSystemMessage(tmp)

        Case 2 'Mensaje por consola a GMS 0.12.1
            tmp = InputBox(JsonLanguage.Item("MENSAJE_ESCRIBIR_MENSAJE"), JsonLanguage.Item("MENSAJE_CONSOLA_GM"))
            If LenB(tmp) Then Call WriteGMMessage(tmp)

        Case 3 'Hablar como NPC 0.12.1
            tmp = InputBox(JsonLanguage.Item("MENSAJE_ESCRIBIR_UN_MENSAJE"), JsonLanguage.Item("MENSAJE_HABLAR_POR_NPC"))
            If LenB(tmp) Then Call WriteTalkAsNPC(tmp)

    End Select

    
    Exit Sub

MensajeriaMenu_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.MensajeriaMenu_Click", Erl)
    Resume Next
    
    Exit Sub
MensajeriaMenu_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.MensajeriaMenu_Click", Erl)
End Sub

Private Sub mnuBorrar_Click()
    On Error Goto mnuBorrar_Click_Err
    
    On Error GoTo mnuBorrar_Click_Err
    

    Dim elitem          As String
    Dim ProximamentTipo As String
    Dim TIPO            As String

    elitem = List1.ListIndex

    If List1.ListIndex < 0 Then Exit Sub
    
    Call ReadNick
    
    ProximamentTipo = General_Field_Read(2, List1.List(List1.ListIndex), "(")
    
    TIPO = General_Field_Read(1, ProximamentTipo, ")")
    
    Call WriteSOSRemove(nick & "Ø" & txtMsg & "Ø" & TIPO)
    
    Call List1.RemoveItem(List1.ListIndex)
    Call List2.RemoveItem(elitem)
    
    txtMsg.Text = vbNullString

    
    Exit Sub

mnuBorrar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.mnuBorrar_Click", Erl)
    Resume Next
    
    Exit Sub
mnuBorrar_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.mnuBorrar_Click", Erl)
End Sub

Private Sub MnuEnviar_Click(Index As Integer)
    On Error Goto MnuEnviar_Click_Err
    
    On Error GoTo MnuEnviar_Click_Err
    

    Dim Coordenadas As String

    nick = Replace(cboListaUsus.Text, " ", "+")

    Select Case Index
            
         Case 0 'Ulla
            Coordenadas = "1 55 45"
            Call ParseUserCommand("/TELEP " & nick & " " & Coordenadas)

        Case 1 'Nix
            Coordenadas = "34 40 85"
            Call ParseUserCommand("/TELEP " & nick & " " & Coordenadas)

        Case 2 'Bander
            Coordenadas = "59 45 45"
            Call ParseUserCommand("/TELEP " & nick & " " & Coordenadas)

        Case 3 'Arghal
            Coordenadas = "151 37 69"
            Call ParseUserCommand("/TELEP " & nick & " " & Coordenadas)

        Case 4 'Otro

            If LenB(nick) <> 0 Then
                Coordenadas = InputBox(JsonLanguage.Item("MENSAJE_INDICAR_POSICION"), JsonLanguage.Item("MENSAJE_TRANSPORTAR_A") & nick)

                If LenB(Coordenadas) <> 0 Then Call ParseUserCommand("/TELEP " & nick & " " & Coordenadas)

            End If

    End Select

    
    Exit Sub

MnuEnviar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.MnuEnviar_Click", Erl)
    Resume Next
    
    Exit Sub
MnuEnviar_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.MnuEnviar_Click", Erl)
End Sub

Private Sub mnuIRa_Click()
    On Error Goto mnuIRa_Click_Err
    
    On Error GoTo mnuIRa_Click_Err
    
    Call WriteGoToChar(ReadField(1, List1.List(List1.ListIndex), Asc("(")))

    
    Exit Sub

mnuIRa_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.mnuIRa_Click", Erl)
    Resume Next
    
    Exit Sub
mnuIRa_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.mnuIRa_Click", Erl)
End Sub

Private Sub mnuDestrabar_Click()
    On Error Goto mnuDestrabar_Click_Err
    On Error GoTo mnuDestrabar_Click_Err
    nick = Replace(List1.Text, " ", "+")
    Call WritePossUser(nick)

    
    Exit Sub

mnuDestrabar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.mnuDestrabar_Click", Erl)
    Resume Next
    
    Exit Sub
mnuDestrabar_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.mnuDestrabar_Click", Erl)
End Sub

Private Sub mnutraer_Click()
    On Error Goto mnutraer_Click_Err
    
    On Error GoTo mnutraer_Click_Err
    
    Call WriteSummonChar(ReadField(1, List1.List(List1.ListIndex), Asc("(")))

    
    Exit Sub

mnutraer_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.mnutraer_Click", Erl)
    Resume Next
    
    Exit Sub
mnutraer_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.mnutraer_Click", Erl)
End Sub

Private Sub mnuInvalida_Click()
    On Error Goto mnuInvalida_Click_Err
    
    On Error GoTo mnuInvalida_Click_Err
    
    nick = ReadField(1, List1.List(List1.ListIndex), Asc("("))
    
    Call ParseUserCommand("/MENSAJEINFORMACION " & nick & "@" & "Su consulta fue rechazada debido a que esta fue catalogada como invalida.")

    ' Lo advertimos
    Call WriteWarnUser(nick, "Consulta a GM's inválida.")
    
    ' Borramos el mensaje de la lista.
    Call mnuBorrar_Click
    
    Exit Sub

mnuInvalida_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.mnuInvalida_Click", Erl)
    Resume Next
    
    Exit Sub
mnuInvalida_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.mnuInvalida_Click", Erl)
End Sub

Private Sub mnuResponder_Click()
    On Error Goto mnuResponder_Click_Err
    
    On Error GoTo mnuResponder_Click_Err
    
    nick = ReadField(1, List1.List(List1.ListIndex), Asc("("))
    tmp = InputBox("Ingrese la respuesta:", "Responder consulta")
    Call ParseUserCommand("/MENSAJEINFORMACION " & nick & "@" & tmp)

    
    Exit Sub

mnuResponder_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.mnuResponder_Click", Erl)
    Resume Next
    
    Exit Sub
mnuResponder_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.mnuResponder_Click", Erl)
End Sub

Private Sub mnuManual_Click()
    On Error Goto mnuManual_Click_Err
    
    On Error GoTo mnuManual_Click_Err
    
    nick = ReadField(1, List1.List(List1.ListIndex), Asc("("))
    Call ParseUserCommand("/MENSAJEINFORMACION " & nick & "@" & "Su consulta fue rechazada debido a que la respuesta se encuentra en el Manual o FAQ de nuestra pagina web. Para mas información visite: www.argentum20.com.ar.")

    
    Exit Sub

mnuManual_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.mnuManual_Click", Erl)
    Resume Next
    
    Exit Sub
mnuManual_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.mnuManual_Click", Erl)
End Sub

Private Sub mnuAccion_Click(Index As Integer)
    On Error Goto mnuAccion_Click_Err
    
    On Error GoTo mnuAccion_Click_Err
    
    nick = cboListaUsus.Text

    If LenB(nick) <> 0 Then

        Select Case Index

            Case 0 ' Informacion General
                Call WriteRequestCharStats(nick)

            Case 1 ' Inventario
                Call WriteRequestCharInventory(nick)

            Case 2 'Skill
                Call WriteRequestCharSkills(nick)

            Case 3 'Atributos
                Call WriteRequestCharInfo(nick)

            Case 4 'Boveda
                Call WriteRequestCharBank(nick)
                Call WriteRequestCharGold(nick)

        End Select

    End If

    
    Exit Sub

mnuAccion_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.mnuAccion_Click", Erl)
    Resume Next
    
    Exit Sub
mnuAccion_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.mnuAccion_Click", Erl)
End Sub

Private Sub mnuAdmin_Click(Index As Integer)
    On Error Goto mnuAdmin_Click_Err
    
    On Error GoTo mnuAdmin_Click_Err
    
    Call cmdAccion_Click(Index)

    
    Exit Sub

mnuAdmin_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.mnuAdmin_Click", Erl)
    Resume Next
    
    Exit Sub
mnuAdmin_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.mnuAdmin_Click", Erl)
End Sub

Private Sub mnuAmbiente_Click(Index As Integer)
    On Error Goto mnuAmbiente_Click_Err
    
    On Error GoTo mnuAmbiente_Click_Err
    
    Call cmdAccion_Click(Index)

    
    Exit Sub

mnuAmbiente_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.mnuAmbiente_Click", Erl)
    Resume Next
    
    Exit Sub
mnuAmbiente_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.mnuAmbiente_Click", Erl)
End Sub

Private Sub mnuBan_Click(Index As Integer)
    On Error Goto mnuBan_Click_Err
    
    On Error GoTo mnuBan_Click_Err
    
    Call cmdAccion_Click(Index)

    
    Exit Sub

mnuBan_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.mnuBan_Click", Erl)
    Resume Next
    
    Exit Sub
mnuBan_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.mnuBan_Click", Erl)
End Sub

Private Sub mnuCarcel_Click(Index As Integer)
    On Error Goto mnuCarcel_Click_Err
    
    On Error GoTo mnuCarcel_Click_Err
    

    If Index = 60 Then
        Call cmdAccion_Click(11)
        Exit Sub

    End If

    nick = cboListaUsus.Text

    Call ParseUserCommand("/CARCEL " & nick & "@encarcelado via panelgm@" & Index)

    
    Exit Sub

mnuCarcel_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.mnuCarcel_Click", Erl)
    Resume Next
    
    Exit Sub
mnuCarcel_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.mnuCarcel_Click", Erl)
End Sub

Private Sub mnuSilencio_Click(Index As Integer)
    On Error Goto mnuSilencio_Click_Err
    
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
    
    Exit Sub
mnuSilencio_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.mnuSilencio_Click", Erl)
End Sub

Private Sub mnuHerramientas_Click(Index As Integer)
    On Error Goto mnuHerramientas_Click_Err
    
    On Error GoTo mnuHerramientas_Click_Err
    
    Call cmdAccion_Click(Index)

    
    Exit Sub

mnuHerramientas_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.mnuHerramientas_Click", Erl)
    Resume Next
    
    Exit Sub
mnuHerramientas_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.mnuHerramientas_Click", Erl)
End Sub

Private Sub mnuIP_Click(Index As Integer)
    On Error Goto mnuIP_Click_Err
    
    On Error GoTo mnuIP_Click_Err
    
    Call cmdAccion_Click(Index)

    
    Exit Sub

mnuIP_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.mnuIP_Click", Erl)
    Resume Next
    
    Exit Sub
mnuIP_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.mnuIP_Click", Erl)
End Sub

Private Sub mnuReload_Click(Index As Integer)
    On Error Goto mnuReload_Click_Err
    
    On Error GoTo mnuReload_Click_Err
    

    Select Case Index

        Case 1 'Reload objetos
            Call WriteReloadObjects

        Case 2 'Reload server.ini
            Call WriteReloadServerIni

        Case 3 'Reload mapas

            Call ParseUserCommand("/RELOAD MAP")
            
        Case 4 'Reload hechizos
            Call WriteReloadSpells

        Case 5 'Reload motd

            Call ParseUserCommand("/RELOADMOTD")
        Case 6 'Reload npcs
            Call WriteReloadNPCs

        Case 7 'Reload sockets

             If MsgBox(JsonLanguage.Item("MENSAJE_REINICIAR_API"), vbYesNo, "Advertencia") = vbYes Then
               '   Call SendData("/RELOAD SOCK")
           End If

    Case 8 'Reload otros

        Call ParseUserCommand("/RELOADOPCIONES")
End Select

    
    Exit Sub

mnuReload_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.mnuReload_Click", Erl)
    Resume Next
    
    Exit Sub
mnuReload_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.mnuReload_Click", Erl)
End Sub

Private Sub MOTD_Click()
    On Error Goto MOTD_Click_Err
    
    On Error GoTo MOTD_Click_Err
    
    Call WriteChangeMOTD 'Cambiar MOTD 0.12.1

    
    Exit Sub

MOTD_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.MOTD_Click", Erl)
    Resume Next
    
    Exit Sub
MOTD_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.MOTD_Click", Erl)
End Sub

Private Sub muyrapido_Click()
    On Error Goto muyrapido_Click_Err
    
    On Error GoTo muyrapido_Click_Err
    
    Call WriteSetSpeed(5#)
    
    Exit Sub

muyrapido_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.muyrapido_Click", Erl)
    Resume Next
    
    Exit Sub
muyrapido_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.muyrapido_Click", Erl)
End Sub

Private Sub Normal_Click()
    On Error Goto Normal_Click_Err
    
    On Error GoTo Normal_Click_Err
    
    Call WriteSetSpeed(1#)
    
    Exit Sub

Normal_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Normal_Click", Erl)
    Resume Next
    
    Exit Sub
Normal_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.Normal_Click", Erl)
End Sub

Private Sub oro_Click()
    On Error Goto oro_Click_Err
    
    On Error GoTo oro_Click_Err
    
    tmp = InputBox("Ingrese el valor de oro que desea editar.", JsonLanguage.Item("MENSAJE_EDICION_USUARIOS"))

    Call ParseUserCommand("/MOD " & cboListaUsus.Text & " ORO " & tmp)
    
    Exit Sub

oro_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.oro_Click", Erl)
    Resume Next
    
    Exit Sub
oro_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.oro_Click", Erl)
End Sub

Private Sub personalizado_Click()
    On Error Goto personalizado_Click_Err
    
    On Error GoTo personalizado_Click_Err
    
    tmp = InputBox("Ingrese evento  Tipo@Duracion@Multiplicacion" & vbCrLf & vbCrLf & "Tipo 1=Multiplica Oro" & vbCrLf & "Tipo 2=Multiplica Experiencia" & vbCrLf & "Tipo 3=Multiplica Recoleccion" & vbCrLf & "Tipo 4=Multiplica Dropeo" & vbCrLf & "Tipo 5=Multiplica Oro y Experiencia" & vbCrLf & "Tipo 6=Multiplica Oro, experiencia y recoleccion" & vbCrLf & "Tipo 7=Multiplica Todo" & vbCrLf & "Duracion= Maximo: 59" & vbCrLf & "Multiplicacion= Maximo 3", "Creacion de nuevo evento")
    Call ParseUserCommand("/CREAREVENTO " & tmp)

    
    Exit Sub

personalizado_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.personalizado_Click", Erl)
    Resume Next
    
    Exit Sub
personalizado_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.personalizado_Click", Erl)
End Sub

Private Sub rapido_Click()
    On Error Goto rapido_Click_Err
    
    On Error GoTo rapido_Click_Err

    Call WriteSetSpeed(2#)

    Exit Sub

rapido_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.rapido_Click", Erl)
    Resume Next
    
    Exit Sub
rapido_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.rapido_Click", Erl)
End Sub

Private Sub Raza_Click()
    On Error Goto Raza_Click_Err
    
    On Error GoTo Raza_Click_Err
    
    tmp = InputBox("Ingrese el valor de raza que desea editar.", JsonLanguage.Item("MENSAJE_EDICION_USUARIOS"))

    Call ParseUserCommand("/MOD " & cboListaUsus.Text & " RAZA " & tmp)
    
    Exit Sub

Raza_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Raza_Click", Erl)
    Resume Next
    
    Exit Sub
Raza_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.Raza_Click", Erl)
End Sub

Private Sub ResetPozos_Click()
    On Error Goto ResetPozos_Click_Err

    Call ParseUserCommand("/RESETPOZOS")
    
    Exit Sub
ResetPozos_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.ResetPozos_Click", Erl)
End Sub

Private Sub SeguroInseguro_Click()
    On Error Goto SeguroInseguro_Click_Err

    Call ParseUserCommand("/MODMAPINFO SEGURO 1")
    Exit Sub
SeguroInseguro_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.SeguroInseguro_Click", Erl)
End Sub

Private Sub SendGlobal_Click()
    On Error Goto SendGlobal_Click_Err
    If LenB(txtMod.Text) Then Call ParseUserCommand("/GMSG " & txtMod.Text)
    txtMod.Text = ""
    txtMod.SetFocus
    Exit Sub
SendGlobal_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.SendGlobal_Click", Erl)
End Sub

Private Sub SkillLibres_Click()
    On Error Goto SkillLibres_Click_Err
    
    On Error GoTo SkillLibres_Click_Err
    
    tmp = InputBox("Ingrese el valor de skills Libres que desea editar.", JsonLanguage.Item("MENSAJE_EDICION_USUARIOS"))

    Call ParseUserCommand("/MOD " & cboListaUsus.Text & " SKILLSLIBRES " & tmp)
    
    Exit Sub

SkillLibres_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.SkillLibres_Click", Erl)
    Resume Next
    
    Exit Sub
SkillLibres_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.SkillLibres_Click", Erl)
End Sub

Private Sub Spawn_Click()
    On Error Goto Spawn_Click_Err
    
    On Error GoTo Spawn_Click_Err
    
    Call WriteSpawnListRequest

    
    Exit Sub

Spawn_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Spawn_Click", Erl)
    Resume Next
    
    Exit Sub
Spawn_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.Spawn_Click", Erl)
End Sub

Private Sub StaffOnline_Click()
    On Error Goto StaffOnline_Click_Err
    
    On Error GoTo StaffOnline_Click_Err
    
    Call WriteOnlineGM '/ONLINEGM 0.12.1

    
    Exit Sub

StaffOnline_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.StaffOnline_Click", Erl)
    Resume Next
    
    Exit Sub
StaffOnline_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.StaffOnline_Click", Erl)
End Sub

Private Sub SubastaEstado_Click()
    On Error Goto SubastaEstado_Click_Err

    Call ParseUserCommand("/SUBASTAACTIVADA")
    Exit Sub
SubastaEstado_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.SubastaEstado_Click", Erl)
End Sub

Private Sub Temporal_Click()
    On Error Goto Temporal_Click_Err
    
    On Error GoTo Temporal_Click_Err

    Dim tmp  As String

    Dim tmp2 As Byte

    tmp2 = InputBox(JsonLanguage.Item("MENSAJE_CANTIDAD_DIAS"), JsonLanguage.Item("TITULO_CANTIDAD_DIAS"))
    tmp = InputBox(JsonLanguage.Item("MENSAJE_MOTIVO"), JsonLanguage.Item("TITULO_MOTIVO"))
    
    If MsgBox(JsonLanguage.Item("MENSAJE_BANEAR_PERSONAJE") & " " & cboListaUsus.Text & " " & tmp2, vbYesNo + vbQuestion) = vbYes Then

        Call WriteBanTemporal(cboListaUsus.Text, tmp, tmp2)

    End If

    Exit Sub

Temporal_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Temporal_Click", Erl)
    Resume Next
    
    Exit Sub
Temporal_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.Temporal_Click", Erl)
End Sub

Private Sub cmdButtonActualizarListaGms_Click()
    On Error Goto cmdButtonActualizarListaGms_Click_Err
    cmdButtonActualizarListaGms.enabled = False
    List1.Clear
    List2.Clear
    Call WriteSOSShowList
    cmdButtonActualizarListaGms.enabled = True
    Exit Sub
cmdButtonActualizarListaGms_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.cmdButtonActualizarListaGms_Click", Erl)
End Sub

Private Sub torneo_cancelar_Click()
    On Error Goto torneo_cancelar_Click_Err
    
    On Error GoTo torneo_cancelar_Click_Err
    
    Call WriteCancelarTorneo
    Call ParseUserCommand("/configlobby end")
    
    Exit Sub

torneo_cancelar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.torneo_cancelar_Click", Erl)
    Resume Next
    
    Exit Sub
torneo_cancelar_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.torneo_cancelar_Click", Erl)
End Sub

Private Sub torneo_comenzar_Click()
    On Error Goto torneo_comenzar_Click_Err
    
    On Error GoTo torneo_comenzar_Click_Err
    
    Call WriteComenzarTorneo

    
    Exit Sub

torneo_comenzar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.torneo_comenzar_Click", Erl)
    Resume Next
    
    Exit Sub
torneo_comenzar_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.torneo_comenzar_Click", Erl)
End Sub

Private Sub txtArma_KeyPress(KeyAscii As Integer)
    On Error Goto txtArma_KeyPress_Err
    If Not IsNumeric(txtArma.Text & Chr(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0
    Exit Sub
txtArma_KeyPress_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.txtArma_KeyPress", Erl)
End Sub

Private Sub txtArma_Change()
    On Error Goto txtArma_Change_Err
    Call ParseUserCommand("/MOD YO" & " Arma " & txtArma.Text)
    Exit Sub
txtArma_Change_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.txtArma_Change", Erl)
End Sub


Private Sub txtBodyYo_KeyPress(KeyAscii As Integer)
    On Error Goto txtBodyYo_KeyPress_Err
    If Not IsNumeric(txtBodyYo.Text & Chr(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0
    Exit Sub
txtBodyYo_KeyPress_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.txtBodyYo_KeyPress", Erl)
End Sub

Private Sub txtBodyYo_Change()
    On Error Goto txtBodyYo_Change_Err
    Call ParseUserCommand("/MOD YO" & " Body " & txtBodyYo.Text)
    Exit Sub
txtBodyYo_Change_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.txtBodyYo_Change", Erl)
End Sub

Private Sub txtCasco_KeyPress(KeyAscii As Integer)
    On Error Goto txtCasco_KeyPress_Err
    If Not IsNumeric(txtCasco.Text & Chr(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0
    Exit Sub
txtCasco_KeyPress_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.txtCasco_KeyPress", Erl)
End Sub

Private Sub txtCasco_Change()
    On Error Goto txtCasco_Change_Err
    Call ParseUserCommand("/MOD YO" & " Casco " & txtCasco.Text)
    Exit Sub
txtCasco_Change_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.txtCasco_Change", Erl)
End Sub

Private Sub txtEscudo_KeyPress(KeyAscii As Integer)
    On Error Goto txtEscudo_KeyPress_Err
    If Not IsNumeric(txtEscudo.Text & Chr(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0
    Exit Sub
txtEscudo_KeyPress_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.txtEscudo_KeyPress", Erl)
End Sub

Private Sub txtEscudo_Change()
    On Error Goto txtEscudo_Change_Err
    Call ParseUserCommand("/MOD YO" & " Escudo " & txtEscudo.Text)
    Exit Sub
txtEscudo_Change_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.txtEscudo_Change", Erl)
End Sub

Private Sub txtHeadNumero_KeyPress(KeyAscii As Integer)
    On Error Goto txtHeadNumero_KeyPress_Err
    If Not IsNumeric(txtHeadNumero.Text & Chr(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0
    Exit Sub
txtHeadNumero_KeyPress_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.txtHeadNumero_KeyPress", Erl)
End Sub

Private Sub txtHeadNumero_Change()
    On Error Goto txtHeadNumero_Change_Err
    Call ParseUserCommand("/MOD YO" & " Head " & txtHeadNumero.Text)
    Exit Sub
txtHeadNumero_Change_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.txtHeadNumero_Change", Erl)
End Sub


Private Sub txtMod_KeyPress(KeyAscii As Integer)
    On Error Goto txtMod_KeyPress_Err
    'If Not IsNumeric(txtMod.Text & Chr(KeyAscii)) And Not KeyAscii = 8 Then KeyAscii = 0
    
    If KeyAscii = vbKeyReturn Then
        Call ParseUserCommand(txtMod.Text)
        txtMod = ""
    End If
    Exit Sub
txtMod_KeyPress_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.txtMod_KeyPress", Erl)
End Sub

Private Sub UnbanCuenta_Click()
    On Error Goto UnbanCuenta_Click_Err
    
    On Error GoTo UnbanCuenta_Click_Err
    
    Call WriteUnBanCuenta(cboListaUsus.Text)

    
    Exit Sub

UnbanCuenta_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.UnbanCuenta_Click", Erl)
    Resume Next
    
    Exit Sub
UnbanCuenta_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.UnbanCuenta_Click", Erl)
End Sub

Private Sub UnbanPersonaje_Click()
    On Error Goto UnbanPersonaje_Click_Err
    
    On Error GoTo UnbanPersonaje_Click_Err
    
    nick = cboListaUsus.Text

    If MsgBox(JsonLanguage.Item("MENSAJEBOX_REMOVE_BAN") & " " & nick, vbYesNo + vbQuestion, "Confirmation") = vbYes Then

        Call WriteUnbanChar(nick)

    End If

    
    Exit Sub

UnbanPersonaje_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.UnbanPersonaje_Click", Erl)
    Resume Next
    
    Exit Sub
UnbanPersonaje_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.UnbanPersonaje_Click", Erl)
End Sub

Private Sub VerPantalla_Click()
    On Error Goto VerPantalla_Click_Err
    Call ParseUserCommand("/SS " & cboListaUsus.Text)
    Exit Sub
VerPantalla_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.VerPantalla_Click", Erl)
End Sub


Private Sub Vida_Click()
    On Error Goto Vida_Click_Err
    
    On Error GoTo Vida_Click_Err
    
    tmp = InputBox("Ingrese el valor de vida que desea editar.", JsonLanguage.Item("MENSAJE_EDICION_USUARIOS"))

    Call ParseUserCommand("/MOD " & cboListaUsus.Text & " HP " & tmp)
    
    Exit Sub

Vida_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.Vida_Click", Erl)
    Resume Next
    
    Exit Sub
Vida_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.Vida_Click", Erl)
End Sub

Private Sub ReadNick()
    On Error Goto ReadNick_Err
    
    On Error GoTo ReadNick_Err
    

    If List1.visible Then
        nick = General_Field_Read(1, List1.List(List1.ListIndex), "(")

        If nick = "" Then Exit Sub
        nick = Left$(nick, Len(nick))
    Else
        nick = General_Field_Read(1, List2.List(List2.ListIndex), "(")

        If nick = "" Then Exit Sub
        nick = Left$(nick, Len(nick))

    End If

    
    Exit Sub

ReadNick_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmPanelGM.ReadNick", Erl)
    Resume Next
    
    Exit Sub
ReadNick_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.ReadNick", Erl)
End Sub

Private Sub YoAcciones_Click(Index As Integer)
    On Error Goto YoAcciones_Click_Err
    
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
    
    Exit Sub
YoAcciones_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.YoAcciones_Click", Erl)
End Sub

Private Sub mnuConsulta_Click()
    On Error Goto mnuConsulta_Click_Err
    
    Dim nick As String
        nick = ReadField(1, List1.List(List1.ListIndex), Asc("("))
    
    If Len(nick) <> 0 Then
        
        Call WriteConsulta(nick)
        
    End If

    Exit Sub
mnuConsulta_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.mnuConsulta_Click", Erl)
End Sub
Public Sub CadenaChat(ByVal chat As String)
    On Error Goto CadenaChat_Err


        Dim Cadena        As String

        Dim partes()      As String

        Dim nombre        As String

        Dim PosicionBarra As Integer

        ' La cadena original
        Cadena = chat
        
        ' Divide la cadena en partes utilizando "Usuarios trabajando:" como separador
        partes = Split(Cadena, "Usuarios trabajando:")

        ' Verifica si hay al menos dos partes en la matriz resultante
            If UBound(partes) >= 1 Then
                
                ' Limpia el contenido actual del ComboBox
                cboListaUsus.Clear
                
                ' Divide la parte en nombres individuales
                Dim Nombres As Variant
                Nombres = Split(partes(1), ",")
                
                ' Agrega cada nombre al ComboBox después de eliminar los espacios adicionales
                Dim i As Integer
                For i = 0 To UBound(Nombres)
                    cboListaUsus.AddItem Trim(Nombres(i))
                Next i
            End If

        ' Divide la cadena en partes utilizando "Control de paquetes -> El usuario" como separador
        partes = Split(Cadena, "Control Paquetes---> El usuario")

        ' Verifica si hay al menos dos partes en la matriz resultante
        If UBound(partes) >= 1 Then
            ' La segunda parte (índice 1) contiene el nombre y otros caracteres
            nombre = partes(1)
            ' Encuentra la posición de la barra vertical "|" en la cadena
            PosicionBarra = InStr(nombre, "|")

            If PosicionBarra > 0 Then
                ' Si se encontró la barra vertical, obtén solo la parte del nombre antes de "|"
                nombre = Left(nombre, PosicionBarra - 1)
                ' Elimina espacios en blanco al principio y al final del nombre
                nombre = Trim(nombre)
                If chkInfoTXT.Value = 1 Then Resultado = GuardarTextoEnArchivo(Cadena, "MacroTotal.txt")
                If chkInfoTXT.Value = 1 Then Resultado = GuardarTextoEnArchivo(Cadena, "MacroDePaquetes.txt")
                If frmPanelgm.chkPaquetes.Value = 1 Then Call WriteCerraCliente(nombre)
            End If
        End If

            ' Divide la cadena en partes utilizando "Control de macro---> El usuario" como separador
            partes = Split(Cadena, "Control de macro---> El usuario")
            
            ' Verifica si hay al menos dos partes en la matriz resultante
            If UBound(partes) >= 1 Then
                ' La segunda parte (índice 1) contiene el nombre y otros caracteres
                nombre = partes(1)
                ' Encuentra la posición de la barra vertical "|" en la cadena
                PosicionBarra = InStr(nombre, "|")
            
                If PosicionBarra > 0 Then
                    ' Si se encontró la barra vertical, obtén solo la parte del nombre antes de "|"
                    nombre = Left(nombre, PosicionBarra - 1)
                    ' Elimina espacios en blanco al principio y al final del nombre
                    nombre = Trim(nombre)
                    
                    ' Declarar TiempoAnterior como Static fuera de la función
                    Static TiempoAnterior As Single

                    ' Verificar si la cadena contiene ciertos textos utilizando Select Case
                    Select Case True
                        Case InStr(Cadena, "Ocultar") > 0
                            If chkInfoTXT.Value = 1 Then Resultado = GuardarTextoEnArchivo(nombre & ",Macro de Ocultar ", "MacroOcultar.txt")
                            'Call ParseUserCommand("/MENSAJEINFORMACION " & nombre & "@" & "INFORMACION: Le recordamos que el uso de macros o programas externos está estrictamente prohibido y puede resultar en sanciones.")
                            If frmPanelgm.chkOcultar.Value = 1 Then
                                If frmPanelgm.cboListaUsus.Text = nombre Then
                                    ' Obtener el tiempo actual en milisegundos
                                    Dim TiempoActual As Single
                                    TiempoActual = Timer
                                    
                                    If TiempoActual - TiempoAnterior < frmPanelgm.txtSegundos Then
                                        Call WriteCerraCliente(nombre)
                                    End If
                                    TiempoAnterior = TiempoActual
                                End If
                            End If
                        Case InStr(Cadena, "UseItemU") > 0
                            If chkInfoTXT.Value = 1 Then Resultado = GuardarTextoEnArchivo(nombre & ",Macro de UsarItem U ", "MacroUseItemU.txt")
                            'Call ParseUserCommand("/MENSAJEINFORMACION " & nombre & "@" & "INFORMACION: Le recordamos que el uso de macros o programas externos está estrictamente prohibido y puede resultar en sanciones.")
                            If frmPanelgm.chkUsarItem.Value = 1 Then Call WriteCerraCliente(nombre)
                        Case InStr(Cadena, "UseItem") > 0
                            If chkInfoTXT.Value = 1 Then Resultado = GuardarTextoEnArchivo(nombre & ",Macro de UsarItem ", "MacroUseItem.txt")
                            'Call ParseUserCommand("/MENSAJEINFORMACION " & nombre & "@" & "INFORMACION: Le recordamos que el uso de macros o programas externos está estrictamente prohibido y puede resultar en sanciones.")
                            If frmPanelgm.chkUsarItem.Value = 1 Then Call WriteCerraCliente(nombre)
                        Case InStr(Cadena, "GuildMessage") > 0
                            If chkInfoTXT.Value = 1 Then Resultado = GuardarTextoEnArchivo(nombre & ",Macro de GuildMessage ", "MacroGuildMessage.txt")
                            'Call ParseUserCommand("/MENSAJEINFORMACION " & nombre & "@" & "INFORMACION: Le recordamos que el uso de macros o programas externos está estrictamente prohibido y puede resultar en sanciones.")
                        Case InStr(Cadena, "LeftClick") > 0
                            Resultado = GuardarTextoEnArchivo(nombre & ",Macro de LeftClick ", "MacroLeftClick.txt")
                            'Call ParseUserCommand("/MENSAJEINFORMACION " & nombre & "@" & "INFORMACION: Le recordamos que el uso de macros o programas externos está estrictamente prohibido y puede resultar en sanciones.")
                            If frmPanelgm.chkLeftClick.Value = 1 Then Call WriteCerraCliente(nombre)
                        Case InStr(Cadena, "ChangeHeading") > 0
                            Resultado = GuardarTextoEnArchivo(nombre & ",Macro de ChangeHeading ", "MacroChangeHeading.txt")
                            'Call ParseUserCommand("/MENSAJEINFORMACION " & nombre & "@" & "INFORMACION: Le recordamos que el uso de macros o programas externos está estrictamente prohibido y puede resultar en sanciones.")
                        Case Else
                            ' Manejar el caso en el que no hay coincidencias
                    End Select
            
                    If frmPanelgm.chkAutoName.Value = 1 Then frmPanelgm.cboListaUsus.Text = nombre
                    If chkInfoTXT.Value = 1 Then Resultado = GuardarTextoEnArchivo(Cadena, "MacroTotal.txt")
                End If
            End If

        ' Divide la cadena en partes utilizando "AntiCheat> El usuario" como separador
        partes = Split(Cadena, "AntiCheat--> El usuario")
            
            ' Verifica si hay al menos dos partes en la matriz resultante
            If UBound(partes) >= 1 Then
                ' La segunda parte (índice 1) contiene el nombre y otros caracteres
                nombre = partes(1)
                ' Encuentra la posición de la barra vertical "|" en la cadena
                PosicionBarra = InStr(nombre, "|")
            
                If PosicionBarra > 0 Then
                    ' Si se encontró la barra vertical, obtén solo la parte del nombre antes de "|"
                    nombre = Left(nombre, PosicionBarra - 1)
                    ' Elimina espacios en blanco al principio y al final del nombre
                    nombre = Trim(nombre)

                    ' Verificar si la cadena contiene ciertos textos utilizando Select Case
                    Select Case True
                        Case InStr(Cadena, "COORDENADAS.") > 0
                            If chkInfoTXT.Value = 1 Then Resultado = GuardarTextoEnArchivo(nombre & ",Macro de Cordenadas", "MacroCoordenadas.txt")
                            'Call ParseUserCommand("/MENSAJEINFORMACION " & nombre & "@" & "INFORMACION: Le recordamos que el uso de macros o programas externos está estrictamente prohibido y puede resultar en sanciones.")
                            If frmPanelgm.chkCoordenadas.Value = 1 Then Call WriteCerraCliente(nombre)
                        Case InStr(Cadena, ").") > 0
                            If chkInfoTXT.Value = 1 Then Resultado = GuardarTextoEnArchivo(nombre & ",Macro de click", "MacroDeClick.txt")
                            'Call ParseUserCommand("/MENSAJEINFORMACION " & nombre & "@" & "INFORMACION: Le recordamos que el uso de macros o programas externos está estrictamente prohibido y puede resultar en sanciones.")
                            If frmPanelgm.chkClicks.Value = 1 Then Call WriteCerraCliente(nombre)
                        Case InStr(Cadena, "INASISTIDO.") > 0
                            If chkInfoTXT.Value = 1 Then Resultado = GuardarTextoEnArchivo(nombre & ",Macro Inasistido", "MacroInasistido.txt")
                            'Call ParseUserCommand("/MENSAJEINFORMACION " & nombre & "@" & "INFORMACION: Le recordamos que el uso de macros o programas externos está estrictamente prohibido y puede resultar en sanciones.")
                            If frmPanelgm.chkInasistido.Value = 1 Then Call WriteCerraCliente(nombre)
                        Case InStr(Cadena, "CARTELEO.") > 0
                            If chkInfoTXT.Value = 1 Then Resultado = GuardarTextoEnArchivo(nombre & ",Macro de Carteleo", "MacroCarteleo.txt")
                            'Call ParseUserCommand("/MENSAJEINFORMACION " & nombre & "@" & "INFORMACION: Le recordamos que el uso de macros o programas externos está estrictamente prohibido y puede resultar en sanciones.")
                            If frmPanelgm.chkCarteleo.Value = 1 Then Call WriteCerraCliente(nombre)
                        Case Else
                            ' Manejar el caso en el que no hay coincidencias
                    End Select
            
                    If frmPanelgm.chkAutoName.Value = 1 Then frmPanelgm.cboListaUsus.Text = nombre
                    If chkInfoTXT.Value = 1 Then Resultado = GuardarTextoEnArchivo(Cadena, "MacroTotal.txt")
                End If
            End If
    Exit Sub
CadenaChat_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.CadenaChat", Erl)
End Sub

Function GuardarTextoEnArchivo(ByVal Cadena As String, ByVal nombreArchivo As String) As Boolean
    On Error Goto GuardarTextoEnArchivo_Err
    On Error GoTo ErrorHandler
    
    Dim fileNumber As Integer
    
    ' Abrir el archivo en modo de adición (agregará contenido sin sobrescribir)
    fileNumber = FreeFile
    Open nombreArchivo For Append As fileNumber
    
    ' Escribir la fecha y hora actual junto con la cadena en el archivo
    Print #fileNumber, Now & " " & Cadena ' O usa vbNewLine en lugar de vbCrLf si lo prefieres
    
    ' Cerrar el archivo
    Close #fileNumber
    
    ' Indicar que la operación se realizó con éxito
    GuardarTextoEnArchivo = True
    
    Exit Function

ErrorHandler:
    ' Si hay un error, indicar que la operación falló
    GuardarTextoEnArchivo = False
    Exit Function
GuardarTextoEnArchivo_Err:
    Call TraceError(Err.Number, Err.Description, "frmPanelGm.GuardarTextoEnArchivo", Erl)
End Function
