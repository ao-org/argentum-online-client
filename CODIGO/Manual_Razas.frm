VERSION 5.00
Begin VB.Form Manual_Mineria 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mineria y Herreria"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6225
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
   ScaleHeight     =   6855
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Aspectos Basicos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   120
      TabIndex        =   22
      Top             =   50
      Visible         =   0   'False
      Width           =   6015
      Begin VB.PictureBox Picture6 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1365
         Left            =   240
         Picture         =   "Manual_Razas.frx":0000
         ScaleHeight     =   1365
         ScaleWidth      =   1470
         TabIndex        =   32
         Top             =   3960
         Width           =   1470
      End
      Begin VB.PictureBox Picture12 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1230
         Left            =   240
         Picture         =   "Manual_Razas.frx":697A
         ScaleHeight     =   1230
         ScaleWidth      =   1395
         TabIndex        =   25
         Top             =   960
         Width           =   1395
      End
      Begin VB.PictureBox Picture11 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   720
         Picture         =   "Manual_Razas.frx":C36C
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   24
         Top             =   2925
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Atras"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         TabIndex        =   23
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Label Label35 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Imagen ilustrativa:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   280
         TabIndex        =   58
         Top             =   3800
         Width           =   1455
      End
      Begin VB.Line Line1 
         X1              =   960
         X2              =   4800
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label20 
         Caption         =   $"Manual_Razas.frx":CBAE
         Height          =   975
         Left            =   120
         TabIndex        =   34
         Top             =   5400
         Width           =   5775
      End
      Begin VB.Label Label19 
         Caption         =   "Con el Martillo de herrero equipado, realizar un click derecho sobre el yunque para que se abra la pestaña de creación de ítems."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   975
         Left            =   1800
         TabIndex        =   33
         Top             =   3960
         Width           =   3855
      End
      Begin VB.Label Label18 
         Caption         =   "5) Ir a la tienda de Armaduras y situarse junto al Yunque"
         Height          =   495
         Left            =   120
         TabIndex        =   31
         Top             =   3600
         Width           =   5775
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Martillo de Herrero : Se utiliza para la fabricación de items a partir de lingotes o minerales pulidos."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   1320
         TabIndex        =   30
         Top             =   3000
         Width           =   4575
      End
      Begin VB.Label Label16 
         Caption         =   "4) Una vez obtenidos todos los lingotes y/o minerales pulidos, se tiene que comprar el Martillo de herrero en la tienda general"
         Height          =   975
         Left            =   120
         TabIndex        =   29
         Top             =   2400
         Width           =   5775
      End
      Begin VB.Label Label15 
         Caption         =   "Luego, hay que seleccionar el mineral deseado y hacerle clic derecho a la fragua para comenzar a crear lingotes o mineral pulido."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   975
         Left            =   1800
         TabIndex        =   28
         Top             =   1200
         Width           =   3855
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Imagen ilustrativa:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   200
         TabIndex        =   27
         Top             =   770
         Width           =   1455
      End
      Begin VB.Label Label27 
         Caption         =   "3) Cuando se obtenga la cantidad deseada, se debe volver a la ciudad y acercarse a la fragua"
         Height          =   975
         Left            =   120
         TabIndex        =   26
         Top             =   300
         Width           =   5775
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "¿Donde minar?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   120
      TabIndex        =   36
      Top             =   50
      Visible         =   0   'False
      Width           =   6015
      Begin VB.PictureBox Picture17 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   240
         Picture         =   "Manual_Razas.frx":CC84
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   46
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox Picture16 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   1440
         Picture         =   "Manual_Razas.frx":DCC6
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   45
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox Picture15 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   840
         Picture         =   "Manual_Razas.frx":ED08
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   44
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox Picture14 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1500
         Left            =   960
         Picture         =   "Manual_Razas.frx":FD4A
         ScaleHeight     =   1500
         ScaleWidth      =   1410
         TabIndex        =   43
         Top             =   2700
         Width           =   1410
      End
      Begin VB.PictureBox Picture13 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1500
         Left            =   2355
         Picture         =   "Manual_Razas.frx":16C7C
         ScaleHeight     =   1500
         ScaleWidth      =   1410
         TabIndex        =   42
         Top             =   2700
         Width           =   1410
      End
      Begin VB.PictureBox Picture10 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1500
         Left            =   3765
         Picture         =   "Manual_Razas.frx":1DBAE
         ScaleHeight     =   1500
         ScaleWidth      =   1410
         TabIndex        =   41
         Top             =   2700
         Width           =   1410
      End
      Begin VB.PictureBox Picture9 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1500
         Left            =   1080
         Picture         =   "Manual_Razas.frx":24AE0
         ScaleHeight     =   1500
         ScaleWidth      =   1410
         TabIndex        =   39
         Top             =   4560
         Width           =   1410
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Minas Rapajik"
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
            Left            =   1560
            TabIndex        =   40
            Top             =   600
            Width           =   2775
         End
      End
      Begin VB.PictureBox Picture7 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1500
         Left            =   3480
         Picture         =   "Manual_Razas.frx":2BA12
         ScaleHeight     =   1500
         ScaleWidth      =   1410
         TabIndex        =   38
         Top             =   4560
         Width           =   1410
      End
      Begin VB.CommandButton Manual_Mineria 
         Caption         =   "Atras"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         TabIndex        =   37
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   $"Manual_Razas.frx":32944
         Height          =   975
         Left            =   120
         TabIndex        =   56
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Los minerales que podremos obtener a través de los yacimientos son los siguientes:"
         Height          =   975
         Left            =   120
         TabIndex        =   55
         Top             =   960
         Width           =   3615
      End
      Begin VB.Label Label32 
         Caption         =   "Los mismos se encuentran distribuidos en diferentes partes de las tierras de RevolucionAO."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2040
         TabIndex        =   54
         Top             =   1440
         Width           =   3735
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Los mapas donde podemos minar son:"
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   2040
         Width           =   3135
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Minas Rapajik (136-137-138) "
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
         Left            =   960
         TabIndex        =   52
         Top             =   2280
         Width           =   4215
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Entrada en mapa 74"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   840
         TabIndex        =   51
         Top             =   2520
         Width           =   4215
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Minas Thyr (285) "
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
         Left            =   1040
         TabIndex        =   50
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Entrada en mapa 85"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1080
         TabIndex        =   49
         Top             =   4400
         Width           =   1455
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Minas Heladas (287) "
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
         Left            =   3360
         TabIndex        =   48
         Top             =   4200
         Width           =   1815
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Entrada en mapa 221"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   3525
         TabIndex        =   47
         Top             =   4400
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Aspectos Basicos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.CommandButton Command3 
         Caption         =   "Indice"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   57
         Top             =   6120
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Siguiente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         TabIndex        =   21
         Top             =   6120
         Width           =   1215
      End
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   3720
         Picture         =   "Manual_Razas.frx":32A15
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   18
         Top             =   5440
         Width           =   495
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   3720
         Picture         =   "Manual_Razas.frx":33A57
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   15
         Top             =   4480
         Width           =   495
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   3720
         Picture         =   "Manual_Razas.frx":34A99
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   12
         Top             =   3520
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   2640
         Picture         =   "Manual_Razas.frx":35ADB
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   8
         Top             =   2280
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   2640
         Picture         =   "Manual_Razas.frx":36B1D
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   2
         Top             =   1560
         Width           =   495
      End
      Begin VB.PictureBox Picture8 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1365
         Left            =   120
         Picture         =   "Manual_Razas.frx":3775F
         ScaleHeight     =   1365
         ScaleWidth      =   1470
         TabIndex        =   1
         Top             =   3480
         Width           =   1470
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "¿Donde minar?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1920
         TabIndex        =   35
         Top             =   850
         Width           =   1455
      End
      Begin VB.Label Label14 
         Caption         =   "90 minerales requeridos para 1 lingote de oro."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   4320
         TabIndex        =   20
         Top             =   5640
         Width           =   1455
      End
      Begin VB.Label Label13 
         Caption         =   "Mineral Oro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4320
         TabIndex        =   19
         Top             =   5400
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "70 minerales requeridos para 1 lingote de plata"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   4320
         TabIndex        =   17
         Top             =   4680
         Width           =   1455
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Mineral Plata"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3960
         TabIndex        =   16
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "50 minerales requeridos para 1 lingote de hierro."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   4320
         TabIndex        =   14
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Mineral Hierro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4320
         TabIndex        =   13
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Estos son los minerales requeridos por cada lingote o mineral pulido deseado:"
         Height          =   2055
         Left            =   1800
         TabIndex        =   11
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   $"Manual_Razas.frx":3E0D9
         Height          =   975
         Left            =   120
         TabIndex        =   10
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label Label5 
         Caption         =   $"Manual_Razas.frx":3E176
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   975
         Left            =   3240
         TabIndex        =   9
         Top             =   2250
         Width           =   2535
      End
      Begin VB.Label Label4 
         Caption         =   "Piquete De Minero : Se utiliza para la extracción de distintos minerales de los yacimientos."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   3240
         TabIndex        =   7
         Top             =   1530
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "1) Comprar los items de Minería:"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   1520
         Width           =   2535
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "¿Cómo fabricar un ítem de herrería?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   5
         Top             =   1200
         Width           =   3375
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Imagen ilustrativa:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   3300
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   $"Manual_Razas.frx":3E201
         Height          =   975
         Left            =   120
         TabIndex        =   4
         Top             =   300
         Width           =   5775
      End
   End
End
Attribute VB_Name = "Manual_Mineria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Frame2.Visible = True
    Frame1.Visible = False

End Sub

Private Sub Command2_Click()
    Frame2.Visible = False
    Frame1.Visible = True

End Sub

Private Sub Command3_Click()
    Manual.Show
    Unload Me

End Sub

Private Sub Label21_Click()
    Frame3.Visible = True
    Frame1.Visible = False

End Sub

Private Sub Manual_Mineria_Click()
    Frame3.Visible = False
    Frame1.Visible = True

End Sub
