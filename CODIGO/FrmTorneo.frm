VERSION 5.00
Begin VB.Form FrmTorneo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Organizacion de evento"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4305
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
   ScaleHeight     =   5865
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Requisitos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      Begin VB.TextBox reglas 
         Alignment       =   2  'Center
         Height          =   885
         Left            =   360
         ScrollBars      =   1  'Horizontal
         TabIndex        =   29
         Text            =   "Prohibido atacarse, tirar invisibilidad, etc"
         Top             =   4080
         Width           =   3375
      End
      Begin VB.TextBox nombre 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   360
         TabIndex        =   27
         Text            =   "Torneo 2vs 2"
         Top             =   3600
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
         Caption         =   "Buscavidas"
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   5160
         Width           =   3615
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   3840
         Width           =   3375
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "Nombre del evento"
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
         Left            =   360
         TabIndex        =   28
         Top             =   3360
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
            Size            =   8.25
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

Private Sub Command1_Click()
    Call WriteCreaerTorneo(nivelmin, nivelmax, cupos, costo, Check1.value, Check2.value, Check3.value, Check4.value, Check5.value, Check6.value, Check7.value, Check8.value, Check9.value, map, x, y, nombre, reglas)

End Sub

