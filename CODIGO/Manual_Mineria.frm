VERSION 5.00
Begin VB.Form Manual_Razas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Manual - Razas"
   ClientHeight    =   6810
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
   ScaleHeight     =   6810
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Razas"
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
         TabIndex        =   27
         Top             =   6120
         Width           =   1215
      End
      Begin VB.PictureBox Picture3 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   990
         Left            =   250
         Picture         =   "Manual_Mineria.frx":0000
         ScaleHeight     =   990
         ScaleWidth      =   780
         TabIndex        =   10
         Top             =   4920
         Width           =   780
      End
      Begin VB.PictureBox Picture2 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1005
         Left            =   240
         Picture         =   "Manual_Mineria.frx":0D11
         ScaleHeight     =   1005
         ScaleWidth      =   810
         TabIndex        =   7
         Top             =   3600
         Width           =   810
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1020
         Left            =   240
         Picture         =   "Manual_Mineria.frx":1B01
         ScaleHeight     =   1020
         ScaleWidth      =   795
         TabIndex        =   5
         Top             =   2400
         Width           =   795
      End
      Begin VB.PictureBox Picture8 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   990
         Left            =   240
         Picture         =   "Manual_Mineria.frx":2920
         ScaleHeight     =   990
         ScaleWidth      =   810
         TabIndex        =   2
         Top             =   1200
         Width           =   810
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
         TabIndex        =   1
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Enano"
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
         Left            =   240
         TabIndex        =   25
         Top             =   4760
         Width           =   795
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Elfo drow"
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
         Left            =   240
         TabIndex        =   24
         Top             =   3440
         Width           =   795
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Elfo"
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
         Left            =   240
         TabIndex        =   23
         Top             =   2240
         Width           =   795
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Humano"
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
         Left            =   240
         TabIndex        =   22
         Top             =   1020
         Width           =   795
      End
      Begin VB.Label Label5 
         Caption         =   $"Manual_Mineria.frx":3708
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
         Height          =   1215
         Left            =   1200
         TabIndex        =   11
         Top             =   4920
         Width           =   4575
      End
      Begin VB.Label Label4 
         Caption         =   $"Manual_Mineria.frx":37FE
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
         Height          =   1215
         Left            =   1200
         TabIndex        =   9
         Top             =   3600
         Width           =   4695
      End
      Begin VB.Label Label3 
         Caption         =   $"Manual_Mineria.frx":3921
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
         Height          =   1215
         Left            =   1200
         TabIndex        =   8
         Top             =   2280
         Width           =   4695
      End
      Begin VB.Label Label1 
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
         Height          =   495
         Left            =   960
         TabIndex        =   6
         Top             =   2760
         Width           =   4695
      End
      Begin VB.Label Label2 
         Caption         =   $"Manual_Mineria.frx":3A21
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label Label9 
         Caption         =   $"Manual_Mineria.frx":3AB8
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
         Left            =   1200
         TabIndex        =   3
         Top             =   1200
         Width           =   4695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Razas"
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
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   6015
      Begin VB.PictureBox Picture4 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2145
         Left            =   240
         Picture         =   "Manual_Mineria.frx":3B7B
         ScaleHeight     =   2145
         ScaleWidth      =   5460
         TabIndex        =   18
         Top             =   3240
         Width           =   5460
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
         TabIndex        =   15
         Top             =   6120
         Width           =   1215
      End
      Begin VB.PictureBox Picture7 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1020
         Left            =   240
         Picture         =   "Manual_Mineria.frx":A5EF
         ScaleHeight     =   1020
         ScaleWidth      =   795
         TabIndex        =   14
         Top             =   480
         Width           =   795
      End
      Begin VB.PictureBox Picture6 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1020
         Left            =   240
         Picture         =   "Manual_Mineria.frx":B3CD
         ScaleHeight     =   1020
         ScaleWidth      =   810
         TabIndex        =   13
         Top             =   1800
         Width           =   810
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tabla de atributos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   2950
         Width           =   5235
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Orco"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1630
         Width           =   795
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Gnomo"
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
         Left            =   240
         TabIndex        =   20
         Top             =   300
         Width           =   795
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Atributos"
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
         Left            =   2520
         TabIndex        =   19
         Top             =   5520
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   $"Manual_Mineria.frx":C1E9
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
         Left            =   1200
         TabIndex        =   17
         Top             =   600
         Width           =   4695
      End
      Begin VB.Label Label8 
         Caption         =   "Pocos inteligentes, Son los más hábiles en el uso del combate con armas, pero también son la raza menos hábil para la magia."
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
         Height          =   495
         Left            =   1200
         TabIndex        =   16
         Top             =   2040
         Width           =   4695
      End
   End
End
Attribute VB_Name = "Manual_Razas"
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
Unload Me
Manual.Show
End Sub

Private Sub Label21_Click()
Unload Me
Manual_Atributos.Show
End Sub
