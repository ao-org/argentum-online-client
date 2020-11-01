VERSION 5.00
Begin VB.Form Manual_Alquimia 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Manual"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Alquimia"
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
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   360
         Picture         =   "Manual_Alquimia.frx":0000
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   9
         Top             =   4080
         Width           =   480
      End
      Begin VB.CommandButton Command1 
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
         Left            =   2280
         TabIndex        =   2
         Top             =   6120
         Width           =   1215
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   2400
         Picture         =   "Manual_Alquimia.frx":04DF
         ScaleHeight     =   615
         ScaleWidth      =   1215
         TabIndex        =   1
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Pociones"
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
         Height          =   255
         Left            =   2160
         TabIndex        =   11
         Top             =   5520
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "4) Cuando obtengas todas las raíces deseadas, equipas la olla, de la misma forma en que equipaste la tijera."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   960
         TabIndex        =   10
         Top             =   4080
         Width           =   4935
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"Manual_Alquimia.frx":1475
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   8
         Top             =   2880
         Width           =   5535
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "2) Comprar una Tijera y Olla en la misma tienda."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Width           =   5535
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   $"Manual_Alquimia.frx":15E2
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   6
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "¿Cómo fabricar pociones?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   5
         Top             =   960
         Width           =   5775
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1)  Dirigirte al comercio “Alquimia” de cualquier ciudad."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   5535
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   $"Manual_Alquimia.frx":166A
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   3
         Top             =   4680
         Width           =   5655
      End
   End
End
Attribute VB_Name = "Manual_Alquimia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
Manual.Show
End Sub

Private Sub Label6_Click()
Unload Me
Manual_Pociones.Show
End Sub
