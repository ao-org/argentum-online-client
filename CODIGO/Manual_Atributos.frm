VERSION 5.00
Begin VB.Form Manual_Atributos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Manual"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6225
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Atributos"
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
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   240
         Picture         =   "Manual_Atributos.frx":0000
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   5
         Top             =   5040
         Width           =   495
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   240
         Picture         =   "Manual_Atributos.frx":074D
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   4
         Top             =   3720
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   240
         Picture         =   "Manual_Atributos.frx":0ED6
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   3
         Top             =   2040
         Width           =   495
      End
      Begin VB.PictureBox Picture17 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   240
         Picture         =   "Manual_Atributos.frx":1500
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   2
         Top             =   600
         Width           =   495
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
         TabIndex        =   1
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   $"Manual_Atributos.frx":1B53
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   960
         TabIndex        =   13
         Top             =   5040
         Width           =   4935
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Constitución"
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
         TabIndex        =   12
         Top             =   4680
         Width           =   5775
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   $"Manual_Atributos.frx":1BDC
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   960
         TabIndex        =   11
         Top             =   3720
         Width           =   4935
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Inteligencia"
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
         TabIndex        =   10
         Top             =   3360
         Width           =   5775
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"Manual_Atributos.frx":1C65
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   960
         TabIndex        =   9
         Top             =   2040
         Width           =   4935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Agilidad"
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
         TabIndex        =   8
         Top             =   1680
         Width           =   5775
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   $"Manual_Atributos.frx":1DCE
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
         Left            =   960
         TabIndex        =   7
         Top             =   600
         Width           =   4935
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fuerza"
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
         TabIndex        =   6
         Top             =   240
         Width           =   5775
      End
   End
End
Attribute VB_Name = "Manual_Atributos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Unload Me
    Manual.Show

End Sub
