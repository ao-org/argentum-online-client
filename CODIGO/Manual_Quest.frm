VERSION 5.00
Begin VB.Form Manual_Quest 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Manual - Quest"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Quest"
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
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2595
         Left            =   120
         TabIndex        =   7
         Top             =   3360
         Width           =   2775
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
      Begin VB.PictureBox Picture17 
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   50
         Picture         =   "Manual_Quest.frx":0000
         ScaleHeight     =   1695
         ScaleWidth      =   5895
         TabIndex        =   1
         Top             =   1200
         Width           =   5895
      End
      Begin VB.Label nivel 
         BackStyle       =   0  'Transparent
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
         Left            =   3000
         TabIndex        =   9
         Top             =   3600
         Width           =   2895
      End
      Begin VB.Label ubicacion 
         BackStyle       =   0  'Transparent
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
         Left            =   3000
         TabIndex        =   8
         Top             =   3360
         Width           =   2895
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de Quest disponibles"
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
         Left            =   120
         TabIndex        =   6
         Top             =   3000
         Width           =   5775
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Repartidos por el mundo podrás encontrar NPCs que te encomendarán misiones a cambio de oro y experiencia o items."
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
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   5535
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Quest"
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
         TabIndex        =   4
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label descripccion 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   3000
         TabIndex        =   3
         Top             =   3840
         Width           =   2895
      End
   End
End
Attribute VB_Name = "Manual_Quest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
Manual.Show , frmmain
End Sub

Private Sub Form_Load()


Dim i As Byte

For i = 1 To NumQuest
    List1.AddItem Quest_Name(i)
Next i

List1.ListIndex = 0
End Sub

Private Sub List1_Click()
ubicacion = "Ubicación: " & NameMaps(PosMap(List1.ListIndex + 1)).name & "(" & PosMap(List1.ListIndex + 1) & ")"
descripccion = "Descripción: " & Quest_Desc(List1.ListIndex + 1)
nivel = "Nivel requerido: " & RequiredLevel(List1.ListIndex + 1)
End Sub
