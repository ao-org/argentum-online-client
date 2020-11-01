VERSION 5.00
Begin VB.Form manual_skill 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Manual - Skills"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Lista de Habilitades"
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
      Top             =   120
      Width           =   6015
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "manual_skill.frx":0000
         Left            =   1680
         List            =   "manual_skill.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2520
         Width           =   2535
      End
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
         Left            =   2160
         TabIndex        =   1
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"manual_skill.frx":0023
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label Label5 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2895
         Left            =   360
         TabIndex        =   3
         Top             =   3000
         Width           =   5415
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de Skills"
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
         Left            =   1800
         TabIndex        =   2
         Top             =   2280
         Width           =   2115
      End
   End
End
Attribute VB_Name = "manual_skill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
Label5 = SkillsDesc(Combo1.ListIndex + 1)
End Sub

Private Sub Form_Load()
Dim i As Byte
Combo1.Clear
For i = 1 To NUMSKILLS
Combo1.AddItem (SkillsNames(i))
Next i
Combo1.ListIndex = 0
End Sub

