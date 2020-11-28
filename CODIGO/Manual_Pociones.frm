VERSION 5.00
Begin VB.Form Manual_Pociones 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Manual"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Pociones"
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.Frame Frame1 
         Caption         =   "Informacion:"
         Height          =   3050
         Left            =   3000
         TabIndex        =   4
         Top             =   1850
         Width           =   2895
         Begin VB.Frame Frame3 
            Caption         =   "Item:"
            Height          =   855
            Left            =   2040
            TabIndex        =   13
            Top             =   1920
            Width           =   735
            Begin VB.PictureBox picture1 
               Appearance      =   0  'Flat
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   465
               Left            =   120
               ScaleHeight     =   32
               ScaleMode       =   0  'User
               ScaleWidth      =   32
               TabIndex        =   14
               Top             =   240
               Width           =   480
            End
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "raices"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   1440
            Width           =   540
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Raices necesarias:"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   1200
            Width           =   1560
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Caracteristicas:"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   2160
            Width           =   1320
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "raices"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   1920
            Width           =   540
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "precio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   960
            Width           =   435
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Skill manualidades necesario:"
            Height          =   195
            Left            =   120
            TabIndex        =   8
            Top             =   1680
            Width           =   2490
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Precio:"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   720
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre: "
            Height          =   195
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   750
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Restaura 35 puntos de vida."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   5
            Top             =   2400
            Width           =   1815
         End
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2955
         Left            =   240
         TabIndex        =   3
         Top             =   1920
         Width           =   2655
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Indice"
         Height          =   495
         Left            =   2400
         TabIndex        =   1
         Top             =   6120
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Lista de pociones disponibles:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1680
         Width           =   5535
      End
      Begin VB.Label Label11 
         Caption         =   "Nota: Los precios sugeridos son con 0 (cero) Skills en Comerciar."
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   5040
         Width           =   5535
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   $"Manual_Pociones.frx":0000
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1455
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5775
      End
   End
End
Attribute VB_Name = "Manual_Pociones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private ListaPociones(1 To 13) As Integer

Private Sub Command2_Click()
    Unload Me
    Manual.Show

End Sub

Private Sub Form_Load()

    Dim i As Byte

    ListaPociones(1) = 37
    ListaPociones(2) = 38
    ListaPociones(3) = 36
    ListaPociones(4) = 39
    ListaPociones(5) = 166
    ListaPociones(6) = 891
    ListaPociones(7) = 894
    ListaPociones(8) = 889
    ListaPociones(9) = 892
    ListaPociones(10) = 893
    ListaPociones(11) = 895
    ListaPociones(12) = 896
    ListaPociones(13) = 1096

    For i = 1 To 13
        List1.AddItem ObjData(ListaPociones(i)).name
    Next i

End Sub

Private Sub Label21_Click()

End Sub

Private Sub List1_Click()
    Label6.Caption = ObjData(ListaPociones(List1.ListIndex + 1)).name
    Label7.Caption = ObjData(ListaPociones(List1.ListIndex + 1)).Valor
    Label8.Caption = ObjData(ListaPociones(List1.ListIndex + 1)).Raices
    Label9.Caption = ObjData(ListaPociones(List1.ListIndex + 1)).SkPociones

    If ObjData(ListaPociones(List1.ListIndex + 1)).Raices = 0 Then
        Label8.Caption = "-"
        Label9.Caption = "-"

    End If

    Label3.Caption = ObjData(ListaPociones(List1.ListIndex + 1)).Texto
    Call Grh_Render_To_Hdc(picture1, ObjData(ListaPociones(List1.ListIndex + 1)).GrhIndex, 0, 0, False)

End Sub
