VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmLogros 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Logros personales"
   ClientHeight    =   5190
   ClientLeft      =   11805
   ClientTop       =   6390
   ClientWidth     =   5235
   BeginProperty Font 
      Name            =   "Verdana"
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
   ScaleHeight     =   5190
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Objetivos disponibles"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.Frame Frame3 
         Caption         =   "Asesino"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   4935
         Begin VB.Frame Frame6 
            Caption         =   "Recompensa"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   3360
            TabIndex        =   21
            Top             =   120
            Width           =   1455
            Begin VB.PictureBox Picture3 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               FillStyle       =   0  'Solid
               ForeColor       =   &H80000008&
               Height          =   480
               Left            =   480
               ScaleHeight     =   480
               ScaleWidth      =   480
               TabIndex        =   22
               Top             =   600
               Width           =   480
            End
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Monedas de oro"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   24
               Top             =   200
               Width           =   1215
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Cant: 50000"
               Height          =   255
               Left            =   120
               TabIndex        =   23
               Top             =   1080
               Width           =   1095
            End
         End
         Begin MSComctlLib.ProgressBar ProgressKill 
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   1080
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   661
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "0/10"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   900
            Width           =   3135
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Mata 10 usuarios para obtener tu primera recompenza."
            Height          =   735
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   3255
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Novato"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         TabIndex        =   5
         Top             =   3360
         Width           =   4935
         Begin VB.Frame Frame5 
            Caption         =   "Recompensa"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   3360
            TabIndex        =   17
            Top             =   120
            Width           =   1455
            Begin VB.PictureBox Picture2 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               FillStyle       =   0  'Solid
               ForeColor       =   &H80000008&
               Height          =   480
               Left            =   480
               ScaleHeight     =   480
               ScaleWidth      =   480
               TabIndex        =   18
               Top             =   600
               Width           =   480
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Cant: 50000"
               Height          =   255
               Left            =   120
               TabIndex        =   20
               Top             =   1080
               Width           =   1095
            End
            Begin VB.Label Label8 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Monedas de oro"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   19
               Top             =   200
               Width           =   1215
            End
         End
         Begin MSComctlLib.ProgressBar ProgressLevel 
            Height          =   375
            Left            =   120
            TabIndex        =   6
            Top             =   1080
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   661
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Alcanza el nivel 10 para obtener tu primera recompenza."
            Height          =   735
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "1/10"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   900
            Width           =   3135
         End
      End
      Begin VB.Frame FrameNpcs 
         Caption         =   "Explordor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4935
         Begin VB.Frame Frame4 
            Caption         =   "Recompensa"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   3360
            TabIndex        =   13
            Top             =   120
            Width           =   1455
            Begin VB.PictureBox Picture1 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               FillStyle       =   0  'Solid
               ForeColor       =   &H80000008&
               Height          =   480
               Left            =   480
               ScaleHeight     =   480
               ScaleWidth      =   480
               TabIndex        =   14
               Top             =   600
               Width           =   480
            End
            Begin VB.Label Label6 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Monedas de oro"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   0
               TabIndex        =   16
               Top             =   200
               Width           =   1455
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Cant: 50000"
               Height          =   255
               Left            =   0
               TabIndex        =   15
               Top             =   1080
               Width           =   1455
            End
         End
         Begin MSComctlLib.ProgressBar ProgressNpcs 
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   1080
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   661
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "24/30"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   4
            Top             =   900
            Width           =   3135
         End
         Begin VB.Label labelNpcs 
            Alignment       =   2  'Center
            Caption         =   "Mata 30 npcs para conseguir tu primer recompenza y empezar a ganar."
            Height          =   855
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   3255
         End
      End
   End
End
Attribute VB_Name = "FrmLogros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
On Error Resume Next
Dim SR As RECT, DR As RECT

SR.Left = 0
SR.Top = 0
SR.Right = 32
SR.bottom = 32

DR.Left = 0
DR.Top = 0
DR.Right = 32
DR.bottom = 32

If NPcLogros.Finalizada Then
    Frame4.Caption = "Reclamar"
    Label1.Caption = "Ya podes reclamar tu recompensa"
Else
    Frame4.Caption = "Recompensa"
    Label1.Caption = NPcLogros.NpcsMatados & "/" & NPcLogros.cant
End If

FrameNpcs.Caption = NPcLogros.Nombre
labelNpcs.Caption = NPcLogros.desc
ProgressNpcs.max = NPcLogros.cant
ProgressNpcs.value = NPcLogros.NpcsMatados
ProgressNpcs.min = 0


    If NPcLogros.TipoRecompensa = 1 Then
        Label6.Caption = ObjData(Val(ReadField(1, NPcLogros.ObjRecompensa, 45))).name
        picture1.ToolTipText = ObjData(Val(ReadField(1, NPcLogros.ObjRecompensa, 45))).name
        Call Grh_Render_To_Hdc(picture1, (ObjData(Val(ReadField(1, NPcLogros.ObjRecompensa, 45))).GrhIndex), 0, 0)
        
        Label7.Caption = "Cant: " & Val(ReadField(2, NPcLogros.ObjRecompensa, 45))
    End If
    
    If NPcLogros.TipoRecompensa = 2 Then
        Label6.Caption = "Monedas de oro"
        Label7.Caption = "Cant: " & NPcLogros.OroRecompensa
        Call Grh_Render_To_Hdc(picture1, (511), 0, 0)
        
    End If
    
    If NPcLogros.TipoRecompensa = 3 Then
    
        Label6.Caption = "Puntos de exp."
        Label7.Caption = NPcLogros.ExpRecompensa & "+ exp."
        Call Grh_Render_To_Hdc(picture1, (19979), 0, 0)
        
    End If
    
    If NPcLogros.TipoRecompensa = 4 Then
        Label6.Caption = "Hechizo"
        Label7.Caption = "Apocalipsis"
        Call Grh_Render_To_Hdc(picture1, (609), 0, 0)
        
    End If




If LevelLogros.Finalizada Then
    Frame5.Caption = "Reclamar"
    Label2.Caption = "Ya podes reclamar tu recompensa"
Else
    Frame5.Caption = "Recompensa"
    Label2.Caption = LevelLogros.NivelUser & "/" & LevelLogros.cant
End If


Frame2.Caption = LevelLogros.Nombre
Label3.Caption = LevelLogros.desc

ProgressLevel.max = LevelLogros.cant
ProgressLevel.value = LevelLogros.NivelUser
ProgressLevel.min = 0



If LevelLogros.TipoRecompensa = 1 Then
        Label8.Caption = ObjData(Val(ReadField(1, NPcLogros.ObjRecompensa, 45))).name
        Picture2.ToolTipText = ObjData(Val(ReadField(1, NPcLogros.ObjRecompensa, 45))).name
        'Call Grh_Render_To_Hdc(Picture2.hdc, (ObjData(Val(ReadField(1, NPcLogros.ObjRecompensa, 45))).grhindex), 0, 0)
        Picture2.Refresh
        Label9.Caption = "Cant: " & Val(ReadField(2, NPcLogros.ObjRecompensa, 45))
    End If
    
    If LevelLogros.TipoRecompensa = 2 Then
        Label8.Caption = "Monedas de oro"
        Label9.Caption = "Cant: " & LevelLogros.OroRecompensa
       ' Call Grh_Render_To_Hdc(Picture2.hdc, (511), 0, 0)
        Picture2.Refresh
    End If
    
    If LevelLogros.TipoRecompensa = 3 Then
    
        Label8.Caption = "Puntos de exp."
        Label9.Caption = LevelLogros.ExpRecompensa & "+ exp."
       ' Call Grh_Render_To_Hdc(Picture2.hdc, (19979), 0, 0)
        Picture2.Refresh
    End If
    
    If LevelLogros.TipoRecompensa = 4 Then
        Label8.Caption = "Hechizo"
        Label9.Caption = "Apocalipsis"
       ' Call Grh_Render_To_Hdc(Picture2.hdc, (609), 0, 0)
        Picture2.Refresh
    End If




If UserLogros.Finalizada Then
    Frame6.Caption = "Reclamar"
    Label5.Caption = "Ya podes reclamar tu recompensa"
Else
    Frame6.Caption = "Recompensa"
    Label5.Caption = UserLogros.UserMatados & "/" & UserLogros.cant
End If



Frame3.Caption = UserLogros.Nombre
Label4.Caption = UserLogros.desc
ProgressKill.max = UserLogros.cant
ProgressKill.value = UserLogros.UserMatados
ProgressKill.min = 0

If UserLogros.TipoRecompensa = 1 Then
        Label11.Caption = ObjData(Val(ReadField(1, NPcLogros.ObjRecompensa, 45))).name
        Picture3.ToolTipText = ObjData(Val(ReadField(1, NPcLogros.ObjRecompensa, 45))).name
       ' Call Grh_Render_To_Hdc(Picture3.hdc, (ObjData(Val(ReadField(1, NPcLogros.ObjRecompensa, 45))).grhindex), 0, 0)
        Picture3.Refresh
        Label10.Caption = "Cant: 1"
    End If
    
    If UserLogros.TipoRecompensa = 2 Then
        Label11.Caption = "Monedas de oro"
        Label10.Caption = "Cant: " & UserLogros.OroRecompensa
       ' Call Grh_Render_To_Hdc(Picture3.hdc, (511), 0, 0)
        Picture3.Refresh
    End If
    
    If UserLogros.TipoRecompensa = 3 Then
    
        Label11.Caption = "Puntos de exp."
        Label10.Caption = UserLogros.ExpRecompensa & "+ exp."
       ' Call Grh_Render_To_Hdc(Picture3.hdc, (19979), 0, 0)
        Picture3.Refresh
    End If
    
    If UserLogros.TipoRecompensa = 4 Then
        Label11.Caption = "Hechizo"
        Label10.Caption = "Apocalipsis"
       ' Call Grh_Render_To_Hdc(Picture3.hdc, (609), 0, 0)
        Picture3.Refresh
    End If

End Sub


Private Sub Picture1_Click()
Call WriteReclamarRecompensa(1)
End Sub

Private Sub Picture2_Click()
Call WriteReclamarRecompensa(3)
End Sub

Private Sub Picture3_Click()
Call WriteReclamarRecompensa(2)
End Sub
