VERSION 5.00
Begin VB.Form MenuGM 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4680
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   312
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image OpcionImg 
      Height          =   360
      Index           =   12
      Left            =   0
      Top             =   4320
      Width           =   1950
   End
   Begin VB.Label OpcionLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "VER PANTALLA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   0
      TabIndex        =   12
      Top             =   4395
      Width           =   1950
   End
   Begin VB.Image OpcionImg 
      Height          =   360
      Index           =   11
      Left            =   0
      Top             =   3960
      Width           =   1950
   End
   Begin VB.Label OpcionLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "VER PROCESOS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   0
      TabIndex        =   11
      Top             =   4035
      Width           =   1950
   End
   Begin VB.Image OpcionImg 
      Height          =   360
      Index           =   10
      Left            =   0
      Top             =   3600
      Width           =   1950
   End
   Begin VB.Label OpcionLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "REVIVIR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   0
      TabIndex        =   10
      Top             =   3675
      Width           =   1950
   End
   Begin VB.Label OpcionLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PENAS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   0
      TabIndex        =   9
      Top             =   3315
      Width           =   1950
   End
   Begin VB.Image OpcionImg 
      Height          =   360
      Index           =   9
      Left            =   0
      Top             =   3240
      Width           =   1950
   End
   Begin VB.Label OpcionLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BANEAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   0
      TabIndex        =   8
      Top             =   2955
      Width           =   1950
   End
   Begin VB.Image OpcionImg 
      Height          =   360
      Index           =   8
      Left            =   0
      Top             =   2880
      Width           =   1950
   End
   Begin VB.Label OpcionLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CARCEL 5 min"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   0
      TabIndex        =   7
      Top             =   2595
      Width           =   1950
   End
   Begin VB.Image OpcionImg 
      Height          =   360
      Index           =   7
      Left            =   0
      Top             =   2520
      Width           =   1950
   End
   Begin VB.Label OpcionLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CONSULTA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   6
      Top             =   2235
      Width           =   1950
   End
   Begin VB.Image OpcionImg 
      Height          =   360
      Index           =   6
      Left            =   0
      Top             =   2160
      Width           =   1950
   End
   Begin VB.Label OpcionLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ECHAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   5
      Top             =   1875
      Width           =   1950
   End
   Begin VB.Image OpcionImg 
      Height          =   360
      Index           =   5
      Left            =   0
      Top             =   1800
      Width           =   1950
   End
   Begin VB.Image OpcionImg 
      Height          =   360
      Index           =   4
      Left            =   0
      Top             =   1440
      Width           =   1950
   End
   Begin VB.Label OpcionLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "INVENTARIO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   4
      Left            =   0
      TabIndex        =   4
      Top             =   1515
      Width           =   1950
   End
   Begin VB.Image OpcionImg 
      Height          =   360
      Index           =   3
      Left            =   0
      Top             =   1080
      Width           =   1950
   End
   Begin VB.Label OpcionLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "INFO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   3
      Top             =   1155
      Width           =   1950
   End
   Begin VB.Image OpcionImg 
      Height          =   360
      Index           =   2
      Left            =   0
      Top             =   720
      Width           =   1950
   End
   Begin VB.Label OpcionLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "NICK2IP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   795
      Width           =   1950
   End
   Begin VB.Image OpcionImg 
      Height          =   360
      Index           =   1
      Left            =   0
      Top             =   360
      Width           =   1950
   End
   Begin VB.Label OpcionLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SILENCIAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   435
      Width           =   1950
   End
   Begin VB.Image OpcionImg 
      Height          =   360
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   1950
   End
   Begin VB.Label OpcionLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SUM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   75
      Width           =   1950
   End
End
Attribute VB_Name = "MenuGM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Over As Integer

Private Sub Form_Load()
    Call Aplicar_Transparencia(Me.hwnd, 180)
    
    Over = -1
End Sub

Private Sub OpcionImg_Click(Index As Integer)
    
    Select Case Index
        Case 0
            Call ParseUserCommand("/SUM")
        Case 1
            Call ParseUserCommand("/SILENCIO " & TargetName & "@" & "15")
        Case 2
            Call ParseUserCommand("/NICK2IP " & TargetName)
        Case 3
            Call ParseUserCommand("/INFO " & TargetName)
        Case 4
            Call ParseUserCommand("/INV " & TargetName)
        Case 5
            Call ParseUserCommand("/ECHAR " & TargetName)
        Case 6
            Call ParseUserCommand("/CONSULTA " & TargetName)
        Case 7
            'Call ParseUserCommand("/CARCEL")' ver ReyarB
            Call WriteJail(TargetName, "Prevencion u ofensa", "5")
        Case 8
            'Call ParseUserCommand("/BAN")' ver ReyarB
            Call WriteBanChar(TargetName, "Incumplimiento de reglas")
        Case 9
            Call ParseUserCommand("/PENAS " & TargetName)
        Case 10
            Call ParseUserCommand("/REVIVIR " & TargetName)
        Case 11
            Call ParseUserCommand("/PROC")
        Case 12
            Call ParseUserCommand("/SS")
    End Select

    Unload Me
    
End Sub

Private Sub OpcionImg_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Over <> Index Then
        If Over >= 0 Then
            OpcionLbl(Over).ForeColor = vbWhite
        End If
        OpcionLbl(Index).ForeColor = vbYellow
        Over = Index
    End If
End Sub

Private Sub OpcionLbl_Click(Index As Integer)
    Call OpcionImg_Click(Index)
End Sub

Private Sub OpcionLbl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call OpcionImg_MouseMove(Index, Button, Shift, x, y)
End Sub

Public Sub LostFocus()
    If Over >= 0 Then
        OpcionLbl(Over).ForeColor = vbWhite
        Over = -1
    End If
End Sub
