VERSION 5.00
Begin VB.Form FrmRanking 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Ranking de Batalla"
   ClientHeight    =   6705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7440
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
   ScaleHeight     =   6705
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   11760
      Width           =   6135
   End
   Begin VB.Label Puesto 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ladder"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   10
      Left            =   5400
      TabIndex        =   10
      Top             =   5820
      Width           =   1500
   End
   Begin VB.Label Puesto 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ladder"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   9
      Left            =   3180
      TabIndex        =   9
      Top             =   5820
      Width           =   1500
   End
   Begin VB.Label Puesto 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ladder"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   8
      Left            =   1010
      TabIndex        =   8
      Top             =   5820
      Width           =   1500
   End
   Begin VB.Label Puesto 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ladder"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   7
      Left            =   5360
      TabIndex        =   7
      Top             =   5015
      Width           =   1500
   End
   Begin VB.Label Puesto 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ladder"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   6
      Left            =   3180
      TabIndex        =   6
      Top             =   5010
      Width           =   1500
   End
   Begin VB.Label Puesto 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ladder"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   5
      Left            =   1010
      TabIndex        =   5
      Top             =   5015
      Width           =   1500
   End
   Begin VB.Label Puesto 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ladder"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   375
      Index           =   4
      Left            =   1010
      TabIndex        =   4
      Top             =   4260
      Width           =   1500
   End
   Begin VB.Label Puesto 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ladder"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   3
      Left            =   5200
      TabIndex        =   3
      Top             =   1110
      Width           =   1500
   End
   Begin VB.Label Puesto 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ladder"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   2
      Left            =   1000
      TabIndex        =   2
      Top             =   1110
      Width           =   1500
   End
   Begin VB.Label Puesto 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ladder"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   3040
      TabIndex        =   0
      Top             =   1110
      Width           =   1500
   End
End
Attribute VB_Name = "FrmRanking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim puntos As Boolean

Private Sub Command1_Click()
    Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If (KeyAscii = 27) Then
        Unload Me

    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim i As Byte

    If puntos Then

        For i = 1 To 10

            If LRanking(i).nombre = "-0" Then
                Puesto(i) = "Vacante"
            Else
                Puesto(i) = LRanking(i).nombre

            End If

        Next i

        puntos = False

    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Unload Me

End Sub

Private Sub Puesto_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Puesto(Index) = LRanking(Index).puntos & " puntos"
    puntos = True

End Sub
