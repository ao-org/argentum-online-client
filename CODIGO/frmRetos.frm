VERSION 5.00
Begin VB.Form frmRetos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retos"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3615
   Icon            =   "frmRetos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   3615
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Apuesta 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   1560
      TabIndex        =   13
      Text            =   "25000"
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Retar 
      Caption         =   "Retar"
      Height          =   375
      Left            =   1080
      TabIndex        =   11
      Top             =   2520
      Width           =   1575
   End
   Begin VB.TextBox Jugador 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   5
      Left            =   2040
      TabIndex        =   10
      Top             =   1560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Jugador 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   3
      Left            =   2040
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Jugador 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   2040
      TabIndex        =   6
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox Jugador 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Jugador 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Jugador 
      Alignment       =   2  'Center
      BackColor       =   &H80000000&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.OptionButton TipoReto 
      Caption         =   "3vs3"
      Height          =   375
      Index           =   2
      Left            =   2280
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.OptionButton TipoReto 
      Caption         =   "2vs2"
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.OptionButton TipoReto 
      Caption         =   "1vs1"
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Apuesta"
      Height          =   255
      Left            =   600
      TabIndex        =   14
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Error 
      Alignment       =   2  'Center
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3000
      Width           =   3375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Equipo 2"
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Equipo 1"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "frmRetos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim JugadoresPorTeam As Integer

Private Sub Apuesta_Change()
    Dim Sel As Integer
    Sel = Apuesta.SelStart

    Dim NewVal As Long
    NewVal = CLng(Abs(Val(Apuesta.Text)))
    
    If NewVal > 100000000 Then
        NewVal = 10000000
    End If

    Apuesta.Text = NewVal
    Apuesta.SelStart = Sel
End Sub

Private Sub Form_Load()
    Jugador(0) = UserName

    JugadoresPorTeam = 1
End Sub

Private Sub Jugador_Change(Index As Integer)
    Error.Caption = vbNullString
End Sub

Private Sub Retar_Click()
    If Validar Then
        Dim Players As String, i As Integer
        ' No incluímos el jugador que crea el reto
        For i = 1 To JugadoresPorTeam * 2 - 1
            Players = Players & Jugador(i).Text & ";"
        Next
        
        Players = Left$(Players, Len(Players) - 1)
    
        Call WriteDuel(Players, CLng(Apuesta.Text))
        Unload Me
    End If
End Sub

Private Sub TipoReto_Click(Index As Integer)
    JugadoresPorTeam = Index + 1

    Dim i As Integer
    For i = 0 To 2
        If i <= Index Then
            Jugador(2 * i).Visible = True
            Jugador(2 * i + 1).Visible = True
        Else
            Jugador(2 * i).Visible = False
            Jugador(2 * i + 1).Visible = False
        End If
    Next

    Error.Caption = vbNullString
End Sub

Private Function Validar() As Boolean
    Dim ErrorStr As String

    Dim i As Integer
    For i = 0 To JugadoresPorTeam * 2 - 1
        If LenB(Jugador(i).Text) = 0 Then
            Error.Caption = "Complete todos los jugadores."
            Exit Function
            
        ElseIf Not ValidarNombre(Jugador(i).Text, ErrorStr) Then
            Error.Caption = "Nombre inválido: """ & Jugador(i).Text & """"
            Exit Function
        End If
    Next
    
    Dim j As Integer
    For i = 0 To JugadoresPorTeam * 2 - 2
        For j = i + 1 To JugadoresPorTeam * 2 - 1
            If Jugador(i).Text = Jugador(j).Text Then
                Error.Caption = "¡No puede haber jugadores repetidos!"
                Exit Function
            End If
        Next
    Next
    
    Error.Caption = vbNullString
    Validar = True
End Function
