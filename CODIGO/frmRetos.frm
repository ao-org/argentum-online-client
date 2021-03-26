VERSION 5.00
Begin VB.Form frmRetos 
   BorderStyle     =   0  'None
   Caption         =   "Retos"
   ClientHeight    =   6120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4365
   Icon            =   "frmRetos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   408
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   291
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Jugador 
      Alignment       =   2  'Center
      BackColor       =   &H000A0A0A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   9
      Left            =   2310
      TabIndex        =   10
      Top             =   4230
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Jugador 
      Alignment       =   2  'Center
      BackColor       =   &H000A0A0A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   8
      Left            =   480
      TabIndex        =   9
      Top             =   4230
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Jugador 
      Alignment       =   2  'Center
      BackColor       =   &H000A0A0A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   7
      Left            =   2310
      TabIndex        =   8
      Top             =   3750
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Jugador 
      Alignment       =   2  'Center
      BackColor       =   &H000A0A0A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   6
      Left            =   480
      TabIndex        =   7
      Top             =   3750
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Jugador 
      Alignment       =   2  'Center
      BackColor       =   &H000A0A0A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   5
      Left            =   2310
      TabIndex        =   6
      Top             =   3270
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Jugador 
      Alignment       =   2  'Center
      BackColor       =   &H000A0A0A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   4
      Left            =   480
      TabIndex        =   5
      Top             =   3270
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Jugador 
      Alignment       =   2  'Center
      BackColor       =   &H000A0A0A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   3
      Left            =   2310
      TabIndex        =   4
      Top             =   2790
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Jugador 
      Alignment       =   2  'Center
      BackColor       =   &H000A0A0A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   2
      Left            =   480
      TabIndex        =   3
      Top             =   2790
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Jugador 
      Alignment       =   2  'Center
      BackColor       =   &H000A0A0A&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   1
      Left            =   2310
      TabIndex        =   2
      Top             =   2310
      Width           =   1575
   End
   Begin VB.TextBox txtPPT 
      Alignment       =   2  'Center
      BackColor       =   &H0014140F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2895
      TabIndex        =   0
      Text            =   "1"
      Top             =   1365
      Width           =   375
   End
   Begin VB.TextBox Apuesta 
      Alignment       =   2  'Center
      BackColor       =   &H0014140F&
      BorderStyle     =   0  'None
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
      Height          =   225
      Left            =   1920
      TabIndex        =   11
      Text            =   "20000"
      Top             =   4845
      Width           =   1455
   End
   Begin VB.TextBox Jugador 
      Alignment       =   2  'Center
      BackColor       =   &H000A0A0A&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Text            =   "Nombre"
      Top             =   2310
      Width           =   1575
   End
   Begin VB.Image Campo 
      Height          =   405
      Index           =   9
      Left            =   2265
      Picture         =   "frmRetos.frx":57E2
      Top             =   4140
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Image Campo 
      Height          =   405
      Index           =   8
      Left            =   435
      Picture         =   "frmRetos.frx":7B96
      Top             =   4140
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Image Campo 
      Height          =   405
      Index           =   7
      Left            =   2265
      Picture         =   "frmRetos.frx":9F4A
      Top             =   3660
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Image Campo 
      Height          =   405
      Index           =   6
      Left            =   435
      Picture         =   "frmRetos.frx":C2FE
      Top             =   3660
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Image Campo 
      Height          =   405
      Index           =   5
      Left            =   2265
      Picture         =   "frmRetos.frx":E6B2
      Top             =   3180
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Image Campo 
      Height          =   405
      Index           =   4
      Left            =   435
      Picture         =   "frmRetos.frx":10A66
      Top             =   3180
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Image Campo 
      Height          =   405
      Index           =   3
      Left            =   2265
      Picture         =   "frmRetos.frx":12E1A
      Top             =   2700
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Image Campo 
      Height          =   405
      Index           =   2
      Left            =   435
      Picture         =   "frmRetos.frx":151CE
      Top             =   2700
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Image Campo 
      Height          =   405
      Index           =   1
      Left            =   2265
      Picture         =   "frmRetos.frx":17582
      Top             =   2220
      Width           =   1680
   End
   Begin VB.Image Campo 
      Height          =   405
      Index           =   0
      Left            =   435
      Picture         =   "frmRetos.frx":19936
      Top             =   2220
      Width           =   1680
   End
   Begin VB.Image Cerrar 
      Height          =   420
      Left            =   3900
      Tag             =   "0"
      Top             =   15
      Width           =   465
   End
   Begin VB.Image SumarJugadores 
      Height          =   315
      Left            =   3465
      Tag             =   "0"
      Top             =   1335
      Width           =   315
   End
   Begin VB.Image RestarJugadores 
      Height          =   315
      Left            =   2385
      Tag             =   "0"
      Top             =   1350
      Width           =   315
   End
   Begin VB.Image Retar 
      Height          =   405
      Left            =   1350
      Tag             =   "0"
      Top             =   5415
      Width           =   1650
   End
   Begin VB.Label Error 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   480
      TabIndex        =   12
      Top             =   5190
      Width           =   3375
   End
End
Attribute VB_Name = "frmRetos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MAX_JUGADORES = 5

Private Sub Apuesta_Change()
    Dim Sel As Integer
    Sel = Apuesta.SelStart

    Dim NewVal As Long
    NewVal = CLng(Abs(Val(Apuesta.Text)))
    
    If NewVal > 100000000 Then
        NewVal = 10000000
        
    ElseIf NewVal < 0 Then
        NewVal = 0
    End If

    Apuesta.Text = NewVal
    Apuesta.SelStart = Sel
End Sub

Private Sub Cerrar_Click()
    Unload Me
End Sub

Private Sub Cerrar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Cerrar.Picture = LoadInterface("boton-cerrar-off.bmp")
    Cerrar.Tag = "1"
End Sub

Private Sub Cerrar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Cerrar.Tag = "0" Then
        Cerrar.Picture = LoadInterface("boton-cerrar-over.bmp")
        Cerrar.Tag = "1"
    End If
End Sub

Private Sub Form_Load()
    Me.Picture = LoadInterface("ventanaretos.bmp")

    Jugador(0) = UserName
    
    Call Aplicar_Transparencia(Me.hWnd, 240)
    Call FormParser.Parse_Form(Me)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Call moverForm(Me.hWnd)

    If Cerrar.Tag = "1" Then
        Set Cerrar.Picture = Nothing
        Cerrar.Tag = "0"
    End If
    
    If RestarJugadores.Tag = "1" Then
        Set RestarJugadores.Picture = Nothing
        RestarJugadores.Tag = "0"
    End If
    
    If SumarJugadores.Tag = "1" Then
        Set SumarJugadores.Picture = Nothing
        SumarJugadores.Tag = "0"
    End If
    
    If Retar.Tag = "1" Then
        Set Retar.Picture = Nothing
        Retar.Tag = "0"
    End If
End Sub

Private Sub Jugador_Change(Index As Integer)
    Error.Caption = vbNullString
End Sub

Private Sub RestarJugadores_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    RestarJugadores.Picture = LoadInterface("boton-sm-menos-off.bmp")
    RestarJugadores.Tag = "1"
End Sub

Private Sub RestarJugadores_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If RestarJugadores.Tag = "0" Then
        RestarJugadores.Picture = LoadInterface("boton-sm-menos-over.bmp")
        RestarJugadores.Tag = "1"
    End If
End Sub

Private Sub RestarJugadores_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim CantidadJugadores As Byte
    CantidadJugadores = Val(txtPPT.Text)
    
    If CantidadJugadores > 1 Then
        txtPPT.Text = CantidadJugadores - 1
    End If
    
    RestarJugadores.Picture = LoadInterface("boton-sm-menos-over.bmp")
    RestarJugadores.Tag = "1"

    Call ActualizarCampos
End Sub

Private Sub Retar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Retar.Picture = LoadInterface("boton-retar-ES-off.bmp")
    Retar.Tag = "1"
End Sub

Private Sub Retar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Retar.Tag = "0" Then
        Retar.Picture = LoadInterface("boton-retar-ES-over.bmp")
        Retar.Tag = "1"
    End If
End Sub

Private Sub Retar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Retar.Picture = LoadInterface("boton-retar-ES-over.bmp")
    Retar.Tag = "1"
End Sub

Private Sub SumarJugadores_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    SumarJugadores.Picture = LoadInterface("boton-sm-mas-off.bmp")
    SumarJugadores.Tag = "1"
End Sub

Private Sub SumarJugadores_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If SumarJugadores.Tag = "0" Then
        SumarJugadores.Picture = LoadInterface("boton-sm-mas-over.bmp")
        SumarJugadores.Tag = "1"
    End If
End Sub

Private Sub Retar_Click()

    If Not Validar Then Exit Sub
    
    Dim Players As String, i As Integer
    
    ' No incluímos el jugador que crea el reto
    For i = 1 To txtPPT.Text * 2 - 1
        Players = Players & Jugador(i).Text & ";"
    Next
        
    Players = Left$(Players, Len(Players) - 1)
    
    Call WriteDuel(Players, CLng(Apuesta.Text))
        
    Unload Me

End Sub

Private Function Validar() As Boolean
    Dim ErrorStr As String

    Dim i        As Integer

    For i = 0 To txtPPT.Text * 2 - 1

        If LenB(Jugador(i).Text) = 0 Then
            Error.Caption = "Complete todos los jugadores."
            Exit Function
            
        ElseIf Not ValidarNombre(Jugador(i).Text, ErrorStr) Then
            Error.Caption = "Nombre inválido: """ & Jugador(i).Text & """"
            Exit Function

        End If

    Next
    
    Dim j As Integer

    For i = 0 To txtPPT.Text * 2 - 2
        For j = i + 1 To txtPPT.Text * 2 - 1

            If Jugador(i).Text = Jugador(j).Text Then
                Error.Caption = "¡No puede haber jugadores repetidos!"
                Exit Function

            End If

        Next
    Next
    
    Error.Caption = vbNullString
    Validar = True

End Function

Private Sub SumarJugadores_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim CantidadJugadores As Byte
    CantidadJugadores = Val(txtPPT.Text)
    
    If CantidadJugadores < MAX_JUGADORES Then
        txtPPT.Text = CantidadJugadores + 1
    End If
    
    RestarJugadores.Picture = LoadInterface("boton-sm-menos-over.bmp")
    RestarJugadores.Tag = "1"
    
    Call ActualizarCampos
End Sub

Private Sub txtPPT_Change()

    Dim CantidadJugadores As Byte
    CantidadJugadores = Val(txtPPT.Text)
    
    If CantidadJugadores < 1 Then
        txtPPT.Text = 1
    
    ElseIf CantidadJugadores > MAX_JUGADORES Then
        txtPPT.Text = MAX_JUGADORES
    End If
    
    Call ActualizarCampos
    
End Sub

Private Sub ActualizarCampos()
    Dim CantidadJugadores As Byte
    CantidadJugadores = Val(txtPPT.Text)
    
    Dim i As Byte
    
    For i = 0 To Jugador.UBound
        If i \ 2 < CantidadJugadores Then
            Jugador(i).Visible = True
            Campo(i).Visible = True
        Else
            Jugador(i).Visible = False
            Campo(i).Visible = False
        End If
    Next
End Sub
