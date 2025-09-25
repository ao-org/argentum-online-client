VERSION 5.00
Begin VB.Form frmRetos 
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   Caption         =   "Retos"
   ClientHeight    =   7305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4365
   Icon            =   "frmRetos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   487
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   291
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPociones 
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
      Left            =   2955
      TabIndex        =   13
      Text            =   "10000"
      Top             =   4695
      Visible         =   0   'False
      Width           =   615
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
      Index           =   9
      Left            =   2310
      TabIndex        =   10
      Top             =   4200
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
      Top             =   4200
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
      Top             =   3720
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
      Top             =   3720
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
      Top             =   3240
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
      Top             =   3240
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
      Top             =   2760
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
      Top             =   2760
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
      Top             =   2280
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
      Left            =   1800
      TabIndex        =   11
      Text            =   "20000"
      Top             =   5490
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
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Image CAMPO_CORTO 
      Height          =   405
      Left            =   2760
      Picture         =   "frmRetos.frx":57E2
      Top             =   4605
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   840
      Top             =   5955
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   465
      Top             =   4665
      Width           =   255
   End
   Begin VB.Image Campo 
      Height          =   405
      Index           =   9
      Left            =   2265
      Picture         =   "frmRetos.frx":6E16
      Top             =   4110
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Image Campo 
      Height          =   405
      Index           =   8
      Left            =   435
      Picture         =   "frmRetos.frx":91CA
      Top             =   4110
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Image Campo 
      Height          =   405
      Index           =   7
      Left            =   2265
      Picture         =   "frmRetos.frx":B57E
      Top             =   3630
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Image Campo 
      Height          =   405
      Index           =   6
      Left            =   435
      Picture         =   "frmRetos.frx":D932
      Top             =   3630
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Image Campo 
      Height          =   405
      Index           =   5
      Left            =   2265
      Picture         =   "frmRetos.frx":FCE6
      Top             =   3150
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Image Campo 
      Height          =   405
      Index           =   4
      Left            =   435
      Picture         =   "frmRetos.frx":1209A
      Top             =   3150
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Image Campo 
      Height          =   405
      Index           =   3
      Left            =   2265
      Picture         =   "frmRetos.frx":1444E
      Top             =   2670
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Image Campo 
      Height          =   405
      Index           =   2
      Left            =   435
      Picture         =   "frmRetos.frx":16802
      Top             =   2670
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Image Campo 
      Height          =   405
      Index           =   1
      Left            =   2265
      Picture         =   "frmRetos.frx":18BB6
      Top             =   2190
      Width           =   1680
   End
   Begin VB.Image Campo 
      Height          =   405
      Index           =   0
      Left            =   435
      Picture         =   "frmRetos.frx":1AF6A
      Top             =   2190
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
      Top             =   6720
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
      Top             =   5760
      Width           =   3375
   End
End
Attribute VB_Name = "frmRetos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'    Argentum 20 - Game Client Program
'    Copyright (C) 2022 - Noland Studios
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'
Option Explicit
Private PocionesRojas As Boolean
Private CaenItems     As Boolean
Private Const MAX_JUGADORES = 5

Private Sub Apuesta_Change()
    Dim Sel As Integer
    Sel = Apuesta.SelStart
    Dim NewVal As Long
    NewVal = CLng(Abs(val(Apuesta.text)))
    If NewVal > 100000000 Then
        NewVal = 10000000
    ElseIf NewVal < 0 Then
        NewVal = 0
    End If
    Apuesta.text = NewVal
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
    CAMPO_CORTO.Picture = LoadInterface("campo-corto.bmp")
    Jugador(0) = userName
    PocionesRojas = False
    CaenItems = False
    Call Aplicar_Transparencia(Me.hWnd, 240)
    Call FormParser.Parse_Form(Me)
End Sub

Private Sub cmdMas_Click()
    If val(txtPociones.text) < 10000 Then
        txtPociones.text = val(txtPociones.text + 1)
    End If
End Sub

Private Sub cmdMenos_Click()
    If val(txtPociones.text) > 0 Then
        txtPociones.text = val(txtPociones.text - 1)
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call MoverForm(Me.hWnd)
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

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If CaenItems Then
        CaenItems = False
    Else
        CaenItems = True
    End If
    If CaenItems = 0 Then
        Image2.Picture = Nothing
    Else
        Image2.Picture = LoadInterface("check-amarillo.bmp")
    End If
    Exit Sub
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If PocionesRojas Then
        PocionesRojas = False
        CAMPO_CORTO.visible = False
        txtPociones.visible = False
    Else
        PocionesRojas = True
        CAMPO_CORTO.visible = True
        txtPociones.visible = True
    End If
    If PocionesRojas = 0 Then
        Image1.Picture = Nothing
    Else
        Image1.Picture = LoadInterface("check-amarillo.bmp")
    End If
    Exit Sub
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
    CantidadJugadores = val(txtPPT.text)
    If CantidadJugadores > 1 Then
        txtPPT.text = CantidadJugadores - 1
    End If
    RestarJugadores.Picture = LoadInterface("boton-sm-menos-over.bmp")
    RestarJugadores.Tag = "1"
    Call ActualizarCampos
End Sub

Private Sub Retar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Retar.Picture = LoadInterface("boton-retar-off.bmp")
    Retar.Tag = "1"
End Sub

Private Sub Retar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Retar.Tag = "0" Then
        Retar.Picture = LoadInterface("boton-retar-over.bmp")
        Retar.Tag = "1"
    End If
End Sub

Private Sub Retar_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Retar.Picture = LoadInterface("boton-retar-over.bmp")
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
    For i = 1 To txtPPT.text * 2 - 1
        Players = Players & Jugador(i).text & ";"
    Next
    Players = Left$(Players, Len(Players) - 1)
    Call WriteDuel(Players, CLng(Apuesta.text), IIf(PocionesRojas, val(txtPociones.text), -1), CaenItems)
    Unload Me
End Sub

Private Function Validar() As Boolean
    Dim ErrorStr As String
    Dim i        As Integer
    For i = 0 To txtPPT.text * 2 - 1
        If LenB(Jugador(i).text) = 0 Then
            Error.Caption = "Complete todos los jugadores."
            Exit Function
        ElseIf Not ValidarNombre(Jugador(i).text, ErrorStr) Then
            Error.Caption = "Nombre inválido: """ & Jugador(i).text & """"
            Exit Function
        End If
    Next
    Dim J As Integer
    For i = 0 To txtPPT.text * 2 - 2
        For J = i + 1 To txtPPT.text * 2 - 1
            If Jugador(i).text = Jugador(J).text Then
                Error.Caption = "¡No puede haber jugadores repetidos!"
                Exit Function
            End If
        Next
    Next
    If PocionesRojas Then
        If val(txtPociones.text) < 0 Or val(txtPociones.text) > 10000000 Then
            Error.Caption = "¡No puedes apostar mas de 10.000.000 de monedas de oro!"
            Exit Function
        End If
    End If
    Error.Caption = vbNullString
    Validar = True
End Function

Private Sub SumarJugadores_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim CantidadJugadores As Byte
    CantidadJugadores = val(txtPPT.text)
    If CantidadJugadores < MAX_JUGADORES Then
        txtPPT.text = CantidadJugadores + 1
    End If
    RestarJugadores.Picture = LoadInterface("boton-sm-menos-over.bmp")
    RestarJugadores.Tag = "1"
    Call ActualizarCampos
End Sub

Private Sub txtPPT_Change()
    Dim CantidadJugadores As Byte
    CantidadJugadores = val(txtPPT.text)
    If CantidadJugadores < 1 Then
        txtPPT.text = 1
    ElseIf CantidadJugadores > MAX_JUGADORES Then
        txtPPT.text = MAX_JUGADORES
    End If
    Call ActualizarCampos
End Sub

Private Sub ActualizarCampos()
    Dim CantidadJugadores As Byte
    CantidadJugadores = val(txtPPT.text)
    Dim i As Byte
    For i = 0 To Jugador.UBound
        If i \ 2 < CantidadJugadores Then
            Jugador(i).visible = True
            Campo(i).visible = True
        Else
            Jugador(i).visible = False
            Campo(i).visible = False
        End If
    Next
End Sub
