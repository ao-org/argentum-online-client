VERSION 5.00
Begin VB.Form frmStatistics 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   8265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5385
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   551
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   359
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblPuntosPesca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4050
      TabIndex        =   16
      Top             =   2385
      Width           =   165
   End
   Begin VB.Image ImgPesca 
      Height          =   495
      Left            =   1350
      Top             =   1170
      Width           =   540
   End
   Begin VB.Image ImgCombate 
      Height          =   495
      Left            =   810
      Top             =   1170
      Width           =   540
   End
   Begin VB.Image ImgEstadisticasPersonaje 
      Height          =   495
      Left            =   255
      Top             =   1170
      Width           =   540
   End
   Begin VB.Image imgCerrar 
      Height          =   375
      Left            =   4920
      Top             =   5
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Elfo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   195
      Index           =   7
      Left            =   3600
      TabIndex        =   15
      Top             =   2880
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Elfo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   195
      Index           =   2
      Left            =   3480
      TabIndex        =   14
      Top             =   3360
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label lblpuntosbattle 
      BackStyle       =   0  'Transparent
      Caption         =   "Neutral"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   195
      Left            =   720
      TabIndex        =   13
      Top             =   6840
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   1725
      Top             =   7500
      Width           =   1935
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   1
      Left            =   2040
      TabIndex        =   12
      Top             =   5025
      Width           =   180
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   2
      Left            =   2040
      TabIndex        =   11
      Top             =   5325
      Width           =   180
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   3
      Left            =   2040
      TabIndex        =   10
      Top             =   5640
      Width           =   180
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   4
      Left            =   2040
      TabIndex        =   9
      Top             =   5940
      Width           =   180
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Paladin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   195
      Index           =   4
      Left            =   2040
      TabIndex        =   8
      Top             =   2920
      Width           =   1260
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Elfo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   195
      Index           =   8
      Left            =   2040
      TabIndex        =   7
      Top             =   3240
      Width           =   1140
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Hombre"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   195
      Index           =   6
      Left            =   2040
      TabIndex        =   6
      Top             =   3525
      Width           =   1260
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "10 min"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   195
      Index           =   5
      Left            =   2040
      TabIndex        =   5
      Top             =   4130
      Width           =   1185
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Neutral"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   195
      Index           =   9
      Left            =   2040
      TabIndex        =   4
      Top             =   3810
      UseMnemonic     =   0   'False
      Width           =   1260
   End
   Begin VB.Label Atri 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "15"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   210
      Index           =   5
      Left            =   2040
      TabIndex        =   3
      Top             =   6240
      Width           =   180
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   195
      Index           =   3
      Left            =   2040
      TabIndex        =   2
      Top             =   4125
      WhatsThisHelpID =   8000
      Width           =   705
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   195
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Top             =   3525
      UseMnemonic     =   0   'False
      Width           =   555
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   195
      Index           =   0
      Left            =   2040
      TabIndex        =   0
      Top             =   3810
      UseMnemonic     =   0   'False
      Width           =   555
   End
End
Attribute VB_Name = "frmStatistics"
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
Private cBotonAceptar               As clsGraphicalButton
Private cBotonCerrar                As clsGraphicalButton
Private cBotonEstadisticasPersonaje As clsGraphicalButton
Private cBotonCombate               As clsGraphicalButton
Private cBotonPesca                 As clsGraphicalButton

Public Sub Iniciar_Labels()
    On Error GoTo Iniciar_Labels_Err
    'Iniciamos los labels con los valores de los atributos y los skills
    Dim i As Integer
    For i = 1 To NUMATRIBUTOS 'Colocado
        Atri(i).Caption = UserAtributos(i)
    Next
    Select Case UserEstadisticas.Alineacion
        Case 0
            Label6(9).Caption = "Criminal"
            Label6(9).ForeColor = RGB(255, 0, 0)
        Case 1
            Label6(9).Caption = "Ciudadano"
            Label6(9).ForeColor = RGB(0, 128, 255)
        Case 2
            Label6(9).Caption = "Caos"
            Label6(9).ForeColor = RGB(128, 0, 0)
        Case 3
            Label6(9).Caption = "Imperial"
            Label6(9).ForeColor = RGB(33, 133, 132)
        Case Else
            Label6(9).Caption = "Desconocido"
    End Select
    With UserEstadisticas
        Label6(0).Caption = .CriminalesMatados 'Colocado
        Label6(1).Caption = .CiudadanosMatados 'Colocado
        Label6(3).Caption = .NpcsMatados
        Label6(4).Caption = .Clase 'Colocado
        Label6(5).Caption = .PenaCarcel & " min"
        Label6(6).Caption = .Genero
        Label6(7).Caption = .VecesQueMoriste
        Label6(8).Caption = .Raza
        lblPuntosPesca.Caption = .PuntosPesca
        lblpuntosbattle.Caption = .BattlePuntos
    End With
    Exit Sub
Iniciar_Labels_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmStatistics.Iniciar_Labels", Erl)
    Resume Next
End Sub

Private Sub Form_Load()
    Iniciar_Labels
    showStatsLabels
    loadButtons
End Sub

Private Sub showCombateLabels()
    'Show combate labels
    Me.Label6(0).visible = True
    Me.Label6(1).visible = True
    Me.Label6(3).visible = True
    'Hide other labels
    Me.Label6(4).visible = False
    Me.Label6(5).visible = False
    Me.Label6(6).visible = False
    Me.Label6(8).visible = False
    Me.Label6(9).visible = False
    Atri(1).visible = False
    Atri(2).visible = False
    Atri(3).visible = False
    Atri(4).visible = False
    Atri(5).visible = False
    Me.lblPuntosPesca.visible = False
End Sub

Private Sub showStatsLabels()
    'Show combate labels
    Me.Label6(0).visible = False
    Me.Label6(1).visible = False
    Me.Label6(3).visible = False
    'Hide other labels
    Me.Label6(4).visible = True
    Me.Label6(5).visible = True
    Me.Label6(6).visible = True
    Me.Label6(8).visible = True
    Me.Label6(9).visible = True
    Atri(1).visible = True
    Atri(2).visible = True
    Atri(3).visible = True
    Atri(4).visible = True
    Atri(5).visible = True
    Me.lblPuntosPesca.visible = False
End Sub

Private Sub showPescaLabels()
    Me.lblPuntosPesca.visible = True
    'Show combate labels
    Me.Label6(0).visible = False
    Me.Label6(1).visible = False
    Me.Label6(3).visible = False
    'Hide other labels
    Me.Label6(4).visible = False
    Me.Label6(5).visible = False
    Me.Label6(6).visible = False
    Me.Label6(8).visible = False
    Me.Label6(9).visible = False
    Atri(1).visible = False
    Atri(2).visible = False
    Atri(3).visible = False
    Atri(4).visible = False
    Atri(5).visible = False
End Sub

Private Sub loadButtons()
    Set cBotonAceptar = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonEstadisticasPersonaje = New clsGraphicalButton
    Set cBotonCombate = New clsGraphicalButton
    Set cBotonPesca = New clsGraphicalButton
    Call cBotonAceptar.Initialize(Image1, "boton-aceptar-default.bmp", "boton-aceptar-over.bmp", "boton-aceptar-off.bmp", Me)
    Call cBotonCerrar.Initialize(imgCerrar, "boton-cerrar-default.bmp", "boton-cerrar-over.bmp", "boton-cerrar-off.bmp", Me)
    Call cBotonEstadisticasPersonaje.Initialize(ImgEstadisticasPersonaje, "boton-personaje-default.bmp", "boton-personaje-over.bmp", "boton-personaje-off.bmp", Me)
    Call cBotonCombate.Initialize(ImgCombate, "boton-retos-default.bmp", "boton-retos-over.bmp", "boton-retos-off.bmp", Me)
    Call cBotonPesca.Initialize(ImgPesca, "boton-pesca-default.bmp", "boton-pesca-over.bmp", "boton-pesca-off.bmp", Me)
End Sub

Private Sub Image1_Click()
    Unload Me
End Sub

Private Sub imgCerrar_Click()
    Unload Me
End Sub

Private Sub ImgCombate_Click()
    showCombateLabels
    Me.Picture = LoadInterface("ventanaestadisticas_combate.bmp")
End Sub

Private Sub ImgEstadisticasPersonaje_Click()
    showStatsLabels
    Me.Picture = LoadInterface("ventanaestadisticas_personaje.bmp")
End Sub

Private Sub ImgPesca_Click()
    showPescaLabels
    Me.Picture = LoadInterface("ventanaestadisticas_pesca.bmp")
End Sub
