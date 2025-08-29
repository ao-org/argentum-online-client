VERSION 5.00
Begin VB.Form frmStatistics 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   8268
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5388
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   689
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   449
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblPuntosPesca 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
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
         Size            =   8.4
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
         Size            =   8.4
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
         Size            =   8.4
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
         Size            =   8.4
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
         Size            =   8.4
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
         Size            =   8.4
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
         Size            =   8.4
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
         Size            =   8.4
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
         Size            =   8.4
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
         Size            =   8.4
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
         Size            =   8.4
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
         Size            =   8.4
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
         Size            =   8.4
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
         Size            =   8.4
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
         Size            =   8.4
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
         Size            =   8.4
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

Private cBotonAceptar As clsGraphicalButton
Private cBotonCerrar As clsGraphicalButton
Private cBotonEstadisticasPersonaje As clsGraphicalButton
Private cBotonCombate As clsGraphicalButton
Private cBotonPesca As clsGraphicalButton
Public Sub Iniciar_Labels()
    On Error Goto Iniciar_Labels_Err
    
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
    
    Exit Sub
Iniciar_Labels_Err:
    Call TraceError(Err.Number, Err.Description, "frmStatistics.Iniciar_Labels", Erl)
End Sub

Private Sub Form_Load()
    On Error Goto Form_Load_Err
    Iniciar_Labels
    showStatsLabels
    loadButtons
    Exit Sub
Form_Load_Err:
    Call TraceError(Err.Number, Err.Description, "frmStatistics.Form_Load", Erl)
End Sub
Private Sub showCombateLabels()
    On Error Goto showCombateLabels_Err
    
    'Show combate labels
    Me.Label6(0).Visible = True
    Me.Label6(1).Visible = True
    Me.Label6(3).Visible = True
    
    'Hide other labels
    
    Me.Label6(4).Visible = False
    Me.Label6(5).Visible = False
    Me.Label6(6).Visible = False
    Me.Label6(8).Visible = False
    Me.Label6(9).Visible = False
    
    Atri(1).Visible = False
    Atri(2).Visible = False
    Atri(3).Visible = False
    Atri(4).Visible = False
    Atri(5).Visible = False
    Me.lblPuntosPesca.Visible = False
    Exit Sub
showCombateLabels_Err:
    Call TraceError(Err.Number, Err.Description, "frmStatistics.showCombateLabels", Erl)
End Sub

Private Sub showStatsLabels()
    On Error Goto showStatsLabels_Err
    
    'Show combate labels
    Me.Label6(0).Visible = False
    Me.Label6(1).Visible = False
    Me.Label6(3).Visible = False
    
    'Hide other labels
    
    Me.Label6(4).Visible = True
    Me.Label6(5).Visible = True
    Me.Label6(6).Visible = True
    Me.Label6(8).Visible = True
    Me.Label6(9).Visible = True
    
    Atri(1).Visible = True
    Atri(2).Visible = True
    Atri(3).Visible = True
    Atri(4).Visible = True
    Atri(5).Visible = True
    Me.lblPuntosPesca.Visible = False
    Exit Sub
showStatsLabels_Err:
    Call TraceError(Err.Number, Err.Description, "frmStatistics.showStatsLabels", Erl)
End Sub

Private Sub showPescaLabels()
    On Error Goto showPescaLabels_Err
        
    Me.lblPuntosPesca.Visible = True
    'Show combate labels
    Me.Label6(0).Visible = False
    Me.Label6(1).Visible = False
    Me.Label6(3).Visible = False
    
    'Hide other labels
    
    Me.Label6(4).Visible = False
    Me.Label6(5).Visible = False
    Me.Label6(6).Visible = False
    Me.Label6(8).Visible = False
    Me.Label6(9).Visible = False
    
    Atri(1).Visible = False
    Atri(2).Visible = False
    Atri(3).Visible = False
    Atri(4).Visible = False
    Atri(5).Visible = False
    Exit Sub
showPescaLabels_Err:
    Call TraceError(Err.Number, Err.Description, "frmStatistics.showPescaLabels", Erl)
End Sub
Private Sub loadButtons()
    On Error Goto loadButtons_Err
       
    Set cBotonAceptar = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonEstadisticasPersonaje = New clsGraphicalButton
    Set cBotonCombate = New clsGraphicalButton
    Set cBotonPesca = New clsGraphicalButton

    Call cBotonAceptar.Initialize(Image1, "boton-aceptar-default.bmp", _
                                                "boton-aceptar-over.bmp", _
                                                "boton-aceptar-off.bmp", Me)
                                                
                                                
    Call cBotonCerrar.Initialize(imgCerrar, "boton-cerrar-default.bmp", _
                                                "boton-cerrar-over.bmp", _
                                                "boton-cerrar-off.bmp", Me)
    
    Call cBotonEstadisticasPersonaje.Initialize(ImgEstadisticasPersonaje, "boton-personaje-default.bmp", _
                                                "boton-personaje-over.bmp", _
                                                "boton-personaje-off.bmp", Me)
    
    Call cBotonCombate.Initialize(ImgCombate, "boton-retos-default.bmp", _
                                                "boton-retos-over.bmp", _
                                                "boton-retos-off.bmp", Me)
    
    Call cBotonPesca.Initialize(ImgPesca, "boton-pesca-default.bmp", _
                                                 "boton-pesca-over.bmp", _
                                                "boton-pesca-off.bmp", Me)
    
    Exit Sub
loadButtons_Err:
    Call TraceError(Err.Number, Err.Description, "frmStatistics.loadButtons", Erl)
End Sub
Private Sub Image1_Click()
    On Error Goto Image1_Click_Err
    Unload Me
    Exit Sub
Image1_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmStatistics.Image1_Click", Erl)
End Sub


Private Sub imgCerrar_Click()
    On Error Goto imgCerrar_Click_Err
    Unload Me
    Exit Sub
imgCerrar_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmStatistics.imgCerrar_Click", Erl)
End Sub

Private Sub ImgCombate_Click()
    On Error Goto ImgCombate_Click_Err
        showCombateLabels
        Me.Picture = LoadInterface("ventanaestadisticas_combate.bmp")
    Exit Sub
ImgCombate_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmStatistics.ImgCombate_Click", Erl)
End Sub

Private Sub ImgEstadisticasPersonaje_Click()
    On Error Goto ImgEstadisticasPersonaje_Click_Err
        showStatsLabels
        Me.Picture = LoadInterface("ventanaestadisticas_personaje.bmp")

    Exit Sub
ImgEstadisticasPersonaje_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmStatistics.ImgEstadisticasPersonaje_Click", Erl)
End Sub

Private Sub ImgPesca_Click()
    On Error Goto ImgPesca_Click_Err
        showPescaLabels
        Me.Picture = LoadInterface("ventanaestadisticas_pesca.bmp")
    Exit Sub
ImgPesca_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmStatistics.ImgPesca_Click", Erl)
End Sub

