VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Argentum20"
   ClientHeight    =   11565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16725
   Icon            =   "frmCrearPersonaje.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   771
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1115
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox lstHogar 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   13680
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   5640
      Width           =   1785
   End
   Begin VB.ComboBox cabeza 
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":000C
      Left            =   840
      List            =   "frmCrearPersonaje.frx":000E
      TabIndex        =   24
      Text            =   "Combo1"
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.ComboBox lstRaza 
      BackColor       =   &H00646401&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      ItemData        =   "frmCrearPersonaje.frx":0010
      Left            =   13680
      List            =   "frmCrearPersonaje.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   4440
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.ComboBox lstGenero 
      BackColor       =   &H00646401&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      ItemData        =   "frmCrearPersonaje.frx":0014
      Left            =   13680
      List            =   "frmCrearPersonaje.frx":001E
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   5040
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.ComboBox lstProfesion 
      BackColor       =   &H00646401&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      ItemData        =   "frmCrearPersonaje.frx":0031
      Left            =   13680
      List            =   "frmCrearPersonaje.frx":0033
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   3840
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   1
      Left            =   7320
      TabIndex        =   28
      Top             =   5250
      Width           =   255
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Height          =   255
      Index           =   0
      Left            =   6360
      TabIndex        =   27
      Top             =   5250
      Width           =   255
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6360
      TabIndex        =   26
      Top             =   4620
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7320
      TabIndex        =   25
      Top             =   4620
      Width           =   255
   End
   Begin VB.Label lbFuerza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6720
      TabIndex        =   23
      Top             =   6630
      Width           =   525
   End
   Begin VB.Label lbAgilidad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6720
      TabIndex        =   22
      Top             =   6855
      Width           =   510
   End
   Begin VB.Label lbConstitucion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6720
      TabIndex        =   21
      Top             =   7530
      Width           =   525
   End
   Begin VB.Label lbInteligencia 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6720
      TabIndex        =   20
      Top             =   7080
      Width           =   525
   End
   Begin VB.Label lbCarisma 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6720
      TabIndex        =   19
      Top             =   7305
      Width           =   510
   End
   Begin VB.Label lbLagaRulzz 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "40"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7200
      TabIndex        =   18
      Top             =   8025
      Width           =   255
   End
   Begin VB.Label SumarFrz 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   17
      Top             =   6600
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   16
      Top             =   6630
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   15
      Top             =   6825
      Width           =   255
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   14
      Top             =   6870
      Width           =   255
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   7245
      TabIndex        =   13
      Top             =   7110
      Width           =   225
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   6480
      TabIndex        =   12
      Top             =   7095
      Width           =   255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   11
      Top             =   7275
      Width           =   255
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   10
      Top             =   7500
      Width           =   255
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7245
      TabIndex        =   9
      Top             =   7290
      Width           =   255
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7245
      TabIndex        =   8
      Top             =   7500
      Width           =   255
   End
   Begin VB.Label modfuerza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+ 0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   7560
      TabIndex        =   7
      Top             =   6600
      Width           =   255
   End
   Begin VB.Label modAgilidad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+ 0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   7560
      TabIndex        =   6
      Top             =   6840
      Width           =   255
   End
   Begin VB.Label modInteligencia 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+ 0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   7560
      TabIndex        =   5
      Top             =   7080
      Width           =   255
   End
   Begin VB.Label modCarisma 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+ 0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   7560
      TabIndex        =   4
      Top             =   7305
      Width           =   255
   End
   Begin VB.Label modConstitucion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+ 0"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   7560
      TabIndex        =   3
      Top             =   7530
      Width           =   255
   End
End
Attribute VB_Name = "frmCrearPersonaje"
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

Dim AnimHead       As Byte

Public SkillPoints As Byte

Function CheckData() As Boolean
    
    On Error GoTo CheckData_Err
    

    If UserRaza = 0 Then
        frmMensaje.Show , frmConnect
        frmMensaje.msg.Caption = "Seleccione la raza del personaje."
        Exit Function

    End If

    If MiCabeza = 0 Then
        frmMensaje.Show , frmConnect
        frmMensaje.msg.Caption = "Seleccione una cabeza para el personaje."
        Exit Function

    End If

    If UserSexo = 0 Then
        frmMensaje.Show , frmConnect
        frmMensaje.msg.Caption = "Seleccione el sexo del personaje."
        Exit Function

    End If
    
    
    
    If UserHogar = 0 Then
        frmMensaje.Show , frmConnect
        frmMensaje.msg.Caption = "Seleccione el hogar del personaje."
        Exit Function

    End If
    

    If UserClase = 0 Then
        frmMensaje.Show , frmConnect
        frmMensaje.msg.Caption = "Seleccione la clase del personaje."
        Exit Function

    End If

    CheckData = True

    
    Exit Function

CheckData_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCrearPersonaje.CheckData", Erl)
    Resume Next
    
End Function

Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single
    
    On Error GoTo RandomNumber_Err
    
    Randomize Timer
    RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound

    If RandomNumber > UpperBound Then RandomNumber = UpperBound

    
    Exit Function

RandomNumber_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCrearPersonaje.RandomNumber", Erl)
    Resume Next
    
End Function

Private Sub Form_Activate()
    g_game_state.state = e_state_createchar_screen
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call FormParser.Parse_Form(Me)
    
    Call SetRGBA(COLOR_RED(0), 255, 0, 0)
    Call SetRGBA(COLOR_RED(1), 255, 0, 0)
    Call SetRGBA(COLOR_RED(2), 255, 0, 0)
    Call SetRGBA(COLOR_RED(3), 255, 0, 0)
    Call SetRGBA(COLOR_GREEN(0), 0, 255, 0)
    Call SetRGBA(COLOR_GREEN(1), 0, 255, 0)
    Call SetRGBA(COLOR_GREEN(2), 0, 255, 0)
    Call SetRGBA(COLOR_GREEN(3), 0, 255, 0)
    

    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCrearPersonaje.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub lstProfesion_Click()
    
    On Error GoTo lstProfesion_Click_Err
    

    'LstFamiliar.Clear
    'Dim i As Byte
    Select Case (lstProfesion.List(lstProfesion.ListIndex))

        Case Is = "Mago"
            RazaRecomendada = "Gnomo / Elfo / Humano / Elfo Drow"
            CPBody = 1
            CPBodyE = 52
            CPArma = 79
            CPGorro = 58
            CPEscudo = 0
            CPAura = "35532:&HDD7C40:0:248"

            'CPHead=
        Case Is = "Paladin"
            RazaRecomendada = "Humano / Elfo Drow / Elfo"
            CPBody = 1
            CPBodyE = 52
            CPArma = 41
            CPGorro = 73
            CPEscudo = 72
            CPAura = "35448:&HFFF306:0:248"

        Case Is = "Cazador"
            RazaRecomendada = "Enano / Humano"
            CPBody = 1
            CPBodyE = 52
            CPArma = 126
            CPGorro = 78
            CPEscudo = 51
            CPAura = "20200:&H904D17:0:248"

        Case Is = "Guerrero"
            RazaRecomendada = "Enano / Humano"
            CPBody = 1
            CPBodyE = 52
            CPArma = 81
            CPGorro = 79
            CPEscudo = 73
            CPAura = "35498:&H8700CE:0:248"

        Case Is = "Bardo"
            RazaRecomendada = "Elfo / Humano"
            CPBody = 1
            CPBodyE = 52
            CPArma = 76
            CPGorro = 43
            CPEscudo = 48
            CPAura = "35445:&H800080:0:248"

        Case Is = "Clerigo"
            RazaRecomendada = "Humano / Elfo Drow / Elfo"
            CPBody = 1
            CPBodyE = 52
            CPArma = 72
            CPGorro = 83
            CPEscudo = 60
            CPAura = "35443:&H83CEDD:0:248"

        Case Is = "Asesino"
            RazaRecomendada = "Humano / ElfoDrow"
            CPBody = 1
            CPBodyE = 52
            CPArma = 40
            CPGorro = 74
            CPEscudo = 58
            CPAura = "35432:&HFB0813:0:248"

        Case Is = "Druida"
            RazaRecomendada = "Humano / Elfo"
            CPBody = 1
            CPBodyE = 52
            CPArma = 100
            CPGorro = 70
            CPEscudo = 40
            CPAura = "35466:&HF622D:0:248"

        Case Is = "Trabajador"
            RazaRecomendada = "Enano / Humano"
            CPBody = 1
            CPBodyE = 52
            CPArma = 143
            CPGorro = 95
            CPEscudo = 0
            CPAura = ""

        Case "Bandido", "Ladrón", "Pirata"
            RazaRecomendada = "Enano / Humano"
            CPBody = 1
            CPBodyE = 52
            CPArma = 143
            CPGorro = 95
            CPEscudo = 0
            CPAura = ""

    End Select
 
    RazaRecomendada = RazaRecomendada

    
    Exit Sub

lstProfesion_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCrearPersonaje.lstProfesion_Click", Erl)
    Resume Next
    
End Sub

Private Sub lstRaza_Click()
    
    On Error GoTo lstRaza_Click_Err
    

    If lstRaza.ListIndex < 0 Then Exit Sub
    
    Dim i As Integer

    i = lstRaza.ListIndex + 1

    Call DameOpciones

    AnimHead = 3

    modfuerza.Caption = IIf(Sgn(ModRaza(i).Fuerza) < 0, "-", "+") & " " & Abs(ModRaza(i).Fuerza)
    modAgilidad.Caption = IIf(Sgn(ModRaza(i).Agilidad) < 0, "-", "+") & " " & Abs(ModRaza(i).Agilidad)
    modInteligencia.Caption = IIf(Sgn(ModRaza(i).Inteligencia) < 0, "-", "+") & " " & Abs(ModRaza(i).Inteligencia)
    modConstitucion.Caption = IIf(Sgn(ModRaza(i).Constitucion) < 0, "-", "+") & " " & Abs(ModRaza(i).Constitucion)
    modCarisma.Caption = IIf(Sgn(ModRaza(i).Carisma) < 0, "-", "+") & " " & Abs(ModRaza(i).Carisma)

    
    Exit Sub

lstRaza_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCrearPersonaje.lstRaza_Click", Erl)
    Resume Next
    
End Sub

Private Sub render_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo render_MouseUp_Err
    

    If x > 331 And x < 347 And y > 412 And y < 424 Then 'Boton izquierda cabezas
        If Cabeza.ListCount = 0 Then Exit Sub
        If Cabeza.ListIndex > 0 Then
            Cabeza.ListIndex = Cabeza.ListIndex - 1

        End If

        If Cabeza.ListIndex = 0 Then
            Cabeza.ListIndex = Cabeza.ListCount - 1

        End If

    End If

    If x > 401 And x < 415 And y > 412 And y < 424 Then 'Boton Derecha cabezas
        If Cabeza.ListCount = 0 Then Exit Sub
        If (Cabeza.ListIndex + 1) <> Cabeza.ListCount Then
            Cabeza.ListIndex = Cabeza.ListIndex + 1

        End If

        If (Cabeza.ListIndex + 1) = Cabeza.ListCount Then
            Cabeza.ListIndex = 0

        End If

    End If

    If x > 348 And x < 408 And y > 511 And y < 523 Then 'Boton Equipar
        CPEquipado = Not CPEquipado

    End If

    If x > 290 And x < 326 And y > 453 And y < 486 Then 'Boton Equipar
        If CPHeading + 1 >= 5 Then
            CPHeading = 1
        Else
            CPHeading = CPHeading + 1

        End If

    End If

    If x > 421 And x < 452 And y > 453 And y < 486 Then 'Boton Equipar
        If CPHeading - 1 <= 0 Then
            CPHeading = 4
        Else
            CPHeading = CPHeading - 1

        End If

    End If

    If x > 548 And x < 560 And y > 258 And y < 271 Then 'Boton Derecha cabezas

        If lstProfesion.ListIndex < lstProfesion.ListCount - 1 Then
            lstProfesion.ListIndex = lstProfesion.ListIndex + 1
        Else
            lstProfesion.ListIndex = 0

        End If

    End If

    If x > 435 And x < 446 And y > 260 And y < 271 Then 'Boton Derecha cabezas

        If lstProfesion.ListIndex - 1 < 0 Then
            lstProfesion.ListIndex = lstProfesion.ListCount - 1
        Else
            lstProfesion.ListIndex = lstProfesion.ListIndex - 1

        End If

    End If

    If x > 548 And x < 560 And y > 304 And y < 323 Then 'Boton Derecha cabezas
        If lstRaza.ListIndex < lstRaza.ListCount - 1 Then
            lstRaza.ListIndex = lstRaza.ListIndex + 1
        Else
            lstRaza.ListIndex = 0

        End If

    End If

    If x > 435 And x < 446 And y > 304 And y < 323 Then 'Boton Derecha cabezas
        If lstRaza.ListIndex - 1 < 0 Then
            lstRaza.ListIndex = lstRaza.ListCount - 1
        Else
            lstRaza.ListIndex = lstRaza.ListIndex - 1

        End If

    End If

    If x > 548 And x < 560 And y > 351 And y < 367 Then 'Boton Derecha cabezas
        If lstGenero.ListIndex < lstGenero.ListCount - 1 Then
            lstGenero.ListIndex = lstGenero.ListIndex + 1
        Else
            lstGenero.ListIndex = 0

        End If

    End If

    If x > 435 And x < 446 And y > 351 And y < 367 Then 'Boton Derecha cabezas
        If lstGenero.ListIndex - 1 < 0 Then
            lstGenero.ListIndex = lstGenero.ListCount - 1
        Else
            lstGenero.ListIndex = lstGenero.ListIndex - 1

        End If

    End If

    If x > 148 And x < 246 And y > 630 And y < 670 Then 'Boton > Volver
        Call Sound.Sound_Play(SND_CLICK)

        UserMap = 307
        AlphaNiebla = 25
        EntradaY = 1
        EntradaX = 1
    
        Call SwitchMap(UserMap)
       
        frmConnect.Visible = True
        g_game_state.state = e_state_account_screen
        Unload Me

    End If

    If x > 731 And x < 829 And y > 630 And y < 670 Then 'Boton > Crear
        Call Sound.Sound_Play(SND_CLICK)

        Dim k As Object

        If StopCreandoCuenta = True Then Exit Sub
            
        If Right$(UserName, 1) = " " Then
            UserName = RTrim$(UserName)

            'MsgBox "Nombre invalido, se han removido los espacios al final del nombre"
        End If
            
        UserRaza = lstRaza.ListIndex + 1
        UserSexo = lstGenero.ListIndex + 1
        UserClase = lstProfesion.ListIndex + 1
        
        UserHogar = lstHogar.ListIndex + 1
            
        UserAtributos(1) = Val(lbFuerza.Caption) + Val(modfuerza.Caption)
        UserAtributos(2) = Val(lbAgilidad.Caption) + Val(modAgilidad.Caption)
        UserAtributos(3) = Val(lbInteligencia.Caption) + Val(modInteligencia.Caption)
        UserAtributos(4) = Val(lbConstitucion.Caption) + Val(modConstitucion.Caption)
        UserAtributos(4) = Val(lbCarisma.Caption) + Val(modCarisma.Caption)
            
        'Ladder Atributos Negativos
        If UserAtributos(1) < 1 Then UserAtributos(1) = 1
        If UserAtributos(2) < 1 Then UserAtributos(2) = 1
        If UserAtributos(3) < 1 Then UserAtributos(3) = 1
        If UserAtributos(4) < 1 Then UserAtributos(4) = 1
        If UserAtributos(5) < 1 Then UserAtributos(5) = 1

        'Barrin 3/10/03
        If CheckData() Then
            UserPassword = CuentaPassword
            'UserEmail = "noseusa@a.com"

            StopCreandoCuenta = True
                
            If Connected Then
                frmMain.ShowFPS.Enabled = True
            End If
            frmConnecting.Show
            Call modNetwork.Connect(IPdelServidor, PuertoDelServidor)
            Call LoginOrConnect(E_MODO.CrearNuevoPj)
        End If

    End If

    
    Exit Sub

render_MouseUp_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCrearPersonaje.render_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub Cabeza_Click()
    
    On Error GoTo Cabeza_Click_Err
    
    MiCabeza = Val(Cabeza.List(Cabeza.ListIndex))
    Call DibujarCPJ(MiCabeza, 3)

    CPHead = MiCabeza

    
    Exit Sub

Cabeza_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCrearPersonaje.Cabeza_Click", Erl)
    Resume Next
    
End Sub
 
Private Sub lstGenero_Click()
    
    On Error GoTo lstGenero_Click_Err
    
    Call DameOpciones
    AnimHead = 3

    
    Exit Sub

lstGenero_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCrearPersonaje.lstGenero_Click", Erl)
    Resume Next
    
End Sub
