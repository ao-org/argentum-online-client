VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "RevolucionAo 1.0"
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
   Begin VB.ComboBox cabeza 
      Height          =   315
      ItemData        =   "frmCrearPersonaje.frx":57E2
      Left            =   840
      List            =   "frmCrearPersonaje.frx":57E4
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
      ItemData        =   "frmCrearPersonaje.frx":57E6
      Left            =   13680
      List            =   "frmCrearPersonaje.frx":57E8
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
      ItemData        =   "frmCrearPersonaje.frx":57EA
      Left            =   13680
      List            =   "frmCrearPersonaje.frx":57F4
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
      ItemData        =   "frmCrearPersonaje.frx":5807
      Left            =   13680
      List            =   "frmCrearPersonaje.frx":5809
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
      Caption         =   "0"
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
      Caption         =   "0"
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
      Caption         =   "0"
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
      Caption         =   "0"
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
      Caption         =   "0"
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
'RevolucionAo 1.0
'Pablo Mercavides

Option Explicit

Dim AnimHead As Byte



Public SkillPoints As Byte
Function CheckData() As Boolean
If UserRaza = 0 Then
    MsgBox "Seleccione la raza del personaje."
    Exit Function
End If

If MiCabeza = 0 Then
    MsgBox "Seleccione una cabeza para el personaje."
    Exit Function
End If


If UserSexo = 0 Then
    MsgBox "Seleccione el sexo del personaje."
    Exit Function
End If

If UserClase = 0 Then
    MsgBox "Seleccione la clase del personaje."
    Exit Function
End If


CheckData = True


End Function
Function RandomNumber(ByVal LowerBound As Variant, ByVal UpperBound As Variant) As Single
    Randomize Timer
    RandomNumber = (UpperBound - LowerBound + 1) * Rnd + LowerBound
    If RandomNumber > UpperBound Then RandomNumber = UpperBound
End Function
Private Sub Form_Activate()
QueRender = 3
End Sub

Private Sub Form_Load()
Call FormParser.Parse_Form(Me)
'Call SwitchMapIAO(281)



'lstProfesion.ListIndex = 0
'LstFamiliar.ListIndex = 0


End Sub



Private Sub lstProfesion_Click()
'LstFamiliar.Clear
'Dim i As Byte
Select Case (lstProfesion.List(lstProfesion.ListIndex))
    Case Is = "Mago"
            RazaRecomendada = "Gnomo/Elfo/Humano/Drow"
            CPBody = 640
            CPBodyE = 641
            CPArma = 79
            CPGorro = 58
            CPEscudo = 0
            CPAura = "35532:&HDD7C40:0:248"
            'CPHead=
    Case Is = "Paladin"
            RazaRecomendada = "Humano/ElfoDrow/Elfo"
            CPBody = 524
            CPBodyE = 525
            CPArma = 41
            CPGorro = 73
            CPEscudo = 72
            CPAura = "35448:&HFFF306:0:248"
    Case Is = "Cazador"
            RazaRecomendada = "Orco/Enano"
            CPBody = 526
            CPBodyE = 527
            CPArma = 126
            CPGorro = 78
            CPEscudo = 51
            CPAura = "20200:&H904D17:0:248"
    Case Is = "Guerrero"
            RazaRecomendada = "Orco/Enano"
            CPBody = 528
            CPBodyE = 529
            CPArma = 81
            CPGorro = 79
            CPEscudo = 73
            CPAura = "35498:&H8700CE:0:248"
    Case Is = "Bardo"
            RazaRecomendada = "Elfo/Humano/Orco"
            CPBody = 636
            CPBodyE = 637
            CPArma = 76
            CPGorro = 43
            CPEscudo = 48
            CPAura = "35445:&H800080:0:248"
    Case Is = "Clerigo"
            RazaRecomendada = "Humano/Elfo Drow/Elfo"
            CPBody = 520
            CPBodyE = 521
            CPArma = 72
            CPGorro = 83
            CPEscudo = 60
            CPAura = "35443:&H83CEDD:0:248"
    Case Is = "Asesino"
            RazaRecomendada = "Humano/ElfoDrow"
            CPBody = 518
            CPBodyE = 519
            CPArma = 40
            CPGorro = 74
            CPEscudo = 58
            CPAura = "35432:&HFB0813:0:248"
    Case Is = "Druida"
            RazaRecomendada = "Humano/Elfo"
            CPBody = 632
            CPBodyE = 633
            CPArma = 100
            CPGorro = 70
            CPEscudo = 40
            CPAura = "35466:&HF622D:0:248"
    Case Is = "Buscavidas"
            RazaRecomendada = "Enano/Orco/Humano"
            CPBody = 602
            CPBodyE = 603
            CPArma = 143
            CPGorro = 95
            CPEscudo = 0
            CPAura = ""
 End Select
 
 RazaRecomendada = "Raza sugerida: " & RazaRecomendada

End Sub

Private Sub lstRaza_Click()
Call DameOpciones
AnimHead = 3
Select Case (lstRaza.List(lstRaza.ListIndex))
    Case Is = "Humano"
        frmCrearPersonaje.modfuerza.Caption = "+ 1"
        frmCrearPersonaje.modConstitucion.Caption = "+ 2"
       frmCrearPersonaje.modAgilidad.Caption = "+ 1"
       frmCrearPersonaje.modInteligencia.Caption = "+ 1"
    Case Is = "Elfo"
        modfuerza.Caption = ""
        modConstitucion.Caption = ""
        modAgilidad.Caption = "+ 2"
        modInteligencia.Caption = "+ 3"
    Case Is = "Elfo Drow"
        modfuerza.Caption = "+ 2"
        modConstitucion.Caption = "+ 1"
        modAgilidad.Caption = ""
        modInteligencia.Caption = "+ 2"
    Case Is = "Enano"
        modfuerza.Caption = "+ 3"
        modConstitucion.Caption = "+ 4"
        modAgilidad.Caption = "- 1"
        modInteligencia.Caption = "- 6"
    Case Is = "Gnomo"
        modfuerza.Caption = "- 5"
        modAgilidad.Caption = "+ 3"
        modInteligencia.Caption = "+ 4"
        modConstitucion.Caption = "- 1"
    Case Is = "Orco"
        modfuerza.Caption = "+ 5"
        modAgilidad.Caption = "- 1"
        modInteligencia.Caption = "- 6"
        modConstitucion.Caption = "+ 3"
End Select
End Sub




Private Sub render_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)


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






'If X > 258 And X < 272 And Y > 413 And Y < 421 Then 'Boton < FUERZA
'    Call Sound.Sound_Play(SND_CLICK)
'    If Not frmCrearPersonaje.lbFuerza.Caption = 6 Then
'    frmCrearPersonaje.lbLagaRulzz.Caption = frmCrearPersonaje.lbLagaRulzz.Caption + 1
'    frmCrearPersonaje.lbFuerza.Caption = frmCrearPersonaje.lbFuerza.Caption - 1
'End If


'End If


'If X > 258 And X < 272 And Y > 442 And Y < 454 Then 'Boton < Agilidad
'Call Sound.Sound_Play(SND_CLICK)
'If frmCrearPersonaje.lbAgilidad.Caption = 6 Then Exit Sub
'frmCrearPersonaje.lbLagaRulzz.Caption = frmCrearPersonaje.lbLagaRulzz.Caption + 1
'frmCrearPersonaje.lbAgilidad.Caption = frmCrearPersonaje.lbAgilidad.Caption - 1

'End If

'If X > 258 And X < 272 And Y > 474 And Y < 483 Then 'Boton < Inteligencia

'Call Sound.Sound_Play(SND_CLICK)
'If frmCrearPersonaje.lbInteligencia.Caption = 6 Then Exit Sub
'frmCrearPersonaje.lbLagaRulzz.Caption = frmCrearPersonaje.lbLagaRulzz.Caption + 1
'frmCrearPersonaje.lbInteligencia.Caption = frmCrearPersonaje.lbInteligencia.Caption - 1
'End If

'If X > 258 And X < 272 And Y > 505 And Y < 517 Then 'Boton < Carisma
'Call Sound.Sound_Play(SND_CLICK)
'If frmCrearPersonaje.lbCarisma.Caption = 6 Then Exit Sub
'frmCrearPersonaje.lbLagaRulzz.Caption = frmCrearPersonaje.lbLagaRulzz.Caption + 1
'frmCrearPersonaje.lbCarisma.Caption = frmCrearPersonaje.lbCarisma.Caption - 1
'End If


'If X > 258 And X < 272 And Y > 500 And Y < 517 Then 'Boton < Constitucion
'    Call Sound.Sound_Play(SND_CLICK)
'    If frmCrearPersonaje.lbConstitucion.Caption = 6 Then Exit Sub
'    frmCrearPersonaje.lbLagaRulzz.Caption = frmCrearPersonaje.lbLagaRulzz.Caption + 1
'    frmCrearPersonaje.lbConstitucion.Caption = frmCrearPersonaje.lbConstitucion.Caption - 1
'End If



'If X > 308 And X < 320 And Y > 411 And Y < 424 Then 'Boton > Fuerza
'Call Sound.Sound_Play(SND_CLICK)
'If Not frmCrearPersonaje.lbLagaRulzz.Caption > 0 Then Exit Sub
'If Not frmCrearPersonaje.lbFuerza.Caption = 18 Then
'frmCrearPersonaje.lbLagaRulzz.Caption = frmCrearPersonaje.lbLagaRulzz.Caption - 1
'frmCrearPersonaje.lbFuerza.Caption = frmCrearPersonaje.lbFuerza.Caption + 1
'End If
'End If


'If X > 308 And X < 320 And Y > 442 And Y < 454 Then 'Boton > Agilidad
'Call Sound.Sound_Play(SND_CLICK)
'If Not frmCrearPersonaje.lbLagaRulzz.Caption > 0 Then Exit Sub
'If Not frmCrearPersonaje.lbAgilidad.Caption = 18 Then
'frmCrearPersonaje.lbLagaRulzz.Caption = frmCrearPersonaje.lbLagaRulzz.Caption - 1
'frmCrearPersonaje.lbAgilidad.Caption = frmCrearPersonaje.lbAgilidad.Caption + 1
'End If

'End If


'If X > 308 And X < 320 And Y > 472 And Y < 485 Then 'Boton > Inteligencia
'Call Sound.Sound_Play(SND_CLICK)
'If Not frmCrearPersonaje.lbLagaRulzz.Caption > 0 Then Exit Sub
'If Not frmCrearPersonaje.lbInteligencia.Caption = 18 Then
'frmCrearPersonaje.lbLagaRulzz.Caption = frmCrearPersonaje.lbLagaRulzz.Caption - 1
'frmCrearPersonaje.lbInteligencia.Caption = frmCrearPersonaje.lbInteligencia.Caption + 1
'End If

'End If

'If X > 308 And X < 320 And Y > 504 And Y < 516 Then 'Boton > Carisma
'Call Sound.Sound_Play(SND_CLICK)
'If Not frmCrearPersonaje.lbLagaRulzz.Caption > 0 Then Exit Sub
'If Not frmCrearPersonaje.lbCarisma.Caption = 18 Then
'frmCrearPersonaje.lbLagaRulzz.Caption = frmCrearPersonaje.lbLagaRulzz.Caption - 1
'frmCrearPersonaje.lbCarisma.Caption = frmCrearPersonaje.lbCarisma.Caption + 1
'End If
'End If

'If X > 308 And X < 320 And Y > 497 And Y < 516 Then 'Boton > Constitucion
'Call Sound.Sound_Play(SND_CLICK)
'If Not frmCrearPersonaje.lbLagaRulzz.Caption > 0 Then Exit Sub
'If Not frmCrearPersonaje.lbConstitucion.Caption = 18 Then
'frmCrearPersonaje.lbLagaRulzz.Caption = frmCrearPersonaje.lbLagaRulzz.Caption - 1
'frmCrearPersonaje.lbConstitucion.Caption = frmCrearPersonaje.lbConstitucion.Caption + 1
'End If
'End If

If x > 148 And x < 246 And y > 630 And y < 670 Then 'Boton > Volver
Call Sound.Sound_Play(SND_CLICK)

    UserMap = 307
    AlphaNiebla = 25
    EntradaY = 1
    EntradaX = 1
    
    Call SwitchMapIAO(UserMap)
       
            'FrmCuenta.Visible = True
            frmConnect.Visible = True
            QueRender = 2
            
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
            
            UserAtributos(1) = Val(lbFuerza.Caption) + Val(modfuerza.Caption)
            UserAtributos(2) = Val(lbAgilidad.Caption) + Val(modAgilidad.Caption)
            UserAtributos(3) = Val(lbInteligencia.Caption) + Val(modInteligencia.Caption)
            UserAtributos(4) = Val(lbConstitucion.Caption) + Val(modConstitucion.Caption)
            
            'Ladder Atributos Negativos
            If UserAtributos(1) < 1 Then UserAtributos(1) = 1
            If UserAtributos(2) < 1 Then UserAtributos(2) = 1
            If UserAtributos(3) < 1 Then UserAtributos(3) = 1
            If UserAtributos(4) < 1 Then UserAtributos(4) = 1
            
            
           

            'Barrin 3/10/03
            If CheckData() Then
                UserPassword = CuentaPassword
                'UserEmail = "noseusa@a.com"

                StopCreandoCuenta = True
                    
                
                    If frmMain.Socket1.Connected Then
                        EstadoLogin = E_MODO.CrearNuevoPj
                        Call Login
                        frmMain.Second.Enabled = True
                        Exit Sub
                    Else
                        EstadoLogin = E_MODO.CrearNuevoPj
                        frmMain.Socket1.HostName = IPdelServidor
                        frmMain.Socket1.RemotePort = PuertoDelServidor
                        frmMain.Socket1.Connect
                    End If
            End If
End If
End Sub

Private Sub Cabeza_Click()
MiCabeza = Val(Cabeza.List(Cabeza.ListIndex))
Call DibujarCPJ(MiCabeza, 3)

CPHead = MiCabeza

End Sub
 
Private Sub lstGenero_Click()
Call DameOpciones
AnimHead = 3

End Sub


Private Sub Timer1_Timer()

End Sub
