VERSION 5.00
Begin VB.Form frmCustomKeys 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuración de controles"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6075
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
   ScaleHeight     =   407
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txConfig 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   22
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   61
      Top             =   1440
      Width           =   1770
   End
   Begin VB.TextBox txConfig 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   21
      Left            =   4125
      Locked          =   -1  'True
      TabIndex        =   58
      Top             =   840
      Width           =   1770
   End
   Begin VB.TextBox txConfig 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   20
      Left            =   4125
      Locked          =   -1  'True
      TabIndex        =   57
      Top             =   240
      Width           =   1770
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   19
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   55
      Top             =   3240
      Width           =   1770
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Moderna"
      Height          =   255
      Left            =   4110
      TabIndex        =   53
      Top             =   4440
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Clásica"
      Height          =   255
      Left            =   4110
      TabIndex        =   52
      Top             =   4080
      Width           =   1215
   End
   Begin VB.ComboBox AccionList3 
      Height          =   315
      ItemData        =   "frmCustomKeys.frx":0000
      Left            =   4110
      List            =   "frmCustomKeys.frx":0013
      Style           =   2  'Dropdown List
      TabIndex        =   49
      Top             =   3360
      Width           =   1815
   End
   Begin VB.ComboBox AccionList1 
      Height          =   315
      ItemData        =   "frmCustomKeys.frx":007B
      Left            =   4110
      List            =   "frmCustomKeys.frx":008E
      Style           =   2  'Dropdown List
      TabIndex        =   46
      Top             =   2160
      Width           =   1815
   End
   Begin VB.ComboBox AccionList2 
      Height          =   315
      ItemData        =   "frmCustomKeys.frx":00F6
      Left            =   4110
      List            =   "frmCustomKeys.frx":0109
      Style           =   2  'Dropdown List
      TabIndex        =   45
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Salir"
      Height          =   495
      Index           =   2
      Left            =   4080
      TabIndex        =   44
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton cmdAccion 
      Caption         =   "Guardar"
      Height          =   495
      Index           =   0
      Left            =   4080
      TabIndex        =   43
      Top             =   4800
      Width           =   1695
   End
   Begin VB.TextBox txConfig 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   0
      Left            =   195
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   240
      Width           =   1770
   End
   Begin VB.TextBox txConfig 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   1
      Left            =   195
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   840
      Width           =   1770
   End
   Begin VB.TextBox txConfig 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   2
      Left            =   195
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   1440
      Width           =   1770
   End
   Begin VB.TextBox txConfig 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   3
      Left            =   195
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   2040
      Width           =   1770
   End
   Begin VB.TextBox txConfig 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   4
      Left            =   195
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   2640
      Width           =   1770
   End
   Begin VB.TextBox txConfig 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   5
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   240
      Width           =   1770
   End
   Begin VB.TextBox txConfig 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   6
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2040
      Width           =   1770
   End
   Begin VB.TextBox txConfig 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   7
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   840
      Width           =   1770
   End
   Begin VB.TextBox txConfig 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   8
      Left            =   195
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   5640
      Width           =   1770
   End
   Begin VB.TextBox txConfig 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   9
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   3840
      Width           =   1770
   End
   Begin VB.TextBox txConfig 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   10
      Left            =   2115
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   5640
      Width           =   1770
   End
   Begin VB.TextBox txConfig 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   11
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1440
      Width           =   1770
   End
   Begin VB.TextBox txConfig 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   12
      Left            =   2175
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   5040
      Width           =   1770
   End
   Begin VB.TextBox txConfig 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   13
      Left            =   195
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   3240
      Width           =   1770
   End
   Begin VB.TextBox txConfig 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   14
      Left            =   195
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3840
      Width           =   1770
   End
   Begin VB.TextBox txConfig 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   15
      Left            =   195
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   4440
      Width           =   1770
   End
   Begin VB.TextBox txConfig 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   16
      Left            =   195
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   5040
      Width           =   1770
   End
   Begin VB.TextBox txConfig 
      Height          =   315
      Index           =   17
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2640
      Width           =   1770
   End
   Begin VB.TextBox txConfig 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   18
      Left            =   2175
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "*"
      Top             =   4440
      Width           =   1770
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "Hablar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   1800
      TabIndex        =   0
      Top             =   8280
      Width           =   3735
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   20
         Left            =   1920
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   19
         Left            =   1920
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Hablar al Clan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Hablar a Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Label lblSalirDel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Domar"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   21
      Left            =   4080
      TabIndex        =   62
      Top             =   1200
      Width           =   465
   End
   Begin VB.Label lblSalirDel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Meditar"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   20
      Left            =   4125
      TabIndex        =   60
      Top             =   0
      Width           =   540
   End
   Begin VB.Label lblSalirDel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Salir del juego"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   19
      Left            =   4125
      TabIndex        =   59
      Top             =   600
      Width           =   1005
   End
   Begin VB.Label lblSalirDel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Clan) Marca de Clan"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   36
      Left            =   2160
      TabIndex        =   56
      Top             =   3000
      Width           =   1500
   End
   Begin VB.Label lblSalirDel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(Clan) Llamada de Clan"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   35
      Left            =   2160
      TabIndex        =   54
      Top             =   2400
      Width           =   1650
   End
   Begin VB.Label lblSalirDel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Configuración rapida:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   32
      Left            =   4110
      TabIndex        =   51
      Top             =   3720
      Width           =   1785
   End
   Begin VB.Label lblSalirDel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Acción Click 3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   31
      Left            =   4110
      TabIndex        =   50
      Top             =   3120
      Width           =   1140
   End
   Begin VB.Label lblSalirDel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Acción Click 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   33
      Left            =   4110
      TabIndex        =   48
      Top             =   1920
      Width           =   1140
   End
   Begin VB.Label lblSalirDel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Acción Click 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   34
      Left            =   4110
      TabIndex        =   47
      Top             =   2520
      Width           =   1140
   End
   Begin VB.Label lblSalirDel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tomar screenshot"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   16
      Left            =   2175
      TabIndex        =   42
      Top             =   4800
      Width           =   1290
   End
   Begin VB.Label lblSalirDel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estadisticas del juego"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   14
      Left            =   2175
      TabIndex        =   41
      Top             =   4200
      Width           =   1545
   End
   Begin VB.Label lblSalirDel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Moverse hacia arriba"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   13
      Left            =   195
      TabIndex        =   40
      Top             =   3000
      Width           =   1500
   End
   Begin VB.Label lblSalirDel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Moverse hacia Derecha"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   11
      Left            =   195
      TabIndex        =   39
      Top             =   4800
      Width           =   1680
   End
   Begin VB.Label lblSalirDel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Moverse hacia Izquierda"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   10
      Left            =   195
      TabIndex        =   38
      Top             =   4200
      Width           =   1755
   End
   Begin VB.Label lblSalirDel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Moverse hacia abajo"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   9
      Left            =   195
      TabIndex        =   37
      Top             =   3600
      Width           =   1485
   End
   Begin VB.Label lblSalirDel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mostrar/ocultar macros"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   8
      Left            =   2160
      TabIndex        =   36
      Top             =   1200
      Width           =   1665
   End
   Begin VB.Label lblSalirDel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mostrar/ocultar nombres"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   4
      Left            =   2160
      TabIndex        =   35
      Top             =   1800
      Width           =   1770
   End
   Begin VB.Label lblSalirDel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ocultarse"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   2115
      TabIndex        =   34
      Top             =   5400
      Width           =   690
   End
   Begin VB.Label lblSalirDel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Actualizar Posición"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   2160
      TabIndex        =   33
      Top             =   3600
      Width           =   1320
   End
   Begin VB.Label lblSalirDel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Robar"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   195
      TabIndex        =   32
      Top             =   5400
      Width           =   435
   End
   Begin VB.Label lblSalirDel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seguro de grupo"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   2160
      TabIndex        =   31
      Top             =   600
      Width           =   1200
   End
   Begin VB.Label lblSalirDel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seguro"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   18
      Left            =   2160
      TabIndex        =   30
      Top             =   0
      Width           =   510
   End
   Begin VB.Label lblSalirDel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Equipar objeto"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   15
      Left            =   210
      TabIndex        =   29
      Top             =   2400
      Width           =   1050
   End
   Begin VB.Label lblSalirDel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usar objeto"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   7
      Left            =   210
      TabIndex        =   28
      Top             =   1800
      Width           =   840
   End
   Begin VB.Label lblSalirDel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tirar objeto"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   17
      Left            =   210
      TabIndex        =   27
      Top             =   1200
      Width           =   840
   End
   Begin VB.Label lblSalirDel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tomar objeto"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   6
      Left            =   210
      TabIndex        =   26
      Top             =   600
      Width           =   960
   End
   Begin VB.Label lblSalirDel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Atacar"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   5
      Left            =   210
      TabIndex        =   25
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblSalirDel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Modo rol"
      Height          =   195
      Index           =   12
      Left            =   7440
      TabIndex        =   24
      Top             =   7920
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "frmCustomKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private TempVars(0 To 32) As Integer

Private Sub cmdAccion_Click(Index As Integer)

    Dim i         As Integer

    Dim bCambio   As Boolean

    Dim Resultado As VbMsgBoxResult

    Select Case Index
    
        Case 0
            Call GuardaConfigEnVariables
            Call SaveRAOInit

        Case 1
            Call LoadDefaultBinds
            Call CargaConfigEnForm
            Call SaveRAOInit

        Case 2
    
            For i = 1 To NUMBINDS

                If TempVars(i - 1) <> BindKeys(i).KeyCode Then
                    bCambio = True
                    Exit For

                End If

            Next

            If bCambio Then
                Resultado = MsgBox("Realizo cambios en la configuración ¿desea guardar antes de salir?", vbQuestion + vbYesNoCancel, "Guardar cambios")

                If Resultado = vbYes Then Call GuardaConfigEnVariables

            End If
        
            If Resultado <> vbCancel Then Unload Me

    End Select

End Sub

Private Sub GuardaConfigEnVariables()

    Dim i As Integer

    For i = 1 To NUMBINDS
        BindKeys(i).name = txConfig(i - 1).Text
        BindKeys(i).KeyCode = TempVars(i - 1)
    Next

    ACCION1 = AccionList1.ListIndex
    ACCION2 = AccionList2.ListIndex
    ACCION3 = AccionList3.ListIndex

End Sub

Private Sub CargaConfigEnForm()

    Dim i As Integer

    For i = 1 To NUMBINDS
        txConfig(i - 1).Text = BindKeys(i).name
        TempVars(i - 1) = BindKeys(i).KeyCode
    Next

    AccionList1.ListIndex = ACCION1
    AccionList2.ListIndex = ACCION2
    AccionList3.ListIndex = ACCION3

End Sub

Private Sub Form_Load()
    Call CargaConfigEnForm
    Call FormParser.Parse_Form(Me)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim i         As Integer

    Dim bCambio   As Boolean

    Dim Resultado As VbMsgBoxResult

    For i = 1 To NUMBINDS

        If TempVars(i - 1) <> BindKeys(i).KeyCode Then
            bCambio = True
            Exit For

        End If

    Next

    If bCambio Then
        Resultado = MsgBox("Realizo cambios en la configuración ¿desea guardar antes de salir?", vbQuestion + vbYesNoCancel, "Guardar cambios")

        If Resultado = vbYes Then Call GuardaConfigEnVariables

    End If

    If Resultado = vbCancel Then Cancel = 1

End Sub

Private Sub Option1_Click()
    Call LoadDefaultBinds
    Call CargaConfigEnForm
    Call SaveRAOInit

End Sub

Private Sub Option2_Click()
    Call LoadDefaultBinds2
    Call CargaConfigEnForm
    Call SaveRAOInit

End Sub

Private Sub txConfig_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

    Dim name As String

    name = txConfig(Index).Text

    If KeyCode > 0 Then
    
        If AlreadyBinded(KeyCode) Then
            Beep
            txConfig(Index).ForeColor = vbRed
            Exit Sub

        End If
    
        If KeyCode = vbKeyShift Then
            name = "Shift"
        ElseIf KeyCode = vbKeyLeft Then
            name = "Flecha Izquierda"
        ElseIf KeyCode = vbKeyRight Then
            name = "Flecha Derecha"
        ElseIf KeyCode = vbKeyDown Then
            name = "Flecha Abajo"
        ElseIf KeyCode = vbKeyUp Then
            name = "Flecha Arriba"
        ElseIf KeyCode = vbKeyControl Then
            name = "Control"
        ElseIf KeyCode = vbKeyPageDown Then
            name = "Page Down"
        ElseIf KeyCode = vbKeyPageUp Then
            name = "Page Up"
        ElseIf KeyCode = vbKeySeparator Then 'Enter teclado numerico
            name = "Intro"
        ElseIf KeyCode = vbKeySpace Then
            name = "Barra Espaciadora"
        ElseIf KeyCode = vbKeyDelete Then
            name = "Delete"
        ElseIf KeyCode = vbKeyEnd Then
            name = "Fin"
        ElseIf KeyCode = vbKeyHome Then
            name = "Inicio"
        ElseIf KeyCode = vbKeyInsert Then
            name = "Insert"
        ElseIf KeyCode = 109 Then
            name = "-"
        ElseIf KeyCode = 112 Then
            name = "F1"
        ElseIf KeyCode = 113 Then
            name = "F2"
        ElseIf KeyCode = 114 Then
            name = "F3"
        ElseIf KeyCode = 115 Then
            name = "F4"
        ElseIf KeyCode = 116 Then
            name = "F5"
        ElseIf KeyCode = 117 Then
            name = "F6"
        ElseIf KeyCode = 118 Then
            name = "F7"
        ElseIf KeyCode = 119 Then
            name = "F8"
        ElseIf KeyCode = 120 Then
            name = "F9"
        ElseIf KeyCode = 121 Then
            name = "F10"
        ElseIf KeyCode = 122 Then
            name = "F11"
        ElseIf KeyCode = 123 Then
            name = "F12"
        ElseIf KeyCode = 44 Then
            name = "Impr. Pant"
        ElseIf KeyCode = 106 Then
            name = "*"
        ElseIf KeyCode = vbKeyNumpad0 Then
            name = "Numerico 0"
        ElseIf KeyCode = vbKeyNumpad1 Then
            name = "Numerico 1"
        ElseIf KeyCode = vbKeyNumpad2 Then
            name = "Numerico 2"
        ElseIf KeyCode = vbKeyNumpad3 Then
            name = "Numerico 3"
        ElseIf KeyCode = vbKeyNumpad4 Then
            name = "Numerico 4"
        ElseIf KeyCode = vbKeyNumpad5 Then
            name = "Numerico 5"
        ElseIf KeyCode = vbKeyNumpad6 Then
            name = "Numerico 6"
        ElseIf KeyCode = vbKeyNumpad7 Then
            name = "Numerico 7"
        ElseIf KeyCode = vbKeyNumpad8 Then
            name = "Numerico 8"
        ElseIf KeyCode = vbKeyNumpad9 Then
            name = "Numerico 9"
        ElseIf KeyCode = vbKeyAdd Then
            name = "Numerico +"
        ElseIf KeyCode = 110 Then
            name = "Numerico ."
        ElseIf KeyCode = 226 Then
            name = "<"
        ElseIf KeyCode = 189 Then
            name = "-"
        ElseIf KeyCode = 188 Then
            name = ","
        ElseIf KeyCode = 190 Then
            name = "."
        Else
    
            name = Chr(KeyCode)

        End If
    
        Call Change_TempKey(Index, KeyCode, name)

    End If

End Sub

Sub Change_TempKey(Index As Integer, KeyCode As Integer, name As String)
    TempVars(Index) = KeyCode
    txConfig(Index).Text = name

End Sub

Function AlreadyBinded(KeyCode As Integer) As Boolean

    Dim i As Integer

    'If (KeyCode >= vbKeyF1 And KeyCode <= vbKeyF12) Or (KeyCode = 44) Or (KeyCode = 106) Then
    'If (KeyCode = 44) Or (KeyCode = 106) Then
    '   AlreadyBinded = True
    '   Exit Function
    'End If

    For i = 1 To NUMBINDS

        If (TempVars(i - 1) = KeyCode) Then
            AlreadyBinded = True
            Exit Function

        End If

    Next i

End Function

