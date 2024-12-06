VERSION 5.00
Begin VB.Form frmCreateBattleground 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crear Battleground"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   390
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
   Icon            =   "frmCreateBattleground.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   383
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   405
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCrear 
      Caption         =   "Crear"
      Height          =   495
      Left            =   4680
      TabIndex        =   26
      Top             =   5160
      Width           =   1215
   End
   Begin VB.TextBox tSize 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   22
      Text            =   "1"
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox tMaxPlayers 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3480
      MaxLength       =   2
      TabIndex        =   20
      Text            =   "20"
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox tMinPlayers 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   18
      Text            =   "2"
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox tMaxLvl 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3480
      MaxLength       =   2
      TabIndex        =   16
      Text            =   "47"
      Top             =   3480
      Width           =   495
   End
   Begin VB.TextBox tMinLvl 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   15
      Text            =   "1"
      Top             =   3480
      Width           =   495
   End
   Begin VB.CommandButton cmb1kl 
      Caption         =   "- 1k"
      Height          =   255
      Left            =   3240
      TabIndex        =   12
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton cmb10kl 
      Caption         =   "- 10k"
      Height          =   255
      Left            =   2640
      TabIndex        =   13
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton cmb10kp 
      Caption         =   "+ 10k"
      Height          =   255
      Left            =   5280
      TabIndex        =   11
      Top             =   3000
      Width           =   615
   End
   Begin VB.CommandButton cmb1kp 
      Caption         =   "+ 1k"
      Height          =   255
      Left            =   4680
      TabIndex        =   10
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox tCosto 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3840
      MaxLength       =   8
      TabIndex        =   9
      Text            =   "0"
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox tPassword 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2640
      MaxLength       =   50
      TabIndex        =   6
      Top             =   2520
      Width           =   3255
   End
   Begin VB.CheckBox chkPassword 
      Caption         =   "Contraseña:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   2520
      Width           =   1455
   End
   Begin VB.ComboBox cmbEquipos 
      Height          =   315
      ItemData        =   "frmCreateBattleground.frx":000C
      Left            =   2640
      List            =   "frmCreateBattleground.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2040
      Width           =   3255
   End
   Begin VB.ComboBox cmbTipo 
      Height          =   315
      ItemData        =   "frmCreateBattleground.frx":002D
      Left            =   2640
      List            =   "frmCreateBattleground.frx":0037
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1560
      Width           =   3255
   End
   Begin VB.TextBox tName 
      Height          =   285
      Left            =   2640
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1080
      Width           =   3255
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "Crear Battleground"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   25
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label lblDivisible 
      Caption         =   "El límite de jugadores debe ser divisible por el tamaño"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3240
      TabIndex        =   24
      Top             =   4440
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Tamaño de equipos:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label Label8 
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   21
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Límite de jugadores:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label Label6 
      Caption         =   "al"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   17
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Límite de nivel:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Costo de inscripción:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Formato de equipos:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Tipo de partida:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Nombre de la partida:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   2415
   End
End
Attribute VB_Name = "frmCreateBattleground"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCrear_Click()
On Error GoTo ErrHandler:

    Dim Settings As t_NewScenearioSettings

    If Len(tName.Text) < 3 Then
        Call MsgBox(JsonLanguage.Item("MENSAJE_NOMBRE_PARTIDA_CORTO"), vbExclamation)
        tName.SetFocus
        Exit Sub
    End If

    Settings.InscriptionFee = Val(tCosto.Text)
    If Settings.InscriptionFee < 0 Or Settings.InscriptionFee > 10000000 Then
        Call MsgBox(JsonLanguage.Item("MENSAJE_COSTO_PARTIDA_INVALIDO"), vbExclamation)
        tCosto.SetFocus
        Exit Sub
    End If

    Settings.MinLevel = Val(tMinLvl.Text)
    Settings.MaxLevel = Val(tMaxLvl.Text)
    If Settings.MinLevel > Settings.MaxLevel Or Settings.MinLevel > 47 Or Settings.MinLevel < 1 Or Settings.MaxLevel > 47 Or Settings.MaxLevel < 1 Then
        Call MsgBox(JsonLanguage.Item("MENSAJE_LIMITES_NIVELES_INVALIDOS"), vbExclamation)
        tMinLvl.SetFocus
        Exit Sub
    End If
    
    Settings.MinPlayers = Val(tMinPlayers.Text)
    Settings.MaxPlayers = Val(tMaxPlayers.Text)
    If Settings.MinPlayers > Settings.MaxPlayers Or Settings.MinPlayers > 40 Or Settings.MinPlayers < 2 Or Settings.MaxPlayers > 40 Or Settings.MaxPlayers < 2 Then
        Call MsgBox(JsonLanguage.Item("MENSAJE_LIMITES_JUGADORES_INVALIDOS"), vbExclamation)

        tMinPlayers.SetFocus
        Exit Sub
    End If
    Settings.TeamSize = Val(tSize.Text)
    If Settings.MinPlayers Mod Settings.TeamSize <> 0 Or Settings.MaxPlayers Mod Settings.TeamSize <> 0 Then
        Call MsgBox(JsonLanguage.Item("MENSAJE_LIMITE_JUGADORES_DIVISIBLE"), vbExclamation)

        tSize.SetFocus
        Exit Sub
    End If
    
    Select Case cmbTipo.ListIndex
        Case 0
            Settings.ScenearioType = 3
        Case 1
            Settings.ScenearioType = 4
    End Select
    
    Select Case cmbEquipos.ListIndex
        Case 0
            Settings.TeamType = 1
        Case 1
            Settings.ScenearioType = 2
    End Select
    
    Call WriteStartLobby(1, Settings, tName.Text, tPassword.Text)
    
    Unload Me
    
    Exit Sub
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "frmCreateBattleGround.btnCrear", Erl)
    Resume Next
End Sub

Private Sub chkPassword_Click()
    tPassword.enabled = chkPassword.Value
    If Not tPassword.enabled Then
        tPassword.Text = ""
    Else
        tPassword.SetFocus
    End If
End Sub


Private Sub cmb10kl_Click()
    Dim Value As Long
    Value = Val(tCosto.Text) - 10000
    If Value < 0 Then Value = 0
    tCosto.Text = Value
End Sub

Private Sub cmb10kp_Click()
    Dim Value As Integer
    Value = Val(tCosto.Text) + 10000
    If Value > 10000000 Then Value = 10000000
    tCosto.Text = Value
End Sub

Private Sub cmb1kl_Click()
    Dim Value As Long
    Value = Val(tCosto.Text) - 1000
    If Value < 0 Then Value = 0
    tCosto.Text = Value
End Sub

Private Sub cmb1kp_Click()
    Dim Value As Long
    Value = Val(tCosto.Text) + 1000
    If Value > 10000000 Then Value = 10000000
    tCosto.Text = Value
End Sub

Private Sub Form_Load()
    cmbTipo.ListIndex = 0
    cmbEquipos.ListIndex = 0
End Sub

Private Sub tMaxLvl_LostFocus()
    Dim Value As Long
    Value = Val(tMaxLvl.Text)
    If Value < 1 Then Value = 1
    If Value > 47 Then Value = 47
    tMaxLvl.Text = Value
End Sub

Private Sub tMaxPlayers_LostFocus()
    Dim Value As Integer
    Value = Val(tMaxPlayers.Text)
    If Value < 2 Then Value = 2
    If Value > 40 Then Value = 40
    tMaxPlayers.Text = Value
End Sub

Private Sub tMinLvl_LostFocus()
    Dim Value As Integer
    Value = Val(tMinLvl.Text)
    If Value < 1 Then Value = 1
    If Value > 47 Then Value = 47
    tMinLvl.Text = Value
End Sub

Private Sub tMinPlayers_LostFocus()
    Dim Value As Integer
    Value = Val(tMinPlayers.Text)
    If Value < 2 Then Value = 2
    If Value > 40 Then Value = 40
    tMinPlayers.Text = Value
End Sub

Private Sub tSize_LostFocus()
    Dim Value As Integer
    Value = Val(tSize.Text)
    If Value < 1 Then Value = 1
    If Value > 20 Then Value = 20
    tSize.Text = Value
    lblDivisible.visible = Value Mod Val(tMaxPlayers.Text) <> 0 Or Value Mod Val(tMaxPlayers.Text) <> 0
End Sub
