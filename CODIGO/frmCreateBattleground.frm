VERSION 5.00
Begin VB.Form frmCreateBattleground 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crear Battleground"
   ClientHeight    =   6255
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
      TabIndex        =   21
      Top             =   5670
      Width           =   1215
   End
   Begin VB.TextBox tSize 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   18
      Text            =   "1"
      Top             =   4440
      Width           =   495
   End
   Begin VB.TextBox tMaxPlayers 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3480
      MaxLength       =   2
      TabIndex        =   16
      Text            =   "20"
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox tMinPlayers 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   14
      Text            =   "2"
      Top             =   3960
      Width           =   495
   End
   Begin VB.TextBox tMaxLvl 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3480
      MaxLength       =   2
      TabIndex        =   12
      Text            =   "47"
      Top             =   3480
      Width           =   495
   End
   Begin VB.TextBox tMinLvl 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   11
      Text            =   "1"
      Top             =   3480
      Width           =   495
   End
   Begin VB.TextBox tCosto 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2640
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
      List            =   "frmCreateBattleground.frx":002F
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
   Begin VB.TextBox tRoundAmount 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2640
      MaxLength       =   3
      TabIndex        =   23
      Text            =   "1"
      Top             =   4920
      Width           =   495
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
      TabIndex        =   22
      Top             =   4440
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblTitle 
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
      TabIndex        =   20
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label lblTeamSize 
      Alignment       =   1  'Right Justify
      Caption         =   "Tamaño de equipos"
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
      Top             =   4440
      Width           =   2415
   End
   Begin VB.Label lblSeparator2 
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
      TabIndex        =   17
      Top             =   3960
      Width           =   255
   End
   Begin VB.Label lblPlayerLimit 
      Alignment       =   1  'Right Justify
      Caption         =   "Límite de jugadores"
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
      TabIndex        =   15
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Label lblSeparator1 
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
      TabIndex        =   13
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label lblLvlRangeLimit 
      Alignment       =   1  'Right Justify
      Caption         =   "Límite de nivel"
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
      TabIndex        =   10
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label lblInscriptionFee 
      Alignment       =   1  'Right Justify
      Caption         =   "Costo de inscripción"
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
   Begin VB.Label lblTeamFormat 
      Alignment       =   1  'Right Justify
      Caption         =   "Formato de equipos"
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
   Begin VB.Label lblEventType 
      Alignment       =   1  'Right Justify
      Caption         =   "Tipo de partida"
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
   Begin VB.Label lblMatchName 
      Alignment       =   1  'Right Justify
      Caption         =   "Nombre de la partida"
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
   Begin VB.Label lblRoundAmount 
      Alignment       =   1  'Right Justify
      Caption         =   "Cantidad de rondas"
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
      TabIndex        =   24
      Top             =   4920
      Width           =   2415
   End
End
Attribute VB_Name = "frmCreateBattleground"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    On Error GoTo ErrHandler
    cmbTipo.Clear
    cmbTipo.AddItem JsonLanguage.Item("MENSAJE_EVENTO_CAPTURA")
    cmbTipo.AddItem JsonLanguage.Item("MENSAJE_EVENTO_CACERIA")
    cmbTipo.AddItem JsonLanguage.Item("MENSAJE_EVENTO_DEATHMATCH")
    cmbTipo.AddItem JsonLanguage.Item("MENSAJE_EVENTO_ABORDAJE")
    cmbEquipos.Clear
    cmbEquipos.AddItem JsonLanguage.Item("MENSAJE_EVENTO_MODALIDAD_RANDOM")
    cmbEquipos.AddItem JsonLanguage.Item("MENSAJE_EVENTO_MODALIDAD_GRUPOS")
    cmbTipo.ListIndex = 0
    cmbEquipos.ListIndex = 0
    tMinLvl.text = 1
    tMaxLvl.text = 47
    tMaxPlayers.text = 32
    tMinPlayers.text = 2
    tSize.text = 1
    tCosto.text = 0
    tRoundAmount.Text = 1
    Exit Sub
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "frmCreateBattleGround.Form_Load", Erl)
    Resume Next
End Sub

Private Sub btnCrear_Click()
    On Error GoTo ErrHandler
    Dim Settings As t_NewScenearioSettings
    
    If Len(tName.text) < 3 Then
        Call MsgBox(JsonLanguage.Item("MENSAJE_NOMBRE_PARTIDA_CORTO"), vbExclamation)
        tName.SetFocus
        Exit Sub
    End If
    
    Settings.InscriptionFee = val(tCosto.text)
    If Settings.InscriptionFee < 0 Or Settings.InscriptionFee > 10000000 Then
        Call MsgBox(JsonLanguage.Item("MENSAJE_COSTO_PARTIDA_INVALIDO"), vbExclamation)
        tCosto.SetFocus
        Exit Sub
    End If
    
    Settings.MinLevel = val(tMinLvl.text)
    Settings.MaxLevel = val(tMaxLvl.text)
    If Settings.MinLevel > Settings.MaxLevel Or Settings.MinLevel > 47 Or Settings.MinLevel < 1 Or Settings.MaxLevel > 47 Or Settings.MaxLevel < 1 Then
        Call MsgBox(JsonLanguage.Item("MENSAJE_LIMITES_NIVELES_INVALIDOS"), vbExclamation)
        tMinLvl.SetFocus
        Exit Sub
    End If
    
    Settings.MinPlayers = val(tMinPlayers.text)
    Settings.MaxPlayers = val(tMaxPlayers.text)
    If Settings.MinPlayers > Settings.MaxPlayers Or Settings.MinPlayers > 32 Or Settings.MinPlayers < 2 Or Settings.MaxPlayers > 32 Or Settings.MaxPlayers < 2 Then
        Call MsgBox(JsonLanguage.Item("MENSAJE_LIMITES_JUGADORES_INVALIDOS"), vbExclamation)
        tMinPlayers.SetFocus
        Exit Sub
    End If
    
    Settings.TeamSize = val(tSize.text)
    If Settings.TeamSize < 1 Or Settings.TeamSize > Settings.MaxPlayers Then
        Call MsgBox(JsonLanguage.Item("MENSAJE_LIMITE_JUGADORES_DIVISIBLE"), vbExclamation)
        tSize.SetFocus
        Exit Sub
    End If
    If Settings.MinPlayers Mod Settings.TeamSize <> 0 Or Settings.MaxPlayers Mod Settings.TeamSize <> 0 Then
        Call MsgBox(JsonLanguage.Item("MENSAJE_LIMITE_JUGADORES_DIVISIBLE"), vbExclamation)
        tSize.SetFocus
        Exit Sub
    End If
    
    Settings.RoundAmount = val(tRoundAmount.Text)
    If Settings.RoundAmount < 1 Or Settings.RoundAmount > 255 Then
        Call MsgBox(JsonLanguage.Item("MENSAJE_RONDAS_INVALIDO"), vbExclamation)
        tRoundAmount.SetFocus
        Exit Sub
    End If

    Select Case cmbTipo.ListIndex
        Case 0
            Settings.ScenearioType = e_EventType.CaptureTheFlag
        Case 1
            Settings.ScenearioType = e_EventType.NpcHunt
        Case 2
            Settings.ScenearioType = e_EventType.DeathMatch
        Case 3
            Settings.ScenearioType = e_EventType.NavalBattle
    End Select
    
    Select Case cmbEquipos.ListIndex
        Case e_TeamTypes.ePremade
            Settings.TeamType = e_TeamTypes.ePremade
        Case e_TeamTypes.eRandom
            Settings.TeamType = e_TeamTypes.eRandom
    End Select
    
    Call WriteStartLobby(1, Settings, tName.text, tPassword.text)
    Unload Me
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "frmCreateBattleGround.btnCrear", Erl)
    Resume Next
End Sub

Private Sub chkPassword_Click()
    tPassword.enabled = chkPassword.value
    If Not tPassword.enabled Then
        tPassword.text = ""
    Else
        tPassword.SetFocus
    End If
End Sub

Private Sub tMaxLvl_Change()
    Dim value As Long
    If tMaxLvl.Text = "" Or Not IsNumeric(tMaxLvl.Text) Then tMaxLvl.Text = "1"
    value = CLng(tMaxLvl.text)
    If value > 47 Then tMaxLvl.Text = "47"
    If value < 1 Then tMaxLvl.Text = "1"
    If value < CLng(tMinLvl.Text) Then tMaxLvl.Text = tMinLvl.Text
    Call ActualizarDivisible
End Sub

Private Sub tMinLvl_Change()
    Dim value As Long
    If tMinLvl.Text = "" Or Not IsNumeric(tMinLvl.Text) Then tMinLvl.Text = "1"
    value = CLng(tMinLvl.text)
    If value > 47 Then tMinLvl.Text = "47"
    If value < 1 Then tMinLvl.Text = "1"
End Sub

Private Sub tMaxPlayers_Change()
    Dim value As Long
    If tMaxPlayers.Text = "" Or Not IsNumeric(tMaxPlayers.Text) Then tMaxPlayers.Text = "2"
    value = CLng(tMaxPlayers.text)
    If value > 32 Then tMaxPlayers.Text = "32"
    If value < 2 Then tMaxPlayers.Text = "2"
    If value < CLng(tMinPlayers.Text) Then tMaxPlayers.Text = tMinPlayers.Text
    Call ActualizarDivisible
End Sub

Private Sub tMinPlayers_Change()
    Dim value As Long
    If tMinPlayers.Text = "" Or Not IsNumeric(tMinPlayers.Text) Then tMinPlayers.Text = "2"
    value = CLng(tMinPlayers.text)
    If value > 32 Then tMinPlayers.Text = "32"
    If value < 2 Then tMinPlayers.Text = "2"
    Call ActualizarDivisible
End Sub

Private Sub tSize_Change()
    Dim value As Long
    If tSize.Text = "" Or Not IsNumeric(tSize.Text) Then tSize.Text = "1"
    value = CLng(tSize.text)
    If value < 1 Then tSize.Text = "1"
    If IsNumeric(tMaxPlayers.Text) Then
        If value > CLng(tMaxPlayers.Text) Then tSize.Text = tMaxPlayers.Text
    End If
    If cmbEquipos.ListIndex = 1 And value <= 1 Then tSize.Text = "2"
    Call ActualizarDivisible
End Sub

Private Sub tRoundAmount_Change()
    Dim value As Long
    If tRoundAmount.Text = "" Or Not IsNumeric(tRoundAmount.Text) Then tRoundAmount.Text = "1"
    value = CLng(tRoundAmount.Text)
    If value > 255 Then tRoundAmount.Text = "255"
    If value < 1 Then tRoundAmount.Text = "1"
End Sub

Private Sub ActualizarDivisible()
    If Not IsNumeric(tSize.Text) Or Not IsNumeric(tMaxPlayers.Text) Or Not IsNumeric(tMinPlayers.Text) Then Exit Sub
    Dim sz As Long, mn As Long, mX As Long
    sz = CLng(tSize.Text)
    mn = CLng(tMinPlayers.Text)
    mX = CLng(tMaxPlayers.Text)
    lblDivisible.visible = (sz > 0) And (mn Mod sz <> 0 Or mX Mod sz <> 0)
End Sub

Private Sub tMinPlayers_LostFocus()
    Call tMinPlayers_Change
End Sub

Private Sub tSize_LostFocus()
    Call tSize_Change
End Sub

Private Sub tMinLvl_LostFocus()
    Call tMinLvl_Change
End Sub

Private Sub tMaxPlayers_LostFocus()
    Call tMaxPlayers_Change
End Sub

Private Sub tMaxLvl_LostFocus()
    Call tMaxLvl_Change
End Sub

Private Sub tRoundAmount_LostFocus()
    Call tRoundAmount_Change
End Sub

Private Sub cmbEquipos_LostFocus()
    Dim value As Long
    value = CLng(tSize.text)
    If cmbEquipos.ListIndex = 1 And value <= 1 Then tSize.Text = "2"
    Call ActualizarDivisible
End Sub
