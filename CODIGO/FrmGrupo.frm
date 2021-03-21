VERSION 5.00
Begin VB.Form FrmGrupo 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Grupo"
   ClientHeight    =   5640
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   6555
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
   ScaleHeight     =   5640
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstGrupo 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1200
      Left            =   3840
      TabIndex        =   0
      Top             =   1680
      Width           =   2190
   End
   Begin VB.Image Command2 
      Height          =   465
      Left            =   3990
      Tag             =   "0"
      Top             =   4800
      Width           =   1905
   End
   Begin VB.Image cmdAbandonar 
      Height          =   480
      Left            =   3930
      Tag             =   "0"
      Top             =   4080
      Width           =   2070
   End
   Begin VB.Image cmdExpulsar 
      Height          =   555
      Left            =   4260
      Tag             =   "0"
      Top             =   3270
      Width           =   585
   End
   Begin VB.Image cmdInvitar 
      Height          =   555
      Left            =   5030
      Tag             =   "0"
      Top             =   3270
      Width           =   585
   End
End
Attribute VB_Name = "FrmGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAbandonar_Click()
    
    On Error GoTo cmdAbandonar_Click_Err
    
    Call WriteAbandonarGrupo
    Unload Me

    
    Exit Sub

cmdAbandonar_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmGrupo.cmdAbandonar_Click", Erl)
    Resume Next
    
End Sub

Private Sub cmdAbandonar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' cmdAbandonar.Picture = LoadInterface(Language + "\grupo_abandonarpress.bmp")
    ' cmdAbandonar.Tag = "1"
End Sub

Private Sub cmdAbandonar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo cmdAbandonar_MouseMove_Err
    

    If cmdAbandonar.Tag = "0" Then
        cmdAbandonar.Picture = LoadInterface(Language + "\grupo_abandonarhover.bmp")
        cmdAbandonar.Tag = "1"

    End If
    
    cmdExpulsar.Picture = Nothing
    cmdExpulsar.Tag = "0"

    cmdInvitar.Picture = Nothing
    cmdInvitar.Tag = "0"

    Command2.Picture = Nothing
    Command2.Tag = "0"

    
    Exit Sub

cmdAbandonar_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmGrupo.cmdAbandonar_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub cmdExpulsar_Click()
    
    On Error GoTo cmdExpulsar_Click_Err
    

    If lstGrupo.ListIndex >= 0 Then
        Call WriteHecharDeGrupo(lstGrupo.ListIndex)
        Unload Me

    End If

    
    Exit Sub

cmdExpulsar_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmGrupo.cmdExpulsar_Click", Erl)
    Resume Next
    
End Sub

Private Sub cmdExpulsar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' cmdExpulsar.Picture = LoadInterface(Language + "\grupo_expulsarpress.bmp")
    '  cmdExpulsar.Tag = "1"
End Sub

Private Sub cmdExpulsar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo cmdExpulsar_MouseMove_Err
    
    
    If cmdExpulsar.Tag = "0" Then
        cmdExpulsar.Picture = LoadInterface(Language + "\grupo_expulsarhover.bmp")
        cmdExpulsar.Tag = "1"

    End If

    cmdInvitar.Picture = Nothing
    cmdInvitar.Tag = "0"

    cmdAbandonar.Picture = Nothing
    cmdAbandonar.Tag = "0"

    Command2.Picture = Nothing
    Command2.Tag = "0"

    
    Exit Sub

cmdExpulsar_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmGrupo.cmdExpulsar_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub cmdInvitar_Click()
    
    On Error GoTo cmdInvitar_Click_Err
    
    Unload Me
    Call WriteInvitarGrupo

    
    Exit Sub

cmdInvitar_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmGrupo.cmdInvitar_Click", Erl)
    Resume Next
    
End Sub

Private Sub cmdInvitar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    'cmdInvitar.Picture = LoadInterface(Language + "\grupo_invitarpress.bmp")
    'cmdInvitar.Tag = "1"
End Sub

Private Sub cmdInvitar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo cmdInvitar_MouseMove_Err
    

    If cmdInvitar.Tag = "0" Then
        cmdInvitar.Picture = LoadInterface(Language + "\grupo_invitarhover.bmp")
        cmdInvitar.Tag = "1"

    End If
    
    cmdExpulsar.Picture = Nothing
    cmdExpulsar.Tag = "0"

    cmdAbandonar.Picture = Nothing
    cmdAbandonar.Tag = "0"

    Command2.Picture = Nothing
    Command2.Tag = "0"

    
    Exit Sub

cmdInvitar_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmGrupo.cmdInvitar_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Command2_Click()
    
    On Error GoTo Command2_Click_Err
    
    Unload Me

    
    Exit Sub

Command2_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmGrupo.Command2_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    '  Command2.Picture = LoadInterface(Language + "\grupo_salirpress.bmp")
    ' Command2.Tag = "1"
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Command2_MouseMove_Err
    

    If Command2.Tag = "0" Then
        Command2.Picture = LoadInterface(Language + "\grupo_salirhover.bmp")
        Command2.Tag = "1"

    End If
    
    cmdExpulsar.Picture = Nothing
    cmdExpulsar.Tag = "0"

    cmdInvitar.Picture = Nothing
    cmdInvitar.Tag = "0"

    cmdAbandonar.Picture = Nothing
    cmdAbandonar.Tag = "0"

    
    Exit Sub

Command2_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmGrupo.Command2_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call FormParser.Parse_Form(Me)

    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmGrupo.Form_Load", Erl)
    Resume Next
    
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Form_KeyPress_Err
    

    If (KeyAscii = 27) Then
        Unload Me

    End If

    
    Exit Sub

Form_KeyPress_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmGrupo.Form_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    cmdExpulsar.Picture = Nothing
    cmdExpulsar.Tag = "0"

    cmdInvitar.Picture = Nothing
    cmdInvitar.Tag = "0"

    cmdAbandonar.Picture = Nothing
    cmdAbandonar.Tag = "0"

    Command2.Picture = Nothing
    Command2.Tag = "0"

    
    Exit Sub

Form_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmGrupo.Form_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub lstGrupo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo lstGrupo_MouseMove_Err
    
    cmdExpulsar.Picture = Nothing
    cmdExpulsar.Tag = "0"

    cmdInvitar.Picture = Nothing
    cmdInvitar.Tag = "0"

    cmdAbandonar.Picture = Nothing
    cmdAbandonar.Tag = "0"

    Command2.Picture = Nothing
    Command2.Tag = "0"

    
    Exit Sub

lstGrupo_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmGrupo.lstGrupo_MouseMove", Erl)
    Resume Next
    
End Sub
