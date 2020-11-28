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
    Call WriteAbandonarGrupo
    Unload Me

End Sub

Private Sub cmdAbandonar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' cmdAbandonar.Picture = LoadInterface("grupo_abandonarpress.bmp")
    ' cmdAbandonar.Tag = "1"
End Sub

Private Sub cmdAbandonar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If cmdAbandonar.Tag = "0" Then
        cmdAbandonar.Picture = LoadInterface("grupo_abandonarhover.bmp")
        cmdAbandonar.Tag = "1"

    End If
    
    cmdExpulsar.Picture = Nothing
    cmdExpulsar.Tag = "0"

    cmdInvitar.Picture = Nothing
    cmdInvitar.Tag = "0"

    Command2.Picture = Nothing
    Command2.Tag = "0"

End Sub

Private Sub cmdExpulsar_Click()

    If lstGrupo.ListIndex >= 0 Then
        Call WriteHecharDeGrupo(lstGrupo.ListIndex)
        Unload Me

    End If

End Sub

Private Sub cmdExpulsar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' cmdExpulsar.Picture = LoadInterface("grupo_expulsarpress.bmp")
    '  cmdExpulsar.Tag = "1"
End Sub

Private Sub cmdExpulsar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If cmdExpulsar.Tag = "0" Then
        cmdExpulsar.Picture = LoadInterface("grupo_expulsarhover.bmp")
        cmdExpulsar.Tag = "1"

    End If

    cmdInvitar.Picture = Nothing
    cmdInvitar.Tag = "0"

    cmdAbandonar.Picture = Nothing
    cmdAbandonar.Tag = "0"

    Command2.Picture = Nothing
    Command2.Tag = "0"

End Sub

Private Sub cmdInvitar_Click()
    Unload Me
    Call WriteInvitarGrupo

End Sub

Private Sub cmdInvitar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    'cmdInvitar.Picture = LoadInterface("grupo_invitarpress.bmp")
    'cmdInvitar.Tag = "1"
End Sub

Private Sub cmdInvitar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If cmdInvitar.Tag = "0" Then
        cmdInvitar.Picture = LoadInterface("grupo_invitarhover.bmp")
        cmdInvitar.Tag = "1"

    End If
    
    cmdExpulsar.Picture = Nothing
    cmdExpulsar.Tag = "0"

    cmdAbandonar.Picture = Nothing
    cmdAbandonar.Tag = "0"

    Command2.Picture = Nothing
    Command2.Tag = "0"

End Sub

Private Sub Command2_Click()
    Unload Me

End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    '  Command2.Picture = LoadInterface("grupo_salirpress.bmp")
    ' Command2.Tag = "1"
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Command2.Tag = "0" Then
        Command2.Picture = LoadInterface("grupo_salirhover.bmp")
        Command2.Tag = "1"

    End If
    
    cmdExpulsar.Picture = Nothing
    cmdExpulsar.Tag = "0"

    cmdInvitar.Picture = Nothing
    cmdInvitar.Tag = "0"

    cmdAbandonar.Picture = Nothing
    cmdAbandonar.Tag = "0"

End Sub

Private Sub Form_Load()
    Call FormParser.Parse_Form(Me)

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)

    If (KeyAscii = 27) Then
        Unload Me

    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdExpulsar.Picture = Nothing
    cmdExpulsar.Tag = "0"

    cmdInvitar.Picture = Nothing
    cmdInvitar.Tag = "0"

    cmdAbandonar.Picture = Nothing
    cmdAbandonar.Tag = "0"

    Command2.Picture = Nothing
    Command2.Tag = "0"

End Sub

Private Sub lstGrupo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    cmdExpulsar.Picture = Nothing
    cmdExpulsar.Tag = "0"

    cmdInvitar.Picture = Nothing
    cmdInvitar.Tag = "0"

    cmdAbandonar.Picture = Nothing
    cmdAbandonar.Tag = "0"

    Command2.Picture = Nothing
    Command2.Tag = "0"

End Sub
