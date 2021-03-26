VERSION 5.00
Begin VB.Form FrmGrupo 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Grupo"
   ClientHeight    =   6525
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   6510
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
   ScaleHeight     =   435
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   434
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstGrupo 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Height          =   2340
      Left            =   3600
      TabIndex        =   0
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Image Aceptar 
      Height          =   420
      Left            =   3525
      Tag             =   "0"
      Top             =   5730
      Width           =   1980
   End
   Begin VB.Image cmdAbandonar 
      Height          =   420
      Left            =   1005
      Tag             =   "0"
      Top             =   5730
      Width           =   1980
   End
   Begin VB.Image Command2 
      Height          =   420
      Left            =   6030
      Tag             =   "0"
      Top             =   15
      Width           =   465
   End
   Begin VB.Image cmdExpulsar 
      Height          =   420
      Left            =   4185
      Tag             =   "0"
      Top             =   4665
      Width           =   465
   End
   Begin VB.Image cmdInvitar 
      Height          =   420
      Left            =   4815
      Tag             =   "0"
      Top             =   4665
      Width           =   465
   End
End
Attribute VB_Name = "FrmGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Aceptar_Click()
    Debug.Print "Era lo mismo que cerrar..."
    Unload Me
End Sub

Private Sub Aceptar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Aceptar.Tag = "0" Then
        Aceptar.Picture = LoadInterface("boton-aceptar-ES-over.bmp")
        Aceptar.Tag = "1"
    End If
End Sub

Private Sub Aceptar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Aceptar.Picture = Nothing
    Aceptar.Tag = "1"
End Sub

Private Sub cmdAbandonar_Click()
    
    On Error GoTo cmdAbandonar_Click_Err
    
    Call WriteAbandonarGrupo
    Unload Me

    
    Exit Sub

cmdAbandonar_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmGrupo.cmdAbandonar_Click", Erl)
    Resume Next
    
End Sub

Private Sub cmdAbandonar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdAbandonar.Picture = LoadInterface("boton-abandonar-es-off.bmp")
    cmdAbandonar.Tag = "1"
End Sub

Private Sub cmdAbandonar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo cmdAbandonar_MouseMove_Err
    

    If cmdAbandonar.Tag = "0" Then
        cmdAbandonar.Picture = LoadInterface("boton-abandonar-es-over.bmp")
        cmdAbandonar.Tag = "1"
    End If
    
    If Aceptar.Tag = "1" Then
        Aceptar.Picture = Nothing
        Aceptar.Tag = "0"
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

Private Sub cmdExpulsar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdExpulsar.Picture = LoadInterface("boton-menos-off.bmp")
    cmdExpulsar.Tag = "1"
End Sub

Private Sub cmdExpulsar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo cmdExpulsar_MouseMove_Err
    
    
    If cmdExpulsar.Tag = "0" Then
        cmdExpulsar.Picture = LoadInterface("boton-menos-over.bmp")
        cmdExpulsar.Tag = "1"

    End If

    cmdInvitar.Picture = Nothing
    cmdInvitar.Tag = "0"

    cmdAbandonar.Picture = Nothing
    cmdAbandonar.Tag = "0"
    
    Aceptar.Picture = Nothing
    Aceptar.Tag = "0"

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

Private Sub cmdInvitar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cmdInvitar.Picture = LoadInterface("boton-mas-off.bmp")
    cmdInvitar.Tag = "1"
End Sub

Private Sub cmdInvitar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo cmdInvitar_MouseMove_Err
    

    If cmdInvitar.Tag = "0" Then
        cmdInvitar.Picture = LoadInterface("boton-mas-over.bmp")
        cmdInvitar.Tag = "1"
    End If
    
    cmdExpulsar.Picture = Nothing
    cmdExpulsar.Tag = "0"

    cmdAbandonar.Picture = Nothing
    cmdAbandonar.Tag = "0"

    Command2.Picture = Nothing
    Command2.Tag = "0"
    
    Aceptar.Picture = Nothing
    Aceptar.Tag = "0"

    
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

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Command2.Picture = LoadInterface("boton-cerrar-off.bmp")
    Command2.Tag = "1"
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo Command2_MouseMove_Err
    

    If Command2.Tag = "0" Then
        Command2.Picture = LoadInterface("boton-cerrar-over.bmp")
        Command2.Tag = "1"
    End If
    
    cmdExpulsar.Picture = Nothing
    cmdExpulsar.Tag = "0"

    cmdInvitar.Picture = Nothing
    cmdInvitar.Tag = "0"

    cmdAbandonar.Picture = Nothing
    cmdAbandonar.Tag = "0"
    
    Aceptar.Picture = Nothing
    Aceptar.Tag = "0"

    
    Exit Sub

Command2_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmGrupo.Command2_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call FormParser.Parse_Form(Me)
    
    Call Aplicar_Transparencia(Me.hWnd, 220)
    
    Me.Picture = LoadInterface("ventanagrupo.bmp")
    
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

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    Call moverForm(Me.hWnd)
    
    cmdExpulsar.Picture = Nothing
    cmdExpulsar.Tag = "0"

    cmdInvitar.Picture = Nothing
    cmdInvitar.Tag = "0"

    cmdAbandonar.Picture = Nothing
    cmdAbandonar.Tag = "0"

    Command2.Picture = Nothing
    Command2.Tag = "0"
    
    Aceptar.Picture = Nothing
    Aceptar.Tag = "0"

    
    Exit Sub

Form_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmGrupo.Form_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub lstGrupo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo lstGrupo_MouseMove_Err
    
    cmdExpulsar.Picture = Nothing
    cmdExpulsar.Tag = "0"

    cmdInvitar.Picture = Nothing
    cmdInvitar.Tag = "0"

    cmdAbandonar.Picture = Nothing
    cmdAbandonar.Tag = "0"

    Command2.Picture = Nothing
    Command2.Tag = "0"
    
    Aceptar.Picture = Nothing
    Aceptar.Tag = "0"

    
    Exit Sub

lstGrupo_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmGrupo.lstGrupo_MouseMove", Erl)
    Resume Next
    
End Sub
