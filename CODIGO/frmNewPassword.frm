VERSION 5.00
Begin VB.Form frmNewPassword 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Cambiar Contraseña de cuenta"
   ClientHeight    =   4425
   ClientLeft      =   0
   ClientTop       =   -180
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   295
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   291
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtNueva 
      Appearance      =   0  'Flat
      BackColor       =   &H0012130D&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   240
      IMEMode         =   3  'DISABLE
      Left            =   990
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2460
      Width           =   2415
   End
   Begin VB.TextBox txtNueva2 
      Appearance      =   0  'Flat
      BackColor       =   &H0012130D&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   240
      IMEMode         =   3  'DISABLE
      Left            =   990
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   3255
      Width           =   2415
   End
   Begin VB.TextBox txtAnterior 
      Appearance      =   0  'Flat
      BackColor       =   &H0012130D&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      IMEMode         =   3  'DISABLE
      Left            =   990
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   1650
      Width           =   2415
   End
   Begin VB.Image btnCerrar 
      Height          =   420
      Left            =   3900
      Tag             =   "0"
      Top             =   15
      Width           =   465
   End
   Begin VB.Image Aceptar 
      Height          =   420
      Left            =   1200
      Tag             =   "0"
      Top             =   3735
      Width           =   1980
   End
End
Attribute VB_Name = "frmNewPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Aceptar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Goto Aceptar_MouseDown_Err
    Aceptar.Picture = LoadInterface("boton-aceptar-ES-off.bmp")
    Aceptar.Tag = "1"
    Exit Sub
Aceptar_MouseDown_Err:
    Call TraceError(Err.Number, Err.Description, "frmNewPassword.Aceptar_MouseDown", Erl)
End Sub

Private Sub btnCerrar_Click()
    On Error Goto btnCerrar_Click_Err
    Unload Me
    Exit Sub
btnCerrar_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmNewPassword.btnCerrar_Click", Erl)
End Sub

Private Sub btnCerrar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Goto btnCerrar_MouseDown_Err
    btnCerrar.Picture = LoadInterface("boton-cerrar-off.bmp")
    btnCerrar.Tag = "1"
    Exit Sub
btnCerrar_MouseDown_Err:
    Call TraceError(Err.Number, Err.Description, "frmNewPassword.btnCerrar_MouseDown", Erl)
End Sub

Private Sub btnCerrar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Goto btnCerrar_MouseMove_Err
    If btnCerrar.Tag = "0" Then
        btnCerrar.Picture = LoadInterface("boton-cerrar-over.bmp")
        btnCerrar.Tag = "1"
    End If
    Exit Sub
btnCerrar_MouseMove_Err:
    Call TraceError(Err.Number, Err.Description, "frmNewPassword.btnCerrar_MouseMove", Erl)
End Sub

Private Sub Form_Load()
    On Error Goto Form_Load_Err
    
    On Error GoTo Form_Load_Err
    
    Call FormParser.Parse_Form(Me)
    
    Call Aplicar_Transparencia(Me.hWnd, 240)
    
    Me.Picture = LoadInterface("ventanacambiarcontrasena.bmp")
    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "frmNewPassword.Form_Load", Erl)
    Resume Next
    
    Exit Sub
Form_Load_Err:
    Call TraceError(Err.Number, Err.Description, "frmNewPassword.Form_Load", Erl)
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error Goto Form_KeyPress_Err
    
    On Error GoTo Form_KeyPress_Err
    

    If (KeyAscii = 27) Then
        Unload Me
    End If

    
    Exit Sub

Form_KeyPress_Err:
    Call RegistrarError(Err.number, Err.Description, "frmNewPassword.Form_KeyPress", Erl)
    Resume Next
    
    Exit Sub
Form_KeyPress_Err:
    Call TraceError(Err.Number, Err.Description, "frmNewPassword.Form_KeyPress", Erl)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Goto Form_MouseMove_Err
    
    On Error GoTo Form_MouseMove_Err
    
    Call MoverForm(Me.hWnd)
    
    If Aceptar.Tag = "1" Then
        Aceptar.Picture = Nothing
        Aceptar.Tag = "0"
    End If
    
    If btnCerrar.Tag = "1" Then
        btnCerrar.Picture = Nothing
        btnCerrar.Tag = "0"
    End If
    
    Exit Sub

Form_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmNewPassword.Form_MouseMove", Erl)
    Resume Next
    
    Exit Sub
Form_MouseMove_Err:
    Call TraceError(Err.Number, Err.Description, "frmNewPassword.Form_MouseMove", Erl)
End Sub

Private Sub Aceptar_Click()
    On Error Goto Aceptar_Click_Err
    
    On Error GoTo Image1_Click_Err

    If txtNueva.Text = "" Then
        Unload Me
    End If

    If txtNueva.Text <> txtNueva2.Text Then
        Call MsgBox("Las contraseñas no coinciden", vbCritical Or vbOKOnly Or vbApplicationModal Or vbDefaultButton1, "Cambiar Contraseña")
        Exit Sub
    End If

    Call WriteChangePassword(txtAnterior.Text, txtNueva.Text)
    Unload Me

    
    Exit Sub

Image1_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmNewPassword.Image1_Click", Erl)
    Resume Next
    
    Exit Sub
Aceptar_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmNewPassword.Aceptar_Click", Erl)
End Sub

Private Sub Aceptar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Goto Aceptar_MouseMove_Err
    
    On Error GoTo Image1_MouseMove_Err
    

    If Aceptar.Tag = "0" Then
        Aceptar.Picture = LoadInterface("boton-aceptar-ES-over.bmp")
        Aceptar.Tag = "1"
    End If

    
    Exit Sub

Image1_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmNewPassword.Image1_MouseMove", Erl)
    Resume Next
    
    Exit Sub
Aceptar_MouseMove_Err:
    Call TraceError(Err.Number, Err.Description, "frmNewPassword.Aceptar_MouseMove", Erl)
End Sub
