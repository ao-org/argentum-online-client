VERSION 5.00
Begin VB.Form frmNewPassword 
   BorderStyle     =   0  'None
   Caption         =   "Cambiar Contraseña de cuenta"
   ClientHeight    =   2685
   ClientLeft      =   0
   ClientTop       =   -180
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   179
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1440
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
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
      Left            =   360
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   930
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   435
      Left            =   930
      Tag             =   "0"
      Top             =   2190
      Width           =   1560
   End
End
Attribute VB_Name = "frmNewPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call FormParser.Parse_Form(Me)
    Me.Picture = LoadInterface("password.bmp")

    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "frmNewPassword.Form_Load", Erl)
    Resume Next
    
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Form_KeyPress_Err
    

    If (KeyAscii = 27) Then
        Unload Me

    End If

    
    Exit Sub

Form_KeyPress_Err:
    Call RegistrarError(Err.number, Err.Description, "frmNewPassword.Form_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    Image1.Picture = Nothing
    Image1.Tag = "0"

    
    Exit Sub

Form_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmNewPassword.Form_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Image1_Click()
    
    On Error GoTo Image1_Click_Err
    

    If Text2.Text = "" Then
        Unload Me

    End If

    If Text2.Text <> Text3.Text Then
        Call MsgBox("Las contraseñas no coinciden", vbCritical Or vbOKOnly Or vbApplicationModal Or vbDefaultButton1, "Cambiar Contraseña")
        Exit Sub

    End If
    
    Call WriteChangePassword(Text1.Text, Text2.Text)
    Unload Me

    
    Exit Sub

Image1_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmNewPassword.Image1_Click", Erl)
    Resume Next
    
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image1_MouseMove_Err
    

    If Image1.Tag = "0" Then
        Image1.Picture = LoadInterface("password_aceptar.bmp")
        Image1.Tag = "1"

    End If

    
    Exit Sub

Image1_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmNewPassword.Image1_MouseMove", Erl)
    Resume Next
    
End Sub
