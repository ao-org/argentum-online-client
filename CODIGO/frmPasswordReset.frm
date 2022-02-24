VERSION 5.00
Begin VB.Form frmPasswordReset 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   5130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5340
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   342
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   356
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPasswordConfirm 
      BackColor       =   &H000D1312&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   2820
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H000D1312&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
      BackColor       =   &H000D1312&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   300
      Left            =   3090
      TabIndex        =   1
      Top             =   1280
      Visible         =   0   'False
      Width           =   1550
   End
   Begin VB.TextBox txtEmail 
      BackColor       =   &H000D1312&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   2280
      Width           =   3855
   End
   Begin VB.Label cmdHaveCode 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000A93D4&
      Height          =   315
      Left            =   1320
      TabIndex        =   5
      Top             =   4400
      Width           =   2715
   End
   Begin VB.Image cmdEnviar 
      Height          =   450
      Left            =   2760
      Top             =   3600
      Width           =   1980
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   600
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label lblSolicitandoCodigo 
      BackStyle       =   0  'Transparent
      Caption         =   "Solicitando cï¿½digo"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   700
      TabIndex        =   4
      Top             =   2600
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "frmPasswordReset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function toggleTextboxs()
    Me.txtEmail.Visible = Not Me.txtEmail.Visible
    Me.txtCodigo.Visible = Not Me.txtCodigo.Visible
    Me.txtPassword.Visible = Not Me.txtPassword.Visible
    Me.txtPasswordConfirm.Visible = Not Me.txtPasswordConfirm.Visible
    cmdHaveCode.Visible = Not cmdHaveCode.Visible
    
    If cmdEnviar.Top = 275 Then
        cmdEnviar.Top = 240
        Image1.Top = 240
    Else
        cmdEnviar.Top = 275
        Image1.Top = 275
    End If
    
    If Me.txtPassword.Visible Then
        Me.Picture = LoadInterface("ventanarecuperarcontrasena2.bmp")
    Else
        Me.Picture = LoadInterface("ventanarecuperarcontrasena.bmp")
    End If
    
End Function

Private Sub cmdEnviar_Click()

    CuentaEmail = Me.txtEmail.Text
    
    If ModAuth.LoginOperation = e_operation.ForgotPassword And Auth_state <> e_state.Idle Then
    
        If Me.txtPassword.Text = "" Or Me.txtPasswordConfirm.Text = "" Or Me.txtCodigo.Text = "" Then
            Call TextoAlAsistente("Falta completar campos.")
            Exit Sub
        End If
        
        If Not isValidEmail(Me.txtEmail.Text) Then
            Call TextoAlAsistente("El email ingresado es inválido.")
            Exit Sub
        End If
        
        If Me.txtPassword.Text <> Me.txtPasswordConfirm.Text Then
            Call TextoAlAsistente("Las contraseï¿½as ingresadas no coinciden.")
            Exit Sub
        End If
        
        ModAuth.LoginOperation = e_operation.ResetPassword
        Call connectToLoginServer
        
    Else
    
        If Not isValidEmail(Me.txtEmail.Text) Then
            Call TextoAlAsistente("El email ingresado es inválido.")
            Exit Sub
        End If
        
        'cmdEnviar.Visible = False
        lblSolicitandoCodigo.Visible = True
        ModAuth.LoginOperation = e_operation.ForgotPassword
        Call connectToLoginServer
    End If
    
End Sub

Private Sub cmdHaveCode_Click()
    If Not isValidEmail(Me.txtEmail.Text) Then
        MsgBox "El email ingresado es inválido."
        Exit Sub
    End If
    
    Call toggleTextboxs
    ModAuth.LoginOperation = e_operation.ForgotPassword
    Auth_state = e_state.RequestForgotPassword
    
End Sub


Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    Me.Picture = LoadInterface("ventanarecuperarcontrasena.bmp")
    
    Me.Left = (frmConnect.Width / 2) - (Me.Width / 2) + frmConnect.Left
    Me.Top = frmConnect.Height - Me.Height - 400 + frmConnect.Top
End Sub

Private Sub Image1_Click()
    Unload Me
End Sub

Public Function isValidEmail(email As String) As Boolean
Dim At As Integer
Dim oneDot As Integer
Dim twoDots As Integer
 
    isValidEmail = True
    At = InStr(1, email, "@", vbTextCompare)
    oneDot = InStr(At + 2, email, ".", vbTextCompare)
    twoDots = InStr(At + 2, email, "..", vbTextCompare)
    If At = 0 Or oneDot = 0 Or Not twoDots = 0 Or Right(email, 1) = "." Then isValidEmail = False
End Function
