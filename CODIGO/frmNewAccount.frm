VERSION 5.00
Begin VB.Form frmNewAccount 
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5136
   ClientLeft      =   11256
   ClientTop       =   10692
   ClientWidth     =   5340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   428
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSurname 
      BackColor       =   &H000D1312&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   2835
      TabIndex        =   1
      Top             =   1605
      Width           =   1815
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H000D1312&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Top             =   1605
      Width           =   1815
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H000D1312&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   2820
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2370
      Width           =   1605
   End
   Begin VB.TextBox txtUsername 
      BackColor       =   &H000D1312&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   2370
      Width           =   1815
   End
   Begin VB.TextBox txtCaptcha 
      BackColor       =   &H000D1312&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Top             =   2970
      Width           =   1095
   End
   Begin VB.TextBox txtValidateMail 
      BackColor       =   &H000D1312&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   690
      TabIndex        =   5
      Top             =   2580
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.TextBox txtCodigo 
      BackColor       =   &H000D1312&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   2955
      TabIndex        =   6
      Top             =   3180
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Label lblResendVerificationCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reenviar código de activación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   600
      TabIndex        =   9
      Top             =   3720
      Visible         =   0   'False
      Width           =   2610
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   4410
      Top             =   2370
      Width           =   255
   End
   Begin VB.Image btnVerValidarCuenta 
      Height          =   300
      Left            =   960
      Top             =   4380
      Width           =   1215
   End
   Begin VB.Image btnCreateAccount 
      Height          =   390
      Left            =   2775
      Top             =   3630
      Width           =   1920
   End
   Begin VB.Image btnRestorePassword 
      Height          =   300
      Left            =   2520
      Top             =   4380
      Width           =   1815
   End
   Begin VB.Label lblCaptchaError 
      BackStyle       =   0  'Transparent
      Caption         =   "Captcha incorrecto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   8880
      TabIndex        =   8
      Top             =   1950
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblCaptcha 
      Alignment       =   2  'Center
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "6 + 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   2970
      Width           =   855
   End
   Begin VB.Image btnCancel 
      Height          =   375
      Left            =   645
      Top             =   3630
      Width           =   1935
   End
   Begin VB.Image btnSendValidarCuenta 
      Height          =   375
      Left            =   2760
      Top             =   4170
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "frmNewAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'    Argentum 20 - Game Client Program
'    Copyright (C) 2022 - Noland Studios
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'
Option Explicit

Private number1 As Byte
Private number2 As Byte
Private equals As Byte

Private cBotonAceptar As clsGraphicalButton
Private cBotonCancelar As clsGraphicalButton
Private cBotonCerrar As clsGraphicalButton

Private Sub btnCancel_Click()
    Unload Me
End Sub
#If PYMMO = 1 Then
Private Sub btnCreateAccount_Click()
    If Val(txtCaptcha.Text) = equals Then
        ModAuth.LoginOperation = e_operation.SignUp
        Call connectToLoginServer
        Call calculateCaptcha
    Else
        Call calculateCaptcha
        Call TextoAlAsistente("Por favor revise el captcha.")
        lblCaptchaError.Visible = True
    End If
End Sub
#End If
#If PYMMO = 0 Then
Private Sub btnCreateAccount_Click()
    CuentaEmail = txtUsername.Text
    CuentaPassword = txtPassword.Text
    
    Call LoginOrConnect(CreandoCuenta)
End Sub
#End If

Private Sub btnRestorePassword_Click()
    Unload Me
    ModAuth.LoginOperation = e_operation.ForgotPassword
    Auth_state = e_state.RequestForgotPassword
    frmPasswordReset.Show , frmConnect
End Sub

Private Sub btnSendValidarCuenta_Click()
        
        txtCodigo.Text = Trim(txtCodigo.Text)
        If txtCodigo.Text <> "" And txtValidateMail.Text <> "" Then
            ModAuth.LoginOperation = e_operation.ValidateAccount
            Call connectToLoginServer
        End If
        
End Sub

Private Sub btnVerValidarCuenta_Click()
    Me.showValidateAccountControls
End Sub



Private Sub Form_Activate()
    Me.Top = frmConnect.Top + frmConnect.Height - Me.Height - 450
    Me.Left = frmConnect.Left + (frmConnect.Width - Me.Width) / 2
End Sub

Private Sub Form_Load()
    Call loadButtons
    Call calculateCaptcha
    Me.Picture = LoadInterface("ventanacrearcuenta.bmp")
End Sub

Private Sub loadButtons()
       
    Set cBotonAceptar = New clsGraphicalButton
    Set cBotonCancelar = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton

    Call cBotonAceptar.Initialize(btnCreateAccount, "boton-crear-cuenta-rojo-default.bmp", _
                                                "boton-crear-cuenta-rojo-over.bmp", _
                                                "boton-crear-cuenta-rojo-off.bmp", Me)
                                                
    Call cBotonCancelar.Initialize(btnCancel, "boton-cancelar-default.bmp", _
                                                "boton-cancelar-over.bmp", _
                                                "boton-cancelar-off.bmp", Me)
                                                
    Call cBotonCerrar.Initialize(btnSendValidarCuenta, "boton-enviar-default.bmp", _
                                                "boton-enviar-over.bmp", _
                                                "boton-enviar-off.bmp", Me)
    
End Sub
Private Sub calculateCaptcha()
    number1 = RandomNumber(0, 9)
    number2 = RandomNumber(0, 9)
    equals = number1 + number2
    lblCaptchaError.Visible = False
    txtCaptcha.Text = ""
    lblCaptcha.Caption = number1 & " + " & number2
End Sub


Public Sub showValidateAccountControls()
    
    Me.Picture = LoadInterface("ventanacrearcuentacodigo.bmp")
    Me.btnSendValidarCuenta.Visible = True
    Me.txtValidateMail.Visible = True
    Me.txtCodigo.Visible = True
    Me.btnCancel.Top = 278
    
    Me.txtUsername.Visible = False
    Me.txtPassword.Visible = False
    Me.txtName.Visible = False
    Me.txtSurname.Visible = False
    Me.txtCaptcha.Visible = False
    Me.lblCaptcha.Visible = False
    Me.lblCaptchaError.Visible = False
    Me.btnVerValidarCuenta.Visible = False
    Me.btnCreateAccount.Visible = False
    Me.btnRestorePassword.Visible = False
    Me.lblResendVerificationCode.Visible = True
    
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.txtPassword.PasswordChar = ""

End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.txtPassword.PasswordChar = "*"

End Sub

Private Sub lblResendVerificationCode_Click()

    If isValidEmail(txtValidateMail) Then
        CuentaEmail = Me.txtValidateMail.Text
        ModAuth.LoginOperation = e_operation.RequestVerificationCode
        Call connectToLoginServer
    Else
        Call TextoAlAsistente("El email ingresado es inválido.")
    End If
    
End Sub
