VERSION 5.00
Begin VB.Form frmNewAccount 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2640
   ClientLeft      =   11970
   ClientTop       =   10620
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCaptcha 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   285
      Left            =   2520
      TabIndex        =   21
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton btnSendValidarCuenta 
      Caption         =   "Enviar"
      Height          =   495
      Left            =   2160
      TabIndex        =   16
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton btnVerValidarCuenta 
      Caption         =   "Validar cuenta"
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton btnCancel 
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton btnCreateAccount 
      Caption         =   "Crear cuenta"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton btnCreateAccountWeb 
      Caption         =   "Crear cuenta web"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Caption         =   "Cuenta creada correctamente, inserte el código de verificación que se le ha enviado por email."
      Height          =   1815
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   5055
      Begin VB.TextBox txtValidateMail 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   255
         Left            =   2160
         TabIndex        =   19
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   255
         Left            =   2880
         TabIndex        =   14
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Código de validación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         TabIndex        =   18
         Top             =   1080
         Width           =   2535
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         TabIndex        =   17
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta creada correctamente, inserte el código de verificación que se le ha enviado por email."
         Height          =   975
         Left            =   120
         TabIndex        =   15
         Top             =   0
         Width           =   4815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   4935
      Begin VB.TextBox txtUsername 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   1320
         TabIndex        =   8
         Top             =   0
         Width           =   2775
      End
      Begin VB.TextBox txtPassword 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000005&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txtName 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   720
         Width           =   2775
      End
      Begin VB.TextBox txtSurname 
         Alignment       =   2  'Center
         BackColor       =   &H80000001&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000005&
         Height          =   285
         Left            =   1320
         TabIndex        =   5
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label lblCaptchaError 
         BackStyle       =   0  'Transparent
         Caption         =   "Captcha incorrecto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3240
         TabIndex        =   22
         Top             =   1470
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblCaptcha 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1400
         TabIndex        =   20
         Top             =   1460
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000A&
         Caption         =   "Email"
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000A&
         Caption         =   "Password"
         Height          =   255
         Left            =   0
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000A&
         Caption         =   "Name"
         Height          =   255
         Left            =   0
         TabIndex        =   10
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000A&
         Caption         =   "Surname"
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmNewAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private number1 As Byte
Private number2 As Byte
Private equals As Byte


Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnCreateAccount_Click()
    If Val(txtCaptcha.Text) = equals Then
        ModAuth.LoginOperation = e_operation.SignUp
        Call connectToLoginServer
        Call calculateCaptcha
    Else
        Call calculateCaptcha
        lblCaptchaError.Visible = True
    End If
End Sub

Private Sub btnCreateAccountWeb_Click()
  Call ShellExecute(0, "Open", "https://ao20.com.ar/", "", App.Path, 1)
End Sub


Public Sub AlternarControllers()
    Frame1.Visible = False
    btnCreateAccount.Visible = False
    btnCreateAccountWeb.Visible = False
    btnVerValidarCuenta.Visible = False
    
    Frame2.Visible = True
    btnSendValidarCuenta.Visible = True
End Sub

Private Sub btnSendValidarCuenta_Click()
    
        ModAuth.LoginOperation = e_operation.ValidateAccount
        Call connectToLoginServer
End Sub

Private Sub btnVerValidarCuenta_Click()
    AlternarControllers
End Sub

Private Sub Form_Load()
    Call calculateCaptcha
End Sub

Private Sub calculateCaptcha()
    number1 = RandomNumber(0, 9)
    number2 = RandomNumber(0, 9)
    equals = number1 + number2
    lblCaptchaError.Visible = False
    txtCaptcha.Text = ""
    lblCaptcha.Caption = number1 & " + " & number2
End Sub
