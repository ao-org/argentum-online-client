VERSION 5.00
Begin VB.Form frmPasswordReset 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   12075
   ClientTop       =   8520
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "x"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox txtPasswordConfirm 
      BackColor       =   &H80000001&
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
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H80000001&
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
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtEmail 
      BackColor       =   &H80000001&
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
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   3975
   End
   Begin VB.Label lblSolicitandoCodigo 
      Caption         =   "Solicitando código"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1440
      TabIndex        =   12
      Top             =   2640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblCodigoRecuperacion 
      Caption         =   "Código de recuperación"
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   720
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblContraseniaConfirmar 
      Caption         =   "Confirmar contraseña"
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   2040
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label lblContrasenia 
      Caption         =   "Contraeña"
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblEmail 
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Recuperación de contraseña"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   50
      Width           =   3975
   End
   Begin VB.Label Label2 
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label cmdEnviar 
      Caption         =   "Enviar"
      Height          =   255
      Left            =   3720
      TabIndex        =   1
      Top             =   2640
      Width           =   495
   End
End
Attribute VB_Name = "frmPasswordReset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function toggleTextboxs()
    Me.lblEmail.Visible = Not Me.lblEmail.Visible
    Me.txtEmail.Visible = Not Me.txtEmail.Visible
    Me.txtCodigo.Visible = Not Me.txtCodigo.Visible
    Me.txtPassword.Visible = Not Me.txtPassword.Visible
    Me.txtPasswordConfirm.Visible = Not Me.txtPasswordConfirm.Visible
    Me.lblCodigoRecuperacion.Visible = Not Me.lblCodigoRecuperacion.Visible
    Me.lblContrasenia.Visible = Not Me.lblContrasenia.Visible
    Me.lblContraseniaConfirmar.Visible = Not Me.lblContraseniaConfirmar.Visible
End Function

Private Sub cmdEnviar_Click()
 
    CuentaEmail = Me.txtEmail.Text
    
    If ModAuth.LoginOperation = e_operation.ForgotPassword Then
        ModAuth.LoginOperation = e_operation.ResetPassword
    Else
        cmdEnviar.Visible = False
        lblSolicitandoCodigo.Visible = True
        ModAuth.LoginOperation = e_operation.ForgotPassword
        
    End If
        
    Call connectToLoginServer
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

