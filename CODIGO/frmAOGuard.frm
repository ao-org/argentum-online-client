VERSION 5.00
Begin VB.Form frmAOGuard 
   Caption         =   "Argentum Guard"
   ClientHeight    =   2280
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
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
   ScaleHeight     =   2280
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSend 
      Caption         =   "Enviar"
      Height          =   360
      Left            =   1680
      TabIndex        =   1
      Top             =   1560
      Width           =   990
   End
   Begin VB.TextBox txtCodigo 
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   3255
   End
End
Attribute VB_Name = "frmAOGuard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSend_Click()
    
    If LenB(txtCodigo.Text) = 0 Then Exit Sub
    
    Call WriteGuardNoticeResponse(txtCodigo.Text, CuentaEmail)
    
End Sub
