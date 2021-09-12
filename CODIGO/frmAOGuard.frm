VERSION 5.00
Begin VB.Form frmAOGuard 
   Appearance      =   0  'Flat
   BackColor       =   &H80000001&
   BorderStyle     =   0  'None
   Caption         =   "Argentum Guard"
   ClientHeight    =   4005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5340
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
   ScaleHeight     =   267
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   356
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Se ha enviado un mensaje a tu casilla de correo. Ingrese el código recibido en la siguiente casilla:"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1200
      TabIndex        =   1
      Top             =   720
      Width           =   3135
   End
   Begin VB.Image cmdSalir 
      Height          =   420
      Left            =   600
      Tag             =   "0"
      Top             =   3000
      Width           =   1980
   End
   Begin VB.Image cmdSend 
      Height          =   420
      Left            =   2880
      Tag             =   "0"
      Top             =   3000
      Width           =   1980
   End
End
Attribute VB_Name = "frmAOGuard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private cBotonSalir As clsGraphicalButton
Private cBotonIngresar As clsGraphicalButton

Private Sub cmdSalir_Click()
    
    IsConnected = False
    Call modNetwork.Disconnect

    Unload Me
    
End Sub

Private Sub Form_Load()

    Call FormParser.Parse_Form(Me)
    
    Call Aplicar_Transparencia(Me.hwnd, 240)
    
    'Me.Picture = LoadInterface("ventanaconectar.bmp")
    Me.Top = FrmLogear.Top
    Me.Left = FrmLogear.Left
    
    Me.txtCodigo.MaxLength = 5
    
    Call LoadButtons
    
End Sub

Private Sub Form_Activate()
    
    If FrmLogear.Visible Then FrmLogear.Hide
    
End Sub

Private Sub LoadButtons()

    Set cBotonSalir = New clsGraphicalButton
    Set cBotonIngresar = New clsGraphicalButton
    
    Call cBotonSalir.Initialize(cmdSalir, "boton-salir-ES-default.bmp", "boton-salir-ES-over.bmp", "boton-salir-ES-off.bmp", Me)
    Call cBotonIngresar.Initialize(cmdSend, "boton-ingresar-ES-default.bmp", "boton-ingresar-ES-over.bmp", "boton-ingresar-ES-off.bmp", Me)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If prgRun Then
        If Not FrmLogear.Visible Then FrmLogear.Show
    End If
    
End Sub

Private Sub cmdSend_Click()
    
    If LenB(txtCodigo.Text) = 0 Then Exit Sub
    
    Call WriteGuardNoticeResponse(txtCodigo.Text)
    
    Unload Me
    
End Sub
