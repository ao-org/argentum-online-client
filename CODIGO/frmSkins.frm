VERSION 5.00
Begin VB.Form frmSkins 
   BorderStyle     =   0  'None
   ClientHeight    =   7245
   ClientLeft      =   16350
   ClientTop       =   5160
   ClientWidth     =   3600
   ControlBox      =   0   'False
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
   ScaleHeight     =   469.422
   ScaleMode       =   0  'User
   ScaleWidth      =   241
   ShowInTaskbar   =   0   'False
   Begin VB.Image cmdCerrar 
      Height          =   420
      Left            =   3120
      Tag             =   "0"
      Top             =   0
      Width           =   465
   End
End
Attribute VB_Name = "frmSkins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Constantes para simular el movimiento de la ventana
Private Const WM_SYSCOMMAND As Long = &H112&
Private Const MOUSE_MOVE    As Long = &HF012&
' Declaraciones de funciones API de Windows
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
' Declaración de variable con evento de inventario gráfico
Public WithEvents InvKeys As clsGrapchicalInventory
Attribute InvKeys.VB_VarHelpID = -1

' Evento al hacer clic en el botón "Cerrar"
Private Sub cmdCerrar_Click()
    ' Oculta el formulario
    Call frmSkins.WalletSkins
End Sub

' Evento al cargar el formulario
Private Sub Form_Load()
    ' Parsea la interfaz del formulario (diseño)
    Call FormParser.Parse_Form(Me)
    ' Aplica transparencia al formulario (valor 240 de opacidad)
    Call Aplicar_Transparencia(Me.hWnd, 240)
    ' Carga la imagen de fondo del formulario desde archivo
    frmSkins.Picture = LoadInterface("ventanaskins.bmp")
    Exit Sub
End Sub

' Permite mover el formulario arrastrándolo con el mouse
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        Call ReleaseCapture
        Call SendMessage(Me.hWnd, WM_SYSCOMMAND, MOUSE_MOVE, 0&)
    End If
End Sub

' Método público para ocultar el formulario
Public Sub WalletSkins()
    On Error GoTo WalletSkins_Err
    frmSkins.visible = False
    Exit Sub
WalletSkins_Err:
    ' Manejo de errores estandarizado
    Call RegistrarError(Err.Number, Err.Description, "frmSkins.WalletSkins", Erl)
    Resume Next
End Sub
