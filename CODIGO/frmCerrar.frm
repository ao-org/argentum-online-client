VERSION 5.00
Begin VB.Form frmCerrar 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image cmdCancelar 
      Height          =   420
      Left            =   630
      Top             =   1750
      Width           =   1980
   End
   Begin VB.Image cmdSalir 
      Height          =   420
      Left            =   640
      Top             =   1180
      Width           =   1980
   End
   Begin VB.Image cmdMenuPrincipal 
      Height          =   420
      Left            =   640
      Top             =   610
      Width           =   1980
   End
End
Attribute VB_Name = "frmCerrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  
'Función para saber si formulario ya es transparente. _
 Se le pasa el Hwnd del formulario en cuestión
 
Public bmoving      As Boolean

Public dX           As Integer

Public dy           As Integer

Private RealizoCambios As String

Private cBotonAceptar As clsGraphicalButton
Private cBotonConstruir As clsGraphicalButton
Private cBotonCerrar As clsGraphicalButton


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  

    If (KeyCode = vbKeyEscape) Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call FormParser.Parse_Form(Me)
    Call Aplicar_Transparencia(Me.hwnd, 240)

    
    Me.Picture = LoadInterface("desconectar.bmp")
    
    Call LoadButtons
    
    Exit Sub
    
Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCerrar.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub LoadButtons()
        
    Set cBotonAceptar = New clsGraphicalButton
    Set cBotonConstruir = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton

    Call cBotonAceptar.Initialize(cmdMenuPrincipal, "boton-mainmenu-ES-default.bmp", _
                                                "boton-mainmenu-ES-over.bmp", _
                                                "boton-mainmenu-ES-off.bmp", Me)
    
    Call cBotonConstruir.Initialize(cmdCancelar, "boton-cancelar-ES-default.bmp", _
                                                "boton-cancelar-ES-over.bmp", _
                                                "boton-cancelar-ES-off.bmp", Me)
                                                
    Call cBotonCerrar.Initialize(cmdSalir, "boton-salir-ES-default.bmp", _
                                                "boton-salir-ES-over.bmp", _
                                                "boton-salir-ES-off.bmp", Me)
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Form_KeyPress_Err
    

    'If (KeyAscii = 27) Then
    '    Unload Me
    'End If

    
    Exit Sub

Form_KeyPress_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCerrar.Form_KeyPress", Erl)
    Resume Next
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdMenuPrincipal_Click()
    Call WriteQuit
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    Call CloseClient
End Sub

