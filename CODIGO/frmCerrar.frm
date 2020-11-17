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
   Begin VB.Image Opcion 
      Height          =   420
      Index           =   2
      Left            =   630
      Top             =   1750
      Width           =   1980
   End
   Begin VB.Image Opcion 
      Height          =   420
      Index           =   1
      Left            =   640
      Top             =   1180
      Width           =   1980
   End
   Begin VB.Image Opcion 
      Height          =   420
      Index           =   0
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
  
'Declaración del Api SetLayeredWindowAttributes que establece _
 la transparencia al form
  
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
  
'Recupera el estilo de la ventana
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
  
'Declaración del Api SetWindowLong necesaria para aplicar un estilo _
 al form antes de usar el Api SetLayeredWindowAttributes
  
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
  
Private Const GWL_EXSTYLE = (-20)

Private Const LWA_ALPHA = &H2

Private Const WS_EX_LAYERED = &H80000

'Función para saber si formulario ya es transparente. _
 Se le pasa el Hwnd del formulario en cuestión
 
Public bmoving      As Boolean

Public dX           As Integer

Public dy           As Integer

' Constantes para SendMessage
Const WM_SYSCOMMAND As Long = &H112&

Const MOUSE_MOVE    As Long = &HF012&

Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long

Private RealizoCambios As String

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST = -1

Private Const HWND_NOTOPMOST = -2

Private Const SWP_NOMOVE = &H2

Private Const SWP_NOSIZE = &H1

Private Sub moverForm()

    Dim res As Long

    ReleaseCapture
    res = SendMessage(Me.hwnd, WM_SYSCOMMAND, MOUSE_MOVE, 0)

End Sub

Public Function Is_Transparent(ByVal hwnd As Long) As Boolean

    On Error Resume Next

    Dim msg As Long

    msg = GetWindowLong(hwnd, GWL_EXSTYLE)

    If (msg And WS_EX_LAYERED) = WS_EX_LAYERED Then
        Is_Transparent = True
    Else
        Is_Transparent = False

    End If

    If Err Then
        Is_Transparent = False

    End If

End Function
  
'Función que aplica la transparencia, se le pasa el hwnd del form y un valor de 0 a 255
Public Function Aplicar_Transparencia(ByVal hwnd As Long, Valor As Integer) As Long

    Dim msg As Long

    On Error Resume Next

    If Valor < 0 Or Valor > 255 Then
        Aplicar_Transparencia = 1
    Else
        msg = GetWindowLong(hwnd, GWL_EXSTYLE)
        msg = msg Or WS_EX_LAYERED
        SetWindowLong hwnd, GWL_EXSTYLE, msg
        'Establece la transparencia
        SetLayeredWindowAttributes hwnd, 0, Valor, LWA_ALPHA
        Aplicar_Transparencia = 0

    End If
  
    If Err Then
        Aplicar_Transparencia = 2

    End If

End Function

Private Sub Form_Load()
    Call FormParser.Parse_Form(Me)
    Call Aplicar_Transparencia(Me.hwnd, 220)
    Me.Picture = LoadInterface("desconectar.bmp")

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Opcion(0).Tag = "0"
    Opcion(0).Picture = Nothing
    Opcion(1).Tag = "0"
    Opcion(1).Picture = Nothing
    Opcion(2).Tag = "0"
    Opcion(2).Picture = Nothing

End Sub

Private Sub Opcion_Click(Index As Integer)

    'Ladder 30/10/2020
    Select Case Index

        Case 0 ' Menu principal
            Call WriteQuit
            Unload Me

        Case 1  'Cerrar juego
            Call CloseClient

        Case 2 'Cancelar
            Unload Me

    End Select

End Sub

Private Sub Opcion_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    'Ladder 30/10/2020
    Select Case Index

        Case 0 ' Menu principal
            Opcion(Index).Picture = LoadInterface("boton-menu-principal-ES-off.bmp")
            Opcion(Index).Tag = "1"

        Case 1  'Cerrar juego
            Opcion(Index).Picture = LoadInterface("boton-salir-ES-off.bmp")
            Opcion(Index).Tag = "1"

        Case 2 'Cancelar
            Opcion(Index).Picture = LoadInterface("boton-cancelar-ES-off.bmp")
            Opcion(Index).Tag = "1"

    End Select

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)

    If (KeyAscii = 27) Then
        Unload Me

    End If

End Sub

Private Sub Opcion_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    'Ladder 30/10/2020
    Select Case Index

        Case 0 ' Menu principal

            If Opcion(Index).Tag = "0" Then
                Opcion(Index).Picture = LoadInterface("boton-menu-principal-ES-over.bmp")
                Opcion(Index).Tag = "1"

            End If

        Case 1  'Cerrar juego

            If Opcion(Index).Tag = "0" Then
                Opcion(Index).Picture = LoadInterface("boton-salir-ES-over.bmp")
                Opcion(Index).Tag = "1"

            End If

        Case 2 'Cancelar

            If Opcion(Index).Tag = "0" Then
                Opcion(Index).Picture = LoadInterface("boton-cancelar-ES-over.bmp")
                Opcion(Index).Tag = "1"

            End If

    End Select

End Sub
