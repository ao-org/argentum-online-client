VERSION 5.00
Begin VB.Form FormOpciones 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4080
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   390
      Index           =   7
      Left            =   240
      Tag             =   "0"
      Top             =   3960
      Width           =   2955
   End
   Begin VB.Image Image1 
      Height          =   390
      Index           =   6
      Left            =   240
      Tag             =   "0"
      Top             =   3360
      Width           =   2955
   End
   Begin VB.Image Image1 
      Height          =   390
      Index           =   5
      Left            =   120
      Tag             =   "0"
      Top             =   3000
      Width           =   2955
   End
   Begin VB.Image Image1 
      Height          =   390
      Index           =   4
      Left            =   240
      Tag             =   "0"
      Top             =   2400
      Width           =   2955
   End
   Begin VB.Image Image1 
      Height          =   390
      Index           =   3
      Left            =   360
      Tag             =   "0"
      Top             =   1320
      Width           =   2955
   End
End
Attribute VB_Name = "FormOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  
'Declaración del Api SetLayeredWindowAttributes que establece _
 la transparencia al form
  
Private Declare Function SetLayeredWindowAttributes Lib "user32" _
                (ByVal hwnd As Long, _
                 ByVal crKey As Long, _
                 ByVal bAlpha As Byte, _
                 ByVal dwFlags As Long) As Long
  
  
'Recupera el estilo de la ventana
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
                (ByVal hwnd As Long, _
                 ByVal nIndex As Long) As Long
  
  
'Declaración del Api SetWindowLong necesaria para aplicar un estilo _
 al form antes de usar el Api SetLayeredWindowAttributes
  
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
               (ByVal hwnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long
  
  
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000
'Función para saber si formulario ya es transparente. _
 Se le pasa el Hwnd del formulario en cuestión
 
Public bmoving As Boolean
Public dX As Integer
Public dy As Integer

' Constantes para SendMessage
Const WM_SYSCOMMAND As Long = &H112&
Const MOUSE_MOVE As Long = &HF012&

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hwnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, lParam As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
  ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
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
Public Function Aplicar_Transparencia(ByVal hwnd As Long, _
                                      Valor As Integer) As Long
  
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
            Unload Me
End If
End Sub

Private Sub Form_Load()
    'Call Aplicar_Transparencia(Me.hwnd, 20)
    Call FormParser.Parse_Form(Me)
    Me.Picture = General_Load_Picture_From_Resource_Ex("opcionesingame.bmp")
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim i As Byte

For i = 0 To 7

'Image1(i).Picture = Nothing
'Image1(i).Tag = "0"
Next i
End Sub

Private Sub Image1_Click(Index As Integer)
    Select Case Index

        
        Case 3
            Call WriteRequestGrupo
            Unload Me
        
        Case 4
            Unload Me
            Call WriteTraerShop
        
        Case 5
                Unload Me
         Call frmOpciones.Init
        
        Case 6
        
            If UserParalizado Or UserInmovilizado Then 'Inmo
                With FontTypes(FontTypeNames.FONTTYPE_WARNING)
                    Call ShowConsoleMsg("No puedes salir estando paralizado.", .red, .green, .blue, .bold, .italic)
                End With
                Exit Sub
            End If
            Call WriteQuit
        
        Case 7
             Unload Me
        
    End Select
End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Index
        Case 0
                Image1(Index).Picture = General_Load_Picture_From_Resource_Ex("estadisticaswidepress.bmp")
                Image1(Index).Tag = "1"
        Case 1
                Image1(Index).Picture = General_Load_Picture_From_Resource_Ex("claneswidepress.bmp")
                Image1(Index).Tag = "1"
        
        Case 2
                Image1(Index).Picture = General_Load_Picture_From_Resource_Ex("manualpress.bmp")
                Image1(Index).Tag = "1"
        
        Case 3
                Image1(Index).Picture = General_Load_Picture_From_Resource_Ex("gruposwidepress.bmp")
                Image1(Index).Tag = "1"
        
        Case 4
                Image1(Index).Picture = General_Load_Picture_From_Resource_Ex("shopdonawidepress.bmp")
                Image1(Index).Tag = "1"

        
        Case 5

                Image1(Index).Picture = General_Load_Picture_From_Resource_Ex("opcionesjuegowidepress.bmp")
                Image1(Index).Tag = "1"
        
        Case 6
                Image1(Index).Picture = General_Load_Picture_From_Resource_Ex("desconectarwidepress.bmp")
                Image1(Index).Tag = "1"
        
        Case 7
                Image1(Index).Picture = General_Load_Picture_From_Resource_Ex("cerrarwidepress.bmp")
                Image1(Index).Tag = "1"
        
    End Select
End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Select Case Index
        Case 0
            If Image1(Index).Tag = "0" Then
                Image1(Index).Picture = General_Load_Picture_From_Resource_Ex("estadisticaswidehover.bmp")
                Image1(Index).Tag = "1"
            End If
        Case 1
            If Image1(Index).Tag = "0" Then
                Image1(Index).Picture = General_Load_Picture_From_Resource_Ex("claneswidehover.bmp")
                Image1(Index).Tag = "1"
            End If
        
        Case 2
            If Image1(Index).Tag = "0" Then
                Image1(Index).Picture = General_Load_Picture_From_Resource_Ex("manualhover.bmp")
                Image1(Index).Tag = "1"
            End If
        
        Case 3
            If Image1(Index).Tag = "0" Then
                Image1(Index).Picture = General_Load_Picture_From_Resource_Ex("gruposwidehover.bmp")
                Image1(Index).Tag = "1"
            End If
        
        Case 4
            If Image1(Index).Tag = "0" Then
                Image1(Index).Picture = General_Load_Picture_From_Resource_Ex("shopdonawidehover.bmp")
                Image1(Index).Tag = "1"
            End If
        
        Case 5
            If Image1(Index).Tag = "0" Then
                Image1(Index).Picture = General_Load_Picture_From_Resource_Ex("opcionesjuegowidehover.bmp")
                Image1(Index).Tag = "1"
            End If
        
        Case 6
            If Image1(Index).Tag = "0" Then
                Image1(Index).Picture = General_Load_Picture_From_Resource_Ex("desconectarwidehover.bmp")
                Image1(Index).Tag = "1"
            End If
        
        Case 7
            If Image1(Index).Tag = "0" Then
                Image1(Index).Picture = General_Load_Picture_From_Resource_Ex("cerrarwidehover.bmp")
                Image1(Index).Tag = "1"
            End If
        
    End Select
End Sub
