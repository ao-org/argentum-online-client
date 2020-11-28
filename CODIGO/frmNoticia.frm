VERSION 5.00
Begin VB.Form frmNoticia 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1665
   ClientLeft      =   2760
   ClientTop       =   -810
   ClientWidth     =   5925
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Asen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Estamos cada dia mas cerca de poder lanzar la beta, estamos muy contentos con esto y esperamos que la disfruten."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   5895
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "20/09/2009"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4560
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Noticia sobre la beta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4935
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Noticias"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5895
   End
End
Attribute VB_Name = "frmNoticia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
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
                                      valor As Integer) As Long
  
Dim msg As Long
  
On Error Resume Next
  
If valor < 0 Or valor > 255 Then
   Aplicar_Transparencia = 1
Else
   msg = GetWindowLong(hwnd, GWL_EXSTYLE)
   msg = msg Or WS_EX_LAYERED
     
   SetWindowLong hwnd, GWL_EXSTYLE, msg
     
   'Establece la transparencia
   SetLayeredWindowAttributes hwnd, 0, valor, LWA_ALPHA
  
   Aplicar_Transparencia = 0
  
End If
  
  
If Err Then
   Aplicar_Transparencia = 2
End If
  
End Function

Private Sub Form_Load()

Call Aplicar_Transparencia(Me.hwnd, 180)
End Sub

