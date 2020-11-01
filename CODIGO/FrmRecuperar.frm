VERSION 5.00
Begin VB.Form FrmRecuperar 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3525
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   3525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox UserCuentatxt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   1500
      Width           =   2535
   End
   Begin VB.TextBox texVer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   600
      MaxLength       =   5
      TabIndex        =   2
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox Email 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   480
      TabIndex        =   1
      Top             =   2110
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   400
      Tag             =   "0"
      Top             =   3360
      Width           =   1080
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   2030
      Tag             =   "0"
      Top             =   3350
      Width           =   1080
   End
   Begin VB.Label valcar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "12345"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   2850
      Width           =   975
   End
End
Attribute VB_Name = "FrmRecuperar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim ValidacionNumber As Long

'Declaraci�n del Api SetLayeredWindowAttributes que establece _
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
  
  
'Declaraci�n del Api SetWindowLong necesaria para aplicar un estilo _
 al form antes de usar el Api SetLayeredWindowAttributes
  
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
               (ByVal hwnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long
  
  
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000
'Funci�n para saber si formulario ya es transparente. _
 Se le pasa el Hwnd del formulario en cuesti�n
  
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
  
'Funci�n que aplica la transparencia, se le pasa el hwnd del form y un valor de 0 a 255
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
Private Sub Form_Load()
Call FormParser.Parse_Form(Me)
Call Aplicar_Transparencia(Me.hwnd, 240)
Me.Picture = LoadInterface("recuperar.bmp")
    ValidacionNumber = RandomNumber(10000, 90000)
    valcar = ValidacionNumber
    UserCuentatxt = CuentaEmail
    Email = CuentaEmail
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Image2.Tag = "1" Then
   Image2.Picture = Nothing
   Image2.Tag = "0"
End If
If Image1.Tag = "1" Then
    Image1.Picture = Nothing
    Image1.Tag = "0"
End If
End Sub




Private Sub Image1_Click()
        Unload Me
frmMasOpciones.Show , frmConnect
frmMasOpciones.Top = frmMasOpciones.Top + 3000
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
        'Image1.Picture = LoadInterface("volverpress.bmp")
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Image1.Tag = "0" Then
        Image1.Picture = LoadInterface("volverhover.bmp")
        Image1.Tag = "1"
    End If
End Sub

Private Sub Image2_Click()
    If UserCuentatxt = "" Then
        Call MensajeAdvertencia("El campo de nombre de la cuenta esta vacio.")
        Exit Sub
    End If
   
    If Email = "" Then
        Call MensajeAdvertencia("El campo de email esta vacio.")
        Exit Sub
    End If

    If texVer = "" Then
        Call MensajeAdvertencia("El campo de texto de verificacion esta vacio.")
        Exit Sub
    End If
    
    If ValidacionNumber <> texVer Then
        Call MensajeAdvertencia("El codigo de verificaci�n es invalido, por favor reintente.")
        Exit Sub
    End If


    EstadoLogin = E_MODO.RecuperandoConstrase�a
    CuentaEmail = CuentaEmail

    If frmmain.Socket1.Connected Then
        frmmain.Socket1.Disconnect
        frmmain.Socket1.Cleanup
        DoEvents
    End If
    frmmain.Socket1.HostName = IPdelServidor
    frmmain.Socket1.RemotePort = PuertoDelServidor
    frmmain.Socket1.Connect
    Unload Me
        
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Image2.Tag = "0" Then
        Image2.Picture = LoadInterface("enviarhover.bmp")
        Image2.Tag = "1"
    End If
End Sub
Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
        'Image2.Picture = LoadInterface("enviarpress.bmp")
End Sub

