VERSION 5.00
Begin VB.Form frmCrearCuenta 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3480
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmCrearCuenta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox texVer 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      IMEMode         =   3  'DISABLE
      Left            =   360
      MaxLength       =   5
      TabIndex        =   2
      Top             =   3920
      Width           =   975
   End
   Begin VB.TextBox Constraseña 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      IMEMode         =   3  'DISABLE
      Left            =   480
      MaxLength       =   20
      PasswordChar    =   "•"
      TabIndex        =   1
      Top             =   2040
      Width           =   2535
   End
   Begin VB.TextBox Email 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   480
      MaxLength       =   320
      TabIndex        =   0
      Top             =   1450
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   1095
      Left            =   120
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   1320
      TabIndex        =   4
      Top             =   3840
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   1520
      Tag             =   "0"
      Top             =   4130
      Width           =   1770
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   330
      Tag             =   "0"
      Top             =   4125
      Width           =   1020
   End
   Begin VB.Label valcar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "96666"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   3880
      Width           =   855
   End
End
Attribute VB_Name = "frmCrearCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim ValidacionNumber As Long

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
  
Public Function Is_Transparent(ByVal hwnd As Long) As Boolean
    
    On Error GoTo Is_Transparent_Err
    

    
  
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
  
    
    Exit Function

Is_Transparent_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCrearCuenta.Is_Transparent", Erl)
    Resume Next
    
End Function
  
'Función que aplica la transparencia, se le pasa el hwnd del form y un valor de 0 a 255
Public Function Aplicar_Transparencia(ByVal hwnd As Long, Valor As Integer) As Long
    
    On Error GoTo Aplicar_Transparencia_Err
    
  
    Dim msg As Long
  
    
  
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
  
    
    Exit Function

Aplicar_Transparencia_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCrearCuenta.Aplicar_Transparencia", Erl)
    Resume Next
    
End Function

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call FormParser.Parse_Form(Me)
    Call Aplicar_Transparencia(Me.hwnd, 240)
    Me.Picture = LoadInterface("crearcuenta.bmp")
    ValidacionNumber = RandomNumber(10000, 90000)

    valcar = ValidacionNumber

    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCrearCuenta.Form_Load", Erl)
    Resume Next
    
End Sub

Private Function CheckearDatos() As Boolean
    
    On Error GoTo CheckearDatos_Err
    

    Dim loopc     As Long

    Dim CharAscii As Integer
    
    If Len(Email.Text) = 0 Or Not CheckMailString(Email.Text) Then
        MsgBox ("Dirección de email invalida")
        Exit Function

    End If
    
    If Len(Constraseña.Text) = 0 Then
        MsgBox ("Ingrese un password.")
        Exit Function

    End If
    
    For loopc = 1 To Len(Constraseña.Text)
        CharAscii = Asc(mid$(Constraseña.Text, loopc, 1))

        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Password inválido. El caractér " & Chr$(CharAscii) & " no está permitido.")
            Exit Function

        End If

    Next loopc
    
    CheckearDatos = True

    
    Exit Function

CheckearDatos_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCrearCuenta.CheckearDatos", Erl)
    Resume Next
    
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Form_MouseMove_Err
    

    If Image2.Tag = "1" Then
        Image2.Picture = Nothing
        Image2.Tag = "0"

    End If

    If Image1.Tag = "1" Then
        Image1.Picture = Nothing
        Image1.Tag = "0"

    End If

    
    Exit Sub

Form_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCrearCuenta.Form_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Image1_Click()
    
    On Error GoTo Image1_Click_Err
    
    Call Sound.Sound_Play(SND_CLICK)

    Unload Me
    frmMasOpciones.Show , frmConnect
    frmMasOpciones.Top = frmMasOpciones.Top + 3000

    
    Exit Sub

Image1_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCrearCuenta.Image1_Click", Erl)
    Resume Next
    
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Image1.Picture = LoadInterface("volverpress.bmp")
    
    On Error GoTo Image1_MouseDown_Err
    
    Image1.Tag = "1"

    
    Exit Sub

Image1_MouseDown_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCrearCuenta.Image1_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image1_MouseMove_Err
    

    If Image1.Tag = "0" Then
        Image1.Picture = LoadInterface("crearcuenta_volver.bmp")
        Image1.Tag = "1"

    End If

    
    Exit Sub

Image1_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCrearCuenta.Image1_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Image2_Click()
    
    On Error GoTo Image2_Click_Err
    
    Call Sound.Sound_Play(SND_CLICK)
    
    If texVer = "" Then
        Call MensajeAdvertencia("El campo de texto de verificacion esta vacio.")
        Exit Sub

    End If
    
    If ValidacionNumber <> texVer Then
        Call MensajeAdvertencia("El codigo de verificación es invalido, por favor reintente.")
        Exit Sub

    End If

    If CheckearDatos Then
        CuentaPassword = Constraseña
        CuentaEmail = Email
    
        EstadoLogin = E_MODO.CreandoCuenta

        If frmMain.Socket1.Connected Then
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
            DoEvents

        End If

        frmMain.Socket1.HostName = IPdelServidor
        frmMain.Socket1.RemotePort = PuertoDelServidor
        frmMain.Socket1.Connect

        'Unload Me
    End If

    
    Exit Sub

Image2_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCrearCuenta.Image2_Click", Erl)
    Resume Next
    
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image2_MouseDown_Err
    

    If Image1.Tag = "0" Then
        ' Image2.Picture = LoadInterface("crearcuentapress.bmp")
        Image2.Tag = "1"

    End If

    
    Exit Sub

Image2_MouseDown_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCrearCuenta.Image2_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image2_MouseMove_Err
    

    If Image2.Tag = "0" Then
        Image2.Picture = LoadInterface("crearcuentahover.bmp")
        Image2.Tag = "1"

    End If

    
    Exit Sub

Image2_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCrearCuenta.Image2_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Label1_Click()
    
    On Error GoTo Label1_Click_Err
    
    texVer.SetFocus

    
    Exit Sub

Label1_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCrearCuenta.Label1_Click", Erl)
    Resume Next
    
End Sub

Private Sub valcar_Click()
    
    On Error GoTo valcar_Click_Err
    
    texVer.SetFocus

    
    Exit Sub

valcar_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmCrearCuenta.valcar_Click", Erl)
    Resume Next
    
End Sub
