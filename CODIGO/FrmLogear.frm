VERSION 5.00
Begin VB.Form FrmLogear 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4005
   ClientLeft      =   6300
   ClientTop       =   0
   ClientWidth     =   5340
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkMode        =   1  'Source
   LinkTopic       =   "frmconnect"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "FrmLogear.frx":0000
   ScaleHeight     =   4005
   ScaleWidth      =   5340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox PasswordTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2820
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Tag             =   "1"
      Top             =   1580
      Width           =   1840
   End
   Begin VB.TextBox NameTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   710
      MaxLength       =   100
      TabIndex        =   0
      Tag             =   "0"
      Top             =   1580
      Width           =   1840
   End
   Begin VB.ComboBox lstServers 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      ItemData        =   "FrmLogear.frx":451CE
      Left            =   720
      List            =   "FrmLogear.frx":451D0
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   2520
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label refuerzolbl 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   640
      Top             =   2150
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   420
      Left            =   2750
      Tag             =   "0"
      Top             =   2055
      Width           =   1980
   End
   Begin VB.Image btnCuenta 
      Height          =   420
      Left            =   630
      Tag             =   "0"
      Top             =   3030
      Width           =   1980
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   2750
      Tag             =   "0"
      Top             =   3030
      Width           =   1980
   End
End
Attribute VB_Name = "FrmLogear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Declaraciï¿½n del Api SetLayeredWindowAttributes que establece _
 la transparencia al form

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

'Recupera el estilo de la ventana
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

'Declaraciï¿½n del Api SetWindowLong necesaria para aplicar un estilo _
 al form antes de usar el Api SetLayeredWindowAttributes

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE = (-20)

Private Const LWA_ALPHA = &H2

Private Const WS_EX_LAYERED = &H80000

'Funciï¿½n para saber si formulario ya es transparente. _
 Se le pasa el Hwnd del formulario en cuestiï¿½n

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

'Funciï¿½n que aplica la transparencia, se le pasa el hwnd del form y un valor de 0 a 255
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

Private Sub btnCuenta_Click()
    Call ShellExecute(0, "Open", "https://www.argentum20.com/", "", App.Path, 1)

End Sub

Private Sub Form_Load()
    'MakeFormTransparent Me, vbBlack
    Call FormParser.Parse_Form(Me)
    Me.Top = Me.Top + 2500
    'Call CargarLst
    Call CargarCuentasGuardadas
    Call Aplicar_Transparencia(Me.hwnd, 220)

    Rem Call SetWindowPos(FrmLogear.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    
    #If DEBUGGING = 1 Then
        lstServers.Visible = True
    #End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If btnCuenta.Tag = "1" Then
        btnCuenta.Picture = Nothing
        btnCuenta.Tag = "0"

    End If

    If Image1.Tag = "1" Then
        Image1.Picture = Nothing
        Image1.Tag = "0"

    End If

    If Image3.Tag = "1" Then
        Image3.Picture = Nothing
        Image3.Tag = "0"

    End If

End Sub

Private Sub Image1_Click()
    Call CloseClient

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Image1.Tag = "0" Then
        Image1.Picture = LoadInterface("boton-salir-ES-over.bmp")
        Image1.Tag = "1"

    End If

    If btnCuenta.Tag = "1" Then
        btnCuenta.Picture = Nothing
        btnCuenta.Tag = "0"

    End If

    If Image3.Tag = "1" Then
        Image3.Picture = Nothing
        Image3.Tag = "0"

    End If

End Sub

Private Sub btnCuenta_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If btnCuenta.Tag = "0" Then
        btnCuenta.Picture = LoadInterface("boton-cuenta-ES-over.bmp")
        btnCuenta.Tag = "1"

    End If

    If Image1.Tag = "1" Then
        Image1.Picture = Nothing
        Image1.Tag = "0"

    End If

    If Image3.Tag = "1" Then
        Image3.Picture = Nothing
        Image3.Tag = "0"

    End If

End Sub

Private Sub Image3_Click()
    Call FormParser.Parse_Form(Me, E_WAIT)

    If IntervaloPermiteConectar Then
        If frmmain.Socket1.Connected Then
            frmmain.Socket1.Disconnect
            frmmain.Socket1.Cleanup
            DoEvents

        End If

        CuentaEmail = NameTxt.Text

        Dim aux As String

        aux = PasswordTxt.Text

        CuentaPassword = aux

        If Image4.Tag = "1" Then
            '       If ExisteCuenta(UserCuenta) Then
            '            Call MensajeAdvertencia("La cuenta ya se encuentra almacenada, no ha sido guardada.")
            ' '            RecordarCheck.value = 0
            '      Else

            CuentaRecordada.nombre = CuentaEmail
            CuentaRecordada.Password = aux
            Call GrabarNuevaCuenta(CuentaEmail, aux)
            '      End If
        Else
            Call ResetearCuentas

        End If

        ' If CuentaRecordada(1).Password <> "" And Val(CuentaRecordada(1).Password) <> Val(PasswordTxt.Text) Then
        '          CuentaRecordada(1).Nombre = UserCuenta
        '           CuentaRecordada(1).Password = aux
        '            Call GrabarNuevaCuenta(UserCuenta, aux)
        '            Call MensajeAdvertencia("Se a almacenado la nueva password.")
        '  End If

        If CheckUserDataLoged() = True Then
            EstadoLogin = E_MODO.IngresandoConCuenta
            frmmain.Socket1.HostName = IPdelServidor
            frmmain.Socket1.RemotePort = PuertoDelServidor
            frmmain.Socket1.Connect

        End If

        ServerIndex = lstServers.ListIndex
        Call SaveRAOInit

    End If

End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Image3.Tag = "0" Then
        Image3.Picture = LoadInterface("boton-ingresar-ES-over.bmp")
        Image3.Tag = "1"

    End If

    If btnCuenta.Tag = "1" Then
        btnCuenta.Picture = Nothing
        btnCuenta.Tag = "0"

    End If

    If Image1.Tag = "1" Then
        Image1.Picture = Nothing
        Image1.Tag = "0"

    End If

End Sub

Private Sub Image4_Click()

    If Image4.Tag = "0" Then
        Image4.Picture = LoadInterface("check-amarillo.bmp")
        Call TextoAlAsistente("¡Recordare la cuenta para la proxima!")
        Image4.Tag = "1"
    Else
        Image4.Picture = Nothing
        Image4.Tag = "0"
        Call TextoAlAsistente("¡No recordare nada!")

    End If

End Sub

Private Sub Label1_Click()

    If Image4.Tag = "0" Then
        Image4.Picture = LoadInterface("check-amarillo.bmp")
        Call TextoAlAsistente("ï¿½Recordare la cuenta para la proxima!")
        Image4.Tag = "1"
    Else
        Image4.Picture = Nothing
        Image4.Tag = "0"
        Call TextoAlAsistente("ï¿½No recordare nada!")

    End If

End Sub

Private Sub lstServers_Click()
    IPdelServidor = ServersLst(lstServers.ListIndex + 1).IP
    PuertoDelServidor = ServersLst(lstServers.ListIndex + 1).puerto
End Sub

Private Sub NameTxt_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        prgRun = False
        End
    
    ElseIf KeyCode = vbKeyReturn Then
        Call Image3_Click

    End If

End Sub

Private Sub PasswordTxt_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        prgRun = False
        End

    ElseIf KeyCode = vbKeyReturn Then
        Call Image3_Click

    End If

End Sub

Private Sub PasswordTxt_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If btnCuenta.Tag = "1" Then
        btnCuenta.Picture = Nothing
        btnCuenta.Tag = "0"

    End If

    If Image1.Tag = "1" Then
        Image1.Picture = Nothing
        Image1.Tag = "0"

    End If

    If Image3.Tag = "1" Then
        Image3.Picture = Nothing
        Image3.Tag = "0"

    End If

End Sub
