VERSION 5.00
Begin VB.Form FrmLogear 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4005
   ClientLeft      =   11865
   ClientTop       =   9450
   ClientWidth     =   5340
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkMode        =   1  'Source
   LinkTopic       =   "frmconnect"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   267
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   356
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox PasswordTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00031413&
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
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   2
      Tag             =   "1"
      Top             =   1560
      Visible         =   0   'False
      Width           =   1840
   End
   Begin VB.TextBox NameTxt 
      Appearance      =   0  'Flat
      BackColor       =   &H00031413&
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
      ScrollBars      =   1  'Horizontal
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
      ItemData        =   "FrmLogear.frx":0000
      Left            =   720
      List            =   "FrmLogear.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2520
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label refuerzolbl 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Image chkRecordar 
      Height          =   255
      Left            =   640
      Top             =   2150
      Width           =   255
   End
   Begin VB.Image cmdIngresar 
      Height          =   420
      Left            =   2750
      Tag             =   "0"
      Top             =   2055
      Width           =   1980
   End
   Begin VB.Image cmdCuenta 
      Height          =   420
      Left            =   645
      Tag             =   "0"
      Top             =   3030
      Width           =   1980
   End
   Begin VB.Image cmdSalir 
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

'Declaracion del Api SetLayeredWindowAttributes que establece _
 la transparencia al form
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

'Recupera el estilo de la ventana
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

'Declaracion del Api SetWindowLong necesaria para aplicar un estilo _
 al form antes de usar el Api SetLayeredWindowAttributes

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000

'Funcion para saber si formulario ya es transparente. _
 Se le pasa el Hwnd del formulario en cuestion

Public bmoving      As Boolean
Public dX           As Integer
Public dy           As Integer

' Constantes para SendMessage
Const WM_SYSCOMMAND As Long = &H112&
Const MOUSE_MOVE    As Long = &HF012&

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long

Private RealizoCambios As String

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Private cBotonSalir As clsGraphicalButton
Private cBotonCuenta As clsGraphicalButton
Private cBotonIngresar As clsGraphicalButton

Private Sub MoverForm()
    
    On Error GoTo moverForm_Err

    Dim res As Long

    ReleaseCapture
    res = SendMessage(Me.hWnd, WM_SYSCOMMAND, MOUSE_MOVE, 0)
    
    Exit Sub

moverForm_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmLogear.moverForm", Erl)
    Resume Next
    
End Sub

Public Function Is_Transparent(ByVal hWnd As Long) As Boolean
    
    On Error GoTo Is_Transparent_Err

    Dim msg As Long

    msg = GetWindowLong(hWnd, GWL_EXSTYLE)

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
    Call RegistrarError(Err.number, Err.Description, "FrmLogear.Is_Transparent", Erl)
    Resume Next
    
End Function

'Funcion que aplica la transparencia, se le pasa el hwnd del form y un valor de 0 a 255
Public Function Aplicar_Transparencia(ByVal hWnd As Long, Valor As Integer) As Long
    On Error GoTo Aplicar_Transparencia_Err
    
    Dim msg As Long

    If Valor < 0 Or Valor > 255 Then
        Aplicar_Transparencia = 1
    Else
        msg = GetWindowLong(hWnd, GWL_EXSTYLE)
        msg = msg Or WS_EX_LAYERED

        SetWindowLong hWnd, GWL_EXSTYLE, msg

        'Establece la transparencia
        SetLayeredWindowAttributes hWnd, 0, Valor, LWA_ALPHA

        Aplicar_Transparencia = 0

    End If

    If Err Then
        Aplicar_Transparencia = 2

    End If

    
    Exit Function

Aplicar_Transparencia_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmLogear.Aplicar_Transparencia", Erl)
    Resume Next
    
End Function

Private Sub cmdCuenta_Click()
    
    On Error GoTo btnCuenta_Click_Err
    
    frmNewAccount.Show , frmConnect

    
    Exit Sub

btnCuenta_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmLogear.btnCuenta_Click", Erl)
    Resume Next
    
End Sub

Private Sub Form_Activate()
    Me.Top = frmConnect.Top + frmConnect.Height - Me.Height - 450
    Me.Left = frmConnect.Left + (frmConnect.Width - Me.Width) / 2
End Sub

Private Sub Form_Load()

    On Error GoTo Form_Load_Err
    
    Call FormParser.Parse_Form(Me)
    
    Call CargarCuentasGuardadas
    Call Aplicar_Transparencia(Me.hwnd, 240)
    
    Me.Picture = LoadInterface("ventanaconectar.bmp")
    
    #If DEBUGGING = 1 Then
        lstServers.Visible = True
    #End If
    
    Me.PasswordTxt.Visible = True

    Call LoadButtons
    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmLogear.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub LoadButtons()

    Set cBotonSalir = New clsGraphicalButton
    Set cBotonCuenta = New clsGraphicalButton
    Set cBotonIngresar = New clsGraphicalButton
    
    Call cBotonSalir.Initialize(cmdSalir, "boton-salir-default.bmp", "boton-salir-over.bmp", "boton-salir-off.bmp", Me)
    Call cBotonCuenta.Initialize(cmdCuenta, "boton-cuenta-default.bmp", "boton-cuenta-over.bmp", "boton-cuenta-off.bmp", Me)
    Call cBotonIngresar.Initialize(cmdIngresar, "boton-ingresar-default.bmp", "boton-ingresar-over.bmp", "boton-ingresar-off.bmp", Me)
End Sub

Private Sub cmdSalir_Click()
    
    Call CloseClient

End Sub

Private Sub cmdIngresar_Click()
    
    On Error GoTo cmdIngresar_Click_Err
    
    Call FormParser.Parse_Form(Me, E_WAIT)

    If IntervaloPermiteConectar Then
        CuentaEmail = NameTxt.Text
        CuentaPassword = PasswordTxt.Text

        If chkRecordar.Tag = "1" Then
            CuentaRecordada.nombre = CuentaEmail
            CuentaRecordada.Password = CuentaPassword
            
            Call GuardarCuenta(CuentaEmail, CuentaPassword)
        Else
            ' Reseteamos los datos de cuenta guardados
            Call GuardarCuenta(vbNullString, vbNullString)
        End If

        If CheckUserDataLoged() = True Then
            ModAuth.LoginOperation = e_operation.Authenticate
            Call LoginOrConnect(E_MODO.IngresandoConCuenta)
        End If

        ServerIndex = lstServers.ListIndex
        
        Call SaveRAOInit

    End If

    Exit Sub

cmdIngresar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmLogear.cmdIngresar_Click", Erl)
    Resume Next
    
End Sub

Private Sub chkRecordar_Click()
    
    On Error GoTo chkRecordar_Click_Err

    If chkRecordar.Tag = "0" Then
        chkRecordar.Picture = LoadInterface("check-amarillo.bmp")
        Call TextoAlAsistente("¡Recordare la cuenta para la proxima!")
        chkRecordar.Tag = "1"
    Else
        chkRecordar.Picture = Nothing
        chkRecordar.Tag = "0"
        Call TextoAlAsistente("¡No recordare nada!")
    End If

    
    Exit Sub

chkRecordar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmLogear.chkRecordar_Click", Erl)
    Resume Next
    
End Sub

Private Sub lstServers_Click()
    
    On Error GoTo lstServers_Click_Err
    
    IPdelServidor = ServersLst(lstServers.ListIndex + 1).IP
    PuertoDelServidor = ServersLst(lstServers.ListIndex + 1).puerto
    IPdelServidorLogin = ServersLst(lstServers.ListIndex + 1).IpLogin
    PuertoDelServidorLogin = ServersLst(lstServers.ListIndex + 1).puertoLogin
    
    Exit Sub

lstServers_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmLogear.lstServers_Click", Erl)
    Resume Next
    
End Sub

Private Sub NameTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo NameTxt_KeyDown_Err

    If KeyCode = 27 Then
        prgRun = False
        End
    ElseIf KeyCode = vbKeyReturn Then
        Call cmdIngresar_Click
    End If
    
    Exit Sub

NameTxt_KeyDown_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmLogear.NameTxt_KeyDown", Erl)
    Resume Next
    
End Sub

Private Sub PasswordTxt_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo PasswordTxt_KeyDown_Err

    If KeyCode = 27 Then
        prgRun = False
        End
    ElseIf KeyCode = vbKeyReturn Then
        Call cmdIngresar_Click
    End If
    
    Exit Sub

PasswordTxt_KeyDown_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmLogear.PasswordTxt_KeyDown", Erl)
    Resume Next
    
End Sub
