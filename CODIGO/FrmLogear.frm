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
    
    On Error GoTo moverForm_Err
    

    Dim res As Long

    ReleaseCapture
    res = SendMessage(Me.hwnd, WM_SYSCOMMAND, MOUSE_MOVE, 0)

    
    Exit Sub

moverForm_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmLogear.moverForm", Erl)
    Resume Next
    
End Sub

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
    Call RegistrarError(Err.number, Err.Description, "FrmLogear.Is_Transparent", Erl)
    Resume Next
    
End Function

'Funciï¿½n que aplica la transparencia, se le pasa el hwnd del form y un valor de 0 a 255
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
    Call RegistrarError(Err.number, Err.Description, "FrmLogear.Aplicar_Transparencia", Erl)
    Resume Next
    
End Function

Private Sub btnCuenta_Click()
    
    On Error GoTo btnCuenta_Click_Err
    
    Call ShellExecute(0, "Open", "https://ao20.com.ar/", "", App.Path, 1)

    
    Exit Sub

btnCuenta_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmLogear.btnCuenta_Click", Erl)
    Resume Next
    
End Sub

Private Sub Form_Load()
    'MakeFormTransparent Me, vbBlack
    
    On Error GoTo Form_Load_Err
    
    Call FormParser.Parse_Form(Me)
    Me.Top = Me.Top + 2500
    'Call CargarLst
    Call CargarCuentasGuardadas
    Call Aplicar_Transparencia(Me.hwnd, 220)
    
    'TODO: Me.Picture = LoadInterface("")

    Rem Call SetWindowPos(FrmLogear.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    
    #If DEBUGGING = 1 Then
        lstServers.Visible = True
    #End If

    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmLogear.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Form_MouseMove_Err
    

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

    
    Exit Sub

Form_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmLogear.Form_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Image1_Click()
    
    On Error GoTo Image1_Click_Err
    
    Call CloseClient

    
    Exit Sub

Image1_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmLogear.Image1_Click", Erl)
    Resume Next
    
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image1_MouseMove_Err
    

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

    
    Exit Sub

Image1_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmLogear.Image1_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub btnCuenta_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo btnCuenta_MouseMove_Err
    

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

    
    Exit Sub

btnCuenta_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmLogear.btnCuenta_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Image3_Click()
    
    On Error GoTo Image3_Click_Err
    
    Call FormParser.Parse_Form(Me, E_WAIT)

    If IntervaloPermiteConectar Then
        If frmMain.Socket1.Connected Then
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
            DoEvents

        End If

        CuentaEmail = NameTxt.Text

        Dim aux As String

        aux = PasswordTxt.Text

        CuentaPassword = aux

        If Image4.Tag = "1" Then

            CuentaRecordada.nombre = CuentaEmail
            CuentaRecordada.Password = aux
            Call RecordarCuenta(CuentaEmail, aux)

        Else
            
            ' Reseteamos los datos de cuenta guardados
            Call RecordarCuenta(vbNullString, vbNullString)

        End If

        If CheckUserDataLoged() = True Then
            EstadoLogin = E_MODO.IngresandoConCuenta
            frmMain.Socket1.HostName = IPdelServidor
            frmMain.Socket1.RemotePort = PuertoDelServidor
            frmMain.Socket1.Connect

        End If

        ServerIndex = lstServers.ListIndex
        Call SaveRAOInit

    End If

    
    Exit Sub

Image3_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmLogear.Image3_Click", Erl)
    Resume Next
    
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image3_MouseMove_Err
    

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

    
    Exit Sub

Image3_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmLogear.Image3_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Image4_Click()
    
    On Error GoTo Image4_Click_Err
    

    If Image4.Tag = "0" Then
        Image4.Picture = LoadInterface("check-amarillo.bmp")
        Call TextoAlAsistente("¡Recordare la cuenta para la proxima!")
        Image4.Tag = "1"
    Else
        Image4.Picture = Nothing
        Image4.Tag = "0"
        Call TextoAlAsistente("¡No recordare nada!")

    End If

    
    Exit Sub

Image4_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmLogear.Image4_Click", Erl)
    Resume Next
    
End Sub

Private Sub Label1_Click()
    
    On Error GoTo Label1_Click_Err
    

    If Image4.Tag = "0" Then
        Image4.Picture = LoadInterface("check-amarillo.bmp")
        Call TextoAlAsistente("¡Recordaré la cuenta para la próxima!")
        Image4.Tag = "1"
    Else
        Image4.Picture = Nothing
        Image4.Tag = "0"
        Call TextoAlAsistente("¡No recordare nada!")

    End If

    
    Exit Sub

Label1_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmLogear.Label1_Click", Erl)
    Resume Next
    
End Sub

Private Sub lstServers_Click()
    
    On Error GoTo lstServers_Click_Err
    
    IPdelServidor = ServersLst(lstServers.ListIndex + 1).IP
    PuertoDelServidor = ServersLst(lstServers.ListIndex + 1).puerto
    
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
        Call Image3_Click

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
        Call Image3_Click

    End If

    
    Exit Sub

PasswordTxt_KeyDown_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmLogear.PasswordTxt_KeyDown", Erl)
    Resume Next
    
End Sub

Private Sub PasswordTxt_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo PasswordTxt_MouseMove_Err
    

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

    
    Exit Sub

PasswordTxt_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmLogear.PasswordTxt_MouseMove", Erl)
    Resume Next
    
End Sub
