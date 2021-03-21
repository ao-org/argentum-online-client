VERSION 5.00
Begin VB.Form FrmBorrarCuenta 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4590
   ClientLeft      =   11835
   ClientTop       =   -150
   ClientWidth     =   3525
   ClipControls    =   0   'False
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   3525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   480
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1500
      Width           =   2535
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
      PasswordChar    =   "•"
      TabIndex        =   1
      Top             =   2130
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
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   480
      MaxLength       =   5
      TabIndex        =   2
      Top             =   3520
      Width           =   780
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   3480
      Width           =   735
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
      Top             =   3480
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   435
      Left            =   320
      Tag             =   "0"
      Top             =   4040
      Width           =   1080
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   1950
      Tag             =   "0"
      Top             =   4050
      Width           =   1035
   End
End
Attribute VB_Name = "FrmBorrarCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim ValidacionNumber As Long

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call FormParser.Parse_Form(Me)
    Call Aplicar_Transparencia(Me.hWnd, 240)
    ValidacionNumber = RandomNumber(10000, 90000)
    valcar = ValidacionNumber
    Me.Picture = LoadInterface(Language + "/borrarcuenta.bmp")

    Call MensajeAdvertencia("Use esta opcion con responsabilidad, una vez borrada la cuenta no se podra volver a recuperar.")
    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmBorrarCuenta.Form_Load", Erl)
    Resume Next
    
End Sub

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
    Call RegistrarError(Err.number, Err.Description, "FrmBorrarCuenta.Form_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Image1_Click()
    
    On Error GoTo Image1_Click_Err
    
    Unload Me
    frmMasOpciones.Show , frmConnect
    frmMasOpciones.Top = frmMasOpciones.Top + 3000

    
    Exit Sub

Image1_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmBorrarCuenta.Image1_Click", Erl)
    Resume Next
    
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    'Image1.Picture = LoadInterface(Language + "volverpress.bmp")
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image1_MouseMove_Err
    

    If Image1.Tag = "0" Then
        Image1.Picture = LoadInterface(Language + "/borrarcuenta_volverhover.bmp")
        Image1.Tag = "1"

    End If

    
    Exit Sub

Image1_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmBorrarCuenta.Image1_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Image2_Click()
    
    On Error GoTo Image2_Click_Err
    

    If Constraseña = "" Then
        Call MensajeAdvertencia("El campo de constraseña esta vacia.")
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
        Call MensajeAdvertencia("El codigo de verificación es invalido, por favor reintente.")
        Exit Sub

    End If
    
    If MsgBox("Esta a punto de borrar la cuenta y todos los personajes que contiene la misma, esta acción es irreversible. ¿Esta seguro?", vbYesNo + vbQuestion, "¡ATENCION!") = vbYes Then
        EstadoLogin = E_MODO.BorrandoCuenta
        CuentaEmail = Email
        CuentaPassword = Constraseña
               
        If frmMain.Socket1.Connected Then
            frmMain.Socket1.Disconnect
            frmMain.Socket1.Cleanup
            DoEvents

        End If

        frmMain.Socket1.HostName = IPdelServidor
        frmMain.Socket1.RemotePort = PuertoDelServidor
        frmMain.Socket1.Connect
        Unload Me

    End If

    
    Exit Sub

Image2_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmBorrarCuenta.Image2_Click", Erl)
    Resume Next
    
End Sub

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    'Image2.Picture = LoadInterface(Language + "borrarpress.bmp")
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image2_MouseMove_Err
    

    If Image2.Tag = "0" Then
        Image2.Picture = LoadInterface(Language + "\borrarhover.bmp")
        Image2.Tag = "1"

    End If

    
    Exit Sub

Image2_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmBorrarCuenta.Image2_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Label1_Click()
    
    On Error GoTo Label1_Click_Err
    
    texVer.SetFocus

    
    Exit Sub

Label1_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmBorrarCuenta.Label1_Click", Erl)
    Resume Next
    
End Sub

Private Sub valcar_Click()
    
    On Error GoTo valcar_Click_Err
    
    texVer.SetFocus

    
    Exit Sub

valcar_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmBorrarCuenta.valcar_Click", Erl)
    Resume Next
    
End Sub
