VERSION 5.00
Begin VB.Form frmMensaje 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4365
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   1  'Blackness
   DrawStyle       =   5  'Transparent
   FontTransparent =   0   'False
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   291
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image imgCerrar 
      Height          =   420
      Left            =   3900
      Tag             =   "0"
      Top             =   15
      Width           =   465
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   1200
      Tag             =   "1"
      Top             =   2535
      Width           =   1980
   End
   Begin VB.Label msg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mensaje"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   1215
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Width           =   3165
      WordWrap        =   -1  'True
   End
   Begin VB.Menu mnuMensaje 
      Caption         =   "Mensaje"
      Visible         =   0   'False
      Begin VB.Menu mnuNormal 
         Caption         =   "Normal"
      End
      Begin VB.Menu mnuGlobal 
         Caption         =   "Global"
      End
      Begin VB.Menu mnuPrivado 
         Caption         =   "Privado"
      End
      Begin VB.Menu mnuGritar 
         Caption         =   "Gritar"
      End
      Begin VB.Menu mnuClan 
         Caption         =   "Clan"
      End
      Begin VB.Menu mnuGMs 
         Caption         =   "Grupo"
      End
   End
End
Attribute VB_Name = "frmMensaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bmoving      As Boolean

Public dX           As Integer

Public dy           As Integer

' Constantes para SendMessage
Const WM_SYSCOMMAND As Long = &H112&

Const MOUSE_MOVE    As Long = &HF012&

Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long

Private RealizoCambios As String

Const HWND_TOPMOST = -1

Const HWND_NOTOPMOST = -2

Const SWP_NOSIZE = &H1

Const SWP_NOMOVE = &H2

Const SWP_NOACTIVATE = &H10

Const SWP_SHOWWINDOW = &H40

Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
'Argentum Online 0.11.6

' función Api para aplicar la transparencia a la ventana
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

' Funciones api para los estilos de la ventana
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'constantes
Private Const GWL_EXSTYLE = (-20)

Private Const LWA_ALPHA = &H2

Private Const WS_EX_LAYERED = &H80000
  
' Función Api SetWindowPos
  
'En el primer parámetro se le pasa el Hwnd de la ventana
'El segundo es la constante que permite hacer el OnTop
'Los parámetros que están en 0 son las coordenadas, o sea la _
 pocición, obviamente opcionales
'El último parámetro es para que al establecer el OnTop la ventana _
 no se mueva de lugar y no se redimensione

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
    Call RegistrarError(Err.number, Err.Description, "frmMensaje.Is_Transparent", Erl)
    Resume Next
    
End Function

'Ladder 21/09/2012
'Cierra el form presionando enter.
Private Sub Form_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Form_KeyPress_Err
    

    If KeyAscii = vbKeyReturn Then
        Unload Me

    End If

    
    Exit Sub

Form_KeyPress_Err:
    Call RegistrarError(Err.number, Err.Description, "frmMensaje.Form_KeyPress", Erl)
    Resume Next
    
End Sub

'Ladder 21/09/2012

Private Sub Form_Load()
    'Call FormParser.Parse_Form(Me)
    
    On Error GoTo Form_Load_Err
    
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    'Call Aplicar_Transparencia(Me.hwnd, 200)
    ''Call Audio.PlayWave(SND_MSG)
    frmMensaje.Picture = LoadInterface(Language + "\mensaje.bmp")
    Me.Caption = "A"
    Call Form_RemoveTitleBar(Me)
    Me.Height = 3190
    Me.Width = 4380
    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "frmMensaje.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    Call MoverForm(Me.hWnd)

    If Image1.Tag = "1" Then
        Image1.Picture = Nothing
        Image1.Tag = "0"
    End If
    
    If imgCerrar.Tag = "1" Then
        imgCerrar.Picture = Nothing
        imgCerrar.Tag = "0"
    End If

    
    Exit Sub

Form_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmMensaje.Form_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Image1_Click()
    'Call Sound.Sound_Play(SND_CLICK)
    
    On Error GoTo Image1_Click_Err
    
    Unload Me
    
    Exit Sub

Image1_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmMensaje.Image1_Click", Erl)
    Resume Next
    
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image1_MouseDown_Err
    
    Image1.Picture = LoadInterface(Language + "\boton-aceptar-ES-off.bmp")
    Image1.Tag = "1"
    
    Exit Sub

Image1_MouseDown_Err:
    Call RegistrarError(Err.number, Err.Description, "frmMensaje.Image1_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image1_MouseMove_Err
    
    If Image1.Tag = "0" Then
        Image1.Picture = LoadInterface(Language + "\boton-aceptar-ES-over.bmp")
        Image1.Tag = "1"
    End If
    
    Exit Sub

Image1_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmMensaje.Image1_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Form_Deactivate()
    ' Me.SetFocus
End Sub

Public Sub PopupMenuMensaje()
    
    On Error GoTo PopupMenuMensaje_Err
    

    Select Case SendingType

        Case 1
            mnuNormal.Checked = True
            mnuGritar.Checked = False
            mnuPrivado.Checked = False
            mnuClan.Checked = False
            mnuGMs.Checked = False
            mnuGlobal.Checked = False

        Case 2
            mnuNormal.Checked = False
            mnuGritar.Checked = True
            mnuPrivado.Checked = False
            mnuClan.Checked = False
            mnuGMs.Checked = False
            mnuGlobal.Checked = False

        Case 3
            mnuNormal.Checked = False
            mnuGritar.Checked = False
            mnuPrivado.Checked = True
            mnuClan.Checked = False
            mnuGMs.Checked = False
            mnuGlobal.Checked = False

        Case 4
            mnuNormal.Checked = False
            mnuGritar.Checked = False
            mnuPrivado.Checked = False
            mnuClan.Checked = True
            mnuGMs.Checked = False
            mnuGlobal.Checked = False

        Case 5
            mnuNormal.Checked = False
            mnuGritar.Checked = False
            mnuPrivado.Checked = False
            mnuClan.Checked = False
            mnuGMs.Checked = True
            mnuGlobal.Checked = False

        Case 6
            mnuNormal.Checked = False
            mnuGritar.Checked = False
            mnuPrivado.Checked = False
            mnuClan.Checked = False
            mnuGMs.Checked = False
            mnuGlobal.Checked = False

        Case 7
            mnuNormal.Checked = False
            mnuGritar.Checked = False
            mnuPrivado.Checked = False
            mnuClan.Checked = False
            mnuGMs.Checked = False
            mnuGlobal.Checked = True

    End Select

    PopUpMenu mnuMensaje

    
    Exit Sub

PopupMenuMensaje_Err:
    Call RegistrarError(Err.number, Err.Description, "frmMensaje.PopupMenuMensaje", Erl)
    Resume Next
    
End Sub

Private Sub imgCerrar_Click()
    
    On Error GoTo imgCerrar_Click_Err
    
    Unload Me
    
    Exit Sub

imgCerrar_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmMensaje.imgCerrar_Click", Erl)
    Resume Next
    
End Sub

Private Sub imgCerrar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo imgCerrar_MouseDown_Err
    
    imgCerrar.Picture = LoadInterface(Language + "\boton-cerrar-off.bmp")
    imgCerrar.Tag = "1"
    
    Exit Sub

imgCerrar_MouseDown_Err:
    Call RegistrarError(Err.number, Err.Description, "frmMensaje.imgCerrar_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub imgCerrar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo imgCerrar_MouseMove_Err
    
    If imgCerrar.Tag = "0" Then
        imgCerrar.Picture = LoadInterface(Language + "\boton-cerrar-over.bmp")
        imgCerrar.Tag = "1"
    End If
    
    Exit Sub

imgCerrar_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmMensaje.imgCerrar_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub mnuNormal_Click()
    
    On Error GoTo mnuNormal_Click_Err
    
    SendingType = 1

    If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus

    
    Exit Sub

mnuNormal_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmMensaje.mnuNormal_Click", Erl)
    Resume Next
    
End Sub

Private Sub mnuGritar_click()
    
    On Error GoTo mnuGritar_click_Err
    
    SendingType = 2

    If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus

    
    Exit Sub

mnuGritar_click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmMensaje.mnuGritar_click", Erl)
    Resume Next
    
End Sub

Private Sub mnuPrivado_click()
    
    On Error GoTo mnuPrivado_click_Err
    
    sndPrivateTo = InputBox("Escriba el usuario con el que desea iniciar una conversación privada", "")

    If sndPrivateTo <> "" Then
        SendingType = 3

        If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus
    Else
        Call MensajeAdvertencia("Debes escribir un usuario válido")

    End If

    
    Exit Sub

mnuPrivado_click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmMensaje.mnuPrivado_click", Erl)
    Resume Next
    
End Sub

Private Sub mnuClan_click()
    
    On Error GoTo mnuClan_click_Err
    
    SendingType = 4

    If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus

    
    Exit Sub

mnuClan_click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmMensaje.mnuClan_click", Erl)
    Resume Next
    
End Sub

Private Sub mnuGMs_click()
    
    On Error GoTo mnuGMs_click_Err
    
    SendingType = 5

    If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus

    
    Exit Sub

mnuGMs_click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmMensaje.mnuGMs_click", Erl)
    Resume Next
    
End Sub

Private Sub mnuGrupo_click()
    
    On Error GoTo mnuGrupo_click_Err
    
    SendingType = 6

    If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus

    
    Exit Sub

mnuGrupo_click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmMensaje.mnuGrupo_click", Erl)
    Resume Next
    
End Sub

Private Sub mnuGlobal_Click()
    
    On Error GoTo mnuGlobal_Click_Err
    
    SendingType = 7

    If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus

    
    Exit Sub

mnuGlobal_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmMensaje.mnuGlobal_Click", Erl)
    Resume Next
    
End Sub

