VERSION 5.00
Begin VB.Form frmOpciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   7350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8040
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   8040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.HScrollBar scrSens 
      Height          =   315
      LargeChange     =   5
      Left            =   4650
      Max             =   20
      Min             =   1
      TabIndex        =   14
      Top             =   5040
      Value           =   10
      Width           =   2415
   End
   Begin VB.HScrollBar Alpha 
      Height          =   315
      LargeChange     =   60
      Left            =   4680
      Max             =   255
      SmallChange     =   2
      TabIndex        =   13
      Top             =   3000
      Value           =   120
      Width           =   2775
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   315
      LargeChange     =   1000
      Left            =   720
      Max             =   0
      Min             =   -4000
      SmallChange     =   2
      TabIndex        =   12
      Top             =   5400
      Width           =   2775
   End
   Begin VB.HScrollBar scrMidi 
      Height          =   315
      LargeChange     =   1000
      Left            =   720
      Max             =   0
      Min             =   -4000
      SmallChange     =   2
      TabIndex        =   11
      Top             =   4500
      Width           =   2775
   End
   Begin VB.HScrollBar scrVolume 
      Height          =   315
      LargeChange     =   1000
      Left            =   720
      Max             =   0
      Min             =   -4000
      SmallChange     =   2
      TabIndex        =   10
      Top             =   3625
      Width           =   2775
   End
   Begin VB.CheckBox Macro 
      Caption         =   "Arriba"
      Height          =   255
      Index           =   0
      Left            =   11400
      TabIndex        =   9
      Top             =   10680
      Width           =   855
   End
   Begin VB.CheckBox Macro 
      Caption         =   "Abajo"
      Height          =   255
      Index           =   1
      Left            =   12600
      TabIndex        =   8
      Top             =   10680
      Width           =   855
   End
   Begin VB.CheckBox Macro 
      Caption         =   "Izquierda"
      Height          =   255
      Index           =   2
      Left            =   11400
      TabIndex        =   7
      Top             =   10920
      Width           =   1095
   End
   Begin VB.CheckBox Macro 
      Caption         =   "Derecha"
      Height          =   255
      Index           =   3
      Left            =   12600
      TabIndex        =   6
      Top             =   10920
      Width           =   1095
   End
   Begin VB.CheckBox Check7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver nombres del mapa"
      Height          =   255
      Left            =   11160
      TabIndex        =   5
      Tag             =   "13"
      Top             =   9600
      Width           =   1935
   End
   Begin VB.CheckBox Check11 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Desvanecimiento de techos"
      Height          =   255
      Left            =   11280
      TabIndex        =   4
      Top             =   9120
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   195
      Left            =   7560
      TabIndex        =   3
      Top             =   10920
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   9600
      Top             =   10080
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Bajar Volumen"
      Height          =   375
      Left            =   7440
      TabIndex        =   1
      Top             =   9960
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Subir Volumen"
      Height          =   375
      Left            =   7440
      TabIndex        =   0
      Top             =   10440
      Width           =   1335
   End
   Begin VB.Image facebook 
      Height          =   390
      Left            =   3300
      Tag             =   "0"
      Top             =   5800
      Width           =   435
   End
   Begin VB.Image instagram 
      Height          =   390
      Left            =   2850
      Tag             =   "0"
      Top             =   5790
      Width           =   435
   End
   Begin VB.Image discord 
      Height          =   390
      Left            =   1890
      Tag             =   "0"
      Top             =   5805
      Width           =   435
   End
   Begin VB.Image cmdChangePassword 
      Height          =   480
      Left            =   5050
      Tag             =   "0"
      Top             =   6550
      Width           =   2760
   End
   Begin VB.Image Command1 
      Height          =   525
      Left            =   370
      Tag             =   "0"
      Top             =   6540
      Width           =   2790
   End
   Begin VB.Image cmdcerrar 
      Height          =   480
      Left            =   3280
      Tag             =   "0"
      Top             =   6800
      Width           =   1755
   End
   Begin VB.Image cmdweb 
      Height          =   390
      Left            =   2370
      Tag             =   "0"
      Top             =   5800
      Width           =   435
   End
   Begin VB.Image cmdayuda 
      Height          =   435
      Left            =   400
      Tag             =   "0"
      Top             =   5790
      Width           =   1395
   End
   Begin VB.Label txtMSens 
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7200
      TabIndex        =   15
      Top             =   5070
      Width           =   375
   End
   Begin VB.Image Check4 
      Height          =   210
      Left            =   4540
      Top             =   4150
      Width           =   180
   End
   Begin VB.Image Check9 
      Height          =   210
      Left            =   4540
      Top             =   3870
      Width           =   180
   End
   Begin VB.Image check1 
      Height          =   210
      Left            =   4560
      Top             =   2190
      Width           =   180
   End
   Begin VB.Image Check5 
      Height          =   210
      Left            =   4560
      Top             =   1900
      Width           =   180
   End
   Begin VB.Image Check6 
      Height          =   210
      Left            =   4560
      Top             =   1600
      Width           =   180
   End
   Begin VB.Image Check2 
      Height          =   210
      Left            =   4420
      Top             =   5840
      Width           =   180
   End
   Begin VB.Image Check3 
      Height          =   210
      Left            =   4420
      Top             =   5520
      Width           =   180
   End
   Begin VB.Image chkInvertir 
      Height          =   210
      Left            =   680
      Top             =   2840
      Width           =   180
   End
   Begin VB.Image chko 
      Height          =   210
      Index           =   2
      Left            =   680
      Top             =   2540
      Width           =   180
   End
   Begin VB.Image chko 
      Height          =   210
      Index           =   3
      Left            =   680
      Top             =   2240
      Width           =   180
   End
   Begin VB.Image chko 
      Height          =   210
      Index           =   1
      Left            =   680
      Top             =   1940
      Width           =   180
   End
   Begin VB.Image chko 
      Height          =   210
      Index           =   0
      Left            =   680
      Top             =   1630
      Width           =   180
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   9000
      TabIndex        =   2
      Top             =   10680
      Width           =   1455
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Argentum Online 0.11.6
'
Option Explicit
Private Bajar As Boolean
Private Subir As Boolean
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

' función Api para aplicar la transparencia a la ventana
Private Declare Function SetLayeredWindowAttributes Lib "user32" _
    (ByVal hwnd As Long, _
     ByVal crKey As Long, _
     ByVal bAlpha As Byte, _
     ByVal dwFlags As Long) As Long

' Funciones api para los estilos de la ventana
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
    (ByVal hwnd As Long, _
     ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, _
     ByVal nIndex As Long, _
     ByVal dwNewLong As Long) As Long

'constantes
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000


Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
  ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
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

Private Sub Alpha_Change()
AlphaMacro = Alpha.value
End Sub

Private Sub Check1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        If OcultarMacrosAlCastear = 1 Then
            OcultarMacrosAlCastear = 0
        Else
            OcultarMacrosAlCastear = 1
        End If
        
    If OcultarMacrosAlCastear = 0 Then
        check1.Picture = Nothing
    Else
        check1.Picture = LoadInterface("config_stick.bmp")
    End If
        
End Sub

Private Sub Check4_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If PermitirMoverse = 1 Then
    PermitirMoverse = 0
Else
    PermitirMoverse = 1
End If

    If PermitirMoverse = 0 Then
        Check4.Picture = Nothing
    Else
        Check4.Picture = LoadInterface("config_stick.bmp")
    End If
End Sub

Private Sub Check5_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If MoverVentana = 1 Then
    MoverVentana = 0
Else
    MoverVentana = 1
End If

    If MoverVentana = 0 Then
        Check5.Picture = Nothing
    Else
        Check5.Picture = LoadInterface("config_stick.bmp")
    End If
End Sub


Private Sub Check2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If CursoresGraficos = 1 Then
    Call WriteVar(App.Path & "\..\Recursos\OUTPUT\" & "raoinit.ini", "OPCIONES", "CursoresGraficos", 0)
    MsgBox "Para que los cambios en esta opción sean reflejados, deberá reiniciar el cliente.", vbQuestion, "Argentum20 - Advertencia" 'hay que poner 20 aniversario
Else
    CursoresGraficos = 1
    Call FormParser.Parse_Form(Me)
    Call WriteVar(App.Path & "\..\Recursos\OUTPUT\" & "raoinit.ini", "OPCIONES", "CursoresGraficos", 1)
    
End If


    If CursoresGraficos = 0 Then
        Check2.Picture = Nothing
    Else
        Check2.Picture = LoadInterface("config_stick.bmp")
    End If
End Sub


Private Sub chkInvertir_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        If InvertirSonido = 1 Then
            InvertirSonido = 0
            
            Sound.InvertirSonido = False
        Else
            InvertirSonido = 1
            Sound.InvertirSonido = True
        End If
        
        If InvertirSonido = 0 Then
            chkInvertir.Picture = Nothing
        Else
            chkInvertir.Picture = LoadInterface("config_stick.bmp")
        End If
End Sub

Private Sub chkInvertir2_Click()

End Sub

Private Sub chkO_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Call Sound.Sound_Play(SND_CLICK)

Select Case Index
    Case 0
       If Musica <> CONST_DESHABILITADA Then
            Sound.Music_Stop
            Musica = CONST_DESHABILITADA
            scrMidi.Enabled = False
        Else
            Musica = CONST_MP3
          scrMidi.Enabled = True
           Sound.NextMusic = MapDat.music_numberHi
             Sound.Fading = 100
        End If
        If Musica = 0 Then
            chko(0).Picture = Nothing
        Else
            chko(0).Picture = LoadInterface("config_stick.bmp")
        End If
    Case 1
        If fX = 1 Then
            fX = 0
            chko(2).Enabled = False
            scrVolume.Enabled = False
            
            Call Sound.Sound_Stop_All
        Else
            fX = 1
            chko(2).Enabled = True
            scrVolume.Enabled = True
        End If
        
        If fX = 0 Then
            chko(1).Picture = Nothing
        Else
            chko(1).Picture = LoadInterface("config_stick.bmp")
        End If
    Case 2
        If FxNavega = 1 Then
            FxNavega = 0
        Else
            FxNavega = 1
        End If
        
        If FxNavega = 0 Then
            chko(2).Picture = Nothing
        Else
            chko(2).Picture = LoadInterface("config_stick.bmp")
        End If
    Case 3
        If AmbientalActivated = 1 Then
                HScroll1.Enabled = False
                AmbientalActivated = 0
                Sound.LastAmbienteActual = 0
                Sound.AmbienteActual = 0
                Sound.Ambient_Stop
        Else
                HScroll1.Enabled = True
                AmbientalActivated = 1
                Call AmbientarAudio(UserMap)
        End If
        If AmbientalActivated = 0 Then
            chko(3).Picture = Nothing
        Else
            chko(3).Picture = LoadInterface("config_stick.bmp")
        End If
End Select

End Sub

Private Sub cmdayuda_Click()
Call FrmGmAyuda.Show
End Sub

Private Sub cmdayuda_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If cmdayuda.Tag = "0" Then
        cmdayuda.Picture = LoadInterface("config_ayuda.bmp")
        cmdayuda.Tag = "1"
    End If
End Sub
Private Sub discord_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If discord.Tag = "0" Then
        discord.Picture = LoadInterface("config_discord.bmp")
        discord.Tag = "1"
    End If
    
cmdweb = Nothing
cmdweb.Tag = "0"
instagram = Nothing
instagram.Tag = "0"
facebook = Nothing
facebook.Tag = "0"
End Sub
Private Sub cmdweb_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If cmdweb.Tag = "0" Then
        cmdweb.Picture = LoadInterface("config_web.bmp")
        cmdweb.Tag = "1"
    End If
    
    discord = Nothing
discord.Tag = "0"
instagram = Nothing
instagram.Tag = "0"
facebook = Nothing
facebook.Tag = "0"
End Sub

Private Sub instagram_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If instagram.Tag = "0" Then
        instagram.Picture = LoadInterface("config_instagram.bmp")
        instagram.Tag = "1"
    End If
    
discord = Nothing
discord.Tag = "0"
cmdweb = Nothing
cmdweb.Tag = "0"
facebook = Nothing
facebook.Tag = "0"
End Sub
Private Sub facebook_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If facebook.Tag = "0" Then
        facebook.Picture = LoadInterface("config_facebook.bmp")
        facebook.Tag = "1"
    End If

discord = Nothing
discord.Tag = "0"
cmdweb = Nothing
cmdweb.Tag = "0"
instagram = Nothing
instagram.Tag = "0"
End Sub
Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Command1.Tag = "0" Then
        Command1.Picture = LoadInterface("config_teclas.bmp")
        Command1.Tag = "1"
    End If
        cmdcerrar = Nothing
cmdcerrar.Tag = "0"
    
End Sub
Private Sub cmdcerrar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If cmdcerrar.Tag = "0" Then
        cmdcerrar.Picture = LoadInterface("config_cerrar.bmp")
        cmdcerrar.Tag = "1"
    End If
    cmdChangePassword = Nothing
cmdChangePassword.Tag = "0"
Command1 = Nothing
Command1.Tag = "0"
End Sub
Private Sub cmdChangePassword_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If cmdChangePassword.Tag = "0" Then
        cmdChangePassword.Picture = LoadInterface("config_contraseñ.bmp")
        cmdChangePassword.Tag = "1"
    End If
    cmdcerrar = Nothing
cmdcerrar.Tag = "0"
End Sub

Private Sub cmdWeb_Click()
ShellExecute Me.hwnd, "open", "https://www.argentum20.com/", "", "", 0
End Sub

Private Sub Command5_Click()
MsgBox ("Proximamente")
End Sub

Private Sub discord_Click()
ShellExecute Me.hwnd, "open", "https://discord.gg/e3juVbF", "", "", 0
End Sub

Private Sub facebook_Click()
ShellExecute Me.hwnd, "open", "https://www.argentum20.com/", "", "", 0
End Sub

Private Sub Form_Load()
    Call Aplicar_Transparencia(Me.hwnd, 240)
    Call FormParser.Parse_Form(Me)
    Me.Picture = LoadInterface("config.bmp")
    
End Sub
Private Sub moverForm()
    Dim res As Long
    ReleaseCapture
    res = SendMessage(Me.hwnd, WM_SYSCOMMAND, MOUSE_MOVE, 0)
End Sub

Private Sub Check3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'    If Check3 Then
    '    SwapMouseButton 1
   ' Else
     '   SwapMouseButton 0
 '   End If
End Sub
Private Sub Check6_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        If FPSFLAG = 1 Then
            FPSFLAG = 0
        Else
            FPSFLAG = 1
        End If
        
        
    If FPSFLAG = 0 Then
        Check6.Picture = Nothing
    Else
        Check6.Picture = LoadInterface("config_stick.bmp")
    End If
End Sub
Private Sub Check9_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        If CopiarDialogoAConsola = 1 Then
            CopiarDialogoAConsola = 0
        Else
            CopiarDialogoAConsola = 1
        End If
        
    If CopiarDialogoAConsola = 0 Then
        Check9.Picture = Nothing
    Else
        Check9.Picture = LoadInterface("config_stick.bmp")
    End If
End Sub

Private Sub Command2_Click()
Bajar = True
Subir = False
Timer1.Enabled = True
End Sub

Private Sub Command3_Click()
Subir = True
Bajar = False
Timer1.Enabled = True
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
moverForm
cmdayuda = Nothing
cmdayuda.Tag = "0"
discord = Nothing
discord.Tag = "0"
cmdweb = Nothing
cmdweb.Tag = "0"
instagram = Nothing
instagram.Tag = "0"
facebook = Nothing
facebook.Tag = "0"
Command1 = Nothing
Command1.Tag = "0"
cmdcerrar = Nothing
cmdcerrar.Tag = "0"
cmdChangePassword = Nothing
cmdChangePassword.Tag = "0"
End Sub
Private Sub cmdcerrar_Click()
    Call GuardarOpciones
    Me.Visible = False
    frmmain.SetFocus
End Sub
Private Sub cmdChangePassword_Click()
    Call frmNewPassword.Show(vbModeless, Me)
End Sub
Private Sub Command1_Click()
    Call frmCustomKeys.Show(vbModeless, Me)
End Sub
Public Sub Init()


    
    If CopiarDialogoAConsola = 0 Then
        Check9.Picture = Nothing
    Else
        Check9.Picture = LoadInterface("config_stick.bmp")
    End If
    
    
    If MoverVentana = 0 Then
        Check5.Picture = Nothing
    Else
        Check5.Picture = LoadInterface("config_stick.bmp")
    End If
    

    If CursoresGraficos = 0 Then
        Check2.Picture = Nothing
    Else
        Check2.Picture = LoadInterface("config_stick.bmp")
    End If



    If PermitirMoverse = 0 Then
        Check4.Picture = Nothing
    Else
        Check4.Picture = LoadInterface("config_stick.bmp")
    End If

    
    If Musica = 0 Then
        chko(0).Picture = Nothing
    Else
        chko(0).Picture = LoadInterface("config_stick.bmp")
    End If
    
    If FxNavega = 0 Then
        chko(2).Picture = Nothing
    Else
        chko(2).Picture = LoadInterface("config_stick.bmp")
    End If
    
    
    If AmbientalActivated = 0 Then
        chko(3).Picture = Nothing
    Else
        chko(3).Picture = LoadInterface("config_stick.bmp")
    End If
    
    If fX = 0 Then
        chko(1).Picture = Nothing
    Else
        chko(1).Picture = LoadInterface("config_stick.bmp")
    End If
    
    
    If InvertirSonido = 0 Then
        chkInvertir.Picture = Nothing
    Else
        chkInvertir.Picture = LoadInterface("config_stick.bmp")
    End If
    
    
    
    If FPSFLAG = 0 Then
        Check6.Picture = Nothing
    Else
        Check6.Picture = LoadInterface("config_stick.bmp")
    End If
    
    
    
    If OcultarMacrosAlCastear = 0 Then
        check1.Picture = Nothing
    Else
        check1.Picture = LoadInterface("config_stick.bmp")
    End If

    
        

    
    scrVolume.value = VolFX
    HScroll1.value = VolAmbient
    scrMidi.value = VolMusic
    
    
    Alpha.value = AlphaMacro
    
    
    scrSens.value = SensibilidadMouse


    
Me.Show vbModeless, frmmain

End Sub

Private Sub HScroll1_Change()
Sound.Ambient_Volume_Set HScroll1.value
VolAmbient = HScroll1.value
End Sub


Private Sub instagram_Click()
ShellExecute Me.hwnd, "open", "https://www.argentum20.com/", "", "", 0
End Sub


Private Sub scrMidi_Change()
Sound.Music_Volume_Set scrMidi.value
Sound.VolumenActualMusicMax = scrMidi.value
VolMusic = Sound.VolumenActualMusicMax
End Sub

Private Sub scrSens_Change()
MouseS = scrSens.value
SensibilidadMouse = MouseS
Call General_Set_Mouse_Speed(MouseS)
txtMSens.Caption = scrSens.value
End Sub

Private Sub scrVolume_Change()
Sound.VolumenActual = scrVolume.value
VolFX = Sound.VolumenActual

End Sub


