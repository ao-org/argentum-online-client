VERSION 5.00
Begin VB.Form frmMasOpciones 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   4575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3480
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
   ScaleHeight     =   4575
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   405
      Index           =   5
      Left            =   1200
      Tag             =   "0"
      Top             =   3990
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   405
      Index           =   4
      Left            =   840
      Tag             =   "0"
      Top             =   3450
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   405
      Index           =   3
      Left            =   420
      Tag             =   "0"
      Top             =   2890
      Width           =   2610
   End
   Begin VB.Image Image1 
      Height          =   405
      Index           =   2
      Left            =   750
      Tag             =   "0"
      Top             =   2360
      Width           =   1995
   End
   Begin VB.Image Image1 
      Height          =   420
      Index           =   1
      Left            =   810
      Tag             =   "0"
      Top             =   1800
      Width           =   1890
   End
   Begin VB.Image Image1 
      Height          =   450
      Index           =   0
      Left            =   930
      Tag             =   "0"
      Top             =   1230
      Width           =   1845
   End
End
Attribute VB_Name = "frmMasOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const SWP_NOMOVE = 2

Const SWP_NOSIZE = 1

Const HWND_TOPMOST = -1

Const HWND_NOTOPMOST = -2 '
  
' Funci�n Api SetWindowPos
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Declaraci�n del Api SetLayeredWindowAttributes que establece _
 la transparencia al form
  
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
  
'Recupera el estilo de la ventana
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
  
'Declaraci�n del Api SetWindowLong necesaria para aplicar un estilo _
 al form antes de usar el Api SetLayeredWindowAttributes
  
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
  
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

Private Sub Form_Activate()
    'SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, _
     SWP_NOMOVE Or SWP_NOSIZE

End Sub

Private Sub Form_Load()
    Call FormParser.Parse_Form(Me)
    Me.Picture = LoadInterface("opcioneslogeo.bmp")
    Call Aplicar_Transparencia(Me.hwnd, 240)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Image1(0).Picture = Nothing

    Image1(1).Picture = Nothing
    Image1(2).Picture = Nothing
    Image1(3).Picture = Nothing
    Image1(4).Picture = Nothing
    Image1(5).Picture = Nothing

    Image1(0).Tag = "0"
    Image1(1).Tag = "0"
    Image1(2).Tag = "0"
    Image1(3).Tag = "0"
    Image1(4).Tag = "0"
    Image1(5).Tag = "0"

End Sub

Private Sub Image1_Click(Index As Integer)
    Call Sound.Sound_Play(SND_CLICK)
    Unload Me

    Select Case Index

        Case 0
            Call Sound.Sound_Play(SND_CLICK)
            Unload Me

            If frmmain.Socket1.Connected Then
                frmmain.Socket1.Disconnect
                frmmain.Socket1.Cleanup
                DoEvents

            End If

            If frmCrearCuenta.Visible = False Then
                Unload frmCrearCuenta

            End If

            frmCrearCuenta.Show , frmConnect
            frmCrearCuenta.Top = frmCrearCuenta.Top + 3000
            
        Case 1
            FrmActivarCuenta.Show , frmConnect
            FrmActivarCuenta.Top = FrmActivarCuenta.Top + 3000

        Case 2
            FrmReenviarMail.Show , frmConnect
            FrmReenviarMail.Top = FrmReenviarMail.Top + 3000

        Case 4
            FrmBorrarCuenta.Show , frmConnect
            FrmBorrarCuenta.Top = FrmBorrarCuenta.Top + 3000

        Case 3
            FrmRecuperar.Show , frmConnect
            FrmRecuperar.Top = FrmRecuperar.Top + 3000

        Case 5
            Unload Me
            FrmLogear.Visible = True
            FrmLogear.Top = FrmLogear.Top + 4000

    End Select

End Sub

Private Sub Image1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    Select Case Index

        Case 0

            If Image1(Index).Tag = "0" Then
                Image1(Index).Picture = LoadInterface("crearcuentawidehover.bmp")
                Image1(Index).Tag = "1"

            End If

        Case 1

            If Image1(Index).Tag = "0" Then
                Image1(Index).Picture = LoadInterface("validarcuentahover.bmp")
                Image1(Index).Tag = "1"

            End If

        Case 2

            If Image1(Index).Tag = "0" Then
                Image1(Index).Picture = LoadInterface("reenviarvalidaci�nhover.bmp")
                Image1(Index).Tag = "1"

            End If
            
        Case 3

            If Image1(Index).Tag = "0" Then
                Image1(Index).Picture = LoadInterface("recuperarcontrase�ahover.bmp")
                Image1(Index).Tag = "1"

            End If
            
        Case 4

            If Image1(Index).Tag = "0" Then
                Image1(Index).Picture = LoadInterface("borrarcuentawidehover.bmp")
                Image1(Index).Tag = "1"

            End If

        Case 5

            If Image1(Index).Tag = "0" Then
                Image1(Index).Picture = LoadInterface("volverwidehover.bmp")
                Image1(Index).Tag = "1"

            End If

    End Select

End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Exit Sub

    Select Case Index

        Case 0
            Image1(Index).Picture = LoadInterface("crearcuentawidepress.bmp")
            Image1(Index).Tag = "1"

        Case 1
            Image1(Index).Picture = LoadInterface("validarcuenta.bmp")
            Image1(Index).Tag = "1"

        Case 2
            Image1(Index).Picture = LoadInterface("reenviarvalidaci�npress.bmp")
            Image1(Index).Tag = "1"

        Case 3
            Image1(Index).Picture = LoadInterface("recuperarcontrase�apress.bmp")
            Image1(Index).Tag = "1"

        Case 4
            Image1(Index).Picture = LoadInterface("borrarcuentapress.bmp")
            Image1(Index).Tag = "1"

        Case 5
            Image1(Index).Picture = LoadInterface("volverwidepress.bmp")
            Image1(Index).Tag = "1"

    End Select

End Sub

Private Sub Image4_Click(Index As Integer)

End Sub
