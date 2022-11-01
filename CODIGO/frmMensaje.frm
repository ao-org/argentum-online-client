VERSION 5.00
Begin VB.Form frmMensaje 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3204
   ClientLeft      =   12
   ClientTop       =   12
   ClientWidth     =   4368
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawMode        =   1  'Blackness
   DrawStyle       =   5  'Transparent
   FontTransparent =   0   'False
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   267
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   364
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image cmdCerrar 
      Height          =   420
      Left            =   3900
      Tag             =   "0"
      Top             =   0
      Width           =   465
   End
   Begin VB.Image cmdAceptar 
      Height          =   420
      Left            =   1200
      Tag             =   "1"
      Top             =   2520
      Width           =   1980
   End
   Begin VB.Label msg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mensaje"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.6
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
'    Argentum 20 - Game Client Program
'    Copyright (C) 2022 - Noland Studios
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'
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

Private cBotonCerrar As clsGraphicalButton
Private cBotonAceptar As clsGraphicalButton
  
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

Private Sub cmdAceptar_Click()
    Unload Me
End Sub

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

Private Sub Form_Load()

    On Error GoTo Form_Load_Err
    If (Not FormParser Is Nothing) Then
        Call FormParser.Parse_Form(Me)
    End If
    
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Me.Picture = LoadInterface("mensaje.bmp")
    Me.Caption = "A"
    Call Form_RemoveTitleBar(Me)
    Me.Height = 3190
    Me.Width = 4380
    
    Call LoadButtons
    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "frmMensaje.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub LoadButtons()
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonAceptar = New clsGraphicalButton
    
    Call cBotonCerrar.Initialize(cmdCerrar, "boton-cerrar-default.bmp", "boton-cerrar-Over.bmp", "boton-cerrar-off.bmp", Me)
    Call cBotonAceptar.Initialize(cmdAceptar, "boton-Aceptar-default.bmp", "boton-Aceptar-Over.bmp", "boton-Aceptar-off.bmp", Me)
End Sub

Private Sub cmdCerrar_Click()
    On Error GoTo cmdCerrar_Click_Err
    
    Unload Me
    
    Exit Sub

cmdCerrar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMensaje.cmdCerrar_Click", Erl)
    Resume Next
    
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

Private Sub mnuGritar_click()
    
    On Error GoTo mnuGritar_click_Err
    
    SendingType = 2

    If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus
    If frmMain.SendTxtCmsg.Visible Then frmMain.SendTxtCmsg.SetFocus
    
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
        If frmMain.SendTxtCmsg.Visible Then frmMain.SendTxtCmsg.SetFocus
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

    If frmMain.SendTxtCmsg.Visible Then frmMain.SendTxtCmsg.SetFocus
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
    If frmMain.SendTxtCmsg.Visible Then frmMain.SendTxtCmsg.SetFocus
    
    Exit Sub

mnuGMs_click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmMensaje.mnuGMs_click", Erl)
    Resume Next
    
End Sub

Private Sub mnuGrupo_click()
    
    On Error GoTo mnuGrupo_click_Err
    
    SendingType = 6

    If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus
    If frmMain.SendTxtCmsg.Visible Then frmMain.SendTxtCmsg.SetFocus
    
    Exit Sub

mnuGrupo_click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmMensaje.mnuGrupo_click", Erl)
    Resume Next
    
End Sub

Private Sub mnuGlobal_Click()
    
    On Error GoTo mnuGlobal_Click_Err
    
    SendingType = 7

    If frmMain.SendTxt.Visible Then frmMain.SendTxt.SetFocus
    If frmMain.SendTxtCmsg.Visible Then frmMain.SendTxtCmsg.SetFocus
    
    Exit Sub

mnuGlobal_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmMensaje.mnuGlobal_Click", Erl)
    Resume Next
    
End Sub
