VERSION 5.00
Begin VB.Form FrmViajes 
   BorderStyle     =   0  'None
   Caption         =   "Formulario de Viaje"
   ClientHeight    =   4425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4485
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
   ScaleHeight     =   4425
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1920
      Left            =   720
      TabIndex        =   0
      Top             =   1080
      Width           =   3090
   End
   Begin VB.Image cmdViajar 
      Height          =   495
      Left            =   1250
      Top             =   3695
      Width           =   2055
   End
   Begin VB.Image cmdCerrar 
      Height          =   375
      Left            =   4000
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "FrmViajes"
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
Private cBtnClose As clsGraphicalButton
Private cBtnTravel As clsGraphicalButton
Private Sub loadButtons()
    Set cBtnClose = New clsGraphicalButton
    Set cBtnTravel = New clsGraphicalButton
    Call cBtnClose.Initialize(cmdCerrar, "boton-cerrar-default.bmp", "boton-cerrar-over.bmp", "boton-cerrar-off.bmp", Me)
    Call cBtnTravel.Initialize(cmdViajar, "boton-viajar-default.bmp", "boton-viajar-over.bmp", "boton-viajar-off.bmp", Me)
End Sub
Private Sub cmdCerrar_Click()
    Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo Form_KeyPress_Err
    If (KeyAscii = 27) Then
        Unload Me
    End If
    Exit Sub
Form_KeyPress_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmViajes.Form_KeyPress", Erl)
    Resume Next
End Sub

Private Sub Form_Load()
    On Error GoTo Form_Load_Err
    Call FormParser.Parse_Form(Me)
    Me.Picture = LoadInterface("ventanaviajes.bmp")
    Call Aplicar_Transparencia(Me.hWnd, 240)
    Call loadButtons
    Exit Sub
Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmViajes.Form_Load", Erl)
    Resume Next
End Sub

Private Sub cmdViajar_Click()
    On Error GoTo cmdViajar_Click_Err
    Dim destino As Byte
    destino = List1.ListIndex + 1
    If destino <= 0 Then Exit Sub
    Unload Me
    Call WriteCompletarViaje(destino, Destinos(destino).costo)
    Exit Sub
cmdViajar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmViajes.cmdViajar_Click", Erl)
    Resume Next
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Form_MouseMove_Err
    Call MoverForm(Me.hWnd)
    Exit Sub
Form_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmViajes.Form_MouseMove", Erl)
    Resume Next
End Sub
