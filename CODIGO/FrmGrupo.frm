VERSION 5.00
Begin VB.Form FrmGrupo 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Grupo"
   ClientHeight    =   6525
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   6510
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
   ScaleHeight     =   435
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   434
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstGrupo 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2130
      Left            =   3600
      TabIndex        =   0
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Image cmdAceptar 
      Height          =   420
      Left            =   3525
      Tag             =   "0"
      Top             =   5730
      Width           =   1980
   End
   Begin VB.Image cmdAbandonar 
      Height          =   420
      Left            =   1005
      Tag             =   "0"
      Top             =   5730
      Width           =   1980
   End
   Begin VB.Image cmdCerrar 
      Height          =   420
      Left            =   6030
      Tag             =   "0"
      Top             =   15
      Width           =   465
   End
   Begin VB.Image cmdExpulsar 
      Height          =   420
      Left            =   4185
      Tag             =   "0"
      Top             =   4665
      Width           =   465
   End
   Begin VB.Image cmdInvitar 
      Height          =   420
      Left            =   4815
      Tag             =   "0"
      Top             =   4665
      Width           =   465
   End
End
Attribute VB_Name = "FrmGrupo"
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

Private cBotonAceptar As clsGraphicalButton
Private cBotonAbandonar As clsGraphicalButton
Private cBotonExpulsar As clsGraphicalButton
Private cBotonInvitar As clsGraphicalButton
Private cBotonCerrar As clsGraphicalButton

Private Sub cmdAceptar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    'Call FormParser.Parse_Form(Me)
    
    Call Aplicar_Transparencia(Me.hwnd, 240)
    
    Me.Picture = LoadInterface("ventanagrupo.bmp")
    
    Call loadButtons
    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmGrupo.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub loadButtons()
    Set cBotonAceptar = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonAbandonar = New clsGraphicalButton
    Set cBotonExpulsar = New clsGraphicalButton
    Set cBotonInvitar = New clsGraphicalButton
    
    Call cBotonAceptar.Initialize(cmdAceptar, "boton-aceptar-default.bmp", _
                                                "boton-aceptar-over.bmp", _
                                                "boton-aceptar-off.bmp", Me)
                                                
    Call cBotonCerrar.Initialize(cmdCerrar, "boton-cerrar-default.bmp", _
                                                "boton-cerrar-over.bmp", _
                                                "boton-cerrar-off.bmp", Me)
                                                
    Call cBotonAbandonar.Initialize(cmdAbandonar, "boton-abandonar-default.bmp", _
                                                "boton-abandonar-over.bmp", _
                                                "boton-abandonar-off.bmp", Me)
                                                
    Call cBotonExpulsar.Initialize(cmdExpulsar, "boton-menos-default.bmp", _
                                                "boton-menos-over.bmp", _
                                                "boton-menos-off.bmp", Me)
                                                
    Call cBotonInvitar.Initialize(cmdInvitar, "boton-mas-default.bmp", _
                                                "boton-mas-over.bmp", _
                                                "boton-mas-off.bmp", Me)
End Sub
Private Sub cmdAbandonar_Click()
    
    On Error GoTo cmdAbandonar_Click_Err
    Call WriteAbandonarGrupo
    Unload Me
    
    Exit Sub

cmdAbandonar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmGrupo.cmdAbandonar_Click", Erl)
    Resume Next
    
End Sub


Private Sub cmdExpulsar_Click()
    
    On Error GoTo cmdExpulsar_Click_Err

    If lstGrupo.ListIndex >= 0 Then
        Call WriteEcharDeGrupo(lstGrupo.ListIndex)
        Unload Me
    End If
    
    Exit Sub

cmdExpulsar_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmGrupo.cmdExpulsar_Click", Erl)
    Resume Next
    
End Sub


Private Sub cmdInvitar_Click()
    
    On Error GoTo cmdInvitar_Click_Err
    
    Unload Me
    Call WriteInvitarGrupo
    
    Exit Sub

cmdInvitar_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmGrupo.cmdInvitar_Click", Erl)
    Resume Next
    
End Sub


Private Sub cmdCerrar_Click()
    
    On Error GoTo Command2_Click_Err
    
    Unload Me
    
    Exit Sub

Command2_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmGrupo.Command2_Click", Erl)
    Resume Next
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Form_KeyPress_Err

    If (KeyAscii = 27) Then
        Unload Me
    End If
    
    Exit Sub

Form_KeyPress_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmGrupo.Form_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    Call moverForm(Me.hWnd)
    
    Exit Sub

Form_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmGrupo.Form_MouseMove", Erl)
    Resume Next
    
End Sub
