VERSION 5.00
Begin VB.Form frmCerrar 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2796
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2796
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image cmdCancelar 
      Height          =   420
      Left            =   630
      Top             =   1750
      Width           =   1980
   End
   Begin VB.Image cmdSalir 
      Height          =   420
      Left            =   640
      Top             =   1180
      Width           =   1980
   End
   Begin VB.Image cmdMenuPrincipal 
      Height          =   420
      Left            =   640
      Top             =   610
      Width           =   1980
   End
End
Attribute VB_Name = "frmCerrar"
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
  
'Función para saber si formulario ya es transparente. _
 Se le pasa el Hwnd del formulario en cuestión
 
Public bmoving      As Boolean

Public dX           As Integer

Public dy           As Integer

Private RealizoCambios As String

Private cBotonAceptar As clsGraphicalButton
Private cBotonConstruir As clsGraphicalButton
Private cBotonCerrar As clsGraphicalButton


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
  

    If (KeyCode = vbKeyEscape) Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call FormParser.Parse_Form(Me)
    Call Aplicar_Transparencia(Me.hwnd, 240)

    
    Me.Picture = LoadInterface("desconectar.bmp")
    
    Call LoadButtons
    
    Exit Sub
    
Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCerrar.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub LoadButtons()
        
    Set cBotonAceptar = New clsGraphicalButton
    Set cBotonConstruir = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton

    Call cBotonAceptar.Initialize(cmdMenuPrincipal, "boton-mainmenu-default.bmp", _
                                                "boton-mainmenu-over.bmp", _
                                                "boton-mainmenu-off.bmp", Me)
    
    Call cBotonConstruir.Initialize(cmdCancelar, "boton-cancelar-default.bmp", _
                                                "boton-cancelar-over.bmp", _
                                                "boton-cancelar-off.bmp", Me)
                                                
    Call cBotonCerrar.Initialize(cmdSalir, "boton-salir-default.bmp", _
                                                "boton-salir-over.bmp", _
                                                "boton-salir-off.bmp", Me)
End Sub


Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdMenuPrincipal_Click()
    Call WriteQuit
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    Call CloseClient
End Sub

