VERSION 5.00
Begin VB.Form frmDeleteChar 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   ClientHeight    =   2892
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4668
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2892
   ScaleWidth      =   4668
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDeleteCharCode 
      Alignment       =   2  'Center
      BackColor       =   &H000D1312&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   300
      Left            =   1560
      TabIndex        =   0
      Top             =   1597
      Width           =   1550
   End
   Begin VB.Image btnCerrar 
      Height          =   375
      Left            =   4200
      Top             =   20
      Width           =   375
   End
   Begin VB.Image btnCancelar 
      Height          =   375
      Left            =   290
      Top             =   2180
      Width           =   1935
   End
   Begin VB.Image btnAceptar 
      Height          =   375
      Left            =   2400
      Top             =   2180
      Width           =   1935
   End
End
Attribute VB_Name = "frmDeleteChar"
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
Private cBotonCancelar As clsGraphicalButton
Private cBotonCerrar As clsGraphicalButton

Private Sub btnAceptar_Click()

    Me.txtDeleteCharCode.Text = Trim(Me.txtDeleteCharCode.Text)

    If Me.txtDeleteCharCode.Text <> "" Then
        ModAuth.LoginOperation = e_operation.ConfirmDeleteChar
        Call connectToLoginServer
        delete_char_validate_code = frmDeleteChar.txtDeleteCharCode.Text
        Unload Me
    Else
        Call MsgBox("El código ingresado es inválido.")
    End If
End Sub

Private Sub btnCerrar_Click()
    Unload Me
End Sub
Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    Me.Picture = LoadInterface("ventacodigoverificacion.bmp")
    
    Call loadButtons
    
End Sub

Private Sub loadButtons()
       
    Set cBotonAceptar = New clsGraphicalButton
    Set cBotonCancelar = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton


    Call cBotonAceptar.Initialize(btnAceptar, "boton-aceptar-default.bmp", _
                                                "boton-aceptar-over.bmp", _
                                                "boton-aceptar-off.bmp", Me)
                                                
    Call cBotonCancelar.Initialize(btnCancelar, "boton-cancelar-default.bmp", _
                                                "boton-cancelar-over.bmp", _
                                                "boton-cancelar-off.bmp", Me)
                                                
    Call cBotonCerrar.Initialize(btnCerrar, "boton-cerrar-default.bmp", _
                                                "boton-cerrar-over.bmp", _
                                                "boton-cerrar-off.bmp", Me)
    
End Sub
