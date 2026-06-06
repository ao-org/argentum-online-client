VERSION 5.00
Begin VB.Form FrmViajes 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formulario de Viaje"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3825
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
   ScaleHeight     =   3000
   ScaleWidth      =   3825
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdViajar 
      Caption         =   "Viajar"
      Height          =   615
      Left            =   720
      TabIndex        =   3
      Top             =   1680
      Width           =   2055
   End
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
      Height          =   1080
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3570
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

