VERSION 5.00
Begin VB.Form frmGuildSol 
   BorderStyle     =   0  'None
   Caption         =   "Solicitud de ingreso"
   ClientHeight    =   3435
   ClientLeft      =   0
   ClientTop       =   -180
   ClientWidth     =   4680
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
   ScaleHeight     =   3435
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   280
      MaxLength       =   400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1500
      Width           =   4095
   End
   Begin VB.Image cmdCerrar 
      Height          =   375
      Left            =   4215
      Top             =   0
      Width           =   375
   End
   Begin VB.Image cmdEnviarSolicitud 
      Height          =   375
      Left            =   700
      Top             =   2800
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"frmSolicitud.frx":0000
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   4095
   End
End
Attribute VB_Name = "frmGuildSol"
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
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Dim CName As String
Private cBotonCerrar As clsGraphicalButton
Private cBotonEnviarSolicitud As clsGraphicalButton

Private Sub loadButtons()
    On Error Goto loadButtons_Err

    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonEnviarSolicitud = New clsGraphicalButton
    
    Call cBotonCerrar.Initialize(cmdCerrar, "boton-cerrar-default.bmp", _
                                                "boton-cerrar-over.bmp", _
                                                "boton-cerrar-off.bmp", Me)
                                                
    Call cBotonEnviarSolicitud.Initialize(cmdEnviarSolicitud, "boton-enviarsolicitud-default.bmp", _
                                                    "boton-enviarsolicitud-over.bmp", _
                                                    "boton-enviarsolicitud-off.bmp", Me)
                                                    
    Exit Sub
loadButtons_Err:
    Call TraceError(Err.Number, Err.Description, "frmSolicitud.loadButtons", Erl)
End Sub

Private Sub cmdCerrar_Click()
    On Error Goto cmdCerrar_Click_Err
    Unload Me
    Exit Sub
cmdCerrar_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmSolicitud.cmdCerrar_Click", Erl)
End Sub

Private Sub cmdEnviarSolicitud_Click()
    On Error Goto cmdEnviarSolicitud_Click_Err
    
    On Error GoTo cmdEnviarSolicitud_Click_Err
    
    Call WriteGuildRequestMembership(CName, Replace(Replace(Text1.Text, ",", ";"), vbCrLf, "º"))

    Unload Me

    
    Exit Sub

cmdEnviarSolicitud_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmGuildSol.cmdEnviarSolicitud_Click", Erl)
    Resume Next
    
    Exit Sub
cmdEnviarSolicitud_Click_Err:
    Call TraceError(Err.Number, Err.Description, "frmSolicitud.cmdEnviarSolicitud_Click", Erl)
End Sub

Public Sub RecieveSolicitud(ByVal GuildName As String)
    On Error Goto RecieveSolicitud_Err
    
    On Error GoTo RecieveSolicitud_Err
    

    CName = GuildName

    
    Exit Sub

RecieveSolicitud_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildSol.RecieveSolicitud", Erl)
    Resume Next
    
    Exit Sub
RecieveSolicitud_Err:
    Call TraceError(Err.Number, Err.Description, "frmSolicitud.RecieveSolicitud", Erl)
End Sub

Private Sub Form_Load()
    On Error Goto Form_Load_Err
    
    On Error GoTo Form_Load_Err
    
    Call FormParser.Parse_Form(Me)
    Me.Picture = LoadInterface("ventanaclanes_solicitud_ingreso.bmp")
    Call Aplicar_Transparencia(Me.hwnd, 240)
    Call loadButtons
    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildSol.Form_Load", Erl)
    Resume Next
    
    Exit Sub
Form_Load_Err:
    Call TraceError(Err.Number, Err.Description, "frmSolicitud.Form_Load", Erl)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Goto Form_MouseMove_Err
    
    On Error GoTo Form_MouseMove_Err
    
    Call MoverForm(Me.hwnd)
    
    Exit Sub

Form_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmGuildSol.Form_MouseMove", Erl)
    Resume Next
    
    Exit Sub
Form_MouseMove_Err:
    Call TraceError(Err.Number, Err.Description, "frmSolicitud.Form_MouseMove", Erl)
End Sub
