VERSION 5.00
Begin VB.Form frmGuildBrief 
   BorderStyle     =   0  'None
   ClientHeight    =   4425
   ClientLeft      =   0
   ClientTop       =   -180
   ClientWidth     =   4500
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Desc 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   750
      Left            =   650
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2800
      Width           =   3200
   End
   Begin VB.Label nombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   270
      Width           =   4245
   End
   Begin VB.Label nivel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel de clan:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2805
      TabIndex        =   7
      Top             =   2240
      Width           =   975
   End
   Begin VB.Label Miembros 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Miembros:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2910
      TabIndex        =   5
      Top             =   1640
      Width           =   765
   End
   Begin VB.Label lider 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Lider:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1005
      TabIndex        =   4
      Top             =   1640
      Width           =   435
   End
   Begin VB.Label lblAlineacion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alineacion:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   825
      TabIndex        =   6
      Top             =   2240
      Width           =   795
   End
   Begin VB.Label fundador 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fundador:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   840
      TabIndex        =   2
      Top             =   1040
      Width           =   765
   End
   Begin VB.Label creacion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de creacion:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2595
      TabIndex        =   3
      Top             =   1055
      Width           =   1395
   End
   Begin VB.Image cmdSolicitarIngreso 
      Height          =   375
      Left            =   630
      Top             =   3720
      Width           =   2775
   End
   Begin VB.Image cmdCerrar 
      Height          =   375
      Left            =   4025
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmGuildBrief"
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

Public EsLeader As Boolean
Private cBotonCerrar As clsGraphicalButton
Private cBotonSolicitarIngreso As clsGraphicalButton
Private Sub loadButtons()

    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonSolicitarIngreso = New clsGraphicalButton
    
    Call cBotonCerrar.Initialize(cmdCerrar, "boton-cerrar-default.bmp", _
                                                "boton-cerrar-over.bmp", _
                                                "boton-cerrar-off.bmp", Me)
                                                
    Call cBotonSolicitarIngreso.Initialize(cmdSolicitarIngreso, "boton-solicitaringreso-default.bmp", _
                                                    "boton-solicitaringreso-over.bmp", _
                                                    "boton-solicitaringreso-off.bmp", Me)
                                                    
End Sub

Private Sub aliado_Click()
    
    On Error GoTo aliado_Click_Err
    
    frmCommet.nombre = Right(nombre.Caption, Len(nombre.Caption) - 7)
    frmCommet.t = TIPO.ALIANZA
    frmCommet.Caption = "Ingrese propuesta de alianza"
    Call frmCommet.Show(vbModal, frmGuildBrief)

    
    Exit Sub

aliado_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildBrief.aliado_Click", Erl)
    Resume Next
    
End Sub

Private Sub cmdcerrar_Click()
    Unload Me
End Sub

Private Sub cmdSolicitarIngreso_Click()
    
    On Error GoTo cmdSolicitarIngreso_Click_Err
    
    Call frmGuildSol.RecieveSolicitud(nombre)
    Call frmGuildSol.Show(vbModal, frmGuildBrief)

    
    Exit Sub

cmdSolicitarIngreso_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmGuildBrief.cmdSolicitarIngreso_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command3_Click()
    
    On Error GoTo Command3_Click_Err
    
    frmCommet.nombre = Right(nombre, Len(nombre) - 7)
    frmCommet.t = TIPO.PAZ
    frmCommet.Caption = "Ingrese propuesta de paz"
    Call frmCommet.Show(vbModal, frmGuildBrief)

    
    Exit Sub

Command3_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildBrief.Command3_Click", Erl)
    Resume Next
    
End Sub
Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
       
    Call FormParser.Parse_Form(Me)
    Me.Picture = LoadInterface("ventanaclanes_detalles.bmp")
    Call Aplicar_Transparencia(Me.hwnd, 240)
    Call loadButtons
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildBrief.Form_Load", Erl)
    Resume Next
    
End Sub
Private Sub Guerra_Click()
    
    On Error GoTo Guerra_Click_Err
    
    Call WriteGuildDeclareWar(Right(nombre.Caption, Len(nombre.Caption) - 7))
    Unload Me

    
    Exit Sub

Guerra_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildBrief.Guerra_Click", Erl)
    Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    Call MoverForm(Me.hwnd)
    
    Exit Sub

Form_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmGuildBrief.Form_MouseMove", Erl)
    Resume Next
    
End Sub
