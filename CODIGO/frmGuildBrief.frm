VERSION 5.00
Begin VB.Form frmGuildBrief 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Detalles del Clan"
   ClientHeight    =   4092
   ClientLeft      =   48
   ClientTop       =   216
   ClientWidth     =   4104
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4092
   ScaleWidth      =   4104
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Solicitar Ingreso"
      Height          =   375
      Left            =   120
      MouseIcon       =   "frmGuildBrief.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   3600
      Width           =   3855
   End
   Begin VB.Frame Frame3 
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   3855
      Begin VB.TextBox Desc 
         Height          =   975
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información del clan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1980
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      Begin VB.Label nivel 
         Caption         =   "Nivel de clan:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   3500
      End
      Begin VB.Label lblAlineacion 
         Caption         =   "Alineacion:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   3500
      End
      Begin VB.Label Miembros 
         Caption         =   "Miembros:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   3500
      End
      Begin VB.Label lider 
         Caption         =   "Lider:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   3500
      End
      Begin VB.Label creacion 
         Caption         =   "Fecha de creacion:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   3500
      End
      Begin VB.Label fundador 
         Caption         =   "Fundador:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3500
      End
      Begin VB.Label nombre 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3500
      End
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

Private Sub Command2_Click()
    
    On Error GoTo Command2_Click_Err
    
    Call frmGuildSol.RecieveSolicitud(Right$(nombre, Len(nombre) - 7))
    Call frmGuildSol.Show(vbModal, frmGuildBrief)

    
    Exit Sub

Command2_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildBrief.Command2_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command3_Click()
    
    On Error GoTo Command3_Click_Err
    
    frmCommet.nombre = Right(nombre.Caption, Len(nombre.Caption) - 7)
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

