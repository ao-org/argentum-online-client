VERSION 5.00
Begin VB.Form frmGuildBrief 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Detalles del Clan"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   4110
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
   ScaleHeight     =   4095
   ScaleWidth      =   4110
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
      Caption         =   "Descripci�n"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
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
      Caption         =   "Informaci�n del clan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
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
'Argentum Online 0.11.6
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Mat�as Fernando Peque�o
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
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit

Public EsLeader As Boolean

Private Sub aliado_Click()
    frmCommet.nombre = Right(nombre.Caption, Len(nombre.Caption) - 7)
    frmCommet.T = TIPO.ALIANZA
    frmCommet.Caption = "Ingrese propuesta de alianza"
    Call frmCommet.Show(vbModal, frmGuildBrief)

End Sub

Private Sub Command2_Click()
    Call frmGuildSol.RecieveSolicitud(Right$(nombre, Len(nombre) - 7))
    Call frmGuildSol.Show(vbModal, frmGuildBrief)

End Sub

Private Sub Command3_Click()
    frmCommet.nombre = Right(nombre.Caption, Len(nombre.Caption) - 7)
    frmCommet.T = TIPO.PAZ
    frmCommet.Caption = "Ingrese propuesta de paz"
    Call frmCommet.Show(vbModal, frmGuildBrief)

End Sub

Private Sub Form_Load()
    Call FormParser.Parse_Form(Me)

End Sub

Private Sub Guerra_Click()
    Call WriteGuildDeclareWar(Right(nombre.Caption, Len(nombre.Caption) - 7))
    Unload Me

End Sub
