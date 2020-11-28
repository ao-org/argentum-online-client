VERSION 5.00
Begin VB.Form frmGuildFoundation 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Fundación de un Clan"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   4095
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Discord - Web del clan  "
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   1560
      Width           =   3855
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   180
         TabIndex        =   1
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3495
      Begin VB.TextBox txtClanName 
         Alignment       =   2  'Center
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   180
         MaxLength       =   30
         TabIndex        =   0
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nota: No se toleraran nombres inapropiados, para más información consultar el reglamento."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   2805
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      MouseIcon       =   "frmGuildFoundation.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   2520
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Siguiente"
      Height          =   375
      Left            =   2280
      MouseIcon       =   "frmGuildFoundation.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   2520
      Width           =   1695
   End
End
Attribute VB_Name = "frmGuildFoundation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub Command1_Click()

If txtClanName.Text = "" Then
    MensajeAdvertencia "¡Ingrese un nombre!"
    Exit Sub
End If

If Len(txtClanName.Text) <= 30 Then
    If Not AsciiValidos(txtClanName) Then
        MensajeAdvertencia "Nombre invalido."
        Exit Sub
    End If
Else
    MensajeAdvertencia "Nombre demasiado extenso."
    Exit Sub
End If


ClanName = txtClanName
Site = Text2
Unload Me
frmGuildDetails.framAlign.Visible = True
frmGuildDetails.Show vbModeless, frmMain
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Deactivate()
Me.SetFocus
End Sub

Private Sub Form_Load()
Call FormParser.Parse_Form(Me)
End Sub
