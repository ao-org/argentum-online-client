VERSION 5.00
Begin VB.Form frmForo 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   5640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox MiMensaje 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Visible         =   0   'False
      Width           =   5400
   End
   Begin VB.TextBox MiMensaje 
      BackColor       =   &H00FFFFFF&
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
      Height          =   4935
      Index           =   1
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   5400
   End
   Begin VB.TextBox Text 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   5385
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   7
      Text            =   "frmForo.frx":0000
      Top             =   600
      Visible         =   0   'False
      Width           =   5430
   End
   Begin VB.ListBox List 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   5325
      ItemData        =   "frmForo.frx":0006
      Left            =   120
      List            =   "frmForo.frx":0008
      TabIndex        =   6
      Top             =   600
      Width           =   5430
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Lista de mensajes"
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
      Left            =   2040
      MouseIcon       =   "frmForo.frx":000A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6240
      Width           =   1560
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
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
      Left            =   3840
      MouseIcon       =   "frmForo.frx":015C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6240
      Width           =   1560
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Dejar Mensaje"
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
      MouseIcon       =   "frmForo.frx":02AE
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6240
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Título"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mensaje"
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
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label Label4 
      Caption         =   "Foro De RevolucionAo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   3
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmForo"
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

Public ForoIndex As Integer
Private Sub Command1_Click()
Dim i
For Each i In Text
    i.Visible = False
Next

Label4.Visible = False

If Not MiMensaje(0).Visible Then
    List.Visible = False
    MiMensaje(0).Visible = True
    MiMensaje(1).Visible = True
    MiMensaje(0).SetFocus
    Command1.Enabled = False
    Label1.Visible = True
    Label2.Visible = True
Else
    Call WriteForumPost(MiMensaje(0).Text, Left$(MiMensaje(1).Text, 450))
    List.AddItem MiMensaje(0).Text
    Load Text(List.ListCount)
    Text(List.ListCount - 1).Text = MiMensaje(1).Text
    List.Visible = True
    MiMensaje(0).Visible = False
    MiMensaje(1).Visible = False
    Command1.Enabled = True
    Label1.Visible = False
    Label2.Visible = False
    Label4.Visible = True
    
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Label4.Visible = True
MiMensaje(0).Visible = False
MiMensaje(1).Visible = False
Command1.Enabled = True
Label1.Visible = False
Label2.Visible = False
Dim i
For Each i In Text
    i.Visible = False
Next
List.Visible = True
End Sub

Private Sub Form_Deactivate()
'Me.SetFocus
End Sub


Private Sub List_Click()
List.Visible = False
Text(List.listIndex).Visible = True

End Sub

Private Sub MiMensaje_Change(Index As Integer)
If Len(MiMensaje(0).Text) <> 0 And Len(MiMensaje(1).Text) <> 0 Then
Command1.Enabled = True
End If

End Sub

