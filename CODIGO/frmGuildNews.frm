VERSION 5.00
Begin VB.Form frmGuildNews 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Noticias de clan"
   ClientHeight    =   4932
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5964
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
   ScaleHeight     =   4932
   ScaleWidth      =   5964
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Caption         =   "Clanes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   3000
      TabIndex        =   10
      Top             =   2400
      Width           =   2895
      Begin VB.CommandButton Command4 
         Caption         =   "Detalles"
         Height          =   375
         Left            =   120
         MouseIcon       =   "frmGuildNews.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   1920
         Width           =   2655
      End
      Begin VB.ListBox guildslist 
         Height          =   1476
         ItemData        =   "frmGuildNews.frx":0152
         Left            =   120
         List            =   "frmGuildNews.frx":0154
         TabIndex        =   11
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Rango de clan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2775
      Begin VB.Label nivel 
         BackStyle       =   0  'Transparent
         Caption         =   "Nivel:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label expacu 
         Caption         =   "Beneficios:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label beneficios 
         BackStyle       =   0  'Transparent
         Caption         =   "No atacarse / Chat de clan / Pedir ayuda (K) / Verse Invisible / Marca de clan / Verse vida / Max miembros: 25"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   255
         Left            =   120
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label expcount 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "400 / 500"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   490
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label porciento 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   490
         Width           =   2415
      End
      Begin VB.Shape EXPBAR 
         BackColor       =   &H000000FF&
         BorderColor     =   &H00FFFFFF&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   210
         Left            =   135
         Top             =   495
         Width           =   960
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Lista de miembros"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   2760
      Begin VB.ListBox miembros 
         Height          =   1884
         ItemData        =   "frmGuildNews.frx":0156
         Left            =   120
         List            =   "frmGuildNews.frx":015D
         TabIndex        =   3
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Noticias de clan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.TextBox news 
         Height          =   1815
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
   End
End
Attribute VB_Name = "frmGuildNews"
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

Private Sub beneficios_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo beneficios_MouseMove_Err
    
    porciento.Visible = True
    expcount.Visible = False

    
    Exit Sub

beneficios_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildNews.beneficios_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Command1_Click()
    
    On Error GoTo Command1_Click_Err
    

    

    Unload Me
    frmMain.SetFocus

    
    Exit Sub

Command1_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildNews.Command1_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command4_Click()
    
    On Error GoTo Command4_Click_Err
    
    Call WriteGuildRequestDetails(guildslist.List(guildslist.ListIndex))

    
    Exit Sub

Command4_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildNews.Command4_Click", Erl)
    Resume Next
    
End Sub

Private Sub Frame5_DragDrop(source As Control, x As Single, y As Single)
    
    On Error GoTo Frame5_DragDrop_Err
    
    porciento.Visible = True
    expcount.Visible = False

    
    Exit Sub

Frame5_DragDrop_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildNews.Frame5_DragDrop", Erl)
    Resume Next
    
End Sub

Private Sub porciento_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo porciento_MouseMove_Err
    
    porciento.Visible = False
    expcount.Visible = True

    
    Exit Sub

porciento_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildNews.porciento_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call FormParser.Parse_Form(Me)

    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "frmGuildNews.Form_Load", Erl)
    Resume Next
    
End Sub

