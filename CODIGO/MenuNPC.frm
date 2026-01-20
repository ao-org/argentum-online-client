VERSION 5.00
Begin VB.Form MenuNPC 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   123
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label OpcionLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LIBERAR TODOS"
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
      Height          =   300
      Index           =   4
      Left            =   0
      TabIndex        =   4
      Top             =   1530
      Width           =   1950
   End
   Begin VB.Label OpcionLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LIBERAR"
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
      Height          =   300
      Index           =   3
      Left            =   0
      TabIndex        =   3
      Top             =   1170
      Width           =   1950
   End
   Begin VB.Image OpcionImg 
      Height          =   360
      Index           =   3
      Left            =   0
      Top             =   1080
      Width           =   1950
   End
   Begin VB.Image OpcionImg 
      Height          =   360
      Index           =   2
      Left            =   0
      Top             =   720
      Width           =   1950
   End
   Begin VB.Label OpcionLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ACOMPAÑAR TODOS"
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
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   810
      Width           =   1950
   End
   Begin VB.Image OpcionImg 
      Height          =   360
      Index           =   1
      Left            =   0
      Top             =   360
      Width           =   1950
   End
   Begin VB.Label OpcionLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ACOMPAÑAR"
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
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   450
      Width           =   1950
   End
   Begin VB.Image OpcionImg 
      Height          =   360
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   1950
   End
   Begin VB.Label OpcionLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "QUIETO"
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
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   90
      Width           =   1950
   End
End
Attribute VB_Name = "MenuNPC"
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
Private Over As Integer

Private Sub Form_Load()
    Call Aplicar_Transparencia(Me.hWnd, 180)
    OpcionLbl(0).Caption = JsonLanguage.Item("FORM_OPCION_5")
    OpcionLbl(1).Caption = JsonLanguage.Item("FORM_OPCION_6")
    OpcionLbl(2).Caption = JsonLanguage.Item("FORM_OPCION_7")
    OpcionLbl(3).Caption = JsonLanguage.Item("FORM_OPCION_8")
    OpcionLbl(4).Caption = JsonLanguage.Item("FORM_OPCION_25")
    Over = -1
End Sub

Private Sub OpcionImg_Click(Index As Integer)
    Select Case Index
        Case 0
            Call ParseUserCommand("/QUIETO")
        Case 1
            Call ParseUserCommand("/ACOMPAÑAR")
        Case 2
            Call ParseUserCommand("/ACOMPAÑARTODOS")
        Case 3
            If ConfirmPetRelease = 1 Then
                PreguntaScreen = JsonLanguage.Item("MENSAJEBOX_LIBERAR_MASCOTA")
                Pregunta = True
                PreguntaLocal = True
                PreguntaNUM = 3
            Else
                Call ParseUserCommand("/LIBERAR")
            End If
        Case 4
            If ConfirmPetRelease = 1 Then
                PreguntaScreen = JsonLanguage.Item("MENSAJEBOX_LIBERAR_TODAS_MASCOTA")
                Pregunta = True
                PreguntaLocal = True
                PreguntaNUM = 4
            Else
                Call ParseUserCommand("/LIBERARTODOS")
            End If
    End Select
    Unload Me
End Sub

Private Sub OpcionImg_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Over <> Index Then
        If Over >= 0 Then
            OpcionLbl(Over).ForeColor = vbWhite
        End If
        OpcionLbl(Index).ForeColor = vbYellow
        Over = Index
    End If
End Sub

Private Sub OpcionLbl_Click(Index As Integer)
    Call OpcionImg_Click(Index)
End Sub

Private Sub OpcionLbl_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call OpcionImg_MouseMove(Index, Button, Shift, x, y)
End Sub

Public Sub LostFocus()
    If Over >= 0 Then
        OpcionLbl(Over).ForeColor = vbWhite
        Over = -1
    End If
End Sub
