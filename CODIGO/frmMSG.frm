VERSION 5.00
Begin VB.Form frmMSG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mensajes de GMs"
   ClientHeight    =   3450
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4785
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   4785
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3360
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4560
   End
   Begin VB.TextBox txtMsg 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2040
      Width           =   4575
   End
   Begin VB.ListBox List2 
      Height          =   1680
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   4575
   End
   Begin VB.Menu menU_usuario 
      Caption         =   "Usuario"
      Visible         =   0   'False
      Begin VB.Menu mnuIR 
         Caption         =   "Ir donde esta el usuario"
      End
      Begin VB.Menu mnutraer 
         Caption         =   "Traer usuario"
      End
      Begin VB.Menu mnuBorrar 
         Caption         =   "Borrar mensaje"
      End
   End
End
Attribute VB_Name = "frmMSG"
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
Dim Nick As String
Dim TIPO As String
Private Const MAX_GM_MSG = 300

Private MisMSG(0 To MAX_GM_MSG) As String
Private Apunt(0 To MAX_GM_MSG) As Integer

Public Sub CrearGMmSg(Nick As String, msg As String)
If List1.ListCount < MAX_GM_MSG Then
        List1.AddItem Nick & "-" & List1.ListCount
        MisMSG(List1.ListCount - 1) = msg
        Apunt(List1.ListCount - 1) = List1.ListCount - 1
End If
End Sub

Private Sub Command1_Click()
Me.Visible = False
List1.Clear
List2.Clear
txtMsg = ""
End Sub

Private Sub Form_Deactivate()
Me.Visible = False
List1.Clear
List2.Clear
txtMsg = ""
End Sub

Private Sub Form_Load()
List1.Clear
List2.Clear
txtMsg = ""
End Sub

Private Sub list1_Click()
Dim ind As Integer
ind = Val(ReadField(2, List1.List(List1.listIndex), Asc("@")))
txtMsg = List2.List(List1.listIndex)
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
    PopUpMenu menU_usuario
End If

End Sub

Private Sub mnuBorrar_Click()
    If List1.listIndex < 0 Then Exit Sub
    Call ReadNick
    Dim ProximamentTipo As String
    ProximamentTipo = General_Field_Read(2, List1.List(List1.listIndex), "(")
    TIPO = General_Field_Read(1, ProximamentTipo, ")")
    Call WriteSOSRemove(Nick & "Ø" & txtMsg & "Ø" & TIPO)
    List1.RemoveItem List1.listIndex
End Sub
Private Sub mnuIR_Click()
    Dim aux As String
    aux = mid$(ReadField(1, List1.List(List1.listIndex), Asc("-")), 10, Len(ReadField(1, List1.List(List1.listIndex), Asc("-"))))
    Call WriteGoToChar(aux)
End Sub
Private Sub mnutraer_Click()
    Dim aux As String
    aux = mid$(ReadField(1, List1.List(List1.listIndex), Asc("-")), 10, Len(ReadField(1, List1.List(List1.listIndex), Asc("-"))))
    Call WriteSummonChar(aux)
End Sub
Private Sub ReadNick()
If List1.Visible Then
    Nick = General_Field_Read(1, List1.List(List1.listIndex), "(")
    If Nick = "" Then Exit Sub
    Nick = Left$(Nick, Len(Nick))
Else
    Nick = General_Field_Read(1, List2.List(List2.listIndex), "(")
    If Nick = "" Then Exit Sub
    Nick = Left$(Nick, Len(Nick))
End If

End Sub

