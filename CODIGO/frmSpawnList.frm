VERSION 5.00
Begin VB.Form frmSpawnList 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Invocar NPC"
   ClientHeight    =   4176
   ClientLeft      =   48
   ClientTop       =   216
   ClientWidth     =   2772
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4176
   ScaleWidth      =   2772
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Filter 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Spawn"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   120
      MouseIcon       =   "frmSpawnList.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3600
      Width           =   2490
   End
   Begin VB.ListBox lstCriaturas 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2904
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2490
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Filtrar:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   660
   End
End
Attribute VB_Name = "frmSpawnList"
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

Public ListaCompleta As Boolean

Private Sub Command1_Click()
    
    On Error GoTo Command1_Click_Err
    
    If lstCriaturas.ListIndex < 0 Then Exit Sub

    Call WriteSpawnCreature(lstCriaturas.ItemData(lstCriaturas.ListIndex))

    
    Exit Sub

Command1_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmSpawnList.Command1_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command2_Click()
    
    On Error GoTo Command2_Click_Err
    
    Unload Me
    
    Exit Sub

Command2_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmSpawnList.Command2_Click", Erl)
    Resume Next
    
End Sub

Private Sub Filter_Change()
    
    On Error GoTo Filter_Change_Err
    
    FillList
    
    Exit Sub

Filter_Change_Err:
    Call RegistrarError(Err.number, Err.Description, "frmSpawnList.Filter_Change", Erl)
    Resume Next
    
End Sub

Public Sub FillList()
    
    On Error GoTo FillList_Err

    lstCriaturas.Clear

    Dim i As Long

    For i = 1 To UBound(NpcData())
        If NpcData(i).name <> "Vacío" Then
            If NpcData(i).PuedeInvocar = 1 Or ListaCompleta Then
                If InStr(1, Tilde(NpcData(i).Name), Tilde(Filter.Text)) Then
                    Call lstCriaturas.AddItem(i & " - " & NpcData(i).Name)
                    lstCriaturas.ItemData(lstCriaturas.NewIndex) = i
                End If
            End If
        End If
    Next i
    
    Exit Sub

FillList_Err:
    Call RegistrarError(Err.number, Err.Description, "frmSpawnList.FillList", Erl)
    Resume Next
    
End Sub

Private Sub Filter_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyEscape Then
    
        Unload Me
    
    ElseIf KeyCode = vbKeyReturn Then
    
        lstCriaturas.ListIndex = 0
        lstCriaturas.SetFocus
    
    End If
    
End Sub


Private Sub lstCriaturas_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
    
        Unload Me
    
    ElseIf KeyCode = vbKeyReturn Then
    
        If lstCriaturas.ListIndex < 0 Then Exit Sub

        Call WriteSpawnCreature(lstCriaturas.ItemData(lstCriaturas.ListIndex))
    
    End If

End Sub
