VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormMovementManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Argentum 20 Game Client
'
'    Copyright (C) 2023 Noland Studios LTD
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
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 Márquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'

Option Explicit

Dim WithEvents Formulario As Form
Attribute Formulario.VB_VarHelpID = -1

Private DragX As Single
Private Dragy As Single
Public LastButtonPressed As clsGraphicalButton

Private LimitY As Integer

Public Sub Initialize(ByRef tForm As Form, Optional ByVal VerticalLimit As Integer = 0)
    On Error Goto Initialize_Err

'Sets the reference to the form

    Set Formulario = tForm
    Set LastButtonPressed = New clsGraphicalButton
    LimitY = VerticalLimit
    Exit Sub
Initialize_Err:
    Call TraceError(Err.Number, Err.Description, "clsFormMovementManager.Initialize", Erl)
End Sub

Private Sub Formulario_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Goto Formulario_MouseDown_Err

'Implements the mouse down event of the form.

    If LimitY > 0 Then
        If y > LimitY Then Exit Sub
    End If

    DragX = x
    Dragy = y
    Exit Sub
Formulario_MouseDown_Err:
    Call TraceError(Err.Number, Err.Description, "clsFormMovementManager.Formulario_MouseDown", Erl)
End Sub

Private Sub Formulario_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Goto Formulario_MouseMove_Err

'Implements the mouse move event of the form.

    If LimitY > 0 Then
        If y > LimitY Then Exit Sub
    End If
    LastButtonPressed.ToggleToNormal
    'If Button > 0 Then
    'Formulario.Left = Formulario.Left + (X - DragX)
    'Formulario.Top = Formulario.Top + (Y - Dragy)
    'End If
    Exit Sub
Formulario_MouseMove_Err:
    Call TraceError(Err.Number, Err.Description, "clsFormMovementManager.Formulario_MouseMove", Erl)
End Sub

