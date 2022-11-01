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
'*****************************************************
'Form Movement Manager
'Designed and implemented by ZaMa
'zaxtor@gmail.com
'*****************************************************

'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************
Option Explicit

Dim WithEvents Formulario As Form
Attribute Formulario.VB_VarHelpID = -1

Private DragX As Single
Private Dragy As Single
Public LastButtonPressed As clsGraphicalButton

Private LimitY As Integer

Public Sub Initialize(ByRef tForm As Form, Optional ByVal VerticalLimit As Integer = 0)
'***************************************************
'Author: ZaMa
'Last Modify Date: 29/12/09
'Sets the reference to the form
'***************************************************
    Set Formulario = tForm
    Set LastButtonPressed = New clsGraphicalButton
    LimitY = VerticalLimit
End Sub

Private Sub Formulario_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'***************************************************
'Author: ZaMa
'Last Modify Date: 29/12/09
'Implements the mouse down event of the form.
'***************************************************
    If LimitY > 0 Then
        If y > LimitY Then Exit Sub
    End If

    DragX = x
    Dragy = y
End Sub

Private Sub Formulario_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'***************************************************
'Author: ZaMa
'Last Modify Date: 01/07/10
'Implements the mouse move event of the form.
'***************************************************
    If LimitY > 0 Then
        If y > LimitY Then Exit Sub
    End If
    LastButtonPressed.ToggleToNormal
    'If Button > 0 Then
    'Formulario.Left = Formulario.Left + (X - DragX)
    'Formulario.Top = Formulario.Top + (Y - Dragy)
    'End If
End Sub

