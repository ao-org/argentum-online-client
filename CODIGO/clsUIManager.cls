VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUIManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Argentum 20 Game Client
'
'    Copyright (C) 2025 Noland Studios LTD
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
' Manages a collection of UI components and preserves render state
Private Components As Collection
' Helper to save UI render states
Private oldShader  As Long
Private oldAlpha   As Long, oldSrc As Long, oldDest As Long

Public Sub Init()
    Set Components = New Collection
End Sub

Public Sub AddComponent(ByVal comp As IUIComponent)
    Components.Add comp
End Sub

Public Sub RenderAll(ByVal dev As Direct3DDevice8)
    ' Render each component
    Dim comp As IUIComponent
    For Each comp In Components
        comp.render dev
    Next
End Sub

Public Sub HandleInput(ByVal mouseX As Long, ByVal mouseY As Long, ByVal mouseButtons As Long)
    Dim comp As IUIComponent
    For Each comp In Components
        comp.HandleInput mouseX, mouseY, mouseButtons
    Next
End Sub
