VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUILabel"
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
Implements IUIComponent
Private m_X     As Long, m_Y As Long
Private m_Text  As String
Private m_Color As Long

Public Sub Init(ByVal text As String, ByVal x As Long, ByVal y As Long, ByVal color As Long)
    m_Text = text: m_X = x: m_Y = y
    m_Color = color
End Sub

' Allow external positioning
Public Sub SetPosition(ByVal x As Long, ByVal y As Long)
    m_X = x: m_Y = y
End Sub

Private Sub IUIComponent_Render(ByVal dev As Direct3DDevice8)
    Dim rc As RECT
    rc.Left = m_X: rc.Top = m_Y
    rc.Right = m_X + 400: rc.Bottom = m_Y + 32
    Call UIRenderer.RenderText(m_Text, rc.Left, rc.Top, m_Color, 1, 128)
End Sub

Private Sub IUIComponent_HandleMouse(ByVal mouseX As Long, ByVal mouseY As Long, ByVal mouseButtons As Long)
    ' no-op
End Sub

Private Sub IUIComponent_HandleKeyboard(ByVal keyCode As Integer, ByVal shiftState As Integer)
    ' no-op
End Sub
