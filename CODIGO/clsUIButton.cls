VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUIButton"
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

' Button component: draws a colored quad with hover and click support
' Requires D3DVertex UDT defined in a standard module

Private m_X As Long, m_Y As Long
Private m_W As Long, m_H As Long
Private m_Text As String
Private m_Font As D3DXFont
Private m_ColorFill As Long, m_HoverColor As Long, m_ColorText As Long
Private m_IsHover As Boolean
Private m_PrevDown As Boolean
Private m_Clicked As Boolean

Public Sub Init(ByVal text As String, ByVal x As Long, ByVal y As Long, _
                ByVal w As Long, ByVal h As Long, ByVal font As D3DXFont, _
                ByVal clrFill As Long, ByVal clrHover As Long, ByVal clrText As Long)
    m_Text = text: m_X = x: m_Y = y
    m_W = w: m_H = h
    Set m_Font = font
    m_ColorFill = clrFill
    m_HoverColor = clrHover
    m_ColorText = clrText
    m_Clicked = False: m_PrevDown = False: m_IsHover = False
End Sub

Public Sub SetPosition(ByVal x As Long, ByVal y As Long)
    m_X = x: m_Y = y
End Sub

Public Property Get Width() As Long
    Width = m_W
End Property
Public Property Get Height() As Long
    Height = m_H
End Property

Public Property Get WasClicked() As Boolean
    WasClicked = m_Clicked
End Property
Public Property Let WasClicked(ByVal val As Boolean)
    m_Clicked = val
End Property

Private Sub IUIComponent_Render(ByVal dev As Direct3DDevice8)
    Dim clr As Long
    clr = IIf(m_IsHover, m_HoverColor, m_ColorFill)
    Dim cols(3) As Long
    cols(0) = clr
    cols(1) = clr
    cols(2) = clr
    cols(3) = clr
    UIRenderer.AddQuad m_X, m_Y, m_W, m_H, cols
    Dim rc As RECT
    rc.Left = m_X: rc.Top = m_Y
    rc.Right = m_X + m_W: rc.Bottom = m_Y + m_H
    m_Font.DrawTextW m_Text, -1, rc, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE, m_ColorText
End Sub
Private Sub IUIComponent_HandleInput(ByVal mouseX As Long, ByVal mouseY As Long, ByVal mouseButtons As Long)
    m_IsHover = (mouseX >= m_X And mouseX <= m_X + m_W And mouseY >= m_Y And mouseY <= m_Y + m_H)
    Dim down As Boolean: down = (mouseButtons <> 0)
    If down And Not m_PrevDown And m_IsHover Then m_Clicked = True
    m_PrevDown = down
End Sub


