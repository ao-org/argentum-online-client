VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUIComponentImpl"
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
Private m_X        As Long
Private m_Y        As Long
Private m_Text     As String
Private m_Color    As Long
Private m_ColorFill As Long
Private m_HoverColor As Long
Private m_Focus    As Boolean
Private m_visible  As Boolean
Private m_W        As Long
Private m_H        As Long

Public Property Let visible(ByVal newState As Boolean)
    m_visible = newState
End Property

Public Property Get visible() As Boolean
    visible = m_visible
End Property

Public Property Let X(ByVal newCoordinate As Long)
    m_X = newCoordinate
End Property

Public Property Get X() As Long
    X = m_X
End Property

Public Property Let Y(ByVal newCoordinate As Long)
    m_Y = newCoordinate
End Property

Public Property Get Y() As Long
    Y = m_Y
End Property

Public Property Let Text(ByVal newText As String)
    m_Text = newText
End Property

Public Property Get Text() As String
    Text = m_Text
End Property

Public Property Let Color(ByVal newColor As Long)
    m_Color = newColor
End Property

Public Property Get Color() As Long
    Color = m_Color
End Property

Public Property Let ColorFill(ByVal newColor As Long)
    m_ColorFill = newColor
End Property

Public Property Get ColorFill() As Long
    ColorFill = m_ColorFill
End Property

Public Property Let HoverColor(ByVal newColor As Long)
    m_HoverColor = newColor
End Property

Public Property Get HoverColor() As Long
    HoverColor = m_HoverColor
End Property

Public Property Let Focused(ByVal newState As Boolean)
    m_Focus = newState
End Property

Public Property Get Focused() As Boolean
    Focused = m_Focus
End Property

Public Sub SetPosition(ByVal newX As Long, ByVal newY As Long)
    m_X = newX
    m_Y = newY
End Sub

Public Sub GetPosition(ByRef currentX As Long, ByRef currentY As Long)
    currentX = m_X
    currentY = m_Y
End Sub

Public Property Let Width(ByVal newWidth As Long)
    m_W = newWidth
End Property

Public Property Get Width() As Long
    Width = m_W
End Property

Public Property Let Height(ByVal newHeight As Long)
    m_H = newHeight
End Property

Public Property Get Height() As Long
    Height = m_H
End Property

Public Sub SetSize(ByVal newWidth As Long, ByVal newHeight As Long)
    m_W = newWidth
    m_H = newHeight
End Sub

Public Sub GetSize(ByRef currentWidth As Long, ByRef currentHeight As Long)
    currentWidth = m_W
    currentHeight = m_H
End Sub

Public Sub InitLabel(ByVal labelText As String, ByVal newX As Long, ByVal newY As Long, ByVal labelColor As Long)
    m_Text = labelText
    m_X = newX
    m_Y = newY
    m_Color = labelColor
    m_Focus = False
End Sub

Public Function IsVisible() As Boolean
    IsVisible = m_visible
End Function

Public Sub SetVisible(ByVal newState As Boolean)
    m_visible = newState
End Sub

Public Sub SetFocus(ByVal newState As Boolean)
    m_Focus = newState
End Sub

Private Sub Class_Initialize()
    m_visible = True
End Sub

Private Sub Class_Terminate()
    m_visible = False
End Sub
