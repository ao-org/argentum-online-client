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
Private m_Impl As clsUIComponentImpl
Private m_Opacity As Integer
Private m_Size As Single
Private m_FontIndex As Integer
Private m_Action  As Integer
Private m_IsHover   As Boolean
Private m_WasHover   As Boolean
Private m_PrevDown  As Boolean
Private m_Clicked   As Boolean

Public Sub Init(ByVal Text As String, ByVal x As Long, ByVal y As Long, ByVal Color As Long, ByVal opacity As Integer, ByVal size As Single, FontIndex, ByVal act As Integer)
    If m_Impl Is Nothing Then
        Set m_Impl = New clsUIComponentImpl
    End If
    Call m_Impl.InitLabel(Text, x, y, Color)
    m_Opacity = opacity
    m_Size = size
    m_Impl.SetSize 100, 20
    m_FontIndex = fontIndex
    m_Action = act
End Sub

Public Function IUIComponent_IsVisible() As Boolean
    IUIComponent_IsVisible = m_Impl.IsVisible
End Function

Public Sub IUIComponent_SetVisible(ByVal newState As Boolean)
    Call m_Impl.SetVisible(newState)
End Sub

' Allow external positioning
Public Sub SetPosition(ByVal x As Long, ByVal y As Long)
    Call m_Impl.SetPosition(x, y)
End Sub

Public Sub GetPosition(ByRef x As Long, ByRef y As Long)
    Call m_Impl.GetPosition(x, y)
End Sub

Public Sub IUIComponent_Render(ByVal dev As Direct3DDevice8)
    If m_Impl Is Nothing Then Exit Sub
    If Not m_Impl.IsVisible Then Exit Sub

    Dim PosX As Long
    Dim PosY As Long
    Call m_Impl.GetPosition(PosX, PosY)

    Dim rc As RECT
    rc.Left = PosX
    rc.Top = PosY
    rc.Right = PosX + 400
    rc.Bottom = PosY + 32
    Call UIRenderer.RenderText(m_Impl.Text, rc.Left, rc.Top, m_Impl.Color, m_FontIndex, m_Opacity, m_Size)
End Sub

Private Sub IUIComponent_HandleMouse(ByVal mouseX As Long, ByVal mouseY As Long, ByVal mouseButtons As Long, ByRef act As Integer)
    m_IsHover = (mouseX >= m_Impl.x And mouseX <= m_Impl.x + m_Impl.Width And mouseY >= m_Impl.y And mouseY <= m_Impl.y + m_Impl.Height)
    Dim down As Boolean: down = (mouseButtons <> 0)
    If m_IsHover And Not m_WasHover Then
        m_WasHover = True
        Call ao20audio.PlayWav(SND_CLICK_OVER)
    Else
        If Not m_IsHover Then
            m_WasHover = False
        End If
    End If
    If down And Not m_PrevDown And m_IsHover Then
        m_Clicked = True
        act = m_Action
        Call m_Impl.SetFocus(True)
        Call ao20audio.PlayWav(SND_CLICK)
    End If
    m_PrevDown = down
End Sub

Private Sub IUIComponent_HandleKeyboard(ByVal keyCode As Long, ByVal shiftState As Long)
    ' Labels do not handle keyboard input.
End Sub

Private Sub IUIComponent_SetFocus(ByVal val As Boolean)
    Call m_Impl.SetFocus(val)
End Sub

Public Property Let Text(ByVal newText As String)
    m_Impl.Text = newText
End Property

Public Property Get Text() As String
    Text = m_Impl.Text
End Property

Public Property Let Color(ByVal newColor As Long)
    m_Impl.Color = newColor
End Property

Public Property Get Color() As Long
    Color = m_Impl.Color
End Property

Public Property Let x(ByVal newX As Long)
    m_Impl.x = newX
End Property

Public Property Get x() As Long
    x = m_Impl.x
End Property

Public Property Let y(ByVal newY As Long)
    m_Impl.y = newY
End Property

Public Property Get y() As Long
    y = m_Impl.y
End Property

Private Sub Class_Initialize()
    Set m_Impl = New clsUIComponentImpl
End Sub

Private Sub Class_Terminate()
    Set m_Impl = Nothing
End Sub
