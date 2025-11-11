VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUIBoxInfo"
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
Private m_Impl      As clsUIComponentImpl
Private m_IsHover   As Boolean
Private m_PrevDown  As Boolean
Private m_Clicked   As Boolean
Private m_marginLeft As Integer
Private m_marginTop As Integer
Private m_Draw   As Boolean

Public Sub Init(ByVal text As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, ByVal clrFill As Long, ByVal clrHover As Long, ByVal clrText As Long)
    If m_Impl Is Nothing Then
        Set m_Impl = New clsUIComponentImpl
    End If

    m_Impl.InitLabel text, x, y, clrText
    m_Impl.SetSize w, h
    m_Impl.ColorFill = clrFill
    m_Impl.HoverColor = clrHover
    m_Clicked = False: m_PrevDown = False: m_IsHover = False
    m_Draw = False
End Sub
Public Sub IUIComponent_SetVisible(ByVal newState As Boolean)
    Call m_Impl.SetVisible(newState)
End Sub

Public Function IUIComponent_IsVisible() As Boolean
    IUIComponent_IsVisible = m_Impl.IsVisible
End Function

Public Sub SetPosition(ByVal x As Long, ByVal y As Long, ByVal d As Boolean)
    Call m_Impl.SetPosition(x, y)
    m_Draw = d
End Sub

Public Property Get Width() As Long
    Width = m_Impl.Width
End Property

Public Property Get Height() As Long
    Height = m_Impl.Height
End Property

Public Property Get WasClicked() As Boolean
    WasClicked = m_Clicked
End Property

Public Property Let WasClicked(ByVal val As Boolean)
    m_Clicked = val
End Property

Private Sub IUIComponent_Render(ByVal dev As Direct3DDevice8)
    If Not m_Draw Then Exit Sub
    Dim clr As Long
    clr = IIf(m_IsHover, m_Impl.HoverColor, m_Impl.ColorFill)

    Dim gliphW As Integer
    Dim gliphH As Integer
    Dim gliphPerRow As Integer
    Dim gliphPerCol As Integer
    gliphW = 7
    gliphH = 12
    gliphPerRow = (m_Impl.Width / gliphW) - 2
    gliphPerCol = m_Impl.Height / gliphH

    Dim cols(3) As Long
    cols(0) = clr
    cols(1) = clr
    cols(2) = clr
    cols(3) = clr

    Dim cols2(3) As Long
    cols2(0) = m_Impl.HoverColor
    cols2(1) = m_Impl.HoverColor
    cols2(2) = m_Impl.HoverColor
    cols2(3) = m_Impl.HoverColor

    UIRenderer.AddQuad m_Impl.X, m_Impl.Y, m_Impl.Width, m_Impl.Height, cols2
'    UIRenderer.AddQuad m_X + m_W - 15, m_Y, 15, 15, cols
    Dim i As Long
    Dim strTextAux As String
    For i = 0 To gliphPerRow
        strTextAux = mid(m_Impl.Text, i * gliphPerRow + 1, gliphPerRow)
        Call UIRenderer.RenderText(strTextAux, m_Impl.X + 10, m_Impl.Y + (gliphH * i) + 15, m_Impl.Color, 1, 255)
    Next i
End Sub

Private Sub IUIComponent_HandleMouse(ByVal mouseX As Long, ByVal mouseY As Long, ByVal mouseButtons As Long, ByRef Hit As Boolean)
    m_IsHover = (mouseX >= m_Impl.X And mouseX <= m_Impl.X + m_Impl.Width And mouseY >= m_Impl.Y And mouseY <= m_Impl.Y + m_Impl.Height)
    Dim down As Boolean: down = (mouseButtons <> 0)
    If down And Not m_PrevDown And m_IsHover Then
        m_Clicked = True
'        Hit = True
        Call m_Impl.SetFocus(True)
    End If
    m_PrevDown = down
End Sub

Private Sub IUIComponent_HandleKeyboard(ByVal KeyCode As Long, ByVal shiftState As Long)
    ' no-op for now
End Sub

Private Sub IUIComponent_SetFocus(ByVal val As Boolean)
    Call m_Impl.SetFocus(val)
End Sub

Private Sub Class_Initialize()
    Set m_Impl = New clsUIComponentImpl
End Sub

Private Sub Class_Terminate()
    Set m_Impl = Nothing
End Sub
