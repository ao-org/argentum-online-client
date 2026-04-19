VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUIFrame"
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
Private m_Action  As Integer
Private m_IsHover   As Boolean
Private m_WasHover   As Boolean
Private m_PrevDown  As Boolean
Private m_Clicked   As Boolean
Private m_MoveOldPosX   As Long
Private m_MoveOldPosY   As Long
Private m_BorderSize   As Long
Private m_elements As Collection


Public Sub Init(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    If m_Impl Is Nothing Then
        Set m_Impl = New clsUIComponentImpl
    End If
    m_Impl.x = x
    m_Impl.y = y
    m_Impl.Width = w
    m_Impl.Height = h
    m_Impl.SetVisible True
    m_MoveOldPosX = 0
    m_MoveOldPosY = 0
    m_BorderSize = 10
    Set m_elements = New Collection
End Sub

Public Sub AddUIElement(ByVal elem As IUIComponent)
    m_elements.Add elem
End Sub

Public Function GetUIElement(ByVal id As Integer) As IUIComponent
    Dim Index As Integer
    Index = 1
    Dim comp As IUIComponent
    For Each comp In m_elements
        If id = Index Then
            GetUIElement comp
        End If
    Next
End Function

Public Sub ChangeState(ByVal s As Integer)
    Dim c As IUIComponent
    Select Case s
        Case 1
            SetVisibleId 3, True
            SetVisibleId 4, True
            SetVisibleId 5, False
            SetVisibleId 6, False
            SetVisibleId 8, False
            SetVisibleId 9, True
        Case 2
            SetVisibleId 5, True
            SetVisibleId 6, True
            SetVisibleId 3, False
            SetVisibleId 4, False
            SetVisibleId 8, True
            SetVisibleId 9, False
    End Select
End Sub

Public Sub SetVisibleId(ByVal i As Integer, ByVal b As Boolean)
    Dim Index As Integer
    Index = 1
    Dim comp As IUIComponent
    For Each comp In m_elements
        If i = Index Then
            comp.SetVisible b
            Return
        End If
    Next
End Sub

Public Function IUIComponent_IsVisible() As Boolean
    IUIComponent_IsVisible = m_Impl.IsVisible
End Function

Public Sub IUIComponent_SetVisible(ByVal newState As Boolean)
    Call m_Impl.SetVisible(newState)
    Dim comp As IUIComponent
    For Each comp In m_elements
        comp.SetVisible newState
    Next
End Sub

Public Sub IUIComponent_SetPosition(ByVal x As Long, ByVal y As Long)
    Call m_Impl.SetPosition(x, y)
End Sub


Public Function IUIComponent_GetPerPosX() As Single
    IUIComponent_GetPerPosX = m_Impl.PerX
End Function

Public Function IUIComponent_GetPerPosY() As Single
    IUIComponent_GetPerPosY = m_Impl.PerY
End Function

Public Function IUIComponent_GetPerW() As Single
    IUIComponent_GetPerW = m_Impl.PerW
End Function

Public Function IUIComponent_GetPerH() As Single
    IUIComponent_GetPerH = m_Impl.PerH
End Function

Public Sub IUIComponent_Render(ByVal dev As Direct3DDevice8)

    If Not m_Impl.IsVisible Then
        Exit Sub
    End If
    Dim texW As Long: Dim texH As Long, Top As Long
    Set m_texPanelBg = UITextures.GetTexture(30, texW, texH)  ' fetch or load from manager cache
    

    
    

    Dim sW As Long, sH As Long, wx As Long, wy As Long
    sW = m_Impl.Width
    sH = m_Impl.Height
    Top = 50
    wx = m_Impl.x
    wy = m_Impl.y
    
    Dim i As Long, texPanel As Direct3DTexture8
    Dim x As Long, y As Long, w As Long, h As Long, b As Long
    b = m_BorderSize
    For i = 13044 To 13052
        Select Case i
            Case 13044
                x = wx + b: y = wy + b: w = sW - (2 * b): h = sH - (2 * b)
            Case 13045
                x = wx: y = wy: w = b: h = b
            Case 13046
                x = wx + b: y = wy: w = sW - b * 2: h = b
            Case 13047
                x = wx + sW - b: y = wy: w = b: h = b
            Case 13048
                x = wx + sW - b: y = wy + b: w = b: h = sH - b * 2
            Case 13049
                x = wx + sW - b: y = wy + sH - b: w = b: h = b
            Case 13050
                x = wx + b: y = wy + sH - b: w = sW - b * 2: h = b
            Case 13051
                x = wx: y = wy + sH - b: w = b: h = b
            Case 13052
                x = wx: y = wy + b: w = b: h = sH - b * 2
        End Select
        Set texPanel = UITextures.GetTexture(i, texW, texH)
        If Not texPanel Is Nothing Then
            UIRenderer.AddTexturedQuad _
                x, y, w, h, _
            texPanel, _
            0!, 0!, 1!, 1!, _
            UI_COLOR_WHITE, 255
        Else
        ' Fallback to solid white if texture missing
            Dim cols2(3) As Long
            cols2(0) = UI_COLOR_WHITE
            cols2(1) = UI_COLOR_WHITE
            cols2(2) = UI_COLOR_WHITE
            cols2(3) = UI_COLOR_WHITE
            UIRenderer.AddQuad wx, wy, sW, sH, cols2, 190
        End If
    Next i
    Dim comp As IUIComponent
    Dim auxX As Single
    Dim auxY As Single
    Dim auxW As Single
    Dim auxH As Single
    For Each comp In m_elements
        auxX = comp.GetPerPosX
        auxY = comp.GetPerPosY
        auxW = comp.GetPerW
        auxH = comp.GetPerH
        If comp.IsVisible Then
            comp.SetSize auxW * m_Impl.Width, auxH * m_Impl.Height
            comp.SetPosition m_Impl.x + (auxX * m_Impl.Width), m_Impl.y + (auxY * m_Impl.Height)
            comp.render dev
        End If
    Next
End Sub

Private Sub IUIComponent_HandleMouse(ByVal mouseX As Long, ByVal mouseY As Long, ByVal mouseButtons As Long, ByRef act As Integer)
    m_IsHover = (mouseX >= m_Impl.x And mouseX <= m_Impl.x + m_Impl.Width And mouseY >= m_Impl.y And mouseY <= m_Impl.y + m_Impl.Height)
    Dim down As Boolean: down = (mouseButtons <> 0)
    Dim isHoverBorderR As Boolean
    isHoverBorderR = (mouseX >= m_Impl.x + m_Impl.Width - m_BorderSize And mouseX <= m_Impl.x + m_Impl.Width)
    Dim isHoverBorderB As Boolean
    isHoverBorderB = (mouseY >= m_Impl.y + m_Impl.Height - m_BorderSize And mouseY <= m_Impl.y + m_Impl.Height)
    Dim comp As IUIComponent
    Dim Action As Integer
    For Each comp In m_elements
        If comp.IsVisible Then
            comp.HandleMouse mouseX, mouseY, mouseButtons, Action
            act = Action
        End If
    Next
    
    If m_IsHover And Not m_WasHover Then
        m_WasHover = True
    Else
        If Not m_IsHover Then
            m_WasHover = False
        End If
    End If
    If down And Not m_PrevDown And m_IsHover Then
        m_Clicked = True
'        act = m_Action
        Call m_Impl.SetFocus(True)
        Call ao20audio.PlayWav(SND_CLICK)
        m_MoveOldPosX = mouseX
        m_MoveOldPosY = mouseY
    End If
    If down And m_PrevDown And m_IsHover Then
        Dim mX As Long: mX = mouseX - m_MoveOldPosX
        Dim mY As Long: mY = mouseY - m_MoveOldPosY
        If isHoverBorderR Then
            m_Impl.SetSize m_Impl.Width + mX, m_Impl.Height
        End If
        If isHoverBorderB Then
            m_Impl.SetSize m_Impl.Width, m_Impl.Height + mY
        End If
        If Not isHoverBorderR And Not isHoverBorderB Then
            m_Impl.SetPosition m_Impl.x + mX, m_Impl.y + mY
        End If
        m_MoveOldPosX = mouseX
        m_MoveOldPosY = mouseY
    End If
    m_PrevDown = down
End Sub






