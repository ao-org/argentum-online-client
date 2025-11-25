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
Private m_Impl      As clsUIComponentImpl
Private m_IsHover   As Boolean
Private m_WasHover   As Boolean
Private m_PrevDown  As Boolean
Private m_Clicked   As Boolean

Public Sub Init(ByVal text As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, ByVal clrFill As Long, ByVal clrHover As Long, ByVal clrText As Long)
    If m_Impl Is Nothing Then
        Set m_Impl = New clsUIComponentImpl
    End If

    m_Impl.InitLabel text, x, y, clrText
    m_Impl.SetSize w, h
    m_Impl.ColorFill = clrFill
    m_Impl.HoverColor = clrHover
    m_Clicked = False: m_PrevDown = False: m_IsHover = False
End Sub

Public Function IUIComponent_IsVisible() As Boolean
    IUIComponent_IsVisible = m_Impl.IsVisible
End Function

Public Sub IUIComponent_SetVisible(ByVal newState As Boolean)
    Call m_Impl.SetVisible(newState)
End Sub

Public Sub SetPosition(ByVal x As Long, ByVal y As Long)
    Call m_Impl.SetPosition(x, y)
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

Public Sub IUIComponent_Render(ByVal dev As Direct3DDevice8)
    Dim clr As Long
    clr = IIf(m_IsHover, m_Impl.HoverColor, m_Impl.ColorFill)
    Dim cols(3) As Long
    cols(0) = clr
    cols(1) = clr
    cols(2) = clr
    cols(3) = clr
'    UIRenderer.AddQuad m_Impl.x, m_Impl.y, m_Impl.Width, m_Impl.Height, cols
    Dim cbTextureL As Direct3DTexture8
    Dim cbTextureC As Direct3DTexture8
    Dim cbTextureR As Direct3DTexture8
    Dim texW As Long: Dim texH As Long: Dim b As Integer
    b = 5
    If (m_IsHover) Then
        Set cbTextureL = UITextures.GetTexture(13015, texW, texH)
        Set cbTextureC = UITextures.GetTexture(13016, texW, texH)
        Set cbTextureR = UITextures.GetTexture(13017, texW, texH)
    Else
        Set cbTextureL = UITextures.GetTexture(13009, texW, texH)
        Set cbTextureC = UITextures.GetTexture(13010, texW, texH)
        Set cbTextureR = UITextures.GetTexture(13011, texW, texH)
    End If
    If Not cbTextureL Is Nothing And Not cbTextureC Is Nothing And Not cbTextureR Is Nothing Then
        UIRenderer.AddTexturedQuad _
            m_Impl.x, m_Impl.y, b, m_Impl.Height, _
        cbTextureL, _
        0!, 0!, 1!, 1!, _
        UI_COLOR_WHITE, 255
        UIRenderer.AddTexturedQuad _
            m_Impl.x + b, m_Impl.y, m_Impl.Width - b, m_Impl.Height, _
        cbTextureC, _
        0!, 0!, 1!, 1!, _
        UI_COLOR_WHITE, 255
        UIRenderer.AddTexturedQuad _
            m_Impl.x + m_Impl.Width - b, m_Impl.y, b, m_Impl.Height, _
        cbTextureR, _
        0!, 0!, 1!, 1!, _
        UI_COLOR_WHITE, 255
    Else
    ' Fallback to solid white if texture missing
        Dim cols2(3) As Long
        cols2(0) = UI_COLOR_WHITE
        cols2(1) = UI_COLOR_WHITE
        cols2(2) = UI_COLOR_WHITE
        cols2(3) = UI_COLOR_WHITE
        UIRenderer.AddQuad m_Impl.x, m_Impl.y, m_Impl.Width, m_Impl.Height, cols2, 190
    End If
    Dim rc As RECT
    rc.Left = m_Impl.x: rc.Top = m_Impl.y
    Dim RECT As RECT
    Dim startX As Integer
    Dim startY As Integer
    startX = m_Impl.x + (m_Impl.Width - (Len(m_Impl.text) * 7)) / 2
    startY = m_Impl.y + (m_Impl.Height - 12) / 2
    rc.Left = startX: rc.Top = startY
    rc.Right = m_Impl.x + m_Impl.Width: rc.Bottom = m_Impl.y + m_Impl.Height
    Call UIRenderer.RenderText(m_Impl.text, rc.Left, rc.Top, m_Impl.color, 1, 255)
End Sub

Private Sub IUIComponent_HandleMouse(ByVal mouseX As Long, ByVal mouseY As Long, ByVal mouseButtons As Long, ByRef Hit As Boolean)
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
        Hit = True
        Call m_Impl.SetFocus(True)
        Call ao20audio.PlayWav(SND_CLICK)
    End If
    m_PrevDown = down
End Sub

Private Sub IUIComponent_HandleKeyboard(ByVal keyCode As Long, ByVal shiftState As Long)
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
