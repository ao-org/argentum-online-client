VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUIInputbox"
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
Private m_Interval  As Integer
Private m_IntervalMax  As Integer
Private m_password  As Boolean
Private m_EmptyText As String
Private m_TextW As Integer
Private m_TextH As Integer
Private m_MarginLeft As Integer
Private m_Action   As Integer

Public Sub Init(ByVal Text As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, ByVal clrFill As Long, ByVal clrHover As Long, ByVal clrText As Long, ByVal pswd As Boolean, ByVal act As Integer)
    If m_Impl Is Nothing Then
        Set m_Impl = New clsUIComponentImpl
    End If
    If Len(Text) > 0 Then m_EmptyText = Text
    m_Impl.InitLabel "", x, y, clrText
    m_Impl.SetSize w, h
    m_Impl.ColorFill = clrFill
    m_Impl.HoverColor = clrHover
    m_Clicked = False: m_PrevDown = False: m_IsHover = False
    m_IntervalMax = 100
    m_Interval = 0
    m_password = pswd
    m_TextW = 7
    m_TextH = 12
    m_MarginLeft = 10
    m_Action = act
End Sub

Public Sub IUIComponent_SetVisible(ByVal newState As Boolean)
    Call m_Impl.SetVisible(newState)
End Sub

Public Sub SetPosition(ByVal x As Long, ByVal y As Long)
    Call m_Impl.SetPosition(x, y)
End Sub

Public Function IUIComponent_IsVisible() As Boolean
    IUIComponent_IsVisible = m_Impl.IsVisible
End Function

Public Property Get Text() As String
    Text = m_Impl.Text
End Property

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
    Dim clr As Long
    Dim strAux As String
    Dim colorTextAux As Integer
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
    Set cbTextureL = UITextures.GetTexture(13019, texW, texH)
    Set cbTextureC = UITextures.GetTexture(13020, texW, texH)
    Set cbTextureR = UITextures.GetTexture(13021, texW, texH)

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
    Dim StartX As Integer
    Dim StartY As Integer
'    StartX = m_Impl.x + (m_Impl.Width - (Len(m_Impl.Text) * m_TextW)) / 2
    StartX = m_Impl.x + m_MarginLeft
    StartY = m_Impl.y + (m_Impl.Height - m_TextH) / 2
    colorTextAux = 255
    m_Interval = m_Interval + 1
    If m_Impl.Focused Then
        If (m_Interval < (m_IntervalMax / 2)) Then
            If m_Impl.Text = "" Then
                strAux = "I"
            Else
                strAux = m_Impl.Text & " I"
            End If
        Else
            strAux = m_Impl.Text
        End If
    Else
        If m_Impl.Text = "" Then
            strAux = m_EmptyText
'            StartX = m_Impl.x + (m_Impl.Width - (Len(m_EmptyText) * m_TextW)) / 2
            colorTextAux = 120
        Else
            strAux = m_Impl.Text
        End If
    End If
    rc.Left = StartX: rc.Top = StartY
    rc.Right = m_Impl.x + m_Impl.Width: rc.Bottom = m_Impl.y + m_Impl.Height
    Call UIRenderer.RenderText(strAux, rc.Left, rc.Top, m_Impl.Color, 1, colorTextAux)
    If m_Interval = m_IntervalMax Then m_Interval = 0
End Sub

Private Sub IUIComponent_HandleMouse(ByVal mouseX As Long, ByVal mouseY As Long, ByVal mouseButtons As Long, ByRef act As Integer)
    m_IsHover = (mouseX >= m_Impl.x And mouseX <= m_Impl.x + m_Impl.Width And mouseY >= m_Impl.y And mouseY <= m_Impl.y + m_Impl.Height)
    Dim down As Boolean: down = (mouseButtons <> 0)
    If down And Not m_PrevDown And m_IsHover Then
        m_Clicked = True
        act = 0
        Call m_Impl.SetFocus(True)
    End If
    m_PrevDown = down

End Sub

Private Sub IUIComponent_HandleKeyboard(ByVal keyCode As Long, ByVal shiftState As Long)
    If m_Impl.Focused Then
'            m_Impl.Text = m_Impl.Text & ChrW(KeyCode)
            Dim ch As String
            ch = KeyCodeToChar(keyCode, shiftState)
        
            Select Case ch
                Case "{BACK}"
                    If Len(m_Impl.Text) > 0 Then m_Impl.Text = Left$(m_Impl.Text, Len(m_Impl.Text) - 1)
        
                Case "{ENTER}"
                    ' Optional: raise event or commit text
        
                Case Else
                    If Len(ch) > 0 Then
                        If m_password Then
                            m_Impl.Text = m_Impl.Text & "*"
                        Else
                            m_Impl.Text = m_Impl.Text & ch
                        End If
                    End If
            End Select
    End If
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
