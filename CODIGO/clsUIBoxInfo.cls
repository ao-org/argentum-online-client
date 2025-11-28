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
Private m_MarginLeft As Integer
Private m_MarginTop As Integer
Private m_Draw   As Boolean
Private m_lblTitle  As clsUILabel
Private m_btnOk  As clsUIButton

Public Sub Init(ByVal Text As String, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, ByVal clrFill As Long, ByVal clrHover As Long, ByVal clrText As Long)
    If m_Impl Is Nothing Then
        Set m_Impl = New clsUIComponentImpl
    End If
    
    Set m_lblTitle = New clsUILabel
    m_lblTitle.Init "INGRESAR", 0, 0, UI_COLOR_WHITE, 255, 0.5, 8
    
    Set m_btnOk = New clsUIButton
    m_btnOk.Init "OK", 0, 0, 150, 40, UI_COLOR_GOLD, &HFF606060, UI_COLOR_WHITE
    
    m_Impl.InitLabel Text, x, y, clrText
    m_Impl.SetSize w, h
    m_Impl.ColorFill = clrFill
    m_Impl.HoverColor = clrHover
    m_Clicked = False: m_PrevDown = False: m_IsHover = False
    m_Draw = False
    m_MarginLeft = 30
    m_MarginTop = 40
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

Public Property Let message(ByVal str As String)
    m_Impl.Text = str
End Property

Public Property Let title(ByVal str As String)
    m_lblTitle.Text = str
End Property

Private Sub IUIComponent_Render(ByVal dev As Direct3DDevice8)
    If Not m_Draw Then Exit Sub
    Dim clr As Long
    clr = IIf(m_IsHover, m_Impl.HoverColor, m_Impl.ColorFill)

    Dim gliphW As Integer
    Dim gliphH As Integer
    Dim gliphPerRow As Integer
    Dim gliphPerCol As Integer
    gliphW = 12
    gliphH = 12
    gliphPerRow = (m_Impl.Width / gliphW) - 2
    gliphPerCol = m_Impl.Height / gliphH

    Dim m_texPanelBg As Direct3DTexture8
        Dim m_texPanelBorder As Direct3DTexture8
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
        b = 10
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
    m_lblTitle.SetPosition wx + (sW - (Len(m_lblTitle.Text) * gliphW)) / 2, wy + 15
    m_lblTitle.IUIComponent_Render dev
    Dim strTextAux As String
'    For i = 0 To gliphPerRow
'        strTextAux = mid(m_Impl.Text, i * gliphPerRow + 1, gliphPerRow)
'        Call UIRenderer.RenderText(m_Impl.Text, wx + (sW - (Len(strTextAux) * 10)) / 2, wy + (gliphH * i) + 50, m_Impl.Color, 8, 255, 0.5)
'    Next i
    Call UIRenderer.RenderText(m_Impl.Text, wx + (sW - (Len(m_Impl.Text) * 5)) / 2, wy + 50, m_Impl.Color, 8, 255, 0.5)
    m_btnOk.SetPosition wx + (sW - m_btnOk.Width) / 2, wy + Me.Height - 55
    m_btnOk.IUIComponent_Render dev
End Sub

Private Sub IUIComponent_HandleMouse(ByVal mouseX As Long, ByVal mouseY As Long, ByVal mouseButtons As Long, ByRef hit As Boolean)
    m_IsHover = (mouseX >= m_Impl.x And mouseX <= m_Impl.x + m_Impl.Width And mouseY >= m_Impl.y And mouseY <= m_Impl.y + m_Impl.Height)
    Dim down As Boolean: down = (mouseButtons <> 0)
    If down And Not m_PrevDown And m_IsHover Then
        m_Clicked = True
'        Hit = True
        Call m_Impl.SetFocus(True)
    End If
    m_PrevDown = down
    m_btnOk.IUIComponent_HandleMouse mouseX, mouseY, mouseButtons, hit
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
