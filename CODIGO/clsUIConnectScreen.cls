VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUIConnectScreen"
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
Private Const WINDOW_WIDTH    As Long = 300
Private Const WINDOW_HEIGHT   As Long = 300
Private Const WINDOW_OFFSET_Y As Long = 100  ' adjust this value to move window lower
Private m_UIMgr               As clsUIManager
Private m_lblTitle             As clsUILabel
Private m_lblPass             As clsUILabel
Private m_cbOk          As clsUICheckbox
Private m_btnConnect          As clsUIButton
Private m_inputboxUser          As clsUIInputbox
Private m_inputboxPass          As clsUIInputbox
Private m_boxInfo          As clsUIBoxInfo
Private m_haracinOk          As Boolean

 
Public Sub Init(ByVal dev As Direct3DDevice8)
    Set m_UIMgr = New clsUIManager: m_UIMgr.Init
    Set m_lblTitle = New clsUILabel
    m_lblTitle.Init "INGRESAR", 0, 0, UI_COLOR_WHITE
    m_UIMgr.AddComponent m_lblTitle
    Set m_btnConnect = New clsUIButton
    m_btnConnect.Init "Conectar", 0, 0, 100, 30, UI_COLOR_GOLD, &HFF606060, UI_COLOR_WHITE
    m_UIMgr.AddComponent m_btnConnect
    Set m_inputboxUser = New clsUIInputbox
    m_inputboxUser.Init "email", 0, 0, 200, 30, &HFFFFFF, &HFFFFFF, UI_COLOR_GOLD, False
    m_UIMgr.AddComponent m_inputboxUser
    Set m_inputboxPass = New clsUIInputbox
    m_inputboxPass.Init "clave", 0, 0, 200, 30, &HFFFFFF, &HFFFFFF, UI_COLOR_GOLD, True
    m_UIMgr.AddComponent m_inputboxPass
    Set m_cbOk = New clsUICheckbox
    m_cbOk.Init "Recordarme", 50, 200, &HFFFFFF, &HFFFF00, UI_COLOR_WHITE
    m_UIMgr.AddComponent m_cbOk
    Set m_boxInfo = New clsUIBoxInfo
    m_boxInfo.Init "HORACIO es un Atosigador de fines del siglo XX, aun activo en 2025. Se caracteriza por hostigar varios recursos a la vez y romper sober_anamente las pelotas ...", 0, 0, 250, 150, &HFFFF00, &HFFFFFF, UI_COLOR_BLACK
    m_UIMgr.AddComponent m_boxInfo
    m_haracinOk = False
End Sub




Public Sub render(ByVal dev As Direct3DDevice8)
    Dim sW As Long, sH As Long, wx As Long, wy As Long
    sW = Render_Connect_Rect.Right - Render_Connect_Rect.Left
    sH = Render_Connect_Rect.Bottom - Render_Connect_Rect.Top
    wx = (sW - WINDOW_WIDTH) \ 2 + Render_Connect_Rect.Left
    wy = (sH - WINDOW_HEIGHT) \ 2 + Render_Connect_Rect.Top + WINDOW_OFFSET_Y

    Call Engine_BeginScene

        ' World/background
        RenderConnect 57, 45, 0, 0

        ' --- UI: textured window background ---
        Dim m_texPanelBg As Direct3DTexture8
        Dim m_texPanelBorder As Direct3DTexture8
        Dim texW As Long: Dim texH As Long
        Set m_texPanelBg = UITextures.GetTexture(30, texW, texH)  ' fetch or load from manager cache

        
        Dim i As Long, texPanel As Direct3DTexture8
        Dim x As Long, y As Long, w As Long, h As Long, b As Long
        b = 20
        For i = 30 To 38
            Select Case i
                Case 30
                    x = wx: y = wy: w = WINDOW_WIDTH: h = WINDOW_HEIGHT
                Case 31
                    x = wx: y = wy: w = b: h = b
                Case 32
                    x = wx + b: y = wy: w = WINDOW_WIDTH - b * 2: h = b
                Case 33
                    x = wx + WINDOW_WIDTH - b: y = wy: w = b: h = b
                Case 34
                    x = wx + WINDOW_WIDTH - b: y = wy + b: w = b: h = WINDOW_HEIGHT - b * 2
                Case 35
                    x = wx + WINDOW_WIDTH - b: y = wy + WINDOW_HEIGHT - b: w = b: h = b
                Case 36
                    x = wx + b: y = wy + WINDOW_HEIGHT - b: w = WINDOW_WIDTH - b * 2: h = b
                Case 37
                    x = wx: y = wy + WINDOW_HEIGHT - b: w = b: h = b
                Case 38
                    x = wx: y = wy + b: w = b: h = WINDOW_HEIGHT - b * 2
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
                UIRenderer.AddQuad wx, wy, WINDOW_WIDTH, WINDOW_HEIGHT, cols2, 190
            End If
        Next i
        
        
        ' Position UI controls
        m_lblTitle.SetPosition wx + 120, wy + 40
        m_inputboxUser.SetPosition wx + 48, wy + 77
        m_inputboxPass.SetPosition wx + 48, wy + 117
        m_cbOk.SetPosition wx + 48, wy + 160
        m_btnConnect.SetPosition wx + (WINDOW_WIDTH - m_btnConnect.Width) \ 2, wy + WINDOW_HEIGHT - 90
        m_boxInfo.SetPosition wx + (WINDOW_WIDTH - m_boxInfo.Width) \ 2, wy + 50, m_haracinOk

        ' Draw UI widgets + batched renderer
        m_UIMgr.RenderAll dev
        UIRenderer.Draw DirectDevice

    Call Engine_EndScene(Render_Connect_Rect, frmConnect.render.hWnd)
End Sub


Public Sub HandleMouse(ByVal mouseX As Long, ByVal mouseY As Long, ByVal mouseButtons As Long)
    m_UIMgr.HandleMouse mouseX, mouseY, mouseButtons
End Sub

Public Sub HandleKeyboard(ByVal KeyCode As Long, ByVal shiftState As Long)
    m_UIMgr.HandleKeyboard KeyCode, shiftState
End Sub

Public Function WasConnectClicked() As Boolean
    WasConnectClicked = m_btnConnect.WasClicked
    If WasConnectClicked Then
        m_btnConnect.WasClicked = False
        m_haracinOk = True
        
        Call SetActiveServer("127.0.0.1", "6501")
        Call DoLogin(m_inputboxUser.Text, m_inputboxPass.Text, True)
    End If
End Function



