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

Private Const WINDOW_WIDTH As Long = 400
Private Const WINDOW_HEIGHT As Long = 200
Private Const WINDOW_OFFSET_Y As Long = 150  ' adjust this value to move window lower
Private m_UIMgr As clsUIManager
Private m_lblUser As clsUILabel
Private m_lblPass As clsUILabel
Private m_btnConnect As clsUIButton
 
    
Public Sub Init(ByVal dev As Direct3DDevice8)
    Set m_UIMgr = New clsUIManager: m_UIMgr.Init

    Set m_lblUser = New clsUILabel
    m_lblUser.Init "Username:", 0, 0, &HFFFFFFFF
    m_UIMgr.AddComponent m_lblUser

    Set m_lblPass = New clsUILabel
    m_lblPass.Init "Password:", 0, 0, &HFFFFFFFF
    m_UIMgr.AddComponent m_lblPass

    Set m_btnConnect = New clsUIButton
    
    m_btnConnect.Init "Connect", 0, 0, 100, 30, &HFF0000, &HFF606060, &HFFFFFFFF
    m_UIMgr.AddComponent m_btnConnect
End Sub

Public Sub render(ByVal dev As Direct3DDevice8)
    Dim sW As Long, sH As Long, wx As Long, wy As Long
    sW = Render_Connect_Rect.Right - Render_Connect_Rect.Left
    sH = Render_Connect_Rect.Bottom - Render_Connect_Rect.Top
    wx = (sW - WINDOW_WIDTH) \ 2 + Render_Connect_Rect.Left
    wy = (sH - WINDOW_HEIGHT) \ 2 + Render_Connect_Rect.Top + WINDOW_OFFSET_Y

    Call Engine_BeginScene
    
        RenderConnect 57, 45, 0, 0
    
        Dim cols(3) As Long
        cols(0) = UI_COLOR_WHITE
        cols(1) = UI_COLOR_WHITE
        cols(2) = UI_COLOR_WHITE
        cols(3) = UI_COLOR_WHITE
        UIRenderer.AddQuad wx, wy, WINDOW_WIDTH, WINDOW_HEIGHT, cols, 190
        
        ' Position UI
        m_lblUser.SetPosition wx + 20, wy + 40
        m_lblPass.SetPosition wx + 20, wy + 80
        m_btnConnect.SetPosition wx + (WINDOW_WIDTH - m_btnConnect.Width) \ 2, wy + WINDOW_HEIGHT - 60
        
        m_UIMgr.RenderAll dev
        UIRenderer.Draw DirectDevice
    
    Call Engine_EndScene(Render_Connect_Rect, frmConnect.render.hWnd)
    
End Sub

Public Sub HandleInput(ByVal mouseX As Long, ByVal mouseY As Long, ByVal mouseButtons As Long)
    m_UIMgr.HandleInput mouseX, mouseY, mouseButtons
End Sub

Public Function WasConnectClicked() As Boolean
    WasConnectClicked = m_btnConnect.WasClicked
    If WasConnectClicked Then m_btnConnect.WasClicked = False
End Function

