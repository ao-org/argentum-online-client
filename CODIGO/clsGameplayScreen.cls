VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGameplayScreen"
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
Private Const WINDOW_WIDTH    As Long = 578
Private Const WINDOW_HEIGHT   As Long = 475
Private Const WINDOW_OFFSET_Y As Long = 0  ' adjust this value to move window lower
Private m_UIMgr               As clsUIManager
Private m_lblClass            As clsUILabel
Private m_lblRace             As clsUILabel
Private m_lblGender           As clsUILabel
Private m_lblStatus           As clsUILabel
Private m_lblJailTime         As clsUILabel
Private m_lblStrength         As clsUILabel
Private m_lblAgility          As clsUILabel
Private m_lblIntelligence     As clsUILabel
Private m_lblConstitution     As clsUILabel
Private m_lblCharisma         As clsUILabel
Private m_UIState             As clsUIComponentImpl

Public Function IsVisible() As Boolean
    IsVisible = m_UIState.visible
End Function

Public Sub SetVisible(ByVal newState As Boolean)
    m_UIState.visible = newState
End Sub

Public Sub Init(ByVal dev As Direct3DDevice8)
    Set m_UIMgr = New clsUIManager
    Call m_UIMgr.Init
    Set m_lblClass = New clsUILabel
    Call m_lblClass.Init("Class", 100, 100, UI_COLOR_WHITE, 255, 0.5, 1, 0)
    Call m_UIMgr.AddComponent(m_lblClass)
    Set m_lblRace = New clsUILabel
    Call m_lblRace.Init("Race", 100, 100, UI_COLOR_WHITE, 255, 0.5, 1, 0)
    Call m_UIMgr.AddComponent(m_lblRace)
    Set m_lblGender = New clsUILabel
    Call m_lblGender.Init("Gender", 100, 100, UI_COLOR_WHITE, 255, 0.5, 1, 0)
    Call m_UIMgr.AddComponent(m_lblGender)
    Set m_lblStatus = New clsUILabel
    Call m_lblStatus.Init("Status", 100, 100, UI_COLOR_WHITE, 255, 0.5, 1, 0)
    Call m_UIMgr.AddComponent(m_lblStatus)
    Set m_lblJailTime = New clsUILabel
    Call m_lblJailTime.Init("JailTime", 100, 100, UI_COLOR_WHITE, 255, 0.5, 1, 0)
    Call m_UIMgr.AddComponent(m_lblJailTime)
    Set m_lblStrength = New clsUILabel
    Call m_lblStrength.Init("Strength", 100, 100, UI_COLOR_WHITE, 255, 0.5, 1, 0)
    Call m_UIMgr.AddComponent(m_lblStrength)
    Set m_lblAgility = New clsUILabel
    Call m_lblAgility.Init("Agility", 100, 100, UI_COLOR_WHITE, 255, 0.5, 1, 0)
    Call m_UIMgr.AddComponent(m_lblAgility)
    Set m_lblIntelligence = New clsUILabel
    Call m_lblIntelligence.Init("Intelligence", 100, 100, UI_COLOR_WHITE, 255, 0.5, 1, 0)
    Call m_UIMgr.AddComponent(m_lblIntelligence)
    Set m_lblConstitution = New clsUILabel
    Call m_lblConstitution.Init("Constitution", 100, 100, UI_COLOR_WHITE, 255, 0.5, 1, 0)
    Call m_UIMgr.AddComponent(m_lblConstitution)
    Set m_lblCharisma = New clsUILabel
    Call m_lblCharisma.Init("CHarisma", 100, 100, UI_COLOR_WHITE, 255, 0.5, 1, 0)
    Call m_UIMgr.AddComponent(m_lblCharisma)
End Sub

Private Sub render_ui(ByVal dev As Direct3DDevice8)
    Dim sW As Long
    Dim sH As Long
    Dim wx As Long
    Dim wy As Long
    sW = 0
    sH = 0
    wx = WINDOW_WIDTH
    wy = WINDOW_HEIGHT
    
    Dim cols2(3) As Long
    cols2(0) = UI_COLOR_WHITE
    cols2(1) = UI_COLOR_WHITE
    cols2(2) = UI_COLOR_WHITE
    cols2(3) = UI_COLOR_WHITE
    Call UIRenderer.AddQuad(0, 0, WINDOW_WIDTH, WINDOW_HEIGHT, cols2, 190)
    Call m_UIMgr.RenderAll(dev)
    Call UIRenderer.Draw(dev)
End Sub



Public Sub render(ByVal dev As Direct3DDevice8)
    Call Engine_BeginScene
            Call engine.render
            'Call render_ui
    Call Engine_EndScene(Render_Main_Rect, 0)
End Sub

Public Sub HandleMouse(ByVal mouseX As Long, ByVal mouseY As Long, ByVal mouseButtons As Long)
    Dim actionID As Integer
    m_UIMgr.HandleMouse mouseX, mouseY, mouseButtons, actionID
End Sub

Public Sub HandleKeyboard(ByVal keyCode As Long, ByVal shiftState As Long)
    m_UIMgr.HandleKeyboard keyCode, shiftState
End Sub

Private Sub Class_Initialize()
    Set m_UIState = New clsUIComponentImpl
End Sub

Private Sub Class_Terminate()
    Set m_UIState = Nothing
End Sub
