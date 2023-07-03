VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGameState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'    Argentum 20 - Game Client Program
'    Copyright (C) 2022 - Noland Studios
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
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'
Option Explicit
Option Base 0

Public Enum e_game_state
    e_state_gameplay_screen = 0
    e_state_connect_screen = 1
    e_state_account_screen = 2
    e_state_createchar_screen = 3
End Enum

Private cur_state As e_game_state

Private Sub Class_Initialize()
cur_state = e_state_connect_screen
End Sub

Public Property Get state() As e_game_state
    state = cur_state
End Property

Public Property Let state(ByVal new_state As e_game_state)
    Dim ScreenInfo As D3DDISPLAYMODE
    If BabelInitialized Then
        If new_state = e_game_state.e_state_gameplay_screen Then
            gameplay_render_offset.x = StartRenderX
            gameplay_render_offset.y = StartRenderY
            
            GameplayDrawAreaRect.Top = StartRenderY
            GameplayDrawAreaRect.Left = StartRenderX
            GameplayDrawAreaRect.Bottom = GameplayDrawAreaRect.Top + Render_Main_Rect.Bottom
            GameplayDrawAreaRect.Right = GameplayDrawAreaRect.Left + Render_Main_Rect.Right
            Call ReloadSettings
        Else
            gameplay_render_offset.x = 0
            gameplay_render_offset.y = 0
            GameplayDrawAreaRect.Top = 0
            GameplayDrawAreaRect.Left = 0
            Call get_game_resolution(ScreenInfo)
            GameplayDrawAreaRect.Bottom = ScreenInfo.Height
            GameplayDrawAreaRect.Right = ScreenInfo.Width
        End If
    Else
        GameplayDrawAreaRect.Top = 0
        GameplayDrawAreaRect.Left = 0
        Call get_game_resolution(ScreenInfo)
        GameplayDrawAreaRect.Bottom = ScreenInfo.Height
        GameplayDrawAreaRect.Right = ScreenInfo.Width
    End If
    RenderCullingRect = GameplayDrawAreaRect
    cur_state = new_state
End Property


