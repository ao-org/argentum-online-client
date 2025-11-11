VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUIComponentImpl"
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
Private m_X     As Long
Private m_Y     As Long
Private m_Text  As String
Private m_Color As Long
Private m_Focus As Boolean
Private m_visible As Boolean

Public Property Let visible(ByVal newState As Boolean)
    m_visible = newState
End Property

Public Property Get visible() As Boolean
    visible = m_visible
End Property

Public Property Let SetX(ByVal newCoordinate As Long)
    m_X = newCoordinate
End Property

Public Property Get GetX()
    GetX = m_X
End Property

Public Property Let SetY(ByVal newCoordinate As Long)
    m_Y = newCoordinate
End Property

Public Property Get GetY()
    GetY = m_Y
End Property







Private Sub Class_Initialize()
    m_visible = True
End Sub

Private Sub Class_Terminate()
    m_visible = False
End Sub
