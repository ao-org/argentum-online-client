VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInstrument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Argentum 20 Game Server
'
'    Copyright (C) 2023 Noland Studios LTD
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

Private Declare Function QueryPerformanceCounter Lib "Kernel32" (x As Currency) As Boolean
Private Declare Function QueryPerformanceFrequency Lib "Kernel32" (x As Currency) As Boolean

Private m_startTime As Currency
Private m_freq As Currency
Private m_overhead As Currency

Public Sub start()
    QueryPerformanceCounter m_startTime
End Sub

Public Function ElapsedSeconds() As Double
    Dim currentTime As Currency
    QueryPerformanceCounter currentTime
    ElapsedSeconds = (currentTime - m_startTime - m_overhead) / m_freq
End Function

Public Function ElapsedMilliseconds() As Double
    ElapsedMilliseconds = ElapsedSeconds * 1000
End Function

Private Sub Class_Initialize()
    QueryPerformanceFrequency m_freq
    Dim ctr1 As Currency
    Dim ctr2 As Currency
    QueryPerformanceCounter ctr1
    QueryPerformanceCounter ctr2
    m_overhead = ctr2 - ctr1
End Sub

