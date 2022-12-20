Attribute VB_Name = "ModSettings"
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

Const CustomSettingsFile As String = "\..\Recursos\OUTPUT\Configuracion.ini"
Const DefaultSettingsFile As String = "\..\Recursos\OUTPUT\DefaultSettings.ini"

Public Function InitializeSettings() As Boolean
    
    If Not FileExist(App.path & DefaultSettingsFile, vbArchive) Then
        InitializeSettings = False
        Exit Function
    End If
    If Not FileExist(App.path & CustomSettingsFile, vbArchive) Then
        Call FileSystem.FileCopy(App.path & DefaultSettingsFile, App.path & CustomSettingsFile)
    End If
    InitializeSettings = True
End Function

Public Function GetSetting(ByVal Section As String, ByVal Name As String) As String
    Dim currentValue As String
    currentValue = GetVar(App.path & CustomSettingsFile, Section, Name)
    If currentValue = "" Then
        currentValue = GetVar(App.path & DefaultSettingsFile, Section, Name)
    End If
    GetSetting = currentValue
End Function

Public Sub SaveSetting(ByVal Section As String, ByVal Name As String, ByVal Value As String)
    Call WriteVar(App.path & CustomSettingsFile, Section, Name, Value)
End Sub


