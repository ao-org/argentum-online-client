Attribute VB_Name = "debug_tools"
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

Public BuildFlags As String

Public Sub Init()
    BuildFlags = GetBuildFlags
End Sub

Public Function GetBuildFlags() As String
    Dim s As String
    '--- Compresion
    #If Compresion = 1 Then
        s = s & "Compresion=1; "
    #Else
        s = s & "Compresion=0; "
    #End If

    '--- Developer
    #If Developer = 1 Then
        s = s & "Developer=1; "
    #Else
        s = s & "Developer=0; "
    #End If

    '--- DEBUGGING
    #If DEBUGGING = 1 Then
        s = s & "DEBUGGING=1; "
    #Else
        s = s & "DEBUGGING=0; "
    #End If

    '--- PYMMO
    #If PYMMO = 1 Then
        s = s & "PYMMO=1; "
    #Else
        s = s & "PYMMO=0; "
    #End If

    '--- ENABLE_ANTICHEAT
    #If ENABLE_ANTICHEAT = 1 Then
        s = s & "ENABLE_ANTICHEAT=1; "
    #Else
        s = s & "ENABLE_ANTICHEAT=0; "
    #End If

    '--- DXUI
    #If DXUI = 1 Then
        s = s & "DXUI=1; "
    #Else
        s = s & "DXUI=0; "
    #End If

    GetBuildFlags = Trim$(s)
End Function

