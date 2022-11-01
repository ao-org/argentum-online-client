Attribute VB_Name = "Carteles"
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

'Carteles
Public cartel    As Boolean

Public Leyenda   As String

Public GrhCartel As Integer

Sub InitCartel(Ley As String, grh As Integer)
    
    On Error GoTo InitCartel_Err
    

    If Not cartel Then
        Leyenda = Ley
        GrhCartel = grh
        cartel = True
    Else
        Exit Sub

    End If

    
    Exit Sub

InitCartel_Err:
    Call RegistrarError(Err.number, Err.Description, "Carteles.InitCartel", Erl)
    Resume Next
    
End Sub

