Attribute VB_Name = "Application"
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
Private Declare Function GetActiveWindow Lib "user32" () As Long

Private Type UltimoError
    Componente As String
    Contador As Byte
    ErrorCode As Long
End Type

''
' Checks if this is the active (foreground) application or not.
'
' @return   True if any of the app's windows are the foreground window, false otherwise.
Public Function IsAppActive() As Boolean
    'Checks if this is the active application or not
    On Error GoTo IsAppActive_Err
    IsAppActive = (GetActiveWindow <> 0)
    Exit Function
IsAppActive_Err:
    Call RegistrarError(Err.Number, Err.Description, "Application.IsAppActive", Erl)
    Resume Next
End Function

Public Sub DeleteFile(ByVal filename As String)
    On Error Resume Next
    If Len(dir$(filename)) > 0 Then
        Kill filename
    End If
End Sub

Public Sub RegistrarError(ByVal Numero As Long, ByVal Descripcion As String, ByVal Componente As String, Optional ByVal Linea As Integer)
    On Error GoTo RegistrarError_Err
    Dim File As Integer: File = FreeFile
    Open GetErrorLogFilename() For Append As #File
    Print #File, "Error: " & Numero
    Print #File, "Descripcion: " & Descripcion
    Print #File, "Componente: " & Componente
    If LenB(Linea) <> 0 Then
        Print #File, "Linea: " & Linea
    End If
    Print #File, "Fecha y Hora: " & Date$ & "-" & Time$
    Print #File, vbNullString
    Close #File
    frmDebug.add_text_tracebox "Error: " & Numero & vbNewLine & "Descripcion: " & Descripcion & vbNewLine & "Componente: " & Componente & vbNewLine & "Linea: " & Linea & _
            vbNewLine & "Fecha y Hora: " & Date$ & "-" & Time$ & vbNewLine
    Exit Sub
RegistrarError_Err:
    Call RegistrarError(Err.Number, Err.Description, "ES.RegistrarError", Erl)
End Sub
