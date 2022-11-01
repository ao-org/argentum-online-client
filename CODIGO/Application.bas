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
End Type: Private HistorialError As UltimoError

''
' Checks if this is the active (foreground) application or not.
'
' @return   True if any of the app's windows are the foreground window, false otherwise.

Public Function IsAppActive() As Boolean
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (maraxus)
    'Last Modify Date: 03/03/2007
    'Checks if this is the active application or not
    '***************************************************
    
    On Error GoTo IsAppActive_Err
    
    IsAppActive = (GetActiveWindow <> 0)

    
    Exit Function

IsAppActive_Err:
    Call RegistrarError(Err.Number, Err.Description, "Application.IsAppActive", Erl)
    Resume Next
    
End Function

Public Sub RegistrarError(ByVal Numero As Long, ByVal Descripcion As String, ByVal Componente As String, Optional ByVal Linea As Integer)
    '**********************************************************
    'Author: Jopi
    'Guarda una descripcion detallada del error en Errores.log
    '**********************************************************
        
        On Error GoTo RegistrarError_Err
    
        
        
        'Si lo del parametro Componente es ES IGUAL, al Componente del anterior error...
100     If Componente = HistorialError.Componente And _
           Numero = HistorialError.ErrorCode Then
       
           'Si ya recibimos error en el mismo componente 10 veces, es bastante probable que estemos en un bucle
            'x lo que no hace falta registrar el error.
102        ' If HistorialError.Contador = 10 Then
           '     Debug.Print "Mismo error"
           '     Debug.Assert False
           '     Exit Sub
           ' End If
        
            'Agregamos el error al historial.
104         HistorialError.Contador = HistorialError.Contador + 1
        
        Else 'Si NO es igual, reestablecemos el contador.

106         HistorialError.Contador = 0
108         HistorialError.ErrorCode = Numero
110         HistorialError.Componente = Componente
            
        End If
    
        'Registramos el error en Errores.log
112     Dim File As Integer: File = FreeFile
        
114     Open App.Path & "\logs\Errores.log" For Append As #File
    
116         Print #File, "Error: " & Numero
118         Print #File, "Descripcion: " & Descripcion
        
120         Print #File, "Componente: " & Componente

122         If LenB(Linea) <> 0 Then
124             Print #File, "Linea: " & Linea
            End If

126         Print #File, "Fecha y Hora: " & Date$ & "-" & Time$
        
128         Print #File, vbNullString
        
130     Close #File
    
132     Debug.Print "Error: " & Numero & vbNewLine & _
                    "Descripcion: " & Descripcion & vbNewLine & _
                    "Componente: " & Componente & vbNewLine & _
                    "Linea: " & Linea & vbNewLine & _
                    "Fecha y Hora: " & Date$ & "-" & Time$ & vbNewLine
        
        Exit Sub

RegistrarError_Err:
        Call RegistrarError(Err.Number, Err.Description, "ES.RegistrarError", Erl)

        
End Sub

