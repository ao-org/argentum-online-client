Attribute VB_Name = "PrevInstance"
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

'Declaration of the Win32 API function for creating /destroying a Mutex, and some types and constants.
Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByRef lpMutexAttributes As SECURITY_ATTRIBUTES, ByVal bInitialOwner As Long, ByVal lpName As String) As Long

Private Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Type SECURITY_ATTRIBUTES

    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long

End Type

Private Const ERROR_ALREADY_EXISTS = 183&

Private mutexHID As Long

''
' Creates a Named Mutex. Private function, since we will use it just to check if a previous instance of the app is running.
'
' @param mutexName The name of the mutex, should be universally unique for the mutex to be created.

Private Function CreateNamedMutex(ByRef mutexName As String) As Boolean
    
    On Error GoTo CreateNamedMutex_Err
    

    '***************************************************
    'Author: Fredy Horacio Treboux (liquid)
    'Last Modification: 01/04/07
    'Last Modified by: Juan Martín Sotuyo Dodero (Maraxus) - Changed Security Atributes to make it work in all OS
    '***************************************************
    Dim sa As SECURITY_ATTRIBUTES
    
    With sa
        .bInheritHandle = 0
        .lpSecurityDescriptor = 0
        .nLength = LenB(sa)

    End With
    
    mutexHID = CreateMutex(sa, False, "Global\" & mutexName)
    
    CreateNamedMutex = Not (Err.LastDllError = ERROR_ALREADY_EXISTS) 'check if the mutex already existed

    
    Exit Function

CreateNamedMutex_Err:
    Call RegistrarError(Err.Number, Err.Description, "PrevInstance.CreateNamedMutex", Erl)
    Resume Next
    
End Function

''
' Checks if there's another instance of the app running, returns True if there is or False otherwise.

Public Function FindPreviousInstance() As Boolean
    
    On Error GoTo FindPreviousInstance_Err
    

    '***************************************************
    'Author: Fredy Horacio Treboux (liquid)
    'Last Modification: 01/04/07
    '
    '***************************************************
    'We try to create a mutex, the name could be anything, but must contain no backslashes.
    If CreateNamedMutex("EnVezDeChitearAyudanos!HablanosAlDiscord") Then
        'There's no other instance running
        FindPreviousInstance = False
    Else
        'There's another instance running
        FindPreviousInstance = True

    End If

    
    Exit Function

FindPreviousInstance_Err:
    Call RegistrarError(Err.Number, Err.Description, "PrevInstance.FindPreviousInstance", Erl)
    Resume Next
    
End Function

''
' Closes the client, allowing other instances to be open.

Public Sub ReleaseInstance()
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 01/04/07
    '
    '***************************************************
    
    On Error GoTo ReleaseInstance_Err
    
    Call ReleaseMutex(mutexHID)
    Call CloseHandle(mutexHID)

    
    Exit Sub

ReleaseInstance_Err:
    Call RegistrarError(Err.Number, Err.Description, "PrevInstance.ReleaseInstance", Erl)
    Resume Next
    
End Sub
