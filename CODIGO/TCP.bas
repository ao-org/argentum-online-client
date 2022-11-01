Attribute VB_Name = "Mod_TCP"
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

Public Warping        As Boolean

Public LlegaronSkills As Boolean
Public LlegaronStats  As Boolean
Public LlegaronAtrib  As Boolean

Public Function PuedoQuitarFoco() As Boolean
    
    On Error GoTo PuedoQuitarFoco_Err
    
    PuedoQuitarFoco = True
    
    Exit Function

PuedoQuitarFoco_Err:
    Call RegistrarError(Err.number, Err.Description, "Mod_TCP.PuedoQuitarFoco", Erl)
    Resume Next
    
End Function
#If PYMMO = 1 Then
Sub LoginOrConnect(ByVal Modo As E_MODO)
    EstadoLogin = Modo
    
    If Auth_state = e_state.AccountLogged Then
        Call Login
    Else
        Call connectToLoginServer
    End If
  
End Sub
#ElseIf PYMMO = 0 Then
Sub LoginOrConnect(ByVal Modo As E_MODO)
    
    EstadoLogin = Modo
    
    If (Not modNetwork.IsConnected) Then
        Call modNetwork.Connect(IPdelServidor, PuertoDelServidor)
    Else
        Call Login
    End If

End Sub
#End If

Sub Login()
    
    On Error GoTo Login_Err
    
    Select Case EstadoLogin
    
        Case E_MODO.Normal
            Call WriteLoginExistingChar
        
        Case E_MODO.CrearNuevoPj
            Call WriteLoginNewChar
#If PYMMO = 0 Then
        Case E_MODO.IngresandoConCuenta
            Call WriteLoginAccount
        
        Case E_MODO.CreandoCuenta
            Call WriteCreateAccount
#End If
    End Select

    Exit Sub

Login_Err:
    Call RegistrarError(Err.number, Err.Description, "Mod_TCP.Login", Erl)
    Resume Next
    
End Sub
