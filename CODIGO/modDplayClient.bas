Attribute VB_Name = "modDplayClient"
'    Argentum 20 - Game Client Program
'    Copyright (C) 2025 - Noland Studios
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
#If DIRECT_PLAY = 1 Then
Option Explicit

Public Const AppGuid = "{5726CF1F-702B-4008-98BC-BF9C95F9E288}"

Public dpc As DirectPlay8Client
Public DPlayEventsForm As frmConnect


Public Sub init_direct_play(ByRef dx As DirectX8)
    Err.Clear
    Debug.Assert dpc Is Nothing
    Debug.Assert Not dx Is Nothing
    Set dpc = dx.DirectPlayClientCreate
    Debug.Assert Err.Number = 0
    dpc.RegisterMessageHandler frmConnect
    Set Protocol_Writes.Writer = New clsNetWriter
    
End Sub

Public Sub shutdown_direct_play()

    'Stop our message handler
    If Not dpc Is Nothing Then dpc.UnRegisterMessageHandler
    'Close down our session
    If Not dpc Is Nothing Then dpc.Close
    Set dpc = Nothing
    'Get rid of our message pump
    'DPlayEventsForm.Un
    
End Sub


Public Sub HandleDPlayError(ByVal ErrNumber As Long, ByVal ErrDescription As String, ByVal place As String, ByVal line As String)
       Select Case err.Number
            Case DPNERR_INVALIDPLAYER
                    Call LogError("The player ID is not recognized as a valid player ID for this game session. " & place & " " & line)
            Case DPNERR_INVALIDPARAM
                    Call LogError("One or more of the parameters passed to the method are invalid." & place & " " & line)
            Case DPNERR_NOTHOST:
                    Call LogError("The client attempted to connect to a nonhost computer. Additionally, this error value may be returned by a nonhost that tried to set the application description. " & place & " " & line)
            Case DPNERR_INVALIDFLAGS
                    Call LogError("The flags passed to this method are invalid. " & place & " " & line)
            Case DPNERR_TIMEDOUT
                    Call LogError("The operation could not complete because it has timed out. " & place & " " & line)
            Case Else
                    Call LogError("Unknown error " & place & " " & line)
        End Select
        err.Clear
End Sub


#End If
