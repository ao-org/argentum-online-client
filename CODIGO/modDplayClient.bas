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
Public dpApp As DPN_APPLICATION_DESC

Public Sub init_direct_play(ByRef dx As DirectX8)
    On Error Goto init_direct_play_Err
    Err.Clear
    CheckAndEnableDirectPlay
    
    Debug.Assert dpc Is Nothing
    Debug.Assert Not dx Is Nothing
    
    Set dpc = dx.DirectPlayClientCreate
    Debug.Assert Err.Number = 0
    frmDebug.add_text_tracebox ("dX.DirectPlayClientCreate OK!")
    
    dpc.RegisterMessageHandler frmConnect
    Set Protocol_Writes.Writer = New clsNetWriter
    
    Dim pInfo As DPN_PLAYER_INFO
    pInfo.Name = "Pablo"
    pInfo.lInfoFlags = DPNINFO_NAME
    dpc.SetClientInfo pInfo
    
    Dim scaps As DPN_SP_CAPS
    scaps = dpc.GetSPCaps(DP8SP_TCPIP)
    
    
     With scaps
        .lBuffersPerThread = 16
        frmDebug.add_text_tracebox ("DPLAY_SP_CAPS:lBuffersPerThread :" & .lBuffersPerThread)
        frmDebug.txtBuffersPerThread.Text = .lBuffersPerThread
        
        frmDebug.add_text_tracebox ("DPLAY_SP_CAPS:lDefaultEnumRetryInterval :" & .lDefaultEnumRetryInterval)
        frmDebug.txtDefaultEnumRetryInterval.Text = .lDefaultEnumRetryInterval
        
        frmDebug.add_text_tracebox ("DPLAY_SP_CAPS:lDefaultEnumTimeout :" & .lDefaultEnumTimeout)
        frmDebug.txtDefaultEnumTimeout.Text = .lDefaultEnumTimeout
        
        frmDebug.add_text_tracebox ("DPLAY_SP_CAPS:lSystemBufferSize :" & .lSystemBufferSize)
        frmDebug.txtSystemBufferSize.Text = .lSystemBufferSize
        
        frmDebug.add_text_tracebox ("DPLAY_SP_CAPS:lNumThreads :" & .lNumThreads)
        frmDebug.txtNumThreads.Text = .lNumThreads
      
    End With
    dpc.SetSPCaps DP8SP_TCPIP, scaps
    Exit Sub
init_direct_play_Err:
    Call TraceError(Err.Number, Err.Description, "modDplayClient.init_direct_play", Erl)
End Sub

Public Sub shutdown_direct_play()
    On Error Goto shutdown_direct_play_Err
    'Stop our message handler
    If Not dpc Is Nothing Then dpc.UnRegisterMessageHandler
    'Close down our session
    If Not dpc Is Nothing Then dpc.Close
    Set dpc = Nothing
    'Get rid of our message pump
    'DPlayEventsForm.Un
    Exit Sub
shutdown_direct_play_Err:
    Call TraceError(Err.Number, Err.Description, "modDplayClient.shutdown_direct_play", Erl)
End Sub


Private Sub CheckAndEnableDirectPlay()
    On Error Goto CheckAndEnableDirectPlay_Err
    If Not IsDirectPlayEnabled() Then
        Dim response As VbMsgBoxResult
        response = MsgBox("DirectPlay is not enabled. Would you like to enable it now?", vbYesNo + vbQuestion, "Enable DirectPlay")
        If response = vbYes Then
            EnableDirectPlay
        Else
            MsgBox "DirectPlay-dependent features may not function correctly.", vbExclamation, "DirectPlay Not Enabled"
        End If
    Else
        frmDebug.add_text_tracebox "DirectPlay Status: DirectPlay is already enabled."
    End If
    Exit Sub
CheckAndEnableDirectPlay_Err:
    Call TraceError(Err.Number, Err.Description, "modDplayClient.CheckAndEnableDirectPlay", Erl)
End Sub

Private Function IsDirectPlayEnabled() As Boolean
    On Error Goto IsDirectPlayEnabled_Err
    On Error Resume Next
    Dim objWMIService As Object
    Dim colFeatures As Object
    Dim objFeature As Variant
    Dim featureName As String
    featureName = "DirectPlay"
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colFeatures = objWMIService.ExecQuery("Select * from Win32_OptionalFeature where Name = '" & featureName & "'")
    For Each objFeature In colFeatures
        If objFeature.Name = featureName And objFeature.InstallState = 1 Then
            IsDirectPlayEnabled = True
            Exit Function
        End If
    Next
    IsDirectPlayEnabled = False
    Exit Function
IsDirectPlayEnabled_Err:
    Call TraceError(Err.Number, Err.Description, "modDplayClient.IsDirectPlayEnabled", Erl)
End Function

Private Sub EnableDirectPlay()
    On Error Goto EnableDirectPlay_Err
    On Error Resume Next
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    shell.Run "dism /online /enable-feature /featurename:DirectPlay /all", 0, True
    MsgBox "DirectPlay has been enabled. Please restart the application.", vbInformation, "DirectPlay Enabled"
    frmDebug.add_text_tracebox "DirectPlay has been enabled. Please restart the application."
    Exit Sub
EnableDirectPlay_Err:
    Call TraceError(Err.Number, Err.Description, "modDplayClient.EnableDirectPlay", Erl)
End Sub


Public Sub HandleDPlayError(ByVal ErrNumber As Long, ByVal ErrDescription As String, ByVal place As String, ByVal line As String)
    On Error Goto HandleDPlayError_Err
       Select Case err.Number
            Case DPNERR_INVALIDPLAYER
                    Call LogError("DPNERR_INVALIDPLAYER: The player ID is not recognized as a valid player ID for this game session. " & place & " " & line)
            Case DPNERR_INVALIDPARAM
                    Call LogError("DPNERR_INVALIDPARAM: One or more of the parameters passed to the method are invalid." & place & " " & line)
            Case DPNERR_NOTHOST:
                    Call LogError("DPNERR_NOTHOST: The client attempted to connect to a nonhost computer. Additionally, this error value may be returned by a nonhost that tried to set the application description. " & place & " " & line)
            Case DPNERR_INVALIDFLAGS
                    Call LogError("DPNERR_INVALIDFLAGS: The flags passed to this method are invalid. " & place & " " & line)
            Case DPNERR_TIMEDOUT
                    Call LogError("DPNERR_TIMEDOUT: The operation could not complete because it has timed out. " & place & " " & line)
            Case DPNERR_NOCONNECTION:
                    Call LogError("DPNERR_NOCONNECTION " & place & " " & line)
            Case DPNERR_INVALIDPASSWORD
                    Call LogError("DPNERR_INVALIDPASSWORD " & place & " " & line)
            Case DPNERR_INVALIDINTERFACE
                    Call LogError("DPNERR_INVALIDINTERFACE " & place & " " & line)
            Case DPNERR_INVALIDAPPLICATION
                    Call LogError("DPNERR_INVALIDAPPLICATION " & place & " " & line)
            Case DPNERR_NOTHOST
                    Call LogError("DPNERR_NOTHOST " & place & " " & line)
            Case DPNERR_SESSIONFULL
                    Call LogError("DPNERR_SESSIONFULL " & place & " " & line)
            Case DPNERR_HOSTREJECTEDCONNECTION
                    Call LogError("DPNERR_HOSTREJECTEDCONNECTION " & place & " " & line)
            Case DPNERR_INVALIDINSTANCE
                    Call LogError("DPNERR_INVALIDINSTANCE " & place & " " & line)
                   
            Case Else
                    Call LogError("Unknown error " & Err.Number & " " & place & " " & line)
        End Select
        err.Clear
    Exit Sub
HandleDPlayError_Err:
    Call TraceError(Err.Number, Err.Description, "modDplayClient.HandleDPlayError", Erl)
End Sub


#End If
