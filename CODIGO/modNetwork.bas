Attribute VB_Name = "modNetwork"
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

#If DIRECT_PLAY = 0 Then
Private Client As Network.Client

Private Type t_FailedIp
    IP As String
    Port As String
End Type

Dim FailedIpList(10) As t_FailedIp
Public FailedListSize As Integer

#If PYMMO = 0 Then
Public Function IsConnected() As Boolean
    On Error Goto IsConnected_Err
    IsConnected = Connected
    Exit Function
IsConnected_Err:
    Call TraceError(Err.Number, Err.Description, "modNetwork.IsConnected", Erl)
End Function
#End If

Public Function IsFailedString(ByVal IP As String, ByVal Port As String)
    On Error Goto IsFailedString_Err
    Dim i As Integer
    For i = 0 To FailedListSize - 1
        If FailedIpList(i).IP = IP And FailedIpList(i).Port = Port Then
            IsFailedString = True
            Exit Function
        End If
    Next i
    Exit Function
IsFailedString_Err:
    Call TraceError(Err.Number, Err.Description, "modNetwork.IsFailedString", Erl)
End Function

Public Sub AddFailedIp(ByVal IP As String, ByVal Port As String)
    On Error Goto AddFailedIp_Err
    FailedIpList(FailedListSize).IP = IP
    FailedIpList(FailedListSize).Port = Port
    FailedListSize = FailedListSize + 1
    Exit Sub
AddFailedIp_Err:
    Call TraceError(Err.Number, Err.Description, "modNetwork.AddFailedIp", Erl)
End Sub

Public Sub Connect(ByVal Address As String, ByVal Service As String)
    On Error Goto Connect_Err
    frmDebug.add_text_tracebox "Connecting to World Server : " & Address & ":" & Service

    If (Address = vbNullString Or Service = vbNullString) Then
        Exit Sub
    End If
    Call Protocol_Writes.Initialize
    
    Set Client = New Network.Client
    Call Client.Attach(AddressOf OnClientConnect, AddressOf OnClientClose, AddressOf OnClientSend, AddressOf OnClientRecv)
    Call Client.Connect(Address, Service)
    Exit Sub
Connect_Err:
    Call TraceError(Err.Number, Err.Description, "modNetwork.Connect", Erl)
End Sub

Public Sub Disconnect()
    On Error Goto Disconnect_Err
If Not Client Is Nothing Then
    Call Client.Close(True)
    Unload frmConnecting
End If
    Exit Sub
Disconnect_Err:
    Call TraceError(Err.Number, Err.Description, "modNetwork.Disconnect", Erl)
End Sub

Public Sub Poll()
    On Error Goto Poll_Err
    If (Client Is Nothing) Then
        Exit Sub
    End If
    GetRemoteError = False
    Call Client.Flush
    Call Client.Poll
    Exit Sub
Poll_Err:
    Call TraceError(Err.Number, Err.Description, "modNetwork.Poll", Erl)
End Sub

Public Sub Send(ByVal Buffer As Network.Writer)
    On Error Goto Send_Err
    If (Connected) Then
        Call Client.Send(False, Buffer)
    End If
    
    Call Buffer.Clear
    Exit Sub
Send_Err:
    Call TraceError(Err.Number, Err.Description, "modNetwork.Send", Erl)
End Sub

Public Sub RetryWithAnotherIp()
    On Error Goto RetryWithAnotherIp_Err
    Call Disconnect
    Call AddFailedIp(IPdelServidor, PuertoDelServidor)
    If FailedListSize = ServerIpCount Then
        Unload frmConnecting
        Exit Sub
    Else
        Do While IsFailedString(IPdelServidor, PuertoDelServidor)
            Call SetDefaultServer
        Loop
    End If
    
    Call modNetwork.Connect(IPdelServidor, PuertoDelServidor)
    Exit Sub
RetryWithAnotherIp_Err:
    Call TraceError(Err.Number, Err.Description, "modNetwork.RetryWithAnotherIp", Erl)
End Sub

#If PYMMO = 1 Then
Private Sub OnClientConnect()
    On Error Goto OnClientConnect_Err
On Error GoTo OnClientConnect_Err:
frmDebug.add_text_tracebox ("Entró OnClientConnect")

If EstadoLogin = E_MODO.CrearNuevoPj Then
    Call LoginOrConnect(E_MODO.CrearNuevoPj)
End If

    Unload frmConnecting
    Connected = True
    Exit Sub
    
OnClientConnect_Err:
    Call RegistrarError(Err.Number, Err.Description, "modNetwork.OnClientConnect", Erl)
    Exit Sub
OnClientConnect_Err:
    Call TraceError(Err.Number, Err.Description, "modNetwork.OnClientConnect", Erl)
End Sub
#ElseIf PYMMO = 0 Then
    
Private Sub OnClientConnect()
    On Error Goto OnClientConnect_Err
On Error GoTo OnClientConnect_Err:
frmDebug.add_text_tracebox ("Entró OnClientConnect")

    Connected = True
    Unload frmConnecting
    Exit Sub
    
OnClientConnect_Err:
    Call RegistrarError(Err.Number, Err.Description, "modNetwork.OnClientConnect", Erl)
    Exit Sub
OnClientConnect_Err:
    Call TraceError(Err.Number, Err.Description, "modNetwork.OnClientConnect", Erl)
End Sub
#End If

Private Sub OnClientClose(ByVal Code As Long)
    On Error Goto OnClientClose_Err
On Error GoTo OnClientClose_Err:
    
    Call Protocol_Writes.Clear

    Call ModLogin.OnClientDisconnect(Code)

    Exit Sub
    
OnClientClose_Err:
    Call RegistrarError(Err.Number, Err.Description, "modNetwork.OnClientClose", Erl)
    Exit Sub
OnClientClose_Err:
    Call TraceError(Err.Number, Err.Description, "modNetwork.OnClientClose", Erl)
End Sub

Private Sub OnClientSend(ByVal Message As Network.Reader)
    On Error Goto OnClientSend_Err
On Error GoTo OnClientSend_Err:
    Exit Sub
    
OnClientSend_Err:
    Call RegistrarError(Err.Number, Err.Description, "modNetwork.OnClientSend", Erl)
    Exit Sub
OnClientSend_Err:
    Call TraceError(Err.Number, Err.Description, "modNetwork.OnClientSend", Erl)
End Sub

Private Sub OnClientRecv(ByVal Message As Network.Reader)
    On Error Goto OnClientRecv_Err
On Error GoTo OnClientRecv_Err:

    Call Protocol.HandleIncomingData(Message)

    Exit Sub
    
OnClientRecv_Err:
    Call RegistrarError(Err.Number, Err.Description, "modNetwork.OnClientRecv", Erl)
    Exit Sub
OnClientRecv_Err:
    Call TraceError(Err.Number, Err.Description, "modNetwork.OnClientRecv", Erl)
End Sub

#Else
'DirectPlay


Public Sub DoSleep(Optional ByVal lMilliSec As Long = 0)
    On Error Goto DoSleep_Err
    'The DoSleep function allows other threads to have a time slice
    'and still keeps the main VB thread alive (since DPlay callbacks
    'run on separate threads outside of VB).
    Sleep lMilliSec
    DoEvents
    Exit Sub
DoSleep_Err:
    Call TraceError(Err.Number, Err.Description, "modNetwork.DoSleep", Erl)
End Sub

Public Sub Poll()
    On Error Goto Poll_Err
' Not needed when using DPLAY
    Exit Sub
Poll_Err:
    Call TraceError(Err.Number, Err.Description, "modNetwork.Poll", Erl)
End Sub

Public Sub OnClientDisconnect(dpnotify As DxVBLibA.DPNMSG_TERMINATE_SESSION, fRejectMsg As Boolean)
    On Error Goto OnClientDisconnect_Err
On Error GoTo OnClientDisconnect_Err:
    frmDebug.add_text_tracebox "DPLAY: OnClientDisconnect"
    Err.Clear
    Connected = False
    
    
    Call ModLogin.OnClientDisconnect(0)
    Call LogOut
    Call Protocol_Writes.Clear
    
    Exit Sub
OnClientDisconnect_Err:
   If Err.Number <> 0 Then
        Call HandleDPlayError(Err.Number, Err.Description, "modnetwork.OnClientDisconnect", Erl)
    End If
    Exit Sub
OnClientDisconnect_Err:
    Call TraceError(Err.Number, Err.Description, "modNetwork.OnClientDisconnect", Erl)
End Sub
    
    
    
Public Sub OnClientConnect(dpnotify As DxVBLibA.DPNMSG_CONNECT_COMPLETE, fRejectMsg As Boolean)
    On Error Goto OnClientConnect_Err
On Error GoTo OnClientConnect_Err:
    frmDebug.add_text_tracebox "DPLAY: OnClientConnect"
    Err.Clear
    
    If EstadoLogin = E_MODO.CrearNuevoPj Then
        Call LoginOrConnect(E_MODO.CrearNuevoPj)
    End If

    
    Connected = True
    Unload frmConnecting
    Exit Sub
OnClientConnect_Err:
   If Err.Number <> 0 Then
        Call HandleDPlayError(Err.Number, Err.Description, "modnetwork.Receive", Erl)
    End If
    Exit Sub
OnClientConnect_Err:
    Call TraceError(Err.Number, Err.Description, "modNetwork.OnClientConnect", Erl)
End Sub

Public Sub Send(ByVal Buffer As clsNetWriter)
    On Error Goto Send_Err
On Error GoTo send_Err:
    Err.Clear
    Writer.send
    Exit Sub
send_Err:
   If Err.Number <> 0 Then
        Call HandleDPlayError(Err.Number, Err.Description, "modnetwork.Receive", Erl)
    End If
    Exit Sub
Send_Err:
    Call TraceError(Err.Number, Err.Description, "modNetwork.Send", Erl)
End Sub

Public Sub Disconnect()
    On Error Goto Disconnect_Err
    Call Protocol_Writes.Clear
    Set Protocol_Writes.Writer = Nothing
    Set Protocol_Writes.Writer = New clsNetWriter
    If Connected Then
        frmDebug.add_text_tracebox "Disconnecting DirecPlay..."
        Connected = False
        DoEvents
        modDplayClient.dpc.Close 0
        Set dpc = Nothing
        frmDebug.add_text_tracebox "Disconnected DirectPlay"
    End If
    
    Exit Sub
Disconnect_Err:
    Call TraceError(Err.Number, Err.Description, "modNetwork.Disconnect", Erl)
End Sub

Public Sub Receive(dpnotify As DxVBLibA.DPNMSG_RECEIVE, fRejectMsg As Boolean)
    On Error Goto Receive_Err
On Error GoTo receive_error:
    Err.Clear
    Call Protocol.HandleIncomingData(dpnotify)
    Exit Sub
receive_error:
   If Err.Number <> 0 Then
        Call HandleDPlayError(Err.Number, Err.Description, "modnetwork.Receive", Erl)
    End If
    Exit Sub
Receive_Err:
    Call TraceError(Err.Number, Err.Description, "modNetwork.Receive", Erl)
End Sub
Public Sub Connect(ByVal Address As String, ByVal Service As String)
    On Error Goto Connect_Err
On Error GoTo connect_error:
    Err.Clear
    
    frmDebug.add_text_tracebox "DPLAY > Connecting to World Server : " & Address & ":" & Service
    

    If (Address = vbNullString Or Service = vbNullString) Then
        Exit Sub
    End If
    
    Dim HostAddr As DirectPlay8Address
    Dim DeviceAddr As DirectPlay8Address
    
    
    
    Err.Clear
    Set HostAddr = DirectX.DirectPlayAddressCreate
    HostAddr.SetSP DP8SP_TCPIP  ' Set the service provider to TCP/IP
    HostAddr.AddComponentLong DPN_KEY_PORT, CLng(Service)
    HostAddr.AddComponentString DPN_KEY_HOSTNAME, Address
    Debug.Assert Err.Number = 0
    Err.Clear
    
    Dim connect_handle As Long
    connect_handle = dpc.Connect(dpApp, HostAddr, DeviceAddr, DPNCONNECT_OKTOQUERYFORADDRESSING, ByVal 0&, 0)
    
    
    If Err.Number <> 0 Then
        Call HandleDPlayError(Err.Number, Err.Description, "modnetwork.Connect", Erl)
    End If
    
    Dim i As Integer
    i = 1
    Do While Not frmConnect.mfGotEvent 'Let's wait for our connectcomplete event
        DoSleep 1000 'Give other threads cpu time
        frmDebug.add_text_tracebox "Trying to connect DPLAY server " & i & "..."
        i = i + 1
    Loop
    Connected = True
    If frmConnect.mfConnectComplete Then
        'We've joined our game
        Set HostAddr = Nothing
        Set DeviceAddr = Nothing
    End If
    Exit Sub
connect_error:
   If Err.Number <> 0 Then
        Call HandleDPlayError(Err.Number, Err.Description, "modnetwork.Receive", Erl)
    End If
    Exit Sub
Connect_Err:
    Call TraceError(Err.Number, Err.Description, "modNetwork.Connect", Erl)
End Sub

#End If
