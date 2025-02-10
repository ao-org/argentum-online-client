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
    IsConnected = Connected
End Function
#End If

Public Function IsFailedString(ByVal IP As String, ByVal Port As String)
    Dim i As Integer
    For i = 0 To FailedListSize - 1
        If FailedIpList(i).IP = IP And FailedIpList(i).Port = Port Then
            IsFailedString = True
            Exit Function
        End If
    Next i
End Function

Public Sub AddFailedIp(ByVal IP As String, ByVal Port As String)
    FailedIpList(FailedListSize).IP = IP
    FailedIpList(FailedListSize).Port = Port
    FailedListSize = FailedListSize + 1
End Sub

Public Sub Connect(ByVal Address As String, ByVal Service As String)
    Debug.Print "Connecting to World Server : " & Address; ":" & Service

    If (Address = vbNullString Or Service = vbNullString) Then
        Exit Sub
    End If
    Call Protocol_Writes.Initialize
    
    Set Client = New Network.Client
    Call Client.Attach(AddressOf OnClientConnect, AddressOf OnClientClose, AddressOf OnClientSend, AddressOf OnClientRecv)
    Call Client.Connect(Address, Service)
End Sub

Public Sub Disconnect()
If Not Client Is Nothing Then
    Call Client.Close(True)
    Unload frmConnecting
End If
End Sub

Public Sub Poll()
    If (Client Is Nothing) Then
        Exit Sub
    End If
    GetRemoteError = False
    Call Client.Flush
    Call Client.Poll
End Sub

Public Sub Send(ByVal Buffer As Network.Writer)
    If (Connected) Then
        Call Client.Send(False, Buffer)
    End If
    
    Call Buffer.Clear
End Sub

Public Sub RetryWithAnotherIp()
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
End Sub

#If PYMMO = 1 Then
Private Sub OnClientConnect()
On Error GoTo OnClientConnect_Err:
Debug.Print ("Entró OnClientConnect")

If EstadoLogin = E_MODO.CrearNuevoPj Then
    Call LoginOrConnect(E_MODO.CrearNuevoPj)
End If

    Unload frmConnecting
    Connected = True
    Exit Sub
    
OnClientConnect_Err:
    Call RegistrarError(Err.Number, Err.Description, "modNetwork.OnClientConnect", Erl)
End Sub
#ElseIf PYMMO = 0 Then
    
Private Sub OnClientConnect()
On Error GoTo OnClientConnect_Err:
Debug.Print ("Entró OnClientConnect")

    Connected = True
    Unload frmConnecting
    Exit Sub
    
OnClientConnect_Err:
    Call RegistrarError(Err.Number, Err.Description, "modNetwork.OnClientConnect", Erl)
End Sub
#End If

Private Sub OnClientClose(ByVal Code As Long)
On Error GoTo OnClientClose_Err:
    
    Call Protocol_Writes.Clear

    Call ModLogin.OnClientDisconnect(Code)

    Exit Sub
    
OnClientClose_Err:
    Call RegistrarError(Err.Number, Err.Description, "modNetwork.OnClientClose", Erl)
End Sub

Private Sub OnClientSend(ByVal Message As Network.Reader)
On Error GoTo OnClientSend_Err:
    Exit Sub
    
OnClientSend_Err:
    Call RegistrarError(Err.Number, Err.Description, "modNetwork.OnClientSend", Erl)
End Sub

Private Sub OnClientRecv(ByVal Message As Network.Reader)
On Error GoTo OnClientRecv_Err:

    Call Protocol.HandleIncomingData(Message)

    Exit Sub
    
OnClientRecv_Err:
    Call RegistrarError(Err.Number, Err.Description, "modNetwork.OnClientRecv", Erl)
End Sub

#Else
'DirectPlay


Public Sub DoSleep(Optional ByVal lMilliSec As Long = 0)
    'The DoSleep function allows other threads to have a time slice
    'and still keeps the main VB thread alive (since DPlay callbacks
    'run on separate threads outside of VB).
    Sleep lMilliSec
    DoEvents
End Sub

Public Sub Poll()
' Not needed when using DPLAY
End Sub

Public Sub OnClientConnect(dpnotify As DxVBLibA.DPNMSG_CONNECT_COMPLETE, fRejectMsg As Boolean)
On Error GoTo OnClientConnect_Err:
    Connected = True
    
    Unload frmConnecting
    Call Login
    
    Exit Sub
OnClientConnect_Err:
    Call RegistrarError(Err.Number, Err.Description, "modNetwork.OnClientConnect", Erl)
End Sub

Public Sub Send(ByVal Buffer As clsNetWriter)
    Writer.send
End Sub

Public Sub Disconnect()
     modDplayClient.dpc.Close 0
    
End Sub

Public Sub Receive(dpnotify As DxVBLibA.DPNMSG_RECEIVE, fRejectMsg As Boolean)
On Error GoTo receive_error:
    Call Protocol.HandleIncomingData(dpnotify)
    Exit Sub
receive_error:
    Call RegistrarError(Err.Number, Err.Description, "modNetwork.Receive", Erl)
End Sub
Public Sub Connect(ByVal Address As String, ByVal Service As String)
    Debug.Print "DPLAY > Connecting to World Server : " & Address; ":" & Service

    If (Address = vbNullString Or Service = vbNullString) Then
        Exit Sub
    End If
    
    Dim HostAddr As DirectPlay8Address
    Dim DeviceAddr As DirectPlay8Address
    
    Dim dpApp As DPN_APPLICATION_DESC
    
    Dim pInfo As DPN_PLAYER_INFO
    pInfo.Name = "Pablo"
    pInfo.lInfoFlags = DPNINFO_NAME
    
    modDplayClient.dpc.SetClientInfo pInfo, DPNOP_SYNC

    
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
        Select Case Err.Number
            Case DPNERR_NOCONNECTION:
            Case DPNERR_INVALIDPASSWORD:
            Case DPNERR_INVALIDFLAGS:
            Case DPNERR_INVALIDINTERFACE:
            Case DPNERR_INVALIDAPPLICATION:
            Case DPNERR_NOTHOST:
            Case DPNERR_SESSIONFULL:
            Case DPNERR_HOSTREJECTEDCONNECTION:
            Case DPNERR_INVALIDINSTANCE:
            Case Else
                Debug.Print "Connect error " & Err.Number
        End Select
    End If
    
    Do While Not frmConnect.mfGotEvent 'Let's wait for our connectcomplete event
        DoSleep 5 'Give other threads cpu time
    Loop
    Connected = True
    If frmConnect.mfConnectComplete Then
        'We've joined our game
        'mfComplete = True
        'mfHost = False
        'Clean up our address
        Set HostAddr = Nothing
        Set DeviceAddr = Nothing
    End If

End Sub

#End If
