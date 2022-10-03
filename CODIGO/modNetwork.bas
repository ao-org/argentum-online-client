Attribute VB_Name = "modNetwork"
Option Explicit

Private Client As Network.Client
#If PYMMO = 0 Then
Public Function IsConnected() As Boolean
    IsConnected = Connected
End Function
#End If

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
End If
End Sub

Public Sub Poll()
    If (Client Is Nothing) Then
        Exit Sub
    End If
    
    Call Client.Flush
    Call Client.Poll
End Sub

Public Sub Send(ByVal Buffer As Network.Writer)
    If (Connected) Then
        Call Client.Send(False, Buffer)
    End If
    
    Call Buffer.Clear
End Sub
#If PYMMO = 1 Then
Private Sub OnClientConnect()
On Error GoTo OnClientConnect_Err:
Debug.Print ("Entró OnClientConnect")

If EstadoLogin = E_MODO.CrearNuevoPj Then
    Call LoginOrConnect(E_MODO.CrearNuevoPj)
End If

   
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
    
    Exit Sub
    
OnClientConnect_Err:
    Call RegistrarError(Err.Number, Err.Description, "modNetwork.OnClientConnect", Erl)
End Sub
#End If

Private Sub OnClientClose(ByVal Code As Long)
On Error GoTo OnClientClose_Err:
    
    Call Protocol_Writes.Clear

    Call frmMain.OnClientDisconnect(Code)

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


