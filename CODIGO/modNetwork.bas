Attribute VB_Name = "modNetwork"
Option Explicit

Private Client As Network.Client

Public Sub Connect(ByVal Address As String, ByVal Service As String)
    If (Address = vbNullString Or Service = vbNullString) Then
        Exit Sub
    End If
    
    Call Protocol_Writes.Initialize
    
    Set Client = New Network.Client
    Call Client.Attach(AddressOf OnClientConnect, AddressOf OnClientClose, AddressOf OnClientSend, AddressOf OnClientRecv)
    Call Client.Connect(Address, Service)
End Sub

Public Sub Disconnect()
    Call Client.Close(True)
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

Private Sub OnClientConnect()
On Error GoTo OnClientConnect_Err:

#If AntiExternos = 1 Then
    XorIndexIn = 0
    XorIndexOut = 0
#End If

    Connected = True
    
    Exit Sub
    
OnClientConnect_Err:
    Call RegistrarError(Err.Number, Err.Description, "modNetwork.OnClientConnect", Erl)
End Sub

Private Sub OnClientClose(ByVal Code As Long)
On Error GoTo OnClientClose_Err:
    
    Call Protocol_Writes.Clear
    
    Call frmMain.OnClientDisconnect(Code <> 0)
    
    Connected = False
    
    Exit Sub
    
OnClientClose_Err:
    Call RegistrarError(Err.Number, Err.Description, "modNetwork.OnClientClose", Erl)
End Sub

Private Sub OnClientSend(ByVal Message As Network.Reader)
On Error GoTo OnClientSend_Err:

    Dim BytesRef() As Byte
    Call Message.GetData(BytesRef) ' Is only a view of the buffer as a SafeArrayPtr ;-)

    #If AntiExternos = 1 Then
        Call Security.XorData(BytesRef, UBound(BytesRef) - 1, XorIndexOut)
    #End If
    
    Exit Sub
    
OnClientSend_Err:
    Call RegistrarError(Err.Number, Err.Description, "modNetwork.OnClientSend", Erl)
End Sub

Private Sub OnClientRecv(ByVal Message As Network.Reader)
On Error GoTo OnClientRecv_Err:

    Dim BytesRef() As Byte
    Call Message.GetData(BytesRef) ' Is only a view of the buffer as a SafeArrayPtr ;-)

    #If AntiExternos = 1 Then
        Call Security.XorData(BytesRef, UBound(BytesRef) - 1, XorIndexIn)
    #End If

    Call Protocol.HandleIncomingData(Message)

    Exit Sub
    
OnClientRecv_Err:
    Call RegistrarError(Err.Number, Err.Description, "modNetwork.OnClientRecv", Erl)
End Sub


