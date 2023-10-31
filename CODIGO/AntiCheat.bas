Attribute VB_Name = "AntiCheat"
Option Explicit

Public Enum e_ACInitResult
    eOk = 0
    eFailedPlatform
    eFAiledConnectAC
End Enum

Public Type t_AntiCheatCallbacks
    SendToServer As Long
    LogMessage As Long
End Type

Private Declare Function InitializeAC Lib "AOACClient.dll" (ByRef Callbacks As t_AntiCheatCallbacks) As Long
Private Declare Sub UnloadAC Lib "AOACClient.dll" ()
Private Declare Sub Update Lib "AOACClient.dll" ()
Private Declare Sub BeginSession Lib "AOACClient.dll" (ByRef userName As String)
Private Declare Sub EndSession Lib "AOACClient.dll" ()
Public Declare Sub HandleRemoteMessage Lib "AOACClient.dll" (ByRef Data As Byte, ByVal DataSize As Integer)


Public Sub InitializeAntiCheat()
On Error GoTo InitializeAC_Err
    Dim InitResult As e_ACInitResult
    Dim Callbacks As t_AntiCheatCallbacks
    Callbacks.SendToServer = FARPROC(AddressOf SendToServerCB)
    Callbacks.LogMessage = FARPROC(AddressOf LogMessageCB)
    InitResult = InitializeAC(Callbacks)
    If InitResult <> eOk Then
        Call DisplayError("El juego se inicio sin activar el anti cheat, debe activarlo para poder conectarse a los servidores.", "")
    End If
    Exit Sub
InitializeAC_Err:
    Call RegistrarError(Err.Number, Err.Description, "AOAC.InitializeAC", Erl)
End Sub

Public Sub UnloadAntiCheat()
    Call UnloadAC
End Sub

Public Sub SendToServerCB(ByVal Data As Long, ByVal DataSize As Long)
    Call WriteAntiCheatMessage(Data, DataSize)
End Sub

Public Sub LogMessageCB(ByRef message As SINGLESTRINGPARAM)
    Dim MessageStr As String
    If message.Len > 0 Then
        MessageStr = GetStringFromPtr(message.Ptr, message.Len)
    End If
    Call LogError(MessageStr)
End Sub

Public Sub HandleAntiCheatServerMessage(ByRef Data() As Byte)
    Call HandleRemoteMessage(Data(0), UBound(Data))
End Sub
Public Sub UpdateAntiCheat()
   Call Update
End Sub

Public Sub BeginAntiCheatSession()
    Call BeginSession(userName)
End Sub

Public Sub EndAntiCheatSession()
    Call EndSession
End Sub
