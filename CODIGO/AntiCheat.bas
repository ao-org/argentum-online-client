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

Public Enum EOS_ELogLevel
    EOS_LOG_Off = 0
    EOS_LOG_Fatal = 100
    EOS_LOG_Error = 200
    EOS_LOG_Warning = 300
    EOS_LOG_Info = 400
    EOS_LOG_Verbose = 500
    EOS_LOG_VeryVerbose = 600
End Enum

Private Declare Function InitializeAC Lib "AOACClient.dll" (ByRef Callbacks As t_AntiCheatCallbacks) As Long
Private Declare Sub UnloadAC Lib "AOACClient.dll" ()
Private Declare Sub Update Lib "AOACClient.dll" ()
Private Declare Sub BeginSession Lib "AOACClient.dll" (ByRef userName As String)
Private Declare Sub EndSession Lib "AOACClient.dll" ()
Public Declare Sub HandleRemoteMessage Lib "AOACClient.dll" (ByRef data As Byte, ByVal DataSize As Integer)
Public Declare Function GetCrc32 Lib "AOACClient.dll" (ByVal Path As String) As Long

Public Sub InitializeAntiCheat()
    On Error Goto InitializeAntiCheat_Err
On Error GoTo InitializeAC_Err
#If ENABLE_ANTICHEAT = 1 Then
    Dim InitResult As e_ACInitResult
    Dim Callbacks As t_AntiCheatCallbacks
    Callbacks.SendToServer = FARPROC(AddressOf SendToServerCB)
    Callbacks.LogMessage = FARPROC(AddressOf LogMessageCB)
    InitResult = InitializeAC(Callbacks)
    If InitResult <> eOk Then
        Call DisplayError(JsonLanguage.Item("MENSAJE_ANTICHEAT_NO_ACTIVADO"), "")
    End If
#End If
Exit Sub
InitializeAC_Err:
    Call RegistrarError(Err.Number, Err.Description, "AOAC.InitializeAC", Erl)
    Exit Sub
InitializeAntiCheat_Err:
    Call TraceError(Err.Number, Err.Description, "AntiCheat.InitializeAntiCheat", Erl)
End Sub

Public Sub UnloadAntiCheat()
    On Error Goto UnloadAntiCheat_Err
#If ENABLE_ANTICHEAT = 1 Then
    Call UnloadAC
#End If
    Exit Sub
UnloadAntiCheat_Err:
    Call TraceError(Err.Number, Err.Description, "AntiCheat.UnloadAntiCheat", Erl)
End Sub

Public Sub SendToServerCB(ByVal Data As Long, ByVal DataSize As Long)
    On Error Goto SendToServerCB_Err
    Call WriteAntiCheatMessage(Data, DataSize)
    Exit Sub
SendToServerCB_Err:
    Call TraceError(Err.Number, Err.Description, "AntiCheat.SendToServerCB", Erl)
End Sub

Public Sub LogMessageCB(ByRef message As SINGLESTRINGPARAM, ByVal Level As Long)
    On Error Goto LogMessageCB_Err
    Dim MessageStr As String
    If message.Len > 0 Then
        MessageStr = GetStringFromPtr(message.Ptr, message.Len)
    End If
    Call LogError(MessageStr)
    Exit Sub
LogMessageCB_Err:
    Call TraceError(Err.Number, Err.Description, "AntiCheat.LogMessageCB", Erl)
End Sub

Public Sub HandleAntiCheatServerMessage(ByRef Data() As Byte)
    On Error Goto HandleAntiCheatServerMessage_Err
#If ENABLE_ANTICHEAT = 1 Then
    Call HandleRemoteMessage(Data(0), UBound(Data))
#End If
    Exit Sub
HandleAntiCheatServerMessage_Err:
    Call TraceError(Err.Number, Err.Description, "AntiCheat.HandleAntiCheatServerMessage", Erl)
End Sub

Public Sub UpdateAntiCheat()
    On Error Goto UpdateAntiCheat_Err
#If ENABLE_ANTICHEAT = 1 Then
   Call Update
#End If
    Exit Sub
UpdateAntiCheat_Err:
    Call TraceError(Err.Number, Err.Description, "AntiCheat.UpdateAntiCheat", Erl)
End Sub

Public Sub BeginAntiCheatSession()
    On Error Goto BeginAntiCheatSession_Err
#If ENABLE_ANTICHEAT = 1 Then
    Call BeginSession(userName)
#End If
    Exit Sub
BeginAntiCheatSession_Err:
    Call TraceError(Err.Number, Err.Description, "AntiCheat.BeginAntiCheatSession", Erl)
End Sub

Public Sub EndAntiCheatSession()
    On Error Goto EndAntiCheatSession_Err
#If ENABLE_ANTICHEAT = 1 Then
    Call EndSession
#End If
    Exit Sub
EndAntiCheatSession_Err:
    Call TraceError(Err.Number, Err.Description, "AntiCheat.EndAntiCheatSession", Erl)
End Sub
