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
On Error GoTo InitializeAC_Err
#If ENABLE_ANTICHEAT = 1 Then
    Dim InitResult As e_ACInitResult
    Dim Callbacks As t_AntiCheatCallbacks
    Callbacks.SendToServer = FARPROC(AddressOf SendToServerCB)
    Callbacks.LogMessage = FARPROC(AddressOf LogMessageCB)
    InitResult = InitializeAC(Callbacks)
    If InitResult <> eOk Then
        Call DisplayError("El juego se inicio sin activar el anti cheat, debe activarlo para poder conectarse a los servidores.", "")
    End If
    Exit Sub
#End If
InitializeAC_Err:
    Call RegistrarError(Err.Number, Err.Description, "AOAC.InitializeAC", Erl)
End Sub

Public Sub UnloadAntiCheat()
#If ENABLE_ANTICHEAT = 1 Then
    Call UnloadAC
#End If
End Sub

Public Sub SendToServerCB(ByVal Data As Long, ByVal DataSize As Long)
    Call WriteAntiCheatMessage(Data, DataSize)
End Sub

Public Sub LogMessageCB(ByRef message As SINGLESTRINGPARAM, ByVal Level As Long)
    Dim MessageStr As String
    If message.Len > 0 Then
        MessageStr = GetStringFromPtr(message.Ptr, message.Len)
    End If
    Call LogError(MessageStr)
End Sub

Public Sub HandleAntiCheatServerMessage(ByRef Data() As Byte)
#If ENABLE_ANTICHEAT = 1 Then
    Call HandleRemoteMessage(Data(0), UBound(Data))
#End If
End Sub

Public Sub UpdateAntiCheat()
#If ENABLE_ANTICHEAT = 1 Then
   Call Update
#End If
End Sub

Public Sub BeginAntiCheatSession()
#If ENABLE_ANTICHEAT = 1 Then
    Call BeginSession(userName)
#End If
End Sub

Public Sub EndAntiCheatSession()
#If ENABLE_ANTICHEAT = 1 Then
    Call EndSession
#End If
End Sub
