Attribute VB_Name = "AntiCheat"
Option Explicit

Public Enum e_ACInitResult
    eOk = 0
    eFailedPlatform
    eFAiledConnectAC
End Enum

Private Declare Function InitializeAC Lib "AOAC.dll" () As Long
Public Declare Sub UnloadAC Lib "AOAC.dll" ()


Public Sub InitializeAntiCheat()
On Error GoTo InitializeAC_Err
    Dim InitResult As e_ACInitResult
    
    InitResult = InitializeAC()
    If InitResult <> eOk Then
        Call DisplayError("El juego se inicio sin activar el anti cheat, debe activarlo para poder conectarse a los servidores.", "")
    End If
    Exit Sub
InitializeAC_Err:
    Call RegistrarError(Err.Number, Err.Description, "AOAC.InitializeAC", Erl)
End Sub
