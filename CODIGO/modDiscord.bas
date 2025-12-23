Attribute VB_Name = "modDiscord"
'*****************************************************************************
' Discord Rich Presence Module for Visual Basic 6
'
' This module provides VB6 declarations and helper functions for the
' DiscordRichPresenceVB6.dll
'
' Usage:
'   1. Copy DiscordRichPresenceVB6.dll and discord-rpc.dll to your application folder
'   2. Add this module to your VB6 project
'   3. Call Discord_Initialize with your Application ID
'   4. Call Discord_Update to update presence
'   5. Setup a Timer (100-250ms interval) to call Discord_RunCallbacks
'   6. Call Discord_Shutdown when closing your application
'
'*****************************************************************************

Option Explicit

' API Declarations for DiscordRichPresenceVB6.dll
' Make sure the DLL is in the same folder as your EXE or in System32
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Const DISCORD_API_ID As String = "1452879113827127377"

' Initialize Discord Rich Presence connection
' Returns:  1 on success, 0 on failure
Public Declare Function InitializeDiscord Lib "DiscordRichPresenceVB6.dll" _
    (ByVal appId As String) As Long

' Update Discord Rich Presence with detailed information
' All parameters are optional (pass vbNullString or empty string to omit)
' Returns: 1 on success, 0 on failure
Public Declare Function UpdatePresence Lib "DiscordRichPresenceVB6.dll" _
    (ByVal state As String, _
     ByVal details As String, _
     ByVal largeImageKey As String, _
     ByVal largeImageText As String, _
     ByVal smallImageKey As String, _
     ByVal smallImageText As String) As Long

' Set timestamp for elapsed time display
' startTime:  Unix timestamp in seconds (use 0 to clear)
' FIXED: Changed from Currency to Double for better compatibility
' Returns: 1 on success, 0 on failure
Public Declare Function SetTimestamp Lib "DiscordRichPresenceVB6.dll" _
    (ByVal startTime As Double) As Long

' Set party size information (current/max players)
' Returns: 1 on success, 0 on failure
Public Declare Function SetPartySize Lib "DiscordRichPresenceVB6.dll" _
    (ByVal currentSize As Long, _
     ByVal maxSize As Long) As Long

' Clear Discord Rich Presence (remove all presence information)
' Returns: 1 on success, 0 on failure
Public Declare Function ClearPresence Lib "DiscordRichPresenceVB6.dll" () As Long

' Shutdown Discord Rich Presence connection
' Call this before application exit
Public Declare Sub ShutdownDiscord Lib "DiscordRichPresenceVB6.dll" ()

' Check if Discord is initialized
' Returns: 1 if initialized, 0 if not
Public Declare Function IsDiscordInitialized Lib "DiscordRichPresenceVB6.dll" () As Long

' *** NEW:  Run Discord callbacks - MUST be called regularly (e.g., via Timer) ***
' Call this every 100-250ms to keep Discord connection alive and process events
' Returns: 1 on success, 0 on failure
Public Declare Function RunCallbacks Lib "DiscordRichPresenceVB6.dll" () As Long

' Get last error message
' Returns:  Pointer to error message string
Public Declare Function GetLastError Lib "DiscordRichPresenceVB6.dll" () As Long

'*****************************************************************************
' Helper Functions
'*****************************************************************************

' Initialize Discord with your Application ID
' Get your Application ID from:  https://discord.com/developers/applications
Public Function Discord_Initialize(ByVal appId As String) As Boolean
    Dim Result As Long
    Result = InitializeDiscord(appId)
    Discord_Initialize = (Result = 1)
End Function

' *** NEW: Run Discord Callbacks ***
' IMPORTANT: Call this regularly (every 100-250ms) via a Timer control
' Example: In Form_Load, set Timer1.Interval = 150
'          In Timer1_Timer event, call Discord_RunCallbacks
Public Function Discord_RunCallbacks() As Boolean
    Dim Result As Long
    Result = RunCallbacks()
    Discord_RunCallbacks = (Result = 1)
End Function

' Update Discord Rich Presence
' All parameters are optional - pass empty string to omit
Public Function Discord_Update( _
    Optional ByVal state As String = "", _
    Optional ByVal details As String = "", _
    Optional ByVal largeImage As String = "", _
    Optional ByVal largeText As String = "", _
    Optional ByVal smallImage As String = "", _
    Optional ByVal smallText As String = "") As Boolean
    
    Dim Result As Long
    Result = UpdatePresence(state, details, largeImage, largeText, smallImage, smallText)
    Discord_Update = (Result = 1)
End Function

' Set elapsed time (time since game started)
' Pass current time to start the timer
' FIXED: Now uses proper timestamp conversion
Public Function Discord_SetStartTime() As Boolean
    Dim Result As Long
    Dim startTime As Double
    
    ' Get current Unix timestamp
    startTime = GetUnixTimestamp()
    Result = SetTimestamp(startTime)
    Discord_SetStartTime = (Result = 1)
End Function

' Clear the timestamp
Public Function Discord_ClearTime() As Boolean
    Dim Result As Long
    Result = SetTimestamp(0)
    Discord_ClearTime = (Result = 1)
End Function

' Set party/player count
Public Function Discord_SetParty(ByVal current As Long, ByVal max As Long) As Boolean
    Dim Result As Long
    Result = SetPartySize(current, max)
    Discord_SetParty = (Result = 1)
End Function

' Clear Discord presence
Public Function Discord_Clear() As Boolean
    Dim Result As Long
    Result = ClearPresence()
    Discord_Clear = (Result = 1)
End Function

' Shutdown Discord connection
Public Sub Discord_Shutdown()
    Call ShutdownDiscord
End Sub

' Check if Discord is connected
Public Function Discord_IsConnected() As Boolean
    Dim Result As Long
    Result = IsDiscordInitialized()
    Discord_IsConnected = (Result = 1)
End Function

' Get current Unix timestamp (seconds since 1970-01-01)
' FIXED: Now returns Double for better precision
Private Function GetUnixTimestamp() As Double
    Dim dateOffset As Date
    Dim secondsSince1970 As Double
    
    dateOffset = DateSerial(1970, 1, 1)
    secondsSince1970 = DateDiff("s", dateOffset, Now)
    
    GetUnixTimestamp = secondsSince1970
End Function

' Get last error message (advanced usage)
Public Function Discord_GetLastError() As String
    Dim ptrError As Long
    Dim errorMsg As String
    
    ptrError = GetLastError()
    If ptrError <> 0 Then
        errorMsg = GetStringFromPointer(ptrError)
        Discord_GetLastError = errorMsg
    Else
        Discord_GetLastError = ""
    End If
End Function

' Helper function to read string from pointer
Private Function GetStringFromPointer(ByVal lpString As Long) As String
    Dim Buffer As String
    Dim Length As Long
    
    If lpString = 0 Then
        GetStringFromPointer = ""
        Exit Function
    End If
    
    ' Allocate buffer
    Buffer = String$(256, 0)
    
    ' Copy string from memory
    Call CopyMemory(ByVal Buffer, ByVal lpString, 256)
    
    ' Find null terminator
    Length = InStr(Buffer, vbNullChar)
    If Length > 0 Then
        GetStringFromPointer = Left$(Buffer, Length - 1)
    Else
        GetStringFromPointer = Buffer
    End If
End Function

Public Function CharStatusToString(ByVal status As Byte) As String
    Select Case status
        Case 0: CharStatusToString = JsonLanguage.Item("MENSAJE_ESTADO_CRIMINAL")
        Case 1: CharStatusToString = JsonLanguage.Item("MENSAJE_ESTADO_CIUDADANO")  ' Ciudadano
        Case 2: CharStatusToString = JsonLanguage.Item("MENSAJE_ESTADO_CAOS")  ' Caos
        Case 3: CharStatusToString = JsonLanguage.Item("MENSAJE_ESTADO_ARMADA")  ' Armada
        Case 4: CharStatusToString = JsonLanguage.Item("MENSAJE_ESTADO_CONSEJO_CAOS")  ' Concilio
        Case 5: CharStatusToString = JsonLanguage.Item("MENSAJE_ESTADO_CONSEJO_REAL")   ' Consejo
    End Select
End Function

