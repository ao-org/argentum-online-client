Attribute VB_Name = "ao20audio_bass_backend"
Option Explicit

#If ENABLE_BASS = 1 Then

' Wrapper around raw BASS declares used by this project for MUSIC streams.
' Scope intentionally limited to music lifecycle:
' - initialize/free BASS device
' - manage one active music stream handle
' - apply runtime music volume
'
' Compressed SFX sample logic is handled in clsAudioEngine (PlayWav path)
' because it needs cache + per-request channel spawning behavior.

Private Const BASS_INIT_DEFAULT    As Long = 0
Private Const BASS_DEFAULT_DEVICE  As Long = -1
Private Const BASS_DEFAULT_FREQ    As Long = 44100
Private Const BASS_ACTIVE_PLAYING  As Long = 1

Private m_BassInitialized    As Boolean
Private m_CurrentMusicStream As Long
Private m_LastBassError      As Long

Private Function ConvertMusicVolumeToBass(ByVal volume As Long) As Single
    Dim gain As Double

    If volume >= 0 Then
        ConvertMusicVolumeToBass = 1!
        Exit Function
    End If

    If volume <= -10000 Then
        ConvertMusicVolumeToBass = 0!
        Exit Function
    End If

    gain = 10 ^ (volume / 2000#)

    If gain < 0# Then gain = 0#
    If gain > 1# Then gain = 1#

    ConvertMusicVolumeToBass = CSng(gain)
End Function

Public Function InitBassAudio(ByVal hWndOwner As Long) As Boolean
    On Error GoTo InitBassAudio_Err

    If m_BassInitialized Then
        InitBassAudio = True
        Exit Function
    End If

    m_LastBassError = 0

    If BASS_Init(BASS_DEFAULT_DEVICE, BASS_DEFAULT_FREQ, BASS_INIT_DEFAULT, hWndOwner, 0) = 0 Then
        m_LastBassError = BASS_ErrorGetCode()
        frmDebug.add_text_tracebox "BASS_Init failed. Error: " & m_LastBassError
        InitBassAudio = False
        Exit Function
    End If

    m_BassInitialized = True
    frmDebug.add_text_tracebox "BASS backend initialized"
    InitBassAudio = True
    Exit Function

InitBassAudio_Err:
    m_LastBassError = Err.Number
    frmDebug.add_text_tracebox "InitBassAudio exception: " & Err.Description
    InitBassAudio = False
End Function

Public Sub BassBackend_StopMusic()
    On Error GoTo BassBackend_StopMusic_Err

    If m_CurrentMusicStream <> 0 Then
        frmDebug.add_text_tracebox "BASS stopping current music stream " & m_CurrentMusicStream
        Call BASS_ChannelStop(m_CurrentMusicStream)
        Call BASS_StreamFree(m_CurrentMusicStream)
        m_CurrentMusicStream = 0
    End If

    Exit Sub
BassBackend_StopMusic_Err:
    frmDebug.add_text_tracebox "BassBackend_StopMusic exception: " & Err.Description
End Sub

Public Sub ShutdownBassAudio()
    On Error GoTo ShutdownBassAudio_Err

    Call BassBackend_StopMusic

    If m_BassInitialized Then
        Call BASS_Free
        m_BassInitialized = False
        frmDebug.add_text_tracebox "BASS backend freed"
    End If

    m_LastBassError = 0
    Exit Sub

ShutdownBassAudio_Err:
    frmDebug.add_text_tracebox "ShutdownBassAudio exception: " & Err.Description
End Sub

Public Function BassBackend_PlayOgg(ByVal FilePath As String, ByVal looping As Boolean, ByVal volume As Long) As Long
    On Error GoTo BassBackend_PlayOgg_Err
    Dim flags As Long

    BassBackend_PlayOgg = 1

    If LenB(FilePath) = 0 Then
        frmDebug.add_text_tracebox "BassBackend_PlayOgg missing path"
        Exit Function
    End If

    If Not FileExist(FilePath, vbArchive) Then
        frmDebug.add_text_tracebox "BassBackend_PlayOgg file not found: " & FilePath
        BassBackend_PlayOgg = 2
        Exit Function
    End If

    If Not InitBassAudio(0) Then
        If m_LastBassError <> 0 Then
            BassBackend_PlayOgg = m_LastBassError
        End If
        Exit Function
    End If

    ' Music path keeps a single active stream by design.
    ' Starting a new track always releases the prior stream first.
    Call BassBackend_StopMusic

    flags = 0
    If looping Then flags = flags Or BASS_SAMPLE_LOOP

    
    m_CurrentMusicStream = BASS_StreamCreateFile(0, StrPtr(FilePath), 0, 0, flags)
    
    If m_CurrentMusicStream = 0 Then
        m_LastBassError = BASS_ErrorGetCode()
        frmDebug.add_text_tracebox "BASS_StreamCreateFile failed. Error: " & m_LastBassError & " Path: " & FilePath
        BassBackend_PlayOgg = m_LastBassError
        Exit Function
    End If

    Call BASS_ChannelSetAttribute(m_CurrentMusicStream, BASS_ATTRIB_VOL, ConvertMusicVolumeToBass(volume))

    If BASS_ChannelPlay(m_CurrentMusicStream, 0) = 0 Then
        m_LastBassError = BASS_ErrorGetCode()
        frmDebug.add_text_tracebox "BASS_ChannelPlay failed. Error: " & m_LastBassError
        Call BASS_StreamFree(m_CurrentMusicStream)
        m_CurrentMusicStream = 0
        BassBackend_PlayOgg = m_LastBassError
        Exit Function
    End If

    BassBackend_PlayOgg = 0
    Exit Function

BassBackend_PlayOgg_Err:
    frmDebug.add_text_tracebox "BassBackend_PlayOgg exception: " & Err.Description
    BassBackend_PlayOgg = Err.Number
End Function

Public Sub BassBackend_SetMusicVolume(ByVal volume As Long)
    On Error GoTo BassBackend_SetMusicVolume_Err

    If m_CurrentMusicStream = 0 Then Exit Sub

    Call BASS_ChannelSetAttribute(m_CurrentMusicStream, BASS_ATTRIB_VOL, ConvertMusicVolumeToBass(volume))
    Exit Sub

BassBackend_SetMusicVolume_Err:
    frmDebug.add_text_tracebox "BassBackend_SetMusicVolume exception: " & Err.Description
End Sub


Public Function BassBackend_IsInitialized() As Boolean
    BassBackend_IsInitialized = m_BassInitialized
End Function

Public Function BassBackend_IsMusicPlaying() As Boolean
    If m_CurrentMusicStream = 0 Then Exit Function
    BassBackend_IsMusicPlaying = (BASS_ChannelIsActive(m_CurrentMusicStream) = BASS_ACTIVE_PLAYING)
End Function

Public Function BassBackend_GetLastError() As Long
    BassBackend_GetLastError = m_LastBassError
End Function

#Else

Public Function InitBassAudio(ByVal hWndOwner As Long) As Boolean
    InitBassAudio = False
End Function

Public Sub ShutdownBassAudio()
End Sub

Public Function BassBackend_PlayOgg(ByVal FilePath As String, ByVal looping As Boolean, ByVal volume As Long) As Long
    BassBackend_PlayOgg = -1
End Function

Public Sub BassBackend_StopMusic()
End Sub

Public Sub BassBackend_SetMusicVolume(ByVal volume As Long)
End Sub


Public Function BassBackend_IsInitialized() As Boolean
    BassBackend_IsInitialized = False
End Function

Public Function BassBackend_IsMusicPlaying() As Boolean
    BassBackend_IsMusicPlaying = False
End Function

Public Function BassBackend_GetLastError() As Long
    BassBackend_GetLastError = -1
End Function

#End If
