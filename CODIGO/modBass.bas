Attribute VB_Name = "modBass"
Option Explicit

Private Declare Function BASS_Init Lib "bass.dll" (ByVal device As Long, ByVal freq As Long, ByVal flags As Long, ByVal win As Long, ByVal clsid As Long) As Long
Private Declare Function BASS_Free Lib "bass.dll" () As Long
Private Declare Function BASS_StreamCreateFile Lib "bass.dll" Alias "BASS_StreamCreateFileA" (ByVal mem As Long, ByVal strFile As String, ByVal offset As Long, ByVal length As Long, ByVal flags As Long) As Long
Private Declare Function BASS_ChannelPlay Lib "bass.dll" (ByVal handle As Long, ByVal restart As Long) As Long
Private Declare Function BASS_ChannelStop Lib "bass.dll" (ByVal handle As Long) As Long
Private Declare Function BASS_StreamFree Lib "bass.dll" (ByVal handle As Long) As Long
Private Declare Function BASS_ChannelSetAttribute Lib "bass.dll" (ByVal handle As Long, ByVal attrib As Long, ByVal value As Single) As Long
Private Declare Function BASS_ChannelIsActive Lib "bass.dll" (ByVal handle As Long) As Long
Private Declare Function BASS_ErrorGetCode Lib "bass.dll" () As Long

Public Const BASS_ATTRIB_VOL       As Long = 2
Public Const BASS_SAMPLE_LOOP      As Long = &H4
Private Const BASS_INIT_DEFAULT    As Long = 0
Private Const BASS_DEFAULT_DEVICE  As Long = -1
Private Const BASS_DEFAULT_FREQ    As Long = 44100
Private Const BASS_ACTIVE_PLAYING  As Long = 1

Private m_BassInitialized    As Boolean
Private m_CurrentMusicStream As Long
Private m_LastBassError      As Long

Private Function ConvertMusicVolumeToBass(ByVal volume As Long) As Single
    Dim bassVolume As Single

    If volume <= 0 Then
        bassVolume = (10000 + volume) / 10000
    ElseIf volume <= 1000 Then
        bassVolume = volume / 1000
    ElseIf volume <= 10000 Then
        bassVolume = volume / 10000
    Else
        bassVolume = 1
    End If

    If bassVolume < 0 Then bassVolume = 0
    If bassVolume > 1 Then bassVolume = 1

    ConvertMusicVolumeToBass = bassVolume
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

Public Function BassBackend_PlayOgg(ByVal filePath As String, ByVal looping As Boolean, ByVal volume As Long) As Long
    On Error GoTo BassBackend_PlayOgg_Err
    Dim flags As Long

    BassBackend_PlayOgg = 1

    If LenB(filePath) = 0 Then
        frmDebug.add_text_tracebox "BassBackend_PlayOgg missing path"
        Exit Function
    End If

    If Not FileExist(filePath, vbArchive) Then
        frmDebug.add_text_tracebox "BassBackend_PlayOgg file not found: " & filePath
        BassBackend_PlayOgg = 2
        Exit Function
    End If

    If Not InitBassAudio(0) Then
        If m_LastBassError <> 0 Then
            BassBackend_PlayOgg = m_LastBassError
        End If
        Exit Function
    End If

    Call BassBackend_StopMusic

    flags = 0
    If looping Then flags = flags Or BASS_SAMPLE_LOOP

    m_CurrentMusicStream = BASS_StreamCreateFile(0, filePath, 0, 0, flags)
    If m_CurrentMusicStream = 0 Then
        m_LastBassError = BASS_ErrorGetCode()
        frmDebug.add_text_tracebox "BASS_StreamCreateFile failed. Error: " & m_LastBassError & " Path: " & filePath
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

Public Function BassBackend_IsMusicPlaying() As Boolean
    If m_CurrentMusicStream = 0 Then Exit Function
    BassBackend_IsMusicPlaying = (BASS_ChannelIsActive(m_CurrentMusicStream) = BASS_ACTIVE_PLAYING)
End Function

Public Function BassBackend_GetLastError() As Long
    BassBackend_GetLastError = m_LastBassError
End Function
