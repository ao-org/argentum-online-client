VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAudioEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'    Argentum 20 - Game Client Program
'    Copyright (C) 2023 - Noland Studios
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
Option Base 0
Option Explicit
Private mDirectX           As DirectX8
Private mDirectSound       As DirectSound8
Private mDirectSoundEnum   As DirectSoundEnum8
Private mBuffer(0 To 9999) As DirectSoundSecondaryBuffer8

Private Enum eMusicBackend
    mbNone = 0
    mbBassOgg = 1
    mbDirectMusicMidi = 2
End Enum

Private mCurrentMusicFile As String
Private mMusicLooping     As Boolean
Private mMusicVolume      As Long
Private mMusicBackend     As eMusicBackend
Private mMusicIsActive    As Boolean
Private mCurrentMidiId    As Integer

Private Type AudioTrack
    Name As String
    label As String
    Buffer As DirectSoundSecondaryBuffer8
End Type

Private mAudioTracks(1 To 30) As AudioTrack

Private Type tCompressedSfx
    Name As String
    FullPath As String
    SampleHandle As Long
End Type

Private mCompressedSfx() As tCompressedSfx
Private mCompressedSfxCount As Integer

Private Type MidiTrack
    Name As String
    directMusicLoader As DirectMusicLoader8
    directMusicPerformance As DirectMusicPerformance8
    directMusicSegment As DirectMusicSegment8
    directMusicSegmentState As DirectMusicSegmentState8
    directMusicPath As DirectMusicAudioPath8
End Type

Private mMidiTrack(1 To 3) As MidiTrack

Private Function GetFileExtension(ByVal fileName As String) As String
    Dim dotPos As Long
    dotPos = InStrRev(fileName, ".")
    If dotPos <= 0 Then Exit Function
    GetFileExtension = LCase$(Mid$(fileName, dotPos))
End Function

Private Function ResolveOggFilePath(ByVal fileName As String, ByRef resolvedPath As String) As Boolean
    resolvedPath = ""
    If LenB(fileName) = 0 Then Exit Function

    #If Compresion = 1 Then
        If Extract_File(ogg, GetCompressedResourcesPath(), fileName, Windows_Temp_Dir, ResourcesPassword, False) Then
            resolvedPath = Windows_Temp_Dir & fileName
        End If
    #Else
        resolvedPath = ao20audio.GetOggFilesPath() & fileName
    #End If

    ResolveOggFilePath = FileExist(resolvedPath, vbArchive)
End Function

Private Function PlayOgg(ByVal fileName As String, Optional ByVal looping As Boolean = False, Optional ByVal volume As Long = 0) As Long
    Dim pathToFile As String

    PlayOgg = -1
    If LenB(fileName) = 0 Then Exit Function

    If Not ResolveOggFilePath(fileName, pathToFile) Then
        #If DEBUG_AUDIO = 1 Then
            frmDebug.add_text_tracebox "PlayOgg missing file: " & filename
        #End If
        Exit Function
    End If

    PlayOgg = BassBackend_PlayOgg(pathToFile, looping, volume)
    #If DEBUG_AUDIO = 1 Then
    If PlayOgg = 0 Then
        frmDebug.add_text_tracebox "PlayOgg started: " & fileName
    Else
        frmDebug.add_text_tracebox "PlayOgg failed. Error: " & PlayOgg
    End If
    #End If
End Function

Private Sub ClearActiveMusicState()
    mCurrentMusicFile = ""
    mMusicLooping = False
    mMusicBackend = mbNone
    mMusicIsActive = False
    mCurrentMidiId = 0
    frmDebug.add_text_tracebox "Music state cleared"
End Sub

Private Sub SetActiveMusicState(ByVal fileName As String, ByVal backend As eMusicBackend, ByVal looping As Boolean, ByVal volume As Long)
    mCurrentMusicFile = fileName
    mMusicBackend = backend
    mMusicLooping = looping
    mMusicVolume = volume
    mMusicIsActive = True
End Sub

Private Function IsSameMusicRequest(ByVal fileName As String, ByVal backend As eMusicBackend, ByVal looping As Boolean) As Boolean
    IsSameMusicRequest = mMusicIsActive And (mMusicBackend = backend) And (StrComp(mCurrentMusicFile, fileName, vbTextCompare) = 0) And (mMusicLooping = looping)
End Function

Private Function DetermineMusicBackendFromFile(ByVal fileName As String, ByRef midiId As Integer) As eMusicBackend
    Dim extension As String

    midiId = 0
    extension = GetFileExtension(fileName)

    Select Case extension
        Case ".ogg"
            DetermineMusicBackendFromFile = mbBassOgg
        Case ".mid"
            midiId = Val(Left$(fileName, Len(fileName) - 4))
            If midiId > 0 Then
                DetermineMusicBackendFromFile = mbDirectMusicMidi
            End If
        Case Else
            DetermineMusicBackendFromFile = mbNone
    End Select
End Function

Private Sub StopAllMusicBackends()
    On Error Resume Next

    frmDebug.add_text_tracebox "StopAllMusicBackends begin. Current backend: " & CStr(mMusicBackend)

    Call BassBackend_StopMusic

    Dim i As Integer
    For i = 1 To UBound(mMidiTrack)
        With mMidiTrack(i)
            If Not .directMusicSegmentState Is Nothing Then
                .directMusicPerformance.StopEx .directMusicSegment, 0, 0
                Set .directMusicSegmentState = Nothing
            End If
        End With
    Next i

    Call ClearActiveMusicState
    frmDebug.add_text_tracebox "StopAllMusicBackends complete"
End Sub

Public Function StopMusic() As Long
    StopMusic = 0
    Call StopAllMusicBackends
End Function

Public Sub SetMusicVolume(ByVal volume As Long)
    mMusicVolume = volume

    Select Case mMusicBackend
        Case mbBassOgg
            Call BassBackend_SetMusicVolume(volume)
        Case mbDirectMusicMidi
            Dim i As Integer
            For i = 1 To UBound(mMidiTrack)
                With mMidiTrack(i)
                    If Not .directMusicSegmentState Is Nothing Then
                        .directMusicPerformance.SetMasterVolume volume
                    End If
                End With
            Next i
    End Select
End Sub

Public Function PlayMusic(ByVal fileName As String, Optional ByVal looping As Boolean = False, Optional ByVal volume As Long = 0) As Long
    Dim backend As eMusicBackend
    Dim midiId As Integer

    PlayMusic = -1
    If LenB(fileName) = 0 Then
        frmDebug.add_text_tracebox "PlayMusic invalid request: empty filename"
        Exit Function
    End If

    backend = DetermineMusicBackendFromFile(fileName, midiId)
    If backend = mbNone Then
        frmDebug.add_text_tracebox "PlayMusic unknown extension or invalid request: " & fileName
        Exit Function
    End If

    If IsSameMusicRequest(fileName, backend, looping) Then
        frmDebug.add_text_tracebox "PlayMusic same track request ignored: " & fileName
        Call SetMusicVolume(volume)
        PlayMusic = 0
        Exit Function
    End If

    If mMusicIsActive Then
        frmDebug.add_text_tracebox "PlayMusic backend switch: " & CStr(mMusicBackend) & " -> " & CStr(backend)
    End If
    Call StopAllMusicBackends

    Select Case backend
        Case mbBassOgg
            #If ENABLE_BASS = 1 Then
                frmDebug.add_text_tracebox "PlayMusic selected backend: BASS/OGG"
                PlayMusic = PlayOgg(fileName, looping, volume)
            #Else
                frmDebug.add_text_tracebox "PlayMusic .ogg ignored because ENABLE_BASS=0"
                PlayMusic = -1
            #End If
        Case mbDirectMusicMidi
            frmDebug.add_text_tracebox "PlayMusic selected backend: DirectMusic/MIDI"
            PlayMusic = PlayMidi(midiId, looping, volume)
    End Select

    If PlayMusic = 0 Then
        Call SetActiveMusicState(fileName, backend, looping, volume)
    Else
        frmDebug.add_text_tracebox "PlayMusic start failed: " & fileName
    End If
End Function

Public Function GetMidiTrackName(ByVal track_id As Integer) As String
    Debug.Assert track_id < UBound(mMidiTrack)
    If track_id < UBound(mMidiTrack) Then
        GetMidiTrackName = mMidiTrack(track_id).Name
    End If
    GetMidiTrackName = ""
End Function

Private Function InitSoundEngine(hWnd As Long) As Long
    On Error Resume Next
    Err.Clear
    Err.Number = 0
    Set mDirectSoundEnum = mDirectX.GetDSEnum
    If Err.Number <> 0 Then
        Call MsgBox(JsonLanguage.Item("MENSAJEBOX_FATAL_ERROR_DIRECTX_ENUM"), vbCritical, App.title)
        frmDebug.add_text_tracebox "Error Number Returned: " & Err.Number
        InitSoundEngine = Err.Number
        Exit Function
    End If
    Set mDirectSound = DirectX.DirectSoundCreate(mDirectSoundEnum.GetGuid(1))
    If Err.Number <> 0 Then
        Call MsgBox(JsonLanguage.Item("MENSAJEBOX_FATAL_ERROR_DIRECTX_CREAR"), vbCritical, App.title)
        frmDebug.add_text_tracebox "Error Number Returned: " & Err.Number
        InitSoundEngine = Err.Number
        Exit Function
    End If
    mDirectSound.SetCooperativeLevel hWnd, DSSCL_NORMAL
    InitSoundEngine = Err.Number
End Function

Private Function InitMusicEngine(dx8 As DirectX8, hWnd As Long) As Long
    On Error Resume Next
    Err.Clear
    With mMidiTrack(1)
        Set .directMusicLoader = dx8.DirectMusicLoaderCreate
        If Err.Number <> 0 Then
            Call MsgBox(JsonLanguage.Item("MENSAJEBOX_FATAL_ERROR_DIRECTX_MUSIC"), vbCritical, App.title)
            frmDebug.add_text_tracebox "Error Number Returned: " & Err.Number
            InitMusicEngine = Err.Number
            Exit Function
        End If
        Set .directMusicPerformance = dx8.DirectMusicPerformanceCreate
        If Err.Number <> 0 Then
            Call MsgBox(JsonLanguage.Item("MENSAJEBOX_FATAL_ERROR_DIRECTX_MUSIC_PERFORMANCE"), vbCritical, App.title)
            frmDebug.add_text_tracebox "Error Number Returned: " & Err.Number
            InitMusicEngine = Err.Number
            Exit Function
        Else
            Dim audio_params As DMUS_AUDIOPARAMS
            .directMusicPerformance.InitAudio hWnd, DMUS_AUDIOF_ALL, audio_params, Nothing, DMUS_APATH_SHARED_STEREOPLUSREVERB, 128
            .directMusicPerformance.SetMasterAutoDownload True
            .directMusicPerformance.SetMasterVolume (1)
        End If
        Set .directMusicPath = .directMusicPerformance.CreateStandardAudioPath(DMUS_APATH_DYNAMIC_3D, 64, True)
    End With
    InitMusicEngine = Err.Number
End Function

Public Function Init(ByRef dx8 As DirectX8, ByVal hWnd As Long) As Long
    Debug.Assert Not dx8 Is Nothing
    Set mDirectX = dx8
    Init = InitSoundEngine(hWnd)
    If Init <> 0 Then
        frmDebug.add_text_tracebox "InitSoundEngine() failed with code " & Init
        Exit Function
    End If
    Init = InitMusicEngine(dx8, hWnd)
    If Init <> 0 Then
        frmDebug.add_text_tracebox "InitSoundEngine() failed with code " & Init
        Exit Function
    End If
    #If ENABLE_BASS = 1 Then
        If Not InitBassAudio(hWnd) Then
            frmDebug.add_text_tracebox "BASS backend unavailable, OGG disabled. Error: " & BassBackend_GetLastError
        End If
    #Else
        frmDebug.add_text_tracebox "BASS backend disabled by compilation define ENABLE_BASS=0"
    #End If
    Init = 0
End Function

Private Function LoadMidi(ByRef track As MidiTrack, ByVal file_str As String, Optional ByVal volume As Long = 0, Optional ByVal balance As Long = 0) As Boolean
    On Error GoTo Error_Handl
    LoadMidi = False
    If (LenB(file_str) <= 0) Then Exit Function
    With track
        #If Compresion = 1 Then
            If Extract_File(Midi, GetCompressedResourcesPath(), file_str & ".mid", Windows_Temp_Dir, ResourcesPassword, False) Then
                Set .directMusicSegment = .directMusicLoader.LoadSegment(Windows_Temp_Dir & file_str & ".mid")
                .directMusicSegment.SetStandardMidiFile
                .directMusicSegment.Download .directMusicPath
            End If
        #Else
            Dim PathToFile As String
            PathToFile = ao20audio.GetMidiFilesPath() & file_str & ".mid"
            If FileExist(PathToFile, vbArchive) Then
                Set .directMusicSegment = .directMusicLoader.LoadSegment(PathToFile)
                .directMusicSegment.SetStandardMidiFile
                .directMusicSegment.Download .directMusicPath
            End If
        #End If
    End With
    LoadMidi = True
    Exit Function
Error_Handl:
    frmDebug.add_text_tracebox "error al cargar musica"
End Function

Private Function CreateWavBufferFromFile(ByVal file_name As String, ByRef new_buffer As DirectSoundSecondaryBuffer8, Optional ByVal Extraido As Boolean = False) As Boolean
    On Error GoTo CreateWavBufferFromFile_Err
    CreateWavBufferFromFile = False
    Debug.Assert Not mDirectSound Is Nothing
    Dim dsbd As DSBUFFERDESC
    dsbd.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME
    dsbd.fxFormat.nFormatTag = WAVE_FORMAT_PCM
    dsbd.fxFormat.nChannels = 2
    dsbd.fxFormat.lSamplesPerSec = 22050
    dsbd.fxFormat.nBitsPerSample = 16
    dsbd.fxFormat.nBlockAlign = dsbd.fxFormat.nBitsPerSample / 8 * dsbd.fxFormat.nChannels
    dsbd.fxFormat.lAvgBytesPerSec = dsbd.fxFormat.lSamplesPerSec * dsbd.fxFormat.nBlockAlign
    Dim filename_with_extension As String: filename_with_extension = file_name & ".wav"
    #If Compresion Then
        If Extraido = False Then
            If Extract_File(wav, GetCompressedResourcesPath(), file_name & ".wav", Windows_Temp_Dir, ResourcesPassword, False) Then
                Set new_buffer = mDirectSound.CreateSoundBufferFromFile(Windows_Temp_Dir & file_name & ".wav", dsbd)
                Delete_File Windows_Temp_Dir & filename_with_extension
                CreateWavBufferFromFile = True
            End If
        Else
            Set new_buffer = mDirectSound.CreateSoundBufferFromFile(Windows_Temp_Dir & file_name & ".wav", dsbd)
            Delete_File Windows_Temp_Dir & filename_with_extension
            CreateWavBufferFromFile = True
        End If
    #Else
        Dim PathToFile As String
        PathToFile = ao20audio.GetWavFilesPath() & filename_with_extension
        If FileExist(PathToFile, vbArchive) Then
            Set new_buffer = mDirectSound.CreateSoundBufferFromFile(PathToFile, dsbd)
            CreateWavBufferFromFile = True
        End If
    #End If
    Exit Function
CreateWavBufferFromFile_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsAudioEngine.CreateWavBufferFromFile " & file_name, Erl)
    Resume Next
End Function

Private Function ResolveOggSfxPath(ByVal Name As String) As String
    Dim fileNameWithExtension As String

    ResolveOggSfxPath = ""
    If LenB(Name) = 0 Then Exit Function

    fileNameWithExtension = Name & ".ogg"

    #If Compresion = 1 Then
        If Extract_File(ogg, GetCompressedResourcesPath(), fileNameWithExtension, Windows_Temp_Dir, ResourcesPassword, False) Then
            ResolveOggSfxPath = Windows_Temp_Dir & fileNameWithExtension
        End If
    #Else
        ResolveOggSfxPath = ao20audio.GetOggFilesPath() & fileNameWithExtension
    #End If

    If Not FileExist(ResolveOggSfxPath, vbArchive) Then
        ResolveOggSfxPath = ""
    End If
End Function

Private Function ConvertFxVolumeToBass(ByVal volume As Long) As Single
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

    ConvertFxVolumeToBass = bassVolume
End Function

Private Function ConvertFxPanToBass(ByVal pan As Long) As Single
    Dim bassPan As Single

    If pan <= -10000 Then
        bassPan = -1
    ElseIf pan >= 10000 Then
        bassPan = 1
    Else
        bassPan = pan / 10000
    End If

    If bassPan < -1 Then bassPan = -1
    If bassPan > 1 Then bassPan = 1

    ConvertFxPanToBass = bassPan
End Function

Private Function FindCompressedSfxIndex(ByVal Name As String, ByVal fullPath As String) As Integer
    Dim i As Integer

    FindCompressedSfxIndex = -1
    If mCompressedSfxCount <= 0 Then Exit Function

    For i = 1 To mCompressedSfxCount
        If (StrComp(mCompressedSfx(i).Name, Name, vbTextCompare) = 0) Or _
           (StrComp(mCompressedSfx(i).FullPath, fullPath, vbTextCompare) = 0) Then
            FindCompressedSfxIndex = i
            Exit Function
        End If
    Next i
End Function

Private Function GetCompressedSfxSample(ByVal Name As String, ByVal fullPath As String) As Long
    On Error GoTo GetCompressedSfxSample_Error
    Dim sampleHandle As Long
    Dim i As Integer

    GetCompressedSfxSample = 0
    If LenB(Name) = 0 Or LenB(fullPath) = 0 Then Exit Function

    i = FindCompressedSfxIndex(Name, fullPath)
    If i > 0 Then
        If mCompressedSfx(i).SampleHandle <> 0 Then
            GetCompressedSfxSample = mCompressedSfx(i).SampleHandle
            #If DEBUG_AUDIO = 1 Then
                frmDebug.add_text_tracebox "PlayWav OGG cache hit: " & Name
            #End If
            Exit Function
        End If
    End If

    sampleHandle = BASS_SampleLoad(0, StrPtr(fullPath), 0, 0, 8, BASS_SAMPLE_OVER_VOL)
    If sampleHandle = 0 Then
        #If DEBUG_AUDIO = 1 Then
            frmDebug.add_text_tracebox "PlayWav OGG sample load failed: " & Name & " Error: " & BASS_ErrorGetCode()
        #End If
        Exit Function
    End If

    If i <= 0 Then
        mCompressedSfxCount = mCompressedSfxCount + 1
        If mCompressedSfxCount = 1 Then
            ReDim mCompressedSfx(1 To 1)
        Else
            ReDim Preserve mCompressedSfx(1 To mCompressedSfxCount)
        End If
        i = mCompressedSfxCount
    End If

    mCompressedSfx(i).Name = Name
    mCompressedSfx(i).FullPath = fullPath
    mCompressedSfx(i).SampleHandle = sampleHandle

    #If DEBUG_AUDIO = 1 Then
        frmDebug.add_text_tracebox "PlayWav OGG sample cached: " & Name
    #End If
    GetCompressedSfxSample = sampleHandle
    Exit Function

GetCompressedSfxSample_Error:
    frmDebug.add_text_tracebox "GetCompressedSfxSample exception: " & Err.Description
End Function

Private Function PlayCompressedSfx(ByVal Name As String, ByVal fullPath As String, ByVal looping As Boolean, ByVal volume As Long, Optional ByVal pan As Long = 0) As Long
    On Error GoTo PlayCompressedSfx_Error
    Dim sampleHandle As Long
    Dim channelHandle As Long
    Dim flags As Long

    PlayCompressedSfx = -1

    If Not InitBassAudio(0) Then
        PlayCompressedSfx = BassBackend_GetLastError
        If PlayCompressedSfx = 0 Then PlayCompressedSfx = -1
        Exit Function
    End If

    sampleHandle = GetCompressedSfxSample(Name, fullPath)
    If sampleHandle = 0 Then
        PlayCompressedSfx = BASS_ErrorGetCode()
        If PlayCompressedSfx = 0 Then PlayCompressedSfx = -1
        Exit Function
    End If

    channelHandle = BASS_SampleGetChannel(sampleHandle, BASS_SAMCHAN_NEW)
    If channelHandle = 0 Then
        PlayCompressedSfx = BASS_ErrorGetCode()
        #If DEBUG_AUDIO = 1 Then
            frmDebug.add_text_tracebox "PlayWav OGG sample channel failed: " & Name & " Error: " & PlayCompressedSfx
        #End If
        Exit Function
    End If

    flags = IIf(looping, BASS_SAMPLE_LOOP, 0)
    Call BASS_ChannelFlags(channelHandle, flags, BASS_SAMPLE_LOOP)
    Call BASS_ChannelSetAttribute(channelHandle, BASS_ATTRIB_VOL, ConvertFxVolumeToBass(volume))
    Call BASS_ChannelSetAttribute(channelHandle, BASS_ATTRIB_PAN, ConvertFxPanToBass(pan))

    If BASS_ChannelPlay(channelHandle, 1) = 0 Then
        PlayCompressedSfx = BASS_ErrorGetCode()
        #If DEBUG_AUDIO = 1 Then
            frmDebug.add_text_tracebox "PlayWav OGG play failed: " & Name & " Error: " & PlayCompressedSfx
        #End If
        Exit Function
    End If

    PlayCompressedSfx = 0
    Exit Function

PlayCompressedSfx_Error:
    frmDebug.add_text_tracebox "PlayCompressedSfx exception: " & Err.Description
End Function

Private Sub FreeCompressedSfxCache()
    On Error GoTo FreeCompressedSfxCache_Error
    Dim i As Integer

    If mCompressedSfxCount <= 0 Then Exit Sub

    For i = 1 To mCompressedSfxCount
        If mCompressedSfx(i).SampleHandle <> 0 Then
            Call BASS_SampleFree(mCompressedSfx(i).SampleHandle)
            mCompressedSfx(i).SampleHandle = 0
        End If
        mCompressedSfx(i).Name = ""
        mCompressedSfx(i).FullPath = ""
    Next i

    Erase mCompressedSfx
    mCompressedSfxCount = 0
    #If DEBUG_AUDIO = 1 Then
        frmDebug.add_text_tracebox "Compressed OGG SFX cache cleared"
    #End If
    Exit Sub

FreeCompressedSfxCache_Error:
    frmDebug.add_text_tracebox "FreeCompressedSfxCache exception: " & Err.Description
End Sub

Private Function StartPlaying(ByRef Buffer As DirectSoundSecondaryBuffer8, ByVal flags As CONST_DSBPLAYFLAGS, ByVal volume As Long, ByVal pan As Long) As Long
    On Error GoTo ErrorHandler
    StartPlaying = False
    Debug.Assert Not Buffer Is Nothing
    Debug.Assert (Buffer.GetStatus And DSBSTATUS_PLAYING) = 0
    Err.Clear
    If Buffer Is Nothing Then
        StartPlaying = -1
        Exit Function
    End If
    If Buffer.GetStatus = DSBSTATUS_PLAYING Then
        StartPlaying = -1
        Exit Function
    End If
    Buffer.SetPan pan
    Buffer.SetVolume volume
    Buffer.Play flags
    StartPlaying = 0
    Exit Function
ErrorHandler:
    StartPlaying = False
End Function

Private Function FindTrackByName(ByVal Name As String, Optional ByVal label As String = "") As Integer
    Dim i As Integer
    For i = 1 To UBound(mAudioTracks)
        With mAudioTracks(i)
            If Not .Buffer Is Nothing Then
                Dim sf As CONST_DSBSTATUSFLAGS: sf = .Buffer.GetStatus
                If (Name = .Name) And (label = .label) And (DSBSTATUS_PLAYING And sf) = DSBSTATUS_PLAYING Then
                    FindTrackByName = i
                    Exit Function
                End If
            End If
        End With
    Next i
    FindTrackByName = -1
End Function

Private Function GetFreeWavTrack() As Integer
    Dim i As Integer
    For i = 1 To UBound(mAudioTracks) - 1
        With mAudioTracks(i)
            If .Buffer Is Nothing Then
                GetFreeWavTrack = i
                Exit Function
            Else
                Dim sf As CONST_DSBSTATUSFLAGS: sf = .Buffer.GetStatus
                'frmdebug.add_text_tracebox "buffer status " & sf
                If (DSBSTATUS_PLAYING And sf) = 0 And (DSBSTATUS_LOOPING And sf) = 0 Then
                    GetFreeWavTrack = i
                    Exit Function
                End If
            End If
        End With
    Next i
    GetFreeWavTrack = -1
End Function

Public Function StopAmbient() As Long
    On Error GoTo StopAmbient_Error_Handl
    StopAmbient = -1
    With mAudioTracks(UBound(mAudioTracks))
        If Not .Buffer Is Nothing Then
            .Buffer.Stop
            .Buffer.SetCurrentPosition 0
        End If
    End With
    StopAmbient = 0
    Exit Function
StopAmbient_Error_Handl:
End Function

Public Function PlayAmbient(ByVal id As Integer, Optional ByVal looping As Boolean = False, Optional ByVal volume As Long = 0, Optional ByVal pan As Long = 0) As Long
    On Error GoTo PlayAmbient_Error_Handl
    Debug.Assert id > 0
    PlayAmbient = -1
    If (id <= 0) Then Exit Function
    If mBuffer(id) Is Nothing Then
        Dim i As Integer
        Call CreateWavBufferFromFile(id, mBuffer(id))
        Debug.Assert Not mBuffer(id) Is Nothing
        If mBuffer(id) Is Nothing Then
            PlayAmbient = -1
            Exit Function
        End If
    End If
    Dim tid As Integer
    tid = UBound(mAudioTracks)
    #If DEBUG_AUDIO = 1 Then
        frmDebug.add_text_tracebox "PlayAmbient, track " & tid & " id: " & id
    #End If
    With mAudioTracks(tid)
        .Name = id
        Set .Buffer = mDirectSound.DuplicateSoundBuffer(ByVal mBuffer(id))
        Call StartPlaying(.Buffer, IIf(looping, DSBPLAY_LOOPING, DSBPLAY_DEFAULT), volume, pan)
    End With
    PlayAmbient = 0
    Exit Function
PlayAmbient_Error_Handl:
End Function

Public Function PlayWav(ByVal prefix_sound_id As String, _
                        Optional ByVal looping As Boolean = False, _
                        Optional ByVal volume As Long = 0, _
                        Optional ByVal pan As Long = 0, _
                        Optional ByVal label As String = "") As Long
    On Error GoTo PlayWav_Error_Handl
    Dim oggPath As String
    Dim sound_id As Integer
    #If ENABLE_BASS = 1 Then
        oggPath = ResolveOggSfxPath(prefix_sound_id)
        If LenB(oggPath) > 0 Then
            #If DEBUG_AUDIO = 1 Then
                frmDebug.add_text_tracebox "PlayWav using OGG sample: " & prefix_sound_id
            #End If
            PlayWav = PlayCompressedSfx(prefix_sound_id, oggPath, looping, volume, pan)
            If PlayWav = 0 Then Exit Function
            #If DEBUG_AUDIO = 1 Then
                frmDebug.add_text_tracebox "PlayWav OGG failed, fallback to WAV: " & prefix_sound_id
            #End If
        Else
            #If DEBUG_AUDIO = 1 Then
                frmDebug.add_text_tracebox "PlayWav fallback to WAV: " & prefix_sound_id
            #End If
            
        End If
    #End If

    sound_id = RemoveFirstThreeIfUnderscore(prefix_sound_id)
    Debug.Assert sound_id > 0
    PlayWav = -1
    If (sound_id <= 0) Then Exit Function
    If mBuffer(sound_id) Is Nothing Then
        Dim i As Integer
        Call CreateWavBufferFromFile(prefix_sound_id, mBuffer(sound_id))
        Debug.Assert Not mBuffer(sound_id) Is Nothing
        If mBuffer(sound_id) Is Nothing Then
            PlayWav = -1
            Exit Function
        End If
    End If
    Dim tid As Integer
    tid = GetFreeWavTrack()
    If (tid > -1) Then
        With mAudioTracks(tid)
            .Name = sound_id
            .label = label
            Set .Buffer = mDirectSound.DuplicateSoundBuffer(ByVal mBuffer(sound_id))
            Call StartPlaying(.Buffer, IIf(looping, DSBPLAY_LOOPING, DSBPLAY_DEFAULT), volume, pan)
        End With
        PlayWav = 0
    End If
    Exit Function
PlayWav_Error_Handl:
End Function

Public Function RemoveFirstThreeIfUnderscore(ByVal s As String) As String
    ' Check for underscore anywhere in the string
    If InStr(s, "_") > 0 Then
        ' If the string is longer than 3 characters, strip off the first 3
        If Len(s) > 3 Then
            RemoveFirstThreeIfUnderscore = mid$(s, 4)
        Else
            ' If it's 3 characters or less, return an empty string
            RemoveFirstThreeIfUnderscore = ""
        End If
    Else
        ' No underscore: return the original string
        RemoveFirstThreeIfUnderscore = s
    End If
End Function

Public Sub ApplyMusicVolume(ByVal volume As Long)
    Call SetMusicVolume(volume)
End Sub

Public Function PlayMidi(ByVal id As Integer, Optional ByVal looping As Boolean = False, Optional ByVal volume As Long = 0) As Long
    On Error GoTo PlayMidi_Error_Handl
    Debug.Assert id > 0
    PlayMidi = -1
    If (id <= 0) Then Exit Function

    If mMusicIsActive And mMusicBackend = mbDirectMusicMidi And mCurrentMidiId = id And mMusicLooping = looping Then
        frmDebug.add_text_tracebox "PlayMidi same track request ignored: " & CStr(id) & ".mid"
        Call SetMusicVolume(volume)
        PlayMidi = 0
        Exit Function
    End If

    If mMusicIsActive Then
        frmDebug.add_text_tracebox "PlayMidi backend switch: " & CStr(mMusicBackend) & " -> " & CStr(mbDirectMusicMidi)
        Call StopAllMusicBackends
    End If

    Dim id_as_str As String: id_as_str = str(id)
    With mMidiTrack(1)
        If .Name <> id_as_str Or .directMusicSegment Is Nothing Then
            .Name = id_as_str
            Call LoadMidi(mMidiTrack(1), id)
        End If

        If .directMusicSegment Is Nothing Then Exit Function

        .directMusicPerformance.StopEx .directMusicSegment, 0, 0
        Set .directMusicSegmentState = .directMusicPerformance.PlaySegmentEx(.directMusicSegment, 0, 0, Nothing, .directMusicPath)
        .directMusicPerformance.SetMasterVolume volume
    End With
    PlayMidi = 0
    mCurrentMidiId = id
    Call SetActiveMusicState(Trim$(id_as_str) & ".mid", mbDirectMusicMidi, looping, volume)
    frmDebug.add_text_tracebox "PlayMidi started: " & Trim$(id_as_str) & ".mid"
    Exit Function
PlayMidi_Error_Handl:
    frmDebug.add_text_tracebox "PlayMidi failed for id: " & CStr(id)
End Function

Public Function StopAllWavsMatchingLabel(ByVal label As String) As Long
    'Stops any wav playing on any track if the label matches the given as a parameter
    On Error GoTo Error_Handl
    StopAllWavsMatchingLabel = -1
    If (label = "") Then Exit Function
    Dim i As Integer
    For i = 1 To UBound(mAudioTracks)
        With mAudioTracks(i)
            If Not .Buffer Is Nothing Then
                Dim sf As CONST_DSBSTATUSFLAGS: sf = .Buffer.GetStatus
                If (label = .label) And (DSBSTATUS_PLAYING And sf) = DSBSTATUS_PLAYING And (DSBSTATUS_LOOPING And sf) = DSBSTATUS_LOOPING Then
                    .Buffer.Stop
                    .Buffer.SetCurrentPosition 0
                    Exit Function
                End If
            End If
        End With
    Next i
    StopAllWavsMatchingLabel = 0
    Exit Function
Error_Handl:
End Function

Public Function StopWav(ByVal sound_id As Integer, Optional ByVal label As String = "") As Long
    On Error GoTo Error_Handl
    Debug.Assert sound_id > 0
    StopWav = -1
    If (sound_id <= 0) Then Exit Function
    Dim tid As Integer: tid = FindTrackByName(sound_id, label)
    If tid > -1 Then
        With mAudioTracks(tid)
            If Not .Buffer Is Nothing Then
                .Buffer.Stop
                .Buffer.SetCurrentPosition 0
            End If
        End With
    End If
    StopWav = 0
    Exit Function
Error_Handl:
End Function

Public Sub StopAllPlayback()
    Dim i As Integer
    For i = 1 To UBound(mAudioTracks)
        With mAudioTracks(i)
            If Not .Buffer Is Nothing Then
                .Buffer.Stop
                .Buffer.SetCurrentPosition 0
            End If
        End With
    Next i
    Call StopAllMusicBackends
    For i = 1 To UBound(mMidiTrack)
        mMidiTrack(i).Name = -1
    Next i
End Sub

Private Sub SetMembersToNothing()
    On Error Resume Next
    Dim i As Integer
    For i = 1 To UBound(mAudioTracks)
        With mAudioTracks(i)
            If Not .Buffer Is Nothing Then
                Set .Buffer = Nothing
            End If
        End With
    Next i
    For i = 1 To UBound(mMidiTrack)
        With mMidiTrack(i)
            Set .directMusicLoader = Nothing
            Set .directMusicPerformance = Nothing
            Set .directMusicSegment = Nothing
            Set .directMusicSegmentState = Nothing
            Set .directMusicPath = Nothing
        End With
    Next i
    For i = LBound(mBuffer) To UBound(mBuffer)
        Set mBuffer(i) = Nothing
    Next i
    Set mDirectX = Nothing
    Set mDirectSound = Nothing
    Set mDirectSoundEnum = Nothing
    Erase mCompressedSfx
    mCompressedSfxCount = 0
End Sub

Public Sub ApplyAmbientVolume(ByVal volume As Long)
    Dim i As Integer
    For i = 1 To UBound(mAudioTracks)
        With mAudioTracks(i)
            If Not .Buffer Is Nothing Then
                Call .Buffer.SetVolume(volume)
            End If
        End With
    Next i
End Sub

Private Sub Class_Initialize()
    SetMembersToNothing
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Err.Clear
    StopMusic
    StopAllPlayback
    FreeCompressedSfxCache
    ShutdownBassAudio
    Dim i As Integer
    For i = 1 To UBound(mMidiTrack)
        With mMidiTrack(i)
            If Not .directMusicPerformance Is Nothing Then
                Call .directMusicPerformance.CloseDown
            End If
        End With
    Next i
    SetMembersToNothing
End Sub
