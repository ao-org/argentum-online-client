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

Private Type AudioTrack
    Name As String
    label As String
    Buffer As DirectSoundSecondaryBuffer8
End Type

Private mAudioTracks(1 To 30) As AudioTrack

Private Type MidiTrack
    Name As String
    directMusicLoader As DirectMusicLoader8
    directMusicPerformance As DirectMusicPerformance8
    directMusicSegment As DirectMusicSegment8
    directMusicSegmentState As DirectMusicSegmentState8
    directMusicPath As DirectMusicAudioPath8
End Type

Private mMidiTrack(1 To 3) As MidiTrack
' Declare the mciSendString API function from winmm.dll.
Private Declare Function mciSendString _
                Lib "winmm.dll" _
                Alias "mciSendStringA" (ByVal lpstrCommand As String, _
                                        ByVal lpstrReturnString As String, _
                                        ByVal uReturnLength As Long, _
                                        ByVal hwndCallback As Long) As Long

Public Function StopMP3() As Long
    StopMP3 = -1
    ' Stop playback of the MP3 file.
    mciSendString "stop mp3", vbNullString, 0, 0
    ' Close the MP3 file to release resources.
    mciSendString "close mp3", vbNullString, 0, 0
    StopMP3 = 1
End Function

Public Function PlayMP3(ByVal filename As String, Optional ByVal looping As Boolean = False, Optional ByVal volume As Long = 0) As Long
    Dim command As String
    Dim ret     As Long
    If filename = "" Then
        PlayMP3 = -1   ' Error: file not found.
        Exit Function
    End If
    ' Stop and close any currently playing MP3.
    mciSendString "stop mp3", vbNullString, 0, 0
    mciSendString "close mp3", vbNullString, 0, 0
    Dim PathToFile As String
    #If Compresion = 1 Then
        If Extract_File(mp3, GetCompressedResourcesPath(), filename, Windows_Temp_Dir, ResourcesPassword, False) Then
            ' Open the MP3 file with an alias "mp3".
            PathToFile = Windows_Temp_Dir & filename
        End If
    #Else
        PathToFile = ao20audio.GetMp3FilesPath() & filename
    #End If
    Debug.Assert FileExist(PathToFile, vbArchive)
    command = "open """ & PathToFile & """ type mpegvideo alias mp3"
    ret = mciSendString(command, vbNullString, 0, 0)
    If ret <> 0 Then
        PlayMP3 = ret
        Exit Function
    End If
    ' If a volume value is provided, set the volume.
    ' Note: The acceptable volume range and behavior may depend on the MCI implementation.
    If volume <> 0 Then
        command = "setaudio mp3 volume to " & volume
        ret = mciSendString(command, vbNullString, 0, 0)
    End If
    ' Play the MP3 file; append the "repeat" flag if looping is desired.
    command = "play mp3"
    If looping Then
        command = command & " repeat"
    End If
    ret = mciSendString(command, vbNullString, 0, 0)
    PlayMP3 = ret
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
    frmDebug.add_text_tracebox "PlayAmbient, track " & tid & " id: " & id
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
    Dim sound_id As Integer
    sound_id = RemoveFirstThreeIfUnderscore(prefix_sound_id)
    Debug.Assert sound_id > 0
    PlayWav = -1
    If (sound_id <= 0) Then Exit Function
    If mBuffer(sound_id) Is Nothing Then
        Dim i As Integer
        Call CreateWavBufferFromFile(prefix_sound_id, mBuffer(sound_id))
        'Debug.Assert Not mBuffer(sound_id) Is Nothing
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
    Dim i As Integer
    For i = 1 To UBound(mMidiTrack)
        With mMidiTrack(i)
            If Not .directMusicSegmentState Is Nothing Then
                .directMusicPerformance.SetMasterVolume volume
            End If
        End With
    Next i
End Sub

Public Function PlayMidi(ByVal id As Integer, Optional ByVal looping As Boolean = False, Optional ByVal volume As Long = 0) As Long
    On Error GoTo PlayMidi_Error_Handl
    Debug.Assert id > 0
    PlayMidi = -1
    If (id <= 0) Then Exit Function
    Dim id_as_str As String: id_as_str = str(id)
    With mMidiTrack(1)
        If .Name <> id_as_str Then
            .Name = id_as_str
            Call LoadMidi(mMidiTrack(1), id)
            .directMusicPerformance.StopEx .directMusicSegment, 0, 0
            Set .directMusicSegmentState = .directMusicPerformance.PlaySegmentEx(.directMusicSegment, 0, 0, Nothing, .directMusicPath)
            .directMusicPerformance.SetMasterVolume volume
        End If
    End With
    PlayMidi = 0
    Exit Function
PlayMidi_Error_Handl:
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
    For i = 1 To UBound(mMidiTrack)
        With mMidiTrack(i)
            .Name = -1
            If Not .directMusicSegmentState Is Nothing Then
                .directMusicPerformance.StopEx .directMusicSegment, 0, 0
            End If
        End With
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
End Sub

Private Sub Class_Initialize()
    SetMembersToNothing
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    Err.Clear
    StopMP3
    StopAllPlayback
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
