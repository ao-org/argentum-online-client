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

Private mDirectX As DirectX8
Private mDirectSound As DirectSound8
Private mDirectSoundEnum As DirectSoundEnum8

Private mBuffer(0 To 9999) As DirectSoundSecondaryBuffer8

Private Type AudioTrack
    name As String
    buffer As DirectSoundSecondaryBuffer8
End Type

Private mAudioTracks(1 To 30) As AudioTrack

Private Type MidiTrack
    name As String
    directMusicLoader As DirectMusicLoader8
    directMusicPerformance As DirectMusicPerformance8
    directMusicSegment As DirectMusicSegment8
    directMusicSegmentState As DirectMusicSegmentState8
    directMusicPath As DirectMusicAudioPath8
End Type

Private mMidiTrack(1 To 3) As MidiTrack

Public Function get_midi_track_name(ByVal track_id As Integer) As String
    Debug.Assert track_id < UBound(mMidiTrack)
    If track_id < UBound(mMidiTrack) Then
        get_midi_track_name = mMidiTrack(track_id).Name
    End If
    get_midi_track_name = ""
End Function

Private Function init_sound_engine(hwnd As Long) As Long
On Error Resume Next
    Err.Clear
    Err.Number = 0
    Set mDirectSoundEnum = mDirectX.GetDSEnum
    If Err.Number <> 0 Then
                Call MsgBox("Fatal error DirectX.GetDSEnum", vbCritical, App.title)
                Debug.Print "Error Number Returned: " & Err.Number
                init_sound_engine = Err.Number
                Exit Function
    End If
    
    Set mDirectSound = DirectX.DirectSoundCreate(mDirectSoundEnum.GetGuid(1))
    If Err.Number <> 0 Then
                Call MsgBox("Fatal error DirectX.DirectSoundCreate", vbCritical, App.title)
                Debug.Print "Error Number Returned: " & Err.Number
                init_sound_engine = Err.Number
                Exit Function
    End If
    
    mDirectSound.SetCooperativeLevel hwnd, DSSCL_NORMAL
    init_sound_engine = Err.Number
End Function

Private Function init_music_engine(dx8 As DirectX8, hwnd As Long) As Long
On Error Resume Next
    Err.Clear
    With mMidiTrack(1)
            Set .directMusicLoader = dx8.DirectMusicLoaderCreate
            If Err.Number <> 0 Then
                Call MsgBox("Fatal error DirectX.DirectMusicLoaderCreate", vbCritical, App.title)
                Debug.Print "Error Number Returned: " & Err.Number
                init_music_engine = Err.Number
                Exit Function
            End If
   
            Set .directMusicPerformance = dx8.DirectMusicPerformanceCreate
            If Err.Number <> 0 Then
                Call MsgBox("Fatal error  dx8.DirectMusicPerformanceCreate", vbCritical, App.title)
                Debug.Print "Error Number Returned: " & Err.Number
                init_music_engine = Err.Number
                Exit Function
            Else
                Dim audio_params As DMUS_AUDIOPARAMS
                .directMusicPerformance.InitAudio hwnd, DMUS_AUDIOF_ALL, audio_params, Nothing, DMUS_APATH_SHARED_STEREOPLUSREVERB, 128
                
                 ' dmp.InitAudio Me.hwnd, DMUS_AUDIOF_ALL, dmusAudio, Nothing, DMUS_APATH_SHARED_STEREOPLUSREVERB, 128
                 .directMusicPerformance.SetMasterAutoDownload True
    
                .directMusicPerformance.SetMasterVolume (1)
            End If
            
            Set .directMusicPath = .directMusicPerformance.CreateStandardAudioPath(DMUS_APATH_DYNAMIC_3D, 64, True)
    End With
  
    init_music_engine = Err.Number
End Function


Public Function init(ByRef dx8 As DirectX8, ByVal hwnd As Long) As Long
    Debug.Assert Not dx8 Is Nothing
    Set mDirectX = dx8
    init = init_sound_engine(hwnd)
    If init <> 0 Then
        Debug.Print "init_sound_engine() failed with code " & init
        Exit Function
    End If
    init = init_music_engine(dx8, hwnd)
    If init <> 0 Then
        Debug.Print "init_sound_engine() failed with code " & init
        Exit Function
    End If
    init = 0
End Function

Private Function load_midi(ByRef track As MidiTrack, ByVal file_str As String, Optional ByVal volume As Long = 0, Optional ByVal balance As Long = 0) As Boolean
On Error GoTo Error_Handl
    
    If (LenB(file_str) <= 0) Then Exit Function
 
    With track
    
        #If Compresion = 1 Then
                If Extract_File(Midi, get_compressed_resources_path(), file_str & ".mid", Windows_Temp_Dir, ResourcesPassword, False) Then
                    Set .directMusicSegment = .directMusicLoader.LoadSegment(Windows_Temp_Dir & file_str & ".mid")
                    .directMusicSegment.SetStandardMidiFile
                    .directMusicSegment.Download .directMusicPath
                End If
        #Else
           Dim PathToFile As String
           PathToFile = ao20audio.get_midi_files_path() & file_str & ".mid"
            If FileExist(PathToFile, vbArchive) Then
                Set .directMusicSegment = .directMusicLoader.LoadSegment(PathToFile)
                .directMusicSegment.SetStandardMidiFile
                .directMusicSegment.Download .directMusicPath
            End If
        #End If
    End With
    
    load_midi = True
    Exit Function

Error_Handl:
    Debug.Print "error al cargar musica"
    load_midi = False

End Function



Private Function create_wav_buffer_from_file(ByVal file_name As Integer, ByRef new_buffer As DirectSoundSecondaryBuffer8, Optional ByVal Extraido As Boolean = False) As Boolean
On Error GoTo create_wav_buffer_from_file_Err
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
            If Extract_File(wav, get_compressed_resources_path(), file_name & ".wav", Windows_Temp_Dir, ResourcesPassword, False) Then
                Set new_buffer = mDirectSound.CreateSoundBufferFromFile(Windows_Temp_Dir & file_name & ".wav", dsbd)
                Delete_File Windows_Temp_Dir & filename_with_extension
                create_wav_buffer_from_file = True
            End If
        Else
            Set new_buffer = mDirectSound.CreateSoundBufferFromFile(Windows_Temp_Dir & file_name & ".wav", dsbd)
            Delete_File Windows_Temp_Dir & filename_with_extension
            create_wav_buffer_from_file = True

        End If
    #Else
        Dim PathToFile As String
        PathToFile = ao20audio.get_wav_files_path() & filename_with_extension
        If FileExist(PathToFile, vbArchive) Then
            Set new_buffer = mDirectSound.CreateSoundBufferFromFile(PathToFile, dsbd)
            create_wav_buffer_from_file = True
        End If

    #End If
   
    Exit Function

create_wav_buffer_from_file_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsAudioEngine.create_wav_buffer_from_file", Erl)
    Resume Next
    
End Function


Private Function start_playing(ByRef buffer As DirectSoundSecondaryBuffer8, ByVal flags As CONST_DSBPLAYFLAGS, ByVal volume As Long, ByVal pan As Long) As Long

On Error GoTo ErrorHandler
    
    Debug.Assert Not buffer Is Nothing

    Debug.Assert (buffer.GetStatus And DSBSTATUS_PLAYING) = 0
    
    Err.Clear
        
    If buffer Is Nothing Then
        start_playing = -1
        Exit Function
    End If
    
    
    If buffer.GetStatus = DSBSTATUS_PLAYING Then
'        buffer.Stop
'        buffer.SetCurrentPosition 0
        start_playing = 11
        Exit Function
    End If
    
    buffer.SetPan pan
    buffer.SetVolume volume
    buffer.Play flags
    
    start_playing = 0

    Exit Function

ErrorHandler:
    start_playing = False

End Function

Private Function find_track_by_name(ByVal Name As String) As Integer
        Dim i As Integer
        For i = 1 To UBound(mAudioTracks)
            With mAudioTracks(i)
                If Not .Buffer Is Nothing Then
                    Dim sf As CONST_DSBSTATUSFLAGS: sf = .Buffer.GetStatus
                    If (Name = .Name) And (DSBSTATUS_PLAYING And sf) = DSBSTATUS_PLAYING And (DSBSTATUS_LOOPING And sf) = DSBSTATUS_LOOPING Then
                        find_track_by_name = i
                        Exit Function
                    End If
                End If
            End With
        Next i
        find_track_by_name = -1
End Function


Private Function get_free_wav_track() As Integer
        Dim i As Integer
        
        For i = 1 To UBound(mAudioTracks)
            With mAudioTracks(i)
                If .buffer Is Nothing Then
                    get_free_wav_track = i
                    Exit Function
                Else
                    Dim sf As CONST_DSBSTATUSFLAGS: sf = .buffer.GetStatus
                    'Debug.Print "buffer status " & sf
                    If (DSBSTATUS_PLAYING And sf) = 0 And (DSBSTATUS_LOOPING And sf) = 0 Then
                        get_free_wav_track = i
                        Exit Function
                    End If
                    
                End If
            End With
        Next i
        get_free_wav_track = -1

End Function

Public Function play_wav(ByVal sound_id As Integer, Optional ByVal looping As Boolean = False, Optional ByVal volume As Long = 0, Optional ByVal pan As Long = 0) As Long
On Error GoTo Error_Handl
    Debug.Assert sound_id > 0
    play_wav = -1
    If (sound_id <= 0) Then Exit Function

    If mBuffer(sound_id) Is Nothing Then
        Dim i As Integer
        Call create_wav_buffer_from_file(sound_id, mBuffer(sound_id))
        Debug.Assert Not mBuffer(sound_id) Is Nothing
        If mBuffer(sound_id) Is Nothing Then
            play_wav = -1
            Exit Function
        End If
    End If
    
    Dim tid As Integer
    tid = get_free_wav_track()
    Debug.Print "play_wav, track " & tid & " id: " & sound_id
    If (tid > -1) Then
        With mAudioTracks(tid)
            .Name = sound_id
            Set .Buffer = mDirectSound.DuplicateSoundBuffer(ByVal mBuffer(sound_id))
            Call start_playing(.Buffer, IIf(looping, DSBPLAY_LOOPING, DSBPLAY_DEFAULT), volume, pan)
        End With
        play_wav = 0
    Else
        Debug.Print "play_wav failed because track id is -1"
    End If
    Exit Function

Error_Handl:

End Function


Public Function play_midi(ByVal id As Integer, Optional ByVal looping As Boolean = False, Optional ByVal volume As Long = 0) As Long
On Error GoTo Error_Handl
    Debug.Assert id > 0
    play_midi = -1
    If (id <= 0) Then Exit Function
    
    With mMidiTrack(1)
        If .Name <> id Then
            .Name = id
            Call load_midi(mMidiTrack(1), id)
            .directMusicPerformance.StopEx .directMusicSegment, 0, 0
            Set .directMusicSegmentState = .directMusicPerformance.PlaySegmentEx(.directMusicSegment, 0, 0, Nothing, .directMusicPath)
        End If
    End With

    play_midi = 0
    Exit Function

Error_Handl:

End Function

Public Function stop_wav(ByVal sound_id As Integer) As Long
    On Error GoTo Error_Handl
    Debug.Assert sound_id > 0
    stop_wav = -1
    If (sound_id <= 0) Then Exit Function

    Dim tid As Integer: tid = find_track_by_name(sound_id)
    If tid > -1 Then
     With mAudioTracks(tid)
            If Not .Buffer Is Nothing Then
                .Buffer.Stop
                .Buffer.SetCurrentPosition 0
            End If
     End With
    End If
        
    Exit Function
Error_Handl:
End Function

Public Sub stop_all_playback()
    Dim i As Integer
    For i = 1 To UBound(mAudioTracks)
        With mAudioTracks(i)
            If Not .buffer Is Nothing Then
                .buffer.Stop
                .buffer.SetCurrentPosition 0
            End If
        End With
    Next i
    
    For i = 1 To UBound(mMidiTrack)
            With mMidiTrack(i)
                .name = -1
                If Not .directMusicSegmentState Is Nothing Then
                       .directMusicPerformance.StopEx .directMusicSegment, 0, 0
                End If
            End With
    Next i
End Sub


Private Sub set_members_to_nothing()
    Call stop_all_playback
    
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
    Set mDirectX = Nothing
    Set mDirectSound = Nothing
    Set mDirectSoundEnum = Nothing
End Sub

Private Sub Class_Initialize()
set_members_to_nothing
End Sub


Private Sub Class_Terminate()
set_members_to_nothing
End Sub
