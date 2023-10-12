Attribute VB_Name = "ao20audio"
'    Argentum 20 - Game Client Program
'    Copyright (C) 2022 - Noland Studios
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
'
'
Option Explicit

Public audio_engine                         As clsAudioEngine

Public MusicEnabled As Byte
Public FxEnabled As Byte
Public AudioEnabled As Byte
    
Public Sub create_audio_engine(ByVal hwnd As Long, ByRef dx8 As DirectX8, ByRef renderer As clsAudioEngine)
On Error GoTo AudioEngineInitErr:
    If AudioEnabled Then
        Set audio_engine = New clsAudioEngine
        Call audio_engine.init(dx8, hwnd)
        Debug.Print "Audio Engine OK"
        Exit Sub
    Else
        Debug.Print "Warning Audio Disabled"
    End If
    
    Exit Sub
AudioEngineInitErr:
    Call MsgBox("Error creating audio engine", vbCritical, "Argentum20")
    Debug.Print "Error Number Returned: " & Err.Number
    End
End Sub
Public Function playwav(ByVal id As Integer, Optional ByVal looping As Boolean = False, Optional ByVal volume As Long = 0, Optional ByVal pan As Long = 0) As Long
    playwav = -1
    If Val(AudioEnabled) > 0 And Val(FxEnabled) > 0 Then
        playwav = ao20audio.audio_engine.play_wav(id, looping, volume, pan)
    End If
End Function
Public Function stopwav(ByVal id As Integer) As Long
   stopwav = -1
    If Val(AudioEnabled) > 0 And Val(FxEnabled) > 0 Then
        stopwav = ao20audio.audio_engine.stop_wav(id)
    End If
End Function

Public Function playmidi(ByVal id As Integer, Optional ByVal looping As Boolean = False, Optional ByVal volume As Long = 0) As Long
    playmidi = -1
    If Val(AudioEnabled) > 0 And Val(MusicEnabled) > 0 Then
        playmidi = ao20audio.audio_engine.play_midi(id, looping, volume)
    End If
End Function
Public Sub stopallplayback()
    If Val(AudioEnabled) > 0 And Val(MusicEnabled) > 0 Then
        Call ao20audio.audio_engine.stop_all_playback
    End If
End Sub
Public Function get_current_midi_name(ByVal track_id As Integer) As String
    If Val(AudioEnabled) > 0 And Val(MusicEnabled) > 0 Then
        get_current_midi_name = ao20audio.audio_engine.get_midi_track_name(track_id)
    End If
End Function
Public Function get_wav_files_path() As String
    get_wav_files_path = App.path & "\..\Recursos\WAV\"
End Function

Public Function get_mp3_files_path() As String
    get_mp3_files_path = App.path & "\MP3\"
End Function

Public Function get_midi_files_path() As String
 get_midi_files_path = App.path & "/../Recursos/midi/"
End Function

Public Function get_compressed_resources_path() As String
 get_compressed_resources_path = App.path & "\..\Recursos\OUTPUT\"
End Function




