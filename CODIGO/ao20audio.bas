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

Public AudioEngine As clsAudioEngine
Public MusicEnabled As Boolean
Public FxEnabled As Boolean
Public AudioEnabled As Boolean
Public AmbientEnabled As Boolean
Public Sub CreateAudioEngine(ByVal hwnd As Long, ByRef dx8 As DirectX8, ByRef renderer As clsAudioEngine)
On Error GoTo AudioEngineInitErr:
    If AudioEnabled Then
        Set AudioEngine = New clsAudioEngine
        Call AudioEngine.init(dx8, hwnd)
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
Public Function StopAmbientAudio() As Long
    StopAmbientAudio = -1
    If AudioEnabled And AmbientEnabled Then
        StopAmbientAudio = ao20audio.AudioEngine.stop_ambient
    End If
End Function
Public Sub PlayAmbientAudio(ByVal UserMap As Long)
    If AudioEnabled And AmbientEnabled Then
        Dim wav As Integer
        If EsNoche Then
            wav = ReadField(1, Val(MapDat.ambient), Asc("-"))
        Else
            wav = ReadField(2, Val(MapDat.ambient), Asc("-"))
            If wav = 0 Then Exit Sub
        End If
    End If
End Sub
Public Function playwav(ByVal id As Integer, Optional ByVal looping As Boolean = False, Optional ByVal volume As Long = 0, Optional ByVal pan As Long = 0) As Long
    playwav = -1
    If AudioEnabled And FxEnabled Then
        PlayWav = ao20audio.AudioEngine.play_wav(id, looping, volume, pan)
    End If
End Function
Public Function stopwav(ByVal id As Integer) As Long
   stopwav = -1
    If AudioEnabled And FxEnabled Then
        StopWav = ao20audio.AudioEngine.stop_wav(id)
    End If
End Function
Public Function playmidi(ByVal id As Integer, Optional ByVal looping As Boolean = False, Optional ByVal volume As Long = 0) As Long
    playmidi = -1
    If AudioEnabled And MusicEnabled Then
        PlayMidi = ao20audio.AudioEngine.play_midi(id, looping, volume)
    End If
End Function
Public Sub stopallplayback()
    If AudioEnabled And MusicEnabled Then
        Call ao20audio.AudioEngine.stop_all_playback
    End If
End Sub
Public Function GetCurrentMidiName(ByVal track_id As Integer) As String
    If AudioEnabled And MusicEnabled Then
        GetCurrentMidiName = ao20audio.AudioEngine.get_midi_track_name(track_id)
    End If
End Function
Public Function GetWavFilesPath() As String
    GetWavFilesPath = App.path & "\..\Recursos\WAV\"
End Function
Public Function GetMp3FilesPath() As String
    GetMp3FilesPath = App.path & "\MP3\"
End Function
Public Function GetMidiFilesPath() As String
    GetMidiFilesPath = App.path & "/../Recursos/midi/"
End Function
Public Function GetCompressedResourcesPath() As String
 GetCompressedResourcesPath = App.path & "\..\Recursos\OUTPUT\"
End Function
Public Function ComputeCharfixVolume(ByRef Pos As Position) As Long
On Error GoTo ComputeCharfixVolumenErr:
    Dim total_distance As Integer, curr_x As Integer, curr_y As Integer
    total_distance = General_Distance_Get(Pos.x, Pos.y, UserPos.x, UserPos.y)
    If (total_distance = 0) Then
        ComputeCharfixVolume = VolFX
    ElseIf total_distance <= 19 Then
        ComputeCharfixVolume = VolFX - (total_distance * 120)
    End If
    If total_distance > 19 Then ComputeCharfixVolume = -4000
    If ComputeCharfixVolume < -4000 Then ComputeCharfixVolume = -4000
    Exit Function
ComputeCharfixVolumenErr:
    Call RegistrarError(Err.Number, Err.Description, "ComputeCharfixVolume", Erl)
    Resume Next
End Function
Public Function ComputeCharfixPan(ByRef Pos As Position) As Long
On Error GoTo ComputeCharfixPanErr:
    Dim total_distance As Integer, position_sgn As Integer, curr_x As Integer, curr_y As Integer
    ComputeCharfixPan = 0
    total_distance = General_Distance_Get(Pos.x, Pos.y, UserPos.x, UserPos.y)
    If InvertirSonido = True Then
        If Pos.x < UserPos.x Then
            position_sgn = -1
        Else
            position_sgn = 1
        End If
    Else
        If Pos.x > UserPos.x Then
            position_sgn = -1
        Else
            position_sgn = 1
        End If
    End If
    If (total_distance = 0) Or (Pos.x = UserPos.x) Then
        ComputeCharfixPan = 0
    ElseIf total_distance < 19 Then
        ComputeCharfixPan = position_sgn * (total_distance * 500)
    Else
        ComputeCharfixPan = position_sgn * 9000
    End If
    Exit Function
ComputeCharfixPanErr:
    Call RegistrarError(Err.Number, Err.Description, "ComputeCharfixPan", Erl)
    Resume Next
End Function

Public Function ComputeCharfixPanByDistance(ByVal total_distance As Integer, position_sgn As Integer) As Long
On Error GoTo ComputeCharfixPanByDistance_err:
    If InvertirSonido Then
        position_sgn = position_sgn * -1
    End If
    If (total_distance = 0) Or (position_sgn = 0) Then
        ComputeCharfixPanByDistance = 0
    ElseIf total_distance < 19 Then
        ComputeCharfixPanByDistance = position_sgn * (total_distance * 500)
    Else
        ComputeCharfixPanByDistance = position_sgn * 9000
    End If
    Exit Function

ComputeCharfixPanByDistance_err:
    Call RegistrarError(Err.Number, Err.Description, "clsSoundEngine.Calculate_Pan_By_Distance", Erl)
    Resume Next
End Function
Public Function ComputeCharfixVolumeByDistance(ByVal distance As Byte) As Long
On Error GoTo ComputeCharfixVolumeByDistance_err:
    distance = Abs(distance)
    If (distance = 0) Then
        ComputeCharfixVolumeByDistance = VolFX
    ElseIf distance <= 19 Then
        ComputeCharfixVolumeByDistance = VolFX - (distance * 120)
    End If
    If distance > 19 Then ComputeCharfixVolumeByDistance = -4000
    If ComputeCharfixVolumeByDistance < -4000 Then ComputeCharfixVolumeByDistance = -4000
    Exit Function
ComputeCharfixVolumeByDistance_err:
    Call RegistrarError(Err.Number, Err.Description, "ComputeCharfixVolumeByDistance", Erl)
    Resume Next
End Function



