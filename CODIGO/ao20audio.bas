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

'Sonidos
Public Const SND_EXCLAMACION   As Integer = 451
Public Const SND_CLICK         As String = 500
Public Const SND_CLICK_OVER    As String = 501
Public Const SND_NAVEGANDO     As Integer = 50
Public Const SND_OVER          As Integer = 0
Public Const SND_DICE          As Integer = 188
Public Const SND_FUEGO         As Integer = 116
Public Const SND_RAIN_IN_LOOP  As Integer = 191
Public Const SND_RAIN_OUT_LOOP As Integer = 194
Public Const SND_RAIN_IN_END   As Integer = 192
Public Const SND_RAIN_OUT_END  As Integer = 195
Public Const SND_NIEVEIN       As Integer = 191
Public Const SND_NIEVEOUT      As Integer = 194
Public Const SND_RESUCITAR     As Integer = 104
Public Const SND_CURAR         As Integer = 101
Public Const SND_DOPA          As Integer = 77
Public Const SND_MEDITATE      As Integer = 158

Public AudioEngine As clsAudioEngine
Public MusicEnabled As Byte
Public FxEnabled As Byte
Public AudioEnabled As Byte
Public AmbientEnabled As Byte
Private CurMusicVolume As Long
Private CurAmbientVolume As Long
Private CurFxVolume As Long

Public Sub CreateAudioEngine(ByVal hwnd As Long, ByRef dx8 As DirectX8, ByRef renderer As clsAudioEngine)
    On Error Goto CreateAudioEngine_Err
On Error GoTo AudioEngineInitErr:
    If AudioEnabled Then
        Set AudioEngine = New clsAudioEngine
        Call AudioEngine.Init(dx8, hwnd)
        frmDebug.add_text_tracebox "Audio Engine OK"
        Exit Sub
    Else
        frmDebug.add_text_tracebox "Warning Audio Disabled"
    End If
    
    Exit Sub
AudioEngineInitErr:
    Call MsgBox(JsonLanguage.Item("MENSAJEBOX_ERROR_CREACION_ENGINE_AUDIO"), vbCritical, "Argentum20")
    frmDebug.add_text_tracebox "Error Number Returned: " & Err.Number
    End
    Exit Sub
CreateAudioEngine_Err:
    Call TraceError(Err.Number, Err.Description, "ao20audio.CreateAudioEngine", Erl)
End Sub
Public Sub SetMusicVolume(ByVal NewVolume As Long)
    On Error Goto SetMusicVolume_Err
    CurMusicVolume = NewVolume
    If AudioEnabled And MusicEnabled And Not AudioEngine Is Nothing Then
        Call ao20audio.AudioEngine.ApplyMusicVolume(NewVolume)
    End If
    Exit Sub
SetMusicVolume_Err:
    Call TraceError(Err.Number, Err.Description, "ao20audio.SetMusicVolume", Erl)
End Sub

Public Sub SetAmbientVolume(ByVal NewVolume As Long)
    On Error Goto SetAmbientVolume_Err
    CurAmbientVolume = NewVolume
    Exit Sub
SetAmbientVolume_Err:
    Call TraceError(Err.Number, Err.Description, "ao20audio.SetAmbientVolume", Erl)
End Sub

Public Sub SetFxVolume(ByVal NewVolume As Long)
    On Error Goto SetFxVolume_Err
    CurFxVolume = NewVolume
    Exit Sub
SetFxVolume_Err:
    Call TraceError(Err.Number, Err.Description, "ao20audio.SetFxVolume", Erl)
End Sub

Public Function StopAmbientAudio() As Long
    On Error Goto StopAmbientAudio_Err
    StopAmbientAudio = -1
    If AudioEnabled And AmbientEnabled And Not AudioEngine Is Nothing Then
        StopAmbientAudio = ao20audio.AudioEngine.StopAmbient
    End If
    Exit Function
StopAmbientAudio_Err:
    Call TraceError(Err.Number, Err.Description, "ao20audio.StopAmbientAudio", Erl)
End Function

Public Sub PlayAmbientAudio(ByVal UserMap As Long)
    On Error Goto PlayAmbientAudio_Err
    If AudioEnabled And AmbientEnabled And Not AudioEngine Is Nothing Then
        Dim wav As Integer
        If EsNoche Then
            wav = ReadField(1, Val(MapDat.ambient), Asc("-"))
        Else
            wav = ReadField(2, Val(MapDat.ambient), Asc("-"))
        End If
        If wav <> 0 Then
            Call ao20audio.AudioEngine.PlayAmbient(wav, True, CurAmbientVolume)
        Else
            Call StopAmbientAudio
        End If
    End If
    Exit Sub
PlayAmbientAudio_Err:
    Call TraceError(Err.Number, Err.Description, "ao20audio.PlayAmbientAudio", Erl)
End Sub

Public Sub PlayWeatherAudio(ByVal id As Integer)
    On Error Goto PlayWeatherAudio_Err
    If AudioEnabled And AmbientEnabled And Not AudioEngine Is Nothing Then
        Call AudioEngine.PlayAmbient(id, True, CurAmbientVolume)
    End If
    Exit Sub
PlayWeatherAudio_Err:
    Call TraceError(Err.Number, Err.Description, "ao20audio.PlayWeatherAudio", Erl)
End Sub

Public Function PlayAmbientWav(ByVal id As Integer, Optional ByVal looping As Boolean = False, Optional ByVal pan As Long = 0, Optional ByVal label As String = "") As Long
    On Error Goto PlayAmbientWav_Err
    PlayAmbientWav = -1
    If AudioEnabled And AmbientEnabled And Not AudioEngine Is Nothing Then
        PlayAmbientWav = ao20audio.AudioEngine.PlayWav(id, looping, CurAmbientVolume, pan, label)
    End If
    Exit Function
PlayAmbientWav_Err:
    Call TraceError(Err.Number, Err.Description, "ao20audio.PlayAmbientWav", Erl)
End Function

Public Function PlayWav(ByVal id As String, Optional ByVal looping As Boolean = False, Optional ByVal volume As Long = 0, Optional ByVal pan As Long = 0, Optional ByVal label As String = "") As Long
    On Error Goto PlayWav_Err
    PlayWav = -1
    If AudioEnabled And FxEnabled And Not AudioEngine Is Nothing Then
        PlayWav = ao20audio.AudioEngine.PlayWav(id, looping, min(CurFxVolume, volume), pan, label)
    End If
    Exit Function
PlayWav_Err:
    Call TraceError(Err.Number, Err.Description, "ao20audio.PlayWav", Erl)
End Function

Public Function StopMP3() As Long
    On Error Goto StopMP3_Err
    StopMP3 = -1
    If AudioEnabled And MusicEnabled And Not AudioEngine Is Nothing Then
        StopMP3 = ao20audio.AudioEngine.StopMP3
    End If
    Exit Function
StopMP3_Err:
    Call TraceError(Err.Number, Err.Description, "ao20audio.StopMP3", Erl)
End Function

Public Function PlayMP3(ByVal filename As String, Optional ByVal looping As Boolean = False, Optional ByVal volume As Long = 0) As Long
    On Error Goto PlayMP3_Err
    PlayMP3 = -1
    If AudioEnabled And MusicEnabled And Not AudioEngine Is Nothing Then
        PlayMP3 = ao20audio.AudioEngine.PlayMP3(FileName, looping, min(CurMusicVolume, volume))
    End If
    Exit Function
PlayMP3_Err:
    Call TraceError(Err.Number, Err.Description, "ao20audio.PlayMP3", Erl)
End Function

Public Function StopWav(ByVal id As String, Optional ByVal label As String = "") As Long
    On Error Goto StopWav_Err
   StopWav = -1
    If AudioEnabled And FxEnabled And Not AudioEngine Is Nothing Then
        StopWav = ao20audio.AudioEngine.StopWav(id, label)
    End If
    Exit Function
StopWav_Err:
    Call TraceError(Err.Number, Err.Description, "ao20audio.StopWav", Erl)
End Function

Public Function StopAllWavsMatchingLabel(ByVal label As String) As Long
    On Error Goto StopAllWavsMatchingLabel_Err
    StopAllWavsMatchingLabel = -1
    If AudioEnabled And FxEnabled And Not AudioEngine Is Nothing Then
        StopAllWavsMatchingLabel = ao20audio.AudioEngine.StopAllWavsMatchingLabel(label)
    End If
    Exit Function
StopAllWavsMatchingLabel_Err:
    Call TraceError(Err.Number, Err.Description, "ao20audio.StopAllWavsMatchingLabel", Erl)
End Function

Public Function PlayMidi(ByVal id As Integer, Optional ByVal looping As Boolean = False, Optional ByVal volume As Long = 0) As Long
    On Error Goto PlayMidi_Err
    PlayMidi = -1
    If AudioEnabled And MusicEnabled And Not AudioEngine Is Nothing Then
        PlayMidi = ao20audio.AudioEngine.PlayMidi(id, looping, CurMusicVolume)
    End If
    Exit Function
PlayMidi_Err:
    Call TraceError(Err.Number, Err.Description, "ao20audio.PlayMidi", Erl)
End Function

Public Sub StopAllPlayback()
    On Error Goto StopAllPlayback_Err
    If AudioEnabled And MusicEnabled And Not AudioEngine Is Nothing Then
        Call ao20audio.AudioEngine.StopAllPlayback
    End If
    Exit Sub
StopAllPlayback_Err:
    Call TraceError(Err.Number, Err.Description, "ao20audio.StopAllPlayback", Erl)
End Sub

Public Function GetCurrentMidiName(ByVal track_id As Integer) As String
    On Error Goto GetCurrentMidiName_Err
    If AudioEnabled And MusicEnabled And Not AudioEngine Is Nothing Then
        GetCurrentMidiName = ao20audio.AudioEngine.GetMidiTrackName(track_id)
    End If
    Exit Function
GetCurrentMidiName_Err:
    Call TraceError(Err.Number, Err.Description, "ao20audio.GetCurrentMidiName", Erl)
End Function

Public Function GetWavFilesPath() As String
    On Error Goto GetWavFilesPath_Err
    GetWavFilesPath = App.path & "\..\Recursos\WAV\"
    Exit Function
GetWavFilesPath_Err:
    Call TraceError(Err.Number, Err.Description, "ao20audio.GetWavFilesPath", Erl)
End Function

Public Function GetMp3FilesPath() As String
    On Error Goto GetMp3FilesPath_Err
    GetMp3FilesPath = App.path & "\..\Recursos\MP3\"
    Exit Function
GetMp3FilesPath_Err:
    Call TraceError(Err.Number, Err.Description, "ao20audio.GetMp3FilesPath", Erl)
End Function

Public Function GetMidiFilesPath() As String
    On Error Goto GetMidiFilesPath_Err
    GetMidiFilesPath = App.path & "\..\Recursos\MIDI\"
    Exit Function
GetMidiFilesPath_Err:
    Call TraceError(Err.Number, Err.Description, "ao20audio.GetMidiFilesPath", Erl)
End Function

Public Function GetCompressedResourcesPath() As String
    On Error Goto GetCompressedResourcesPath_Err
 GetCompressedResourcesPath = App.path & "\..\Recursos\OUTPUT\"
    Exit Function
GetCompressedResourcesPath_Err:
    Call TraceError(Err.Number, Err.Description, "ao20audio.GetCompressedResourcesPath", Erl)
End Function

Public Function ComputeCharFxVolume(ByRef Pos As Position) As Long
    On Error Goto ComputeCharFxVolume_Err
On Error GoTo ComputeCharFxVolumenErr:
    Dim total_distance As Integer
    total_distance = General_Distance_Get(Pos.x, Pos.y, UserPos.x, UserPos.y)
    ComputeCharFxVolume = ComputeCharFxVolumeByDistance(total_distance)
    Exit Function
ComputeCharFxVolumenErr:
    Call RegistrarError(Err.Number, Err.Description, "ComputeCharFxVolume", Erl)
    Resume Next
    Exit Function
ComputeCharFxVolume_Err:
    Call TraceError(Err.Number, Err.Description, "ao20audio.ComputeCharFxVolume", Erl)
End Function

Public Function ComputeCharFxPan(ByRef Pos As Position) As Long
    On Error Goto ComputeCharFxPan_Err
On Error GoTo ComputeCharFxPanErr:
    Dim total_distance As Integer, position_sgn As Integer, curr_x As Integer, curr_y As Integer
    ComputeCharFxPan = 0
    total_distance = General_Distance_Get(Pos.x, Pos.y, UserPos.x, UserPos.y)
    If InvertirSonido = False Then
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
        ComputeCharFxPan = 0
    ElseIf total_distance < 19 Then
        ComputeCharFxPan = position_sgn * (total_distance * 500)
    Else
        ComputeCharFxPan = position_sgn * 9000
    End If
    Exit Function
ComputeCharFxPanErr:
    Call RegistrarError(Err.Number, Err.Description, "ComputeCharFxPan", Erl)
    Resume Next
    Exit Function
ComputeCharFxPan_Err:
    Call TraceError(Err.Number, Err.Description, "ao20audio.ComputeCharFxPan", Erl)
End Function

Public Function ComputeCharFxPanByDistance(ByVal total_distance As Integer, position_sgn As Integer) As Long
    On Error Goto ComputeCharFxPanByDistance_Err
On Error GoTo ComputeCharFxPanByDistance_err:
    If InvertirSonido Then
        position_sgn = position_sgn * -1
    End If
    If (total_distance = 0) Or (position_sgn = 0) Then
        ComputeCharFxPanByDistance = 0
    ElseIf total_distance < 19 Then
        ComputeCharFxPanByDistance = position_sgn * (total_distance * 500)
    Else
        ComputeCharFxPanByDistance = position_sgn * 9000
    End If
    Exit Function

ComputeCharFxPanByDistance_err:
    Call RegistrarError(Err.Number, Err.Description, "clsSoundEngine.Calculate_Pan_By_Distance", Erl)
    Resume Next
    Exit Function
ComputeCharFxPanByDistance_Err:
    Call TraceError(Err.Number, Err.Description, "ao20audio.ComputeCharFxPanByDistance", Erl)
End Function

Public Function ComputeCharFxVolumeByDistance(ByVal distance As Byte) As Long
    On Error Goto ComputeCharFxVolumeByDistance_Err
On Error GoTo ComputeCharFxVolumeByDistance_err:
    distance = Abs(distance)
    If distance < 20 Then
        ComputeCharFxVolumeByDistance = VolFX - distance * 120
        If ComputeCharFxVolumeByDistance < -4000 Then ComputeCharFxVolumeByDistance = -4000
    Else
        ComputeCharFxVolumeByDistance = -4000
    End If
    Exit Function
ComputeCharFxVolumeByDistance_err:
    Call RegistrarError(Err.Number, Err.Description, "ComputeCharFxVolumeByDistance", Erl)
    Resume Next
    Exit Function
ComputeCharFxVolumeByDistance_Err:
    Call TraceError(Err.Number, Err.Description, "ao20audio.ComputeCharFxVolumeByDistance", Erl)
End Function


