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
Public AudioEngine             As clsAudioEngine
Public MusicEnabled            As Byte
Public FxEnabled               As Byte
Public AudioEnabled            As Byte
Public AmbientEnabled          As Byte
Public FxStepsEnabled          As Byte
Private CurMusicVolume         As Long
Private CurAmbientVolume       As Long
Private CurFxVolume            As Long
Private CurStepsVolume         As Long

Public Sub CreateAudioEngine(ByVal hWnd As Long, ByRef dx8 As DirectX8, ByRef renderer As clsAudioEngine)
    On Error GoTo AudioEngineInitErr:
    If AudioEnabled Then
        Set AudioEngine = New clsAudioEngine
        Call AudioEngine.Init(dx8, hWnd)
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
End Sub

Public Sub SetMusicVolume(ByVal NewVolume As Long)
    CurMusicVolume = NewVolume
    If AudioEnabled And MusicEnabled And Not AudioEngine Is Nothing Then
        Call ao20audio.AudioEngine.ApplyMusicVolume(NewVolume)
    End If
End Sub

Public Sub SetAmbientVolume(ByVal NewVolume As Long)
    CurAmbientVolume = NewVolume
    If AudioEnabled And AmbientEnabled And Not AudioEngine Is Nothing Then
        Call ao20audio.AudioEngine.ApplyAmbientVolume(NewVolume)
    End If
End Sub

Public Sub SetFxVolume(ByVal NewVolume As Long)
    CurFxVolume = NewVolume
End Sub

Public Sub SetVolumeSteps(ByVal NewVolume As Long)
    CurStepsVolume = NewVolume
End Sub

Public Function StopAmbientAudio() As Long
    StopAmbientAudio = -1
    If AudioEnabled > 0 And Not AudioEngine Is Nothing Then
        StopAmbientAudio = ao20audio.AudioEngine.StopAmbient
    End If
End Function

Public Sub PlayAmbientAudio(ByVal UserMap As Long)
    If AudioEnabled = 0 Or AmbientEnabled = 0 Or AudioEngine Is Nothing Then
        Exit Sub
    End If
    Dim wav As Integer
    If EsNoche Then
        wav = ReadField(1, val(MapDat.ambient), Asc("-"))
    Else
        wav = ReadField(2, val(MapDat.ambient), Asc("-"))
    End If
    If wav <> 0 Then
        Call ao20audio.AudioEngine.PlayAmbient(wav, True, CurAmbientVolume)
    Else
        Call StopAmbientAudio
    End If
End Sub

Public Sub PlayWeatherAudio(ByVal id As Integer)
    If AudioEnabled And AmbientEnabled And Not AudioEngine Is Nothing Then
        Call AudioEngine.PlayAmbient(id, True, CurAmbientVolume)
    End If
End Sub

Public Function PlayAmbientWav(ByVal id As Integer, Optional ByVal looping As Boolean = False, Optional ByVal pan As Long = 0, Optional ByVal label As String = "") As Long
    PlayAmbientWav = -1
    If AudioEnabled And AmbientEnabled And Not AudioEngine Is Nothing Then
        PlayAmbientWav = ao20audio.AudioEngine.PlayWav(id, looping, CurAmbientVolume, pan, label)
    End If
End Function

Public Function PlayWav(ByVal id As String, _
                        Optional ByVal looping As Boolean = False, _
                        Optional ByVal volume As Long = 0, _
                        Optional ByVal pan As Long = 0, _
                        Optional ByVal label As String = "") As Long
    PlayWav = -1
    If AudioEnabled And FxEnabled And Not AudioEngine Is Nothing Then
        PlayWav = ao20audio.AudioEngine.PlayWav(id, looping, min(CurFxVolume, volume), pan, label)
    End If
End Function
Public Function PlayFx(ByVal id As String, _
                       ByVal category As eFxCategory, _
                       Optional ByVal looping As Boolean = False, _
                       Optional ByVal volume As Long = 0, _
                       Optional ByVal pan As Long = 0, _
                       Optional ByVal label As String = "") As Long
    PlayFx = -1
    If AudioEngine Is Nothing Or AudioEnabled = 0 Then Exit Function

    Dim effVol As Long
    Select Case category
        Case eFxSteps
            If FxStepsEnabled = 0 Then Exit Function
            effVol = min(CurStepsVolume, volume)

        Case eFxAmbient
            If AmbientEnabled = 0 Then Exit Function
            effVol = min(CurAmbientVolume, volume)

        Case Else
            If FxEnabled = 0 Then Exit Function
            effVol = min(CurFxVolume, volume)
    End Select

    PlayFx = AudioEngine.PlayWav(id, looping, effVol, pan, label)
End Function

Public Function StopMP3() As Long
    StopMP3 = -1
    If AudioEnabled And MusicEnabled And Not AudioEngine Is Nothing Then
        StopMP3 = ao20audio.AudioEngine.StopMP3
    End If
End Function

Public Function PlayMP3(ByVal filename As String, Optional ByVal looping As Boolean = False, Optional ByVal volume As Long = 0) As Long
    PlayMP3 = -1
    If AudioEnabled And MusicEnabled And Not AudioEngine Is Nothing Then
        PlayMP3 = ao20audio.AudioEngine.PlayMP3(filename, looping, min(CurMusicVolume, volume))
    End If
End Function

Public Function StopWav(ByVal id As String, Optional ByVal label As String = "") As Long
    StopWav = -1
    If AudioEnabled And FxEnabled And Not AudioEngine Is Nothing Then
        StopWav = ao20audio.AudioEngine.StopWav(id, label)
    End If
End Function

Public Function StopAllWavsMatchingLabel(ByVal label As String) As Long
    StopAllWavsMatchingLabel = -1
    If AudioEnabled And FxEnabled And Not AudioEngine Is Nothing Then
        StopAllWavsMatchingLabel = ao20audio.AudioEngine.StopAllWavsMatchingLabel(label)
    End If
End Function

Public Function PlayMidi(ByVal id As Integer, Optional ByVal looping As Boolean = False, Optional ByVal volume As Long = 0) As Long
    PlayMidi = -1
    If AudioEnabled And MusicEnabled And Not AudioEngine Is Nothing Then
        PlayMidi = ao20audio.AudioEngine.PlayMidi(id, looping, CurMusicVolume)
    End If
End Function

Public Sub StopAllPlayback()
    If AudioEnabled And MusicEnabled And Not AudioEngine Is Nothing Then
        Call ao20audio.AudioEngine.StopAllPlayback
    End If
End Sub

Public Function GetCurrentMidiName(ByVal track_id As Integer) As String
    If AudioEnabled And MusicEnabled And Not AudioEngine Is Nothing Then
        GetCurrentMidiName = ao20audio.AudioEngine.GetMidiTrackName(track_id)
    End If
End Function

Public Function GetWavFilesPath() As String
    GetWavFilesPath = App.path & "\..\Recursos\WAV\"
End Function

Public Function GetMp3FilesPath() As String
    GetMp3FilesPath = App.path & "\..\Recursos\MP3\"
End Function

Public Function GetMidiFilesPath() As String
    GetMidiFilesPath = App.path & "\..\Recursos\MIDI\"
End Function

Public Function GetCompressedResourcesPath() As String
    GetCompressedResourcesPath = App.path & "\..\Recursos\OUTPUT\"
End Function

Public Function ComputeCharFxVolume(ByRef Pos As Position) As Long
    On Error GoTo ComputeCharFxVolumenErr:
    Dim total_distance As Integer
    total_distance = General_Distance_Get(Pos.x, Pos.y, UserPos.x, UserPos.y)
    ComputeCharFxVolume = ComputeVolumeByDistance(eFxGeneral, total_distance)
    Exit Function
ComputeCharFxVolumenErr:
    Call RegistrarError(Err.Number, Err.Description, "ComputeCharFxVolume", Erl)
    Resume Next
End Function

Public Function ComputeCharFxPan(ByRef Pos As Position) As Long
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
End Function

Public Function ComputeCharFxPanByDistance(ByVal total_distance As Integer, position_sgn As Integer) As Long
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
End Function
Public Function ComputeVolumeByDistance(ByVal category As eFxCategory, ByVal distance As Integer) As Long
    On Error GoTo ComputeVolumeByDistance_err
    distance = Abs(distance)

    Dim base As Long
    Select Case category
        Case eFxSteps:   base = VolSteps
        Case eFxAmbient: base = VolAmbient
        Case Else:       base = VolFX
    End Select

    If distance < 20 Then
        ComputeVolumeByDistance = base - distance * 120
        If ComputeVolumeByDistance < -4000 Then ComputeVolumeByDistance = -4000
    Else
        ComputeVolumeByDistance = -4000
    End If
    Exit Function
ComputeVolumeByDistance_err:
    Call RegistrarError(Err.Number, Err.Description, "ComputeVolumeByDistance", Erl)
End Function

Public Function ComputeVolumeAtPos(ByVal category As eFxCategory, ByRef Pos As Position) As Long
    On Error GoTo ComputeVolumeAtPos_err
    Dim total_distance As Integer
    total_distance = General_Distance_Get(Pos.x, Pos.y, UserPos.x, UserPos.y)
    ComputeVolumeAtPos = ComputeVolumeByDistance(category, total_distance)
    Exit Function
ComputeVolumeAtPos_err:
    Call RegistrarError(Err.Number, Err.Description, "ComputeVolumeAtPos", Erl)
End Function
