Attribute VB_Name = "Animations"
Option Explicit
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
Type tAnimationPlaybackState
    PlaybackState As ePlaybackType
    CurrentGrh As Long
    CurrentFrame As Long
    ActiveClip As Long
    ComposedAnimation As Long
    CurrentClipLoops As Long
    LastFrameTime As Long
    ElapsedTime As Long
    Alpha As Boolean
    AlphaValue As Byte
    Fx As Long 'support
End Type

Public Sub UpdateAnimation(ByRef animationState As tAnimationPlaybackState)
On Error GoTo UpdateAnimation_Err
    Dim DeltaTime As Long
    DeltaTime = GetTickCount() - animationState.LastFrameTime
    animationState.LastFrameTime = GetTickCount()
    animationState.ElapsedTime = animationState.ElapsedTime + DeltaTime
    If animationState.Fx <> 0 Then
        Call UpdateFx(animationState)
    Else
        Call UpdateClip(animationState)
    End If
    Exit Sub
UpdateAnimation_Err:
    Call RegistrarError(Err.Number, Err.Description, "animations.UpdateAnimation", Erl)
    Resume Next
End Sub

Sub UpdateClip(ByRef animationState As tAnimationPlaybackState)
    Dim DeltaTime As Long
    With ComposedFxData(animationState.ComposedAnimation).Clips(animationState.ActiveClip)
        If (animationState.ElapsedTime >= .ClipTime) Then
            DeltaTime = animationState.ElapsedTime Mod .ClipTime
            animationState.CurrentClipLoops = animationState.CurrentClipLoops + 1
            If (.LoopCount >= 0 And animationState.CurrentClipLoops >= .LoopCount) Then
                Call StartNextClip(animationState)
            End If
            animationState.ElapsedTime = DeltaTime
        End If
        Dim progress As Single
        progress = animationState.ElapsedTime / .ClipTime
        If .Playback = Backward Then
            progress = 1 - progress
        End If
        With GrhData(animationState.CurrentGrh)
            animationState.CurrentFrame = (.NumFrames - 1) * progress + 1
        End With
        
    End With
End Sub

Sub UpdateFx(ByRef animationState As tAnimationPlaybackState)
    On Error GoTo UpdateFx_Err
    Dim DeltaTime As Long
100     With GrhData(animationState.CurrentGrh)
102         If (animationState.ElapsedTime >= .speed) Then
104             DeltaTime = animationState.ElapsedTime Mod .speed
106             If (animationState.CurrentClipLoops = 0) Then
108                 animationState.PlaybackState = Stopped
                    Exit Sub
                End If
110             animationState.CurrentClipLoops = animationState.CurrentClipLoops - 1
112             animationState.ElapsedTime = DeltaTime
            End If
            Dim progress As Single
114         progress = animationState.ElapsedTime / .speed
116         animationState.CurrentFrame = (.NumFrames - 1) * progress + 1
        End With
        Exit Sub
UpdateFx_Err:
118     Call RegistrarError(Err.Number, Err.Description, "animations.UpdateFx", Erl)
120     Resume Next
End Sub

Sub Initialize(ByRef animationState As tAnimationPlaybackState)
    animationState.CurrentFrame = 1
    animationState.CurrentClipLoops = 0
    animationState.ElapsedTime = 0
    animationState.PlaybackState = Forward
    animationState.Alpha = False
    animationState.AlphaValue = 180
    animationState.LastFrameTime = GetTickCount()
End Sub

Public Function IsLoopActive(ByRef animationState As tAnimationPlaybackState) As Boolean
    If animationState.PlaybackState = Stopped Or animationState.PlaybackState = Pause Then
        IsLoopActive = False
        Exit Function
    End If
    With ComposedFxData(animationState.ComposedAnimation).Clips(animationState.ActiveClip)
        IsLoopActive = .LoopCount < 0
    End With
End Function

Public Sub StartFx(ByRef animationState As tAnimationPlaybackState, ByVal Fx As Long, Optional ByVal loopC As Integer = 0)
    On Error GoTo StartFx_Err
100     If Fx = 0 Then
102         animationState.PlaybackState = Stopped
            Exit Sub
        End If
        Dim AnimationId As Integer
104     AnimationId = FxToAnimationMap(Fx)
106     If AnimationId > 0 Then
            'dont restart a looping animation
108         If AnimationId = animationState.ComposedAnimation And IsLoopActive(animationState) Then
                Exit Sub
            End If
110         animationState.ComposedAnimation = AnimationId
112         Call StartAnimation(animationState, animationState.ComposedAnimation)
            Exit Sub
        End If
114     Call Initialize(animationState)
116     animationState.Fx = Fx
118     animationState.CurrentGrh = FxData(Fx).Animacion
120     animationState.CurrentClipLoops = loopC
        Exit Sub
StartFx_Err:
122     Call RegistrarError(Err.Number, Err.Description, "animations.StartFx", Erl)
124     Resume Next
End Sub

Sub StartAnimation(ByRef animationState As tAnimationPlaybackState, ByVal composedAnimationIndex As Long)
On Error GoTo StartAnimation_Err
    If composedAnimationIndex = 0 Then
        animationState.PlaybackState = Stopped
        Exit Sub
    End If
    animationState.ComposedAnimation = composedAnimationIndex
    animationState.ActiveClip = 1
    animationState.CurrentGrh = FxData(ComposedFxData(composedAnimationIndex).Clips(animationState.ActiveClip).fX).Animacion
    Call Initialize(animationState)
    animationState.LastFrameTime = GetTickCount()
    animationState.Fx = 0
    Exit Sub
StartAnimation_Err:
    Call RegistrarError(Err.Number, Err.Description, "animations.StartAnimation", Erl)
    Resume Next
End Sub

Public Sub StartNextClip(ByRef animationState As tAnimationPlaybackState)
On Error GoTo StartNextClip_Err
    Call ChangeToClip(animationState, animationState.ActiveClip + 1)
    Exit Sub
StartNextClip_Err:
    Call RegistrarError(Err.Number, Err.Description, "animations.StartNextClip", Erl)
    Resume Next
End Sub

Public Sub ChangeToClip(ByRef animationState As tAnimationPlaybackState, ByVal clipIndex As Integer)
    On Error GoTo ChangeToClip_Err
100     If animationState.Fx > 0 Then
102         animationState.PlaybackState = Stopped
            Exit Sub
        End If
104     animationState.ActiveClip = clipIndex
106     animationState.CurrentClipLoops = 0
108     With ComposedFxData(animationState.ComposedAnimation)
110         If animationState.ActiveClip > UBound(.Clips) Then
112             animationState.PlaybackState = Stopped
                Exit Sub
            End If
114         animationState.CurrentGrh = FxData(.Clips(animationState.ActiveClip).Fx).Animacion
116         animationState.CurrentFrame = 1
118         animationState.ElapsedTime = 0
120         animationState.LastFrameTime = GetTickCount()
122         animationState.ElapsedTime = 0
        End With
        Exit Sub
ChangeToClip_Err:
124     Call RegistrarError(Err.Number, Err.Description, "animations.ChangeToClip composed animation: " & animationState.ComposedAnimation & " and clip index: " & clipIndex, Erl)
126     Resume Next
End Sub

Public Function GetFx(ByRef animationState As tAnimationPlaybackState) As Integer
    If animationState.Fx > 0 Then
        GetFx = animationState.Fx
    Else
        GetFx = ComposedFxData(animationState.ComposedAnimation).Clips(animationState.ActiveClip).Fx
    End If
End Function
