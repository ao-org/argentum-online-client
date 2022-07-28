Attribute VB_Name = "Animations"
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
    Dim detalTime As Long
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
    With GrhData(animationState.CurrentGrh)
        If (animationState.ElapsedTime >= .speed) Then
            DeltaTime = animationState.ElapsedTime Mod .speed
            If (animationState.CurrentClipLoops = 0) Then
                PlaybackState = Stopped
                Exit Sub
            End If
            animationState.CurrentClipLoops = animationState.CurrentClipLoops - 1
            animationState.ElapsedTime = DeltaTime
        End If
        Dim progress As Single
        progress = animationState.ElapsedTime / .speed
        animationState.CurrentFrame = (.NumFrames - 1) * progress + 1
    End With
    Exit Sub
UpdateFx_Err:
    Call RegistrarError(Err.Number, Err.Description, "animations.UpdateFx", Erl)
    Resume Next
End Sub

Sub Initialize(ByRef animationState As tAnimationPlaybackState)
    animationState.CurrentFrame = 1
    animationState.CurrentClipLoops = 0
    animationState.ElapsedTime = 0
    animationState.PlaybackState = Forward
    animationState.Alpha = 180
    animationState.LastFrameTime = GetTickCount()
End Sub

Public Sub StartFx(ByRef animationState As tAnimationPlaybackState, ByVal Fx As Long, Optional ByVal loopC As Integer = 0)
On Error GoTo StartFx_Err
    If Fx = 0 Then
        animationState.PlaybackState = Stopped
        Exit Sub
    End If
    animationState.ComposedAnimation = FxToAnimationMap(Fx)
    If animationState.ComposedAnimation > 0 Then
        Call StartAnimation(animationState, animationState.ComposedAnimation)
        Exit Sub
    End If
    Call Initialize(animationState)
    animationState.Fx = Fx
    animationState.CurrentGrh = FxData(Fx).Animacion
    animationState.CurrentClipLoops = loopC
    Exit Sub
StartFx_Err:
    Call RegistrarError(Err.Number, Err.Description, "animations.StartFx", Erl)
    Resume Next
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
    If animationState.Fx > 0 Then
        animationState.PlaybackState = Stopped
        Exit Sub
    End If
    animationState.ActiveClip = clipIndex
    animationState.CurrentClipLoops = 0
    With ComposedFxData(animationState.ComposedAnimation)
        If animationState.ActiveClip > UBound(.Clips) Then
            animationState.PlaybackState = Stopped
            Exit Sub
        End If
        animationState.CurrentGrh = FxData(.Clips(animationState.ActiveClip).fX).Animacion
        animationState.CurrentFrame = 1
        animationState.ElapsedTime = 0
        animationState.LastFrameTime = GetTickCount()
        animationState.ElapsedTime = 0
    End With
    Exit Sub
ChangeToClip_Err:
    Call RegistrarError(Err.Number, Err.Description, "animations.ChangeToClip", Erl)
    Resume Next
End Sub

Public Function GetFx(ByRef animationState As tAnimationPlaybackState) As Integer
    If animationState.Fx > 0 Then
        GetFx = animationState.Fx
    Else
        GetFx = ComposedFxData(animationState.ComposedAnimation).Clips(animationState.ActiveClip).Fx
    End If
End Function
