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
End Type


Public Sub UpdateAnimation(ByRef animationState As tAnimationPlaybackState)
On Error GoTo UpdateAnimation_Err
    Dim detalTime As Long
    DeltaTime = GetTickCount() - animationState.LastFrameTime
    ElapsedTime = ElapsedTime + DeltaTime
    With ComposedFxData(animationState.ComposedAnimation).Clips(animationState.ActiveClip)
        If (ElapsedTime > .ClipTime) Then
            DeltaTime = ElapsedTime - .ClipTime
            animationState.CurrentClipLoops = animationState.CurrentClipLoops + 1
            If (.LoopCount > 0 And animationState.CurrentClipLoops >= .LoopCount) Then
                Call StartNextClip(animationState)
            End If
            animationState.ElapsedTime = DeltaTime
        End If
        Dim progress As Single
        progress = ElapsedTime / .ClipTime
        If .Playback = Backward Then
            progress = 1 - progress
        End If
        With GrhData(animationState.CurrentGrh)
            animationState.CurrentFrame = .NumFrames * progress
        End With
        
    End With
UpdateAnimation_Err:
    Call RegistrarError(Err.Number, Err.Description, "animations.UpdateAnimation", Erl)
    Resume Next
End Sub

Public Sub StartAnimation(ByRef animationState As tAnimationPlaybackState, ByVal composedAnimationIndex As Long)
On Error GoTo StartAnimation_Err
    animationState.PlaybackState = Forward
    animationState.ComposedAnimation = composedAnimationIndex
    animationState.ActiveClip = 1
    animationState.CurrentGrh = ComposedFxData(composedAnimationIndex).Clips(animationState.ActiveClip).Fx
    animationState.CurrentFrame = GrhData(animationState.CurrentGrh).Frames(1)
    animationState.CurrentClipLoops = 0
StartAnimation_Err:
    Call RegistrarError(Err.Number, Err.Description, "animations.StartAnimation", Erl)
    Resume Next
End Sub

Public Sub StartNextClip(ByRef animationState As tAnimationPlaybackState)
On Error GoTo StartNextClip_Err
    animationState.ActiveClip = animationState.ActiveClip + 1
    animationState.CurrentClipLoops = 0
    With ComposedFxData(animationState.ComposedAnimation)
        If animationState.ActiveClip >= UBound(.Clips) Then
            animationState.PlaybackState = Complete
            Exit Sub
        End If
        animationState.CurrentGrh = .Clips(animationState.ActiveClip).Fx
        
        animationState.CurrentFrame = GrhData(animationState.CurrentGrh).Frames(1)
        LastFrameTime = GetTickCount()
        ElapsedTime = 0
    End With
StartNextClip_Err:
    Call RegistrarError(Err.Number, Err.Description, "animations.StartNextClip", Erl)
    Resume Next
End Sub
