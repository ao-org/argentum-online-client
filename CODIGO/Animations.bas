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
    animationState.LastFrameTime = GetTickCount()
    animationState.ElapsedTime = animationState.ElapsedTime + DeltaTime
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
    Exit Sub
UpdateAnimation_Err:
    Call RegistrarError(Err.Number, Err.Description, "animations.UpdateAnimation", Erl)
    Resume Next
End Sub

Public Sub StartAnimation(ByRef animationState As tAnimationPlaybackState, ByVal composedAnimationIndex As Long)
On Error GoTo StartAnimation_Err
    animationState.PlaybackState = Forward
    animationState.ComposedAnimation = composedAnimationIndex
    animationState.ActiveClip = 1
    animationState.CurrentGrh = FxData(ComposedFxData(composedAnimationIndex).Clips(animationState.ActiveClip).fX).Animacion
    animationState.CurrentFrame = 1
    animationState.CurrentClipLoops = 0
    animationState.ElapsedTime = 0
    animationState.LastFrameTime = GetTickCount()
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
        LastFrameTime = GetTickCount()
        ElapsedTime = 0
    End With
    Exit Sub
ChangeToClip_Err:
    Call RegistrarError(Err.Number, Err.Description, "animations.StartNextClip", Erl)
    Resume Next
End Sub
