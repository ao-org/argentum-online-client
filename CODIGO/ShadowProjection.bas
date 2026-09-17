Attribute VB_Name = "ShadowProjection"
Option Explicit

Public Const SHADOW_MIN_DEPTH As Single = 2!
Public Const SHADOW_MIN_ALPHA As Byte = 48

Public Type ShadowProjectionData
    BaseLeftX As Single
    BaseRightX As Single
    BaseY As Single
    TopLeftX As Single
    TopRightX As Single
    TopY As Single
    SourceBottomV As Single
End Type

Public Function Shadow_Project(ByVal x As Single, ByVal width As Single, ByVal height As Single, ByVal pivotY As Single, ByVal directionX As Single, ByVal directionY As Single) As ShadowProjectionData
    Dim result As ShadowProjectionData
    If width < 1 Then width = 1
    If height < 1 Then height = 1
    If pivotY < 0 Then pivotY = 0
    If pivotY > height - 1 Then pivotY = height - 1
    If directionY < SHADOW_MIN_DEPTH Then directionY = SHADOW_MIN_DEPTH
    result.BaseLeftX = x
    result.BaseRightX = x + width
    result.BaseY = pivotY
    result.TopLeftX = x + directionX
    result.TopRightX = x + width + directionX
    result.TopY = pivotY - directionY
    result.SourceBottomV = (pivotY + 1) / height
    If result.SourceBottomV > 1 Then result.SourceBottomV = 1
    Shadow_Project = result
End Function

Public Function Shadow_SunDirectionX(ByVal hour As Double) As Single
    Dim normalizedHour As Double
    normalizedHour = hour - Fix(hour / 24#) * 24#
    If normalizedHour < 0 Then normalizedHour = normalizedHour + 24#
    Shadow_SunDirectionX = Sin((normalizedHour - 12#) * 3.14159265358979 / 12#) * 48!
End Function

Public Function Shadow_CurrentHour() As Double
    Dim dayLen As Long
    dayLen = WorldTime_DayLenMs()
    If dayLen > 0 Then Shadow_CurrentHour = CDbl(WorldTime_Ms()) / CDbl(dayLen) * 24#
End Function

Public Function Shadow_BodyPivot(ByVal bodyShadowOffsetY As Long, ByVal bodyHeight As Long) As Single
    If bodyShadowOffsetY > 0 Then
        Shadow_BodyPivot = bodyShadowOffsetY
    ElseIf bodyHeight > 0 Then
        Shadow_BodyPivot = bodyHeight - 1
    Else
        Shadow_BodyPivot = -1
    End If
End Function

Public Function Shadow_LocalDirectionX(ByVal lightX As Long, ByVal lightY As Long, ByVal targetX As Long, ByVal targetY As Long, ByVal lightRange As Long) As Single
    If lightRange < 1 Then Exit Function
    If Abs(targetX - lightX) > lightRange Or Abs(targetY - lightY) > lightRange Then Exit Function
    Shadow_LocalDirectionX = ((targetX - lightX) - (targetY - lightY)) * 48!
    If Shadow_LocalDirectionX > 96 Then Shadow_LocalDirectionX = 96
    If Shadow_LocalDirectionX < -96 Then Shadow_LocalDirectionX = -96
End Function

Public Function Shadow_LocalDepth(ByVal lightX As Long, ByVal lightY As Long, ByVal targetX As Long, ByVal targetY As Long, ByVal lightRange As Long) As Single
    If lightRange < 1 Then Exit Function
    If Abs(targetX - lightX) > lightRange Or Abs(targetY - lightY) > lightRange Then Exit Function
    Shadow_LocalDepth = Abs((targetX - lightX) + (targetY - lightY)) * 16!
    If Shadow_LocalDepth < SHADOW_MIN_DEPTH Then Shadow_LocalDepth = SHADOW_MIN_DEPTH
End Function