Attribute VB_Name = "Group"
Option Explicit

Const StartY = 250
Const StartX = 10
Const SpacingY = 60
Const GroupBackgroundGrh = 29335
Const HpBarStartX = 34
Const HpBarEndX = 99
Const HpBarStartY = 21
Const TextStartX = 33
Const TextEndX = 100
Const TextStartY = 5
Const HeadOffsetX = 0
Const HeadOffsetY = -5
Const FrameWidth = 104
Const FrameHeight = 32
Const AnimationSpeed = 0.03

Const HideShowRectWidth = 10
Const HideShowRectHeight = 40
Public Const HideArrowGrh = 29548
Const ShowArrowGrh = 29549

Public Type t_GroupEntry
    CharIndex As Integer
    Name As String
    GroupId As Integer
    MinHp As Integer
    MaxHp As Integer
    Shield As Long
    RenderArea As RECT
    Head As HeadData
End Type

Public GroupSize As Byte
Public GroupMembers() As t_GroupEntry
Public HideShowRect As RECT
Public Hide As Boolean
Public CurrentPivot As Single
Public LastFrameTime As Long
Public AnimationActive As Boolean
Public ActiveArrowGrh As Long
Public Sub Clear()
    GroupSize = 0
End Sub

Public Sub UpdateRenderArea()
    Dim i  As Integer
    Dim DrawCount As Integer
    Dim RenderStartY As Integer
    RenderStartY = StartY - (SpacingY * (GroupSize - 1)) / 2
    For i = 0 To GroupSize - 1
        If GroupMembers(i).CharIndex <> UserCharIndex Then
            GroupMembers(i).RenderArea.Left = StartX
            GroupMembers(i).RenderArea.Top = RenderStartY + SpacingY * DrawCount
            GroupMembers(i).RenderArea.Right = GroupMembers(i).RenderArea.Left + FrameWidth
            GroupMembers(i).RenderArea.Bottom = GroupMembers(i).RenderArea.Top + FrameHeight
            DrawCount = DrawCount + 1
        End If
    Next i
    Engine_Draw_Box StartX - HideShowRectWidth / 2, StartY - HideShowRectHeight / 2, _
                    HideShowRectWidth, HideShowRectHeight, RGBA_From_Comp(128, 128, 128, 255)
    HideShowRect.Top = StartY - HideShowRectHeight / 2
    HideShowRect.Left = StartX - HideShowRectWidth / 2 - 2
    HideShowRect.Right = HideShowRect.Left + HideShowRectWidth
    HideShowRect.Bottom = HideShowRect.Top + HideShowRectHeight
    CurrentPivot = 0
    ActiveArrowGrh = IIf(Hide, HideArrowGrh, ShowArrowGrh)
End Sub

Public Sub RenderGroup()
    Dim i  As Integer
    Dim temp_array(3) As RGBA
    Dim HpBarSize As Single
    Dim ShieldBardSize As Single
    Call RGBAList(temp_array, 255, 255, 255, 50)
    If AnimationActive Then
        Dim CurrTime As Long
        CurrTime = GetTickCount()
        CurrentPivot = CurrentPivot + IIf(Hide, AnimationSpeed, -AnimationSpeed) * (CurrTime - LastFrameTime)
        If Hide Then
            If CurrentPivot >= FrameWidth + StartX + gameplay_render_offset.x Then
                CurrentPivot = FrameWidth + StartX + gameplay_render_offset.x
                AnimationActive = False
            End If
        Else
            If CurrentPivot <= 0 Then
                CurrentPivot = 0
                AnimationActive = False
            End If
        End If
    End If
    If GroupSize < 1 Then Exit Sub
    For i = 0 To GroupSize - 1
        With GroupMembers(i)
            If .CharIndex <> UserCharIndex Then
                Call Draw_GrhIndex(GroupBackgroundGrh, .RenderArea.Left - CurrentPivot + gameplay_render_offset.x, .RenderArea.Top + gameplay_render_offset.y)
                Call Draw_Grh(.Head.Head(E_Heading.south), .RenderArea.Left - HeadOffsetX - CurrentPivot + gameplay_render_offset.x, .RenderArea.Top + HeadOffsetY + gameplay_render_offset.y, 1, 0, COLOR_WHITE, False, 0, 0)
                Dim HpPlusShield As Long
                HpPlusShield = .MaxHp + .Shield
                HpBarSize = .MinHp / HpPlusShield
                HpBarSize = HpBarSize * (HpBarEndX - HpBarStartX)
                Engine_Draw_Box .RenderArea.Left + HpBarStartX - CurrentPivot + gameplay_render_offset.x, .RenderArea.Top + HpBarStartY + gameplay_render_offset.y, HpBarSize, 3, RGBA_From_Comp(178, 0, 0, 160)
                If .Shield > 0 Then
                    ShieldBardSize = .Shield / HpPlusShield
                    ShieldBardSize = ShieldBardSize * (HpBarEndX - HpBarStartX)
                    Engine_Draw_Box .RenderArea.Left + HpBarStartX - CurrentPivot + HpBarSize + gameplay_render_offset.x, .RenderArea.Top + HpBarStartY + gameplay_render_offset.y, ShieldBardSize, 3, RGBA_From_Comp(162, 108, 16, 160)
                End If
                Call Engine_Text_Render(.Name, .RenderArea.Left + TextStartX - CurrentPivot + gameplay_render_offset.x, .RenderArea.Top + TextStartY + gameplay_render_offset.y, temp_array, 1, True, 0, 128)
            End If
        End With
    Next i
    Call Draw_GrhIndex(ActiveArrowGrh, HideShowRect.Left + gameplay_render_offset.x, HideShowRect.Top + gameplay_render_offset.y)
End Sub

Public Function HandleMouseInput(ByVal x As Integer, ByVal y As Integer) As Boolean
    Dim i As Integer
    If GroupSize < 1 Then Exit Function
    For i = 0 To GroupSize - 1
        If PointIsInsideRect(x, y, GroupMembers(i).RenderArea) Then
            If UsingSkill = magia Then
                HandleMouseInput = True
                If MainTimer.Check(TimersIndex.CastSpell) Then
                    Call WriteActionOnGroupFrame(GroupMembers(i).GroupId)
                    UsaLanzar = False
                    UsingSkill = 0
                    If CursoresGraficos = 0 Then
                        frmMain.MousePointer = vbDefault
                    End If
                End If
            End If
            Exit Function
        End If
    Next i
    If PointIsInsideRect(x, y, HideShowRect) Then
        AnimationActive = True
        LastFrameTime = GetTickCount()
        Hide = Not Hide
        ActiveArrowGrh = IIf(Hide, HideArrowGrh, ShowArrowGrh)
    End If
End Function

