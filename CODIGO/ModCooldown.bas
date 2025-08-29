Attribute VB_Name = "modCooldowns"
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

Const CdDrawSize As Integer = 32
Const HalfCDDrawSize As Integer = CdDrawSize / 2
Private colorCooldown As RGBA

Public Sub InitializeEffectArrays()
    On Error Goto InitializeEffectArrays_Err
    ReDim BuffList.EffectList(ACTIVE_EFFECT_LIST_SIZE) As t_ActiveEffect
    ReDim DeBuffList.EffectList(ACTIVE_EFFECT_LIST_SIZE) As t_ActiveEffect
    ReDim CDList.EffectList(ACTIVE_EFFECT_LIST_SIZE) As t_ActiveEffect
    Exit Sub
InitializeEffectArrays_Err:
    Call TraceError(Err.Number, Err.Description, "ModCooldown.InitializeEffectArrays", Erl)
End Sub

Private Sub DrawEffectCd(ByVal x As Integer, ByVal y As Integer, ByRef Effect As t_ActiveEffect, ByVal currTime As Long, ByRef colors() As RGBA)
    On Error Goto DrawEffectCd_Err
    Dim Grh As Grh
    Dim angle As Single
    Call InitGrh(Grh, Effect.Grh)
    Call Grh_Render_Advance(grh, x - HalfCDDrawSize, y - HalfCDDrawSize, CdDrawSize, CdDrawSize, colors)
    If Effect.duration > -1 Then
        angle = (currTime - Effect.startTime) * 360 / Effect.duration
    Else
        angle = 0
    End If
    Call Engine_Draw_Load(x, y, CdDrawSize, CdDrawSize, colorCooldown, angle)
    If Effect.StackCount > 1 Then
        RenderText Effect.StackCount, x - 5, y + HalfCDDrawSize - 12, COLOR_WHITE, 4, False
    End If
    Exit Sub
DrawEffectCd_Err:
    Call TraceError(Err.Number, Err.Description, "ModCooldown.DrawEffectCd", Erl)
End Sub

Public Sub renderCooldowns(ByVal x As Integer, ByVal y As Integer)
    On Error Goto renderCooldowns_Err
    Dim Item As clsCooldown
    Call SetRGBA(colorCooldown, 125, 125, 125, 120)
    Dim i As Integer
    Dim CurrTime As Long
    Dim colors(3) As RGBA
    Call SetRGBA(colors(0), 255, 255, 255, 255)
    Call SetRGBA(colors(1), 255, 255, 255, 255)
    Call SetRGBA(colors(2), 255, 255, 255, 255)
    Call SetRGBA(colors(3), 255, 255, 255, 255)
    CurrTime = GetTickCount()
    Call UpdateEffectTime(BuffList, CurrTime)
    Call UpdateEffectTime(DeBuffList, CurrTime)
    Call UpdateEffectTime(CDList, CurrTime)
    Dim CurrentX As Integer
    CurrentX = x
    Dim Margin As Integer
    Margin = 5
    If BuffList.EffectCount > 0 Then
        For i = 0 To BuffList.EffectCount - 1
            Call DrawEffectCd(CurrentX, y, BuffList.EffectList(i), CurrTime, colors)
            CurrentX = CurrentX - CdDrawSize - Margin
        Next i
        CurrentX = x
        y = y + CdDrawSize + Margin
    End If
    
    If DeBuffList.EffectCount > 0 Then
        For i = 0 To DeBuffList.EffectCount - 1
            Call DrawEffectCd(CurrentX, y, DeBuffList.EffectList(i), CurrTime, colors)
            CurrentX = CurrentX - CdDrawSize
        Next i
        CurrentX = x
        y = y + CdDrawSize + Margin
    End If
    If CDList.EffectCount > 0 Then
        y = Render_Main_Rect.Bottom - CdDrawSize - Margin + gameplay_render_offset.y
        For i = 0 To CDList.EffectCount - 1
            Call DrawEffectCd(CurrentX, y, CDList.EffectList(i), CurrTime, colors)
            CurrentX = CurrentX - CdDrawSize - Margin
        Next i
        CurrentX = x
    End If
    Exit Sub
renderCooldowns_Err:
    Call TraceError(Err.Number, Err.Description, "ModCooldown.renderCooldowns", Erl)
End Sub

Public Sub renderCooldownsInventory(ByVal x As Integer, ByVal y As Integer, ByVal cdProgress As Single)
    On Error Goto renderCooldownsInventory_Err
    Call SetRGBA(colorCooldown, 50, 25, 15, 170)
    x = x + 16
    y = y + 16
    Dim currTime As Long
    Dim colores() As RGBA
    Dim progress As Single
    ReDim colores(3)
    Call SetRGBA(colores(0), 255, 255, 255, 125)
    Call SetRGBA(colores(1), 255, 255, 255, 125)
    Call SetRGBA(colores(2), 255, 255, 255, 125)
    Call SetRGBA(colores(3), 255, 255, 255, 125)
    
    If cdProgress >= 1 Then
       Set cooldown_ataque = Nothing
    Else
        Call Engine_Draw_Load(x, y, 32, 32, colorCooldown, 360 * cdProgress)
        x = x - 36
        i = i + 1
    End If

    Exit Sub
renderCooldownsInventory_Err:
    Call TraceError(Err.Number, Err.Description, "ModCooldown.renderCooldownsInventory", Erl)
End Sub

Public Sub AddOrUpdateEffect(ByRef EffectList As t_ActiveEffectList, ByRef Effect As t_ActiveEffect)
    On Error Goto AddOrUpdateEffect_Err
On Error GoTo AddEffect_Err
    Dim Index As Integer
    Index = FindEffectIndex(EffectList, Effect)
    If Index > -1 Then
        EffectList.EffectList(Index) = Effect
        Exit Sub
    End If
100 If Not IsArrayInitialized(EffectList.EffectList) Then
104     ReDim EffectList.EffectList(ACTIVE_EFFECT_LIST_SIZE) As t_ActiveEffect
    ElseIf EffectList.EffectCount >= UBound(EffectList.EffectList) Then
108     ReDim Preserve EffectList.EffectList(EffectList.EffectCount * 1.2) As t_ActiveEffect
    End If
116 EffectList.EffectList(EffectList.EffectCount) = Effect
120 EffectList.EffectCount = EffectList.EffectCount + 1
    Exit Sub
AddEffect_Err:
      Call RegistrarError(Err.Number, Err.Description, "modCooldowns.AddEffect", Erl)
    Exit Sub
AddOrUpdateEffect_Err:
    Call TraceError(Err.Number, Err.Description, "ModCooldown.AddOrUpdateEffect", Erl)
End Sub

Public Function FindEffectIndex(ByRef EffectList As t_ActiveEffectList, ByRef Effect As t_ActiveEffect) As Integer
    On Error Goto FindEffectIndex_Err
    FindEffectIndex = -1
    Dim i As Integer
100 For i = 0 To EffectList.EffectCount - 1
106     If EffectList.EffectList(i).TypeId = Effect.TypeId And _
            EffectList.EffectList(i).Id = Effect.Id Then
110         FindEffectIndex = i
            Exit Function
        End If
    Next i
    Exit Function
FindEffectIndex_Err:
    Call TraceError(Err.Number, Err.Description, "ModCooldown.FindEffectIndex", Erl)
End Function

Public Sub RemoveEffect(ByRef EffectList As t_ActiveEffectList, ByRef Effect As t_ActiveEffect)
    On Error Goto RemoveEffect_Err
On Error GoTo RemoveEffect_Err
    Dim Index As Integer
    Index = FindEffectIndex(EffectList, Effect)
    If Index > -1 Then
        EffectList.EffectList(Index) = EffectList.EffectList(EffectList.EffectCount - 1)
        EffectList.EffectCount = EffectList.EffectCount - 1
    End If
    Exit Sub
RemoveEffect_Err:
      Call RegistrarError(Err.Number, Err.Description, "modCooldowns.RemoveEffect", Erl)
    Exit Sub
RemoveEffect_Err:
    Call TraceError(Err.Number, Err.Description, "ModCooldown.RemoveEffect", Erl)
End Sub

Public Sub UpdateEffectTime(ByRef EffectList As t_ActiveEffectList, ByVal CurrTime As Long)
    On Error Goto UpdateEffectTime_Err
    Dim i As Integer
    Dim PendingTime As Long
    i = 0
    While i < EffectList.EffectCount
        PendingTime = EffectList.EffectList(i).Duration - (CurrTime - EffectList.EffectList(i).StartTime)
        If EffectList.EffectList(i).Duration > -1 And PendingTime < 1 Then
            Call RemoveEffect(EffectList, EffectList.EffectList(i))
        Else
            i = i + 1
        End If
    Wend
    Exit Sub
UpdateEffectTime_Err:
    Call TraceError(Err.Number, Err.Description, "ModCooldown.UpdateEffectTime", Erl)
End Sub

Public Sub ResetAllCd()
    On Error Goto ResetAllCd_Err
    BuffList.EffectCount = 0
    DeBuffList.EffectCount = 0
    CDList.EffectCount = 0
    Exit Sub
ResetAllCd_Err:
    Call TraceError(Err.Number, Err.Description, "ModCooldown.ResetAllCd", Erl)
End Sub

