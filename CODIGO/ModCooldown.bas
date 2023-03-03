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

Private colorCooldown As RGBA

Private Sub DrawEffectCd(ByVal x As Integer, ByVal y As Integer, ByRef Effect As e_ActiveEffect, ByVal CurrTime As Long, ByRef colors() As RGBA)
    Dim Grh As Grh
    Dim angle As Single
    Call InitGrh(Grh, Effect.Grh)
    Call Grh_Render_Advance(Grh, x - 16, y - 16, 32, 32, colors)
    If Effect.duration > -1 Then
        angle = (currTime - Effect.startTime) * 360 / Effect.duration
    Else
        angle = 0
    End If
    Call Engine_Draw_Load(x, y, 32, 32, colorCooldown, angle)
End Sub

Public Sub renderCooldowns(ByVal x As Integer, ByVal y As Integer)
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
            CurrentX = CurrentX - 32 - Margin
        Next i
        CurrentX = x
        y = y + 32 + Margin
    End If
    
    If DeBuffList.EffectCount > 0 Then
        For i = 0 To DeBuffList.EffectCount - 1
            Call DrawEffectCd(CurrentX, y, DeBuffList.EffectList(i), CurrTime, colors)
            CurrentX = CurrentX - 32
        Next i
        CurrentX = x
        y = y + 32 + Margin
    End If
    If CDList.EffectCount > 0 Then
        y = Render_Main_Rect.Bottom - 32 - Margin
        For i = 0 To CDList.EffectCount - 1
            Call DrawEffectCd(CurrentX, y, CDList.EffectList(i), CurrTime, colors)
            CurrentX = CurrentX - 32 - Margin
        Next i
        CurrentX = x
    End If
End Sub

Public Sub renderCooldownsInventory(ByVal x As Integer, ByVal y As Integer, ByVal cdProgress As Single)
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

End Sub

Public Sub AddOrUpdateEffect(ByRef EffectList As t_ActiveEffectList, ByRef Effect As e_ActiveEffect)
On Error GoTo AddEffect_Err
    Dim Index As Integer
    Index = FindEffectIndex(EffectList, Effect)
    If Index > -1 Then
        EffectList.EffectList(Index) = Effect
        Exit Sub
    End If
100 If Not IsArrayInitialized(EffectList.EffectList) Then
104     ReDim EffectList.EffectList(ACTIVE_EFFECT_LIST_SIZE) As e_ActiveEffect
    ElseIf EffectList.EffectCount >= UBound(EffectList.EffectList) Then
108     ReDim Preserve EffectList.EffectList(EffectList.EffectCount * 1.2) As e_ActiveEffect
    End If
116 EffectList.EffectList(EffectList.EffectCount) = Effect
120 EffectList.EffectCount = EffectList.EffectCount + 1
    Exit Sub
AddEffect_Err:
      Call RegistrarError(Err.Number, Err.Description, "modCooldowns.AddEffect", Erl)
End Sub

Public Function FindEffectIndex(ByRef EffectList As t_ActiveEffectList, ByRef Effect As e_ActiveEffect) As Integer
    FindEffectIndex = -1
    Dim i As Integer
100 For i = 0 To EffectList.EffectCount - 1
106     If EffectList.EffectList(i).TypeId = Effect.TypeId And _
            EffectList.EffectList(i).Id = Effect.Id Then
110         FindEffectIndex = i
            Exit Function
        End If
    Next i
End Function

Public Sub RemoveEffect(ByRef EffectList As t_ActiveEffectList, ByRef Effect As e_ActiveEffect)
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
End Sub

Public Sub UpdateEffectTime(ByRef EffectList As t_ActiveEffectList, ByVal CurrTime As Long)
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
End Sub

Public Sub ResetAllCd()
    BuffList.EffectCount = 0
    DeBuffList.EffectCount = 0
    CDList.EffectCount = 0
End Sub

