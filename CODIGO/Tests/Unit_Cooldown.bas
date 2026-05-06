Attribute VB_Name = "Unit_Cooldown"
Option Explicit

' ==========================================================================
' Cooldown System Test Suite
' Tests clsCooldown initialization (totalTime, iconGrh, initialTime) and
' ModCooldown active effect slot tracking (AddOrUpdateEffect, RemoveEffect,
' FindEffectIndex, ResetAllCd).
'
' ==========================================================================

#If UNIT_TEST = 1 Then

Public Sub test_suite_cooldown()
    Call UnitTesting.RunTest("cd_init_stores_total_time", test_cd_init_stores_total_time())
    Call UnitTesting.RunTest("cd_init_stores_icon_grh", test_cd_init_stores_icon_grh())
    Call UnitTesting.RunTest("cd_init_records_initial_time", test_cd_init_records_initial_time())
    Call UnitTesting.RunTest("cd_add_effect_increments_count", test_cd_add_effect_increments_count())
    Call UnitTesting.RunTest("cd_find_effect_returns_index", test_cd_find_effect_returns_index())
    Call UnitTesting.RunTest("cd_remove_effect_decrements_count", test_cd_remove_effect_decrements_count())
    Call UnitTesting.RunTest("cd_reset_all_clears_lists", test_cd_reset_all_clears_lists())
    Call UnitTesting.RunTest("cd_update_effect_overwrites", test_cd_update_effect_overwrites())
    Call UnitTesting.RunTest("cd_pbt_init_stores_values", test_cd_pbt_init_stores_values())
End Sub

' Cooldown_Initialize stores totalTime correctly
Private Function test_cd_init_stores_total_time() As Boolean
    On Error GoTo Fail

    Dim cd As New clsCooldown
    Call cd.Cooldown_Initialize(5000, 100)

    test_cd_init_stores_total_time = (cd.totalTime = 5000)
    Exit Function
Fail:
    test_cd_init_stores_total_time = False
End Function

' Cooldown_Initialize stores iconGrh correctly
Private Function test_cd_init_stores_icon_grh() As Boolean
    On Error GoTo Fail

    Dim cd As New clsCooldown
    Call cd.Cooldown_Initialize(3000, 42)

    test_cd_init_stores_icon_grh = (cd.iconGrh = 42)
    Exit Function
Fail:
    test_cd_init_stores_icon_grh = False
End Function

' Cooldown_Initialize records initialTime as current tick count
Private Function test_cd_init_records_initial_time() As Boolean
    On Error GoTo Fail

    Dim beforeTick As Long
    Dim afterTick As Long
    Dim cd As New clsCooldown

    beforeTick = GetTickCount()
    Call cd.Cooldown_Initialize(2000, 10)
    afterTick = GetTickCount()

    ' initialTime should be between beforeTick and afterTick (inclusive)
    test_cd_init_records_initial_time = (cd.initialTime >= beforeTick) And _
                                        (cd.initialTime <= afterTick)
    Exit Function
Fail:
    test_cd_init_records_initial_time = False
End Function

' AddOrUpdateEffect tracks active cooldown slots
Private Function test_cd_add_effect_increments_count() As Boolean
    On Error GoTo Fail

    ' Save original state
    Dim origCount As Integer
    origCount = CDList.EffectCount

    ' Ensure arrays are initialized
    Call modCooldowns.InitializeEffectArrays

    ' Reset to known state
    CDList.EffectCount = 0

    ' Add an effect
    Dim Effect As t_ActiveEffect
    Effect.TypeId = 1
    Effect.id = 100
    Effect.duration = 5000
    Effect.startTime = GetTickCount()
    Effect.Grh = 50

    Call modCooldowns.AddOrUpdateEffect(CDList, Effect)

    Dim result As Boolean
    result = (CDList.EffectCount = 1)

    ' Restore original state
    CDList.EffectCount = origCount

    test_cd_add_effect_increments_count = result
    Exit Function
Fail:
    ' Restore on error
    CDList.EffectCount = origCount
    test_cd_add_effect_increments_count = False
End Function

' FindEffectIndex returns correct index for tracked effect
Private Function test_cd_find_effect_returns_index() As Boolean
    On Error GoTo Fail

    ' Save original state
    Dim origCount As Integer
    origCount = CDList.EffectCount

    ' Ensure arrays are initialized
    Call modCooldowns.InitializeEffectArrays

    ' Reset to known state
    CDList.EffectCount = 0

    ' Add two effects
    Dim effect1 As t_ActiveEffect
    effect1.TypeId = 1
    effect1.id = 200
    effect1.duration = 3000
    effect1.startTime = GetTickCount()
    effect1.Grh = 10

    Dim effect2 As t_ActiveEffect
    effect2.TypeId = 2
    effect2.id = 300
    effect2.duration = 4000
    effect2.startTime = GetTickCount()
    effect2.Grh = 20

    Call modCooldowns.AddOrUpdateEffect(CDList, effect1)
    Call modCooldowns.AddOrUpdateEffect(CDList, effect2)

    ' Find the second effect
    Dim idx As Integer
    idx = modCooldowns.FindEffectIndex(CDList, effect2)

    Dim result As Boolean
    result = (idx = 1)

    ' Restore original state
    CDList.EffectCount = origCount

    test_cd_find_effect_returns_index = result
    Exit Function
Fail:
    CDList.EffectCount = origCount
    test_cd_find_effect_returns_index = False
End Function

' RemoveEffect decrements count and removes the slot
Private Function test_cd_remove_effect_decrements_count() As Boolean
    On Error GoTo Fail

    ' Save original state
    Dim origCount As Integer
    origCount = CDList.EffectCount

    ' Ensure arrays are initialized
    Call modCooldowns.InitializeEffectArrays

    ' Reset to known state
    CDList.EffectCount = 0

    ' Add two effects
    Dim effect1 As t_ActiveEffect
    effect1.TypeId = 1
    effect1.id = 400
    effect1.duration = 3000
    effect1.startTime = GetTickCount()
    effect1.Grh = 10

    Dim effect2 As t_ActiveEffect
    effect2.TypeId = 2
    effect2.id = 500
    effect2.duration = 4000
    effect2.startTime = GetTickCount()
    effect2.Grh = 20

    Call modCooldowns.AddOrUpdateEffect(CDList, effect1)
    Call modCooldowns.AddOrUpdateEffect(CDList, effect2)

    ' Remove the first effect
    Call modCooldowns.RemoveEffect(CDList, effect1)

    Dim result As Boolean
    result = (CDList.EffectCount = 1)

    ' Verify the remaining effect is effect2
    If result Then
        result = (CDList.EffectList(0).TypeId = 2) And _
                 (CDList.EffectList(0).id = 500)
    End If

    ' Restore original state
    CDList.EffectCount = origCount

    test_cd_remove_effect_decrements_count = result
    Exit Function
Fail:
    CDList.EffectCount = origCount
    test_cd_remove_effect_decrements_count = False
End Function

' ResetAllCd clears all effect lists
Private Function test_cd_reset_all_clears_lists() As Boolean
    On Error GoTo Fail

    ' Save original state
    Dim origBuff As Integer
    Dim origDeBuff As Integer
    Dim origCD As Integer
    origBuff = BuffList.EffectCount
    origDeBuff = DeBuffList.EffectCount
    origCD = CDList.EffectCount

    ' Ensure arrays are initialized
    Call modCooldowns.InitializeEffectArrays

    ' Add an effect to each list
    Dim Effect As t_ActiveEffect
    Effect.TypeId = 1
    Effect.id = 600
    Effect.duration = 5000
    Effect.startTime = GetTickCount()
    Effect.Grh = 30

    Call modCooldowns.AddOrUpdateEffect(BuffList, Effect)

    Effect.id = 601
    Call modCooldowns.AddOrUpdateEffect(DeBuffList, Effect)

    Effect.id = 602
    Call modCooldowns.AddOrUpdateEffect(CDList, Effect)

    ' Reset all
    Call modCooldowns.ResetAllCd

    Dim result As Boolean
    result = (BuffList.EffectCount = 0) And _
             (DeBuffList.EffectCount = 0) And _
             (CDList.EffectCount = 0)

    ' Restore original state
    BuffList.EffectCount = origBuff
    DeBuffList.EffectCount = origDeBuff
    CDList.EffectCount = origCD

    test_cd_reset_all_clears_lists = result
    Exit Function
Fail:
    BuffList.EffectCount = origBuff
    DeBuffList.EffectCount = origDeBuff
    CDList.EffectCount = origCD
    test_cd_reset_all_clears_lists = False
End Function

' AddOrUpdateEffect overwrites existing effect with same TypeId/id
Private Function test_cd_update_effect_overwrites() As Boolean
    On Error GoTo Fail

    ' Save original state
    Dim origCount As Integer
    origCount = CDList.EffectCount

    ' Ensure arrays are initialized
    Call modCooldowns.InitializeEffectArrays

    ' Reset to known state
    CDList.EffectCount = 0

    ' Add an effect
    Dim Effect As t_ActiveEffect
    Effect.TypeId = 5
    Effect.id = 700
    Effect.duration = 3000
    Effect.startTime = GetTickCount()
    Effect.Grh = 40

    Call modCooldowns.AddOrUpdateEffect(CDList, Effect)

    ' Update the same effect with new duration
    Effect.duration = 9000
    Effect.Grh = 99

    Call modCooldowns.AddOrUpdateEffect(CDList, Effect)

    ' Count should still be 1 (updated, not added)
    Dim result As Boolean
    result = (CDList.EffectCount = 1)

    ' Verify the updated values
    If result Then
        result = (CDList.EffectList(0).duration = 9000) And _
                 (CDList.EffectList(0).Grh = 99)
    End If

    ' Restore original state
    CDList.EffectCount = origCount

    test_cd_update_effect_overwrites = result
    Exit Function
Fail:
    CDList.EffectCount = origCount
    test_cd_update_effect_overwrites = False
End Function

' Feature: unit-test-coverage, Property 16: Cooldown initialization stores values
Private Function test_cd_pbt_init_stores_values() As Boolean
    On Error GoTo Fail

    Dim i As Long
    Dim duration As Long
    Dim icon As Long
    Dim cd As clsCooldown

    For i = 1 To 110
        duration = i * 100
        icon = i * 5

        Set cd = New clsCooldown
        Call cd.Cooldown_Initialize(duration, icon)

        If cd.totalTime <> duration Then
            test_cd_pbt_init_stores_values = False
            Exit Function
        End If

        If cd.iconGrh <> icon Then
            test_cd_pbt_init_stores_values = False
            Exit Function
        End If

        Set cd = Nothing
    Next i

    test_cd_pbt_init_stores_values = True
    Exit Function
Fail:
    test_cd_pbt_init_stores_values = False
End Function

#End If
