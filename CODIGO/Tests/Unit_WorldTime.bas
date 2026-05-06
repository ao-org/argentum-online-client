Attribute VB_Name = "Unit_WorldTime"
Option Explicit

' ==========================================================================
' WorldTime Test Suite
' Tests the modWorldTime module: day-cycle initialization, elapsed-time
' normalization, get/set round-trips, second consistency, and range
' invariants for HandleHora and PrepareHora.
'
' NOTE: WorldTime_Ms depends on GetTickCountRaw() internally, so tests
' cannot assert exact ms values. Instead we verify range invariants
' (WorldTime_Ms in [0, dayLenMs-1]), relational invariants
' (WorldTime_Sec = WorldTime_Ms() \ 1000), and get/set round-trips.
' ==========================================================================

#If UNIT_TEST = 1 Then

Public Function test_suite_world_time() As Boolean
    Call UnitTesting.RunTest("wt_init_ms_range", test_wt_init_ms_range())
    Call UnitTesting.RunTest("wt_init_clamp", test_wt_init_clamp())
    Call UnitTesting.RunTest("wt_day_len_get_set", test_wt_day_len_get_set())
    Call UnitTesting.RunTest("wt_sec_consistency", test_wt_sec_consistency())
    Call UnitTesting.RunTest("wt_handle_hora_range", test_wt_handle_hora_range())
    Call UnitTesting.RunTest("wt_prepare_hora_range", test_wt_prepare_hora_range())
    test_suite_world_time = True
End Function

' Uses 60000ms (1 min) as representative day length since WorldTime_Ms
' depends on GetTickCountRaw() internally — we can only verify range invariants.
Private Function test_wt_init_ms_range() As Boolean
    On Error GoTo Fail
    Dim dayLenMs As Long
    dayLenMs = 60000
    Call WorldTime_Init(dayLenMs, 0)
    Dim ms As Long
    ms = WorldTime_Ms()
    test_wt_init_ms_range = (ms >= 0 And ms <= dayLenMs - 1)
    Exit Function
Fail:
    test_wt_init_ms_range = False
End Function

Private Function test_wt_init_clamp() As Boolean
    On Error GoTo Fail
    ' Test with zero
    Call WorldTime_Init(0, 0)
    Dim clamped1 As Long
    clamped1 = WorldTime_DayLenMs()
    ' Test with negative
    Call WorldTime_Init(-500, 0)
    Dim clamped2 As Long
    clamped2 = WorldTime_DayLenMs()
    test_wt_init_clamp = (clamped1 = 1 And clamped2 = 1)
    Exit Function
Fail:
    test_wt_init_clamp = False
End Function

Private Function test_wt_day_len_get_set() As Boolean
    On Error GoTo Fail
    ' Init to a known state first
    Call WorldTime_Init(60000, 0)
    ' Set a positive value and verify round-trip
    Call WorldTime_SetDayLenMs(120000)
    Dim got As Long
    got = WorldTime_DayLenMs()
    Dim positiveOk As Boolean
    positiveOk = (got = 120000)
    ' Set zero - should clamp to 1
    Call WorldTime_SetDayLenMs(0)
    Dim gotZero As Long
    gotZero = WorldTime_DayLenMs()
    Dim zeroOk As Boolean
    zeroOk = (gotZero = 1)
    ' Set negative - should clamp to 1
    Call WorldTime_SetDayLenMs(-100)
    Dim gotNeg As Long
    gotNeg = WorldTime_DayLenMs()
    Dim negOk As Boolean
    negOk = (gotNeg = 1)
    test_wt_day_len_get_set = (positiveOk And zeroOk And negOk)
    Exit Function
Fail:
    test_wt_day_len_get_set = False
End Function

Private Function test_wt_sec_consistency() As Boolean
    On Error GoTo Fail
    Call WorldTime_Init(60000, 0)
    Dim ms As Long
    ms = WorldTime_Ms()
    Dim sec As Long
    sec = WorldTime_Sec()
    test_wt_sec_consistency = (sec = ms \ 1000)
    Exit Function
Fail:
    test_wt_sec_consistency = False
End Function

Private Function test_wt_handle_hora_range() As Boolean
    On Error GoTo Fail
    Dim dayLenMs As Long
    dayLenMs = 60000
    Dim elapsedMs As Long
    elapsedMs = 30000
    Call WorldTime_HandleHora(elapsedMs, dayLenMs)
    Dim ms As Long
    ms = WorldTime_Ms()
    test_wt_handle_hora_range = (ms >= 0 And ms <= dayLenMs - 1)
    Exit Function
Fail:
    test_wt_handle_hora_range = False
End Function

Private Function test_wt_prepare_hora_range() As Boolean
    On Error GoTo Fail
    Dim dayLenMs As Long
    dayLenMs = 60000
    Call WorldTime_Init(dayLenMs, 0)
    Dim outElapsedMs As Long
    Dim outDayLenMs As Long
    Call WorldTime_PrepareHora(outElapsedMs, outDayLenMs)
    Dim rangeOk As Boolean
    rangeOk = (outElapsedMs >= 0 And outElapsedMs <= dayLenMs - 1)
    Dim dayLenOk As Boolean
    dayLenOk = (outDayLenMs = dayLenMs)
    test_wt_prepare_hora_range = (rangeOk And dayLenOk)
    Exit Function
Fail:
    test_wt_prepare_hora_range = False
End Function

#End If
