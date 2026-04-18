Attribute VB_Name = "Unit_Instrument"
Option Explicit

' ==========================================================================
' Instrument Test Suite
' Tests clsInstrument.cls: timing via QueryPerformanceCounter
' Verifies ElapsedSeconds >= 0, ElapsedMilliseconds consistency,
' and monotonicity of successive readings.
'
' Requirements: 5.1, 5.2, 5.3
' ==========================================================================

#If UNIT_TEST = 1 Then

Public Sub test_suite_instrument()
    Call UnitTesting.RunTest("inst_elapsed_nonneg", test_elapsed_nonneg())
    Call UnitTesting.RunTest("inst_ms_consistency", test_ms_consistency())
    Call UnitTesting.RunTest("inst_monotonicity", test_monotonicity())
End Sub

' Requirement 5.1: After Start, ElapsedSeconds >= 0
Private Function test_elapsed_nonneg() As Boolean
    On Error GoTo Fail
    Dim inst As New clsInstrument
    Call inst.Start
    test_elapsed_nonneg = (inst.ElapsedSeconds >= 0)
    Exit Function
Fail:
    test_elapsed_nonneg = False
End Function

' Requirement 5.2: ElapsedMilliseconds ? ElapsedSeconds * 1000 (tolerance 1ms)
Private Function test_ms_consistency() As Boolean
    On Error GoTo Fail
    Dim inst As New clsInstrument
    Call inst.Start
    Dim sec As Double
    Dim ms As Double
    sec = inst.ElapsedSeconds
    ms = inst.ElapsedMilliseconds
    ' ms should be close to sec * 1000; allow 1ms tolerance for the
    ' tiny elapsed time between the two calls
    test_ms_consistency = (Abs(ms - sec * 1000) <= 1)
    Exit Function
Fail:
    test_ms_consistency = False
End Function

' Requirement 5.3: Second ElapsedSeconds reading >= first (monotonicity)
Private Function test_monotonicity() As Boolean
    On Error GoTo Fail
    Dim inst As New clsInstrument
    Call inst.Start
    Dim first As Double
    Dim second As Double
    first = inst.ElapsedSeconds
    second = inst.ElapsedSeconds
    test_monotonicity = (second >= first)
    Exit Function
Fail:
    test_monotonicity = False
End Function

#End If
