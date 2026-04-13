Attribute VB_Name = "Unit_ElapsedTime"
Option Explicit

' ==========================================================================
' ElapsedTime Test Suite
' Tests timing utilities: tick elapsed calculation (normal and 32-bit
' wraparound), tick-after comparison, and positive modulo with edge cases
' (zero/negative modulus, negative dividend).
' ==========================================================================

#If UNIT_TEST = 1 Then

' Runs all elapsed-time-related unit tests.
Public Sub test_suite_elapsed_time()
    Call UnitTesting.RunTest("elapsed_normal", test_elapsed_normal())
    Call UnitTesting.RunTest("elapsed_wraparound", test_elapsed_wraparound())
    Call UnitTesting.RunTest("elapsed_tick_after", test_tick_after())
    Call UnitTesting.RunTest("elapsed_posmod_values", test_posmod_values())
    Call UnitTesting.RunTest("elapsed_posmod_zero_mod", test_posmod_zero_mod())
End Sub

' Verifies TicksElapsed returns the correct delta for a normal
' (non-wrapping) tick pair: TicksElapsed(100, 200) = 100.
Private Function test_elapsed_normal() As Boolean
    On Error GoTo Fail
    test_elapsed_normal = (TicksElapsed(100, 200) = 100#)
    Exit Function
Fail:
    test_elapsed_normal = False
End Function

' Verifies TicksElapsed handles 32-bit wraparound correctly when
' currentTick < startTick. Uses startTick=&H7FFFFFFF, currentTick=100.
Private Function test_elapsed_wraparound() As Boolean
    On Error GoTo Fail
    Dim result As Double
    result = TicksElapsed(&H7FFFFFFF, 100)
    ' When currentTick < startTick the function uses:
    ' (TICKS32 - startTick) + currentTick = (4294967296 - 2147483647) + 100 = 2147483749
    test_elapsed_wraparound = (result > 0)
    Exit Function
Fail:
    test_elapsed_wraparound = False
End Function

' Verifies TickAfter returns True when a >= b, and False when a < b.
Private Function test_tick_after() As Boolean
    On Error GoTo Fail
    Dim okAfter As Boolean
    okAfter = (TickAfter(200, 100) = True)
    Dim okEqual As Boolean
    okEqual = (TickAfter(100, 100) = True)
    Dim okBefore As Boolean
    okBefore = (TickAfter(50, 100) = False)
    test_tick_after = (okAfter And okEqual And okBefore)
    Exit Function
Fail:
    test_tick_after = False
End Function

' Verifies PosMod returns correct results for positive, negative,
' and zero dividend values: PosMod(7,5)=2, PosMod(-1,5)=4, PosMod(0,5)=0.
Private Function test_posmod_values() As Boolean
    On Error GoTo Fail
    Dim ok1 As Boolean: ok1 = (PosMod(7, 5) = 2)
    Dim ok2 As Boolean: ok2 = (PosMod(-1, 5) = 4)
    Dim ok3 As Boolean: ok3 = (PosMod(0, 5) = 0)
    test_posmod_values = (ok1 And ok2 And ok3)
    Exit Function
Fail:
    test_posmod_values = False
End Function

' Verifies PosMod returns 0 when the modulus is zero or negative.
Private Function test_posmod_zero_mod() As Boolean
    On Error GoTo Fail
    Dim ok1 As Boolean: ok1 = (PosMod(5, 0) = 0)
    Dim ok2 As Boolean: ok2 = (PosMod(5, -1) = 0)
    test_posmod_zero_mod = (ok1 And ok2)
    Exit Function
Fail:
    test_posmod_zero_mod = False
End Function

#End If
