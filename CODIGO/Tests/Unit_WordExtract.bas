Attribute VB_Name = "Unit_WordExtract"
Option Explicit

' ==========================================================================
' WordExtract Test Suite
' Tests System.bas: HiWord and LoWord functions for extracting the high and
' low 16-bit words from a Long value.
' ==========================================================================

#If UNIT_TEST = 1 Then

Public Sub test_suite_word_extract()
    ' Example-based tests
    Call UnitTesting.RunTest("we_hiword_10002", test_hiword_10002())
    Call UnitTesting.RunTest("we_loword_10002", test_loword_10002())
    Call UnitTesting.RunTest("we_hiword_zero", test_hiword_zero())
    Call UnitTesting.RunTest("we_loword_zero", test_loword_zero())
    
    ' Property-based tests
    Call UnitTesting.RunTest("we_pbt_recombination", test_pbt_recombination())
End Sub

' --------------------------------------------------------------------------
' Example-based tests
' --------------------------------------------------------------------------

' HiWord(&H00010002) = 1
Private Function test_hiword_10002() As Boolean
    On Error GoTo Fail
    test_hiword_10002 = (System.HiWord(&H10002) = 1)
    Exit Function
Fail:
    test_hiword_10002 = False
End Function

' LoWord(&H00010002) = 2
Private Function test_loword_10002() As Boolean
    On Error GoTo Fail
    test_loword_10002 = (System.LoWord(&H10002) = 2)
    Exit Function
Fail:
    test_loword_10002 = False
End Function

' HiWord(0) = 0
Private Function test_hiword_zero() As Boolean
    On Error GoTo Fail
    test_hiword_zero = (System.HiWord(0) = 0)
    Exit Function
Fail:
    test_hiword_zero = False
End Function

' LoWord(0) = 0
Private Function test_loword_zero() As Boolean
    On Error GoTo Fail
    test_loword_zero = (System.LoWord(0) = 0)
    Exit Function
Fail:
    test_loword_zero = False
End Function

' --------------------------------------------------------------------------
' Property-based tests
' --------------------------------------------------------------------------

' Feature: unit-test-coverage-tier4, Property 8: HiWord/LoWord recombination
Private Function test_pbt_recombination() As Boolean
    On Error GoTo Fail
    
    Dim i As Long
    Dim n As Long
    Dim recombined As Long
    
    For i = 1 To 110
        n = i * 19463
        
        ' Guard: only test non-negative Long values (0..&H7FFFFFFF)
        If n < 0 Then n = n And &H7FFFFFFF
        
        recombined = CLng(System.HiWord(n)) * &H10000 + (CLng(System.LoWord(n)) And &HFFFF&)
        
        If recombined <> n Then
            test_pbt_recombination = False
            Exit Function
        End If
    Next i
    
    test_pbt_recombination = True
    Exit Function
Fail:
    test_pbt_recombination = False
End Function

#End If
