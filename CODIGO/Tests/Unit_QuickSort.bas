Attribute VB_Name = "Unit_QuickSort"
Option Explicit

' ==========================================================================
' QuickSort Test Suite
' Tests the General_Quick_Sort procedure from modCompression.bas:
' already-sorted input, reverse-sorted input, duplicates, single element,
' size invariant (element count preserved), and string sorting.
' ==========================================================================

#If UNIT_TEST = 1 Then

Public Function test_suite_quick_sort() As Boolean
    Call UnitTesting.RunTest("sort_already_sorted", test_sort_already_sorted())
    Call UnitTesting.RunTest("sort_reverse", test_sort_reverse())
    Call UnitTesting.RunTest("sort_duplicates", test_sort_duplicates())
    Call UnitTesting.RunTest("sort_single_element", test_sort_single_element())
    Call UnitTesting.RunTest("sort_size_invariant", test_sort_size_invariant())
    Call UnitTesting.RunTest("sort_strings", test_sort_strings())
    test_suite_quick_sort = True
End Function

Private Function test_sort_already_sorted() As Boolean
    On Error GoTo Fail
    Dim arr As Variant
    arr = Array(1, 2, 3, 4, 5)
    Call General_Quick_Sort(arr, LBound(arr), UBound(arr))
    test_sort_already_sorted = (arr(0) = 1 And arr(1) = 2 And arr(2) = 3 And arr(3) = 4 And arr(4) = 5)
    Exit Function
Fail:
    test_sort_already_sorted = False
End Function

Private Function test_sort_reverse() As Boolean
    On Error GoTo Fail
    Dim arr As Variant
    arr = Array(5, 4, 3, 2, 1)
    Call General_Quick_Sort(arr, LBound(arr), UBound(arr))
    test_sort_reverse = (arr(0) = 1 And arr(1) = 2 And arr(2) = 3 And arr(3) = 4 And arr(4) = 5)
    Exit Function
Fail:
    test_sort_reverse = False
End Function

Private Function test_sort_duplicates() As Boolean
    On Error GoTo Fail
    Dim arr As Variant
    arr = Array(3, 1, 2, 3, 1)
    Call General_Quick_Sort(arr, LBound(arr), UBound(arr))
    test_sort_duplicates = (arr(0) = 1 And arr(1) = 1 And arr(2) = 2 And arr(3) = 3 And arr(4) = 3)
    Exit Function
Fail:
    test_sort_duplicates = False
End Function

Private Function test_sort_single_element() As Boolean
    On Error GoTo Fail
    Dim arr As Variant
    arr = Array(42)
    Call General_Quick_Sort(arr, LBound(arr), UBound(arr))
    test_sort_single_element = (arr(0) = 42)
    Exit Function
Fail:
    test_sort_single_element = False
End Function

Private Function test_sort_size_invariant() As Boolean
    On Error GoTo Fail
    Dim arr As Variant
    arr = Array(9, 3, 7, 1, 5, 8, 2)
    Dim countBefore As Long
    countBefore = UBound(arr) - LBound(arr) + 1
    Call General_Quick_Sort(arr, LBound(arr), UBound(arr))
    Dim countAfter As Long
    countAfter = UBound(arr) - LBound(arr) + 1
    test_sort_size_invariant = (countBefore = countAfter)
    Exit Function
Fail:
    test_sort_size_invariant = False
End Function

Private Function test_sort_strings() As Boolean
    On Error GoTo Fail
    Dim arr As Variant
    arr = Array("cherry", "apple", "banana")
    Call General_Quick_Sort(arr, LBound(arr), UBound(arr))
    test_sort_strings = (arr(0) = "apple" And arr(1) = "banana" And arr(2) = "cherry")
    Exit Function
Fail:
    test_sort_strings = False
End Function

#End If
