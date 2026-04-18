Attribute VB_Name = "Unit_StringBuilderExt"
Option Explicit

' ==========================================================================
' StringBuilder Extension Test Suite
' Tests cStringBuilder.cls: Append, AppendNL, Insert, Remove, Find, Clear,
' Length, Capacity, toString, TheString, HeapMinimize, and auto-expansion.
'
' Requirements: 1.1, 1.2, 1.3, 1.4, 1.5, 1.6, 1.7, 1.8, 1.9, 1.10,
'               1.11, 1.12, 1.13, 1.15
' ==========================================================================

#If UNIT_TEST = 1 Then

Public Sub test_suite_stringbuilder_ext()
    ' Example-based tests
    Call UnitTesting.RunTest("sb_init_state", test_init_state())
    Call UnitTesting.RunTest("sb_append_length_and_content", test_append_length_and_content())
    Call UnitTesting.RunTest("sb_multiple_append", test_multiple_append())
    Call UnitTesting.RunTest("sb_append_nl", test_append_nl())
    Call UnitTesting.RunTest("sb_insert_valid", test_insert_valid())
    Call UnitTesting.RunTest("sb_insert_out_of_range", test_insert_out_of_range())
    Call UnitTesting.RunTest("sb_remove_valid", test_remove_valid())
    Call UnitTesting.RunTest("sb_remove_out_of_range", test_remove_out_of_range())
    Call UnitTesting.RunTest("sb_find_existing", test_find_existing())
    Call UnitTesting.RunTest("sb_find_non_existing", test_find_non_existing())
    Call UnitTesting.RunTest("sb_clear", test_clear())
    Call UnitTesting.RunTest("sb_thestring_assignment", test_thestring_assignment())
    Call UnitTesting.RunTest("sb_heap_minimize", test_heap_minimize())
    Call UnitTesting.RunTest("sb_append_beyond_chunksize", test_append_beyond_chunksize())
    
    ' Property-based tests
    Call UnitTesting.RunTest("sb_pbt_append_concat_invariant", test_pbt_append_concat_invariant())
    Call UnitTesting.RunTest("sb_pbt_thestring_roundtrip", test_pbt_thestring_roundtrip())
    Call UnitTesting.RunTest("sb_pbt_heapminimize_preserves", test_pbt_heapminimize_preserves())
End Sub

' Requirement 1.1: New StringBuilder has Length=0 and toString=""
Private Function test_init_state() As Boolean
    On Error GoTo Fail
    Dim sb As New cStringBuilder
    test_init_state = (sb.length = 0) And (sb.ToString = "")
    Exit Function
Fail:
    test_init_state = False
End Function

' Requirement 1.2: Append increments Length and toString contains appended string
Private Function test_append_length_and_content() As Boolean
    On Error GoTo Fail
    Dim sb As New cStringBuilder
    Call sb.Append("Hello")
    test_append_length_and_content = (sb.length = 5) And (sb.ToString = "Hello")
    Exit Function
Fail:
    test_append_length_and_content = False
End Function

' Requirement 1.3: Multiple Append produces concatenation in order
Private Function test_multiple_append() As Boolean
    On Error GoTo Fail
    Dim sb As New cStringBuilder
    Call sb.Append("Hello")
    Call sb.Append(" ")
    Call sb.Append("World")
    test_multiple_append = (sb.ToString = "Hello World") And (sb.length = 11)
    Exit Function
Fail:
    test_multiple_append = False
End Function

' Requirement 1.4: AppendNL appends string followed by vbCrLf
Private Function test_append_nl() As Boolean
    On Error GoTo Fail
    Dim sb As New cStringBuilder
    Call sb.AppendNL("Line1")
    test_append_nl = (sb.ToString = "Line1" & vbCrLf)
    Exit Function
Fail:
    test_append_nl = False
End Function

' Requirement 1.5: Insert at valid index inserts without losing content
Private Function test_insert_valid() As Boolean
    On Error GoTo Fail
    Dim sb As New cStringBuilder
    Call sb.Append("HelloWorld")
    ' Insert at index 5 (0-based character position)
    Call sb.Insert(5, " ")
    test_insert_valid = (sb.ToString = "Hello World") And (sb.length = 11)
    Exit Function
Fail:
    test_insert_valid = False
End Function

' Requirement 1.6: Insert out of range raises error 9
Private Function test_insert_out_of_range() As Boolean
    On Error GoTo ErrExpected
    Dim sb As New cStringBuilder
    Call sb.Append("Hello")
    Call sb.Insert(999, "X")
    ' If we reach here, no error was raised — test fails
    test_insert_out_of_range = False
    Exit Function
ErrExpected:
    test_insert_out_of_range = (Err.Number = 9)
End Function

' Requirement 1.7: Remove valid characters, Length reduced
Private Function test_remove_valid() As Boolean
    On Error GoTo Fail
    Dim sb As New cStringBuilder
    Call sb.Append("Hello World")
    ' Remove 1 character at index 5 (the space)
    Call sb.Remove(5, 1)
    test_remove_valid = (sb.ToString = "HelloWorld") And (sb.length = 10)
    Exit Function
Fail:
    test_remove_valid = False
End Function

' Requirement 1.8: Remove out of range raises error 9
Private Function test_remove_out_of_range() As Boolean
    On Error GoTo ErrExpected
    Dim sb As New cStringBuilder
    Call sb.Append("Hello")
    Call sb.Remove(999, 1)
    ' If we reach here, no error was raised — test fails
    test_remove_out_of_range = False
    Exit Function
ErrExpected:
    test_remove_out_of_range = (Err.Number = 9)
End Function

' Requirement 1.9: Find existing string returns correct 1-based position
Private Function test_find_existing() As Boolean
    On Error GoTo Fail
    Dim sb As New cStringBuilder
    Call sb.Append("Hello World")
    ' "World" starts at position 7 (1-based)
    test_find_existing = (sb.Find("World") = 7)
    Exit Function
Fail:
    test_find_existing = False
End Function

' Requirement 1.10: Find non-existing string returns 0
Private Function test_find_non_existing() As Boolean
    On Error GoTo Fail
    Dim sb As New cStringBuilder
    Call sb.Append("Hello World")
    test_find_non_existing = (sb.Find("XYZ") = 0)
    Exit Function
Fail:
    test_find_non_existing = False
End Function

' Requirement 1.11: Clear resets Length=0 and toString=""
Private Function test_clear() As Boolean
    On Error GoTo Fail
    Dim sb As New cStringBuilder
    Call sb.Append("Some content")
    Call sb.Clear
    test_clear = (sb.length = 0) And (sb.ToString = "")
    Exit Function
Fail:
    test_clear = False
End Function

' Requirement 1.12: TheString assignment — toString returns same string, Length correct
Private Function test_thestring_assignment() As Boolean
    On Error GoTo Fail
    Dim sb As New cStringBuilder
    sb.TheString = "Assigned"
    test_thestring_assignment = (sb.ToString = "Assigned") And (sb.length = 8)
    Exit Function
Fail:
    test_thestring_assignment = False
End Function

' Requirement 1.13: HeapMinimize reduces Capacity, toString unchanged
Private Function test_heap_minimize() As Boolean
    On Error GoTo Fail
    Dim sb As New cStringBuilder
    ' Use a small ChunkSize so we can force multiple expansions
    sb.ChunkSize = 16
    ' Append a large string to force expansion
    Call sb.Append(String$(200, "A"))
    ' Now clear most of the content and set a small string
    Call sb.Clear
    Call sb.Append("Small")
    
    Dim capBefore As Long
    capBefore = sb.Capacity
    
    Call sb.HeapMinimize
    
    Dim capAfter As Long
    capAfter = sb.Capacity
    
    test_heap_minimize = (capAfter <= capBefore) And (sb.ToString = "Small")
    Exit Function
Fail:
    test_heap_minimize = False
End Function

' Requirement 1.15: Append beyond ChunkSize auto-expands, toString complete
Private Function test_append_beyond_chunksize() As Boolean
    On Error GoTo Fail
    Dim sb As New cStringBuilder
    ' Set a small ChunkSize to force expansion
    sb.ChunkSize = 8
    
    ' Append a string much larger than ChunkSize
    Dim bigStr As String
    bigStr = String$(100, "X")
    Call sb.Append(bigStr)
    
    test_append_beyond_chunksize = (sb.ToString = bigStr) And (sb.length = 100)
    Exit Function
Fail:
    test_append_beyond_chunksize = False
End Function

' Feature: unit-test-coverage-tier4, Property 1: Append concatenation invariant
' Validates: Requirements 1.2, 1.3
Private Function test_pbt_append_concat_invariant() As Boolean
    On Error GoTo Fail

    Dim i As Long
    Dim j As Long
    Dim n As Long
    Dim sb As cStringBuilder
    Dim expected As String
    Dim part As String
    Dim totalLen As Long

    For i = 1 To 110
        Set sb = New cStringBuilder
        expected = vbNullString
        totalLen = 0

        ' N strings to append (1..20), cycling
        n = ((i - 1) Mod 20) + 1

        For j = 1 To n
            part = String$(j, Chr$(65 + (j Mod 26)))
            Call sb.Append(part)
            expected = expected & part
            totalLen = totalLen + Len(part)
        Next j

        If sb.ToString <> expected Then
            test_pbt_append_concat_invariant = False
            Exit Function
        End If

        If sb.Length <> totalLen Then
            test_pbt_append_concat_invariant = False
            Exit Function
        End If

        Set sb = Nothing
    Next i

    test_pbt_append_concat_invariant = True
    Exit Function
Fail:
    test_pbt_append_concat_invariant = False
End Function

' Feature: unit-test-coverage-tier4, Property 2: TheString/toString round-trip
' Validates: Requirements 1.12, 1.14
Private Function test_pbt_thestring_roundtrip() As Boolean
    On Error GoTo Fail

    Dim i As Long
    Dim sb As cStringBuilder
    Dim s As String

    For i = 1 To 110
        Set sb = New cStringBuilder
        s = String$(i, Chr$(32 + (i Mod 95)))

        sb.TheString = s

        If sb.ToString <> s Then
            test_pbt_thestring_roundtrip = False
            Exit Function
        End If

        Set sb = Nothing
    Next i

    test_pbt_thestring_roundtrip = True
    Exit Function
Fail:
    test_pbt_thestring_roundtrip = False
End Function

' Feature: unit-test-coverage-tier4, Property 3: HeapMinimize content preservation
' Validates: Requirements 1.13
Private Function test_pbt_heapminimize_preserves() As Boolean
    On Error GoTo Fail

    Dim i As Long
    Dim sb As cStringBuilder
    Dim content As String
    Dim capBefore As Long
    Dim capAfter As Long

    For i = 1 To 110
        Set sb = New cStringBuilder
        sb.ChunkSize = 16

        ' Append strings of increasing length to force expansion
        content = String$(i * 3, Chr$(65 + (i Mod 26)))
        Call sb.Append(content)

        capBefore = sb.Capacity

        Call sb.HeapMinimize

        capAfter = sb.Capacity

        If sb.ToString <> content Then
            test_pbt_heapminimize_preserves = False
            Exit Function
        End If

        If capAfter > capBefore Then
            test_pbt_heapminimize_preserves = False
            Exit Function
        End If

        Set sb = Nothing
    Next i

    test_pbt_heapminimize_preserves = True
    Exit Function
Fail:
    test_pbt_heapminimize_preserves = False
End Function

#End If
