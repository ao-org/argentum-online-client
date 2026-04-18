Attribute VB_Name = "Unit_Language"
Option Explicit

' ==========================================================================
' Language / FileToString Integration Test Suite
' Tests ModLanguage.bas: FileToString read/write round-trip, empty file
' handling, and validation of Languages/*.json files (1-5) including
' JSON.parse producing valid Dictionary objects.
'
' Requirements: 7.1, 7.2, 7.3, 7.4, 7.5
' ==========================================================================

#If UNIT_TEST = 1 Then

Public Sub test_suite_language()
    ' Example-based tests
    Call UnitTesting.RunTest("lang_fts_roundtrip", test_fts_roundtrip())
    Call UnitTesting.RunTest("lang_fts_empty", test_fts_empty())
    Call UnitTesting.RunTest("lang_json_read_1", test_json_read(1))
    Call UnitTesting.RunTest("lang_json_read_2", test_json_read(2))
    Call UnitTesting.RunTest("lang_json_read_3", test_json_read(3))
    Call UnitTesting.RunTest("lang_json_read_4", test_json_read(4))
    Call UnitTesting.RunTest("lang_json_read_5", test_json_read(5))
    Call UnitTesting.RunTest("lang_json_parse_1", test_json_parse(1))
    Call UnitTesting.RunTest("lang_json_parse_2", test_json_parse(2))
    Call UnitTesting.RunTest("lang_json_parse_3", test_json_parse(3))
    Call UnitTesting.RunTest("lang_json_parse_4", test_json_parse(4))
    Call UnitTesting.RunTest("lang_json_parse_5", test_json_parse(5))
    
    ' Property-based test
    Call UnitTesting.RunTest("lang_pbt_fts_roundtrip", test_pbt_fts_roundtrip())
End Sub

' --------------------------------------------------------------------------
' Example-based tests
' --------------------------------------------------------------------------

' Requirement 7.1: FileToString round-trip with known content
Private Function test_fts_roundtrip() As Boolean
    On Error GoTo Fail
    Dim tmpPath As String
    tmpPath = App.path & "\test_temp_lang.txt"
    
    Dim expected As String
    expected = "Hello World 12345"
    
    ' Write known content to temp file
    Dim f As Integer: f = FreeFile
    Open tmpPath For Output As #f
    Print #f, expected;
    Close #f
    
    ' Read back with FileToString
    Dim result As String
    result = ModLanguage.FileToString(tmpPath)
    
    test_fts_roundtrip = (result = expected)
    
    ' Cleanup
    On Error Resume Next
    Kill tmpPath
    On Error GoTo 0
    Exit Function
Fail:
    On Error Resume Next
    Kill tmpPath
    On Error GoTo 0
    test_fts_roundtrip = False
End Function

' Requirement 7.2: FileToString with empty file returns ""
Private Function test_fts_empty() As Boolean
    On Error GoTo Fail
    Dim tmpPath As String
    tmpPath = App.path & "\test_temp_lang.txt"
    
    ' Create empty file
    Dim f As Integer: f = FreeFile
    Open tmpPath For Output As #f
    Close #f
    
    ' Read back with FileToString
    Dim result As String
    result = ModLanguage.FileToString(tmpPath)
    
    test_fts_empty = (result = "")
    
    ' Cleanup
    On Error Resume Next
    Kill tmpPath
    On Error GoTo 0
    Exit Function
Fail:
    On Error Resume Next
    Kill tmpPath
    On Error GoTo 0
    test_fts_empty = False
End Function

' Requirement 7.3: Each Languages/*.json can be read without error, content not empty
Private Function test_json_read(ByVal langId As Long) As Boolean
    On Error GoTo Fail
    Dim langPath As String
    langPath = App.path & "\Languages\" & langId & ".json"
    
    Dim content As String
    content = ModLanguage.FileToString(langPath)
    
    test_json_read = (Len(content) > 0)
    Exit Function
Fail:
    test_json_read = False
End Function

' Requirement 7.4: Each Languages/*.json parses to valid Dictionary (not Nothing)
Private Function test_json_parse(ByVal langId As Long) As Boolean
    On Error GoTo Fail
    Dim langPath As String
    langPath = App.path & "\Languages\" & langId & ".json"
    
    Dim content As String
    content = ModLanguage.FileToString(langPath)
    
    Dim dict As Object
    Set dict = JSON.parse(content)
    
    test_json_parse = (Not dict Is Nothing)
    Exit Function
Fail:
    test_json_parse = False
End Function

' --------------------------------------------------------------------------
' Property-based tests
' --------------------------------------------------------------------------

' Feature: unit-test-coverage-tier4, Property 10: FileToString write/read round-trip
' Validates: Requirements 7.1
Private Function test_pbt_fts_roundtrip() As Boolean
    On Error GoTo Fail
    Dim tmpPath As String
    tmpPath = App.path & "\test_temp_lang.txt"
    
    Dim i As Long
    Dim expected As String
    Dim result As String
    Dim f As Integer
    
    For i = 1 To 110
        ' Generate non-empty ASCII string of varying length
        expected = String$(((i Mod 50) + 1), Chr$(32 + (i Mod 95)))
        
        ' Write to temp file
        f = FreeFile
        Open tmpPath For Output As #f
        Print #f, expected;
        Close #f
        
        ' Read back with FileToString
        result = ModLanguage.FileToString(tmpPath)
        
        If result <> expected Then
            ' Cleanup on failure
            On Error Resume Next
            Kill tmpPath
            On Error GoTo 0
            test_pbt_fts_roundtrip = False
            Exit Function
        End If
        
        ' Cleanup each iteration
        On Error Resume Next
        Kill tmpPath
        On Error GoTo 0
    Next i
    
    test_pbt_fts_roundtrip = True
    Exit Function
Fail:
    On Error Resume Next
    Kill tmpPath
    On Error GoTo 0
    test_pbt_fts_roundtrip = False
End Function

#End If
