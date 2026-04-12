Attribute VB_Name = "Unit_IniManager"
Option Explicit

' ==========================================================================
' IniManager Test Suite
' Tests the clsIniManager class: Initialize, GetValue, ChangeValue,
' KeyExists, NodesCount, EntriesCount, and DumpFile round-trip.
' ==========================================================================

#If UNIT_TEST = 1 Then

Private Const TEMP_INI_PATH As String = "test_temp.ini"
Private Const TEMP_DUMP_PATH As String = "test_temp_dump.ini"

' Writes a temporary INI file with known sections and keys for testing.
Private Sub write_temp_ini(ByVal path As String)
    Dim hFile As Integer
    hFile = FreeFile
    Open path For Output As hFile
    Print #hFile, "[General]"
    Print #hFile, "Name=TestApp"
    Print #hFile, "Version=1.0"
    Print #hFile, "Debug=0"
    Print #hFile, ""
    Print #hFile, "[Network]"
    Print #hFile, "Host=127.0.0.1"
    Print #hFile, "Port=7666"
    Print #hFile, ""
    Print #hFile, "[Display]"
    Print #hFile, "Width=800"
    Print #hFile, "Height=600"
    Print #hFile, ""
    Close hFile
End Sub

' Cleans up temporary files created during testing.
Private Sub delete_temp_files()
    On Error Resume Next
    Kill App.path & "\" & TEMP_INI_PATH
    Kill App.path & "\" & TEMP_DUMP_PATH
    On Error GoTo 0
End Sub

Public Sub test_suite_ini_manager()
    Call write_temp_ini(App.path & "\" & TEMP_INI_PATH)
    
    Call UnitTesting.RunTest("ini_get_value", test_ini_get_value())
    Call UnitTesting.RunTest("ini_key_exists", test_ini_key_exists())
    Call UnitTesting.RunTest("ini_change_existing", test_ini_change_existing())
    Call UnitTesting.RunTest("ini_change_new_section", test_ini_change_new_section())
    Call UnitTesting.RunTest("ini_nodes_count", test_ini_nodes_count())
    Call UnitTesting.RunTest("ini_entries_count", test_ini_entries_count())
    Call UnitTesting.RunTest("ini_dump_round_trip", test_ini_dump_round_trip())
    
    Call delete_temp_files
End Sub

Private Function test_ini_get_value() As Boolean
    On Error GoTo Fail
    Dim mgr As New clsIniManager
    Call mgr.Initialize(App.path & "\" & TEMP_INI_PATH)
    
    test_ini_get_value = (mgr.GetValue("General", "Name") = "TestApp") And _
                         (mgr.GetValue("Network", "Port") = "7666")
    Exit Function
Fail:
    test_ini_get_value = False
End Function

Private Function test_ini_key_exists() As Boolean
    On Error GoTo Fail
    Dim mgr As New clsIniManager
    Call mgr.Initialize(App.path & "\" & TEMP_INI_PATH)
    
    test_ini_key_exists = (mgr.KeyExists("General") = True) And _
                          (mgr.KeyExists("NonExistent") = False)
    Exit Function
Fail:
    test_ini_key_exists = False
End Function

Private Function test_ini_change_existing() As Boolean
    On Error GoTo Fail
    Dim mgr As New clsIniManager
    Call mgr.Initialize(App.path & "\" & TEMP_INI_PATH)
    
    Call mgr.ChangeValue("General", "Name", "UpdatedApp")
    
    test_ini_change_existing = (mgr.GetValue("General", "Name") = "UpdatedApp")
    Exit Function
Fail:
    test_ini_change_existing = False
End Function

Private Function test_ini_change_new_section() As Boolean
    On Error GoTo Fail
    Dim mgr As New clsIniManager
    Call mgr.Initialize(App.path & "\" & TEMP_INI_PATH)
    
    Call mgr.ChangeValue("NewSection", "NewKey", "NewValue")
    
    test_ini_change_new_section = (mgr.GetValue("NewSection", "NewKey") = "NewValue")
    Exit Function
Fail:
    test_ini_change_new_section = False
End Function

Private Function test_ini_nodes_count() As Boolean
    On Error GoTo Fail
    Dim mgr As New clsIniManager
    Call mgr.Initialize(App.path & "\" & TEMP_INI_PATH)
    
    ' The temp INI has 3 sections: General, Network, Display
    test_ini_nodes_count = (mgr.NodesCount = 3)
    Exit Function
Fail:
    test_ini_nodes_count = False
End Function

Private Function test_ini_entries_count() As Boolean
    On Error GoTo Fail
    Dim mgr As New clsIniManager
    Call mgr.Initialize(App.path & "\" & TEMP_INI_PATH)
    
    ' General has 3 entries (Name, Version, Debug), Network has 2 (Host, Port)
    test_ini_entries_count = (mgr.EntriesCount("General") = 3) And _
                             (mgr.EntriesCount("Network") = 2)
    Exit Function
Fail:
    test_ini_entries_count = False
End Function

Private Function test_ini_dump_round_trip() As Boolean
    On Error GoTo Fail
    Dim mgr As New clsIniManager
    Call mgr.Initialize(App.path & "\" & TEMP_INI_PATH)
    
    ' Dump to a second temp file
    Dim dumpPath As String
    dumpPath = App.path & "\" & TEMP_DUMP_PATH
    Call mgr.DumpFile(dumpPath)
    
    ' Re-load from the dumped file
    Dim mgr2 As New clsIniManager
    Call mgr2.Initialize(dumpPath)
    
    ' Verify values match across both instances
    test_ini_dump_round_trip = (mgr2.GetValue("General", "Name") = "TestApp") And _
                               (mgr2.GetValue("Network", "Port") = "7666") And _
                               (mgr2.GetValue("Display", "Width") = "800") And _
                               (mgr2.NodesCount = 3)
    Exit Function
Fail:
    test_ini_dump_round_trip = False
End Function

#End If
