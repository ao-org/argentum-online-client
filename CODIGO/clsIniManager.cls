VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIniManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Argentum 20 Game Client
'
'    Copyright (C) 2023 Noland Studios LTD
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
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 Márquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'
Option Explicit
Option Base 0

''
'Structure that contains a value and it's key in a INI file
'
' @param    key String containing the key associated to the value.
' @param    value String containing the value of the INI entry.
' @see      MainNode
'
Private Type ChildNode
    key As String
    value As String
End Type

''
'Structure that contains all info under a tag in a INI file.
'Such tags are indicated with the "[" and "]" characters.
'
' @param    name String containing the text within the "[" and "]" characters.
'It's the key used when searching for a main section of the INI data.
' @param    values Array of ChildNodes, each containing a value entry along with it's key.
' @param    numValues Number of entrys in the main node.
Private Type MainNode
    Name As String
    values() As ChildNode
    numValues As Integer
End Type

''
'Containts all Main sections of the loaded INI file
Private fileData() As MainNode
''
'Stores the total number of main sections in the loaded INI file
Private MainNodes  As Long

''
'Default constructor. Does nothing.
Private Sub Class_Initialize()
End Sub

''
'Destroy every array and deallocates al memory.
'
Private Sub Class_Terminate()
    On Error GoTo Class_Terminate_Err
    Dim i As Long
    'Clean up
    If MainNodes Then
        For i = 1 To MainNodes - 1
            Erase fileData(i).values
        Next i
        Erase fileData
    End If
    MainNodes = 0
    Exit Sub
Class_Terminate_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsIniManager.Class_Terminate", Erl)
    Resume Next
End Sub

''
'Loads a INI file so it's values can be read. Must be called before being able to use GetValue.
'
' @param    file Complete path of the INI file to be loaded.
' @see      GetValue
Public Sub Initialize(ByVal File As String)
    On Error GoTo Initialize_Err
    'Opens the requested file and loads it's data into memory
    Dim Handle As Integer
    Dim text   As String
    Dim Pos    As Long
    'Prevent memory losses if we are attempting to reload a file....
    Call Class_Terminate
    'Get a free handle and start reading line by line until the end
    Handle = FreeFile
    Open File For Input As Handle
    Do Until EOF(Handle)
        Line Input #Handle, text
        'Is it null??
        If Len(text) Then
            'If it starts with '[' it is a main node or nothing (GetPrivateProfileStringA works this way), otherwise it's a value
            If Left$(text, 1) = "[" Then
                'If it has an ending ']' it's a main node, otherwise it's nothing
                Pos = InStr(2, text, "]")
                If Pos Then
                    'Add a main node
                    ReDim Preserve fileData(MainNodes) As MainNode
                    fileData(MainNodes).Name = UCase$(Trim$(mid$(text, 2, Pos - 2)))
                    MainNodes = MainNodes + 1
                End If
            Else
                'So it's a value. Check if it has a '=', otherwise it's nothing
                Pos = InStr(2, text, "=")
                If Pos Then
                    'Is it under any main node??
                    If MainNodes Then
                        With fileData(MainNodes - 1)
                            'Add it to the main node's value
                            ReDim Preserve .values(.numValues) As ChildNode
                            .values(.numValues).value = Right$(text, Len(text) - Pos)
                            .values(.numValues).key = UCase$(Left$(text, Pos - 1))
                            .numValues = .numValues + 1
                        End With
                    End If
                End If
            End If
        End If
    Loop
    Close Handle
    If MainNodes Then
        'Sort main nodes to allow binary search
        Call SortMainNodes(0, MainNodes - 1)
        'Sort values of each node to allow binary search
        Dim i As Long
        For i = 0 To MainNodes - 1
            If fileData(i).numValues Then Call SortChildNodes(fileData(i), 0, fileData(i).numValues - 1)
        Next i
    End If
    Exit Sub
Initialize_Err:
    Close Handle
    Call RegistrarError(Err.Number, Err.Description & vbNewLine & "Archivo: " & File, "clsIniManager.Initialize", Erl)
    Resume Next
End Sub

''
'Sorts all child nodes within the given MainNode alphabetically by their keys. Uses quicksort.
'
' @param    Node The MainNode whose values are to be sorted.
' @param    first The first index to consider when sorting.
' @param    last The last index to be considered when sorting.
Private Sub SortChildNodes(ByRef Node As MainNode, ByVal First As Integer, ByVal Last As Integer)
    On Error GoTo SortChildNodes_Err
    'Sorts the list of values in a given MainNode using quicksort,
    'this allows the use of Binary Search for faster searches
    Dim min  As Integer      'First item in the list
    Dim max  As Integer      'Last item in the list
    Dim comp As String      'Item used to compare
    Dim temp As ChildNode
    min = First
    max = Last
    With Node
        comp = .values((min + max) \ 2).key
        Do While min <= max
            Do While .values(min).key < comp And min < Last
                min = min + 1
            Loop
            Do While .values(max).key > comp And max > First
                max = max - 1
            Loop
            If min <= max Then
                temp = .values(min)
                .values(min) = .values(max)
                .values(max) = temp
                min = min + 1
                max = max - 1
            End If
        Loop
    End With
    If First < max Then SortChildNodes Node, First, max
    If min < Last Then SortChildNodes Node, min, Last
    Exit Sub
SortChildNodes_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsIniManager.SortChildNodes", Erl)
    Resume Next
End Sub

''
'Sorts all main nodes in the loaded INI file alphabetically by their names. Uses quicksort.
'
' @param    first The first index to consider when sorting.
' @param    last The last index to be considered when sorting.
Private Sub SortMainNodes(ByVal First As Integer, ByVal Last As Integer)
    On Error GoTo SortMainNodes_Err
    'Sorts the MainNodes list using quicksort,
    'this allows the use of Binary Search for faster searches
    Dim min  As Integer      'First item in the list
    Dim max  As Integer      'Last item in the list
    Dim comp As String      'Item used to compare
    Dim temp As MainNode
    min = First
    max = Last
    comp = fileData((min + max) \ 2).Name
    Do While min <= max
        Do While fileData(min).Name < comp And min < Last
            min = min + 1
        Loop
        Do While fileData(max).Name > comp And max > First
            max = max - 1
        Loop
        If min <= max Then
            temp = fileData(min)
            fileData(min) = fileData(max)
            fileData(max) = temp
            min = min + 1
            max = max - 1
        End If
    Loop
    If First < max Then SortMainNodes First, max
    If min < Last Then SortMainNodes min, Last
    Exit Sub
SortMainNodes_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsIniManager.SortMainNodes", Erl)
    Resume Next
End Sub

''
'Searches for a given key within a given main section and if it exists retrieves it's value, otherwise a null string
'
' @param    Main The name of the main section in which we will be searching.
' @param    key The key of the value we are looking for.
' @returns  The value asociated with the given key under the requeted main section of the INI file or a null string if it's not found.
Public Function GetValue(ByVal Main As String, ByVal key As String)
    On Error GoTo GetValue_Err
    'Returns a value if the key and main node exist, or a nullstring otherwise
    Dim i      As Long
    Dim J      As Long
    Dim retval As String
    'Search for the main node
    i = FindMain(UCase$(Main))
    If i >= 0 Then
        'If valid, binary search among keys
        J = FindKey(fileData(i), UCase$(key))
        'If we found it we return it
        If J >= 0 Then retval = fileData(i).values(J).value
        GetValue = retval
    End If
    Exit Function
GetValue_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsIniManager.GetValue", Erl)
    Resume Next
End Function

''
'Searches for a given key within a given main section and if it exists retrieves it's value, otherwise a null string
'
' @param    Main The name of the main section in which we will be searching.
' @param    key The key of the value we are looking for.
' @returns  The value asociated with the given key under the requeted main section of the INI file or a null string if it's not found.
Public Sub ChangeValue(ByVal Main As String, ByVal key As String, ByVal value As String)
    On Error GoTo ChangeValue_Err
    'If the key and main node exist, changes the value
    Dim i As Long
    Dim J As Long
    'Search for the main node
    i = FindMain(UCase$(Main))
    If i < 0 Then i = AddNode(Main)
    'If valid, binary search among keys
    J = FindKey(fileData(i), UCase$(key))
    If J < 0 Then
        Call AddKey(Main, key, value)
    Else
        'If we found it we change it
        fileData(i).values(J).value = value
    End If
    Exit Sub
ChangeValue_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsIniManager.ChangeValue", Erl)
    Resume Next
End Sub

''
'Searches for a given key within a given main node and returns the index in which it's stored or the negation of the index in which it should be if not found.
'
' @param    Node The MainNode among whose value entries we will be searching.
' @param    key The key of the value we are looking for.
' @returns  The index in which the value with the key we are looking for is stored or the negation of the index in which it should be if not found.
Private Function FindKey(ByRef Node As MainNode, ByVal key As String) As Long
    On Error GoTo FindKey_Err
    'Returns the index of the value which key matches the requested one,
    'or the negation of the position were it should be if not found
    Dim min As Long
    Dim max As Long
    Dim mid As Long
    min = 0
    max = Node.numValues - 1
    Do While min <= max
        mid = (min + max) \ 2
        If Node.values(mid).key < key Then
            min = mid + 1
        ElseIf Node.values(mid).key > key Then
            max = mid - 1
        Else
            'We found it
            FindKey = mid
            Exit Function
        End If
    Loop
    'Not found, return the negation of the position where it should be
    '(all higher values are to the right of the list and lower values are to the left)
    FindKey = Not min
    Exit Function
FindKey_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsIniManager.FindKey", Erl)
    Resume Next
End Function

''
'Searches for a main section with the given name within the loaded INI file and returns the index in which it's stored or the negation of the index in which it should be if not found.
'
' @param    name The name of the MainNode we are looking for.
' @returns  The index in which the main section we are looking for is stored or the negation of the index in which it should be if not found.
Private Function FindMain(ByVal Name As String) As Long
    On Error GoTo FindMain_Err
    'Returns the index of the MainNode which name matches the requested one,
    'or the negation of the position were it should be if not found
    Dim min As Long
    Dim max As Long
    Dim mid As Long
    min = 0
    max = MainNodes - 1
    Do While min <= max
        mid = (min + max) \ 2
        If fileData(mid).Name < Name Then
            min = mid + 1
        ElseIf fileData(mid).Name > Name Then
            max = mid - 1
        Else
            'We found it
            FindMain = mid
            Exit Function
        End If
    Loop
    'Not found, return the negation of the position where it should be
    '(all higher values are to the right of the list and lower values are to the left)
    FindMain = Not min
    Exit Function
FindMain_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsIniManager.FindMain", Erl)
    Resume Next
End Function

''
'Checks wether a given key exists or not.
'
' @param    name    The name of the element whose existance is being checked.
' @returns  True if the key exists, false otherwise.
Public Function KeyExists(ByVal Name As String) As Boolean
    'Returns true of the key exists, false otherwise.
    On Error GoTo KeyExists_Err
    KeyExists = FindMain(UCase$(Name)) >= 0
    Exit Function
KeyExists_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsIniManager.KeyExists", Erl)
    Resume Next
End Function

Private Function AddNode(ByRef Name As String) As Long
    On Error GoTo AddNode_Err
    Dim i       As Long
    Dim MainPos As Long
    Name = UCase$(Name)
    'Add a main node
    ReDim Preserve fileData(MainNodes) As MainNode
    If MainNodes Then
        i = MainNodes - 1
        MainPos = Not FindMain(Name)
        Do While i >= MainPos
            fileData(i + 1) = fileData(i)
            i = i - 1
        Loop
        fileData(MainPos).numValues = 0
        Erase fileData(MainPos).values
    End If
    fileData(MainPos).Name = Name
    MainNodes = MainNodes + 1
    AddNode = MainPos
    Exit Function
AddNode_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsIniManager.AddNode", Erl)
    Resume Next
End Function

Private Function AddKey(ByRef Main As String, ByRef key As String, Optional ByRef value As String = vbNullString) As Long
    On Error GoTo AddKey_Err
    Dim MainPos As Long
    Dim KeyPos  As Long
    Dim i       As Long
    Main = UCase$(Main)
    key = UCase$(key)
    MainPos = FindMain(Main)
    If MainPos < 0 Then MainPos = AddNode(Main)
    With fileData(MainPos)
        'Add it to the main node's value
        ReDim Preserve .values(.numValues) As ChildNode
        If .numValues > 0 Then
            i = .numValues - 1
            KeyPos = Not FindKey(fileData(MainPos), key)
            Do While i >= KeyPos
                .values(i + 1) = .values(i)
                i = i - 1
            Loop
        End If
        .values(KeyPos).key = key
        .values(KeyPos).value = value
        .numValues = .numValues + 1
    End With
    AddKey = KeyPos
    Exit Function
AddKey_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsIniManager.AddKey", Erl)
    Resume Next
End Function

Public Sub DumpFile(ByRef File As String)
    On Error GoTo DumpFile_Err
    Dim hFile As Integer
    Dim i     As Long
    Dim J     As Long
    hFile = FreeFile()
    Open File For Output As hFile
    For i = 0 To MainNodes - 1
        With fileData(i)
            Print #hFile, "[" & .Name & "]"
            For J = 0 To .numValues - 1
                Print #hFile, .values(J).key & "=" & .values(J).value
            Next J
            Print #hFile, ""
        End With
    Next i
    Close hFile
    Exit Sub
DumpFile_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsIniManager.DumpFile", Erl)
    Resume Next
End Sub

Public Function GetKey(ByVal Main As String, ByVal Index As Long) As String
    On Error GoTo errhandler
    Dim i As Long
    'Search for the main node
    i = FindMain(UCase$(Main))
    If i >= 0 Then
        GetKey = fileData(i).values(Index).key
    End If
    Exit Function
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "clsIniManager.GetKey", Erl)
    Resume Next
End Function

Public Property Get EntriesCount(ByVal Main As String) As Long
    On Error GoTo errhandler
    Dim i As Long
    'Search for the main node
    i = FindMain(UCase$(Main))
    If i >= 0 Then
        EntriesCount = fileData(i).numValues
    End If
    Exit Function
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "clsIniManager.GetKey", Erl)
    Resume Next
End Property

Public Sub GetPair(ByVal Main As String, ByVal Index As Long, ByRef key As String, ByRef value As String)
    On Error GoTo errhandler
    Dim i As Long
    'Search for the main node
    i = FindMain(UCase$(Main))
    If i >= 0 Then
        With fileData(i).values(Index)
            key = .key
            value = .value
        End With
    End If
    Exit Sub
errhandler:
    Call RegistrarError(Err.Number, Err.Description, "clsIniManager.GetPair", Erl)
    Resume Next
End Sub

Public Function GetNode(ByVal i As Long) As String
    GetNode = fileData(i).Name
End Function

Public Property Get NodesCount() As Long
    NodesCount = MainNodes
End Property
