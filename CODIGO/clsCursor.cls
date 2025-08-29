VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'    Argentum 20 - Game Client Program
'    Copyright (C) 2022 - Noland Studios
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
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'
Option Explicit
 
Public Enum CursorType

    E_NORMAL = 0
    E_ATTACK
    E_ARROW
    E_CAST
    E_WAIT
    E_SHOOT
    E_SHIP
    E_agArrar

End Enum
 
Private Const NUM_CURSORS = 7
 
Private hndlList(0 To NUM_CURSORS) As IPictureDisp
 
Public Sub Parse_Form(ByRef aFrm As Form, Optional ByVal cType As CursorType = E_NORMAL)
    On Error Goto Parse_Form_Err
    On Error GoTo Parse_Form_Err

        ' Exit if CursoresGraficos is disabled
        If CursoresGraficos = 0 Then Exit Sub

20      Dim aControl As Control
30      Dim lngHandle As Long

        ' Determine the cursor handle based on the type
40      lngHandle = GetCursorHandle(cType)

140     ' Loop through each control on the form
        For Each aControl In aFrm.Controls
150         On Error Resume Next ' Minimal error handling inside loop
160         aControl.MouseIcon = hndlList(cType)
170         aControl.MousePointer = vbCustom
180         On Error GoTo Parse_Form_Err
        Next

190     ' Set the form's mouse properties
200     aFrm.MouseIcon = hndlList(cType)
210     aFrm.MousePointer = vbCustom

220     Exit Sub

Parse_Form_Err:
        ' Log the error with the exact line number where it occurred
240     Call RegistrarError(Err.Number, Err.Description & " | Line: " & Erl, "clsCursor.Parse_Form", Erl)
250     Resume Next
    Exit Sub
Parse_Form_Err:
    Call TraceError(Err.Number, Err.Description, "clsCursor.Parse_Form", Erl)
End Sub

' Helper function to determine the cursor handle based on type
Private Function GetCursorHandle(ByVal cType As CursorType) As Long
    On Error Goto GetCursorHandle_Err
    Select Case cType
        Case E_WAIT
            GetCursorHandle = vbHourglass
        Case E_NORMAL
            GetCursorHandle = vbDefault
        Case E_ATTACK, E_ARROW, E_CAST, E_SHOOT, E_SHIP, E_agArrar
            GetCursorHandle = vbCrosshair
        Case Else
            ' Default fallback to prevent uninitialized return values
            GetCursorHandle = vbDefault
    End Select
    Exit Function
GetCursorHandle_Err:
    Call TraceError(Err.Number, Err.Description, "clsCursor.GetCursorHandle", Erl)
End Function

Public Sub Init()
    On Error Goto Init_Err
    
    On Error GoTo Init_Err
    
 
    Set hndlList(E_NORMAL) = LoadResPicture("104", vbResCursor)
    Set hndlList(E_ATTACK) = LoadResPicture("102", vbResCursor)
    Set hndlList(E_ARROW) = LoadResPicture("101", vbResCursor)
    Set hndlList(E_CAST) = LoadResPicture("103", vbResCursor)
    Set hndlList(E_WAIT) = LoadResPicture("107", vbResCursor)
    Set hndlList(E_SHOOT) = LoadResPicture("106", vbResCursor)
    Set hndlList(E_SHIP) = LoadResPicture("105", vbResCursor)
    Set hndlList(E_agArrar) = LoadResPicture("108", vbResCursor)
 
    
    Exit Sub

Init_Err:
    Call RegistrarError(Err.number, Err.Description, "clsCursor.Init", Erl)
    Resume Next
    
    Exit Sub
Init_Err:
    Call TraceError(Err.Number, Err.Description, "clsCursor.Init", Erl)
End Sub
 
