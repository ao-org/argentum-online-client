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
    
    On Error GoTo Parse_Form_Err
    

    If CursoresGraficos = 0 Then Exit Sub

    
     
    Dim aControl  As Control

    Dim lngHandle As Long
 
    Select Case cType

        Case E_NORMAL
            lngHandle = vbDefault

        Case E_ATTACK
            lngHandle = vbCrosshair

        Case E_ARROW
            lngHandle = vbCrosshair

        Case E_CAST
            lngHandle = vbCrosshair

        Case E_WAIT
            lngHandle = vbHourglass

        Case E_SHOOT
            lngHandle = vbCrosshair

        Case E_SHIP
            lngHandle = vbCrosshair

        Case E_agArrar
            lngHandle = vbCrosshair

    End Select
 
    For Each aControl In aFrm.Controls

        Select Case TypeName(aControl)

            Case "Timer", "Socket", "Shape", "Menu"

            Case Else
                aControl.MouseIcon = hndlList(cType)
                aControl.MousePointer = vbCustom

        End Select

    Next
 
    aFrm.MouseIcon = hndlList(cType)
    aFrm.MousePointer = vbCustom
 
    
    Exit Sub

Parse_Form_Err:
    Call RegistrarError(Err.number, Err.Description, "clsCursor.Parse_Form", Erl)
    Resume Next
    
End Sub
 
Public Function GetDefaultCursor(ByRef aFrm As Form) As CursorType
    
    On Error GoTo GetDefaultCursor_Err
    
 
    Dim lngHandle As Long
 
    Select Case aFrm.MouseIcon.handle

        Case hndlList(E_NORMAL).handle
            GetDefaultCursor = E_NORMAL

        Case hndlList(E_ATTACK).handle
            GetDefaultCursor = E_ATTACK

        Case hndlList(E_ARROW).handle
            GetDefaultCursor = E_ARROW

        Case hndlList(E_CAST).handle
            GetDefaultCursor = E_CAST

        Case hndlList(E_WAIT).handle
            GetDefaultCursor = E_WAIT

        Case hndlList(E_SHOOT).handle
            GetDefaultCursor = E_SHOOT

        Case hndlList(E_SHIP).handle
            GetDefaultCursor = E_SHIP

        Case hndlList(E_agArrar).handle
            GetDefaultCursor = E_agArrar

    End Select
 
    
    Exit Function

GetDefaultCursor_Err:
    Call RegistrarError(Err.number, Err.Description, "clsCursor.GetDefaultCursor", Erl)
    Resume Next
    
End Function
 
Public Sub Init()
    
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
    
End Sub
 
