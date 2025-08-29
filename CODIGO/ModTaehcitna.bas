Attribute VB_Name = "ModTaehcitna"
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

Private Const MAX_COMPROBACIONES As Byte = 4

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private ContadorMacroClicks(1 To MAX_COMPROBACIONES) As Position

Public LastSentPosX As Integer
Public LastSentPosY As Integer

Public Function ComprobarPosibleMacro(ByVal MouseX As Integer, ByVal MouseY As Integer) As Boolean
    On Error Goto ComprobarPosibleMacro_Err
    Call CopyMemory(ContadorMacroClicks(2), ContadorMacroClicks(1), Len(ContadorMacroClicks(1)) * (MAX_COMPROBACIONES - 1))
    
    ContadorMacroClicks(1).x = MouseX
    ContadorMacroClicks(1).y = MouseY
    
    Dim i As Byte
    
    For i = 1 To MAX_COMPROBACIONES
        If ContadorMacroClicks(i).x <> MouseX Or ContadorMacroClicks(i).y <> MouseY Then
            ComprobarPosibleMacro = False
            Exit Function
        End If
    Next i
    
    
    ComprobarPosibleMacro = True
    Call generarLogMacrero
    Exit Function
ComprobarPosibleMacro_Err:
    Call TraceError(Err.Number, Err.Description, "ModTaehcitna.ComprobarPosibleMacro", Erl)
End Function

Private Sub generarLogMacrero()
    On Error Goto generarLogMacrero_Err
    Call WriteLogMacroClickHechizo(tMacro.inasistidoPosFija)
    Exit Sub
generarLogMacrero_Err:
    Call TraceError(Err.Number, Err.Description, "ModTaehcitna.generarLogMacrero", Erl)
End Sub

Public Sub CountPacketIterations(ByRef packetControl As t_packetControl, ByVal expectedAverage As Double)
    On Error Goto CountPacketIterations_Err

    Dim delta As Long
    Dim actualcount As Long
    
    actualcount = GetTickCount()
    
    delta = actualcount - packetControl.last_count
    
    If delta < 40 Then Exit Sub
    
    packetControl.last_count = actualcount
    
    Call alterIndex(packetControl)
    
    packetControl.iterations(10) = delta
    Dim percentageDiff As Double, average As Double
    percentageDiff = getPercentageDiff(packetControl)
    average = getAverage(packetControl)
   ' frmdebug.add_text_tracebox "Delta: " & delta & " Average: " & average
    If percentageDiff < 5 Then
        'frmdebug.add_text_tracebox "DIFF: " & getPercentageDiff(packetControl)
        'Call AddtoRichTextBox(frmMain.RecTxt, "DIFF: " & getPercentageDiff(packetControl), 255, 200, 0, True)
        Call WriteRepeatMacro
        'frmdebug.add_text_tracebox "DIFF: " & getPercentageDiff(packetControl)
    End If
    
    If average > 20 And average < expectedAverage Then
        Call WriteRepeatMacro
    End If
    
    Exit Sub
CountPacketIterations_Err:
    Call TraceError(Err.Number, Err.Description, "ModTaehcitna.CountPacketIterations", Erl)
End Sub
Private Function getPercentageDiff(ByRef packetControl As t_packetControl) As Double
    On Error Goto getPercentageDiff_Err

    Dim i As Long, min As Long, max As Long
    
    min = packetControl.iterations(1)
    max = packetControl.iterations(1)
    
    Dim count As Long
    
    For i = 1 To 10
    
        If packetControl.iterations(i) < min Then
            min = packetControl.iterations(i)
        End If
        
        If packetControl.iterations(i) > max Then
            max = packetControl.iterations(i)
        End If
        
    Next i
    
    getPercentageDiff = 100 - ((min * 100) / max)
    
    Exit Function
getPercentageDiff_Err:
    Call TraceError(Err.Number, Err.Description, "ModTaehcitna.getPercentageDiff", Erl)
End Function

Private Function getAverage(ByRef packetControl As t_packetControl) As Double
    On Error Goto getAverage_Err

    Dim i As Long, suma As Long
    
    For i = 1 To UBound(packetControl.iterations)
        suma = suma + packetControl.iterations(i)
    Next i
    
    getAverage = suma / UBound(packetControl.iterations)
    
    Exit Function
getAverage_Err:
    Call TraceError(Err.Number, Err.Description, "ModTaehcitna.getAverage", Erl)
End Function

Private Sub alterIndex(ByRef packetControl As t_packetControl)
    On Error Goto alterIndex_Err
    Dim i As Long
    
    For i = 1 To 10 ' packetControl.cant_iterations
        If i < 10 Then 'packetControl.cant_iterations Then
            packetControl.iterations(i) = packetControl.iterations(i + 1)
        End If
    Next i
    Exit Sub
alterIndex_Err:
    Call TraceError(Err.Number, Err.Description, "ModTaehcitna.alterIndex", Erl)
End Sub



