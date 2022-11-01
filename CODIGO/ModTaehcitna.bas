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

Public Function ComprobarPosibleMacro(ByVal MouseX As Integer, ByVal MouseY As Integer) As Boolean
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
End Function

Private Sub generarLogMacrero()
    Call WriteLogMacroClickHechizo(tMacro.inasistidoPosFija)
End Sub

Public Sub CountPacketIterations(ByRef packetControl As t_packetControl, ByVal expectedAverage As Double)

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
   ' Debug.Print "Delta: " & delta & " Average: " & average
    If percentageDiff < 5 Then
        'Debug.Print "DIFF: " & getPercentageDiff(packetControl)
        'Call AddtoRichTextBox(frmMain.RecTxt, "DIFF: " & getPercentageDiff(packetControl), 255, 200, 0, True)
        Call WriteRepeatMacro
        'Debug.Print "DIFF: " & getPercentageDiff(packetControl)
    End If
    
    If average > 20 And average < expectedAverage Then
        Call WriteRepeatMacro
    End If
    
End Sub
Private Function getPercentageDiff(ByRef packetControl As t_packetControl) As Double

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
    
End Function

Private Function getAverage(ByRef packetControl As t_packetControl) As Double

    Dim i As Long, suma As Long
    
    For i = 1 To UBound(packetControl.iterations)
        suma = suma + packetControl.iterations(i)
    Next i
    
    getAverage = suma / UBound(packetControl.iterations)
    
End Function

Private Sub alterIndex(ByRef packetControl As t_packetControl)
    Dim i As Long
    
    For i = 1 To 10 ' packetControl.cant_iterations
        If i < 10 Then 'packetControl.cant_iterations Then
            packetControl.iterations(i) = packetControl.iterations(i + 1)
        End If
    Next i
End Sub

Public Sub efectoSangre()
    

        
    If Seguido = 1 Then
        Dim mouse As POINTAPI
        Dim MainLeft As Long
        Dim MainTop As Long
        Dim MainWidth As Long
        Dim MainHeight As Long
        
        MainWidth = frmMain.Width / 15
        MainHeight = frmMain.Height / 15
        
        
        GetCursorPos mouse
        
        MainLeft = frmMain.Left / 15
        MainTop = frmMain.Top / 15
        If mouse.x > MainLeft And mouse.y > MainTop And mouse.x < MainWidth + MainLeft And mouse.y < MainHeight + MainTop Then
            Cheat_X = mouse.x - MainLeft
            Cheat_Y = mouse.y - MainTop
            Call WriteSendPosSeguimiento(Cheat_X, Cheat_Y)
            'Debug.Print "X: " & mouse.x - MainLeft & "|Y: " & mouse.y - MainTop & "|Main Pos X: " & MainLeft / 15 & "|Main Pos Y: " & MainTop / 15
        End If
    End If
        
End Sub

