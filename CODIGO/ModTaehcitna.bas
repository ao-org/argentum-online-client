Attribute VB_Name = "ModTaehcitna"
Option Explicit

Private Const MAX_COMPROBACIONES As Byte = 4
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
    Call WriteLogMacroClickHechizo
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
    Debug.Print "Delta: " & delta & " Average: " & average
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
