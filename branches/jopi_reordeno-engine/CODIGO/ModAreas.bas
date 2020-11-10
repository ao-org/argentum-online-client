Attribute VB_Name = "ModAreas"
Option Explicit

'LAS GUARDAMOS PARA PROCESAR LOS MPs y sabes si borrar personajes
Public MinLimiteX      As Integer

Public MaxLimiteX      As Integer

Public MinLimiteY      As Integer

Public MaxLimiteY      As Integer

Private Const AREA_DIM As Byte = 12

Public Sub CambioDeArea(ByVal x As Byte, ByVal y As Byte)

    Dim loopX As Long, loopY As Long
    
    MinLimiteX = (x \ AREA_DIM - 1) * AREA_DIM
    MaxLimiteX = MinLimiteX + (AREA_DIM * 3) - 1
    
    MinLimiteY = (y \ AREA_DIM - 1) * AREA_DIM
    MaxLimiteY = MinLimiteY + (AREA_DIM * 3) - 1
    
    For loopX = 1 To 100
        For loopY = 1 To 100
            
            If (loopY < MinLimiteY) Or (loopY > MaxLimiteY) Or (loopX < MinLimiteX) Or (loopX > MaxLimiteX) Then
                'Erase NPCs
                
                If MapData(loopX, loopY).charindex > 0 Then
                    If MapData(loopX, loopY).charindex <> UserCharIndex Then
                        Call EraseChar(MapData(loopX, loopY).charindex)

                    End If

                End If
                
                'Erase OBJs
                MapData(loopX, loopY).ObjGrh.GrhIndex = 0
                
            End If
        
        Next loopY
    Next loopX
    
    Call RefreshAllChars

End Sub
