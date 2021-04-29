Attribute VB_Name = "ModTaehcitna"
Option Explicit

Private Const MAX_COMPROBACIONES As Byte = 4
Private ContadorMacroClicks(1 To MAX_COMPROBACIONES) As Position

Public Sub ComprobarPosibleMacro(ByVal MouseX As Integer, ByVal MouseY As Integer)
    Call CopyMemory(ContadorMacroClicks(2), ContadorMacroClicks(1), Len(ContadorMacroClicks(1)) * (MAX_COMPROBACIONES - 1))
    
    ContadorMacroClicks(1).x = MouseX
    ContadorMacroClicks(1).y = MouseY
    
    Dim i As Byte
    
    For i = 1 To MAX_COMPROBACIONES
        If ContadorMacroClicks(i).x <> MouseX Or ContadorMacroClicks(i).y <> MouseY Then Exit Sub
    Next i
    
    Call CloseClient
    
End Sub
