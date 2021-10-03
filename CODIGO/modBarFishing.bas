Attribute VB_Name = "modBarFishing"
Option Explicit

Public showBarFishing                   As Boolean

Public barFocusFishingPOS               As tBarFocusFishing

Type tBarFocusFishing

    'pos actual. PosY es fija, PosX es variable
    pos_x       As Integer
    pos_y       As Integer
    
    'intervalo de movimiento
    lastMove    As Long
    typeMove    As Integer
    
    'totalidad de la barra
    max         As Long
    min         As Long
    
    'area verde
    minOk       As Long
    maxOk       As Long
    
End Type

Private center_X                         As Integer

Private center_Y                         As Integer

Private Const GRH_BAR_FISHING            As Long = 60666

Private Const GRH_BAR_FOCUS_FISHING      As Long = 60667

Private Const MIN_TIME_BAR_MOVE          As Long = 30 'ms

Private GRH_BAR                          As grh

Private GRH_FOCUS_BAR                    As grh

'@Agus: Seteamos las posiciones fijas de la barra
Public Sub setPositionBarFishing()

    center_X = (frmMain.renderer.Width / 2) - (GrhData(GRH_BAR_FISHING).pixelWidth / 2)
    center_Y = (frmMain.renderer.Height / 2) - (GrhData(GRH_BAR_FISHING).pixelHeight / 2)
    
    Call InitGrh(GRH_BAR, GRH_BAR_FISHING)
    Call InitGrh(GRH_FOCUS_BAR, GRH_BAR_FOCUS_FISHING)
    
    With barFocusFishingPOS
    
        .max = GrhData(GRH_BAR_FISHING).pixelWidth * 0.9
        .min = GrhData(GRH_BAR_FISHING).pixelWidth * 0.15
        
        .minOk = center_X / 2.6
        .maxOk = center_X / 1.4
        
        .pos_x = RandomNumber(.minOk, .maxOk * 1.1)
    
    End With

End Sub

'@Agus: Resistencia del pez
Private Sub fishResistance()

Dim time    As Long
Dim value   As Integer

time = GetTickCount()

With barFocusFishingPOS

        Call Draw_Grh(GRH_FOCUS_BAR, center_X + .pos_x, center_Y + 8, 0, 0, COLOR_WHITE, False)

        If (time - .lastMove) >= MIN_TIME_BAR_MOVE Then
            
            .lastMove = time
            
            If .pos_x < .max And .pos_x > .min Then
            
                If .pos_x > .maxOk Then
                    value = 1
                Else
                    value = -1
                End If
                
            End If
            
            .typeMove = value
            .pos_x = .pos_x + value
        
        End If

End With

End Sub

Public Sub userResistance(ByVal dir As Integer)

With barFocusFishingPOS

        If dir > 0 And .pos_x >= .max Then Exit Sub
        If dir < 0 And .pos_x <= .min Then Exit Sub
        
        
        .pos_x = .pos_x + (dir * 8)
        
        Call Draw_Grh(GRH_FOCUS_BAR, center_X + .pos_x, center_Y + 8, 0, 0, COLOR_WHITE, False)

End With


End Sub

Private Sub fishingStory()

With barFocusFishingPOS

    If .pos_x < .maxOk And .pos_x > .minOk Then 'dentro de la zona de éxito
        Engine_Text_Render "¡Muy bien!, sigue así y obtendrás un lindo pez", center_X + 18, center_Y + 40, COLOR_WHITE, 1
    Else
        Engine_Text_Render "¡Ten cuidado, el pez se está escapando!", center_X + 18, center_Y + 40, COLOR_WHITE, 1
    End If

End With

End Sub

Public Sub renderBarFishing()
 
    If showBarFishing Then

        Call Draw_Grh(GRH_BAR, center_X, center_Y, 0, 0, COLOR_WHITE, False)
        Call fishResistance
        Call fishingStory

    End If

End Sub


