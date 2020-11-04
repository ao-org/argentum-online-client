Attribute VB_Name = "ModClient"
'RevolucionAo 1.0
'Pablo Mercavides
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_EXSTYLE As Long = (-20)
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long
Private Const WS_EX_TRANSPARENT As Long = &H20&
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&
Public Sub Make_Transparent_Richtext(ByVal hwnd As Long)
'If Win2kXP Then
    Call SetWindowLong(hwnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
End Sub
Public Sub NameMapa(ByVal Map As Long)
    'Dim DarNombreMapa As String

    'DarNombreMapa = DarNameMapa(Map)
    frmmain.NameMapa.Caption = MapDat.map_name
    
    engine.Letter_Set 0, MapDat.map_name
    
    'engine.Map_Letter_Fade_Set 1, 0
    

End Sub
Public Sub PrintToConsole(Text As String, Optional ByVal red As Integer = -1, Optional ByVal green As Integer, Optional ByVal blue As Integer, Optional ByVal bold As Boolean, Optional ByVal italic As Boolean, Optional ByVal bCrLf As Boolean, Optional ByVal FontTypeIndex As Byte = 0)
    
    Dim bUrl As Boolean
    
    With frmmain.RecTxt
        
        '.SelFontName = "Tahoma"
       ' .SelFontSize = 8
        
        If FontTypeIndex <= 0 Then
            
           ' bUrl = True
           ' EnableUrlDetect
            
            If (Len(.Text)) > 20000 Then .Text = vbNullString
            .SelStart = Len(frmmain.RecTxt.Text)
            .SelLength = 0
            
            
        
            .SelBold = IIf(bold, True, False)
            .SelItalic = IIf(italic, True, False)
            
            If Not red = -1 Then .SelColor = RGB(red, green, blue)
    
            .SelText = IIf(bCrLf, Text, Text & vbCrLf)
            
        Else
            If (Len(.Text)) > 20000 Then .Text = vbNullString
            
            'If FontTypeIndex = FONTTYPE_SERVER Then Text = "Servidor> " & Text
            
        '   bUrl = (FontTypeIndex = FONTTYPE_SERVER Or FontTypeIndex = FONTTYPE_TALK Or _
                FontTypeIndex = FONTTYPE_GUILDMSG Or FontTypeIndex = FONTTYPE_PIEL Or _
                FontTypeIndex = FONTTYPE_PIEL2)
                        
            'If bUrl Then EnableUrlDetect
            
            .SelStart = Len(frmmain.RecTxt.Text)
            .SelLength = 0

            .SelBold = FontTypes(FontTypeIndex).bold
            .SelItalic = FontTypes(FontTypeIndex).italic
            
            If Not red = -1 Then .SelColor = RGB(FontTypes(FontTypeIndex).red, FontTypes(FontTypeIndex).green, FontTypes(FontTypeIndex).blue)
    
            .SelText = IIf(bCrLf, Text, Text & vbCrLf)
            
        End If
    End With
    

    
End Sub

