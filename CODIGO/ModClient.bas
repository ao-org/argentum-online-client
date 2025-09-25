Attribute VB_Name = "ModClient"
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
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_EXSTYLE As Long = (-20)
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WS_EX_TRANSPARENT As Long = &H20&
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_ALPHA = &H2&

Public Sub Make_Transparent_Richtext(ByVal hWnd As Long)
    'If Win2kXP Then
    On Error GoTo Make_Transparent_Richtext_Err
    Call SetWindowLong(hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
    Exit Sub
Make_Transparent_Richtext_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModClient.Make_Transparent_Richtext", Erl)
    Resume Next
End Sub

Public Sub NameMapa(ByVal map As Long)
    On Error GoTo NameMapa_Err
    frmMain.NameMapa.Caption = MapDat.map_name
    If QueRender = 0 Then
        Letter_Set 0, MapDat.map_name
    End If
    Exit Sub
NameMapa_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModClient.NameMapa", Erl)
    Resume Next
End Sub

Public Sub PrintToConsole(text As String, _
                          Optional ByVal red As Integer = -1, _
                          Optional ByVal green As Integer, _
                          Optional ByVal blue As Integer, _
                          Optional ByVal bold As Boolean, _
                          Optional ByVal italic As Boolean, _
                          Optional ByVal bCrLf As Boolean, _
                          Optional ByVal FontTypeIndex As Byte = 0)
    On Error GoTo PrintToConsole_Err
    Dim bUrl As Boolean
    With frmMain.RecTxt
        '.SelFontName = "Tahoma"
        ' .SelFontSize = 8
        If FontTypeIndex <= 0 Then
            ' bUrl = True
            ' EnableUrlDetect
            If (Len(.text)) > 20000 Then .text = vbNullString
            .SelStart = Len(frmMain.RecTxt.text)
            .SelLength = 0
            .SelBold = IIf(bold, True, False)
            .SelItalic = IIf(italic, True, False)
            If Not red = -1 Then .SelColor = RGB(red, green, blue)
            .SelText = IIf(bCrLf, text, text & vbCrLf)
        Else
            If (Len(.text)) > 20000 Then .text = vbNullString
            'If FontTypeIndex = FONTTYPE_SERVER Then Text = "Servidor> " & Text
            '   bUrl = (FontTypeIndex = FONTTYPE_SERVER Or FontTypeIndex = FONTTYPE_TALK Or _
                FontTypeIndex = FONTTYPE_GUILDMSG Or FontTypeIndex = FONTTYPE_PIEL Or _
                FontTypeIndex = FONTTYPE_PIEL2)
            'If bUrl Then EnableUrlDetect
            .SelStart = Len(frmMain.RecTxt.text)
            .SelLength = 0
            .SelBold = FontTypes(FontTypeIndex).bold
            .SelItalic = FontTypes(FontTypeIndex).italic
            If Not red = -1 Then .SelColor = RGB(FontTypes(FontTypeIndex).red, FontTypes(FontTypeIndex).green, FontTypes(FontTypeIndex).blue)
            .SelText = IIf(bCrLf, text, text & vbCrLf)
        End If
    End With
    Exit Sub
PrintToConsole_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModClient.PrintToConsole", Erl)
    Resume Next
End Sub
