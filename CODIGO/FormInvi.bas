Attribute VB_Name = "ModuloFunciones"
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
'Dieser Source stammt von http://www.activevb.de
'und kann frei verwendet werden. Für eventuelle Schäden
'wird nicht gehaftet.
'
'Um Fehler oder Fragen zu klären, nutzen Sie bitte unser Forum.
'Ansonsten viel Spaß und Erfolg mit diesem Source!

'Code von Benjamin Wilger
'Benjamin@ActiveVB.de
'Copyright (C) 2001

Option Explicit

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Const RGN_OR As Long = 2&

Private Declare Sub OleTranslateColor Lib "oleaut32.dll" (ByVal clr As Long, ByVal hpal As Long, ByRef lpcolorref As Long)

Private Type BITMAPINFOHEADER

    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long

End Type

Private Type RGBQUAD

    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte

End Type

Private Type BITMAPINFO

    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD

End Type

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByVal lplpVoid As Long, ByVal Handle As Long, ByVal dw As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Const BI_RGB         As Long = 0&

Private Const DIB_RGB_COLORS As Long = 0&

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Const LWA_COLORKEY  As Long = &H1&

Private Const GWL_EXSTYLE   As Long = (-20&)

Private Const WS_EX_LAYERED As Long = &H80000

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Public Const WM_NCLBUTTONDOWN As Long = &HA1&

Public Const HTCAPTION        As Long = 2&

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function ReleaseCapture Lib "user32" () As Long

'Remove Title Bar

Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const GWL_STYLE = (-16)

Private Const WS_CAPTION = &HC00000

Private Const SWP_FRAMECHANGED = &H20

Private Const SWP_NOMOVE = &H2

Private Const SWP_NOZORDER = &H4

Private Const SWP_NOSIZE = &H1
 
Private Const EM_LINEINDEX = &HBB

Private Const EM_LINELENGTH = &HC1

Private Const EM_GETLINE = &HC4
 
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lPara As Long) As Long

Private Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lPara As String) As Long

Private Const LWA_ALPHA = &H2

' Constantes para SendMessage
Const WM_SYSCOMMAND As Long = &H112&

Const MOUSE_MOVE    As Long = &HF012&
 
Public Function TextoLineaRichTextBox(ByVal pControl As RichTextBox, ByVal pLinea As Long) As String
    
    On Error GoTo TextoLineaRichTextBox_Err
    

    

    Dim vLongitudLinea As Long, vNumeroLinea As Long

    Dim vTemporal      As String
   
    TextoLineaRichTextBox = ""
   
    vNumeroLinea = SendMessageLong(pControl.hWnd, EM_LINEINDEX, pLinea, 0&)
    vLongitudLinea = SendMessageLong(pControl.hWnd, EM_LINELENGTH, vNumeroLinea, 0&) + 1
    vTemporal = String$((vLongitudLinea + 2), 0)
   
    Mid$(vTemporal, 1, 1) = Chr$(vLongitudLinea And &HFF)
    Mid$(vTemporal, 2, 1) = Chr$(vLongitudLinea \ &H100)
   
    vLongitudLinea = SendMessageString(pControl.hWnd, EM_GETLINE, pLinea, vTemporal)
   
    If (vLongitudLinea > 0) Then
        TextoLineaRichTextBox = Left$(vTemporal, vLongitudLinea)

    End If

    
    Exit Function

TextoLineaRichTextBox_Err:
    Call RegistrarError(Err.number, Err.Description, "ModuloFunciones.TextoLineaRichTextBox", Erl)
    Resume Next
    
End Function
 
' Sacado de https://www.vbforums.com/showthread.php?379880-RESOLVED-Remove-Title-Bar-Off-Of-Form-Using-API-s
' Borro algunas partes innecesarias (WyroX)
Public Sub Form_RemoveTitleBar(F As Form)
    
    On Error GoTo Form_RemoveTitleBar_Err
    

    Dim Style As Long

    ' Get window's current style bits.
    Style = GetWindowLong(F.hWnd, GWL_STYLE)
    ' Set the style bit for the title off.
    Style = Style And Not WS_CAPTION

    ' Send the new style to the window.
    SetWindowLong F.hWnd, GWL_STYLE, Style

    ' Repaint the window.
    'SetWindowPos f.hwnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE
    
    Exit Sub

Form_RemoveTitleBar_Err:
    Call RegistrarError(Err.number, Err.Description, "ModuloFunciones.Form_RemoveTitleBar", Erl)
    Resume Next
    
End Sub

Public Function MakeFormTransparent(Frm As Form, ByVal lngTransColor As Long)
    
    On Error GoTo MakeFormTransparent_Err
    

    Dim hRegion  As Long

    Dim WinStyle As Long
    
    'Systemfarben ggf. in RGB-Werte übersetzen
    If lngTransColor < 0 Then OleTranslateColor lngTransColor, 0&, lngTransColor

    'Ab Windows 2000/98 geht das relativ einfach per API
    'Mit IsFunctionExported wird geprüft, ob die Funktion
    'SetLayeredWindowAttributes unter diesem Betriebsystem unterstützt wird.
    If IsFunctionExported("SetLayeredWindowAttributes", "user32") Then
        'Den Fenster-Stil auf "Layered" setzen
        WinStyle = GetWindowLong(Frm.hWnd, GWL_EXSTYLE)
        WinStyle = WinStyle Or WS_EX_LAYERED
        SetWindowLong Frm.hWnd, GWL_EXSTYLE, WinStyle
        SetLayeredWindowAttributes Frm.hWnd, lngTransColor, 0&, LWA_COLORKEY
        
    Else 'Manuell die Region erstellen und übernehmen
        hRegion = RegionFromBitmap(Frm, lngTransColor)
        SetWindowRgn Frm.hWnd, hRegion, True
        DeleteObject hRegion

    End If

    
    Exit Function

MakeFormTransparent_Err:
    Call RegistrarError(Err.number, Err.Description, "ModuloFunciones.MakeFormTransparent", Erl)
    Resume Next
    
End Function

Private Function RegionFromBitmap(picSource As Object, ByVal lngTransColor As Long) As Long
    
    On Error GoTo RegionFromBitmap_Err
    

    Dim lngRetr      As Long, lngHeight As Long, lngWidth As Long

    Dim lngRgnFinal  As Long, lngRgnTmp As Long

    Dim lngStart     As Long

    Dim x            As Long, y As Long

    Dim hdc          As Long
    
    Dim bi24BitInfo  As BITMAPINFO

    Dim iBitmap      As Long

    Dim BWidth       As Long

    Dim BHeight      As Long

    Dim iDC          As Long

    Dim PicBits()    As Byte

    Dim Col          As Long

    Dim OldScaleMode As ScaleModeConstants
    
    OldScaleMode = picSource.ScaleMode
    picSource.ScaleMode = vbPixels
    
    hdc = picSource.hdc
    lngWidth = picSource.ScaleWidth '- 1
    lngHeight = picSource.ScaleHeight - 1

    BWidth = (picSource.ScaleWidth \ 4) * 4 + 4
    BHeight = picSource.ScaleHeight

    'Bitmap-Header
    With bi24BitInfo.bmiHeader
        .biBitCount = 24
        .biCompression = BI_RGB
        .biPlanes = 1
        .biSize = Len(bi24BitInfo.bmiHeader)
        .biWidth = BWidth
        .biHeight = BHeight + 1

    End With

    'ByteArrays in der erforderlichen Größe anlegen
    ReDim PicBits(0 To bi24BitInfo.bmiHeader.biWidth * 3 - 1, 0 To bi24BitInfo.bmiHeader.biHeight - 1)
    
    iDC = CreateCompatibleDC(hdc)
    'Gerätekontextunabhängige Bitmap (DIB) erzeugen
    iBitmap = CreateDIBSection(iDC, bi24BitInfo, DIB_RGB_COLORS, ByVal 0&, ByVal 0&, ByVal 0&)
    'iBitmap in den neuen DIB-DC wählen
    Call SelectObject(iDC, iBitmap)
    'hDC des Quell-Fensters in den hDC der DIB kopieren
    Call BitBlt(iDC, 0, 0, bi24BitInfo.bmiHeader.biWidth, bi24BitInfo.bmiHeader.biHeight, hdc, 0, 0, vbSrcCopy)
    'Gerätekontextunabhängige Bitmap in ByteArrays kopieren
    Call GetDIBits(hdc, iBitmap, 0, bi24BitInfo.bmiHeader.biHeight, PicBits(0, 0), bi24BitInfo, DIB_RGB_COLORS)
    
    'Wir brauchen nur den Array, also können wir die Bitmap direkt wieder löschen.
    
    'DIB-DC
    Call DeleteDC(iDC)
    'Bitmap
    Call DeleteObject(iBitmap)

    lngRgnFinal = CreateRectRgn(0, 0, 0, 0)

    For y = 0 To lngHeight
        x = 0

        Do While x < lngWidth
            Do While x < lngWidth And RGB(PicBits(x * 3 + 2, lngHeight - y + 1), PicBits(x * 3 + 1, lngHeight - y + 1), PicBits(x * 3, lngHeight - y + 1)) = lngTransColor
                
                x = x + 1
            Loop

            If x <= lngWidth Then
                lngStart = x

                Do While x < lngWidth And RGB(PicBits(x * 3 + 2, lngHeight - y + 1), PicBits(x * 3 + 1, lngHeight - y + 1), PicBits(x * 3, lngHeight - y + 1)) <> lngTransColor
                    x = x + 1
                Loop

                If x + 1 > lngWidth Then x = lngWidth
                lngRgnTmp = CreateRectRgn(lngStart, y, x, y + 1)
                lngRetr = CombineRgn(lngRgnFinal, lngRgnFinal, lngRgnTmp, RGN_OR)
                DeleteObject lngRgnTmp

            End If

        Loop
    Next

    picSource.ScaleMode = OldScaleMode
    RegionFromBitmap = lngRgnFinal

    
    Exit Function

RegionFromBitmap_Err:
    Call RegistrarError(Err.number, Err.Description, "ModuloFunciones.RegionFromBitmap", Erl)
    Resume Next
    
End Function

'Code von vbVision:
'Diese Funktion überprüft, ob die angegebene Function von einer DLL exportiert wird.
Private Function IsFunctionExported(ByVal sFunction As String, ByVal sModule As String) As Boolean
    
    On Error GoTo IsFunctionExported_Err
    

    

    Dim hMod As Long, lpFunc As Long, bLibLoaded As Boolean
    
    'Handle der DLL erhalten
    hMod = GetModuleHandle(sModule)

    If hMod = 0 Then 'Falls DLL nicht registriert ...
        hMod = LoadLibrary(sModule) 'DLL in den Speicher laden.

        If hMod Then bLibLoaded = True

    End If
    
    If hMod Then
        If GetProcAddress(hMod, sFunction) Then IsFunctionExported = True

    End If
    
    If bLibLoaded Then Call FreeLibrary(hMod)

    
    Exit Function

IsFunctionExported_Err:
    Call RegistrarError(Err.number, Err.Description, "ModuloFunciones.IsFunctionExported", Erl)
    Resume Next
    
End Function

Public Function SuperMid(ByVal strMain As String, str1 As String, str2 As String, Optional reverse As Boolean) As String

    'DESCRIPTION: Extract the portion of a string between the two substrings defined in str1 and str2.
    'DEVELOPER: Ryan Wells (wellsr.com)
    'HOW TO USE: - Pass the argument your main string and the 2 strings you want to find in the main string.
    ' - This function will extract the values between the end of your first string and the beginning
    ' of your next string.
    ' - If the optional boolean "reverse" is true, an InStrRev search will occur to find the last
    ' instance of the substrings in your main string.
    Dim i As Integer, j As Integer, temp As Variant

    On Error GoTo errhandler:

    If reverse = True Then
        i = InStrRev(strMain, str1)
        j = InStrRev(strMain, str2)

        If Abs(j - i) < Len(str1) Then j = InStrRev(strMain, str2, i)
        If i = j Then 'try to search 2nd half of string for unique match
            j = InStrRev(strMain, str2, i - 1)

        End If

    Else
        i = InStr(1, strMain, str1)
        j = InStr(1, strMain, str2)

        If Abs(j - i) < Len(str1) Then j = InStr(i + Len(str1), strMain, str2)
        If i = j Then 'try to search 2nd half of string for unique match
            j = InStr(i + 1, strMain, str2)

        End If

    End If

    If i = 0 And j = 0 Then GoTo errhandler:
    If j = 0 Then j = Len(strMain) + Len(str2) 'just to make it arbitrarily large
    If i = 0 Then i = Len(strMain) + Len(str1) 'just to make it arbitrarily large
    If i > j And j <> 0 Then 'swap order
        temp = j
        j = i
        i = temp
        temp = str2
        str2 = str1
        str1 = temp

    End If

    i = i + Len(str1)
    SuperMid = mid(strMain, i, j - i)
    Exit Function
errhandler:
    SuperMid = "A"

    'MsgBox "Error extracting strings. Check your input" & vbNewLine & vbNewLine & "Aborting", , "Strings not found"
End Function

'Función que aplica la transparencia, se le pasa el hwnd del form y un valor de 0 a 255
Public Function Aplicar_Transparencia(ByVal hWnd As Long, Valor As Integer) As Long
    
    On Error GoTo Aplicar_Transparencia_Err
    
  
    Dim msg As Long
  
    
  
    If Valor < 0 Or Valor > 255 Then
        Aplicar_Transparencia = 1
    Else
        msg = GetWindowLong(hWnd, GWL_EXSTYLE)
        msg = msg Or WS_EX_LAYERED
     
        SetWindowLong hWnd, GWL_EXSTYLE, msg
     
        'Establece la transparencia
        SetLayeredWindowAttributes hWnd, 0, Valor, LWA_ALPHA
  
        Aplicar_Transparencia = 0
  
    End If
  
    If Err Then
        Aplicar_Transparencia = 2

    End If
  
    
    Exit Function

Aplicar_Transparencia_Err:
    Call RegistrarError(Err.number, Err.Description, "ModuloFunciones.Aplicar_Transparencia", Erl)
    Resume Next
    
End Function

Public Sub MoverForm(ByVal hWnd As Long)
    
    On Error GoTo moverForm_Err
    

    Dim res As Long

    ReleaseCapture
    res = SendMessage(hWnd, WM_SYSCOMMAND, MOUSE_MOVE, 0)

    
    Exit Sub

moverForm_Err:
    Call RegistrarError(Err.number, Err.Description, "ModuloFunciones.MoverForm", Erl)
    Resume Next
    
End Sub
