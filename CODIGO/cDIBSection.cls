VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDIBSection"
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

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Type SAFEARRAYBOUND

    cElements As Long
    lLbound As Long

End Type

Private Type SAFEARRAY2D

    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND

End Type

Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" (Ptr() As Any) As Long

Private Type RGBQUAD

    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte

End Type

Private Type BITMAPINFOHEADER '40 bytes

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

Private Type BITMAPINFO

    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD

End Type

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long

' Note - this is not the declare in the API viewer - modify lplpVoid to be
' Byref so we get the pointer back:
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

Private Const BI_RGB = 0&

Private Const DIB_RGB_COLORS = 0 '  color table in RGBs

Private Type BITMAP

    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long

End Type

Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

' Clipboard functions:
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function CloseClipboard Lib "user32" () As Long

Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long

Private Declare Function EmptyClipboard Lib "user32" () As Long

Private Const CF_BITMAP = 2

' Handle to the current DIBSection:
Private m_hDIb    As Long

' Handle to the old bitmap in the DC, for clear up:
Private m_hBmpOld As Long

' Handle to the Device context holding the DIBSection:
Private m_hDC     As Long

' Address of memory pointing to the DIBSection's bits:
Private m_lPtr    As Long

' Type containing the Bitmap information:
Private m_tBI     As BITMAPINFO

Public Function CopyToClipboard(Optional ByVal bAsDIB As Boolean = True) As Boolean
    
    On Error GoTo CopyToClipboard_Err
    

    Dim lhDCDesktop As Long

    Dim lhDC        As Long

    Dim lhBmpOld    As Long

    Dim hObj        As Long

    Dim lFmt        As Long

    lhDCDesktop = GetDC(GetDesktopWindow())

    If (lhDCDesktop <> 0) Then
        lhDC = CreateCompatibleDC(lhDCDesktop)

        If (lhDC <> 0) Then
            If (bAsDIB) Then
                MsgBox "I don't know how to put a DIB on the clipboard! Copy as bitmap instead!!!"
                ' Create a duplicate DIBSection and copy
                ' to the clipboard:
                'LSet tBI = m_tBI
                'hDibCopy = CreateDIBSection( _
                '        lhDC, _
                '        m_tBI, _
                '        DIB_RGB_COLORS, _
                '        lPtr, _
                '        0, 0)
                'If (hDibCopy <> 0) Then
                '    lhBmpOld = SelectObject(lhDC, hObj)
                '    BitBlt lhDC, 0, 0, Width, Height, m_hDC, 0, 0, vbSrcCopy
                '    SelectObject lhDC, lhBmpOld
                '    lFmt = CF_DIB
                '
                '     '....
                                    
                'Else
                '    hObj = 0
                'End If
            Else
                ' Create a compatible bitmap and copy to
                ' the clipboard:
                hObj = CreateCompatibleBitmap(lhDCDesktop, Width, Height)

                If (hObj <> 0) Then
                    lhBmpOld = SelectObject(lhDC, hObj)
                    PaintPicture lhDC
                    SelectObject lhDC, lhBmpOld
                    lFmt = CF_BITMAP

                    ' Now set the clipboard to the bitmap:
                    If (OpenClipboard(0) <> 0) Then
                        EmptyClipboard

                        If (SetClipboardData(lFmt, hObj) <> 0) Then
                            CopyToClipboard = True

                        End If

                        CloseClipboard

                    End If

                End If

            End If

            DeleteDC lhDC

        End If

        DeleteDC lhDCDesktop

    End If

    
    Exit Function

CopyToClipboard_Err:
    Call RegistrarError(Err.number, Err.Description, "cDIBSection.CopyToClipboard", Erl)
    Resume Next
    
End Function

Public Function CreateDIB(ByVal lhDC As Long, ByVal lWidth As Long, ByVal lHeight As Long, ByRef hDib As Long) As Boolean
    
    On Error GoTo CreateDIB_Err
    

    With m_tBI.bmiHeader
        .biSize = Len(m_tBI.bmiHeader)
        .biWidth = lWidth
        .biHeight = lHeight
        .biPlanes = 1
        .biBitCount = 24
        .biCompression = BI_RGB
        .biSizeImage = BytesPerScanLine * .biHeight

    End With

    hDib = CreateDIBSection(lhDC, m_tBI, DIB_RGB_COLORS, m_lPtr, 0, 0)
    CreateDIB = (hDib <> 0)

    
    Exit Function

CreateDIB_Err:
    Call RegistrarError(Err.number, Err.Description, "cDIBSection.CreateDIB", Erl)
    Resume Next
    
End Function

Public Function CreateFromPicture(ByRef picThis As StdPicture)
    
    On Error GoTo CreateFromPicture_Err
    

    Dim lhDC        As Long

    Dim lhDCDesktop As Long

    Dim lhBmpOld    As Long

    Dim tBMP        As BITMAP

    Dim lhWnd       As Long
    
    GetObjectAPI picThis.handle, Len(tBMP), tBMP

    If (Create(tBMP.bmWidth, tBMP.bmHeight)) Then
        lhWnd = GetDesktopWindow()
        lhDCDesktop = GetDC(lhWnd)

        If (lhDCDesktop <> 0) Then
            lhDC = CreateCompatibleDC(lhDCDesktop)
            ReleaseDC lhWnd, lhDCDesktop

            If (lhDC <> 0) Then
                lhBmpOld = SelectObject(lhDC, picThis.handle)
                LoadPictureBlt lhDC
                SelectObject lhDC, lhBmpOld
                DeleteDC lhDC

            End If

        End If

    End If

    
    Exit Function

CreateFromPicture_Err:
    Call RegistrarError(Err.number, Err.Description, "cDIBSection.CreateFromPicture", Erl)
    Resume Next
    
End Function

Public Function Create(ByVal lWidth As Long, ByVal lHeight As Long) As Boolean
    
    On Error GoTo Create_Err
    
    ClearUp
    m_hDC = CreateCompatibleDC(0)

    If (m_hDC <> 0) Then
        If (CreateDIB(m_hDC, lWidth, lHeight, m_hDIb)) Then
            m_hBmpOld = SelectObject(m_hDC, m_hDIb)
            Create = True
        Else
            DeleteDC m_hDC
            m_hDC = 0

        End If

    End If

    
    Exit Function

Create_Err:
    Call RegistrarError(Err.number, Err.Description, "cDIBSection.Create", Erl)
    Resume Next
    
End Function

Public Property Get BytesPerScanLine() As Long
    ' Scans must align on dword boundaries:
    
    On Error GoTo BytesPerScanLine_Err
    
    BytesPerScanLine = (m_tBI.bmiHeader.biWidth * 3 + 3) And &HFFFFFFFC

    
    Exit Property

BytesPerScanLine_Err:
    Call RegistrarError(Err.number, Err.Description, "cDIBSection.BytesPerScanLine", Erl)
    Resume Next
    
End Property

Public Property Get Width() As Long
    
    On Error GoTo Width_Err
    
    Width = m_tBI.bmiHeader.biWidth

    
    Exit Property

Width_Err:
    Call RegistrarError(Err.number, Err.Description, "cDIBSection.Width", Erl)
    Resume Next
    
End Property

Public Property Get Height() As Long
    
    On Error GoTo Height_Err
    
    Height = m_tBI.bmiHeader.biHeight

    
    Exit Property

Height_Err:
    Call RegistrarError(Err.number, Err.Description, "cDIBSection.Height", Erl)
    Resume Next
    
End Property

Public Sub LoadPictureBlt(ByVal lhDC As Long, Optional ByVal lSrcLeft As Long = 0, Optional ByVal lSrcTop As Long = 0, Optional ByVal lSrcWidth As Long = -1, Optional ByVal lSrcHeight As Long = -1, Optional ByVal eRop As RasterOpConstants = vbSrcCopy)
    
    On Error GoTo LoadPictureBlt_Err
    

    If lSrcWidth < 0 Then lSrcWidth = m_tBI.bmiHeader.biWidth
    If lSrcHeight < 0 Then lSrcHeight = m_tBI.bmiHeader.biHeight
    BitBlt m_hDC, 0, 0, lSrcWidth, lSrcHeight, lhDC, lSrcLeft, lSrcTop, eRop

    
    Exit Sub

LoadPictureBlt_Err:
    Call RegistrarError(Err.number, Err.Description, "cDIBSection.LoadPictureBlt", Erl)
    Resume Next
    
End Sub

Public Sub PaintPicture(ByVal lhDC As Long, Optional ByVal lDestLeft As Long = 0, Optional ByVal lDestTop As Long = 0, Optional ByVal lDestWidth As Long = -1, Optional ByVal lDestHeight As Long = -1, Optional ByVal lSrcLeft As Long = 0, Optional ByVal lSrcTop As Long = 0, Optional ByVal eRop As RasterOpConstants = vbSrcCopy)
    
    On Error GoTo PaintPicture_Err
    

    If (lDestWidth < 0) Then lDestWidth = m_tBI.bmiHeader.biWidth
    If (lDestHeight < 0) Then lDestHeight = m_tBI.bmiHeader.biHeight
    BitBlt lhDC, lDestLeft, lDestTop, lDestWidth, lDestHeight, m_hDC, lSrcLeft, lSrcTop, eRop

    
    Exit Sub

PaintPicture_Err:
    Call RegistrarError(Err.number, Err.Description, "cDIBSection.PaintPicture", Erl)
    Resume Next
    
End Sub

Public Property Get hdc() As Long
    
    On Error GoTo hdc_Err
    
    hdc = m_hDC

    
    Exit Property

hdc_Err:
    Call RegistrarError(Err.number, Err.Description, "cDIBSection.hdc", Erl)
    Resume Next
    
End Property

Public Property Get hDib() As Long
    
    On Error GoTo hDib_Err
    
    hDib = m_hDIb

    
    Exit Property

hDib_Err:
    Call RegistrarError(Err.number, Err.Description, "cDIBSection.hDib", Erl)
    Resume Next
    
End Property

Public Property Get DIBSectionBitsPtr() As Long
    
    On Error GoTo DIBSectionBitsPtr_Err
    
    DIBSectionBitsPtr = m_lPtr

    
    Exit Property

DIBSectionBitsPtr_Err:
    Call RegistrarError(Err.number, Err.Description, "cDIBSection.DIBSectionBitsPtr", Erl)
    Resume Next
    
End Property

Public Sub RandomiseBits(Optional ByVal bGray As Boolean = False)
    
    On Error GoTo RandomiseBits_Err
    

    Dim bDib() As Byte

    Dim x      As Long, y As Long

    Dim lC     As Long

    Dim tSA    As SAFEARRAY2D

    Dim xEnd   As Long
    
    ' Get the bits in the from DIB section:
    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = m_tBI.bmiHeader.biHeight
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = BytesPerScanLine()
        .pvData = m_lPtr

    End With

    CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4

    ' random:
    Randomize Timer
    
    xEnd = (Width - 1) * 3

    If (bGray) Then

        For y = 0 To m_tBI.bmiHeader.biHeight - 1
            For x = 0 To xEnd Step 3
                lC = Rnd * 255
                bDib(x, y) = lC
                bDib(x + 1, y) = lC
                bDib(x + 2, y) = lC
            Next x
        Next y

    Else

        For x = 0 To xEnd Step 3
            For y = 0 To m_tBI.bmiHeader.biHeight - 1
                bDib(x, y) = 0
                bDib(x + 1, y) = Rnd * 255
                bDib(x + 2, y) = Rnd * 255
            Next y
        Next x

    End If
    
    ' Clear the temporary array descriptor
    ' (This does not appear to be necessary, but
    ' for safety do it anyway)
    CopyMemory ByVal VarPtrArray(bDib), 0&, 4
    
    
    Exit Sub

RandomiseBits_Err:
    Call RegistrarError(Err.number, Err.Description, "cDIBSection.RandomiseBits", Erl)
    Resume Next
    
End Sub

Public Sub ClearUp()
    
    On Error GoTo ClearUp_Err
    

    If (m_hDC <> 0) Then
        If (m_hDIb <> 0) Then
            SelectObject m_hDC, m_hBmpOld
            DeleteObject m_hDIb

        End If

        DeleteDC m_hDC

    End If

    m_hDC = 0
    m_hDIb = 0
    m_hBmpOld = 0
    m_lPtr = 0

    
    Exit Sub

ClearUp_Err:
    Call RegistrarError(Err.number, Err.Description, "cDIBSection.ClearUp", Erl)
    Resume Next
    
End Sub

Public Function Resample(ByVal lNewHeight As Long, ByVal lNewWidth As Long) As cDIBSection
    
    On Error GoTo Resample_Err
    

    Dim cDib As cDIBSection

    Set cDib = New cDIBSection

    If cDib.Create(lNewWidth, lNewHeight) Then
        If (lNewWidth <> m_tBI.bmiHeader.biWidth) Or (lNewHeight <> m_tBI.bmiHeader.biHeight) Then
            ' Change in size, do resample:
            ResampleDib cDib
        Else
            ' No size change so just return a copy:
            cDib.LoadPictureBlt m_hDC

        End If

        Set Resample = cDib

    End If

    
    Exit Function

Resample_Err:
    Call RegistrarError(Err.number, Err.Description, "cDIBSection.Resample", Erl)
    Resume Next
    
End Function

Private Function ResampleDib(ByRef cDibTo As cDIBSection) As Boolean
    
    On Error GoTo ResampleDib_Err
    

    Dim bDibFrom() As Byte

    Dim bDibTo()   As Byte

    Dim tSAFrom    As SAFEARRAY2D

    Dim tSATo      As SAFEARRAY2D

    ' Get the bits in the from DIB section:
    With tSAFrom
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = m_tBI.bmiHeader.biHeight
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = BytesPerScanLine()
        .pvData = m_lPtr

    End With

    CopyMemory ByVal VarPtrArray(bDibFrom()), VarPtr(tSAFrom), 4

    ' Get the bits in the to DIB section:
    With tSATo
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = cDibTo.Height
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = cDibTo.BytesPerScanLine()
        .pvData = cDibTo.DIBSectionBitsPtr

    End With

    CopyMemory ByVal VarPtrArray(bDibTo()), VarPtr(tSATo), 4

    Dim xScale As Single

    Dim yScale As Single

    Dim x      As Long, y As Long, xEnd As Long, xOut As Long

    Dim fX     As Single, fY As Single

    Dim ifY    As Long, ifX As Long

    Dim dX     As Single, dy As Single

    Dim R      As Long, r1 As Single, r2 As Single, r3 As Single, r4 As Single

    Dim G      As Long, g1 As Single, g2 As Single, g3 As Single, g4 As Single

    Dim B      As Long, b1 As Single, b2 As Single, b3 As Single, b4 As Single

    Dim ir1    As Long, ig1 As Long, ib1 As Long

    Dim ir2    As Long, ig2 As Long, ib2 As Long

    xScale = (Width - 1) / cDibTo.Width
    yScale = (Height - 1) / cDibTo.Height
    
    xEnd = cDibTo.Width - 1
        
    For y = 0 To cDibTo.Height - 1
        
        fY = y * yScale
        ifY = Int(fY)
        dy = fY - ifY
        
        For x = 0 To xEnd
            fX = x * xScale
            ifX = Int(fX)
            dX = fX - ifX
            
            ifX = ifX * 3
            ' Interpolate using the four nearest pixels in the source
            b1 = bDibFrom(ifX, ifY)
            g1 = bDibFrom(ifX + 1, ifY)
            r1 = bDibFrom(ifX + 2, ifY)
            b2 = bDibFrom(ifX + 3, ifY)
            g2 = bDibFrom(ifX + 4, ifY)
            r2 = bDibFrom(ifX + 5, ifY)
            b3 = bDibFrom(ifX, ifY + 1)
            g3 = bDibFrom(ifX + 1, ifY + 1)
            r3 = bDibFrom(ifX + 2, ifY + 1)
            b4 = bDibFrom(ifX + 3, ifY + 1)
            g4 = bDibFrom(ifX + 4, ifY + 1)
            r4 = bDibFrom(ifX + 5, ifY + 1)
            
            ' Interplate in x direction:
            ir1 = r1 * (1 - dy) + r3 * dy
            ig1 = g1 * (1 - dy) + g3 * dy
            ib1 = b1 * (1 - dy) + b3 * dy
            ir2 = r2 * (1 - dy) + r4 * dy
            ig2 = g2 * (1 - dy) + g4 * dy
            ib2 = b2 * (1 - dy) + b4 * dy
            ' Interpolate in y:
            R = ir1 * (1 - dX) + ir2 * dX
            G = ig1 * (1 - dX) + ig2 * dX
            B = ib1 * (1 - dX) + ib2 * dX
            
            ' Set output:
            If (R < 0) Then R = 0
            If (R > 255) Then R = 255
            If (G < 0) Then G = 0
            If (G > 255) Then G = 255
            If (B < 0) Then B = 0
            If (B > 255) Then
                B = 255

            End If

            xOut = x * 3
            bDibTo(xOut, y) = B
            bDibTo(xOut + 1, y) = G
            bDibTo(xOut + 2, y) = R
            
        Next x
        
    Next y

    ' Clear the temporary array descriptor
    ' (This does not appear to be necessary, but
    ' for safety do it anyway)
    CopyMemory ByVal VarPtrArray(bDibFrom), 0&, 4
    CopyMemory ByVal VarPtrArray(bDibTo), 0&, 4

    
    Exit Function

ResampleDib_Err:
    Call RegistrarError(Err.number, Err.Description, "cDIBSection.ResampleDib", Erl)
    Resume Next
    
End Function

Private Sub Class_Terminate()
    
    On Error GoTo Class_Terminate_Err
    
    ClearUp

    
    Exit Sub

Class_Terminate_Err:
    Call RegistrarError(Err.number, Err.Description, "cDIBSection.Class_Terminate", Erl)
    Resume Next
    
End Sub

