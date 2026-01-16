Attribute VB_Name = "mPosCrypto"
Option Explicit

Public Type tPosEnc
    cx As Byte
    cy As Byte
End Type

Public gPosKey(0 To 31) As Byte
Public gNoncePrefix(0 To 5) As Byte
Public gPosEpoch As Long
Public gPosUpdateCount As Long

Public Sub PosCrypto_InitFromSessionToken(ByVal token As String)
    Dim keyHash() As Byte
    Dim nonceHash() As Byte
    Dim tokenBytes() As Byte
    Dim i As Long

    Str2ByteArr "poskey|" & token, tokenBytes
    keyHash = hashBytes(tokenBytes, API_HASH_SHA256)
    If UBound(keyHash) >= 31 Then
        For i = 0 To 31
            gPosKey(i) = keyHash(i)
        Next i
    End If

    Str2ByteArr "posnonce|" & token, tokenBytes
    nonceHash = hashBytes(tokenBytes, API_HASH_SHA256)
    If UBound(nonceHash) >= 5 Then
        For i = 0 To 5
            gNoncePrefix(i) = nonceHash(i)
        Next i
    End If

    gPosEpoch = 0
    gPosUpdateCount = 0
End Sub

Private Sub BuildPosNonce(ByRef nonce() As Byte, ByVal charIndex As Integer, ByVal epoch As Long)
    Dim i As Long
    ReDim nonce(0 To 11)
    For i = 0 To 5
        nonce(i) = gNoncePrefix(i)
    Next i
    nonce(6) = CByte(charIndex And &HFF)
    nonce(7) = CByte((charIndex And &HFF00) \ &H100)
    nonce(8) = CByte(epoch And &HFF)
    nonce(9) = CByte((epoch And &HFF00) \ &H100)
    nonce(10) = CByte((epoch And &HFF0000) \ &H10000)
    nonce(11) = CByte((epoch And &HFF000000) \ &H1000000)
End Sub

Private Sub PosSetWithEpoch(ByRef p As tPosEnc, ByVal charIndex As Integer, ByVal x As Byte, ByVal y As Byte, ByVal epoch As Long)
    Dim inputBytes(0 To 1) As Byte
    Dim nonce() As Byte
    Dim outputBytes() As Byte

    inputBytes(0) = x
    inputBytes(1) = y
    BuildPosNonce nonce, charIndex, epoch
    outputBytes = cipherStreamBytes(inputBytes, gPosKey, nonce, API_SC_CHACHA20, 0)
    If UBound(outputBytes) >= 1 Then
        p.cx = outputBytes(0)
        p.cy = outputBytes(1)
    End If
End Sub

Private Sub PosGetWithEpoch(ByRef p As tPosEnc, ByVal charIndex As Integer, ByVal epoch As Long, ByRef x As Byte, ByRef y As Byte)
    Dim inputBytes(0 To 1) As Byte
    Dim nonce() As Byte
    Dim outputBytes() As Byte

    inputBytes(0) = p.cx
    inputBytes(1) = p.cy
    BuildPosNonce nonce, charIndex, epoch
    outputBytes = cipherStreamBytes(inputBytes, gPosKey, nonce, API_SC_CHACHA20, 0)
    If UBound(outputBytes) >= 1 Then
        x = outputBytes(0)
        y = outputBytes(1)
    End If
End Sub

Public Sub PosSet(ByRef p As tPosEnc, ByVal charIndex As Integer, ByVal x As Byte, ByVal y As Byte)
    PosSetWithEpoch p, charIndex, x, y, gPosEpoch
End Sub

Public Sub PosGet(ByRef p As tPosEnc, ByVal charIndex As Integer, ByRef x As Byte, ByRef y As Byte)
    PosGetWithEpoch p, charIndex, gPosEpoch, x, y
End Sub

Public Sub PosRotateEpoch(ByVal newEpoch As Long)
    Dim i As Long
    Dim x As Byte
    Dim y As Byte

    If newEpoch = gPosEpoch Then Exit Sub
    For i = LBound(charlist) To UBound(charlist)
        If charlist(i).active = 1 Then
            PosGetWithEpoch charlist(i).PosEnc, i, gPosEpoch, x, y
            PosSetWithEpoch charlist(i).PosEnc, i, x, y, newEpoch
        End If
    Next i
    gPosEpoch = newEpoch
End Sub

Public Sub PosMaybeRotate()
    gPosUpdateCount = gPosUpdateCount + 1
    If (gPosUpdateCount And &HFF) = 0 Then
        PosRotateEpoch gPosEpoch + 1
    End If
End Sub
