Attribute VB_Name = "modDX8Requires"
'RevolucionAo 1.0
'Pablo Mercavides
Option Explicit

'*******************************************************
' CARGA DE TEXTURAS
'*******************************************************
Public SurfaceDB As clsTexManager

Public Type D3D8Textures
    Texture As Direct3DTexture8
    texwidth As Long
    texheight As Long
End Type

'To get free bytes in drive
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, FreeBytesToCaller As Currency, BytesTotal As Currency, FreeBytesTotal As Currency) As Long

'To get free bytes in RAM
Private pUdtMemStatus As MEMORYSTATUS

Private Type MEMORYSTATUS

    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long

End Type

Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
'*******************************************************
' FIN - CARGA DE TEXTURAS
'*******************************************************

'*******************************************************
' MOTOR GRAFICO
'*******************************************************
' No matter what you do with DirectX8, you will need to start with
' the DirectX8 object. You will need to create a new instance of
' the object, using the New keyword, rather than just getting a
' pointer to it, since there's nowhere to get a pointer from yet (duh!).
Public DirectX As New DirectX8

' The D3DX8 object contains lots of helper functions, mostly math
' to make Direct3D alot easier to use. Notice we create a new
' instance of the object using the New keyword.
Public DirectD3D8 As D3DX8
Public DirectD3D As Direct3D8

' The Direct3DDevice8 represents our rendering device, which could
' be a hardware or a software device. The great thing is we still
' use the same object no matter what it is
Public DirectDevice As Direct3DDevice8

' The D3DDISPLAYMODE type structure that holds
' the information about your current display adapter.
Public DispMode  As D3DDISPLAYMODE
    
' The D3DPRESENT_PARAMETERS type holds a description of the way
' in which DirectX will display it's rendering.
Public D3DWindow As D3DPRESENT_PARAMETERS

Public SpriteBatch As New clsBatch

Public Projection As D3DMATRIX
Public IdentityMatrix As D3DMATRIX

Public Type TYPE_VERTEX

    x       As Single
    y       As Single
    z       As Single

    Color   As RGBA

    TX      As Single
    TY      As Single

End Type

Public Const PI As Single = 3.14159265358979

'*******************************************************
' FIN - MOTOR GRAFICO
'*******************************************************


Public Function General_Bytes_To_Megabytes(Bytes As Double) As Double

    Dim dblAns As Double

    dblAns = (Bytes / 1024) / 1024
    General_Bytes_To_Megabytes = format(dblAns, "###,###,##0.00")

End Function

Public Function General_Get_Free_Ram() As Double

    'Return Value in Megabytes
    Dim dblAns As Double

    GlobalMemoryStatus pUdtMemStatus
    dblAns = pUdtMemStatus.dwAvailPhys
    General_Get_Free_Ram = General_Bytes_To_Megabytes(dblAns)

End Function

Public Function General_Get_Free_Ram_Bytes() As Long
    GlobalMemoryStatus pUdtMemStatus
    General_Get_Free_Ram_Bytes = pUdtMemStatus.dwAvailPhys

End Function

Public Function ARGB(ByVal R As Long, ByVal G As Long, ByVal B As Long, ByVal A As Long) As Long
        
    Dim c As Long
        
    If A > 127 Then
        A = A - 128
        c = A * 2 ^ 24 Or &H80000000
        c = c Or R * 2 ^ 16
        c = c Or G * 2 ^ 8
        c = c Or B
    Else
        c = A * 2 ^ 24
        c = c Or R * 2 ^ 16
        c = c Or G * 2 ^ 8
        c = c Or B

    End If
    
    ARGB = c

End Function

