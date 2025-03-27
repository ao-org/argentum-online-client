VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTexManager"
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

Const HASH_TABLE_SIZE As Long = 337
Const E_OUTOFMEMORY = 7

Private Type SURFACE_ENTRY_DYN
    FileName As Integer
    UltimoAcceso As Long
    Texture As Direct3DTexture8
    size As Long
    texture_width As Integer
    texture_height As Integer
End Type

Private Type HashNode
    surfaceCount As Integer
    SurfaceEntry() As SURFACE_ENTRY_DYN
End Type

Private TexList(HASH_TABLE_SIZE - 1) As HashNode

Private mD3D                         As D3DX8
Private device                       As Direct3DDevice8
Private mMaxEntries                  As Integer
Private mFreeMemoryBytes             As Long


Private Sub Class_Terminate()
    On Error GoTo Class_Terminate_Err
    Dim i As Long
    Dim j As Long
    'Destroy every surface in memory
    For i = 0 To HASH_TABLE_SIZE - 1
        With TexList(i)
            For j = 1 To .surfaceCount
                Set .SurfaceEntry(j).Texture = Nothing
            Next j
            'Destroy the arrays
            Erase .SurfaceEntry
        End With

    Next i


    Exit Sub

Class_Terminate_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsTexManager.Class_Terminate", Erl)
    Resume Next

End Sub

Public Function GetTexture(ByVal FileName As Integer, ByRef textwidth As Long, ByRef textheight As Long) As Direct3DTexture8
On Error GoTo GetTexture_Err
    Debug.Assert FileName > 0
    Set GetTexture = Nothing
    textwidth = 0
    textheight = 0
    Dim i As Long
    ' Search the index on the list
    With TexList(FileName Mod HASH_TABLE_SIZE)
        For i = 1 To .surfaceCount
            If .SurfaceEntry(i).FileName = FileName Then
                .SurfaceEntry(i).UltimoAcceso = GetTickCount()
                textwidth = .SurfaceEntry(i).texture_width
                textheight = .SurfaceEntry(i).texture_height
                Set GetTexture = .SurfaceEntry(i).Texture
                Exit Function
            End If
        Next i
    End With
    'Not in memory, load it!
    Set GetTexture = CreateDirect3dTexture(FileName, textwidth, textheight)
    Exit Function

GetTexture_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsTexManager.GetTexture", Erl)
    Resume Next
End Function

Public Function Init(ByRef D3D8 As D3DX8, ByRef d3d_device As Direct3DDevice8, ByVal MaxMemory As Long) As Boolean
On Error GoTo Init_Err

    Set mD3D = D3D8
    Set device = d3d_device
    mFreeMemoryBytes = MaxMemory
    Init = True

    Exit Function
Init_Err:
    Call RegistrarError(Err.Number, Err.Description, "clsTexManager.Init", Erl)
    Resume Next

End Function

Private Function LoadTexture(ByVal FileName As String, ByRef Dest As Direct3DTexture8) As Long

On Error GoTo LoadTexture_ErrHandler

    Dim bytArr() As Byte

    #If Compresion = 1 Then
        If Not Extract_File_To_Memory(Graphics, App.path & "\..\Recursos\OUTPUT\", LTrim(FileName) & ".png", bytArr, ResourcesPassword) Then
            frmDebug.add_text_tracebox "¡No se puede cargar el grafico numero " & FileName & "!"
            Exit Function
        End If

        Set Dest = mD3D.CreateTextureFromFileInMemoryEx( _
            device, bytArr(0), UBound(bytArr) + 1, _
            D3DX_DEFAULT, D3DX_DEFAULT, 1, 0, _
            D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT, _
            D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, &HFF000000, _
            ByVal 0, ByVal 0 _
        )
    #Else
        Dim PathToFile As String
        PathToFile = App.path & "\..\Recursos\Graficos\" & LTrim(FileName) & ".png"

        If Not FileExist(PathToFile, vbArchive) Then
            frmDebug.add_text_tracebox "¡No se puede cargar el grafico numero " & PathToFile & "!"
            Exit Function
        End If

        Set Dest = mD3D.CreateTextureFromFileEx( _
            device, PathToFile, _
            0, 0, 1, 0, _
            D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT, _
            D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, &HFF000000, _
            ByVal 0, ByVal 0 _
        )
    #End If

    Debug.Assert (Not (Dest Is Nothing))

    Exit Function

LoadTexture_ErrHandler:
    LoadTexture = Err.Number

End Function
Private Sub ReleaseMemory()
    Dim i As Long
    For i = 0 To 100
       Call RemoveLRU
    Next
End Sub
Private Function CreateDirect3dTexture(ByVal FileNum As Integer, ByRef TextureWidth As Long, ByRef TextureHeight As Long) As Direct3DTexture8

On Error GoTo ErrHandler

    Dim Texture         As Direct3DTexture8
    Dim surfaceDesc     As D3DSURFACE_DESC
    Dim loadResult      As Long
    Set CreateDirect3dTexture = Nothing
    TextureWidth = 0
    TextureHeight = 0
    loadResult = LoadTexture(str(FileNum), Texture)
    
    Select Case loadResult
        Case D3DERR_INVALIDCALL
            Debug.Assert False
            frmDebug.add_text_tracebox "LoadTexture failed with D3DERR_INVALIDCALL"
        Case D3DERR_NOTAVAILABLE
            Debug.Assert False
            frmDebug.add_text_tracebox "LoadTexture failed with D3DERR_NOTAVAILABLE"
        Case D3DXERR_INVALIDDATA
            Debug.Assert False
            frmDebug.add_text_tracebox "LoadTexture failed with D3DXERR_INVALIDDATA"
        Case D3DERR_OUTOFVIDEOMEMORY
        Case E_OUTOFMEMORY
            Call ReleaseMemory
            'Try to load the texture again, if it fails we've run out of options.
            loadResult = LoadTexture(str(FileNum), Texture)
    End Select
    
    If loadResult <> D3D_OK Then
       Call RegistrarError(loadResult, "Unhandled error", "clsTexManager.LoadTexture", 0)
       Exit Function
    Else
        Call Texture.GetLevelDesc(0, surfaceDesc)
    
        With TexList(FileNum Mod HASH_TABLE_SIZE)
            .surfaceCount = .surfaceCount + 1
            ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
    
            With .SurfaceEntry(.surfaceCount)
                .FileName = FileNum
                .UltimoAcceso = GetTickCount()
                Set .Texture = Texture
                .texture_width = surfaceDesc.Width
                .texture_height = surfaceDesc.Height
                .size = surfaceDesc.size
            End With
        End With
    
        'Keep track of how many memory we've been using.
        mFreeMemoryBytes = mFreeMemoryBytes - surfaceDesc.size
    
        'Ensure all return values are filled.
        TextureWidth = surfaceDesc.Width
        TextureHeight = surfaceDesc.Height
        Set CreateDirect3dTexture = Texture
    End If
        
    

    Exit Function

ErrHandler:
    frmDebug.add_text_tracebox "ERROR EN GRHLOAD>" & FileNum & ".png"

End Function

Public Function CreateTexture(ByVal Width As Long, ByVal Height As Long) As Direct3DTexture8
On Error GoTo ErrHandler
    Dim Texture As Direct3DTexture8
    Set Texture = mD3D.CreateTexture(device, Width, Height, 1, 0, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT)

    Set CreateTexture = Texture
    Exit Function
ErrHandler:
    frmDebug.add_text_tracebox "Failed to generate texture, " & Err.Description
End Function

Public Sub SetTextureData(ByRef Texture As Direct3DTexture8, ByRef Bytes() As Byte, ByVal size As Long, _
                          ByVal TextureWidth As Long, ByVal DrawWidth As Integer, _
                          ByVal StartY As Integer, ByVal endY As Integer)
On Error GoTo ErrHandler
    Debug.Assert Not Texture Is Nothing
    Dim lr As D3DLOCKED_RECT
    Dim RenderArea As Rect
    RenderArea.Left = 0
    RenderArea.Top = StartY
    RenderArea.Bottom = endY
    RenderArea.Right = DrawWidth
    Call Texture.LockRect(0, lr, ByVal RenderArea, 0)
    Dim BufferPos As Long
    Dim row As Long
    Dim destPtr As Long
    destPtr = lr.pBits
    For row = 0 To endY
        Call DXCopyMemory(ByVal destPtr, Bytes(BufferPos), DrawWidth * BytesPerPixel)
        BufferPos = BufferPos + DrawWidth * BytesPerPixel
        destPtr = destPtr + lr.Pitch
    Next row
    Call Texture.UnlockRect(0)
   Exit Sub
ErrHandler:
    Call RegistrarError(Err.Number, Err.Description, "clsTexManager.SetTextureData", Erl)
End Sub

Private Function RemoveLRU() As Boolean

        On Error GoTo RemoveLRU_Err

        Dim LRUi         As Long
        Dim LRUj         As Long
        Dim LRUtime      As Long
        Dim i            As Long
        Dim j            As Long

        Dim surface_desc As D3DSURFACE_DESC
100     LRUtime = GetTickCount()

        'Check out through the whole list for the least recently used
102     For i = 0 To HASH_TABLE_SIZE - 1
104         With TexList(i)
106             For j = 1 To .surfaceCount
108                 If LRUtime > .SurfaceEntry(j).UltimoAcceso And .SurfaceEntry(j).size > 0 Then
110                     LRUi = i
112                     LRUj = j
114                     LRUtime = .SurfaceEntry(j).UltimoAcceso

                    End If
116             Next j
            End With
118     Next i

        'Retrieve the surface desc
120     Call TexList(LRUi).SurfaceEntry(LRUj).Texture.GetLevelDesc(0, surface_desc)

        'Remove it
122     Set TexList(LRUi).SurfaceEntry(LRUj).Texture = Nothing
124     TexList(LRUi).SurfaceEntry(LRUj).FileName = 0

        'Move back the list (if necessary)
126     If LRUj Then
128         RemoveLRU = True
130         With TexList(LRUi)
132             For j = LRUj To .surfaceCount - 1
134                 .SurfaceEntry(j) = .SurfaceEntry(j + 1)
136             Next j
138             .surfaceCount = .surfaceCount - 1
140             If .surfaceCount Then
142                 ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
                Else
144                 Erase .SurfaceEntry
                End If
            End With
        End If
        'Update the used bytes
146     mFreeMemoryBytes = mFreeMemoryBytes + surface_desc.size
        Exit Function

RemoveLRU_Err:
     Call RegistrarError(Err.Number, Err.Description, "clsTexManager.RemoveLRU", Erl)
     Resume Next
End Function
