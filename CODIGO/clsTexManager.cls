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

Private mCantidadGraficos            As Integer

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
    Call RegistrarError(Err.number, Err.Description, "clsTexManager.Class_Terminate", Erl)
    Resume Next
    
End Sub

Public Function GetTexture(ByVal FileName As Integer, ByRef textwidth As Long, ByRef textheight As Long) As Direct3DTexture8
    
    On Error GoTo GetTexture_Err
    

    If FileName = 0 Then
        Debug.Print "0 GRH ATMPT TO BE LOADED"
        Exit Function

    End If

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
    Call RegistrarError(Err.number, Err.Description, "clsTexManager.GetTexture", Erl)
    Resume Next
    
End Function

Public Function Init(ByRef D3D8 As D3DX8, ByRef d3d_device As Direct3DDevice8, ByVal MaxMemory As Long) As Boolean
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 5/04/2005
    '
    '**************************************************************
    
    On Error GoTo Init_Err
    
    mCantidadGraficos = 0
    
    'Seteamos el objeto
    Set mD3D = D3D8
    Set device = d3d_device
    mFreeMemoryBytes = MaxMemory
    
    Init = True

    
    Exit Function

Init_Err:
    Call RegistrarError(Err.number, Err.Description, "clsTexManager.Init", Erl)
    Resume Next
    
End Function

Private Function CreateDirect3dTexture(ByVal FileNum As Integer, ByRef texture_width As Long, ByRef texture_height As Long) As Direct3DTexture8
On Error GoTo ErrHandler
    Dim surface_desc As D3DSURFACE_DESC
    Dim Index        As Integer
    Dim DirArchico   As String
    Dim bytArr()     As Byte
    
    Set CreateDirect3dTexture = Nothing
    texture_width = 0
    texture_height = 0
    
    Index = FileNum Mod HASH_TABLE_SIZE

    With TexList(Index)
        .surfaceCount = .surfaceCount + 1
        ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN

        With .SurfaceEntry(.surfaceCount)
            .FileName = FileNum
            .UltimoAcceso = GetTickCount()
            #If Compresion = 1 Then
                If Not Extract_File_To_Memory(Graphics, App.path & "\..\Recursos\OUTPUT\", LTrim(str(FileNum)) & ".png", bytArr, ResourcesPassword) Then
                    Debug.Print "¡No se puede cargar el grafico numero " & FileNum & "!"
                    Exit Function
                End If
                Set .Texture = mD3D.CreateTextureFromFileInMemoryEx(device, bytArr(0), UBound(bytArr) + 1, D3DX_DEFAULT, D3DX_DEFAULT, 1, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, &HFF000000, ByVal 0, ByVal 0)
            #Else
                Dim PathToFile As String
                PathToFile = App.path & "\..\Recursos\Graficos\" & LTrim(str(FileNum)) & ".png"
                If Not FileExist(PathToFile, vbArchive) Then
                    Debug.Print "¡No se puede cargar el grafico numero " & PathToFile & "!"
                    Exit Function
                End If

                Set .Texture = mD3D.CreateTextureFromFileEx(device, PathToFile, 0, 0, _
                     1, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, &HFF000000, ByVal 0, ByVal 0)
            #End If
                
            .Texture.GetLevelDesc 0, surface_desc
            .texture_width = surface_desc.Width
            .texture_height = surface_desc.Height
            .size = surface_desc.size
            texture_width = .texture_width
            texture_height = .texture_height
            
            Set CreateDirect3dTexture = TexList(Index).SurfaceEntry(TexList(Index).surfaceCount).Texture
        End With

    End With

    Exit Function

errhandler:
    Debug.Print "ERROR EN GRHLOAD>" & FileNum & ".png"

End Function

Public Function CreateTexture(ByVal Width As Long, ByVal Height As Long) As Direct3DTexture8
On Error GoTo ErrHandler
    Dim Texture As Direct3DTexture8
    Set Texture = mD3D.CreateTexture(device, Width, Height, 1, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED)
    
    Set CreateTexture = Texture
    Exit Function
ErrHandler:
    Debug.Print "Failed to generate texture, " & Err.Description
End Function

Public Sub SetTextureData(ByRef Texture As Direct3DTexture8, ByRef Bytes() As Byte, ByVal size As Long, _
                          ByVal TextureWidth As Long, ByVal DrawWidth As Integer, _
                          ByVal startY As Integer, ByVal endY As Integer)
On Error GoTo ErrHandler
    Debug.Assert Not Texture Is Nothing
    Dim lr As D3DLOCKED_RECT
    Dim RenderArea As RECT
    RenderArea.Left = 0
    RenderArea.Top = startY
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

Public Function CreateTextureFromData(ByRef Bytes() As Byte, ByVal Size As Long, ByVal Width As Long, ByVal Height As Long) As Direct3DTexture8
On Error GoTo ErrHandler
    Dim Texture As Direct3DTexture8
    Set Texture = mD3D.CreateTextureFromFileInMemoryEx(device, Bytes(0), UBound(Bytes) + 1, Width, Height, 1, 0, D3DFMT_A8R8G8B8, D3DPOOL_DEFAULT, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, &HFF000000, ByVal 0, ByVal 0)
    Set CreateTextureFromData = Texture
    Debug.Assert Not Texture Is Nothing
    Exit Function
ErrHandler:
    Debug.Print "Failed to generate texture from memory, " & Err.Description
End Function


Private Function RemoveLRU() As Boolean
    
    On Error GoTo RemoveLRU_Err
    

    '**************************************************************
    'Author: Juan Mart?n Sotuyo Dodero
    'Last Modify Date: 3/06/2006
    'Removes the Least Recently Used surface to make some room for new ones
    '**************************************************************
    Dim LRUi         As Long

    Dim LRUj         As Long

    Dim LRUtime      As Long

    Dim i            As Long

    Dim j            As Long

    Dim surface_desc As D3DSURFACE_DESC
    
    LRUtime = GetTickCount()
    
    'Check out through the whole list for the least recently used
    For i = 0 To HASH_TABLE_SIZE - 1

        With TexList(i)

            For j = 1 To .surfaceCount

                If LRUtime > .SurfaceEntry(j).UltimoAcceso Then
                    LRUi = i
                    LRUj = j
                    LRUtime = .SurfaceEntry(j).UltimoAcceso

                End If

            Next j

        End With

    Next i
    
    'Retrieve the surface desc
    Call TexList(LRUi).SurfaceEntry(LRUj).Texture.GetLevelDesc(0, surface_desc)
    
    'Remove it
    Set TexList(LRUi).SurfaceEntry(LRUj).Texture = Nothing
    TexList(LRUi).SurfaceEntry(LRUj).FileName = 0
    
    'Move back the list (if necessary)
    If LRUj Then
        RemoveLRU = True
        
        With TexList(LRUi)

            For j = LRUj To .surfaceCount - 1
                .SurfaceEntry(j) = .SurfaceEntry(j + 1)
            Next j
            
            .surfaceCount = .surfaceCount - 1

            If .surfaceCount Then
                ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
            Else
                Erase .SurfaceEntry

            End If

        End With

    End If
    
    'Update the used bytes
    mFreeMemoryBytes = mFreeMemoryBytes + surface_desc.size

    
    Exit Function

RemoveLRU_Err:
    Call RegistrarError(Err.number, Err.Description, "clsTexManager.RemoveLRU", Erl)
    Resume Next
    
End Function


