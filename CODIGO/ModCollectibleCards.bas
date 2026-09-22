Attribute VB_Name = "ModCollectibleCards"
Option Explicit

Public Sub DrawCollectibleCard()
    Call CollectibleCardRender(frmCollectibleCard.picCollectibleCard, ObjData(G_LasSelectedObjIndex).CollectibleCardImgPathing, 0, 0, 439, 600, 0, 0, 439, 600)
End Sub

Public Sub CollectibleCardRender(ByRef pic As PictureBox, _
                                         ByVal TextureFileName As String, _
                                         ByVal DestX As Long, _
                                         ByVal DestY As Long, _
                                         ByVal destWidth As Long, _
                                         ByVal destHeight As Long, _
                                         ByVal srcX As Long, _
                                         ByVal srcY As Long, _
                                         ByVal srcWidth As Long, _
                                         ByVal srcHeight As Long, _
                                         Optional ByVal ClearColor As Long = &H0)
    On Error GoTo CollectibleCardRender_Err
    
    If LenB(Trim$(TextureFileName)) = 0 Then
        frmDebug.add_text_tracebox "Card texture path is empty"
        Exit Sub
    End If

    Dim Texture As Direct3DTexture8
    Dim texwidth As Long
    Dim texheight As Long
    
    Set Texture = SurfaceDB.GetInterfaceTexture(TextureFileName, texwidth, texheight)
    If Texture Is Nothing Then
        Debug.Print "Failed to load card texture: " & TextureFileName
        frmDebug.add_text_tracebox "Failed to load card texture: " & TextureFileName
        Exit Sub
    End If

    Dim DestRect As RECT
    
    With DestRect
        .Left = 0
        .Top = 0
        .Right = destWidth
        .Bottom = destHeight
    End With
    
    Call Engine_BeginScene
    
    Call Batch_Textured_Box_File(DestX, DestY, destWidth, destHeight, srcX, srcY, TextureFileName, COLOR_WHITE, False, 0, _
                                 CSng(destWidth) / CSng(srcWidth), CSng(destHeight) / CSng(srcHeight))
    
    Call Engine_EndScene(DestRect, pic.hWnd)
    

    Exit Sub

CollectibleCardRender_Err:
    Call RegistrarError(Err.Number, Err.Description, "TileEngine.CollectibleCardRender", Erl)
    frmDebug.add_text_tracebox "Error in CollectibleCardRender: " & Err.Description
    Resume Next
End Sub


