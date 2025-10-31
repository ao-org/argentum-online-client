VERSION 5.00
Begin VB.Form frmShopAO20 
   BorderStyle     =   0  'None
   Caption         =   "Tienda AO20"
   ClientHeight    =   7125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   475
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   432
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstItemShopFilter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H80000000&
      Height          =   3150
      Left            =   360
      TabIndex        =   5
      Top             =   2400
      Width           =   3255
   End
   Begin VB.ListBox lstItemsShop 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H80000000&
      Height          =   3150
      Left            =   360
      TabIndex        =   3
      Top             =   2400
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.PictureBox PictureItemShop 
      Appearance      =   0  'Flat
      BackColor       =   &H80000001&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   4770
      ScaleHeight     =   31
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   31
      TabIndex        =   2
      Top             =   2880
      Width           =   495
   End

   Begin VB.PictureBox picUserPreview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000001&
      ForeColor       =   &H80000008&
      Height          =   1905
      Left            =   4095
      ScaleHeight     =   127
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   124
      TabIndex        =   6
      Top             =   3600
      Width           =   1860
   End
   Begin VB.TextBox txtFindObj 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000000&
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   1620
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Una vez realizada la transacción, reloguee su personaje por seguridad"
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   4095
      TabIndex        =   4
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   6000
      Top             =   0
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   3600
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   1080
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   3240
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblCredits 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   240
      Left            =   5460
      TabIndex        =   1
      Top             =   1560
      Width           =   165
   End
End
Attribute VB_Name = "frmShopAO20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

'Vista previa completa con control de overrides y registro de depuración.
Private mPreviewHeading As E_Heading
Private mPreviewBodyOverride As Long
Private mPreviewHelmetObjNum As Long
Private mLastDrawnBody As Long
Private mLastDrawnHead As Long

Private Const cHelmetXOffset As Long = -2

Private Function NormalizeHeading(ByVal value As E_Heading) As E_Heading
    If value < E_Heading.NORTH Then
        NormalizeHeading = E_Heading.WEST
    ElseIf value > E_Heading.WEST Then
        NormalizeHeading = E_Heading.NORTH
    Else
        NormalizeHeading = value
    End If
End Function

Private Sub Form_Load()
    Me.Picture = LoadInterface("ventanatiendaao20.bmp")
    Me.KeyPreview = True
    mPreviewHeading = E_Heading.south
    mPreviewBodyOverride = 0
    mPreviewHelmetObjNum = 0
    mLastDrawnBody = 0
    mLastDrawnHead = 0
    Label1.Caption = JsonLanguage.Item("MENSAJE_TRANSACCION_RELOGUEO")
    Call RefreshSelectedItemPreview
End Sub

Private Sub Image2_Click()
    Unload Me
End Sub

Private Sub Image3_Click()
    'Antes de enviar al servidor hago una pre consulta de los créditos en cliente
    Dim obj_to_buy As ObjDatas
    If Not GetSelectedShopObject(obj_to_buy) Then Exit Sub
    If credits_shopAO20 >= obj_to_buy.Valor Then
        Call writeBuyShopItem(obj_to_buy.ObjNum)
    Else
        Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_NO_TIENE_CREDITOS_SUFICIENTES"), 255, 0, 0, True)
    End If
End Sub

Private Sub Image4_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Form_KeyDown_Err

    Dim previousHeading As E_Heading
    previousHeading = mPreviewHeading

    Select Case KeyCode
        Case vbKeyLeft
            mPreviewHeading = NormalizeHeading(mPreviewHeading - 1)
        Case vbKeyRight
            mPreviewHeading = NormalizeHeading(mPreviewHeading + 1)
        Case vbKeyUp
            mPreviewHeading = E_Heading.NORTH
        Case vbKeyDown
            mPreviewHeading = E_Heading.south
        Case Else
            Exit Sub
    End Select

    Debug.Print "[ShopPreview] Form_KeyDown - Heading de " & previousHeading & " a " & mPreviewHeading

    Call DrawUserPreview
    Exit Sub

Form_KeyDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmShopAO20.Form_KeyDown", Erl)
End Sub

Private Sub lstItemShopFilter_Click()
    Debug.Print "[ShopPreview] lstItemShopFilter_Click - index=" & lstItemShopFilter.ListIndex
    Call RefreshSelectedItemPreview
End Sub
Private Sub lstItemShopFilter_KeyUp(KeyCode As Integer, Shift As Integer)
    Debug.Print "[ShopPreview] lstItemShopFilter_KeyUp - index=" & lstItemShopFilter.ListIndex
    Call RefreshSelectedItemPreview
End Sub

Private Sub DrawUserPreview()
    On Error GoTo DrawUserPreview_Err

    Dim bodyIndex As Long
    Dim headIndex As Long
    Call GetUserAppearance(bodyIndex, headIndex)

    If mPreviewBodyOverride > 0 Then
        Debug.Print "[ShopPreview] DrawUserPreview - override body=" & mPreviewBodyOverride
        bodyIndex = mPreviewBodyOverride
    End If

    If bodyIndex <= 0 Then
        Debug.Print "[ShopPreview] DrawUserPreview - cuerpo inválido, se evita limpiar"
        Exit Sub
    End If

    Dim heading As E_Heading
    heading = NormalizeHeading(mPreviewHeading)

    Dim bodyFrame As Long
    bodyFrame = ResolveBodyFrame(bodyIndex, heading)
    If bodyFrame <= 0 Then
        Debug.Print "[ShopPreview] DrawUserPreview - sin frame de cuerpo válido"
        Exit Sub
    End If

    Call picUserPreview.Cls
    Debug.Print "[ShopPreview] DrawUserPreview - limpiando picture"

    Dim bodyWidth As Long
    Dim bodyHeight As Long
    bodyWidth = GrhData(bodyFrame).pixelWidth
    bodyHeight = GrhData(bodyFrame).pixelHeight

    Dim bodyX As Long
    Dim bodyY As Long
    bodyX = (picUserPreview.ScaleWidth - bodyWidth) \ 2
    bodyY = min(picUserPreview.ScaleHeight - bodyHeight + BodyData(bodyIndex).HeadOffset.y \ 2, _
                (picUserPreview.ScaleHeight - bodyHeight) \ 2)

    Call Grh_Render_To_Hdc(picUserPreview, bodyFrame, bodyX, bodyY, False, RGB(11, 11, 11))

    Dim headFrame As Long
    Dim headX As Long
    Dim headY As Long
    headFrame = ResolveHeadFrame(headIndex, heading)
    If headFrame > 0 Then
        headX = (picUserPreview.ScaleWidth - GrhData(headFrame).pixelWidth) \ 2 + 1
        headY = bodyY + bodyHeight - GrhData(headFrame).pixelHeight + BodyData(bodyIndex).HeadOffset.y
        Call Grh_Render_To_HdcSinBorrar(picUserPreview, headFrame, headX, headY, False)
    Else
        Debug.Print "[ShopPreview] DrawUserPreview - sin cabeza válida"
    End If

    If mPreviewHelmetObjNum > 0 Then
        Dim helmetFrame As Long
        helmetFrame = ResolveHelmetFrame(mPreviewHelmetObjNum, heading)
        If helmetFrame > 0 Then
            Dim helmetWidth As Long
            Dim helmetHeight As Long
            helmetWidth = GrhData(helmetFrame).pixelWidth
            helmetHeight = GrhData(helmetFrame).pixelHeight

            Dim helmetX As Long
            Dim helmetY As Long
            helmetX = ((picUserPreview.ScaleWidth - helmetWidth) \ 2) + cHelmetXOffset
            helmetY = bodyY + bodyHeight - helmetHeight + BodyData(bodyIndex).HeadOffset.y
            Call Grh_Render_To_HdcSinBorrar(picUserPreview, helmetFrame, helmetX, helmetY, False)
            Debug.Print "[ShopPreview] DrawUserPreview - casco dibujado obj=" & mPreviewHelmetObjNum & _
                        " frame=" & helmetFrame & " en (" & helmetX & "," & helmetY & ")"
        Else
            Debug.Print "[ShopPreview] DrawUserPreview - casco sin frame válido (obj=" & mPreviewHelmetObjNum & ")"
        End If
    End If

    mLastDrawnBody = bodyIndex
    mLastDrawnHead = headIndex

    Debug.Print "[ShopPreview] DrawUserPreview - cuerpo=" & bodyIndex & " cabeza=" & headIndex & _
                " heading=" & heading
    Call picUserPreview.Refresh
    Exit Sub

DrawUserPreview_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmShopAO20.DrawUserPreview", Erl)
End Sub

Private Function ResolveBodyFrame(ByVal bodyIndex As Long, ByVal heading As E_Heading) As Long
    On Error GoTo ResolveBodyFrame_Err

    Dim lower As Long
    Dim upper As Long
    On Error Resume Next
    lower = LBound(BodyData)
    upper = UBound(BodyData)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    On Error GoTo ResolveBodyFrame_Err

    If bodyIndex < lower Or bodyIndex > upper Then Exit Function

    Dim grhIndex As Long
    grhIndex = BodyData(bodyIndex).Walk(heading).GrhIndex
    ResolveBodyFrame = GetGrhFrame(grhIndex, 1)
    Exit Function

ResolveBodyFrame_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmShopAO20.ResolveBodyFrame", Erl)
    ResolveBodyFrame = 0
End Function

Private Function ResolveHeadFrame(ByVal headIndex As Long, ByVal heading As E_Heading) As Long
    On Error GoTo ResolveHeadFrame_Err

    If headIndex <= 0 Then Exit Function

    Dim lower As Long
    Dim upper As Long
    On Error Resume Next
    lower = LBound(HeadData)
    upper = UBound(HeadData)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    On Error GoTo ResolveHeadFrame_Err

    If headIndex < lower Or headIndex > upper Then Exit Function

    Dim grhIndex As Long
    grhIndex = HeadData(headIndex).Head(heading).GrhIndex
    ResolveHeadFrame = GetGrhFrame(grhIndex, 1)
    Exit Function

ResolveHeadFrame_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmShopAO20.ResolveHeadFrame", Erl)
    ResolveHeadFrame = 0
End Function

Private Sub RefreshSelectedItemPreview()
    On Error GoTo RefreshSelectedItemPreview_Err

    mPreviewBodyOverride = 0
    mPreviewHelmetObjNum = 0

    Dim shopIndex As Long
    shopIndex = SelectedShopIndex()
    If shopIndex = 0 Then
        Call PictureItemShop.Cls
        Debug.Print "[ShopPreview] RefreshSelectedItemPreview - sin selección"
        Call DrawUserPreview
        Exit Sub
    End If

    Dim objNum As Long
    objNum = ObjShop(shopIndex).ObjNum
    If objNum < LBound(ObjData) Or objNum > UBound(ObjData) Then
        Call PictureItemShop.Cls
        Debug.Print "[ShopPreview] RefreshSelectedItemPreview - ObjNum inválido: " & objNum
        Call DrawUserPreview
        Exit Sub
    End If

    Dim itemFrame As Long
    itemFrame = GetGrhFrame(ObjData(objNum).GrhIndex, 1)
    If itemFrame > 0 Then
        Call Grh_Render_To_Hdc(PictureItemShop, itemFrame, 0, 0, False)
        Debug.Print "[ShopPreview] RefreshSelectedItemPreview - itemFrame=" & itemFrame
    Else
        Call PictureItemShop.Cls
        Debug.Print "[ShopPreview] RefreshSelectedItemPreview - sin frame para ObjNum=" & objNum
    End If

    Dim objType As eObjType
    objType = ObjData(objNum).ObjType

    Dim resolvedBody As Long
    resolvedBody = ResolvePreviewBody(objNum)

    Select Case objType
        Case eObjType.otArmadura, eObjType.otSkinsArmours, eObjType.otMonturas
            If resolvedBody > 0 Then
                mPreviewBodyOverride = resolvedBody
                Debug.Print "[ShopPreview] RefreshSelectedItemPreview - override cuerpo=" & resolvedBody
            Else
                Debug.Print "[ShopPreview] RefreshSelectedItemPreview - override cuerpo inválido para obj=" & objNum
            End If
    End Select

    If objType = eObjType.otCASCO Or objType = eObjType.otSkinsHelmets Then
        mPreviewHelmetObjNum = objNum
        Debug.Print "[ShopPreview] RefreshSelectedItemPreview - casco seleccionado obj=" & objNum
    End If

    Debug.Print "[ShopPreview] RefreshSelectedItemPreview - objNum=" & objNum & " tipo=" & objType
    Call DrawUserPreview
    Exit Sub

RefreshSelectedItemPreview_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmShopAO20.RefreshSelectedItemPreview", Erl)
End Sub

Private Function SelectedShopIndex() As Long
    On Error GoTo SelectedShopIndex_Err

    Dim listIndex As Long
    listIndex = lstItemShopFilter.ListIndex
    If listIndex < 0 Then Exit Function

    Dim dataIndex As Long
    If listIndex >= 0 And listIndex < lstItemShopFilter.ListCount Then
        dataIndex = lstItemShopFilter.ItemData(listIndex)
    End If

    If dataIndex <= 0 Then
        dataIndex = listIndex + 1
    End If

    Dim upper As Long
    On Error Resume Next
    upper = UBound(ObjShop)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    On Error GoTo SelectedShopIndex_Err

    If dataIndex < 1 Or dataIndex > upper Then Exit Function

    SelectedShopIndex = dataIndex
    Exit Function

SelectedShopIndex_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmShopAO20.SelectedShopIndex", Erl)
End Function

Private Function GetSelectedShopObject(ByRef obj_to_buy As ObjDatas) As Boolean
    On Error GoTo GetSelectedShopObject_Err

    Dim shopIndex As Long
    shopIndex = SelectedShopIndex()
    If shopIndex = 0 Then Exit Function

    Dim objNum As Long
    objNum = ObjShop(shopIndex).ObjNum
    If objNum < LBound(ObjData) Or objNum > UBound(ObjData) Then Exit Function

    obj_to_buy = ObjData(objNum)
    obj_to_buy.ObjNum = objNum
    obj_to_buy.Valor = ObjShop(shopIndex).Valor
    obj_to_buy.Name = ObjShop(shopIndex).Name
    GetSelectedShopObject = True
    Exit Function

GetSelectedShopObject_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmShopAO20.GetSelectedShopObject", Erl)
End Function

Private Sub GetUserAppearance(ByRef bodyIndex As Long, ByRef headIndex As Long)
    On Error GoTo GetUserAppearance_Err

    bodyIndex = UserBody
    headIndex = UserHead

    Dim charLower As Long
    Dim charUpper As Long
    On Error Resume Next
    charLower = LBound(charlist)
    charUpper = UBound(charlist)
    If Err.Number = 0 Then
        If UserCharIndex >= charLower And UserCharIndex <= charUpper Then
            With charlist(UserCharIndex)
                If .iBody > 0 Then
                    bodyIndex = .iBody
                ElseIf .Body.BodyIndex > 0 Then
                    bodyIndex = .Body.BodyIndex
                End If
                If .IHead > 0 Then headIndex = .IHead
            End With
        End If
    Else
        Err.Clear
    End If
    On Error GoTo GetUserAppearance_Err

    Dim bodyLower As Long
    Dim bodyUpper As Long
    On Error Resume Next
    bodyLower = LBound(BodyData)
    bodyUpper = UBound(BodyData)
    If Err.Number <> 0 Then
        Err.Clear
        bodyIndex = 0
    ElseIf bodyIndex < bodyLower Or bodyIndex > bodyUpper Then
        bodyIndex = 0
    End If

    Dim headLower As Long
    Dim headUpper As Long
    On Error Resume Next
    headLower = LBound(HeadData)
    headUpper = UBound(HeadData)
    If Err.Number <> 0 Then
        Err.Clear
        headIndex = 0
    ElseIf headIndex < headLower Or headIndex > headUpper Then
        headIndex = 0
    End If
    On Error GoTo GetUserAppearance_Err

    Exit Sub

GetUserAppearance_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmShopAO20.GetUserAppearance", Erl)
    bodyIndex = 0
    headIndex = 0
End Sub

Private Function ResolvePreviewBody(ByVal objNum As Long) As Long
    On Error GoTo ResolvePreviewBody_Err

    If objNum < LBound(ObjData) Or objNum > UBound(ObjData) Then Exit Function

    Dim obj As ObjDatas
    obj = ObjData(objNum)

    Dim candidate As Long
    Select Case obj.ObjType
        Case eObjType.otArmadura, eObjType.otSkinsArmours, eObjType.otMonturas
            candidate = ResolveBodyByRace(obj)
            If candidate = 0 Then candidate = ResolveBodyByIndexValue(obj.GrhIndex)
        Case Else
            candidate = ResolveBodyByIndexValue(obj.GrhIndex)
    End Select

    If candidate > 0 Then
        Dim lower As Long
        Dim upper As Long
        On Error Resume Next
        lower = LBound(BodyData)
        upper = UBound(BodyData)
        If Err.Number <> 0 Then
            Err.Clear
            candidate = 0
        ElseIf candidate < lower Or candidate > upper Then
            candidate = 0
        End If
    End If
    On Error GoTo ResolvePreviewBody_Err

    ResolvePreviewBody = candidate
    Exit Function

ResolvePreviewBody_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmShopAO20.ResolvePreviewBody", Erl)
    ResolvePreviewBody = 0
End Function

Private Function ResolveBodyByRace(ByRef obj As ObjDatas) As Long
    On Error GoTo ResolveBodyByRace_Err

    Dim race As Long
    race = UserStats.Raza
    Dim gender As Long
    gender = UserStats.Sexo

    If gender <> eGenero.Mujer Then gender = eGenero.Hombre

    Select Case race
        Case eRaza.Humano
            If gender = eGenero.Mujer Then
                ResolveBodyByRace = obj.RopajeHumana
            Else
                ResolveBodyByRace = obj.RopajeHumano
            End If
        Case eRaza.Elfo
            If gender = eGenero.Mujer Then
                ResolveBodyByRace = obj.RopajeElfa
            Else
                ResolveBodyByRace = obj.RopajeElfo
            End If
        Case eRaza.ElfoOscuro
            If gender = eGenero.Mujer Then
                ResolveBodyByRace = obj.RopajeElfaOscura
            Else
                ResolveBodyByRace = obj.RopajeElfoOscuro
            End If
        Case eRaza.Gnomo
            If gender = eGenero.Mujer Then
                ResolveBodyByRace = obj.RopajeGnoma
            Else
                ResolveBodyByRace = obj.RopajeGnomo
            End If
        Case eRaza.Enano
            If gender = eGenero.Mujer Then
                ResolveBodyByRace = obj.RopajeEnana
            Else
                ResolveBodyByRace = obj.RopajeEnano
            End If
        Case eRaza.Orco
            If gender = eGenero.Mujer Then
                ResolveBodyByRace = obj.RopajeOrca
            Else
                ResolveBodyByRace = obj.RopajeOrco
            End If
    End Select
    Exit Function

ResolveBodyByRace_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmShopAO20.ResolveBodyByRace", Erl)
    ResolveBodyByRace = 0
End Function

Private Function ResolveBodyByIndexValue(ByVal value As Long) As Long
    On Error GoTo ResolveBodyByIndexValue_Err

    If value <= 0 Then Exit Function

    Dim lower As Long
    Dim upper As Long
    On Error Resume Next
    lower = LBound(BodyData)
    upper = UBound(BodyData)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    On Error GoTo ResolveBodyByIndexValue_Err

    If value >= lower And value <= upper Then
        ResolveBodyByIndexValue = value
        Exit Function
    End If

    If value <= UBound(GrhData) Then
        Dim i As Long
        For i = lower To upper
            Dim walkGrh As Long
            walkGrh = BodyData(i).Walk(E_Heading.south).GrhIndex
            If walkGrh = value Then
                ResolveBodyByIndexValue = i
                Exit Function
            End If
            If walkGrh > 0 And walkGrh <= UBound(GrhData) Then
                Dim frame As Long
                frame = GetGrhFrame(walkGrh, 1)
                If frame = value Then
                    ResolveBodyByIndexValue = i
                    Exit Function
                End If
            End If
        Next i
    End If
    Exit Function

ResolveBodyByIndexValue_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmShopAO20.ResolveBodyByIndexValue", Erl)
    ResolveBodyByIndexValue = 0
End Function

Private Function GetGrhFrame(ByVal baseGrh As Long, Optional ByVal frameIndex As Long = 1) As Long
    On Error GoTo GetGrhFrame_Err

    If baseGrh <= 0 Or baseGrh > UBound(GrhData) Then Exit Function

    Dim numFrames As Long
    numFrames = GrhData(baseGrh).NumFrames

    If numFrames <= 1 Then
        GetGrhFrame = baseGrh
    Else
        If frameIndex <= 0 Or frameIndex > numFrames Then frameIndex = 1
        GetGrhFrame = GrhData(baseGrh).Frames(frameIndex)
        If GetGrhFrame = 0 Then
            GetGrhFrame = GrhData(baseGrh).Frames(1)
        End If
    End If
    Exit Function

GetGrhFrame_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmShopAO20.GetGrhFrame", Erl)
    GetGrhFrame = 0
End Function

Private Function ResolveHelmetFrame(ByVal objNum As Long, ByVal heading As E_Heading) As Long
    On Error GoTo ResolveHelmetFrame_Err

    If objNum < LBound(ObjData) Or objNum > UBound(ObjData) Then Exit Function

    Dim baseGrh As Long
    baseGrh = ObjData(objNum).GrhIndex
    If baseGrh <= 0 Then Exit Function

    If heading < E_Heading.NORTH Or heading > E_Heading.WEST Then
        heading = E_Heading.south
    End If

    Dim cascoLower As Long
    Dim cascoUpper As Long
    On Error Resume Next
    cascoLower = LBound(CascoAnimData)
    cascoUpper = UBound(CascoAnimData)
    If Err.Number = 0 Then
        If baseGrh >= cascoLower And baseGrh <= cascoUpper Then
            Dim cascoGrh As Long
            cascoGrh = CascoAnimData(baseGrh).Head(heading).GrhIndex
            If cascoGrh > 0 Then
                ResolveHelmetFrame = GetGrhFrame(cascoGrh, 1)
                Exit Function
            End If
        End If
    Else
        Err.Clear
    End If
    On Error GoTo ResolveHelmetFrame_Err

    ResolveHelmetFrame = GetGrhFrame(baseGrh, heading)
    Exit Function

ResolveHelmetFrame_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmShopAO20.ResolveHelmetFrame", Erl)
    ResolveHelmetFrame = 0
End Function

Private Sub txtFindObj_Change()
    lstItemShopFilter.Clear
    Dim i As Long
    Dim displayText As String
    Dim addedIndex As Long
    For i = 1 To UBound(ObjShop)
        If InStr(1, ObjShop(i).Name, txtFindObj.text, 1) > 0 Then
            displayText = ObjShop(i).Name & " ( " & JsonLanguage.Item("MENSAJE_VALOR") & ObjShop(i).Valor & " )"
            lstItemShopFilter.AddItem displayText
            addedIndex = lstItemShopFilter.NewIndex
            If addedIndex >= 0 Then
                lstItemShopFilter.ItemData(addedIndex) = i
            End If
        End If
    Next i
    Call RefreshSelectedItemPreview
End Sub
