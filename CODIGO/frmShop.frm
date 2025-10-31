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

'Vista previa simplificada: deshabilitamos la lógica adicional de overrides.
Private mPreviewHeading As E_Heading
'Private mPreviewBodyOverride As Long
'Private mPreviewHelmetObjNum As Long
'Private mLastDrawnBody As Long
'Private mLastDrawnHead As Long

Private Sub Form_Load()
    Me.Picture = LoadInterface("ventanatiendaao20.bmp")
    'Vista previa simplificada: se deshabilita la rotación y overrides.
    mPreviewHeading = E_Heading.south
    'mPreviewBodyOverride = 0
    'mPreviewHelmetObjNum = 0
    'mLastDrawnBody = 0
    'mLastDrawnHead = 0
    Label1.Caption = JsonLanguage.Item("MENSAJE_TRANSACCION_RELOGUEO")
    Call DrawUserPreview
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
    'Vista previa simplificada: sin rotación mediante teclado.
End Sub
Private Sub lstItemShopFilter_Click()
    Call DrawUserPreview
End Sub
Private Sub lstItemShopFilter_KeyUp(KeyCode As Integer, Shift As Integer)
    Call DrawUserPreview
End Sub
Private Sub DrawUserPreview()
    On Error GoTo DrawUserPreview_Err

    Call picUserPreview.Cls

    Dim bodyIndex As Long
    Dim headIndex As Long
    Call GetUserAppearance(bodyIndex, headIndex)

    If bodyIndex <= 0 Then Exit Sub

    Call DibujarNPC(picUserPreview, headIndex, bodyIndex, mPreviewHeading)
    Call picUserPreview.Refresh
    Exit Sub

DrawUserPreview_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmShopAO20.DrawUserPreview", Erl)
End Sub
Private Sub RefreshSelectedItemPreview()
    'Vista previa simplificada: solo redibuja al usuario.
    Call DrawUserPreview
End Sub
#If False Then
Private Sub RefreshSelectedItemPreview()
    On Error GoTo RefreshSelectedItemPreview_Err

    mPreviewBodyOverride = 0
    mPreviewHelmetObjNum = 0

    Dim shopIndex As Long
    shopIndex = SelectedShopIndex()
    If shopIndex = 0 Then
        Call PictureItemShop.Cls
        Call DrawUserPreview
        Exit Sub
    End If

    Dim objNum As Long
    objNum = ObjShop(shopIndex).ObjNum
    If objNum < LBound(ObjData) Or objNum > UBound(ObjData) Then
        Call PictureItemShop.Cls
        Call DrawUserPreview
        Exit Sub
    End If

    Dim itemFrame As Long
    itemFrame = GetGrhFrame(ObjData(objNum).GrhIndex, 1)
    If itemFrame > 0 Then
        Call Grh_Render_To_Hdc(PictureItemShop, itemFrame, 0, 0, False)
    Else
        Call PictureItemShop.Cls
    End If

    Dim objType As eObjType
    objType = ObjData(objNum).ObjType

    Dim resolvedBody As Long
    resolvedBody = ResolvePreviewBody(objNum)

    Select Case objType
        Case eObjType.otArmadura, eObjType.otSkinsArmours, eObjType.otMonturas
            If resolvedBody > 0 Then mPreviewBodyOverride = resolvedBody
    End Select

    If objType = eObjType.otCASCO Or objType = eObjType.otSkinsHelmets Then
        mPreviewHelmetObjNum = objNum
    End If

    Call DrawUserPreview
    Exit Sub

RefreshSelectedItemPreview_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmShopAO20.RefreshSelectedItemPreview", Erl)
End Sub

#End If
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
