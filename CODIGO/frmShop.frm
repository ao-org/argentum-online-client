VERSION 5.00
Begin VB.Form frmShopAO20 
   BorderStyle     =   0  'None
   Caption         =   "Tienda AO20"
   ClientHeight    =   7125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   KeyPreview      =   -1  'True
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
Private mPreviewHeading As E_Heading
Private mPreviewOverrideBodyIndex As Integer
Private mPreviewHelmetGrhIndex As Long
Private mPreviewMountGrhIndex As Long

Private Sub Form_Load()
    Me.Picture = LoadInterface("ventanatiendaao20.bmp")
    Label1.Caption = JsonLanguage.Item("MENSAJE_TRANSACCION_RELOGUEO")
    mPreviewHeading = E_Heading.south
    mPreviewOverrideBodyIndex = 0
    mPreviewHelmetGrhIndex = 0
    mPreviewMountGrhIndex = 0
    Call DrawUserPreview
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo Form_KeyDown_Err

    Dim newHeading As E_Heading
    newHeading = mPreviewHeading

    Select Case KeyCode
        Case vbKeyUp
            newHeading = E_Heading.NORTH
        Case vbKeyRight
            newHeading = E_Heading.EAST
        Case vbKeyDown
            newHeading = E_Heading.south
        Case vbKeyLeft
            newHeading = E_Heading.WEST
        Case Else
            Exit Sub
    End Select

    If newHeading <> mPreviewHeading Then
        mPreviewHeading = newHeading
        Call DrawUserPreview
    End If
    Exit Sub

Form_KeyDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmShopAO20.Form_KeyDown", Erl)
    Resume Next
End Sub

Private Sub Image2_Click()
    Unload Me
End Sub

Private Sub Image3_Click()
    'Antes de enviar al servidor hago una pre consulta de los créditos en cliente
    Dim obj_to_buy As ObjDatas
    Dim i          As Long
    obj_to_buy = ObjShop(Me.lstItemShopFilter.ListIndex + 1)
    Dim obj_name As String
    obj_name = Split(lstItemShopFilter.text, " (")(0)
    For i = 1 To UBound(ObjShop)
        If obj_name = ObjShop(i).Name Then
            obj_to_buy = ObjData(ObjShop(i).ObjNum)
            obj_to_buy.ObjNum = ObjShop(i).ObjNum
            obj_to_buy.Valor = ObjShop(i).Valor
            Exit For
        End If
    Next i
    If credits_shopAO20 >= obj_to_buy.Valor Then
        Call writeBuyShopItem(obj_to_buy.ObjNum)
    Else
        Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("MENSAJE_NO_TIENE_CREDITOS_SUFICIENTES"), 255, 0, 0, True)
    End If
End Sub

Private Sub Image4_Click()
    Unload Me
End Sub

Private Sub lstItemShopFilter_Click()
    Dim Grh      As Long
    Dim i        As Long
    Dim obj_name As String
    Dim objNum   As Long
    Dim objType  As Integer
    obj_name = Split(lstItemShopFilter.text, " (")(0)
    For i = 1 To UBound(ObjShop)
        If obj_name = ObjShop(i).Name Then
            objNum = ObjShop(i).ObjNum
            Exit For
        End If
    Next i

    If objNum > 0 And objNum <= UBound(ObjData) Then
        Grh = ObjData(objNum).GrhIndex
        objType = ObjData(objNum).ObjType
    Else
        objNum = 0
        Grh = 0
        objType = 0
    End If

    Call Grh_Render_To_Hdc(PictureItemShop, Grh, 0, 0, False)
    Call UpdateUserPreviewSelection(objNum, objType)
End Sub

Private Sub txtFindObj_Change()
    lstItemShopFilter.Clear
    Dim i As Long
    For i = 1 To UBound(ObjShop)
        If InStr(1, ObjShop(i).Name, txtFindObj.text, 1) > 0 Then
            Call frmShopAO20.lstItemShopFilter.AddItem(ObjShop(i).Name & " ( " & JsonLanguage.Item("MENSAJE_VALOR") & ObjShop(i).Valor & " )")
        End If
    Next i
End Sub

Private Sub DrawUserPreview()
    On Error GoTo DrawUserPreview_Err

    If UserCharIndex <= 0 Then Exit Sub
    If picUserPreview.Width = 0 Or picUserPreview.Height = 0 Then Exit Sub

    Dim bodyIndex As Integer
    Dim headIndex As Integer
    Dim heading As E_Heading

    bodyIndex = charlist(UserCharIndex).iBody
    headIndex = charlist(UserCharIndex).IHead

    If mPreviewOverrideBodyIndex > 0 Then
        If mPreviewOverrideBodyIndex >= LBound(BodyData) And mPreviewOverrideBodyIndex <= UBound(BodyData) Then
            bodyIndex = mPreviewOverrideBodyIndex
        End If
    End If

    If bodyIndex < LBound(BodyData) Or bodyIndex > UBound(BodyData) Then Exit Sub

    heading = mPreviewHeading
    If heading < E_Heading.NORTH Or heading > E_Heading.WEST Then heading = E_Heading.south

    Dim bodyGrh As Long
    Dim bodyFrame As Long
    Dim mountFrame As Long
    Dim mountDrawn As Boolean
    Dim backgroundColor As Long

    bodyGrh = BodyData(bodyIndex).Walk(heading).GrhIndex
    bodyFrame = ResolvePreviewGrhFrame(bodyGrh)
    backgroundColor = RGB(11, 11, 11)

    If mPreviewMountGrhIndex > 0 Then
        mountFrame = ResolvePreviewGrhFrame(mPreviewMountGrhIndex)
        If mountFrame > 0 Then
            Dim mountX As Integer
            Dim mountY As Integer
            mountX = (picUserPreview.ScaleWidth - GrhData(mountFrame).pixelWidth) \ 2
            mountY = picUserPreview.ScaleHeight - GrhData(mountFrame).pixelHeight
            Call Grh_Render_To_Hdc(picUserPreview, mountFrame, mountX, mountY, False, backgroundColor)
            mountDrawn = True
        End If
    End If

    Dim bodyX As Integer
    Dim bodyY As Integer
    If bodyFrame > 0 Then
        bodyX = (picUserPreview.ScaleWidth - GrhData(bodyFrame).pixelWidth) \ 2
        bodyY = min(picUserPreview.ScaleHeight - GrhData(bodyFrame).pixelHeight + BodyData(bodyIndex).HeadOffset.y \ 2, (picUserPreview.ScaleHeight - GrhData(bodyFrame).pixelHeight) \ 2)
        If mountDrawn Then
            Call Grh_Render_To_HdcSinBorrar(picUserPreview, bodyFrame, bodyX, bodyY, False)
        Else
            Call Grh_Render_To_Hdc(picUserPreview, bodyFrame, bodyX, bodyY, False, backgroundColor)
            mountDrawn = True
        End If
    ElseIf Not mountDrawn Then
        picUserPreview.Cls
    End If

    Dim headFrame As Long
    Dim headX As Integer
    Dim headY As Integer
    Dim headOffsetX As Integer
    Dim headDrawn As Boolean
    If headIndex >= LBound(HeadData) And headIndex <= UBound(HeadData) Then
        headFrame = ResolvePreviewGrhFrame(HeadData(headIndex).Head(heading).GrhIndex)
    End If

    If headFrame > 0 And bodyFrame > 0 Then
        If bodyIndex >= LBound(BodyData) And bodyIndex <= UBound(BodyData) Then
            headOffsetX = BodyData(bodyIndex).HeadOffset.x - BodyData(bodyIndex).BodyOffset.x
        Else
            headOffsetX = 0
        End If

        headX = bodyX + headOffsetX
        headY = bodyY + GrhData(bodyFrame).pixelHeight - GrhData(headFrame).pixelHeight + BodyData(bodyIndex).HeadOffset.y
        Call Grh_Render_To_HdcSinBorrar(picUserPreview, headFrame, headX, headY, False)
        headDrawn = True
    End If

    If mPreviewHelmetGrhIndex > 0 Then
        Dim helmetFrame As Long
        helmetFrame = ResolvePreviewGrhFrame(mPreviewHelmetGrhIndex)
        If helmetFrame > 0 Then
            Dim helmetX As Integer
            Dim helmetY As Integer
            If headDrawn Then
                helmetX = headX
                helmetY = headY
            Else
                If bodyFrame > 0 Then
                    If bodyIndex >= LBound(BodyData) And bodyIndex <= UBound(BodyData) Then
                        headOffsetX = BodyData(bodyIndex).HeadOffset.x - BodyData(bodyIndex).BodyOffset.x
                    Else
                        headOffsetX = 0
                    End If
                    helmetX = bodyX + headOffsetX
                Else
                    helmetX = (picUserPreview.ScaleWidth - GrhData(helmetFrame).pixelWidth) \ 2
                End If
                If bodyFrame > 0 Then
                    helmetY = bodyY + GrhData(bodyFrame).pixelHeight - GrhData(helmetFrame).pixelHeight + BodyData(bodyIndex).HeadOffset.y
                Else
                    helmetY = (picUserPreview.ScaleHeight - GrhData(helmetFrame).pixelHeight) \ 2
                End If
            End If
            helmetX = helmetX - 2
            Call Grh_Render_To_HdcSinBorrar(picUserPreview, helmetFrame, helmetX, helmetY, False)
        End If
    End If

    Exit Sub

DrawUserPreview_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmShopAO20.DrawUserPreview", Erl)
    Resume Next
End Sub

Private Sub UpdateUserPreviewSelection(ByVal objNum As Long, ByVal objType As Integer)
    On Error GoTo UpdateUserPreviewSelection_Err

    mPreviewOverrideBodyIndex = 0
    mPreviewHelmetGrhIndex = 0
    mPreviewMountGrhIndex = 0

    If objNum > 0 Then
        Select Case objType
            Case eObjType.otArmadura, eObjType.otSkinsArmours
                mPreviewOverrideBodyIndex = ResolvePreviewBodyIndex(objNum, objType)
            Case eObjType.otCASCO, eObjType.otSkinsHelmets
                mPreviewHelmetGrhIndex = ObjData(objNum).GrhIndex
            Case eObjType.otMonturas
                mPreviewMountGrhIndex = ResolvePreviewMountGrh(objNum)
        End Select
    End If

    Call DrawUserPreview
    Exit Sub

UpdateUserPreviewSelection_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmShopAO20.UpdateUserPreviewSelection", Erl)
    Resume Next
End Sub

Private Function ResolvePreviewBodyIndex(ByVal objNum As Long, ByVal objType As Integer) As Integer
    On Error GoTo ResolvePreviewBodyIndex_Err

    Dim candidate As Long
    Dim race As eRaza
    Dim gender As eGenero

    If objNum <= 0 Then GoTo ResolvePreviewBodyIndex_CleanExit
    If objNum > UBound(ObjData) Then GoTo ResolvePreviewBodyIndex_CleanExit

    race = UserStats.Raza
    gender = UserStats.Sexo

    If race >= eRaza.Humano And race <= eRaza.Orco And gender >= eGenero.Hombre And gender <= eGenero.Mujer Then
        candidate = ObjData(objNum).PreviewBody(race, gender)
        If candidate = 0 Then
            candidate = ShopPreview_GetBodyOverride(objNum, race, gender)
        End If
    End If

    If candidate = 0 Then
        candidate = ObjData(objNum).PreviewDefaultBody
        If candidate = 0 Then
            candidate = ShopPreview_GetDefaultBody(objNum)
        End If
    End If

    If candidate = 0 And objType <> eObjType.otSkinsArmours Then
        candidate = ObjData(objNum).GrhIndex
    End If

    If candidate >= LBound(BodyData) And candidate <= UBound(BodyData) Then
        ResolvePreviewBodyIndex = candidate
        Exit Function
    End If

ResolvePreviewBodyIndex_CleanExit:
    ResolvePreviewBodyIndex = 0
    Exit Function

ResolvePreviewBodyIndex_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmShopAO20.ResolvePreviewBodyIndex", Erl)
    ResolvePreviewBodyIndex = 0
End Function

Private Function ResolvePreviewMountGrh(ByVal objNum As Long) As Long
    On Error GoTo ResolvePreviewMountGrh_Err

    Dim candidate As Long

    If objNum <= 0 Then GoTo ResolvePreviewMountGrh_CleanExit

    candidate = ObjData(objNum).GrhIndex
    If candidate > 0 And candidate <= UBound(GrhData) Then
        ResolvePreviewMountGrh = candidate
        Exit Function
    End If

ResolvePreviewMountGrh_CleanExit:
    ResolvePreviewMountGrh = 0
    Exit Function

ResolvePreviewMountGrh_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmShopAO20.ResolvePreviewMountGrh", Erl)
    ResolvePreviewMountGrh = 0
End Function

Private Function ResolvePreviewGrhFrame(ByVal grhIndex As Long) As Long
    On Error GoTo ResolvePreviewGrhFrame_Err

    If grhIndex <= 0 Or grhIndex > UBound(GrhData) Then GoTo ResolvePreviewGrhFrame_CleanExit

    If GrhData(grhIndex).NumFrames > 0 Then
        ResolvePreviewGrhFrame = GrhData(grhIndex).Frames(1)
    Else
        ResolvePreviewGrhFrame = grhIndex
    End If
    Exit Function

ResolvePreviewGrhFrame_CleanExit:
    ResolvePreviewGrhFrame = 0
    Exit Function

ResolvePreviewGrhFrame_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmShopAO20.ResolvePreviewGrhFrame", Erl)
    ResolvePreviewGrhFrame = 0
End Function
