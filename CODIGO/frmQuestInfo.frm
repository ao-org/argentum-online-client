VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmQuestInfo 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6552
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12312
   LinkTopic       =   "Form1"
   ScaleHeight     =   6552
   ScaleWidth      =   12312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox Text1 
      Height          =   2895
      Left            =   3950
      TabIndex        =   10
      Top             =   2400
      Width           =   3130
      _ExtentX        =   5525
      _ExtentY        =   5101
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmQuestInfo.frx":0000
   End
   Begin VB.ListBox lstQuests 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2676
      Left            =   10560
      TabIndex        =   6
      Top             =   8040
      Width           =   2835
   End
   Begin VB.PictureBox picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   10580
      ScaleHeight     =   32
      ScaleMode       =   0  'User
      ScaleWidth      =   32
      TabIndex        =   4
      Top             =   1830
      Width           =   480
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1560
      Left            =   7320
      TabIndex        =   1
      Top             =   1680
      Width           =   2230
      _ExtentX        =   3937
      _ExtentY        =   2752
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   14737632
      BackColor       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Criatura"
         Object.Width           =   3106
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Cantidad"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Index"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Tipo"
         Object.Width           =   0
      EndProperty
      Picture         =   "frmQuestInfo.frx":0082
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   2280
      Left            =   9840
      TabIndex        =   3
      Top             =   3000
      Width           =   1965
      _ExtentX        =   3471
      _ExtentY        =   4022
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   14737632
      BackColor       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Criatura"
         Object.Width           =   2294
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Cantidad"
         Object.Width           =   1113
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Index"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Tipo"
         Object.Width           =   0
      EndProperty
      Picture         =   "frmQuestInfo.frx":0953
   End
   Begin MSComctlLib.ListView ListViewQuest 
      Height          =   2640
      Left            =   600
      TabIndex        =   7
      Top             =   1920
      Width           =   2835
      _ExtentX        =   4995
      _ExtentY        =   4657
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Quest"
         Object.Width           =   4588
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "estado"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "id"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.PictureBox PlayerView 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1185
      Left            =   7330
      ScaleHeight     =   79
      ScaleMode       =   0  'User
      ScaleWidth      =   146
      TabIndex        =   2
      Top             =   3830
      Width           =   2190
   End
   Begin VB.Label lblRepetible 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Misión repetible"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   2040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label npclbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   7320
      TabIndex        =   8
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   11760
      Top             =   0
      Width           =   495
   End
   Begin VB.Label objetolbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9840
      TabIndex        =   5
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   6400
      Tag             =   "0"
      Top             =   5730
      Width           =   1980
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   3880
      Tag             =   "0"
      Top             =   5730
      Width           =   1980
   End
   Begin VB.Label titulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mision"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000EA4EB&
      Height          =   255
      Left            =   3840
      TabIndex        =   0
      Top             =   1800
      Width           =   3375
   End
End
Attribute VB_Name = "FrmQuestInfo"
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
Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Me.Picture = LoadInterface("ventananuevamision.bmp")

    Text1.BackColor = RGB(11, 11, 11)
    PlayerView.BackColor = RGB(11, 11, 11)
    picture1.BackColor = RGB(19, 14, 11)
    Call Aplicar_Transparencia(Me.hWnd, 240)
    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmQuestInfo.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Form_KeyPress_Err
    

    If (KeyAscii = 27) Then
        Unload Me

    End If
    
    
    Exit Sub

Form_KeyPress_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmQuestInfo.Form_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    MoverForm Me.hWnd
    
    Image1.Picture = Nothing
    Image1.Tag = 0
    Image2.Picture = Nothing
    Image2.Tag = 0

    
    Exit Sub

Form_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmQuestInfo.Form_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image1_MouseMove_Err
    

    If Image1.Tag = "0" Then
        Image1.Picture = LoadInterface("boton-rechazar-over.bmp")
        Image1.Tag = "1"

    End If

    
    Exit Sub

Image1_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmQuestInfo.Image1_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image1_MouseUp_Err
    
    Unload Me

    Exit Sub

Image1_MouseUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmQuestInfo.Image1_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image2_MouseMove_Err
    

    If Image2.Tag = "0" Then
        Image2.Picture = LoadInterface("boton-aceptar-over.bmp")
        Image2.Tag = "1"

    End If

    
    Exit Sub

Image2_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmQuestInfo.Image2_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image2_MouseUp_Err
    


    If ListViewQuest.SelectedItem.Index > 0 Then
            Call WriteQuestAccept(ListViewQuest.SelectedItem.Index)
        Unload Me
    End If
    
    Exit Sub

Image2_MouseUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmQuestInfo.Image2_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub Image3_Click()
    
    On Error GoTo Image3_Click_Err
    Unload Me
    
    Exit Sub

Image3_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmQuestInfo.Image3_Click", Erl)
    Resume Next
    
End Sub

Public Sub ListView1_Click()
    
    On Error GoTo ListView1_Click_Err
    
    If ListView1.SelectedItem Is Nothing Then Exit Sub

    If ListView1.SelectedItem.SubItems(2) <> "" Then
        If ListView1.SelectedItem.SubItems(3) = 0 Then
            PlayerView.BackColor = RGB(11, 11, 11)
            Call DibujarNPC(PlayerView, NpcData(ListView1.SelectedItem.SubItems(2)).Head, NpcData(ListView1.SelectedItem.SubItems(2)).Body, 3)

            npclbl.Caption = NpcData(ListView1.SelectedItem.SubItems(2)).Name & " (" & ListView1.SelectedItem.SubItems(1) & ")"
    
        Else

            Dim x As Long

            Dim y As Long
        
            x = (PlayerView.ScaleWidth - GrhData(ObjData(ListView1.SelectedItem.SubItems(2)).GrhIndex).pixelWidth) / 2
            y = (PlayerView.ScaleHeight - GrhData(ObjData(ListView1.SelectedItem.SubItems(2)).GrhIndex).pixelHeight) / 2
            Call Grh_Render_To_Hdc(PlayerView, ObjData(ListView1.SelectedItem.SubItems(2)).GrhIndex, x, y, False, RGB(11, 11, 11))
        
            npclbl.Caption = ObjData(ListView1.SelectedItem.SubItems(2)).Name & " (" & ListView1.SelectedItem.SubItems(1) & ")"
    
        End If

    End If

    
    Exit Sub

ListView1_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmQuestInfo.ListView1_Click", Erl)
    Resume Next
    
End Sub

Public Sub ListView2_Click()
    
    On Error GoTo ListView2_Click_Err

    If ListView2.SelectedItem Is Nothing Then Exit Sub

    If ListView2.SelectedItem.SubItems(2) <> "" Then
 
        Call Grh_Render_To_Hdc(picture1, ObjData(ListView2.SelectedItem.SubItems(2)).GrhIndex, 0, 0, False, RGB(19, 14, 11))
    
    End If
    
    objetolbl.Caption = ObjData(ListView2.SelectedItem.SubItems(2)).Name & vbCrLf & " (" & ListView2.SelectedItem.SubItems(1) & ")"

    
    Exit Sub

ListView2_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmQuestInfo.ListView2_Click", Erl)
    Resume Next
    
End Sub

Public Sub ShowQuest(ByVal QuestIndex As Integer)

    Call ListViewQuest_ItemClick(ListViewQuest.ListItems(QuestIndex))

End Sub


Private Sub ListViewQuest_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    On Error GoTo ListViewQuest_ItemClick_Err
    
    If ListViewQuest.SelectedItem Is Nothing Then Exit Sub
    
    If Len(ListViewQuest.SelectedItem.SubItems(2)) <> 0 Then
        
        Dim QuestIndex As Byte

        QuestIndex = ListViewQuest.SelectedItem.SubItems(2)
        
        FrmQuestInfo.ListView2.ListItems.Clear
        FrmQuestInfo.ListView1.ListItems.Clear
            
        FrmQuestInfo.titulo.Caption = QuestList(QuestIndex).nombre
        
        FrmQuestInfo.Text1.Text = ""
        PlayerView.BackColor = RGB(11, 11, 11)
        picture1.BackColor = RGB(19, 14, 11)
        PlayerView.Refresh
        picture1.Refresh
        npclbl.Caption = ""
        objetolbl.Caption = ""
        
        lblRepetible.Visible = QuestList(QuestIndex).Repetible = 1
                
        
                'tmpStr = tmpStr & "Detalles: " & .ReadASCIIString & vbCrLf
                'tmpStr = tmpStr & "Nivel requerido: " & .ReadByte & vbCrLf
                
                
        
                  If QuestList(QuestIndex).RequiredQuest <> 0 Then
                    FrmQuestInfo.Text1.Text = ""
                    Call AddtoRichTextBox(Text1, QuestList(QuestIndex).desc & vbCrLf & vbCrLf & "Requisitos: " & vbCrLf & "Nivel requerido: " & QuestList(QuestIndex).RequiredLevel & vbCrLf & "Quest: " & QuestList(QuestList(QuestIndex).RequiredQuest).nombre, 128, 128, 128)
                Else
                    FrmQuestInfo.Text1.Text = ""
                    Call AddtoRichTextBox(Text1, QuestList(QuestIndex).desc & vbCrLf & vbCrLf & "Requisitos: " & vbCrLf & "Nivel requerido: " & QuestList(QuestIndex).RequiredLevel & vbCrLf, 128, 128, 128)
                
                End If
               
               

                
                
                If UBound(QuestList(QuestIndex).RequiredNPC) > 0 Then 'Hay NPCs
                    If UBound(QuestList(QuestIndex).RequiredNPC) > 5 Then
                        FrmQuestInfo.ListView1.FlatScrollBar = False
                    Else
                        FrmQuestInfo.ListView1.FlatScrollBar = True
                    End If
                    
                    
                    For i = 1 To UBound(QuestList(QuestIndex).RequiredNPC)
                                                

                            Dim subelemento As ListItem
    
                            Set subelemento = FrmQuestInfo.ListView1.ListItems.Add(, , NpcData(QuestList(QuestIndex).RequiredNPC(i).NpcIndex).Name)
                           
                            subelemento.SubItems(1) = QuestList(QuestIndex).RequiredNPC(i).Amount
                            subelemento.SubItems(2) = QuestList(QuestIndex).RequiredNPC(i).NpcIndex
                            subelemento.SubItems(3) = 0

    
                    Next i
    
                End If
                    
    
                If LBound(QuestList(QuestIndex).RequiredOBJ) > 0 Then  'Hay OBJs
    
                    For i = 1 To UBound(QuestList(QuestIndex).RequiredOBJ)
                        Set subelemento = FrmQuestInfo.ListView1.ListItems.Add(, , ObjData(QuestList(QuestIndex).RequiredOBJ(i).OBJIndex).Name)
                        subelemento.SubItems(1) = QuestList(QuestIndex).RequiredOBJ(i).Amount
                        subelemento.SubItems(2) = QuestList(QuestIndex).RequiredOBJ(i).OBJIndex
                        subelemento.SubItems(3) = 1
                    Next i
    
                End If
        
               
                If QuestList(QuestIndex).RewardGLD <> 0 Then
                     Set subelemento = FrmQuestInfo.ListView2.ListItems.Add(, , "Oro")
                     subelemento.SubItems(1) = BeautifyBigNumber(QuestList(QuestIndex).RewardGLD)
                     subelemento.SubItems(2) = 12
                     subelemento.SubItems(3) = 0
                End If
                
                
                If QuestList(QuestIndex).RewardEXP <> 0 Then
                    Set subelemento = FrmQuestInfo.ListView2.ListItems.Add(, , "Experiencia")
                    subelemento.SubItems(1) = BeautifyBigNumber(QuestList(QuestIndex).RewardEXP)
                    subelemento.SubItems(2) = 608
                    subelemento.SubItems(3) = 1
                End If
               

                If UBound(QuestList(QuestIndex).RewardOBJ) > 0 Then
                
                    
                    For i = 1 To UBound(QuestList(QuestIndex).RewardOBJ)

                                                                   
                        Set subelemento = FrmQuestInfo.ListView2.ListItems.Add(, , ObjData(QuestList(QuestIndex).RewardOBJ(i).OBJIndex).Name)
                           
                        subelemento.SubItems(1) = QuestList(QuestIndex).RewardOBJ(i).Amount
                        subelemento.SubItems(2) = QuestList(QuestIndex).RewardOBJ(i).OBJIndex
                        subelemento.SubItems(3) = 1
                               
               
                    Next i
    
                End If
                
    Call ListView1_Click
    Call ListView2_Click

    End If
    
    Exit Sub

ListViewQuest_ItemClick_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmQuestInfo.ListViewQuest_ItemClick", Erl)
    Resume Next
    
End Sub

Private Sub lstQuests_Click()
    
    On Error GoTo lstQuests_Click_Err
    
Dim QuestIndex As Byte

QuestIndex = Val(ReadField(1, lstQuests.List(lstQuests.ListIndex), Asc("-")))

FrmQuestInfo.ListView2.ListItems.Clear
FrmQuestInfo.ListView1.ListItems.Clear
            
                FrmQuestInfo.titulo.Caption = QuestList(QuestIndex).nombre
               
                FrmQuestInfo.Text1.Text = ""
                Call AddtoRichTextBox(Text1, QuestList(QuestIndex).desc & vbCrLf & "Nivel requerido: " & QuestList(QuestIndex).RequiredLevel & vbCrLf, 128, 128, 128)
                'tmpStr = tmpStr & "Detalles: " & .ReadASCIIString & vbCrLf
                'tmpStr = tmpStr & "Nivel requerido: " & .ReadByte & vbCrLf
               
               

                
                
                If UBound(QuestList(QuestIndex).RequiredNPC) > 0 Then 'Hay NPCs
                    If UBound(QuestList(QuestIndex).RequiredNPC) > 5 Then
                        FrmQuestInfo.ListView1.FlatScrollBar = False
                    Else
                        FrmQuestInfo.ListView1.FlatScrollBar = True
               
                    End If
                    
                    
                    For i = 1 To UBound(QuestList(QuestIndex).RequiredNPC)
                                                

                            Dim subelemento As ListItem
    
                            Set subelemento = FrmQuestInfo.ListView1.ListItems.Add(, , NpcData(QuestList(QuestIndex).RequiredNPC(i).NpcIndex).Name)
                           
                            subelemento.SubItems(1) = QuestList(QuestIndex).RequiredNPC(i).Amount
                            subelemento.SubItems(2) = QuestList(QuestIndex).RequiredNPC(i).NpcIndex
                            subelemento.SubItems(3) = 0

    
                    Next i
    
                End If
                    
    
                If LBound(QuestList(QuestIndex).RequiredOBJ) > 0 Then  'Hay OBJs
    
                    For i = 1 To UBound(QuestList(QuestIndex).RequiredOBJ)
                        Set subelemento = FrmQuestInfo.ListView1.ListItems.Add(, , ObjData(QuestList(QuestIndex).RequiredOBJ(i).OBJIndex).Name)
                        subelemento.SubItems(1) = QuestList(QuestIndex).RequiredOBJ(i).Amount
                        subelemento.SubItems(2) = QuestList(QuestIndex).RequiredOBJ(i).OBJIndex
                        subelemento.SubItems(3) = 1
                    Next i
    
                End If
        
               
                Set subelemento = FrmQuestInfo.ListView2.ListItems.Add(, , "Oro")
                           
                subelemento.SubItems(1) = QuestList(QuestIndex).RewardGLD
                subelemento.SubItems(2) = 12
                subelemento.SubItems(3) = 0
               
                Set subelemento = FrmQuestInfo.ListView2.ListItems.Add(, , "Experiencia")
                           
                subelemento.SubItems(1) = QuestList(QuestIndex).RewardEXP
                subelemento.SubItems(2) = 608
                subelemento.SubItems(3) = 1
               

                If UBound(QuestList(QuestIndex).RewardOBJ) > 0 Then
                
                    
                    For i = 1 To UBound(QuestList(QuestIndex).RewardOBJ)

                                                                   
                        Set subelemento = FrmQuestInfo.ListView2.ListItems.Add(, , ObjData(QuestList(QuestIndex).RewardOBJ(i).OBJIndex).Name)
                           
                        subelemento.SubItems(1) = QuestList(QuestIndex).RewardOBJ(i).Amount
                        subelemento.SubItems(2) = QuestList(QuestIndex).RewardOBJ(i).OBJIndex
                        subelemento.SubItems(3) = 1
                               
               
                    Next i
    
                End If


    
    Exit Sub

lstQuests_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmQuestInfo.lstQuests_Click", Erl)
    Resume Next
    
End Sub

Private Sub RichTextBox1_Change()

End Sub

