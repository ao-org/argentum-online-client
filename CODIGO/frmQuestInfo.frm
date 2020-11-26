VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form FrmQuestInfo 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12255
   LinkTopic       =   "Form1"
   Picture         =   "frmQuestInfo.frx":0000
   ScaleHeight     =   6510
   ScaleWidth      =   12255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3255
      Left            =   3120
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "frmQuestInfo.frx":103D2C
      Top             =   2040
      Width           =   3375
   End
   Begin VB.ListBox lstQuests 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3345
      Left            =   480
      TabIndex        =   8
      Top             =   1800
      Width           =   2115
   End
   Begin VB.PictureBox picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   10440
      ScaleHeight     =   32
      ScaleMode       =   0  'User
      ScaleWidth      =   32
      TabIndex        =   5
      Top             =   1800
      Width           =   480
   End
   Begin VB.PictureBox PlayerView 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1380
      Left            =   6720
      ScaleHeight     =   92
      ScaleMode       =   0  'User
      ScaleWidth      =   165
      TabIndex        =   3
      Top             =   3960
      Width           =   2475
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1500
      Left            =   6720
      TabIndex        =   2
      Top             =   1680
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   2646
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   14737632
      BackColor       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
      Picture         =   "frmQuestInfo.frx":104375
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   2520
      Left            =   9600
      TabIndex        =   4
      Top             =   2760
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   4445
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FlatScrollBar   =   -1  'True
      _Version        =   393217
      ForeColor       =   14737632
      BackColor       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Criatura"
         Object.Width           =   2646
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
      Picture         =   "frmQuestInfo.frx":104C46
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   11760
      Top             =   0
      Width           =   495
   End
   Begin VB.Label npclbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "adsdasda"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6840
      TabIndex        =   7
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label objetolbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "info_obj"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   9720
      TabIndex        =   6
      Top             =   2520
      Width           =   1935
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   6360
      Tag             =   "0"
      Top             =   5760
      Width           =   1980
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   3840
      Tag             =   "0"
      Top             =   5760
      Width           =   1980
   End
   Begin VB.Label detalle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "asdsaad"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   3255
      Left            =   3120
      TabIndex        =   1
      Top             =   2160
      Width           =   3405
   End
   Begin VB.Label titulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mision"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3120
      TabIndex        =   0
      Top             =   1680
      Width           =   3495
   End
End
Attribute VB_Name = "FrmQuestInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    'Me.Picture = LoadInterface("mision.bmp")

Text1.BackColor = RGB(11, 11, 11)
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)

    If (KeyAscii = 27) Then
        Unload Me

    End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Image1.Picture = Nothing
    Image1.Tag = 0
    Image2.Picture = Nothing
    Image2.Tag = 0

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Image1.Tag = "0" Then
        Image1.Picture = LoadInterface("boton-rechazar-es-over.bmp")
        Image1.Tag = "1"

    End If

End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Unload Me

End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Image2.Tag = "0" Then
        Image2.Picture = LoadInterface("boton-aceptar-ES-over.bmp")
        Image2.Tag = "1"

    End If

End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If lstQuests.ListIndex + 1 > 0 Then
        Call WriteQuestAccept(lstQuests.ListIndex + 1)
        Unload Me
    End If
End Sub

Private Sub Image3_Click()
Unload Me
End Sub

Public Sub ListView1_Click()

    If ListView1.SelectedItem.SubItems(2) <> "" Then
        If ListView1.SelectedItem.SubItems(3) = 0 Then
            Call DibujarBody(ListView1.SelectedItem.SubItems(2), 3)
      
            npclbl.Caption = NpcData(ListView1.SelectedItem.SubItems(2)).Name & " (" & ListView1.SelectedItem.SubItems(1) & ")"
    
        Else

            Dim x As Long

            Dim y As Long
        
            x = (PlayerView.ScaleWidth - GrhData(ListView1.SelectedItem.SubItems(2)).pixelWidth) / 2
            y = (PlayerView.ScaleHeight - GrhData(ListView1.SelectedItem.SubItems(2)).pixelHeight) / 2
            Call Grh_Render_To_Hdc(PlayerView, ObjData(ListView1.SelectedItem.SubItems(2)).GrhIndex, x, y, False)
        
            npclbl.Caption = ObjData(ListView1.SelectedItem.SubItems(2)).Name & " (" & ListView1.SelectedItem.SubItems(1) & ")"
    
        End If

    End If

End Sub

Sub DibujarBody(ByVal MyBody As Integer, Optional ByVal Heading As Byte = 3)

    On Error Resume Next

    Dim grh As grh

    grh = BodyData(NpcData(MyBody).Body).Walk(3)

    Dim x    As Long

    Dim y    As Long

    Dim grhH As grh

    grhH = HeadData(NpcData(MyBody).Head).Head(3)

    x = (PlayerView.ScaleWidth - GrhData(grh.GrhIndex).pixelWidth) / 2
    y = (PlayerView.ScaleHeight - GrhData(grh.GrhIndex).pixelHeight) / 2
    Call Grh_Render_To_Hdc(PlayerView, GrhData(grh.GrhIndex).Frames(1), x, y, False)

    If NpcData(MyBody).Head <> 0 Then
        x = (PlayerView.ScaleWidth - GrhData(grhH.GrhIndex).pixelWidth) / 2
        y = (PlayerView.ScaleHeight - GrhData(grhH.GrhIndex).pixelHeight) / 2 + 8 + BodyData(NpcData(MyBody).Body).HeadOffset.y
        Call Grh_Render_To_HdcSinBorrar(PlayerView, GrhData(grhH.GrhIndex).Frames(1), x, y, False)

    End If

End Sub

Public Sub ListView2_Click()

    If ListView2.SelectedItem.SubItems(2) <> "" Then
 
        Call Grh_Render_To_Hdc(picture1, ObjData(ListView2.SelectedItem.SubItems(2)).GrhIndex, 0, 0, False)
    
    End If
    
    objetolbl.Caption = ObjData(ListView2.SelectedItem.SubItems(2)).Name & vbCrLf & " (" & ListView2.SelectedItem.SubItems(1) & ")"

End Sub

Private Sub lstQuests_Click()
Dim QuestIndex As Byte

QuestIndex = Val(ReadField(1, lstQuests.List(lstQuests.ListIndex), Asc("-")))

FrmQuestInfo.ListView2.ListItems.Clear
FrmQuestInfo.ListView1.ListItems.Clear
            
                FrmQuestInfo.titulo.Caption = QuestList(QuestIndex).nombre
               
                
                FrmQuestInfo.detalle.Caption = QuestList(QuestIndex).desc & vbCrLf & "Nivel requerido: " & QuestList(QuestIndex).RequiredLevel & vbCrLf
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


End Sub

