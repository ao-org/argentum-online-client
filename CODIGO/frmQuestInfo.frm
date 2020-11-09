VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmQuestInfo 
   BackColor       =   &H00000709&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8115
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   8115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   6370
      ScaleHeight     =   32
      ScaleMode       =   0  'User
      ScaleWidth      =   32
      TabIndex        =   5
      Top             =   1640
      Width           =   480
   End
   Begin VB.PictureBox PlayerView 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1500
      Left            =   3000
      ScaleHeight     =   100
      ScaleMode       =   0  'User
      ScaleWidth      =   157
      TabIndex        =   3
      Top             =   3310
      Width           =   2355
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1125
      Left            =   2895
      TabIndex        =   2
      Top             =   1560
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   1984
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   12632256
      BackColor       =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
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
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   1845
      Left            =   5640
      TabIndex        =   4
      Top             =   2805
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   3254
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
         Name            =   "Arial"
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
         Object.Width           =   2119
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
      Picture         =   "frmQuestInfo.frx":0000
   End
   Begin VB.Label npclbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   495
      Left            =   2960
      TabIndex        =   7
      Top             =   3160
      Width           =   2295
   End
   Begin VB.Label objetolbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      Top             =   2360
      Width           =   1935
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   5080
      Tag             =   "0"
      Top             =   5260
      Width           =   1470
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   1980
      Tag             =   "0"
      Top             =   5250
      Width           =   1665
   End
   Begin VB.Label detalle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Left            =   435
      TabIndex        =   1
      Top             =   1605
      Width           =   2295
   End
   Begin VB.Label titulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   1230
      Width           =   2175
   End
End
Attribute VB_Name = "FrmQuestInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    'Me.Picture = LoadInterface("mision.bmp")

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Image1.Picture = Nothing
    Image1.Tag = 0
    Image2.Picture = Nothing
    Image2.Tag = 0

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Image1.Tag = "0" Then
        Image1.Picture = LoadInterface("quest_volver.bmp")
        Image1.Tag = "1"

    End If

End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Unload Me

End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Image2.Tag = "0" Then
        Image2.Picture = LoadInterface("quest_acepto.bmp")
        Image2.Tag = "1"

    End If

End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call WriteQuestAccept
    Unload Me

End Sub

Public Sub ListView1_Click()

    If ListView1.SelectedItem.SubItems(2) <> "" Then
        If ListView1.SelectedItem.SubItems(3) = 0 Then
            Call DibujarBody(ListView1.SelectedItem.SubItems(2), 3)
      
            npclbl.Caption = NpcData(ListView1.SelectedItem.SubItems(2)).name & " (" & ListView1.SelectedItem.SubItems(1) & ")"
    
        Else

            Dim x As Long

            Dim y As Long
        
            x = PlayerView.ScaleWidth / 2 - GrhData(ListView1.SelectedItem.SubItems(2)).pixelWidth / 2
            y = PlayerView.ScaleHeight / 2 - GrhData(ListView1.SelectedItem.SubItems(2)).pixelHeight / 2
            Call Grh_Render_To_Hdc(PlayerView, ObjData(ListView1.SelectedItem.SubItems(2)).GrhIndex, x, y, False)
        
            npclbl.Caption = ObjData(ListView1.SelectedItem.SubItems(2)).name & " (" & ListView1.SelectedItem.SubItems(1) & ")"
    
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

    x = PlayerView.ScaleWidth / 2 - GrhData(grh.GrhIndex).pixelWidth / 2
    y = PlayerView.ScaleHeight / 2 - GrhData(grh.GrhIndex).pixelHeight / 2
    Call Grh_Render_To_Hdc(PlayerView, GrhData(grh.GrhIndex).Frames(1), x, y, False)

    If NpcData(MyBody).Head <> 0 Then
        x = PlayerView.ScaleWidth / 2 - GrhData(grhH.GrhIndex).pixelWidth / 2
        y = PlayerView.ScaleHeight / 2 - GrhData(grhH.GrhIndex).pixelHeight + 8 + BodyData(NpcData(MyBody).Body).HeadOffset.y / 2
        Call Grh_Render_To_HdcSinBorrar(PlayerView, GrhData(grhH.GrhIndex).Frames(1), x, y, False)

    End If

End Sub

Public Sub ListView2_Click()

    If ListView2.SelectedItem.SubItems(2) <> "" Then
 
        Call Grh_Render_To_Hdc(picture1, ObjData(ListView2.SelectedItem.SubItems(2)).GrhIndex, 0, 0, False)
    
    End If
    
    objetolbl.Caption = ObjData(ListView2.SelectedItem.SubItems(2)).name & vbCrLf & " (" & ListView2.SelectedItem.SubItems(1) & ")"

End Sub
