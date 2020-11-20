VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form FrmQuestInfo 
   BackColor       =   &H00000709&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10365
   LinkTopic       =   "Form1"
   Picture         =   "frmQuestInfo.frx":0000
   ScaleHeight     =   6510
   ScaleWidth      =   10365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   8520
      ScaleHeight     =   32
      ScaleMode       =   0  'User
      ScaleWidth      =   32
      TabIndex        =   5
      Top             =   1700
      Width           =   480
   End
   Begin VB.PictureBox PlayerView 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1380
      Left            =   4830
      ScaleHeight     =   92
      ScaleMode       =   0  'User
      ScaleWidth      =   165
      TabIndex        =   3
      Top             =   3990
      Width           =   2475
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1500
      Left            =   4850
      TabIndex        =   2
      Top             =   1700
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
      Picture         =   "frmQuestInfo.frx":DB8F4
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   3000
      Left            =   7680
      TabIndex        =   4
      Top             =   2280
      Width           =   2200
      _ExtentX        =   3889
      _ExtentY        =   5292
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
      Picture         =   "frmQuestInfo.frx":DC1C5
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   9840
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
      Left            =   4920
      TabIndex        =   7
      Top             =   3800
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
      Height          =   375
      Left            =   7800
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   5440
      Tag             =   "0"
      Top             =   5730
      Width           =   1980
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   2920
      Tag             =   "0"
      Top             =   5730
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
      Height          =   3375
      Left            =   540
      TabIndex        =   1
      Top             =   1920
      Width           =   4005
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
      Left            =   380
      TabIndex        =   0
      Top             =   1680
      Width           =   4335
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
    Call WriteQuestAccept
    Unload Me

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
        
            x = PlayerView.ScaleWidth / 2 - GrhData(ListView1.SelectedItem.SubItems(2)).pixelWidth / 2
            y = PlayerView.ScaleHeight / 2 - GrhData(ListView1.SelectedItem.SubItems(2)).pixelHeight / 2
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
    
    objetolbl.Caption = ObjData(ListView2.SelectedItem.SubItems(2)).Name & vbCrLf & " (" & ListView2.SelectedItem.SubItems(1) & ")"

End Sub

