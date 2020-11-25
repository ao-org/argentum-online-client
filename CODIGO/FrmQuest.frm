VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form FrmQuests 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Quest"
   ClientHeight    =   6495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmQuest.frx":0000
   ScaleHeight     =   6495
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PlayerView 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   7560
      ScaleHeight     =   97
      ScaleMode       =   0  'User
      ScaleWidth      =   157
      TabIndex        =   7
      Top             =   3960
      Width           =   2355
   End
   Begin VB.PictureBox picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   360
      ScaleHeight     =   32
      ScaleMode       =   0  'User
      ScaleWidth      =   32
      TabIndex        =   6
      Top             =   5760
      Visible         =   0   'False
      Width           =   480
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
      Height          =   1200
      Left            =   550
      TabIndex        =   1
      Top             =   1800
      Width           =   2115
   End
   Begin VB.TextBox txtInfo 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   3670
      Left            =   12600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2400
      Width           =   4335
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1550
      Left            =   7485
      TabIndex        =   4
      Top             =   1680
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   2725
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      TextBackground  =   -1  'True
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
         Object.Width           =   3353
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Cantidad"
         Object.Width           =   973
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
      Picture         =   "FrmQuest.frx":DB8F4
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   1365
      Left            =   480
      TabIndex        =   5
      Top             =   3960
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   2408
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
         Object.Width           =   2560
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Cantidad"
         Object.Width           =   1376
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
      Picture         =   "FrmQuest.frx":E87FE
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   9960
      Top             =   0
      Width           =   375
   End
   Begin VB.Label npclbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "asdadsada"
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
      Height          =   375
      Left            =   7560
      TabIndex        =   8
      Top             =   3775
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   2920
      Tag             =   "0"
      Top             =   5730
      Width           =   1980
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   5450
      Tag             =   "0"
      Top             =   5730
      Width           =   1980
   End
   Begin VB.Label titulo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "¡No tenes misiones!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   1680
      Width           =   4335
   End
   Begin VB.Label detalle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   3375
      Left            =   3120
      TabIndex        =   2
      Top             =   2040
      Width           =   4160
   End
End
Attribute VB_Name = "FrmQuests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Image1.Picture = Nothing
    Image1.Tag = 0
    Image2.Picture = Nothing
    Image2.Tag = 0

End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)

    If (KeyAscii = 27) Then
        Unload Me

    End If

End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Image1.Tag = "0" Then
        Image1.Picture = LoadInterface("boton-abandonar-es-over.bmp")
        Image1.Tag = "1"

    End If

End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Unload Me

End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Image2.Tag = "0" Then
        Image2.Picture = LoadInterface("boton-aceptar-ES-over.bmp")
        Image2.Tag = "1"

    End If

End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If lstQuests.ListCount = 0 Then
        MsgBox "No tienes ninguna mision!", vbOKOnly + vbExclamation
        Exit Sub

    End If

    'Chequeamos si tiene algun item seleccionado.
    If lstQuests.ListIndex < 0 Then
        MsgBox "Primero debes seleccionar una mision!", vbOKOnly + vbExclamation
        Exit Sub

    End If
            
    Select Case MsgBox("Estas seguro que deseas abandonar la mision?", vbYesNo + vbExclamation)

        Case vbYes  'Boton Si.
            'Enviamos el paquete para abandonar la quest
            Call WriteQuestAbandon(lstQuests.ListIndex + 1)
            detalle.Caption = ""
            titulo.Caption = ""
            picture1.Refresh
            PlayerView.Refresh
            ListView1.ListItems.Clear
            ListView2.ListItems.Clear

        Case vbNo   'Boton NO.
            'Como selecciono que no, no hace nada.
            Exit Sub

    End Select

End Sub

Private Sub Image3_Click()
Unload Me
End Sub

Public Sub ListView1_Click()

    On Error Resume Next

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

    On Error Resume Next

    If ListView2.SelectedItem.SubItems(2) <> "" Then
 
        Call Grh_Render_To_Hdc(picture1, ObjData(ListView2.SelectedItem.SubItems(2)).GrhIndex, 0, 0, False)
    
    End If

End Sub

Public Sub lstQuests_Click()

    If lstQuests.ListIndex < 0 Then Exit Sub
    
    Call WriteQuestDetailsRequest(lstQuests.ListIndex + 1)

End Sub

