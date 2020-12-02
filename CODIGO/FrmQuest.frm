VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form FrmQuests 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Quest"
   ClientHeight    =   6555
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmQuest.frx":0000
   ScaleHeight     =   6555
   ScaleWidth      =   12255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox detalle 
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
      Height          =   3375
      Left            =   3880
      LinkItem        =   "detalle"
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   1920
      Width           =   3260
   End
   Begin VB.PictureBox PlayerView 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   7320
      ScaleHeight     =   97
      ScaleMode       =   0  'User
      ScaleWidth      =   149
      TabIndex        =   6
      Top             =   3960
      Width           =   2235
   End
   Begin VB.PictureBox picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   10560
      ScaleHeight     =   32
      ScaleMode       =   0  'User
      ScaleWidth      =   32
      TabIndex        =   5
      Top             =   1840
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
      Height          =   3345
      Left            =   560
      TabIndex        =   1
      Top             =   1800
      Width           =   2955
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
      Height          =   1545
      Left            =   7340
      TabIndex        =   3
      Top             =   1680
      Width           =   2200
      _ExtentX        =   3889
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
         Object.Width           =   2506
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Cantidad"
         Object.Width           =   1201
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
      Picture         =   "FrmQuest.frx":103D2C
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   2325
      Left            =   9885
      TabIndex        =   4
      Top             =   2880
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   4101
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
         Object.Width           =   2295
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Cantidad"
         Object.Width           =   882
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
      Picture         =   "FrmQuest.frx":110C36
   End
   Begin VB.Label objetolbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "                                     "
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
      Height          =   435
      Left            =   9960
      TabIndex        =   9
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   11760
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
      Left            =   7320
      TabIndex        =   7
      Top             =   3780
      Width           =   2200
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   3880
      Tag             =   "0"
      Top             =   5730
      Width           =   1980
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   6400
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
      ForeColor       =   &H000EA4EB&
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   1680
      Width           =   3255
   End
End
Attribute VB_Name = "FrmQuests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
detalle.BackColor = RGB(11, 11, 11)
PlayerView.BackColor = RGB(11, 11, 11)
picture1.BackColor = RGB(19, 14, 11)
    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmQuests.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    Image1.Picture = Nothing
    Image1.Tag = 0
    Image2.Picture = Nothing
    Image2.Tag = 0

    
    Exit Sub

Form_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmQuests.Form_MouseMove", Erl)
    Resume Next
    
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Form_KeyPress_Err
    

    If (KeyAscii = 27) Then
        Unload Me

    End If

    
    Exit Sub

Form_KeyPress_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmQuests.Form_KeyPress", Erl)
    Resume Next
    
End Sub
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image1_MouseMove_Err
    

    If Image1.Tag = "0" Then
        Image1.Picture = LoadInterface("boton-abandonar-es-over.bmp")
        Image1.Tag = "1"

    End If

    
    Exit Sub

Image1_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmQuests.Image1_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image2_MouseUp_Err
    
    Unload Me

    
    Exit Sub

Image2_MouseUp_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmQuests.Image2_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image2_MouseMove_Err
    

    If Image2.Tag = "0" Then
        Image2.Picture = LoadInterface("boton-aceptar-ES-over.bmp")
        Image2.Tag = "1"

    End If

    
    Exit Sub

Image2_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmQuests.Image2_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image1_MouseUp_Err
    

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
            detalle.Text = ""
            titulo.Caption = ""
            picture1.Refresh
            PlayerView.Refresh
            ListView1.ListItems.Clear
            ListView2.ListItems.Clear

        Case vbNo   'Boton NO.
            'Como selecciono que no, no hace nada.
            Exit Sub

    End Select

    
    Exit Sub

Image1_MouseUp_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmQuests.Image1_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub Image3_Click()
    
    On Error GoTo Image3_Click_Err
    
Unload Me
    
    Exit Sub

Image3_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmQuests.Image3_Click", Erl)
    Resume Next
    
End Sub

Public Sub ListView1_Click()
    
    On Error GoTo ListView1_Click_Err
    

    

    If ListView1.SelectedItem.SubItems(2) <> "" Then
        If ListView1.SelectedItem.SubItems(3) = 0 Then
            Call DibujarBody(ListView1.SelectedItem.SubItems(2), 3)
      
            npclbl.Caption = NpcData(ListView1.SelectedItem.SubItems(2)).Name & " (" & ListView1.SelectedItem.SubItems(1) & ")"
        Else

            Dim x As Long

            Dim y As Long
        
            x = (PlayerView.ScaleWidth - GrhData(ListView1.SelectedItem.SubItems(2)).pixelWidth) / 2
            y = (PlayerView.ScaleHeight - GrhData(ListView1.SelectedItem.SubItems(2)).pixelHeight) / 2
            
            Call Grh_Render_To_HdcSinBorrar(PlayerView, ObjData(ListView1.SelectedItem.SubItems(2)).GrhIndex, x, y, False, RGB(11, 11, 11))
        
            npclbl.Caption = ObjData(ListView1.SelectedItem.SubItems(2)).Name & " (" & ListView1.SelectedItem.SubItems(1) & ")"
    
        End If

    End If

    
    Exit Sub

ListView1_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmQuests.ListView1_Click", Erl)
    Resume Next
    
End Sub

Sub DibujarBody(ByVal MyBody As Integer, Optional ByVal Heading As Byte = 3)
    
    On Error GoTo DibujarBody_Err
    

    Dim grh As grh

    grh = BodyData(NpcData(MyBody).Body).Walk(3)

    Dim x    As Long

    Dim y    As Long

    Dim grhH As grh

    grhH = HeadData(NpcData(MyBody).Head).Head(3)

    x = (PlayerView.ScaleWidth - GrhData(grh.GrhIndex).pixelWidth) / 2
    y = (PlayerView.ScaleHeight - GrhData(grh.GrhIndex).pixelHeight) / 2
    Call Grh_Render_To_HdcSinBorrar(PlayerView, GrhData(grh.GrhIndex).Frames(1), x, y, False, RGB(11, 11, 11))

    If NpcData(MyBody).Head <> 0 Then
        x = (PlayerView.ScaleWidth - GrhData(grhH.GrhIndex).pixelWidth) / 2
        y = (PlayerView.ScaleHeight - GrhData(grhH.GrhIndex).pixelHeight) / 2 + 8 + BodyData(NpcData(MyBody).Body).HeadOffset.y
        Call Grh_Render_To_HdcSinBorrar(PlayerView, GrhData(grhH.GrhIndex).Frames(1), x, y, False, RGB(11, 11, 11))

    End If

    
    Exit Sub

DibujarBody_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmQuests.DibujarBody", Erl)
    Resume Next
    
End Sub

Public Sub ListView2_Click()
    
    On Error GoTo ListView2_Click_Err
    

    

    If ListView2.SelectedItem.SubItems(2) <> "" Then
 
        Call Grh_Render_To_HdcSinBorrar(picture1, ObjData(ListView2.SelectedItem.SubItems(2)).GrhIndex, 0, 0, False, RGB(11, 11, 11))
        picture1.Visible = True
        
        objetolbl.Caption = ObjData(ListView2.SelectedItem.SubItems(2)).Name & vbCrLf & " (" & ListView2.SelectedItem.SubItems(1) & ")"
    
    End If

    
    Exit Sub

ListView2_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmQuests.ListView2_Click", Erl)
    Resume Next
    
End Sub

Public Sub lstQuests_Click()
    
    On Error GoTo lstQuests_Click_Err
    

    If lstQuests.ListIndex < 0 Then Exit Sub
    
    Call WriteQuestDetailsRequest(lstQuests.ListIndex + 1)

    
    Exit Sub

lstQuests_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmQuests.lstQuests_Click", Erl)
    Resume Next
    
End Sub

