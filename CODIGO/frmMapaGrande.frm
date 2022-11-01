VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMapaGrande 
   Appearance      =   0  'Flat
   BackColor       =   &H80000006&
   BorderStyle     =   0  'None
   ClientHeight    =   10788
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11568
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   899
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   964
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox PlayerView 
      Appearance      =   0  'Flat
      BackColor       =   &H000A0A0A&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   8520
      ScaleHeight     =   89
      ScaleMode       =   0  'User
      ScaleWidth      =   177
      TabIndex        =   10
      Top             =   8760
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   8685
      ScaleHeight     =   33.032
      ScaleMode       =   0  'User
      ScaleWidth      =   32
      TabIndex        =   9
      Top             =   7530
      Width           =   480
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   960
      Left            =   8520
      TabIndex        =   3
      Top             =   3900
      Width           =   2595
      _ExtentX        =   4572
      _ExtentY        =   1693
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   657930
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Criatura"
         Object.Width           =   3575
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "Cantidad"
         Object.Width           =   818
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Index"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ListView listdrop 
      Height          =   780
      Left            =   9405
      TabIndex        =   8
      Top             =   7380
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   1376
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   657930
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "item"
         Object.Width           =   3177
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "grhindex"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.PictureBox picMap 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8910
      Left            =   480
      ScaleHeight     =   743
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   1410
      Width           =   7680
      Begin VB.Shape Shape2 
         BackColor       =   &H00FF0000&
         BorderColor     =   &H00C00000&
         FillColor       =   &H0000FFFF&
         Height          =   480
         Left            =   600
         Top             =   600
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Shape llamadadeclan 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         Height          =   180
         Left            =   240
         Shape           =   3  'Circle
         Tag             =   "0"
         Top             =   120
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         Height          =   75
         Left            =   4800
         Shape           =   1  'Square
         Top             =   4800
         Width           =   75
      End
      Begin VB.Shape lblAllies 
         BorderColor     =   &H000000C0&
         FillColor       =   &H0000FFFF&
         Height          =   405
         Left            =   1920
         Top             =   2880
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label lblPos 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   180
         TabIndex        =   1
         Top             =   8220
         Width           =   60
      End
      Begin VB.Image imgSwitchWorld 
         Height          =   435
         Index           =   1
         Left            =   8790
         Tag             =   "0"
         Top             =   0
         Width           =   435
      End
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8760
      TabIndex        =   15
      Top             =   6480
      Width           =   2250
   End
   Begin VB.Label Label8 
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
      Height          =   255
      Left            =   8550
      TabIndex        =   14
      Top             =   10080
      Width           =   2655
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   9000
      TabIndex        =   13
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   8760
      TabIndex        =   12
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "La informacion del mapa esta aquí."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   8520
      TabIndex        =   11
      Top             =   1800
      Width           =   2670
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   9795
      Top             =   2385
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   8565
      Top             =   2385
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   420
      Left            =   11100
      Tag             =   "0"
      Top             =   0
      Width           =   465
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8760
      TabIndex        =   7
      Top             =   5685
      Width           =   2250
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8760
      TabIndex        =   6
      Top             =   5955
      UseMnemonic     =   0   'False
      Width           =   2250
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8760
      TabIndex        =   5
      Top             =   6210
      Width           =   2250
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   8760
      TabIndex        =   4
      Top             =   5430
      Width           =   2250
   End
   Begin VB.Label lblMapInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Mapa Desconocido"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   555
      Index           =   0
      Left            =   8490
      TabIndex        =   2
      Top             =   1320
      Width           =   2715
   End
End
Attribute VB_Name = "frmMapaGrande"
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

Public bmoving      As Boolean

Public dX           As Integer

Public dy           As Integer

Public Referencias  As Boolean

Private Const TILE_SIZE = 27

Private Const MAPAS_ANCHO = 19

Private Const MAPAS_ALTO = 22

Private Sub Form_Activate()
    Call CargarDatosMapa(UserMap)
    If ListView1.ListItems.count > 0 Then
        Call ListView1_ItemClick(ListView1.ListItems.Item(1))
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo Form_KeyDown_Err
    
    If KeyCode = 27 Then
        Me.Visible = False
    End If
    
    Exit Sub

Form_KeyDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMapaGrande.Form_KeyDown", Erl)
    Resume Next
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    ListView1.BackColor = RGB(7, 7, 7)
    listdrop.BackColor = RGB(7, 7, 7)
    lblMapInfo(0).ForeColor = RGB(235, 164, 14)
    
    Call FormParser.Parse_Form(Me)
    Call Aplicar_Transparencia(Me.hWnd, 240)
    
   ' picMap.Picture = LoadInterface("mapa.bmp")
    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMapaGrande.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    MoverForm Me.hwnd
    Image1 = Nothing
    
    If Image1.Tag = "1" Then
        Image1.Picture = Nothing
        Image1.Tag = "0"
    End If
    
    Exit Sub

Form_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMapaGrande.Form_MouseMove", Erl)
    Resume Next
    
End Sub

Sub DibujarHead(ByVal MyHead As Integer, ByVal yoff As Integer, Optional ByVal Heading As Byte = 3)
    
    On Error GoTo DibujarHead_Err
    

    Dim grh As grh

    grh = HeadData(MyHead).Head(3)

    Dim x As Long

    Dim y As Long

    x = PlayerView.ScaleWidth / 2 - GrhData(grh.GrhIndex).pixelWidth / 2
    y = PlayerView.ScaleHeight / 2 - GrhData(grh.GrhIndex).pixelHeight + yoff / 2
    Call Grh_Render_To_Hdc(PlayerView, grh.GrhIndex, x - 1, y, False)

    
    Exit Sub

DibujarHead_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMapaGrande.DibujarHead", Erl)
    Resume Next
    
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Image1.Picture = LoadInterface("boton-cerrar-off.bmp")
    Image1.Tag = "1"
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Image1.Tag = "0" Then
        Image1.Picture = LoadInterface("boton-cerrar-over.bmp")
        Image1.Tag = "1"
    End If
End Sub

Private Sub ListView1_beforeEdit(ByVal Columna As Integer, Cancel As Boolean)
    
    On Error GoTo ListView1_beforeEdit_Err
    

    If Columna > 5 Then
        Cancel = True

    End If

    
    Exit Sub

ListView1_beforeEdit_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMapaGrande.ListView1_beforeEdit", Erl)
    Resume Next
    
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Image1_MouseUp_Err
    
    Me.Visible = False

    
    Exit Sub

Image1_MouseUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMapaGrande.Image1_MouseUp", Erl)
    Resume Next
    
End Sub

Private Sub Image2_Click()
    
    On Error GoTo Image2_Click_Err

    If WorldActual = 1 Then
        WorldActual = 2
        Image2.Picture = LoadInterface("check-amarillo.bmp")
    Else
        WorldActual = 1
        Image2.Picture = Nothing
    End If
    
    ActualizarPosicionMapa

    picMap.Picture = LoadInterface("mapa" & WorldActual & ".bmp")
    
    Exit Sub

Image2_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMapaGrande.Image2_Click", Erl)
    Resume Next
    
End Sub

Private Sub Image3_Click()
    
    On Error GoTo Image3_Click_Err
    

    If Dungeon Then Exit Sub

    If Referencias Then
        picMap.Picture = LoadInterface("mapa" & WorldActual & ".bmp")
        Image3.Picture = Nothing
        Referencias = False
    Else
        Referencias = True
        picMap.Picture = LoadInterface("mapa_referencias.bmp")
        Image3.Picture = LoadInterface("check-amarillo.bmp")

    End If

    
    Exit Sub

Image3_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMapaGrande.Image3_Click", Erl)
    Resume Next
    
End Sub

Private Sub Label6_Click()
    
    On Error GoTo Label6_Click_Err
    
    Call Image2_Click

    
    Exit Sub

Label6_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMapaGrande.Label6_Click", Erl)
    Resume Next
    
End Sub

Private Sub Label7_Click()
    
    On Error GoTo Label7_Click_Err
    
    Call Image3_Click

    
    Exit Sub

Label7_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMapaGrande.Label7_Click", Erl)
    Resume Next
    
End Sub

Private Sub listdrop_Click()
    
    On Error GoTo listdrop_Click_Err
    

    

    'Picture1.Refresh
    Picture1.BackColor = vbBlack
    Picture1.Refresh

    'Call Grh_Render_To_Hdc(Picture1, ObjData(NpcData(ListView1.SelectedItem.SubItems(2)).QuizaDropea(listdrop.SelectedItem.Index)).grhindex, 0, 0, False)
    If listdrop.ListItems.count <= 0 Then Exit Sub
    Call Grh_Render_To_Hdc(Picture1, listdrop.SelectedItem.SubItems(1), 0, 0, False)

    
    Exit Sub

listdrop_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMapaGrande.listdrop_Click", Erl)
    Resume Next
    
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    On Error GoTo ListView1_Click_Err

    Label8.Caption = ""
    Picture1.Refresh
    
    If ListView1.ListItems.count <= 0 Then Exit Sub

    Label8.Caption = NpcData(ListView1.SelectedItem.SubItems(2)).Name
    
    Call DibujarNPC(Me.PlayerView, NpcData(ListView1.SelectedItem.SubItems(2)).Head, NpcData(ListView1.SelectedItem.SubItems(2)).Body)

    Dim i As Byte

    Label2.Caption = "Vida: " & NpcData(ListView1.SelectedItem.SubItems(2)).Hp & " puntos"
    Label3.Caption = "Experiencia: " & NpcData(ListView1.SelectedItem.SubItems(2)).exp & " puntos"
    Label4.Caption = "Oro: " & NpcData(ListView1.SelectedItem.SubItems(2)).oro & " monedas"
    Label5.Caption = "Ataque: " & NpcData(ListView1.SelectedItem.SubItems(2)).MinHit & "/" & NpcData(ListView1.SelectedItem.SubItems(2)).MaxHit
    Label9.Caption = "Exp. de clan: " & NpcData(ListView1.SelectedItem.SubItems(2)).ExpClan & " puntos"
    listdrop.ListItems.Clear
    
    ListView1.ToolTipText = NpcData(ListView1.SelectedItem.SubItems(2)).Name

    If ListView1.SelectedItem.SubItems(2) <> "" Then
    
        If NpcData(ListView1.SelectedItem.SubItems(2)).NumQuiza <> 0 Then

            '    Call Grh_Render_To_Hdc(Picture1, ObjData(NpcData(ListView1.SelectedItem.SubItems(2)).QuizaDropea(1)).grhindex,' 0, 0, False)
            If NpcData(ListView1.SelectedItem.SubItems(2)).NumQuiza = 0 Then Exit Sub

            For i = 1 To NpcData(ListView1.SelectedItem.SubItems(2)).NumQuiza
                '  listdrop.ListItems.Add(1).Text = ObjData((NpcData(ListView1.SelectedItem.SubItems(2)).QuizaDropea(i))).name
                'listdrop.ListItems.Add(1).SubItems(2) = ObjData((NpcData(ListView1.SelectedItem.SubItems(2)).QuizaDropea(i))).grhindex
            
                ' Dim subelemento As ListItem

                Dim subelemento As ListItem

                Set subelemento = listdrop.ListItems.Add(, , ObjData((NpcData(ListView1.SelectedItem.SubItems(2)).QuizaDropea(i))).Name)

                subelemento.SubItems(1) = ObjData((NpcData(ListView1.SelectedItem.SubItems(2)).QuizaDropea(i))).GrhIndex

            Next i

            Call listdrop_Click

        End If

    End If
        
    
    Exit Sub

ListView1_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMapaGrande.ListView1_Click", Erl)
    Resume Next
    
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub picMap_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo picMap_MouseDown_Err
    

    

    lblAllies.Visible = True

    Dim PosX As Integer

    Dim PosY As Integer

    Dim Mapa As Integer
    
    'lblAllies.top = Y * 18 / 32
    'lblAllies.left = X * 14 / 32
    
    If x >= llamadadeclan.Left And x <= llamadadeclan.Left + llamadadeclan.Width And y >= llamadadeclan.Top And y <= llamadadeclan.Top + llamadadeclan.Height Then
        AddtoRichTextBox frmMain.RecTxt, "Ubicación de tu compañero de clan que solicita ayuda: (" & LLamadaDeclanMapa & "-" & LLamadaDeclanX & "-" & LLamadaDeclanY & ").", 2, 51, 223, 1, 1

    End If

    ' Para obtener las coordenadas (x, y) del "slot" divido la posición del cursor
    ' por el tamaño de los tiles y me quedo solo con la parte entera
    PosX = Int(x / TILE_SIZE) ' PosX = Valor entero entre 0 y (MAPAS_ANCHO - 1)
    PosY = Int(y / TILE_SIZE) ' PosY = Valor entero entre 0 y (MAPAS_ALTO - 1)
    
    ' Uso estas coordeandas para calcular el índice del mapa
    Mapa = PosX + PosY * MAPAS_ANCHO + 1 ' +1 porque los mapas empiezan en 1
    
    ' Luego multiplico por TILE_SIZE para tener la posición final en donde poner el indicador
    PosX = PosX * TILE_SIZE
    PosY = PosY * TILE_SIZE


        If Mundo(WorldActual).MapIndice(Mapa) = 0 Then Exit Sub
        Call CargarDatosMapa(Mundo(WorldActual).MapIndice(Mapa))
        lblMapInfo(0) = MapDat.map_name & "(" & Mundo(WorldActual).MapIndice(Mapa) & ")"
        
        If Button = vbRightButton Then
            Call ParseUserCommand("/TELEP YO " & Mundo(WorldActual).MapIndice(Mapa) & " " & 50 & " " & 50)

        End If

    Label2.Caption = ""
    Label3.Caption = ""
    Label4.Caption = ""
    Label5.Caption = ""
    Label9.Caption = ""
    listdrop.ListItems.Clear
    
    ListView1.SetFocus
    'listdrop.SetFocus
    If ListView1.ListItems.count > 0 Then
        Call ListView1_ItemClick(ListView1.ListItems.Item(1))
    End If
    Call listdrop_Click

    lblAllies.Top = PosY
    lblAllies.Left = PosX

    
    Exit Sub

picMap_MouseDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMapaGrande.picMap_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub picMap_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo picMap_MouseMove_Err
    
    MoverForm Me.hwnd

    
    Exit Sub

picMap_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmMapaGrande.picMap_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub ActualizarPosicion(ByVal map As Integer)
    Dim x As Long, y As Long

    x = (map - 1) Mod MAPAS_ANCHO
    y = Int((map - 1) / MAPAS_ANCHO)

    Shape1.Top = y * TILE_SIZE + (UserPos.y * TILE_SIZE / 100) - Shape1.Height \ 2
    Shape1.Left = x * TILE_SIZE + (UserPos.x * TILE_SIZE / 100) - Shape1.Width \ 2
    
    Shape1.Visible = True
End Sub

Public Sub ActualizarPosicionMapa()
    Dim Index As Integer

    For Index = 1 To Mundo(WorldActual).Ancho * Mundo(WorldActual).Alto

        If Mundo(WorldActual).MapIndice(Index) = UserMap Then
            Call ActualizarPosicion(Index)
            Exit Sub
        End If
    Next
    
    Shape1.Visible = False
End Sub

Public Sub CalcularPosicionMAPA()
    
    On Error GoTo CalcularPosicionMAPA_Err
    
    frmMapaGrande.lblMapInfo(0) = MapDat.map_name & "(" & UserMap & ")"

    If NameMaps(UserMap).desc <> "" Then
        frmMapaGrande.Label1.Caption = NameMaps(UserMap).desc
    Else
        frmMapaGrande.Label1.Caption = "Sin información relevante."

    End If

    Dim i       As Integer
    Dim j       As Byte

    Dim Encontre As Boolean
    
    
    For j = 1 To TotalWorlds
        For i = 1 To Mundo(j).Ancho * Mundo(j).Alto
    
            If Mundo(j).MapIndice(i) = UserMap Then
                idmap = i
                Encontre = True
                frmMapaGrande.picMap.Picture = LoadInterface("mapa" & j & ".bmp")
                WorldActual = j

                If j = 1 Then
                    frmMapaGrande.Image2.Picture = Nothing
                Else
                    frmMapaGrande.Image2.Picture = LoadInterface("check-amarillo.bmp")
                End If
                
                Call ActualizarPosicion(idmap)

                Dim x As Long, y As Long
                x = (idmap - 1) Mod MAPAS_ANCHO
                y = Int((idmap - 1) / MAPAS_ANCHO)
                lblAllies.Top = y * TILE_SIZE
                lblAllies.Left = x * TILE_SIZE
                lblAllies.Visible = True

                Exit For
            End If
        Next i
        
        If Encontre Then
            Exit For
        End If
    Next j
    
    If Encontre = False Then
        If Not frmMapaGrande.Visible Then
            WorldActual = 1
            frmMapaGrande.picMap.Picture = LoadInterface("mapa1.bmp")
            frmMapaGrande.Image2.Picture = Nothing
        End If

    End If
    
    Call CargarDatosMapa(UserMap)
    
    Exit Sub

CalcularPosicionMAPA_Err:
    Call RegistrarError(Err.Number, Err.Description, "ModUtils.CalcularPosicionMAPA", Erl)
    Resume Next
    
End Sub
