VERSION 5.00
Begin VB.Form FrmSastre 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Trabajar de sastre"
   ClientHeight    =   7125
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmSastre.frx":0000
   ScaleHeight     =   7125
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox cantidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   250
      Left            =   3960
      MaxLength       =   3
      TabIndex        =   4
      Text            =   "1"
      Top             =   5640
      Width           =   465
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
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
      ItemData        =   "FrmSastre.frx":96C60
      Left            =   5655
      List            =   "FrmSastre.frx":96C67
      TabIndex        =   3
      Top             =   3840
      Width           =   330
   End
   Begin VB.ListBox List1 
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
      Left            =   4020
      TabIndex        =   2
      Top             =   3840
      Width           =   1605
   End
   Begin VB.ListBox lstArmas 
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
      Height          =   2760
      Left            =   480
      TabIndex        =   1
      Top             =   2160
      Width           =   3120
   End
   Begin VB.PictureBox picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   4780
      ScaleHeight     =   32
      ScaleMode       =   0  'User
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   1850
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   315
      Left            =   4680
      Tag             =   "0"
      Top             =   5640
      Width           =   255
   End
   Begin VB.Image Image2 
      Height          =   315
      Left            =   3360
      Tag             =   "0"
      Top             =   5640
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   6000
      Tag             =   "0"
      Top             =   0
      Width           =   510
   End
   Begin VB.Label desc 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   5
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Image Command3 
      Height          =   400
      Left            =   3530
      Tag             =   "0"
      Top             =   6345
      Width           =   1950
   End
   Begin VB.Image Command4 
      Height          =   400
      Left            =   1080
      Tag             =   "0"
      Top             =   6345
      Width           =   1950
   End
   Begin VB.Image Command2 
      Height          =   585
      Left            =   2190
      Tag             =   "0"
      Top             =   1320
      Width           =   660
   End
   Begin VB.Image Command1 
      Height          =   585
      Left            =   1140
      Tag             =   "0"
      Top             =   1320
      Width           =   660
   End
End
Attribute VB_Name = "FrmSastre"
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
Private indice As Byte

Const PielTigreBengalaIndex = 1145
Const PielTigreIndex = 4339
Const BlackWolfIndex = 1146



Private Sub Command1_Click()
    
    On Error GoTo Command1_Click_Err
    

    Dim i As Byte

    indice = 1
    lstArmas.Clear

    For i = 1 To UBound(SastreRopas())

        If SastreRopas(i).Index = 0 Then Exit For
        lstArmas.AddItem (ObjData(SastreRopas(i).Index).Name)
    Next i

    Command1.Picture = LoadInterface("sastreria_vestimentahover.bmp")
    Command2.Picture = Nothing

    
    Exit Sub

Command1_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmSastre.Command1_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command2_Click()
    
    On Error GoTo Command2_Click_Err
    

    Dim i As Byte

    indice = 2
    lstArmas.Clear

    For i = 1 To UBound(SastreGorros())

        If SastreGorros(i).Index = 0 Then Exit For
        lstArmas.AddItem (ObjData(SastreGorros(i).Index).Name)
    Next i
    
    Command2.Picture = LoadInterface("sastreria_gorroshover.bmp")
    Command1.Picture = Nothing

    
    Exit Sub

Command2_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmSastre.Command2_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command3_Click()
    
    On Error GoTo Command3_Click_Err
    

    

    If indice = 1 Then
        If cantidad > 1 Then
            UserMacro.Intervalo = gIntervals.BuildWork
            UserMacro.cantidad = cantidad
            UserMacro.TIPO = 3
            UserMacro.Index = SastreRopas(lstArmas.ListIndex + 1).Index
            AddtoRichTextBox frmMain.RecTxt, "Comienzas a trabajar.", 2, 51, 223, 1, 1
            UserMacro.Activado = True
            frmMain.MacroLadder.Interval = gIntervals.BuildWork
            frmMain.MacroLadder.Enabled = True
        Else
            Call WriteCraftSastre(SastreRopas(lstArmas.ListIndex + 1).Index)

            If frmMain.macrotrabajo.Enabled Then MacroBltIndex = SastreRopas(lstArmas.ListIndex + 1).Index

        End If

        Unload Me
    ElseIf indice = 2 Then

        If cantidad > 1 Then
            UserMacro.cantidad = cantidad
            UserMacro.TIPO = 3
            UserMacro.Index = SastreGorros(lstArmas.ListIndex + 1).Index
            AddtoRichTextBox frmMain.RecTxt, "Comienzas a trabajar.", 2, 51, 223, 1, 1
            UserMacro.Intervalo = gIntervals.BuildWork
            UserMacro.Activado = True
            frmMain.MacroLadder.Interval = gIntervals.BuildWork
            frmMain.MacroLadder.Enabled = True
        Else
            Call WriteCraftSastre(SastreGorros(lstArmas.ListIndex + 1).Index)

            If frmMain.macrotrabajo.Enabled Then MacroBltIndex = SastreGorros(lstArmas.ListIndex + 1).Index

        End If

    End If

    Unload Me

    
    Exit Sub

Command3_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmSastre.Command3_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command4_Click()
    
    On Error GoTo Command4_Click_Err
    
    Unload Me

    
    Exit Sub

Command4_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmSastre.Command4_Click", Erl)
    Resume Next
    
End Sub


Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Command4_MouseMove_Err
    

    If Command4.Tag = "0" Then
        Command4.Picture = LoadInterface("boton-cancelar-over.bmp")
        Command4.Tag = "1"

    End If

    Command3.Picture = Nothing
    Command3.Tag = "0"

    
    Exit Sub

Command4_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmSastre.Command4_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call FormParser.Parse_Form(Me)
    indice = 1

    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmSastre.Form_Load", Erl)
    Resume Next
    
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Form_KeyPress_Err
    

    If (KeyAscii = 27) Then
        Unload Me

    End If

    
    Exit Sub

Form_KeyPress_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmSastre.Form_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Command3_MouseMove_Err
    

    If Command3.Tag = "0" Then
        Command3.Picture = LoadInterface("boton-elaborar-over.bmp")
        Command3.Tag = "1"

    End If

    Command4.Picture = Nothing
    Command4.Tag = "0"

    
    Exit Sub

Command3_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmSastre.Command3_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Form_MouseMove_Err
    

    If indice <> 1 Then
        Command1.Picture = Nothing
        Command1.Tag = "0"

    End If

    If indice <> 2 Then
        Command2.Picture = Nothing
        Command2.Tag = "0"

    End If

    Command3.Picture = Nothing
    Command3.Tag = "0"
    Command4.Picture = Nothing
    Command4.Tag = "0"

    
    Exit Sub

Form_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmSastre.Form_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Image1_Click()
    On Error GoTo Command4_Click_Err
    
    Unload Me

    
    Exit Sub

Command4_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmSastre.Command4_Click", Erl)
    Resume Next
End Sub

Private Sub Image2_Click()
    If cantidad > 0 Then
        cantidad = cantidad - 1
    Else
        Exit Sub
    End If


End Sub

Private Sub Image3_Click()
    cantidad = cantidad + 1
End Sub

Private Sub List1_Click()
    
    On Error GoTo List1_Click_Err
    

    

    Dim grh As Long

    If List1.ListIndex = 0 Then
        grh = 697
    ElseIf List1.ListIndex = 1 Then
        grh = 699
    ElseIf List1.ListIndex = 2 Then
        grh = 698
    ElseIf List1.ListIndex = 3 Then
        Grh = ObjData(BlackWolfIndex).GrhIndex
    ElseIf List1.ListIndex = 4 Then
        Grh = ObjData(PielTigreIndex).GrhIndex
    ElseIf List1.ListIndex = 5 Then
        Grh = ObjData(PielTigreBengalaIndex).GrhIndex
        

    End If

    Call Grh_Render_To_Hdc(picture1, grh, 0, 0, False)

    
    Exit Sub

List1_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmSastre.List1_Click", Erl)
    Resume Next
    
End Sub

Private Sub lstArmas_Click()
    
    On Error GoTo lstArmas_Click_Err
    

    Dim Obj As ObjDatas
    
    If indice = 1 Then
        Obj = ObjData(SastreRopas(lstArmas.ListIndex + 1).Index)
    ElseIf indice = 2 Then
        Obj = ObjData(SastreGorros(lstArmas.ListIndex + 1).Index)
    Else
        Exit Sub
    End If

    List1.Clear
    List2.Clear
        
        List1.AddItem "Piel de lobo"
        List2.AddItem Obj.PielLobo
    
        List1.AddItem "Piel de oso pardo"
        List2.AddItem Obj.PielOsoPardo
  
        List1.AddItem "Piel de oso polar"
        List2.AddItem Obj.PielOsoPolar
       
        List1.AddItem "Piel de lobo negro"
        List2.AddItem Obj.PielLoboNegro
      
        List1.AddItem "Piel de Tigre"
        List2.AddItem Obj.PielTigre
    
        List1.AddItem "Piel Tigre Bengala"
        List2.AddItem Obj.PielTigreBengala
      
    Call Grh_Render_To_Hdc(picture1, Obj.GrhIndex, 0, 0)
    
    desc.Caption = "Defensa: " & Obj.MinDef & "/" & Obj.MaxDef

    
    Exit Sub

lstArmas_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmSastre.lstArmas_Click", Erl)
    Resume Next
    
End Sub
