VERSION 5.00
Begin VB.Form frmHerrero 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Trabajar con Herreria"
   ClientHeight    =   5610
   ClientLeft      =   0
   ClientTop       =   -90
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Height          =   2955
      Left            =   720
      TabIndex        =   5
      Top             =   2150
      Width           =   2400
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
      Height          =   810
      Left            =   3840
      TabIndex        =   4
      Top             =   3000
      Width           =   1605
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
      Height          =   810
      ItemData        =   "frmHerrero.frx":0000
      Left            =   5535
      List            =   "frmHerrero.frx":0007
      TabIndex        =   3
      Top             =   3000
      Width           =   525
   End
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
      Left            =   3820
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "1"
      Top             =   4520
      Width           =   660
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   4750
      ScaleHeight     =   32
      ScaleMode       =   0  'User
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   1760
      Width           =   480
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
      TabIndex        =   1
      Top             =   3840
      Width           =   2055
   End
   Begin VB.Image Command6 
      Height          =   450
      Left            =   4590
      Top             =   4420
      Width           =   1740
   End
   Begin VB.Image Command5 
      Height          =   465
      Left            =   3960
      Top             =   5000
      Width           =   2130
   End
   Begin VB.Image Command2 
      Height          =   510
      Left            =   2640
      Stretch         =   -1  'True
      Top             =   1460
      Width           =   570
   End
   Begin VB.Image Command1 
      Height          =   600
      Left            =   1980
      OLEDropMode     =   1  'Manual
      Top             =   1450
      Width           =   615
   End
   Begin VB.Image Command4 
      Height          =   600
      Left            =   1360
      Tag             =   "0"
      Top             =   1440
      Width           =   615
   End
   Begin VB.Image Command3 
      Height          =   600
      Left            =   680
      Top             =   1460
      Width           =   615
   End
End
Attribute VB_Name = "frmHerrero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez
Dim Index As Byte

Option Explicit

Private Sub Command1_Click()
    
    On Error GoTo Command1_Click_Err
    
    Index = 1

    Dim i As Byte

    lstArmas.Clear

    For i = 1 To UBound(ArmasHerrero())

        If ArmasHerrero(i).Index = 0 Then Exit For
        Call frmHerrero.lstArmas.AddItem(ObjData(ArmasHerrero(i).Index).Name)
    Next i
    
    Command1.Picture = LoadInterface("herreria_armashover.bmp")
    Command3.Picture = Nothing
    Command2.Picture = Nothing
    Command4.Picture = Nothing

    
    Exit Sub

Command1_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmHerrero.Command1_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' Command1.Picture = LoadInterface("herreria_armaspress.bmp")
    ' Command1.Tag = "1"
End Sub

Private Sub Command2_Click()
    
    On Error GoTo Command2_Click_Err
    
    Index = 2

    Dim i As Byte

    lstArmas.Clear

    For i = 0 To UBound(ArmadurasHerrero())

        If ArmadurasHerrero(i).Index = 0 Then Exit For
        If ObjData(ArmadurasHerrero(i).Index).ObjType = 3 Then
            Call frmHerrero.lstArmas.AddItem(ObjData(ArmadurasHerrero(i).Index).Name)

        End If

    Next i

    Command2.Picture = LoadInterface("herreria_armadurashover.bmp")
    Command1.Picture = Nothing
    Command3.Picture = Nothing
    Command4.Picture = Nothing

    
    Exit Sub

Command2_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmHerrero.Command2_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command3_Click()
    
    On Error GoTo Command3_Click_Err
    
    Index = 3

    lstArmas.Clear

    Dim i As Byte

    For i = 0 To UBound(CascosHerrero())

        If CascosHerrero(i).Index = 0 Then Exit For
        Call frmHerrero.lstArmas.AddItem(ObjData(CascosHerrero(i).Index).Name)
    Next i

    Command3.Picture = LoadInterface("herreria_cascoshover.bmp")

    Command1.Picture = Nothing
    Command2.Picture = Nothing
    Command4.Picture = Nothing

    
    Exit Sub

Command3_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmHerrero.Command3_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command4_Click()
    
    On Error GoTo Command4_Click_Err
    
    Index = 4

    lstArmas.Clear

    Dim i As Byte

    For i = 0 To UBound(EscudosHerrero())

        If EscudosHerrero(i).Index = 0 Then Exit For
        Call frmHerrero.lstArmas.AddItem(ObjData(EscudosHerrero(i).Index).Name)
    Next i

    Command4.Picture = LoadInterface("herreria_escudoshover.bmp")
    Command1.Picture = Nothing
    Command2.Picture = Nothing
    Command3.Picture = Nothing

    
    Exit Sub

Command4_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmHerrero.Command4_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command5_Click()
    
    On Error GoTo Command5_Click_Err
    
    Unload Me

    
    Exit Sub

Command5_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmHerrero.Command5_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command6_Click()
    
    On Error GoTo Command6_Click_Err
    

    

    If Index = 1 Then

        If cantidad > 1 Then
            UserMacro.cantidad = cantidad
            UserMacro.TIPO = 4
            UserMacro.Index = ArmasHerrero(lstArmas.ListIndex + 1).Index
            AddtoRichTextBox frmMain.RecTxt, "Comienzas a trabajar.", 2, 51, 223, 1, 1
            UserMacro.Intervalo = IntervaloTrabajo
            UserMacro.Activado = True
            frmMain.MacroLadder.Interval = IntervaloTrabajo
            frmMain.MacroLadder.Enabled = True
        Else
            Call WriteCraftBlacksmith(ArmasHerrero(lstArmas.ListIndex + 1).Index)

            If frmMain.macrotrabajo.Enabled Then MacroBltIndex = ArmasHerrero(lstArmas.ListIndex + 1).Index

        End If

        Unload Me
    ElseIf Index = 2 Then
    
        If cantidad > 1 Then
            UserMacro.cantidad = cantidad
            UserMacro.TIPO = 4
            UserMacro.Index = ArmadurasHerrero(lstArmas.ListIndex).Index
            AddtoRichTextBox frmMain.RecTxt, "Comienzas a trabajar.", 2, 51, 223, 1, 1
            UserMacro.Intervalo = IntervaloTrabajo
            UserMacro.Activado = True
            frmMain.MacroLadder.Interval = IntervaloTrabajo
            frmMain.MacroLadder.Enabled = True
        Else
            Call WriteCraftBlacksmith(ArmadurasHerrero(lstArmas.ListIndex).Index)

            If frmMain.macrotrabajo.Enabled Then MacroBltIndex = ArmadurasHerrero(lstArmas.ListIndex).Index
            
        End If
        
        Unload Me
    ElseIf Index = 3 Then
    
        If cantidad > 1 Then
            UserMacro.cantidad = cantidad
            UserMacro.TIPO = 4
            UserMacro.Index = CascosHerrero(lstArmas.ListIndex).Index
            AddtoRichTextBox frmMain.RecTxt, "Comienzas a trabajar.", 2, 51, 223, 1, 1
            UserMacro.Intervalo = IntervaloTrabajo
            UserMacro.Activado = True
            frmMain.MacroLadder.Interval = IntervaloTrabajo
            frmMain.MacroLadder.Enabled = True
        Else
            Call WriteCraftBlacksmith(CascosHerrero(lstArmas.ListIndex).Index)

            If frmMain.macrotrabajo.Enabled Then MacroBltIndex = CascosHerrero(lstArmas.ListIndex).Index
            
        End If
        
        Unload Me
    ElseIf Index = 4 Then
    
        If cantidad > 1 Then
            UserMacro.cantidad = cantidad
            UserMacro.TIPO = 4
            UserMacro.Index = EscudosHerrero(lstArmas.ListIndex).Index
            AddtoRichTextBox frmMain.RecTxt, "Comienzas a trabajar.", 2, 51, 223, 1, 1
            UserMacro.Intervalo = IntervaloTrabajo
            UserMacro.Activado = True
            frmMain.MacroLadder.Interval = IntervaloTrabajo
            frmMain.MacroLadder.Enabled = True
        Else
            Call WriteCraftBlacksmith(EscudosHerrero(lstArmas.ListIndex).Index)

            If frmMain.macrotrabajo.Enabled Then MacroBltIndex = EscudosHerrero(lstArmas.ListIndex).Index
            
        End If
        
        Unload Me

    End If

    
    Exit Sub

Command6_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmHerrero.Command6_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command6_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    'Command6.Picture = LoadInterface("trabajar_construirpress.bmp")
    ' Command6.Tag = "1"
End Sub

Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Command6_MouseMove_Err
    

    If Command6.Tag = "0" Then
        Command6.Picture = LoadInterface("trabajar_construirhover.bmp")
        Command6.Tag = "1"

    End If
    
    Command5.Picture = Nothing
    Command5.Tag = "0"

    
    Exit Sub

Command6_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmHerrero.Command6_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Command5_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    'Command5.Picture = LoadInterface("trabajar_salirpress.bmp")
    'Command5.Tag = "1"
End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Command5_MouseMove_Err
    

    If Command5.Tag = "0" Then
        Command5.Picture = LoadInterface("trabajar_salirhover.bmp")
        Command5.Tag = "1"

    End If

    Command6.Picture = Nothing
    Command6.Tag = "0"

    
    Exit Sub

Command5_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmHerrero.Command5_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Command4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    'Command4.Picture = LoadInterface("herreria_escudospress.bmp")
    'Command4.Tag = "1"
End Sub

Private Sub Command3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' 'Command3.Picture = LoadInterface("herreria_cascospress.bmp")
    ' Command3.Tag = "1"
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    'Command2.Picture = LoadInterface("herreria_armaduraspress.bmp")
    ' Command2.Tag = "1"
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call FormParser.Parse_Form(Me)
    Index = 3

    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "frmHerrero.Form_Load", Erl)
    Resume Next
    
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Form_KeyPress_Err
    

    If (KeyAscii = 27) Then
        Unload Me

    End If

    
    Exit Sub

Form_KeyPress_Err:
    Call RegistrarError(Err.number, Err.Description, "frmHerrero.Form_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Form_MouseMove_Err
    

    Command5.Picture = Nothing
    Command5.Tag = "0"

    Command6.Picture = Nothing
    Command6.Tag = "0"

    
    Exit Sub

Form_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "frmHerrero.Form_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub List1_Click()
    
    On Error GoTo List1_Click_Err
    

    

    Dim SR As RECT, DR As RECT

    SR.Left = 0
    SR.Top = 0
    SR.Right = 32
    SR.Bottom = 32

    DR.Left = 0
    DR.Top = 0
    DR.Right = 32
    DR.Bottom = 32

    Dim grh As Long

    If List1.ListIndex = 0 Then
        grh = 724
    ElseIf List1.ListIndex = 1 Then
        grh = 725
    ElseIf List1.ListIndex = 2 Then
        grh = 723

    End If

    Call Grh_Render_To_Hdc(picture1, grh, 0, 0, False)

    
    Exit Sub

List1_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmHerrero.List1_Click", Erl)
    Resume Next
    
End Sub

Private Sub lstArmas_Click()
    
    On Error GoTo lstArmas_Click_Err
    

    

    Dim SR As RECT, DR As RECT

    SR.Left = 0
    SR.Top = 0
    SR.Right = 32
    SR.Bottom = 32

    DR.Left = 0
    DR.Top = 0
    DR.Right = 32
    DR.Bottom = 32

    List1.Clear
    List2.Clear
    List1.AddItem ("Lingote de hierro")
    List1.AddItem ("Lingo de plata")
    List1.AddItem ("Lingotes de Oro")

    If Index = 1 Then
        Call Grh_Render_To_Hdc(picture1, ObjData(ArmasHerrero(lstArmas.ListIndex + 1).Index).GrhIndex, 0, 0)
        List2.AddItem (ObjData(ArmasHerrero(lstArmas.ListIndex + 1).Index).LingH)
        List2.AddItem (ObjData(ArmasHerrero(lstArmas.ListIndex + 1).Index).LingP)
        List2.AddItem (ObjData(ArmasHerrero(lstArmas.ListIndex + 1).Index).LingO)
        desc.Caption = "Golpe: " & ObjData(ArmasHerrero(lstArmas.ListIndex + 1).Index).MinHit & "/" & ObjData(ArmasHerrero(lstArmas.ListIndex + 1).Index).MaxHit
    
    ElseIf Index = 2 Then
        Call Grh_Render_To_Hdc(picture1, ObjData(ArmadurasHerrero(lstArmas.ListIndex).Index).GrhIndex, 0, 0)
        List2.AddItem (ObjData(ArmadurasHerrero(lstArmas.ListIndex).Index).LingH)
        List2.AddItem (ObjData(ArmadurasHerrero(lstArmas.ListIndex).Index).LingP)
        List2.AddItem (ObjData(ArmadurasHerrero(lstArmas.ListIndex).Index).LingO)
        desc.Caption = "Defensa: " & ObjData(ArmadurasHerrero(lstArmas.ListIndex).Index).MinDef & "/" & ObjData(ArmadurasHerrero(lstArmas.ListIndex).Index).MaxDef
    ElseIf Index = 3 Then
        Call Grh_Render_To_Hdc(picture1, ObjData(CascosHerrero(lstArmas.ListIndex).Index).GrhIndex, 0, 0)
        List2.AddItem (ObjData(CascosHerrero(lstArmas.ListIndex).Index).LingH)
        List2.AddItem (ObjData(CascosHerrero(lstArmas.ListIndex).Index).LingP)
        List2.AddItem (ObjData(CascosHerrero(lstArmas.ListIndex).Index).LingO)
        desc.Caption = "Defensa: " & ObjData(CascosHerrero(lstArmas.ListIndex).Index).MinDef & "/" & ObjData(CascosHerrero(lstArmas.ListIndex).Index).MaxDef
    ElseIf Index = 4 Then
        Call Grh_Render_To_Hdc(picture1, ObjData(EscudosHerrero(lstArmas.ListIndex).Index).GrhIndex, 0, 0)
        List2.AddItem (ObjData(EscudosHerrero(lstArmas.ListIndex).Index).LingH)
        List2.AddItem (ObjData(EscudosHerrero(lstArmas.ListIndex).Index).LingP)
        List2.AddItem (ObjData(EscudosHerrero(lstArmas.ListIndex).Index).LingO)
        desc.Caption = "Defensa: " & ObjData(EscudosHerrero(lstArmas.ListIndex).Index).MinDef & "/" & ObjData(EscudosHerrero(lstArmas.ListIndex).Index).MaxDef

    End If

    picture1.Visible = True

    
    Exit Sub

lstArmas_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "frmHerrero.lstArmas_Click", Erl)
    Resume Next
    
End Sub

