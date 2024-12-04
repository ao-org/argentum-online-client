VERSION 5.00
Begin VB.Form frmAlqui 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Trabajar con alquimista"
   ClientHeight    =   7080
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   4
      Top             =   1890
      Width           =   480
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
      Left            =   3840
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "1"
      Top             =   5670
      Width           =   660
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
      ItemData        =   "frmAlqui.frx":0000
      Left            =   5520
      List            =   "frmAlqui.frx":0007
      TabIndex        =   2
      Top             =   3840
      Width           =   435
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
      Left            =   4080
      TabIndex        =   1
      Top             =   3840
      Width           =   1245
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
      Height          =   3540
      Left            =   585
      TabIndex        =   0
      Top             =   1480
      Width           =   2910
   End
   Begin VB.Image salir 
      Height          =   375
      Left            =   6000
      Top             =   0
      Width           =   495
   End
   Begin VB.Image cmdMasMenos 
      Height          =   315
      Index           =   0
      Left            =   3300
      Tag             =   "0"
      Top             =   5640
      Width           =   315
   End
   Begin VB.Image cmdMasMenos 
      Height          =   315
      Index           =   1
      Left            =   4680
      Tag             =   "0"
      Top             =   5640
      Width           =   315
   End
   Begin VB.Label desc 
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
      Height          =   735
      Left            =   4080
      TabIndex        =   5
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Image Command3 
      Height          =   400
      Left            =   3530
      Tag             =   "0"
      Top             =   6350
      Width           =   1950
   End
   Begin VB.Image Command4 
      Height          =   400
      Left            =   1080
      Tag             =   "0"
      Top             =   6350
      Width           =   1950
   End
End
Attribute VB_Name = "frmAlqui"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Argentum 20 Game Client
'
'    Copyright (C) 2023 Noland Studios LTD
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
'
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
'    This program was based on Argentum Online 0.11.6
'    Copyright (C) 2002 Márquez Pablo Ignacio
'
'    Argentum Online is based on Baronsoft's VB6 Online RPG
'    You can contact the original creator of ORE at aaron@baronsoft.com
'    for more information about ORE please visit http://www.baronsoft.com/
'
'
'

Private Sub cmdMasMenos_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo cmdMasMenos_MouseDown_Err
    

    Call ao20audio.PlayWav(SND_CLICK)

    Select Case Index

        Case 0
            cmdMasMenos(Index).Picture = LoadInterface("boton-sm-menos-off.bmp")
            cantidad.Text = str((Val(cantidad.Text) - 1))
            m_Increment = -1

        Case 1
            cmdMasMenos(Index).Picture = LoadInterface("boton-sm-mas-off.bmp")
            cantidad.Text = str((Val(cantidad.Text) + 1))
            m_Increment = 1

    End Select

    tmrNumber.Interval = 10
    tmrNumber.enabled = True

    
    Exit Sub

cmdMasMenos_MouseDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmBancoCuenta.cmdMasMenos_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub Command3_Click()
    
    On Error GoTo Command3_Click_Err
    

    
    
    If lstArmas.ListIndex < 0 Then
        MsgBox "Debes seleccionar un objeto de la lista"
        Exit Sub

    End If

    If cantidad > 1 Then
        UserMacro.cantidad = cantidad
        UserMacro.TIPO = 1
        UserMacro.Index = ObjAlquimista(lstArmas.ListIndex + 1)
        AddtoRichTextBox frmMain.RecTxt, JsonLanguage.Item("MENSAJE_COMIENZAS_A_TRABAJAR"), 2, 51, 223, 1, 1
        UserMacro.Intervalo = IntervaloTrabajo
        UserMacro.Activado = True
        frmMain.MacroLadder.Interval = gIntervals.BuildWork
        frmMain.MacroLadder.enabled = True
    Else
        Call WriteCraftAlquimista(ObjAlquimista(lstArmas.ListIndex + 1))

        If frmMain.macrotrabajo.enabled Then MacroBltIndex = ObjAlquimista(lstArmas.ListIndex + 1)
    
    End If

    Unload Me

    
    Exit Sub

Command3_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmAlqui.Command3_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command4_Click()
    
    On Error GoTo Command4_Click_Err
    
    Unload Me

    
    Exit Sub

Command4_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmAlqui.Command4_Click", Erl)
    Resume Next
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Form_KeyPress_Err
    

    If (KeyAscii = 27) Then
        Unload Me

    End If

    
    Exit Sub

Form_KeyPress_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmAlqui.Form_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call FormParser.Parse_Form(Me)

    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmAlqui.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    Command4.Picture = Nothing
    Command4.Tag = "0"
    Command3.Picture = Nothing
    Command3.Tag = "0"

    
    Exit Sub

Form_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmAlqui.Form_MouseMove", Erl)
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
    Call Grh_Render_To_Hdc(picture1, 21926, 0, 0, False)
    
    Exit Sub

List1_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmAlqui.List1_Click", Erl)
    Resume Next
    
End Sub

Private Sub Command3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    ' Command3.Picture = LoadInterface("trabajar_construirpress.bmp")
    '  Command3.Tag = "1"
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
    Call RegistrarError(Err.Number, Err.Description, "frmAlqui.Command3_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Command4_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    '                Command4.Picture = LoadInterface("trabajar_salirpress.bmp")
    '                Command4.Tag = "1"
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
    Call RegistrarError(Err.Number, Err.Description, "frmAlqui.Command4_MouseMove", Erl)
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
    Call frmAlqui.List1.Clear
    Call frmAlqui.List2.Clear
    
    If (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).Raices) > 0 Then
        frmAlqui.List1.AddItem ("Raices")
        frmAlqui.List2.AddItem (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).Raices)
    End If
    If (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).Botella) > 0 Then
        frmAlqui.List1.AddItem ("Botella")
        frmAlqui.List2.AddItem (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).Botella)
    End If
    If (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).Cuchara) > 0 Then
        frmAlqui.List1.AddItem ("Cuchara")
        frmAlqui.List2.AddItem (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).Cuchara)
    End If
    
    If (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).Mortero) > 0 Then
        frmAlqui.List1.AddItem ("Mortero")
        frmAlqui.List2.AddItem (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).Mortero)
    End If
    
    If (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).FrascoAlq) > 0 Then
        frmAlqui.List1.AddItem ("FrascoAlq")
        frmAlqui.List2.AddItem (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).FrascoAlq)
    End If
    
    If (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).FrascoElixir) > 0 Then
        frmAlqui.List1.AddItem ("FrascoElixir")
        frmAlqui.List2.AddItem (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).FrascoElixir)
    End If
    
    If (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).Dosificador) > 0 Then
        frmAlqui.List1.AddItem ("Dosificador")
        frmAlqui.List2.AddItem (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).Dosificador)
    End If
    
    If (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).Orquidea) > 0 Then
        frmAlqui.List1.AddItem ("Orquidea")
        frmAlqui.List2.AddItem (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).Orquidea)
    End If
    
    If (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).Carmesi) > 0 Then
        frmAlqui.List1.AddItem ("Carmesi")
        frmAlqui.List2.AddItem (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).Carmesi)
    End If
    
    If (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).HongoDeLuz) > 0 Then
        frmAlqui.List1.AddItem ("HongoDeLuz")
        frmAlqui.List2.AddItem (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).HongoDeLuz)
    End If
    
    If (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).Esporas) > 0 Then
        frmAlqui.List1.AddItem ("Esporas")
        frmAlqui.List2.AddItem (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).Esporas)
    End If

    If (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).Tuna) > 0 Then
        frmAlqui.List1.AddItem ("Tuna")
        frmAlqui.List2.AddItem (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).Tuna)
    End If

    If (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).Cala) > 0 Then
        frmAlqui.List1.AddItem ("Cala")
        frmAlqui.List2.AddItem (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).Cala)
    End If
    
    If (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).ColaDeZorro) > 0 Then
        frmAlqui.List1.AddItem ("ColaDeZorro")
        frmAlqui.List2.AddItem (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).ColaDeZorro)
    End If
    
    If (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).FlorOceano) > 0 Then
        frmAlqui.List1.AddItem ("FlorOceano")
        frmAlqui.List2.AddItem (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).FlorOceano)
    End If
    
    If (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).FlorRoja) > 0 Then
        frmAlqui.List1.AddItem ("FlorRoja")
        frmAlqui.List2.AddItem (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).FlorRoja)
    End If
    
    If (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).Hierva) > 0 Then
        frmAlqui.List1.AddItem ("Hierva")
        frmAlqui.List2.AddItem (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).Hierva)
    End If
    
    If (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).HojasDeRin) > 0 Then
        frmAlqui.List1.AddItem ("HojasDeRin")
        frmAlqui.List2.AddItem (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).HojasDeRin)
    End If
    
    If (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).HojasRojas) > 0 Then
        frmAlqui.List1.AddItem ("HojasRojas")
        frmAlqui.List2.AddItem (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).HojasRojas)
    End If
    
    If (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).SemillasPros) > 0 Then
        frmAlqui.List1.AddItem ("SemillasPros")
        frmAlqui.List2.AddItem (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).SemillasPros)
    End If
    
    If (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).Pimiento) > 0 Then
        frmAlqui.List1.AddItem ("Pimiento")
        frmAlqui.List2.AddItem (ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).Pimiento)
    End If
    
    desc.Caption = ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).Texto

    Call Grh_Render_To_Hdc(picture1, ObjData(ObjAlquimista(lstArmas.ListIndex + 1)).GrhIndex, 0, 0, False)
    
    
    Exit Sub

lstArmas_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmAlqui.lstArmas_Click", Erl)
    Resume Next
    
End Sub

Private Sub salir_Click()
    On Error GoTo salir_Click_Err
    
    Unload Me

    
    Exit Sub

salir_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmBancoCuenta.salir_Click", Erl)
    Resume Next
End Sub
