VERSION 5.00
Begin VB.Form FrmKeyInv 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3600
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   235
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox interface 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   495
      MousePointer    =   99  'Custom
      ScaleHeight     =   88
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   219
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1245
      Width           =   2625
   End
   Begin VB.Label NombreLlave 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   240
      TabIndex        =   1
      Top             =   2295
      Width           =   3135
   End
   Begin VB.Image cmdCerrar 
      Height          =   420
      Left            =   3150
      Tag             =   "0"
      Top             =   15
      Width           =   465
   End
End
Attribute VB_Name = "FrmKeyInv"
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

Const WM_SYSCOMMAND As Long = &H112&
Const MOUSE_MOVE    As Long = &HF012&
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long

Public WithEvents InvKeys As clsGrapchicalInventory
Attribute InvKeys.VB_VarHelpID = -1

Private Sub cmdcerrar_Click()
    
    On Error GoTo cmdcerrar_Click_Err
    
    frmMain.CerrarLlavero
    
    Exit Sub

cmdcerrar_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmKeyInv.cmdcerrar_Click", Erl)
    Resume Next
    
End Sub

Private Sub cmdCerrar_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo cmdCerrar_MouseDown_Err
    
    cmdCerrar.Picture = LoadInterface("boton-cerrar-off.bmp")
    cmdCerrar.Tag = "1"
    
    Exit Sub

cmdCerrar_MouseDown_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmKeyInv.cmdCerrar_MouseDown", Erl)
    Resume Next
    
End Sub

Private Sub cmdCerrar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo cmdCerrar_MouseMove_Err
    
    If cmdCerrar.Tag = "0" Then
        cmdCerrar.Picture = LoadInterface("boton-cerrar-over.bmp")
        cmdCerrar.Tag = "1"
    End If
    
    Exit Sub

cmdCerrar_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmKeyInv.cmdCerrar_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub Form_Activate()
    
    On Error GoTo Form_Activate_Err
    
    If InvKeys.OBJIndex(1) = 0 Then
        NombreLlave.Caption = "Aquí aparecerán las llaves que consigas"
    End If
    
    Exit Sub

Form_Activate_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmKeyInv.Form_Activate", Erl)
    Resume Next
    
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    
    On Error GoTo Form_KeyPress_Err
    

    If (KeyAscii = 27) Then
        Unload Me

    End If

    
    Exit Sub

Form_KeyPress_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmKeyInv.Form_KeyPress", Erl)
    Resume Next
    
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call FormParser.Parse_Form(Me)
    Me.Picture = LoadInterface("ventanallavero.bmp")
    cmdCerrar.Picture = Nothing
    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmKeyInv.Form_Load", Erl)
    Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo Form_MouseMove_Err
    
    If InvKeys.OBJIndex(1) <> 0 Then
        NombreLlave.Caption = vbNullString
    End If
    
    ReleaseCapture
    Call SendMessage(Me.hwnd, WM_SYSCOMMAND, MOUSE_MOVE, 0)
    
    If cmdCerrar.Tag = "1" Then
        cmdCerrar.Picture = Nothing
        cmdCerrar.Tag = "0"
    End If
    
    Exit Sub

Form_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmKeyInv.Form_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub interface_DblClick()
    
    On Error GoTo interface_DblClick_Err
    
    If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Exit Sub

    If InvKeys.IsItemSelected Then
        Call WriteUseKey(InvKeys.SelectedItem)
    End If
    
    Exit Sub

interface_DblClick_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmKeyInv.interface_DblClick", Erl)
    Resume Next
    
End Sub

Private Sub interface_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo interface_MouseMove_Err
    
    Dim Slot As Integer
    Slot = InvKeys.GetSlot(x, y)
    
    If Slot <> 0 Then
        If InvKeys.OBJIndex(Slot) <> 0 Then
            NombreLlave.Caption = InvKeys.ItemName(Slot)
        End If
    End If
    
    If cmdCerrar.Tag = "1" Then
        cmdCerrar.Picture = Nothing
        cmdCerrar.Tag = "0"
    End If
    
    Exit Sub

interface_MouseMove_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmKeyInv.interface_MouseMove", Erl)
    Resume Next
    
End Sub

Private Sub interface_Paint()
    
    On Error GoTo interface_Paint_Err
    
    InvKeys.ReDraw
    
    Exit Sub

interface_Paint_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmKeyInv.interface_Paint", Erl)
    Resume Next
    
End Sub
