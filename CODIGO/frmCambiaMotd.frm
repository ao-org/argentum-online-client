VERSION 5.00
Begin VB.Form frmCambiaMotd 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   """ZMOTD"""
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkItalic 
      Caption         =   "Cursiva"
      Height          =   255
      Left            =   2880
      TabIndex        =   12
      Top             =   4320
      Width           =   855
   End
   Begin VB.CheckBox chkBold 
      Caption         =   "Negrita"
      Height          =   255
      Left            =   1080
      TabIndex        =   11
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton cmdMarron 
      Caption         =   "Marron"
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdVerde 
      Caption         =   "Verde"
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdMorado 
      Caption         =   "Morado"
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdAmarillo 
      Caption         =   "Amarillo"
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdGris 
      Caption         =   "Gris"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdBlanco 
      Caption         =   "Blanco"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdRojo 
      Caption         =   "Rojo"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdAzul 
      BackColor       =   &H00FF0000&
      Caption         =   "Azul"
      Height          =   375
      Left            =   600
      MaskColor       =   &H00FF0000&
      TabIndex        =   3
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   4680
      Width           =   3855
   End
   Begin VB.TextBox txtMotd 
      Height          =   2415
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   660
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "No olvides agregar los colores al final de cada línea (ver tabla de abajo)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   60
      Width           =   4455
   End
End
Attribute VB_Name = "frmCambiaMotd"
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
Option Explicit

Private Sub cmdOk_Click()
    On Error GoTo cmdOk_Click_Err
    Dim t() As String
    Dim i   As Long, n As Long, Pos As Long
    If Len(txtMotd.text) >= 2 Then
        If Right$(txtMotd.text, 2) = vbCrLf Then txtMotd.text = Left$(txtMotd.text, Len(txtMotd.text) - 2)
    End If
    t = Split(txtMotd.text, vbCrLf)
    'hola~1~1~1~1~1
    For i = LBound(t) To UBound(t)
        n = 0
        Pos = InStr(1, t(i), "~")
        Do While Pos > 0 And Pos < Len(t(i))
            n = n + 1
            Pos = InStr(Pos + 1, t(i), "~")
        Loop
        If n <> 5 Then
            MsgBox JsonLanguage.Item("MENSAJEBOX_ERROR_FORMATO_LINEA") & i + 1 & "."
            Exit Sub
        End If
    Next i
    Call WriteSetMOTD(txtMotd.text)
    Unload Me
    Exit Sub
cmdOk_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCambiaMotd.cmdOk_Click", Erl)
    Resume Next
End Sub

'A partir de Command2_Click son todos buttons para agregar color al texto
Private Sub cmdAzul_Click()
    On Error GoTo cmdAzul_Click_Err
    txtMotd.text = txtMotd & "~50~70~250~" & CStr(chkBold.value) & "~" & CStr(chkItalic.value)
    Exit Sub
cmdAzul_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCambiaMotd.cmdAzul_Click", Erl)
    Resume Next
End Sub

Private Sub cmdRojo_Click()
    On Error GoTo cmdRojo_Click_Err
    txtMotd.text = txtMotd & "~255~0~0~" & CStr(chkBold.value) & "~" & CStr(chkItalic.value)
    Exit Sub
cmdRojo_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCambiaMotd.cmdRojo_Click", Erl)
    Resume Next
End Sub

Private Sub cmdBlanco_Click()
    On Error GoTo cmdBlanco_Click_Err
    txtMotd.text = txtMotd & "~255~255~255~" & CStr(chkBold.value) & "~" & CStr(chkItalic.value)
    Exit Sub
cmdBlanco_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCambiaMotd.cmdBlanco_Click", Erl)
    Resume Next
End Sub

Private Sub cmdGris_Click()
    On Error GoTo cmdGris_Click_Err
    txtMotd.text = txtMotd & "~157~157~157~" & CStr(chkBold.value) & "~" & CStr(chkItalic.value)
    Exit Sub
cmdGris_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCambiaMotd.cmdGris_Click", Erl)
    Resume Next
End Sub

Private Sub cmdAmarillo_Click()
    On Error GoTo cmdAmarillo_Click_Err
    txtMotd.text = txtMotd & "~244~244~0~" & CStr(chkBold.value) & "~" & CStr(chkItalic.value)
    Exit Sub
cmdAmarillo_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCambiaMotd.cmdAmarillo_Click", Erl)
    Resume Next
End Sub

Private Sub cmdMorado_Click()
    On Error GoTo cmdMorado_Click_Err
    txtMotd.text = txtMotd & "~128~0~128~" & CStr(chkBold.value) & "~" & CStr(chkItalic.value)
    Exit Sub
cmdMorado_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCambiaMotd.cmdMorado_Click", Erl)
    Resume Next
End Sub

Private Sub cmdVerde_Click()
    On Error GoTo cmdVerde_Click_Err
    txtMotd.text = txtMotd & "~23~104~26~" & CStr(chkBold.value) & "~" & CStr(chkItalic.value)
    Exit Sub
cmdVerde_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCambiaMotd.cmdVerde_Click", Erl)
    Resume Next
End Sub

Private Sub cmdMarron_Click()
    On Error GoTo cmdMarron_Click_Err
    txtMotd.text = txtMotd & "~97~58~31~" & CStr(chkBold.value) & "~" & CStr(chkItalic.value)
    Exit Sub
cmdMarron_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCambiaMotd.cmdMarron_Click", Erl)
    Resume Next
End Sub
