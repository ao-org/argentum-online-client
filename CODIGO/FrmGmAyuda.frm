VERSION 5.00
Begin VB.Form FrmGmAyuda 
   BorderStyle     =   0  'None
   Caption         =   "Formulario de mensaje a administradores"
   ClientHeight    =   6528
   ClientLeft      =   3228
   ClientTop       =   1332
   ClientWidth     =   7044
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   544
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   587
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdButtonEntrarForo 
      Caption         =   "Reportar bug en foro [Sugerido]"
      Height          =   495
      Left            =   5280
      MaskColor       =   &H00404080&
      TabIndex        =   8
      Top             =   840
      Width           =   1695
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "Consulta regular"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   7080
      TabIndex        =   7
      Top             =   1320
      Width           =   1695
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "Descargo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   8700
      TabIndex        =   6
      Top             =   1410
      Width           =   975
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "Acusación"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   9900
      TabIndex        =   5
      Top             =   1410
      Width           =   1095
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "Sugerencia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   8700
      TabIndex        =   4
      Top             =   1650
      Width           =   1095
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "Reporte de bug"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   7080
      TabIndex        =   3
      Top             =   1650
      Width           =   1455
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "Otro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   7080
      TabIndex        =   2
      Top             =   1890
      Width           =   855
   End
   Begin VB.OptionButton optConsulta 
      Caption         =   "Queja"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   8700
      TabIndex        =   1
      Top             =   1890
      Width           =   1755
   End
   Begin VB.TextBox txtMotivo 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
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
      Height          =   2175
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Image cmdCerrar 
      Height          =   420
      Left            =   6585
      Tag             =   "0"
      Top             =   0
      Width           =   420
   End
   Begin VB.Image cmdCancelar 
      Height          =   360
      Left            =   1335
      Tag             =   "0"
      Top             =   5715
      Width           =   1920
   End
   Begin VB.Image cmdEnviarMensaje 
      Height          =   360
      Left            =   3855
      Tag             =   "0"
      Top             =   5715
      Width           =   1920
   End
   Begin VB.Image optConsult 
      Height          =   210
      Index           =   5
      Left            =   3915
      Tag             =   "0"
      Top             =   3765
      Width           =   225
   End
   Begin VB.Image optConsult 
      Height          =   210
      Index           =   2
      Left            =   3915
      Tag             =   "0"
      Top             =   3405
      Width           =   225
   End
   Begin VB.Image optConsult 
      Height          =   210
      Index           =   6
      Left            =   3915
      Tag             =   "0"
      Top             =   3045
      Width           =   225
   End
   Begin VB.Image optConsult 
      Height          =   210
      Index           =   3
      Left            =   3915
      Tag             =   "0"
      Top             =   2670
      Width           =   225
   End
   Begin VB.Image optConsult 
      Height          =   210
      Index           =   1
      Left            =   3915
      Tag             =   "0"
      Top             =   2310
      Width           =   225
   End
   Begin VB.Image optConsult 
      Height          =   210
      Index           =   4
      Left            =   3915
      Tag             =   "0"
      Top             =   1950
      Width           =   225
   End
   Begin VB.Image optConsult 
      Height          =   210
      Index           =   0
      Left            =   3915
      Tag             =   "0"
      Top             =   1590
      Width           =   225
   End
End
Attribute VB_Name = "FrmGmAyuda"
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

Private cBotonEnviarMensaje As clsGraphicalButton
Private cBotonCancelar As clsGraphicalButton
Private cBotonCerrar As clsGraphicalButton

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdButtonEntrarForo_Click()
    Call ShellExecute(0, "Open", "https://www.elmesonhostigado.com/foro", "", App.Path, 1)
End Sub

Private Sub Form_Load()
    
    On Error GoTo Form_Load_Err
    
    Call Aplicar_Transparencia(Me.hwnd, 240)

    'Call FormParser.Parse_Form(Me)
    Me.Picture = LoadInterface("ventanagm.bmp")
    
    Call LoadButtons

    
    Exit Sub

Form_Load_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmGmAyuda.Form_Load", Erl)
    Resume Next
    
End Sub
Private Sub LoadButtons()
    Set cBotonEnviarMensaje = New clsGraphicalButton
    Set cBotonCancelar = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton
    
    Call cBotonEnviarMensaje.Initialize(cmdEnviarMensaje, "boton-enviar-default.bmp", _
                                                "boton-enviar-over.bmp", _
                                                "boton-enviar-off.bmp", Me)
    
    Call cBotonCancelar.Initialize(cmdCancelar, "boton-cancelar-default.bmp", _
                                                "boton-cancelar-over.bmp", _
                                                "boton-cancelar-off.bmp", Me)
                                                
    Call cBotonCerrar.Initialize(cmdCerrar, "boton-cerrar-default.bmp", _
                                                "boton-cerrar-over.bmp", _
                                                "boton-cerrar-off.bmp", Me)
                                                
End Sub

Private Sub cmdEnviarMensaje_Click()
    
    On Error GoTo cmdEnviarMensaje_Click_Err
    

    If txtMotivo.Text = "" Then
        Call AddtoRichTextBox(frmMain.RecTxt, "Debes escribir tu mensaje.", 255, 255, 255, False, False, False)
        Exit Sub
    ElseIf DarIndiceElegido = -1 Then
        Call AddtoRichTextBox(frmMain.RecTxt, "Debes elegir el motivo de tu consulta.", 255, 255, 255, False, False, False)
        Exit Sub
    Else
        Call WriteQuestionGM(txtMotivo.Text, optConsulta(DarIndiceElegido).Caption)
        Unload Me

    End If

    
    Exit Sub

cmdEnviarMensaje_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmGmAyuda.cmdEnviarMensaje_Click", Erl)
    Resume Next
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoverForm Me.hwnd
End Sub

Private Sub optConsult_Click(Index As Integer)
    
    On Error GoTo optConsult_Click_Err
    

    Dim i As Integer

    For i = 0 To 6

        If i <> Index Then
            optConsult(i).Picture = Nothing
            optConsult(i).Tag = 0
        Else
            optConsult(i).Picture = LoadInterface("radio-on.bmp")
            optConsult(i).Tag = 1

        End If

    Next i

    Select Case Index

        Case 0
            Call AddtoRichTextBox(frmMain.RecTxt, "¡Por favor explique correctamente el motivo de su consulta!", 255, 255, 255, False, False, False)

        Case 1
            Call AddtoRichTextBox(frmMain.RecTxt, "Deje el nombre del personaje del que está pidiendo descargo por una medida, conjunto con el administrador que está relacionado con ella.", 255, 255, 255, False, False, False)

        Case 2
            Call AddtoRichTextBox(frmMain.RecTxt, "Se dará prioridad a su consulta enviando un mensaje a los administradores conectados, por favor utilize ésta opción responsablemente.", 255, 255, 255, False, False, False)

        Case 3
            Call AddtoRichTextBox(frmMain.RecTxt, "Su sugerencia SERÁ leída por un miembro del staff, y será tomada en cuenta para futuros cambios.", 255, 255, 255, False, False, False)

        Case 4
            Call AddtoRichTextBox(frmMain.RecTxt, "Explique de la forma más detallada la forma de repetir el error. El staff de programación lo resolverá lo antes posible.", 255, 255, 255, False, False, False)

        Case 5
            Call AddtoRichTextBox(frmMain.RecTxt, "Deje la mayor cantidad de datos posibles, esta opción es para consultas que no entran en otras secciónes.", 255, 255, 255, False, False, False)

        Case 6
            Call AddtoRichTextBox(frmMain.RecTxt, "En caso de ser una queja hacia un miembro del staff, deje principalmente hacia quien esta referida la queja, y los motivos dejando todos los detalles posibles", 255, 255, 255, False, False, False)

    End Select

    
    Exit Sub

optConsult_Click_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmGmAyuda.optConsult_Click", Erl)
    Resume Next
    
End Sub

Private Function DarIndiceElegido() As Integer
    
    On Error GoTo DarIndiceElegido_Err
    

    Dim i As Integer

    For i = 0 To 6

        If optConsult(i).Tag = 1 Then
            DarIndiceElegido = i
            Exit Function

        End If

    Next i

    DarIndiceElegido = -1

    
    Exit Function

DarIndiceElegido_Err:
    Call RegistrarError(Err.number, Err.Description, "FrmGmAyuda.DarIndiceElegido", Erl)
    Resume Next
    
End Function

