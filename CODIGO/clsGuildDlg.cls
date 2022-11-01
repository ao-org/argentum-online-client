VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGuildDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
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

Private Const p_SEGUNDOS = 10

Private Const p_CANTIDADDIALOGOS As Byte = 7

Private Type t_GuildDlg

    Texto As String
    status As Byte
End Type

Private Segundos As Byte
Private p_Dialogos() As t_GuildDlg
Private offset       As Single
Private p_Activo     As Boolean
Private usedFont     As StdFont
Private Alpha        As Single
Private force_visible As Boolean
Private Const DIALOG_HEIGHT As Integer = 14


''
' Removes all dialogs from screen.

Public Sub RemoveDialogs()
    
    On Error GoTo RemoveDialogs_Err
    

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/02/07
    'Removes all dialogs
    '***************************************************
    Dim i As Long
    
    For i = 1 To p_CANTIDADDIALOGOS
        p_Dialogos(i).Texto = vbNullString
        p_Dialogos(i).status = 0
    Next i

    
    Exit Sub

RemoveDialogs_Err:
    Call RegistrarError(Err.number, Err.Description, "clsGuildDlg.RemoveDialogs", Erl)
    Resume Next
    
End Sub

''
' Retrieves the number of dialogs to be displayed on screen.
'
' @return   The number of dialogs to be displayed on screen.

Public Property Get CantidadDialogos() As Byte
    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/01/07
    'Retrieves the number of dialogs to be displayed on screen
    '***************************************************
    
    On Error GoTo CantidadDialogos_Err
    
    CantidadDialogos = p_CANTIDADDIALOGOS

    
    Exit Property

CantidadDialogos_Err:
    Call RegistrarError(Err.number, Err.Description, "clsGuildDlg.CantidadDialogos", Erl)
    Resume Next
    
End Property

Public Property Let Activo(ByVal v As Boolean)
    
    On Error GoTo Activo_Err
    
    p_Activo = v

    
    Exit Property

Activo_Err:
    Call RegistrarError(Err.number, Err.Description, "clsGuildDlg.Activo", Erl)
    Resume Next
    
End Property

Public Property Get Activo() As Boolean
    
    On Error GoTo Activo_Err
    
    Activo = p_Activo

    
    Exit Property

Activo_Err:
    Call RegistrarError(Err.number, Err.Description, "clsGuildDlg.Activo", Erl)
    Resume Next
    
End Property

Private Sub Class_Initialize()
    
    On Error GoTo Class_Initialize_Err
    
    ReDim p_Dialogos(1 To p_CANTIDADDIALOGOS) As t_GuildDlg
    p_Activo = True
    
    Exit Sub

Class_Initialize_Err:
    Call RegistrarError(Err.number, Err.Description, "clsGuildDlg.Class_Initialize", Erl)
    Resume Next
    
End Sub

Public Sub Draw()
    
    On Error GoTo Draw_Err

    Dim i As Long
    Dim y As Integer
    

    If Not p_Activo Then Exit Sub
    
    Dim Color(3) As RGBA
    Call RGBAList(Color, 222, 194, 112)
    
    Dim dY As Single
    Dim alphaOffset As Integer
    Dim Texto(2) As String
    alphaOffset = DIALOG_HEIGHT * (p_CANTIDADDIALOGOS - 2)
    
    'If p_Dialogos(1).Texto <> vbNullString Then Call Engine_Draw_Box(0, frmMain.renderer.ScaleHeight - 105, 300, 105, RGBA_From_Comp(0, 0, 0, Alpha * 0.3))
    
    For i = 1 To p_CANTIDADDIALOGOS
        If Trim$(p_Dialogos(i).Texto) <> vbNullString Then
            dY = (DIALOG_HEIGHT * i - offset - alphaOffset) / alphaOffset
            y = 88 - DIALOG_HEIGHT * i + offset
             If dy < 0 Then
                Call Engine_Text_Render_No_Ladder(p_Dialogos(i).Texto, 5, -i * 14 + 115 + offset - 22, Color, p_Dialogos(i).status, 1, False, 0, Alpha)
            Else
                Call Engine_Text_Render_No_Ladder(p_Dialogos(i).Texto, 5, -i * 14 + 115 + offset - 22, Color, p_Dialogos(i).status, 1, False, 0, Alpha * (1 - dy))
            End If
        End If

    Next i
    
    If Not force_visible And Segundos <= 0 Then
        If Alpha > 0 Then
            Alpha = Alpha - timerTicksPerFrame * 20
            If Alpha < 0 Then Alpha = 0
        Else
            'RemoveDialogs
        End If
    End If
    
    If offset > 0 Then
        offset = offset - timerTicksPerFrame * 2
        If offset <= 0 Then
            offset = 0
        End If
    End If
    
    
    Exit Sub

Draw_Err:
    Call RegistrarError(Err.number, Err.Description, "clsGuildDlg.Draw", Erl)
    Resume Next
    
End Sub

Public Sub PassTimer()
    
    On Error GoTo PassTimer_Err
    

    If Not p_Activo Then Exit Sub

    If Segundos > 0 Then
        Segundos = Segundos - 1
    End If
        
    Exit Sub

PassTimer_Err:
    Call RegistrarError(Err.number, Err.Description, "clsGuildDlg.PassTimer", Erl)
    Resume Next
    
End Sub

''
' Splits dialogs into lines fitting properly the render area and inserts them

Public Sub PushBackText(ByVal s As String, ByVal status As Byte)
    
    On Error GoTo PushBackText_Err

    '***************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last Modification: 04/07/08
    'Splits dialogs into lines fitting properly the render area and inserts them
    '***************************************************
    Dim outputText() As String
    Dim i As Long
    Dim separatorPos As Long
    
    outputText = Prepare_Multiline_Text(s, 290)
    
    separatorPos = InStr(1, outputText(0), ">")
    If separatorPos > 0 Then
        outputText(0) = Chr(1) & Left(outputText(0), separatorPos) & Chr(1) & Right(outputText(0), Len(outputText(0)) - separatorPos)
    End If
    
    'Show message and continue with the rest
    For i = 0 To UBound(outputText)
        Call RealPushBackText(outputText(i), status)
    Next i
        
    Exit Sub

PushBackText_Err:
    Call RegistrarError(Err.number, Err.Description, "clsGuildDlg.PushBackText", Erl)
    Resume Next
    
End Sub

Private Sub RealPushBackText(ByVal s As String, ByVal status As Byte)
    
    On Error GoTo RealPushBackText_Err

    Dim i     As Integer

    
    For i = (p_CANTIDADDIALOGOS - 1) To 1 Step -1
        p_Dialogos(i + 1) = p_Dialogos(i)
    Next i
    
    p_Dialogos(1).Texto = s
    p_Dialogos(1).status = status
    Segundos = p_SEGUNDOS
    offset = DIALOG_HEIGHT
    Alpha = 255
    Exit Sub

RealPushBackText_Err:
    Call RegistrarError(Err.number, Err.Description, "clsGuildDlg.RealPushBackText", Erl)
    Resume Next
    
End Sub

''
' Retrieves the font being used to render dialogs.
'
' @return   The font being used to render dialogs.

Public Property Get font() As StdFont
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 04/27/06
    'Retrieves the font being used to render dialogs.
    '**************************************************************
    
    On Error GoTo font_Err
    
    font = usedFont

    
    Exit Property

font_Err:
    Call RegistrarError(Err.number, Err.Description, "clsGuildDlg.font", Erl)
    Resume Next
    
End Property

''
' Sets the font being used to render dialogs.
'
' @param    nValue  The new font to be used to render dialogs.

Public Property Let font(ByRef nValue As StdFont)
    '**************************************************************
    'Author: Juan Martín Sotuyo Dodero
    'Last Modify Date: 04/27/06
    'Sets the font being used to render dialogs.
    '**************************************************************
    
    On Error GoTo font_Err
    
    Set usedFont = nValue

    
    Exit Property

font_Err:
    Call RegistrarError(Err.number, Err.Description, "clsGuildDlg.font", Erl)
    Resume Next
End Property

' Toggle visibility of guild dialogs when user opens guild chatbox
Public Sub toggle_dialogs_visibility(ByVal visible As Boolean)

    force_visible = visible
    ' Reset the alpha to make the chat visible
    Alpha = 255

End Sub

